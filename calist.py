"""Calist — compile device inspection forms into one equipment register.

Extracts fixed cells from many device-inspection Excel forms and compiles them
into a copy of a pre-defined template.

The device type is read from the filename: everything after the first "-" is
scanned for its leading letters, so "Clinic-AGH001.xlsx" yields the code "AGH".
That code selects a cell map from device_config.py, which says where each field
lives on that device's form. Data is always read from the first worksheet.

Output columns (template column B onwards):
    Device | Manufacturer | Model | S.N | Location | Code | Date | Status
"""

from __future__ import annotations

import logging
import os
import re
import threading
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, Iterator

import xlrd
from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_to_tuple

from device_config import DEVICE_CONFIGS

# ──────────────────────────────────────────────────────────────────────────────
# Configuration
# ──────────────────────────────────────────────────────────────────────────────

#: Fields written to the template, in column order (template column B onwards).
FIELDS = ["Device", "Manufacturer", "Model", "S.N", "Location", "Code", "Date", "Status"]

TEMPLATE_START_ROW = 4          # first data row; rows 1-3 are headers
TEMPLATE_START_COL = 2          # column B (1=A, 2=B, ...)

OUTPUT_NAME = "device list.xlsx"
SUPPORTED_EXTENSIONS = (".xlsx", ".xlsm", ".xls")

#: A file whose code isn't in DEVICE_CONFIGS is skipped with an error. Set this
#: to False to fall back to UNKNOWN_FALLBACK instead (the old behaviour), which
#: emits a row of whatever happens to be in A1:A6.
SKIP_UNKNOWN_CODES = True

UNKNOWN_FALLBACK = {
    "device_name": "Unknown Device",
    "cells": {
        "Manufacturer": "A1", "Model": "A2", "S.N": "A3",
        "Location": "A4", "Date": "A5", "Status": "A6",
    },
}

#: Device pairs permitted to share one serial number, because the second row is
#: generated from the same physical unit. Names must match device_config.py.
ALLOWED_SHARED_SN_PAIRS = {
    frozenset({"Patient Monitor", "NIBP"}),
    frozenset({"Vital Sign (SPO2 Module)", "Vital Sign (NIBP Module)"}),
}

Record = dict[str, str]

log = logging.getLogger("aggregator")


# ──────────────────────────────────────────────────────────────────────────────
# Structured results
#
# The pipeline reports progress twice over: as log records (the seam the GUI
# and headless callers both read), and as these dataclasses, which carry the
# same facts in a form a table can render. Neither replaces the other.
# ──────────────────────────────────────────────────────────────────────────────

#: A file's state. "ready" is pre-flight only — it means the filename resolved
#: to a known device, not that the workbook has been opened.
READY = "ready"
OK = "ok"
UNSUPPORTED = "unsupported"
UNKNOWN_CODE = "unknown_code"
ERROR = "error"
CANCELLED = "cancelled"

#: Statuses that mean a file contributed nothing to the register.
PROBLEM_STATUSES = frozenset({UNSUPPORTED, UNKNOWN_CODE, ERROR})


@dataclass
class FileOutcome:
    """What happened to one source file."""

    filename: str
    path: str
    status: str
    device_code: str | None = None
    device_name: str | None = None
    detail: str = ""
    rows: int = 0               # 1, or 2 where a second_row was generated

    @property
    def is_problem(self) -> bool:
        return self.status in PROBLEM_STATUSES


@dataclass
class RunResult:
    """The outcome of one full run."""

    output_path: Path | None = None
    rows_written: int = 0
    outcomes: list[FileOutcome] = field(default_factory=list)
    duplicates_removed: int = 0
    second_rows_added: int = 0
    cancelled: bool = False
    error: str | None = None

    @property
    def succeeded(self) -> bool:
        return self.output_path is not None

    @property
    def files_read(self) -> int:
        return sum(1 for o in self.outcomes if o.status == OK)

    @property
    def problems(self) -> list[FileOutcome]:
        return [o for o in self.outcomes if o.is_problem]


#: Called after each file with (outcome, index, total). Fires on the worker
#: thread, so a GUI caller must marshal back to its own loop.
ProgressHook = Callable[[FileOutcome, int, int], None]


# ──────────────────────────────────────────────────────────────────────────────
# Value normalisation
# ──────────────────────────────────────────────────────────────────────────────

def clean(value: object) -> str:
    """Render a raw cell value as the string that belongs in the output.

    Straight str() of whatever Excel hands back. Note this means a real date
    cell is written as "2024-01-15 00:00:00" and a numeric serial as
    "123456.0"; the forms in use hold these as text, so it does not arise.
    """
    return "" if value is None else str(value).strip()


# ──────────────────────────────────────────────────────────────────────────────
# Filename → device code
# ──────────────────────────────────────────────────────────────────────────────

def classify_file(filepath: str) -> FileOutcome:
    """Work out what a file *would* produce, without opening it.

    Extension check plus a filename-to-config lookup — no workbook I/O, so this
    is cheap enough to run on every file the moment it is added. That is what
    lets an unrecognised device code surface before a long run rather than
    after it.
    """
    path = Path(filepath)
    filename = path.name

    if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        return FileOutcome(filename, filepath, UNSUPPORTED,
                           detail=f"Unsupported format '{path.suffix}'")

    code = extract_device_code(filename)
    config = DEVICE_CONFIGS.get(code)

    if config is None:
        shown = f"'{code}'" if code else "none found"
        return FileOutcome(filename, filepath, UNKNOWN_CODE, device_code=code,
                           detail=f"code {shown} is not in the device table")

    name = config["device_name"]
    rows = 2 if "second_row" in config else 1
    if rows == 2:
        name = f"{name}  +{config['second_row']['device_name']}"

    return FileOutcome(filename, filepath, READY, device_code=code,
                       device_name=name, rows=rows)


def extract_device_code(filename: str) -> str | None:
    """Pull the device-type code out of a filename.

    Everything after the first "-" is used when one is present, otherwise the
    whole stem; the leading run of letters is the code.

        "Clinic-AGH001.xlsx" -> "AGH"
        "VNT023.xlsx"        -> "VNT"
    """
    stem = Path(filename).stem
    _, separator, tail = stem.partition("-")
    match = re.match(r"[A-Za-z]+", tail if separator else stem)
    return match.group(0).upper() if match else None


# ──────────────────────────────────────────────────────────────────────────────
# Reading source workbooks
# ──────────────────────────────────────────────────────────────────────────────

CellGetter = Callable[[str], object]


@contextmanager
def _open_source(filepath: str) -> Iterator[CellGetter]:
    """Yield a function that reads a cell (by A1 reference) from the first sheet.

    Hides the difference between openpyxl (.xlsx/.xlsm) and xlrd (.xls) so that
    both formats share one extraction path.
    """
    if Path(filepath).suffix.lower() == ".xls":
        workbook = xlrd.open_workbook(filepath)
        sheet = workbook.sheet_by_index(0)

        def get(ref: str) -> object:
            row, col = coordinate_to_tuple(ref)
            try:
                # Raw value: xlrd reports a date as a bare Excel serial number,
                # which clean() will render as e.g. "45306.0".
                return sheet.cell_value(row - 1, col - 1)
            except IndexError:
                return None

        yield get
    else:
        # data_only=True returns the *cached* result of any formula. A file
        # written by a script and never opened in Excel has no cache, so those
        # cells come back as None.
        workbook = load_workbook(filepath, data_only=True)
        try:
            sheet = workbook.worksheets[0]
            yield lambda ref: sheet[ref].value
        finally:
            workbook.close()


def read_record(filepath: str, cell_map: dict[str, str]) -> Record:
    """Read every mapped cell out of one source file."""
    with _open_source(filepath) as get:
        return {
            field: clean(get(ref)) if ref else ""
            for field, ref in cell_map.items()
        }


# ──────────────────────────────────────────────────────────────────────────────
# Building records
# ──────────────────────────────────────────────────────────────────────────────

def build_second_row(record: Record, second_row: dict) -> Record:
    """Derive the sub-module row that accompanies a two-row device.

    Same data as the parent, but with its own device name, its status taken
    from the parent's Status2 cell, and its code token rewritten.
    """
    extra = dict(record)
    extra["Device"] = second_row["device_name"]
    extra["Status"] = record.get("Status2", "")
    extra["_row_order"] = 1

    old_token, new_token = second_row["code_replace"]
    # count=1: replace only the device code, never a site prefix that happens
    # to contain the same letters.
    extra["Code"] = re.sub(re.escape(old_token), new_token, extra["Code"], count=1)
    return extra


def extract_records(
    source_files: Iterable[str],
    on_file: ProgressHook | None = None,
    cancel: threading.Event | None = None,
) -> tuple[list[Record], list[FileOutcome]]:
    """Read every source file, returning the records and a per-file outcome.

    ``on_file(outcome, index, total)`` fires after each file. ``cancel`` is
    checked between files, so a long run can be stopped without waiting for it
    to finish; files not reached are reported as CANCELLED.
    """
    ordered = sorted(source_files)
    total = len(ordered)
    records: list[Record] = []
    outcomes: list[FileOutcome] = []

    def report(outcome: FileOutcome, index: int) -> None:
        outcomes.append(outcome)
        if on_file:
            on_file(outcome, index, total)

    for index, filepath in enumerate(ordered, start=1):
        if cancel is not None and cancel.is_set():
            for remaining in ordered[index - 1:]:
                outcomes.append(FileOutcome(os.path.basename(remaining), remaining,
                                            CANCELLED, detail="Run cancelled"))
            log.warning("Cancelled — %d file(s) not processed.", total - index + 1)
            break

        filename = os.path.basename(filepath)
        pre = classify_file(filepath)

        if pre.status == UNSUPPORTED:
            log.warning("%s — %s", filename, pre.detail)
            report(pre, index)
            continue

        if pre.status == UNKNOWN_CODE:
            if SKIP_UNKNOWN_CODES:
                log.error("%s — code '%s' is not in device_config.py; file skipped",
                          filename, pre.device_code)
                report(pre, index)
                continue
            log.warning("%s — code '%s' not found, using fallback cell map",
                        filename, pre.device_code)
            config = UNKNOWN_FALLBACK
        else:
            config = DEVICE_CONFIGS[pre.device_code]

        try:
            record = read_record(filepath, config["cells"])
            record["Code"] = Path(filename).stem
            record["Device"] = config["device_name"]

            # Devices with a probe or tube carry a second serial; show both.
            if record.get("S.N2", "").strip():
                record["S.N"] = f"{record['S.N']}\n({record['S.N2']})"

            # Sort keys, so ordering never has to re-derive the code tokens.
            record["_group"] = record["Code"]
            record["_row_order"] = 0

            records.append(record)
            rows = 1
            log.info(
                "[OK]    %s  →  %s (%s)  |  Model: %s  |  S/N: %s  |  Status: %s",
                filename, config["device_name"], pre.device_code,
                record.get("Model", ""), record.get("S.N", ""), record.get("Status", ""),
            )

            if "second_row" in config:
                extra = build_second_row(record, config["second_row"])
                records.append(extra)
                rows = 2
                log.info(
                    "        ↳ 2nd row added: %s  |  Code: %s  |  Status: %s",
                    extra["Device"], extra["Code"], extra["Status"],
                )

            report(FileOutcome(filename, filepath, OK, pre.device_code,
                               config["device_name"], rows=rows), index)

        except Exception as exc:
            log.error("%s — %s", filename, exc)
            report(FileOutcome(filename, filepath, ERROR, pre.device_code,
                               pre.device_name, detail=str(exc)), index)

    return records, outcomes


# ──────────────────────────────────────────────────────────────────────────────
# Ordering and de-duplication
# ──────────────────────────────────────────────────────────────────────────────

def natural_key(text: str) -> tuple:
    """Sort key that orders embedded numbers numerically (AGH9 before AGH10).

    Each part is tagged with a type rank so text and numbers never end up being
    compared against each other.
    """
    return tuple(
        (1, int(part)) if part.isdigit() else (0, part)
        for part in re.split(r"(\d+)", text.lower())
    )


def sort_records(records: list[Record]) -> list[Record]:
    """Order by parent code, keeping each generated sub-module row behind it."""
    return sorted(records, key=lambda r: (natural_key(r.get("_group", "")),
                                          r.get("_row_order", 0)))


def deduplicate_records(records: list[Record]) -> list[Record]:
    """Drop records that repeat a serial number.

    A device and its generated sub-module row are allowed to share one serial
    (they are the same physical unit); any third record with that serial is
    still dropped. Blank serials are always kept.
    """
    seen: dict[str, list[str]] = {}
    kept: list[Record] = []

    for record in records:
        serial = record.get("S.N", "").strip()
        device = record.get("Device", "")

        if not serial:
            kept.append(record)
            continue

        if serial not in seen:
            seen[serial] = [device]
            kept.append(record)
            continue

        existing = seen[serial]
        if len(existing) == 1 and frozenset(existing + [device]) in ALLOWED_SHARED_SN_PAIRS:
            existing.append(device)
            kept.append(record)
        else:
            log.warning("Skipped duplicate S/N '%s' for '%s' (already recorded by %s)",
                        serial, device, existing)

    return kept


# ──────────────────────────────────────────────────────────────────────────────
# Writing output
# ──────────────────────────────────────────────────────────────────────────────

def resolve_output_path(source_files: list[str], template_file: str) -> Path:
    """Pick where the output goes, refusing to write over the template.

    Windows paths are case-insensitive, so an output called "device list.xlsx"
    would silently overwrite a template named "Device List.xlsx" sitting in the
    same folder.
    """
    output_path = Path(source_files[0]).parent / OUTPUT_NAME
    same_file = (os.path.normcase(os.path.abspath(output_path))
                 == os.path.normcase(os.path.abspath(template_file)))
    if same_file:
        raise ValueError(
            f"The output file ({OUTPUT_NAME}) would overwrite the template. "
            f"Move the template out of {output_path.parent}, or rename it."
        )
    return output_path


def write_output(records: list[Record], template_file: str, output_path: Path) -> None:
    """Fill a copy of the template with the records and save it."""
    workbook = load_workbook(template_file)
    try:
        sheet = workbook.active
        for index, record in enumerate(records):
            row = TEMPLATE_START_ROW + index
            sheet.cell(row=row, column=1, value=index + 1)
            for offset, name in enumerate(FIELDS):
                sheet.cell(row=row, column=TEMPLATE_START_COL + offset,
                           value=record.get(name, ""))
        workbook.save(output_path)
    finally:
        workbook.close()


# ──────────────────────────────────────────────────────────────────────────────
# Orchestration
# ──────────────────────────────────────────────────────────────────────────────

def process_files(
    source_files: list[str],
    template_file: str,
    *,
    deduplicate: bool = False,
    on_file: ProgressHook | None = None,
    cancel: threading.Event | None = None,
) -> RunResult:
    """Run the full pipeline and report what happened.

    Always returns a RunResult; check ``.succeeded`` or ``.output_path``. A
    cancelled run writes nothing and comes back with ``cancelled=True``.
    """
    rule = "─" * 55
    result = RunResult()

    if not source_files:
        result.error = "No source files selected."
        log.error("%s", result.error)
        return result

    log.info(rule)
    log.info("Template : %s", os.path.basename(template_file))
    log.info("Sources  : %d file(s) selected", len(source_files))
    log.info(rule)

    try:
        output_path = resolve_output_path(source_files, template_file)
    except ValueError as exc:
        result.error = str(exc)
        log.error("%s", exc)
        return result

    records, result.outcomes = extract_records(source_files, on_file, cancel)
    result.second_rows_added = sum(1 for o in result.outcomes if o.rows == 2)
    records = sort_records(records)

    if cancel is not None and cancel.is_set():
        result.cancelled = True
        log.warning("Run cancelled — no file written.")
        return result

    if deduplicate:
        before = len(records)
        records = deduplicate_records(records)
        result.duplicates_removed = before - len(records)
        log.info("%d duplicate record(s) removed.", result.duplicates_removed) \
            if result.duplicates_removed else log.info("No duplicates found.")

    if not records:
        result.error = "No valid records extracted. No file created."
        log.warning("%s", result.error)
        return result

    try:
        write_output(records, template_file, output_path)
    except Exception as exc:
        result.error = f"Could not save output file: {exc}"
        log.error("%s", result.error)
        return result

    result.output_path = output_path
    result.rows_written = len(records)

    log.info(rule)
    log.info("✔ Success! %d record(s) written.", len(records))
    log.info("NEW FILE SAVED AT: %s", output_path)
    return result


# ──────────────────────────────────────────────────────────────────────────────
# Entry point
# ──────────────────────────────────────────────────────────────────────────────

def main() -> None:
    """Launch the desktop app.

    The UI is imported here rather than at module scope so that importing this
    module — from a test, a script, or another tool — costs nothing and pulls
    in no GUI toolkit.
    """
    from ui import run
    run()


if __name__ == "__main__":
    main()
