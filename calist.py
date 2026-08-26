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
import sys
import threading
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable, Iterator

import xlrd
from openpyxl import load_workbook
from openpyxl.styles import Font
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

#: Where the reference template lives relative to the app.
TEMPLATE_NAME = "Device List.xlsx"

#: The single source of the version number. calist.spec reads it straight out
#: of this file to stamp the executable's Windows version resource, so the
#: About box and the file's Properties tab cannot drift apart.
__version__ = "1.2.0"

#: Authorship. Written into every register and into the workbook's document
#: properties, so the credit travels with the file rather than living only in
#: the app that made it.
AUTHOR_NAME = "Ahmed Gehad"
AUTHOR_EMAIL = "ahmedgehad2112@gmail.com"
ATTRIBUTION = f"{AUTHOR_NAME} · {AUTHOR_EMAIL}"


def bundled_template() -> Path | None:
    """The reference template shipped with the app, if it can be found.

    A frozen build unpacks its data files to ``sys._MEIPASS``; running from a
    checkout, they sit beside this module. Returning it as a default means a
    freshly downloaded copy is usable without hunting for a template first.
    """
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    candidate = base / "template" / TEMPLATE_NAME
    return candidate if candidate.is_file() else None


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
BAD_FORMAT = "bad_format"
ERROR = "error"
CANCELLED = "cancelled"

#: Statuses that mean a file contributed nothing to the register.
PROBLEM_STATUSES = frozenset({UNSUPPORTED, UNKNOWN_CODE, BAD_FORMAT, ERROR})


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

# ──────────────────────────────────────────────────────────────────────────────
# Optional filename-format check
#
#   G302  -  AGH001  -  0425
#   │        │          └── MMYY: month 01-12, then a two-digit year
#   │        └───────────── device code + unit number
#   └────────────────────── site code: letters then digits
#
# The happy path is ONE precompiled match and nothing else — no splitting, no
# allocation — because this runs on every file the moment it is added. The
# per-part diagnosis below only runs for names that already failed, which costs
# nothing on a folder where everything is named correctly.
# ──────────────────────────────────────────────────────────────────────────────

FILENAME_EXAMPLE = "G302-AGH001-0425"

_NAME_RE = re.compile(r"[A-Za-z]+\d+-[A-Za-z]+\d+-(?:0[1-9]|1[0-2])\d{2}\Z")
_PART_RE = re.compile(r"[A-Za-z]+\d+\Z")
_DATE_RE = re.compile(r"\d{4}\Z")


def check_filename_format(stem: str) -> str | None:
    """None if the stem matches the house format, else why it doesn't."""
    if _NAME_RE.match(stem):
        return None
    return _explain_bad_filename(stem)


def _explain_bad_filename(stem: str) -> str:
    """Say which part is wrong. Only reached for names that already failed."""
    parts = stem.split("-")
    if len(parts) != 3:
        return (f"expected 3 parts like {FILENAME_EXAMPLE}, "
                f"found {len(parts)}")

    site, device, date = parts
    if not _PART_RE.match(site):
        return f"site code '{site}' should be letters then digits, like G302"
    if not _PART_RE.match(device):
        return f"device code '{device}' should be letters then digits, like AGH001"
    if not _DATE_RE.match(date):
        return f"date '{date}' should be 4 digits (MMYY), like 0425"
    if not 1 <= int(date[:2]) <= 12:
        return f"month '{date[:2]}' in '{date}' is not between 01 and 12"
    return f"does not match {FILENAME_EXAMPLE}"


def classify_file(filepath: str, strict_names: bool = False) -> FileOutcome:
    """Work out what a file *would* produce, without opening it.

    Extension check, optional filename-format check, and a filename-to-config
    lookup — no workbook I/O, so this is cheap enough to run on every file the
    moment it is added. That is what lets a bad name or an unrecognised device
    surface before a long run rather than after it.

    With ``strict_names`` the house format is enforced first: when it is on,
    the user has asked for that shape specifically, so a malformed name is the
    finding worth reporting even if a device code could still be salvaged.
    """
    path = Path(filepath)
    filename = path.name

    if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        return FileOutcome(filename, filepath, UNSUPPORTED,
                           detail=f"Unsupported format '{path.suffix}'")

    if strict_names:
        problem = check_filename_format(path.stem)
        if problem:
            return FileOutcome(filename, filepath, BAD_FORMAT, detail=problem)

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
    strict_names: bool = False,
) -> tuple[list[Record], list[FileOutcome]]:
    """Read every source file, returning the records and a per-file outcome.

    ``on_file(outcome, index, total)`` fires after each file. ``cancel`` is
    checked between files, so a long run can be stopped without waiting for it
    to finish; files not reached are reported as CANCELLED. ``strict_names``
    enforces the house filename format, skipping anything that breaks it.
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
        pre = classify_file(filepath, strict_names)

        if pre.status == BAD_FORMAT:
            log.error("%s — %s; file skipped", filename, pre.detail)
            report(pre, index)
            continue

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

        stamp_attribution(sheet, workbook, len(records))
        workbook.save(output_path)
    finally:
        workbook.close()


def stamp_attribution(sheet, workbook, record_count: int) -> None:
    """Sign the register, visibly and in the file's own properties.

    A register gets emailed, printed and filed long after the app that made it
    is out of sight, so the credit belongs in the document rather than only in
    the tool. One line, a blank row clear of the data, plus the Excel document
    properties that show under File → Properties.
    """
    row = TEMPLATE_START_ROW + record_count + 1
    made_on = datetime.now().strftime("%d/%m/%Y")

    cell = sheet.cell(row=row, column=TEMPLATE_START_COL)
    cell.value = f"Generated by Calist — {ATTRIBUTION} — {made_on}"
    cell.font = Font(italic=True, size=9, color="FF808080")

    workbook.properties.creator = ATTRIBUTION
    workbook.properties.lastModifiedBy = ATTRIBUTION
    workbook.properties.title = "Equipment register"
    workbook.properties.description = (
        f"Compiled from {record_count} row(s) by Calist — {ATTRIBUTION}")


# ──────────────────────────────────────────────────────────────────────────────
# Orchestration
# ──────────────────────────────────────────────────────────────────────────────

def process_files(
    source_files: list[str],
    template_file: str,
    *,
    deduplicate: bool = False,
    strict_names: bool = False,
    on_file: ProgressHook | None = None,
    cancel: threading.Event | None = None,
) -> RunResult:
    """Run the full pipeline and report what happened.

    Always returns a RunResult; check ``.succeeded`` or ``.output_path``. A
    cancelled run writes nothing and comes back with ``cancelled=True``.
    ``strict_names`` skips any file whose name breaks the house format.
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

    records, result.outcomes = extract_records(source_files, on_file, cancel,
                                               strict_names)
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
