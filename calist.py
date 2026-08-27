"""Calist — compile device inspection forms into one equipment register.

Extracts fixed cells from many device-inspection Excel forms and compiles them
into a copy of a pre-defined template.

The device type is read from the filename: everything after the first "-" is
scanned for its leading letters, so "Clinic-AGH001.xlsx" yields the code "AGH".
That code selects a cell map from device_config.py, which says where each field
lives on that device's form. Data is read from the first worksheet that holds
anything — the X-ray workbook opens on an empty macro stub, and reading that
one made every X-ray come out blank.

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
from openpyxl.cell.cell import MergedCell
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
__version__ = "1.3.2"

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
    #: The serial as read from the form. Empty until the file is actually
    #: opened — pre-flight resolves a device from the filename alone and never
    #: touches the workbook, so it cannot know this.
    serial: str = ""
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


def _sheet_is_blank(sheet) -> bool:
    """True only for an openpyxl sheet with no cells whatsoever.

    Conservative on purpose: a sheet carrying styling but no values reports a
    dimension larger than A1 and is therefore kept, so this can only ever skip
    a tab that could not have held the data.
    """
    return (sheet.max_row <= 1 and sheet.max_column <= 1
            and sheet["A1"].value is None)


def _pick_sheet(sheets: list, is_blank) -> object:
    """The first sheet that holds anything at all.

    The rule used to be "always index 0", which is right for almost every form.
    The X-ray workbook opens on an empty ``Waveform Dialog`` stub left behind
    by its macros, with the real form on the next tab — so every mapped cell
    read blank, on every X-ray, with no error anywhere.

    Skipping is deliberately limited to sheets that are *unambiguously* empty.
    Choosing by name would be worse: the data sheet is variously called Device
    data, Data entry, Inserting data and Data device across the real forms, so
    a name list would go stale, while "has no cells at all" cannot be the sheet
    a cell map describes.
    """
    for sheet in sheets:
        if not is_blank(sheet):
            return sheet
    return sheets[0]


@contextmanager
def _open_source(filepath: str) -> Iterator[CellGetter]:
    """Yield a function that reads a cell (by A1 reference) from the first sheet.

    Hides the difference between openpyxl (.xlsx/.xlsm) and xlrd (.xls) so that
    both formats share one extraction path.

    Both readers resolve **merged cells**. The forms draw each answer as a box
    spanning two columns, and a merged range stores its value only in the
    top-left cell — every other cell in the range reads as empty. So a cell map
    pointing at the second column of a box (``L17`` of a merged ``K17:L17``)
    silently produced a blank field. Looking through to the anchor makes the
    map work whichever cell of the box it names.
    """
    if Path(filepath).suffix.lower() == ".xls":
        # formatting_info is what carries the merge list. It costs more memory
        # and some files refuse it, which must not turn into a read failure —
        # without it merged_cells is simply empty and behaviour is as before.
        try:
            workbook = xlrd.open_workbook(filepath, formatting_info=True)
        except Exception:
            log.debug("No formatting info for %s; merges unresolved",
                      filepath, exc_info=True)
            workbook = xlrd.open_workbook(filepath)
        sheet = _pick_sheet(workbook.sheets(),
                            lambda s: not (s.nrows and s.ncols))
        merged = getattr(sheet, "merged_cells", ()) or ()

        def get(ref: str) -> object:
            row, col = coordinate_to_tuple(ref)

            def at(r: int, c: int) -> object:
                try:
                    # Raw value: xlrd reports a date as a bare Excel serial
                    # number, which clean() renders as e.g. "45306.0".
                    return sheet.cell_value(r - 1, c - 1)
                except IndexError:
                    return None

            value = at(row, col)
            if value not in (None, ""):
                return value
            # xlrd ranges are 0-based with exclusive upper bounds.
            for rlo, rhi, clo, chi in merged:
                if rlo <= row - 1 < rhi and clo <= col - 1 < chi:
                    return at(rlo + 1, clo + 1)
            return value

        yield get
    else:
        # data_only=True returns the *cached* result of any formula. A file
        # written by a script and never opened in Excel has no cache, so those
        # cells come back as None.
        workbook = load_workbook(filepath, data_only=True)
        try:
            sheet = _pick_sheet(workbook.worksheets, _sheet_is_blank)

            def get(ref: str) -> object:
                cell = sheet[ref]
                # openpyxl hands back a MergedCell for exactly the cells that
                # are covered by a range but are not its anchor. Testing the
                # type rather than an empty value means an ordinary blank cell
                # reads as blank, precisely as before, and never picks up the
                # text of some unrelated block it happens to sit inside.
                if not isinstance(cell, MergedCell):
                    return cell.value
                for rng in sheet.merged_cells.ranges:
                    if (rng.min_row <= cell.row <= rng.max_row
                            and rng.min_col <= cell.column <= rng.max_col):
                        return sheet.cell(rng.min_row, rng.min_col).value
                return None

            yield get
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
            # Kept apart from Code because build_second_row rewrites that.
            record["_source"] = filename

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

            # A dual-serial device carries its second number on its own line
            # for the spreadsheet; a table row wants it on one.
            serial = record.get("S.N", "").replace("\n", "  /  ")
            report(FileOutcome(filename, filepath, OK, pre.device_code,
                               config["device_name"], serial=serial,
                               rows=rows), index)

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


def source_name(record: Record) -> str:
    """The file a record came from, for messages that send the user to it.

    ``_source`` is set when the record is read and survives into a generated
    sub-module row, which is the case ``Code`` cannot cover: build_second_row
    rewrites the code token, so a sub-module's Code names no file on disk.
    """
    return record.get("_source") or record.get("Code", "?")


def deduplicate_records(records: list[Record]) -> list[Record]:
    """Drop records that repeat a serial number.

    A device and its generated sub-module row are allowed to share one serial
    (they are the same physical unit); any third record with that serial is
    still dropped. Blank serials are always kept.

    What gets logged is the two **filenames** involved, because that is what
    the user has to go and open. The device type does not identify which form
    to look at when a round holds a dozen of the same model.
    """
    seen: dict[str, list[tuple[str, str]]] = {}      # serial → [(device, file)]
    kept: list[Record] = []

    for record in records:
        serial = record.get("S.N", "").strip()
        device = record.get("Device", "")
        source = source_name(record)

        if not serial:
            kept.append(record)
            continue

        if serial not in seen:
            seen[serial] = [(device, source)]
            kept.append(record)
            continue

        existing = seen[serial]
        devices = [d for d, _ in existing]
        if (len(existing) == 1
                and frozenset(devices + [device]) in ALLOWED_SHARED_SN_PAIRS):
            existing.append((device, source))
            kept.append(record)
        else:
            first = existing[0][1]
            same = " (the same file)" if first == source else ""
            log.warning(
                "Duplicate serial '%s' — skipped %s, already recorded by %s%s",
                serial, source, first, same)

    return kept


# ──────────────────────────────────────────────────────────────────────────────
# Writing output
# ──────────────────────────────────────────────────────────────────────────────

def resolve_output_path(source_files: list[str], template_file: str,
                        output_dir: str | os.PathLike | None = None) -> Path:
    """Pick where the output goes, refusing to write over the template.

    ``output_dir`` is the folder the user chose; without one the register lands
    beside the first source file, which is the long-standing default.

    Windows paths are case-insensitive, so an output called "device list.xlsx"
    would silently overwrite a template named "Device List.xlsx" sitting in the
    same folder.
    """
    folder = Path(output_dir) if output_dir else Path(source_files[0]).parent
    output_path = folder / OUTPUT_NAME
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
    output_dir: str | os.PathLike | None = None,
    on_file: ProgressHook | None = None,
    cancel: threading.Event | None = None,
) -> RunResult:
    """Run the full pipeline and report what happened.

    Always returns a RunResult; check ``.succeeded`` or ``.output_path``. A
    cancelled run writes nothing and comes back with ``cancelled=True``.
    ``strict_names`` skips any file whose name breaks the house format.
    ``output_dir`` overrides where the register is written.
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
        output_path = resolve_output_path(source_files, template_file, output_dir)
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

def inspect_form(filepath: str) -> int:
    """Print what every mapped cell of one form actually reads.

    The answer to "why is this field blank?". Cell maps are written by reading
    coordinates off a paper form, and a form that has been re-laid-out since
    shows up here immediately: a field reading '' next to a cell reference is
    a map that no longer matches the file.

    Returns a process exit code, so a form with unreadable fields fails loudly
    when this is run from a script.
    """
    path = Path(filepath)
    outcome = classify_file(filepath)
    print(f"file   : {path.name}")
    print(f"device : {outcome.device_name or '—'} ({outcome.device_code or '—'})")

    if outcome.status != READY:
        print(f"\nCannot inspect: {outcome.detail}")
        return 1

    cells = DEVICE_CONFIGS[outcome.device_code]["cells"]
    record = read_record(filepath, cells)
    blank = [f for f, v in record.items() if not v]

    print(f"\n{'field':<14}{'cell':<7}value")
    print("─" * 60)
    for field, ref in cells.items():
        value = record[field].replace("\n", " / ")
        print(f"{field:<14}{ref:<7}{value if value else '(blank)'}")

    if path.suffix.lower() != ".xls":
        workbook = load_workbook(filepath, data_only=True)
        sheet = _pick_sheet(workbook.worksheets, _sheet_is_blank)
        merges = sorted(str(r) for r in sheet.merged_cells.ranges)
        skipped = [ws.title for ws in workbook.worksheets
                   if ws is not sheet and _sheet_is_blank(ws)][:3]
        print(f"\nsheet read: {sheet.title!r} "
              f"(of {len(workbook.worksheets)}: "
              f"{', '.join(ws.title for ws in workbook.worksheets[:6])}"
              f"{' …' if len(workbook.worksheets) > 6 else ''})")
        if skipped and workbook.worksheets[0] is not sheet:
            print(f"skipped empty leading tab(s): {', '.join(skipped)}")
        print(f"merged ranges: {len(merges)}")
        if merges:
            print("  " + ", ".join(merges[:24])
                  + (" …" if len(merges) > 24 else ""))

    if blank:
        print(f"\n{len(blank)} field(s) read blank: {', '.join(blank)}")
        print("Either the form leaves them empty, or device_config.py points at "
              "the wrong cell for this layout.")
        return 1

    print("\nEvery mapped field read a value.")
    return 0


def main() -> None:
    """Launch the desktop app, or inspect a single form.

        python calist.py                    launch the app
        python calist.py --inspect FORM     dump what each mapped cell reads

    The UI is imported inside the launch branch rather than at module scope so
    that importing this module — from a test, a script, or another tool —
    costs nothing and pulls in no GUI toolkit. --inspect keeps that property.
    """
    args = sys.argv[1:]
    if args and args[0] == "--inspect":
        if len(args) != 2:
            print("usage: python calist.py --inspect <form.xlsx>")
            raise SystemExit(2)
        logging.basicConfig(level=logging.WARNING, format="%(levelname)s %(message)s")
        raise SystemExit(inspect_form(args[1]))

    from ui import run
    run()


if __name__ == "__main__":
    main()
