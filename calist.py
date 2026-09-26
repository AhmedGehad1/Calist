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
import time
import zipfile
from contextlib import contextmanager
from dataclasses import dataclass, field, replace
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable, Iterator
from xml.etree.ElementTree import iterparse

import xlrd
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.numbers import BUILTIN_FORMATS, is_date_format
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.utils.datetime import MAC_EPOCH, WINDOWS_EPOCH, from_excel

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
__version__ = "1.7.0"

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
ERROR = "error"
CANCELLED = "cancelled"
#: Not a device form, and the "real device forms only" switch is on: a device
#: list, a template, a name with no customer code. Hidden, not a problem.
LEFT_OUT = "left_out"
#: A copy of a form already in the run — "G302-AGH001-0425 (2)" beside
#: "G302-AGH001-0425", the same serial. Only the original is written.
COPY = "copy"

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
    #: The serial as read from the form. Empty until the file is actually
    #: opened — pre-flight resolves a device from the filename alone and never
    #: touches the workbook, so it cannot know this.
    serial: str = ""
    detail: str = ""
    rows: int = 0               # 1, or 2 where a second_row was generated
    #: The name the register writes in its Code column — the filename with its
    #: typing slips undone (see [repair_name]). The file itself keeps its name.
    code: str = ""

    @property
    def is_problem(self) -> bool:
        return self.status in PROBLEM_STATUSES


@dataclass
class Duplicate:
    """A record dropped because its serial number was already recorded.

    The two filenames are the point: the user has to open both to work out
    which one is wrong, and the device type cannot tell them apart when a round
    holds a dozen of the same model.
    """

    serial: str
    device: str
    dropped: str                # the file whose row was dropped
    kept: str                   # the file that recorded the serial first


@dataclass
class RunResult:
    """The outcome of one full run."""

    output_path: Path | None = None
    rows_written: int = 0
    outcomes: list[FileOutcome] = field(default_factory=list)
    duplicates_removed: int = 0
    #: The dropped records themselves. The log carries the same facts; a
    #: summary needs them in a form it can render without scraping text.
    duplicates: list[Duplicate] = field(default_factory=list)
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

    @property
    def left_out(self) -> list[FileOutcome]:
        """Files the "real device forms only" switch kept out of the register."""
        return [o for o in self.outcomes if o.status == LEFT_OUT]

    @property
    def copies(self) -> list[FileOutcome]:
        """Copies of a form already in the run, left out in favour of it."""
        return [o for o in self.outcomes if o.status == COPY]


#: Called after each file with (outcome, index, total). Fires on the worker
#: thread, so a GUI caller must marshal back to its own loop.
ProgressHook = Callable[[FileOutcome, int, int], None]


# ──────────────────────────────────────────────────────────────────────────────
# Value normalisation
# ──────────────────────────────────────────────────────────────────────────────

def plural(count: int, word: str, many: str | None = None) -> str:
    """"1 file", "2 files" — counts in the log read as sentences, not "file(s)"."""
    return f"{count:,} {word if count == 1 else (many or word + 's')}"


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
# Repairing a filename's typing slips
#
#   G302  -  AGH001  -  0425
#   │        │          └── MMYY: month 01-12, then a two-digit year
#   │        └───────────── device code + unit number
#   └────────────────────── customer (site) code: one letter, then digits
#
# By the owner's decision **the files on disk are never renamed**. What is
# repaired is the name Calist *uses*: the register's Code column, and the id the
# archive export gives a record. The slips are the archive's own, measured over
# its 89,538 workbooks:
#
#   ".G414-CA002-0426"            a leading dot, quote or invisible mark
#   "Copy of B14-BB036-0126"      "Copy of " / "Final " in front
#   "G119-CE007 -0525 "           spaces around a dash, or before ".xlsx"
#   "k116-AX001-0123"             lower case
#   "JO8-AGH001-1025"             a letter O typed for the zero of J08
#   "F29-AI041-0625 Pending"      anything after the date: pending, (2), FAIL
#   "H39-AG025-0233"              a date that cannot be one — 2033, month 15,
#                                 "00324" — taken from the form's own Date
#
# A name that is not the house shape at all ("GE  TEC850  OR  SQAX00871",
# "5071938426") is left exactly as it is: there is nothing to repair it toward.
#
# Pure and I/O-free, like [classify_file], because it runs on every file the
# moment it is added. Only the impossible date needs the form, and that is
# settled after the read by [settle_name_date].
# ──────────────────────────────────────────────────────────────────────────────

FILENAME_EXAMPLE = "G302-AGH001-0425"

#: Site code, device code, and the dash after them. Kept for tools/audit.py.
_CODES_RE = re.compile(r"[A-Za-z]+\d+-[A-Za-z]+\d+-")

#: What may sit in front of a name without being part of it: spaces, dots,
#: quotes, dashes, and the invisible direction marks a copy out of WhatsApp or
#: an Arabic folder leaves behind ("⁯D38-AGH178-0224"). Deliberately a
#: list rather than "anything but a letter": Arabic text in front of a name is
#: a different kind of file, not a slip.
_LEAD = (r"[\s.'\"`´‘’\-_~,;·"
         r"​-‏‪-‮⁠-⁯﻿]*")

_PREFIX = r"(?P<prefix>(?:(?:copy\s+of|final)\s+)*)"

#: The device part must end where it seems to: "H59-BZ00F-1025" is not the
#: 000 template "BZ00" with an F after it, it is a name nobody can repair.
_HOUSE_NAME_RE = re.compile(
    rf"{_LEAD}{_PREFIX}"
    r"(?P<site>[A-Za-z]+\d+)\s*-\s*(?P<tag>[A-Za-z]+\d+)(?![A-Za-z0-9])"
    r"(?:\s*-\s*(?P<date>\d+))?"
    r"(?P<tail>.*)\Z",
    re.IGNORECASE | re.DOTALL,
)

#: Just the customer code at the front, for a name whose device part is past
#: repair: it still says whose device it is.
_SITE_FIRST_RE = re.compile(rf"{_LEAD}{_PREFIX}(?P<site>[A-Za-z]+\d+)\s*-",
                            re.IGNORECASE)

#: "JO8" and "BO3": the second letter is an O typed for a zero. Every customer
#: code in Hospital Codes.xlsx but one is a single letter and digits.
_SITE_O_RE = re.compile(r"([A-Z])O(\d+)\Z")
_CUSTOMER_RE = re.compile(r"[A-Z]\d+\Z")


def _unslip(site: str) -> str:
    """"JO8" -> "J08". Anything else is returned as it is."""
    slipped = _SITE_O_RE.match(site)
    return f"{slipped[1]}0{slipped[2]}" if slipped else site

#: A device numbered 000 (or 00, or 0) is the blank the round's forms are
#: copied from — "D38-BZ000-0126-SINO" is the Sino syringe template.
_TEMPLATE_TAG_RE = re.compile(r"[A-Z]+0+\Z")

#: Text after the date worth a person's attention — "pending", "(2)" — as
#: opposed to a run of dots or dashes, which is a slip of the finger.
_WORDS_RE = re.compile(r"[^\W_]")

#: Where a date in the name comes from. See [NameRepair.date_source].
DATE_NONE, DATE_NAME, DATE_CHECK, DATE_FORM, DATE_UNSETTLED = (
    "", "name", "check", "form", "unsettled")


@dataclass(frozen=True)
class NameRepair:
    """A filename with its typing slips undone. See the section comment."""

    #: The stem as it is on disk.
    original: str
    #: The repaired stem — what the register's Code column says.
    name: str
    site: str = ""
    tag: str = ""
    #: The date in the name, MMYY, once settled.
    date: str = ""
    #: Text after the date that [name] leaves out: " (2)", "-pending".
    tail: str = ""
    #: "Copy of " / "Final " taken off the front.
    prefix: str = ""
    typed_site: str = ""
    typed_date: str = ""
    #: DATE_NAME, the name's own; DATE_CHECK, not a date — waiting for the
    #: form's; DATE_FORM, taken from it; DATE_UNSETTLED, the form had none.
    date_source: str = DATE_NONE
    form_date: str = ""
    #: For a name past repair only: whether it still starts with a customer
    #: code ("H59-BZ00F-1025"). A repaired name answers from [site].
    customer: bool = False

    @property
    def house(self) -> bool:
        """Whether the name has the house shape at all — site, then device."""
        return bool(self.site)

    @property
    def has_customer_code(self) -> bool:
        if self.site:
            return bool(_CUSTOMER_RE.match(self.site))
        return self.customer

    @property
    def is_template(self) -> bool:
        return bool(_TEMPLATE_TAG_RE.match(self.tag))

    @property
    def is_copy(self) -> bool:
        """Text was cut to reach this name, so another file may carry it."""
        return bool(self.tail.strip() or self.prefix)

    @property
    def with_tail(self) -> str:
        """The repaired name with its tail kept — for files whose extra text
        is what tells them apart ("-SEVO" and "-Iso", two vaporizers)."""
        if not self.date:
            return self.name
        return f"{self.site}-{self.tag}-{self.date}{self.tail.rstrip()}"

    @property
    def notes(self) -> list[str]:
        """Each repair that changes what the name *says*. Spaces, case and a
        stray leading dot do not count: nobody needs to check those."""
        notes = []
        if self.prefix:
            notes.append(f"{self.prefix.strip()!r} taken off the front")
        if self.site != self.typed_site.upper():
            notes.append(f"customer code {self.typed_site!r} has a letter O "
                         f"for a zero, read as {self.site}")
        if self.date_source == DATE_FORM:
            notes.append(f"the date in the name, {self.typed_date!r}, is not a "
                         f"real date — {self.date} is from the form's date, "
                         f"{self.form_date!r}")
        elif self.date_source == DATE_UNSETTLED:
            notes.append(f"the date in the name, {self.typed_date!r}, is not a "
                         f"real date, and the form has no date to take instead")
        if _WORDS_RE.search(self.tail):
            notes.append(f"{self.tail.strip()!r} after the date left out")
        return notes


def _real_name_date(digits: str, latest_year: int) -> bool:
    return (len(digits) == 4 and 1 <= int(digits[:2]) <= 12
            and 2000 + int(digits[2:]) <= latest_year)


def repair_name(stem: str, latest_year: int | None = None) -> NameRepair:
    """Undo the typing slips in a filename stem. See the section comment.

    ``latest_year`` is the last year a name's date may be: this year for the
    app, the round after the folder's for the archive export. A later one is a
    slip — "0329" typed for a March 2026 visit.
    """
    match = _HOUSE_NAME_RE.match(stem)
    if not match:
        # Nothing to repair toward. The customer code is still worth knowing,
        # for the "real device forms only" switch — see [left_out_because].
        first = _SITE_FIRST_RE.match(stem)
        return NameRepair(stem, stem, customer=bool(
            first and _CUSTOMER_RE.match(_unslip(first["site"].upper()))))

    latest = datetime.now().year if latest_year is None else latest_year
    typed_site = match["site"]
    site = _unslip(typed_site.upper())
    tag = match["tag"].upper()
    date = match["date"] or ""

    if not date:
        # No date, so nothing to cut after it: "G302-AGH001-june" is written
        # as it is, less its surrounding spaces.
        return NameRepair(stem, f"{site}-{tag}{match['tail'].rstrip()}", site, tag,
                          prefix=match["prefix"], typed_site=typed_site)

    source = DATE_NAME if _real_name_date(date, latest) else DATE_CHECK
    return NameRepair(stem, f"{site}-{tag}-{date}", site, tag, date,
                      tail=match["tail"], prefix=match["prefix"],
                      typed_site=typed_site, typed_date=date, date_source=source)


_TIME_PART_RE = re.compile(r"\s+\d{1,2}:\d{2}(:\d{2})?\s*\Z")


def month_year(text: str) -> str | None:
    """A form's Date as MMYY, or None when it is not a date.

    Takes the shapes the forms write — "11-03-2026", "2026-03-11 00:00:00",
    the bare number xlrd returns for a real date cell in an .xls, and a date an
    engineer forced to text with a leading apostrophe.
    """
    value = _TIME_PART_RE.sub("", (text or "").strip().lstrip("'"))
    if _EXCEL_DATE_NUMBER.match(value):
        try:
            when = from_excel(float(value))
        except (ValueError, OverflowError):
            return None
        return f"{when.month:02d}{when.year % 100:02d}"
    if not is_calendar_date(value):
        return None
    parts = re.findall(r"\d+", value)
    if len(parts[0]) == 4:
        year, month = int(parts[0]), int(parts[1])
    else:
        month, year = int(parts[1]), int(parts[2])
    return f"{month:02d}{year % 100:02d}"


def settle_name_date(repair: NameRepair, form_date: str) -> NameRepair:
    """Replace a name's impossible date with the form's own, if it has one.

    By the owner's decision the form's date is trusted whenever it reads as a
    date. When it does not, the name keeps the digits it was typed with — and
    says so, because a Code nobody can check is worse than one marked.
    """
    if repair.date_source != DATE_CHECK:
        return repair
    mmyy = month_year(form_date)
    if mmyy is None:
        return replace(repair, date_source=DATE_UNSETTLED)
    return replace(repair, name=f"{repair.site}-{repair.tag}-{mmyy}", date=mmyy,
                   date_source=DATE_FORM, form_date=form_date.strip())


def left_out_because(repair: NameRepair) -> str:
    """Why the "real device forms only" switch keeps this file out, or "".

    Only what the name alone can tell. A device list or a blank template named
    like a form is found when it is read — see [read_best]'s ``forms_only``.
    """
    if not repair.has_customer_code:
        return f"no customer code in the name (like {FILENAME_EXAMPLE[:4]})"
    if repair.is_template:
        return f"device number {repair.tag} — a template, not a device"
    return ""


def classify_file(filepath: str, real_forms_only: bool = False) -> FileOutcome:
    """Work out what a file *would* produce, without opening it.

    Extension check, the optional "real device forms only" check, and a
    filename-to-config lookup — no workbook I/O, so this is cheap enough to run
    on every file the moment it is added. That is what lets a bad name or an
    unrecognised device surface before a long run rather than after it.

    With ``real_forms_only`` a file whose name has no customer code, or whose
    device number is 000, is LEFT_OUT before its device is looked up: the user
    asked for device forms only, and those are not.
    """
    path = Path(filepath)
    filename = path.name

    if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        return FileOutcome(filename, filepath, UNSUPPORTED,
                           detail=f"Unsupported format '{path.suffix}'")

    repair = repair_name(path.stem)
    if real_forms_only:
        why = left_out_because(repair)
        if why:
            return FileOutcome(filename, filepath, LEFT_OUT, detail=why,
                               code=repair.name)

    code = extract_device_code(filename)
    config = DEVICE_CONFIGS.get(code)

    if config is None:
        shown = f"'{code}'" if code else "none found"
        return FileOutcome(filename, filepath, UNKNOWN_CODE, device_code=code,
                           detail=f"code {shown} is not in the device table",
                           code=repair.name)

    name = config["device_name"]
    rows = 2 if "second_row" in config else 1
    if rows == 2:
        name = f"{name}  +{config['second_row']['device_name']}"

    return FileOutcome(filename, filepath, READY, device_code=code,
                       device_name=name, rows=rows, code=repair.name)


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
# Finding forms in a folder
# ──────────────────────────────────────────────────────────────────────────────

#: Excel drops a lock file beside every workbook someone has open, named "~$"
#: plus the workbook's name. They are not workbooks — opening one fails with
#: "File is not a zip file" — so a folder anyone is working in used to produce
#: a row of failures for files the user never chose.
LOCK_PREFIX = "~$"

#: Windows FILE_ATTRIBUTE_HIDDEN. os.scandir fills st_file_attributes straight
#: from the directory listing on Windows, so testing it costs no extra call.
_HIDDEN = 0x2


def is_source_file(name: str) -> bool:
    """Whether a filename found by scanning a folder is worth opening.

    Deliberately not applied to files the user picked by hand: choosing a .docx
    should still report "Unsupported format", because saying nothing at all
    about a file someone explicitly selected is worse than an error row.
    """
    lowered = name.lower()
    return (not name.startswith(LOCK_PREFIX)
            and lowered != OUTPUT_NAME.lower()
            and lowered.endswith(SUPPORTED_EXTENSIONS))


def find_source_files(
    folder: str | os.PathLike,
    cancel: threading.Event | None = None,
) -> Iterator[str]:
    """Yield every form under ``folder``, depth first.

    os.scandir hands back entries that already know file-from-directory and,
    on Windows, their attributes too — so the extension test runs before any
    stat and nothing is materialised in full. ``Path.rglob("*")`` followed by
    ``is_file()`` costs a stat per entry instead: tolerable on a local disk,
    minutes on a network share, and all of it on whichever thread called it.

    Unreadable sub-folders are skipped rather than raising, so one permission
    error in a deep tree cannot lose the rest of the round.
    """
    stack = [os.fspath(folder)]
    while stack:
        if cancel is not None and cancel.is_set():
            return
        try:
            entries = os.scandir(stack.pop())
        except OSError as exc:
            log.debug("Skipped a folder while scanning: %s", exc)
            continue

        with entries:
            while True:
                try:
                    entry = next(entries)
                except StopIteration:
                    break
                except OSError as exc:              # pragma: no cover
                    log.debug("Stopped reading a folder: %s", exc)
                    break

                try:
                    if entry.is_dir(follow_symlinks=False):
                        stack.append(entry.path)
                        continue
                    if not is_source_file(entry.name):
                        continue
                    if getattr(entry.stat(), "st_file_attributes", 0) & _HIDDEN:
                        continue
                except OSError:                     # vanished mid-scan
                    continue
                yield entry.path


# ──────────────────────────────────────────────────────────────────────────────
# Reading source workbooks
# ──────────────────────────────────────────────────────────────────────────────

CellGetter = Callable[[str], object]

# ── The .xlsx/.xlsm package, read directly ────────────────────────────────────
#
# load_workbook() costs ~500ms on these forms — it parses every worksheet, the
# whole styles table, the drawings and the calc chain, to reach seven cells.
# Reading the package directly costs ~2ms: the workbook relationships, one
# sheet's bytes, and the shared strings or the styles only when a wanted cell
# turns out to need them.
#
# Everything below reproduces openpyxl's own answers deliberately — which sheet
# it would have picked, how it renders a number or a date, how it looks through
# a merged range. A whole-grid comparison against the old reader is the check
# that keeps it honest; see the reader tests in test_calist.py.

_MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_DOC_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PKG_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"

_WORKBOOK_PART = "xl/workbook.xml"
_WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"
_SHARED_STRINGS_PART = "xl/sharedStrings.xml"
_STYLES_PART = "xl/styles.xml"

#: One cell element, matched from the offset of its own opening tag.
#:
#: The lazy quantifier on the attribute run is load-bearing. Greedy, it eats
#: the "/" of a self-closing ``<c r="E18" s="168"/>``, then matches the ">"
#: branch and swallows everything up to the *next* cell's ``</c>`` — silently
#: returning a neighbouring cell's value. test_a_self_closing_cell_does_not_
#: swallow_the_next_one pins this.
_CELL_RE = re.compile(rb'<c r="[A-Z]+\d+"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
_CELL_TYPE_RE = re.compile(rb'\bt="([^"]+)"')
_CELL_STYLE_RE = re.compile(rb'\bs="(\d+)"')
_VALUE_RE = re.compile(rb'<v[^>]*>(.*?)</v>', re.S)
_TEXT_RUN_RE = re.compile(rb'<t[^>]*>(.*?)</t>', re.S)
_MERGE_RE = re.compile(rb'<mergeCell\s+ref="([^"]+)"')
_OUTSIDE_A1_RE = re.compile(rb'<c r="(?!A1")')

_ENTITY_RE = re.compile(rb'&(?:#(\d+)|#x([0-9a-fA-F]+)|(amp|lt|gt|quot|apos));')
_NAMED_ENTITIES = {b"amp": b"&", b"lt": b"<", b"gt": b">",
                   b"quot": b'"', b"apos": b"'"}


def _unescape(raw: bytes) -> str:
    """Decode the entities an XML parser would have decoded for us."""
    if b"&" not in raw:
        return raw.decode("utf-8")

    def replace(match: re.Match) -> bytes:
        decimal, hexadecimal, named = match.groups()
        if decimal:
            return chr(int(decimal)).encode("utf-8")
        if hexadecimal:
            return chr(int(hexadecimal, 16)).encode("utf-8")
        return _NAMED_ENTITIES[named]

    return _ENTITY_RE.sub(replace, raw).decode("utf-8")


def _cast_number(text: str) -> int | float:
    """openpyxl's rule for a numeric cell, reproduced.

    Kept here rather than imported from openpyxl's private worksheet reader so
    that a version bump cannot quietly change how a number reaches clean() —
    "123456" must stay an int and "123456.0" a float, because that difference
    is what ends up written into the register.
    """
    if "." in text or "E" in text or "e" in text:
        return float(text)
    return int(text)


def _sheet_has_content(data: bytes) -> bool:
    """True unless this worksheet is unambiguously empty.

    The openpyxl equivalent asked whether max_row/max_column reach past A1, or
    A1 itself holds a value. Both count a *styled* but valueless cell, so this
    looks for cell elements, not for values — a sheet carrying only formatting
    is kept, exactly as before.
    """
    if _OUTSIDE_A1_RE.search(data):
        return True
    at = data.find(b'<c r="A1"')
    if at < 0:
        return False
    match = _CELL_RE.match(data, at)
    body = match.group(2) if match else None
    return bool(body) and (b"<v" in body or b"<is" in body)


class _XlsxSource:
    """One worksheet of an .xlsx/.xlsm package, held as raw XML."""

    def __init__(self, filepath: str):
        self._archive = zipfile.ZipFile(filepath)
        try:
            self._parts, self._epoch = self._workbook_parts()
            self._data = self._populated_sheet(self._parts)
            self._check_attribute_order()
            self._merges = self._merged_ranges()
            self._strings: list[str] | None = None
            self._date_styles: dict[int, bool] = {}
        except Exception:
            self._archive.close()
            raise

    def _check_attribute_order(self) -> None:
        """Excel, LibreOffice and openpyxl all write r as a cell's first
        attribute, and the lookups below rely on it. Two memchr scans
        (~0.02 ms) prove it for the file in hand rather than assuming it,
        because the failure would otherwise be a silently blank field.
        """
        if self._data.count(b"<c ") != self._data.count(b'<c r="'):
            raise ValueError(
                "cell references are not written where this reader can "
                "find them (unexpected attribute order)")

    def select_sheet(self, name: str) -> bool:
        """Re-point this source at another tab of the same workbook.

        Used only when the sheet chosen on open turns out not to carry the
        device details — GC files put them on 'Data entry' behind a 'Report'
        tab, BM files on 'cover page'. Re-reading one part of an already-open
        archive is far cheaper than opening the file again.
        """
        for candidate, part in self._parts:
            if candidate != name:
                continue
            self._data = self._archive.read(part)
            self._check_attribute_order()
            self._merges = self._merged_ranges()
            self.sheet_name = name
            return True
        return False

    def close(self) -> None:
        self._archive.close()

    # ── package structure ────────────────────────────────────────────────────

    def _workbook_parts(self) -> tuple[list[tuple[str, str]], datetime]:
        """The worksheet names and parts, in the order openpyxl would list them.

        Two rules matter and are easy to get wrong. An older .xlsm carries its
        VBA modules as <sheet> entries with an empty r:id — openpyxl drops
        those with a warning, so they must not shift the ordering here.
        Chartsheets go to wb.chartsheets rather than wb.worksheets and are
        skipped; **dialogsheets are not** — openpyxl counts them, and the X-ray
        workbook opens on one.
        """
        targets: dict[str, tuple[str, str]] = {}
        with self._archive.open(_WORKBOOK_RELS) as handle:
            for _, element in iterparse(handle):
                if element.tag == _PKG_REL_NS + "Relationship":
                    kind = (element.get("Type") or "").rsplit("/", 1)[-1]
                    targets[element.get("Id")] = (kind, element.get("Target"))

        parts: list[tuple[str, str]] = []
        epoch = WINDOWS_EPOCH
        with self._archive.open(_WORKBOOK_PART) as handle:
            for _, element in iterparse(handle):
                if element.tag == _MAIN_NS + "workbookPr":
                    if element.get("date1904") in ("1", "true"):
                        epoch = MAC_EPOCH
                elif element.tag == _MAIN_NS + "sheet":
                    rel = element.get(_DOC_REL_NS + "id")
                    if not rel:
                        continue
                    kind, target = targets.get(rel, (None, None))
                    if target is None or kind == "chartsheet":
                        continue
                    parts.append((element.get("name") or "", _resolve_part(target)))

        if not parts:
            raise ValueError("workbook has no worksheets")
        return parts, epoch

    def _populated_sheet(self, parts: list[tuple[str, str]]) -> bytes:
        """The first worksheet that holds anything at all.

        Always reading sheet one is right for almost every form. The X-ray
        workbook opens on an empty ``Waveform Dialog`` stub left behind by its
        macros, with the real form on the next tab — so every mapped cell read
        blank, on every X-ray, with no error anywhere.

        The names are kept for --inspect, which has to be able to say which tab
        it actually read; a diagnostic that names the wrong sheet is worse than
        none.
        """
        self.sheet_names = [name for name, _ in parts]
        self.sheet_name = self.sheet_names[0]
        self.skipped_sheets: list[str] = []

        first: bytes | None = None
        for name, part in parts:
            data = self._archive.read(part)
            if first is None:
                first = data
            if _sheet_has_content(data):
                self.sheet_name = name
                return data
            self.skipped_sheets.append(name)

        self.skipped_sheets = []
        return first if first is not None else b""

    def merged_refs(self) -> list[str]:
        """Every merged range on the sheet that was read, as A1 references."""
        return sorted(
            f"{get_column_letter(left)}{top}:{get_column_letter(right)}{bottom}"
            for top, left, bottom, right in self._merges)

    def _merged_ranges(self) -> list[tuple[int, int, int, int]]:
        ranges = []
        for ref in _MERGE_RE.findall(self._data):
            start, _, end = ref.decode().partition(":")
            top, left = coordinate_to_tuple(start)
            bottom, right = coordinate_to_tuple(end or start)
            ranges.append((top, left, bottom, right))
        return ranges

    # ── cells ────────────────────────────────────────────────────────────────

    def _anchor(self, ref: str) -> str:
        """Look a reference through to the top-left of any range covering it.

        The forms draw each answer as a box spanning two columns, and a merged
        range stores its value only in that top-left cell. A cell map naming
        the second column of a box (``L17`` of a merged ``K17:L17``) otherwise
        reads blank with no error anywhere — which is how the Ultrasound serial
        number disappeared when that form was re-laid-out.
        """
        row, column = coordinate_to_tuple(ref)
        for top, left, bottom, right in self._merges:
            if top <= row <= bottom and left <= column <= right:
                if row == top and column == left:
                    return ref
                return f"{get_column_letter(left)}{top}"
        return ref

    def values(self, refs: Iterable[str]) -> dict[str, object]:
        """Read the given references, resolving merges and shared strings."""
        anchors = {ref: self._anchor(ref) for ref in refs}

        raw: dict[str, tuple[str, int, bytes | None]] = {}
        for anchor in set(anchors.values()):
            at = self._data.find(b'<c r="' + anchor.encode() + b'"')
            if at < 0:
                continue
            match = _CELL_RE.match(self._data, at)
            if match is None:
                continue
            attributes = match.group(1) or b""
            kind = _CELL_TYPE_RE.search(attributes)
            style = _CELL_STYLE_RE.search(attributes)
            raw[anchor] = (kind.group(1).decode() if kind else "n",
                           int(style.group(1)) if style else 0,
                           match.group(2))

        return {ref: self._value(*raw[anchor]) if anchor in raw else None
                for ref, anchor in anchors.items()}

    def _value(self, kind: str, style_id: int, body: bytes | None) -> object:
        """Render one cell the way openpyxl's data_only reader would."""
        if not body:
            return None

        if kind == "inlineStr":
            if b"<rPh" in body:
                # Phonetic runs; openpyxl drops them and keeps the base text.
                # Rather than guess, say so — a wrong field is worse than a
                # named failure.
                raise ValueError("cell carries phonetic runs")
            return "".join(_unescape(run) for run in _TEXT_RUN_RE.findall(body))

        # data_only semantics: a formula cell yields its *cached* result, so a
        # file written by a script and never opened in Excel has none.
        match = _VALUE_RE.search(body)
        if match is None:
            return None
        text = match.group(1)

        if kind == "s":
            return self._shared_string(int(text))
        if kind == "b":
            return bool(int(text))
        if kind in ("str", "e", "d"):
            return _unescape(text)

        number = _cast_number(text.decode())
        if style_id and self._is_date_style(style_id):
            try:
                return from_excel(number, self._epoch)
            except (OverflowError, ValueError):
                return "#VALUE!"
        return number

    # ── side parts, read only when a wanted cell needs them ──────────────────

    def _shared_string(self, index: int) -> str:
        if self._strings is None:
            self._strings = self._read_shared_strings()
        try:
            return self._strings[index]
        except IndexError:
            return ""

    def _read_shared_strings(self) -> list[str]:
        try:
            handle = self._archive.open(_SHARED_STRINGS_PART)
        except KeyError:
            return []
        strings: list[str] = []
        runs: list[str] = []
        with handle:
            for _, element in iterparse(handle, ("end",)):
                if element.tag == _MAIN_NS + "t":
                    runs.append(element.text or "")
                elif element.tag == _MAIN_NS + "si":
                    strings.append("".join(runs))
                    runs.clear()
                    element.clear()
        return strings

    def _is_date_style(self, style_id: int) -> bool:
        """Whether this style's number format makes the cell a date.

        Only reached for a numeric cell that a cell map actually asks for, so
        the styles table — 184KB on these forms — is usually never read at all.
        """
        if style_id not in self._date_styles:
            self._date_styles = self._read_date_styles()
        return self._date_styles.get(style_id, False)

    def _read_date_styles(self) -> dict[int, bool]:
        try:
            handle = self._archive.open(_STYLES_PART)
        except KeyError:
            # No styles table at all: nothing can be a date, and a missing
            # part must not take the whole file down.
            return {}

        formats = dict(BUILTIN_FORMATS)
        applied: list[int] = []
        inside = False
        with handle:
            for event, element in iterparse(handle, ("start", "end")):
                if event == "start":
                    if element.tag == _MAIN_NS + "cellXfs":
                        inside = True
                    continue
                if element.tag == _MAIN_NS + "numFmt":
                    formats[int(element.get("numFmtId"))] = element.get("formatCode")
                elif element.tag == _MAIN_NS + "cellXfs":
                    break          # dxfs follows, and can be far larger
                elif element.tag == _MAIN_NS + "xf" and inside:
                    applied.append(int(element.get("numFmtId", 0)))
                element.clear()
        return {index: is_date_format(formats.get(number) or "")
                for index, number in enumerate(applied)}


def _resolve_part(target: str) -> str:
    """A relationship target as a package path."""
    if target.startswith("/"):
        return target[1:]
    return target if target.startswith("xl/") else "xl/" + target


# ── .xls, through xlrd ────────────────────────────────────────────────────────

class _XlsSource:
    """One sheet of a legacy .xls, read through xlrd."""

    def __init__(self, filepath: str):
        # formatting_info is what carries the merge list. It costs more memory
        # and some files refuse it, which must not turn into a read failure —
        # without it merged_cells is simply empty and behaviour is as before.
        try:
            self._workbook = xlrd.open_workbook(filepath, formatting_info=True,
                                                on_demand=True)
        except Exception:
            log.debug("No formatting info for %s; merges unresolved",
                      filepath, exc_info=True)
            self._workbook = xlrd.open_workbook(filepath, on_demand=True)
        self._sheet = self._populated_sheet()
        self._merges = getattr(self._sheet, "merged_cells", ()) or ()
        self.sheet_names = self._workbook.sheet_names()
        self.sheet_name = self._sheet.name

    def select_sheet(self, name: str) -> bool:
        """Re-point this source at another sheet — see _XlsxSource.select_sheet."""
        for index in range(self._workbook.nsheets):
            if self._workbook.sheet_names()[index] != name:
                continue
            self._sheet = self._workbook.sheet_by_index(index)
            self._merges = getattr(self._sheet, "merged_cells", ()) or ()
            self.sheet_name = name
            return True
        return False

    def close(self) -> None:
        try:
            self._workbook.release_resources()
        except Exception:                                # pragma: no cover
            pass

    def _populated_sheet(self):
        """The first sheet holding anything, loading one at a time.

        workbook.sheets() would load every sheet up front, which is the whole
        cost on_demand=True exists to avoid.
        """
        first = None
        for index in range(self._workbook.nsheets):
            sheet = self._workbook.sheet_by_index(index)
            if first is None:
                first = sheet
            if sheet.nrows and sheet.ncols:
                return sheet
            if sheet is not first:
                self._workbook.unload_sheet(index)
        if first is None:
            raise ValueError("workbook has no sheets")
        return first

    def values(self, refs: Iterable[str]) -> dict[str, object]:
        return {ref: self._value(ref) for ref in refs}

    def _value(self, ref: str) -> object:
        row, column = coordinate_to_tuple(ref)
        value = self._at(row, column)
        if value not in (None, ""):
            return value
        # xlrd ranges are 0-based with exclusive upper bounds.
        for top, bottom, left, right in self._merges:
            if top <= row - 1 < bottom and left <= column - 1 < right:
                return self._at(top + 1, left + 1)
        return value

    def _at(self, row: int, column: int) -> object:
        try:
            # Raw value: xlrd reports a date as a bare Excel serial number,
            # which clean() renders as e.g. "45306.0".
            return self._sheet.cell_value(row - 1, column - 1)
        except IndexError:
            return None


#: The first bytes of each workbook format. A file is read by what it IS, not by
#: what its name says: the archive holds .xlsx workbooks saved with a .xls name
#: (and the reverse), which failed outright when routed by extension.
_ZIP_MAGIC = b"PK\x03\x04"
_OLE2_MAGIC = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"


def _open_workbook(filepath: str):
    try:
        with open(filepath, "rb") as handle:
            head = handle.read(8)
    except OSError:
        head = b""
    if head.startswith(_ZIP_MAGIC):
        return _XlsxSource(filepath)
    if head.startswith(_OLE2_MAGIC):
        return _XlsSource(filepath)
    # Neither signature: let the extension decide, so the error a genuinely
    # broken file raises is the same one it always did.
    if Path(filepath).suffix.lower() == ".xls":
        return _XlsSource(filepath)
    return _XlsxSource(filepath)


@contextmanager
def _open_source(filepath: str) -> Iterator[CellGetter]:
    """Yield a function that reads a cell (by A1 reference) from one workbook.

    A compatibility seam, not the hot path: read_record asks for the whole cell
    map in one call, because resolving merges once for the batch is what makes
    a form cost two milliseconds instead of five hundred.
    """
    source = _open_workbook(filepath)
    try:
        cache: dict[str, object] = {}

        def get(ref: str) -> object:
            if ref not in cache:
                cache.update(source.values([ref]))
            return cache.get(ref)

        yield get
    finally:
        source.close()


def read_record(filepath: str, cell_map: dict[str, str]) -> Record:
    """Read every mapped cell out of one source file."""
    source = _open_workbook(filepath)
    try:
        values = source.values({ref for ref in cell_map.values() if ref})
    finally:
        source.close()
    return {
        field: clean(values.get(ref)) if ref else ""
        for field, ref in cell_map.items()
    }


# ──────────────────────────────────────────────────────────────────────────────
# Reading a form whose map no longer fits
#
# Measured over the 2025 and 2026 rounds — 42,826 forms — the configured map
# works for 94.7% of them. The rest are not broken files: they are forms that
# were re-laid-out between visits, and the shift is uniform, usually one to
# three rows. `alt_cells` records the layouts we know; this is the net for the
# ones nobody has written down yet.
# ──────────────────────────────────────────────────────────────────────────────

#: A serial the site assigned because the device carries none, e.g. BALANCE001.
_ASSIGNED_SERIAL_RE = re.compile(r"^[A-Za-z][A-Za-z ]*\.?\s*\d+$")

#: The test is *any* digit, not a trailing one: STX21170332PA and NV03124H are
#: perfectly good manufacturer serials that end in a letter. What separates a
#: serial from a location typed into the wrong cell is having numbers at all.
_HAS_DIGIT_RE = re.compile(r"\d")

#: Written where a device genuinely has no serial and none was assigned.
#:
#: "0.0" and "0.00" are here because a numeric zero cell arrives as a float and
#: renders that way — a Baby Warmer cover sheet filled in with zeros throughout
#: otherwise reads as a real serial and gets imported as a device.
_PLACEHOLDERS = {"", "-", "--", "n.a", "na", "n/a", "none", "null",
                 "0", "0.0", "0.00", "00"}

#: A date, in any shape these forms use. Excluded by name because a misaligned
#: map lands on a date more often than on anything else, and "16-01-2026" would
#: otherwise pass a digits-only test and be imported as a serial number.
_DATE_LIKE_RE = re.compile(r"^\s*\d{1,4}\s*[-/.]\s*\d{1,2}\s*[-/.]\s*\d{1,4}\s*$")


#: A short decimal is a measurement, not a serial: a layout that has slipped
#: onto a test table reads "7.1", "25.8" or "0.99" — which have digits and so
#: passed as serials. Kept narrow on purpose: a serial exported as a float
#: keeps its length ("1873206.0"), and real serials do look like "37.254"
#: (a Therapeutic Ultrasound) — so only one to three digits before the point
#: and one or two after count as a reading — and never a whole number ending
#: in ".0", which is how an .xls hands back a numeric serial: a Nebulizer's
#: serial 115 arrives as "115.0", and calling that a reading lost 251 fields.
_READING_RE = re.compile(r"^\s*\d{1,3}\.(?!0+\s*$)\d{1,2}\s*$")

#: A computed value — a spreadsheet's float, "0.38271946501827364" — is not a
#: serial either. An old Defibrillator .xls opens on its test sheet, where the
#: map lands on an uncertainty column. Twelve decimals, not fewer: real serials
#: carry up to eight ("1742.60318475", and "437.20981", "6.20417" on 180 forms).
_COMPUTED_RE = re.compile(r"^\s*-?\d*\.\d{12,}\s*$")


def is_calendar_date(text: str) -> bool:
    """True only for text that is an actual calendar date.

    The shape alone is not enough. Water Bath serial "2110/04/98" has the
    shape of a date and is not one — year 2110, day 98 — and treating it as a
    date made the reader reject the right layout and go looking for another.
    """
    match = _DATE_LIKE_RE.match(text or "")
    if not match:
        return False
    parts = [int(p) for p in re.findall(r"\d+", text)]
    if len(parts) != 3:
        return False
    if len(re.findall(r"\d+", text)[0]) == 4:          # 2024-01-15
        year, month, day = parts
    else:                                              # 15-01-2024, 15/1/24
        day, month, year = parts
    if year < 100:
        year += 2000
    return 1 <= month <= 12 and 1 <= day <= 31 and 1950 <= year <= 2099


def classify_serial(serial: str) -> str:
    """One of: blank, placeholder, assigned, real, suspect.

    ``suspect`` is the finding that matters: the cell holds something that is
    not a serial by any reading — the archive has infusion pumps whose serial
    cell says "ICU" — which is how a misaligned map announces itself.
    """
    value = serial.strip()
    if not value:
        return "blank"
    if value.lower() in _PLACEHOLDERS:
        return "placeholder"
    if is_calendar_date(value):
        return "suspect"
    if _READING_RE.match(value) or _COMPUTED_RE.match(value):
        return "suspect"
    if not _HAS_DIGIT_RE.search(value):
        return "suspect"
    if _ASSIGNED_SERIAL_RE.match(value):
        return "assigned"
    return "real"


def plausible(record: Record) -> bool:
    """A record is plausible when the serial looks like one and a model is set.

    Two fields rather than one: a misaligned map that happens to land on some
    other number would pass a serial-only check, and one that lands on a label
    would pass a model-only check. Requiring both is what separates layouts.
    """
    if not record:
        return False
    if classify_serial(clean(record.get("S.N"))) in ("suspect", "blank"):
        return False
    model = clean(record.get("Model"))
    if not model:
        return False
    # A record that is placeholders all the way through is not a reading of
    # anything. Some cover sheets are filled in with "0" throughout, and a
    # serial of "0" classifies as a placeholder rather than a blank — so
    # without this the fallback "rescues" them into a row of zeroes.
    if (model.lower() in _PLACEHOLDERS
            and clean(record.get("S.N")).lower() in _PLACEHOLDERS):
        return False
    # A model is never its own serial. A Phototherapy layout that has moved
    # eight rows lands both references on one stray cell — the text "D38" —
    # and a record reading "D38" / "D38" passed every other test here.
    if " ".join(model.lower().split()) == " ".join(
            clean(record.get("S.N")).lower().split()):
        return False
    return True


#: The labels these forms print beside each field.
_FIELD_LABEL_PATTERNS: list[tuple[str, re.Pattern]] = [
    ("S.N", re.compile(r"^serial\s*(no\.?|number)?\s*:?\s*$|^s\.?\s*n\.?\s*:?\s*$", re.I)),
    ("Model", re.compile(r"^model\s*:?\s*$", re.I)),
    ("Manufacturer", re.compile(r"^manufacturer\s*:?\s*$", re.I)),
    ("Location", re.compile(r"^location\s*:?\s*$", re.I)),
]

#: Headings that introduce the instrument doing the calibrating.
#:
#: This is not a nicety. A Hemodialysis cover page prints the reference meter
#: ABOVE the device under test:
#:
#:     A15  Calibration Device   G15 Model   H15 EMIS      <- the calibrator
#:     A22  Device information
#:     A24  Device name          G24 Model   H24 AK96      <- the actual device
#:
#: so "take the first Model: label" records the calibrator's model and serial
#: as the dialysis machine's. Position cannot be trusted; the headings can.
_GEAR_HEADING = re.compile(
    r"calibration\s*device|reference\s*(meter|equipment|instrument|standard)"
    r"|test\s*equipment|standard\s*used|m\s*&\s*te|calibrator\b|equipment\s*used"
    r"|measuring\s*equipment|traceab",
    re.I,
)
#: The heading that opens the block describing the device being calibrated.
#: "Equipment Data" is the certificate's own heading for the device block. Left
#: out, the certificate's "...traceable..." sentence above it read as a
#: calibrator heading and the whole block was skipped — so no certificate tab
#: could ever rescue a form, and FC's only device block was unreadable.
_DEVICE_HEADING = re.compile(
    r"device\s*(information|data|under\s*test)|equipment\s*under|equipment\s*data", re.I)

_LABEL_COLUMNS = "ABCDEFGHIJKLMN"
_LABEL_MAX_ROW = 95

#: Other things printed on these sheets that are captions, never values.
#:
#: The trailing ``|.*:`` alternative is the load-bearing one. Naming captions
#: individually did not scale: the list already held "safety" and still let
#: "Syringe brand:", "Contact Person Name:" and "Phone No.:" through into the
#: Status column, because those were simply not on it. **Anything ending in a
#: colon is a caption** — no status the archive uses ends in one, and "Pass."
#: is handled by the trailing-stop strip in `normalise_status`.
_OTHER_CAPTIONS = re.compile(
    r"^\s*(date\s*(of\s*receipt)?|issue\s*date.*|status|type|class|"
    r"prev\.?\s*calib\.?|next\s*calib\.?.*|tested\s*by|entered\s*by|"
    r"revised\s*by|safety|remark s?|accessories)\s*:?\s*$"
    r"|^.*\S\s*:\s*$",
    re.I,
)


def _is_a_label(text: str) -> bool:
    """True when a cell holds a caption rather than an answer."""
    return (any(pattern.match(text) for _, pattern in _FIELD_LABEL_PATTERNS)
            or bool(_OTHER_CAPTIONS.match(text)))


# ──────────────────────────────────────────────────────────────────────────────
# What a Status cell says
#
# This vocabulary used to live in firebase_export, which imports this module.
# It moved here because the register needs it too: a mapped Status cell that
# lands on a caption — "Safety:", "Syringe brand:" — was being written into the
# register verbatim, and recognising a real status is what makes it possible to
# refuse one. firebase_export re-exports every name below, so nothing there
# changed.
#
# Keeping ONE vocabulary matters more than where it lives. Two copies would
# drift, and the register and the phone would then disagree about what the same
# cell said.
# ──────────────────────────────────────────────────────────────────────────────

#: A status cell that means the device passed. Anything else — "Fail", a blank,
#: a comment — is recorded as not passed, because only an explicit pass is one.
_PASS_VALUES = {"pass", "passed", "ok", "accepted", "conform", "conforms"}

#: Status cells that unambiguously mean the device failed outright.
_FAIL_VALUES = {"fail", "failed", "faulty"}

#: The device was brought into specification by calibrating it. Its own
#: outcome, and deliberately not added to `_PASS_VALUES`.
_CALIBRATED_VALUES = {"calibrated"}

#: Misspellings that turn up in the archive often enough to matter, found by
#: counting every Status text the classifier refused across all 87,492 forms.
#: Only unambiguous ones: "p" or "ci" could be anything, so they stay refused.
_TYPOS = {
    "passs": "pass",
    "psss": "pass",
    "caibrated": "calibrated",
    "limeted non": "limited non",
}

#: A limitation with no qualifier. "Limited Calibrated" is recorded the same
#: way, by the owner's decision: kept simple rather than read as calibrated.
_LIMITED_VALUES = {"limited", "limited calibrated"}

#: The app's six outcomes, exactly as ``CalibrationStatus.wire`` spells them in
#: ``lib/models/calibration_status.dart``. The two are related only by agreeing
#: on these strings, so a test on each side pins them.
STATUS_PASS = "pass"
STATUS_CALIBRATED = "calibrated"
STATUS_LIMITED_NON = "limited non"
#: A limitation the form does not qualify — "Limited", "Limited Calibrated".
#: Its own outcome rather than a guess at one of the two above.
STATUS_LIMITED = "limited"
STATUS_LIMITED_FAIL = "limited fail"
STATUS_FAIL = "fail"

#: Worse is higher. An unqualified "limited" ranks above "limited non": a
#: limitation nobody said was harmless is not assumed to be.
_SEVERITY = {STATUS_PASS: 0, STATUS_CALIBRATED: 1, STATUS_LIMITED_NON: 2,
             STATUS_LIMITED: 3, STATUS_LIMITED_FAIL: 4, STATUS_FAIL: 5}


def normalise_status(raw: str | None) -> str | None:
    """One Status cell's text as one of the app's outcomes, or None.

    **None is an answer, not a failure.** It means "this cell does not say", and
    the caller leaves the outcome out rather than guessing. Guessing would be
    worse: an unreadable cell classed as a pass puts a date on next year's
    certificate that the device never earned.
    """
    text = " ".join((raw or "").strip().lower().split())
    # A trailing full stop ("Pass.") and the known misspellings.
    text = text.rstrip(" .")
    text = _TYPOS.get(text, text)
    if not text:
        return None
    if text in _PASS_VALUES:
        return STATUS_PASS
    if text in _CALIBRATED_VALUES:
        return STATUS_CALIBRATED
    if text in _FAIL_VALUES:
        return STATUS_FAIL
    if "limited" in text:
        if "fail" in text:
            return STATUS_LIMITED_FAIL
        # "limited non" is the sheet's own spelling; "none" is how it gets
        # typed.
        if "non" in text.split("limited", 1)[1]:
            return STATUS_LIMITED_NON
        # Unqualified: its own outcome, never promoted to either of the two
        # above. Anything else with "limited" in it names nothing we know.
        if text in _LIMITED_VALUES:
            return STATUS_LIMITED
    return None


#: A form's own printed "Status" label — "Status", "Status:", "STATUS".
_STATUS_LABEL = re.compile(r"^\s*status[\s:.]*$", re.I)

#: Answers that belong to a *different* question on these sheets, and so can
#: never be a calibration verdict.
#:
#: "Safety:" is printed beside the status box on most of these forms and its
#: answer is a size. A cell map that lands one column off therefore reports
#: "Large" as the device's outcome — which reads as a real answer to anyone
#: scanning the register, and is not one. The caption guard cannot catch these
#: because they are values, not captions.
_NOT_A_VERDICT = {"small", "large", "n.a", "na", "n/a", "---", "----"}

#: An engineer's code — "JTE-", "STE-18", "JTE-40" — from the "Tested by" and
#: "Revised by" boxes beside the status. Never a verdict: on older Centrifuge
#: forms the map's K25 is "Revised by", captioned ABOVE the box where the
#: caption check cannot see it, and 45 registers rows read "JTE-" as a status.
_ENGINEER_CODE = re.compile(r"^\s*[a-z]te\s*-\s*\d*\s*$", re.I)

#: A bare number in a status box is a test reading the map landed on ("191.2"
#: joules, "5.1" volts), never a verdict — except 0, which is what a formula
#: copying an empty box shows, and which the owner reads as "not tested".
_A_READING = re.compile(r"^\s*-?(?!0+(\.0+)?\s*$)\d+(\.\d+)?\s*$")

#: The columns a status box is ever found in. Deliberately the same set the
#: export has always used, so both sides read a form identically.
_STATUS_COLS = "ABCDEFGHIJKL"


def status_from_grid(grid: dict) -> str:
    """The Status box's text, found from the form's own label. "" if none.

    The fallback for a form whose *mapped* Status cell is blank or holds
    something that is not a status. The cell maps are right for most forms but
    not all, and a box that has moved is the common case: measured over 81
    device maps, 33 of them have forms where the box sits in a different
    *column*, not merely a different row — so an offset search cannot find it
    and only the printed label can.

    Right of the label first, then below, **each direction on its own**: the
    first filled cell in a direction decides that direction, so a "Comment:"
    label to the right does not hide the box underneath. Cells repeating the
    label are its own merged span — merges are resolved — and are stepped over.

    **Only text [normalise_status] recognises is returned.** That is what makes
    searching safe, and it is what keeps the legend off the answer: these forms
    print "Pass | Fail | Limited non | limited fail" across columns B-E on the
    answer's own row, and every one of those cells reads as a status. The
    search never starts from them because it starts from the *label*.
    """
    text = {
        ref: (str(value).strip() if value is not None else "")
        for ref, value in grid.items()
    }

    labels = []
    for ref, value in text.items():
        if not value or not _STATUS_LABEL.match(value):
            continue
        match = re.match(r"([A-Z]+)(\d+)$", ref)
        if match and match.group(1) in _STATUS_COLS:
            labels.append(
                (int(match.group(2)), _STATUS_COLS.index(match.group(1)))
            )

    directions = (
        [(0, 1), (0, 2), (0, 3), (0, 4)],   # to the right
        [(1, 0), (2, 0), (3, 0)],           # below
    )
    for row, col in sorted(labels):         # reading order
        for offsets in directions:
            for down, right in offsets:
                column = col + right
                if column >= len(_STATUS_COLS):
                    break
                value = text.get(f"{_STATUS_COLS[column]}{row + down}", "")
                if not value or _STATUS_LABEL.match(value):
                    continue
                if normalise_status(value):
                    return value
                break                       # this direction holds no status
    return ""


def locate_by_labels(values: dict[str, object]) -> dict[str, str]:
    """Work out where the fields sit from the form's own printed labels.

    ``values`` is a whole-grid read of A1:N95. Returns a cell map, which may be
    partial — a caller should check it produced a plausible record before
    trusting it.

    Only cells inside the device's own section are considered; see
    ``_GEAR_HEADING``. Where the form has no such headings at all, the first
    match top-down wins, which is what every single-block form wants.
    """
    text = {ref: clean(value) for ref, value in values.items()}

    device_row = gear_row = None
    for row in range(1, _LABEL_MAX_ROW + 1):
        for column in _LABEL_COLUMNS:
            body = text.get(f"{column}{row}", "")
            if not body:
                continue
            if device_row is None and _DEVICE_HEADING.search(body):
                device_row = row
            if _GEAR_HEADING.search(body) and (gear_row is None or row < gear_row):
                gear_row = row

    def in_device_section(row: int) -> bool:
        if device_row is not None:
            # The device block runs from its heading to the next gear heading
            # below it, or to the end of the sheet.
            end = gear_row if gear_row is not None and gear_row > device_row else None
            return row >= device_row and (end is None or row < end)
        if gear_row is not None:
            return row < gear_row
        return True

    found: dict[str, str] = {}
    for row in range(1, _LABEL_MAX_ROW + 1):
        if not in_device_section(row):
            continue
        for index, column in enumerate(_LABEL_COLUMNS):
            body = text.get(f"{column}{row}", "")
            if not body:
                continue
            for field, pattern in _FIELD_LABEL_PATTERNS:
                if field in found or not pattern.match(body):
                    continue
                # Merges make a label repeat across its own span, so the value
                # is the first cell to the right whose text differs — and which
                # is not itself a label. The Therapeutic Ultrasound form heads a
                # table with "Model | S.N." at row 11 and puts the real fields
                # at row 69; without this the model reads as the string "S.N.".
                for other in _LABEL_COLUMNS[index + 1:]:
                    candidate = text.get(f"{other}{row}", "")
                    if not candidate or candidate == body:
                        continue
                    if _is_a_label(candidate):
                        break
                    found[field] = f"{other}{row}"
                    break
    return found


#: Captions for the device's own date, most telling first: the data tab's
#: "Date of receipt:", a cover page's "Test Date", a certificate's "Calib.
#: Date:". Never "Issue Date", "Last Cal. Date", "Prev. Calib." or "Next Calib.
#: On" — those are other dates, and the last two are a year away from it.
_DATE_CAPTIONS = (re.compile(r"^\s*date\s*of\s*receipt\s*:?\s*$", re.I),
                  re.compile(r"^\s*test\s*date\s*:?\s*$", re.I),
                  re.compile(r"^\s*calib(\.|ration)?\s*date\s*:?\s*$", re.I))


def date_by_caption(values: dict[str, object]) -> str:
    """The date printed right of the device's date caption, or "".

    ``values`` is a whole-grid read, as for [locate_by_labels]. Only a value
    that is unmistakably a date is taken: the cell right of a caption is
    sometimes the next caption, or blank.
    """
    text = {ref: clean(value) for ref, value in values.items()}
    for caption in _DATE_CAPTIONS:
        for ref in sorted((r for r, t in text.items() if caption.match(t)),
                          key=coordinate_to_tuple):
            row, column = coordinate_to_tuple(ref)
            for c in range(column + 1, len(_LABEL_COLUMNS) + 1):
                value = text.get(f"{get_column_letter(c)}{row}", "")
                if not value or caption.match(value):
                    continue
                if _looks_like_date(value):
                    return value
                break
    return ""


# ──────────────────────────────────────────────────────────────────────────────
# Repairing a typed date
#
# By the owner's decision the register repairs an engineer's slip in a date —
# "21--01-2024", "10-04-02025", "12/19/2022" — and never the form itself. A
# repair is made only when the value is not a date as typed, and only to the
# one date that the file's own name confirms: its month and year are the
# `-0124` at the end of `D22-AGH001-0124`, each part may differ from them by
# one slipped digit at most, and the day is the engineer's. A name without
# that date gets no repair. A test PERIOD ("30/04/2023 - 18/06/2023", which
# the certificates print too) is not a slip: the register takes its start.
# The cell is marked, and its note says what the form says.
# ──────────────────────────────────────────────────────────────────────────────

_FILED_RE = re.compile(r"[A-Za-z]+\d+-[A-Za-z]+\d+-(0[1-9]|1[0-2])(\d{2})")
_ONE_DATE = r"(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})"
_PERIOD_RE = re.compile(rf"^\s*{_ONE_DATE}\s*(?:-|–|—|…|\.{{2,3}}|to)\s*{_ONE_DATE}\s*$", re.I)


def _a_date(day: int, month: int, year: int) -> str | None:
    try:
        datetime(year, month, day)
    except ValueError:
        return None
    return f"{day:02d}-{month:02d}-{year}" if 1950 <= year <= 2099 else None


def _one_slip(typed: str, meant: str) -> bool:
    """Whether ``typed`` is ``meant`` or one digit away from it."""
    if typed == meant:
        return True
    if abs(len(typed) - len(meant)) > 1:
        return False
    if len(typed) == len(meant):
        return sum(a != b for a, b in zip(typed, meant)) == 1
    short, long = sorted((typed, meant), key=len)
    return any(long[:i] + long[i + 1:] == short for i in range(len(long)))


def repair_date(typed: str, filepath: str) -> tuple[str, str] | None:
    """The date a mistyped Date means, and why — or None when it cannot be
    settled. See the section comment above. A date as typed is never touched,
    whatever the file name says."""
    if _looks_like_date(typed or ""):
        return None
    period = _PERIOD_RE.match(typed or "")
    if period:
        day, month, year = (int(g) for g in period.groups()[:3])
        start = _a_date(day, month, year)
        end = _a_date(*(int(g) for g in period.groups()[3:]))
        if start and end:
            return start, (f"The form gives a period, {typed.strip()!r}; "
                           f"the register shows its start.")
        return None

    filed = _FILED_RE.search(Path(filepath).stem)
    groups = re.findall(r"\d+", typed or "")
    if not filed or not 2 <= len(groups) <= 3:
        return None
    month, year = filed.group(1), f"20{filed.group(2)}"
    if int(year) > datetime.now().year:
        # "H39-AG025-0233" is itself a slip — 2033 is still to come — so it
        # confirms nothing; the name is the one repaired from the form.
        return None
    if len(groups) == 3:
        day, typed_month, typed_year = groups
    elif len(groups[1]) == 6:           # "15-092025": month and year run together
        day, typed_month, typed_year = groups[0], groups[1][:2], groups[1][2:]
    elif len(groups[0]) == 4:           # "0712-2025": day and month run together
        day, typed_month, typed_year = groups[0][:2], groups[0][2:], groups[1]
    else:
        return None
    if len(day) > 2:
        return None
    # Read as typed first, only the separators cleaned: "21--03-2024" is 21
    # March, whatever the name says — and when it is not the name's month, it
    # is left alone rather than bent to fit.
    as_typed = (_a_date(int(day), int(typed_month), int(typed_year))
                if len(typed_month) <= 2 and len(typed_year.lstrip("0")) == 4 else None)
    if as_typed:
        if as_typed[3:] != f"{month}-{year}":
            return None
        return as_typed, f"Repaired: the form says {typed.strip()!r}."
    # Only digits that cannot be a date as typed — month 20, year 205 — are
    # read against the name, one slipped digit at most.
    readings = set()
    if ((_one_slip(typed_month.zfill(2), month) or _one_slip(typed_month, month))
            and _one_slip(typed_year, year)):
        readings.add(_a_date(int(day), int(month), int(year)))
    if (len(groups) == 3 and int(day) == int(month) and len(typed_month) <= 2
            and typed_year == year):    # "12/19/2022": month first
        readings.add(_a_date(int(typed_month), int(month), int(year)))
    readings.discard(None)
    if len(readings) != 1:
        return None
    return readings.pop(), f"Repaired: the form says {typed.strip()!r}."


def _repair_date_field(record: Record, filepath: str) -> None:
    typed = record.get("Date", "")
    if not typed or _looks_like_date(typed):
        return
    repaired = repair_date(typed, filepath)
    if repaired:
        record["Date"], reason = repaired
        record.setdefault("_flags", {})["Date"] = {"show": None, "reason": reason}


def _uniform_offset(configured: dict[str, str],
                    located: dict[str, str]) -> int | None:
    """How far the form shifted, when every located field moved together.

    An inserted row moves *everything* below it, so a form whose Model, serial,
    manufacturer and location all sit one row lower has moved its Date and
    Status too. Those two carry no label the locator can match — Status is
    captioned above its value rather than beside it — so they can only be
    reached by applying the offset the other four agree on.

    Returns None unless the agreement is exact: same column, same distance, for
    every field found. A partial or ragged match means something other than a
    shift, and guessing there is how a caption ends up in the register.
    """
    offsets = set()
    for field, ref in located.items():
        old = configured.get(field)
        if not old:
            return None
        old_column, old_row = coordinate_to_tuple(old)[1], coordinate_to_tuple(old)[0]
        new_column, new_row = coordinate_to_tuple(ref)[1], coordinate_to_tuple(ref)[0]
        if old_column != new_column:
            return None
        offsets.add(new_row - old_row)
    if len(offsets) == 1:
        moved = offsets.pop()
        return moved or None
    return None


def _shift(ref: str, rows: int) -> str:
    row, column = coordinate_to_tuple(ref)
    return f"{get_column_letter(column)}{row + rows}"


def _read_with(source, cell_map: dict[str, str]) -> Record:
    values = source.values({ref for ref in cell_map.values() if ref})
    return {field: clean(values.get(ref)) if ref else ""
            for field, ref in cell_map.items()}


# ──────────────────────────────────────────────────────────────────────────────
# Does the form's own caption agree with the layout?
#
# `plausible()` asks whether the values *look* right, and a shifted layout can
# pass that: on an older Baby Warmer form the map reads the model into
# Manufacturer and the ward into Model, and both look like perfectly ordinary
# text. What a shifted layout cannot fake is the caption printed beside each
# box. So before a layout is trusted, the caption to the left of its Model,
# Serial and Manufacturer cells is read — and a caption that names a DIFFERENT
# field rejects the layout.
#
# Only a contradiction counts. A missing caption, or one this does not
# recognise, proves nothing either way and blocks nothing — which is what keeps
# forms without printed captions reading exactly as before.
# ──────────────────────────────────────────────────────────────────────────────

_CAPTIONS: dict[str, re.Pattern] = {
    "Model": re.compile(r"^\s*model\s*(no\.?|number|name)?\s*[:.]?\s*$", re.I),
    "S.N": re.compile(
        r"^\s*(serial\s*(no\.?|number|#)?|s\s*[./]?\s*n\s*\.?\s*(o\.?)?)\s*[:.]?\s*$", re.I),
    "Manufacturer": re.compile(
        r"^\s*(manufacturer|manufacture|make|brand)\s*[:.]?\s*$", re.I),
    "Location": re.compile(r"^\s*(location|department|dept\.?|ward)\s*[:.]?\s*$", re.I),
    "Date": re.compile(
        r"^\s*(date(\s*of\s*receipt)?|calib\.?\s*date|test\s*date)\s*[:.]?\s*$", re.I),
    "Status": re.compile(r"^\s*status\s*[:.]?\s*$", re.I),
}

#: Other captions printed beside boxes on these forms. A value whose caption is
#: one of these is not the field that was wanted either.
_FOREIGN_CAPTION = re.compile(
    r"^\s*(status|next\s*calib.*|prev\.?\s*calib.*|last\s*cal.*|device\s*name"
    r"|equipment\s*type|type|class|safety|tested\s*by|entered\s*by|revised\s*by"
    r"|accessories|comments?|remarks?)\s*[:.]?\s*$", re.I)

#: How far left of a value its caption may sit. The standard form prints the
#: Serial caption four to five columns left of its box (G..J beside K).
_CAPTION_REACH = 6

#: The fields whose captions decide whether a layout is trusted.
_CHECKED_CAPTIONS = ("Model", "S.N", "Manufacturer")


def _caption_refs(ref: str) -> list[str]:
    """The cells left of a value on its own row, nearest first."""
    row, column = coordinate_to_tuple(ref)
    return [f"{get_column_letter(c)}{row}"
            for c in range(column - 1, max(0, column - 1 - _CAPTION_REACH), -1)]


def caption_verdict(field: str, value: str, row_text: list[str]) -> str:
    """'confirms', 'contradicts' or 'unknown' — one field's caption.

    ``row_text`` is the text of the cells left of the value, nearest first.
    The value's own merged span repeats its text leftwards and is stepped over.
    """
    own = " ".join((value or "").lower().split())
    for text in row_text:
        if not text or " ".join(text.lower().split()) == own:
            continue
        if _CAPTIONS[field].match(text):
            return "confirms"
        if _FOREIGN_CAPTION.match(text) or any(
                pattern.match(text) for name, pattern in _CAPTIONS.items()
                if name != field):
            return "contradicts"
        return "unknown"                # a value, or a caption we don't know
    return "unknown"


def layout_verdict(source, cells: dict[str, str], record: Record) -> str:
    """'contradicted' when any checked caption names another field;
    'confirmed' when at least one agrees and none contradicts; else 'unknown'.
    """
    wanted = {field: cells[field] for field in _CHECKED_CAPTIONS if cells.get(field)}
    if not wanted:
        return "unknown"
    lefts = {field: _caption_refs(ref) for field, ref in wanted.items()}
    values = source.values({r for refs in lefts.values() for r in refs})
    verdicts = [caption_verdict(field, record.get(field, ""),
                                [clean(values.get(r)) for r in refs])
                for field, refs in lefts.items()]
    if "contradicts" in verdicts:
        return "contradicted"
    return "confirmed" if "confirms" in verdicts else "unknown"


#: How far above and below the mapped Status cell to look for the real box.
#:
#: **This number is the whole cost of the fallback**, because `values()` scans
#: the sheet once per reference it is asked for — so the price is linear in the
#: band, and reading the whole A1:N95 grid to find one cell is what a naive
#: version does. Measured over 1,200 real forms, warm:
#:
#:     whole grid   50.9 s      band 24   14.2 s
#:     band 32      15.6 s      band 18   13.1 s   <- here
#:     baseline      9.0 s      band 12   12.1 s
#:
#: and what each band recovers, over 3,004 forms:
#:
#:     whole grid   2548        band 18   2499     band 12   2480
#:
#: 18 is the knee: 98% of everything the whole grid finds, for a quarter of its
#: cost. The median form is unaffected either way (2.5 ms -> 4.3 ms) because
#: only the 19.5% whose mapped cell missed ever pay, and 76% of those come back
#: with a real status.
_STATUS_BAND_ROWS = 18


def _status_band(anchor: str) -> set[str]:
    """The cells worth reading to find a Status box that has moved.

    Every column, because a moved box is usually in a different one, but only
    the rows near where the map said it would be.
    """
    match = re.match(r"([A-Z]+)(\d+)$", anchor)
    if not match:
        return set()
    row = int(match.group(2))
    lo = max(1, row - _STATUS_BAND_ROWS)
    hi = min(_LABEL_MAX_ROW, row + _STATUS_BAND_ROWS)
    # _STATUS_COLS, not _LABEL_COLUMNS: status_from_grid indexes into A-L and
    # ignores anything beyond, so reading M and N would cost and never be read.
    return {f"{c}{r}" for r in range(lo, hi + 1) for c in _STATUS_COLS}


def _settle_status(source, cells: dict[str, str], record: Record) -> None:
    """Refuse a caption in the Status column, and find a box that has moved.

    Modifies ``record`` in place. Three outcomes, in order:

    1. The value already **reads as a status** — kept untouched. This is the
       common path and it costs one set lookup. It is also why a device whose
       box never moves is completely unaffected: 10 of the 81 mapped Status
       cells read on every form sampled, and for those this function stops here
       every time.
    2. Otherwise the box is looked up by the form's **own printed label**. A
       moved box is usually in a different column, not merely a different row —
       33 device maps have forms like that — so an offset search cannot find it
       and only the label can. See [status_from_grid].
    3. Nothing found: the original text is kept, **unless it is a caption or an
       answer to a different question**, which is dropped. "Safety:" and
       "Syringe brand:" were reaching the register as a device's verdict, and a
       map one column out reports that question's answer, "Large". A blank is
       honest and neither of those was.

    Only ``Status`` is searched, and only on a form with a single box. Where a
    device has a ``Status2`` the two verdicts belong to different modules and
    one printed "Status" label cannot say which is which — so those forms get
    the caption guard and nothing else, rather than a confident wrong answer.
    """
    single_box = not cells.get("Status2")

    for field in ("Status", "Status2"):
        raw = record.get(field, "")
        if not raw or normalise_status(raw):
            continue                        # blank stays blank; a status stays

        engineer = _ENGINEER_CODE.match(raw)
        if (engineer or _is_a_label(raw) or _A_READING.match(raw)
                or raw.strip().lower().rstrip(".") in _NOT_A_VERDICT):
            record[field] = ""              # never EXPORT a caption, or a size
            # …but the register keeps what the form said, marked for checking,
            # by the owner's decision: nothing is dropped from it silently.
            record.setdefault("_flags", {})[field] = {
                "show": raw,
                "reason": "Not a verdict — an engineer's code, from the "
                          "'Tested by' / 'Revised by' box beside the status."
                          if engineer else
                          "Not a verdict — the status box holds a caption or "
                          "the answer to a different question."}

    anchor = cells.get("Status", "")
    if not anchor or not single_box or normalise_status(record.get("Status", "")):
        # No box on this form at all (CT, MRI, Dexa), two boxes, or one that
        # already read. Nothing to look for, and no read to pay for.
        return

    found = status_from_grid(source.values(_status_band(anchor)))
    if found:
        record["Status"] = clean(found)
        # A real verdict was found after all: nothing left to mark.
        record.get("_flags", {}).pop("Status", None)


def read_best(filepath: str, config: dict,
              extra: dict[str, str] | None = None,
              forms_only: bool = False) -> tuple[Record, str]:
    """Read a form, falling back through every layout we know before giving up.

    Returns the record and how it was obtained: ``"primary"``, ``"alt N"``,
    ``"labels"``, ``"labels on <sheet>"``, or ``"none"``.

    **The configured map always wins when it produces a plausible record**, so
    this can only rescue a file the map got wrong; it can never change one it
    already got right. Every later step costs an extra read, and 94.7% of forms
    stop at the first.

    The Status column is settled separately afterwards by [_settle_status],
    because a map can place the identity block perfectly and still miss the
    verdict box — `plausible()` asks for a serial and a model, so a form whose
    Status alone has moved never reaches any of the fallbacks below.

    ``extra`` is merged into every read, for a caller that wants more of the
    sheet in the same pass — firebase_export takes the client header that way
    rather than opening all 47,000 workbooks twice.

    ``forms_only`` refuses a form with nothing filled in, raising LeftOut —
    the "real device forms only" switch. Without it a blank form named for a
    device is read as it always was, a row of blanks.
    """
    cells = {**config["cells"], **(extra or {})}
    source = _open_workbook(filepath)
    try:
        if _is_device_list(source, filepath):
            raise NotAForm("a device list (register), not an inspection form")

        if config.get("source") == "word":
            # The workbook itself cannot be trusted for identity at all: the
            # Mammography QC template prints "Model" beside the X-ray TUBE,
            # which a label search would record as the device.
            record = {field: "" for field in cells}
            word = read_word_certificate(filepath)
            how = "none"
            if word:
                record.update({f: word.get(f, "") for f in cells if f in word})
                how = "word certificate"
        else:
            opening = getattr(source, "sheet_name", None)
            record, how = _best_layout(source, cells, config, extra)
            if how == "primary" and config.get("field_alternates"):
                _pick_field_cells(source, cells, config["field_alternates"], record)
            if (how.startswith("labels on") and opening is not None
                    and _CERTIFICATE_TAB.search(how)):
                # Identity came from the certificate, which carries no verdict:
                # the verdict box is on the form's own tab.
                source.select_sheet(opening)
            _settle_status(source, cells, record)
            if how == "none":
                # No layout fits and no tab carries the fields: the last place
                # left is the Word certificate some devices keep beside the
                # workbook. Identity only — no certificate carries a verdict,
                # so Status stays whatever the workbook said.
                word = read_word_certificate(filepath)
                if word and plausible(word):
                    for field in ("Manufacturer", "Model", "S.N", "Location", "Date"):
                        if field in cells:
                            record[field] = word.get(field, "")
                    how = "word certificate"

        if config.get("conclusion_tab"):
            _status_from_conclusion(source, config["conclusion_tab"], record)
        if not has_device_name(filepath):
            _refuse_if_not_this_device(source, config, record, filepath)
        elif forms_only and not _holds_identity(record) and not _workbook_holds_identity(
                source, config["cells"]):
            raise LeftOut("nothing is filled in — a blank form or template", record)
        if cells.get("Date"):
            _repair_date_field(record, filepath)
        record["_flags"] = {**_flags_for(record, how, cells),
                            **record.get("_flags", {})}
        return record, how
    finally:
        source.close()


def _pick_field_cells(source, cells: dict[str, str],
                      alternates: dict[str, list[str]], record: Record) -> None:
    """For one field, choose between cells by what the form's caption says.

    For a field that moved on SOME forms but not others — Lab Oven's Location
    is K20 on 68 forms and K19 on two — no single cell is right everywhere,
    and swapping the map would trade 68 fixes for two new errors. The form
    says which is which: take the first candidate whose caption confirms the
    field and which holds something; failing that, keep the mapped cell.
    """
    for field, others in alternates.items():
        if field not in cells or not cells[field]:
            continue
        candidates = [cells[field], *others]
        lefts = {ref: _caption_refs(ref) for ref in candidates}
        values = source.values(set(candidates) | {r for refs in lefts.values() for r in refs})
        verdicts = {ref: caption_verdict(field, clean(values.get(ref)),
                                         [clean(values.get(r)) for r in lefts[ref]])
                    for ref in candidates}
        confirmed = [ref for ref in candidates
                     if clean(values.get(ref)) and verdicts[ref] == "confirms"]
        chosen = confirmed[0] if confirmed else None
        if field == "Date" and not (chosen and _looks_like_date(clean(values.get(chosen)))):
            # An engineer sometimes types a range into one date box
            # ("14-24/09/2023"). A real date in any candidate the form does not
            # caption as something else beats it.
            chosen = next((ref for ref in candidates
                           if verdicts[ref] != "contradicts"
                           and _looks_like_date(clean(values.get(ref)))), chosen)
        if chosen is None and verdicts[cells[field]] == "contradicts":
            # The mapped cell is certainly another box — Centrifuge's K25 is
            # "Tested by", so its "JTE-40" is an engineer, not a verdict. The
            # box the form captions as this field wins even when it is empty:
            # an honest blank beats a confident wrong value.
            chosen = next((ref for ref in candidates[1:]
                           if verdicts[ref] == "confirms"), None)
        if chosen is not None:
            record[field] = clean(values.get(chosen))


def _status_from_conclusion(source, tab_pattern: str, record: Record) -> None:
    """Take the verdict from a Conclusion sentence, where the form has one.

    The sentence wins over the mapped status box: on a Hemodialysis "Final"
    printout the box the map lands on is one test line's result, while the
    Conclusion is the verdict on the device.
    """
    sentence = conclusion_on(source, tab_pattern)
    verdict = verdict_of_sentence(sentence or "")
    if verdict is None:
        return
    record["Status"] = verdict
    if verdict == "":
        record.setdefault("_flags", {})["Status"] = {
            "show": sentence, "reason": "The conclusion gives readings only — no verdict."}


#: Reasons a value is kept in the register but marked for checking. Firebase
#: never sees these: it reads the named fields only, never `_flags`.
_DATE_TIME_RE = re.compile(r"^\s*\d{4}-\d{2}-\d{2}(\s+\d{2}:\d{2}:\d{2})?\s*$")
_EXCEL_DATE_NUMBER = re.compile(r"^\s*[2-6]\d{4}(\.0+)?\s*$")


def _looks_like_date(text: str) -> bool:
    """A date as these forms write it — including a stray leading apostrophe
    ("'10-02-2024", a date an engineer forced to text) and the bare serial
    number xlrd returns for a real date cell in an .xls."""
    text = (text or "").strip().lstrip("'")
    return bool(is_calendar_date(text) or _DATE_TIME_RE.match(text)
                or _EXCEL_DATE_NUMBER.match(text))


def _flags_for(record: Record, how: str, cells: dict[str, str]) -> dict:
    flags: dict[str, dict] = {}

    def flag(field: str, reason: str) -> None:
        if cells.get(field) and field not in flags:
            flags[field] = {"show": None, "reason": reason}

    identity = ("Manufacturer", "Model", "S.N", "Location", "Date")
    if how == "none":
        for field in identity:
            if clean(record.get(field)):
                flag(field, "No known layout fits this form — this is whatever "
                            "sits in the mapped cell.")
    elif how.startswith("labels"):
        for field in identity:
            if clean(record.get(field)):
                flag(field, "Read by finding the form's printed labels; this "
                            "layout is not in the device table yet.")

    serial = clean(record.get("S.N"))
    if serial and classify_serial(serial) == "suspect":
        flag("S.N", "Does not look like a serial number (a date, or no digits).")
    for field in ("Model", "Manufacturer"):
        value = clean(record.get(field))
        if value and _is_a_label(value):
            flag(field, "A caption printed on the form, not a value.")
        elif value and is_calendar_date(value):
            flag(field, "A date, not a " + ("model" if field == "Model" else "maker") + ".")
    model, maker = clean(record.get("Model")), clean(record.get("Manufacturer"))
    if (model and maker and model.lower() not in _PLACEHOLDERS
            and model.lower() == maker.lower()):
        flag("Manufacturer", "Same as the Model — one of the two is probably misread.")
    date = clean(record.get("Date"))
    if date and date.lower() not in _PLACEHOLDERS and not _looks_like_date(date):
        flag("Date", "Not a date.")
    for field in ("Status", "Status2"):
        value = clean(record.get(field))
        # "0" is an empty verdict box copying an empty test sheet — not wrong.
        if value and value not in ("0", "0.0") and not normalise_status(value):
            flag(field, "Not a recognised verdict.")
    return flags


#: A certificate tab: "Certificate", "ECG.certificate", and on FC just "cert".
_CERTIFICATE_TAB = re.compile(r"cert", re.I)


class NotAForm(ValueError):
    """The workbook is not an inspection form at all — e.g. a device list."""


class LeftOut(NotAForm):
    """A form with nothing filled in, refused only when the caller asked for
    real device forms — see [read_best]'s ``forms_only``.

    Carries what was read all the same: the archive export numbers devices by
    serial across every file, left out or not, and must see what it always saw.
    """

    def __init__(self, reason: str, record: Record | None = None) -> None:
        super().__init__(reason)
        self.record = record or {}


class ReadAsOther(NotAForm):
    """The name carries no device code, and the form's own Equipment Type
    names exactly one other device. The caller reads it again as that one."""

    def __init__(self, code: str, stated: str) -> None:
        super().__init__(f"its form says it is a {stated!r}")
        self.code = code
        self.stated = stated


# ──────────────────────────────────────────────────────────────────────────────
# The Word certificate beside a workbook
#
# Some devices keep their identity only in a same-named Word certificate: the
# workbook is a vendor's QC template (Mammography) or an electrical-safety
# printout (Hemodialysis "Final"), and its own cover page, where it has one,
# was never filled in. The certificate prints the calibrator's block FIRST —
# "Dose meter Model: 07-492 … S.N: 108357" — and the device's after an
# "Equipment Data" heading, so only text after that heading is read.
#
# Standard library only: a .docx is a zip of XML, and the body text of a
# Word 97 .doc is stored as plain 8-bit or UTF-16 runs, which is enough for
# fixed labels like these. No verdict is ever read from one — none carries it.
# ──────────────────────────────────────────────────────────────────────────────

_WORD_LABELS = (
    ("Model", r"Model"),
    ("S.N", r"Serial\s*No\.?"),
    ("Manufacturer", r"Manufacturer"),
    ("Location", r"Location"),
)
_WORD_STOP = (r"(?=\s*(?:Equipment\s*Type|Model|Serial\s*No\.?|Manufacturer|Location"
              r"|Prev\.?\s*Calib\.?|Date\s*of\s*receipt|Next\s*Calib\.?\s*On"
              r"|Test\s*Location|Temperature|Relative\s*Humidity)\s*:"
              r"|\s*Environmental\s*conditions|$)")
_WORD_DATE = re.compile(r"Test\s*Date\s*:\s*(\d{1,2}\s*[-/.]\s*\d{1,2}\s*[-/.]\s*\d{2,4})", re.I)
_WORD_RECEIPT = re.compile(
    r"Date\s*of\s*receipt\s*:\s*(\d{1,2}\s*[-/.]\s*\d{1,2}\s*[-/.]\s*\d{2,4})", re.I)


def word_certificate_path(filepath: str) -> str | None:
    stem = os.path.splitext(filepath)[0]
    for extension in (".docx", ".doc", ".DOCX", ".DOC"):
        if os.path.isfile(stem + extension):
            return stem + extension
    return None


def _word_text(path: str) -> str:
    if path.lower().endswith(".docx"):
        with zipfile.ZipFile(path) as archive:
            xml = archive.read("word/document.xml")
        text = _unescape(re.sub(rb"<[^>]+>", b" ", xml))
        return " ".join(text.split())
    raw = Path(path).read_bytes()
    best = ""
    for encoding in ("cp1252", "utf-16-le"):
        text = " ".join(re.sub(r"[^\x20-\x7E]+", " ",
                               raw.decode(encoding, "ignore")).split())
        if text.count("Equipment Data") > best.count("Equipment Data"):
            best = text
    return best


def read_word_certificate(filepath: str) -> Record | None:
    """The device block of the Word certificate beside ``filepath``, or None."""
    path = word_certificate_path(filepath)
    if not path:
        return None
    try:
        text = _word_text(path)
    except Exception:                                   # noqa: BLE001
        log.debug("Unreadable Word certificate %s", path, exc_info=True)
        return None
    start = text.find("Equipment Data")
    if start < 0:
        return None
    block = text[start:start + 900]
    record: Record = {}
    for field, label in _WORD_LABELS:
        match = re.search(rf"\b{label}\s*:\s*(.*?){_WORD_STOP}", block)
        if match:
            record[field] = clean(match.group(1))[:80]
    date = _WORD_DATE.search(text) or _WORD_RECEIPT.search(block)
    if date:
        record["Date"] = re.sub(r"\s+", "", date.group(1))
    return record if record.get("S.N") or record.get("Model") else None


# ──────────────────────────────────────────────────────────────────────────────
# The verdict written as a sentence
#
# Hemodialysis workbooks state their verdict in a Conclusion on the "Final"
# tab, not in a status box, and its row moves (A119 to A132 across the
# archive). Three wordings cover every one of them.
# ──────────────────────────────────────────────────────────────────────────────

_CONCLUSION_LABEL = re.compile(r"^\s*conclusions?\s*:?", re.I)
_PASSED_SENTENCE = re.compile(r"\bpassed\s+all\s+tests?\b", re.I)
_NO_VERDICT_SENTENCE = re.compile(r"device\s+readings|no\s+accepted\s+range", re.I)
_FAILED_SENTENCE = re.compile(r"\bfail(ed|s)?\b", re.I)


def verdict_of_sentence(sentence: str) -> str | None:
    """A conclusion sentence as a status word, "" for "no verdict", None if
    the sentence says nothing this recognises."""
    if not sentence:
        return None
    if normalise_status(sentence):                  # "…Limited Calibrated, non…"
        return {STATUS_LIMITED_NON: "Limited non", STATUS_LIMITED: "Limited",
                STATUS_LIMITED_FAIL: "Limited fail", STATUS_PASS: "Pass",
                STATUS_FAIL: "Fail", STATUS_CALIBRATED: "Calibrated"}[
                    normalise_status(sentence)]
    if _NO_VERDICT_SENTENCE.search(sentence):
        return ""
    if _PASSED_SENTENCE.search(sentence):
        return "Pass"
    if _FAILED_SENTENCE.search(sentence):
        return "Fail"
    return None


def conclusion_on(source, tab_pattern: str) -> str | None:
    """The Conclusion sentence on the first tab matching ``tab_pattern``."""
    tab = next((t for t in getattr(source, "sheet_names", [])
                if re.search(tab_pattern, t, re.I)), None)
    if tab is None or not source.select_sheet(tab):
        return None
    grid = {f"{c}{r}" for r in range(100, 181) for c in "ABCDEFGH"}
    values = {ref: clean(v) for ref, v in source.values(grid).items()}
    labels = sorted((int(ref[1:]), ref) for ref, text in values.items()
                    if text and _CONCLUSION_LABEL.match(text) and ref[0] in "AB")
    if not labels:
        return None
    row, ref = labels[0]
    # The sentence may share the label's own cell: "Conclusion: Tested …".
    rest = _CONCLUSION_LABEL.sub("", values[ref], count=1).strip()
    if len(rest) > 2:
        return rest
    for r in range(row, row + 5):
        for column in "ABCDEFGH":
            text = values.get(f"{column}{r}", "")
            if len(text) > 2 and not _CONCLUSION_LABEL.match(text):
                return text
    return None


#: The column headings of a register: Calist's own output, and the hand-made
#: device lists that copy it (which write "Device name" and "SN").
_REGISTER_HEADINGS = {"device", "manufacturer", "model", "s.n", "location", "code"}
_HEADING_SPELLINGS = {"device name": "device", "sn": "s.n", "s/n": "s.n",
                      "serial": "s.n", "serial no": "s.n"}

#: A device's number written the house way — `G302-AGH001` — anywhere in the
#: name, so "Copy of D41-AA002-0624" still counts as named for its device.
_DEVICE_NAME_RE = re.compile(r"[A-Za-z]+\d+-[A-Za-z]+\d+")


def has_device_name(filepath: str) -> bool:
    """Whether the filename carries a site and device number, `G302-AGH001`.

    One without it was named some other way — by serial (`CN84017253.xlsx`),
    by brand (`GE  Tec 850  SQAB01358  OR.xlsx`), or left as a template
    (`BB.xlsx`, `000-AGH000-0000.xlsx`) — and its "code" is only its first
    letters: a Philips serial reads as a Microwave, a GE vaporizer as a
    Temperature Calibration Tester.
    """
    return bool(_DEVICE_NAME_RE.search(Path(filepath).stem))


def _is_device_list(source, filepath: str) -> bool:
    """Whether this workbook is a register rather than a form.

    Device lists sit among the forms and are named like them —
    "Device List H23-AS-1023.xlsx" resolves to a Centrifuge, "Al Arbaeen
    Device List.xlsx" to a Phototherapy (from "Al") — so the name cannot tell
    them apart, and reading one as a device writes a row of whichever cells the
    map lands on. The heading row can. It is only looked for when the sheet or
    the file is called a list, or the file is not named for a device ("Al
    baeerat.xlsx" is a site's list), so an ordinary form pays nothing.
    """
    name = (getattr(source, "sheet_name", "") or "").strip().lower()
    if (name != "list" and "list" not in Path(filepath).stem.lower()
            and has_device_name(filepath)):
        return False
    values = source.values({f"{c}{r}" for r in range(1, 7) for c in "ABCDEFGHIJ"})
    for row in range(1, 7):
        headings = set()
        for c in "ABCDEFGHIJ":
            text = " ".join(clean(values.get(f"{c}{row}")).lower().rstrip(".").split())
            headings.add(_HEADING_SPELLINGS.get(text, text))
        if len(_REGISTER_HEADINGS & headings) >= 5:
            return True
    return False


# ──────────────────────────────────────────────────────────────────────────────
# Files not named for a device
#
# By the owner's decision these are skipped, like device lists, when the form
# either holds nothing at all or says outright that it is another device. Both
# checks run only for a file whose name carries no `G302-AGH001` — about one
# form in a hundred — so a properly named form pays nothing and is never
# skipped, even when it is blank.
# ──────────────────────────────────────────────────────────────────────────────

#: The caption before the form's own statement of what the device is:
#: "Equipment Type:" in the certificate, "Device Type:" on some data tabs.
_TYPE_CAPTION = re.compile(r"^\s*(equipment|device)\s*type\s*:?\s*$", re.I)

#: Spellings that name the same thing, so "Pulse Oximeter Module" is the
#: "Vital Sign (SPO2 Module)" it is, and "Lab Pipette" the "Pipet".
_TYPE_SYNONYMS = ((re.compile(r"pulse\s*oximeter"), "spo2"),
                  (re.compile(r"pipette"), "pipet"),
                  (re.compile(r"x\s*-?\s*ray"), "xray"),
                  (re.compile(r"c\s*-\s*arm"), "carm"))
_TYPE_FILLER = {"device", "machine", "unit", "system", "the", "and", "of", "for",
                "module", "set", "analyser", "analyzer"}


def _type_words(text: str) -> set[str]:
    text = (text or "").lower()
    for pattern, word in _TYPE_SYNONYMS:
        text = pattern.sub(word, text)
    return {w for w in re.findall(r"[a-z0-9]+", text) if w not in _TYPE_FILLER}


def names_device(device_name: str, stated: str) -> bool:
    """Whether a typed device type names this device: one's words contain the
    other's. "Syringe Pump" is a "Syringe"; "Anesthesia (Ventilator)" is an
    "Anesthesia" and a "Ventilator", and not a "Temperature Calibration Tester".
    """
    mine, theirs = _type_words(device_name), _type_words(stated)
    return bool(mine and theirs) and (mine <= theirs or theirs <= mine)


def stated_device_type(source) -> str:
    """The device type the form states beside its own "Equipment Type:" or
    "Device Type:" caption, on any tab; "" when it states none.

    Leaves the source on the sheet it started on.

    Certificates first, where "Equipment Type:" is printed, and the caption
    looked for in A–F only (it sits in B or C) before its row is read: every
    reference costs a scan of the sheet, and reading the whole A1:N60 grid of
    each tab doubled the time of the forms that come here.
    """
    captions = {f"{c}{r}" for r in range(1, 61) for c in "ABCDEF"}
    opening = getattr(source, "sheet_name", None)
    tabs = sorted(getattr(source, "sheet_names", []),
                  key=lambda tab: not _CERTIFICATE_TAB.search(tab))
    try:
        for name in tabs:
            if not source.select_sheet(name):
                continue
            found = source.values(captions)
            for ref in sorted((r for r, v in found.items() if _TYPE_CAPTION.match(clean(v))),
                              key=coordinate_to_tuple):
                row, column = coordinate_to_tuple(ref)
                right = [f"{get_column_letter(c)}{row}"
                         for c in range(column + 1, len(_LABEL_COLUMNS) + 1)]
                values = source.values(right)
                for cell in right:
                    text = clean(values.get(cell))
                    if text and not _TYPE_CAPTION.match(text):
                        if classify_serial(text) != "placeholder":
                            return text
                        break
    finally:
        if opening is not None:
            source.select_sheet(opening)
    return ""


def _is_identity_text(text: str) -> bool:
    return (bool(re.search(r"[^\W\d_]", text or ""))
            and classify_serial(text) != "placeholder" and not _is_a_label(text))


def _holds_identity(record: Record) -> bool:
    """A real serial, or a model or maker written in words — "-113" is not."""
    return (classify_serial(record.get("S.N", "")) in ("real", "assigned")
            or any(_is_identity_text(record.get(field, ""))
                   for field in ("Model", "Manufacturer")))


def _identity_parts(record: Record) -> int:
    return ((classify_serial(record.get("S.N", "")) in ("real", "assigned"))
            + sum(_is_identity_text(record.get(field, ""))
                  for field in ("Model", "Manufacturer")))


def _workbook_holds_identity(source, cells: dict[str, str]) -> bool:
    """Whether any tab's printed labels have a device identity beside them —
    the check that a form which read nothing is really blank, not just laid
    out some way no map knows (a certificate with a serial and a maker but a
    model of "0").

    Two of serial, model and maker, not one: the `safety` tab of a blank
    template prints the safety analyzer's own serial beside "S.N" — a
    calibrator, not the device.
    """
    grid = {f"{c}{r}": f"{c}{r}" for r in range(1, _LABEL_MAX_ROW + 1)
            for c in _LABEL_COLUMNS}
    opening = getattr(source, "sheet_name", None)
    try:
        for name in getattr(source, "sheet_names", []) or [opening]:
            if name is not None and not source.select_sheet(name):
                continue
            located = locate_by_labels(source.values(set(grid)))
            if located and _identity_parts(
                    _read_with(source, {f: located.get(f, "") for f in cells})) >= 2:
                return True
    finally:
        if opening is not None:
            source.select_sheet(opening)
    return False


#: A folder an anesthesia machine is filed under. See [_stale_certificate].
_ANESTHESIA_FOLDER = re.compile(r"an(a)?esthesia", re.I)


def _stale_certificate(code: str, filepath: str) -> bool:
    """Whether a certificate naming this device is the template's, not the
    device's: "Ventilator" on a file filed under an Anesthesia folder.

    680 Anesthesia forms carry "Ventilator" as their Equipment Type, and the
    21 GE Carestations named by brand ("GE  OR  SM617300058MA") are anesthesia
    machines, not ventilators. By the owner's decision these are skipped, as
    they always were, rather than read as one. A Carescape filed under
    "Ventilator" is one, and is still read as it.
    """
    return code == "AM" and any(_ANESTHESIA_FOLDER.search(part)
                                for part in Path(filepath).parent.parts)


def _refuse_if_not_this_device(source, config: dict, record: Record,
                               filepath: str = "") -> None:
    """Deal with a file not named for a device whose form is someone else's or
    no one's. Returns quietly when the form is this device's.

    By the owner's decision the form's own Equipment Type decides the device
    for these files: when it names exactly one device the map knows, this
    raises ReadAsOther and the caller reads the file again as that device
    ("GE  TEC850  OR  SQAX00871" is a Vaporizer, not the Temperature Calibration
    Tester its "GE" reads as). When it could be several — "Patient Monitor" is
    AG, AGH and VAGH — nothing is guessed and the file is refused, naming them.
    """
    stated = stated_device_type(source)
    own = config.get("device_name", "")
    if stated and not names_device(own, stated):
        named = sorted(code for code, other in DEVICE_CONFIGS.items()
                       if names_device(other.get("device_name", ""), stated))
        if len(named) == 1 and not _stale_certificate(named[0], filepath):
            raise ReadAsOther(named[0], stated)
        if named:
            choices = (f" — which could be any of {', '.join(named)}"
                       if len(named) > 1 else "")
            raise NotAForm(f"not named for a device, and its form says it is a "
                           f"{stated!r}, not a {own!r}{choices}")
    if not _holds_identity(record) and not _workbook_holds_identity(
            source, config["cells"]):
        raise NotAForm("not named for a device, and nothing is filled in "
                       "(a blank form or template)")


def _best_layout(source, cells: dict[str, str], config: dict,
                 extra: dict[str, str] | None) -> tuple[Record, str]:
    """Try every written-down layout, then the form's own labels. See read_best.

    A layout is accepted only when its values are plausible AND the form's own
    captions do not contradict it — see [layout_verdict]. The second test is
    what catches a form that has moved a couple of rows and still reads
    plausible-looking text from the wrong boxes.
    """
    record = _read_with(source, cells)
    if plausible(record) and layout_verdict(source, cells, record) != "contradicted":
        return record, "primary"

    for index, alternate in enumerate(config.get("alt_cells", []), start=1):
        layout = {**alternate, **(extra or {})}
        candidate = _read_with(source, layout)
        if (plausible(candidate)
                and layout_verdict(source, layout, candidate) != "contradicted"):
            return candidate, f"alt {index}"

    # Nothing written down fits. Ask the form where its fields are.
    grid = {f"{c}{r}": f"{c}{r}"
            for r in range(1, _LABEL_MAX_ROW + 1) for c in _LABEL_COLUMNS}
    sheet = source.values(set(grid))
    located = locate_by_labels(sheet)
    if located:
        candidate = {**cells, **located}
        # Carry Date and Status along when the form has simply shifted:
        # they have no label to find, and leaving them behind is how the
        # caption "Status:" ends up in the register instead of "Calibrated".
        rows = _uniform_offset(cells, located)
        if rows is not None:
            for field in ("Date", "Status", "Status2", "S.N2"):
                if cells.get(field):
                    candidate[field] = _shift(cells[field], rows)
        # A form whose block moved to other COLUMNS too has no uniform offset,
        # so a field with no label of its own is placed from a located one:
        # Therapeutic Ultrasound prints its Date two rows above its Model,
        # wherever that block sits — without this its date read the caption
        # "Test Parameter" on 95 forms.
        for field, (anchor, down) in config.get("label_offsets", {}).items():
            if located.get(anchor):
                candidate[field] = _shift(located[anchor], down)
        found = _read_with(source, candidate)
        if plausible(found):
            # Neither placement reached a date: ask the form's date caption.
            # Only a non-date is ever replaced, so a right date cannot change.
            if cells.get("Date") and not _looks_like_date(found.get("Date", "")):
                found["Date"] = date_by_caption(sheet) or found.get("Date", "")
            return found, "labels"

    # Still nothing: the device details may be on another tab entirely.
    #
    # Remember which sheet we started on. Walking the other tabs leaves the
    # source pointing at the last one tried, and the caller reads the Status
    # box off whatever sheet it is handed — so a failed search must put the
    # source back, or the verdict gets read from a tab the record did not come
    # from.
    opening_sheet = getattr(source, "sheet_name", None)
    # The opening tab's captions confirm the map, and only the identity
    # failed — typically a serial nobody entered. The map's Status and Date
    # boxes are then the form's own, whichever tab supplies the identity.
    map_confirmed = layout_verdict(source, cells, record) == "confirmed"
    fallback: tuple[Record, str] | None = None
    # Ordinary tabs first, exactly as before; a certificate tab only last.
    # A certificate repeats the identity but never the verdict, so a rescue
    # from the form's own data tab keeps its Status — reaching the certificate
    # first lost the verdict on forms the data tab used to read in full.
    tabs = sorted(getattr(source, "sheet_names", []),
                  key=lambda tab: bool(_CERTIFICATE_TAB.search(tab)))
    for name in tabs:
        if name == opening_sheet:
            continue
        if not source.select_sheet(name):
            continue
        sheet = source.values(set(grid))
        located = locate_by_labels(sheet)
        if not located:
            continue
        # Only what was actually located on THIS sheet. The configured
        # references describe a different tab, so carrying them across
        # reads whatever happens to sit at those coordinates here — which
        # is how a Nebulizer's date came back as "Gas Flow Analyser".
        # A blank field is honest; a confident wrong one is not.
        found = _read_with(source, {field: located.get(field, "")
                                    for field in cells})
        if not plausible(found):
            continue
        # A certificate or cover page carries neither the verdict nor the
        # form's own date; the opening tab, which the map describes, does. So
        # keep what the map read there — and only a value that is
        # unmistakably a date or a verdict — for a rescue from a certificate,
        # or when the opening tab's captions confirm the map. Not otherwise:
        # then the opening tab may be a test sheet, and a Balance form's
        # mapped cell there reads a test line's "Fail".
        if map_confirmed or _CERTIFICATE_TAB.search(name):
            if not found.get("Date") and _looks_like_date(record.get("Date", "")):
                found["Date"] = record["Date"]
            for field in ("Status", "Status2"):
                if (cells.get(field) and not normalise_status(found.get(field))
                        and normalise_status(record.get(field))):
                    found[field] = record[field]
        # Still no date: this tab prints its own, beside "Date of receipt:"
        # (a data tab), "Test Date" (a cover page) or "Calib. Date:" (a
        # certificate). The label search places no Date, so every rescue from
        # another tab used to leave it blank — 190 forms across the archive.
        if cells.get("Date") and not found.get("Date"):
            found["Date"] = date_by_caption(sheet)
        # A certificate's typed placeholder serial is a worse answer than a
        # real serial on another tab: keep looking, and settle for it last.
        if classify_serial(found.get("S.N", "")) in ("real", "assigned"):
            return found, f"labels on {name!r}"
        fallback = fallback or (found, name)

    if fallback is not None:
        # The walk has moved on; the Status search that follows reads the
        # current tab, so go back to the one this record came from.
        found, name = fallback
        source.select_sheet(name)
        return found, f"labels on {name!r}"
    if opening_sheet is not None:
        source.select_sheet(opening_sheet)
    return record, "none"


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
    # The module row shares the parent's identity marks, but its Status is the
    # parent's Status2 — so it takes that mark, never the parent's own.
    flags = {f: v for f, v in record.get("_flags", {}).items()
             if f not in ("Status", "Status2")}
    if "Status2" in record.get("_flags", {}):
        flags["Status"] = record["_flags"]["Status2"]
    extra["_flags"] = flags

    old_token, new_token = second_row["code_replace"]
    # count=1: replace only the device code, never a site prefix that happens
    # to contain the same letters.
    extra["Code"] = re.sub(re.escape(old_token), new_token, extra["Code"], count=1)
    return extra


#: In quiet mode, how many files pass between progress lines. Low enough that
#: a stalled run still looks different from a fast one, high enough that forty
#: thousand forms produce a readable log rather than forty thousand lines.
HEARTBEAT_EVERY = 500


def _heartbeat(index: int, total: int, started: float) -> None:
    elapsed = time.monotonic() - started
    rate = index / elapsed if elapsed > 0 else 0
    log.info("        read %s of %s  ·  %.0f/s", f"{index:,}", f"{total:,}", rate)


@dataclass
class _FormRead:
    """One form as read, kept until every form is in so copies can be
    settled against their originals — see [_settle_copies]."""

    record: Record
    second: Record | None
    config: dict
    outcome: FileOutcome
    repair: NameRepair
    #: Why the Code cell is marked, sentence by sentence.
    notes: list[str]
    dropped: bool = False


def _serial_key(record: Record) -> str:
    """A serial as two copies of one form would both write it, or "" when
    there is none to compare — blank, or a placeholder like "N.A"."""
    serial = record.get("S.N", "").split("\n")[0].strip()
    if classify_serial(serial) in ("blank", "placeholder"):
        return ""
    return re.sub(r"\.0+$", "", serial).upper()


def _settle_copies(reads: list[_FormRead]) -> None:
    """Decide what happens to files whose repaired names collide.

    Cutting "(2)" or " - Copy" off a name can land it on a file already in the
    run: "D38-AGH090-0225 (2)" beside "D38-AGH090-0225". By the owner's
    decision only the original is kept — **when it is the same device**. A copy
    whose serial matches the original's, or that has none, is dropped and
    listed. One with a different serial is a different device that was filed
    under a copy's name ("F21-BZ010-0624 (2)" is pump 14007615, the original
    14007302) and is kept, marked.

    Where no file is the plain original, the extra text is what tells the files
    apart — "G341-AB00-0626-SEVO" and "-Iso" are two vaporizers — so none of
    them is cut.
    """
    by_name: dict[str, list[_FormRead]] = {}
    for read in reads:
        by_name.setdefault(read.record["Code"].upper(), []).append(read)

    for members in by_name.values():
        if len(members) < 2 or not any(m.repair.is_copy for m in members):
            continue
        originals = [m for m in members if not m.repair.is_copy]
        copies = [m for m in members if m.repair.is_copy]

        if not originals:
            for member in members:
                kept = replace(member.repair, name=member.repair.with_tail, tail="")
                member.repair = kept
                member.notes = [n for n in member.notes if not n.endswith(
                    "after the date left out")]
                _set_code(member, kept.name)
            continue

        known = {_serial_key(o.record) for o in originals} - {""}
        first = originals[0].outcome.filename
        for copy in copies:
            serial = _serial_key(copy.record)
            if not serial or serial in known:
                copy.dropped = True
                copy.outcome.status = COPY
                copy.outcome.detail = (
                    f"a copy of {first}" + (f", same serial {serial}" if serial
                                            else ", with no serial of its own")
                    + " — only the original is in the register")
                log.warning("%s — %s", copy.outcome.filename, copy.outcome.detail)
            else:
                theirs = ", ".join(sorted(known)) or "none"
                copy.notes.append(
                    f"Same name as {first} once the extra text is left out, but "
                    f"a different serial ({serial}, not {theirs}) — kept as a "
                    f"device of its own")


def _set_code(read: _FormRead, code: str) -> None:
    """Write a read's Code — and its module row's, which is derived from it."""
    read.record["Code"] = read.record["_group"] = code
    read.outcome.code = code
    if read.second is not None:
        second = build_second_row(read.record, read.config["second_row"])
        read.second["Code"] = second["Code"]
        read.second["_group"] = code


def _mark_code(read: _FormRead) -> None:
    """Colour the Code cell when the name was repaired in a way worth checking,
    and say what the file on disk is actually called."""
    if not read.notes:
        return
    text = "; ".join(read.notes)
    reason = f"{text[0].upper()}{text[1:]}. The file is {read.outcome.filename!r}."
    for record in (read.record, read.second):
        if record is not None:
            record.setdefault("_flags", {})["Code"] = {"show": None, "reason": reason}


def extract_records(
    source_files: Iterable[str],
    on_file: ProgressHook | None = None,
    cancel: threading.Event | None = None,
    real_forms_only: bool = False,
    quiet: bool = False,
) -> tuple[list[Record], list[FileOutcome]]:
    """Read every source file, returning the records and a per-file outcome.

    ``on_file(outcome, index, total)`` fires after each file. ``cancel`` is
    checked between files, so a long run can be stopped without waiting for it
    to finish; files not reached are reported as CANCELLED.
    ``real_forms_only`` leaves out anything that is not a device form — a name
    with no customer code, a 000 template, a device list, a blank form — as
    LEFT_OUT; see [classify_file] and [read_best].

    Every record's Code is its file's name with the typing slips undone — see
    [repair_name]. Copies of a form already in the run are settled at the end,
    once every original has been read; see [_settle_copies].

    ``quiet`` drops the line-per-file commentary in favour of a heartbeat.
    Problems are still reported in full — it is the forty thousand successes
    that make a log unreadable, never the failures.
    """
    ordered = sorted(source_files)
    total = len(ordered)
    reads: list[_FormRead] = []
    outcomes: list[FileOutcome] = []
    started = time.monotonic()

    def report(outcome: FileOutcome, index: int) -> None:
        outcomes.append(outcome)
        if on_file:
            on_file(outcome, index, total)

    for index, filepath in enumerate(ordered, start=1):
        if cancel is not None and cancel.is_set():
            for remaining in ordered[index - 1:]:
                outcomes.append(FileOutcome(os.path.basename(remaining), remaining,
                                            CANCELLED, detail="Run cancelled"))
            log.warning("Cancelled — %s not processed.",
                        plural(total - index + 1, "file"))
            break

        filename = os.path.basename(filepath)
        pre = classify_file(filepath, real_forms_only)

        if pre.status == LEFT_OUT:
            if not quiet:
                log.info("%s — left out: %s", filename, pre.detail)
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
            code = pre.device_code
            notes: list[str] = []
            try:
                record, layout = read_best(filepath, config, forms_only=real_forms_only)
            except ReadAsOther as other:
                # The name gives no device; the form's own Equipment Type does.
                code, config = other.code, DEVICE_CONFIGS[other.code]
                record, layout = read_best(filepath, config, forms_only=real_forms_only)
                notes.append(f"the name carries no device code, and the form's "
                             f"Equipment Type says {other.stated!r} — read as a "
                             f"{config['device_name']}")
                log.info("%s — read as a %s (%s): its form says %r",
                         filename, config["device_name"], code, other.stated)
            if layout not in ("primary", "none"):
                # Worth saying out loud: a form reaching here has been
                # re-laid-out since its map was written, and the register only
                # has a row for it because the fallback found the fields.
                log.info("%s — read using %s (the form has moved since its "
                         "cell map was written)", filename, layout)

            repair = settle_name_date(repair_name(Path(filename).stem),
                                      clean(record.get("Date")))
            notes += repair.notes
            record["Code"] = repair.name
            record["Device"] = config["device_name"]

            # Devices with a probe or tube carry a second serial; show both.
            if record.get("S.N2", "").strip():
                record["S.N"] = f"{record['S.N']}\n({record['S.N2']})"

            # Sort keys, so ordering never has to re-derive the code tokens.
            record["_group"] = record["Code"]
            record["_row_order"] = 0
            # Kept apart from Code because build_second_row rewrites that.
            record["_source"] = filename

            rows = 1
            if quiet:
                if index % HEARTBEAT_EVERY == 0:
                    _heartbeat(index, total, started)
            else:
                log.info(
                    "[OK]    %s  →  %s (%s)  |  Model: %s  |  S/N: %s  |  Status: %s",
                    filename, config["device_name"], code,
                    record.get("Model", ""), record.get("S.N", ""),
                    record.get("Status", ""),
                )
                if repair.name != Path(filename).stem:
                    log.info("        ↳ written as %s", repair.name)

            second = None
            if "second_row" in config:
                second = build_second_row(record, config["second_row"])
                rows = 2
                if not quiet:
                    log.info(
                        "        ↳ 2nd row added: %s  |  Code: %s  |  Status: %s",
                        second["Device"], second["Code"], second["Status"],
                    )

            # A dual-serial device carries its second number on its own line
            # for the spreadsheet; a table row wants it on one.
            serial = record.get("S.N", "").replace("\n", "  /  ")
            outcome = FileOutcome(filename, filepath, OK, code,
                                  config["device_name"], serial=serial,
                                  rows=rows, code=repair.name)
            reads.append(_FormRead(record, second, config, outcome, repair, notes))
            report(outcome, index)

        except NotAForm as exc:
            # Not a failure — the file is simply not a form. With the switch on
            # it is left out like any other non-form; otherwise it is reported
            # as unsupported, so it shows among the problems without looking
            # like something broke.
            if real_forms_only:
                if not quiet:
                    log.info("%s — left out: %s", filename, exc)
                status = LEFT_OUT
            else:
                log.warning("%s — %s; skipped", filename, exc)
                status = UNSUPPORTED
            report(FileOutcome(filename, filepath, status, pre.device_code,
                               pre.device_name, detail=str(exc).capitalize(),
                               code=pre.code), index)

        except Exception as exc:
            log.error("%s — %s", filename, exc)
            report(FileOutcome(filename, filepath, ERROR, pre.device_code,
                               pre.device_name, detail=str(exc), code=pre.code),
                   index)

    _settle_copies(reads)
    records: list[Record] = []
    for read in reads:
        if read.dropped:
            continue
        _mark_code(read)
        records.append(read.record)
        if read.second is not None:
            records.append(read.second)
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


def deduplicate_records(records: list[Record],
                        dropped: list[Duplicate] | None = None) -> list[Record]:
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
            if dropped is not None:
                dropped.append(Duplicate(serial, device, source, first))

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
    """Fill a copy of the template with the records and save it.

    On a Turbo-scale run this is the long tail: ~2.5s to fill 38,000 rows and
    ~7s for openpyxl to serialise them. A GUI caller sees stalls of up to half
    a second in there — measured, and they sit inside workbook.save(), so
    yielding the GIL around the fill loop does not touch them (tried; it moved
    the worst stall 483ms → 416ms and cost a second on the write). What the
    window needs is therefore to *say* what it is doing, which process_files
    logs before calling this, not to pretend the pause is not happening.
    """
    workbook = load_workbook(template_file)
    try:
        sheet = workbook.active
        for index, record in enumerate(records):
            row = TEMPLATE_START_ROW + index
            sheet.cell(row=row, column=1, value=index + 1)
            flags = record.get("_flags") or {}
            for offset, name in enumerate(FIELDS):
                mark = flags.get(name)
                value = record.get(name, "")
                if mark and mark.get("show") is not None:
                    value = mark["show"]
                cell = sheet.cell(row=row, column=TEMPLATE_START_COL + offset,
                                  value=value)
                if mark:
                    _mark_for_checking(cell, mark["reason"])

        stamp_attribution(sheet, workbook, len(records))
        workbook.save(output_path)
    finally:
        workbook.close()


#: The colour of a value kept in the register but not trusted. Amber rather
#: than red: it is a question for a person, not a verdict on the device.
FLAG_FILL = PatternFill("solid", fgColor="FFE699")


def _mark_for_checking(cell, reason: str) -> None:
    """Colour a doubtful value amber and say why, in a hover note.

    The column layout is untouched on purpose — anything that reads the
    register by position keeps working, and a row with nothing doubtful looks
    exactly as it always did.
    """
    cell.fill = FLAG_FILL
    note = Comment(f"Calist: {reason}", "Calist")
    note.width, note.height = 260, 70
    cell.comment = note


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
    real_forms_only: bool = False,
    output_dir: str | os.PathLike | None = None,
    on_file: ProgressHook | None = None,
    cancel: threading.Event | None = None,
    quiet: bool = False,
) -> RunResult:
    """Run the full pipeline and report what happened.

    Always returns a RunResult; check ``.succeeded`` or ``.output_path``. A
    cancelled run writes nothing and comes back with ``cancelled=True``.
    ``real_forms_only`` leaves out anything that is not a device form — see
    [extract_records].
    ``output_dir`` overrides where the register is written. ``quiet`` trades
    the line-per-file log for a heartbeat, which is what makes a run of tens of
    thousands readable.
    """
    rule = "─" * 55
    result = RunResult()

    if not source_files:
        result.error = "No source files selected."
        log.error("%s", result.error)
        return result

    log.info(rule)
    log.info("Template : %s", os.path.basename(template_file))
    log.info("Sources  : %s selected", plural(len(source_files), "file"))
    log.info(rule)

    try:
        output_path = resolve_output_path(source_files, template_file, output_dir)
    except ValueError as exc:
        result.error = str(exc)
        log.error("%s", exc)
        return result

    records, result.outcomes = extract_records(source_files, on_file, cancel,
                                               real_forms_only, quiet)
    result.second_rows_added = sum(1 for o in result.outcomes
                                   if o.rows == 2 and o.status == OK)
    records = sort_records(records)
    if result.copies:
        log.info("%s left out — the same device as a form already in the "
                 "register.", plural(len(result.copies), "copy", "copies"))
    if result.left_out:
        log.info("%s left out — not device forms.",
                 plural(len(result.left_out), "file"))

    if cancel is not None and cancel.is_set():
        result.cancelled = True
        log.warning("Run cancelled — no file written.")
        return result

    if deduplicate:
        before = len(records)
        records = deduplicate_records(records, result.duplicates)
        result.duplicates_removed = before - len(records)
        log.info("%s removed.", plural(result.duplicates_removed, "duplicate record")) \
            if result.duplicates_removed else log.info("No duplicates found.")

    if not records:
        result.error = "No valid records extracted. No file created."
        log.warning("%s", result.error)
        return result

    log.info("Writing %s…", plural(len(records), "row"))
    try:
        write_output(records, template_file, output_path)
    except Exception as exc:
        result.error = f"Could not save output file: {exc}"
        log.error("%s", result.error)
        return result

    result.output_path = output_path
    result.rows_written = len(records)

    log.info(rule)
    log.info("✔ Success! %s written.", plural(len(records), "record"))
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
    if outcome.code and outcome.code != path.stem:
        print(f"code   : {outcome.code}   (the name as the register writes it)")
    print(f"device : {outcome.device_name or '—'} ({outcome.device_code or '—'})")

    if outcome.status != READY:
        print(f"\nCannot inspect: {outcome.detail}")
        return 1

    config = DEVICE_CONFIGS[outcome.device_code]
    cells = config["cells"]
    record = read_record(filepath, cells)
    blank = [f for f, v in record.items() if not v]

    print(f"\n{'field':<14}{'cell':<7}value")
    print("─" * 60)
    for field, ref in cells.items():
        value = record[field].replace("\n", " / ")
        print(f"{field:<14}{ref:<7}{value if value else '(blank)'}")

    # What the pipeline would actually record, which is not the same thing once
    # the configured map stops fitting the form.
    try:
        best, layout = read_best(filepath, config)
    except ReadAsOther as other:
        config = DEVICE_CONFIGS[other.code]
        print(f"\nits form says it is a {other.stated!r} — the register reads it "
              f"as a {config['device_name']} ({other.code})")
        best, layout = read_best(filepath, config)
    except NotAForm as exc:
        print(f"\nnot written to the register: {exc}")
        return 1
    if layout != "primary":
        print(f"\nthe configured map does not fit this form — read using: {layout}")
        if layout == "none":
            print("no layout, alternate or printed label produced a usable record")
        else:
            for field in ("Model", "S.N", "Manufacturer", "Location"):
                if best.get(field):
                    print(f"  {field:<14}{best[field]}")

    if path.suffix.lower() != ".xls":
        # Asked of the reader itself, not of a second opinion: the whole point
        # of this dump is to say what the pipeline saw.
        source = _open_workbook(filepath)
        try:
            merges = source.merged_refs()
            names = source.sheet_names
            print(f"\nsheet read: {source.sheet_name!r} "
                  f"(of {len(names)}: {', '.join(names[:6])}"
                  f"{' …' if len(names) > 6 else ''})")
            if source.skipped_sheets:
                print("skipped empty leading tab(s): "
                      f"{', '.join(source.skipped_sheets[:3])}")
            print(f"merged ranges: {len(merges)}")
            if merges:
                print("  " + ", ".join(merges[:24])
                      + (" …" if len(merges) > 24 else ""))
        finally:
            source.close()

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
        # The report draws a box-rule and the forms carry Arabic site names,
        # and the Windows console still defaults to cp1252, which can encode
        # neither. Without this, --inspect does all its work and then dies on
        # the first print — on the one tool you reach for when a field is
        # blank, which is exactly when you need to see the output.
        for stream in (sys.stdout, sys.stderr):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except (AttributeError, ValueError):   # not a real console
                pass
        logging.basicConfig(level=logging.WARNING, format="%(levelname)s %(message)s")
        raise SystemExit(inspect_form(args[1]))

    from ui import run
    run()


if __name__ == "__main__":
    main()
