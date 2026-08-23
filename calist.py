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
import tkinter as tk
from contextlib import contextmanager
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk
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
    progress_callback: Callable[[], None] | None = None,
) -> list[Record]:
    """Read every source file and return the records to be written."""
    records: list[Record] = []

    for filepath in sorted(source_files):
        filename = os.path.basename(filepath)
        extension = Path(filepath).suffix.lower()

        if extension not in SUPPORTED_EXTENSIONS:
            log.warning("%s — unsupported format '%s'", filename, extension)
            if progress_callback:
                progress_callback()
            continue

        code_key = extract_device_code(filename)
        config = DEVICE_CONFIGS.get(code_key)

        if config is None:
            if SKIP_UNKNOWN_CODES:
                log.error(
                    "%s — code '%s' is not in device_config.py; file skipped",
                    filename, code_key,
                )
                if progress_callback:
                    progress_callback()
                continue
            log.warning("%s — code '%s' not found, using fallback cell map",
                        filename, code_key)
            config = UNKNOWN_FALLBACK

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
            log.info(
                "[OK]    %s  →  %s (%s)  |  Model: %s  |  S/N: %s  |  Status: %s",
                filename, config["device_name"], code_key,
                record.get("Model", ""), record.get("S.N", ""), record.get("Status", ""),
            )

            if "second_row" in config:
                extra = build_second_row(record, config["second_row"])
                records.append(extra)
                log.info(
                    "        ↳ 2nd row added: %s  |  Code: %s  |  Status: %s",
                    extra["Device"], extra["Code"], extra["Status"],
                )

        except Exception as exc:
            log.error("%s — %s", filename, exc)

        if progress_callback:
            progress_callback()

    return records


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
            for offset, field in enumerate(FIELDS):
                sheet.cell(row=row, column=TEMPLATE_START_COL + offset,
                           value=record.get(field, ""))
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
    progress_callback: Callable[[], None] | None = None,
) -> Path | None:
    """Run the full pipeline. Returns the output path, or None on failure."""
    rule = "─" * 55

    if not source_files:
        log.error("No source files selected.")
        return None

    log.info(rule)
    log.info("Template : %s", os.path.basename(template_file))
    log.info("Sources  : %d file(s) selected", len(source_files))
    log.info(rule)

    try:
        output_path = resolve_output_path(source_files, template_file)
    except ValueError as exc:
        log.error("%s", exc)
        return None

    records = sort_records(extract_records(source_files, progress_callback))

    if deduplicate:
        before = len(records)
        records = deduplicate_records(records)
        removed = before - len(records)
        log.info("%d duplicate record(s) removed.", removed) if removed \
            else log.info("No duplicates found.")

    if not records:
        log.warning("No valid records extracted. No file created.")
        return None

    try:
        write_output(records, template_file, output_path)
    except Exception as exc:
        log.error("Could not save output file: %s", exc)
        return None

    log.info(rule)
    log.info("✔ Success! %d record(s) written.", len(records))
    log.info("NEW FILE SAVED AT: %s", output_path)
    return output_path


# ──────────────────────────────────────────────────────────────────────────────
# GUI
# ──────────────────────────────────────────────────────────────────────────────

# Catppuccin Mocha
BG = "#1e1e2e"
SURFACE = "#313244"
TEXT = "#cdd6f4"
SUBTLE = "#a6adc8"
BLUE = "#89b4fa"
GREEN = "#a6e3a1"
RED = "#f38ba8"
PEACH = "#fab387"
GREY = "#6c7086"

FONT = "Segoe UI"


class TkLogHandler(logging.Handler):
    """Routes log records to a Tk text widget from any thread."""

    def __init__(self, widget: scrolledtext.ScrolledText):
        super().__init__()
        self.widget = widget

    def emit(self, record: logging.LogRecord) -> None:
        message = self.format(record)
        # Tk is not thread-safe; hop back to the main loop before touching it.
        self.widget.after(0, self._append, message)

    def _append(self, message: str) -> None:
        self.widget.config(state="normal")
        self.widget.insert("end", message + "\n")
        self.widget.see("end")
        self.widget.config(state="disabled")


class StatusFormatter(logging.Formatter):
    """Tags warnings and errors, leaving ordinary progress lines unadorned."""

    LABELS = {logging.WARNING: "WARN", logging.ERROR: "ERROR",
              logging.CRITICAL: "FATAL"}

    def format(self, record: logging.LogRecord) -> str:
        message = record.getMessage()
        label = self.LABELS.get(record.levelno)
        return f"[{label}]  {message}" if label else message


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Calist")
        self.configure(bg=BG)
        self.geometry("820x580")
        self.minsize(640, 460)

        self._source_files: set[str] = set()
        self._template_file = tk.StringVar(value="")
        self._dedup = tk.BooleanVar(value=False)
        self._done_count = 0

        self._build_ui()
        self._attach_logging()

    # ── construction ─────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        self._build_header()
        self._build_toolbar()
        self._build_selection_labels()
        self._build_log_panel()
        self._build_progress()
        self._build_footer()

    def _build_header(self) -> None:
        tk.Label(self, text="Calist", font=(FONT, 17, "bold"),
                 bg=BG, fg=TEXT).pack(pady=(18, 2))
        tk.Label(self, text="Compile device inspection forms into one equipment register",
                 font=(FONT, 9), bg=BG, fg=SUBTLE).pack(pady=(0, 12))

    def _build_toolbar(self) -> None:
        bar = tk.Frame(self, bg=BG)
        bar.pack(fill="x", padx=14, pady=6)

        self._btn_sources = self._button(bar, "📂  Select Source Files", self._pick_sources, BLUE)
        self._btn_sources.pack(side="left", padx=(0, 8))

        self._btn_template = self._button(bar, "📄  Select Template File", self._pick_template, GREEN)
        self._btn_template.pack(side="left", padx=(0, 8))

        self._btn_process = self._button(bar, "▶  Process Data", self._run_processing,
                                         RED, state="disabled")
        self._btn_process.pack(side="left")

        tk.Checkbutton(bar, text="Remove duplicate S/N", variable=self._dedup,
                       font=(FONT, 9), bg=BG, fg=TEXT, selectcolor=SURFACE,
                       activebackground=BG, activeforeground=TEXT,
                       cursor="hand2").pack(side="left", padx=(12, 0))

    def _build_selection_labels(self) -> None:
        self._lbl_sources = tk.Label(self, text="No source files selected.",
                                     font=(FONT, 9), bg=BG, fg=SUBTLE, anchor="w")
        self._lbl_sources.pack(fill="x", padx=14, pady=(4, 0))

        self._lbl_template = tk.Label(self, text="No template selected.",
                                      font=(FONT, 9), bg=BG, fg=SUBTLE, anchor="w")
        self._lbl_template.pack(fill="x", padx=14, pady=(0, 6))

    def _build_log_panel(self) -> None:
        frame = tk.Frame(self, bg=SURFACE, bd=1, relief="sunken")
        frame.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        tk.Label(frame, text=" Status Log", font=(FONT, 9, "bold"),
                 bg=SURFACE, fg=TEXT, anchor="w").pack(fill="x")

        self._log_box = scrolledtext.ScrolledText(
            frame, wrap="word", font=("Consolas", 9), bg=BG, fg=TEXT,
            insertbackground=TEXT, relief="flat", state="disabled", height=14,
        )
        self._log_box.pack(fill="both", expand=True, padx=4, pady=(0, 4))

    def _build_progress(self) -> None:
        frame = tk.Frame(self, bg=BG)
        frame.pack(fill="x", padx=14, pady=(0, 4))

        self._prog_label = tk.Label(frame, text="", font=(FONT, 8),
                                    bg=BG, fg=SUBTLE, anchor="w")
        self._prog_label.pack(fill="x")

        self._progress = ttk.Progressbar(frame, orient="horizontal", mode="determinate")
        self._progress.pack(fill="x")

    def _build_footer(self) -> None:
        frame = tk.Frame(self, bg=BG)
        frame.pack(anchor="e", padx=14, pady=(0, 10))

        self._btn_clear_sources = self._button(frame, "🗑  Clear Sources",
                                               self._clear_sources, PEACH, width=15)
        self._btn_clear_sources.pack(side="left", padx=(0, 8))

        self._button(frame, "🗑  Clear Log", self._clear_log, GREY, width=12).pack(side="left")

    def _button(self, parent, text, command, colour, state="normal", width=None) -> tk.Button:
        options = dict(
            text=text, command=command, font=(FONT, 9, "bold"),
            bg=colour, fg=BG, activebackground=colour, activeforeground=BG,
            relief="flat", cursor="hand2", padx=10, pady=6, state=state,
        )
        if width:
            options["width"] = width
        return tk.Button(parent, **options)

    def _attach_logging(self) -> None:
        handler = TkLogHandler(self._log_box)
        handler.setFormatter(StatusFormatter())
        log.addHandler(handler)
        log.setLevel(logging.INFO)
        log.propagate = False

    # ── actions ──────────────────────────────────────────────────────────────

    def _pick_sources(self) -> None:
        chosen = filedialog.askopenfilenames(
            title="Select Source Excel Files",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")],
        )
        if not chosen:
            return
        self._source_files.update(chosen)
        self._lbl_sources.config(
            text=f"{len(self._source_files)} source file(s) selected.", fg=BLUE)
        log.info("Added %d file(s). Total: %d", len(chosen), len(self._source_files))
        self._refresh_process_btn()

    def _pick_template(self) -> None:
        chosen = filedialog.askopenfilename(
            title="Select Template Excel File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not chosen:
            return
        self._template_file.set(chosen)
        self._lbl_template.config(text=f"Template: {os.path.basename(chosen)}", fg=GREEN)
        log.info("Template selected: %s", os.path.basename(chosen))
        self._refresh_process_btn()

    def _clear_sources(self) -> None:
        self._source_files.clear()
        self._lbl_sources.config(text="No source files selected.", fg=SUBTLE)
        log.info("Cleared all source files.")
        self._refresh_process_btn()

    def _clear_log(self) -> None:
        self._log_box.config(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.config(state="disabled")

    def _refresh_process_btn(self) -> None:
        ready = bool(self._source_files) and bool(self._template_file.get())
        self._btn_process.config(state="normal" if ready else "disabled")

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        for button in (self._btn_sources, self._btn_template, self._btn_clear_sources):
            button.config(state=state)
        self._btn_process.config(
            state="disabled" if busy else "normal",
            text="⏳  Processing…" if busy else "▶  Process Data",
        )

    def _run_processing(self) -> None:
        self._set_busy(True)

        files = sorted(self._source_files)
        total = len(files)
        self._done_count = 0
        self._progress.config(maximum=total, value=0)
        self._prog_label.config(text=f"0 / {total} files processed")

        def on_progress() -> None:
            self._done_count += 1
            done = self._done_count      # capture now; the lambda runs later
            self.after(0, lambda: (
                self._progress.config(value=done),
                self._prog_label.config(
                    text=f"{done} / {total} files processed  ({total - done} remaining)"),
            ))

        def worker() -> None:
            try:
                output_path = process_files(
                    files, self._template_file.get(),
                    deduplicate=self._dedup.get(),
                    progress_callback=on_progress,
                )
            except Exception as exc:                 # never let the thread die silently
                log.exception("Unexpected failure: %s", exc)
                output_path = None
            self.after(0, self._on_done, output_path)

        threading.Thread(target=worker, daemon=True).start()

    def _on_done(self, output_path: Path | None) -> None:
        self._set_busy(False)
        self._refresh_process_btn()
        if output_path:
            self._prog_label.config(text=f"✔ Done — saved to {output_path.name}")
        else:
            self._prog_label.config(text="✖ Finished with errors — see the log.")
            messagebox.showerror(
                "Processing failed",
                "No output file was produced. Check the status log for details.",
            )


def main() -> None:
    App().mainloop()


if __name__ == "__main__":
    main()
