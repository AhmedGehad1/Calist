# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

**Calist** — a single-window Tkinter desktop app that harvests fixed cells out of many medical-device
inspection Excel forms and compiles them into one flat equipment register.

## Commands

```powershell
python calist.py                # launch the GUI (the only entry point)
python -m pytest                # run the test suite (pip install pytest first)
python -c "import calist"       # import check without opening a window
pip install -r requirements.txt # openpyxl + xlrd
```

`xlrd` 2.x reads **only** `.xls`; `.xlsx`/`.xlsm` go through `openpyxl`.

## Architecture

Two modules, one direction: [calist.py](calist.py) imports `DEVICE_CONFIGS` from
[device_config.py](device_config.py), which is pure data.

The pipeline is a chain of small functions orchestrated by
[`process_files()`](calist.py#L343), which does no work itself:

1. [`extract_device_code()`](calist.py#L89) — filename stem, split on the first `-`, leading letters of
   the right-hand part. `"Clinic-AGH001.xlsx"` → `"AGH"`.
2. `DEVICE_CONFIGS[code]["cells"]` — maps field names to A1 refs.
3. [`read_record()`](calist.py#L144) — reads those cells via [`_open_source()`](calist.py#L112), a
   context manager that hides the openpyxl/xlrd split behind one `get(ref)` function. **Always
   worksheet index 0.**
4. [`clean()`](calist.py#L75) — renders raw cell values as output strings.
5. [`build_second_row()`](calist.py#L157) — for configs with a `second_row` block, emits a sub-module
   row of the same physical unit.
6. [`sort_records()`](calist.py#L261) → optional [`deduplicate_records()`](calist.py#L267) →
   [`write_output()`](calist.py#L323).

### Logging is the seam between core and GUI

Core functions call `log.info/warning/error` on the module logger; they contain no reference to Tkinter.
The GUI attaches a [`TkLogHandler`](calist.py#L411) that marshals records onto the main thread via
`widget.after(0, ...)`. **Keep it that way** — anything that reaches into the GUI from pipeline code
breaks headless use and the tests.

### Value normalisation

`clean()` is a plain `str(value).strip()`, with `None` → `""`. This is deliberate: the forms in use hold
dates and serials as text, so nothing needs converting.

It does mean non-text cells pass through in Excel's own form — a real date cell becomes
`"2024-01-15 00:00:00"`, a numeric serial `"123456.0"`, and in `.xls` a date is the bare serial number
(`"45306.0"`) because xlrd returns it as a float. If a form ever starts using real date or number cells,
that's where to fix it. Tests lock the current behaviour in, so a change there is a deliberate one.

`load_workbook(data_only=True)` returns *cached* formula results — a file written by a script and never
opened in Excel yields `None` for those cells.

### Sorting is data-driven, not hard-coded

Records carry `_group` (the parent's Code) and `_row_order` (0 = parent, 1 = generated sub-module), set
at creation. `sort_records` sorts on those, so adding a two-row device needs only a config entry — no
code change.

`natural_key` orders embedded numbers numerically. Device codes are zero-padded (`AGH001`), so plain
string sort already gives the right order and this is a no-op on real filenames; it only matters if an
unpadded code ever appears. Safe to drop if that's never going to happen.

The `_`-prefixed keys never reach the template because the writer iterates `FIELDS`.

### Template contract

[template/Device List.xlsx](template/Device%20List.xlsx): sheet `list`, headers row 3, data from row 4,
column A = index,
columns B–I = `FIELDS`. The writer uses `workbook.active`, so the template's active sheet must be the
data sheet. Note the template ships with column A pre-numbered down to row 84; nothing clears those, so
output files always show ~81 numbered-but-empty trailing rows.

### `device_config.py`

51 of 57 forms share one layout, so cell maps are built by [`form()`](device_config.py#L19) rather than
written out: `form(row, status)` where `row` is the Model row and the rest sit at fixed offsets
(Manufacturer `row+2`, S.N `val{row}`, Location `val{row+2}`, Date `row-date_gap`). `col`/`val` switch
the column pair (E/K default, D/J and F/L exist); `date_gap=4` covers forms with an extra line above the
Date; `extra={...}` adds `S.N2`/`Status2` or overrides a cell.

Adding a device is one line. Only genuinely different forms (`AK`, `CF`) get literal dicts — that
asymmetry is deliberate, so odd forms stand out.

## Coupling to keep in mind

`ALLOWED_SHARED_SN_PAIRS` ([calist.py:61](calist.py#L61)) holds `device_name` strings verbatim from
`device_config.py`; renaming a device there breaks the exemption that lets a Patient Monitor and its
NIBP row share a serial. `test_second_row_names_are_covered_by_the_dedup_exemptions` guards this — run
the tests after renaming anything.

## Behaviour worth knowing

- An unknown device code **skips the file** with an error rather than emitting a junk row. Set
  `SKIP_UNKNOWN_CODES = False` ([calist.py:49](calist.py#L49)) to restore the old A1:A6 fallback.
- Output is `device list.xlsx` beside the first source file. `resolve_output_path()` refuses to run if
  that would overwrite the template (Windows paths are case-insensitive, so it would otherwise clobber
  `Device List.xlsx`).
- Because the output lands *among* the sources, selecting a whole folder twice feeds the previous run's
  output back in as an input. Its code resolves to `DEVICE`, which isn't in `DEVICE_CONFIGS`, so it is
  skipped with an error.
- Per-file failures are logged and skipped; a partial output is still written. Read the status log.

## Open data questions

Flagged in `device_config.py` and the README, unresolved — they need checking against the paper forms:

- `EU` (Lab Oven) has Location `K19`; every other standard form puts it at `row+2` = `K20`.
- `CF` (Baby Warmer) has Model *below* Manufacturer, inverted vs. every other form.
- `CA` is named `"X-ray ()"` and is otherwise identical to `BF`.
- `AO` and `CK` are both named `"Infrared"` with different Status cells.

Device names are reproduced verbatim in output, spelling slips included (`Protien Analyzer`,
`Tornique`). Correcting them changes the text written into every register, so treat it as a deliberate
data change, not a typo fix.

## Repo

Public GitHub repo at `AhmedGehad1/Calist`, default branch `main`. CI runs the test suite on
windows-latest across Python 3.10–3.12 ([.github/workflows/tests.yml](.github/workflows/tests.yml)).
There is deliberately **no LICENSE file** — all rights reserved.

Never commit completed inspection forms; they carry real device and site data. `.gitignore` covers the
common patterns plus Calist's own `device list.xlsx` output.
