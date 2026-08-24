<h1 align="center">Calist</h1>

<p align="center">
  <em>Compile a folder of device inspection forms into one equipment register.</em>
</p>

<p align="center">
  <a href="https://github.com/AhmedGehad1/Calist/actions/workflows/tests.yml">
    <img alt="Tests" src="https://github.com/AhmedGehad1/Calist/actions/workflows/tests.yml/badge.svg">
  </a>
  <img alt="Python" src="https://img.shields.io/badge/python-3.10%2B-blue">
  <img alt="Devices" src="https://img.shields.io/badge/device%20types-57-brightgreen">
  <img alt="Platform" src="https://img.shields.io/badge/platform-Windows-lightgrey">
</p>

<p align="center">
  <img src="docs/ui-2-review.png" alt="Calist with a folder of inspection forms loaded" width="860">
</p>

---

## The problem

Biomedical equipment inspections are recorded one Excel form per device. A hospital wing might produce
several hundred of them in a round, and each one holds the same six facts — manufacturer, model, serial,
location, date, status — in six fixed cells.

The catch is that **the cells are in a different place on every device's form.** A defibrillator's
serial is in `K15`; an ultrasound's is in `L17`; a baby incubator's is in `L72`. Building the equipment
register by hand means opening every file, remembering which layout you're looking at, and copying six
cells out of it.

## What Calist does

Point it at a folder of completed inspection forms and a blank register template. It reads the device
type from each filename, looks up that device's cell layout, pulls the six fields out, and writes one
row per device into a copy of your template — sorted, numbered, and optionally de-duplicated.

Several hundred forms take a few seconds.

## How it works

```
Clinic-AGH001.xlsx
        │
        ├─ 1. read device code from filename ──────────►  "AGH"
        │
        ├─ 2. look up its layout in device_config.py ──►  Model E18, S.N K18, …
        │
        ├─ 3. read those cells from sheet 1 ───────────►  {Model: "MX450", S.N: "SN-100", …}
        │
        ├─ 4. generate the sub-module row, if any ─────►  + a second "NIBP" row
        │
        └─ 5. sort, de-duplicate, write into template ─►  device list.xlsx
```

## Install

```bash
git clone https://github.com/AhmedGehad1/Calist.git
cd Calist
pip install -r requirements.txt
python calist.py
```

Requires Python 3.10 or newer. `openpyxl` handles `.xlsx`/`.xlsm`, `xlrd` handles legacy `.xls`, and
`customtkinter` draws the interface.

Installing `tkinterdnd2` as well lets you drag a folder straight onto the window; without it the drop
zone is click-only and nothing else changes.

## Usage

**Point it at the folder.** Everything else follows from that.

<table>
<tr>
<td width="50%"><img src="docs/ui-1-hero.png" alt="Empty state"></td>
<td width="50%"><img src="docs/ui-3-working.png" alt="Building the register"></td>
</tr>
<tr>
<td><b>1 · Add your devices.</b> Picking a folder pulls in every Excel file inside it, subfolders
included. Each one is checked on arrival, so an unrecognised device code shows up in the table
straight away — before you commit to a run, not after it.</td>
<td><b>2 · Watch it work.</b> Rows turn green as each device is read, with a progress bar, the file
currently open, and an estimate of the time left. Cancel stops it cleanly without writing anything.</td>
</tr>
</table>

<p align="center">
  <img src="docs/ui-4-results.png" alt="Results" width="720">
</p>

**3 · Collect the register.** The results card says what was built and exactly where it went, with
**Open register** and **Reveal in folder** one click away. The table filters itself to anything that
needs attention.

The register is saved as `device list.xlsx` **next to the first source file** — the destination is
shown in the *Saves to* row throughout, so it is never a surprise. If a register is already there,
the row says so before you build.

The template is remembered between sessions, so choosing it is a first-run step only. Duplicate
serial removal is optional and off by default.

### Keyboard

`Ctrl+O` add a folder · `Ctrl+Enter` build · `Esc` cancel a run · `Delete` remove selected rows ·
double-click a row to reveal that file.

### Filename convention

The device type comes from the filename, so this part matters:

```
Clinic-AGH001.xlsx
       └┬┘
        └──  everything after the first "-", leading letters only  →  AGH
```

If there is no `-`, the whole name is used (`VNT023.xlsx` → `VNT`). A file whose code isn't in the
device table is **skipped with an error** rather than silently producing a junk row.

## The device table

All 57 layouts live in [`device_config.py`](device_config.py). Because 51 of the 57 forms are the same
layout at a different row offset, they're built by a helper rather than written out by hand:

```python
"DG": {"device_name": "CBC Analyzer", "cells": form(18, "H32")},
```

`form(row, status)` takes the row holding the **Model** and derives the rest from it:

```
                      form(18, "H32")
      Date          E16   ← row - 2
      Model         E18   ← row          the anchor
      Manufacturer  E20   ← row + 2
      S.N           K18   ← row,     value column
      Location      K20   ← row + 2, value column
      Status        H32   ← given explicitly, it moves the most
```

Keyword arguments cover the variations:

| Argument | Use when | Example |
|---|---|---|
| `col=` / `val=` | the form uses a different column pair | `form(32, "F41", col="D", val="J")` |
| `date_gap=4` | an extra line sits above the Date | `form(26, "G35", date_gap=4)` |
| `extra={...}` | a second serial, second status, or a one-off cell | `form(17, "H30", extra={"S.N2": "L21"})` |

Genuinely different forms (`AK`, `CF`) are written out as literal dicts, so the odd ones stand out
instead of hiding in a wall of near-identical blocks.

### Adding a device type

One line. Find the Model cell on the form, note the Status cell, and add:

```python
"XY": {"device_name": "Your Device", "cells": form(<model row>, "<status cell>")},
```

## Two-row devices

Some units are inspected as one device but recorded as two — a patient monitor and its NIBP module
share a chassis and a serial number, but each gets its own status and its own line in the register.
A `second_row` block generates that line automatically:

```python
"AGH": {
    "device_name": "Patient Monitor",
    "cells": form(18, "D39", extra={"Status2": "J39"}),
    "second_row": {"device_name": "NIBP", "code_replace": ("AGH", "AGCB")},
},
```

The generated row copies the parent's data, takes its status from `Status2`, and rewrites the code
(`Clinic-AGH001` → `Clinic-AGCB001`). It always sorts directly beneath its parent.

## Duplicate serial numbers

With **Remove duplicate S/N** ticked, a repeated serial is dropped and logged. The exception is a
device and its own generated sub-module — they legitimately share one serial because they are one
physical unit. A *third* record on that serial is still dropped.

Blank serials are always kept; several devices with no serial recorded shouldn't collapse into one row.

## Supported devices

| | | |
|---|---|---|
| `AA` Anesthesia | `AV` Elisa reader | `EC` Laminar flow |
| `AB` Vaporizer | `AX` Lab Incubator | `ED` Heart lung Machine |
| `AC` Defibrillator | `BB` Ultrasound | `EE` Flowmeter |
| `AD` Pacemaker | `BF` X-ray | `EO` Pipet |
| `AE` ESU | `BL` Autoclave | `EP` Refrigerator |
| `AF` ECG | `BP` Balance | `EU` Lab Oven |
| `AGH` Patient Monitor | `BV` Blood gas analyzer | `EV` Blood Mixer |
| `AH` SPO2 | `BZ` Syringe | `EY` Freezer |
| `AI` Infusion | `CA` X-ray () | `FE` Nebulizer |
| `AJ` Suction | `CB` Digital blood pressure | `FG` ACT |
| `AK` Baby Incubator | `CE` Sphygmomanometer | `FI` Hormone Analyzer |
| `AL` Phototherapy | `CF` Baby Warmer | `FJ` OR Table |
| `AM` Ventilator | `CK` Infrared | `FQ` C-pap |
| `AN` Thermo | `DA` Shaker | `GC` Portable Data Logger |
| `AO` Infrared | `DG` CBC Analyzer | `GD` Protien Analyzer |
| `AQ` Water Bath | `DL` Sealing Machine | `GI` Bacteria Analyzer |
| `AR` Electrolyte Analyzer | `DO` O2 conc | `GK` Tornique |
| `AS` Centrifuge | `DV` OR light | `GP` Holter machines |
| `AU` Chemistry analyzer | `EA` C-Arm | `VAH` Vital Sign (SPO2 Module) |

<sub>Names are reproduced exactly as they appear in the register output — including the
spelling slips noted below.</sub>

## Development

```bash
pip install pytest
python -m pytest
```

35 tests covering the logic with no I/O — filename parsing, value handling, ordering, de-duplication,
second-row generation and pre-flight classification — plus end-to-end runs against workbooks built on
the fly, and integrity checks over the device table itself. One of those asserts that every two-row
device's names are covered by the shared-serial exemption, so renaming a device can't silently break
de-duplication.

The suite imports `calist` only, never `ui`, so it needs no display and no GUI toolkit.

### Layout

```
calist.py                      the pipeline — no GUI toolkit imported
ui.py                          the desktop interface (customtkinter)
device_config.py               the 57 device layouts and the form() helper
test_calist.py                 the test suite
template/Device List.xlsx      reference register template
```

The pipeline knows nothing about the interface. `calist.py` imports no GUI toolkit at all — `ui.py`
attaches a logging handler to pick up progress, and reads the structured `FileOutcome` / `RunResult`
values the pipeline returns. So the whole thing drives headlessly:

```python
from calist import process_files

result = process_files(["Clinic-AGH001.xlsx"], "Device List.xlsx", deduplicate=True)
print(result.rows_written, "rows ->", result.output_path)
for problem in result.problems:
    print(problem.filename, problem.detail)
```

`process_files` also takes `on_file=` for per-file progress and `cancel=` (a `threading.Event`) to
stop a long run — a cancelled run writes nothing.

To check what a filename *would* resolve to without opening it:

```python
from calist import classify_file
classify_file("Clinic-AGH001.xlsx").device_name   # 'Patient Monitor  +NIBP'
```

## Known issues

Four entries in the device table need checking against the paper forms:

- `EU` (Lab Oven) puts Location at `K19`; every other standard form puts it at `K20`.
- `CF` (Baby Warmer) has Model *below* Manufacturer — inverted compared with all 56 others.
- `CA` is identical to `BF` and its name is unfinished in the source (`"X-ray ()"`).
- `AO` and `CK` are both named `Infrared` with different Status cells.

Three device names carry spelling slips that reach the register output — `Protien Analyzer`,
`Tornique`, and `X-ray ()`. They're left as-is deliberately: correcting them changes the text written
into every historical register, so it should be a deliberate call rather than a drive-by fix.

Also: because the output lands among the source files, selecting the same folder twice will feed the
previous run's output back in. Its code resolves to `DEVICE`, which isn't in the table, so it's skipped
with an error rather than corrupting the register.

## License

No license is granted. The source is published for reference; all rights are reserved by the author.
If you'd like to use it, please get in touch.
