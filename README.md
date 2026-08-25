<div align="center">

<img src="docs/calist-icon.png" alt="Calist" width="96">

# Calist

**Turn a folder of device inspection forms into one equipment register — in seconds.**

[![Download](https://img.shields.io/badge/⬇%20Download%20for%20Windows-2ea44f?style=for-the-badge)](https://github.com/AhmedGehad1/Calist/releases/latest/download/Calist.exe)

[![Tests](https://github.com/AhmedGehad1/Calist/actions/workflows/tests.yml/badge.svg)](https://github.com/AhmedGehad1/Calist/actions/workflows/tests.yml)
[![Release](https://github.com/AhmedGehad1/Calist/actions/workflows/release.yml/badge.svg)](https://github.com/AhmedGehad1/Calist/actions/workflows/release.yml)
![Python](https://img.shields.io/badge/python-3.10%20|%203.11%20|%203.12-blue)
![Device types](https://img.shields.io/badge/device%20types-57-brightgreen)
![Tests](https://img.shields.io/badge/tests-35%20passing-brightgreen)
![Platform](https://img.shields.io/badge/platform-Windows%2010%20|%2011-lightgrey)

<img src="docs/ui-2-review.png" alt="Calist with a folder of inspection forms loaded" width="880">

</div>

---

Biomedical equipment inspections are recorded one Excel form per device. A hospital wing produces
several hundred in a round, and every one holds the same six facts — manufacturer, model, serial,
location, date, status.

The catch: **those six cells sit in a different place on every device's form.** A defibrillator's
serial is in `K15`. An ultrasound's is in `L17`. A baby incubator's is in `L72`. Compiling the
register by hand means opening every file, working out which layout you're looking at, and copying
six cells — several hundred times.

Calist knows all 57 layouts. Point it at the folder and it does the round in about six seconds.

<div align="center">

| Doing it by hand | With Calist |
|---|---|
| Open 300 files one at a time | Pick the folder once |
| Remember 57 different cell layouts | Recognised automatically from the filename |
| Discover a bad file at row 214 | Flagged before the run starts |
| Retype serials and hope | Read straight from the cell |
| An afternoon | **~6 seconds** |

</div>

## Contents

[Download](#download) · [Features](#features) · [How it works](#how-it-works) ·
[Performance](#performance) · [Filename convention](#filename-convention) ·
[The device table](#the-device-table) · [Two-row devices](#two-row-devices) ·
[Duplicate serials](#duplicate-serial-numbers) · [Headless use](#headless-use) ·
[Development](#development) · [Known issues](#known-issues)

## Features

| | |
|---|---|
| **Folder-first** | Pick one folder and every Excel file inside it is pulled in, subfolders included. No hand-picking 300 files. |
| **Validated before you build** | Every file is resolved to a device the moment it's added. An unrecognised code shows up in the table *immediately* — not 200 files into a run. |
| **You can see it working** | Rows turn green one by one as each device is read, with a progress bar, the file currently open, and time remaining. |
| **Cancel any time** | Stops cleanly between files and writes nothing. |
| **The destination is never a surprise** | The *Saves to* row shows exactly where the register will land, before you commit — and warns if a register is already there. |
| **Handles two-row devices** | A patient monitor and its NIBP module share a chassis but need separate lines. Calist generates the second row and sorts it beneath its parent. |
| **Smart duplicate removal** | Optional. Drops repeated serials, but knows a device and its own sub-module legitimately share one. |
| **Nothing to set up** | The register template ships inside the executable, so a freshly downloaded copy works on first launch. Swap in your own whenever you like. |
| **Remembers your setup** | Template, last folder and preferences persist between sessions. |
| **Old and new Excel** | `.xlsx`, `.xlsm` via openpyxl; legacy `.xls` via xlrd. |
| **Scriptable** | The pipeline imports no GUI toolkit, so it drives headlessly from Python. |

## Download

### [⬇ Download Calist.exe](https://github.com/AhmedGehad1/Calist/releases/latest/download/Calist.exe)

**One file. No installer, no Python, nothing to configure.** Download it, double-click, and point it
at your folder of forms. The register template is built in, so the first run works immediately.

Works on Windows 10 and 11. About 12 MB. Every release is built and tested automatically by
[GitHub Actions](https://github.com/AhmedGehad1/Calist/actions/workflows/release.yml) from the
source in this repository — see [all releases](https://github.com/AhmedGehad1/Calist/releases).

> **First launch:** Windows shows *"Windows protected your PC"* because the file isn't
> code-signed — a certificate costs a few hundred dollars a year, and this is a free tool.
> Click **More info → Run anyway**. You only see it once.

<details>
<summary><b>Run from source instead</b></summary>

<br>

```bash
git clone https://github.com/AhmedGehad1/Calist.git
cd Calist
pip install -r requirements.txt
python calist.py
```

Python 3.10 or newer. `openpyxl`, `xlrd` and `customtkinter` install from `requirements.txt`.

`pip install tkinterdnd2` additionally enables dragging a folder straight onto the window; without
it the drop zone is click-only and nothing else changes.

</details>

<details>
<summary><b>Build the executable yourself</b></summary>

<br>

```bash
pip install pyinstaller
pyinstaller calist.spec --noconfirm --clean
```

The result is `dist/Calist.exe`. [`calist.spec`](calist.spec) is the build recipe; two settings in it
are load-bearing and commented as such — `hiddenimports=["ui"]`, because the interface is imported
lazily and PyInstaller's static analysis cannot see it, and `collect_data_files("customtkinter")`,
because CustomTkinter loads its themes and fonts from disk at runtime.

</details>

## How it works

<table>
<tr>
<td width="50%"><img src="docs/ui-1-hero.png" alt="Adding devices"></td>
<td width="50%"><img src="docs/ui-3-working.png" alt="Building the register"></td>
</tr>
<tr valign="top">
<td>

**1 · Add your devices**

Pick a folder and every Excel file inside it comes in, subfolders included.

Each one is checked on arrival, so an unrecognised device code appears in the
table straight away — before you commit to a run.

</td>
<td>

**2 · Watch it work**

Rows turn green as each device is read, with the file currently open and an
estimate of the time left.

Cancel stops cleanly without writing anything.

</td>
</tr>
</table>

<div align="center">
<img src="docs/ui-4-results.png" alt="Results" width="760">
</div>

**3 · Collect the register** — the results card says what was built and exactly where it went, with
**Open register** and **Reveal in folder** one click away. The table filters itself down to anything
that needs attention.

The register is saved as `device list.xlsx` **beside the first source file**. That location is shown
in the *Saves to* row the whole time, so it's never a surprise.

<details>
<summary><b>What happens to a single form</b></summary>

```
Clinic-AGH001.xlsx
        │
        ├─ 1. read the device code from the filename ──►  "AGH"
        │
        ├─ 2. look up its layout in device_config.py ──►  Model E18, S.N K18, …
        │
        ├─ 3. read those cells from sheet 1 ───────────►  {Model: "MX450", S.N: "SN-100", …}
        │
        ├─ 4. generate the sub-module row, if any ─────►  + a second "NIBP" row
        │
        └─ 5. sort, de-duplicate, write to template ───►  device list.xlsx
```

</details>

### Keyboard

| | |
|---|---|
| <kbd>Ctrl</kbd>+<kbd>O</kbd> | Add a folder |
| <kbd>Ctrl</kbd>+<kbd>Enter</kbd> | Build the register |
| <kbd>Esc</kbd> | Cancel a run |
| <kbd>Delete</kbd> | Remove selected rows |
| Double-click | Reveal that file in Explorer |

## Performance

Measured on 300 forms spanning five different device layouts, on a normal desktop machine:

| | |
|---|---|
| Full run, 300 forms → 360 rows | **~5.5 seconds** (≈55 forms/sec) |
| Pre-flight validation, 300 filenames | **under 5 ms** (~14 µs each) |

Pre-flight costs so little because it never opens a workbook — it resolves the filename against the
device table and nothing more. That's what makes validating-on-add practical even for a large folder.

## Filename convention

The device type comes from the filename, so this part matters:

```
Clinic-AGH001.xlsx
       └┬┘
        └──  everything after the first "-", leading letters only  →  AGH
```

If there's no `-`, the whole name is used (`VNT023.xlsx` → `VNT`). A file whose code isn't in the
device table is **skipped with an error** rather than silently producing a junk row.

## The device table

All 57 layouts live in [`device_config.py`](device_config.py). Because 51 of the 57 forms are the
same layout at a different row offset, they're built by a helper rather than written out by hand:

```python
"DG": {"device_name": "CBC Analyzer", "cells": form(18, "H32")},
```

`form(row, status)` takes the row holding the **Model** and derives the rest from it:

```
                      form(18, "H32")

      Date          E16   ← row - 2
      Model         E18   ← row            the anchor
      Manufacturer  E20   ← row + 2
      S.N           K18   ← row,     value column
      Location      K20   ← row + 2, value column
      Status        H32   ← given explicitly; it moves the most
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

<details>
<summary><b>All 57 supported devices</b></summary>

<br>

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

<sub>Names appear exactly as they're written into the register — including the spelling slips noted
under <a href="#known-issues">Known issues</a>.</sub>

</details>

## Two-row devices

Some units are inspected as one device but recorded as two. A patient monitor and its NIBP module
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

With **Remove duplicate serial numbers** switched on, a repeated serial is dropped and logged.

The exception is a device and its own generated sub-module — they legitimately share one serial,
because they are one physical unit. A *third* record on that serial is still dropped.

Blank serials are always kept: several devices with no serial recorded shouldn't collapse into one row.

## Headless use

The pipeline knows nothing about the interface — `calist.py` imports no GUI toolkit at all. So the
whole thing drives from Python:

```python
from calist import process_files

result = process_files(["Clinic-AGH001.xlsx"], "Device List.xlsx", deduplicate=True)

print(result.rows_written, "rows ->", result.output_path)
for problem in result.problems:
    print(problem.filename, problem.detail)
```

`process_files` also takes `on_file=` for per-file progress and `cancel=` (a `threading.Event`) to
stop a long run — a cancelled run writes nothing.

To check what a filename *would* resolve to, without opening it:

```python
from calist import classify_file

classify_file("Clinic-AGH001.xlsx").device_name   # 'Patient Monitor  +NIBP'
classify_file("Clinic-ZZZ999.xlsx").status        # 'unknown_code'
```

## Development

```bash
pip install pytest
python -m pytest
```

35 tests, run in CI against Python 3.10, 3.11 and 3.12. They cover the logic with no I/O — filename
parsing, value handling, ordering, de-duplication, second-row generation and pre-flight
classification — plus end-to-end runs against workbooks built on the fly, and integrity checks over
the device table itself.

One of those asserts that every two-row device's names are covered by the shared-serial exemption,
so renaming a device can't silently break de-duplication.

The suite imports `calist` only, never `ui`, so it needs no display and no GUI toolkit.

### Project layout

```
calist.py                    the pipeline — imports no GUI toolkit
ui.py                        the desktop interface (customtkinter)
device_config.py             the 57 device layouts and the form() helper
test_calist.py               the test suite
template/Device List.xlsx    reference register template
```

The two halves meet in exactly two places: the pipeline emits log records that `ui.py` picks up with
a handler, and it returns structured `FileOutcome` / `RunResult` values that the table renders.

## Known issues

Four entries in the device table need checking against the paper forms:

- `EU` (Lab Oven) puts Location at `K19`; every other standard form puts it at `K20`.
- `CF` (Baby Warmer) has Model *below* Manufacturer — inverted compared with all 56 others.
- `CA` is identical to `BF`, and its name is unfinished in the source (`"X-ray ()"`).
- `AO` and `CK` are both named `Infrared` with different Status cells.

Three device names carry spelling slips that reach the register output — `Protien Analyzer`,
`Tornique` and `X-ray ()`. They're left as-is deliberately: correcting them changes the text written
into every historical register, so it should be a considered call rather than a drive-by fix.

Because the output lands among the source files, selecting the same folder twice will feed the
previous run's output back in. Its code resolves to `DEVICE`, which isn't in the table, so it's
skipped with an error rather than corrupting the register.

## License

No license is granted. The source is published for reference; all rights are reserved by the author.
If you'd like to use it, please get in touch.
