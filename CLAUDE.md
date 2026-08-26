# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

**Calist** — a single-window Tkinter desktop app that harvests fixed cells out of many medical-device
inspection Excel forms and compiles them into one flat equipment register.

## Commands

```powershell
python calist.py                # launch the app
python -m pytest                # run the test suite (105 tests)
python -c "import calist"       # pipeline import check — pulls in no GUI
pip install -r requirements.txt # openpyxl + xlrd + customtkinter

pyinstaller calist.spec --noconfirm              # -> dist/Calist.exe (one file)
$env:CALIST_ONEDIR=1; pyinstaller calist.spec    # -> dist/Calist/    (folder)
python docs/make_icon.py                         # redraw the app icon
```

`xlrd` 2.x reads **only** `.xls`; `.xlsx`/`.xlsm` go through `openpyxl`.

## Architecture

Four modules, strictly one direction:

```
ui.py  ──imports──>  calist.py  ──imports──>  device_config.py
(customtkinter)      (pipeline)               (pure data)
   │
   └──imports──>  access.py     (daily PIN gate — pure, standalone)
```

`access.py` imports nothing from the rest of the app, so the gate is testable without a display and
cannot be broken by a pipeline change.

**`calist.py` must never import a GUI toolkit.** `main()` imports `ui` lazily inside the function
body, so `python calist.py` still launches the app while `import calist` stays GUI-free — which is
what lets the test suite run without a display. There is a test-adjacent check for this:
`python -c "import calist, sys; assert 'tkinter' not in sys.modules"`.

The two meet in two places, and nowhere else:

- **Logging** — the pipeline emits `log.info/warning/error`; `ui.TkLogHandler` routes records into
  the details drawer.
- **Structured results** — `FileOutcome` and `RunResult` carry the same facts in a form the table
  renders. Neither channel replaces the other; keep both fed.

The pipeline is a chain of small functions orchestrated by
[`process_files()`](calist.py#L462), which does no work itself:

1. [`extract_device_code()`](calist.py#L188) — filename stem, split on the first `-`, leading letters
   of the right-hand part. `"Clinic-AGH001.xlsx"` → `"AGH"`.
2. `DEVICE_CONFIGS[code]["cells"]` — maps field names to A1 refs.
3. [`read_record()`](calist.py#L243) — reads those cells via [`_open_source()`](calist.py#L211), a
   context manager that hides the openpyxl/xlrd split behind one `get(ref)` function. **Always
   worksheet index 0.**
4. [`clean()`](calist.py#L142) — renders raw cell values as output strings.
5. [`build_second_row()`](calist.py#L256) — for configs with a `second_row` block, emits a sub-module
   row of the same physical unit.
6. [`sort_records()`](calist.py#L380) → optional [`deduplicate_records()`](calist.py#L386) →
   [`write_output()`](calist.py#L442).

### Pre-flight

`classify_file(path, strict_names=False)` resolves a filename to a device **without opening the
workbook** — extension check, optional format check, code extraction, config lookup. The UI runs it
on every file the moment it is added, which is how a bad name or an unrecognised code surfaces
before a long run instead of after it. It must stay I/O-free; a test asserts it works on a path that
does not exist.

### Filename format check (`strict_names`)

Optional, off by default, surfaced as a switch next to the dedup one. Enforces
`G302-AGH001-0425` — site code (letters then digits), device code and number, then MMYY with the
month range-checked.

The performance shape is deliberate and worth preserving:

- `check_filename_format()` is **one precompiled `_NAME_RE.match()`** on the accepting path — no
  splitting, no allocation. ~0.7 µs, so 300 names re-validate in ~0.2 ms and the table can refresh
  on the same click that flips the switch.
- `_explain_bad_filename()` does the per-part diagnosis and is **only reached for names that already
  failed**. A correctly named folder never pays for it. Keep it that way — moving the diagnosis onto
  the hot path would make toggling feel sluggish on a large folder.

With the switch on, the format is checked **before** the device lookup: the user has asked for that
shape specifically, so a malformed name is the finding worth reporting even when a device code could
still be salvaged. `test_format_is_checked_before_the_device_code` pins this.

### Threading rules (both of these have already caused bugs)

The pipeline runs on a worker thread. Two hard rules:

1. **Never read a Tk variable off the main thread.** `self._template.get()` inside the worker raises
   `RuntimeError: main thread is not in main loop`. `App._start` captures `template`, `deduplicate`
   and the file list into plain Python values *before* spawning the thread — keep it that way.
2. **Never touch a widget from the worker.** `on_file` pushes onto `App._events` (a `queue.Queue`);
   `App._drain` polls it on the main thread via `after()`. `TkLogHandler.emit` does the same with
   its own queue. Batching there is also what keeps a few hundred forms from redrawing the table
   once per file.

`widget.after(0, ...)` called from a worker thread is the common shortcut and is *not* safe — it
registers a Tcl command from outside the main loop. Use the queues.

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

Duplicate warnings name the two **files**, not the device type — a round holding a dozen of the same
model makes the type useless for finding the form to open. `source_name()` reads `_source`, which is
set in `extract_records` and survives into a generated sub-module row; `Code` cannot serve because
`build_second_row` rewrites its device token, so a sub-module's Code names no file on disk.

`ALLOWED_SHARED_SN_PAIRS` ([calist.py:60](calist.py#L60)) holds `device_name` strings verbatim from
`device_config.py`; renaming a device there breaks the exemption that lets a Patient Monitor and its
NIBP row share a serial. `test_second_row_names_are_covered_by_the_dedup_exemptions` guards this — run
the tests after renaming anything.

## Behaviour worth knowing

- An unknown device code **skips the file** with an error rather than emitting a junk row. Set
  `SKIP_UNKNOWN_CODES = False` ([calist.py:48](calist.py#L48)) to restore the old A1:A6 fallback.
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

## Daily PIN gate (`access.py`)

The app is locked until the day's four-digit code is entered. The code is derived from the local
date, so nothing is stored and nothing is distributed:

```
base = day*17 + month*31 + year*11 + SECRET_KEY   (8374)
pin  = (base * base // 10) % 10000, zero-padded
```

**The `// 10` is not cosmetic.** Taking the last four digits of a square straight off reaches only
1044 of the 10000 codes — the final digits of a square cannot be arbitrary — and left 38 codes in
2026 repeating on a second date. Shifting one digit widens that to 5784 and halves the repeats to
19. `test_the_shifted_digit_is_what_widens_the_keyspace` and
`test_codes_rarely_repeat_within_a_year` both fail if it is dropped.

Unlock state rides in the same `settings.json` the UI already writes, as `unlocked_on` holding an
ISO date. That gives the midnight reset for free — a date that is not today means locked, and so
does a **missing** key, which is why deleting the settings file locks the app rather than opening
it. `test_missing_state_is_locked_not_open` pins that direction.

### Never withdraw the CTk root before the first mainloop

The lock is a **panel inside the main window** (`LockPanel`, shown by `App.show_lock()`), not a
Toplevel over a hidden root. That is not a style choice — the obvious version is broken.

CustomTkinter's `CTk.withdraw()` sets `_withdraw_called_before_window_exists` when the window has
never been shown. `CTk.mainloop()` then takes its first-show branch, which calls
`_windows_set_titlebar_color()`; with `_window_exists` still False that hides the window and does
not restore it (the saved state is `None`), and the same flag stops mainloop calling `deiconify()`.
Result: the process runs with a permanently invisible window. It looks exactly like the app opening
and instantly closing.

So: one root, created once, never withdrawn. `show_lock()` grid-removes the main widgets (removing
them from the tab order too, not merely covering them) and `_on_unlocked()` puts back exactly what
it took away. `run()` is three lines and calls no window-state methods at all.

Three more things worth not undoing:

- **Rate limiting is load-bearing.** Even shifted, 5784 of 10000 values are reachable, so unlimited
  guessing would still fall eventually. After 5 failures a cooldown starts at 30s and doubles,
  capped at 15 min, and it is **persisted** — closing the dialog must not shed it, and a correct
  code entered during a cooldown is still refused.
- **`App.__init__` schedules the day-watcher, it does not call it.** Calling it there raises a second
  lock prompt behind the one `run()` puts up.
- **The watcher never re-locks mid-run** (`self._cancel is None`). Taking the window away during a
  build would lose the user's work and protect nothing, since the run was authorised that morning.

Limits, so nobody mistakes this for more than it is: the formula is in the source and the repo is
public, and the system clock is the only authority on the date. It raises the bar for casual use; it
is not a licensing system.

## Attribution

`AUTHOR_NAME` / `AUTHOR_EMAIL` / `ATTRIBUTION` in [calist.py](calist.py). `stamp_attribution()`
writes a signature line one blank row below the last record and sets the workbook's document
properties, so credit travels with the register rather than living only in the app. This is the one
place output deliberately differs from the pre-lock builds — earlier work verified register output
byte-identical twice, so a diff here is expected, not a regression.

## `ui.py`

One window, three states swapped in the same layout by `_enter_setup` / `_enter_working` /
`_enter_results`. Not a wizard — this is a tool the same person runs repeatedly, and steps tax every
repeat run.

### Everything lives on one scrollable page

`App._build` puts every block inside a `CTkScrollableFrame` (`self._page`). This replaced a fixed
window grid where the table row carried the only weight — so every pixel the results card or the
details drawer needed came straight out of the device list, which collapsed to a sliver once a run
finished with the drawer open. Three things follow from the change, and all three are load-bearing:

- **Nothing stretches on its own.** A scrollable page is exactly as tall as its contents, so
  `grid_rowconfigure(..., weight=1)` inside it does nothing. `_fit_to_window` is what hands spare
  window height to whichever block can use it — the hero while the app is empty, the table once it
  is not — by converting the slack into whole Treeview rows (`ROW_PX`) or hero pixels. It is
  incremental and self-correcting rather than computing exact chrome heights, clamped by
  `MIN_ROWS`/`MAX_ROWS` and `MIN_HERO_H`, and capped by `MAX_FIT_PASSES` so a layout that cannot
  land on an exact fit stops rescheduling itself. Call `_on_resize()` (debounced) after anything
  that changes the page's shape — `_show_action`, `_toggle_log` and `_refresh_intake` all do.
- **The hero needs `grid_propagate(False)` and an explicit height.** Left to size itself it collapses
  to its contents, and the empty state is meant to fill the window.
- **`CTkScrollableFrame` steals the mouse wheel.** It binds `<MouseWheel>` with `bind_all` and
  scrolls itself for any event whose widget chain reaches its canvas — which is every widget on the
  page, including the device table and the log box. `_wheel_over()` puts a widget-level binding on
  each of those that scrolls the widget and returns `"break"`; widget bindings run before `all`
  bindings, so the scroll stays where the pointer is. Add any future scrollable widget to that list.

`show_lock()` now grid-removes the whole page rather than a list of individual widgets, and root row
0 keeps `weight=1` permanently so the lock panel fills the window in its place.

- **Adding devices is the hero.** With nothing loaded, `_refresh_intake` shows the hero panel and
  hides the table; once devices are in, the slim bar takes over and the table appears.
- **The destination is always on screen, and now selectable.** The *Saves to* row shows where the
  register will land *before* the run, and warns when it would replace an existing file.
  `shorten_path()` elides the middle of long paths, never the tail — the deepest folders and
  filename are what the user reads. **Select folder** sets `output_dir`, which `_start` captures on
  the main thread (threading rule 1) and passes to `process_files`; with none set the register lands
  beside the first source file exactly as it always did. `_remember` drops a stored folder that no
  longer exists, so a deleted or unmounted destination cannot fail the next run.
- **The table carries the serial number**, between Device type and Status. It is blank (`—`) until
  the run actually opens the file — `classify_file` is I/O-free and resolves a device from the
  filename alone, so it cannot know a serial. `FileOutcome.serial` is filled in `extract_records`,
  where a dual-serial device's two numbers are joined onto one line for the table.
- **Live per-row status is the anti-frozen signal**, more than the progress bar — users watch their
  own filenames resolve.
- Inputs are frozen during a run (`_set_inputs_enabled`), so the settings can't describe a build
  other than the one happening.
- The `ttk.Treeview` is styled to match CTk (`style_treeview`). It stays a Treeview rather than
  stacked CTk frames because it routinely holds hundreds of rows.
- Drag-and-drop is optional: `HAS_DND` gates a `TkinterDnD.DnDWrapper` mixin on the root. Absent the
  package, the drop zone is click-only and nothing else changes. **Untested** — `tkinterdnd2` is not
  installed here.
- Template, last folder and the dedup switch persist to `%APPDATA%\Calist\settings.json`; both read
  and write are best-effort and must never raise.

## Packaging (`calist.spec`) and the antivirus problem

Windows Defender flagged the v1.1.x download as **`Trojan:Win32/Sabsik.FL.A!ml`**. The `!ml` suffix
means a machine-learning verdict rather than a signature match — a false positive, and a well-known
one for PyInstaller. Three properties of that build drove it. Undoing any of the first two brings it
back:

1. **A one-file build self-extracts.** It unpacks into `%TEMP%\_MEIxxxx` at launch and executes from
   there, which is what packed malware does and is the heaviest single signal. So the spec builds
   **two shapes**, and both are published: `pyinstaller calist.spec` gives the one-file
   `dist/Calist.exe`, and `CALIST_ONEDIR=1 pyinstaller calist.spec` gives the folder
   `dist/Calist/`, which is zipped as `Calist-windows.zip` and is the one that gets through.
2. **The executable carried no version resource at all** — every field empty. The spec now generates
   one, ASCII-only (Explorer, Task Manager and the SmartScreen prompt read it back, and a non-ASCII
   byte renders as mojibake there). CI fails the build if `ProductName` or `CompanyName` come out
   empty, because that regression is invisible until users are already being blocked.
3. **It is unsigned.** Only a certificate fixes that one, and there isn't one. Say so plainly rather
   than implying the other two measures are a complete fix.

Also load-bearing: **`upx=False`**. UPX-packing an unsigned binary is one of the strongest heuristic
signals there is, and it only saves a few MB.

`__version__` in [calist.py](calist.py#L78) is the single source of the version number. The spec
reads it with a regex rather than importing the module, so a build never depends on the app's
runtime imports resolving; `CALIST_VERSION` overrides it for CI. The release workflow **refuses to
build a tag that disagrees with it**, so the Properties tab can't claim a different version from the
release it came from.

Asset names (`Calist.exe`, `Calist-windows.zip`) must stay stable across releases — the
`releases/latest/download/<name>` permalinks in the README are built from them.

## The icon

[docs/make_icon.py](docs/make_icon.py) draws it; run it to regenerate (needs Pillow, which is a dev
dependency only and deliberately not in `requirements.txt`). Two things there are deliberate:

- **Every size is drawn at its own geometry, not downsampled from one master.** The four-row list
  reads at 256px and turns to mush at 16px, so 16/24/32 drop to three rows with much thicker
  strokes. `Image.save` would re-render every entry from the base image, throwing that away — the
  finished renditions go in through `append_images` instead.
- **The tile has a lifted rim.** The fill is near-black and so is the Windows 11 taskbar; without a
  lighter edge the icon dissolves into its own background.

`ui.app_icon()` resolves it the same way `calist.bundled_template()` does, which is why the spec
keeps the `docs/` prefix when bundling. `App._apply_icon` applies it **twice** — CustomTkinter
finishes setting the window up on its first mainloop pass, and that re-show drops an icon assigned
during `__init__`, reverting to Tk's default feather.

## Repo

Public GitHub repo at `AhmedGehad1/Calist`, default branch `main`. CI runs the test suite on
windows-latest across Python 3.10–3.12 ([.github/workflows/tests.yml](.github/workflows/tests.yml)).
There is deliberately **no LICENSE file** — all rights reserved.

Never commit completed inspection forms; they carry real device and site data. `.gitignore` covers the
common patterns plus Calist's own `device list.xlsx` output.
