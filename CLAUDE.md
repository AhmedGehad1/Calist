# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

**Calist** — a single-window Tkinter desktop app that harvests fixed cells out of many medical-device
inspection Excel forms and compiles them into one flat equipment register.

## Commands

```powershell
python calist.py                # launch the app
python calist.py --inspect FORM # dump what each mapped cell of one form reads
pip install -r requirements.txt # openpyxl + xlrd + customtkinter

python -m pytest                            # the whole suite (419 tests)
python -m pytest test_calist.py             # one file
python -m pytest -k merged                  # one topic, by substring
python -m pytest test_calist.py::test_a_merged_cell_reads_through_to_its_anchor
python -c "import calist, sys; assert 'tkinter' not in sys.modules"   # GUI-free rule

pyinstaller calist.spec --noconfirm              # -> dist/Calist.exe (one file)
$env:CALIST_ONEDIR=1; pyinstaller calist.spec    # -> dist/Calist/    (folder)
python docs/make_icon.py                         # redraw the app icon
```

Five test files, all runnable without a display: `test_calist.py` (the pipeline),
`test_firebase_export.py` (the archive export), `test_access.py` (the PIN gate),
`test_settings.py` and `test_device_config.py` (the cell maps the Calystra app also writes).
There is no linter configured.

**Releases are built by CI, not locally.** Bump `__version__` in `calist.py`, then
`git tag vX.Y.Z && git push origin vX.Y.Z` — [release.yml](.github/workflows/release.yml) runs the
tests, builds both shapes, and publishes them. It **refuses a tag that disagrees with
`__version__`**, so bump first. Building locally is only for checking the spec; note that
PyInstaller refuses to run at all if the obsolete `pathlib` backport is installed
(`pip uninstall pathlib`).

`.xlsx`/`.xlsm` are read by Calist's own targeted reader (see *Reading a form* below); `xlrd` 2.x
reads **only** `.xls`. openpyxl is now a **write-side dependency only** — `write_output` is the one
place that calls `load_workbook`, though the reader still borrows its date and number-format
helpers so that values render exactly as they always did.

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
[`process_files()`](calist.py#L1165), which does no work itself:

1. [`extract_device_code()`](calist.py#L305) — filename stem, split on the first `-`, leading letters
   of the right-hand part. `"Clinic-AGH001.xlsx"` → `"AGH"`. The register's Code is the stem with
   its typing slips undone — see *Filenames are repaired, never renamed*.
2. `DEVICE_CONFIGS[code]["cells"]` — maps field names to A1 refs.
3. [`read_best()`](calist.py#L998) — asks [`_XlsxSource`](calist.py#L491) (or
   [`_XlsSource`](calist.py#L751) for `.xls`) for the whole cell map at once, falling through to
   `alt_cells` and then to the form's own printed labels when the map no longer fits — see *When
   the map no longer fits the form* below. Reads the **first non-empty worksheet** to begin with;
   see *Which sheet gets read*.
4. [`clean()`](calist.py#L203) — renders raw cell values as output strings.
5. [`build_second_row()`](calist.py#L861) — for configs with a `second_row` block, emits a sub-module
   row of the same physical unit.
6. [`sort_records()`](calist.py#L1022) → optional [`deduplicate_records()`](calist.py#L1038) →
   [`write_output()`](calist.py#L1112).

### Reading a form (the hot path — do not undo these)

Reading one `.xlsx` used to cost **~500 ms** (median 382 ms, worst 2.2 s): `load_workbook` parses
every worksheet, the whole `styles.xml`, the drawings and the calc chain in order to reach seven
cells. A 300-form round spent two and a half minutes doing it.

[`_XlsxSource`](calist.py#L491) goes at the package directly and costs **~2 ms** (mean 4.5 ms).
Measured against the old reader over every readable sample workbook: **82 workbooks, 171,200 cells
(15,518 carrying values), 0 differences**, with 9,843 of those reads going through a merged
non-anchor cell — 5,180 of them returning real text.

Five things there are load-bearing:

- **The lazy quantifier in `_CELL_RE`.** `<c r="E18"([^>]*?)(?:/>|>(.*?)</c>)`. Greedy, `[^>]*` eats
  the `/` of a self-closing `<c r="E18" s="168"/>`, takes the `>` branch, and swallows everything up
  to the *next* cell's `</c>` — so E18 silently returns F18's value. Pinned by
  `test_a_self_closing_cell_does_not_swallow_the_next_one`.
- **The `<c ` vs `<c r="` count gate.** All 625,586 cell tags in the sample corpus write `r` first,
  and the lookups rely on it. Two `bytes.count` calls (~0.02 ms) prove it per file instead of
  assuming it; a file that fails the gate raises rather than reading blank.
- **Anything the reader cannot make sense of raises.** There is no fallback reader, so a silent
  blank field — the exact failure `--inspect` exists to hunt — would be the alternative. A raise
  becomes a visible `ERROR` row via the per-file `try/except` in `extract_records`.
- **`sharedStrings.xml` and `styles.xml` are read lazily**, and only for the cells a map actually
  asks for. Scoping the styles trigger to the wanted cells rather than every cell on the sheet is
  most of the win: `styles.xml` is 184 KB on these forms and was being parsed for 20 files in 30.
- **No thread or process pool.** Measured twice: threads gave 1.2× against the old reader and
  **nothing** against the new one (1000 forms: 2.05 s at one thread, 2.11 s at two, 2.34 s at
  eight). After decompression the work is pure Python and GIL-bound, and at 2 ms a form the pool
  overhead exceeds the parse.

`.xls` goes through xlrd with `on_demand=True`, walking sheets one at a time — `workbook.sheets()`
would load every sheet and cancel the benefit. 58 ms → 21 ms median.

**How to prove a change to the reader.** The unit tests cover the shapes; what covers the *forms* is
a whole-grid diff against the previous reader. Before touching it, dump every cell in rows 1–80 ×
columns A–T of every workbook in a real folder through `calist._open_source` to JSON; after the
change, re-read and compare. That is what produced the 171,200-cell figure above, and it is the only
thing that catches a field going quietly blank — which is the failure mode this reader has already
had twice (the Ultrasound serial, the X-ray tab). `_open_source` is kept as a per-reference
compatibility seam for exactly this; `read_record` batches instead.

Where the time goes now, on 20,000 forms: **28 s read, 0.5 s sort, 0.2 s dedup, 9.6 s write.**
`write_output` is the long tail and there is no cheap fix — the stalls are inside
`workbook.save()`, so yielding the GIL around the fill loop does nothing (tried: worst stall
483 ms → 416 ms, and it cost a second on the write). `process_files` logs `Writing N row(s)…` and
the progress card says *Writing the register…* instead, which is the honest answer.

### Pre-flight

`classify_file(path, real_forms_only=False)` resolves a filename to a device **without opening the
workbook** — extension check, the optional real-device-forms check, code extraction, config lookup,
and the repaired name (`FileOutcome.code`). The UI runs it
on every file the moment it is added, which is how a bad name or an unrecognised code surfaces
before a long run instead of after it. It must stay I/O-free; a test asserts it works on a path that
does not exist.

### Filenames are repaired, never renamed (`repair_name`)

**By the owner's decision the files on disk are never renamed.** What is repaired is the name Calist
*uses* — the register's Code column, and the id the Firebase export gives a record. Measured over
the archive's 89,538 workbooks, the slips `repair_name` undoes are:

| Slip | Example | Repaired | Marked? |
|---|---|---|---|
| leading dot, quote, backtick, invisible mark | `.G414-CA002-0426` | `G414-CA002-0426` | no |
| spaces around a dash, before `.xlsx`; lower case | `G119-CE007 -0525 `, `k116-AX001-0123` | `G119-CE007-0525` | no |
| `Copy of ` / `Final ` in front | `Copy of B14-BB036-0126` | `B14-BB036-0126` | yes |
| a letter O for the zero of the customer code | `JO8-AGH001-1025` | `J08-AGH001-1025` | yes |
| anything after the date | `G302-BB001-0326-pending`, ` (2)`, ` FAIL` | `G302-BB001-0326` | yes (dots/dashes only: no) |
| a date that cannot be one | `G302-BB001-0329` (2029), `1525`, `00324` | from the form's Date: `-0326` | yes |

"Marked" is the amber cell and note of `_flags`, on **Code**, naming the real file — only for a
repair that changes what the name *says*; a space is not worth a person's time. A name that is not
the house shape at all (`GE  TEC850  OR  SQAX00871`, `5071938426`, `H58-BM010` with no date) is left
exactly as it is: there is nothing to repair it toward. `test_a_name_is_written_with_its_slips_undone`
holds every example. **Proof: on all 89,538 archive names, 0 of the 78,923 already-correct names
change.**

Rules that are load-bearing:

- **An impossible date is taken from the form, and the form is trusted** (owner's decision) — month
  and all, whenever its Date reads as a date. `repair_name` is I/O-free, so it only marks such a
  name `DATE_CHECK`; `settle_name_date` finishes it after the read. A form with no date leaves the
  typed digits, marked. "Impossible" means not four digits, month not 01–12, or a year after
  `latest_year` — this year in the app, the folder's year + 1 in the export (the old export rule).
- **A name dated in the future confirms nothing.** `repair_date` used the name's MMYY to fix a form's
  mistyped date; for `H39-AG025-0233` that would bend `01-02-20233` into 2033, so it now refuses.
- **`_LEAD` is a list, not "anything but a letter".** Arabic text in front of a name is a different
  kind of file, not a slip.
- **The device part must end where it seems to.** `H59-BZ00F-1025` is not the `000` template
  `BZ00` with an F after it; it is a name nobody can repair, left as it is.

**Copies** (`_settle_copies`, per run). Cutting text can land a name on a file already in the run:
`D38-AGH090-0225 (2)` beside `D38-AGH090-0225`. By the owner's decision **only the original is
kept — when it is the same device**:

- a copy whose serial matches the original's, or that has none, is dropped: status `COPY`, listed,
  and its module row goes with it;
- a copy with a **different** serial is a different device filed under a copy's name (`F21-BZ010-0624
  (2)` is pump 14007615, the original 14007302) — kept, marked;
- where **no** file is the plain original, the extra text is what tells them apart (`G341-AB00-0626-SEVO`
  and `-Iso` are two vaporizers; `-3.5 MHZ` / `-7.5 MHZ` two probes) — so none of them is cut.

In the archive that was 77 same-serial copies, 9 different-serial ones and 23 groups with no plain
original.

### The "real device forms only" switch (`real_forms_only`)

It **replaced the three-position filename check** (off / codes / full), by the owner's decision.
Off by default. On, anything that is not a device form is `LEFT_OUT` — hidden from the table, not a
problem, listed under Details with the reason, and never in the register:

- **no customer code** — the name must start (after the slips above) with **one letter then
  digits**: 834 of the 835 codes in Hospital Codes.xlsx are. `AG035-SMB176400052HA` is a device code
  and a serial, `DG169-EC001-0524` a typo nobody can repair; both are left out. ≈8,900 archive files;
- **device number all zeros** — `BZ000`, `AGH00`, `BP0` — the blank a round's forms are copied from,
  even when a serial was typed in (owner's decision). 172 files;
- **a device list** (`_is_device_list`, as before) and **a blank form**: nothing filled in on the
  record or any tab (`read_best(forms_only=True)` raises `LeftOut`). These need the workbook, so
  they drop out during the run rather than when added.

The name checks (`left_out_because`) are I/O-free and run in `classify_file`, **before** the device
lookup: the user asked for device forms, so a non-form is the finding, not its code. With the switch
off a blank form named for a device is read as it always was, a row of blanks.

Old settings files hold the replaced check as `name_check` (0–2) or `strict_names` (bool);
`ui.forms_only_setting` reads any non-off position as on. `_remember` **drops** both old keys rather
than writing them alongside — an older build reading `strict_names` would enforce the filename format
this switch replaced.

### In the Firebase export

The same repair and the same rule, by the owner's decision ("both"):

- **Ids come from the repaired name** whenever it carries a real date — `.G414-CA002-0426`, which used
  to be `x_2026_<hash of its path>`, is now `G414-CA002-0426`. `_adopt_repair` compares first, so a
  name the repair leaves alone keeps exactly the id and the attention it had
  (`test_a_name_that_needs_no_repair_keeps_exactly_what_it_had`). The old reading is kept as
  `ParsedForm.legacy` / `legacy_id`, and is what stands when an impossible date finds no date on the
  form.
- **The rule is always on in the export** (there is no switch there): `form.excluded` says why, and
  `exported()` is what reaches `assign_ids` and Firestore. `settle_copies` is `_settle_copies` for the
  archive, grouped by base id — and it only judges forms `deepen` actually opened, because the dry run
  reads a sample and an unread serial is not an empty one.
- **No existing record moves** (`assign_ids_stable`). Letting repaired files into the numbering as
  equals moved **333** records the repair never touched, measured: `G114-BP001-00324`, once its date
  was fixed, took `BP001` from the `G114-BP001-0223` that had held it for years, and a repaired file
  taking a fresh number first pushed every later renumbered device up by one. So `assign_ids` —
  unchanged, byte for byte — numbers **every scanned file's old identity**, excluded files included
  (they always took part), and every file the repair left alone keeps that id. Only then does
  `_fit_newcomers` place the files whose identity the repair created: the same device (site, tag,
  serial) joins its tag; anything else gets a number above every number in use. `deepen` still
  reads every known-device file, left out or not, so the serials numbering sees are the ones it
  always saw (`LeftOut` carries the record it read). Ordering tricks were tried first and are not
  enough; do not go back to them.
- **An engineer's deletion follows its file**: a tombstone on the legacy id also stops the repaired one.
- **`sweep_replaced` deletes what the run no longer produces** — the owner chose "replace" for repaired
  ids and "delete" for excluded files. A record goes only if the export made it (`imported`), its
  `sourcePath` is a file this run scanned (so `--year`/`--limit` cannot reach anything else), nothing
  was written under its id, and **nobody changed it in the app** (form data, tests, an uploader or
  engineer on it — kept and listed for a person). It streams the collection with `select`, like
  `upload_pending`, so no index is needed.
- **The dry run lists every candidate** in `dry-run-…-deletions.txt` (old id → new id, or old id and
  why). Only `--push --yes` deletes. Workbook copies in Storage are not deleted.

**Proof, against the whole archive** (87,409 exported files, old export vs new, serials from the
audit snapshot): **0 regressions** — all 79,821 files the repair leaves alone keep exactly their id;
273 repaired files get their own id; 7,315 are left out (6,767 no customer code, 330 copies, 172
templates, 46 blank forms). The register side, old reader vs new on 3,851 files (every renamed or
unnamed file plus a control sample): 0 regressions, 1,665 Codes repaired, 76 files read as the
device their certificate names that used to be skipped (97, less the 21 Carestations above). On the 116 folders where a cleaned name
collides: 87 copies dropped, 9 kept for a different serial, 55 kept with their text.

### Threading rules (both of these have already caused bugs)

Two things run on worker threads: the build (`App._start`) and the folder scan
(`App._start_scan`). Two hard rules, and they apply to both:

1. **Never read a Tk variable off the main thread.** `self._template.get()` inside the worker raises
   `RuntimeError: main thread is not in main loop`. `App._start` captures `template`, `deduplicate`,
   `forms_only`, `turbo` and the file list into plain Python values *before* spawning the thread;
   `_start_scan` does the same with `forms_only`. Keep it that way.
2. **Never touch a widget from the worker.** `on_file` pushes onto `App._events` (a `queue.Queue`);
   `App._drain` polls it on the main thread via `after()`. The scan pushes batches of 200 onto the
   same queue. `TkLogHandler.emit` does the same with its own queue. Batching there is also what
   keeps a few hundred forms from redrawing the table once per file.

`_drain` reschedules itself while **either** `self._cancel` or `self._scan` is set, so a scan and a
build both stay drained; `_cancel_run` stops whichever is in flight, and `_on_close` sets the scan's
event too, or its thread keeps walking a network share after the window is gone.

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

### Which sheet gets read (this has already lost a whole device)

The rule was **always `worksheets[0]`**, and it is right for almost every form. The X-ray workbook
is the exception: it opens on an empty `Waveform Dialog` stub left behind by its macros, with the
real form on the next tab. So every mapped cell of every X-ray read blank — silently, because a
missing value is indistinguishable from a form that was left empty.

`_XlsxSource._populated_sheet()` takes the **first sheet that holds anything at all**, and
`_sheet_has_content()` is deliberately strict: a sheet with any cell outside `A1`, or a value in
`A1`, is kept — cell *elements*, not values, so a tab carrying only formatting still counts. This
can only ever skip a tab that could not have held the data.

Two rules about *which* parts count as sheets, both easy to undo by accident:

- An older `.xlsm` lists its VBA modules as `<sheet r:id="">` entries with no part behind them.
  They are dropped, and must not shift the ordering.
- **Chartsheets are skipped; dialogsheets are not.** openpyxl counts dialogsheets in
  `wb.worksheets`, and the X-ray workbook opens on one — filtering by relationship type instead of
  by emptiness would pick a different tab than the app used to.

**Do not switch this to matching sheet names.** Across the real templates the data sheet is called
`Device data`, `Device Data`, `Data entry`, `Inserting data`, `Inserting Data` and `Data device`,
and several workbooks carry *both* a `Device data` and a `Data entry` tab with different layouts —
so a name list would pick the wrong one and would go stale on the next form revision.

Verified: the part-resolution rule agrees with `load_workbook(...).worksheets[0]` on all 82
readable sample workbooks, 0 disagreements.

### When the map no longer fits the form (`read_best`)

Measured over the 2025 and 2026 rounds — **42,826 forms** — the configured map works for **94.7%**
of them. The rest are not broken files: the form was re-laid-out between visits and everything below
the inserted row moved. Nothing announced it, because a shifted map reads blank and a blank field is
indistinguishable from one an engineer left empty.

[`read_best()`](calist.py#L998) tries, in order, and **stops at the first plausible record**:

1. the configured `cells` — where 94.7% stop, at no extra cost;
2. each `alt_cells` entry, the layouts already written down;
3. the fields located from **the form's own printed labels**, on the same sheet;
4. the same, across the workbook's other tabs.

Steps 3–4 are the net for shifts nobody has recorded yet. `alt_cells` still earns its place: it is
cheaper (no whole-grid read) and auditable.

**An alternate whose identity cells repeat the primary's can never be reached**, because an alternate
is only tried when the whole record is implausible — and if the primary's Model and S.N read, so do
its twin's. Three such entries existed (`BZ`, `AL`, `VAGH`), each written to name a *different Status
cell*, which is not something this mechanism can express. That is now `_settle_status`'s job.
`test_an_alternate_that_repeats_the_primary_identity_is_never_added` pins it.

**The configured map always wins when it produces a plausible record**, so this can only rescue a
file the map got wrong and can never change one it already got right.

Four things there are load-bearing:

- **`plausible()` needs the serial *and* the model.** A misaligned map landing on some other number
  passes a serial-only check; one landing on a caption passes a model-only check. It also rejects a
  record that is placeholders throughout — a cover sheet filled in with `0` has a serial that
  classifies as a *placeholder*, not a blank, and without that check the fallback "rescues" it into
  a row of zeroes.
- **A shift moves everything, so Date and Status move too.** They carry no label the locator can
  match — Status is captioned *above* its value, not beside it — so `_uniform_offset()` derives the
  distance from the four labelled fields and applies it to `Date`, `Status`, `Status2` and `S.N2`.
  It returns `None` unless every located field moved the same distance in the same column. Without
  this, `EP` wrote the caption `"Status:"` into the register instead of `"Calibrated"`.
- **On another sheet, carry nothing across.** The configured references describe a different tab, so
  reading them there returns whatever happens to sit at those coordinates — which is how a
  Nebulizer's date came back as `"Gas Flow Analyser"`. Only what was located on that sheet is used;
  the rest stay blank, because a blank field is honest and a confident wrong one is not.
- **A value that is itself a caption is not a value.** The Therapeutic Ultrasound form heads a table
  `Model | S.N.` at row 11 and puts the real fields at row 69; taking the first match read the model
  as the string `"S.N."`.

### The Status box moves on its own (`_settle_status`)

**`read_best` cannot fix a wrong Status, and this is the trap.** `plausible()` asks for a serial and
a model, so a map that places the identity block perfectly and misses the verdict box produces a
*plausible* record — the primary wins, no fallback is ever reached, and the miss is silent. Every
one of these was found that way, not by anything failing:

- `DG` read `H32`, a row the form does not reach: **0 of 953**.
- `AM` read `G33`, which serves the older template: 148 of 250 forms use `G31`.
- `AK` read the Model from `D72`, which is *Next calibration* — a **date in the Model column**.
- `BZ`, `AL`, `GC`, `FW`, `DU` had alternates pointing at column J/K, where `Safety:`,
  `Syringe brand:` and `Contact Person Name:` are printed. Those captions were the register's verdict
  on ~1,450 Phototherapy forms alone.

So `read_best` settles `Status` separately, after the layout search, on every route it can return by.
Three steps, in order:

1. The value **already reads as a status** → kept untouched. One set lookup, no extra read. This is
   why the 10 devices whose box never moves (`AGH`, `FJ`, `CF`, `AG`, `BC`, `CK`, `CZ`, `DE`, `FD`,
   `GI`) are completely unaffected, and why this can never change a file the map got right.
2. Otherwise the box is found from the form's **own printed `Status:` label** — `status_from_grid`.
3. Nothing found → the original text is kept, **unless it is a caption or an answer to a different
   question** (`_NOT_A_VERDICT`: `Small`, `Large`, `N.A`, `----`), **an engineer's code**
   (`_ENGINEER_CODE`: `JTE-`, `STE-18`) or **a bare number** (`_A_READING`: `191.2` joules, `5.1`),
   which is dropped to `""` — and kept, amber, in the register. On older Centrifuge forms the map's
   `K25` is the "Revised by" box, captioned *above* it where the caption check cannot see, and 45 rows
   read `JTE-` as a verdict. **`0` is not a reading**: see below.

Four things there are load-bearing:

- **Label-anchored, never an offset search.** Boxes move *sideways* more than they move down:
  measured over all 81 mapped Status cells, 33 devices have forms whose box is in a different
  column (`AO` `H32`→`A31`, `BM` `K22`→`J17`, `ED` `G43`→`J38`). Head to head on the forms where
  the mapped cell misses, the label found **230** that a same-column search missed, against **1**
  the other way — and where they disagreed, the offset search was returning the **legend**. These
  forms print `Pass | Fail | Limited non | limited fail` across columns **B–E on the answer's own
  row**, so every one of those cells reads as a status. Only starting from the label avoids them.
- **`_STATUS_BAND_ROWS = 18` is the whole cost.** `values()` scans the sheet once per reference, so
  the price is linear in the band. Reading the whole `A1:N95` grid to find one cell took a form from
  2.5 ms to 14 ms — a five-fold regression on the path the reader exists to keep fast. 18 rows
  recovers 2,499 of the 2,548 the whole grid finds, for a quarter of the cost. Only the 19.5% whose
  mapped cell missed ever pay, and 76% of those come back with a status.
- **A two-box device gets the caption guard and nothing else.** One printed `Status` label cannot say
  whether it belongs to the ECG or the NIBP, so filling a missing one from the nearest label would
  put one module's verdict against the other. `Status2` devices skip step 2 entirely.
- **A device with no Status box skips everything.** `CT`, `MRI`, `Dexa` and the air mattress map
  `Status` to `""` deliberately; with no anchor there is nothing to search around and no read to pay.

Proven the way this repo proves reader changes — whole-corpus, before and after, 3,004 forms:
**0 regressions**, 617 forms gained a status, 95 captions went to zero. That harness has since
become `tools/audit.py` — see *The archive audit* below.

**A status of `0` is an empty verdict, not a wrong one.** On Patient Monitor forms the status boxes
are formulas copying a test sheet (`='NIBP Test sheet'!I70`), and a copy of an empty cell shows `0`.
The owner has confirmed it: `0` means the module was not tested. It is never flagged, and the audit
counts it as missing, not wrong. (A lowercase `pass` beside the ECG box, identical on every form, is
template text — not a verdict to recover.)

### The caption check: a layout must be verified, not just plausible

`plausible()` asks whether values *look* right, and a shifted layout can pass it. On an older Baby
Warmer form the current map reads the model into Manufacturer, the ward into Model and `N.A` into
Serial — ordinary text, a placeholder serial, all "plausible". What a shifted layout cannot fake is
the **caption printed to the left of each box**. So `_best_layout` accepts a layout only when
`layout_verdict()` does not come back `contradicted`: the caption beside Model, Serial or
Manufacturer names a *different* field (`Location:` beside what the map calls Model).

**Only a contradiction counts.** A missing or unrecognised caption proves nothing and blocks
nothing, which is what keeps every caption-less form reading exactly as before. Measured on every
Baby Warmer form against its certificate: **191 corrected, 0 regressed**, and it picks the right
layout even within 2025, where both are in use.

**It costs about 0.4 ms a form** — 18 extra references (six cells left of three values), and a
reference with no cell element makes `values()` scan to the end of the sheet. Measured on a fixed
2,000-form archive sample: median 4.3 → 4.8 ms, total +7% with the Word and Conclusion fallbacks
included. If that ever matters, bound each lookup to its `<row>` element — a reader change, so it
needs the whole-grid proof described under *Reading a form*.

The same check chooses **per field** where one field moved on some forms only: `field_alternates`
in a config lists candidate cells, and `_pick_field_cells` takes the one whose caption names the
field (`EU` Location, `BZ` Date, `AS` Status, `BB` Date). A mapped cell whose caption names another
box outright loses even to an *empty* captioned box — Centrifuge's `K25` is "Tested by", so its
`JTE-40` is an engineer, and the honest answer is a blank Status. For a Date, a real date in any
candidate the form does not caption as something else beats a non-date: Ultrasound 2024 forms print
`Date of receipt` in F16, with F15 empty.

`plausible()` also refuses a record whose Model equals its Serial: a Phototherapy layout eight rows
down lands both on the literal text `D38`. And `classify_serial()` treats a **short decimal** —
one to three digits, then one or two after the point (`7.1`, `25.8`, `0.99`) — as a test reading,
not a serial. Keep it that narrow: Therapeutic Ultrasound serials look like `37.254`, `1873206.0` is a
float-exported one, and widening the rule to either made the reader reject correct layouts. A
**computed float** (`0.38271946501827364`: an old Defibrillator `.xls` opens on its test sheet, and the
map lands on an uncertainty column) is not a serial either — but only from **twelve** decimals: real
serials carry up to eight (`1742.60318475`; `437.20981` and `6.20417` on 180 forms).

**Placing a field that has no label.** When a form's block moves columns as well as rows there is
no uniform offset, so a field without a printed label cannot follow the others. `label_offsets` in
a config places it from a located one: Therapeutic Ultrasound's Date is two rows above its Model,
wherever that block sits (`{"Date": ("Model", -2)}`). Mapping that block's rows directly was tried
first and read test readings on 105 fields — the label search was already right about identity.

**Other tabs: ordinary ones first, certificates last.** The certificate repeats the device's
identity but never its verdict. Once `Equipment Data` became a recognised heading, the other-tab
search began stopping at the certificate before reaching the data tab it used to read in full —
and the verdict went with it. So `_best_layout` walks non-certificate tabs first, and within the
walk a placeholder serial loses to a real serial on a later tab: a Suction Unit's `cover report`
copies an empty box as `0.0`, and the certificate after it has the real number.

A rescue keeps the map's Date and Status from the form's own tab — when the rescuing tab has none
and they are unmistakably a date and a verdict — in two cases only: the rescue came from a
**certificate**, or the opening tab's **captions confirm the map** (`map_confirmed`), which is the
common case of a form whose only fault is a serial nobody entered; its identity then comes from the
cover page and its verdict from its own box — 408 fields across the archive. Not otherwise: when a
form opens on a test sheet, the map's Status cell there is a test line, and a Balance form once read
`Fail` from it (`test_a_rescue_from_a_data_tab_takes_its_verdict_from_that_tab`).

**The date is found by its own caption.** `locate_by_labels` places no Date, so every rescue from
another tab — and a same-tab labels read whose block moved columns — used to leave it blank:
`date_by_caption` takes the date beside `Date of receipt:` (a data tab), `Test Date` (a cover page)
or `Calib. Date:` (a certificate), in that order, and **never** `Issue Date`, `Last Cal. Date`,
`Prev. Calib.` or `Next Calib. On` — other dates, the last two a year away. It only ever fills a Date
that is blank or not a date, so it cannot change one that read; 612 dates filled, 0 changed.

### When no layout fits: the Word certificate, then the Conclusion

Some devices keep their identity only in a **same-named Word certificate** beside the workbook.
Hemodialysis "Final" workbooks are an electrical-safety printout, and older templates (Pacemaker,
Ultrasound, Heart-lung …) have a cover page that was never filled in. `read_best` reads it as a
**last resort**: the configured map, then alternates, then printed labels, then other tabs (a
filled cover page), and only then `read_word_certificate`. That is zero cost for normal forms.

Two rules there are load-bearing:

- **Read only after `Equipment Data`.** The certificate prints the calibrator's block first — `Dose
  meter Model: 07-492 … S.N: 108357` — the same trap as the Hemodialysis cover page above.
- **Standard library only.** A `.docx` is zipped XML; the body text of a Word 97 `.doc` sits in the
  file as plain 8-bit or UTF-16 runs, enough for fixed labels. 47 of 48 BD certificates are `.doc`.

No Word certificate carries a verdict. Hemodialysis states its verdict as a **Conclusion sentence**
on the `Final` tab, whose row moves (A119–A132), so `conclusion_tab` in its config anchors on the word
*Conclusion*. `verdict_of_sentence` knows the only three wordings in the archive: "…Passed all Test…"
→ Pass, "…Limited Calibrated, non calibrated item…" → Limited non, and "…the device readings and
there's no accepted range…" → no verdict (kept in the register, marked).

### Device lists are not forms

Device-List workbooks sit among the forms and are named like them: `Device List H23-AS-1023.xlsx`
resolves to a Centrifuge, `Al Arbaeen Device List.xlsx` to a Phototherapy (from "Al"). Reading one
writes whatever the map lands on. `_is_device_list` recognises a register by its heading row — only
looked for when the sheet or file is called a list, or the file is not named for a device (`Al
baeerat.xlsx` is a site's list, headed `Device name` and `SN`), so an ordinary form pays nothing — and
`read_best` raises `NotAForm`, which `extract_records` reports as *unsupported* (or *left out*, with
the real-device-forms switch on) and the Firebase export leaves out. 24 such files were being read as
devices.

### Files not named for a device (`has_device_name`)

A name without a site and device number — `G302-AGH001` anywhere in it, so `Copy of D41-AA002-0624`
still counts — was named some other way: by serial (`CN84017253.xlsx`, a Philips monitor), by brand
(`GE  Tec 850  SQAB01358  OR.xlsx`), or left a template (`BB.xlsx`, `000-AGH000-0000.xlsx`). Its
"code" is only its first letters, so a Philips serial read as a **Microwave** and a GE vaporizer as a
**Temperature Calibration Tester**. By the owner's decision, `_refuse_if_not_this_device` deals with
such a file in two cases, and only for these names (≈1 form in 100 pays):

- **its form says it is another device**: the value beside its own `Equipment Type:` / `Device
  Type:` names another mapped device and not this one (`names_device`: one's words contain the
  other's, with `Pulse Oximeter` ≡ `SPO2`, `Pipette` ≡ `Pipet`). When it names **exactly one** device
  the map knows, `ReadAsOther` hands it back and `extract_records` reads the file again **as that
  device** — the GE TEC850 files are Vaporizers — with Code kept as the file's own name, marked, and
  nothing invented. When it could be several (`Patient Monitor` is AG, AGH and VAGH) nothing is
  guessed: `NotAForm`, naming them. `BZ009.xlsx` whose certificate says *Syringe Pump* is a Syringe,
  and is read as one. **One exception, by the owner's decision** (`_stale_certificate`):
  *Ventilator* on a file filed under an **Anesthesia** folder is the template's text, not the
  device's — the 21 GE Carestations named by brand are anesthesia machines — so those are skipped
  as they always were. The 5 Carescape R860s filed under *Ventilator* are read as ventilators.

  **Only for names like these.** On a properly named file the certificate's type is not trusted over
  the name: across the archive it disagrees on 3,457 forms, and the models show the certificates are
  stale — the CT forms that say *Nebulizer* are Siemens Somatoms, the Anesthesia forms that say
  *Ventilator* are GE Avances.
- **nothing is filled in**: no real serial and no model or maker in words, on the record *or on any
  tab* — and on another tab it takes **two** of serial, model and maker, because a blank template's
  `safety` tab prints the safety analyzer's own serial beside `S.N`.

Cost: nothing on a named form (one regex on the name). On the others, `stated_device_type` looks at
certificate tabs first and for the caption in A–F only, then reads just its row — the first version
read A1:N60 of every tab and put 10% on a whole run; now it is ≈4%, median unchanged.

**A properly named form is never skipped, even blank** — unless the real-device-forms switch is on,
which leaves out a blank form whatever its name. Renaming the files was considered and
rejected: in the brand-named ones the file name is the only record of the device's identity (their
device tab is empty), and the Firebase document id is derived from the name.

### Typed-date repairs (the register only — the form is never changed)

By the owner's decision `_repair_date_field` repairs an engineer's slip in a Date that is **not a
date as typed**, and marks the cell with what the form says. `repair_date`:

- **A period** (`30/04/2023 - 18/06/2023`, `06-09-2023…11-09-2023`, `… to …`) — which the
  certificates print too, so not a slip — becomes its **start**.
- **Otherwise the file name decides.** Read as typed first, separators cleaned (`21--01-2024`,
  `11-06-2023.`, `15-092025.`): taken only if its month and year are the name's `-0124`. Only digits
  that cannot be a date as typed (month `20`, year `205`, `02025`, `1023`) are read against the name,
  one slipped digit per part at most. **Two readings, no repair**: `12/19/2022` is 19-12 month-first
  or 12-12 one slip. `21--03-2024` in a `-0124` file is plainly March, and is left alone rather than
  bent to fit. A name without the date gets no repair.

Files are also opened by **what they are**, not their extension (`_open_workbook` checks the first
bytes): the archive holds `.xlsx` workbooks saved under a `.xls` name, which failed outright.

### The register marks what it keeps but doubts (`_flags`)

**By the owner's decision, the register never drops a value silently.** A caption in the Status
box, a serial that is not one, a record no layout fits — each is written as found, the cell filled
amber, with an Excel note saying why (`_mark_for_checking`). The column layout is untouched.

This lives in an underscore side-channel, `record["_flags"]` — the same pattern as `_group` and
`_source`, which never reaches `FIELDS`. **The marks never reach the Firebase export**: it reads
only named fields (`record.get("S.N")` …), and `_settle_status` still blanks a caption in the
named field while recording the raw text in `_flags["Status"]["show"]` for the register alone.
The *fixes* do reach it, by the owner's decision — a corrected map, a repaired date, a status that
was a caption or a reading, a skipped file — which is why the audit proves every one of them.
`build_second_row` hands a module row its own Status2 mark, never the parent's Status mark.

### The archive audit (`tools/audit.py`)

The only thing that finds a field going quietly blank or wrong is reading the real forms, so the
check is a tool, not a one-off:

```powershell
python tools/audit.py "D:\MedCal Pro"                        # ~17 min, all four years
python tools/audit.py --report <snapshot>.json.gz            # rebuild the workbook
python tools/audit.py --diff <before>.json.gz <after>.json.gz  # the proof: 0 regressions
```

It reads every form the way the register does, then asks the form independently: the
**Certificate tab**, whose `Equipment Data` values are formulas naming the device-tab cell they copy
(Excel rewrites them when rows are inserted, so each certificate is a per-form cell map), the
device tab's printed labels, shape rules, and the same serial across years. Findings are MAP
ERROR, WRONG VALUE (systemic → a map problem, isolated → a form problem), MISSING, UNDETERMINED or
EXPECTED; the workbook has a clickable row per finding and a stratified sample to confirm by eye.

`--diff` re-judges **both** snapshots with today's rules, so only what the reader lands can differ.
A regression is a field that was OK and is not now, or whose confirmed value changed. A file now
refused as a device list is not one — it is listed by name instead, so each refusal is seen.

It is **read-only against the forms** and writes only to `<root>/_Calist audit`. That is enforced,
not assumed: `_save_outside_customers` refuses any destination inside a `Customers` folder. It
exists because a shadowed loop variable once saved the report over a customer's workbook.

Two lessons it has already taught, both worth keeping in mind for any map change:

- **Prove a cell change by value, on every form, `.xls` included.** Formula evidence alone once
  marked a Baby Warmer map change "safe" that would have broken every 2025–26 form — they are `.xls`,
  whose formulas cannot be read, so "no evidence" looked like "no breakage".
- **A swap that fixes many and breaks a few is not a fix.** `EU` and `BZ` would each have regressed
  a handful of forms; `field_alternates` fixes all of them.

### Never read the calibrator as the device

These sheets describe **two** instruments: the device being calibrated and the one doing the
calibrating. Both blocks print `Model:` and `Serial No.:`. A Hemodialysis cover page prints the
reference meter **first**:

```
A15  Calibration Device    G15 Model   H15 EMIS       <- the calibrator
A22  Device information
A24  Device name           G24 Model   H24 AK96       <- the dialysis machine
A26  Manufacturer  Gambro  G26 Serial  H26 11195
```

so "first `Model:` label wins" records `EMIS` / `9D2749` as the machine's. **Position cannot be
trusted; the headings can.** `locate_by_labels()` reads `_DEVICE_HEADING` and `_GEAR_HEADING` first
and only considers cells inside the device's own section. A form with no such headings — every
single-block form — behaves as before, first match top-down.

`test_the_calibrator_is_never_read_as_the_device` pins this. A register row carrying a Fluke's
serial looks right and is wrong, and nothing downstream would catch it.

### Merged cells (this has already lost a field)

The forms draw each answer as a **box spanning two columns**. Excel stores a merged range's value
only in its top-left cell; every other cell in the range reads as empty. So a cell map naming the
second column of a box — `L17` of a merged `K17:L17` — silently produced a blank field, with no
error anywhere. That is how the Ultrasound serial number disappeared when that form was re-laid-out
from the F/L columns onto E/K.

Both readers resolve it. `_XlsxSource._anchor()` finds the range covering a wanted reference and
redirects to its top-left cell; a reference that *is* the anchor, or that sits in no range at all,
is left alone, so an ordinary blank cell still reads blank and can never pick up the text of some
unrelated merged block it happens to sit inside.

The xlrd side needs `formatting_info=True` to see merges at all. It costs memory and some files
refuse it, so it falls back to a plain open — where `merged_cells` is empty and behaviour is
unchanged. xlrd ranges are 0-based with exclusive upper bounds; the `.xlsx` reader's are 1-based
inclusive.

**`--inspect` is how you check a cell map against a real form** without guessing:

```powershell
python calist.py --inspect "G302-BB001-0526.xlsx"
```

It prints every mapped field, the cell it reads, the value that comes back, the sheet name and the
merged ranges, and exits non-zero if anything read blank. Reach for it first whenever a field is
empty, and use it on the forms behind the open data questions below.

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

`ALLOWED_SHARED_SN_PAIRS` ([calist.py:71](calist.py#L71)) holds `device_name` strings verbatim from
`device_config.py`; renaming a device there breaks the exemption that lets a Patient Monitor and its
NIBP row share a serial. `test_second_row_names_are_covered_by_the_dedup_exemptions` guards this — run
the tests after renaming anything.

## Behaviour worth knowing

- An unknown device code **skips the file** with an error rather than emitting a junk row. Set
  `SKIP_UNKNOWN_CODES = False` ([calist.py:59](calist.py#L59)) to restore the old A1:A6 fallback.
- Output is `device list.xlsx` beside the first source file. `resolve_output_path()` refuses to run if
  that would overwrite the template (Windows paths are case-insensitive, so it would otherwise clobber
  `Device List.xlsx`).
- Because the output lands *among* the sources, selecting a whole folder twice would feed the
  previous run's output back in as an input. `is_source_file()` drops it by name during a scan; a
  copy picked by hand still resolves to `DEVICE`, which isn't in `DEVICE_CONFIGS`, so it is skipped
  with an error.
- **Excel lock files are skipped during a folder scan.** Excel drops a `~$`-prefixed copy beside any
  workbook someone has open; they are not workbooks, and the real sample tree holds seven of them.
  The filter applies to scanning only — a file picked *by hand* is still classified, because saying
  nothing at all about a file someone explicitly selected is worse than an "Unsupported format" row.
- Per-file failures are logged and skipped; a partial output is still written. Read the status log.

## Open data questions

None remain open from the list that used to be here. The 2026 archive audit settled both with the
forms' own certificates: **`EU` Location is `K20`** like every standard form (68 forms), with `K19`
kept as a caption-chosen candidate for the two B14 forms that print it there; and **`CF` is not
inverted** — it has two layouts two rows apart, both still in use in 2025, told apart per form by
the caption check. `CA` (was `"X-ray ()"`, now Dental X-Ray) and the duplicate `AO`/`CK` "Infrared"
naming were resolved earlier.

**`BD` (Mammography) is read only from its Word certificate — never from its workbook.** Every BD
workbook carries the identical header `GE / Alpha st / Gona Hospital`, survey date 2012, across many
site codes: the vendor QC template's boilerplate. Its "Model" caption belongs to the X-ray *tube*. A
cell map would put the same fabricated maker and model into every row and look entirely plausible,
which is why BD has `"source": "word"` and no cells at all. 48 of 53 have a certificate; the other 5
stay unread. `test_mammography_is_read_only_from_its_word_certificate` pins this.

`FC` (High Flow Nasal Cannula) is mapped: 128 of its 131 forms are a 2023 batch whose only device
block is on a tab called `cert`, reached by the other-tab label search. `BJ` (Auto Refractometer)
remains unmapped: one archive file, and no coherent field block.

Device names are reproduced verbatim in output, spelling slips included (`Protien Analyzer`,
`Tornique`). Correcting them changes the text written into every register, so treat it as a deliberate
data change, not a typo fix — the owner has been asked and has not said yes.

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
repeat run. `_enter_scanning` reuses the working card for a folder walk, which has no total to count
towards until it has finished walking.

### It runs on Windows 10 too

The development machine is Windows 11; the machines Calist is used on include Windows 10. A Win11-only
font or window attribute looks right here and fails **silently** there, and the GUI-free test suite
cannot catch it. Every visual choice needs its Windows 10 path:

- **Fonts:** `Segoe UI Variable` is Win11-only — fall back to `Segoe UI`. `Cascadia Mono` is not a
  Win10 system font (Windows Terminal's copy is private to it) — fall back to `Consolas`. Tk
  substitutes an unknown family without a word, so check `tkinter.font.families()` rather than
  assuming.
- **Icons:** `Segoe Fluent Icons` is Win11-only. `Segoe MDL2 Assets` ships on both, so use it
  everywhere and only codepoints it has — one icon set on every machine.
- **Window chrome:** DWM attributes added in Windows 11 (caption colour, corner preference, Mica)
  are best-effort extras that must never raise.
- **Screen size:** older Win10 machines often run 1366×768 at 100% scaling; check layouts there.
- **WebView2** is not guaranteed on an offline Win10 machine — worth remembering before any move
  to a web-rendered UI.

### Turbo

A round switch beside the wordmark, which turns red and grows a flame (`TurboFlame`, canvas
polygons — **Pillow is a dev-only dependency and must not become a runtime one**). It exists for
runs of tens of thousands, where the per-file UI work is the cost rather than the reading:

- **No table.** `_refresh_table` and `_update_row` return early, and the table's rows are *deleted*
  rather than merely hidden, so they stop costing memory for the session. `_build_summary`'s panel
  takes the same grid slot.
- **No per-file event.** `on_file` enqueues only problems, the last file, and a tick at most every
  100 ms. Forty thousand queue items and forty thousand row updates are what used to stop the window
  breathing.
- **Log problems only** (`process_files(quiet=True)`): failures, unknown codes and duplicate serials
  in full, successes collapsed to a heartbeat every `HEARTBEAT_EVERY` files. 20,000 forms produce
  about a thousand log lines instead of twenty thousand.
- The summary renders from `RunResult` alone — counts, unknown codes grouped by code, duplicate
  serials, and the files that failed. Nothing is scraped back out of the log, which is why
  `deduplicate_records` also collects `Duplicate` records instead of only warning.

Turbo persists in `settings.json` and is captured into a plain bool by `_start` before the worker
begins (threading rule 1). `_on_run_done` writes the outcomes back into `self._files` in Turbo too —
they are stored by reference, so it costs a dict slot each, and without it switching Turbo off after
a run would show a half-stale table.

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
  `MIN_ROWS`/`MAX_ROWS` and `MIN_HERO_H`/`MAX_HERO_H`, and capped by `MAX_FIT_PASSES` so a layout
  that cannot land on an exact fit stops rescheduling itself. Call `_on_resize()` (debounced) after
  anything that changes the page's shape — `_show_action`, `_toggle_log` and `_refresh_intake` all
  do.
- **Two unit systems meet in the hero branch.** `CTkFrame.configure(height=)` and `cget("height")`
  speak CustomTkinter's *logical* units; `slack` and every `winfo_*` measurement are *device*
  pixels, 1.25× apart on this display. Comparing one against the other made a 380 cap render as
  475 and stopped the loop settling. Convert with
  `ctk.ScalingTracker.get_widget_scaling(widget)` and keep each side in its own units. The table
  branch is exempt: a Treeview's `height` is a row count and `ROW_PX` is the raw ttk `rowheight`,
  so both are already device pixels.
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
  hides the table; once devices are in, the slim bar takes over and the table appears (or, in Turbo,
  the summary panel).
- **Folders are scanned on a worker.** `calist.find_source_files()` walks with `os.scandir`, whose
  entries already know file-from-directory — `Path.rglob("*")` plus `is_file()` cost a stat per
  entry, and all of it ran on the main thread, which on a network share is a window that stops
  repainting with no way to stop it. Results arrive in batches of 200 through `App._events`, the
  count updates live, and Cancel works. 20,000 forms resolve in under a second with the window
  still live.
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
- **"Real device forms only"** is the switch beside dedup. A `LEFT_OUT` file is kept in `_files` (so
  switching off brings it back) but never drawn: `_visible_outcomes` skips it, the counts exclude it,
  `_start` does not hand it to the run, and `_log_left_out` lists every one with its reason under
  Details — once per intake or toggle, not per refresh. A dropped copy stays in the table as
  *Copy — left out*.
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

`__version__` in [calist.py](calist.py#L86) is the single source of the version number. The spec
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
