# Product

<!-- impeccable:product-schema 1 -->

## Platform

windows-desktop

A native Windows desktop app — Python 3.10–3.12, CustomTkinter 6 on Tk 8.6 — shipped
as a PyInstaller build (one file, and a folder zip). Not a website: web-only guidance
(CSS, responsive breakpoints, browser APIs) does not apply.

## Users

Biomedical engineers who inspect and calibrate medical equipment on hospital and
clinic rounds. Every device inspected gets its own Excel inspection form. After a
round, the engineer points Calist at the round's folder to compile those forms into
one equipment register. They know the forms, the device codes and the filename rules
well, run the app repeatedly, and want speed and dense information, not hand-holding.

A typical run is 100–1,000 forms (one hospital round). Archive runs of tens of
thousands happen, rarely.

## Product Purpose

Turn a folder of inspection forms into one clean, sorted, de-duplicated equipment
register (`device list.xlsx`, filled into the `Device List.xlsx` template) without
opening a file by hand. Success is a register that is right — no field quietly blank
or wrong — and a round whose problem files are found and fixed quickly.

## Positioning

Calist already knows the cell layout of 95 device types' forms, recognises each
form from its filename's device code before opening anything, and finds a field by
the form's own printed captions when a form was re-laid-out. Where it doubts a value
it keeps it and marks it for checking in the register instead of dropping it.

## Operating Context

- Machines: mostly laptops at 1366×768 and 100% scaling; Windows 10 and Windows 11;
  some offline. Every visual choice needs a Windows 10 path (fonts, icons, window
  chrome).
- Files: a round folder, often with subfolders per ward or department, on a local
  disk or a network share. Filenames follow `SITE-DEVICE###-MMYY`, e.g.
  `G302-AGH001-0425`.
- The register is opened in Excel afterwards; values Calist doubts are filled amber
  with a note.
- A daily four-digit access code unlocks the app each morning.
- The workflow is repeated: add a round, check the problem files, build, open the
  register. No wizard; every extra step taxes every repeat run.

## Capabilities and Constraints

- Add a folder (scanned on a worker thread) or single files; every file is checked
  from its name the moment it is added — unknown device code, unsupported file,
  not a device form — before any run.
- Build the register: live per-file status, progress, cancel; results with counts;
  open the register or show it in its folder.
- Settings: the register template, the save folder, remove duplicate serial numbers,
  real device forms only, light or dark.
- Two-row devices (a monitor and its NIBP module) produce a module row.
- Turbo: a lighter mode for runs of tens of thousands (no per-file table).
- Terminology users know: round, form, register, device code, site code, serial
  number (S.N), module row, copy, left out, unknown code.
- Constraints: English only. Pillow is not available at runtime (no image files
  beyond the icon). Icons come from the Segoe MDL2 Assets font. The device table is
  a native Tk Treeview because it holds hundreds of rows. The UI must never block
  on file I/O.

## Brand Commitments

- Name: Calist. App icon: a dark tile with a four-row list (`docs/calist-icon.png`).
- Credit: "Built by Ahmed Gehad" appears in the app (footer and About) and every
  register is signed with the author's name and email.
- README tagline: "An afternoon of Excel, done in six seconds."

## Evidence on Hand

- README figures: 95 device types; ~6 seconds for a round; measured on 42,826 real
  forms from the 2025 and 2026 rounds.
- No testimonials, customer names or screenshots of real forms may be shown:
  inspection forms carry real device and site data, and the repository is public.
  Screenshots use invented demo data only.

## Product Principles

1. Never quietly wrong. Anything doubtful is shown, marked or listed — never hidden.
2. Problems first. The files that need fixing are the first thing the screen helps
   you find, and the fix is one step away.
3. Fast on repeat. The same person runs this every round; no step that is not needed.
4. Works on every engineer's machine: Windows 10, small laptop screens, offline.
5. Speak the engineers' language: device codes, serials, rounds — dense, not chatty.

## Accessibility & Inclusion

WCAG 2 AA contrast in both themes (enforced by `test_theme.py`). Status is never
carried by colour alone: every status has a glyph and words. Full keyboard use for
the main flow, with visible focus.
