---
version: 1
slug: "ui-py"
primary_target: "ui.py"
related_targets: []
---

# Surface brief: Calist main window (ui.py)

Scope: the single Calist window — empty state, loaded round, problems, build in
progress, results, Turbo, settings drawer, Details log, lock screen.
Mode: Operate. Native Windows desktop app (Python + CustomTkinter 6 on Tk 8.6),
Windows 10 and 11, mostly 1366×768 laptops at 100%. English only.

Audience and job: biomedical engineers compiling a round of 100–1,000 inspection
forms into one equipment register, repeatedly. Task: add a round, find and fix the
problem files, build, open the register.

Constraints: no Pillow at runtime (icons are Segoe MDL2 Assets glyphs; images are
PhotoImage-drawn); the device table is a ttk Treeview; status never by colour alone;
WCAG AA in both themes (test_theme.py); Turbo (no per-file table) from 1,000 forms.

Unresolved: none in the window itself. Register-file wording ("Compiled from N
row(s)") is the owner's call and outside this surface.

## Direction contract

THESIS: A pre-flight checklist for a round: problems first, the build one click
away. It refuses the category default of a settings form over a log, and the
card-grid dashboard.

OWN-WORLD: Equipment-room graphite (#121719 dark / #eef3f3 light) with
theatre-scrub teal (#3cc2b0 / #0b7a6e) reserved for the primary action, the
current selection, and live state (progress, the entered PIN, a focused field). Verdict colours green/amber/red, always paired with a dot or
glyph and words. One family, Segoe UI Variable (Segoe UI on Windows 10); MDL2
icons; 10 px cards, 7 px controls; 1 px hairlines; no shadows, no gradients.

STORY: The engineer adds the round and sees at once how many forms, how many need
a look, and which device codes are wrong; picks one problem group, fixes those
files, builds, and opens the register — never hunting for where it was saved.

FIRST VIEWPORT: Loaded: header (wordmark, round name, Details, settings gear);
toolbar with teal "Add folder", Add files, Re-check and Clear all left, search and
All/Problems right; a status line with the round's counts, an amber warning glyph,
"N need a look" and one chip per problem group; the device table fills the
rest; a fixed one-line footer with "Saves to <path>" left and a teal 180×40
"Build register" right. Empty: one card, folder glyph, "Add a round of inspection forms",
primary "Add folder".

FORM: Owner-pinned direction A, Clinical precision. No concept-seed roll (a
pinned direction beats the roll), so no seed key. Evidence of the pick: the owner
chose A through the structured question at Gate 2, 26 Sep 2026, after real
screenshots of all three shortlisted directions in both themes
(docs/review/gate2/GATE-2-A.png, -B.png, -D.png, local review material; B and D
were then deleted from the code). Signature interaction: a problem-group chip
filters the table to exactly those files, names the filter at the line's right
end, and offers Show all. Motion: the settings drawer slides in 160 ms, ease-out;
nothing else animates without a state change.

FINISH: unreviewed and undocumented is unfinished; this build ends with the finish review, the verdict, DESIGN.md, and every shipping raster carrying its provenance
