---
name: Calist
description: A pre-flight checklist for a round of inspection forms, in clinical graphite and scrub teal.
colors:
  bg-dark: "#121719"
  bg-light: "#eef3f3"
  surface-dark: "#192024"
  surface-light: "#ffffff"
  surface-2-dark: "#1f272c"
  surface-2-light: "#f4f8f8"
  chrome-dark: "#161c1f"
  chrome-light: "#ffffff"
  row-alt-dark: "#232c31"
  row-alt-light: "#eef4f4"
  border-dark: "#2c373d"
  border-light: "#d2dddd"
  text-dark: "#e4ecee"
  text-light: "#132123"
  muted-dark: "#9cabb1"
  muted-light: "#4a5d61"
  faint-dark: "#7b8b92"
  faint-light: "#6f8286"
  accent-dark: "#3cc2b0"
  accent-light: "#0b7a6e"
  accent-hover-dark: "#5ad3c2"
  accent-hover-light: "#086358"
  on-accent-dark: "#04201c"
  on-accent-light: "#ffffff"
  selection-dark: "#1d3b39"
  selection-light: "#cdebe6"
  highlight-dark: "#1b3a37"
  highlight-light: "#d9f1ed"
  on-highlight-dark: "#bff0e8"
  on-highlight-light: "#0b4d45"
  success-dark: "#4cc781"
  success-light: "#1b7f43"
  warning-dark: "#e6b04a"
  warning-light: "#8f5f00"
  caution-dark: "#e6b04a"
  caution-light: "#b87400"
  warn-tint-dark: "#2b2717"
  warn-tint-light: "#fcf3df"
  danger-dark: "#f06d6d"
  danger-light: "#c0302b"
  danger-tint-dark: "#2f1f21"
  danger-tint-light: "#fbe7e6"
  turbo-dark: "#ff7a4d"
  turbo-light: "#c2451c"
typography:
  display-lock:
    fontFamily: "Segoe UI Variable Display, Segoe UI"
    fontSize: "24px"
    fontWeight: 700
  display-empty:
    fontFamily: "Segoe UI Variable Display, Segoe UI"
    fontSize: "22px"
    fontWeight: 700
  wordmark:
    fontFamily: "Segoe UI Variable Display, Segoe UI"
    fontSize: "20px"
    fontWeight: 700
  title:
    fontFamily: "Segoe UI Variable Display, Segoe UI"
    fontSize: "19px"
    fontWeight: 700
  verdict:
    fontFamily: "Segoe UI Variable Display, Segoe UI"
    fontSize: "17px"
    fontWeight: 700
  action-large:
    fontFamily: "Segoe UI Variable Display, Segoe UI"
    fontSize: "14px"
    fontWeight: 700
  round-name:
    fontFamily: "Segoe UI Variable Text, Segoe UI"
    fontSize: "15px"
    fontWeight: 400
  body:
    fontFamily: "Segoe UI Variable Text, Segoe UI"
    fontSize: "13px"
    fontWeight: 400
  body-strong:
    fontFamily: "Segoe UI Variable Text, Segoe UI"
    fontSize: "13px"
    fontWeight: 700
  label:
    fontFamily: "Segoe UI Variable Text, Segoe UI"
    fontSize: "12px"
    fontWeight: 400
  credit:
    fontFamily: "Segoe UI Variable Text, Segoe UI"
    fontSize: "11px"
    fontWeight: 400
  table:
    fontFamily: "Segoe UI Variable Text, Segoe UI"
    fontSize: "10pt"
    fontWeight: 400
  table-heading:
    fontFamily: "Segoe UI Variable Text, Segoe UI"
    fontSize: "9pt"
    fontWeight: 400
  log:
    fontFamily: "Cascadia Mono, Consolas"
    fontSize: "11px"
    fontWeight: 400
  icon:
    fontFamily: "Segoe MDL2 Assets"
    fontSize: "14px"
rounded:
  check: "3px"
  control-inner: "5px"
  control: "7px"
  card: "10px"
spacing:
  xs: "4px"
  sm: "6px"
  md: "8px"
  lg: "12px"
  xl: "16px"
  drawer: "22px"
  page: "24px"
components:
  button-primary:
    backgroundColor: "{colors.accent-dark}"
    textColor: "{colors.on-accent-dark}"
    typography: "{typography.body-strong}"
    rounded: "{rounded.control}"
    height: "36px"
  button-primary-hover:
    backgroundColor: "{colors.accent-hover-dark}"
    textColor: "{colors.on-accent-dark}"
  button-primary-footer:
    backgroundColor: "{colors.accent-dark}"
    textColor: "{colors.on-accent-dark}"
    typography: "{typography.action-large}"
    rounded: "{rounded.control}"
    width: "180px"
    height: "40px"
  button-secondary:
    backgroundColor: "{colors.surface-2-dark}"
    textColor: "{colors.text-dark}"
    typography: "{typography.body}"
    rounded: "{rounded.control}"
    height: "36px"
  button-secondary-hover:
    backgroundColor: "{colors.border-dark}"
    textColor: "{colors.text-dark}"
  button-ghost:
    textColor: "{colors.muted-dark}"
    typography: "{typography.body}"
    rounded: "{rounded.control}"
    height: "36px"
  button-ghost-hover:
    backgroundColor: "{colors.surface-2-dark}"
    textColor: "{colors.muted-dark}"
  button-disabled:
    backgroundColor: "{colors.surface-2-dark}"
    textColor: "{colors.faint-dark}"
    rounded: "{rounded.control}"
  button-icon:
    textColor: "{colors.muted-dark}"
    typography: "{typography.icon}"
    rounded: "{rounded.control}"
    size: "32px"
  chip:
    backgroundColor: "{colors.surface-dark}"
    textColor: "{colors.text-dark}"
    typography: "{typography.body}"
    rounded: "{rounded.control}"
    height: "28px"
  chip-picked:
    backgroundColor: "{colors.highlight-dark}"
    textColor: "{colors.on-highlight-dark}"
    typography: "{typography.body}"
    rounded: "{rounded.control}"
    height: "28px"
  toggle:
    backgroundColor: "{colors.surface-2-dark}"
    textColor: "{colors.muted-dark}"
    typography: "{typography.label}"
    rounded: "{rounded.control}"
    height: "32px"
  toggle-selected:
    backgroundColor: "{colors.accent-dark}"
    textColor: "{colors.on-accent-dark}"
    rounded: "{rounded.control-inner}"
    width: "76px"
    height: "24px"
  checkbox:
    backgroundColor: "{colors.accent-dark}"
    textColor: "{colors.on-accent-dark}"
    rounded: "{rounded.check}"
    size: "20px"
  input-search:
    backgroundColor: "{colors.surface-2-dark}"
    textColor: "{colors.text-dark}"
    typography: "{typography.label}"
    rounded: "{rounded.control}"
    width: "290px"
    height: "30px"
  card:
    backgroundColor: "{colors.surface-dark}"
    rounded: "{rounded.card}"
  table-row:
    backgroundColor: "{colors.surface-dark}"
    textColor: "{colors.text-dark}"
    typography: "{typography.table}"
    height: "26px"
  table-row-stripe:
    backgroundColor: "{colors.row-alt-dark}"
  table-row-selected:
    backgroundColor: "{colors.selection-dark}"
    textColor: "{colors.text-dark}"
  table-row-warn:
    backgroundColor: "{colors.warn-tint-dark}"
    textColor: "{colors.text-dark}"
  table-row-error:
    backgroundColor: "{colors.danger-tint-dark}"
    textColor: "{colors.text-dark}"
  table-row-muted:
    textColor: "{colors.faint-dark}"
  footer:
    backgroundColor: "{colors.chrome-dark}"
    padding: "12px 24px"
  settings-drawer:
    backgroundColor: "{colors.chrome-dark}"
    padding: "18px 22px"
    width: "380px"
  keypad-key:
    backgroundColor: "{colors.surface-2-dark}"
    textColor: "{colors.text-dark}"
    typography: "{typography.title}"
    rounded: "{rounded.control}"
    width: "76px"
    height: "52px"
---

# Design System: Calist

## Overview

**Creative North Star: "The Equipment-Room Checklist"**

Calist is a pre-flight checklist for a round of inspection forms: problems first, the build one click away. The world is the graphite of an equipment room and the teal of theatre scrubs. Graphite does all the structural work, in a few close tonal steps from window to card to input; teal marks the action to take and the thing currently chosen; green, amber and red carry verdicts and nothing else. The screen is dense and quiet on purpose, because the same engineer runs it every round and reads filenames, device codes and serials, not decoration.

It is a native Windows 10 and 11 desktop app (CustomTkinter 6 on Tk 8.6), not a web page. Every value here is a CustomTkinter or ttk setting in `theme.py` and `ui.py`, and every colour exists as a light and a dark pair; dark is what a fresh install opens in. Depth is carried by 1px hairlines and tonal steps, never by shadows or gradients. One family (Segoe UI Variable, falling back to Segoe UI on Windows 10) and one icon font (Segoe MDL2 Assets, on both) keep every machine drawing the same window.

The window refuses the category default of a settings form over a log, and the card-grid dashboard. It is one page: header, toolbar, a status line of counts and problem chips, the device table, and a fixed footer that always says where the register will be saved.

**Key Characteristics:**
- Graphite surfaces in tonal steps; 1px hairlines instead of shadows.
- Teal reserved for the primary action and the current choice.
- Every status told three ways: colour, a dot or MDL2 glyph, and words.
- WCAG 2 AA in both themes, enforced by `test_theme.py` against `CONTRAST_RULES`.
- One sans family and one icon font, each with a Windows 10 path.
- Dense: 26px table rows, 28–36px controls, a 1366×768 laptop shows 15+ forms.
- One animation: the settings drawer's 160ms ease-out slide.

## Colors

A cool graphite neutral ramp with one teal accent and three verdict hues, all defined as (light, dark) pairs in `theme.CLINICAL`.

### Primary
- **Theatre-Scrub Teal** (`accent-dark` / `accent-light`): the fill of the primary action (Add folder, Build register, Open register), the selected segment of a Toggle, a checked checkbox, the edge of the focused search box, the progress bar and filled PIN dots. Dark mode puts near-black text on it (`on-accent-dark`); white on a mid blue was the 1.x button and failed contrast at 3.2:1.
- **Deep Scrub Teal** (`accent-hover-*`): hover of the primary action and of a selected Toggle segment.
- **Selection Wash** (`selection-*`): the background of a selected table row; text stays `text`.
- **Picked Mint** (`highlight-*` with `on-highlight-*`): a problem-group chip that is the table's current filter.

### Tertiary (verdicts)
- **Pass Green** (`success-*`): the "read" count dot, the check glyph beside "Register built" and "Nothing needs a look", the ok row dot.
- **Amber** (`warning-*` for text, `caution-*` for dots and glyphs): a row that needs a look, the "need a look" glyph, a Cancelled verdict, a destination that will be replaced, the lock cooldown message. Caution may be brighter than text may be.
- **Amber Wash** (`warn-tint-*`): the background of a row that needs a look.
- **Alarm Red** (`danger-*`): a row that could not be read, a Not built verdict, a wrong PIN, and the hover of Cancel.
- **Red Wash** (`danger-tint-*`): the background of a row that could not be read.
- **Turbo Orange** (`turbo-*`): only the MDL2 bolt that says Turbo is on, in the status line and in the working footer.

### Neutral
- **Equipment-Room Graphite** (`bg-*`): the window and the lock screen.
- **Card** (`surface-*`): the empty card, the table card, the Details drawer, the lock card, plain chips.
- **Recessed** (`surface-2-*`): inputs, secondary buttons, the Toggle track, keypad keys, the disabled button fill, scrollbar thumbs of the page.
- **Chrome** (`chrome-*`): the fixed footer and the settings drawer.
- **Stripe** (`row-alt-*`): every other plain table row.
- **Hairline** (`border-*`): every 1px border and rule, the table scrollbar thumb.
- **Ink** (`text-*`), **Muted** (`muted-*`, labels, counts, ghost buttons, icon buttons), **Faint** (`faint-*`, hints, placeholders, disabled labels, the credit line, dimmed rows).

### Named Rules
**The Scrub Rule.** Teal marks what to press and what is chosen: the primary action, the selected segment, the checked box, the picked chip, the selected row, the focused field, progress. It is never a heading colour, a decoration, or a status.

**The Three Witnesses Rule.** A state is never colour alone. Every table row has a status dot and status words; every verdict has an MDL2 glyph and words; every count dot has a count and a noun.

**The Faint Floor Rule.** `faint` is the one text colour held only to 3:1. It is for hints, placeholders, disabled labels and the credit; nothing a user must read to act is drawn in it.

**The Pair Rule.** A colour is added to `TOKENS` and to the palette as a (light, dark) pair with its contrast pair in `CONTRAST_RULES`, or it is not added. Tk-native widgets (the table, its tags, hairlines) take the resolved side via `live()` and are re-applied on every theme switch.

## Typography

**Display Font:** Segoe UI Variable Display (with Segoe UI on Windows 10), bold
**Body Font:** Segoe UI Variable Text (with Segoe UI on Windows 10)
**Label/Mono Font:** Cascadia Mono (with Consolas on Windows 10), the Details log only
**Icons:** Segoe MDL2 Assets, on Windows 10 and 11

**Character:** The platform's own UI face, in one family with a bolder optical size for the few names and verdicts; it reads as a Windows tool, not a web page in a frame. Families are chosen at startup from `tkinter.font.families()`, because Tk substitutes a missing family silently.

### Hierarchy
- **Display** (700, 22–24px): the lock card's "Calist" (24) and the empty card's "Add a round of inspection forms" (22). Once per screen.
- **Wordmark** (700, 20px): "Calist" at the header's left.
- **Title** (700, 19px): the drawer's "Settings"; the same size carries keypad digits.
- **Verdict** (700, 17px): "Register built", "Not built", "Cancelled", beside a 22px glyph.
- **Large action** (700, 14–15px display face): the footer's Build register and Open register (14), the empty card's Add folder (15).
- **Round name** (400, 15px): the round beside the wordmark, after a vertical hairline.
- **Body** (400 or 700, 13px): button labels (primary bold), the counts line ("30 forms" and "N need a look" bold), drawer values, checkbox labels.
- **Label** (400, 12px): the footer, chips' secondary lines, drawer section labels and explanations, search text, Toggle segments, filter note.
- **Credit** (400, 11px): "Built by Ahmed Gehad" in the footer.
- **Table** (400, 10pt body; 9pt headings in `muted`): ttk sizes in points, so they track Windows scaling differently from the CustomTkinter pixel sizes above.
- **Log** (mono 11px, `muted`): the Details drawer, no wrap.

### Named Rules
**The One Family Rule.** One sans family, two optical sizes, bold only for names, counts that lead, verdicts and the primary action. The mono face is for the log and nowhere else.

**The Windows 10 Path Rule.** Every family is a candidate list whose last entry ships with Windows 10 (`BODY_FONTS`, `DISPLAY_FONTS`, `MONO_FONTS`); icons come only from Segoe MDL2 Assets, and only codepoints in `theme.ICONS`, which the tests check exist in the font. Segoe Fluent Icons is never used.

## Layout

One window, one scrollable page, blocks stacked on a single column with a 24px gutter (`spacing.page`) and 8px between blocks. Minimum window 1000×600; designed and checked at 1366×768 100%, also at 1920×1080 125%.

- **Header:** wordmark alone, a vertical hairline, the round name; Details (ghost) and the settings icon button at the right.
- **Toolbar:** Add folder (primary), Add files, Re-check (secondary), Clear all (ghost) at the left with 6px gaps; the search box and the All/Problems Toggle at the right. All controls 34px.
- **Status line:** counts with 8px dots, 16px apart; the amber warning glyph and "N need a look"; up to four problem-group chips 28px high, 6px apart, the rest folded into one "N more" chip; at the right end, what the table is filtered to and a ghost "Show all".
- **Table card:** takes all spare height. `_fit_to_window` converts slack into whole 26px rows, between 5 and 40.
- **Empty state:** the table and toolbar are replaced by one card that absorbs the spare height (300–1400), its stack centred.
- **Footer:** pinned to the window, not the page, above a hairline; 12px vertical and 24px horizontal padding; one line in every state (idle: Saves to, path, hint, credit, Build register 180×40; working: stage, current file, ETA, a 6px progress bar, Cancel; results: verdict, breakdown, saved path, Open register, Show in folder, Start over).
- **Settings drawer:** 380px, over the page from the right edge, 22px inner padding, sections separated by hairlines.
- **Density:** Turbo (1,000 forms or more) removes the per-file table rows and the All filter; the status line lists problems only.

## Elevation & Depth

Flat. No shadows and no gradients anywhere; CustomTkinter draws neither and none is faked. Depth is tonal: graphite window, a slightly lighter card, a recessed input tone, and a separate chrome tone for the footer and drawer, each edged with a 1px `border` hairline. The drawer reads as a layer only through its own chrome fill and a full-height vertical hairline on its left edge.

### Named Rules
**The Hairline Rule.** Separation is a 1px `border` line or a tonal step, never a shadow, glow or gradient.

## Shapes

Two radii. Cards (empty card, table card, Details drawer, lock card) round at 10px; every control (buttons, chips, search box, Toggle track, keypad keys) at 7px. Nested shapes step down: a Toggle segment is 5px inside its 7px track, a checkbox 3px. Dots are true circles: 8px count dots, 14px PIN dots, and table status dots drawn pixel by pixel into a PhotoImage at 30% of the row height, so they stay round at any scaling. The footer and drawer are square to the window edge.

## Components

### Buttons
Quiet and exact; the same action looks the same everywhere because every button is one of five roles (`BUTTON_ROLES`).
- **Shape:** 7px corners; 36px by default, 34px in the toolbar, 40px in the footer, 44px in the empty card, 52px on the keypad.
- **Primary:** teal fill, `on-accent` text, bold. Add folder, Build register, Open register.
- **Secondary:** recessed fill, `text`, 1px hairline edge; hover turns to the border tone. Cancel alone hovers red.
- **Ghost:** no fill, `muted` text, hover to the recessed tone. Details, Clear all, Start over, Show all, keypad Clear.
- **Icon button:** a ghost button holding one MDL2 glyph, 32–34px square (24px inside the search box).
- **Focus:** a 2px ring in the `text` colour, shown only when focus arrives by keyboard; Enter and Space press (`keyboard_ready`).
- **Disabled:** recessed fill and `faint` label, and out of the Tab order (`set_enabled`); a disabled primary never keeps its teal.

### Chips
- **Style:** a card-toned button with a hairline edge, 28px, label then the count after three spaces ("Unknown code QQQ   2").
- **State:** picked turns to the mint `highlight` fill with `on-highlight` text and filters the table to exactly those files; the status line's right end names the filter and offers Show all.

### Toggle
Mutually exclusive choices (All/Problems, Dark/Light): a recessed 32px track with a hairline, segments 76px wide; the selected one filled teal, the others `muted` text on no fill.

### Checkbox
On/off settings as a 20px box with a 2px `muted` edge, teal fill and `on-accent` check when on, a 12px `muted` explanation indented under the label.

### Inputs / Fields
- **Style:** the search box is a recessed 7px field with a hairline edge, an MDL2 search glyph in `faint`, 12px text and a `faint` placeholder.
- **Focus:** the edge turns teal.
- **Clear:** a 24px icon button that is always laid out and only shown when there is text, so the box never shifts.

### Cards / Containers
- **Corner Style:** 10px.
- **Background:** `surface`.
- **Shadow Strategy:** none; see Elevation & Depth.
- **Border:** 1px `border`.
- **Internal Padding:** the table sits 8px in at the left, 6px at the top; the lock card 32–40px.

### Device Table (signature)
A native ttk Treeview styled flat: `surface` background, 10pt rows 26px high, 9pt `muted` headings, no borders. Columns File, Device, Serial number (an em dash until read), Status, plus a narrow first column holding only the status dot. Plain rows stripe with `row-alt`; a row needing a look takes the amber wash and a row that could not be read the red wash, and a tinted row is never also striped; copies and left-out files are dimmed to `faint`. Selection is the teal wash. A 12px CTk scrollbar sits inside the card.

### Status Line and Verdict
Counts first ("30 forms" bold, "26 ready" with a faint dot, "24 read" with a green dot, "N left out" with a faint dot), then the amber MDL2 warning glyph and bold "N need a look", then the chips. After a build the footer's left end shows the verdict: a 22px MDL2 glyph in the verdict colour and 17px bold words.

### Lock Screen
The only place the app's mark appears inside the window: a centred card on the graphite window holding the 48px mark (drawn per scaling, 48/60/72/96), "Calist" at 24px, "Enter today's access code", four 14px PIN dots that fill teal, a 3×4 keypad of 76×52 secondary keys (Clear as a ghost word, backspace as its MDL2 glyph), and the hint and email beneath.

### Motion
The settings drawer slides in and out over 160ms in eight 20ms frames with a cubic ease-out. Nothing else animates without a state change.

## Do's and Don'ts

### Do:
- **Do** show the wordmark alone in the header; the app's mark appears only on the lock screen and as the window and taskbar icon (owner's decision).
- **Do** use the MDL2 open-folder glyph (40px, `muted`) on the empty card, not the app's mark (owner's decision).
- **Do** pair every status colour with a dot or MDL2 glyph and words.
- **Do** add every new colour as a (light, dark) pair with a `CONTRAST_RULES` entry, and resolve it with `live()` for Tk-native widgets.
- **Do** build every button with `make_button` in one of the five roles, and toggle it with `set_enabled`.
- **Do** keep the footer one line in every state, pinned to the window, with the save path always visible.
- **Do** keep cards at 10px and controls at 7px, edged with a 1px hairline.

### Don't:
- **Don't** use a CTkSwitch for an on/off setting: its knob is larger than its track and vanishes into the drawer in one mode or the other. Use a checkbox.
- **Don't** use a CTkSegmentedButton: it takes one text colour for every segment. Use `Toggle`.
- **Don't** draw shadows, glows or gradients.
- **Don't** use teal for headings, decoration or a status.
- **Don't** draw anything a user must read in `faint`.
- **Don't** use a Windows 11-only font or Segoe Fluent Icons; every family ends on a Windows 10 fallback.
- **Don't** stand a text character in for an icon; glyphs come from `theme.ICONS`.
- **Don't** animate anything but the settings drawer.
