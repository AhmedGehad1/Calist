"""Calist's look, as data: colours, type and icons.

Pure data and arithmetic — no tkinter, no customtkinter — so the test suite can
check every colour pair for contrast and every icon for existence without a
display, and `import theme` is as GUI-free as `import calist`.

Colours are (light, dark) pairs, the shape CustomTkinter switches between on
its own when the appearance mode changes. Tk-native widgets (the device table,
its scrollbar, the flame canvas) cannot take a pair; ui.py resolves those for
the current mode and re-applies them on every switch.

Everything here has to hold on Windows 10 as well as 11 (see CLAUDE.md): fonts
come as candidate lists that fall back to what Windows 10 ships, and icons are
drawn only from Segoe MDL2 Assets, which both carry.
"""

from __future__ import annotations

from dataclasses import dataclass

Colour = tuple[str, str]            # (light, dark)

LIGHT, DARK = 0, 1

# ──────────────────────────────────────────────────────────────────────────────
# Colour
# ──────────────────────────────────────────────────────────────────────────────

#: Every token a palette must define. The contrast tests walk this, so a token
#: added here without a pair in a palette fails loudly rather than rendering in
#: CustomTkinter's default blue.
TOKENS = (
    "bg",               # the window
    "surface",          # cards
    "surface_2",        # the table, inputs, secondary buttons
    "chrome",           # the footer bar and the settings drawer
    "row_alt",          # every other table row
    "border",
    "text",
    "muted",            # secondary text: labels, counts
    "faint",            # hints and disabled text — the one text colour below AA
    "accent",           # the primary action
    "accent_hover",
    "on_accent",        # text on the primary action
    "selection",        # a selected table row
    "highlight",        # the one field a direction gives its signature colour
    "on_highlight",
    "success",
    "warning",          # warning text
    "caution",          # a warning dot or glyph: brighter than text may be
    "warn_tint",        # the background of a row that needs a look
    "danger",
    "danger_tint",      # the background of a row that could not be read
    "turbo",
    "turbo_hover",
)

#: Clinical precision. The graphite of an equipment room and the teal of
#: theatre scrubs; the verdict colours carry the meaning and nothing else is
#: loud. The primary action is teal with dark text on it in dark mode — white
#: on blue is exactly the button 1.x shipped, and it failed contrast (3.2:1).
CLINICAL: dict[str, Colour] = {
    "bg":           ("#eef3f3", "#121719"),
    "surface":      ("#ffffff", "#192024"),
    "surface_2":    ("#f4f8f8", "#1f272c"),
    "chrome":       ("#ffffff", "#161c1f"),
    "row_alt":      ("#eef4f4", "#232c31"),
    "border":       ("#d2dddd", "#2c373d"),
    "text":         ("#132123", "#e4ecee"),
    "muted":        ("#4a5d61", "#9cabb1"),
    "faint":        ("#6f8286", "#7b8b92"),
    "accent":       ("#0b7a6e", "#3cc2b0"),
    "accent_hover": ("#086358", "#5ad3c2"),
    "on_accent":    ("#ffffff", "#04201c"),
    "selection":    ("#cdebe6", "#1d3b39"),
    "highlight":    ("#d9f1ed", "#1b3a37"),
    "on_highlight": ("#0b4d45", "#bff0e8"),
    "success":      ("#1b7f43", "#4cc781"),
    "warning":      ("#8f5f00", "#e6b04a"),
    "caution":      ("#b87400", "#e6b04a"),
    "warn_tint":    ("#fcf3df", "#2b2717"),
    "danger":       ("#c0302b", "#f06d6d"),
    "danger_tint":  ("#fbe7e6", "#2f1f21"),
    "turbo":        ("#c2451c", "#ff7a4d"),
    "turbo_hover":  ("#a53a17", "#ff9670"),
}

def pick(colour: Colour | str, dark: bool) -> str:
    """One side of a pair, for a widget that cannot take both."""
    if isinstance(colour, str):
        return colour
    return colour[DARK if dark else LIGHT]


def _channel(value: int) -> float:
    c = value / 255
    return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4


def luminance(hex_colour: str) -> float:
    """WCAG 2 relative luminance of a #rrggbb colour."""
    h = hex_colour.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * _channel(r) + 0.7152 * _channel(g) + 0.0722 * _channel(b)


def contrast_ratio(a: str, b: str) -> float:
    """WCAG 2 contrast ratio, 1.0 (none) to 21.0 (black on white)."""
    la, lb = sorted((luminance(a), luminance(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


#: The pairs the contrast tests hold every palette to, in both modes:
#: (foreground token, background tokens, minimum ratio). 4.5 is WCAG AA for
#: body text; 3.0 is AA for large text and for anything that is not text — a
#: status dot, an icon, the edge of the primary button. `faint` is held only
#: to 3.0 because it is reserved for hints and disabled labels, which WCAG
#: exempts; nothing a user must read to act is ever drawn in it.
CONTRAST_RULES = (
    ("text", ("bg", "surface", "surface_2", "chrome", "row_alt"), 4.5),
    ("muted", ("bg", "surface", "surface_2", "chrome", "row_alt"), 4.5),
    ("faint", ("bg", "surface", "surface_2", "chrome", "row_alt"), 3.0),
    ("on_accent", ("accent", "accent_hover"), 4.5),
    ("accent", ("bg", "surface", "chrome"), 3.0),
    ("on_highlight", ("highlight",), 4.5),
    ("success", ("surface_2", "row_alt", "chrome"), 3.0),
    ("warning", ("bg", "surface", "surface_2", "row_alt", "chrome"), 3.0),
    ("danger", ("bg", "surface", "surface_2", "row_alt", "chrome"), 3.0),
    ("text", ("selection", "warn_tint", "danger_tint"), 4.5),
    ("muted", ("warn_tint", "danger_tint"), 4.5),
    ("caution", ("surface", "row_alt", "warn_tint"), 3.0),
    ("danger", ("danger_tint",), 3.0),
)


# ──────────────────────────────────────────────────────────────────────────────
# Type
# ──────────────────────────────────────────────────────────────────────────────

#: Candidates, best first. The first of each exists only on Windows 11; the
#: fallback is what a Windows 10 machine has. Tk substitutes an unknown family
#: without a word, which is why these are chosen from the installed list
#: rather than assumed.
BODY_FONTS = ("Segoe UI Variable Text", "Segoe UI")
DISPLAY_FONTS = ("Segoe UI Variable Display", "Segoe UI")
MONO_FONTS = ("Cascadia Mono", "Consolas")
ICON_FONTS = ("Segoe MDL2 Assets",)

#: What each list falls back to when nothing in it is installed.
FALLBACK_BODY = "Segoe UI"
FALLBACK_MONO = "Consolas"


def pick_family(available, candidates, fallback: str | None = None) -> str:
    """The first candidate that is installed, else `fallback`, else the last.

    `available` is whatever tkinter.font.families() returned — passed in, so
    this stays testable without a display.
    """
    installed = {name.lower() for name in available}
    for name in candidates:
        if name.lower() in installed:
            return name
    return fallback if fallback is not None else candidates[-1]


# ──────────────────────────────────────────────────────────────────────────────
# Icons
# ──────────────────────────────────────────────────────────────────────────────

#: Glyphs from Segoe MDL2 Assets — a font on every Windows 10 and 11 machine,
#: so icons cost no image files, no Pillow, and stay sharp at any scaling.
#: Segoe Fluent Icons (Windows 11 only) is deliberately not used. The test
#: suite checks each of these is really in the font.
ICONS = {
    "add": "",
    "cancel": "",
    "settings": "",
    "search": "",
    "forward": "",
    "refresh": "",
    "lock": "",
    "check": "",
    "backspace": "",
    "chevron_right": "",
    "warning": "",
    "error": "",
    "page": "",
    "folder_open": "",
    "clear": "",
    "open_file": "",
    "folder": "",
    "info": "",
    "bolt": "",
    "sun": "",
    "moon": "",
    "filter": "",
    "copy": "",
}


# ──────────────────────────────────────────────────────────────────────────────
# The look
# ──────────────────────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class Direction:
    """The look in one place: palette, type and shape."""

    name: str
    palette: dict
    #: Font candidates, best first, each ending on a Windows 10 font.
    body: tuple
    display: tuple
    mono: tuple
    #: Corner radius of cards, and of buttons and inputs.
    radius: int
    control_radius: int


#: Clinical precision — chosen by the owner over "Certificate" and "Bold" from
#: real screenshots of all three, in both themes (Gate 2, September 2026).
ACTIVE = Direction("Clinical precision", CLINICAL, BODY_FONTS, DISPLAY_FONTS,
                   MONO_FONTS, radius=10, control_radius=7)

#: The palette the app draws with.
PALETTE: dict[str, Colour] = ACTIVE.palette
