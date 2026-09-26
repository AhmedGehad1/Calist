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
    "border",
    "text",
    "muted",            # secondary text: labels, counts
    "faint",            # hints and disabled text — the one text colour below AA
    "accent",           # the primary action
    "accent_hover",
    "on_accent",        # text on the primary action
    "selection",        # a selected table row
    "success",
    "warning",
    "danger",
    "turbo",
    "turbo_hover",
)

#: Today's colours, token for token. Dark is exactly what 1.x shipped, so a
#: screenshot before and after moving to tokens must not differ by a pixel.
CLASSIC: dict[str, Colour] = {
    "bg":           ("#f4f5f8", "#17171b"),
    "surface":      ("#ffffff", "#1f1f25"),
    "surface_2":    ("#eef0f4", "#26262e"),
    "border":       ("#d8dbe3", "#33333d"),
    "text":         ("#1a1b22", "#e8e8ee"),
    "muted":        ("#565a6b", "#9494a4"),
    "faint":        ("#83879a", "#6a6a78"),
    "accent":       ("#2f6fe4", "#4c8dff"),
    "accent_hover": ("#245bc2", "#3b74d9"),
    "on_accent":    ("#ffffff", "#ffffff"),
    "selection":    ("#d6e4ff", "#2f4f86"),
    "success":      ("#1a7f37", "#3fb950"),
    "warning":      ("#9a6700", "#d9a020"),
    "danger":       ("#cf222e", "#f2585f"),
    "turbo":        ("#d93a1c", "#ff4d2d"),
    "turbo_hover":  ("#b8301a", "#ff6f52"),
}

#: The palette the app draws with.
PALETTE: dict[str, Colour] = CLASSIC


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
    ("text", ("bg", "surface", "surface_2"), 4.5),
    ("muted", ("bg", "surface", "surface_2"), 4.5),
    ("faint", ("bg", "surface", "surface_2"), 3.0),
    ("on_accent", ("accent", "accent_hover"), 4.5),
    ("accent", ("bg", "surface"), 3.0),
    ("success", ("surface_2",), 3.0),
    ("warning", ("bg", "surface", "surface_2"), 3.0),
    ("danger", ("bg", "surface", "surface_2"), 3.0),
    ("text", ("selection",), 4.5),
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
