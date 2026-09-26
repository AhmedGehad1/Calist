"""Tests for theme.py — colours, type and icons. Run with: python -m pytest

All of these run without a display: theme.py is pure data, which is the point
of keeping it apart from ui.py.
"""

import subprocess
import sys

import pytest

import theme


# ── colour ───────────────────────────────────────────────────────────────────

def test_the_palette_defines_every_token_in_both_modes():
    missing = [t for t in theme.TOKENS if t not in theme.PALETTE]
    assert not missing
    for token, pair in theme.PALETTE.items():
        assert isinstance(pair, tuple) and len(pair) == 2, token
        for side in pair:
            assert side.startswith("#") and len(side) == 7, (token, side)


def _contrast_failures(palette) -> set[tuple[str, str, str]]:
    failures = set()
    for mode, label in ((theme.LIGHT, "light"), (theme.DARK, "dark")):
        for fg, backs, minimum in theme.CONTRAST_RULES:
            for back in backs:
                ratio = theme.contrast_ratio(palette[fg][mode], palette[back][mode])
                if ratio < minimum:
                    failures.add((fg, back, label))
    return failures


def test_every_pairing_meets_its_contrast_minimum():
    """WCAG AA, in light and in dark: text 4.5, non-text and large text 3.0.

    1.x's own palette missed this three times (white on its blue button was
    3.2:1); the replacement has to pass every pairing, in both modes.
    """
    assert _contrast_failures(theme.PALETTE) == set()


def test_the_look_falls_back_to_windows_10_fonts():
    for candidates in (theme.ACTIVE.body, theme.ACTIVE.display, theme.ACTIVE.mono):
        assert candidates[-1] in {"Segoe UI", "Consolas"}, candidates


def test_contrast_ratio_matches_the_wcag_reference_points():
    assert theme.contrast_ratio("#000000", "#ffffff") == pytest.approx(21.0)
    assert theme.contrast_ratio("#ffffff", "#ffffff") == pytest.approx(1.0)
    # The textbook AA boundary: #767676 on white is the lightest grey that passes.
    assert theme.contrast_ratio("#767676", "#ffffff") == pytest.approx(4.54, abs=0.01)


def test_pick_takes_the_side_for_the_mode_and_passes_plain_colours_through():
    assert theme.pick(("#ffffff", "#000000"), dark=True) == "#000000"
    assert theme.pick(("#ffffff", "#000000"), dark=False) == "#ffffff"
    assert theme.pick("#123456", dark=True) == "#123456"


# ── type ─────────────────────────────────────────────────────────────────────

def test_windows_11_fonts_are_used_where_installed():
    installed = ["Arial", "Segoe UI", "Segoe UI Variable Text", "Cascadia Mono"]
    assert theme.pick_family(installed, theme.BODY_FONTS) == "Segoe UI Variable Text"
    assert theme.pick_family(installed, theme.MONO_FONTS) == "Cascadia Mono"


def test_a_windows_10_machine_falls_back_to_what_it_ships():
    """Segoe UI Variable and Cascadia Mono are not Windows 10 system fonts."""
    windows10 = ["Arial", "Consolas", "Segoe UI", "Segoe MDL2 Assets"]
    assert theme.pick_family(windows10, theme.BODY_FONTS) == "Segoe UI"
    assert theme.pick_family(windows10, theme.DISPLAY_FONTS) == "Segoe UI"
    assert theme.pick_family(windows10, theme.MONO_FONTS) == "Consolas"


def test_family_matching_ignores_case_and_an_empty_list_falls_back():
    assert theme.pick_family(["segoe ui"], theme.BODY_FONTS) == "Segoe UI"
    assert theme.pick_family([], theme.BODY_FONTS, theme.FALLBACK_BODY) == "Segoe UI"


def test_every_candidate_list_ends_on_a_windows_10_font():
    windows10 = {"Segoe UI", "Consolas", "Segoe MDL2 Assets"}
    for candidates in (theme.BODY_FONTS, theme.DISPLAY_FONTS, theme.MONO_FONTS,
                       theme.ICON_FONTS):
        assert candidates[-1] in windows10, candidates


# ── icons ────────────────────────────────────────────────────────────────────

def test_icons_are_single_private_use_glyphs():
    for name, glyph in theme.ICONS.items():
        assert len(glyph) == 1 and 0xE700 <= ord(glyph) <= 0xF8FF, name


def _missing_glyphs(font: str, chars: str) -> list[str] | None:
    """Characters `font` has no glyph for, via GDI; None when it is not installed."""
    import ctypes
    from ctypes import wintypes

    gdi, user = ctypes.windll.gdi32, ctypes.windll.user32
    # Handles are 64-bit; left undeclared, ctypes passes them as 32-bit ints.
    HANDLE = wintypes.HANDLE
    user.GetDC.restype = HANDLE
    user.GetDC.argtypes = [wintypes.HWND]
    user.ReleaseDC.argtypes = [wintypes.HWND, HANDLE]
    gdi.CreateFontW.restype = HANDLE
    gdi.SelectObject.restype = HANDLE
    gdi.SelectObject.argtypes = [HANDLE, HANDLE]
    gdi.DeleteObject.argtypes = [HANDLE]
    gdi.GetTextFaceW.argtypes = [HANDLE, ctypes.c_int, wintypes.LPWSTR]
    gdi.GetGlyphIndicesW.argtypes = [HANDLE, wintypes.LPCWSTR, ctypes.c_int,
                                     ctypes.POINTER(wintypes.WORD), wintypes.DWORD]
    dc = user.GetDC(None)
    handle = gdi.CreateFontW(-16, 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 0, 0, font)
    old = gdi.SelectObject(dc, handle)
    try:
        face = ctypes.create_unicode_buffer(64)
        gdi.GetTextFaceW(dc, 64, face)
        if face.value.lower() != font.lower():
            return None                                     # substituted
        indices = (wintypes.WORD * len(chars))()
        # GGI_MARK_NONEXISTING_GLYPHS: a missing glyph comes back as 0xFFFF.
        gdi.GetGlyphIndicesW(dc, chars, len(chars), indices, 1)
        return [c for c, i in zip(chars, indices) if i == 0xFFFF]
    finally:
        gdi.SelectObject(dc, old)
        gdi.DeleteObject(handle)
        user.ReleaseDC(None, dc)


@pytest.mark.skipif(sys.platform != "win32", reason="GDI is Windows-only")
def test_every_icon_exists_in_segoe_mdl2_assets():
    """A codepoint the font lacks renders as an empty box — on every machine."""
    missing = _missing_glyphs("Segoe MDL2 Assets", "".join(theme.ICONS.values()))
    if missing is None:
        pytest.skip("Segoe MDL2 Assets is not installed here")
    names = [n for n, g in theme.ICONS.items() if g in missing]
    assert not names, f"not in Segoe MDL2 Assets: {names}"


# ── the rule that keeps it testable ──────────────────────────────────────────

def test_importing_theme_pulls_in_no_gui_toolkit():
    code = "import theme, sys; assert 'tkinter' not in sys.modules"
    subprocess.run([sys.executable, "-c", code], check=True)
