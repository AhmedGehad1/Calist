"""Calist desktop interface.

All customtkinter code lives here; calist.py stays a pure pipeline and knows
nothing about this module. The two meet at the logging module — the pipeline
emits log records, TkLogHandler routes them into the details drawer — and at
the structured FileOutcome / RunResult values the table renders.

The window carries one layout through three states:

    setup    choose forms and a template; every file pre-validated on arrival
    working  per-row status streams in live, with progress and a cancel
    results  what was built, where it went, and one click to open it
"""

from __future__ import annotations

import json
import logging
import os
import queue
import subprocess
import sys
import threading
import time
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

import access
import calist
import theme
import ui_state
from calist import (ATTRIBUTION, AUTHOR_EMAIL, AUTHOR_NAME, CANCELLED, COPY,
                    ERROR, LEFT_OUT, OK, READY, UNKNOWN_CODE, UNSUPPORTED,
                    FileOutcome, RunResult)

# Drag-and-drop is a bonus, never a requirement: without tkinterdnd2 the drop
# zone is simply click-only.
try:
    from tkinterdnd2 import DND_FILES, REFUSE_DROP, TkinterDnD
    HAS_DND = True
except Exception:                                    # pragma: no cover
    HAS_DND = False

# ──────────────────────────────────────────────────────────────────────────────
# Appearance
# ──────────────────────────────────────────────────────────────────────────────

# Every colour is a (light, dark) pair from theme.py. CustomTkinter widgets
# take the pair and switch by themselves; Tk-native ones (the device table, its
# scrollbar, the flame) take live(pair) and are re-coloured by _apply_theme.
#: The look this build draws: palette, type, shape and signature pieces.
DIR = theme.ACTIVE
_P = theme.PALETTE
BG = _P["bg"]                   # window
SURFACE = _P["surface"]         # cards, the table
SURFACE_2 = _P["surface_2"]     # inputs, secondary buttons
CHROME = _P["chrome"]           # the footer and the settings drawer
ROW_ALT = _P["row_alt"]         # every other table row
BORDER = _P["border"]
TEXT = _P["text"]
MUTED = _P["muted"]
FAINT = _P["faint"]

PRIMARY = _P["accent"]
PRIMARY_HOVER = _P["accent_hover"]
ON_ACCENT = _P["on_accent"]
SELECTION = _P["selection"]
HIGHLIGHT = _P["highlight"]
ON_HIGHLIGHT = _P["on_highlight"]
SUCCESS = _P["success"]
WARNING = _P["warning"]
CAUTION = _P["caution"]         # a warning dot or glyph
WARN_TINT = _P["warn_tint"]     # a row that needs a look
DANGER = _P["danger"]
DANGER_TINT = _P["danger_tint"]  # a row that could not be read

#: Turbo: the flat bolt that marks a run of 1,000 forms or more.
TURBO = _P["turbo"]
TURBO_HOVER = _P["turbo_hover"]

#: The appearance a fresh install opens in.
DEFAULT_APPEARANCE = "dark"


def live(colour) -> str:
    """The side of a (light, dark) pair the window is showing right now."""
    return theme.pick(colour, dark=ctk.get_appearance_mode() == "Dark")

#: The "real device forms only" switch. On, anything that is not a device form
#: — a device list, a 000 template, a name with no customer code — is left out
#: of the table and the register, and listed under Details instead.
FORMS_ONLY_LABEL = "Real device forms only"

# Families are chosen from what is installed (resolve_fonts), because Tk
# substitutes a missing family without a word. These are the Windows 10
# fallbacks until then.
FONT = "Segoe UI"
DISPLAY = "Segoe UI"
MONO = "Consolas"
ICON_FONT = "Segoe MDL2 Assets"

#: Families that exist only on Windows 11. `--win10` screenshots and
#: force_windows10_fonts() pretend they are absent.
WINDOWS_11_ONLY = frozenset({"Segoe UI Variable Text", "Segoe UI Variable Display",
                             "Segoe UI Variable Small", "Cascadia Mono",
                             "Cascadia Code", "Segoe Fluent Icons"})
_windows10_fonts = False


def force_windows10_fonts() -> None:
    """Draw with the fonts a Windows 10 machine has, to preview it here."""
    global _windows10_fonts
    _windows10_fonts = True


def resolve_fonts(root) -> None:
    """Pick each role's family from the installed ones, once, at startup."""
    global FONT, DISPLAY, MONO
    import tkinter.font as tkfont
    installed = set(tkfont.families(root))
    if _windows10_fonts:
        installed -= WINDOWS_11_ONLY
    FONT = theme.pick_family(installed, DIR.body, theme.FALLBACK_BODY)
    DISPLAY = theme.pick_family(installed, DIR.display, theme.FALLBACK_BODY)
    MONO = theme.pick_family(installed, DIR.mono, theme.FALLBACK_MONO)


def body_font(size: int = 13, weight: str = "normal") -> ctk.CTkFont:
    return ctk.CTkFont(FONT, size, weight)


def display_font(size: int, weight: str = "bold") -> ctk.CTkFont:
    return ctk.CTkFont(DISPLAY, size, weight)


def mono_font(size: int = 11) -> ctk.CTkFont:
    return ctk.CTkFont(MONO, size)


def icon_font(size: int = 14) -> ctk.CTkFont:
    return ctk.CTkFont(ICON_FONT, size)


#: Tags of rows that carry no background of their own, so can be striped.
PLAIN_TAGS = ("ready", "ok")

#: Row status → (label, treeview tag)
STATUS_DISPLAY = {
    READY: ("Ready", "ready"),
    OK: ("Read", "ok"),
    UNKNOWN_CODE: ("Unknown code", "warn"),
    UNSUPPORTED: ("Not a form Calist reads", "warn"),
    ERROR: ("Could not be read", "error"),
    CANCELLED: ("Cancelled", "muted"),
    COPY: ("Copy, left out", "muted"),
    LEFT_OUT: ("Left out", "muted"),
}


def status_text(outcome: FileOutcome) -> str:
    """What the Status column says: the state first, then the one fact that
    says what to do about it."""
    label = STATUS_DISPLAY.get(outcome.status, (outcome.status, ""))[0]
    if outcome.status == UNKNOWN_CODE and outcome.device_code:
        return f"Unknown code {outcome.device_code}"
    if outcome.status == UNSUPPORTED:
        suffix = Path(outcome.filename).suffix.lower()
        return f"{label} ({suffix})" if suffix else label
    if outcome.detail and outcome.status not in (READY, OK):
        return f"{label}: {outcome.detail}"
    return label


def result_breakdown(result: RunResult) -> str:
    """One line on what went into the register, and what did not."""
    parts = [f"{result.rows_written:,} rows from {result.files_read:,} forms"]
    if result.second_rows_added:
        parts.append(f"{result.second_rows_added:,} module rows")
    if result.problems:
        parts.append(f"{len(result.problems):,} need a look")
    if result.copies:
        parts.append(f"{len(result.copies):,} "
                     f"cop{'ies' if len(result.copies) != 1 else 'y'} left out")
    if result.left_out:
        parts.append(f"{len(result.left_out):,} not device forms")
    if result.duplicates_removed:
        parts.append(f"{result.duplicates_removed:,} duplicate serial"
                     f"{'s' if result.duplicates_removed != 1 else ''} removed")
    return ", ".join(parts)

SETTINGS_FILE = (Path(os.environ.get("APPDATA") or Path.home())
                 / "Calist" / "settings.json")

#: How often an open window re-checks whether the calendar date has moved on.
NEW_DAY_CHECK_MS = 30_000


#: The mark's renditions, drawn by docs/make_icon.py: the lock screen's 48
#: logical pixels at 100%, 125%, 150% and 200%.
MARK_SIZES = (48, 60, 72, 96)


def mark_image(widget, size: int = 48) -> tk.PhotoImage | None:
    """The app's mark at `size` logical pixels on this display, or None.

    Tk can only shrink an image by dropping whole pixels, which shreds a small
    rounded tile, so each scaling has its own drawn file and the nearest one
    at or above the wanted size is used. Best-effort, like the window icon.
    """
    wanted = size * ctk.ScalingTracker.get_widget_scaling(widget)
    size = next((n for n in MARK_SIZES if n >= wanted - 0.5), MARK_SIZES[-1])
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    try:
        return tk.PhotoImage(master=widget, file=str(base / "docs" /
                                                     f"calist-mark-{size}.png"))
    except Exception:
        log_ui.debug("Could not load the mark", exc_info=True)
        return None


def app_icon() -> Path | None:
    """The window icon, whether frozen or running from a checkout.

    Same resolution as ``calist.bundled_template()``: a frozen build unpacks
    its data files under ``sys._MEIPASS``, and the spec keeps the ``docs/``
    prefix so one lookup covers both cases.
    """
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    candidate = base / "docs" / "calist.ico"
    return candidate if candidate.is_file() else None


# ──────────────────────────────────────────────────────────────────────────────
# Settings
# ──────────────────────────────────────────────────────────────────────────────

def load_settings() -> dict:
    """Best-effort read of remembered preferences; never raises.

    Read as utf-8-sig, not utf-8: Notepad and Windows PowerShell both write a
    UTF-8 BOM, and a BOM makes json.loads raise. That would be swallowed here
    and silently reset every saved preference. utf-8-sig strips a BOM when
    present and is identical to utf-8 when it is not.
    """
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8-sig"))
    except Exception:
        return {}


def forms_only_setting(settings: dict) -> bool:
    """The "real device forms only" switch, as remembered.

    It replaced the three-position filename check. A settings file from before
    holds that under `name_check` (0-2) or `strict_names` (a bool); any
    position but off meant "be strict", which is what this switch is now.
    """
    remembered = settings.get("real_forms_only")
    if remembered is None:
        remembered = settings.get("name_check", settings.get("strict_names", False))
    try:
        return bool(int(remembered))
    except (TypeError, ValueError):
        return False


def appearance_setting(settings: dict) -> str:
    """"dark" or "light", as remembered; anything else is the default."""
    value = str(settings.get("appearance", "")).lower()
    return value if value in ("dark", "light") else DEFAULT_APPEARANCE


def save_settings(data: dict) -> None:
    """Best-effort write. A read-only profile must not break the app."""
    try:
        SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
        SETTINGS_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except Exception:
        log_ui.debug("Could not save settings to %s", SETTINGS_FILE, exc_info=True)


log_ui = logging.getLogger("calist.ui")


# ──────────────────────────────────────────────────────────────────────────────
# Logging bridge
# ──────────────────────────────────────────────────────────────────────────────

class TkLogHandler(logging.Handler):
    """Routes log records to a Tk text widget from any thread.

    emit() only touches a queue — never a widget. A poller running on the main
    thread drains it. Calling widget.after() from a worker thread is the usual
    shortcut, but it reaches into Tk from outside the main loop; batching
    through a queue is both correct and cheaper when a run produces hundreds of
    lines in quick succession.
    """

    POLL_MS = 120

    def __init__(self, widget: ctk.CTkTextbox):
        super().__init__()
        self.widget = widget
        self._pending: queue.Queue[str] = queue.Queue()
        self._stopped = False
        self.widget.after(self.POLL_MS, self._pump)

    def emit(self, record: logging.LogRecord) -> None:
        self._pending.put(self.format(record))

    def close(self) -> None:
        self._stopped = True
        super().close()

    def _pump(self) -> None:
        if self._stopped:
            return
        lines = []
        try:
            while True:
                lines.append(self._pending.get_nowait())
        except queue.Empty:
            pass

        try:
            if lines:
                self.widget.configure(state="normal")
                self.widget.insert("end", "\n".join(lines) + "\n")
                self.widget.see("end")
                self.widget.configure(state="disabled")
            self.widget.after(self.POLL_MS, self._pump)
        except tk.TclError:
            self._stopped = True            # window went away mid-poll


class StatusFormatter(logging.Formatter):
    """Tags warnings and errors, leaving ordinary progress lines unadorned."""

    LABELS = {logging.WARNING: "WARN", logging.ERROR: "ERROR",
              logging.CRITICAL: "FATAL"}

    def format(self, record: logging.LogRecord) -> str:
        message = record.getMessage()
        label = self.LABELS.get(record.levelno)
        return f"[{label}]  {message}" if label else message


# ──────────────────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────────────────

def shorten_path(path: Path | str, limit: int = 74) -> str:
    """Trim a long path from the middle, keeping the drive and the tail.

    The destination folder is the one thing the user most needs to read, so
    when it will not fit, the deepest folders and the filename survive and the
    middle is elided — never the end.
    """
    text = str(path)
    if len(text) <= limit:
        return text

    parts = Path(text).parts
    if len(parts) <= 2:
        return text

    drive, rest = parts[0], list(parts[1:])
    kept: list[str] = []
    for part in reversed(rest):
        candidate = os.sep.join(reversed([*kept, part]))
        if kept and len(drive) + 2 + len(candidate) > limit:
            break
        kept.append(part)

    return f"{drive}…{os.sep}" + os.sep.join(reversed(kept))


def human_duration(seconds: float) -> str:
    if seconds < 1:
        return "moments"
    if seconds < 60:
        return f"{int(seconds)}s"
    minutes, secs = divmod(int(seconds), 60)
    return f"{minutes}m {secs:02d}s"


def tint_title_bar(window) -> None:
    """Colour the title bar and the window's edge to match it, on Windows 11.

    DWMWA_BORDER_COLOR (34), DWMWA_CAPTION_COLOR (35) and DWMWA_TEXT_COLOR (36)
    exist from Windows 11 only. Windows 10 answers with an error code and keeps the dark or light
    title bar CustomTkinter already set, which is the right fallback — so the
    result is ignored, and nothing here may raise.
    """
    if sys.platform != "win32":
        return
    try:
        import ctypes
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
        for attribute, colour in ((34, BORDER), (35, BG), (36, TEXT)):
            rgb = live(colour).lstrip("#")
            # COLORREF is 0x00BBGGRR.
            value = ctypes.c_int(int(rgb[4:6] + rgb[2:4] + rgb[0:2], 16))
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, attribute, ctypes.byref(value), ctypes.sizeof(value))
    except Exception:
        log_ui.debug("Could not tint the title bar", exc_info=True)


def reveal_in_explorer(path: Path) -> None:
    """Open the containing folder with the file selected."""
    if sys.platform == "win32":
        subprocess.run(["explorer", "/select,", str(path)])
    elif sys.platform == "darwin":
        subprocess.run(["open", "-R", str(path)])
    else:
        subprocess.run(["xdg-open", str(path.parent)])


def open_file(path: Path) -> None:
    if sys.platform == "win32":
        os.startfile(path)                             # noqa: S606
    elif sys.platform == "darwin":
        subprocess.run(["open", str(path)])
    else:
        subprocess.run(["xdg-open", str(path)])


#: Device-table row height at 100% scaling. Dense enough that a 1366x768
#: laptop shows 15 or more forms at once, loose enough for the dot to breathe.
ROW_HEIGHT = 26


def table_row_height(widget) -> int:
    """ROW_HEIGHT in device pixels. ttk's rowheight is raw pixels and does not
    follow CustomTkinter's scaling, so without this a 125% display gets rows
    too short for their own text."""
    return round(ROW_HEIGHT * ctk.ScalingTracker.get_widget_scaling(widget))


def style_treeview(root=None) -> None:
    """Make ttk.Treeview match the surrounding CTk surfaces.

    A Treeview rather than stacked CTk frames because this table routinely
    holds hundreds of rows, which native Tk handles and a pile of CTk widgets
    would not.
    """
    style = ttk.Style()
    style.theme_use("clam")

    # ttk takes one colour, not a pair — so this resolves the pair for the
    # current appearance, and _apply_theme calls it again on every switch.
    style.configure(
        "Calist.Treeview",
        background=live(SURFACE), fieldbackground=live(SURFACE),
        foreground=live(TEXT), borderwidth=0, relief="flat", font=(FONT, 10),
        rowheight=table_row_height(root) if root is not None else ROW_HEIGHT,
    )
    style.configure(
        "Calist.Treeview.Heading",
        background=live(SURFACE), foreground=live(MUTED), relief="flat",
        borderwidth=0, padding=(4, 6), font=(FONT, 9),
    )
    style.map("Calist.Treeview",
              background=[("selected", live(SELECTION))],
              foreground=[("selected", live(TEXT))])
    style.map("Calist.Treeview.Heading", background=[("active", live(SURFACE))])
    # Drop the default border box.
    style.layout("Calist.Treeview",
                 [("Calist.Treeview.treearea", {"sticky": "nswe"})])


def status_dots(root, row_px: int) -> dict[str, tk.PhotoImage]:
    """A small filled circle per status tag, for the table's #0 column.

    Drawn pixel by pixel into a PhotoImage — no image files, no Pillow — at a
    size that follows the row height, so it stays crisp at any scaling.
    PhotoImage has on/off transparency only, so the edge is decided per pixel
    from 4x4 samples rather than blended.
    """
    size = max(7, round(row_px * 0.30))
    radius = size / 2
    dots = {}
    for tag, colour in (("ready", FAINT), ("ok", SUCCESS), ("warn", CAUTION),
                        ("error", DANGER), ("muted", BORDER)):
        image = tk.PhotoImage(master=root, width=size, height=size)
        fill = live(colour)
        for y in range(size):
            for x in range(size):
                inside = sum(
                    (x + (i + .5) / 4 - radius) ** 2 + (y + (j + .5) / 4 - radius) ** 2
                    <= radius ** 2 for i in range(4) for j in range(4))
                if inside >= 8:
                    image.put(fill, (x, y))
        dots[tag] = image
    return dots


# ── controls ─────────────────────────────────────────────────────────────────
#
# Every button is one of a few roles, so the same action looks the same
# everywhere and a disabled button never looks pressable: CustomTkinter keeps
# a disabled button's fill, which left 1.x's "Build register" bright blue while
# it did nothing.

BUTTON_ROLES = {
    "primary":   dict(fg_color=PRIMARY, hover_color=PRIMARY_HOVER,
                      text_color=ON_ACCENT, border_width=0),
    "secondary": dict(fg_color=SURFACE_2, hover_color=BORDER, text_color=TEXT,
                      border_width=1, border_color=BORDER),
    "ghost":     dict(fg_color="transparent", hover_color=SURFACE_2,
                      text_color=MUTED, border_width=0),
    "chip":      dict(fg_color=SURFACE, hover_color=SURFACE_2, text_color=TEXT,
                      border_width=1, border_color=BORDER),
    "picked":    dict(fg_color=HIGHLIGHT, hover_color=HIGHLIGHT,
                      text_color=ON_HIGHLIGHT, border_width=1,
                      border_color=HIGHLIGHT),
}
DISABLED_LOOK = dict(fg_color=SURFACE_2, hover_color=SURFACE_2,
                     text_color_disabled=FAINT)


def make_button(master, text: str, command, *, role: str = "secondary",
                width: int = 120, height: int = 36, font=None) -> ctk.CTkButton:
    look = BUTTON_ROLES[role]
    button = ctk.CTkButton(master, text=text, command=command, width=width,
                           height=height, corner_radius=DIR.control_radius,
                           font=font or body_font(13, "bold" if role == "primary"
                                                  else "normal"),
                           text_color_disabled=FAINT, **look)
    button._calist_role = role
    keyboard_ready(button)
    return button


def keyboard_ready(button: ctk.CTkButton) -> None:
    """Let Tab reach the button, show where focus is, and press on Enter/Space.

    CustomTkinter buttons take no keyboard focus, and route their own events
    to an inner canvas — so this works on the frame itself. A mouse click does
    not move focus, so the ring appears only for someone using the keyboard.
    The ring is the text colour: it reads on the teal primary as well as on
    the plain buttons.
    """
    tk.Frame.configure(button, takefocus=1)

    def focused(_event) -> None:
        if str(button.cget("state")) != "disabled":
            button.configure(border_width=2, border_color=TEXT)

    def unfocused(_event) -> None:
        look = BUTTON_ROLES.get(getattr(button, "_calist_role", "secondary"), {})
        button.configure(border_width=look.get("border_width", 0),
                         border_color=look.get("border_color", BORDER))

    # Tab focuses the frame; CustomTkinter's own focus_set() hands focus to
    # the inner text label. Either way the ring and the keys must work.
    targets = [button] + [w for w in (getattr(button, "_text_label", None),) if w]
    for target in targets:
        tk.Misc.bind(target, "<FocusIn>", focused, add="+")
        tk.Misc.bind(target, "<FocusOut>", unfocused, add="+")
        for key in ("<Return>", "<KP_Enter>", "<space>"):
            tk.Misc.bind(target, key, lambda _e: button.invoke(), add="+")


def set_enabled(button: ctk.CTkButton, enabled: bool) -> None:
    """Enable or disable a button, and make it look it."""
    role = getattr(button, "_calist_role", "secondary")
    # A disabled button leaves the Tab order: there is nothing to press.
    tk.Frame.configure(button, takefocus=1 if enabled else 0)
    if enabled:
        button.configure(state="normal", **BUTTON_ROLES[role])
    else:
        look = dict(DISABLED_LOOK)
        if role == "ghost":
            # A hover colour may not be "transparent"; a disabled button never
            # shows its hover anyway, so it is simply left as it was.
            look = dict(fg_color="transparent", text_color_disabled=FAINT)
        button.configure(state="disabled", **look)


def make_icon_button(master, icon: str, command, *, size: int = 34,
                     glyph_size: int = 14) -> ctk.CTkButton:
    """A square button carrying one Segoe MDL2 Assets glyph."""
    button = ctk.CTkButton(master, text=theme.ICONS[icon], command=command,
                           width=size, height=size, corner_radius=DIR.control_radius,
                           font=icon_font(glyph_size), fg_color="transparent",
                           hover_color=SURFACE_2, text_color=MUTED)
    button._calist_role = "ghost"
    keyboard_ready(button)
    return button


def make_check(master, text: str, variable, command) -> ctk.CTkCheckBox:
    """An on/off setting, as a checkbox.

    Not a CTkSwitch: CustomTkinter draws a switch's knob larger than its
    track, so whatever colour the knob is, it vanishes into the drawer in
    some look and mode — white on the white drawer, black on the black one.
    A bordered box with the accent's own check mark reads everywhere.
    """
    return ctk.CTkCheckBox(master, text=text, variable=variable, command=command,
                           font=body_font(13), text_color=TEXT,
                           checkbox_width=20, checkbox_height=20, border_width=2,
                           corner_radius=max(DIR.control_radius // 2, 3),
                           border_color=MUTED, hover_color=PRIMARY_HOVER,
                           fg_color=PRIMARY, checkmark_color=ON_ACCENT)


class Toggle(ctk.CTkFrame):
    """Mutually exclusive choices — All/Problems, Dark/Light.

    Not a CTkSegmentedButton: that takes one text colour for every segment,
    and no single colour reads on both the accent-filled selection and the
    plain segments in every look (D's yellow wants black text, its plain
    segments white).
    """

    def __init__(self, master, values, command=None, height: int = 32):
        super().__init__(master, fg_color=SURFACE_2, border_width=1,
                         border_color=BORDER, corner_radius=DIR.control_radius)
        self._command = command
        self._value = values[0]
        self._buttons = {}
        inner = max(DIR.control_radius - 2, 2)
        for column, value in enumerate(values):
            button = ctk.CTkButton(self, text=value, width=76, height=height - 8,
                                   corner_radius=inner, font=body_font(12),
                                   command=lambda v=value: self._pick(v))
            button.grid(row=0, column=column, padx=(4 if column == 0 else 0, 4),
                        pady=4)
            keyboard_ready(button)
            self._buttons[value] = button
        self._paint()

    def get(self) -> str:
        return self._value

    def set(self, value: str | None) -> None:
        """Select `value`; None shows no selection (another filter is on)."""
        if value is None or value in self._buttons:
            self._value = value
            self._paint()

    def _pick(self, value: str) -> None:
        if value != self._value:
            self.set(value)
            if self._command:
                self._command(value)

    def _paint(self) -> None:
        for value, button in self._buttons.items():
            on = value == self._value
            button.configure(fg_color=PRIMARY if on else "transparent",
                             hover_color=PRIMARY_HOVER if on else BORDER,
                             text_color=ON_ACCENT if on else MUTED)


# customtkinter needs a mixin to cooperate with tkinterdnd2's root.
if HAS_DND:                                            # pragma: no cover
    class _Root(ctk.CTk, TkinterDnD.DnDWrapper):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            # tkdnd is a native library loaded into Tcl. If it is missing or
            # blocked (an antivirus quarantining one DLL), the window must
            # still open — click-only, as without the package — not crash.
            try:
                self.TkdndVersion = TkinterDnD._require(self)
                self.dnd_ready = True
            except Exception:
                self.dnd_ready = False
                log_ui.debug("Drag and drop unavailable", exc_info=True)
else:
    _Root = ctk.CTk


# ──────────────────────────────────────────────────────────────────────────────
# Lock screen
# ──────────────────────────────────────────────────────────────────────────────

class LockPanel(ctk.CTkFrame):
    """Asks for the day's PIN, as a panel that covers the whole window.

    Deliberately NOT a Toplevel over a withdrawn root. CustomTkinter's CTk
    tracks whether its window has ever been shown, and calling withdraw()
    before the first mainloop() sets a flag that makes its own first-show
    routine hide the window and never bring it back — the app ends up alive but
    invisible. Keeping one always-visible root and swapping what is inside it
    sidesteps that entirely, and removes the startup flicker as a bonus.

    Owns no persistence: it reports back through ``result`` and ``state_dict``,
    and the caller decides what to save.
    """

    DOTS = 4

    def __init__(self, master, state: dict, on_done):
        super().__init__(master, fg_color=BG, corner_radius=0)

        self.state_dict = dict(state)
        self.result = False
        self._on_done = on_done
        self._entry = ""

        self._build()
        self._refresh_dots()
        self._tick_cooldown()

    def take_focus(self) -> None:
        """Route typing here without stealing focus from the whole desktop."""
        self.focus_set()
        self.winfo_toplevel().bind("<Key>", self._on_key)

    def release_focus(self) -> None:
        self.winfo_toplevel().unbind("<Key>")

    # ── layout ───────────────────────────────────────────────────────────────

    def _build(self) -> None:
        """One card, centred: the mark, the name, four dots and a keypad.

        The keys are the same buttons as the rest of the app. Clear is a word
        and backspace its glyph, so no key is a character standing in for an
        icon.
        """
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)          # centre the card

        card = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=DIR.radius,
                            border_width=1, border_color=BORDER)
        card.grid(row=1, column=0)
        card.grid_columnconfigure(0, weight=1)
        self._card = card

        self._mark = mark_image(self, 48)
        self._logo = tk.Label(card, image=self._mark, bd=0, highlightthickness=0,
                              bg=live(SURFACE))
        if self._mark is not None:
            self._logo.grid(row=0, column=0, pady=(34, 12))
        ctk.CTkLabel(card, text="Calist", text_color=TEXT, font=display_font(24)
                     ).grid(row=1, column=0)
        ctk.CTkLabel(card, text="Enter today's access code", text_color=MUTED,
                     font=body_font(13)).grid(row=2, column=0, pady=(2, 0))

        # PIN dots
        self._dots = ctk.CTkFrame(card, fg_color="transparent")
        self._dots.grid(row=3, column=0, pady=(22, 4))
        self._dot_widgets = []
        for i in range(self.DOTS):
            dot = ctk.CTkFrame(self._dots, width=14, height=14, corner_radius=7,
                               fg_color=SURFACE_2, border_width=1,
                               border_color=BORDER)
            dot.grid(row=0, column=i, padx=8)
            dot.grid_propagate(False)
            self._dot_widgets.append(dot)

        self._message = ctk.CTkLabel(card, text="", text_color=DANGER,
                                     font=body_font(12))
        self._message.grid(row=4, column=0, pady=(4, 8))

        # Keypad
        pad = ctk.CTkFrame(card, fg_color="transparent")
        pad.grid(row=5, column=0, padx=40)
        keys = [("1", 0, 0), ("2", 0, 1), ("3", 0, 2),
                ("4", 1, 0), ("5", 1, 1), ("6", 1, 2),
                ("7", 2, 0), ("8", 2, 1), ("9", 2, 2),
                ("C", 3, 0), ("0", 3, 1), ("<", 3, 2)]
        self._keys = []
        for label, r, c in keys:
            if label == "C":
                btn = make_button(pad, "Clear", lambda: self._press("C"),
                                  role="ghost", width=76, height=52,
                                  font=body_font(13))
            elif label == "<":
                btn = make_button(pad, theme.ICONS["backspace"],
                                  lambda: self._press("<"), role="ghost",
                                  width=76, height=52, font=icon_font(18))
            else:
                btn = make_button(pad, label, lambda k=label: self._press(k),
                                  role="secondary", width=76, height=52,
                                  font=display_font(19))
            btn.grid(row=r, column=c, padx=5, pady=5)
            self._keys.append(btn)

        self._hint = ctk.CTkLabel(
            card, text=f"The code changes daily. Ask {AUTHOR_NAME} for today's.",
            text_color=MUTED, font=body_font(12))
        self._hint.grid(row=6, column=0, padx=32, pady=(20, 0))
        ctk.CTkLabel(card, text=AUTHOR_EMAIL, text_color=FAINT,
                     font=body_font(12)).grid(row=7, column=0, pady=(2, 30))

    # ── entry ────────────────────────────────────────────────────────────────

    def _refresh_dots(self) -> None:
        for i, dot in enumerate(self._dot_widgets):
            filled = i < len(self._entry)
            dot.configure(fg_color=PRIMARY if filled else SURFACE_2,
                          border_color=PRIMARY if filled else BORDER)

    def _press(self, key: str) -> None:
        if self._locked_out():
            return
        if key == "C":
            self._entry = ""
        elif key == "<":
            self._entry = self._entry[:-1]
        elif key.isdigit() and len(self._entry) < self.DOTS:
            self._entry += key
            self._message.configure(text="")

        self._refresh_dots()
        if len(self._entry) == self.DOTS:
            self.after(120, self._submit)          # let the last dot paint

    def _on_key(self, event) -> None:
        if event.char.isdigit():
            self._press(event.char)
        elif event.keysym in ("BackSpace", "Delete"):
            self._press("<")

    def _submit(self) -> None:
        if len(self._entry) != self.DOTS or self._locked_out():
            return

        if access.verify_pin(self._entry):
            self.state_dict = access.mark_unlocked(self.state_dict)
            self.result = True
            self.release_focus()
            self._on_done(self)
            return

        self.state_dict = access.record_failure(self.state_dict)
        self._entry = ""
        self._refresh_dots()

        wait = access.cooldown_remaining(self.state_dict)
        if wait:
            self._tick_cooldown()
        else:
            left = access.attempts_left(self.state_dict)
            note = f"  ({left} left)" if left <= 2 else ""
            self._message.configure(text=f"That code is not right{note}",
                                    text_color=DANGER)

    # ── cooldown ─────────────────────────────────────────────────────────────

    def _locked_out(self) -> bool:
        return access.cooldown_remaining(self.state_dict) > 0

    def _tick_cooldown(self) -> None:
        """Count the penalty down in place, disabling the pad while it runs."""
        wait = access.cooldown_remaining(self.state_dict)
        for btn in self._keys:
            set_enabled(btn, not wait)

        if wait:
            self._message.configure(
                text=f"Too many attempts — wait {human_duration(wait)}",
                text_color=WARNING)
            self.after(500, self._tick_cooldown)


# ──────────────────────────────────────────────────────────────────────────────
# Application
# ──────────────────────────────────────────────────────────────────────────────

class App(_Root):
    def __init__(self) -> None:
        super().__init__()
        resolve_fonts(self)

        self.title("Calist")
        self.minsize(1000, 600)
        self._zoom_job: str | None = None
        self._size_to_screen()
        self.configure(fg_color=BG)
        self._apply_icon()
        self.after(300, lambda: tint_title_bar(self))

        self._settings = load_settings()
        self._files: dict[str, FileOutcome] = {}       # path → latest outcome
        self._template = tk.StringVar(value=self._initial_template())
        self._dedup = tk.BooleanVar(value=self._settings.get("deduplicate", False))
        self._forms_only = tk.BooleanVar(value=forms_only_setting(self._settings))
        #: Whether the window is in Turbo: decided from the file count
        #: (ui_state.is_turbo), and captured once more by _start for the run.
        self._turbo_mode = False
        self._cancel: threading.Event | None = None
        self._scan: threading.Event | None = None
        self._scan_from = 0
        self._events: queue.Queue[tuple] = queue.Queue()
        self._result: RunResult | None = None
        self._lock: LockPanel | None = None
        #: A save folder chosen in Settings. For this round only (owner's
        #: decision): Clear all and a restart go back to the round's folder, so
        #: January's choice cannot quietly catch February's register.
        self._outdir = tk.StringVar(value="")
        #: What the user added, in order: (path, is_folder). The register's
        #: default folder comes from the first (ui_state.origin_folder), and
        #: Re-check walks them all again.
        self._sources: list[tuple[str, bool]] = []
        self._fit_job: str | None = None
        self._fit_passes = 0
        self._started_at = 0.0
        #: (files done, monotonic time) samples behind the ETA.
        self._marks: list[tuple[int, float]] = []
        #: Problems seen so far in the run in flight, for Turbo's live summary.
        self._live_problems: list[FileOutcome] = []
        self._live_render = 0.0
        self._log_open = False
        #: The problem group the banner is filtering the table to, if any.
        self._group: ui_state.Group | None = None
        self._search_job: str | None = None
        self._panel_open = False
        self._panel_job: str | None = None
        #: (rows shown, forms in the round) as of the last table refresh.
        self._table_counts = (0, 0)
        #: Pending stand-down of an armed "Clear all", if one is armed.
        self._clear_job: str | None = None
        #: Status dots for the table, drawn per theme (see status_dots).
        self._dots: dict[str, tk.PhotoImage] = {}
        #: One-pixel rules. CustomTkinter will not draw a frame that thin, so
        #: these are Tk frames, re-coloured by _apply_theme.
        self._hairlines: list[tk.Frame] = []
        #: Raw Treeview row height, in device pixels — scaled with the display.
        self.ROW_PX = table_row_height(self)

        style_treeview(self)
        self._build()
        self._attach_logging()
        self._enter_setup()

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Control-o>", lambda _e: self._add_folder())
        # Ctrl+Enter is always the footer's main action: Build before a run,
        # Open register after one.
        self.bind("<Control-Return>", lambda _e: self._primary_action())
        self.bind("<Control-f>", lambda _e: self._focus_search())
        self.bind("<Control-comma>", lambda _e: self._toggle_settings())
        self.bind("<F5>", lambda _e: self._recheck())
        self.bind("<Escape>", self._on_escape)

        # Scheduled, not called: the initial unlock belongs to run(), and
        # checking here as well would raise a second prompt behind the first.
        self.after(NEW_DAY_CHECK_MS, self._watch_for_new_day)

    def _apply_icon(self) -> None:
        """Put the Calist mark on the window and the taskbar.

        Applied twice on purpose. CustomTkinter finishes setting the window up
        on its first mainloop pass, and on Windows that re-show drops an icon
        assigned during ``__init__`` — the app reverts to Tk's default feather.
        Re-applying once the loop is running makes it stick.

        Best-effort throughout: a missing or unreadable icon costs the app its
        logo, which is never worth failing to start over.
        """
        icon = app_icon()
        if icon is None:
            return

        def apply() -> None:
            try:
                self.iconbitmap(str(icon))
            except Exception:
                log_ui.debug("Could not set the window icon", exc_info=True)

        apply()
        self.after(250, apply)

    #: The window a normal screen opens with, in CustomTkinter's logical units.
    WINDOW_W, WINDOW_H = 1180, 780
    #: At or below this many logical pixels of screen height — a 1366x768
    #: laptop, the machine most engineers carry — the window opens maximised,
    #: because every row it gives up is a row of the device table.
    SMALL_SCREEN_H = 800

    def _size_to_screen(self) -> None:
        """Open centred on a roomy screen, maximised on a laptop.

        Maximising is scheduled, never done here: CustomTkinter hides and
        re-shows the window on its first mainloop pass, and a state set before
        that is not the state it restores.
        """
        scale = ctk.ScalingTracker.get_window_scaling(self) or 1.0
        screen_w = self.winfo_screenwidth() / scale
        screen_h = self.winfo_screenheight() / scale
        width = int(min(self.WINDOW_W, screen_w - 40))
        height = int(min(self.WINDOW_H, screen_h - 90))
        x = max(0, int((screen_w - width) / 2))
        y = max(0, int((screen_h - height) / 2) - 20)
        self.geometry(f"{width}x{height}+{x}+{y}")
        if screen_h <= self.SMALL_SCREEN_H:
            self._zoom_job = self.after(80, lambda: self.state("zoomed"))

    def _initial_template(self) -> str:
        """The remembered template, else the one shipped with the app.

        Falling back matters most for a freshly downloaded copy: without it the
        first run stalls on "choose a template" with nothing obvious to choose.
        It also covers a remembered template that has since been moved.
        """
        remembered = self._settings.get("template", "")
        if remembered and Path(remembered).is_file():
            return remembered
        shipped = calist.bundled_template()
        return str(shipped) if shipped else ""

    # ── layout ───────────────────────────────────────────────────────────────
    #
    # The page scrolls; the footer does not. What the user acts on — where the
    # register goes, and the button that builds it — is pinned to the bottom
    # of the window, so no state (a long table, the Details log, a result) can
    # push it out of sight. On a 1366x768 laptop the old single page put the
    # result's buttons below the fold.

    #: Side margin of every block, in logical units.
    PAD = 24

    def _build(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._page = ctk.CTkScrollableFrame(
            self, fg_color=BG, corner_radius=0,
            scrollbar_button_color=SURFACE_2, scrollbar_button_hover_color=BORDER)
        self._page.grid(row=0, column=0, sticky="nsew")
        self._page.grid_columnconfigure(0, weight=1)

        self._build_header()
        self._build_intake()
        self._build_banner()
        self._build_table()
        self._build_log_drawer()
        self._build_footer()
        self._build_settings_panel()
        self._enable_drop()

        # Reflow when the window is resized, and once at startup after the
        # first real geometry is known.
        self.bind("<Configure>", self._on_resize)
        self.after(120, self._fit_to_window)

    def _build_header(self) -> None:
        bar = ctk.CTkFrame(self._page, fg_color="transparent")
        bar.grid(row=0, column=0, sticky="ew", padx=self.PAD, pady=(12, 8))
        bar.grid_columnconfigure(2, weight=1)
        self._header = bar

        # The wordmark alone: the owner preferred it without the app's mark
        # beside it (the mark stays on the lock screen and the taskbar).
        self._lbl_mark = ctk.CTkLabel(bar, text="Calist", text_color=TEXT,
                                      font=display_font(20))
        self._lbl_mark.grid(row=0, column=0, sticky="w")

        # Which round this is — the folder the forms came from. Orientation
        # for the person who has three rounds open in Explorer.
        self._round_rule = self._hairline(bar, vertical=True)
        self._lbl_round = ctk.CTkLabel(bar, text="", text_color=TEXT,
                                       font=body_font(15))

        tools = ctk.CTkFrame(bar, fg_color="transparent")
        tools.grid(row=0, column=3, sticky="e")
        self._btn_details = make_button(tools, "Details", self._toggle_log,
                                        role="ghost", width=84, height=32)
        self._btn_details.grid(row=0, column=0, padx=(0, 4))
        self._btn_settings = make_icon_button(tools, "settings",
                                              self._toggle_settings, size=32)
        self._btn_settings.grid(row=0, column=1)

    def _build_intake(self) -> None:
        """Adding a round is the whole point of the app, so it leads.

        With nothing loaded the hero fills the window; once forms are in it
        collapses to a toolbar — adding, re-checking, searching and the
        All/Problems filter on one line — and hands the space to the table.
        """
        self._intake = ctk.CTkFrame(self._page, fg_color="transparent")
        self._intake.grid(row=1, column=0, sticky="nsew", padx=self.PAD, pady=(0, 8))
        self._intake.grid_columnconfigure(0, weight=1)
        self._intake.grid_rowconfigure(0, weight=1)

        # ── hero ─────────────────────────────────────────────────────────────
        # Fixed height with propagation off, so _fit_to_window can grow the hero
        # into spare window space. Left to size itself it would collapse to its
        # contents, and the empty state is meant to fill the window.
        self._hero = ctk.CTkFrame(self._intake, fg_color=SURFACE,
                                  corner_radius=DIR.radius, border_width=1,
                                  border_color=BORDER, height=self.MIN_HERO_H)
        self._hero.grid(row=0, column=0, sticky="nsew")
        self._hero.grid_propagate(False)
        self._hero.grid_columnconfigure(0, weight=1)
        self._hero.grid_rowconfigure(0, weight=1)
        self._hero.grid_rowconfigure(6, weight=1)      # centres the stack

        # The folder glyph, by the owner's choice over the app's mark here.
        ctk.CTkLabel(self._hero, text=theme.ICONS["folder_open"], text_color=MUTED,
                     font=icon_font(40)).grid(row=1, column=0, pady=(0, 14))
        droppable = getattr(self, "dnd_ready", False)
        ctk.CTkLabel(self._hero, text="Drop a round's folder here" if droppable
                     else "Add a round of inspection forms",
                     text_color=TEXT, font=display_font(22)).grid(row=2, column=0)
        ctk.CTkLabel(self._hero, text=("Or choose it below. " if droppable else
                                       "Choose the round's folder. ")
                     + "Calist reads every form in it, subfolders included.",
                     text_color=MUTED, font=body_font(13)).grid(row=3, column=0,
                                                                pady=(6, 0))

        buttons = ctk.CTkFrame(self._hero, fg_color="transparent")
        buttons.grid(row=4, column=0, pady=(22, 0))
        make_button(buttons, "Add folder", self._add_folder, role="primary",
                    width=170, height=44, font=display_font(15)
                    ).pack(side="left", padx=(0, 10))
        make_button(buttons, "Add files", self._add_files, role="secondary",
                    width=130, height=44, font=body_font(14)).pack(side="left")

        ctk.CTkLabel(self._hero, text="Forms named like G302-AGH001-0425 are "
                     "recognised from the name, before anything is opened.",
                     text_color=FAINT, font=body_font(12)
                     ).grid(row=5, column=0, pady=(18, 0))

        # ── toolbar, once forms are in ───────────────────────────────────────
        bar = ctk.CTkFrame(self._intake, fg_color="transparent")
        bar.grid_columnconfigure(4, weight=1)
        self._toolbar = bar
        self._btn_slim_folder = make_button(bar, "Add folder", self._add_folder,
                                            role="primary", width=112, height=34)
        self._btn_slim_folder.grid(row=0, column=0, padx=(0, 6))
        self._btn_slim_files = make_button(bar, "Add files", self._add_files,
                                           role="secondary", width=96, height=34)
        self._btn_slim_files.grid(row=0, column=1, padx=(0, 6))
        self._btn_recheck = make_button(bar, "Re-check", self._recheck,
                                        role="secondary", width=92, height=34)
        self._btn_recheck.grid(row=0, column=2, padx=(0, 2))
        self._btn_clear = make_button(bar, "Clear all", self._arm_clear,
                                      role="ghost", width=112, height=34)
        self._btn_clear.grid(row=0, column=3)

        # Search: every word must appear somewhere in the row (ui_state.matches).
        box = ctk.CTkFrame(bar, fg_color=SURFACE_2, corner_radius=DIR.control_radius,
                           border_width=1, border_color=BORDER)
        box.grid(row=0, column=5, sticky="e", padx=(12, 8))
        self._search_box = box
        ctk.CTkLabel(box, text=theme.ICONS["search"], text_color=FAINT,
                     font=icon_font(13), width=18).grid(row=0, column=0, padx=(10, 2))
        self._search = ctk.CTkEntry(
            box, width=290, height=30, border_width=0, fg_color=SURFACE_2,
            text_color=TEXT, placeholder_text="Search file, device, serial or code (Ctrl+F)",
            placeholder_text_color=FAINT, font=body_font(12))
        self._search.grid(row=0, column=1, padx=(0, 4), pady=1)
        self._search.bind("<KeyRelease>", self._on_search_typed)
        self._search.bind("<FocusIn>", lambda _e: box.configure(border_color=PRIMARY))
        self._search.bind("<FocusOut>", lambda _e: box.configure(border_color=BORDER))
        self._btn_search_clear = make_icon_button(box, "cancel", self._clear_search,
                                                  size=24, glyph_size=10)
        # Always placed, only shown when there is text: a button that appears
        # would widen the box and shift it left as the user starts typing.
        self._btn_search_clear.grid(row=0, column=2, padx=(0, 4))
        self._show_search_clear(False)

        self._filter = Toggle(bar, ("All", "Problems"),
                              command=self._on_filter_picked)
        self._filter.grid(row=0, column=6, sticky="e")


    def _build_banner(self) -> None:
        """The round in one line: its counts, then its problems, grouped.

        A round's worth of one mistake — fourteen forms named for a device
        code nobody added — is one chip, and one click shows exactly those
        files. The right end says what the table is filtered to, and how to
        clear it.
        """
        self._banner = ctk.CTkFrame(self._page, fg_color="transparent")
        self._banner.grid_columnconfigure(1, weight=1)

    def _build_table(self) -> None:
        wrap = ctk.CTkFrame(self._page, fg_color=SURFACE, corner_radius=DIR.radius,
                            border_width=1, border_color=BORDER)
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_rowconfigure(0, weight=1)
        self._table_wrap = wrap

        body = tk.Frame(wrap, bg=live(SURFACE), highlightthickness=0, bd=0)
        body.grid(row=0, column=0, sticky="nsew", padx=(8, 4), pady=(6, 8))
        self._table_body = body
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(0, weight=1)

        # The #0 column carries a status dot and nothing else. The dot, the
        # status words and — for a problem — a tinted row say what state a row
        # is in; the text itself stays neutral and readable.
        self._tree = ttk.Treeview(
            body, style="Calist.Treeview", show=("tree", "headings"),
            selectmode="extended", columns=("file", "device", "serial", "status"),
            height=self.MIN_ROWS,
        )
        self._tree.heading("file", text="File", anchor="w")
        self._tree.heading("device", text="Device", anchor="w")
        self._tree.heading("serial", text="Serial number", anchor="w")
        self._tree.heading("status", text="Status", anchor="w")
        dot = self.ROW_PX
        self._tree.column("#0", width=dot, minwidth=dot, stretch=False, anchor="center")
        # ttk widths are raw pixels while the row font scales, so at 125% the
        # unscaled columns cut "High Flow Nasal Cannula" short.
        scale = ctk.ScalingTracker.get_widget_scaling(self)
        for name, width, least, stretch in (("file", 270, 160, False),
                                            ("device", 200, 120, False),
                                            ("serial", 160, 100, False),
                                            ("status", 260, 150, True)):
            self._tree.column(name, width=round(width * scale),
                              minwidth=round(least * scale), anchor="w",
                              stretch=stretch)
        self._tree.grid(row=0, column=0, sticky="nsew")

        bar = ctk.CTkScrollbar(body, command=self._tree.yview, width=12,
                               fg_color=SURFACE, button_color=BORDER,
                               button_hover_color=FAINT)
        bar.grid(row=0, column=1, sticky="ns", padx=(4, 0))
        self._tree.configure(yscrollcommand=bar.set)
        # The page binds <MouseWheel> on every descendant, so without this the
        # wheel would scroll the page instead of the list under the pointer.
        self._tree.bind("<MouseWheel>", self._wheel_over(self._tree))

        self._colour_tags()

        self._tree.bind("<Double-1>", self._reveal_selected)
        self._tree.bind("<Delete>", self._remove_selected)
        self._tree.bind("<Button-3>", self._row_menu)
        self._menu = tk.Menu(self, tearoff=0)
        self._menu.add_command(label="Open the form", command=self._open_selected)
        self._menu.add_command(label="Show in folder", command=self._reveal_selected)
        self._menu.add_separator()
        self._menu.add_command(label="Copy file name", command=self._copy_names)
        self._menu.add_separator()
        self._menu.add_command(label="Remove from the list", command=self._remove_selected)

        self._empty = ctk.CTkLabel(body, text="", text_color=FAINT,
                                   font=body_font(13), justify="center")

    def _build_log_drawer(self) -> None:
        self._drawer = ctk.CTkFrame(self._page, fg_color=SURFACE,
                                    corner_radius=DIR.radius, border_width=1,
                                    border_color=BORDER)
        self._drawer.grid_columnconfigure(0, weight=1)
        self._drawer.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(self._drawer, text="Details", text_color=TEXT,
                     font=body_font(13, "bold")).grid(row=0, column=0, sticky="w",
                                                      padx=14, pady=(10, 0))
        self._log_box = ctk.CTkTextbox(
            self._drawer, height=170, fg_color=SURFACE, text_color=MUTED,
            font=mono_font(11), wrap="none", activate_scrollbars=True)
        self._log_box.grid(row=1, column=0, sticky="nsew", padx=8, pady=(2, 8))
        self._log_box.configure(state="disabled")
        self._log_box.bind("<MouseWheel>", self._wheel_over(self._log_box))

    def _build_footer(self) -> None:
        """The action bar: where the register goes, and what to do next.

        One line in every state, so the table keeps the height. It is pinned
        to the window, not the page, so nothing can scroll it out of sight.
        """
        self._footer = ctk.CTkFrame(self, fg_color=CHROME, corner_radius=0)
        self._footer.grid(row=1, column=0, sticky="ew")
        self._footer.grid_columnconfigure(0, weight=1)
        self._hairline(self._footer).grid(row=0, column=0, sticky="ew")

        self._action = ctk.CTkFrame(self._footer, fg_color="transparent")
        self._action.grid(row=1, column=0, sticky="ew", padx=self.PAD, pady=12)
        self._action.grid_columnconfigure(0, weight=1)

        # — idle: where it saves and what is ready; the credit; the button
        self._idle = ctk.CTkFrame(self._action, fg_color="transparent")
        self._idle.grid_columnconfigure(4, weight=1)
        ctk.CTkLabel(self._idle, text="Saves to", text_color=MUTED,
                     font=body_font(12)).grid(row=0, column=0, sticky="w", padx=(0, 8))
        self._lbl_dest = ctk.CTkLabel(self._idle, text="", text_color=TEXT,
                                      anchor="w", font=body_font(12))
        self._lbl_dest.grid(row=0, column=1, sticky="w")
        self._lbl_dest_note = ctk.CTkLabel(self._idle, text="", text_color=WARNING,
                                           font=body_font(12))
        self._lbl_dest_note.grid(row=0, column=2, sticky="w", padx=(8, 0))
        self._lbl_hint = ctk.CTkLabel(self._idle, text="", text_color=MUTED,
                                      anchor="w", font=body_font(12))
        self._lbl_hint.grid(row=0, column=3, sticky="w", padx=(20, 0))
        self._lbl_credit = ctk.CTkLabel(self._idle, text=f"Built by {AUTHOR_NAME}",
                                        text_color=FAINT, font=body_font(11))
        self._lbl_credit.grid(row=0, column=5, sticky="e", padx=(12, 16))
        self._btn_build = make_button(self._idle, "Build register", self._start,
                                      role="primary", width=180, height=40,
                                      font=display_font(14))
        self._btn_build.grid(row=0, column=6, sticky="e")

        # — working: stage, the file in hand, progress; cancel
        self._busy = ctk.CTkFrame(self._action, fg_color="transparent")
        self._busy.grid_columnconfigure(2, weight=1)
        self._lbl_bolt = ctk.CTkLabel(self._busy, text=theme.ICONS["bolt"],
                                      text_color=TURBO, font=icon_font(15))
        self._lbl_stage = ctk.CTkLabel(self._busy, text="", text_color=TEXT,
                                       anchor="w", font=body_font(13, "bold"))
        self._lbl_stage.grid(row=0, column=1, sticky="w")
        self._lbl_current = ctk.CTkLabel(self._busy, text="", text_color=MUTED,
                                         anchor="w", font=body_font(12))
        self._lbl_current.grid(row=0, column=2, sticky="w", padx=(14, 0))
        self._lbl_eta = ctk.CTkLabel(self._busy, text="", text_color=MUTED,
                                     anchor="e", font=body_font(12))
        self._lbl_eta.grid(row=0, column=3, sticky="e", padx=(0, 16))
        self._btn_cancel = make_button(self._busy, "Cancel", self._cancel_run,
                                       role="secondary", width=110, height=40)
        self._btn_cancel.configure(hover_color=DANGER)
        self._btn_cancel.grid(row=0, column=4, rowspan=2, sticky="e")
        self._bar = ctk.CTkProgressBar(self._busy, height=6, corner_radius=3,
                                       progress_color=PRIMARY, fg_color=SURFACE_2)
        self._bar.set(0)
        self._bar.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(6, 0),
                       padx=(0, 16))

        # — results: the verdict and its facts; what to do with the register
        self._done = ctk.CTkFrame(self._action, fg_color="transparent")
        self._done.grid_columnconfigure(1, weight=1)
        self._verdict = ctk.CTkFrame(self._done, fg_color="transparent")
        self._verdict.grid(row=0, column=0, rowspan=2, sticky="w")
        self._lbl_breakdown = ctk.CTkLabel(self._done, text="", text_color=MUTED,
                                           anchor="w", font=body_font(12))
        self._lbl_breakdown.grid(row=0, column=1, sticky="w", padx=(16, 0))
        self._lbl_saved = ctk.CTkLabel(self._done, text="", text_color=TEXT,
                                       anchor="w", font=body_font(12))
        self._lbl_saved.grid(row=1, column=1, sticky="w", padx=(16, 0))
        actions = ctk.CTkFrame(self._done, fg_color="transparent")
        actions.grid(row=0, column=2, rowspan=2, sticky="e")
        self._btn_open = make_button(actions, "Open register", self._open_result,
                                     role="primary", width=148, height=40,
                                     font=display_font(14))
        self._btn_open.grid(row=0, column=0, padx=(0, 8))
        self._btn_reveal = make_button(actions, "Show in folder", self._reveal_result,
                                       role="secondary", width=128, height=40)
        self._btn_reveal.grid(row=0, column=1, padx=(0, 8))
        make_button(actions, "Start over", self._enter_setup, role="ghost",
                    width=92, height=40).grid(row=0, column=2)

    #: Width of the settings drawer, in logical units.
    PANEL_W = 380

    def _build_settings_panel(self) -> None:
        """Settings slide in from the right, over the page.

        They are set once and rarely touched again, so they no longer sit
        between the table and the button that builds from it. Where the
        register goes stays on the footer, where it cannot be missed.
        """
        panel = ctk.CTkFrame(self, width=self.PANEL_W, fg_color=CHROME,
                             corner_radius=0, border_width=0)
        panel.grid_propagate(False)
        panel.grid_columnconfigure(1, weight=1)
        self._panel = panel
        # Its left edge is a rule the full height of the window, so the drawer
        # reads as a layer over the table rather than a patch of it.
        self._hairline(panel, vertical=True).grid(row=0, column=0, rowspan=20,
                                                  sticky="ns")
        panel = ctk.CTkFrame(panel, fg_color="transparent")
        panel.grid(row=0, column=1, rowspan=20, sticky="nsew")
        panel.grid_columnconfigure(0, weight=1)
        pad = 22

        top = ctk.CTkFrame(panel, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=pad, pady=(18, 10))
        top.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(top, text="Settings", text_color=TEXT,
                     font=display_font(19)).grid(row=0, column=0, sticky="w")
        make_icon_button(top, "cancel", self._toggle_settings, size=32
                         ).grid(row=0, column=1, sticky="e")

        def section(row: int, title: str) -> ctk.CTkFrame:
            frame = ctk.CTkFrame(panel, fg_color="transparent")
            frame.grid(row=row, column=0, sticky="ew", padx=pad, pady=(12, 0))
            frame.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(frame, text=title, text_color=MUTED, anchor="w",
                         font=body_font(12)).grid(row=0, column=0, columnspan=2,
                                                  sticky="w")
            return frame

        def rule(row: int) -> None:
            self._hairline(panel).grid(row=row, column=0, sticky="ew", padx=pad,
                                       pady=(16, 0))

        template = section(1, "Register template")
        self._lbl_template = ctk.CTkLabel(template, text="", text_color=TEXT,
                                          anchor="w", font=body_font(13))
        self._lbl_template.grid(row=1, column=0, sticky="w", pady=(2, 0))
        self._btn_change = make_button(template, "Change", self._pick_template,
                                       role="secondary", width=86, height=32)
        self._btn_change.grid(row=1, column=1, sticky="e")

        dest = section(2, "Save this round's register in")
        self._lbl_panel_dest = ctk.CTkLabel(dest, text="", text_color=TEXT,
                                            anchor="w", justify="left",
                                            wraplength=self.PANEL_W - 2 * pad - 110,
                                            font=body_font(13))
        self._lbl_panel_dest.grid(row=1, column=0, sticky="w", pady=(2, 0))
        self._btn_dest = make_button(dest, "Choose", self._pick_output_folder,
                                     role="secondary", width=86, height=32)
        self._btn_dest.grid(row=1, column=1, sticky="e")
        self._btn_dest_reset = make_button(dest, "Use the round's folder",
                                           self._use_round_folder, role="ghost",
                                           width=0, height=28)

        rule(3)
        switches = ctk.CTkFrame(panel, fg_color="transparent")
        switches.grid(row=4, column=0, sticky="ew", padx=pad, pady=(12, 0))
        switches.grid_columnconfigure(0, weight=1)
        self._switch_dedup = make_check(switches, "Remove duplicate serial numbers",
                                         self._dedup, self._remember)
        self._switch_dedup.grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(switches, text="Keeps the first form for each serial and "
                     "lists the rest under Details.", text_color=MUTED,
                     anchor="w", justify="left", wraplength=self.PANEL_W - 2 * pad - 32,
                     font=body_font(12)).grid(row=1, column=0, sticky="w",
                                              padx=(28, 0), pady=(2, 12))
        self._switch_forms = make_check(switches, FORMS_ONLY_LABEL, self._forms_only,
                                         self._on_forms_only_toggled)
        self._switch_forms.grid(row=2, column=0, sticky="w")
        ctk.CTkLabel(switches, text="Leaves out device lists, blank templates and "
                     "names without a site code.", text_color=MUTED, anchor="w",
                     justify="left", wraplength=self.PANEL_W - 2 * pad - 32,
                     font=body_font(12)).grid(row=3, column=0, sticky="w",
                                              padx=(28, 0), pady=(2, 0))

        rule(5)
        look = section(6, "Appearance")
        self._appearance = Toggle(look, ("Dark", "Light"),
                                  command=lambda v: self._apply_theme(v.lower()))
        self._appearance.set(ctk.get_appearance_mode())
        self._appearance.grid(row=1, column=0, sticky="w", pady=(6, 0))

        rule(7)
        about = ctk.CTkFrame(panel, fg_color="transparent")
        about.grid(row=8, column=0, sticky="ew", padx=pad, pady=(14, 0))
        ctk.CTkLabel(about, text=f"Calist {calist.__version__}", text_color=TEXT,
                     anchor="w", font=body_font(13, "bold")).grid(row=0, column=0,
                                                                  sticky="w")
        ctk.CTkLabel(about, text=f"Built by {AUTHOR_NAME}", text_color=MUTED,
                     anchor="w", font=body_font(12)).grid(row=1, column=0, sticky="w")
        ctk.CTkLabel(about, text=AUTHOR_EMAIL, text_color=MUTED, anchor="w",
                     font=body_font(12)).grid(row=2, column=0, sticky="w")
        ctk.CTkLabel(about, text="Every register Calist builds is signed with "
                     "this credit.", text_color=FAINT, anchor="w", justify="left",
                     wraplength=self.PANEL_W - 2 * pad, font=body_font(12)
                     ).grid(row=3, column=0, sticky="w", pady=(6, 0))

    @staticmethod
    def _wheel_over(widget):
        """Give `widget` the wheel while the pointer is over it.

        CTkScrollableFrame binds <MouseWheel> with bind_all and scrolls itself
        for any event whose widget chain reaches its canvas — which is every
        widget on the page. A widget binding runs first, so returning "break"
        here keeps the scroll where the user is pointing.

        CustomTkinter 6 exempts its own CTkTextbox from that, but not a ttk
        Treeview, so the device table still needs this.
        """
        def handler(event):
            widget.yview_scroll(-int(event.delta / 120), "units")
            return "break"
        return handler

    def _colour_tags(self) -> None:
        """Row colours. Treeview tags take one colour, not a pair.

        A row that needs a look is tinted amber, one that could not be read
        red — findable while scrolling a thousand rows — and the dot and the
        words say the same thing without colour. Rows that contribute nothing
        (a copy, a file left out, a cancelled read) are dimmed. Every other
        plain row is striped; a row never carries two background tags, because
        which one ttk shows is not the one you would expect.
        """
        self._tree.tag_configure("stripe", background=live(ROW_ALT))
        for tag in ("ready", "ok"):
            self._tree.tag_configure(tag, foreground=live(TEXT))
        self._tree.tag_configure("warn", foreground=live(TEXT),
                                 background=live(WARN_TINT))
        self._tree.tag_configure("error", foreground=live(TEXT),
                                 background=live(DANGER_TINT))
        self._tree.tag_configure("muted", foreground=live(FAINT))
        self._dots = status_dots(self, self.ROW_PX)

    def _apply_theme(self, mode: str) -> None:
        """Switch between light and dark, Tk-native widgets included.

        CustomTkinter re-colours its own widgets from their (light, dark)
        pairs; everything else here was given one resolved colour and has to be
        given the other.
        """
        ctk.set_appearance_mode(mode)
        style_treeview(self)
        self._colour_tags()
        self._table_body.configure(bg=live(SURFACE))
        for line in self._hairlines:
            if line.winfo_exists():
                line.configure(bg=live(BORDER))
        self._refresh_table()
        if self._result is not None:
            self._show_verdict(self._result)
        self._remember()
        self.after(50, lambda: tint_title_bar(self))

    def _attach_logging(self) -> None:
        handler = TkLogHandler(self._log_box)
        handler.setFormatter(StatusFormatter())
        calist.log.addHandler(handler)
        calist.log.setLevel(logging.INFO)
        calist.log.propagate = False

    def _hairline(self, master, vertical: bool = False) -> tk.Frame:
        line = tk.Frame(master, bg=live(BORDER), bd=0, highlightthickness=0,
                        width=1 if vertical else 0, height=0 if vertical else 1)
        self._hairlines.append(line)
        return line

    # ── state transitions ────────────────────────────────────────────────────

    def _show_action(self, which: ctk.CTkFrame) -> None:
        for frame in (self._idle, self._busy, self._done):
            frame.grid_forget()
        which.grid(row=0, column=0, sticky="ew")
        self._on_resize()

    def _enter_setup(self) -> None:
        self._result = None
        self._cancel = None
        self._scan = None
        # Reset any previous run's per-file statuses back to pre-flight, under
        # whatever the current settings are.
        forms_only = bool(self._forms_only.get())
        for path in list(self._files):
            self._files[path] = calist.classify_file(path, forms_only)
        self._filter.set("All")
        self._show_action(self._idle)
        self._set_inputs_enabled(True)
        self._refresh_all()

    def _set_inputs_enabled(self, enabled: bool) -> None:
        """Freeze the inputs while a run is in flight.

        Changing the template or the device set mid-run would describe a build
        that isn't the one actually happening.
        """
        for widget in (self._btn_change, self._btn_dest, self._btn_dest_reset,
                       self._btn_clear, self._btn_slim_folder, self._btn_slim_files,
                       self._btn_recheck):
            set_enabled(widget, enabled)
        for switch in (self._switch_dedup, self._switch_forms):
            switch.configure(state="normal" if enabled else "disabled")

    def _enter_working(self, total: int) -> None:
        self._show_action(self._busy)
        self._set_inputs_enabled(False)
        self._bar.set(0)
        self._lbl_stage.configure(text=f"Reading form 0 of {total:,}")
        self._lbl_eta.configure(text="")
        self._lbl_current.configure(text="")
        set_enabled(self._btn_cancel, True)
        self._btn_cancel.configure(text="Cancel")
        if self._turbo_mode:
            self._lbl_bolt.grid(row=0, column=0, padx=(0, 6))
        else:
            self._lbl_bolt.grid_remove()

    def _enter_scanning(self) -> None:
        """The busy footer again, but with no total — a walk cannot know one
        until it has finished walking."""
        self._show_action(self._busy)
        self._set_inputs_enabled(False)
        self._bar.set(0)
        self._lbl_stage.configure(text="Looking for inspection forms…")
        self._lbl_eta.configure(text="")
        self._lbl_current.configure(text="")
        set_enabled(self._btn_cancel, True)
        self._btn_cancel.configure(text="Cancel")
        self._lbl_bolt.grid_remove()

    def _enter_results(self, result: RunResult) -> None:
        self._result = result
        self._lbl_bolt.grid_remove()
        self._show_action(self._done)
        self._set_inputs_enabled(True)

        built = result.succeeded and not result.cancelled
        set_enabled(self._btn_open, built)
        set_enabled(self._btn_reveal, built)
        if result.cancelled:
            self._lbl_breakdown.configure(
                text=f"Stopped after {result.files_read:,} of "
                     f"{len(result.outcomes):,} forms. Nothing was written.")
            self._lbl_saved.configure(text="")
        elif not result.succeeded:
            self._lbl_breakdown.configure(
                text=result.error or "Open Details to see what went wrong.")
            self._lbl_saved.configure(text="")
        else:
            self._lbl_breakdown.configure(text=result_breakdown(result))
            self._lbl_saved.configure(
                text=f"Saved to {shorten_path(result.output_path, 64)}")
        self._show_verdict(result)

        if result.problems and not self._turbo_mode:
            self._filter.set("Problems")
        self._refresh_all()

    def _show_verdict(self, result: RunResult) -> None:
        """The verdict: a glyph in the verdict colour, and plain words."""
        for child in self._verdict.winfo_children():
            child.destroy()
        if result.cancelled:
            words, colour, glyph = "Cancelled", WARNING, "warning"
        elif not result.succeeded:
            words, colour, glyph = "Not built", DANGER, "error"
        else:
            words, colour, glyph = "Register built", SUCCESS, "check"
        ctk.CTkLabel(self._verdict, text=theme.ICONS[glyph], text_color=colour,
                     font=icon_font(22)).grid(row=0, column=0, padx=(0, 10))
        ctk.CTkLabel(self._verdict, text=words, text_color=TEXT,
                     font=display_font(17)).grid(row=0, column=1, sticky="w")

    # ── file intake ──────────────────────────────────────────────────────────

    #: Forms handed to the main thread per batch during a folder scan. Small
    #: enough that the window stays live on a huge tree, large enough that the
    #: queue is not the bottleneck.
    SCAN_BATCH = 200

    def _add_paths(self, paths: list[str], *, record: bool = True) -> None:
        """Absorb a batch of paths. Folders are walked on a worker thread.

        Each pick is remembered in order (`record`), for the register's
        default folder and for Re-check — which passes record=False, since it
        is walking picks already remembered.
        """
        forms_only = bool(self._forms_only.get())   # threading rule 1: read it here
        before = len(self._files)
        folders, singles = [], []
        for raw in paths:
            path = Path(raw)
            if record and path.exists():
                pick = (str(path), path.is_dir())
                if pick not in self._sources:
                    self._sources.append(pick)
            if path.is_dir():
                folders.append(str(path))
            elif path.is_file():
                # Deliberately unfiltered: a file the user picked by hand earns
                # an answer, even if that answer is "Unsupported format".
                singles.append(str(path))

        added = self._absorb(singles, forms_only)
        if folders:
            self._start_scan(folders, forms_only, before)
        else:
            self._finish_intake(added)

    def _absorb(self, filepaths: list[str], forms_only: bool) -> int:
        added = 0
        for key in filepaths:
            if key not in self._files:
                self._files[key] = calist.classify_file(key, forms_only)
                added += 1
        return added

    def _finish_intake(self, added: int) -> None:
        if added:
            log_ui.debug("Added %d files", added)
            calist.log.info("Added %s. Total: %s.", calist.plural(added, "form"),
                            f"{len(self._files):,}")
        self._enter_setup()
        if added:
            self._log_left_out()

    def _log_left_out(self) -> None:
        """List every file the switch keeps out, and why, under Details.

        Once per intake or toggle rather than on every refresh — the table
        hides these files, so this is the only place they can be seen.
        """
        left = sorted((o for o in self._files.values() if o.status == LEFT_OUT),
                      key=lambda o: o.filename.lower())
        if not left:
            return
        calist.log.info("Left out %s that %s not device forms:",
                        calist.plural(len(left), "file"),
                        "is" if len(left) == 1 else "are")
        for outcome in left:
            calist.log.info("   %s — %s", outcome.filename, outcome.detail)

    def _start_scan(self, folders: list[str], forms_only: bool, before: int) -> None:
        """Walk folders on a worker, so a big or networked tree cannot freeze.

        The walk used to run inline: a stat per entry, the whole listing built
        before anything was filtered, all on the main thread. On a share that
        is a window that stops repainting with no way to stop it.
        """
        cancel = threading.Event()
        self._scan = cancel
        self._scan_from = before
        self._enter_scanning()

        def worker() -> None:
            batch: list[FileOutcome] = []
            try:
                for folder in folders:
                    for filepath in calist.find_source_files(folder, cancel):
                        # Classifying here keeps the main thread free; it is
                        # filename-only, so it opens nothing.
                        batch.append(calist.classify_file(filepath, forms_only))
                        if len(batch) >= self.SCAN_BATCH:
                            self._events.put(("scan", batch, 0, 0))
                            batch = []
                    if cancel.is_set():
                        break
            except Exception as exc:               # never die silently
                calist.log.error("Could not finish scanning: %s", exc)
            if batch:
                self._events.put(("scan", batch, 0, 0))
            self._events.put(("scanned", cancel.is_set(), 0, 0))

        threading.Thread(target=worker, daemon=True).start()
        self._drain()

    def _on_scan_done(self, stopped: bool) -> None:
        self._scan = None
        added = len(self._files) - self._scan_from
        if stopped:
            calist.log.warning("Scan stopped — %s added so far.",
                               calist.plural(added, "form"))
        self._finish_intake(added)

    def _add_folder(self) -> None:
        start = self._settings.get("last_folder", "")
        chosen = filedialog.askdirectory(title="Choose a folder of inspection forms",
                                         initialdir=start or None)
        if chosen:
            self._settings["last_folder"] = chosen
            self._remember()
            self._add_paths([chosen])

    def _add_files(self) -> None:
        chosen = filedialog.askopenfilenames(
            title="Choose inspection forms",
            initialdir=self._settings.get("last_folder") or None,
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")],
        )
        if chosen:
            self._settings["last_folder"] = os.path.dirname(chosen[0])
            self._remember()
            self._add_paths(list(chosen))

    def _enable_drop(self) -> None:                   # pragma: no cover
        """Take drops anywhere on the window.

        Only the root is registered: OLE walks up from the window under the
        pointer to the nearest registered one, so every widget inside counts.
        """
        if not getattr(self, "dnd_ready", False):
            return
        self.drop_target_register(DND_FILES)
        self.dnd_bind("<<DropEnter>>", self._on_drag_enter)
        self.dnd_bind("<<DropPosition>>", self._on_drag_over)
        self.dnd_bind("<<DropLeave>>", self._on_drag_leave)
        self.dnd_bind("<<Drop>>", self._on_drop)

    def _drop_refused(self) -> bool:
        return (self._lock is not None or self._cancel is not None
                or self._scan is not None)

    def _on_drag_enter(self, event):                  # pragma: no cover
        if self._drop_refused():
            return REFUSE_DROP
        # Say where it will go: the empty card, or the table once forms are in.
        target = self._table_wrap if self._files else self._hero
        target.configure(border_color=PRIMARY, border_width=2)
        return event.action

    def _on_drag_over(self, event):                   # pragma: no cover
        return REFUSE_DROP if self._drop_refused() else event.action

    def _on_drag_leave(self, event=None):             # pragma: no cover
        for target in (self._table_wrap, self._hero):
            target.configure(border_color=BORDER, border_width=1)
        return getattr(event, "action", None)

    def _on_drop(self, event) -> None:                 # pragma: no cover
        self._on_drag_leave()
        if self._drop_refused():
            return REFUSE_DROP
        self._add_paths(list(self.tk.splitlist(event.data)))
        return event.action

    #: How long "Clear the list?" waits for its second click.
    CLEAR_CONFIRM_MS = 3000

    def _arm_clear(self) -> None:
        """Clear all asks once, in place.

        It sits beside Re-check and throws away the whole round's list, so a
        slip must not cost the round. No dialog: the button itself asks, and
        quietly stands down if the second click does not come. The files on
        disk are never touched either way.
        """
        if self._clear_job is not None:
            self.after_cancel(self._clear_job)
            self._disarm_clear()
            self._clear_files()
            return
        self._btn_clear.configure(text="Clear the list?", text_color=DANGER)
        self._clear_job = self.after(self.CLEAR_CONFIRM_MS, self._disarm_clear)

    def _disarm_clear(self) -> None:
        self._clear_job = None
        self._btn_clear.configure(text="Clear all", text_color=MUTED)

    def _clear_files(self) -> None:
        self._files.clear()
        self._sources.clear()
        self._outdir.set("")              # a chosen folder is for one round
        self._group = None
        calist.log.info("Cleared all devices.")
        self._enter_setup()

    def _recheck(self) -> None:
        """Walk every folder and file added again, after fixing forms outside.

        Renamed files are found under their new names and deleted ones drop
        out; the round's folder and any chosen save folder stay as they are.
        """
        if not self._sources or self._cancel is not None or self._scan is not None:
            return
        gone = [path for path, _ in self._sources if not Path(path).exists()]
        self._files.clear()
        self._group = None
        calist.log.info("Re-checking what was added (%s)…",
                        calist.plural(len(self._sources), "pick"))
        for path in gone:
            calist.log.warning("No longer there: %s", path)
        self._add_paths([path for path, _ in self._sources], record=False)

    def _remove_selected(self, _event=None) -> None:
        for iid in self._tree.selection():
            self._files.pop(iid, None)
        self._enter_setup()

    def _reveal_selected(self, _event=None) -> None:
        selection = self._tree.selection()
        if selection and Path(selection[0]).exists():
            reveal_in_explorer(Path(selection[0]))

    def _pick_template(self) -> None:
        chosen = filedialog.askopenfilename(
            title="Choose the register template",
            initialdir=os.path.dirname(self._template.get()) or None,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if chosen:
            self._template.set(chosen)
            self._remember()
            calist.log.info("Template selected: %s", os.path.basename(chosen))
            self._refresh_all()

    def _remember(self) -> None:
        # The old filename-check keys are dropped, not kept alongside: an older
        # build reading `strict_names` would enforce the filename format this
        # switch replaced. Without them it reads the check as off. `turbo`
        # went the same way when Turbo became automatic.
        for stale in ("name_check", "strict_names", "turbo"):
            self._settings.pop(stale, None)
        self._settings.update(deduplicate=bool(self._dedup.get()),
                              real_forms_only=bool(self._forms_only.get()),
                              appearance=ctk.get_appearance_mode().lower())

        # A chosen save folder lasts one round, so it is never written down —
        # and one an older build did write is dropped.
        self._settings.pop("output_dir", None)
        if self._outdir.get() and not Path(self._outdir.get()).is_dir():
            self._outdir.set("")

        # Never persist the built-in template. In a frozen build it lives in
        # PyInstaller's temp extraction folder, which is deleted on exit and
        # gets a new random name next launch — so the stored path would be dead
        # before it was ever read back. Leaving the key out means the fallback
        # in _initial_template() resolves it fresh each time.
        chosen = self._template.get()
        if chosen and not self._is_bundled_template(chosen):
            self._settings["template"] = chosen
        else:
            self._settings.pop("template", None)

        save_settings(self._settings)

    @staticmethod
    def _is_bundled_template(path: str) -> bool:
        shipped = calist.bundled_template()
        if not shipped:
            return False
        return os.path.normcase(os.path.abspath(path)) == \
            os.path.normcase(os.path.abspath(str(shipped)))

    # ── Turbo ────────────────────────────────────────────────────────────────

    def _on_forms_only_toggled(self) -> None:
        """Re-check every loaded name against the switch.

        Cheap enough to do inline — the check reads the name alone, so even a
        few thousand devices re-resolve at once and the table updates on the
        same click. Device lists and blank forms named like devices can only
        be told by opening them, so those drop out during the run.
        """
        self._remember()
        on = bool(self._forms_only.get())
        calist.log.info("Real device forms only: %s", "on" if on else "off")
        self._enter_setup()
        if on:
            self._log_left_out()

    # ── search, filters, the drawer ──────────────────────────────────────────

    def _on_search_typed(self, event=None) -> None:
        """Filter as the user types, once they pause — not on every key."""
        if event is not None and event.keysym == "Escape":
            self._clear_search()
            return
        self._show_search_clear(bool(self._search.get()))
        if self._search_job is not None:
            self.after_cancel(self._search_job)
        self._search_job = self.after(120, self._run_search)

    def _run_search(self) -> None:
        self._search_job = None
        self._refresh_table()
        self._refresh_banner()

    def _clear_search(self) -> None:
        self._search.delete(0, "end")
        # CustomTkinter restores a placeholder only on focus-out; a clear made
        # while the box is not focused would otherwise leave it blank.
        restore = getattr(self._search, "_activate_placeholder", None)
        if restore is not None and self.focus_get() is not self._search._entry:
            restore()
        self._show_search_clear(False)
        self._refresh_table()
        self.focus_set()

    def _show_search_clear(self, shown: bool) -> None:
        set_enabled(self._btn_search_clear, shown)
        self._btn_search_clear.configure(text=theme.ICONS["cancel"] if shown else "")

    def _focus_search(self) -> None:
        if self._files and not self._turbo_mode:
            self._search.focus_set()

    def _pick_group(self, group: ui_state.Group | None) -> None:
        """Show only one kind of problem — or, clicked again, everything.

        While a group is picked, All/Problems shows neither: the group is the
        filter, and the status line names it with a way back.
        """
        same = group is not None and self._group is not None \
            and group.key == self._group.key
        self._group = None if same else group
        self._filter.set(None if self._group is not None else "All")
        self._refresh_table()
        self._refresh_banner()

    def _on_filter_picked(self, _value: str) -> None:
        self._group = None
        self._refresh_table()
        self._refresh_banner()

    def _filter_problems(self) -> None:
        self._group = None
        self._filter.set("Problems")
        self._refresh_table()
        self._refresh_banner()

    def _show_all(self) -> None:
        """Drop every filter: the group, Problems, and the search."""
        self._group = None
        self._filter.set("All")
        self._search.delete(0, "end")
        self._clear_search()
        self._refresh_banner()

    def _on_escape(self, _event=None) -> None:
        """Esc backs out one step: the drawer, then a filter, then a run."""
        if self._panel_open:
            self._toggle_settings()
        elif self._search.get() or self._group is not None \
                or self._filter.get() != "All":
            self._show_all()
        else:
            self._cancel_run()

    def _toggle_settings(self) -> None:
        """Slide the settings drawer in or out (about 160 ms, eased)."""
        if self._lock is not None:
            return
        self._panel_open = not self._panel_open
        if self._panel_job is not None:
            self.after_cancel(self._panel_job)
            self._panel_job = None
        if self._panel_open:
            self._refresh_settings()
            self._panel.place(relx=1.0, x=self.PANEL_W, rely=0, relheight=1.0,
                              anchor="ne")
            self._panel.lift()
        self._slide_panel(0)

    #: Frames of the drawer's slide; 8 at 20 ms is about 160 ms.
    SLIDE_FRAMES = 8

    def _slide_panel(self, step: int) -> None:
        # Ease out: fast at first, settling at the end.
        t = (step + 1) / self.SLIDE_FRAMES
        eased = 1 - (1 - t) ** 3
        shown = eased if self._panel_open else 1 - eased
        self._panel.place_configure(x=round(self.PANEL_W * (1 - shown)))
        if step + 1 < self.SLIDE_FRAMES:
            self._panel_job = self.after(20, lambda: self._slide_panel(step + 1))
        else:
            self._panel_job = None
            if not self._panel_open:
                self._panel.place_forget()

    # ── rendering ────────────────────────────────────────────────────────────

    def _refresh_all(self) -> None:
        self._refresh_intake()
        self._refresh_table()
        self._refresh_banner()
        self._refresh_settings()
        self._refresh_hint()

    # ── fitting the page to the window ───────────────────────────────────────
    #
    # A scrollable page is exactly as tall as its contents, so nothing stretches
    # to fill the window on its own. These two hand the leftover space to
    # whichever block can use it — the hero while the app is empty, the device
    # table once it is not — and give it back when the window shrinks, at which
    # point the scrollbar takes over rather than any block being crushed.

    MIN_ROWS, MAX_ROWS = 5, 40
    #: The hero absorbs the spare height down to the footer, so the empty
    #: window is one card, not a card and a dead band. MIN is a little above
    #: the glyph-title-buttons stack.
    MIN_HERO_H, MAX_HERO_H = 300, 1400
    #: Cap on the settle loop below, so a layout that cannot converge on an
    #: exact fit stops oscillating instead of rescheduling itself forever.
    MAX_FIT_PASSES = 6

    def _on_resize(self, event=None) -> None:
        # <Configure> fires many times during a drag, and our own resizing
        # fires it again; coalesce into one pass once the drag settles.
        if event is not None and event.widget is not self:
            return
        self._fit_passes = 0
        if self._fit_job is not None:
            self.after_cancel(self._fit_job)
        self._fit_job = self.after(70, self._fit_to_window)

    def _show_page_scrollbar(self, needed: bool) -> None:
        """The page's scrollbar only when the page actually scrolls: beside the
        table's own it is a second bar and a gutter for nothing."""
        bar = getattr(self._page, "_scrollbar", None)
        if bar is None:
            return
        if needed and not bar.winfo_ismapped():
            bar.grid()
        elif not needed and bar.winfo_ismapped():
            bar.grid_remove()

    def _fit_to_window(self) -> None:
        self._fit_job = None
        if self._lock is not None or not self._page.winfo_ismapped():
            return
        try:
            viewport = self._page._parent_canvas.winfo_height()
        except Exception:                                # pragma: no cover
            return
        if viewport < 120 or self._fit_passes >= self.MAX_FIT_PASSES:
            return
        self._fit_passes += 1

        # Positive slack means unused window; negative means we overflow and
        # the scrollbar is carrying the difference.
        slack = viewport - self._page.winfo_reqheight()
        self._show_page_scrollbar(slack < 0)

        if self._files:
            # A small surplus is left alone, but a small overflow costs a row:
            # it would otherwise bring the page's scrollbar out for a few pixels.
            if 0 <= slack < self.ROW_PX:
                return
            current = int(self._tree.cget("height"))
            wanted = max(self.MIN_ROWS,
                         min(self.MAX_ROWS, current + slack // self.ROW_PX))
            if wanted != current:
                self._tree.configure(height=wanted)
                self._fit_job = self.after(30, self._fit_to_window)
        else:
            if abs(slack) < 10:
                return
            # Two unit systems meet here. CTkFrame.configure(height=) takes
            # CustomTkinter's logical units and cget gives them back, while
            # slack and every winfo_* measurement are device pixels — 1.25x
            # apart on this display. Mixing them made a 380 cap render as 475
            # and kept the loop from ever settling.
            scale = ctk.ScalingTracker.get_widget_scaling(self._hero)
            current = self._hero.cget("height")
            wanted = min(self.MAX_HERO_H,
                         max(self.MIN_HERO_H, current + slack / scale))
            if abs(wanted - current) >= 4:
                self._hero.configure(height=wanted)
                self._fit_job = self.after(30, self._fit_to_window)

    def _refresh_intake(self) -> None:
        """Hand the vertical space to whichever of hero/table matters now."""
        counts = ui_state.count(self._files.values())
        if self._cancel is None:                 # a run keeps what it started as
            self._turbo_mode = ui_state.is_turbo(counts.total)

        # The round is what was added — the same folder the register goes to
        # — and only falls back to the folder the forms share.
        origin = ui_state.origin_folder(self._sources) if self._files else ""
        name = (os.path.basename(origin.rstrip("\/")) if origin else "") or (
            ui_state.round_name(self._files) if self._files else "")
        self.title(f"{name} — Calist" if name else "Calist")
        if name:
            self._round_rule.grid(row=0, column=1, sticky="ns", padx=14, pady=4)
            self._lbl_round.configure(text=name)
            self._lbl_round.grid(row=0, column=2, sticky="w")
        else:
            self._round_rule.grid_remove()
            self._lbl_round.grid_remove()

        if self._files:
            self._hero.grid_remove()
            self._toolbar.grid(row=0, column=0, sticky="ew")
            self._intake.grid_configure(sticky="ew")
            self._intake.grid_rowconfigure(0, weight=0)
            self._table_wrap.grid(row=3, column=0, sticky="nsew", padx=self.PAD)
            # Turbo lists problems only, so there is nothing for All to show.
            if self._turbo_mode:
                self._filter.grid_remove()
            else:
                self._filter.grid()
        else:
            self._toolbar.grid_remove()
            self._table_wrap.grid_remove()
            self._hero.grid()
            self._intake.grid_configure(sticky="nsew")
            self._intake.grid_rowconfigure(0, weight=1)

        self._on_resize()

    def _render_counts(self, parent, counts: ui_state.Counts) -> int:
        """The round's counts, as the start of the status line. Returns the
        next free column."""
        done = counts.read > 0
        items = [(None, f"{counts.total:,} forms"),
                 (SUCCESS if done else FAINT, f"{counts.read if done else counts.ready:,} "
                           f"{'read' if done else 'ready'}")]
        if counts.left_out:
            items.append((FAINT, f"{counts.left_out:,} left out"))
        for column, (colour, words) in enumerate(items):
            chip = ctk.CTkFrame(parent, fg_color="transparent")
            chip.grid(row=0, column=column, padx=(0, 16))
            if colour is not None:
                ctk.CTkFrame(chip, width=8, height=8, corner_radius=4,
                             fg_color=colour).grid(row=0, column=0, padx=(0, 6))
            ctk.CTkLabel(chip, text=words,
                         text_color=TEXT if colour is None else MUTED,
                         font=body_font(13, "bold" if colour is None else "normal")
                         ).grid(row=0, column=1)
        return len(items)

    #: Problem groups shown as chips; the rest are one "more" chip, so a round
    #: with many unknown codes cannot push the line off the window.
    MAX_CHIPS = 4

    def _refresh_banner(self) -> None:
        """The status line: the round's counts, then its problems, grouped,
        then what the table is filtered to."""
        for child in self._banner.winfo_children():
            child.destroy()
        if not self._files:
            self._banner.grid_remove()
            return
        self._banner.grid(row=2, column=0, sticky="ew", padx=self.PAD, pady=(0, 8))

        outcomes = list(self._files.values())
        counts = ui_state.count(outcomes)
        groups = ui_state.problem_groups(outcomes)
        # A filter on a group that has since been fixed must not stick.
        if self._group is not None:
            self._group = next((g for g in groups if g.key == self._group.key), None)
            if self._group is None:
                self._filter.set("All")

        line = ctk.CTkFrame(self._banner, fg_color="transparent")
        line.grid(row=0, column=0, sticky="w")
        column = self._render_counts(line, counts)

        if groups:
            files = sum(g.count for g in groups)
            ctk.CTkLabel(line, text=theme.ICONS["warning"], text_color=CAUTION,
                         font=icon_font(14)).grid(row=0, column=column, padx=(4, 6))
            ctk.CTkLabel(line, text=f"{files:,} need a look", text_color=TEXT,
                         font=body_font(13, "bold")
                         ).grid(row=0, column=column + 1, padx=(0, 10))
            column += 2
            for group in groups[:self.MAX_CHIPS]:
                picked = self._group is not None and group.key == self._group.key
                make_button(line, f"{group.label}   {group.count:,}",
                            lambda g=group: self._pick_group(g),
                            role="picked" if picked else "chip", height=28, width=0
                            ).grid(row=0, column=column, padx=(0, 6))
                column += 1
            rest = groups[self.MAX_CHIPS:]
            if rest:
                more = sum(g.count for g in rest)
                make_button(line, f"{len(rest)} more   {more:,}",
                            lambda: self._filter_problems(), role="chip",
                            height=28, width=0).grid(row=0, column=column)
        elif counts.total:
            ctk.CTkLabel(line, text=theme.ICONS["check"], text_color=SUCCESS,
                         font=icon_font(13)).grid(row=0, column=column, padx=(4, 6))
            ctk.CTkLabel(line, text="Nothing needs a look", text_color=MUTED,
                         font=body_font(13)).grid(row=0, column=column + 1)

        self._render_filter_note()
        self._on_resize()

    def _render_filter_note(self) -> None:
        """The right end of the status line: what the table is showing.

        The frame is made only when it will hold something: an empty CTkFrame
        falls back to 200x200 and would open a band above and below the line.
        """
        counts = ui_state.count(self._files.values())
        shown, total = self._table_counts
        if not (self._turbo_mode or shown != total or counts.copies):
            return
        side = ctk.CTkFrame(self._banner, fg_color="transparent")
        side.grid(row=0, column=1, sticky="e")
        if self._turbo_mode:
            ctk.CTkLabel(side, text=theme.ICONS["bolt"], text_color=TURBO,
                         font=icon_font(13)).grid(row=0, column=0, padx=(0, 6))
            ctk.CTkLabel(side, text=f"Turbo: {ui_state.TURBO_AT:,} forms or more, "
                         "so only problems are listed", text_color=MUTED,
                         font=body_font(12)).grid(row=0, column=1)
            return
        if shown != total:
            what = f"Showing {shown:,} of {total:,}"
            if self._group is not None:
                what += f": {self._group.label}"
            elif self._filter.get() == "Problems":
                what += ": problems only"
            query = " ".join(self._search.get().split())
            if query:
                what += f' matching "{query}"'
            ctk.CTkLabel(side, text=what,
                         text_color=MUTED, font=body_font(12)
                         ).grid(row=0, column=0, padx=(0, 6))
            make_button(side, "Show all", self._show_all, role="ghost", height=28,
                        width=0).grid(row=0, column=1)
            return
        if counts.copies:
            make_button(side, f"{counts.copies:,} "
                        f"cop{'ies' if counts.copies != 1 else 'y'} left out, "
                        "see Details", self._show_details, role="ghost",
                        height=28, width=0).grid(row=0, column=0)

    def _show_details(self) -> None:
        if not self._log_open:
            self._toggle_log()

    def _visible_outcomes(self) -> list[tuple[str, FileOutcome]]:
        return ui_state.visible(
            self._files.items(),
            problems_only=self._turbo_mode or self._filter.get() == "Problems",
            group=self._group, query=self._search.get(), status_text=status_text)

    def _refresh_table(self) -> None:
        self._tree.delete(*self._tree.get_children())
        rows = self._visible_outcomes()

        for index, (path, outcome) in enumerate(rows):
            tag = STATUS_DISPLAY.get(outcome.status, (outcome.status, "muted"))[1]
            # Only plain rows are striped: a problem row's tint is its
            # background, and a second background tag would override it.
            tags = (tag, "stripe") if index % 2 and tag in PLAIN_TAGS else (tag,)
            self._tree.insert(
                "", "end", iid=path, tags=tags, image=self._dots.get(tag, ""),
                values=(outcome.filename, outcome.device_name or "—",
                        outcome.serial or "—", status_text(outcome)))

        total = sum(1 for o in self._files.values() if o.status != LEFT_OUT)
        self._table_counts = (len(rows), total)
        if rows:
            self._empty.grid_remove()
        else:
            self._empty.configure(
                text="No problems so far." if self._turbo_mode
                else "Nothing matches. Clear the search, or choose All."
                if total else "No forms yet.")
            self._empty.grid(row=0, column=0)

    def _destination(self) -> tuple[Path | None, str]:
        """Where the register will land, and any warning about it."""
        if not self._files or not self._template.get():
            return None, ""
        try:
            # resolve_output_path only reads the first entry, and this runs on
            # every refresh — sorting tens of thousands of paths for it twice
            # over is work nobody asked for.
            path = calist.resolve_output_path([min(self._files)],
                                              self._template.get(),
                                              self._save_folder() or None)
        except ValueError as exc:
            return None, str(exc)
        return path, "Replaces the existing file" if path.exists() else ""

    def _save_folder(self) -> str:
        """Where the register goes: the folder chosen for this round, else the
        round's own folder (ui_state.origin_folder)."""
        return self._outdir.get() or ui_state.origin_folder(self._sources)

    def _use_round_folder(self) -> None:
        self._outdir.set("")
        self._refresh_settings()
        self._refresh_hint()

    def _pick_output_folder(self) -> None:
        """Choose where the register is saved.

        Without a choice the register lands beside the first source file, which
        is the original behaviour and stays the default — this only overrides
        it. A folder that has gone missing falls back rather than erroring.
        """
        current = self._outdir.get()
        initial = current if current and Path(current).is_dir() else \
            self._settings.get("last_folder", "")
        chosen = filedialog.askdirectory(
            parent=self, title="Select a folder to save the register in",
            initialdir=initial or None)
        if not chosen:
            return
        self._outdir.set(os.path.normpath(chosen))
        self._remember()
        self._refresh_settings()
        self._refresh_hint()

    def _refresh_settings(self) -> None:
        template = self._template.get()
        if template:
            self._lbl_template.configure(text=os.path.basename(template),
                                         text_color=TEXT)
        else:
            self._lbl_template.configure(text="Not chosen yet", text_color=FAINT)

        path, note = self._destination()
        if path:
            self._lbl_dest.configure(text=shorten_path(path, 70),
                                     text_color=WARNING if note else TEXT)
            self._lbl_dest_note.configure(text=f"({note.lower()})" if note else "")
        elif note:
            self._lbl_dest.configure(text=note, text_color=DANGER)
            self._lbl_dest_note.configure(text="")
        else:
            self._lbl_dest.configure(text="the round's folder, once forms are added",
                                     text_color=FAINT)
            self._lbl_dest_note.configure(text="")

        chosen = self._outdir.get()
        if chosen:
            self._lbl_panel_dest.configure(text=shorten_path(chosen, 40))
            self._btn_dest_reset.grid(row=2, column=0, columnspan=2, sticky="w",
                                      pady=(6, 0))
        else:
            origin = ui_state.origin_folder(self._sources)
            self._lbl_panel_dest.configure(
                text=f"The round's folder ({os.path.basename(origin)})" if origin
                else "The round's folder: the first folder you add")
            self._btn_dest_reset.grid_remove()

    def _refresh_hint(self) -> None:
        if not self._files:
            hint, ready = "Add forms to build a register", False
        elif not self._template.get():
            hint, ready = "Choose a register template in Settings", False
        else:
            path, note = self._destination()
            if path is None:
                hint, ready = note, False
            else:
                usable = sum(1 for o in self._files.values() if o.status == READY)
                if not usable:
                    hint, ready = "No recognised forms to build from", False
                else:
                    hint = (f"{usable:,} form{'s' if usable != 1 else ''} "
                            f"ready. Ctrl+Enter builds")
                    ready = True

        self._lbl_hint.configure(text=hint, text_color=MUTED if ready else FAINT)
        set_enabled(self._btn_build, ready)

    # ── the run ──────────────────────────────────────────────────────────────

    def _start(self) -> None:
        if self._btn_build.cget("state") == "disabled":
            return

        # Read every Tk variable HERE, on the main thread. Tk state must not be
        # touched from the worker — doing so raises "main thread is not in main
        # loop" at best, and corrupts the interpreter at worst.
        # A file the switch already left out stays out of the run too: the
        # worker would only classify it again and log it a second time.
        files = sorted(path for path, outcome in self._files.items()
                       if outcome.status != LEFT_OUT)
        template = self._template.get()
        deduplicate = bool(self._dedup.get())
        forms_only = bool(self._forms_only.get())
        output_dir = self._save_folder() or None
        turbo = ui_state.is_turbo(len(files))
        self._turbo_mode = turbo

        total = len(files)
        cancel = threading.Event()
        self._cancel = cancel
        self._started_at = time.monotonic()
        self._marks = [(0, self._started_at)]
        self._live_problems = []
        self._live_render = 0.0
        self._enter_working(total)
        self._remember()

        # In Turbo only problems and a throttled tick cross the thread
        # boundary. Forty thousand queue items and forty thousand row updates
        # are the whole reason the window used to stop breathing.
        last = [0.0]

        def on_file(outcome: FileOutcome, index: int, _total: int) -> None:
            # Worker thread: queue only, never a widget.
            if not turbo:
                self._events.put(("file", outcome, index, total))
                return
            now = time.monotonic()
            if outcome.is_problem or index == total or now - last[0] >= 0.1:
                last[0] = now
                self._events.put(("file", outcome, index, total))

        def worker() -> None:
            try:
                result = calist.process_files(
                    files, template, deduplicate=deduplicate,
                    real_forms_only=forms_only, output_dir=output_dir,
                    on_file=on_file, cancel=cancel, quiet=turbo,
                )
            except Exception as exc:                   # never die silently
                calist.log.exception("Unexpected failure: %s", exc)
                result = RunResult(error=str(exc))
            self._events.put(("done", result, 0, 0))

        threading.Thread(target=worker, daemon=True).start()
        self._drain()

    def _drain(self) -> None:
        """Main-thread poller for worker events.

        Batching here also keeps the table responsive: a few hundred forms
        arrive as a handful of redraws rather than one per file.
        """
        latest: tuple | None = None
        finished: RunResult | None = None
        scanned = False
        stopped = False
        found = 0

        try:
            while True:
                kind, payload, index, total = self._events.get_nowait()
                if kind == "file":
                    self._files[payload.path] = payload
                    self._update_row(payload)
                    if payload.is_problem:
                        self._live_problems.append(payload)
                    latest = (payload, index, total)
                elif kind == "scan":
                    for outcome in payload:
                        if outcome.path not in self._files:
                            self._files[outcome.path] = outcome
                    found += len(payload)
                elif kind == "scanned":
                    scanned, stopped = True, payload
                else:
                    finished = payload
        except queue.Empty:
            pass

        if found and not scanned:
            self._lbl_current.configure(
                text=f"{len(self._files) - self._scan_from} devices found")

        if latest:
            self._update_progress(*latest)
            # The banner's groups follow the run, twice a second at most.
            if time.monotonic() - self._live_render > 0.5:
                self._live_render = time.monotonic()
                self._refresh_banner()

        if finished is not None:
            self._on_run_done(finished)
        elif scanned:
            self._on_scan_done(stopped)
        elif self._cancel is not None or self._scan is not None:
            self.after(60, self._drain)

    def _update_row(self, outcome: FileOutcome) -> None:
        tag = STATUS_DISPLAY.get(outcome.status, (outcome.status, "muted"))[1]
        if self._turbo_mode:
            # Turbo lists problems only: one that surfaces mid-run is added.
            if outcome.is_problem and not self._tree.exists(outcome.path):
                self._tree.insert("", "end", iid=outcome.path, tags=(tag,),
                                  image=self._dots.get(tag, ""),
                                  values=(outcome.filename, outcome.device_name or "—",
                                          outcome.serial or "—", status_text(outcome)))
            return
        if not self._tree.exists(outcome.path):
            return
        # The stripe belongs to the row's position, and only a plain row has one.
        striped = self._tree.index(outcome.path) % 2 and tag in PLAIN_TAGS
        self._tree.item(outcome.path, tags=(tag, "stripe") if striped else (tag,),
                        image=self._dots.get(tag, ""),
                        values=(outcome.filename, outcome.device_name or "—",
                                outcome.serial or "—", status_text(outcome)))

    #: How many recent progress marks the ETA is derived from. A whole-run
    #: average is skewed for the rest of the run by whatever came first, and a
    #: folder mixing .xls (~21ms a form) with .xlsx (~2ms) skews it badly.
    ETA_WINDOW = 12

    def _update_progress(self, outcome: FileOutcome, index: int, total: int) -> None:
        if not self._turbo_mode and self._tree.exists(outcome.path):
            self._tree.see(outcome.path)

        self._bar.set(index / total if total else 0)
        self._lbl_stage.configure(text=f"Reading form {index:,} of {total:,}")
        self._lbl_current.configure(text=outcome.filename)

        now = time.monotonic()
        self._marks.append((index, now))
        del self._marks[:-self.ETA_WINDOW]

        if index >= total:
            # Every form is read; what is left is sorting, de-duplicating and
            # writing, which on tens of thousands of rows takes long enough
            # that "Reading device 20,001 of 20,001" would look stuck.
            self._lbl_stage.configure(text="Writing the register…")
            self._lbl_eta.configure(text="finishing up")
            return
        if index < 3:
            return

        first_index, first_at = self._marks[0]
        done, span = index - first_index, now - first_at
        rate = done / span if done and span > 0 else 0
        if rate > 0:
            self._lbl_eta.configure(
                text=f"about {human_duration((total - index) / rate)} left"
                     f"   ·   {rate:,.0f}/s")

    def _on_run_done(self, result: RunResult) -> None:
        # Stored by reference, not copied, so this costs a dict slot per file
        # and nothing else — and it is what keeps the table truthful if Turbo
        # is switched off after a run that never drew one.
        for outcome in result.outcomes:
            self._files[outcome.path] = outcome
        self._cancel = None
        self._enter_results(result)

    def _cancel_run(self) -> None:
        """Stop whichever long job is in flight — a build, or a folder scan."""
        for job, stage in ((self._cancel, "Finishing the current device…"),
                           (self._scan, "Stopping the scan…")):
            if job is not None and not job.is_set():
                job.set()
                self._btn_cancel.configure(state="disabled", text="Stopping…")
                self._lbl_stage.configure(text=stage)
                return

    # ── result actions ───────────────────────────────────────────────────────

    def _open_result(self) -> None:
        if self._result and self._result.output_path:
            self._safely(open_file, self._result.output_path)

    def _reveal_result(self) -> None:
        if self._result and self._result.output_path:
            self._safely(reveal_in_explorer, self._result.output_path)

    def _safely(self, action, target: Path) -> None:
        try:
            action(target)
        except Exception as exc:
            messagebox.showerror("Could not open", f"{target}\n\n{exc}")

    # ── misc ─────────────────────────────────────────────────────────────────

    def _row_menu(self, event) -> None:
        """Right-click a form: the next step with a problem file is opening it."""
        row = self._tree.identify_row(event.y)
        if not row or self._cancel is not None:
            return
        if row not in self._tree.selection():
            self._tree.selection_set(row)
        self._menu.tk_popup(event.x_root, event.y_root)

    def _open_selected(self) -> None:
        for path in self._tree.selection()[:1]:
            if Path(path).exists():
                self._safely(open_file, Path(path))

    def _copy_names(self) -> None:
        names = [Path(path).name for path in self._tree.selection()]
        if names:
            self.clipboard_clear()
            self.clipboard_append("\n".join(names))

    def _primary_action(self) -> None:
        if self._done.winfo_ismapped():
            self._open_result()
        elif self._idle.winfo_ismapped():
            self._start()

    def _toggle_log(self) -> None:
        self._log_open = not self._log_open
        if self._log_open:
            self._drawer.grid(row=4, column=0, sticky="ew", padx=self.PAD,
                              pady=(8, 12))
            self._btn_details.configure(text="Hide details")
            # It opens below the table: bring it into view.
            self.after(80, lambda: self._page._parent_canvas.yview_moveto(1.0))
        else:
            self._drawer.grid_forget()
            self._btn_details.configure(text="Details")
        self._on_resize()

    # ── the lock ─────────────────────────────────────────────────────────────

    def show_lock(self) -> None:
        """Cover the window with the PIN panel and take the app out of reach.

        The page and the footer are removed from the layout, not merely hidden
        behind the panel, so nothing back there can be reached by tabbing to
        it. The root is never withdrawn — see the class docstring on LockPanel
        for why that matters.
        """
        if getattr(self, "_lock", None) is not None:
            return

        if self._panel_open:
            self._panel_open = False
            self._panel.place_forget()
        self._page.grid_remove()
        self._footer.grid_remove()
        self._lock = LockPanel(self, self._settings, self._on_unlocked)
        self._lock.grid(row=0, column=0, rowspan=2, sticky="nsew")
        self._lock.take_focus()

    def _on_unlocked(self, panel: "LockPanel") -> None:
        self._settings.update(panel.state_dict)
        save_settings(self._settings)

        panel.destroy()
        self._lock = None
        self._page.grid()
        self._footer.grid()

        self._refresh_all()
        self.after(80, self._fit_to_window)
        calist.log.info("Unlocked for %s", date.today().isoformat())

    def _watch_for_new_day(self) -> None:
        """Re-lock once the calendar date moves on.

        Checked while the app sits open, but never during a build: taking the
        interface away mid-run would throw away the user's work and protects
        nothing, since the run was already authorised this morning.
        """
        if (self._lock is None and self._cancel is None and self._scan is None
                and not access.is_unlocked_today(self._settings)):
            calist.log.warning("A new day has started — the access code is needed again.")
            self.show_lock()

        self.after(NEW_DAY_CHECK_MS, self._watch_for_new_day)

    def _show_about(self) -> None:
        """About lives at the foot of Settings."""
        if not self._panel_open:
            self._toggle_settings()

    def _on_close(self) -> None:
        if self._cancel is not None and not self._cancel.is_set():
            if not messagebox.askokcancel(
                    "Still working",
                    "A register is still being built. Close anyway?"):
                return
            self._cancel.set()
        # A scan is interruptible work nobody would miss, so it goes quietly —
        # but it still has to be told to stop, or its thread keeps walking a
        # network share after the window is gone.
        if self._scan is not None:
            self._scan.set()
        self._remember()
        self.destroy()


def run() -> None:
    """Launch the app, behind the daily lock."""
    # Match Tk's coordinate space to physical pixels so the window is crisp on
    # scaled displays.
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            pass

    ctk.set_appearance_mode(appearance_setting(load_settings()))
    ctk.set_default_color_theme("blue")

    # One window, shown once, never withdrawn — see LockPanel for why that
    # matters with CustomTkinter.
    app = App()
    if not access.is_unlocked_today(app._settings):
        app.show_lock()
    app.mainloop()


if __name__ == "__main__":
    run()
