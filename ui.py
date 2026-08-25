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
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

import access
import calist
from calist import (ATTRIBUTION, AUTHOR_EMAIL, AUTHOR_NAME, BAD_FORMAT,
                    CANCELLED, ERROR, FILENAME_EXAMPLE, OK, READY,
                    UNKNOWN_CODE, UNSUPPORTED, FileOutcome, RunResult)

# Drag-and-drop is a bonus, never a requirement: without tkinterdnd2 the drop
# zone is simply click-only.
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except Exception:                                    # pragma: no cover
    HAS_DND = False

# ──────────────────────────────────────────────────────────────────────────────
# Appearance
# ──────────────────────────────────────────────────────────────────────────────

BG = "#17171b"          # window
SURFACE = "#1f1f25"     # cards
SURFACE_2 = "#26262e"   # table, inputs
BORDER = "#33333d"
TEXT = "#e8e8ee"
MUTED = "#9494a4"
FAINT = "#6a6a78"

PRIMARY = "#4c8dff"
PRIMARY_HOVER = "#3b74d9"
SUCCESS = "#3fb950"
WARNING = "#d9a020"
DANGER = "#f2585f"

FONT = "Segoe UI"
MONO = "Consolas"

#: Row status → (label, treeview tag)
STATUS_DISPLAY = {
    READY: ("Ready", "ready"),
    OK: ("Read", "ok"),
    UNKNOWN_CODE: ("Unknown code", "warn"),
    UNSUPPORTED: ("Unsupported file", "warn"),
    BAD_FORMAT: ("Name format", "warn"),
    ERROR: ("Failed", "error"),
    CANCELLED: ("Cancelled", "muted"),
}

SETTINGS_FILE = (Path(os.environ.get("APPDATA") or Path.home())
                 / "Calist" / "settings.json")

#: How often an open window re-checks whether the calendar date has moved on.
NEW_DAY_CHECK_MS = 30_000


# ──────────────────────────────────────────────────────────────────────────────
# Settings
# ──────────────────────────────────────────────────────────────────────────────

def load_settings() -> dict:
    """Best-effort read of remembered preferences; never raises."""
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


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


def style_treeview() -> None:
    """Make ttk.Treeview match the surrounding CTk surfaces.

    A Treeview rather than stacked CTk frames because this table routinely
    holds hundreds of rows, which native Tk handles and a pile of CTk widgets
    would not.
    """
    style = ttk.Style()
    style.theme_use("clam")

    style.configure(
        "Calist.Treeview",
        background=SURFACE_2, fieldbackground=SURFACE_2, foreground=TEXT,
        rowheight=32, borderwidth=0, relief="flat", font=(FONT, 10),
    )
    style.configure(
        "Calist.Treeview.Heading",
        background=SURFACE, foreground=MUTED, relief="flat",
        borderwidth=0, padding=(10, 8), font=(FONT, 9, "bold"),
    )
    style.map("Calist.Treeview",
              background=[("selected", "#2f4f86")],
              foreground=[("selected", TEXT)])
    style.map("Calist.Treeview.Heading", background=[("active", SURFACE_2)])
    # Drop the default border box.
    style.layout("Calist.Treeview",
                 [("Calist.Treeview.treearea", {"sticky": "nswe"})])

    style.configure("Calist.Vertical.TScrollbar",
                    background=SURFACE_2, troughcolor=SURFACE,
                    bordercolor=SURFACE, arrowcolor=MUTED,
                    borderwidth=0, relief="flat")


# customtkinter needs a mixin to cooperate with tkinterdnd2's root.
if HAS_DND:                                            # pragma: no cover
    class _Root(ctk.CTk, TkinterDnD.DnDWrapper):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self.TkdndVersion = TkinterDnD._require(self)
else:
    _Root = ctk.CTk


# ──────────────────────────────────────────────────────────────────────────────
# Lock screen
# ──────────────────────────────────────────────────────────────────────────────

class LockDialog(ctk.CTkToplevel):
    """Asks for the day's PIN. Modal, and the only way past it is the code.

    Owns no persistence: it reports the updated settings dict back through
    ``result`` and ``state``, and the caller decides what to save.
    """

    DOTS = 4

    def __init__(self, master, state: dict):
        super().__init__(master)

        self.state_dict = dict(state)
        self.result = False
        self._entry = ""

        self.title("Calist")
        self.geometry("400x580")
        self.resizable(False, False)
        self.configure(fg_color=BG)
        self.protocol("WM_DELETE_WINDOW", self._give_up)

        self._build()
        self._refresh_dots()
        self._tick_cooldown()

        # Modal: hold focus until this is answered.
        self.transient(master)
        self.grab_set()
        self.bind("<Key>", self._on_key)
        self.focus_force()

    # ── layout ───────────────────────────────────────────────────────────────

    def _build(self) -> None:
        self.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self, text="Calist", text_color=TEXT,
                     font=ctk.CTkFont(FONT, 26, "bold")
                     ).grid(row=0, column=0, pady=(38, 2))
        ctk.CTkLabel(self, text="Enter today's access code", text_color=MUTED,
                     font=ctk.CTkFont(FONT, 13)).grid(row=1, column=0)

        # PIN dots
        self._dots = ctk.CTkFrame(self, fg_color="transparent")
        self._dots.grid(row=2, column=0, pady=(26, 6))
        self._dot_widgets = []
        for i in range(self.DOTS):
            dot = ctk.CTkFrame(self._dots, width=18, height=18, corner_radius=9,
                               fg_color=SURFACE_2, border_width=1,
                               border_color=BORDER)
            dot.grid(row=0, column=i, padx=9)
            dot.grid_propagate(False)
            self._dot_widgets.append(dot)

        self._message = ctk.CTkLabel(self, text="", text_color=DANGER,
                                     font=ctk.CTkFont(FONT, 12))
        self._message.grid(row=3, column=0, pady=(6, 10))

        # Keypad
        pad = ctk.CTkFrame(self, fg_color="transparent")
        pad.grid(row=4, column=0)
        keys = [("1", 0, 0), ("2", 0, 1), ("3", 0, 2),
                ("4", 1, 0), ("5", 1, 1), ("6", 1, 2),
                ("7", 2, 0), ("8", 2, 1), ("9", 2, 2),
                ("C", 3, 0), ("0", 3, 1), ("<", 3, 2)]
        self._keys = []
        for label, r, c in keys:
            muted = label in ("C", "<")
            btn = ctk.CTkButton(
                pad, text=label, width=82, height=62, corner_radius=12,
                fg_color=SURFACE if muted else SURFACE_2,
                hover_color=BORDER, text_color=MUTED if muted else TEXT,
                font=ctk.CTkFont(FONT, 20 if not muted else 16,
                                 "bold" if not muted else "normal"),
                command=lambda k=label: self._press(k))
            btn.grid(row=r, column=c, padx=6, pady=6)
            self._keys.append(btn)

        self._hint = ctk.CTkLabel(
            self, text=f"The code changes daily. Ask {AUTHOR_NAME} for today's.",
            text_color=FAINT, font=ctk.CTkFont(FONT, 11))
        self._hint.grid(row=5, column=0, pady=(16, 2))
        ctk.CTkLabel(self, text=AUTHOR_EMAIL, text_color=FAINT,
                     font=ctk.CTkFont(FONT, 11)).grid(row=6, column=0, pady=(0, 20))

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
        elif event.keysym == "Escape":
            self._give_up()

    def _submit(self) -> None:
        if len(self._entry) != self.DOTS or self._locked_out():
            return

        if access.verify_pin(self._entry):
            self.state_dict = access.mark_unlocked(self.state_dict)
            self.result = True
            self.grab_release()
            self.destroy()
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
            btn.configure(state="disabled" if wait else "normal")

        if wait:
            self._message.configure(
                text=f"Too many attempts — wait {human_duration(wait)}",
                text_color=WARNING)
            self.after(500, self._tick_cooldown)

    def _give_up(self) -> None:
        self.result = False
        self.grab_release()
        self.destroy()


# ──────────────────────────────────────────────────────────────────────────────
# Application
# ──────────────────────────────────────────────────────────────────────────────

class App(_Root):
    def __init__(self) -> None:
        super().__init__()

        self.title("Calist")
        self.geometry("1020x700")
        self.minsize(880, 600)
        self.configure(fg_color=BG)

        self._settings = load_settings()
        self._files: dict[str, FileOutcome] = {}       # path → latest outcome
        self._template = tk.StringVar(value=self._initial_template())
        self._dedup = tk.BooleanVar(value=self._settings.get("deduplicate", False))
        self._strict = tk.BooleanVar(value=self._settings.get("strict_names", False))
        self._cancel: threading.Event | None = None
        self._events: queue.Queue[tuple] = queue.Queue()
        self._result: RunResult | None = None
        self._started_at = 0.0
        self._log_open = False

        style_treeview()
        self._build()
        self._attach_logging()
        self._enter_setup()

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Control-o>", lambda _e: self._add_folder())
        self.bind("<Control-Return>", lambda _e: self._start())
        self.bind("<Escape>", lambda _e: self._cancel_run())

        # Scheduled, not called: the initial unlock belongs to run(), and
        # checking here as well would raise a second prompt behind the first.
        self.after(NEW_DAY_CHECK_MS, self._watch_for_new_day)

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

    def _build(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        # Row 1 (hero) and row 2 (table) trade the vertical slack between them;
        # _refresh_intake decides which one has it.
        self._build_header()
        self._build_intake()
        self._build_table()
        self._build_settings()
        self._build_action()
        self._build_log_drawer()

    def _build_header(self) -> None:
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.grid(row=0, column=0, sticky="ew", padx=28, pady=(22, 14))
        bar.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(bar, text="Calist", text_color=TEXT,
                     font=ctk.CTkFont(FONT, 26, "bold")).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(bar, text="Compile inspection forms into one equipment register",
                     text_color=MUTED, font=ctk.CTkFont(FONT, 13)
                     ).grid(row=1, column=0, sticky="w", pady=(2, 0))

        buttons = ctk.CTkFrame(bar, fg_color="transparent")
        buttons.grid(row=0, column=1, rowspan=2, sticky="e")

        self._btn_details = ctk.CTkButton(
            buttons, text="Details", width=88, height=32, corner_radius=8,
            fg_color=SURFACE, hover_color=SURFACE_2, text_color=MUTED,
            font=ctk.CTkFont(FONT, 12), command=self._toggle_log,
        )
        self._btn_details.grid(row=0, column=0, padx=(0, 8))

        ctk.CTkButton(
            buttons, text="About", width=76, height=32, corner_radius=8,
            fg_color=SURFACE, hover_color=SURFACE_2, text_color=MUTED,
            font=ctk.CTkFont(FONT, 12), command=self._show_about,
        ).grid(row=0, column=1)

        ctk.CTkLabel(bar, text=f"Built by {AUTHOR_NAME}", text_color=FAINT,
                     font=ctk.CTkFont(FONT, 11)
                     ).grid(row=2, column=1, sticky="e", pady=(4, 0))

    def _build_intake(self) -> None:
        """Adding devices is the whole point of the app, so it leads.

        With nothing loaded the hero fills the window; once devices are in it
        collapses to a slim bar and hands the space to the table.
        """
        self._intake = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=14,
                                    border_width=1, border_color=BORDER)
        self._intake.grid(row=1, column=0, sticky="nsew", padx=28, pady=(0, 14))
        self._intake.grid_columnconfigure(0, weight=1)
        self._intake.grid_rowconfigure(0, weight=1)

        # ── hero ─────────────────────────────────────────────────────────────
        self._hero = ctk.CTkFrame(self._intake, fg_color="transparent")
        self._hero.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        self._hero.grid_columnconfigure(0, weight=1)
        self._hero.grid_rowconfigure(0, weight=1)
        self._hero.grid_rowconfigure(7, weight=1)      # centres the stack

        badge = ctk.CTkFrame(self._hero, fg_color=SURFACE_2, corner_radius=22,
                             width=76, height=76, border_width=1,
                             border_color=BORDER)
        badge.grid(row=1, column=0, pady=(0, 18))
        badge.grid_propagate(False)
        ctk.CTkLabel(badge, text="+", text_color=PRIMARY,
                     font=ctk.CTkFont(FONT, 38, "bold")).place(relx=.5, rely=.46,
                                                               anchor="center")

        ctk.CTkLabel(self._hero, text="Add your devices", text_color=TEXT,
                     font=ctk.CTkFont(FONT, 24, "bold")).grid(row=2, column=0)
        ctk.CTkLabel(
            self._hero,
            text=("Drop a folder here, or browse for one, to build the register"
                  if HAS_DND else
                  "Choose the folder holding your device inspection forms"),
            text_color=MUTED, font=ctk.CTkFont(FONT, 13),
        ).grid(row=3, column=0, pady=(6, 0))

        buttons = ctk.CTkFrame(self._hero, fg_color="transparent")
        buttons.grid(row=4, column=0, pady=(22, 0))
        ctk.CTkButton(buttons, text="Add folder", width=190, height=50,
                      corner_radius=11, fg_color=PRIMARY, hover_color=PRIMARY_HOVER,
                      font=ctk.CTkFont(FONT, 15, "bold"), command=self._add_folder
                      ).pack(side="left", padx=(0, 12))
        ctk.CTkButton(buttons, text="Add single files", width=150, height=50,
                      corner_radius=11, fg_color=SURFACE_2, hover_color=BORDER,
                      text_color=TEXT, font=ctk.CTkFont(FONT, 14),
                      command=self._add_files).pack(side="left")

        ctk.CTkLabel(
            self._hero,
            text="Picking a folder adds every Excel file inside it, "
                 "including files in subfolders.",
            text_color=FAINT, font=ctk.CTkFont(FONT, 12),
        ).grid(row=5, column=0, pady=(20, 0))
        ctk.CTkLabel(
            self._hero,
            text="Each one is checked on arrival, so an unrecognised device "
                 "shows up before you build — not after.",
            text_color=FAINT, font=ctk.CTkFont(FONT, 12),
        ).grid(row=6, column=0, pady=(3, 0))

        # ── slim bar ─────────────────────────────────────────────────────────
        self._drop_slim = ctk.CTkFrame(self._intake, fg_color="transparent")
        self._drop_slim.grid_columnconfigure(2, weight=1)
        self._btn_slim_folder = ctk.CTkButton(
            self._drop_slim, text="+  Add folder", width=124, height=34,
            corner_radius=8, fg_color=PRIMARY, hover_color=PRIMARY_HOVER,
            font=ctk.CTkFont(FONT, 12, "bold"), command=self._add_folder)
        self._btn_slim_folder.grid(row=0, column=0, padx=(0, 8))

        self._btn_slim_files = ctk.CTkButton(
            self._drop_slim, text="+  Add files", width=112, height=34,
            corner_radius=8, fg_color=SURFACE_2, hover_color=BORDER,
            text_color=TEXT, font=ctk.CTkFont(FONT, 12), command=self._add_files)
        self._btn_slim_files.grid(row=0, column=1)

        self._lbl_summary = ctk.CTkLabel(self._drop_slim, text="", text_color=MUTED,
                                         anchor="w", font=ctk.CTkFont(FONT, 12))
        self._lbl_summary.grid(row=0, column=2, sticky="w", padx=16)

        self._btn_clear = ctk.CTkButton(
            self._drop_slim, text="Clear all", width=90, height=34, corner_radius=8,
            fg_color="transparent", hover_color=SURFACE_2, text_color=MUTED,
            font=ctk.CTkFont(FONT, 12), command=self._clear_files,
        )
        self._btn_clear.grid(row=0, column=3)

        if HAS_DND:                                    # pragma: no cover
            self._intake.drop_target_register(DND_FILES)
            self._intake.dnd_bind("<<Drop>>", self._on_drop)

    def _build_table(self) -> None:
        wrap = ctk.CTkFrame(self, fg_color=SURFACE_2, corner_radius=12,
                            border_width=1, border_color=BORDER)
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_rowconfigure(1, weight=1)
        self._table_wrap = wrap

        head = ctk.CTkFrame(wrap, fg_color="transparent")
        head.grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 8))
        head.grid_columnconfigure(1, weight=1)

        self._lbl_table = ctk.CTkLabel(head, text="Devices", text_color=TEXT,
                                       font=ctk.CTkFont(FONT, 13, "bold"))
        self._lbl_table.grid(row=0, column=0, sticky="w")

        self._filter = ctk.CTkSegmentedButton(
            head, values=["All", "Problems"], width=180, height=30,
            font=ctk.CTkFont(FONT, 11), command=lambda _v: self._refresh_table(),
            fg_color=SURFACE, selected_color=PRIMARY,
            selected_hover_color=PRIMARY_HOVER, unselected_color=SURFACE,
            unselected_hover_color=BORDER,
        )
        self._filter.set("All")
        self._filter.grid(row=0, column=2, sticky="e")

        body = tk.Frame(wrap, bg=SURFACE_2, highlightthickness=0, bd=0)
        body.grid(row=1, column=0, sticky="nsew", padx=6, pady=(0, 10))
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(0, weight=1)

        self._tree = ttk.Treeview(
            body, style="Calist.Treeview", show="headings", selectmode="extended",
            columns=("file", "device", "status"),
        )
        self._tree.heading("file", text="FILE", anchor="w")
        self._tree.heading("device", text="DEVICE TYPE", anchor="w")
        self._tree.heading("status", text="STATUS", anchor="w")
        self._tree.column("file", width=300, minwidth=180, anchor="w", stretch=False)
        self._tree.column("device", width=250, minwidth=150, anchor="w", stretch=False)
        self._tree.column("status", width=330, minwidth=160, anchor="w", stretch=True)
        self._tree.grid(row=0, column=0, sticky="nsew")

        bar = ttk.Scrollbar(body, orient="vertical", command=self._tree.yview,
                            style="Calist.Vertical.TScrollbar")
        bar.grid(row=0, column=1, sticky="ns")
        self._tree.configure(yscrollcommand=bar.set)

        self._tree.tag_configure("ready", foreground=MUTED)
        self._tree.tag_configure("ok", foreground=SUCCESS)
        self._tree.tag_configure("warn", foreground=WARNING)
        self._tree.tag_configure("error", foreground=DANGER)
        self._tree.tag_configure("muted", foreground=FAINT)

        self._tree.bind("<Double-1>", self._reveal_selected)
        self._tree.bind("<Delete>", self._remove_selected)

        self._empty = ctk.CTkLabel(
            body, text="No devices yet.\nAdd a folder to get started.",
            text_color=FAINT, font=ctk.CTkFont(FONT, 13), justify="center")

    def _build_settings(self) -> None:
        panel = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=12,
                             border_width=1, border_color=BORDER)
        panel.grid(row=3, column=0, sticky="ew", padx=28, pady=(14, 0))
        panel.grid_columnconfigure(1, weight=1)

        # Template row
        ctk.CTkLabel(panel, text="Template", text_color=MUTED, width=76, anchor="w",
                     font=ctk.CTkFont(FONT, 12)).grid(row=0, column=0, sticky="w",
                                                      padx=(18, 10), pady=(14, 6))
        self._lbl_template = ctk.CTkLabel(panel, text="", text_color=TEXT, anchor="w",
                                          font=ctk.CTkFont(FONT, 12))
        self._lbl_template.grid(row=0, column=1, sticky="w", pady=(14, 6))
        self._btn_change = ctk.CTkButton(
            panel, text="Change", width=84, height=30, corner_radius=8,
            fg_color=SURFACE_2, hover_color=BORDER, text_color=TEXT,
            font=ctk.CTkFont(FONT, 12), command=self._pick_template)
        self._btn_change.grid(row=0, column=2, padx=(10, 18), pady=(14, 6))

        # Destination row — the answer to "where did my file go?"
        ctk.CTkLabel(panel, text="Saves to", text_color=MUTED, width=76, anchor="w",
                     font=ctk.CTkFont(FONT, 12)).grid(row=1, column=0, sticky="w",
                                                      padx=(18, 10), pady=(0, 6))
        self._lbl_dest = ctk.CTkLabel(panel, text="", text_color=TEXT, anchor="w",
                                      font=ctk.CTkFont(FONT, 12))
        self._lbl_dest.grid(row=1, column=1, sticky="w", pady=(0, 6))
        self._btn_dest = ctk.CTkButton(
            panel, text="Open folder", width=110, height=30, corner_radius=8,
            fg_color=SURFACE_2, hover_color=BORDER, text_color=TEXT,
            font=ctk.CTkFont(FONT, 12), command=self._open_destination_folder)
        self._btn_dest.grid(row=1, column=2, padx=(10, 18), pady=(0, 6))

        self._switch_dedup = ctk.CTkSwitch(
            panel, text="Remove duplicate serial numbers", variable=self._dedup,
            font=ctk.CTkFont(FONT, 12), text_color=TEXT, progress_color=PRIMARY,
            button_color=TEXT, fg_color=BORDER, command=self._remember,
        )
        self._switch_dedup.grid(row=2, column=0, columnspan=3, sticky="w",
                                padx=(18, 0), pady=(4, 2))

        self._switch_strict = ctk.CTkSwitch(
            panel, text=f"Accept only filenames like  {FILENAME_EXAMPLE}",
            variable=self._strict, font=ctk.CTkFont(FONT, 12), text_color=TEXT,
            progress_color=PRIMARY, button_color=TEXT, fg_color=BORDER,
            command=self._on_strict_toggled,
        )
        self._switch_strict.grid(row=3, column=0, columnspan=3, sticky="w",
                                 padx=(18, 0), pady=(2, 2))

        self._lbl_format = ctk.CTkLabel(
            panel, text="site code · device code and number · month and year",
            text_color=FAINT, anchor="w", font=ctk.CTkFont(FONT, 11))
        self._lbl_format.grid(row=4, column=0, columnspan=3, sticky="w",
                              padx=(66, 0), pady=(0, 14))

    def _build_action(self) -> None:
        self._action = ctk.CTkFrame(self, fg_color="transparent")
        self._action.grid(row=4, column=0, sticky="ew", padx=28, pady=(14, 18))
        self._action.grid_columnconfigure(0, weight=1)

        # — idle: hint on the left, primary button on the right
        self._idle = ctk.CTkFrame(self._action, fg_color="transparent")
        self._idle.grid_columnconfigure(0, weight=1)
        self._lbl_hint = ctk.CTkLabel(self._idle, text="", text_color=MUTED,
                                      anchor="w", font=ctk.CTkFont(FONT, 12))
        self._lbl_hint.grid(row=0, column=0, sticky="w")
        self._btn_build = ctk.CTkButton(
            self._idle, text="Build register", width=180, height=44,
            corner_radius=10, fg_color=PRIMARY, hover_color=PRIMARY_HOVER,
            font=ctk.CTkFont(FONT, 14, "bold"), command=self._start)
        self._btn_build.grid(row=0, column=1, sticky="e")

        # — working: progress, current file, cancel
        self._busy = ctk.CTkFrame(self._action, fg_color="transparent")
        self._busy.grid_columnconfigure(0, weight=1)
        self._lbl_stage = ctk.CTkLabel(self._busy, text="", text_color=TEXT,
                                       anchor="w", font=ctk.CTkFont(FONT, 13, "bold"))
        self._lbl_stage.grid(row=0, column=0, sticky="w")
        self._lbl_eta = ctk.CTkLabel(self._busy, text="", text_color=MUTED,
                                     anchor="e", font=ctk.CTkFont(FONT, 12))
        self._lbl_eta.grid(row=0, column=1, sticky="e", padx=(0, 12))
        self._btn_cancel = ctk.CTkButton(
            self._busy, text="Cancel", width=100, height=36, corner_radius=9,
            fg_color=SURFACE_2, hover_color=DANGER, text_color=TEXT,
            font=ctk.CTkFont(FONT, 12, "bold"), command=self._cancel_run)
        self._btn_cancel.grid(row=0, column=2, rowspan=2, sticky="e")

        self._bar = ctk.CTkProgressBar(self._busy, height=8, corner_radius=4,
                                       progress_color=PRIMARY, fg_color=SURFACE_2)
        self._bar.set(0)
        self._bar.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0),
                       padx=(0, 12))
        self._lbl_current = ctk.CTkLabel(self._busy, text="", text_color=FAINT,
                                         anchor="w", font=ctk.CTkFont(MONO, 11))
        self._lbl_current.grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # — results card
        self._done = ctk.CTkFrame(self._action, fg_color=SURFACE, corner_radius=12,
                                  border_width=1, border_color=BORDER)
        self._done.grid_columnconfigure(1, weight=1)
        self._lbl_verdict = ctk.CTkLabel(self._done, text="", text_color=TEXT,
                                         anchor="w", font=ctk.CTkFont(FONT, 15, "bold"))
        self._lbl_verdict.grid(row=0, column=0, columnspan=2, sticky="w",
                               padx=18, pady=(14, 0))
        self._lbl_breakdown = ctk.CTkLabel(self._done, text="", text_color=MUTED,
                                           anchor="w", font=ctk.CTkFont(FONT, 12))
        self._lbl_breakdown.grid(row=1, column=0, columnspan=2, sticky="w",
                                 padx=18, pady=(2, 8))
        self._lbl_saved = ctk.CTkLabel(self._done, text="", text_color=TEXT,
                                       anchor="w", font=ctk.CTkFont(MONO, 11))
        self._lbl_saved.grid(row=2, column=0, columnspan=2, sticky="w",
                             padx=18, pady=(0, 12))

        actions = ctk.CTkFrame(self._done, fg_color="transparent")
        actions.grid(row=3, column=0, columnspan=2, sticky="ew", padx=18, pady=(0, 16))
        actions.grid_columnconfigure(2, weight=1)
        self._btn_open = ctk.CTkButton(
            actions, text="Open register", width=150, height=38, corner_radius=9,
            fg_color=PRIMARY, hover_color=PRIMARY_HOVER,
            font=ctk.CTkFont(FONT, 13, "bold"), command=self._open_result)
        self._btn_open.grid(row=0, column=0, padx=(0, 10))
        self._btn_reveal = ctk.CTkButton(
            actions, text="Reveal in folder", width=150, height=38, corner_radius=9,
            fg_color=SURFACE_2, hover_color=BORDER, text_color=TEXT,
            font=ctk.CTkFont(FONT, 13), command=self._reveal_result)
        self._btn_reveal.grid(row=0, column=1)
        ctk.CTkButton(actions, text="Start over", width=120, height=38,
                      corner_radius=9, fg_color="transparent", hover_color=SURFACE_2,
                      text_color=MUTED, font=ctk.CTkFont(FONT, 13),
                      command=self._enter_setup).grid(row=0, column=3, sticky="e")

    def _build_log_drawer(self) -> None:
        self._drawer = ctk.CTkFrame(self, fg_color=SURFACE_2, corner_radius=12,
                                    border_width=1, border_color=BORDER)
        self._drawer.grid_columnconfigure(0, weight=1)
        self._drawer.grid_rowconfigure(0, weight=1)

        self._log_box = ctk.CTkTextbox(
            self._drawer, height=180, fg_color=SURFACE_2, text_color=MUTED,
            font=ctk.CTkFont(MONO, 11), wrap="none", activate_scrollbars=True)
        self._log_box.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self._log_box.configure(state="disabled")

    def _attach_logging(self) -> None:
        handler = TkLogHandler(self._log_box)
        handler.setFormatter(StatusFormatter())
        calist.log.addHandler(handler)
        calist.log.setLevel(logging.INFO)
        calist.log.propagate = False

    # ── state transitions ────────────────────────────────────────────────────

    def _show_action(self, which: ctk.CTkFrame) -> None:
        for frame in (self._idle, self._busy, self._done):
            frame.grid_forget()
        which.grid(row=0, column=0, sticky="ew")

    def _enter_setup(self) -> None:
        self._result = None
        self._cancel = None
        # Reset any previous run's per-file statuses back to pre-flight, under
        # whatever the current settings are.
        strict = bool(self._strict.get())
        for path in list(self._files):
            self._files[path] = calist.classify_file(path, strict)
        self._filter.set("All")
        self._show_action(self._idle)
        self._set_inputs_enabled(True)
        self._refresh_all()

    def _set_inputs_enabled(self, enabled: bool) -> None:
        """Freeze the inputs while a run is in flight.

        Changing the template or the device set mid-run would describe a build
        that isn't the one actually happening.
        """
        state = "normal" if enabled else "disabled"
        for widget in (self._btn_change, self._switch_dedup, self._switch_strict,
                       self._btn_clear, self._btn_slim_folder, self._btn_slim_files):
            widget.configure(state=state)

    def _enter_working(self, total: int) -> None:
        self._show_action(self._busy)
        self._set_inputs_enabled(False)
        self._bar.set(0)
        self._lbl_stage.configure(text=f"Reading device 0 of {total}")
        self._lbl_eta.configure(text="")
        self._lbl_current.configure(text="")
        self._btn_cancel.configure(state="normal", text="Cancel")

    def _enter_results(self, result: RunResult) -> None:
        self._result = result
        self._show_action(self._done)
        self._set_inputs_enabled(True)

        if result.cancelled:
            self._lbl_verdict.configure(text="Cancelled", text_color=WARNING)
            self._lbl_breakdown.configure(
                text=f"Stopped after {result.files_read} of "
                     f"{len(result.outcomes)} devices. Nothing was written.")
            self._lbl_saved.configure(text="")
            self._btn_open.configure(state="disabled")
            self._btn_reveal.configure(state="disabled")
        elif not result.succeeded:
            self._lbl_verdict.configure(text="Could not build the register",
                                        text_color=DANGER)
            self._lbl_breakdown.configure(
                text=result.error or "See Details for what went wrong.")
            self._lbl_saved.configure(text="")
            self._btn_open.configure(state="disabled")
            self._btn_reveal.configure(state="disabled")
        else:
            bits = [f"{result.files_read} read"]
            if result.second_rows_added:
                bits.append(f"{result.second_rows_added} module rows added")
            if result.problems:
                bits.append(f"{len(result.problems)} skipped")
            if result.duplicates_removed:
                bits.append(f"{result.duplicates_removed} duplicates removed")

            self._lbl_verdict.configure(text="Register built", text_color=SUCCESS)
            self._lbl_breakdown.configure(
                text=f"{result.rows_written} rows from "
                     f"{len(result.outcomes)} devices  ·  " + "  ·  ".join(bits))
            self._lbl_saved.configure(
                text=f"Saved to   {shorten_path(result.output_path, 72)}")
            self._btn_open.configure(state="normal")
            self._btn_reveal.configure(state="normal")

        if result.problems:
            self._filter.set("Problems")
        self._refresh_all()

    # ── file intake ──────────────────────────────────────────────────────────

    def _add_paths(self, paths: list[str]) -> None:
        """Classify and absorb a batch of paths; folders are scanned."""
        added = 0
        strict = bool(self._strict.get())
        for raw in paths:
            path = Path(raw)
            if path.is_dir():
                for found in sorted(path.rglob("*")):
                    if (found.is_file()
                            and found.suffix.lower() in calist.SUPPORTED_EXTENSIONS
                            and found.name.lower() != calist.OUTPUT_NAME.lower()):
                        key = str(found)
                        if key not in self._files:
                            self._files[key] = calist.classify_file(key, strict)
                            added += 1
            elif path.is_file():
                key = str(path)
                if key not in self._files:
                    self._files[key] = calist.classify_file(key, strict)
                    added += 1

        if added:
            log_ui.debug("Added %d file(s)", added)
            calist.log.info("Added %d device(s). Total: %d", added, len(self._files))
        self._enter_setup()

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

    def _on_drop(self, event) -> None:                 # pragma: no cover
        self._add_paths(list(self.tk.splitlist(event.data)))

    def _clear_files(self) -> None:
        self._files.clear()
        calist.log.info("Cleared all devices.")
        self._enter_setup()

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
        self._settings.update(template=self._template.get(),
                              deduplicate=bool(self._dedup.get()),
                              strict_names=bool(self._strict.get()))
        save_settings(self._settings)

    def _on_strict_toggled(self) -> None:
        """Re-check every loaded name against the new setting, immediately.

        Cheap enough to do inline — the format check is a single precompiled
        match, so even a few hundred devices re-resolve in well under a
        millisecond and the table updates on the same click.
        """
        self._remember()
        calist.log.info("Filename format check %s",
                        "on" if self._strict.get() else "off")
        self._enter_setup()

    # ── rendering ────────────────────────────────────────────────────────────

    def _refresh_all(self) -> None:
        self._refresh_intake()
        self._refresh_table()
        self._refresh_settings()
        self._refresh_hint()

    def _refresh_intake(self) -> None:
        """Hand the vertical space to whichever of hero/table matters now."""
        if self._files:
            self._hero.grid_remove()
            self._drop_slim.grid(row=0, column=0, sticky="ew", padx=16, pady=12)
            self._intake.grid_configure(sticky="ew")
            self._intake.grid_rowconfigure(0, weight=0)
            self.grid_rowconfigure(1, weight=0)
            self.grid_rowconfigure(2, weight=1)
            self._table_wrap.grid(row=2, column=0, sticky="nsew", padx=28)

            ready = sum(1 for o in self._files.values() if o.status == READY)
            done = sum(1 for o in self._files.values() if o.status == OK)
            problems = sum(1 for o in self._files.values() if o.is_problem)

            counted = f"{done} read" if done else f"{ready} recognised"
            text = f"{len(self._files)} devices  ·  {counted}"
            if problems:
                text += f"  ·  {problems} need attention"
            self._lbl_summary.configure(
                text=text, text_color=WARNING if problems else MUTED)
        else:
            self._drop_slim.grid_remove()
            self._table_wrap.grid_remove()
            self._hero.grid()
            self._intake.grid_configure(sticky="nsew")
            self._intake.grid_rowconfigure(0, weight=1)
            self.grid_rowconfigure(1, weight=1)
            self.grid_rowconfigure(2, weight=0)

    def _visible_outcomes(self) -> list[tuple[str, FileOutcome]]:
        items = sorted(self._files.items(), key=lambda kv: kv[1].filename.lower())
        if self._filter.get() == "Problems":
            items = [kv for kv in items if kv[1].is_problem]
        return items

    def _refresh_table(self) -> None:
        self._tree.delete(*self._tree.get_children())
        rows = self._visible_outcomes()

        for path, outcome in rows:
            label, tag = STATUS_DISPLAY.get(outcome.status, (outcome.status, "muted"))
            detail = f"{label} — {outcome.detail}" if outcome.detail else label
            self._tree.insert(
                "", "end", iid=path, tags=(tag,),
                values=(outcome.filename, outcome.device_name or "—", detail))

        if rows:
            self._empty.grid_remove()
        else:
            self._empty.grid(row=0, column=0)

        if self._filter.get() == "Problems":
            self._lbl_table.configure(text=f"Devices  ({len(rows)} needing attention)")
        else:
            self._lbl_table.configure(text=f"Devices  ({len(self._files)})")

    def _destination(self) -> tuple[Path | None, str]:
        """Where the register will land, and any warning about it."""
        if not self._files or not self._template.get():
            return None, ""
        try:
            path = calist.resolve_output_path(sorted(self._files),
                                              self._template.get())
        except ValueError as exc:
            return None, str(exc)
        return path, "Replaces the existing file" if path.exists() else ""

    def _refresh_settings(self) -> None:
        template = self._template.get()
        if template:
            self._lbl_template.configure(text=os.path.basename(template),
                                         text_color=TEXT)
        else:
            self._lbl_template.configure(text="Not chosen yet", text_color=FAINT)

        path, note = self._destination()
        if path:
            text = shorten_path(path)
            if note:
                text += f"      ({note})"
            self._lbl_dest.configure(text=text, text_color=WARNING if note else TEXT)
            self._btn_dest.configure(state="normal")
        elif note:
            self._lbl_dest.configure(text=note, text_color=DANGER)
            self._btn_dest.configure(state="disabled")
        else:
            self._lbl_dest.configure(
                text="Shown once devices and a template are chosen", text_color=FAINT)
            self._btn_dest.configure(state="disabled")

    def _refresh_hint(self) -> None:
        if not self._files:
            hint, ready = "Add devices to continue", False
        elif not self._template.get():
            hint, ready = "Choose a register template to continue", False
        else:
            path, note = self._destination()
            if path is None:
                hint, ready = note, False
            else:
                usable = sum(1 for o in self._files.values() if o.status == READY)
                if not usable:
                    hint, ready = "No recognised devices to build from", False
                else:
                    hint = (f"Ready to build the register from {usable} device"
                            f"{'s' if usable != 1 else ''}")
                    ready = True

        self._lbl_hint.configure(text=hint, text_color=MUTED if ready else FAINT)
        self._btn_build.configure(state="normal" if ready else "disabled")

    # ── the run ──────────────────────────────────────────────────────────────

    def _start(self) -> None:
        if self._btn_build.cget("state") == "disabled":
            return

        # Read every Tk variable HERE, on the main thread. Tk state must not be
        # touched from the worker — doing so raises "main thread is not in main
        # loop" at best, and corrupts the interpreter at worst.
        files = sorted(self._files)
        template = self._template.get()
        deduplicate = bool(self._dedup.get())
        strict_names = bool(self._strict.get())

        total = len(files)
        cancel = threading.Event()
        self._cancel = cancel
        self._started_at = time.monotonic()
        self._enter_working(total)
        self._remember()

        def on_file(outcome: FileOutcome, index: int, _total: int) -> None:
            # Worker thread: queue only, never a widget.
            self._events.put(("file", outcome, index, total))

        def worker() -> None:
            try:
                result = calist.process_files(
                    files, template, deduplicate=deduplicate,
                    strict_names=strict_names, on_file=on_file, cancel=cancel,
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

        try:
            while True:
                kind, payload, index, total = self._events.get_nowait()
                if kind == "file":
                    self._files[payload.path] = payload
                    self._update_row(payload)
                    latest = (payload, index, total)
                else:
                    finished = payload
        except queue.Empty:
            pass

        if latest:
            self._update_progress(*latest)

        if finished is not None:
            self._on_run_done(finished)
        elif self._cancel is not None:
            self.after(60, self._drain)

    def _update_row(self, outcome: FileOutcome) -> None:
        if not self._tree.exists(outcome.path):
            return
        label, tag = STATUS_DISPLAY.get(outcome.status, (outcome.status, "muted"))
        detail = f"{label} — {outcome.detail}" if outcome.detail else label
        self._tree.item(outcome.path, tags=(tag,),
                        values=(outcome.filename, outcome.device_name or "—", detail))

    def _update_progress(self, outcome: FileOutcome, index: int, total: int) -> None:
        if self._tree.exists(outcome.path):
            self._tree.see(outcome.path)

        self._bar.set(index / total if total else 0)
        self._lbl_stage.configure(text=f"Reading device {index} of {total}")
        self._lbl_current.configure(text=outcome.filename)

        elapsed = time.monotonic() - self._started_at
        if 3 <= index < total:
            remaining = (elapsed / index) * (total - index)
            self._lbl_eta.configure(text=f"about {human_duration(remaining)} left")
        elif index >= total:
            self._lbl_eta.configure(text="finishing up")

    def _on_run_done(self, result: RunResult) -> None:
        for outcome in result.outcomes:
            self._files[outcome.path] = outcome
        self._cancel = None
        self._enter_results(result)

    def _cancel_run(self) -> None:
        if self._cancel is not None and not self._cancel.is_set():
            self._cancel.set()
            self._btn_cancel.configure(state="disabled", text="Stopping…")
            self._lbl_stage.configure(text="Finishing the current device…")

    # ── result actions ───────────────────────────────────────────────────────

    def _open_result(self) -> None:
        if self._result and self._result.output_path:
            self._safely(open_file, self._result.output_path)

    def _reveal_result(self) -> None:
        if self._result and self._result.output_path:
            self._safely(reveal_in_explorer, self._result.output_path)

    def _open_destination_folder(self) -> None:
        path, _ = self._destination()
        if path is None:
            return
        target = path if path.exists() else path.parent
        self._safely(reveal_in_explorer if path.exists() else open_file, target)

    def _safely(self, action, target: Path) -> None:
        try:
            action(target)
        except Exception as exc:
            messagebox.showerror("Could not open", f"{target}\n\n{exc}")

    # ── misc ─────────────────────────────────────────────────────────────────

    def _toggle_log(self) -> None:
        self._log_open = not self._log_open
        if self._log_open:
            self.grid_rowconfigure(5, weight=0)
            self._drawer.grid(row=5, column=0, sticky="ew", padx=28, pady=(0, 18))
            self._btn_details.configure(text="Hide details")
        else:
            self._drawer.grid_forget()
            self._btn_details.configure(text="Details")

    def _watch_for_new_day(self) -> None:
        """Re-lock once the calendar date moves on.

        Checked while the app sits open, but never during a build: taking the
        window away mid-run would throw away the user's work and protects
        nothing, since the run was already authorised this morning.
        """
        if not access.is_unlocked_today(self._settings) and self._cancel is None:
            calist.log.warning("A new day has started — the access code is needed again.")
            self.withdraw()
            if unlock(self):
                self.deiconify()
                self.lift()
            else:
                self.destroy()
                return

        self.after(NEW_DAY_CHECK_MS, self._watch_for_new_day)

    def _show_about(self) -> None:
        messagebox.showinfo(
            "About Calist",
            f"Calist — compile device inspection forms into one equipment register.\n\n"
            f"Built by {AUTHOR_NAME}\n{AUTHOR_EMAIL}\n\n"
            f"Every register Calist produces is signed with this attribution.",
            parent=self)

    def _on_close(self) -> None:
        if self._cancel is not None and not self._cancel.is_set():
            if not messagebox.askokcancel(
                    "Still working",
                    "A register is still being built. Close anyway?"):
                return
            self._cancel.set()
        self._remember()
        self.destroy()


def unlock(app: "App") -> bool:
    """Show the lock screen unless today's code has already been entered.

    Returns False when the user closed it without unlocking, which means the
    app should not open at all.
    """
    if access.is_unlocked_today(app._settings):
        return True

    dialog = LockDialog(app, app._settings)
    app.wait_window(dialog)

    # Keep whatever the dialog recorded — the unlock stamp on success, the
    # failure count and cooldown otherwise, so closing the window is not a way
    # to shed a penalty.
    app._settings.update(dialog.state_dict)
    save_settings(app._settings)
    return dialog.result


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

    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    app = App()
    app.withdraw()                     # stay hidden until the code is accepted
    if not unlock(app):
        app.destroy()
        return

    app.deiconify()
    app.lift()
    app.mainloop()


if __name__ == "__main__":
    run()
