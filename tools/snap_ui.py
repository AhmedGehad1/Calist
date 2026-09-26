"""Render every screen of Calist to PNG, for design review.

    python tools/snap_ui.py                         # dark, 1366x768 laptop, all states
    python tools/snap_ui.py --theme light --size 1920x1080 --scale 1.25
    python tools/snap_ui.py --state results --label before-ctk6

Development only: it needs Pillow (for ImageGrab), which is deliberately not a
runtime dependency. Every row it shows is invented — no real form, serial or
site ever reaches a screenshot, because these end up in the README.

The window is sized and scaled to *emulate* the target screen rather than to
match this one: a 1366x768 laptop runs at 100%, so on a 125% monitor the
widget scaling, window scaling and Tk's point size are all pulled back to
1.0. Otherwise the review would show a roomier layout than the engineers get.

It never touches the real settings file (a temporary one stands in, already
unlocked for today) and never calls _on_close, so nothing it does is saved.
"""

from __future__ import annotations

import argparse
import ctypes
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

#: Where screenshots go. Ignored by git: review material, not documentation.
OUT = ROOT / "docs" / "review"

#: Taskbar plus caption: what a maximised window loses from the screen height.
CHROME_H = 71

ROUND = Path(r"C:\Rounds\January round")

#: (folder, filename, serial). Codes are real device codes so the pre-flight
#: resolves them; sites, numbers and serials are invented.
DEVICES = [
    ("Ward", "W01-AC004-0925.xlsx", "DF-20481"),
    ("Ward", "W01-AC005-0925.xlsx", "DF-20517"),
    ("Ward", "W01-AGH011-0925.xlsx", "PM-88213"),
    ("Ward", "W01-AGH012-0925.xlsx", "PM-88240"),
    ("Ward", "W01-BZ014-0925.xlsx", "SY-40177"),
    ("Ward", "W01-BZ015-0925.xlsx", "SY-40178"),
    ("Ward", "W01-AO021-0925.xlsx", "TH-1180"),
    ("Lab", "L02-DG003-0925.xlsx", "CB-7735"),
    ("Lab", "L02-EP006-0925.xlsx", "RF-5502"),
    ("Lab", "L02-BB006-0925.xlsx", "US-66120"),
    ("Theatre", "T03-AA010-0925.xlsx", "AN-31907"),
    ("Theatre", "T03-AE011-0925.xlsx", "ES-9021"),
    ("Theatre", "T03-AM009-0925.xlsx", "VT-12094"),
    ("Theatre", "T03-AB002-0925.xlsx", "VP-4410"),
    ("NICU", "N04-CF005-0925.xlsx", "BW-3316"),
    ("NICU", "N04-AK002-0925.xlsx", "BI-8820"),
    ("NICU", "N04-AL008-0925.xlsx", "PT-2291"),
    ("NICU", "N04-FC012-0925.xlsx", "HF-0457"),
    ("ICU", "I05-AGH020-0925.xlsx", "PM-88301"),
    ("ICU", "I05-AGH021-0925.xlsx", "PM-88302"),
    ("ICU", "I05-AM014-0925.xlsx", "VT-12188"),
    ("ICU", "I05-AF007-0925.xlsx", "EC-55012"),
    ("ICU", "I05-AF008-0925.xlsx", "EC-55013"),
    ("ICU", "I05-BZ030-0925.xlsx", "SY-40290"),
    # Problems — the unknown codes, an unsupported file, a failure, a copy.
    ("Ward", "W01-QQQ100-0925.xlsx", ""),
    ("Ward", "W01-QQQ101-0925.xlsx", ""),
    ("Lab", "L02-ZZZ099-0925.xlsx", ""),
    ("Lab", "L02-BB007-0925.docx", ""),
    ("Theatre", "T03-AE012-0925.xlsx", ""),         # fails to read
    ("ICU", "I05-AF007-0925 (2).xlsx", "EC-55012"),  # a copy
]

FAILED = "T03-AE012-0925.xlsx"
COPY = "I05-AF007-0925 (2).xlsx"


# ── environment ──────────────────────────────────────────────────────────────

def system_dpi() -> float:
    """The monitor's scaling, read the same way Tk will see it."""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return ctypes.windll.user32.GetDpiForSystem() / 96
    except Exception:
        return 1.0


def window_bounds(app) -> tuple[int, int, int, int]:
    """The visible window, title bar included, in physical pixels."""
    hwnd = ctypes.windll.user32.GetParent(app.winfo_id())
    rect = (ctypes.c_long * 4)()
    # DWMWA_EXTENDED_FRAME_BOUNDS: the frame as drawn, without the invisible
    # resize borders GetWindowRect would include.
    if ctypes.windll.dwmapi.DwmGetWindowAttribute(
            hwnd, 9, ctypes.byref(rect), ctypes.sizeof(rect)) == 0:
        return tuple(rect)
    return (app.winfo_rootx(), app.winfo_rooty(),
            app.winfo_rootx() + app.winfo_width(),
            app.winfo_rooty() + app.winfo_height())


def settle(app, seconds: float = 0.9) -> None:
    """Let every after() pass run — _fit_to_window takes several."""
    end = time.monotonic() + seconds
    while time.monotonic() < end:
        app.update()
        time.sleep(0.02)


# ── fake data ────────────────────────────────────────────────────────────────

def paths() -> list[str]:
    return [str(ROUND / folder / name) for folder, name, _ in DEVICES]


def serial_of(filename: str) -> str:
    known = next((s for _, name, s in DEVICES if name == filename), None)
    if known is not None:
        return known
    import zlib                            # a generated name: invent one, stably
    return f"SN-{zlib.crc32(filename.encode()) % 90000 + 10000}"


def read_outcome(calist, path: str):
    """What the run would report for one fake file."""
    outcome = calist.classify_file(path)
    name = outcome.filename
    if outcome.status != calist.READY:
        return outcome
    if name == FAILED:
        outcome.status = calist.ERROR
        outcome.detail = "the workbook is password-protected"
    elif name == COPY:
        outcome.status = calist.COPY
        outcome.detail = "same device as I05-AF007-0925.xlsx"
    else:
        outcome.status = calist.OK
        outcome.serial = serial_of(name)
        outcome.rows = 2 if "AGH" in name else 1
    return outcome


def fake_result(calist, outcomes):
    from calist import Duplicate, RunResult
    rows = sum(o.rows for o in outcomes if o.status == calist.OK)
    return RunResult(
        output_path=ROUND / calist.OUTPUT_NAME,
        rows_written=rows - 1,
        outcomes=outcomes,
        duplicates_removed=1,
        duplicates=[Duplicate("PM-88301", "Patient Monitor",
                              "I05-AGH021-0925.xlsx", "I05-AGH020-0925.xlsx")],
        second_rows_added=sum(1 for o in outcomes
                              if o.status == calist.OK and o.rows == 2),
    )


# ── states ───────────────────────────────────────────────────────────────────

def load(app, calist) -> None:
    app._files.clear()
    app._sources[:] = [(str(ROUND), True)]      # "Add folder" on the round
    for path in paths():
        app._files[path] = calist.classify_file(path)
    app._enter_setup()


def st_empty(app, calist):
    app._files.clear()
    app._enter_setup()


def st_loaded(app, calist):
    load(app, calist)


def st_problems(app, calist):
    load(app, calist)
    app._filter.set("Problems")
    app._refresh_all()


def st_working(app, calist):
    load(app, calist)
    files = sorted(app._files)
    total = len(files)
    app._enter_working(total)
    app._started_at = time.monotonic() - 3
    app._marks = [(0, app._started_at)]
    done = int(total * 0.6)
    for index, path in enumerate(files[:done], start=1):
        outcome = read_outcome(calist, path)
        app._files[path] = outcome
        app._update_row(outcome)
        app._marks.append((index, app._started_at + index * 0.1))
        app._update_progress(outcome, index, total)


def st_results(app, calist):
    load(app, calist)
    outcomes = [read_outcome(calist, p) for p in sorted(app._files)]
    for outcome in outcomes:
        app._files[outcome.path] = outcome
    app._started_at = time.monotonic() - 2.4
    app._enter_results(fake_result(calist, outcomes))


def st_results_all(app, calist):
    st_results(app, calist)
    app._filter.set("All")
    app._refresh_all()


def st_details(app, calist):
    st_results(app, calist)
    # What a real run writes to Details; the screenshot run does not read forms.
    for line in (
            "Template : Device List.xlsx",
            "Sources  : 30 files selected",
            "[WARN]  L02-ZZZ099-0925.xlsx: unknown device code 'ZZZ'",
            "[WARN]  W01-QQQ100-0925.xlsx: unknown device code 'QQQ'",
            "[ERROR]  T03-AE012-0925.xlsx: the workbook is password-protected",
            "Copy left out: I05-AF007-0925 (2).xlsx (same device as I05-AF007-0925.xlsx)",
            "[WARN]  Duplicate serial PM-88301: I05-AGH021-0925.xlsx (already in "
            "I05-AGH020-0925.xlsx)",
            "Writing 27 rows…",
            "Saved C:\Rounds\January round\device list.xlsx"):
        calist.log.info(line)
    if not app._log_open:
        app._toggle_log()


def st_search(app, calist):
    load(app, calist)
    app._search.insert(0, "I05 monitor")
    app._on_search_typed()
    app._run_search()


def st_group(app, calist):
    load(app, calist)
    import ui_state
    app._pick_group(ui_state.problem_groups(app._files.values())[0])


def st_clear(app, calist):
    """Clear all, armed: the button asks before it throws the list away."""
    load(app, calist)
    app._arm_clear()


def st_settings(app, calist):
    load(app, calist)
    app._dedup.set(True)                  # one setting on, one off
    app._toggle_settings()


def st_turbo(app, calist):
    """A 1,200-form archive run, part way through: Turbo takes over."""
    app._files.clear()
    for n in range(1200):
        folder, name, _ = DEVICES[n % 24]
        stem, ext = name.rsplit(".", 1)
        site, code, month = stem.split("-")
        path = str(ROUND / folder / f"{site}-{code[:-3]}{n:03d}-{month}.{ext}")
        app._files[path] = calist.classify_file(path)
    for path in paths()[24:27]:                      # a few unknown codes
        app._files[path] = calist.classify_file(path)
    app._enter_setup()
    files = sorted(app._files)
    total = len(files)
    app._turbo_mode = True
    app._cancel = __import__("threading").Event()
    app._enter_working(total)
    app._started_at = time.monotonic() - 3
    app._marks = [(0, app._started_at)]
    for index, path in enumerate(files[:500], start=1):
        outcome = read_outcome(calist, path)
        app._files[path] = outcome
        if outcome.is_problem:
            app._live_problems.append(outcome)
        app._marks.append((index, app._started_at + index * 0.006))
    app._update_progress(outcome, 500, total)
    app._refresh_table()
    app._refresh_banner()


def st_lock(app, calist):
    app.show_lock()


STATES = {
    "empty": st_empty,
    "loaded": st_loaded,
    "problems": st_problems,
    "group": st_group,
    "search": st_search,
    "settings": st_settings,
    "clear": st_clear,
    "working": st_working,
    "results": st_results,
    "results-all": st_results_all,
    "details": st_details,
    "turbo": st_turbo,
    "lock": st_lock,
}


# ── main ─────────────────────────────────────────────────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("--theme", choices=("dark", "light"), default="dark")
    parser.add_argument("--size", default="1366x768",
                        help="the screen to emulate, e.g. 1366x768 or 1920x1080")
    parser.add_argument("--scale", type=float, default=1.0,
                        help="that screen's Windows scaling, e.g. 1.0 or 1.25")
    parser.add_argument("--state", action="append", choices=sorted(STATES),
                        help="render only these states (repeatable)")
    parser.add_argument("--label", default="current",
                        help="subfolder of docs/review to write into")
    parser.add_argument("--win10", action="store_true",
                        help="force the fonts a Windows 10 machine falls back to")
    args = parser.parse_args()

    screen_w, screen_h = (int(n) for n in args.size.lower().split("x"))
    dpi = system_dpi()

    import customtkinter as ctk
    from PIL import ImageGrab

    import access
    import calist
    import ui

    # A temporary settings file, unlocked for today, so the day-watcher can
    # never pull the lock over a capture and nothing real is read or written.
    workdir = Path(tempfile.mkdtemp(prefix="calist-snap-"))
    ui.SETTINGS_FILE = workdir / "settings.json"
    ui.save_settings(access.mark_unlocked({}))
    if args.win10 and hasattr(ui, "force_windows10_fonts"):
        ui.force_windows10_fonts()

    ctk.set_appearance_mode(args.theme)
    ctk.set_default_color_theme("blue")
    ctk.set_widget_scaling(args.scale / dpi)
    ctk.set_window_scaling(args.scale / dpi)

    app = ui.App()
    if getattr(app, "_zoom_job", None):       # a laptop-sized screen: we size it
        app.after_cancel(app._zoom_job)
    # Tk sizes point fonts (the device table) from its own scaling, which
    # CustomTkinter's settings do not reach.
    app.tk.call("tk", "scaling", args.scale * 96 / 72)
    ui.style_treeview(app)
    app._colour_tags()

    width = round(screen_w / args.scale)
    height = round((screen_h - CHROME_H * args.scale) / args.scale)
    app.geometry(f"{width}x{height}+0+0")
    app.attributes("-topmost", True)
    settle(app, 1.5)

    folder = OUT / args.label
    folder.mkdir(parents=True, exist_ok=True)
    tag = f"{args.theme}-{screen_w}x{screen_h}" + ("-win10" if args.win10 else "")

    for name in args.state or list(STATES):
        if app._lock is not None:            # the lock state leaves it up
            app._lock.destroy()
            app._lock = None
            app._page.grid()
            app._footer.grid()
        if app._panel_open:                  # so is the settings drawer
            app._panel_open = False
            app._panel.place_forget()
        app._cancel = None
        app._turbo_mode = False
        app._group = None
        app._clear_search()
        if app._log_open:
            app._toggle_log()
        if app._clear_job is not None:       # an armed Clear all stands down
            app.after_cancel(app._clear_job)
            app._disarm_clear()
        STATES[name](app, calist)
        settle(app)
        path = folder / f"{tag}-{name}.png"
        # The frame's last row is Windows 11's 1px border, and whatever is on
        # screen behind the window shows through it — not the app's pixels.
        left, top, right, bottom = window_bounds(app)
        ImageGrab.grab(bbox=(left, top, right, bottom - 1), all_screens=True).save(path)
        print(path.relative_to(ROOT))

    app.destroy()
    return 0


if __name__ == "__main__":
    sys.exit(main())
