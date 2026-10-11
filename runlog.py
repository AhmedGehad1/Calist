"""
The run log: one line per Build press, crash, PIN lockout and launch, sent to
the owner's Google Sheet for diagnosis.

It is not a core function, so it is built to be invisible:

- **Never in the way.** Nothing here runs on the Tk thread but a dict being
  queued. A build's record is written by the build's own worker, after the
  result is on screen; everything else, including every network byte and the
  first import of ssl and urllib (about 27 ms, measured), happens on one
  daemon thread that waits while a build or folder scan runs and exits when
  there is nothing left to send.
- **Like TCP.** Each record is a file in %LOCALAPPDATA%\\Calist\\outbox until
  the sheet answers with that record's own id. A lost answer means the record
  is sent again, and the sheet answers the second copy as a duplicate (still
  with the id) instead of writing a second row. Nothing is retried faster than
  1, 2, 4 … 60 minutes apart.
- **Silent.** Every error is swallowed. The logger is "calist.runlog", which
  the Details drawer never listens to, with propagation off.
- **Dormant without an address.** The address and token are baked into
  official builds by calist.spec (module calist_endpoint, generated from CI
  secrets) or given in CALIST_RUNLOG_URL / CALIST_RUNLOG_TOKEN. A checkout and
  the test suite have neither, so they create no folder, no file and no
  thread.

The receiving end is tools/runlog_webapp/Code.gs; README "What Calist sends"
says what a record holds. Importing this module costs its own bytecode: no
ssl, urllib.request, socket, calist or GUI toolkit is loaded until needed, and
test_runlog.py pins that in a subprocess.
"""

from __future__ import annotations

import json
import logging
import os
import queue
import re
import sys
import threading
import time
import traceback
from datetime import datetime

log = logging.getLogger("calist.runlog")
log.addHandler(logging.NullHandler())
log.propagate = False

#: The record layout; tools/runlog_webapp/Code.gs refuses any other.
SCHEMA = 1

#: Records kept per kind while a PC cannot reach the sheet. A crash loop can
#: fill its own allowance but never push out a build row.
CAPS = {"build": 1000, "crash": 50, "lockout": 50, "launch": 50}
MAX_AGE_DAYS = 180

#: Seconds to stay quiet after a build or scan ends, so the results screen
#: paints before anything else wants the interpreter.
GRACE_S = 5.0
#: Seconds between two records, unless the sheet asks for more.
MIN_INTERVAL_S = 3.0
#: After the n-th failure in a row: BACKOFF_BASE * 2**(n-1), at most BACKOFF_CAP.
BACKOFF_BASE = 60.0
BACKOFF_CAP = 3600.0
TIMEOUT_S = 60.0
#: An answer longer than this is not the script's: a portal page, a proxy.
MAX_REPLY = 64 * 1024

ERROR_CHARS = 300
FRAMES_KEPT = 12
CRASHES_PER_SESSION = 10
CRASHES_PER_SIGNATURE = 3
#: Real device codes are at most four letters (all 111 in the table). A longer
#: run of letters from a file name is more likely a word, or part of a name,
#: so it is counted under "(long)" and never sent as text.
CODE_LETTERS = 4
TOP_CODES = 5

#: Where an answer may come from. The POST goes to the first host; Google
#: answers with a redirect to the second, where the script's reply waits.
GOOGLE_HOSTS = ("script.google.com", "script.googleusercontent.com")
LOOPBACK_HOSTS = ("127.0.0.1", "localhost")

_RECORD_ID = re.compile(r"[0-9a-f]{32}")
_INSTALL_ID = re.compile(r"[0-9a-f]{16}")
_OUTBOX_NAME = re.compile(r"(\d{13})-([a-z]+)-([0-9a-f]{32})\.json")

#: The moment this process started, near enough: what a launch row reports.
LAUNCHED_AT = time.time()


# ──────────────────────────────────────────────────────────────────────────────
# State shared by the threads. Every field is guarded by _lock unless noted.
# ──────────────────────────────────────────────────────────────────────────────

_lock = threading.Lock()
_thread: threading.Thread | None = None
#: Set by kick() while the sender runs; checked under _lock before it exits,
#: so a record written after its last look at the outbox is never stranded.
_wake = threading.Event()
#: Set when no build or scan is running.
_idle = threading.Event()
_idle.set()
_busy = 0
_busy_ended = 0.0
#: Records made on the Tk thread, waiting for the sender to write them.
_pending: queue.SimpleQueue = queue.SimpleQueue()
_launch_due = False
_launch_done = False
#: Sleeps go through this event, so they cost no polling.
_nap = threading.Event()
#: Set only by _reset (the tests), to make a running sender return.
_closing = False
_crash_seen: dict[tuple, int] = {}
_crash_total = 0
#: The sheet's wish, from the last answer that carried one (sender only).
_interval = 0.0

_UNSET = object()
_endpoint = _UNSET
_install: str | None = None
_app: dict | None = None
#: state.json as last read or written. Loaded off the Tk thread only.
_state: dict | None = None
_state_lock = threading.Lock()


# ──────────────────────────────────────────────────────────────────────────────
# Entry points. None of them raises, and each returns at once when there is
# nowhere to send.
# ──────────────────────────────────────────────────────────────────────────────

def record_build(result, target: str, took_s: float) -> None:
    """A Build press, from the build's worker after its result was posted.

    ``target`` is the register's intended path, for a build that failed
    before saving; ``took_s`` is the worker's own clock around process_files.
    """
    try:
        if not _usable(io=True):
            return
        _write(_envelope("build", time.time(), build_data(result, target, took_s)))
    except Exception:
        log.debug("Could not record a build", exc_info=True)


def record_crash(where: str, exc_type, exc, tb) -> None:
    """An exception nobody else reported: a Tk callback ("tk"), a worker that
    caught it ("build", "scan"), or the start itself ("startup").

    On the Tk thread the record is only queued; the sender writes it. A
    startup crash is written at once, because the process is on its way out.
    """
    try:
        if not _usable(io=False):
            return
        data = crash_data(where, exc_type, exc, tb)
        if not _crash_allowed((data["type"], (data["frames"] or [""])[-1])):
            return
        if where != "startup" and threading.current_thread() is threading.main_thread():
            _pending.put(("crash", time.time(), data))
            kick()
        else:
            _write(_envelope("crash", time.time(), data))
    except Exception:
        log.debug("Could not record a crash", exc_info=True)


def record_lockout(failures: int, wait_s: int) -> None:
    """A PIN cooldown starting (Tk thread: queued only). Never the digits."""
    try:
        if not _usable(io=False):
            return
        _pending.put(("lockout", time.time(),
                      {"failures": int(failures), "wait_s": int(wait_s)}))
        kick()
    except Exception:
        log.debug("Could not record a lockout", exc_info=True)


def start() -> None:
    """Called once, a little after launch (Tk thread). Starting the sender is
    all it does; the sender writes the launch row and sends what is waiting."""
    global _launch_due
    try:
        if not _usable(io=False):
            return
        with _lock:
            if _launch_done:
                return
            _launch_due = True
        kick()
    except Exception:
        log.debug("Could not start the run log", exc_info=True)


def hold() -> None:
    """A build or scan is starting: the sender waits until it is over."""
    global _busy
    with _lock:
        _busy += 1
        _idle.clear()


def release() -> None:
    """A build or scan has ended: the sender may go on, after GRACE_S."""
    global _busy, _busy_ended
    with _lock:
        _busy = max(0, _busy - 1)
        if not _busy:
            _busy_ended = time.monotonic()
            _idle.set()
    kick()


def kick() -> None:
    """Start the sender, or tell the running one that there is new work."""
    global _thread
    try:
        if not _usable(io=False):
            return
        with _lock:
            if _thread is not None:
                _wake.set()
                return
            _wake.clear()
            _thread = threading.Thread(target=_sender, name="calist-runlog",
                                       daemon=True)
            _thread.start()
    except Exception:
        log.debug("Could not start the run log sender", exc_info=True)


# ──────────────────────────────────────────────────────────────────────────────
# Where to send, and whether to
# ──────────────────────────────────────────────────────────────────────────────

def endpoint() -> tuple[str, str] | None:
    """(url, token), or None when this copy has nowhere to send.

    The environment wins, for the owner's own testing; otherwise the module
    calist.spec generates from the CI secrets. Resolved once per process.
    """
    global _endpoint
    if _endpoint is _UNSET:
        _endpoint = _resolve_endpoint()
    return _endpoint


def _resolve_endpoint() -> tuple[str, str] | None:
    url = os.environ.get("CALIST_RUNLOG_URL", "").strip()
    token = os.environ.get("CALIST_RUNLOG_TOKEN", "").strip()
    override = bool(url)
    if not override:
        try:
            import calist_endpoint as baked        # generated at build time only
        except Exception:
            return None
        url = str(getattr(baked, "URL", "") or "").strip()
        token = str(getattr(baked, "TOKEN", "") or "").strip()
    return (url, token) if valid_url(url, loopback_ok=override) else None


def valid_url(url: str, loopback_ok: bool = False) -> bool:
    """A web app's /exec address on script.google.com, or, for tests and the
    owner's own override, plain http on this machine. urllib would happily
    open file:// and ftp:// too, which is why this is checked at all."""
    from urllib.parse import urlsplit             # loaded by calist anyway
    try:
        parts = urlsplit(url)
        host = parts.hostname
    except ValueError:
        return False
    if parts.scheme == "https" and host == GOOGLE_HOSTS[0]:
        return parts.path.endswith("/exec")
    return loopback_ok and parts.scheme == "http" and host in LOOPBACK_HOSTS


def _redirect_ok(url: str, loopback_ok: bool) -> bool:
    from urllib.parse import urlsplit
    try:
        parts = urlsplit(url)
        host = parts.hostname
    except ValueError:
        return False
    if parts.scheme == "https" and host in GOOGLE_HOSTS:
        return True
    return loopback_ok and parts.scheme == "http" and host in LOOPBACK_HOSTS


def _data_dir() -> str | None:
    base = os.environ.get("LOCALAPPDATA")
    return os.path.join(base, "Calist") if base else None


def _outbox_dir() -> str | None:
    base = _data_dir()
    return os.path.join(base, "outbox") if base else None


def _usable(io: bool) -> bool:
    """Somewhere to send, somewhere to keep it, and the sheet has not said stop.

    With ``io`` False (the Tk thread) the stop is read only from memory; the
    sender loads it from disk and drops anything queued in the meantime.
    """
    if endpoint() is None or _data_dir() is None:
        return False
    state = _state if _state is not None or not io else _load_state()
    return not _stopped(state)


def _stopped(state: dict | None) -> bool:
    try:
        return float((state or {}).get("disable_until", 0)) > time.time()
    except (TypeError, ValueError):
        return False


# ──────────────────────────────────────────────────────────────────────────────
# What a record says. Pure functions of what they are given.
# ──────────────────────────────────────────────────────────────────────────────

def outcome_of(result) -> str:
    """Saved / Cancelled / Nothing to write / Failed, from the result's shape
    rather than its message, so a reworded message changes nothing."""
    from calist import OK
    if result.output_path is not None:
        return "Saved"
    if result.cancelled:
        return "Cancelled"
    if result.error and result.outcomes and not any(
            o.status == OK for o in result.outcomes):
        return "Nothing to write"
    return "Failed"


def build_data(result, target: str, took_s: float) -> dict:
    """The build row's own fields, counted in one pass over the outcomes.

    Read only after process_files has returned: settling copies turns an
    outcome already reported as read into a copy (calist._settle_copies), so
    the files_read the footer shows is only final then.
    """
    from calist import COPY, ERROR, LEFT_OUT, OK, UNKNOWN_CODE, UNSUPPORTED
    tally = {OK: 0, ERROR: 0, UNKNOWN_CODE: 0, UNSUPPORTED: 0, LEFT_OUT: 0, COPY: 0}
    codes: dict[str | None, int] = {}
    for outcome in result.outcomes:
        if outcome.status in tally:
            tally[outcome.status] += 1
        if outcome.status == UNKNOWN_CODE:
            codes[outcome.device_code] = codes.get(outcome.device_code, 0) + 1
    outcome = outcome_of(result)
    return {
        "outcome": outcome,
        "devices": tally[OK],
        "rows": result.rows_written,
        "took_s": round(float(took_s), 2),
        "read_s": round(float(result.read_seconds), 2),
        "write_s": round(float(result.write_seconds), 2),
        "saved_to": str(result.output_path or target or ""),
        "errors": tally[ERROR],
        "unknown": tally[UNKNOWN_CODE],
        "unsupported": tally[UNSUPPORTED],
        "left_out": tally[LEFT_OUT],
        "copies": tally[COPY],
        "dupes_removed": result.duplicates_removed,
        "unknown_codes": format_codes(codes),
        "error": scrub(result.error or "") if outcome == "Failed" else "",
    }


def format_codes(counts: dict) -> str:
    """'GZ×3, HE×1, +2 more': the device codes Calist did not know, most
    common first. A file with no code counts as '?'; see CODE_LETTERS."""
    tidy: dict[str, int] = {}
    for code, n in counts.items():
        if not code or not (str(code).isascii() and str(code).isalpha()):
            key = "?"
        elif len(code) > CODE_LETTERS:
            key = "(long)"
        else:
            key = str(code).upper()
        tidy[key] = tidy.get(key, 0) + int(n)
    ranked = sorted(tidy.items(), key=lambda item: (-item[1], item[0]))
    shown = ", ".join(f"{code}×{n}" for code, n in ranked[:TOP_CODES])
    more = len(ranked) - TOP_CODES
    return shown + (f", +{more} more" if more > 0 else "")


#: A path in quotes ('C:\\Users\\…', "\\\\server\\share\\…"): the whole quoted
#: text goes. Then a bare one, from its drive or share to the end of the
#: phrase: losing the rest of a sentence is better than half a folder name.
_QUOTED_PATH = re.compile(r"""(['"])(?:[A-Za-z]:[\\/]|\\\\)(?:(?!\1).)*\1""")
_BARE_PATH = re.compile(r"""(?:\b[A-Za-z]:[\\/]|\\\\(?=\w))[^'"<>|\r\n]*""")


def scrub(text: str, limit: int = ERROR_CHARS) -> str:
    """One line, at most ``limit`` characters, with every path replaced.

    Error messages carry paths (the register being saved, a form that would
    not open), and those name a Windows user and a customer's folder."""
    text = _QUOTED_PATH.sub(lambda m: f"{m.group(1)}<path>{m.group(1)}", str(text))
    text = _BARE_PATH.sub("<path>", text)
    return " ".join(text.split())[:limit]


def frames(tb, keep: int = FRAMES_KEPT) -> list[str]:
    """'ui.py:2532 worker' per frame, innermost last. walk_tb reads no source
    file, so this is safe on the Tk thread; basenames keep user paths out."""
    out = [f"{os.path.basename(frame.f_code.co_filename)}:{lineno} "
           f"{frame.f_code.co_name}" for frame, lineno in traceback.walk_tb(tb)]
    return out[-keep:]


def crash_data(where: str, exc_type, exc, tb) -> dict:
    name = getattr(exc_type, "__qualname__", None) or getattr(exc_type, "__name__", "?")
    module = getattr(exc_type, "__module__", "") or ""
    if module and module != "builtins":
        name = f"{module}.{name}"
    try:
        message = str(exc) if exc is not None else ""
    except Exception:
        message = ""
    return {"where": str(where)[:20], "type": str(name)[:80],
            "message": scrub(message), "frames": frames(tb)}


def _crash_allowed(signature: tuple) -> bool:
    """A callback that fails on every tick must not become a thousand rows."""
    global _crash_total
    with _lock:
        if (_crash_total >= CRASHES_PER_SESSION
                or _crash_seen.get(signature, 0) >= CRASHES_PER_SIGNATURE):
            return False
        _crash_seen[signature] = _crash_seen.get(signature, 0) + 1
        _crash_total += 1
        return True


# ──────────────────────────────────────────────────────────────────────────────
# Who and where
# ──────────────────────────────────────────────────────────────────────────────

def install_id() -> str:
    """This Windows user's random id on this PC, made once.

    Random rather than read from the registry: an unsigned exe that reads the
    machine id and sends it with user and PC names looks like the
    fingerprinting antivirus models hunt for. The 'x' mode means two Calists
    starting together cannot both write one."""
    global _install
    if _install:
        return _install
    folder = _data_dir()
    path = os.path.join(folder, "install-id")
    try:
        with open(path, encoding="ascii") as f:
            value = f.read().strip()
        if _INSTALL_ID.fullmatch(value):
            _install = value
            return value
        os.remove(path)                          # damaged: make a new one
    except (FileNotFoundError, UnicodeDecodeError):
        pass
    os.makedirs(folder, exist_ok=True)
    value = os.urandom(8).hex()
    try:
        with open(path, "x", encoding="ascii") as f:
            f.write(value)
    except FileExistsError:
        with open(path, encoding="ascii") as f:
            value = f.read().strip()
    _install = value
    return value


def user_name() -> str:
    name = os.environ.get("USERNAME", "")
    if not name:
        try:
            name = os.getlogin()
        except OSError:
            name = ""
    return name[:64]


def pc_name() -> str:
    name = os.environ.get("COMPUTERNAME", "")
    if not name:
        try:
            import socket
            name = socket.gethostname()
        except Exception:
            name = ""
    return name[:64]


def build_shape() -> str:
    """one-file (unpacks to %TEMP%\\_MEIxxxx), folder, or source."""
    if not getattr(sys, "frozen", False):
        return "source"
    unpacked = os.path.basename(getattr(sys, "_MEIPASS", "") or "")
    return "one-file" if unpacked.startswith("_MEI") else "folder"


def ram_gb() -> float | None:
    try:
        import ctypes

        class MemoryStatus(ctypes.Structure):
            _fields_ = [("dwLength", ctypes.c_ulong), ("dwMemoryLoad", ctypes.c_ulong),
                        ("ullTotalPhys", ctypes.c_ulonglong),
                        ("ullAvailPhys", ctypes.c_ulonglong),
                        ("ullTotalPageFile", ctypes.c_ulonglong),
                        ("ullAvailPageFile", ctypes.c_ulonglong),
                        ("ullTotalVirtual", ctypes.c_ulonglong),
                        ("ullAvailVirtual", ctypes.c_ulonglong),
                        ("ullAvailExtendedVirtual", ctypes.c_ulonglong)]

        status = MemoryStatus()
        status.dwLength = ctypes.sizeof(status)
        if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status)):
            return round(status.ullTotalPhys / 2 ** 30, 1)
    except Exception:
        pass
    return None


def app_info() -> dict:
    global _app
    if _app is None:
        try:
            from calist import __version__ as version
        except Exception:
            version = "?"
        _app = {"version": version, "shape": build_shape(),
                "cpus": os.cpu_count() or 0, "ram_gb": ram_gb()}
    return dict(_app)


def _iso(timestamp: float) -> str:
    """Local time with its UTC offset: '2026-10-11T14:32:05+03:00'."""
    return datetime.fromtimestamp(timestamp).astimezone().isoformat(timespec="seconds")


def _envelope(kind: str, at: float, data: dict) -> dict:
    return {"v": SCHEMA, "kind": kind, "id": os.urandom(16).hex(), "at": _iso(at),
            "pc": {"install": install_id(), "name": pc_name(), "user": user_name()},
            "app": app_info(), "data": data}


# ──────────────────────────────────────────────────────────────────────────────
# The outbox: one file per record, named so that sorting is oldest first.
# ──────────────────────────────────────────────────────────────────────────────

def _write(record: dict) -> None:
    """Write a record where only a finished file can be seen (tmp, then
    os.replace), then keep its kind within CAPS and MAX_AGE_DAYS."""
    folder = _outbox_dir()
    os.makedirs(folder, exist_ok=True)
    name = f"{int(time.time() * 1000):013d}-{record['kind']}-{record['id']}.json"
    temporary = os.path.join(folder, name + ".tmp")
    with open(temporary, "w", encoding="ascii") as f:
        json.dump(record, f, ensure_ascii=True, separators=(",", ":"))
    os.replace(temporary, os.path.join(folder, name))
    _trim(folder, record["kind"])


def outbox_names(folder: str | None = None) -> list[str]:
    folder = folder or _outbox_dir()
    try:
        return sorted(n for n in os.listdir(folder) if _OUTBOX_NAME.fullmatch(n))
    except OSError:
        return []


def _trim(folder: str, kind: str) -> None:
    names = outbox_names(folder)
    oldest_kept = time.time() * 1000 - MAX_AGE_DAYS * 86_400_000
    stale = [n for n in names if int(n[:13]) < oldest_kept]
    mine = [n for n in names if _kind_of(n) == kind and n not in stale]
    excess = mine[:max(0, len(mine) - CAPS.get(kind, 50))]
    gone: dict[str, int] = {}
    for name in stale + excess:
        if _remove(os.path.join(folder, name)):
            gone[_kind_of(name)] = gone.get(_kind_of(name), 0) + 1
    if gone:
        _count_dropped(gone)


def _sweep_temporaries(folder: str) -> None:
    """A .tmp left by a process killed mid-write, an hour or more ago."""
    try:
        for name in os.listdir(folder):
            path = os.path.join(folder, name)
            if name.endswith(".tmp") and time.time() - os.path.getmtime(path) > 3600:
                _remove(path)
    except OSError:
        pass


def _kind_of(name: str) -> str:
    match = _OUTBOX_NAME.fullmatch(name)
    return match.group(2) if match else "?"


def _remove(path: str) -> bool:
    try:
        os.remove(path)
        return True
    except OSError:                    # gone already, or held for a moment
        return False


# ──────────────────────────────────────────────────────────────────────────────
# state.json: drops still to report, the last send error, the sheet's stop.
# ──────────────────────────────────────────────────────────────────────────────

def _state_path() -> str:
    return os.path.join(_data_dir(), "state.json")


def _load_state() -> dict:
    global _state
    try:
        with open(_state_path(), encoding="utf-8") as f:
            state = json.load(f)
        if not isinstance(state, dict):
            state = {}
    except (OSError, ValueError):
        state = {}
    _state = state
    return state


def _update_state(change) -> None:
    global _state
    with _state_lock:
        state = dict(_load_state())
        change(state)
        try:
            os.makedirs(_data_dir(), exist_ok=True)
            temporary = _state_path() + ".tmp"
            with open(temporary, "w", encoding="utf-8") as f:
                json.dump(state, f)
            os.replace(temporary, _state_path())
        except OSError:
            log.debug("Could not save the run log state", exc_info=True)
        _state = state


def _count_dropped(counts: dict[str, int]) -> None:
    def change(state):
        dropped = dict(state.get("dropped") or {})
        for kind, n in counts.items():
            dropped[kind] = int(dropped.get(kind, 0)) + n
        state["dropped"] = dropped
    _update_state(change)


def _health() -> dict:
    """Sent with every record: drops not yet reported, and why sending last
    failed, so a PC that keeps failing shows up instead of going quiet."""
    state = _state or {}
    try:
        dropped = {str(k): int(v) for k, v in (state.get("dropped") or {}).items()
                   if int(v) > 0}
        failures = int(state.get("failures", 0) or 0)
    except (AttributeError, TypeError, ValueError):     # a damaged state.json
        dropped, failures = {}, 0
    error = str(state.get("last_error", "") or "")
    return {"dropped": dropped,
            "last_error": f"{error} ×{failures}" if error and failures else ""}


def _acked(sent: dict) -> None:
    """The drops a record reported are reported; the failure run is over."""
    state = _state or {}
    if not sent.get("dropped") and not state.get("last_error"):
        return

    def change(state):
        dropped = dict(state.get("dropped") or {})
        for kind, n in sent.get("dropped", {}).items():
            left = int(dropped.get(kind, 0)) - int(n)
            if left > 0:
                dropped[kind] = left
            else:
                dropped.pop(kind, None)
        state["dropped"] = dropped
        state.pop("last_error", None)
        state.pop("failures", None)
    _update_state(change)


def _failed(why: str) -> None:
    def change(state):
        try:
            state["failures"] = int(state.get("failures", 0) or 0) + 1
        except (TypeError, ValueError):
            state["failures"] = 1
        state["last_error"] = str(why)[:80]
    _update_state(change)


def _apply_settings(reply) -> None:
    """The sheet can stop every copy for a while, or slow them down."""
    global _interval
    settings = reply.get("cfg") if isinstance(reply, dict) else None
    if not isinstance(settings, dict):
        return
    interval = settings.get("min_interval_s")
    if isinstance(interval, (int, float)) and interval >= 0:
        _interval = min(float(interval), BACKOFF_CAP)
    if settings.get("stop") is True:
        hours = settings.get("stop_hours")
        hours = float(hours) if isinstance(hours, (int, float)) and hours > 0 else 24.0
        until = time.time() + min(hours, 24 * 30) * 3600
        _update_state(lambda state: state.__setitem__("disable_until", until))


# ──────────────────────────────────────────────────────────────────────────────
# The answer
# ──────────────────────────────────────────────────────────────────────────────

def read_reply(status: int, body: bytes) -> dict | None:
    """The script's JSON answer, or None for anything else: a sign-in or
    portal page, an empty body, a proxy's error."""
    if status != 200 or len(body) > MAX_REPLY:
        return None
    try:
        reply = json.loads(body.decode("utf-8"))
    except ValueError:
        return None
    return reply if isinstance(reply, dict) else None


def is_ack(reply, sent_id: str) -> bool:
    """The receipt, and the only thing that deletes a record: ok, and this
    record's own id. A duplicate answer counts, because it carries the id."""
    return (isinstance(reply, dict) and reply.get("ok") is True
            and reply.get("id") == sent_id)


def is_rejection(reply, sent_id: str) -> bool:
    """The script will never take this record (bad token, unknown kind): keep
    it and it would block every record behind it."""
    return (isinstance(reply, dict) and reply.get("ok") is False
            and reply.get("reject") is True and reply.get("id") in (sent_id, None))


def backoff(failures: int) -> float:
    """Seconds to wait after the n-th failure in a row: 1, 2, 4 … 60 minutes."""
    return min(BACKOFF_BASE * 2 ** max(0, failures - 1), BACKOFF_CAP)


class _Transport:
    """urllib, built on the sender thread only: importing ssl and urllib and
    making a context is the expensive part, so it is done once per run."""

    def __init__(self, url: str, token: str):
        import ssl
        import urllib.error
        import urllib.request

        loopback = not url.startswith("https://")
        allowed = _redirect_ok

        class Redirects(urllib.request.HTTPRedirectHandler):
            # urllib follows a redirect anywhere, ftp included; Google's
            # answer only ever comes from its own two hosts.
            def redirect_request(self, req, fp, code, msg, headers, newurl):
                if not allowed(newurl, loopback):
                    raise urllib.error.HTTPError(newurl, code, "redirect refused",
                                                 headers, fp)
                return super().redirect_request(req, fp, code, msg, headers, newurl)

        handlers = [Redirects()]
        if loopback:                  # the tests' fake: never via a proxy
            handlers.append(urllib.request.ProxyHandler({}))
        else:
            handlers.append(urllib.request.HTTPSHandler(
                context=ssl.create_default_context()))
        self.url = url
        self.token = token
        self.version = app_info()["version"]
        self._request = urllib.request.Request
        self._http_error = urllib.error.HTTPError
        self._opener = urllib.request.build_opener(*handlers)

    def post(self, body: bytes) -> tuple[dict | None, str]:
        """(reply, why it is not one). urllib turns the POST into a GET when
        Google redirects it with a 302, which is exactly what is needed."""
        request = self._request(self.url, data=body, headers={
            "Content-Type": "application/json",
            "User-Agent": f"Calist/{self.version}"})
        try:
            with self._opener.open(request, timeout=TIMEOUT_S) as response:
                status = response.status
                raw = response.read(MAX_REPLY + 1)
        except self._http_error as exc:
            exc.close()
            return None, f"HTTP {exc.code}"
        except Exception as exc:
            return None, type(exc).__name__
        reply = read_reply(status, raw)
        return reply, "" if reply is not None else f"not a receipt (HTTP {status})"


# ──────────────────────────────────────────────────────────────────────────────
# The sender
# ──────────────────────────────────────────────────────────────────────────────

def _sender() -> None:
    global _thread
    try:
        _run()
    except BaseException:                         # a daemon must never shout
        log.debug("The run log sender stopped", exc_info=True)
    finally:
        with _lock:
            if _thread is threading.current_thread():
                _thread = None


def _run() -> None:
    target = endpoint()
    folder = _outbox_dir()
    if target is None or folder is None:
        return
    _load_state()
    _sweep_temporaries(folder)
    transport = None
    while not _closing:
        _write_queued()
        if _stopped(_state):
            _drop_queued()
            _finished(force=True)
            return
        names = outbox_names(folder)
        if not names:
            if _finished():
                return
            continue
        _wait_until_idle()
        if _closing:
            return
        if transport is None:
            try:
                transport = _Transport(*target)
            except Exception as exc:         # e.g. the Windows certificate store
                _failed(type(exc).__name__)
                _nap.wait(backoff(_failed_count()))
                continue
            _wait_until_idle()               # the imports took a moment of the GIL
        if _send(transport, folder, names[0]):
            _nap.wait(max(MIN_INTERVAL_S, _interval))
        else:
            _nap.wait(backoff(_failed_count()))


def _write_queued() -> None:
    """The launch row and anything the Tk thread queued, written here."""
    global _launch_due, _launch_done
    with _lock:
        launch = _launch_due and not _launch_done
        _launch_due = False
        _launch_done = _launch_done or launch
    if launch:
        _safely_write("launch", LAUNCHED_AT, {})
    while True:
        try:
            kind, at, data = _pending.get_nowait()
        except queue.Empty:
            return
        _safely_write(kind, at, data)


def _safely_write(kind: str, at: float, data: dict) -> None:
    try:
        _write(_envelope(kind, at, data))
    except Exception:
        log.debug("Could not write a %s record", kind, exc_info=True)
        _count_dropped({kind: 1})


def _drop_queued() -> None:
    while True:
        try:
            _pending.get_nowait()
        except queue.Empty:
            return


def _finished(force: bool = False) -> bool:
    """The exit decision, under the lock kick() takes. Work that arrived since
    the last look at the outbox has set _wake, so the loop goes round again."""
    global _thread
    with _lock:
        if not force and (_wake.is_set() or not _pending.empty() or _launch_due):
            _wake.clear()
            return False
        if _thread is threading.current_thread():
            _thread = None
        return True


def _wait_until_idle() -> None:
    """Wait out a build or scan, then GRACE_S more for the results to paint."""
    while not _closing:
        _idle.wait()
        left = GRACE_S - (time.monotonic() - _busy_ended)
        if left <= 0 and _idle.is_set():
            return
        if left > 0:
            _nap.wait(left)


def _failed_count() -> int:
    try:
        return max(1, int((_state or {}).get("failures", 1) or 1))
    except (TypeError, ValueError):
        return 1


def _send(transport: _Transport, folder: str, name: str) -> bool:
    """One record, start to finish. True when it is settled: received, or
    gone for good (another Calist sent it, damaged, refused by the script)."""
    path = os.path.join(folder, name)
    kind = _kind_of(name)
    try:
        with open(path, encoding="utf-8") as f:
            record = json.load(f)
        if not (isinstance(record, dict)
                and _RECORD_ID.fullmatch(str(record.get("id", "")))):
            raise ValueError("not a record")
    except FileNotFoundError:
        return True
    except ValueError:                             # damaged: it can never be sent
        if _remove(path):
            _count_dropped({kind: 1})
        return True
    except OSError as exc:                         # held for a moment (antivirus)
        _failed(type(exc).__name__)
        return False
    health = _health()
    try:
        payload = dict(record, token=transport.token, health=health)
        body = json.dumps(payload, ensure_ascii=True,
                          separators=(",", ":")).encode("ascii")
    except (TypeError, ValueError):                # cannot be sent as it is
        if _remove(path):
            _count_dropped({kind: 1})
        return True
    reply, why = transport.post(body)
    _apply_settings(reply)
    if is_ack(reply, record["id"]):
        _remove(path)
        _acked(health)
        return True
    if is_rejection(reply, record["id"]):
        if _remove(path):
            _count_dropped({kind: 1})
        return True
    _failed(why or "not a receipt")
    return False


def _reset() -> None:
    """Stop any sender and go back to a fresh process: for the tests only."""
    global _thread, _busy, _busy_ended, _launch_due, _launch_done, _crash_total
    global _interval, _endpoint, _install, _app, _state, _closing
    _closing = True
    _nap.set()
    _idle.set()
    running = _thread
    if running is not None and running is not threading.current_thread():
        running.join(timeout=10)
    with _lock:
        _thread = None
        _busy = 0
        _busy_ended = 0.0
        _launch_due = _launch_done = False
        _crash_total = 0
        _crash_seen.clear()
    _wake.clear()
    _nap.clear()
    _drop_queued()
    _interval = 0.0
    _endpoint = _UNSET
    _install = None
    _app = None
    _state = None
    _closing = False
