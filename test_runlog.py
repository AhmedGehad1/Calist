"""Tests for runlog.py, the run log. Run with: python -m pytest

No display and no network: the sender talks to a fake of the Apps Script web
app on 127.0.0.1, shaped like the real one. The POST runs the script and is
answered with a 302, and the script's JSON receipt waits at the redirect
target for the GET that urllib turns the POST into.
"""

import http.server
import json
import os
import re
import subprocess
import sys
import threading
import time
from pathlib import Path
from types import SimpleNamespace

import pytest

import calist
import runlog
from calist import (COPY, ERROR, LEFT_OUT, OK, UNKNOWN_CODE, UNSUPPORTED,
                    FileOutcome, RunResult)

ROOT = Path(__file__).resolve().parent
TOKEN = "t0ken"


@pytest.fixture(autouse=True)
def isolated(tmp_path, monkeypatch):
    """Its own LOCALAPPDATA, no address unless the test gives one, short waits."""
    runlog._reset()
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "local"))
    monkeypatch.delenv("CALIST_RUNLOG_URL", raising=False)
    monkeypatch.delenv("CALIST_RUNLOG_TOKEN", raising=False)
    monkeypatch.setattr(runlog, "GRACE_S", 0.0)
    monkeypatch.setattr(runlog, "MIN_INTERVAL_S", 0.0)
    monkeypatch.setattr(runlog, "BACKOFF_BASE", 0.05)
    monkeypatch.setattr(runlog, "BACKOFF_CAP", 0.2)
    monkeypatch.setattr(runlog, "TIMEOUT_S", 2.0)
    yield
    runlog._reset()


def outcome(name, status, **kw):
    return FileOutcome(filename=name, path=rf"C:\Round\{name}", status=status, **kw)


def saved_result(**kw):
    """A round as the footer would show it: 4 rows from 3 forms, and the rest."""
    outcomes = [
        outcome("W01-AC004-0126.xlsx", OK, rows=1),
        outcome("W01-AC005-0126.xlsx", OK, rows=1),
        outcome("W01-AGH011-0126.xlsx", OK, rows=2),
        outcome("W01-AE012-0126.xlsx", ERROR, detail="password-protected"),
        outcome("W01-GZ001-0126.xlsx", UNKNOWN_CODE, device_code="GZ"),
        outcome("W01-GZ002-0126.xlsx", UNKNOWN_CODE, device_code="GZ"),
        outcome("5071938426.xlsx", UNKNOWN_CODE, device_code=None),
        outcome("notes.docx", UNSUPPORTED),
        outcome("Device List.xlsx", LEFT_OUT),
        outcome("W01-AC004-0126 (2).xlsx", COPY),
    ]
    values = dict(output_path=Path(r"D:\Rounds\January\device list.xlsx"),
                  rows_written=4, outcomes=outcomes, duplicates_removed=1,
                  read_seconds=1.234, write_seconds=0.5)
    values.update(kw)
    return RunResult(**values)


def wait_for(condition, timeout=5.0):
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        if condition():
            return True
        time.sleep(0.01)
    return condition()


def settled():
    """Nothing waiting and the sender gone: what the end of a send looks like."""
    return not runlog.outbox_names() and runlog._thread is None


# ── what a build row says ─────────────────────────────────────────────────────

def test_a_build_row_says_what_the_footer_says():
    data = runlog.build_data(saved_result(), target="", took_s=2.0)

    assert data == {
        "outcome": "Saved", "devices": 3, "rows": 4,
        "took_s": 2.0, "read_s": 1.23, "write_s": 0.5,
        "saved_to": r"D:\Rounds\January\device list.xlsx",
        "errors": 1, "unknown": 3, "unsupported": 1, "left_out": 1, "copies": 1,
        "dupes_removed": 1, "unknown_codes": "GZ×2, ?×1", "error": "",
    }


def test_devices_are_the_forms_read_after_copies_settle():
    """_settle_copies turns an outcome already reported as read into a copy, so
    the count is taken from the finished result: the footer's number."""
    result = saved_result()
    result.outcomes[1].status = COPY
    assert runlog.build_data(result, "", 1.0)["devices"] == result.files_read == 2


LOCKED = ("Could not save output file: [Errno 13] Permission denied: "
          "'C:\\\\Users\\\\bob\\\\January round\\\\device list.xlsx'")


@pytest.mark.parametrize("result, expected", [
    (saved_result(), "Saved"),
    (saved_result(output_path=None, rows_written=0, cancelled=True), "Cancelled"),
    (RunResult(error="No valid records extracted. No file created.",
               outcomes=[outcome("notes.docx", UNSUPPORTED)]), "Nothing to write"),
    (saved_result(output_path=None, rows_written=0, error=LOCKED), "Failed"),
    (RunResult(error="the worker itself fell over"), "Failed"),
    (RunResult(error="Output path would overwrite the template."), "Failed"),
])
def test_every_way_a_build_can_end_has_an_outcome(result, expected):
    assert runlog.outcome_of(result) == expected


def test_a_failed_save_keeps_its_target_and_loses_its_path():
    result = saved_result(output_path=None, rows_written=0, error=LOCKED)
    data = runlog.build_data(result, target=r"D:\Jan\device list.xlsx", took_s=1)

    assert data["saved_to"] == r"D:\Jan\device list.xlsx"
    assert data["error"] == ("Could not save output file: [Errno 13] "
                             "Permission denied: '<path>'")


def test_nothing_to_write_matches_the_real_pipeline(tmp_path):
    """The classifier reads a result's shape; this pins it to process_files."""
    template = tmp_path / "tpl" / "Device List.xlsx"
    result = calist.process_files([str(tmp_path / "notes.docx")], str(template))
    assert runlog.outcome_of(result) == "Nothing to write"


def test_no_file_name_or_path_leaves_in_a_build_row():
    result = saved_result(outcomes=[
        outcome("SECRETSITE-AC001-0126.xlsx", OK, rows=1, serial="SN-SECRET",
                detail="SECRET detail", code="SECRETSITE-AC001-0126"),
        outcome("SECRETSITE-ZZ001-0126.xlsx", ERROR, detail="SECRET error"),
    ])
    sent = json.dumps(runlog.build_data(result, "", 1.0))
    assert "SECRET" not in sent


def test_unknown_codes_are_counted_never_spelled_out_when_long():
    counts = {"GZ": 3, "HE": 1, None: 1, "HASSAN": 2, "AB": 1, "CD": 1, "EF": 1,
              "G1": 1}
    assert runlog.format_codes(counts) == "GZ×3, (long)×2, ?×2, AB×1, CD×1, +2 more"
    assert runlog.format_codes({}) == ""


@pytest.mark.parametrize("text, expected", [
    (LOCKED, "Could not save output file: [Errno 13] Permission denied: '<path>'"),
    ("cannot open D:\\Rounds\\January round\\x.xlsx now", "cannot open <path>"),
    ('"\\\\server\\share\\form.xlsx" is missing', '"<path>" is missing'),
    ("line one\n  line two", "line one line two"),
    ("x" * 500, "x" * runlog.ERROR_CHARS),
])
def test_error_text_is_one_line_capped_and_pathless(text, expected):
    assert runlog.scrub(text) == expected


# ── the envelope ──────────────────────────────────────────────────────────────

def test_the_envelope_is_pinned_and_json_safe():
    """Its keys are the protocol with tools/runlog_webapp/Code.gs: a rename here
    is a rename there."""
    record = runlog._envelope("build", 1_760_000_000.0, {"outcome": "Saved"})

    assert set(record) == {"v", "kind", "id", "at", "pc", "app", "data"}
    assert set(record["pc"]) == {"install", "name", "user"}
    assert set(record["app"]) == {"version", "shape", "cpus", "ram_gb"}
    assert record["v"] == runlog.SCHEMA == 1
    assert re.fullmatch(r"[0-9a-f]{32}", record["id"])
    assert re.fullmatch(r"[0-9a-f]{16}", record["pc"]["install"])
    assert re.fullmatch(r"\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d[+-]\d\d:\d\d", record["at"])
    assert record["app"]["version"] == calist.__version__
    assert record["app"]["shape"] == "source"
    assert json.loads(json.dumps(record)) == record


def test_the_install_id_is_made_once_and_survives_a_restart(tmp_path):
    first = runlog.install_id()
    runlog._install = None                       # a new process
    assert runlog.install_id() == first

    (tmp_path / "local" / "Calist" / "install-id").write_text("garbage")
    runlog._install = None
    replaced = runlog.install_id()
    assert re.fullmatch(r"[0-9a-f]{16}", replaced) and replaced != first


def test_crash_frames_are_basenames_innermost_last():
    def inner():
        raise ValueError(r"bad cell in C:\Users\bob\form.xlsx")

    try:
        inner()
    except ValueError as exc:
        data = runlog.crash_data("tk", type(exc), exc, exc.__traceback__)

    assert data["type"] == "ValueError"
    assert data["message"] == "bad cell in <path>"
    assert data["frames"][-1].startswith("test_runlog.py:")
    assert data["frames"][-1].endswith(" inner")
    assert all("\\" not in frame and "/" not in frame for frame in data["frames"])


def test_one_bad_callback_cannot_flood_the_crash_tab(monkeypatch):
    monkeypatch.setattr(runlog, "CRASHES_PER_SESSION", 4)
    same = [runlog._crash_allowed(("TclError", "ui.py:865 tick")) for _ in range(5)]
    assert same == [True, True, True, False, False]
    assert runlog._crash_allowed(("KeyError", "ui.py:10 x"))
    assert not runlog._crash_allowed(("OSError", "ui.py:20 y"))      # session cap


# ── where to send ─────────────────────────────────────────────────────────────

@pytest.mark.parametrize("url, loopback_ok, ok", [
    ("https://script.google.com/macros/s/AKfy/exec", False, True),
    ("https://script.google.com/macros/s/AKfy/dev", False, False),
    ("http://script.google.com/macros/s/AKfy/exec", False, False),
    ("https://evil.example/macros/s/AKfy/exec", False, False),
    ("file:///C:/Windows/win.ini", True, False),
    ("ftp://script.google.com/exec", True, False),
    ("http://127.0.0.1:8080/exec", False, False),
    ("http://127.0.0.1:8080/exec", True, True),
    ("not a url", True, False),
])
def test_only_googles_exec_address_or_loopback_is_accepted(url, loopback_ok, ok):
    assert runlog.valid_url(url, loopback_ok) is ok


def test_the_environment_wins_over_the_baked_address(monkeypatch):
    baked = SimpleNamespace(URL="https://script.google.com/macros/s/AKfy/exec",
                            TOKEN="baked")
    monkeypatch.setitem(sys.modules, "calist_endpoint", baked)
    assert runlog._resolve_endpoint() == (baked.URL, "baked")

    monkeypatch.setenv("CALIST_RUNLOG_URL", "http://127.0.0.1:9/exec")
    monkeypatch.setenv("CALIST_RUNLOG_TOKEN", "mine")
    assert runlog._resolve_endpoint() == ("http://127.0.0.1:9/exec", "mine")


def test_an_address_without_its_token_is_no_address(monkeypatch, tmp_path):
    """Every record would be refused, and a refused record is deleted: without
    the token there is nowhere to send, rather than somewhere to lose them."""
    monkeypatch.setenv("CALIST_RUNLOG_URL", "http://127.0.0.1:9/exec")
    assert runlog._resolve_endpoint() is None
    runlog.record_build(saved_result(), "", 1.0)
    assert not (tmp_path / "local").exists()

    monkeypatch.setitem(sys.modules, "calist_endpoint", SimpleNamespace(
        URL=f"https://script.google.com/macros/s/{FAKE_ID}/exec", TOKEN=""))
    monkeypatch.delenv("CALIST_RUNLOG_URL")
    assert runlog._resolve_endpoint() is None


def test_the_check_says_when_only_the_token_is_missing(monkeypatch, capsys):
    monkeypatch.setenv("CALIST_RUNLOG_URL", "http://127.0.0.1:9/exec")
    assert runlog.check() == 2
    assert "CALIST_RUNLOG_TOKEN is not set" in capsys.readouterr().out


def test_a_baked_loopback_address_is_refused(monkeypatch):
    """Loopback is for the tests and the owner's own override, never a build."""
    monkeypatch.setitem(sys.modules, "calist_endpoint",
                        SimpleNamespace(URL="http://127.0.0.1:9/exec", TOKEN="x"))
    assert runlog._resolve_endpoint() is None


@pytest.mark.skipif(getattr(sys, "frozen", False), reason="checkout only")
def test_a_checkout_has_no_baked_address():
    import importlib.util
    assert importlib.util.find_spec("calist_endpoint") is None


def test_without_an_address_nothing_is_made_or_started(tmp_path):
    before = threading.active_count()
    runlog.record_build(saved_result(), "", 1.0)
    runlog.record_lockout(6, 30)
    runlog.record_crash("tk", ValueError, ValueError("x"), None)
    runlog.start()
    runlog.hold()
    runlog.release()

    assert not (tmp_path / "local").exists()
    assert runlog._thread is None
    assert threading.active_count() == before


# ── the outbox ────────────────────────────────────────────────────────────────

@pytest.fixture
def address(monkeypatch):
    """An address nobody answers at: records stay in the outbox."""
    monkeypatch.setenv("CALIST_RUNLOG_URL", "http://127.0.0.1:9/exec")
    monkeypatch.setenv("CALIST_RUNLOG_TOKEN", TOKEN)


def outbox(tmp_path):
    return tmp_path / "local" / "Calist" / "outbox"


def test_a_record_appears_whole_or_not_at_all(address, tmp_path, monkeypatch):
    real_replace = os.replace

    def refuse(*_args):
        raise OSError("disk full")

    monkeypatch.setattr(os, "replace", refuse)
    runlog.record_build(saved_result(), "", 1.0)            # must not raise
    monkeypatch.setattr(os, "replace", real_replace)

    assert runlog.outbox_names() == []
    stray = list(outbox(tmp_path).glob("*.tmp"))
    assert len(stray) == 1

    old = time.time() - 7200
    os.utime(stray[0], (old, old))
    runlog._sweep_temporaries(str(outbox(tmp_path)))
    assert not list(outbox(tmp_path).glob("*.tmp"))


def test_records_wait_oldest_first(address, monkeypatch):
    ticks = iter(range(1_760_000_000, 1_770_000_000, 7))     # 2025, a tick apart
    monkeypatch.setattr(runlog.time, "time", lambda: float(next(ticks)))
    for _ in range(3):
        runlog.record_build(saved_result(), "", 1.0)

    names = runlog.outbox_names()
    assert len(names) == 3 and names == sorted(names)


def test_each_kind_keeps_its_own_allowance_and_counts_what_it_drops(address,
                                                                     monkeypatch):
    monkeypatch.setattr(runlog, "CAPS", {"build": 3, "crash": 1, "lockout": 1,
                                         "launch": 1})
    for _ in range(5):
        runlog.record_build(saved_result(), "", 1.0)
    runlog._write(runlog._envelope("crash", time.time(), {}))

    kinds = [runlog._kind_of(name) for name in runlog.outbox_names()]
    assert kinds.count("build") == 3 and kinds.count("crash") == 1
    assert runlog._load_state()["dropped"] == {"build": 2}
    assert runlog._health()["dropped"] == {"build": 2}


def test_records_older_than_the_limit_are_dropped(address, tmp_path):
    folder = outbox(tmp_path)
    folder.mkdir(parents=True)
    ancient = folder / f"{1_000_000_000_000:013d}-build-{'a' * 32}.json"
    ancient.write_text("{}")

    runlog.record_build(saved_result(), "", 1.0)

    assert not ancient.exists()
    assert runlog._load_state()["dropped"] == {"build": 1}


def test_an_unwritable_data_folder_is_shrugged_off(address, tmp_path, monkeypatch):
    blocker = tmp_path / "not-a-folder"
    blocker.write_text("")
    monkeypatch.setenv("LOCALAPPDATA", str(blocker))
    runlog.record_build(saved_result(), "", 1.0)
    runlog.record_crash("build", OSError, OSError("x"), None)


# ── the receipt ───────────────────────────────────────────────────────────────

ID = "c" * 32


@pytest.mark.parametrize("reply, receipt", [
    ({"ok": True, "id": ID, "dup": False}, True),
    ({"ok": True, "id": ID, "dup": True}, True),           # a duplicate, with its id
    ({"ok": True, "dup": True}, False),                     # a bare duplicate
    ({"ok": True, "id": "d" * 32}, False),                  # someone else's
    ({"ok": "true", "id": ID}, False),
    ({"ok": False, "id": ID, "retry": True}, False),
    (["ok", ID], False),
    (None, False),
])
def test_only_this_records_own_id_is_a_receipt(reply, receipt):
    assert runlog.is_ack(reply, ID) is receipt


@pytest.mark.parametrize("status, body", [
    (200, b"<html><body>Sign in to continue</body></html>"),
    (200, b""),
    (204, b""),
    (200, b'["ok"]'),
    (200, b'{"ok": true}' + b" " * runlog.MAX_REPLY),
    (500, b'{"ok": true, "id": "x"}'),
], ids=["sign-in page", "empty", "204", "a list", "too long", "server error"])
def test_a_page_that_is_not_the_scripts_answer_is_no_reply(status, body):
    assert runlog.read_reply(status, body) is None


def test_backoff_doubles_from_a_minute_to_an_hour(monkeypatch):
    monkeypatch.setattr(runlog, "BACKOFF_BASE", 60.0)
    monkeypatch.setattr(runlog, "BACKOFF_CAP", 3600.0)
    waits = [runlog.backoff(n) for n in range(1, 9)]
    assert waits == [60, 120, 240, 480, 960, 1920, 3600, 3600]


def test_the_sender_exits_only_when_nothing_arrived_since_its_last_look():
    runlog._thread = threading.current_thread()
    runlog._wake.set()                       # a record written after the last look
    assert runlog._finished() is False
    assert runlog._thread is threading.current_thread()

    runlog._pending.put(("lockout", 0.0, {}))
    assert runlog._finished() is False
    runlog._drop_queued()

    assert runlog._finished() is True
    assert runlog._thread is None


# ── the fake web app ──────────────────────────────────────────────────────────

class FakeWebApp:
    """tools/runlog_webapp/Code.gs as Google serves it, on 127.0.0.1."""

    def __init__(self):
        self.rows = {}                 # the sheet: record id -> record
        self.posts = []
        self.gets = []
        self.mode = "ok"               # ok | html | 307 | slow | 404 | retry
        self.lose_acks = 0             # this many receipts are lost on the way
        self.reject_next = 0
        self.settings = {"stop": False, "stop_hours": 24, "min_interval_s": 0}
        self._answers = {}
        app = self

        class Handler(http.server.BaseHTTPRequestHandler):
            def log_message(self, *_args):
                pass

            def _reply(self, status, body=b"", kind="application/json"):
                self.send_response(status)
                self.send_header("Content-Type", kind)
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def do_POST(self):
                body = self.rfile.read(int(self.headers.get("Content-Length", 0)))
                app.posts.append({"body": body,
                                  "type": self.headers.get("Content-Type"),
                                  "agent": self.headers.get("User-Agent")})
                if app.mode == "404":           # an address that is no deployment
                    return self._reply(404, b"Sorry, unable to open the file", "text/html")
                if app.mode == "slow":
                    time.sleep(3)
                if app.mode == "html":
                    return self._reply(200, b"<html>Sign in</html>", "text/html")
                if app.mode == "307":
                    self.send_response(307)
                    self.send_header("Location", f"{app.base}/echo/307")
                    self.send_header("Content-Length", "0")
                    return self.end_headers()
                key = str(len(app.posts))
                app._answers[key] = app.answer(json.loads(body))
                self.send_response(302)
                self.send_header("Location", f"{app.base}/echo/{key}")
                self.send_header("Content-Length", "0")
                self.end_headers()

            def do_GET(self):
                if self.path.endswith("/exec"):    # doGet, the health check
                    if app.mode == "404":
                        return self._reply(404, b"Sorry, unable to open the file",
                                           "text/html")
                    return self._reply(200, json.dumps(
                        {"ok": True, "service": "calist-runlog", "v": 1}).encode())
                app.gets.append({"path": self.path,
                                 "length": self.headers.get("Content-Length")})
                answer = app._answers.pop(self.path.rsplit("/", 1)[-1], None)
                if answer is None:
                    return self._reply(404, b"Sorry, unable to open the file", "text/html")
                if app.lose_acks:
                    app.lose_acks -= 1
                    return self._reply(500, b"oops", "text/plain")
                self._reply(200, json.dumps(answer).encode())

        self.server = http.server.ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self.server.daemon_threads = True
        self.base = f"http://127.0.0.1:{self.server.server_address[1]}"
        self.url = f"{self.base}/macros/s/fake/exec"
        threading.Thread(target=self.server.serve_forever, daemon=True).start()

    def answer(self, record):
        """What Code.gs does with a record: reject, recognise, or append."""
        rid = record.get("id")
        if self.mode == "retry":
            return {"ok": False, "retry": True, "id": rid, "why": "busy",
                    "cfg": self.settings}
        if record.get("token") != TOKEN:          # Code.gs: the token first ...
            return {"ok": False, "reject": True, "id": rid, "why": "token",
                    "cfg": self.settings}
        install = (record.get("pc") or {}).get("install", "")
        if not re.fullmatch(r"[0-9a-f]{16}", str(install)):       # ... then this
            return {"ok": False, "reject": True, "id": rid, "why": "install",
                    "cfg": self.settings}
        if self.reject_next:                      # a record the script can't take
            self.reject_next -= 1
            return {"ok": False, "reject": True, "id": rid, "why": "kind",
                    "cfg": self.settings}
        if rid in self.rows:
            return {"ok": True, "id": rid, "dup": True, "pc": 1, "cfg": self.settings}
        self.rows[rid] = record
        return {"ok": True, "id": rid, "dup": False, "pc": 1, "cfg": self.settings}

    def stop(self):
        self.server.shutdown()
        self.server.server_close()


@pytest.fixture
def webapp(monkeypatch):
    app = FakeWebApp()
    monkeypatch.setenv("CALIST_RUNLOG_URL", app.url)
    monkeypatch.setenv("CALIST_RUNLOG_TOKEN", TOKEN)
    yield app
    runlog._reset()                     # the sender goes before the server does
    app.stop()


def test_a_build_reaches_the_sheet_once(webapp):
    runlog.record_build(saved_result(), "", 1.5)
    runlog.kick()

    assert wait_for(settled)
    assert len(webapp.rows) == 1 and len(webapp.posts) == 1
    record = next(iter(webapp.rows.values()))
    assert record["kind"] == "build" and record["token"] == TOKEN
    assert record["data"]["devices"] == 3 and record["data"]["took_s"] == 1.5
    assert record["health"] == {"dropped": {}, "last_error": ""}
    assert webapp.posts[0]["type"] == "application/json"
    assert webapp.posts[0]["agent"] == f"Calist/{calist.__version__}"
    # urllib followed Google's 302 with a GET that carries no body.
    assert len(webapp.gets) == 1 and webapp.gets[0]["length"] in (None, "0")


def test_a_lost_receipt_is_answered_as_a_duplicate_not_a_second_row(webapp):
    webapp.lose_acks = 1
    runlog.record_build(saved_result(), "", 1.0)
    runlog.kick()

    assert wait_for(settled)
    assert len(webapp.rows) == 1
    assert len(webapp.posts) == 2          # sent again, recognised, deleted


def test_a_page_that_says_200_is_not_a_receipt(webapp):
    webapp.mode = "html"                   # a hotel or hospital login page
    runlog.record_build(saved_result(), "", 1.0)
    runlog.kick()

    assert wait_for(lambda: len(webapp.posts) >= 2)
    assert len(runlog.outbox_names()) == 1
    assert runlog._state["last_error"] == "not a receipt (HTTP 200)"

    webapp.mode = "ok"
    assert wait_for(settled)
    record = next(iter(webapp.rows.values()))
    # The row that finally got through says why the earlier tries failed.
    assert record["health"]["last_error"].startswith("not a receipt (HTTP 200) ×")
    assert "last_error" not in runlog._load_state()


def test_a_307_is_not_followed(webapp):
    webapp.mode = "307"
    runlog.record_build(saved_result(), "", 1.0)
    runlog.kick()

    assert wait_for(lambda: len(webapp.posts) >= 1)
    assert wait_for(lambda: (runlog._state or {}).get("last_error")
                    == "HTTP 307 from 127.0.0.1")
    assert webapp.gets == [] and len(runlog.outbox_names()) == 1


def test_the_scripts_own_reason_is_kept(webapp):
    """'Try again' from the script (its lock was busy) says why in state.json."""
    webapp.mode = "retry"
    runlog.record_build(saved_result(), "", 1.0)
    runlog.kick()

    assert wait_for(lambda: (runlog._state or {}).get("last_error") == "script: busy")
    assert len(runlog.outbox_names()) == 1


# ── python runlog.py --check ──────────────────────────────────────────────────

def test_the_check_without_an_address_says_how_to_set_one(capsys):
    assert runlog.check() == 2
    assert "CALIST_RUNLOG_URL" in capsys.readouterr().out


def test_the_check_sends_what_is_waiting_and_says_so(webapp, capsys):
    runlog.record_build(saved_result(), "", 1.0)
    runlog._write(runlog._envelope("launch", time.time(), {}))

    assert runlog.check() == 0
    out = capsys.readouterr().out
    assert "web app   answers" in out
    assert out.count("  sent") == 2 and "all sent" in out
    assert len(webapp.rows) == 2 and runlog.outbox_names() == []
    assert TOKEN not in out                       # the token is never printed


def test_the_check_tests_the_token_before_sending_anything(webapp, monkeypatch,
                                                          capsys):
    """The owner's second try: the token was the instructions' placeholder."""
    monkeypatch.setenv("CALIST_RUNLOG_TOKEN", "PASTE-YOUR-TOKENS-VALUE-HERE")
    runlog.record_build(saved_result(), "", 1.0)

    assert runlog.check() == 1
    out = capsys.readouterr().out
    assert "token     REFUSED by the script" in out
    assert "still the placeholder" in out and "Nothing was sent" in out
    assert len(runlog.outbox_names()) == 1 and webapp.rows == {}


def test_the_check_proves_the_token_with_nothing_waiting_and_writes_no_row(
        webapp, capsys):
    assert runlog.check() == 0
    out = capsys.readouterr().out
    assert "token     accepted by the script" in out and "nothing waiting" in out
    assert webapp.rows == {}


def test_the_check_never_calls_a_refusal_sent(webapp, capsys):
    """It once printed 'all sent' after two refusals."""
    webapp.reject_next = 1
    runlog.record_build(saved_result(), "", 1.0)

    assert runlog.check() == 1
    out = capsys.readouterr().out
    assert "refused (kind)" in out and "0 sent, 1 dropped" in out
    assert "all sent" not in out


def test_the_check_names_an_address_that_is_no_deployment(webapp, capsys):
    """What the owner's first try hit: a 404 means the address is not a live
    Web app deployment, and the records stay for the right one."""
    webapp.mode = "404"
    runlog.record_build(saved_result(), "", 1.0)

    assert runlog.check() == 1
    out = capsys.readouterr().out
    assert "HTTP 404 from 127.0.0.1" in out and "Manage deployments" in out
    assert len(runlog.outbox_names()) == 1


#: Built from parts: a deployment-shaped address written out here would trip
#: test_no_deployment_address_is_committed, which is the point of that test.
FAKE_ID = "AKfycb" + "x" * 24 + "0123"


@pytest.mark.parametrize("url, shown", [
    (f"https://script.google.com/macros/s/{FAKE_ID}/exec",
     "https://script.google.com/macros/s/AKfycb...0123/exec"),
    ("http://127.0.0.1:9/exec", "http://127.0.0.1:9/exec"),
])
def test_the_check_shows_only_enough_of_the_address_to_compare(url, shown):
    assert runlog.masked(url) == shown


def test_a_slow_server_costs_a_bounded_wait(webapp, monkeypatch):
    monkeypatch.setattr(runlog, "TIMEOUT_S", 0.3)
    webapp.mode = "slow"
    runlog.record_build(saved_result(), "", 1.0)
    started = time.monotonic()
    runlog.kick()

    assert wait_for(lambda: (runlog._state or {}).get("last_error"), timeout=3)
    assert time.monotonic() - started < 2.5
    assert len(runlog.outbox_names()) == 1


def test_a_refused_record_is_dropped_and_counted_then_reported(webapp):
    webapp.reject_next = 1
    runlog.record_build(saved_result(), "", 1.0)
    runlog.kick()
    assert wait_for(settled)
    assert webapp.rows == {}
    assert runlog._load_state()["dropped"] == {"build": 1}

    runlog.record_build(saved_result(), "", 1.0)
    runlog.kick()
    assert wait_for(settled)
    record = next(iter(webapp.rows.values()))
    assert record["health"]["dropped"] == {"build": 1}
    assert runlog._load_state()["dropped"] == {}


def test_a_wrong_token_keeps_the_records_until_it_is_right(webapp, monkeypatch):
    """What the owner's first real check hit: refused for the token, both
    records deleted. A token is a setting to fix, not a bad record."""
    monkeypatch.setenv("CALIST_RUNLOG_TOKEN", "not-the-token")
    runlog.record_build(saved_result(), "", 1.0)
    runlog.kick()

    assert wait_for(lambda: (runlog._state or {}).get("last_error")
                    == "the script refused the token")
    assert len(runlog.outbox_names()) == 1 and webapp.rows == {}

    runlog._reset()                                    # the next start, token fixed
    monkeypatch.setenv("CALIST_RUNLOG_TOKEN", TOKEN)
    runlog.kick()
    assert wait_for(settled)
    record = next(iter(webapp.rows.values()))
    assert record["health"]["dropped"] == {}           # nothing was lost
    assert record["health"]["last_error"].startswith("the script refused the token ×")


def test_the_sheet_can_stop_every_copy_and_it_lasts_a_restart(webapp, tmp_path):
    webapp.settings = {"stop": True, "stop_hours": 1, "min_interval_s": 0}
    runlog.record_build(saved_result(), "", 1.0)
    runlog.kick()
    assert wait_for(settled)
    assert runlog._load_state()["disable_until"] > time.time() + 3000

    runlog._reset()                                       # a new process
    runlog.record_build(saved_result(), "", 1.0)
    runlog.record_lockout(6, 30)
    runlog.start()
    assert wait_for(lambda: runlog._thread is None)
    assert runlog.outbox_names() == [] and len(webapp.posts) == 1


def test_nothing_is_sent_while_a_build_runs(webapp):
    runlog.hold()
    runlog.record_build(saved_result(), "", 1.0)
    runlog.kick()
    time.sleep(0.3)
    assert webapp.posts == []

    runlog.release()
    assert wait_for(settled)
    assert len(webapp.rows) == 1


def test_the_tk_thread_only_queues(webapp, monkeypatch):
    """Lockouts, Tk crashes and the launch come from the Tk thread; the files
    are written by the sender, never there."""
    writers = []
    real_write = runlog._write

    def watched(record):
        writers.append(threading.current_thread() is threading.main_thread())
        real_write(record)

    monkeypatch.setattr(runlog, "_write", watched)
    runlog.record_lockout(7, 60)
    try:
        raise KeyError("x")
    except KeyError as exc:
        runlog.record_crash("tk", type(exc), exc, exc.__traceback__)
    runlog.start()

    assert wait_for(lambda: len(webapp.rows) == 3)
    assert writers and not any(writers)
    kinds = sorted(r["kind"] for r in webapp.rows.values())
    assert kinds == ["crash", "launch", "lockout"]
    lockout = next(r for r in webapp.rows.values() if r["kind"] == "lockout")
    assert lockout["data"] == {"failures": 7, "wait_s": 60}


# ── importing it costs nothing ────────────────────────────────────────────────

def test_importing_runlog_is_light_and_does_nothing(tmp_path):
    env = dict(os.environ, LOCALAPPDATA=str(tmp_path),
               CALIST_RUNLOG_URL="http://127.0.0.1:9/exec")
    out = subprocess.run(
        [sys.executable, "-c",
         "import sys, threading, runlog; "
         "heavy = [m for m in ('tkinter', 'customtkinter', 'ssl', 'http.client', "
         "'urllib.request', 'email', 'uuid', 'getpass', 'socket', 'calist', "
         "'openpyxl') if m in sys.modules]; "
         "print(heavy, threading.active_count())"],
        capture_output=True, text=True, env=env, cwd=ROOT)
    assert out.returncode == 0, out.stderr
    assert out.stdout.strip() == "[] 1"
    assert list(tmp_path.iterdir()) == []


def test_no_deployment_address_is_committed():
    """The repository is public: a real /exec address would be anyone's to
    write to. It lives in a CI secret; placeholders here stay short."""
    try:
        listed = subprocess.run(["git", "ls-files", "-z"], capture_output=True,
                                cwd=ROOT, check=True).stdout
    except (OSError, subprocess.CalledProcessError):
        pytest.skip("not a git checkout")
    deployment = re.compile(r"/macros/s/[A-Za-z0-9_-]{20,}")
    texts = (".py", ".md", ".yml", ".yaml", ".gs", ".json", ".spec", ".txt")
    found = [name for name in listed.decode("utf-8").split("\0")
             if name.endswith(texts) and (ROOT / name).is_file()
             and deployment.search((ROOT / name).read_text("utf-8", errors="ignore"))]
    assert found == []
