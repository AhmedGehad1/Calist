"""Tests for firebase_export.py. Run with: python -m pytest

Only the pure parts -- nothing here opens a network connection or touches
Firebase. The module's other behaviour is exercised by its own --dry-run and
--verify-maps modes against the real archive.
"""

import os
from pathlib import Path

import pytest

from device_config import DEVICE_CONFIGS
import firebase_export
from firebase_export import reachable


# ── reachable: the Windows MAX_PATH escape ────────────────────────────────────
#
# Six workbooks in the archive nest past 260 characters, because an automated
# site files its ultrasound probes under folders carrying the full model and
# serial. os.walk finds them; stat and open then raise FileNotFoundError on a
# file that is plainly there. That killed a whole import run at the
# cost-estimate stage, after 34 minutes of reading and before a single document
# was written.

WINDOWS_ONLY = pytest.mark.skipif(os.name != "nt", reason="MAX_PATH is a Windows limit")


def test_short_path_is_left_alone():
    # The prefix is ugly and leaks into error messages, so it is only applied
    # where it is actually needed.
    assert reachable("D:/MedCal Pro/x.xlsx") == "D:/MedCal Pro/x.xlsx"


@WINDOWS_ONLY
def test_long_path_gets_the_prefix():
    long = "D:/MedCal Pro/" + "a" * 260 + "/x.xlsx"
    assert reachable(long).startswith("\\\\?\\")


@WINDOWS_ONLY
def test_prefix_is_not_applied_twice():
    once = reachable("D:/MedCal Pro/" + "a" * 260 + "/x.xlsx")
    assert reachable(once) == once


@WINDOWS_ONLY
def test_prefixed_path_is_absolute():
    # \\?\ disables all path normalisation, so a relative path behind it would
    # never resolve -- it has to be made absolute first.
    assert "/" not in reachable("./" + "a" * 260 + "/x.xlsx").replace("\\\\?\\", "")


@WINDOWS_ONLY
def test_a_deep_file_can_actually_be_read(tmp_path):
    """The point of the whole thing: a real file past the limit opens."""
    deep = tmp_path
    while len(str(deep)) < 270:
        deep = deep / "nested-folder-name"
    os.makedirs(reachable(str(deep)), exist_ok=True)

    target = str(deep / "workbook.xlsx")
    assert len(target) > 260, "test needs a path past the limit to be meaningful"

    with open(reachable(target), "wb") as handle:
        handle.write(b"x" * 1234)

    with pytest.raises(OSError):
        Path(target).stat()          # what used to happen
    assert Path(reachable(target)).stat().st_size == 1234


# ── device aliases ────────────────────────────────────────────────────────────
#
# VAGH and GJAF are typos rather than device types. Both were reported as
# unrecognised and imported without readings until their layouts were read off
# the forms.


def test_vagh_shares_agh_cells_by_reference():
    # Shared by reference so the two can never drift apart, exactly as AG does.
    assert DEVICE_CONFIGS["VAGH"]["cells"] is DEVICE_CONFIGS["AGH"]["cells"]


def test_vagh_rewrites_its_own_code_for_the_second_row():
    # AGH's ("AGH" -> "AGCB") would turn VAGH001 into VAGCB001, so VAGH needs
    # its own substitution or the NIBP row silently keeps the parent's code.
    was, now = DEVICE_CONFIGS["VAGH"]["second_row"]["code_replace"]
    assert "VAGH001".replace(was, now) == "AGCB001"


def test_gjaf_is_an_aortic_balloon_not_an_ecg():
    # The form prints "ECG Performance Test", because a balloon pump triggers
    # off the ECG signal. That heading is a section, not the device -- reading
    # it as the identity gives ECG, which is wrong.
    assert DEVICE_CONFIGS["GJAF"]["device_name"] == "Aortic balloon"
    assert DEVICE_CONFIGS["GJAF"]["device_name"] != DEVICE_CONFIGS["AF"]["device_name"]


def test_gjaf_does_not_use_af_cells():
    # A different device, and a different layout: AF puts the Date at D30 and
    # the Status at F41.
    assert DEVICE_CONFIGS["GJAF"]["cells"] != DEVICE_CONFIGS["AF"]["cells"]
    assert DEVICE_CONFIGS["GJAF"]["cells"]["Date"] == "D28"


def test_gjaf_carries_its_second_layout_as_an_alternate():
    # These files come in two layouts, one anchored at row 32 and a later one
    # at row 60.
    dates = [alt["Date"] for alt in DEVICE_CONFIGS["GJAF"]["alt_cells"]]
    assert "D56" in dates


# ── tombstones ────────────────────────────────────────────────────────────────
#
# When an engineer deletes a calibration in the app, the record goes and a
# tombstone is written in its place. Without one the deletion would not stick:
# this import is idempotent and keyed on the filename, so the next run recreates
# whatever is on disk and the engineer watches a device they removed come back.


class _FakeSnap:
    def __init__(self, doc_id):
        self.id = doc_id


class _FakeCollection:
    def __init__(self, ids, error=None):
        self._ids = ids
        self._error = error

    def stream(self):
        if self._error:
            raise self._error
        return [_FakeSnap(i) for i in self._ids]


class _FakeDb:
    def __init__(self, ids=(), error=None):
        self._ids = ids
        self._error = error

    def collection(self, name):
        assert name == "deletions"
        return _FakeCollection(self._ids, self._error)


def test_tombstones_are_read_as_a_set_of_ids():
    assert firebase_export.load_tombstones(_FakeDb(["G302-BP001-0326", "x"])) == {
        "G302-BP001-0326",
        "x",
    }


def test_no_tombstones_is_an_empty_set_not_an_error():
    assert firebase_export.load_tombstones(_FakeDb()) == set()


def test_an_unreadable_tombstone_collection_does_not_stop_the_import(capsys):
    """Fails soft, loudly.

    Being unable to honour a deletion is a smaller problem than being unable to
    import at all -- but a silent empty set would be indistinguishable from
    "nobody has deleted anything", so it has to say so.
    """
    result = firebase_export.load_tombstones(_FakeDb(error=RuntimeError("no auth")))

    assert result == set()
    out = capsys.readouterr().out
    assert "WARNING" in out
    assert "may reappear" in out


# ── the write guard ───────────────────────────────────────────────────────────


def test_upload_only_refuses_without_yes(tmp_path, capsys):
    """The write guard has to hold on the new mode too.

    --upload-only skips the archive scan, so it reaches the point of writing far
    sooner than --push does. It must still refuse without --yes, and it must
    refuse *before* opening a connection — which is what makes this testable
    with no credentials present.
    """
    assert firebase_export.main([str(tmp_path), "--upload-only"]) == 3
    assert "Refusing to write without --yes" in capsys.readouterr().out
