"""Tests for firebase_export.py. Run with: python -m pytest

Only the pure parts -- nothing here opens a network connection or touches
Firebase. The module's other behaviour is exercised by its own --dry-run and
--verify-maps modes against the real archive.
"""

import os
from pathlib import Path

import pytest

from device_config import DEVICE_CONFIGS
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


def test_gjaf_is_an_ecg():
    assert DEVICE_CONFIGS["GJAF"]["device_name"] == DEVICE_CONFIGS["AF"]["device_name"]


def test_gjaf_does_not_share_af_cells():
    # Its layout genuinely differs -- AF puts the Date at D30, GJAF at D28 --
    # so sharing by reference would be wrong here even though it is right for
    # VAGH.
    assert DEVICE_CONFIGS["GJAF"]["cells"] != DEVICE_CONFIGS["AF"]["cells"]
    assert DEVICE_CONFIGS["GJAF"]["cells"]["Date"] == "D28"


def test_gjaf_carries_its_second_layout_as_an_alternate():
    # These files come in two layouts, one anchored at row 32 and a later one
    # at row 60.
    dates = [alt["Date"] for alt in DEVICE_CONFIGS["GJAF"]["alt_cells"]]
    assert "D56" in dates
