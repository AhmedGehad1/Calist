"""Tests for firebase_export.py. Run with: python -m pytest

Only the pure parts -- nothing here opens a network connection or touches
Firebase. The module's other behaviour is exercised by its own --dry-run and
--verify-maps modes against the real archive.
"""

import os
import sys
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

    # The assertion that matters: the prefixed path always opens.
    assert Path(reachable(target)).stat().st_size == 1234

    # Whether the *bare* path fails is a property of the machine, not of the
    # code: MAX_PATH is lifted by the LongPathsEnabled policy, which the GitHub
    # runner sets and most workstations do not. Asserting it unconditionally
    # turned a green suite red on CI and blocked a release, so it is recorded
    # only where it actually applies.
    try:
        Path(target).stat()
    except OSError:
        pass                          # the usual case, and why reachable exists
    else:
        assert sys.getwindowsversion().major >= 10, "long paths on an old Windows"


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


# ── the form header ───────────────────────────────────────────────────────────
#
# Every form names its client and its contact, and none of it was ever imported:
# clientAddress, contactName and contactPhone were empty on all 82,287 records,
# because the only source consulted was Hospital Codes.xlsx.
#
# Nothing about the position is stable -- the client name sits at C3, C5, C6,
# C11 or D6 depending on the device type, the contact at A30, G32, F44, H42 or
# B28. The *labels* are identical across every form in the archive, so the
# fields are found by label and one rule covers all 57 maps.


def test_finds_the_name_and_the_address_two_rows_below():
    grid = {
        "C6": "Client Name:", "E6": "Alpha Eye Center",
        "C8": "Client Address:", "E8": "Tolba Awaida St, Zagazig",
    }
    assert firebase_export.header_from_grid(grid) == {
        "client_name": "Alpha Eye Center",
        "client_address": "Tolba Awaida St, Zagazig",
        "contact_name": "",
        "contact_phone": "",
    }


def test_the_label_may_sit_anywhere():
    # Same rule, a different device type's layout.
    grid = {
        "D11": "Client Name", "F11": "Sohag Oncology Institute",
        "D13": "Client Address", "F13": "Kornish Al Nile, Sohag",
    }
    assert firebase_export.header_from_grid(grid)["client_name"] == "Sohag Oncology Institute"


def test_the_labels_own_merged_span_is_skipped():
    """The reason this is not just "take the next non-empty cell".

    Calist resolves merged cells, so a label merged across C6:D6 reports its own
    text for D6 as well. Taking the first non-empty neighbour would return
    "Client Name:" as the hospital's name -- which is exactly what the first
    version of this did.
    """
    grid = {
        "C6": "Client Name:", "D6": "Client Name:", "E6": "Alpha Eye Center",
        "C8": "Client Address:", "D8": "Client Address:", "E8": "Zagazig",
    }
    out = firebase_export.header_from_grid(grid)
    assert out["client_name"] == "Alpha Eye Center"
    assert out["client_address"] == "Zagazig"


def test_a_form_with_no_client_header_yields_nothing():
    # Some device types put their test data on the sheet the reader picks and
    # the header elsewhere. The caller falls back to the code list.
    assert not any(firebase_export.header_from_grid({"A16": "Test parameter (mmHg)"}).values())


def test_a_name_with_no_address_still_gives_the_name():
    out = firebase_export.header_from_grid({"C6": "Client Name:", "E6": "Alpha Eye Center"})
    assert out["client_name"] == "Alpha Eye Center"
    assert out["client_address"] == ""


def test_finds_the_contact_and_their_number():
    # Contacts sit far lower than the client header and often in column A --
    # A30, G32, F44, H42, B28 across five device types -- so position tells you
    # nothing and the label tells you everything.
    grid = {
        "A30": "Contact Person Name:", "C30": "Mohamed Sayed",
        "A31": "Phone No.:", "C31": "01284466683",
    }
    out = firebase_export.header_from_grid(grid)
    assert out["contact_name"] == "Mohamed Sayed"
    assert out["contact_phone"] == "01284466683"


def test_a_phone_number_keeps_its_leading_zero():
    # Read as text, never coerced to a number -- the same reason the Excel
    # writer uses TextCellValue for phones.
    grid = {"H43": "Phone No.:", "J43": "01204715812"}
    assert firebase_export.header_from_grid(grid)["contact_phone"] == "01204715812"


def test_client_and_contact_are_found_in_the_same_pass():
    grid = {
        "C6": "Client Name:", "E6": "Alpha Eye Center",
        "C8": "Client Address:", "E8": "Zagazig",
        "A30": "Contact Person Name:", "C30": "Mohamed Sayed",
        "A31": "Phone No.:", "C31": "01284466683",
    }
    out = firebase_export.header_from_grid(grid)
    assert all(out.values()), out


def test_the_code_list_wins_on_the_name_and_the_form_supplies_the_address():
    # The code list carries the official spelling, but has no addresses at all.
    customer = firebase_export.build_customer(
        "G302", {"name": "Official Name"}, "hospital", ("Form Name", "Some Street")
    )
    assert customer["name"] == "Official Name"
    assert customer["address"] == "Some Street"


def test_the_form_names_a_site_the_code_list_has_never_heard_of():
    # 25 sites are missing from Hospital Codes.xlsx entirely, carrying 7,610
    # calibrations between them. Their forms know perfectly well who they are.
    customer = firebase_export.build_customer("F22", None, "hospital", ("Isis", "Luxor"))
    assert customer["name"] == "Isis"
    assert customer["address"] == "Luxor"


def test_a_rebuilt_document_never_mentions_storagePath():
    """Because every write is a merge, and None would overwrite a real path.

    push_storage sets storagePath once a workbook is actually uploaded. If
    build_document sent an explicit None, a plain re-import would wipe it from
    every record that had one -- 41,818 of them -- and silently detach the whole
    archive from its files. A field that is not mentioned survives a merge.
    """
    form = firebase_export.ParsedForm(
        path="D:/x/G302-BP001-0326.xlsx",
        filename="G302-BP001-0326.xlsx",
        year=2026,
        site="G302",
        tag="BP001",
        device_code="BP",
        serial="SN-1",
    )
    assert "storagePath" not in firebase_export.build_document(form, None)


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
