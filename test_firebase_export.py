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


def test_vagh_reads_its_verdicts_off_the_vital_signs_template():
    """VAGH's cells are VAH's, not AGH's, however its code and name read.

    These files were sharing AGH's map by reference, on the assumption that a
    stray "V" was the only difference. It was not: all 12 are filled in on the
    vital signs template and carry their verdicts at G38/J38, while AGH's
    D39/J39 are empty on every one -- so both statuses imported blank.

    Nothing caught it because the identity block is identical on the two
    templates, so the record looked entirely plausible.
    """
    cells = DEVICE_CONFIGS["VAGH"]["cells"]
    assert cells["Status"] == DEVICE_CONFIGS["VAH"]["cells"]["Status"] == "G38"
    assert cells["Status2"] == DEVICE_CONFIGS["VAH"]["cells"]["Status2"] == "J38"
    assert cells["Status"] != DEVICE_CONFIGS["AGH"]["cells"]["Status"]


def test_an_alternate_that_repeats_the_primary_identity_is_never_added():
    """An alternate is only tried when the whole record is implausible.

    So one whose Model and S.N match the primary's can never be reached -- it
    is dead weight that reads as a written-down layout. Two used to exist: BZ's
    fourth entry and VAGH's, both there only to name a different Status cell.
    A Status box that moves on its own is found by its printed label instead.
    """
    for code, config in DEVICE_CONFIGS.items():
        identity = ("Model", "S.N", "Manufacturer", "Location")
        primary = tuple(config["cells"].get(f, "") for f in identity)
        for index, alternate in enumerate(config.get("alt_cells", [])):
            assert tuple(alternate.get(f, "") for f in identity) != primary, \
                f"{code} alt {index} repeats the primary's identity cells"


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


# ── the app's own uploads ─────────────────────────────────────────────────────
#
# The MedCal Pro app now puts the certificate it just filed into the bucket
# under `filed/`, so a colleague can open it. This tool never writes there --
# it only clears the twin, because there are two moments when nothing else can:
# when the archive copy supersedes it, and when the engineer deleted the record.


def test_a_filed_path_matches_the_app_that_writes_it():
    """Byte-for-byte what `filedWorkbookPath` in upload_queue.dart produces.

    The two are related only by agreeing on this string. Nothing checks it at
    build time, so it is checked here.
    """
    assert firebase_export.filed_blob_name("B03-SN21953-2026", "B03", 2026) == (
        "filed/2026/B03/B03-SN21953-2026.xlsx"
    )


def test_a_filed_path_is_never_the_archive():
    """The one property that must hold whatever gets passed in.

    A record id and a customer code both reach this from data, so neither can
    be trusted to be a clean path segment.
    """
    hostile = firebase_export.filed_blob_name("../../etc/passwd", "../archive", 2026)

    assert hostile.startswith("filed/2026/")
    assert ".." not in hostile
    assert hostile.count("/") == 3


def test_an_empty_customer_code_still_produces_a_usable_key():
    assert firebase_export.filed_blob_name("r1", "  ", 2026) == (
        "filed/2026/UNKNOWN/r1.xlsx"
    )


class _FakeBlob:
    def __init__(self, present):
        self._present = present
        self.deleted = False

    def exists(self):
        return self._present

    def delete(self):
        self.deleted = True


class _FakeBucket:
    def __init__(self, present=True, raises=None):
        self._present = present
        self._raises = raises
        self.asked = []
        self.blobs = {}

    def blob(self, name):
        self.asked.append(name)
        if self._raises:
            raise self._raises
        self.blobs[name] = _FakeBlob(self._present)
        return self.blobs[name]


def test_clearing_a_twin_deletes_it_and_says_so():
    bucket = _FakeBucket(present=True)

    assert firebase_export.clear_filed_twin(bucket, "r1", "B03", 2026) is True
    assert bucket.asked == ["filed/2026/B03/r1.xlsx"]
    assert bucket.blobs["filed/2026/B03/r1.xlsx"].deleted


def test_a_record_that_never_had_a_twin_is_not_an_error():
    """The overwhelming majority. They came off the archive drive, not a phone."""
    bucket = _FakeBucket(present=False)

    assert firebase_export.clear_filed_twin(bucket, "r1", "B03", 2026) is False
    assert not bucket.blobs["filed/2026/B03/r1.xlsx"].deleted


def test_a_failed_cleanup_does_not_stop_the_import(capsys):
    """An orphaned object is untidy. A halted import is not."""
    bucket = _FakeBucket(raises=RuntimeError("bucket unreachable"))

    assert firebase_export.clear_filed_twin(bucket, "r1", "B03", 2026) is False
    assert "could not clear the filed copy" in capsys.readouterr().out


# ── outcomes ──────────────────────────────────────────────────────────────────
#
# The importer used to collapse each form's Status cell to `passed`, and read
# only the first box. A patient monitor recorded as "limited non" therefore
# imported as `passed: False`, and the app's previous-calibration lookup threw
# the whole record away — so a monitor the archive knew as AGH002 was filed as
# a brand-new AGH005. `statuses` carries what the form actually said.

from firebase_export import build_document, build_statuses, normalise_status


def _form(**fields):
    base = dict(
        path="D:/x/G181-AGH002-0825.xlsx",
        filename="G181-AGH002-0825.xlsx",
        year=2025,
        site="G181",
        tag="AGH002",
        device_code="AGH",
        serial="AQ-17157179",
    )
    base.update(fields)
    return firebase_export.ParsedForm(**base)


def test_the_four_outcomes_are_the_apps_own_strings():
    """Pinned against `CalibrationStatus.wire` in lib/models/calibration_status.dart.

    The app parses exactly these and nothing else. A drift here would not
    fail loudly: every backfilled outcome would quietly stop parsing and the
    app would fall back to `passed`, which is the bug this exists to fix.
    """
    assert (
        firebase_export.STATUS_PASS,
        firebase_export.STATUS_CALIBRATED,
        firebase_export.STATUS_LIMITED_NON,
        firebase_export.STATUS_LIMITED,
        firebase_export.STATUS_LIMITED_FAIL,
        firebase_export.STATUS_FAIL,
    ) == ("pass", "calibrated", "limited non", "limited", "limited fail", "fail")


def test_everything_passed_already_accepted_is_a_pass():
    # Nothing `passed` called a pass may become anything else.
    for value in firebase_export._PASS_VALUES:
        assert normalise_status(value.upper()) == "pass", value


@pytest.mark.parametrize("raw, expected", [
    ("Calibrated", "calibrated"),
    ("  CALIBRATED ", "calibrated"),
    ("Limited Non", "limited non"),
    ("limited none", "limited non"),
    ("  LIMITED   non ", "limited non"),
    ("Limited Fail", "limited fail"),
    ("limited-fail", "limited fail"),
    ("Fail", "fail"),
    ("Failed", "fail"),
    ("Faulty", "fail"),
])
def test_status_cells_read_as_the_outcome_they_name(raw, expected):
    assert normalise_status(raw) == expected


@pytest.mark.parametrize("raw", [None, "", "limited sometimes", "N.A", "----", "see comment"])
def test_a_cell_that_names_no_outcome_is_left_out_not_guessed(raw):
    # Guessing "pass" would put a date on next year's certificate that the
    # device never earned.
    assert normalise_status(raw) is None


def test_a_monitor_carries_both_its_verdicts():
    # D39 is ECG, J39 is NIBP. The importer never read J39 at all.
    document = build_document(_form(status="pass", status2="limited non"), None)
    assert document["statuses"] == {"ecg": "pass", "nibp": "limited non"}


def test_the_record_that_hid_agh002_now_says_what_it_was():
    document = build_document(_form(status="limited non", status2="pass"), None)
    # Unchanged, so a phone on the old build sees exactly what it saw before.
    assert document["passed"] is False
    # And the new build can tell a limitation from a failure.
    assert document["statuses"] == {"ecg": "limited non", "nibp": "pass"}


def test_a_monitor_box_that_names_nothing_is_dropped_on_its_own():
    document = build_document(_form(status="see comment", status2="fail"), None)
    assert document["statuses"] == {"nibp": "fail"}


def test_calibrated_is_its_own_outcome_and_leaves_passed_alone():
    form = _form(device_code="BP", tag="BP001", status="Calibrated")
    document = build_document(form, None)
    assert document["statuses"] == {"overall": "calibrated"}
    # What a phone on the old build sees is unchanged.
    assert document["passed"] is False


def test_a_calibrated_box_ranks_after_a_pass_but_before_a_limitation():
    assert build_statuses(_form(device_code="GC", status="Pass",
                                status2="Calibrated")) == {"overall": "calibrated"}


def test_a_single_status_form_is_one_overall_verdict():
    form = _form(device_code="BP", tag="BP001", status="Pass")
    assert build_statuses(form) == {"overall": "pass"}


def test_a_vital_signs_monitor_keeps_both_its_modules():
    # G38 is the SpO2 module, J38 the NIBP module. Collapsing them into one
    # verdict would hide which module failed.
    form = _form(device_code="VAH", status="OK", status2="Faulty")
    assert build_statuses(form) == {"spo2": "pass", "nibp": "fail"}


@pytest.mark.parametrize("code", ["AGH", "AG", "VAGH"])
def test_every_code_a_patient_monitor_is_filed_under_keeps_both(code):
    form = _form(device_code=code, status="limited fail", status2="pass")
    assert build_statuses(form) == {"ecg": "limited fail", "nibp": "pass"}


def test_every_form_with_a_second_status_box_says_what_both_boxes_are():
    """A cell map with a Status2 and no entry would collapse two verdicts to one.

    Nothing would fail: the device would import with a single worst-of status,
    and the history would never show which of its two modules had the problem.
    """
    def has_second_box(config):
        maps = [config["cells"], *config.get("alt_cells", [])]
        return any("Status2" in cells for cells in maps)

    two_box = {code for code, config in DEVICE_CONFIGS.items()
               if has_second_box(config)}
    assert two_box, "expected at least the patient monitor"
    missing = two_box - set(firebase_export._PER_TEST_STATUS)
    assert not missing, f"no meaning given for the second status box of {missing}"


def test_nothing_classifiable_leaves_statuses_out_of_the_write():
    # Absent, not empty: under a merge an absent field leaves the server's
    # answer standing, and the app falls back to `passed`.
    assert "statuses" not in build_document(_form(status="", status2=""), None)


def test_a_rebuilt_document_never_sends_an_empty_map():
    """An empty map under merge=True replaces what the server holds.

    `tests` and `formData` were sent as {} on every run, which would erase the
    form data of any imported record an engineer had since corrected in the app.
    """
    document = build_document(_form(status="pass"), None)
    assert "tests" not in document
    assert "formData" not in document
    assert not [key for key, value in document.items() if value == {}]


def test_the_histogram_counts_each_box_and_what_it_becomes():
    forms = [
        _form(status="Pass", status2="Limited Non"),
        _form(status="Pass", status2="see comment"),
        _form(device_code="BP", tag="BP001", status="OK"),
    ]
    rows, summary = firebase_export.status_histogram(forms)

    assert ("AGH", "Status", "Pass", "pass", 2) in rows
    assert ("AGH", "Status2", "Limited Non", "limited non", 1) in rows
    assert ("AGH", "Status2", "see comment", None, 1) in rows
    # BP has no second box, so none is counted for it.
    assert not [row for row in rows if row[0] == "BP" and row[1] == "Status2"]
    assert summary == {"forms": 3, "with_statuses": 3, "two_box_with_both": 1,
                       "by_label": 0}


# ── the statuses-only backfill ────────────────────────────────────────────────
#
# Re-running the full push to add one field would rewrite every field of every
# record and re-derive every device number — and if the files on disk have
# changed since the last import, a device could be renumbered. The backfill
# writes `statuses` and nothing else, and only onto a record that exists and
# whose serial matches.

from firebase_export import collect_statuses, push_statuses


class _Ref:
    def __init__(self, doc_id):
        self.id = doc_id


class _StoredSnap:
    def __init__(self, doc_id, data):
        self.id = doc_id
        self._data = data
        self.exists = data is not None

    def to_dict(self):
        return None if self._data is None else dict(self._data)


class _Batch:
    def __init__(self, db):
        self._db = db
        self._ops = []

    def update(self, ref, data):
        self._ops.append((ref.id, data))

    def commit(self):
        self._db.commits.append(list(self._ops))


class _StoreCollection:
    def __init__(self, db, name):
        self._db = db
        self._name = name

    def document(self, doc_id):
        return _Ref(doc_id)

    def stream(self):
        assert self._name == "deletions"
        return [_FakeSnap(i) for i in self._db.deleted]


class _StoreDb:
    """Just enough of firebase_admin's client: reads, batched updates, tombstones."""

    def __init__(self, docs, deleted=()):
        self.docs = docs
        self.deleted = list(deleted)
        self.commits = []
        self.read = []

    def collection(self, name):
        return _StoreCollection(self, name)

    def get_all(self, refs, field_paths=None):
        for ref in refs:
            self.read.append(ref.id)
            yield _StoredSnap(ref.id, self.docs.get(ref.id))

    def batch(self):
        return _Batch(self)

    @property
    def written(self):
        return {doc_id: data for commit in self.commits for doc_id, data in commit}


def _record(doc_id, serial="AQ-17157179", status="pass", status2="", code="AGH"):
    form = _form(device_code=code, serial=serial, status=status, status2=status2)
    form.doc_id = doc_id
    return form


def test_one_entry_per_record_and_the_worse_copy_wins():
    found = collect_statuses([
        _record("G181-AGH002-0825", "AQ-1", "pass", "pass"),
        # A duplicate copy of the same visit that says something worse.
        _record("G181-AGH002-0825", "AQ-1", "pass", "limited non"),
        _record("G181-AGH003-0825", "KN-9", "", ""),
    ])
    assert found == {
        "G181-AGH002-0825": ("AQ-1", {"ecg": "pass", "nibp": "limited non"},
                             {"D:/x/G181-AGH002-0825.xlsx"}),
    }


def test_a_form_that_never_got_an_id_is_not_collected():
    assert collect_statuses([_record("", status="pass")]) == {}


def test_only_the_statuses_field_is_written():
    db = _StoreDb({"A": {"serialUpper": "AQ-1", "siteDevice": "G181-AGH002"}})
    counts = push_statuses({"A": ("AQ-1", {"ecg": "pass", "nibp": "fail"})}, db)

    assert db.written == {"A": {"statuses": {"ecg": "pass", "nibp": "fail"}}}
    assert counts["updated"] == 1


def test_a_record_whose_serial_differs_is_left_alone():
    # The id shifted — a different device now answers to it. Writing would put
    # one monitor's verdict on another.
    db = _StoreDb({"A": {"serialUpper": "KN-97048765"}})
    counts = push_statuses({"A": ("AQ-17157179", {"ecg": "fail"})}, db)

    assert db.written == {}
    assert counts["mismatched"] == 1


def test_a_record_the_server_does_not_have_is_never_created():
    db = _StoreDb({})
    counts = push_statuses({"A": ("AQ-1", {"ecg": "pass"})}, db)

    assert db.written == {}
    assert counts["missing"] == 1


def test_a_deleted_record_is_neither_read_nor_written():
    db = _StoreDb({"A": {"serialUpper": "AQ-1"}}, deleted=["A"])
    counts = push_statuses({"A": ("AQ-1", {"ecg": "pass"})}, db)

    assert db.written == {}
    assert db.read == []
    assert counts["deleted"] == 1


def test_every_record_is_reached_across_batches(monkeypatch):
    monkeypatch.setattr(firebase_export, "BATCH_SIZE", 2)
    docs = {f"R{i}": {"serialUpper": f"S{i}"} for i in range(5)}
    found = {f"R{i}": (f"S{i}", {"overall": "pass"}) for i in range(5)}
    db = _StoreDb(docs)

    counts = push_statuses(found, db)

    assert counts["updated"] == 5
    assert len(db.commits) == 3
    assert set(db.written) == set(docs)


def test_statuses_only_refuses_without_yes_before_connecting(
    tmp_path, monkeypatch, capsys
):
    form = _record("G181-AGH002-0825", status="pass", status2="limited non")
    monkeypatch.setattr(firebase_export, "scan",
                        lambda root, only_year=None, limit=0: [form])
    monkeypatch.setattr(firebase_export, "deepen",
                        lambda forms, sample=0, progress=None: len(forms))
    monkeypatch.setattr(firebase_export, "assign_ids", lambda forms: (0, 0))

    def no_connection(project):
        raise AssertionError("connected without --yes")

    monkeypatch.setattr(firebase_export, "connect", no_connection)

    assert firebase_export.main([str(tmp_path), "--statuses-only"]) == 3
    out = capsys.readouterr().out
    assert "the `statuses` field only" in out
    assert "Refusing to write" in out


# ── finding the status box by its label ───────────────────────────────────────
#
# The balance map reads G30, where the app's own template puts the status.
# Older balance forms print "Status" at G25 and the answer at G27, so ~70% of
# the archive's balances read blank. The form's own label is what finds it.

from firebase_export import status_from_grid


def test_the_status_box_is_found_two_rows_under_its_label():
    grid = {"G25": "Status", "G27": "pass", "C30": "Comment:"}
    assert status_from_grid(grid) == "pass"


def test_a_label_merged_down_over_two_rows_is_stepped_over():
    # Calist resolves merges, so a label spanning G25:G26 reports at both.
    assert status_from_grid({"G25": "Status", "G26": "Status", "G27": "Fail"}) == "Fail"


def test_the_status_box_can_sit_to_the_right_of_its_label():
    grid = {"E30": "Status:", "F30": "Status:", "G30": "Calibrated"}
    assert status_from_grid(grid) == "Calibrated"


def test_a_neighbour_that_is_not_a_status_is_never_taken():
    assert status_from_grid({"G25": "Status", "G27": "-----"}) == ""
    assert status_from_grid({"G25": "Status", "H25": "Comment:", "G27": "Large"}) == ""


def test_a_label_to_the_right_does_not_hide_the_box_below():
    grid = {"G25": "Status", "J25": "Comment:", "G27": "limited non"}
    assert status_from_grid(grid) == "limited non"


def test_no_label_means_no_status():
    assert status_from_grid({"G27": "pass"}) == ""


def test_a_single_box_form_falls_back_to_its_labelled_box():
    form = _form(device_code="BP", tag="BP001", status="", status_by_label="Pass")
    assert build_statuses(form) == {"overall": "pass"}


def test_the_mapped_cell_wins_whenever_it_reads():
    form = _form(device_code="BP", tag="BP001", status="Fail", status_by_label="Pass")
    assert build_statuses(form) == {"overall": "fail"}


def test_a_two_box_device_never_uses_the_label():
    # The nearest label could be the other module's.
    form = _form(device_code="AGH", status="", status2="pass", status_by_label="fail")
    assert build_statuses(form) == {"nibp": "pass"}


def test_the_fallback_never_changes_passed():
    # `passed` is what a phone on the old build reads, and stays as it was.
    form = _form(device_code="BP", tag="BP001", status="", status_by_label="Pass")
    document = build_document(form, None)
    assert document["passed"] is False
    assert document["statuses"] == {"overall": "pass"}


# ── identity by the exact source file ─────────────────────────────────────────
#
# The original import read the wrong cell as the serial on some forms: 749
# records store none, 138 store "PASS", a ward name or a date. The serial check
# alone refused them all, although they are the same file. `sourcePath` — the
# exact file each record was built from — proves it; a file *name* would not,
# because renumbered siblings share one.


@pytest.mark.parametrize("raw, expected", [
    ("Passs", "pass"), ("Psss", "pass"), ("Pass.", "pass"),
    ("Caibrated", "calibrated"), ("limeted non", "limited non"),
])
def test_the_archive_s_common_misspellings_are_read(raw, expected):
    assert normalise_status(raw) == expected


@pytest.mark.parametrize("raw", ["p", "ci", "0", "0.0", "."])
def test_a_scrap_that_could_be_anything_stays_refused(raw):
    assert normalise_status(raw) is None


def test_a_record_with_no_serial_on_the_server_is_matched_by_its_file():
    db = _StoreDb({"A": {"serialUpper": "", "sourcePath": "D:/x/A.xlsx"}})
    counts = push_statuses({"A": ("TH001", {"overall": "pass"}, {"D:/x/A.xlsx"})}, db)

    assert db.written == {"A": {"statuses": {"overall": "pass"}}}
    assert counts["updated"] == 1


def test_a_misread_serial_on_the_server_does_not_hide_the_same_file():
    db = _StoreDb({"A": {"serialUpper": "PASS", "sourcePath": "D:/x/A.xlsx"}})
    counts = push_statuses({"A": ("17029", {"overall": "fail"}, {"D:/x/A.xlsx"})}, db)

    assert counts["updated"] == 1


def test_a_different_file_with_a_different_serial_is_still_refused():
    # A renumbered sibling: the id now belongs to another file of the same name.
    db = _StoreDb({"A": {"serialUpper": "243058", "sourcePath": "D:/y/A.xlsx"}})
    counts = push_statuses({"A": ("20946339", {"overall": "pass"}, {"D:/x/A.xlsx"})}, db)

    assert db.written == {}
    assert counts["mismatched"] == 1


def test_an_empty_serial_on_both_sides_is_not_proof_of_anything():
    db = _StoreDb({"A": {"serialUpper": "", "sourcePath": "D:/y/other.xlsx"}})
    counts = push_statuses({"A": ("", {"overall": "pass"}, {"D:/x/A.xlsx"})}, db)

    assert db.written == {}
    assert counts["mismatched"] == 1


# ── "Limited", and the monitor template's three boxes ─────────────────────────

from firebase_export import module_statuses_from_grid


def test_an_unqualified_limited_is_its_own_outcome():
    assert normalise_status("Limited") == "limited"
    assert normalise_status("LIMITED") == "limited"
    # By the owner's decision, recorded the same way rather than as calibrated.
    assert normalise_status("Limited Calibrated") == "limited"
    # The qualified ones keep their meaning.
    assert normalise_status("Limited Fail") == "limited fail"
    assert normalise_status("Limited Non") == "limited non"


def test_limited_ranks_above_limited_non_and_below_limited_fail():
    assert build_statuses(_form(device_code="GC", status="limited non",
                                status2="Limited")) == {"overall": "limited"}
    assert build_statuses(_form(device_code="GC", status="Limited",
                                status2="limited fail")) == {"overall": "limited fail"}


_ROW = {"D38": "Ecg Status:", "G38": "Spo2 Status:", "J38": "NIBP Status:"}


def test_the_three_boxes_are_read_under_their_labels():
    grid = {**_ROW, "D39": "Pass", "G39": "0.0", "J39": "FAIL"}
    assert module_statuses_from_grid(grid) == {"ecg": "Pass", "nibp": "FAIL"}


def test_a_module_the_device_does_not_have_is_skipped():
    # "----" under ECG is the template's own "no such module". It decides the
    # ECG box; the "pass" further down belongs to something else.
    grid = {**_ROW, "D39": "----", "D40": "pass", "G39": "Pass"}
    assert module_statuses_from_grid(grid) == {"spo2": "Pass"}


def test_one_module_never_takes_its_neighbour_s_answer():
    # Nothing under ECG; to its right lie the SpO2 label and then SpO2's answer.
    grid = {"D38": "Ecg Status:", "G38": "Spo2 Status:", "H38": "Pass"}
    assert module_statuses_from_grid(grid) == {"spo2": "Pass"}


def test_a_label_merged_across_two_cells_is_stepped_over():
    grid = {"D38": "Ecg Status:", "D39": "Ecg Status:", "D40": "limited"}
    assert module_statuses_from_grid(grid) == {"ecg": "limited"}


def test_a_form_not_on_the_template_has_no_module_boxes():
    assert module_statuses_from_grid({"G25": "Status", "G27": "pass"}) == {}


def test_a_spo2_device_on_the_monitor_template_keeps_each_box():
    form = _form(device_code="AH", tag="AH001", status="0.0",
                 module_statuses={"spo2": "Pass", "nibp": "fail"})
    assert build_statuses(form) == {"spo2": "pass", "nibp": "fail"}


def test_a_monitor_gains_its_spo2_box_and_a_box_its_map_missed():
    form = _form(device_code="AGH", status="0", status2="pass",
                 module_statuses={"ecg": "Pass", "spo2": "PASS"})
    assert build_statuses(form) == {"nibp": "pass", "ecg": "pass", "spo2": "pass"}


def test_on_a_monitor_the_mapped_box_wins_over_the_label():
    form = _form(device_code="AGH", status="fail", status2="pass",
                 module_statuses={"ecg": "pass"})
    assert build_statuses(form)["ecg"] == "fail"


def test_a_vital_signs_form_whose_map_lands_on_the_labels_reads_below_them():
    form = _form(device_code="VAH", status="Spo2 Status:", status2="NIBP Status:",
                 module_statuses={"spo2": "Pass", "nibp": "pass"})
    assert build_statuses(form) == {"spo2": "pass", "nibp": "pass"}
