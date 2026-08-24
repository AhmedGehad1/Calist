"""Tests for the pure logic in calist.py. Run with: python -m pytest

These cover the parts with no I/O: filename parsing, value cleaning, ordering,
de-duplication, second-row generation and pre-flight classification, plus a
couple of end-to-end runs against workbooks built on the fly.
"""

import threading
from datetime import datetime

import pytest
from openpyxl import Workbook, load_workbook

import calist
from device_config import DEVICE_CONFIGS, form


# ── extract_device_code ───────────────────────────────────────────────────────

def test_code_after_first_hyphen():
    assert calist.extract_device_code("Clinic-AGH001.xlsx") == "AGH"


def test_code_without_hyphen_uses_whole_stem():
    assert calist.extract_device_code("VNT023.xlsx") == "VNT"


def test_code_is_uppercased():
    assert calist.extract_device_code("site-agh001.xlsx") == "AGH"


def test_code_is_none_when_no_letters():
    assert calist.extract_device_code("123.xlsx") is None


# ── clean ─────────────────────────────────────────────────────────────────────

def test_clean_none_is_blank():
    assert calist.clean(None) == ""


def test_clean_strips_text():
    assert calist.clean("  ABC  ") == "ABC"


def test_clean_stringifies_numbers_as_is():
    """Deliberate: a numeric cell keeps Excel's float form."""
    assert calist.clean(123456.0) == "123456.0"
    assert calist.clean(1.5) == "1.5"


def test_clean_stringifies_dates_as_is():
    """Deliberate: a real date cell keeps datetime's default repr."""
    assert calist.clean(datetime(2024, 1, 15)) == "2024-01-15 00:00:00"


# ── natural_key ───────────────────────────────────────────────────────────────

def test_numbers_sort_numerically():
    codes = ["AGH10", "AGH9", "AGH1"]
    assert sorted(codes, key=calist.natural_key) == ["AGH1", "AGH9", "AGH10"]


def test_natural_key_never_compares_text_with_numbers():
    # Mixed shapes must not raise TypeError.
    sorted(["AGH1", "AGHX", "1AGH", "AGH"], key=calist.natural_key)


# ── second row ────────────────────────────────────────────────────────────────

def test_second_row_takes_status2_and_rewrites_code():
    parent = {"Code": "AGH001", "Status": "OK", "Status2": "Faulty", "_row_order": 0}
    child = calist.build_second_row(
        parent, {"device_name": "NIBP", "code_replace": ("AGH", "AGCB")})
    assert child["Device"] == "NIBP"
    assert child["Status"] == "Faulty"
    assert child["Code"] == "AGCB001"
    assert child["_row_order"] == 1


def test_second_row_replaces_code_token_only_once():
    parent = {"Code": "AGH-AGH001", "Status": "", "Status2": "", "_row_order": 0}
    child = calist.build_second_row(
        parent, {"device_name": "NIBP", "code_replace": ("AGH", "AGCB")})
    assert child["Code"] == "AGCB-AGH001"


def test_second_row_does_not_mutate_parent():
    parent = {"Code": "AGH001", "Status": "OK", "Status2": "Faulty", "_row_order": 0}
    calist.build_second_row(
        parent, {"device_name": "NIBP", "code_replace": ("AGH", "AGCB")})
    assert parent["Status"] == "OK" and parent["Code"] == "AGH001"


# ── sorting ───────────────────────────────────────────────────────────────────

def test_sub_module_row_follows_its_parent():
    records = [
        {"_group": "AGH002", "_row_order": 0, "Device": "Patient Monitor"},
        {"_group": "AGH001", "_row_order": 1, "Device": "NIBP"},
        {"_group": "AGH001", "_row_order": 0, "Device": "Patient Monitor"},
    ]
    ordered = calist.sort_records(records)
    assert [(r["_group"], r["_row_order"]) for r in ordered] == [
        ("AGH001", 0), ("AGH001", 1), ("AGH002", 0),
    ]


# ── deduplicate_records ───────────────────────────────────────────────────────

def _rec(device, serial):
    return {"Device": device, "S.N": serial}


def test_blank_serials_are_always_kept():
    records = [_rec("ECG", ""), _rec("ECG", ""), _rec("ECG", "")]
    assert len(calist.deduplicate_records(records)) == 3


def test_plain_duplicate_is_dropped():
    records = [_rec("ECG", "SN1"), _rec("ECG", "SN1")]
    assert len(calist.deduplicate_records(records)) == 1


def test_monitor_and_nibp_may_share_a_serial():
    records = [_rec("Patient Monitor", "SN1"), _rec("NIBP", "SN1")]
    assert len(calist.deduplicate_records(records)) == 2


def test_third_record_on_a_shared_serial_is_dropped():
    records = [_rec("Patient Monitor", "SN1"), _rec("NIBP", "SN1"), _rec("ECG", "SN1")]
    kept = calist.deduplicate_records(records)
    assert [r["Device"] for r in kept] == ["Patient Monitor", "NIBP"]


def test_unpaired_devices_may_not_share_a_serial():
    records = [_rec("ECG", "SN1"), _rec("Ultrasound", "SN1")]
    assert len(calist.deduplicate_records(records)) == 1


def test_vital_sign_modules_may_share_a_serial():
    records = [_rec("Vital Sign (SPO2 Module)", "SN1"),
               _rec("Vital Sign (NIBP Module)", "SN1")]
    assert len(calist.deduplicate_records(records)) == 2


# ── config integrity ──────────────────────────────────────────────────────────

def test_form_produces_the_documented_layout():
    assert form(18, "H32") == {
        "Manufacturer": "E20", "Model": "E18", "S.N": "K18",
        "Location": "K20", "Date": "E16", "Status": "H32",
    }


def test_every_config_has_the_required_fields():
    required = {"Manufacturer", "Model", "S.N", "Location", "Date", "Status"}
    for code, config in DEVICE_CONFIGS.items():
        assert "device_name" in config, code
        assert required <= set(config["cells"]), code


def test_second_row_names_are_covered_by_the_dedup_exemptions():
    """A device rename must not silently break the shared-serial rule."""
    for code, config in DEVICE_CONFIGS.items():
        if "second_row" not in config:
            continue
        pair = frozenset({config["device_name"],
                          config["second_row"]["device_name"]})
        assert pair in calist.ALLOWED_SHARED_SN_PAIRS, code


def test_configs_with_second_row_define_status2():
    for code, config in DEVICE_CONFIGS.items():
        if "second_row" in config:
            assert "Status2" in config["cells"], code


# ── classify_file (pre-flight, no workbook I/O) ───────────────────────────────

def test_classify_recognises_a_known_device():
    outcome = calist.classify_file("Clinic-AC003.xlsx")
    assert outcome.status == calist.READY
    assert outcome.device_code == "AC"
    assert outcome.device_name == "Defibrillator"
    assert outcome.rows == 1
    assert not outcome.is_problem


def test_classify_reports_two_rows_for_module_devices():
    outcome = calist.classify_file("Clinic-AGH001.xlsx")
    assert outcome.rows == 2
    assert "NIBP" in outcome.device_name


def test_classify_flags_an_unknown_code():
    outcome = calist.classify_file("Clinic-ZZZ999.xlsx")
    assert outcome.status == calist.UNKNOWN_CODE
    assert outcome.is_problem
    assert "ZZZ" in outcome.detail


def test_classify_flags_an_unsupported_extension():
    outcome = calist.classify_file("notes.txt")
    assert outcome.status == calist.UNSUPPORTED
    assert outcome.is_problem


def test_classify_does_not_open_the_file():
    """It must work on a path that does not exist — that is the whole point."""
    outcome = calist.classify_file(r"X:\nowhere\Clinic-AF001.xlsx")
    assert outcome.status == calist.READY
    assert outcome.device_name == "ECG"


# ── end-to-end ────────────────────────────────────────────────────────────────

def _form(path, cells):
    wb = Workbook()
    ws = wb.active
    for ref, value in cells.items():
        ws[ref] = value
    wb.save(path)
    return str(path)


@pytest.fixture
def workspace(tmp_path):
    """A source folder, a template, and three forms (one a two-row device)."""
    src = tmp_path / "src"
    src.mkdir()
    tpl_dir = tmp_path / "tpl"
    tpl_dir.mkdir()

    template = tpl_dir / "Device List.xlsx"
    wb = Workbook()
    ws = wb.active
    for col, name in enumerate(["No."] + calist.FIELDS, start=1):
        ws.cell(row=3, column=col, value=name)
    wb.save(template)

    files = [
        _form(src / "Clinic-AC001.xlsx",
              {"E17": "Zoll", "E15": "R-Series", "K15": "SN-1",
               "K17": "ER", "E13": "01/01/2024", "G24": "Working"}),
        _form(src / "Clinic-AGH002.xlsx",
              {"E20": "Philips", "E18": "MX450", "K18": "SN-2", "K20": "ICU",
               "E16": "02/01/2024", "D39": "Working", "J39": "Faulty"}),
        _form(src / "Clinic-ZZZ003.xlsx", {"A1": "unknown"}),
    ]
    return files, str(template)


def test_run_reports_every_file_and_writes_the_register(workspace):
    files, template = workspace
    result = calist.process_files(files, template)

    assert result.succeeded
    assert result.output_path.exists()
    assert len(result.outcomes) == 3
    assert result.files_read == 2                 # the ZZZ file is skipped
    assert result.second_rows_added == 1          # AGH generates its NIBP row
    assert result.rows_written == 3               # 2 devices + 1 module row
    assert [o.status for o in result.problems] == [calist.UNKNOWN_CODE]


def test_progress_hook_fires_once_per_file_in_order(workspace):
    files, template = workspace
    seen = []
    calist.process_files(files, template,
                         on_file=lambda o, i, t: seen.append((i, t, o.status)))

    assert [i for i, _, _ in seen] == [1, 2, 3]
    assert {t for _, t, _ in seen} == {3}


def test_cancelling_before_the_run_writes_nothing(workspace):
    files, template = workspace
    cancel = threading.Event()
    cancel.set()

    result = calist.process_files(files, template, cancel=cancel)

    assert result.cancelled
    assert result.output_path is None
    assert not result.succeeded
    assert all(o.status == calist.CANCELLED for o in result.outcomes)


def test_cancelling_partway_stops_early(workspace):
    files, template = workspace
    cancel = threading.Event()

    # Trip the flag as soon as the first file has been handled.
    def on_file(outcome, index, total):
        if index == 1:
            cancel.set()

    result = calist.process_files(files, template, on_file=on_file, cancel=cancel)

    assert result.cancelled
    assert result.output_path is None
    assert result.files_read == 1
    assert any(o.status == calist.CANCELLED for o in result.outcomes)


def test_refuses_to_overwrite_the_template(tmp_path):
    """The template living beside the sources must not be clobbered."""
    src = tmp_path / "src"
    src.mkdir()
    template = src / calist.OUTPUT_NAME       # same folder, same name
    wb = Workbook()
    wb.save(template)
    form_path = _form(src / "Clinic-AC001.xlsx", {"E15": "R-Series"})

    result = calist.process_files([form_path], str(template))

    assert not result.succeeded
    assert "template" in (result.error or "").lower()
    assert template.stat().st_size > 0


def test_written_rows_land_where_the_template_expects_them(workspace):
    files, template = workspace
    result = calist.process_files(files, template)

    wb = load_workbook(result.output_path)
    ws = wb.active
    first = [ws.cell(row=calist.TEMPLATE_START_ROW, column=c).value
             for c in range(1, len(calist.FIELDS) + 2)]
    wb.close()

    assert first[0] == 1                              # index in column A
    assert first[1] == "Defibrillator"                # Device in column B
    assert "SN-1" in first[4]                         # S.N in column E
