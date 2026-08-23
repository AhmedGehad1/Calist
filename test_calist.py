"""Tests for the pure logic in calist.py. Run with: python -m pytest

These cover the parts with no I/O: filename parsing, value cleaning, ordering,
de-duplication and second-row generation.
"""

from datetime import datetime

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
