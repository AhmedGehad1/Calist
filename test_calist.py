"""Tests for the pure logic in calist.py. Run with: python -m pytest

These cover the parts with no I/O: filename parsing, value cleaning, ordering,
de-duplication, second-row generation and pre-flight classification, plus a
couple of end-to-end runs against workbooks built on the fly.
"""

import logging
import os
import re
import threading
import zipfile
from pathlib import Path
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


# The duplicate warning has to name the two files, because that is what the
# user opens to resolve it — the device type does not identify which form to
# look at when a round holds a dozen of the same model.

def _sourced(device, serial, source):
    return {"Device": device, "S.N": serial, "_source": source}


def test_duplicate_warning_names_both_files(caplog):
    records = [_sourced("ECG", "SN1", "D23-AGH001-0225.xlsx"),
               _sourced("ECG", "SN1", "D23-AGH007-0225.xlsx")]
    with caplog.at_level(logging.WARNING, logger="aggregator"):
        calist.deduplicate_records(records)

    message = caplog.text
    assert "D23-AGH001-0225.xlsx" in message
    assert "D23-AGH007-0225.xlsx" in message
    assert "SN1" in message


def test_duplicate_warning_falls_back_to_the_code_without_a_source():
    # Records built by older callers carry no _source; Code still names the file.
    assert calist.source_name({"Code": "D23-AGH001-0225"}) == "D23-AGH001-0225"
    assert calist.source_name({}) == "?"


def test_a_sub_module_row_reports_the_file_it_came_from():
    """build_second_row rewrites Code, so only _source still names the file."""
    parent = {"Device": "Patient Monitor", "S.N": "SN1", "Code": "Ward-AGH001",
              "_source": "Ward-AGH001.xlsx", "Status2": "Working"}
    extra = calist.build_second_row(
        parent, {"device_name": "NIBP", "code_replace": ("AGH", "NIB")})

    assert extra["Code"] != parent["Code"]          # the token was rewritten
    assert calist.source_name(extra) == "Ward-AGH001.xlsx"


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


# ── filename repair ───────────────────────────────────────────────────────────
#
# By the owner's decision the files are never renamed; the name Calist writes
# and exports is repaired. Every example is a real archive filename.

@pytest.mark.parametrize("stem, name", [
    ("G302-AGH001-0425", "G302-AGH001-0425"),         # already right: untouched
    (".G414-CA002-0426", "G414-CA002-0426"),          # a leading dot
    ("'D33-AD005-0623", "D33-AD005-0623"),            # a leading apostrophe
    ("\u206fD38-AGH178-0224", "D38-AGH178-0224"),     # an invisible mark
    ("Copy of B14-BB036-0126", "B14-BB036-0126"),
    ("Final G317-BM001-1224", "G317-BM001-1224"),
    ("G119-CE007 -0525", "G119-CE007-0525"),          # spaces around a dash
    ("D20-AGH039- 0824", "D20-AGH039-0824"),
    ("H05-AC016-0924  ", "H05-AC016-0924"),           # spaces before .xlsx
    ("k116-AX001-0123", "K116-AX001-0123"),           # lower case
    ("JO8-AGH001-1025", "J08-AGH001-1025"),           # a letter O for a zero
    ("g302-bb001-0326-pending", "G302-BB001-0326"),
    ("F29-AI055-0625 (303) pending downstream error", "F29-AI055-0625"),
    ("D38-AGH090-0225 (2)", "D38-AGH090-0225"),
    ("K86-CE013-0323 - Copy", "K86-CE013-0323"),
    ("J11-AGH027-1224..........", "J11-AGH027-1224"),
    ("F30-FF001-0724applied parts safety report", "F30-FF001-0724"),
])
def test_a_name_is_written_with_its_slips_undone(stem, name):
    assert calist.repair_name(stem, 2026).name == name


@pytest.mark.parametrize("stem", [
    "GE      TEC850     Main OR   SQAX00871",   # named by brand
    "5071938426",                               # named by serial
    "\u200f\u0645\u0633\u062a\u0646\u062f Ahmed",  # Arabic text in front is not a slip
    "H59-BZ00F-1025",                           # a device number past repair
    "G302-AGH001-june",                         # no date to cut after
    "H58-BM010",                                # no date at all
    "Clinic-AGH001",
])
def test_a_name_that_is_not_the_house_shape_is_left_exactly_as_it_is(stem):
    assert calist.repair_name(stem, 2026).name == stem


def test_o01_is_a_customer_code_not_a_slip():
    """The O-for-zero repair is for a letter O *after* the site's letter."""
    assert calist.repair_name("O01-CD043-0224", 2026).name == "O01-CD043-0224"
    assert calist.repair_name("BO3-CF005-0923", 2026).name == "B03-CF005-0923"


@pytest.mark.parametrize("stem", [
    "G302-BB001-0329",      # 2029 — the owner's own example
    "H39-AG025-0233",
    "H96-AGH022-1525",      # month 15
    "G161-EP001-2023",      # month 20
    "G307-CA001-0000",
    "G114-BP001-00324",     # five digits
    "B04-ED003-923",        # three
    "D38-AI008-02240000000000",
])
def test_a_date_that_cannot_be_one_waits_for_the_form(stem):
    repair = calist.repair_name(stem, 2026)
    assert repair.date_source == calist.DATE_CHECK
    assert repair.name == stem      # nothing is guessed before the form is read


def test_the_latest_year_is_the_line():
    assert calist.repair_name("G302-BB001-0326", 2026).date_source == calist.DATE_NAME
    assert calist.repair_name("G302-BB001-0327", 2026).date_source == calist.DATE_CHECK
    assert calist.repair_name("G302-BB001-0327", 2027).date_source == calist.DATE_NAME


def test_the_forms_own_date_settles_an_impossible_one():
    """"0329" for a device calibrated 11-03-2026 is -0326."""
    repair = calist.settle_name_date(calist.repair_name("G302-BB001-0329", 2026),
                                     "11-03-2026")
    assert repair.name == "G302-BB001-0326"
    assert repair.date_source == calist.DATE_FORM
    assert "'0329'" in repair.notes[0] and "11-03-2026" in repair.notes[0]


def test_the_form_is_trusted_whatever_year_it_gives():
    """By the owner's decision: the form's date, whenever it reads as one."""
    repair = calist.settle_name_date(calist.repair_name("H12-BB006-04720", 2023),
                                     "15-07-2020")
    assert repair.name == "H12-BB006-0720"


def test_a_form_with_no_date_leaves_the_name_as_typed_and_says_so():
    repair = calist.settle_name_date(calist.repair_name("K28-FC095-012300", 2026),
                                     "")
    assert repair.name == "K28-FC095-012300"
    assert repair.date_source == calist.DATE_UNSETTLED
    assert "no date to take" in repair.notes[0]


def test_a_real_date_in_the_name_is_never_replaced_by_the_forms():
    repair = calist.settle_name_date(calist.repair_name("G302-BB001-0326", 2026),
                                     "11-09-2025")
    assert repair.name == "G302-BB001-0326"


@pytest.mark.parametrize("text, mmyy", [
    ("11-03-2026", "0326"),
    ("11/3/2026", "0326"),
    ("2026-03-11 00:00:00", "0326"),     # a real date cell in an .xlsx
    ("45306.0", "0124"),                 # xlrd's number for 15-01-2024
    ("'10-02-2024", "0224"),             # forced to text
    ("25-20-2023", None),                # month 20
    ("Pass", None),
    ("", None),
])
def test_a_forms_date_as_mmyy(text, mmyy):
    assert calist.month_year(text) == mmyy


def test_only_repairs_that_change_what_the_name_says_are_marked():
    """A space, a dot or lower case are not worth a person's time."""
    for stem in ("G119-CE007 -0525", ".G414-CA002-0426", "k116-AX001-0123",
                 "\u206fD38-AGH178-0224", "J11-AGH027-1224.........."):
        assert calist.repair_name(stem, 2026).notes == [], stem
    for stem in ("Copy of B14-BB036-0126", "JO8-AGH001-1025",
                 "F29-AI041-0625 Pending", "D38-AGH090-0225 (2)"):
        assert calist.repair_name(stem, 2026).notes, stem


def test_a_copy_is_a_name_that_lost_text():
    assert calist.repair_name("D38-AGH090-0225 (2)", 2026).is_copy
    assert calist.repair_name("Copy of B14-BB036-0126", 2026).is_copy
    assert calist.repair_name("D02-BZ118-1225 --", 2026).is_copy
    assert not calist.repair_name("H05-AC016-0924 ", 2026).is_copy
    assert not calist.repair_name(".G414-CA002-0426", 2026).is_copy


# ── the "real device forms only" switch ───────────────────────────────────────

@pytest.mark.parametrize("stem, why", [
    ("GE      TEC850     Main OR   SQAX00871", "customer code"),
    ("5071938426", "customer code"),
    ("000-AGH000-0000", "customer code"),
    ("AG035-SMB176400052HA", "customer code"),   # a device code, not a site
    ("Clinic-AGH001", "customer code"),
    ("CD001", "customer code"),
    ("D38-BZ000-0126-SINO", "template"),
    ("D26-AGH000-0225", "template"),
    ("G132-BP0", "template"),
    ("B14-CF00-1223", "template"),
])
def test_what_the_switch_leaves_out_by_name(stem, why):
    assert why in calist.left_out_because(calist.repair_name(stem, 2026))


@pytest.mark.parametrize("stem", [
    "G302-AGH001-0425", ".G414-CA002-0426", "JO8-AGH001-1025",
    "H58-BM010",            # no date, but a customer's device
    "H59-BZ00F-1025",       # device number past repair, customer code intact
    "D12-BZ010-0224",
])
def test_what_the_switch_lets_in(stem):
    assert calist.left_out_because(calist.repair_name(stem, 2026)) == ""


def test_the_switch_is_off_by_default():
    """An old-style name stays acceptable unless the caller asks otherwise."""
    assert calist.classify_file("Clinic-AGH001.xlsx").status == calist.READY


def test_the_switch_leaves_a_non_form_out_and_it_is_not_a_problem():
    outcome = calist.classify_file("Clinic-AGH001.xlsx", real_forms_only=True)
    assert outcome.status == calist.LEFT_OUT
    assert not outcome.is_problem
    assert "customer code" in outcome.detail


def test_the_switch_passes_a_good_name_through_to_the_device_lookup():
    outcome = calist.classify_file("G302-AGH001-0425.xlsx", real_forms_only=True)
    assert outcome.status == calist.READY
    assert outcome.device_code == "AGH"
    assert outcome.rows == 2


def test_the_switch_is_asked_before_the_device_code():
    """The user asked for device forms; a non-form is the finding, not its code."""
    outcome = calist.classify_file("nonsense-ZZZ999.xlsx", real_forms_only=True)
    assert outcome.status == calist.LEFT_OUT


def test_pre_flight_carries_the_repaired_name():
    outcome = calist.classify_file("C:/nowhere/.G414-CA002-0426.xlsm")
    assert outcome.code == "G414-CA002-0426"


def test_house_format_still_yields_the_right_device_code():
    assert calist.extract_device_code("G302-AGH001-0425.xlsx") == "AGH"
    assert calist.extract_device_code("G302-VAH012-1226.xlsx") == "VAH"


# ── merged cells ──────────────────────────────────────────────────────────────
#
# The forms draw each answer as a box spanning two columns. A merged range
# stores its value only in the top-left cell, so a map naming the second column
# read blank — which is how the Ultrasound serial went missing when that form
# was re-laid-out onto the E/K columns.

def _merged_form(path, cells, merges):
    wb = Workbook()
    ws = wb.active
    for ref, value in cells.items():
        ws[ref] = value
    for rng in merges:
        ws.merge_cells(rng)
    wb.save(path)
    return str(path)


def test_a_merged_cell_reads_through_to_its_anchor(tmp_path):
    path = _merged_form(tmp_path / "f.xlsx", {"K17": "6061439WX0"}, ["K17:L17"])
    assert calist.read_record(path, {"S.N": "L17"}) == {"S.N": "6061439WX0"}


def test_naming_the_anchor_itself_still_works(tmp_path):
    path = _merged_form(tmp_path / "f.xlsx", {"K17": "6061439WX0"}, ["K17:L17"])
    assert calist.read_record(path, {"S.N": "K17"}) == {"S.N": "6061439WX0"}


def test_an_ordinary_empty_cell_stays_empty(tmp_path):
    """Only cells actually inside a merge resolve; blank means blank."""
    path = _merged_form(tmp_path / "f.xlsx", {"A1": "header"}, ["A1:Z1"])
    assert calist.read_record(path, {"Status": "H30"}) == {"Status": ""}


def test_the_ultrasound_probe_serial_survives_a_merged_layout(tmp_path):
    """The whole BB form, drawn as two-column boxes on the E/K columns."""
    path = _merged_form(
        tmp_path / "G302-BB001-0526.xlsx",
        {"E15": "10-05-2026", "E17": "Versana Essential", "E19": "GE",
         "K17": "6061439WX0", "K19": "Clinics", "K21": "982693WX4",
         "H30": "Working"},
        [f"{c}{r}:{chr(ord(c) + 1)}{r}"
         for r in (15, 17, 19, 21) for c in ("E", "K")])

    record = calist.read_record(path, DEVICE_CONFIGS["BB"]["cells"])

    assert record["S.N"] == "6061439WX0"
    assert record["S.N2"] == "982693WX4"
    assert record["Model"] == "Versana Essential"
    assert record["Manufacturer"] == "GE"
    assert record["Location"] == "Clinics"


def test_the_two_ultrasound_serials_end_up_on_two_lines(tmp_path):
    """The register cell reads '<device serial>\\n(<probe serial>)'."""
    src = tmp_path / "src"
    src.mkdir()
    template = tmp_path / "Device List.xlsx"
    wb = Workbook()
    ws = wb.active
    for col, name in enumerate(["No."] + calist.FIELDS, start=1):
        ws.cell(row=3, column=col, value=name)
    wb.save(template)

    form_path = _merged_form(
        src / "G302-BB001-0526.xlsx",
        {"E15": "10-05-2026", "E17": "Versana Essential", "E19": "GE",
         "K17": "6061439WX0", "K19": "Clinics", "K21": "982693WX4",
         "H30": "Working"},
        [f"{c}{r}:{chr(ord(c) + 1)}{r}"
         for r in (15, 17, 19, 21) for c in ("E", "K")])

    result = calist.process_files([form_path], str(template))

    assert result.succeeded
    written = load_workbook(result.output_path).active
    serial_col = calist.TEMPLATE_START_COL + calist.FIELDS.index("S.N")
    assert written.cell(row=calist.TEMPLATE_START_ROW,
                        column=serial_col).value == "6061439WX0\n(982693WX4)"


# ── which sheet gets read ─────────────────────────────────────────────────────
#
# "Always the first tab" was right for every form but one: the X-ray workbook
# opens on an empty "Waveform Dialog" stub left by its macros, so every mapped
# cell read blank on every X-ray.

def _multi_sheet_form(path, sheets):
    """sheets: list of (title, {ref: value}) in file order."""
    wb = Workbook()
    wb.remove(wb.active)
    for title, cells in sheets:
        ws = wb.create_sheet(title)
        for ref, value in cells.items():
            ws[ref] = value
    wb.save(path)
    return str(path)


def test_an_empty_leading_sheet_is_skipped(tmp_path):
    path = _multi_sheet_form(tmp_path / "f.xlsx", [
        ("Waveform Dialog", {}),
        ("Data entry", {"K18": "XR-77341"}),
    ])
    assert calist.read_record(path, {"S.N": "K18"}) == {"S.N": "XR-77341"}


def test_a_populated_first_sheet_still_wins(tmp_path):
    """Never skip past a sheet that holds data, even if a later one does too."""
    path = _multi_sheet_form(tmp_path / "f.xlsx", [
        ("Device data", {"K18": "the right one"}),
        ("Data entry", {"K18": "the wrong one"}),
    ])
    assert calist.read_record(path, {"S.N": "K18"}) == {"S.N": "the right one"}


def test_a_workbook_of_empty_sheets_does_not_crash(tmp_path):
    path = _multi_sheet_form(tmp_path / "f.xlsx", [("one", {}), ("two", {})])
    assert calist.read_record(path, {"S.N": "K18"}) == {"S.N": ""}


def test_the_xray_reads_through_a_stub_tab_to_its_form(tmp_path):
    """The real BF shape: empty macro tab first, merged boxes on Data entry."""
    wb = Workbook()
    wb.remove(wb.active)
    wb.create_sheet("Waveform Dialog")
    ws = wb.create_sheet("Data entry")
    for ref, value in {"E16": "10-05-2026", "E18": "Multix Impact",
                       "E20": "Siemens", "K18": "XR-77341",
                       "K20": "TUBE-99812", "K22": "Radiology",
                       "J27": "Pass"}.items():
        ws[ref] = value
    for row in (16, 18, 20, 22):
        ws.merge_cells(f"E{row}:F{row}")
        ws.merge_cells(f"K{row}:L{row}")
    path = tmp_path / "G302-BF001-0526.xlsx"
    wb.save(path)

    record = calist.read_record(str(path), DEVICE_CONFIGS["BF"]["cells"])

    assert record["S.N"] == "XR-77341"
    assert record["S.N2"] == "TUBE-99812"
    assert record["Model"] == "Multix Impact"
    assert record["Manufacturer"] == "Siemens"
    assert record["Location"] == "Radiology"
    assert record["Status"] == "Pass"


# ── reading the package directly ──────────────────────────────────────────────
#
# The reader goes at the sheet XML rather than through load_workbook, which
# took ~500ms a form because it parses every tab and the whole styles table to
# reach seven cells. These pin the parts of that which a fixture built with
# openpyxl cannot reach on its own.

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"


def _handmade_xlsx(path, sheets, shared=None, styles=None):
    """Write an .xlsx entry by entry, without openpyxl.

    Two shapes matter here and openpyxl cannot produce either. It writes every
    string as an *inline* string, so a fixture built with it never exercises
    the shared-string table — which is what Excel itself writes and what every
    real form uses. And it cannot emit the ``<sheet r:id="">`` entries an older
    macro-enabled workbook carries for its VBA modules.

    ``sheets`` is a list of ``(name, rel_id, kind, body_xml)``. A blank rel_id
    writes the broken entry, with no part behind it.
    """
    entries, rels = [], []
    for index, (name, rel_id, kind, body) in enumerate(sheets, start=1):
        if not rel_id:
            continue
        part = f"{kind}s/sheet{index}.xml"
        rels.append(f'<Relationship Id="{rel_id}" Target="{part}" '
                    f'Type="{_REL}/{kind}"/>')
        entries.append((f"xl/{part}",
                        f'<?xml version="1.0"?><worksheet xmlns="{_NS}">'
                        f'{body}</worksheet>'))

    listed = "".join(
        f'<sheet name="{name}" sheetId="{i}" r:id="{rel_id}"/>'
        for i, (name, rel_id, _, _) in enumerate(sheets, start=1))

    entries.append(("xl/workbook.xml",
                    f'<?xml version="1.0"?><workbook xmlns="{_NS}" '
                    f'xmlns:r="{_REL}"><sheets>{listed}</sheets></workbook>'))
    entries.append(("xl/_rels/workbook.xml.rels",
                    f'<?xml version="1.0"?><Relationships xmlns="{_PKG}">'
                    f'{"".join(rels)}</Relationships>'))
    if shared is not None:
        items = "".join(f"<si><t>{text}</t></si>" for text in shared)
        entries.append(("xl/sharedStrings.xml",
                        f'<?xml version="1.0"?><sst xmlns="{_NS}" '
                        f'count="{len(shared)}">{items}</sst>'))
    if styles is not None:
        applied = "".join(f'<xf numFmtId="{n}"/>' for n in styles)
        entries.append(("xl/styles.xml",
                        f'<?xml version="1.0"?><styleSheet xmlns="{_NS}">'
                        f'<cellXfs count="{len(styles)}">{applied}</cellXfs>'
                        f'<dxfs count="0"/></styleSheet>'))

    with zipfile.ZipFile(path, "w") as archive:
        for name, text in entries:
            archive.writestr(name, text)
    return str(path)


def _rows(*cells):
    return f'<row r="1">{"".join(cells)}</row>'


def test_a_shared_string_is_looked_up(tmp_path):
    """What Excel actually writes: the text lives in a separate table."""
    path = _handmade_xlsx(
        tmp_path / "f.xlsx",
        [("Device data", "rId1", "worksheet",
          f'<sheetData><row r="18"><c r="K18" t="s"><v>1</v></c></row>'
          f'</sheetData>')],
        shared=["not this one", "6061439WX0"])
    assert calist.read_record(path, {"S.N": "K18"}) == {"S.N": "6061439WX0"}


def test_a_self_closing_cell_does_not_swallow_the_next_one(tmp_path):
    """A styled-but-empty cell must read blank, not steal its neighbour.

    ``<c r="E18" s="168"/>`` is how Excel writes a formatted empty cell, and
    matching it loosely runs straight past the "/" and on to the *next* cell's
    "</c>" — so E18 silently returned F18's value.
    """
    path = _handmade_xlsx(
        tmp_path / "f.xlsx",
        [("Device data", "rId1", "worksheet",
          '<sheetData><row r="18"><c r="E18" s="168"/>'
          '<c r="F18" t="s"><v>0</v></c></row></sheetData>')],
        shared=["the neighbour"])
    record = calist.read_record(path, {"Model": "E18", "S.N": "F18"})
    assert record == {"Model": "", "S.N": "the neighbour"}


def test_a_vba_module_entry_does_not_shift_the_sheet_order(tmp_path):
    """Older .xlsm files list their macro modules as sheets with no r:id."""
    path = _handmade_xlsx(tmp_path / "f.xlsm", [
        ("MainModule", "", None, ""),
        ("Device data", "rId1", "worksheet",
         '<sheetData><row r="18"><c r="K18" t="str"><v>XR-77341</v></c>'
         '</row></sheetData>'),
    ])
    assert calist.read_record(path, {"S.N": "K18"}) == {"S.N": "XR-77341"}


def test_a_leading_chartsheet_is_not_counted_as_a_sheet(tmp_path):
    """Chartsheets are not worksheets; openpyxl keeps them apart and so do we.

    Asserted on the sheet list rather than on a value, because a chartsheet
    part holds no cells and would be skipped as empty either way — the list is
    where dropping the rule would actually show, and it is what --inspect
    prints.
    """
    path = _handmade_xlsx(tmp_path / "f.xlsx", [
        ("Trend", "rId1", "chartsheet", ""),
        ("Device data", "rId2", "worksheet",
         '<sheetData><row r="18"><c r="K18" t="str"><v>XR-77341</v></c>'
         '</row></sheetData>'),
    ])
    source = calist._open_workbook(path)
    try:
        assert source.sheet_names == ["Device data"]
        assert source.sheet_name == "Device data"
    finally:
        source.close()
    assert calist.read_record(path, {"S.N": "K18"}) == {"S.N": "XR-77341"}


def test_a_dialogsheet_is_still_a_candidate_sheet(tmp_path):
    """The X-ray workbook opens on one, and it must be skipped for emptiness
    rather than by type — openpyxl counts dialogsheets among its worksheets."""
    path = _handmade_xlsx(tmp_path / "f.xlsm", [
        ("Waveform Dialog", "rId1", "dialogsheet",
         '<sheetData><row r="1"><c r="B1" t="str"><v>stub</v></c></row>'
         '</sheetData>'),
        ("Data entry", "rId2", "worksheet",
         '<sheetData><row r="18"><c r="K18" t="str"><v>never reached</v></c>'
         '</row></sheetData>'),
    ])
    assert calist.read_record(path, {"S.N": "B1"}) == {"S.N": "stub"}


def test_a_date_formatted_number_becomes_a_date(tmp_path):
    """The styles table is only opened for a numeric cell a map asks for, so
    this pins that the lazy read still finds the number format."""
    path = _handmade_xlsx(
        tmp_path / "f.xlsx",
        [("Device data", "rId1", "worksheet",
          '<sheetData><row r="16">'
          '<c r="E16" s="1"><v>45306</v></c>'
          '<c r="F16" s="0"><v>45306</v></c>'
          '</row></sheetData>')],
        styles=[0, 14])          # 14 is the built-in short date format
    record = calist.read_record(path, {"Date": "E16", "S.N": "F16"})
    assert record == {"Date": "2024-01-15 00:00:00", "S.N": "45306"}


def test_escaped_text_is_decoded(tmp_path):
    path = _handmade_xlsx(
        tmp_path / "f.xlsx",
        [("Device data", "rId1", "worksheet",
          '<sheetData><row r="18"><c r="K18" t="str">'
          '<v>Smith &amp; Sons &lt;GmbH&gt; #40</v></c></row></sheetData>')])
    assert calist.read_record(path, {"S.N": "K18"}) == {
        "S.N": "Smith & Sons <GmbH> #40"}


def test_an_unreadable_file_raises_rather_than_reading_blank(tmp_path):
    """A silent blank field is the failure mode --inspect exists to hunt, so
    anything the reader cannot make sense of has to be loud."""
    path = tmp_path / "G302-BB001-0526.xlsx"
    path.write_bytes(b"this is not a zip")
    with pytest.raises(Exception):
        calist.read_record(str(path), {"S.N": "K18"})



# ── finding forms in a folder ─────────────────────────────────────────────────
#
# The intake used to walk with rglob("*") and stat every entry, on the UI
# thread. It also picked up the "~$" lock files Excel leaves beside any open
# workbook, which are not workbooks at all — so a folder someone was working in
# produced a row of failures for files nobody chose.

def _tree(root, names):
    for name in names:
        path = root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"x")
    return root


def test_lock_files_are_not_forms():
    assert not calist.is_source_file("~$G302-AGH001-0425.xlsx")
    assert calist.is_source_file("G302-AGH001-0425.xlsx")


def test_a_previous_run_output_is_not_a_form():
    """The register lands among the sources, so selecting the folder twice
    would otherwise feed the last run's output back in as an input."""
    assert not calist.is_source_file(calist.OUTPUT_NAME)
    assert not calist.is_source_file(calist.OUTPUT_NAME.upper())


def test_only_excel_files_are_found(tmp_path):
    _tree(tmp_path, ["a.xlsx", "b.xls", "c.xlsm", "notes.txt", "sheet.csv"])
    found = {Path(p).name for p in calist.find_source_files(tmp_path)}
    assert found == {"a.xlsx", "b.xls", "c.xlsm"}


def test_the_scan_reaches_sub_folders(tmp_path):
    _tree(tmp_path, ["top.xlsx", "ward/one.xlsx", "ward/lab/two.xlsx"])
    found = {Path(p).name for p in calist.find_source_files(tmp_path)}
    assert found == {"top.xlsx", "one.xlsx", "two.xlsx"}


def test_the_scan_skips_lock_files(tmp_path):
    _tree(tmp_path, ["G302-AGH001-0425.xlsx", "~$G302-AGH001-0425.xlsx"])
    found = [Path(p).name for p in calist.find_source_files(tmp_path)]
    assert found == ["G302-AGH001-0425.xlsx"]


def test_the_scan_stops_when_cancelled(tmp_path):
    _tree(tmp_path, [f"f{n}.xlsx" for n in range(20)])
    cancel = threading.Event()
    cancel.set()
    assert list(calist.find_source_files(tmp_path, cancel)) == []


def test_an_unreadable_folder_does_not_lose_the_rest(tmp_path, monkeypatch):
    """One permission error deep in a tree must not cost the whole round."""
    _tree(tmp_path, ["good.xlsx", "locked/hidden.xlsx"])
    real = os.scandir

    def refuse(path):
        if Path(path).name == "locked":
            raise PermissionError(13, "Access is denied")
        return real(path)

    monkeypatch.setattr(os, "scandir", refuse)
    found = [Path(p).name for p in calist.find_source_files(tmp_path)]
    assert found == ["good.xlsx"]

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



# ── quiet mode and the structured duplicate report ────────────────────────────
#
# Turbo runs tens of thousands of forms, where a line per file makes the log
# useless and the duplicate warnings are the thing worth keeping.

def test_duplicates_are_reported_as_data_not_only_as_a_log_line():
    records = [_sourced("ECG", "SN1", "D23-AGH001-0225.xlsx"),
               _sourced("ECG", "SN1", "D23-AGH007-0225.xlsx")]
    dropped = []
    calist.deduplicate_records(records, dropped)

    assert len(dropped) == 1
    hit = dropped[0]
    assert hit.serial == "SN1"
    assert hit.dropped == "D23-AGH007-0225.xlsx"
    assert hit.kept == "D23-AGH001-0225.xlsx"


def test_collecting_duplicates_is_optional():
    """Callers that only want the log must not have to pass a list."""
    records = [_sourced("ECG", "SN1", "a.xlsx"), _sourced("ECG", "SN1", "b.xlsx")]
    assert len(calist.deduplicate_records(records)) == 1


def test_quiet_mode_drops_the_line_per_file(workspace, caplog):
    files, template = workspace
    with caplog.at_level(logging.INFO, logger="aggregator"):
        calist.process_files(files, template, quiet=True)
    assert "[OK]" not in caplog.text


def test_quiet_mode_still_reports_every_problem(workspace, caplog):
    """It is the successes that make a log unreadable, never the failures."""
    files, template = workspace
    with caplog.at_level(logging.INFO, logger="aggregator"):
        calist.process_files(files, template, quiet=True)
    assert "ZZZ" in caplog.text                  # the unknown-code form


def test_a_loud_run_still_names_every_file(workspace, caplog):
    files, template = workspace
    with caplog.at_level(logging.INFO, logger="aggregator"):
        calist.process_files(files, template)
    assert "[OK]" in caplog.text


def test_quiet_mode_writes_the_same_register(workspace, tmp_path):
    """Logging less must not change a single row."""
    files, template = workspace
    (tmp_path / "loud").mkdir()
    (tmp_path / "quiet").mkdir()

    loud = calist.process_files(files, template, output_dir=tmp_path / "loud")
    quiet = calist.process_files(files, template, output_dir=tmp_path / "quiet",
                                 quiet=True)

    assert loud.succeeded and quiet.succeeded

    assert loud.rows_written == quiet.rows_written
    rows = [[c.value for c in row]
            for path in (loud.output_path, quiet.output_path)
            for row in load_workbook(path).active.iter_rows(min_row=4, max_row=6)]
    assert rows[:3] == rows[3:]


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


def test_real_forms_only_leaves_badly_named_files_out_of_a_run(tmp_path):
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

    good = _form(src / "G302-AC001-0425.xlsx",
                 {"E17": "Zoll", "E15": "R-Series", "K15": "SN-1",
                  "K17": "ER", "E13": "01/01/2024", "G24": "Working"})
    bad = _form(src / "Clinic-AC002.xlsx",
                {"E17": "Zoll", "E15": "R-Series", "K15": "SN-2",
                 "K17": "ER", "E13": "01/01/2024", "G24": "Working"})

    loose = calist.process_files([good, bad], str(template))
    assert loose.files_read == 2

    strict = calist.process_files([good, bad], str(template), real_forms_only=True)
    assert strict.files_read == 1
    assert strict.problems == []
    assert [o.filename for o in strict.left_out] == ["Clinic-AC002.xlsx"]
    assert strict.rows_written == 1


def test_second_row_code_survives_the_house_format():
    """G302-AGH001-0425 must become G302-AGCB001-0425, not mangle the rest."""
    parent = {"Code": "G302-AGH001-0425", "Status": "OK", "Status2": "Faulty",
              "_row_order": 0}
    child = calist.build_second_row(
        parent, {"device_name": "NIBP", "code_replace": ("AGH", "AGCB")})
    assert child["Code"] == "G302-AGCB001-0425"


def test_register_is_signed_with_the_author(workspace):
    """Credit must travel with the file, not just live in the app."""
    files, template = workspace
    result = calist.process_files(files, template)

    wb = load_workbook(result.output_path)
    ws = wb.active
    footer_row = calist.TEMPLATE_START_ROW + result.rows_written + 1
    footer = ws.cell(row=footer_row, column=calist.TEMPLATE_START_COL).value
    creator = wb.properties.creator
    wb.close()

    assert footer and calist.AUTHOR_NAME in footer
    assert calist.AUTHOR_EMAIL in footer
    assert "Calist" in footer
    assert calist.AUTHOR_NAME in creator          # File -> Properties in Excel


def test_attribution_sits_clear_of_the_data(workspace):
    """A blank row between the last record and the signature."""
    files, template = workspace
    result = calist.process_files(files, template)

    wb = load_workbook(result.output_path)
    ws = wb.active
    last_data = calist.TEMPLATE_START_ROW + result.rows_written - 1
    gap = ws.cell(row=last_data + 1, column=calist.TEMPLATE_START_COL).value
    wb.close()

    assert gap in (None, ""), "the signature must not touch the last record"


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


# ── reading a form whose map no longer fits ───────────────────────────────────
#
# Measured over the 2025 and 2026 archives: the configured map works for 94.7%
# of 42,826 forms. The rest were re-laid-out between rounds — a uniform shift of
# one to three rows — and read blank with nothing said. These pin the fallback
# chain, and the two traps that make it dangerous to write naively.

_STANDARD = {"Manufacturer": "E20", "Model": "E18", "S.N": "K18",
             "Location": "K20", "Date": "E16", "Status": "H32"}


def _device_form(path, row, *, sheet="Device data", extra=None, before=None):
    """A form in the usual shape, with its field block anchored at `row`."""
    wb = Workbook()
    wb.remove(wb.active)
    for title, cells in (before or []):
        ws = wb.create_sheet(title)
        for ref, value in cells.items():
            ws[ref] = value
    ws = wb.create_sheet(sheet)
    ws[f"C{row}"] = "Model:"
    ws[f"E{row}"] = "Perfusor Space"
    ws[f"I{row}"] = "Serial No.:"
    ws[f"K{row}"] = "148271"
    ws[f"C{row + 2}"] = "Manufacturer:"
    ws[f"E{row + 2}"] = "B.Braun"
    ws[f"I{row + 2}"] = "Location"
    ws[f"K{row + 2}"] = "ICU"
    for ref, value in (extra or {}).items():
        ws[ref] = value
    wb.save(path)
    return str(path)


def test_the_configured_map_wins_when_it_fits(tmp_path):
    path = _device_form(tmp_path / "f.xlsx", 18)
    record, how = calist.read_best(path, {"cells": _STANDARD})
    assert how == "primary"
    assert record["Model"] == "Perfusor Space"


def test_an_alternate_layout_is_tried_next(tmp_path):
    path = _device_form(tmp_path / "f.xlsx", 19)          # shifted one row down
    config = {"cells": _STANDARD,
              "alt_cells": [{**_STANDARD, "Model": "E19", "S.N": "K19",
                             "Manufacturer": "E21", "Location": "K21"}]}
    record, how = calist.read_best(path, config)
    assert how == "alt 1"
    assert record["S.N"] == "148271"


def test_the_printed_labels_rescue_an_unrecorded_shift(tmp_path):
    """No alt_cells for this offset — the form's own labels have to find it."""
    path = _device_form(tmp_path / "f.xlsx", 15)
    record, how = calist.read_best(path, {"cells": _STANDARD})
    assert how == "labels"
    assert record["Model"] == "Perfusor Space"
    assert record["S.N"] == "148271"
    assert record["Manufacturer"] == "B.Braun"
    assert record["Location"] == "ICU"


def test_the_device_block_is_found_on_another_sheet(tmp_path):
    """GC forms put a Report tab in front; BM keeps the details on a cover."""
    path = _device_form(tmp_path / "f.xlsx", 24, sheet="cover page",
                        before=[("Report", {"A1": "Electrical Safety Test",
                                            "A2": "Protective Earth"})])
    record, how = calist.read_best(path, {"cells": _STANDARD})
    assert how == "labels on 'cover page'"
    assert record["Model"] == "Perfusor Space"


def test_the_calibrator_is_never_read_as_the_device(tmp_path):
    """The trap this whole guard exists for.

    A Hemodialysis cover page prints the reference meter ABOVE the device under
    test, so "first Model: label wins" records the calibrator's model and serial
    as the machine's. The headings are the only reliable signal.
    """
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("cover page")
    ws["A15"] = "Calibration Device"
    ws["G15"] = "Model"
    ws["H15"] = "EMIS"                                   # the Fluke
    ws["G18"] = "Serial No."
    ws["H18"] = "9D2749"
    ws["A22"] = "Device information"
    ws["G24"] = "Model"
    ws["H24"] = "AK96"                                   # the dialysis machine
    ws["A26"] = "Manufacturer"
    ws["C26"] = "Gambro"
    ws["G26"] = "Serial No."
    ws["H26"] = "11195"
    path = tmp_path / "f.xlsx"
    wb.save(path)

    record, how = calist.read_best(str(path), {"cells": _STANDARD})

    assert record["Model"] == "AK96", "read the calibrator instead of the device"
    assert record["S.N"] == "11195"
    assert record["Manufacturer"] == "Gambro"
    assert "EMIS" not in record.values()
    assert "9D2749" not in record.values()


def test_a_table_heading_is_not_mistaken_for_a_field(tmp_path):
    """Therapeutic Ultrasound heads a table 'Model | S.N.' at row 11 and puts
    the real fields at row 69. Taking the first match read the model as "S.N.".
    """
    path = _device_form(tmp_path / "f.xlsx", 69,
                        extra={"A11": "Model", "E11": "S.N."})
    record, how = calist.read_best(path, {"cells": _STANDARD})
    assert record["Model"] == "Perfusor Space"
    assert record["S.N"] == "148271"


def test_a_form_of_placeholders_is_not_rescued(tmp_path):
    """Some cover sheets are filled in with 0 throughout. A serial of "0"
    classifies as a placeholder rather than a blank, so without an explicit
    check the fallback "rescues" them into a row of zeroes.
    """
    # Named for a device: a blank form with a proper name is still read (and
    # reports "none"); only one not named for a device is skipped.
    path = _device_form(tmp_path / "G302-AGH001-0425.xlsx", 15)
    wb = load_workbook(path)
    ws = wb["Device data"]
    for ref in ("E15", "K15", "E17", "K17"):
        ws[ref] = "0"
    wb.save(path)

    _record, how = calist.read_best(str(path), {"cells": _STANDARD})
    assert how == "none"


def test_nothing_readable_reports_none(tmp_path):
    path = _multi_sheet_form(tmp_path / "G302-AGH001-0425.xlsx",
                             [("Sheet", {"A1": "nothing"})])
    _record, how = calist.read_best(path, {"cells": _STANDARD})
    assert how == "none"


def test_plausible_rejects_a_serial_that_is_not_one():
    assert not calist.plausible({"S.N": "ICU", "Model": "MX450"})
    assert not calist.plausible({"S.N": "16-01-2026", "Model": "MX450"})
    assert not calist.plausible({"S.N": "SN1", "Model": ""})
    assert calist.plausible({"S.N": "SN1", "Model": "MX450"})


def test_classify_serial_sorts_the_shapes_apart():
    assert calist.classify_serial("") == "blank"
    assert calist.classify_serial("N.A") == "placeholder"
    assert calist.classify_serial("ICU") == "suspect"
    assert calist.classify_serial("16-01-2026") == "suspect"
    assert calist.classify_serial("BALANCE001") == "assigned"
    assert calist.classify_serial("STX21170332PA") == "real"


# ── the device table after the archive validation ─────────────────────────────

def test_every_cell_reference_is_a_real_a1_reference():
    """Catches a typo in a coordinate, which would otherwise read blank."""
    from openpyxl.utils.cell import coordinate_to_tuple
    for code, config in DEVICE_CONFIGS.items():
        maps = [config["cells"], *config.get("alt_cells", [])]
        for cell_map in maps:
            for field, ref in cell_map.items():
                if not ref:
                    continue                      # deliberately unmapped field
                coordinate_to_tuple(ref)          # raises if malformed
                assert re.fullmatch(r"[A-Z]{1,3}\d{1,4}", ref), f"{code}.{field}={ref}"


def test_mammography_is_read_only_from_its_word_certificate():
    """BD must never be read from its workbook.

    All BD workbooks carry the same header — GE / Alpha st / Gona Hospital,
    survey date 2012 — across many different site codes: the vendor QC
    template's boilerplate, not the device. A cell map would write the same
    fabricated manufacturer and model into every row and look entirely
    plausible, so BD stayed unmapped until the 2026 archive audit found where
    the device really is: a same-named Word certificate beside each workbook
    (48 of 53). No cell of the workbook may be mapped.
    """
    config = DEVICE_CONFIGS["BD"]
    assert config["source"] == "word"
    assert not any(config["cells"].values())
    assert not config.get("alt_cells")


def test_the_devices_added_from_the_archive_are_all_named():
    """Every code mapped here is one the site's master list actually names."""
    from device_names import DEVICE_NAMES
    added = {"FA", "EQ", "FV", "FM", "FR", "FF", "DE", "FD", "DB", "CZ", "GM",
             "BQ", "EZ", "FU", "FT", "CP", "FW", "DU", "AY", "BW", "BX", "FP",
             "DW", "DS", "GH", "DX", "BC", "CD", "EN"}
    assert added <= set(DEVICE_CONFIGS), added - set(DEVICE_CONFIGS)
    assert added <= set(DEVICE_NAMES), added - set(DEVICE_NAMES)


def test_a_form_with_no_status_field_maps_status_to_nothing():
    """CT, MRI, Dexa and the air mattress end at "Tested By" — no Status box.

    Mapping Status to a borrowed coordinate would read whatever sits there.
    """
    for code in ("BW", "BX", "FP", "DW"):
        assert DEVICE_CONFIGS[code]["cells"]["Status"] == "", code


# ── The Status column: captions, and a box that has moved ─────────────────────
#
# Two separate faults met here. A cell map can place the identity block
# perfectly and still miss the verdict box, and `plausible()` asks only for a
# serial and a model — so read_best never falls through and the miss is silent.
# Where the map then landed on a caption, that caption was written into the
# register as the device's outcome.


class _Grid:
    """Stands in for an open workbook: answers values() from a dict."""

    def __init__(self, cells):
        self._cells = cells
        self.reads = 0

    def values(self, refs):
        self.reads += 1
        return {ref: self._cells.get(ref, "") for ref in refs}


def test_a_caption_is_never_written_as_a_status():
    """A caption is printed beside the box, never in it.

    A syringe form that fell through to an alternate had its verdict recorded
    as 'Syringe brand:' on 50 of 61 sampled files, and 'Safety:' on 7 more. A
    blank is honest; that was not. None of these may reach the register.
    """
    for caption in ("Safety:", "Syringe brand:", "Contact Person Name:",
                    "Phone No.:", "Status", "Status:"):
        record = {"Status": caption}
        calist._settle_status(_Grid({}), {"Status": "G35"}, record)
        assert record["Status"] == "", caption


def test_an_answer_to_a_different_question_is_never_written_as_a_status():
    """A map one column out reports the Safety size, which reads like an answer.

    "Large" is not a caption, so the caption guard cannot catch it — and left
    alone it looks entirely plausible to somebody reading a register.
    """
    for value in ("Large", "Small", "N.A", "----"):
        record = {"Status": value}
        calist._settle_status(_Grid({}), {"Status": "K22"}, record)
        assert record["Status"] == "", value


def test_a_status_that_already_reads_is_never_second_guessed():
    """Step 1 returns before anything is searched.

    This is what keeps a device whose box never moves completely unaffected —
    10 of the 81 mapped Status cells read on every form sampled — and it is why
    the fallback cannot change a file the map already got right, even with
    another status sitting next to it.
    """
    grid = _Grid({"G35": "Pass", "G33": "Status:", "G34": "Fail"})
    record = {"Status": "Pass"}
    calist._settle_status(grid, {"Status": "G35"}, record)
    assert record["Status"] == "Pass"
    assert grid.reads == 0, "a readable status must cost no extra read"


def test_a_box_that_moved_sideways_is_found_by_its_label():
    """Boxes move across columns, not just down rows.

    Measured over 81 device maps: 33 have forms where the box sits in a
    different column. An offset search in the mapped column cannot find those,
    which is why the form's own printed label is what anchors the search.
    """
    grid = _Grid({"H32": "", "I22": "Status:", "K22": "Calibrated"})
    record = {"Status": ""}
    calist._settle_status(grid, {"Status": "H32"}, record)
    assert record["Status"] == "Calibrated"


def test_the_printed_legend_is_never_mistaken_for_the_answer():
    """These forms print the options across B-E on the answer's own row.

    Every one of those cells reads as a status, so a search that starts
    anywhere but the label will happily return one of them. The answer is the
    box the label points at.
    """
    grid = _Grid({
        "G33": "Status:",
        "B34": "Pass", "C34": "Fail", "D34": "Limited non", "E34": "limited fail",
        "G34": "Fail",                      # the real answer
    })
    record = {"Status": ""}
    calist._settle_status(grid, {"Status": "G35"}, record)
    assert record["Status"] == "Fail"


def test_a_two_box_device_is_never_given_one_verdict_for_both():
    """One printed "Status" label cannot say which module it belongs to.

    A patient monitor's D39 is the ECG and J39 the NIBP. Filling a missing one
    from the nearest label would put one module's verdict against the other.
    """
    grid = _Grid({"G30": "Status:", "G31": "Pass"})
    record = {"Status": "", "Status2": ""}
    calist._settle_status(grid, {"Status": "D39", "Status2": "J39"}, record)
    assert record == {"Status": "", "Status2": ""}
    assert grid.reads == 0


def test_a_blank_status_stays_blank_when_the_form_says_nothing():
    """A form nobody filled in must not acquire a verdict from anywhere."""
    grid = _Grid({"E18": "BC-5000", "K18": "SN-9"})
    record = {"Status": ""}
    calist._settle_status(grid, {"Status": "G35"}, record)
    assert record["Status"] == ""


# ── The cell maps these faults were found through ────────────────────────────


def test_the_maps_that_read_the_wrong_cell_stay_fixed():
    """Each of these read blank, or read the wrong thing, on real forms.

    Pinned because they are one-character edits that a later tidy-up could undo
    without anything failing: a wrong Status reads blank, and a blank is
    indistinguishable from a form nobody filled in.
    """
    # CBC: H32 is a row this form does not reach. 0 of 300 files read it.
    assert DEVICE_CONFIGS["DG"]["cells"]["Status"] == "K22"
    assert DEVICE_CONFIGS["DG"]["alt_cells"][0]["Status"] == "G29"
    # Ventilator: G33 serves the older template; 148 of 250 use G31.
    assert DEVICE_CONFIGS["AM"]["cells"]["Status"] == "G31"
    # Vital signs filed under a patient-monitor code: the verdicts are at
    # G38/J38, and AGH's D39/J39 are empty on all 12 files.
    assert DEVICE_CONFIGS["VAGH"]["cells"]["Status"] == "G38"
    assert DEVICE_CONFIGS["VAGH"]["cells"]["Status2"] == "J38"


def test_the_incubator_reads_the_model_below_the_manufacturer():
    """This form is inverted, and the map was not.

    Model sat where "Next calibration" is printed, so the register showed a
    date in the Model column. The alternate could not rescue it: that cell and
    the serial both held values, so the record looked plausible and read_best
    stopped at the primary.
    """
    for cells in [DEVICE_CONFIGS["AK"]["cells"]] + DEVICE_CONFIGS["AK"]["alt_cells"]:
        model, manufacturer = cells["Model"], cells["Manufacturer"]
        assert model[0] == manufacturer[0]
        assert int(model[1:]) == int(manufacturer[1:]) + 2, cells


def test_every_status_cell_is_a_real_reference_or_deliberately_empty():
    """A Status is either a cell, or "" for a form that has no box at all.

    There is no useful rule about *which* column: the legend that makes B-E
    dangerous on the F/G-answer forms does not exist on the patient monitor,
    whose answer genuinely sits at D39 and reads on every file. So the legend
    is guarded where it can actually be hit — at the search, by anchoring on
    the printed label — not by a column rule that would be false here.
    """
    for code, config in DEVICE_CONFIGS.items():
        maps = [("cells", config["cells"])]
        maps += [(f"alt {i}", a) for i, a in enumerate(config.get("alt_cells", []))]
        for name, cells in maps:
            status = cells.get("Status", "")
            if status:
                assert re.fullmatch(r"[A-Z]+\d+", status), f"{code} {name}: {status}"


# ── The 2026 archive audit: forms that moved, and what the register marks ────
#
# Every test below comes from a failure found by reading all 81,000 forms of
# the 2023-2026 archive (tools/audit.py) — each one silent in the register
# until then. Workbooks are built on the fly: the repository is public and no
# customer form may ever be committed.

def _block(ws, row, values, captions=True, col="E", val="K"):
    """The standard device block — Model / Serial on `row`, Manufacturer /
    Location two rows below — with or without its printed captions."""
    mid = chr(ord(val) - 4)
    if captions:
        ws[f"A{row}"], ws[f"{mid}{row}"] = "Model:", "Serial No.:"
        ws[f"A{row + 2}"], ws[f"{mid}{row + 2}"] = "Manufacturer:", "Location:"
    ws[f"{col}{row}"], ws[f"{val}{row}"] = values["Model"], values["S.N"]
    ws[f"{col}{row + 2}"], ws[f"{val}{row + 2}"] = values["Manufacturer"], values["Location"]


def test_a_workbook_is_read_by_what_it_is_not_what_it_is_called(tmp_path):
    """Eleven archive files are .xlsx workbooks saved under a .xls name."""
    path = _form(tmp_path / "G302-AC001-0425.xls", {"E15": "R-Series", "K15": "SN-77"})
    record, _ = calist.read_best(path, DEVICE_CONFIGS["AC"])
    assert record["S.N"] == "SN-77"


def test_a_serial_shaped_like_a_date_is_only_a_date_if_it_is_one():
    """Water Bath serial 2110/04/98 — year 2110, day 98 — made the reader
    reject the right layout."""
    assert calist.classify_serial("2110/04/98") == "real"
    assert calist.classify_serial("16-01-2026") == "suspect"
    assert calist.is_calendar_date("2024-01-15")
    assert not calist.is_calendar_date("99/99/99")


def test_a_model_is_never_its_own_serial():
    """A Phototherapy layout eight rows down lands both on the text D38."""
    assert not calist.plausible({"S.N": "D38", "Model": "D38"})
    assert calist.plausible({"S.N": "BXP-7823-146", "Model": "Bililed Max +"})


def test_a_layout_whose_captions_name_other_fields_is_rejected(tmp_path):
    """The Baby Warmer failure: two layouts, and the wrong one reads text
    that looks like a perfectly good record — the maker as the model, the
    ward as the serial. Only the captions beside the boxes give it away."""
    wb = Workbook()
    _block(wb.active, 18, {"Model": "MX450", "S.N": "SN-1234",
                           "Manufacturer": "Philips", "Location": "Room 12"})
    path = tmp_path / "G302-XX001-0425.xlsx"
    wb.save(path)
    config = {"device_name": "Test", "cells": form(20, "K30"),
              "alt_cells": [form(18, "K30")]}
    # The map two rows low reads Model=Philips beside "Manufacturer:" and
    # S.N="Room 12" beside "Location:" — plausible values, wrong captions.
    record, how = calist.read_best(str(path), config)
    assert how == "alt 1"
    assert (record["Model"], record["S.N"]) == ("MX450", "SN-1234")
    assert record["Manufacturer"] == "Philips"


def test_a_form_without_captions_keeps_its_mapped_layout(tmp_path):
    """Only a contradiction counts: no caption proves nothing either way."""
    wb = Workbook()
    _block(wb.active, 18, {"Model": "MX450", "S.N": "SN-1234",
                           "Manufacturer": "Philips", "Location": "ICU"},
           captions=False)
    path = tmp_path / "G302-XX001-0425.xlsx"
    wb.save(path)
    record, how = calist.read_best(str(path), {"device_name": "Test",
                                               "cells": form(18, "K30")})
    assert how == "primary" and record["Model"] == "MX450"


def test_a_field_is_taken_from_the_cell_its_caption_names(tmp_path):
    """Lab Oven: Location is K20 on 68 forms and K19 on two."""
    wb = Workbook()
    ws = wb.active
    _block(ws, 18, {"Model": "UN55", "S.N": "AR1805V042",
                    "Manufacturer": "Memmert", "Location": ""})
    ws["G20"] = None
    ws["G19"], ws["K19"] = "Location:", "Lab"
    path = tmp_path / "B14-EU001-0226.xlsx"
    wb.save(path)
    record, _ = calist.read_best(str(path), DEVICE_CONFIGS["EU"])
    assert record["Location"] == "Lab"


def test_an_empty_box_captioned_status_beats_a_tested_by_box(tmp_path):
    """Centrifuge .xls variant: K25 is "Tested by", and JTE-40 is an
    engineer, not the verdict. The Status box is empty: say so."""
    wb = Workbook()
    ws = wb.active
    _block(ws, 18, {"Model": "5702R", "S.N": "5703XQ815492",
                    "Manufacturer": "Eppendorf", "Location": "Lab"})
    ws["J22"] = "Status:"
    ws["G25"], ws["K25"] = "Tested by:", "JTE-40"
    path = tmp_path / "H62-AS002-0824.xlsx"
    wb.save(path)
    record, _ = calist.read_best(str(path), DEVICE_CONFIGS["AS"])
    assert record["Status"] == ""


def test_a_device_list_is_not_read_as_a_device(tmp_path):
    """A file named "Device List H23-AS-1023" resolves to a Centrifuge."""
    wb = Workbook()
    ws = wb.active
    ws.title = "list"
    for column, heading in enumerate(["No.", *calist.FIELDS], start=1):
        ws.cell(row=3, column=column, value=heading)
    path = tmp_path / "Device List H23-AS-1023.xlsx"
    wb.save(path)
    with pytest.raises(calist.NotAForm):
        calist.read_best(str(path), DEVICE_CONFIGS["AS"])
    _, outcomes = calist.extract_records([str(path)])
    assert outcomes[0].status == calist.UNSUPPORTED
    assert "device list" in outcomes[0].detail.lower()


_CERTIFICATE_TEXT = (
    "Calibration Data Calibration Eng. JTE-40 Test Date: 15-06-2023 "
    "Calibration Device: Digital KVP meter Dose meter Model : 07-492 "
    "S.N : 108357 Calibration Method: MECL-WI-20 Attachments 2 pages "
    "Equipment Data Equipment Type: Mammography Model: FDR3500DRLH "
    "Serial No.: 6481207 Manufacturer: FujiFilm Location: Radiology "
    "Prev. Calib.: N.A Date of receipt: 15-06-2023 Next Calib. On: 15-06-2024 "
    "Environmental conditions Temperature: 24 C")


def _docx(path, text):
    runs = "".join(f"<w:r><w:t>{word} </w:t></w:r>" for word in text.split())
    xml = ('<?xml version="1.0"?><w:document xmlns:w="http://schemas.'
           'openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>'
           f"{runs}</w:p></w:body></w:document>")
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", xml)
    return str(path)


def test_the_word_certificate_is_read_after_equipment_data_only(tmp_path):
    """It prints the dose meter FIRST; that must never become the device."""
    _docx(tmp_path / "F22-BD001-0623.docx", _CERTIFICATE_TEXT)
    record = calist.read_word_certificate(str(tmp_path / "F22-BD001-0623.xlsm"))
    assert record["Model"] == "FDR3500DRLH"
    assert record["S.N"] == "6481207"
    assert record["Manufacturer"] == "FujiFilm"
    assert record["Location"] == "Radiology"
    assert record["Date"] == "15-06-2023"


def test_a_legacy_doc_certificate_is_read_too(tmp_path):
    """47 of the 48 are Word 97 .doc: plain text runs among binary."""
    (tmp_path / "F22-BD001-0623.doc").write_bytes(
        b"\xd0\xcf\x11\xe0\x00\x01" + _CERTIFICATE_TEXT.encode("cp1252") + b"\x00\x07")
    record = calist.read_word_certificate(str(tmp_path / "F22-BD001-0623.xlsm"))
    assert (record["Model"], record["S.N"]) == ("FDR3500DRLH", "6481207")


def test_mammography_ignores_its_workbook_entirely(tmp_path):
    """The QC template prints "Model" beside the X-ray TUBE."""
    wb = Workbook()
    wb.active["B4"], wb.active["E4"], wb.active["F4"] = "Model", "Model", "Tube-99"
    wb.save(tmp_path / "F22-BD001-0623.xlsx")
    _docx(tmp_path / "F22-BD001-0623.docx", _CERTIFICATE_TEXT)
    record, how = calist.read_best(str(tmp_path / "F22-BD001-0623.xlsx"),
                                   DEVICE_CONFIGS["BD"])
    assert how == "word certificate"
    assert record["Model"] == "FDR3500DRLH"


def test_a_form_no_layout_fits_is_read_from_its_word_certificate(tmp_path):
    """Hemodialysis "Final" printouts carry no device block at all."""
    wb = Workbook()
    wb.active.title = "Final"
    wb.active["A30"], wb.active["E20"] = "2. Protective Earth Resistance", "Line1-Line2"
    wb.save(tmp_path / "H36-BM008-0923.xlsx")
    _docx(tmp_path / "H36-BM008-0923.docx", _CERTIFICATE_TEXT)
    record, how = calist.read_best(str(tmp_path / "H36-BM008-0923.xlsx"),
                                   DEVICE_CONFIGS["BM"])
    assert how == "word certificate"
    assert record["S.N"] == "6481207"


@pytest.mark.parametrize("sentence, verdict", [
    ("Tested Equipment ONLY Passed all Test and Related Parameters Stated in "
     "this Report", "Pass"),
    ("The device under test is Limited Calibrated, non calibrated item "
     "(Electrical Safety)", "Limited non"),
    ("The stated data is the device readings and there's no accepted range "
     "and reference.", ""),
])
def test_the_hemodialysis_verdict_comes_from_its_conclusion(tmp_path, sentence, verdict):
    """The only three wordings in the archive; the row moves (A119-A132)."""
    wb = Workbook()
    wb.active.title = "Data entry"
    final = wb.create_sheet("Final")
    final["K22"] = "Pass"                   # one test line's result, not the verdict
    final["A127"], final["A128"] = "Conclusion:", sentence
    path = tmp_path / "H36-BM001-0923.xlsx"
    wb.save(path)
    record, _ = calist.read_best(str(path), DEVICE_CONFIGS["BM"])
    assert record["Status"] == verdict


def _template(tmp_path):
    template = tmp_path / "Device List.xlsx"
    wb = Workbook()
    for column, heading in enumerate(["No.", *calist.FIELDS], start=1):
        wb.active.cell(row=3, column=column, value=heading)
    wb.save(template)
    return str(template)


def _cell(sheet, field):
    return sheet.cell(row=calist.TEMPLATE_START_ROW,
                      column=calist.TEMPLATE_START_COL + calist.FIELDS.index(field))


def test_a_doubtful_value_is_kept_amber_with_a_note(tmp_path):
    record = {"Device": "Syringe", "Model": "FA313", "S.N": "04-01-2022",
              "Status": "Pass",
              "_flags": {"S.N": {"show": None, "reason": "Not a serial."}}}
    out = tmp_path / "out.xlsx"
    calist.write_output([record], _template(tmp_path), out)
    sheet = load_workbook(out).active
    serial = _cell(sheet, "S.N")
    assert serial.value == "04-01-2022"
    assert serial.fill.fgColor.rgb.endswith("FFE699")
    assert "Not a serial" in serial.comment.text
    model = _cell(sheet, "Model")
    assert model.comment is None and model.fill.fill_type is None


def test_a_caption_in_the_status_box_reaches_the_register_but_not_the_export(tmp_path):
    """Kept and marked in the register, by the owner's decision; still blank
    for the Firebase export, which reads the named field only."""
    wb = Workbook()
    ws = wb.active
    _block(ws, 18, {"Model": "Bililed", "S.N": "BXP-7266-318",
                    "Manufacturer": "Novos", "Location": "NICU"}, col="D", val="J")
    ws["F37"] = "Safety:"
    path = tmp_path / "D38-AL010-0224.xlsx"
    wb.save(path)
    config = {"device_name": "Phototherapy", "cells": form(18, "F37", col="D", val="J")}
    record, _ = calist.read_best(str(path), config)
    assert record["Status"] == ""                            # the export
    assert record["_flags"]["Status"]["show"] == "Safety:"   # the register


def test_an_empty_verdict_box_showing_zero_is_not_flagged(tmp_path):
    """The NIBP box copies an empty test sheet and shows 0: empty, not wrong."""
    wb = Workbook()
    ws = wb.active
    _block(ws, 18, {"Model": "MX450", "S.N": "SN-2",
                    "Manufacturer": "Philips", "Location": "ICU"})
    ws["D39"], ws["J39"] = "Pass", 0
    path = tmp_path / "H05-AGH023-0923.xlsx"
    wb.save(path)
    record, _ = calist.read_best(str(path), DEVICE_CONFIGS["AGH"])
    assert "Status2" not in record["_flags"]


def test_the_module_row_takes_its_own_status_mark_not_the_parents():
    parent = {"Code": "AGH001", "Status": "Pass", "Status2": "Safety:",
              "_row_order": 0,
              "_flags": {"Status2": {"show": "Safety:", "reason": "Not a verdict."},
                         "Model": {"show": None, "reason": "Check."}}}
    module = calist.build_second_row(
        parent, {"device_name": "NIBP", "code_replace": ("AGH", "AGCB")})
    assert module["_flags"]["Status"]["show"] == "Safety:"
    assert "Model" in module["_flags"] and "Status2" not in module["_flags"]


# ── Regressions the archive audit caught in its own fixes ─────────────────────
#
# The before/after diff over all 81,000 forms found these in the first round
# of fixes. Each is pinned so it cannot come back quietly.

def test_a_reading_is_not_a_serial_but_a_long_decimal_can_be():
    """Test tables read "7.1" or "0.99"; a Therapeutic Ultrasound's
    serial looks like "37.254", and a float-exported serial is "1873206.0"."""
    for reading in ("7.1", "25.8", "0.99", "6.9"):
        assert calist.classify_serial(reading) == "suspect", reading
    # A whole number ending ".0" is how an .xls hands back a numeric serial:
    # a Nebulizer's serial 115 arrives as "115.0".
    for serial in ("37.254", "1873206.0", "104357", "115.0", "99.0"):
        assert calist.classify_serial(serial) != "suspect", serial


def test_a_date_with_no_label_is_placed_from_the_located_model(tmp_path):
    """Therapeutic Ultrasound's block moves rows AND columns, so no uniform
    offset exists; its Date sits two rows above its Model wherever that is.
    Without this the Date read the caption "Test Parameter" on 95 forms."""
    wb = Workbook()
    ws = wb.active
    ws.title = "data entry"
    ws["E16"] = "Test Parameter"
    _block(ws, 69, {"Model": "Sono plus 490", "S.N": "37.254",
                    "Manufacturer": "Enraf Nonius", "Location": "Physiotherapy"},
           col="D", val="J")
    ws["D67"] = "06-03-2026"
    path = tmp_path / "F11-BN001-0326.xlsx"
    wb.save(path)
    record, how = calist.read_best(str(path), DEVICE_CONFIGS["BN"])
    assert how == "labels"
    assert record["Date"] == "06-03-2026"
    assert record["S.N"] == "37.254"


def test_the_data_tab_is_searched_before_the_certificate(tmp_path):
    """The certificate repeats identity but never the verdict. Reaching it
    first lost the Status of forms the data tab used to read in full."""
    wb = Workbook()
    wb.active.title = "Report"
    wb.active["B2"] = "Calibration report"      # the real Report tab is not empty
    cert = wb.create_sheet("Certificate")
    cert["D26"] = "Equipment Data"
    _block(cert, 29, {"Model": "Erkameter3000", "S.N": "17518302",
                      "Manufacturer": "Erka", "Location": "Clinics"})
    data = wb.create_sheet("Inserting Data")
    _block(data, 18, {"Model": "Erkameter3000", "S.N": "17518302",
                      "Manufacturer": "Erka", "Location": "Clinics"})
    data["G58"], data["G59"] = "Status", "pass"
    path = tmp_path / "G225-CE003-0226.xlsx"
    wb.save(path)
    record, how = calist.read_best(str(path), DEVICE_CONFIGS["CE"])
    assert how == "labels on 'Inserting Data'"
    assert record["Status"] == "pass"


def test_a_placeholder_serial_loses_to_a_real_one_on_a_later_tab(tmp_path):
    """Two SPO2 forms: the cover report, searched first, copies an empty
    serial box as "0"; the certificate after it has the real serial. Take
    the real one, rather than stopping at the first tab that reads at all."""
    wb = Workbook()
    wb.active.title = "Report"
    wb.active["B2"] = "Calibration report"
    cert = wb.create_sheet("Certificate")
    cert["D26"] = "Equipment Data"
    _block(cert, 29, {"Model": "Rad-5", "S.N": "708246193",
                      "Manufacturer": "Masimo", "Location": "ICU"})
    cover = wb.create_sheet("cover report")
    _block(cover, 20, {"Model": "Rad-5", "S.N": "0",
                       "Manufacturer": "Masimo", "Location": "ICU"})
    path = tmp_path / "K116-AH001-0123.xlsx"
    wb.save(path)
    record, how = calist.read_best(str(path), DEVICE_CONFIGS["AH"])
    assert how == "labels on 'Certificate'"
    assert record["S.N"] == "708246193"


def test_a_typed_date_range_loses_to_a_real_date(tmp_path):
    """An engineer typed "14-24/09/2023" into one of two date boxes."""
    wb = Workbook()
    ws = wb.active
    _block(ws, 15, {"Model": "Valleylab FT10", "S.N": "T9A31574DX",
                    "Manufacturer": "Covidien", "Location": "OR"})
    ws["A11"], ws["E11"] = "Date:", "25-09-2023"
    ws["A13"], ws["E13"] = "Date:", "14-24/09/2023"
    path = tmp_path / "K140-AE001-0923.xlsx"
    wb.save(path)
    record, _ = calist.read_best(str(path), DEVICE_CONFIGS["AE"])
    assert record["Date"] == "25-09-2023"


def test_a_rescue_from_a_data_tab_takes_its_verdict_from_that_tab(tmp_path):
    """When the form opens on a test sheet, the map's Status cell there is a
    test line's result. A Balance form read "Fail" from it once."""
    wb = Workbook()
    report = wb.active
    report.title = "Report"
    report["B2"], report["G30"] = "Test results", "Fail"
    data = wb.create_sheet("Device Data")
    _block(data, 18, {"Model": "PB3002", "S.N": "B123456789",
                      "Manufacturer": "Mettler", "Location": "Lab"})
    data["F30"], data["G30"] = "Status:", "Pass"
    path = tmp_path / "G241-BP004-0125.xlsx"
    wb.save(path)
    record, how = calist.read_best(str(path), DEVICE_CONFIGS["BP"])
    assert how == "labels on 'Device Data'"
    assert record["Status"] == "Pass"


def test_a_confirmed_map_keeps_its_verdict_when_a_cover_page_supplies_the_identity(tmp_path):
    """An Infusion form whose serial was never entered: the map fits (its
    captions agree), so the identity comes from the cover page — and the
    verdict and date the map read on the form's own tab come with it.
    Losing them cost 408 fields across the archive."""
    wb = Workbook()
    data = wb.active
    data.title = "Data entry"
    _block(data, 27, {"Model": "Benefusion VP3", "S.N": "",
                      "Manufacturer": "Mindray", "Location": "NICU"})
    data["E25"], data["G36"] = "27-11-2025", "Pass"
    cover = wb.create_sheet("cover page")
    _block(cover, 20, {"Model": "Benefusion VP3", "S.N": "0",
                       "Manufacturer": "Mindray", "Location": "NICU"})
    path = tmp_path / "H67-AI000-1125.xlsx"
    wb.save(path)
    record, how = calist.read_best(str(path), DEVICE_CONFIGS["AI"])
    assert how == "labels on 'cover page'"
    assert record["Status"] == "Pass"
    assert record["Date"] == "27-11-2025"


# ── Files not named for a device ──────────────────────────────────────────────
#
# By the owner's decision these are skipped, like device lists, when the form
# is blank or says outright that it is another device. A properly named form
# is never skipped, even when blank.

def _certificate(ws, equipment_type, row=26):
    ws[f"D{row}"] = "Equipment Data"
    ws[f"B{row + 2}"], ws[f"D{row + 2}"] = "Equipment Type:", equipment_type


def test_a_file_named_by_serial_is_skipped_when_its_form_is_another_device(tmp_path):
    """"CN84017253" is a Philips serial, and "CN" reads as the Microwave's
    code. The form's own certificate says Patient Monitor."""
    wb = Workbook()
    data = wb.active
    data.title = "Device data"
    _block(data, 18, {"Model": "CM12", "S.N": "CN84017253",
                      "Manufacturer": "PHILIPS", "Location": "ICU"})
    _certificate(wb.create_sheet("Certificate"), "Patient Monitor")
    path = tmp_path / "CN84017253.xlsx"
    wb.save(path)
    with pytest.raises(calist.NotAForm, match="Patient Monitor"):
        calist.read_best(str(path), DEVICE_CONFIGS["CN"])


def test_a_file_named_by_serial_is_read_when_its_form_agrees(tmp_path):
    """"BZ009.xlsx" is not a device name, but a Syringe Pump is a Syringe."""
    wb = Workbook()
    data = wb.active
    data.title = "Device data"
    _block(data, 26, {"Model": "Agilia SP MC", "S.N": "23841956",
                      "Manufacturer": "Fresenius", "Location": "ICU"})
    _certificate(wb.create_sheet("Certificate"), "Syringe Pump", row=40)
    path = tmp_path / "BZ009.xlsx"
    wb.save(path)
    record, _ = calist.read_best(str(path), DEVICE_CONFIGS["BZ"])
    assert record["S.N"] == "23841956"


def _blank_ultrasound(path):
    """Captions and nothing else — plus the safety tab's own analyzer serial,
    which is a calibrator's and must not make the form count as filled in."""
    wb = Workbook()
    data = wb.active
    data.title = "Device data"
    data["A17"], data["H17"] = "Model:", "Serial No.:"
    data["A19"], data["H19"] = "Manufacturer:", "Location:"
    safety = wb.create_sheet("safety")
    safety["A10"], safety["C10"] = "S.N", "4271836"
    wb.save(path)
    return str(path)


def test_a_blank_form_not_named_for_a_device_is_skipped(tmp_path):
    path = _blank_ultrasound(tmp_path / "BB.xlsx")
    with pytest.raises(calist.NotAForm, match="nothing is filled in"):
        calist.read_best(path, DEVICE_CONFIGS["BB"])


def test_a_blank_form_named_for_a_device_is_still_read(tmp_path):
    path = _blank_ultrasound(tmp_path / "G302-BB001-0425.xlsx")
    _record, how = calist.read_best(path, DEVICE_CONFIGS["BB"])
    assert how == "none"


def test_a_sites_device_list_is_refused_whatever_it_is_called(tmp_path):
    """"Al baeerat.xlsx" is a site's register — headed "Device name" and "SN"
    — and "Al" reads as the Phototherapy code."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Al baeerat"
    ws.append(["Al Bairat Medical"])
    ws.append([])
    ws.append(["No.", "Device name", "Manufacturer", "Model", "SN", "Location",
               "test date", "code", "status"])
    ws.append([1, "Balance", "mediRhein", "MRCMT26", "N.A(1)", "Clinics",
               "27-12-2023", "G60-BP001-1223", "Pass"])
    path = tmp_path / "Al baeerat.xlsx"
    wb.save(path)
    with pytest.raises(calist.NotAForm, match="device list"):
        calist.read_best(str(path), DEVICE_CONFIGS["AL"])


# ── Dates found by their caption; statuses that are not verdicts ──────────────

def test_a_rescue_from_a_cover_page_takes_the_cover_pages_test_date(tmp_path):
    """The label search places no Date, so a rescue from another tab left it
    blank on 190 forms. The cover page prints it beside "Test Date" — near a
    next-calibration date and an issue date, which are not it."""
    wb = Workbook()
    data = wb.active
    data.title = "Data entry"
    data["A2"] = "Performance Test"
    cover = wb.create_sheet("cover report")
    _block(cover, 20, {"Model": "Tec 7", "S.N": "CKSA01234",
                       "Manufacturer": "Datex-Ohmeda", "Location": "OR"})
    cover["A16"], cover["C16"] = "Next Calib. On:", "05-01-2024"
    cover["A25"], cover["C25"] = "Issue Date/ Issue No.", "06-01-2023"
    cover["A27"], cover["C27"] = "Test Date", "05-01-2023"
    path = tmp_path / "H15-AB001-0123.xlsx"
    wb.save(path)
    record, how = calist.read_best(str(path), DEVICE_CONFIGS["AB"])
    assert how == "labels on 'cover report'"
    assert record["Date"] == "05-01-2023"


def test_an_ultrasound_date_one_row_low_is_read_from_its_captioned_box(tmp_path):
    """Some 2024 Ultrasound forms print "Date of receipt" in row 16, not 15."""
    wb = Workbook()
    ws = wb.active
    _block(ws, 17, {"Model": "Logiq P7", "S.N": "LP7000812",
                    "Manufacturer": "GE", "Location": "Outpatient"}, col="F", val="L")
    ws["C14"], ws["F14"] = "Issue Date/ Issue No", "09-07-2024"
    ws["C16"], ws["F16"] = "Date of receipt:", "10-07-2024"
    path = tmp_path / "F22-BB005-0724.xlsx"
    wb.save(path)
    record, how = calist.read_best(str(path), DEVICE_CONFIGS["BB"])
    assert how == "primary"
    assert record["Date"] == "10-07-2024"


def test_an_engineers_code_is_never_a_status(tmp_path):
    """Older Centrifuge forms: the map's K25 is the "Revised by" box, its
    caption above it, and "JTE-" was the verdict on 45 forms. The Status box
    was left empty, so the honest status is blank — kept amber in the register."""
    wb = Workbook()
    ws = wb.active
    _block(ws, 18, {"Model": "PLC-05", "S.N": "2000915",
                    "Manufacturer": "Gemmy", "Location": "Lab"})
    ws["I22"] = "Status:"
    ws["B24"], ws["G24"], ws["K24"] = "Tested By:", "Entered by:", "Revised by:"
    ws["B25"], ws["G25"], ws["K25"] = "JTE-", "JTE-", "JTE-"
    path = tmp_path / "G11-AS010-1223.xlsx"
    wb.save(path)
    record, _ = calist.read_best(str(path), DEVICE_CONFIGS["AS"])
    assert record["Status"] == ""
    assert record["_flags"]["Status"]["show"] == "JTE-"


@pytest.mark.parametrize("raw, status", [(191.2, ""), (5.1, ""), (0, "0"), ("Pass", "Pass")])
def test_a_reading_in_the_status_box_is_not_a_verdict_but_zero_is_left_alone(
        tmp_path, raw, status):
    """"191.2" joules is a test line the map landed on. "0" is an empty box
    copied by a formula — the owner's "not tested" — and stays as it is."""
    wb = Workbook()
    ws = wb.active
    _block(ws, 14, {"Model": "Rad-5", "S.N": "708246193",
                    "Manufacturer": "Masimo", "Location": "ICU"})
    ws["G27"] = raw
    path = tmp_path / "H12-AH007-0720.xlsx"
    wb.save(path)
    record, _ = calist.read_best(str(path), DEVICE_CONFIGS["AH"])
    assert record["Status"] == status


def test_a_computed_float_is_not_a_serial_but_a_long_real_serial_is():
    """An old Defibrillator opens on its test sheet, where the map read an
    uncertainty column. Real serials carry up to eight decimals."""
    assert calist.classify_serial("0.38271946501827364") == "suspect"
    for serial in ("1742.60318475", "437.20981", "6.20417", "1093.006172"):
        assert calist.classify_serial(serial) == "real", serial


# ── Typed-date repairs (the register only; the form is never changed) ─────────

@pytest.mark.parametrize("typed, name, repaired", [
    ("21--01-2024", "D22-AGH001-0124", "21-01-2024"),
    ("11-06-2023.", "D33-EY001-0623", "11-06-2023"),
    ("10-04-02025", "K184-FE001-0425", "10-04-2025"),
    ("08-12-205", "D02-BZ114-1225", "08-12-2025"),
    ("25-20-2023", "G152-BP004-1023", "25-10-2023"),
    ("0712-2025", "F33-AGH009-1225", "07-12-2025"),
    ("30/04/2023 - 18/06/2023", "K117-AM001-0823", "30-04-2023"),
    ("13-02-2024 to 11-03-2024", "K140-AM001-0224", "13-02-2024"),
])
def test_a_typed_date_slip_is_repaired_as_the_file_name_confirms(typed, name, repaired):
    assert calist.repair_date(typed, f"{name}.xlsx")[0] == repaired


@pytest.mark.parametrize("typed, name", [
    ("21-05-2024", "D22-AGH001-0124"),     # a date as typed is never touched
    ("21--03-2024", "D22-AGH001-0124"),    # plainly 21 March; the name says January
    ("12/19/2022", "A04-AM007-1222"),      # 19-12 month-first, or 12-12 one slip
    ("21--01-2024", "CN84017253"),         # no month in the name to confirm it
    ("=E14", "D38-AGH248-0224"),
])
def test_a_date_that_cannot_be_settled_is_left_as_typed(typed, name):
    assert calist.repair_date(typed, f"{name}.xlsx") is None


def test_a_repaired_date_reaches_the_register_with_what_was_typed(tmp_path):
    wb = Workbook()
    ws = wb.active
    _block(ws, 18, {"Model": "MX450", "S.N": "DE12345",
                    "Manufacturer": "Philips", "Location": "ICU"})
    ws["E16"] = "21--01-2024"
    path = tmp_path / "D22-AGH001-0124.xlsx"
    wb.save(path)
    record, _ = calist.read_best(str(path), DEVICE_CONFIGS["AGH"])
    assert record["Date"] == "21-01-2024"
    out = tmp_path / "out.xlsx"
    calist.write_output([record], _template(tmp_path), out)
    date = _cell(load_workbook(out).active, "Date")
    assert date.value == "21-01-2024"
    assert date.fill.fgColor.rgb.endswith("FFE699")
    assert "21--01-2024" in date.comment.text


def test_a_form_whose_block_moved_columns_takes_the_date_beside_its_caption(tmp_path):
    """The block sits in D/J, not E/K, so no uniform offset exists and the
    map's Date cell reads nothing. The form prints its date beside "Date of
    receipt:" — and only a non-date is ever replaced this way."""
    wb = Workbook()
    ws = wb.active
    _block(ws, 30, {"Model": "MX450", "S.N": "DE12345",
                    "Manufacturer": "Philips", "Location": "ICU"}, col="D", val="J")
    ws["A28"], ws["D28"] = "Date of receipt:", "14-03-2024"
    path = tmp_path / "D22-AGH001-0324.xlsx"
    wb.save(path)
    record, how = calist.read_best(str(path), DEVICE_CONFIGS["AGH"])
    assert how == "labels"
    assert record["Date"] == "14-03-2024"


# ── Repaired names in the register ────────────────────────────────────────────
#
# The files keep their names; the register's Code column carries the repaired
# one, amber with a note naming the real file whenever the repair changed what
# the name says.

def _build(files, tmp_path, **options):
    """A run with its template and register kept out of the forms' folder."""
    for folder in ("tpl", "out"):
        (tmp_path / folder).mkdir(exist_ok=True)
    return calist.process_files(files, _template(tmp_path / "tpl"),
                                output_dir=tmp_path / "out", **options)


def _defib(path, serial="SN-1", date="11-03-2026", model="R-Series"):
    """A Defibrillator (AC) form, filled in."""
    return _form(path, {"E17": "Zoll", "E15": model, "K15": serial,
                        "K17": "ER", "E13": date, "G24": "Pass"})


def _register_rows(result):
    sheet = load_workbook(result.output_path).active
    rows = []
    for row in range(calist.TEMPLATE_START_ROW, sheet.max_row + 1):
        code = sheet.cell(row=row, column=calist.TEMPLATE_START_COL
                          + calist.FIELDS.index("Code"))
        if code.value and not str(code.value).startswith("Generated by"):
            rows.append(code)
    return rows


def test_the_register_writes_the_repaired_name(tmp_path):
    form_path = _defib(tmp_path / ".G302-AC001-0326.xlsx")
    result = _build([form_path], tmp_path)
    [code] = _register_rows(result)
    assert code.value == "G302-AC001-0326"
    assert code.comment is None                 # a leading dot is not worth a note
    assert result.outcomes[0].code == "G302-AC001-0326"


def test_text_cut_from_the_name_is_marked_with_the_real_filename(tmp_path):
    form_path = _defib(tmp_path / "G302-AC001-0326-pending.xlsx")
    result = _build([form_path], tmp_path)
    [code] = _register_rows(result)
    assert code.value == "G302-AC001-0326"
    assert code.fill.fgColor.rgb.endswith("FFE699")
    assert "'-pending' after the date left out" in code.comment.text
    assert "G302-AC001-0326-pending.xlsx" in code.comment.text


def test_an_impossible_date_is_taken_from_the_form(tmp_path):
    """The owner's example: -0329 on a device calibrated 11-03-2026."""
    form_path = _defib(tmp_path / "G302-AC001-0329.xlsx", date="11-03-2026")
    result = _build([form_path], tmp_path)
    [code] = _register_rows(result)
    assert code.value == "G302-AC001-0326"
    assert "'0329'" in code.comment.text


def test_an_impossible_date_with_no_date_on_the_form_is_kept_and_marked(tmp_path):
    form_path = _defib(tmp_path / "G302-AC001-1525.xlsx", date="")
    result = _build([form_path], tmp_path)
    [code] = _register_rows(result)
    assert code.value == "G302-AC001-1525"
    assert "no date to take" in code.comment.text


def test_a_name_dated_in_the_future_repairs_no_form_date(tmp_path):
    """"H39-AG025-0233" is itself a slip; it must not bend a form's mistyped
    "01-02-20233" into 2033."""
    assert calist.repair_date("01-02-20233", "H39-AC025-0233.xlsx") is None
    assert calist.repair_date("01-02-20233", "H39-AC025-0223.xlsx") is not None


def test_a_copy_with_the_same_serial_is_left_out(tmp_path):
    original = _defib(tmp_path / "D38-AC090-0225.xlsx", serial="SN-9")
    copy = _defib(tmp_path / "D38-AC090-0225 (2).xlsx", serial="SN-9")
    result = _build([original, copy], tmp_path)
    assert [c.value for c in _register_rows(result)] == ["D38-AC090-0225"]
    [dropped] = result.copies
    assert dropped.filename == "D38-AC090-0225 (2).xlsx"
    assert "D38-AC090-0225.xlsx" in dropped.detail
    assert result.files_read == 1


def test_a_copy_with_no_serial_is_left_out_too(tmp_path):
    original = _defib(tmp_path / "D38-AC090-0225.xlsx", serial="SN-9")
    copy = _defib(tmp_path / "D38-AC090-0225 - Copy.xlsx", serial="N.A")
    result = _build([original, copy], tmp_path)
    assert len(_register_rows(result)) == 1
    assert [o.filename for o in result.copies] == ["D38-AC090-0225 - Copy.xlsx"]


def test_a_copy_with_a_different_serial_is_a_different_device(tmp_path):
    """"F21-BZ010-0624 (2)" is pump 14007615; the original is 14007302."""
    original = _defib(tmp_path / "F21-AC010-0624.xlsx", serial="14007302")
    other = _defib(tmp_path / "F21-AC010-0624 (2).xlsx", serial="14007615")
    result = _build([original, other], tmp_path)
    codes = _register_rows(result)
    assert [c.value for c in codes] == ["F21-AC010-0624", "F21-AC010-0624"]
    assert result.copies == []
    marked = [c for c in codes if c.comment is not None]
    assert len(marked) == 1
    assert "different serial" in marked[0].comment.text
    assert "F21-AC010-0624 (2).xlsx" in marked[0].comment.text


def test_a_lone_copy_is_simply_cleaned(tmp_path):
    copy = _defib(tmp_path / "D38-AC090-0225 (2).xlsx")
    result = _build([copy], tmp_path)
    [code] = _register_rows(result)
    assert code.value == "D38-AC090-0225"
    assert result.copies == []


def test_with_no_plain_original_the_extra_text_is_what_tells_them_apart(tmp_path):
    """"G341-AB00-0626-SEVO" and "-Iso" are two vaporizers."""
    sevo = _defib(tmp_path / "G341-AC001-0626-SEVO.xlsx", serial="A1")
    iso = _defib(tmp_path / "G341-AC001-0626-Iso.xlsx", serial="B2")
    result = _build([sevo, iso], tmp_path)
    codes = sorted(c.value for c in _register_rows(result))
    assert codes == ["G341-AC001-0626-Iso", "G341-AC001-0626-SEVO"]
    assert result.copies == []


def test_a_dropped_copy_takes_its_module_row_with_it(tmp_path):
    """A monitor's NIBP row belongs to the file; dropping the file drops both."""
    wb = Workbook()
    _block(wb.active, 18, {"Model": "MX450", "S.N": "DE1", "Manufacturer": "Philips",
                           "Location": "ICU"})
    wb.active["E16"] = "11-02-2026"
    wb.save(tmp_path / "D38-AGH001-0226.xlsx")
    wb.save(tmp_path / "D38-AGH001-0226 (Pending).xlsx")
    files = [str(tmp_path / "D38-AGH001-0226.xlsx"),
             str(tmp_path / "D38-AGH001-0226 (Pending).xlsx")]
    result = _build(files, tmp_path)
    assert result.rows_written == 2
    assert result.second_rows_added == 1
    assert [c.value for c in _register_rows(result)] == ["D38-AGH001-0226",
                                                         "D38-AGCB001-0226"]


# ── The device a form's own certificate names ─────────────────────────────────

def _vaporizer_named_by_brand(tmp_path):
    """"GE  TEC850  OR  SQAX00871": "GE" reads as the Temperature Calibration
    Tester's code, and the certificate says Vaporizer."""
    wb = Workbook()
    data = wb.active
    data.title = "Device data"
    _block(data, 18, {"Model": "Tec 850", "S.N": "SQAX00871",
                      "Manufacturer": "GE", "Location": "OR"})
    data["E16"] = "05-01-2025"
    _certificate(wb.create_sheet("Certificate"), "Vaporizer", row=40)
    path = tmp_path / "GE      TEC850     Main OR   SQAX00871.xlsx"
    wb.save(path)
    return str(path)


def test_read_best_hands_back_the_device_the_certificate_names(tmp_path):
    path = _vaporizer_named_by_brand(tmp_path)
    with pytest.raises(calist.ReadAsOther) as found:
        calist.read_best(path, DEVICE_CONFIGS["GE"])
    assert found.value.code == "AB"


def test_a_file_named_by_brand_is_read_as_its_certificates_device(tmp_path):
    path = _vaporizer_named_by_brand(tmp_path)
    result = _build([path], tmp_path)
    assert result.files_read == 1
    assert result.outcomes[0].device_code == "AB"
    sheet = load_workbook(result.output_path).active
    assert _cell(sheet, "Device").value == "Vaporizer"
    assert _cell(sheet, "S.N").value == "SQAX00871"
    code = _cell(sheet, "Code")
    assert code.value == Path(path).stem          # nothing invented
    assert "'Vaporizer'" in code.comment.text


def test_with_the_switch_on_a_file_named_by_brand_is_left_out(tmp_path):
    path = _vaporizer_named_by_brand(tmp_path)
    result = _build([path], tmp_path, real_forms_only=True)
    assert [o.status for o in result.outcomes] == [calist.LEFT_OUT]


def test_a_certificate_naming_several_devices_decides_nothing(tmp_path):
    """"Patient Monitor" is AG, AGH and VAGH: skipped, and the choices named."""
    wb = Workbook()
    data = wb.active
    data.title = "Device data"
    _block(data, 18, {"Model": "CM12", "S.N": "CN84017253",
                      "Manufacturer": "PHILIPS", "Location": "ICU"})
    _certificate(wb.create_sheet("Certificate"), "Patient Monitor")
    path = tmp_path / "CN84017253.xlsx"
    wb.save(path)
    with pytest.raises(calist.NotAForm, match="AGH") as refused:
        calist.read_best(str(path), DEVICE_CONFIGS["CN"])
    assert not isinstance(refused.value, calist.ReadAsOther)


# ── What the switch leaves out once a form is opened ──────────────────────────

def test_with_the_switch_on_a_blank_form_named_for_a_device_is_left_out(tmp_path):
    path = _blank_ultrasound(tmp_path / "G302-BB001-0425.xlsx")
    with pytest.raises(calist.LeftOut):
        calist.read_best(path, DEVICE_CONFIGS["BB"], forms_only=True)
    result = _build([path], tmp_path, real_forms_only=True)
    assert [o.status for o in result.outcomes] == [calist.LEFT_OUT]
    assert not result.problems


def test_with_the_switch_off_a_blank_named_form_is_read_as_always(tmp_path):
    path = _blank_ultrasound(tmp_path / "G302-BB001-0425.xlsx")
    result = _build([path], tmp_path)
    assert result.files_read == 1


def test_with_the_switch_on_a_device_list_is_left_out_not_a_problem(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "list"
    ws.append(["No.", "Device", "Manufacturer", "Model", "S.N", "Location", "Code"])
    ws.append([1, "Balance", "mediRhein", "MRCMT26", "N.A(1)", "Clinics", "G60-BP001-1223"])
    path = tmp_path / "G60-BP001-1223.xlsx"
    wb.save(path)
    result = _build([str(path)], tmp_path,
                                  real_forms_only=True)
    assert [o.status for o in result.outcomes] == [calist.LEFT_OUT]
    loose = _build([str(path)], tmp_path)
    assert [o.status for o in loose.outcomes] == [calist.UNSUPPORTED]


def _ge_named_by_brand(folder):
    """A GE machine named by brand, whose certificate says "Ventilator"."""
    folder.mkdir(parents=True)
    wb = Workbook()
    data = wb.active
    data.title = "Device data"
    _block(data, 18, {"Model": "Carestation 620", "S.N": "SM617300058MA",
                      "Manufacturer": "GE", "Location": "OR"})
    _certificate(wb.create_sheet("Certificate"), "Ventilator", row=40)
    path = folder / "GE        OR       SM617300058MA.xlsx"
    wb.save(path)
    return str(path)


def test_ventilator_on_an_anesthesia_machine_is_the_templates_not_the_devices(tmp_path):
    """680 Anesthesia forms say "Ventilator"; a Carestation filed under
    Anesthesia is skipped as it always was, not read as a ventilator."""
    path = _ge_named_by_brand(tmp_path / "Kom ombo" / "Anesthesia")
    with pytest.raises(calist.NotAForm, match="Ventilator") as refused:
        calist.read_best(path, DEVICE_CONFIGS["GE"])
    assert not isinstance(refused.value, calist.ReadAsOther)


def test_a_ventilator_filed_as_one_is_still_read_as_one(tmp_path):
    path = _ge_named_by_brand(tmp_path / "Kom ombo" / "Ventilator")
    with pytest.raises(calist.ReadAsOther) as found:
        calist.read_best(path, DEVICE_CONFIGS["GE"])
    assert found.value.code == "AM"
