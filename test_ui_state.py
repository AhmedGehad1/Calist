"""Tests for ui_state.py — what the window shows. Run with: python -m pytest

No display needed: ui_state decides, ui.py draws.
"""

import subprocess
import sys

from calist import (COPY, ERROR, LEFT_OUT, OK, READY, UNKNOWN_CODE, UNSUPPORTED,
                    FileOutcome)

import ui_state


def outcome(name, status=READY, **kw):
    return FileOutcome(filename=name, path=rf"C:\Round\{name}", status=status, **kw)


ROUND = [
    outcome("W01-AC004-0925.xlsx", OK, device_name="Defibrillator",
            device_code="AC", serial="DF-20481"),
    outcome("W01-AGH011-0925.xlsx", READY, device_name="Patient Monitor",
            device_code="AGH"),
    outcome("W01-QQQ100-0925.xlsx", UNKNOWN_CODE, device_code="QQQ"),
    outcome("W01-QQQ101-0925.xlsx", UNKNOWN_CODE, device_code="QQQ"),
    outcome("L02-ZZZ099-0925.xlsx", UNKNOWN_CODE, device_code="ZZZ"),
    outcome("L02-BB007-0925.docx", UNSUPPORTED),
    outcome("T03-AE012-0925.xlsx", ERROR, device_name="ESU",
            detail="the workbook is password-protected"),
    outcome("I05-AF007-0925 (2).xlsx", COPY, device_name="ECG"),
    outcome("Device List.xlsx", LEFT_OUT),
]


# ── Turbo ────────────────────────────────────────────────────────────────────

def test_turbo_starts_at_one_thousand_files():
    """The owner's threshold: 1,000 forms or more."""
    assert not ui_state.is_turbo(999)
    assert ui_state.is_turbo(1000)
    assert ui_state.is_turbo(20000)


def test_the_register_goes_to_the_first_folder_picked():
    """The owner's rule, in pick order — not alphabetical order."""
    picks = [(r"C:\Rounds\January round", True), (r"C:\Other", True)]
    assert ui_state.origin_folder(picks) == r"C:\Rounds\January round"


def test_or_beside_the_first_file_picked():
    picks = [(r"C:\Rounds\Ward\W01-AC004-0925.xlsx", False),
             (r"C:\Rounds\Lab\L02-BB006-0925.xlsx", False)]     # Lab sorts first
    assert ui_state.origin_folder(picks) == r"C:\Rounds\Ward"


def test_a_folder_picked_after_files_does_not_move_it():
    picks = [(r"C:\Rounds\Ward\a.xlsx", False), (r"C:\Rounds", True)]
    assert ui_state.origin_folder(picks) == r"C:\Rounds\Ward"
    assert ui_state.origin_folder([]) == ""


def test_the_round_is_named_for_the_folder_every_form_shares():
    assert ui_state.round_name([r"C:\Rounds\January round\Ward\a.xlsx",
                                r"C:\Rounds\January round\Lab\b.xlsx"]) == "January round"
    assert ui_state.round_name([r"C:\Rounds\January round\a.xlsx"]) == "January round"
    assert ui_state.round_name([r"C:\a.xlsx", r"D:\b.xlsx"]) == ""    # two drives
    assert ui_state.round_name([r"C:\a.xlsx"]) == ""                  # a drive root
    assert ui_state.round_name([]) == ""


# ── the problem banner ───────────────────────────────────────────────────────

def test_problems_are_grouped_by_kind_and_unknown_codes_by_code():
    groups = ui_state.problem_groups(ROUND)
    assert [(g.label, g.count) for g in groups] == [
        ("Unknown code QQQ", 2),        # the most common code first
        ("Unknown code ZZZ", 1),
        ("Not a form Calist reads", 1),
        ("Could not be read", 1),
    ]


def test_copies_and_left_out_files_are_not_problems():
    assert ui_state.problem_groups([outcome("a (2).xlsx", COPY),
                                    outcome("list.xlsx", LEFT_OUT)]) == []


def test_a_clean_round_has_no_groups():
    assert ui_state.problem_groups(ROUND[:2]) == []


def test_counts_leave_the_left_out_files_out_of_the_total():
    c = ui_state.count(ROUND)
    assert (c.total, c.ready, c.read, c.problems, c.copies, c.left_out) == \
        (8, 1, 1, 5, 1, 1)


# ── search and filters ───────────────────────────────────────────────────────

def test_search_matches_every_word_anywhere_in_the_row():
    defib = ROUND[0]
    assert ui_state.matches(defib, "defib")
    assert ui_state.matches(defib, "W01 df-204")          # two fields, any case
    assert ui_state.matches(defib, "  ")                   # blank finds all
    assert not ui_state.matches(defib, "defib monitor")    # every word must hit


def test_search_also_reads_the_status_the_table_shows():
    failed = ROUND[6]
    assert ui_state.matches(failed, "password")            # the detail
    assert ui_state.matches(failed, "failed", status_text="Failed")


def test_the_table_never_shows_left_out_files_and_sorts_by_name():
    rows = ui_state.visible([(o.path, o) for o in ROUND])
    names = [o.filename for _, o in rows]
    assert "Device List.xlsx" not in names
    assert names == sorted(names, key=str.lower)


def test_filters_combine():
    items = [(o.path, o) for o in ROUND]
    qqq = ui_state.problem_groups(ROUND)[0]
    assert len(ui_state.visible(items, problems_only=True)) == 5
    assert [o.filename for _, o in ui_state.visible(items, group=qqq)] == \
        ["W01-QQQ100-0925.xlsx", "W01-QQQ101-0925.xlsx"]
    assert [o.filename for _, o in ui_state.visible(items, group=qqq,
                                                    query="101")] == \
        ["W01-QQQ101-0925.xlsx"]


def test_importing_ui_state_pulls_in_no_gui_toolkit():
    code = "import ui_state, sys; assert 'tkinter' not in sys.modules"
    subprocess.run([sys.executable, "-c", code], check=True)
