"""Tests for the daily PIN gate. Run with: python -m pytest

All pure — no display, no files, no clock dependence beyond dates passed in.
"""

from datetime import date, datetime, timedelta

import pytest

import access


# ── the formula ───────────────────────────────────────────────────────────────

@pytest.mark.parametrize("when, expected", [
    (date(2026, 8, 25), "5688"),
    (date(2026, 8, 26), "2250"),
    (date(2026, 8, 27), "8868"),
    (date(2026, 8, 28), "5545"),
    (date(2026, 1, 6), "0884"),       # leading zero must survive
    (date(2026, 1, 10), "0132"),
])
def test_known_pins(when, expected):
    assert access.daily_pin(when) == expected


def test_pin_is_always_four_characters():
    """A short result must be padded, not printed bare (542 -> '0542')."""
    day = date(2025, 1, 1)
    for _ in range(1200):
        pin = access.daily_pin(day)
        assert len(pin) == 4 and pin.isdigit(), (day, pin)
        day += timedelta(days=1)


def test_pin_matches_the_specified_arithmetic():
    when = date(2026, 8, 25)
    base = when.day * 17 + when.month * 31 + when.year * 11 + access.SECRET_KEY
    assert access.daily_pin(when) == f"{(base * base // 10) % 10000:04d}"


def test_the_shifted_digit_is_what_widens_the_keyspace():
    """Guards the // 10 — dropping it silently shrinks the code space by 5.5x.

    Only the last four digits of a square can be reached without the shift, so
    losing it would quietly make guessing five times easier.
    """
    shifted = {(n * n // 10) % 10000 for n in range(200_000)}
    unshifted = {(n * n) % 10000 for n in range(200_000)}
    assert len(shifted) == 5784
    assert len(unshifted) == 1044
    assert len(shifted) > len(unshifted) * 5


def test_codes_rarely_repeat_within_a_year():
    """38 dates collided before the shift; this pins the improvement."""
    from collections import Counter
    day, counts = date(2026, 1, 1), Counter()
    while day <= date(2026, 12, 31):
        counts[access.daily_pin(day)] += 1
        day += timedelta(days=1)
    repeated = [pin for pin, n in counts.items() if n > 1]
    assert len(repeated) == 19


def test_consecutive_days_differ():
    day = date(2025, 1, 1)
    for _ in range(800):
        assert access.daily_pin(day) != access.daily_pin(day + timedelta(days=1))
        day += timedelta(days=1)


# ── verification ──────────────────────────────────────────────────────────────

def test_correct_pin_is_accepted():
    when = date(2026, 8, 25)
    assert access.verify_pin(access.daily_pin(when), when)


def test_surrounding_whitespace_is_forgiven():
    when = date(2026, 8, 25)
    assert access.verify_pin(f"  {access.daily_pin(when)}  ", when)


@pytest.mark.parametrize("entered", ["6888", "689", "68899", "", "abcd", "    ", None])
def test_wrong_or_malformed_input_is_rejected(entered):
    assert not access.verify_pin(entered, date(2026, 8, 25))


def test_yesterdays_pin_does_not_work_today():
    today = date(2026, 8, 25)
    assert not access.verify_pin(access.daily_pin(today - timedelta(days=1)), today)


# ── unlock state ──────────────────────────────────────────────────────────────

def test_a_fresh_state_is_locked():
    assert not access.is_unlocked_today({}, date(2026, 8, 25))


def test_missing_state_is_locked_not_open():
    """Deleting the settings file must lock the app, never unlock it."""
    assert not access.is_unlocked_today(None, date(2026, 8, 25))


def test_unlocking_holds_for_the_rest_of_that_day():
    today = date(2026, 8, 25)
    state = access.mark_unlocked({}, today)
    assert access.is_unlocked_today(state, today)


def test_crossing_midnight_relocks():
    today = date(2026, 8, 25)
    state = access.mark_unlocked({}, today)
    assert not access.is_unlocked_today(state, today + timedelta(days=1))


def test_unlocking_preserves_other_settings():
    state = access.mark_unlocked({"template": r"C:\x.xlsx", "deduplicate": True},
                                 date(2026, 8, 25))
    assert state["template"] == r"C:\x.xlsx"
    assert state["deduplicate"] is True


def test_clear_unlock_relocks_without_touching_preferences():
    state = access.mark_unlocked({"template": "t"}, date(2026, 8, 25))
    state = access.clear_unlock(state)
    assert not access.is_unlocked_today(state, date(2026, 8, 25))
    assert state["template"] == "t"


# ── rate limiting ─────────────────────────────────────────────────────────────

def test_early_failures_do_not_trigger_a_cooldown():
    now = datetime(2026, 8, 25, 9, 0, 0)
    state = {}
    for _ in range(access.FREE_ATTEMPTS):
        state = access.record_failure(state, now)
    assert access.cooldown_remaining(state, now) == 0


def test_cooldown_starts_after_the_free_attempts():
    now = datetime(2026, 8, 25, 9, 0, 0)
    state = {}
    for _ in range(access.FREE_ATTEMPTS + 1):
        state = access.record_failure(state, now)
    assert access.cooldown_remaining(state, now) == access.COOLDOWN_BASE


def test_cooldown_doubles_then_stops_growing():
    now = datetime(2026, 8, 25, 9, 0, 0)
    state = {}
    for _ in range(access.FREE_ATTEMPTS):
        state = access.record_failure(state, now)

    waits = []
    for _ in range(12):
        state = access.record_failure(state, now)
        waits.append(access.cooldown_remaining(state, now))

    assert waits[0] == access.COOLDOWN_BASE
    assert waits[1] == access.COOLDOWN_BASE * 2
    assert waits[2] == access.COOLDOWN_BASE * 4
    assert max(waits) == access.COOLDOWN_MAX          # capped, never unbounded
    assert waits == sorted(waits)                     # never goes backwards


def test_cooldown_expires_with_time():
    now = datetime(2026, 8, 25, 9, 0, 0)
    state = {}
    for _ in range(access.FREE_ATTEMPTS + 1):
        state = access.record_failure(state, now)
    later = now + timedelta(seconds=access.COOLDOWN_BASE + 1)
    assert access.cooldown_remaining(state, later) == 0


def test_a_successful_unlock_clears_the_penalty():
    now = datetime(2026, 8, 25, 9, 0, 0)
    state = {}
    for _ in range(access.FREE_ATTEMPTS + 3):
        state = access.record_failure(state, now)
    assert access.cooldown_remaining(state, now) > 0

    state = access.mark_unlocked(state, now.date())
    assert access.cooldown_remaining(state, now) == 0
    assert access.attempts_left(state) == access.FREE_ATTEMPTS


def test_attempts_left_counts_down():
    now = datetime(2026, 8, 25, 9, 0, 0)
    state = {}
    assert access.attempts_left(state) == access.FREE_ATTEMPTS
    state = access.record_failure(state, now)
    assert access.attempts_left(state) == access.FREE_ATTEMPTS - 1


def test_state_is_json_safe():
    """It rides in settings.json, so everything stored must serialise."""
    import json
    now = datetime(2026, 8, 25, 9, 0, 0)
    state = access.record_failure(access.mark_unlocked({}, now.date()), now)
    assert json.loads(json.dumps(state)) == state


# ── the gate imports no GUI ───────────────────────────────────────────────────

def test_access_pulls_in_no_toolkit():
    import subprocess, sys
    out = subprocess.run(
        [sys.executable, "-c",
         "import access, sys; "
         "assert 'tkinter' not in sys.modules; "
         "assert 'customtkinter' not in sys.modules; print('clean')"],
        capture_output=True, text=True)
    assert out.returncode == 0, out.stderr
