"""Daily PIN gate for Calist.

The PIN is derived from the local date, so nothing is stored and nothing has to
be distributed: the author can work out any day's code, tell a colleague, and
that code stops working when the date rolls over.

    base = day*17 + month*31 + year*11 + SECRET_KEY
    pin  = base squared, final digit dropped, then four digits, zero-padded

Dropping that final digit matters more than it looks. Taking the last four
digits of a square straight off only ever reaches 1044 of the 10000 possible
codes, because the last digits of a square cannot be arbitrary — and it left 38
codes in 2026 repeating on a second date. Shifting one digit widens the space to
5784 and halves the repeats.

Pure logic only — no GUI toolkit, no file I/O. Unlock state is passed in and out
as a plain dict so ui.py keeps ownership of persistence.

What this is and is not
-----------------------
This raises the bar for casual use by people who would otherwise just open the
app. It is not a licensing system: the formula lives in the source, and the
system clock is the only authority on what day it is. Anyone able to read the
code or set their clock can get in.
"""

from __future__ import annotations

import hmac
from datetime import date, datetime

#: Changing this invalidates every previously issued PIN.
SECRET_KEY = 8374

#: Wrong entries tolerated before a cooldown starts.
FREE_ATTEMPTS = 5

#: First cooldown, in seconds; doubles per failure after that.
COOLDOWN_BASE = 30

#: Ceiling on the cooldown, so a run of typos never bricks the whole day.
COOLDOWN_MAX = 15 * 60

_UNLOCKED_ON = "unlocked_on"
_FAILURES = "pin_failures"
_LOCKED_UNTIL = "pin_locked_until"


# ──────────────────────────────────────────────────────────────────────────────
# The PIN itself
# ──────────────────────────────────────────────────────────────────────────────

def daily_pin(on: date | None = None) -> str:
    """The four-digit code for a given date (today by default)."""
    day = on or date.today()
    base = day.day * 17 + day.month * 31 + day.year * 11 + SECRET_KEY
    # // 10 drops the square's final digit before the four are taken; see the
    # module docstring for why that is not cosmetic.
    return f"{(base * base // 10) % 10000:04d}"


def verify_pin(entered: str, on: date | None = None) -> bool:
    """True if ``entered`` is the code for that date.

    Compared with compare_digest so the check does not leak how much of the PIN
    was right through timing. Whitespace is forgiven; anything else is not.
    """
    if entered is None:
        return False
    return hmac.compare_digest(str(entered).strip(), daily_pin(on))


# ──────────────────────────────────────────────────────────────────────────────
# Unlock state
#
# Held in the same settings dict ui.py already persists. The stored value is the
# ISO date it was unlocked on, which gives the midnight reset for free: a date
# that is not today means locked, and so does a missing key. Deleting the
# settings file therefore locks the app rather than opening it.
# ──────────────────────────────────────────────────────────────────────────────

def is_unlocked_today(state: dict, today: date | None = None) -> bool:
    stamp = (state or {}).get(_UNLOCKED_ON)
    if not stamp:
        return False
    return stamp == (today or date.today()).isoformat()


def mark_unlocked(state: dict, today: date | None = None) -> dict:
    """Record a successful unlock and clear any failure history."""
    state = dict(state or {})
    state[_UNLOCKED_ON] = (today or date.today()).isoformat()
    state.pop(_FAILURES, None)
    state.pop(_LOCKED_UNTIL, None)
    return state


def clear_unlock(state: dict) -> dict:
    """Force a re-lock, leaving every other preference untouched."""
    state = dict(state or {})
    state.pop(_UNLOCKED_ON, None)
    return state


# ──────────────────────────────────────────────────────────────────────────────
# Rate limiting
#
# Only ~1044 four-digit values are reachable (the last four digits of a square
# cannot be arbitrary), so unlimited guessing would fall to a few hundred tries.
# The cooldown is persisted, so closing the window does not clear it.
# ──────────────────────────────────────────────────────────────────────────────

def record_failure(state: dict, now: datetime | None = None) -> dict:
    """Count a wrong entry and start or extend the cooldown."""
    state = dict(state or {})
    moment = now or datetime.now()
    failures = int(state.get(_FAILURES, 0)) + 1
    state[_FAILURES] = failures

    if failures > FREE_ATTEMPTS:
        step = failures - FREE_ATTEMPTS - 1
        wait = min(COOLDOWN_BASE * (2 ** step), COOLDOWN_MAX)
        state[_LOCKED_UNTIL] = moment.timestamp() + wait
    return state


def cooldown_remaining(state: dict, now: datetime | None = None) -> int:
    """Whole seconds still to wait before another attempt is accepted."""
    until = (state or {}).get(_LOCKED_UNTIL)
    if not until:
        return 0
    left = float(until) - (now or datetime.now()).timestamp()
    return max(0, int(left + 0.999))          # round up, so 0.2s reads as 1s


def attempts_left(state: dict) -> int:
    """Free attempts remaining before a cooldown kicks in."""
    return max(0, FREE_ATTEMPTS - int((state or {}).get(_FAILURES, 0)))
