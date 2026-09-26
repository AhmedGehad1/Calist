"""What the window shows, decided without a window.

The rules behind the device table, the problem banner and Turbo live here as
plain functions over FileOutcome values, so they are tested like the pipeline
is — with no display — and ui.py only draws what these decide. Imports calist
(GUI-free) and nothing from tkinter.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field

from calist import (COPY, ERROR, LEFT_OUT, OK, READY, UNKNOWN_CODE, UNSUPPORTED,
                    FileOutcome)

#: From this many files on, the window switches to Turbo by itself: no
#: per-file table, problems-only logging. The owner's threshold — the top of a
#: normal round (100–1,000), so an older Windows 10 laptop never has to draw a
#: table it would struggle with.
TURBO_AT = 1000


def is_turbo(count: int) -> bool:
    """Whether a run of `count` files goes without the per-file table."""
    return count >= TURBO_AT


def origin_folder(picks) -> str:
    """Where the register goes when nobody chose a folder.

    `picks` are (path, is_folder) pairs in the order the user added them. The
    owner's rule: the first folder picked, or — when files were picked one by
    one — the folder of the first of those. Never the alphabetically first
    form: a round with Lab and Ward subfolders used to land its register in
    Lab, because "Lab" sorts before "Ward".
    """
    for path, is_folder in picks:
        return path if is_folder else os.path.dirname(path)
    return ""


def round_name(paths) -> str:
    """The round's folder: the deepest folder every form sits under.

    Forms in `January round\\Ward` and `January round\\Lab` are the January
    round. Forms from two drives, or straight off a drive's root, have no
    shared folder worth naming.
    """
    folders = {os.path.dirname(p) for p in paths}
    if not folders:
        return ""
    try:
        common = os.path.commonpath(sorted(folders))
    except ValueError:                                  # different drives
        return ""
    return os.path.basename(common)


# ── the problem banner ───────────────────────────────────────────────────────

@dataclass(frozen=True)
class Group:
    """One kind of problem, and the files it covers.

    `key` is what the table filters on; `label` is what the banner says. An
    unknown code is grouped by the code, because fourteen files named for a
    device nobody has added are one finding, not fourteen.
    """

    key: tuple[str, str]
    label: str
    paths: tuple[str, ...] = field(default=())

    @property
    def count(self) -> int:
        return len(self.paths)


def problem_groups(outcomes) -> list[Group]:
    """The round's problems, most common unknown code first, then the rest."""
    codes: dict[str, list[str]] = {}
    unsupported: list[str] = []
    failed: list[str] = []
    for outcome in outcomes:
        if outcome.status == UNKNOWN_CODE:
            codes.setdefault(outcome.device_code or "?", []).append(outcome.path)
        elif outcome.status == UNSUPPORTED:
            unsupported.append(outcome.path)
        elif outcome.status == ERROR:
            failed.append(outcome.path)

    groups = [Group(("code", code), f"Unknown code {code}", tuple(paths))
              for code, paths in sorted(codes.items(),
                                        key=lambda kv: (-len(kv[1]), kv[0]))]
    if unsupported:
        groups.append(Group(("unsupported", ""), "Not a form Calist reads",
                            tuple(unsupported)))
    if failed:
        groups.append(Group(("failed", ""), "Could not be read", tuple(failed)))
    return groups


@dataclass(frozen=True)
class Counts:
    """The round at a glance. `left_out` files are never in `total`."""

    total: int = 0
    ready: int = 0
    read: int = 0
    problems: int = 0
    copies: int = 0
    left_out: int = 0


def count(outcomes) -> Counts:
    tally = {"ready": 0, "read": 0, "problems": 0, "copies": 0, "left_out": 0}
    total = 0
    for outcome in outcomes:
        if outcome.status == LEFT_OUT:
            tally["left_out"] += 1
            continue
        total += 1
        if outcome.status == READY:
            tally["ready"] += 1
        elif outcome.status == OK:
            tally["read"] += 1
        elif outcome.status == COPY:
            tally["copies"] += 1
        if outcome.is_problem:
            tally["problems"] += 1
    return Counts(total=total, **tally)


# ── what the table shows ─────────────────────────────────────────────────────

def matches(outcome: FileOutcome, query: str, status_text: str = "") -> bool:
    """Whether a search finds this file.

    Every word must appear somewhere in the row — file name, device, serial,
    the register's Code, the device code or the status — so "icu defib" finds
    the ICU's defibrillators. Case and extra spaces do not matter.
    """
    words = query.lower().split()
    if not words:
        return True
    haystack = " ".join(filter(None, (
        outcome.filename, outcome.device_name, outcome.serial, outcome.code,
        outcome.device_code, outcome.detail, status_text,
    ))).lower()
    return all(word in haystack for word in words)


def visible(items, *, problems_only: bool = False, group: Group | None = None,
            query: str = "", status_text=lambda outcome: ""):
    """The rows the table draws, in its order.

    `items` are (path, outcome) pairs. A left-out file is never drawn — the
    user asked not to see them, and Details lists them with the reason.
    """
    wanted = set(group.paths) if group is not None else None
    rows = []
    for path, outcome in items:
        if outcome.status == LEFT_OUT:
            continue
        if problems_only and not outcome.is_problem:
            continue
        if wanted is not None and path not in wanted:
            continue
        if query and not matches(outcome, query, status_text(outcome)):
            continue
        rows.append((path, outcome))
    rows.sort(key=lambda kv: kv[1].filename.lower())
    return rows
