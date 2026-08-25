"""Tests for settings persistence. Run with: python -m pytest

ui.py imports customtkinter, which needs a display, so these are skipped where
one is unavailable rather than failing the suite.
"""

import json

import pytest

ui = pytest.importorskip("ui", reason="needs a display and customtkinter")


@pytest.fixture
def settings_file(tmp_path, monkeypatch):
    path = tmp_path / "settings.json"
    monkeypatch.setattr(ui, "SETTINGS_FILE", path)
    return path


PREFS = {"unlocked_on": "2026-08-25", "deduplicate": True, "strict_names": True}


def test_reads_plain_utf8(settings_file):
    settings_file.write_text(json.dumps(PREFS), encoding="utf-8")
    assert ui.load_settings() == PREFS


def test_reads_utf8_with_a_bom(settings_file):
    """Notepad and Windows PowerShell both write a BOM.

    Reading as plain utf-8 makes json.loads raise, which load_settings swallows
    — silently wiping every saved preference and relocking the app.
    """
    settings_file.write_text(json.dumps(PREFS), encoding="utf-8-sig")
    assert settings_file.read_bytes()[:3] == b"\xef\xbb\xbf"
    assert ui.load_settings() == PREFS


def test_missing_file_gives_empty_settings(settings_file):
    assert not settings_file.exists()
    assert ui.load_settings() == {}


def test_corrupt_file_gives_empty_settings_instead_of_raising(settings_file):
    settings_file.write_text("{not json at all", encoding="utf-8")
    assert ui.load_settings() == {}


def test_round_trip(settings_file):
    ui.save_settings(PREFS)
    assert ui.load_settings() == PREFS


def test_save_never_raises_on_an_unwritable_path(tmp_path, monkeypatch):
    """A locked-down profile must not stop the app working."""
    monkeypatch.setattr(ui, "SETTINGS_FILE",
                        tmp_path / "nope\x00bad" / "settings.json")
    ui.save_settings(PREFS)          # must not raise


def test_saved_file_has_no_bom(settings_file):
    """What we write must be readable by anything, BOM-free."""
    ui.save_settings(PREFS)
    assert settings_file.read_bytes()[:3] != b"\xef\xbb\xbf"
