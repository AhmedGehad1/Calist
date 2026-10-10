"""The window's layout passes, without a window. Run with: python -m pytest

On first open the page used to scroll until the window was resized: two layout
passes ran in the same tick, both measured the page before Tk had applied the
other's change, and both grew the empty card by the whole gap — then both
shrank it — until the pass budget ran out with the page 219 px too tall.

These drive App's own layout methods on a stand-in that behaves like Tk where
it matters: a new height is applied only when idle work runs, which happens
between timer callbacks due at different moments, or on update_idletasks.

ui.py imports customtkinter, which needs a display, so these are skipped where
one is unavailable rather than failing the suite (as test_settings.py does).
"""

import pytest

ui = pytest.importorskip("ui", reason="needs a display and customtkinter")

SCALE = 1.25            # the development display; logical → device pixels
VIEWPORT = 669          # the page's height in the 1000x600 window, measured
CHROME = 75             # header and paddings: 450 px of page at a 300 card


@pytest.fixture(autouse=True)
def _scaling(monkeypatch):
    monkeypatch.setattr(ui.ctk.ScalingTracker, "get_widget_scaling",
                        lambda widget: SCALE)


class _Bar:
    def __init__(self):
        self.shown = False

    def winfo_ismapped(self):
        return self.shown

    def grid(self):
        self.shown = True

    def grid_remove(self):
        self.shown = False


class _Hero:
    """The empty-state card: configured now, laid out at idle."""

    def __init__(self, height):
        self.height = self.applied = height

    def cget(self, _option):
        return self.height

    def configure(self, height):
        self.height = height


class _Canvas:
    def __init__(self, height):
        self.height = height

    def winfo_height(self):
        return self.height


class _Page:
    def __init__(self, hero, viewport):
        self._hero = hero
        self._parent_canvas = _Canvas(viewport)
        self._scrollbar = _Bar()

    def winfo_ismapped(self):
        return True

    def winfo_reqheight(self):
        return int(CHROME + self._hero.applied * SCALE)


class _Window:
    """Just enough of App for its layout methods, with a fake event loop."""

    MAX_FIT_PASSES = ui.App.MAX_FIT_PASSES
    MIN_HERO_H, MAX_HERO_H = ui.App.MIN_HERO_H, ui.App.MAX_HERO_H
    ROW_PX = 26

    _on_resize = ui.App._on_resize
    _schedule_fit = ui.App._schedule_fit
    _fit_to_window = ui.App._fit_to_window
    _show_page_scrollbar = ui.App._show_page_scrollbar

    def __init__(self, viewport=VIEWPORT, hero=300):
        self._lock = None
        self._files = {}
        self._fit_job = None
        self._fit_passes = 0
        self._hero = _Hero(hero)
        self._page = _Page(self._hero, viewport)
        self.now = 0
        self.timers = []                # (due, order, job, callback)
        self.order = 0
        self.passes_run = 0

    # Tk's timers and idle work
    def after(self, delay, callback):
        self.order += 1
        job = f"after#{self.order}"
        self.timers.append((self.now + delay, self.order, job, callback))
        return job

    def after_cancel(self, job):
        self.timers = [t for t in self.timers if t[2] != job]

    def update_idletasks(self):
        self._hero.applied = self._hero.height

    def run(self, limit=100):
        """Fire timers in order; idle work runs whenever time moves on."""
        while self.timers and limit:
            limit -= 1
            self.timers.sort()
            due, _, _, callback = self.timers.pop(0)
            if due > self.now:
                self.update_idletasks()
                self.now = due
            if callback == self._fit_to_window:
                self.passes_run += 1
            callback()
        self.update_idletasks()

    def slack(self):
        return self._page._parent_canvas.height - self._page.winfo_reqheight()


def test_startup_and_unlock_leave_one_pass_pending():
    window = _Window()
    window._schedule_fit(120, fresh=True)        # _build
    window._on_resize()                          # _refresh_intake
    window._on_resize()                          # the unlock
    assert len(window.timers) == 1


def test_the_empty_window_settles_to_fit_without_scrolling():
    window = _Window()
    window._schedule_fit(120, fresh=True)
    window._on_resize()
    window.run()
    assert 0 <= window.slack() < 10
    assert not window._page._scrollbar.shown
    assert window.passes_run <= 2


def test_two_passes_in_one_tick_cannot_overshoot():
    """What 2.0.2 did at startup: a second pass in the same tick. Each now
    measures after the other's change is applied, so the card is sized once."""
    window = _Window()
    window._fit_to_window()
    window._fit_to_window()
    window.run()
    assert 0 <= window.slack() < 10
    assert not window._page._scrollbar.shown


def test_out_of_passes_the_scrollbar_still_matches_the_page():
    """The bar used to keep the state of the last pass that ran."""
    window = _Window(hero=600)                   # 75 + 750 px: overflows
    window._fit_passes = window.MAX_FIT_PASSES
    window._fit_to_window()
    assert window._page._scrollbar.shown
    assert not window.timers                     # out of budget: no new pass

    window._hero.configure(height=400)           # something else made it fit
    window._fit_to_window()
    assert not window._page._scrollbar.shown


def test_a_resize_still_refits():
    window = _Window()
    window.run()
    window._page._parent_canvas.height = 900     # dragged taller
    window._on_resize()
    window.run()
    assert 0 <= window.slack() < 10
