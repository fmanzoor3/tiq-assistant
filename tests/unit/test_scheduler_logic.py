"""Regression tests for the scheduler's fire/dedup/suppression logic.

These guard the reliability fixes that make lunch / end-of-day popups actually
appear even after laptop sleep, and that suppress reminders on non-working days.
"""

import os
from datetime import date

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

# Skip the whole module if PyQt6 isn't installed in this environment.
pytest.importorskip("PyQt6")

from PyQt6.QtWidgets import QApplication  # noqa: E402

from tiq_assistant.core.models import SessionType, ScheduleConfig  # noqa: E402
from tiq_assistant.desktop.scheduler import SchedulerManager  # noqa: E402


@pytest.fixture(scope="module")
def qapp():
    app = QApplication.instance() or QApplication([])
    yield app


@pytest.fixture
def scheduler(qapp):
    sm = SchedulerManager()
    sm._config = ScheduleConfig()
    return sm


def _collect(scheduler):
    fired = []
    scheduler.morning_popup_due.connect(lambda: fired.append("M"))
    scheduler.afternoon_popup_due.connect(lambda: fired.append("A"))
    return fired


def test_fire_is_deduped_within_a_day(qapp, scheduler):
    # This test targets dedup, not the working-day suppression (which has its
    # own tests). Force suppression off so it passes on any real calendar date.
    scheduler._should_fire_today = lambda _today: True
    fired = _collect(scheduler)
    scheduler._fire(SessionType.MORNING)
    scheduler._fire(SessionType.MORNING)  # second call must be ignored
    qapp.processEvents()
    assert fired == ["M"]


def test_weekend_is_suppressed(scheduler):
    # 2026-08-01 is a Saturday.
    assert scheduler._should_fire_today(date(2026, 8, 1)) is False


def test_full_day_holiday_is_suppressed(scheduler):
    # 2026-01-01 is a full-day national holiday.
    assert scheduler._should_fire_today(date(2026, 1, 1)) is False


def test_normal_weekday_fires(scheduler):
    # 2026-07-27 is a Monday, not a holiday.
    assert scheduler._should_fire_today(date(2026, 7, 27)) is True


def test_snooze_bypasses_daily_dedup(qapp, scheduler):
    scheduler._should_fire_today = lambda _today: True  # ignore weekend/holiday
    fired = _collect(scheduler)
    # Mark afternoon as already shown today.
    scheduler._fire(SessionType.AFTERNOON)
    qapp.processEvents()
    # A snooze re-fire must still emit even though it was "shown".
    scheduler._snooze_fire_afternoon()
    qapp.processEvents()
    assert fired.count("A") == 2


def test_shown_set_prunes_previous_days(scheduler):
    scheduler._shown.add((date(2020, 1, 1), SessionType.MORNING))
    scheduler._prune_shown(date(2026, 7, 27))
    assert (date(2020, 1, 1), SessionType.MORNING) not in scheduler._shown
