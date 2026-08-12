"""Regression: adding a meeting must not create duplicates across reopens."""

import os
from datetime import date, datetime

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
pytest.importorskip("PyQt6")

from PyQt6.QtWidgets import QApplication  # noqa: E402

import tiq_assistant.storage.sqlite_store as store_mod  # noqa: E402
from tiq_assistant.storage.sqlite_store import SQLiteStore  # noqa: E402
from tiq_assistant.core.models import Project, OutlookMeeting, SessionType  # noqa: E402


@pytest.fixture(scope="module")
def qapp():
    return QApplication.instance() or QApplication([])


@pytest.fixture
def store(tmp_path, monkeypatch):
    s = SQLiteStore(db_path=tmp_path / "t.db")
    monkeypatch.setattr(store_mod, "_store", s)
    return s


def _meeting():
    return OutlookMeeting(
        subject="Standup",
        start_datetime=datetime(2026, 8, 4, 10, 0),
        end_datetime=datetime(2026, 8, 4, 10, 30),
        match_confidence=1.0,
    )


def test_meeting_not_added_twice_across_reopens(qapp, store):
    from tiq_assistant.desktop.windows.day_entry_dialog import DayEntryDialog

    p = Project(name="Proj", ticket_number="1", keywords=[])
    store.save_project(p)
    m = _meeting()
    m.matched_project_id = p.id
    target = date(2026, 8, 4)

    def count():
        return len(store.get_entries(start_date=target, end_date=target))

    # First open: add it.
    d1 = DayEntryDialog(target_date=target, session=SessionType.FULL_DAY,
                        outlook_meetings=[m])
    d1._add_single_meeting(0)
    assert count() == 1

    # Reopen: the row should be disabled and re-adding must be a no-op.
    d2 = DayEntryDialog(target_date=target, session=SessionType.FULL_DAY,
                        outlook_meetings=[m])
    assert d2._meetings_table.cellWidget(0, 6).text() == "Added"
    assert not d2._meetings_table.cellWidget(0, 6).isEnabled()
    d2._add_single_meeting(0)  # forced programmatic add
    assert count() == 1  # still one, no duplicate
