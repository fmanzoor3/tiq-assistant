"""Tests for holiday add/edit store methods (manual editing support)."""

import os
from datetime import date

import pytest

from tiq_assistant.storage.sqlite_store import SQLiteStore
from tiq_assistant.core.exceptions import StorageError


@pytest.fixture
def store(tmp_path):
    return SQLiteStore(db_path=tmp_path / "t.db")


def test_add_and_get(store):
    store.add_holiday(date(2026, 5, 1), "Labour Day", "full_day")
    holidays = store.get_holidays()
    assert len(holidays) == 1
    assert holidays[0]["holiday_date"] == date(2026, 5, 1)
    assert holidays[0]["holiday_type"] == "full_day"


def test_add_duplicate_date_raises(store):
    store.add_holiday(date(2026, 5, 1), "A", "full_day")
    with pytest.raises(StorageError):
        store.add_holiday(date(2026, 5, 1), "B", "half_day")


def test_update_changes_date_and_type(store):
    store.add_holiday(date(2026, 3, 19), "Arife", "half_day")
    hid = store.get_holidays()[0]["id"]
    store.update_holiday(hid, date(2026, 3, 20), "Bayram", "full_day")
    h = store.get_holidays()[0]
    assert h["holiday_date"] == date(2026, 3, 20)
    assert h["name"] == "Bayram"
    assert h["holiday_type"] == "full_day"


def test_update_into_existing_date_raises(store):
    store.add_holiday(date(2026, 5, 1), "A", "full_day")
    store.add_holiday(date(2026, 5, 2), "B", "full_day")
    b_id = [h["id"] for h in store.get_holidays() if h["name"] == "B"][0]
    with pytest.raises(StorageError):
        store.update_holiday(b_id, date(2026, 5, 1), "B", "full_day")


def test_seed_only_when_empty(store):
    assert store.seed_default_holidays_if_empty() > 0
    # Second call must not add anything (no re-adding deleted defaults).
    assert store.seed_default_holidays_if_empty() == 0


def test_deleting_a_default_holiday_sticks(tmp_path, monkeypatch):
    """Regression: deleting a built-in half-day must make it a normal workday.

    HolidayService reads the global store, so point that at a temp DB.
    """
    import tiq_assistant.storage.sqlite_store as store_mod
    from tiq_assistant.core.holidays import HolidayService

    s = SQLiteStore(db_path=tmp_path / "hol.db")
    monkeypatch.setattr(store_mod, "_store", s)

    s.seed_default_holidays_if_empty()
    svc = HolidayService()  # reads the (now temp) global DB
    assert svc.get_expected_hours(date(2026, 5, 25)) == 4  # half-day initially

    h = [x for x in s.get_holidays() if x["holiday_date"] == date(2026, 5, 25)][0]
    s.delete_holiday(h["id"])
    svc.reload_from_database()

    assert not svc.is_holiday(date(2026, 5, 25))
    assert svc.get_expected_hours(date(2026, 5, 25)) == 8  # now a full workday

    # And it must not reappear when a fresh service loads.
    assert not HolidayService().is_holiday(date(2026, 5, 25))
