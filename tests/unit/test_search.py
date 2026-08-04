"""Tests for entry search (relevance ranking + term matching)."""

from datetime import date

import pytest

from tiq_assistant.storage.sqlite_store import SQLiteStore
from tiq_assistant.core.models import (
    TimesheetEntry, ActivityCode, EntryStatus, EntrySource,
)


@pytest.fixture
def store(tmp_path):
    return SQLiteStore(db_path=tmp_path / "t.db")


def _add(store, d, desc, project="YAPAY ZEKA SUPPORT", ticket="9000"):
    store.save_entry(TimesheetEntry(
        consultant_id="F", entry_date=d, hours=2, project_name=project,
        ticket_number=ticket, activity_code=ActivityCode.GLST, location="ANKARA",
        description=desc, status=EntryStatus.DRAFT, source=EntrySource.MANUAL,
    ))


def test_requires_all_terms(store):
    _add(store, date(2026, 8, 1), "Claude design discussion")
    _add(store, date(2026, 8, 2), "Claude API work")          # no "design"
    _add(store, date(2026, 8, 3), "UI design tweaks")          # no "claude"
    results = store.search_entries("claude design")
    assert len(results) == 1
    assert results[0].description == "Claude design discussion"


def test_exact_phrase_ranks_first(store):
    _add(store, date(2026, 7, 15), "Reviewed Claude API design docs")  # split
    _add(store, date(2026, 8, 1), "Claude design discussion")          # phrase
    results = store.search_entries("claude design")
    assert results[0].description == "Claude design discussion"


def test_matches_project_and_ticket(store):
    _add(store, date(2026, 8, 1), "General work", project="Agentbot", ticket="2019135")
    assert len(store.search_entries("Agentbot")) == 1
    assert len(store.search_entries("2019135")) == 1


def test_empty_query_returns_recent(store):
    for i in range(3):
        _add(store, date(2026, 8, 1 + i), f"entry {i}")
    results = store.search_entries("")
    assert len(results) == 3


def test_no_match_returns_empty(store):
    _add(store, date(2026, 8, 1), "something else")
    assert store.search_entries("nonexistent term xyz") == []
