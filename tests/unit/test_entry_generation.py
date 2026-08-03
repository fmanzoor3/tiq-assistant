"""Tests for LLM-driven entry generation (project mapping, phrasing, parsing)."""

from datetime import date

import pytest

from tiq_assistant.storage.sqlite_store import SQLiteStore
from tiq_assistant.core.models import Project
from tiq_assistant.services.entry_generation_service import (
    EntryGenerationService, load_llm_config, save_llm_config,
)
from tiq_assistant.integrations.llm_client import LLMClient, LLMError


class _FakeLLM(LLMClient):
    def __init__(self, payload):
        self._payload = payload

    def chat(self, messages, **kwargs):
        self.last_system = messages[0]["content"]
        return self._payload


@pytest.fixture
def store(tmp_path):
    s = SQLiteStore(db_path=tmp_path / "t.db")
    default = Project(name="YAPAY ZEKA SUPPORT", ticket_number="9000", keywords=[])
    agentbot = Project(name="Agentbot", ticket_number="2019135", keywords=["agentbot"])
    opt = Project(name="Optimization Project", ticket_number="3030", keywords=["optimization"])
    for p in (default, agentbot, opt):
        s.save_project(p)
    us = s.get_settings()
    us.default_project_id = default.id
    s.save_settings(us)
    return s


def _svc(store, payload):
    return EntryGenerationService(store=store, llm=_FakeLLM(payload))


def test_named_projects_map_to_tickets(store):
    payload = (
        '{"entries":['
        '{"project":"Agentbot","hours":4,"description":"Fixed feedback"},'
        '{"project":"Optimization Project","hours":4,"description":"Tuned params"}]}'
    )
    res = _svc(store, payload).generate("...", date(2026, 8, 3), 8)
    assert res.entries[0].project_name == "Agentbot"
    assert res.entries[0].ticket_number == "2019135"
    assert res.entries[1].project_name == "Optimization Project"


def test_null_project_falls_back_to_default(store):
    payload = '{"entries":[{"project":null,"hours":8,"description":"Agentbot: feedback"}]}'
    res = _svc(store, payload).generate("...", date(2026, 8, 3), 8)
    assert res.entries[0].project_name == "YAPAY ZEKA SUPPORT"
    assert res.entries[0].description.startswith("Agentbot:")


def test_unknown_project_falls_back_to_default(store):
    payload = '{"entries":[{"project":"Nonexistent","hours":3,"description":"x"}]}'
    res = _svc(store, payload).generate("...", date(2026, 8, 3), 3)
    assert res.entries[0].project_name == "YAPAY ZEKA SUPPORT"


def test_parses_json_inside_code_fence(store):
    payload = 'Here:\n```json\n{"entries":[{"project":"Agentbot","hours":2,"description":"y"}]}\n```'
    res = _svc(store, payload).generate("...", date(2026, 8, 3), 2)
    assert res.entries[0].project_name == "Agentbot"
    assert res.entries[0].hours == 2


def test_invalid_json_raises(store):
    with pytest.raises(LLMError):
        _svc(store, "no json here at all").generate("...", date(2026, 8, 3), 3)


def test_empty_transcript_raises(store):
    with pytest.raises(LLMError):
        _svc(store, "{}").generate("   ", date(2026, 8, 3), 3)


def test_config_round_trip(store):
    cfg = load_llm_config(store)
    assert cfg.enabled is False  # off by default
    cfg.enabled = True
    cfg.base_url = "https://example/v1"
    cfg.model = "qwen-test"
    save_llm_config(cfg, store)
    reloaded = load_llm_config(store)
    assert reloaded.enabled is True
    assert reloaded.base_url == "https://example/v1"
    assert reloaded.model == "qwen-test"


def test_hours_are_integers_min_one(store):
    payload = '{"entries":[{"project":null,"hours":0,"description":"x"},{"project":null,"hours":2.6,"description":"y"}]}'
    res = _svc(store, payload).generate("...", date(2026, 8, 3), 3)
    assert res.entries[0].hours == 1     # 0 -> min 1
    assert res.entries[1].hours == 3     # 2.6 -> rounded


def test_day_context_is_included_in_prompt(store):
    """Existing entries, meetings and recent days should reach the LLM prompt."""
    from datetime import timedelta
    from tiq_assistant.core.models import (
        TimesheetEntry, ActivityCode, EntryStatus, EntrySource,
    )

    target = date(2026, 5, 20)
    store.save_entry(TimesheetEntry(
        consultant_id="F", entry_date=target, hours=3, project_name="Agentbot",
        ticket_number="2019135", activity_code=ActivityCode.GLST, location="ANKARA",
        description="Agentbot: triage", status=EntryStatus.DRAFT, source=EntrySource.MANUAL,
    ))
    store.save_entry(TimesheetEntry(
        consultant_id="F", entry_date=target - timedelta(days=1), hours=8,
        project_name="Agentbot", ticket_number="2019135", activity_code=ActivityCode.GLST,
        location="ANKARA", description="Agentbot: feature work",
        status=EntryStatus.DRAFT, source=EntrySource.MANUAL,
    ))

    svc = EntryGenerationService(store=store)
    ctx = svc.gather_day_context(target)
    assert len(ctx["existing_entries"]) == 1
    assert len(ctx["recent_context"]) == 1

    fake = _FakeLLM('{"entries":[{"project":"Agentbot","hours":5,"description":"Agentbot: more work"}]}')
    svc2 = EntryGenerationService(store=store, llm=fake)
    svc2.generate(
        "did more work", target, 5,
        existing_entries=ctx["existing_entries"],
        meetings=ctx["meetings"],
        recent_context=ctx["recent_context"],
    )
    # The user prompt (2nd message) is built by generate(); confirm it carried context.
    # _FakeLLM stores the system prompt; rebuild user prompt to assert structure.
    user_prompt = svc2._build_user_prompt(
        "did more work", 5, target,
        existing_entries=ctx["existing_entries"],
        meetings=ctx["meetings"],
        recent_context=ctx["recent_context"],
    )
    assert "ALREADY logged" in user_prompt
    assert "previous few days" in user_prompt
    assert "20 May 2026" in user_prompt
