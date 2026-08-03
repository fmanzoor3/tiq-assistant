"""Turn a spoken/typed description of the day into draft timesheet entries.

Given free text like "I worked on Agentbot fixing feedback for a few hours and
spent the rest on optimization", this service asks the local LLM to produce
structured entries that:

- assign the correct project (defaulting to YAPAY ZEKA SUPPORT when the user
  doesn't name a real project, or explicitly says "under the default project"),
- follow the user's phrasing conventions (short, task-focused, and the
  "Area: task" prefix for support work booked under the default project),
- estimate integer hours per task, scaled to fill the day's remaining target.

The LLM only ever selects from the user's real project list and returns JSON;
all field mapping (ticket number, activity code, location) is done in Python so
the model can't invent codes. Everything is returned as unsaved drafts for the
user to review in the top-up dialog.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from datetime import date
from typing import Optional

from tiq_assistant.core.models import (
    TimesheetEntry, Project, ActivityCode, EntryStatus, EntrySource, UserSettings,
)
from tiq_assistant.storage.sqlite_store import SQLiteStore, get_store
from tiq_assistant.integrations.llm_client import LLMClient, LLMConfig, LLMError


DEFAULT_PROJECT_NAME = "YAPAY ZEKA SUPPORT"


@dataclass
class GeneratedEntry:
    """One proposed entry before it becomes a TimesheetEntry."""
    project_id: Optional[str]
    project_name: Optional[str]
    ticket_number: Optional[str]
    hours: int
    description: str
    activity_code: ActivityCode = ActivityCode.GLST
    is_meeting: bool = False


@dataclass
class GenerationResult:
    entries: list[GeneratedEntry] = field(default_factory=list)
    raw_response: str = ""
    notes: str = ""


class EntryGenerationService:
    def __init__(
        self,
        store: Optional[SQLiteStore] = None,
        llm: Optional[LLMClient] = None,
    ):
        self.store = store or get_store()
        self._llm = llm

    # --------------------------------------------------------------- prompt

    def _build_system_prompt(self, projects: list[Project], settings: UserSettings) -> str:
        # Few-shot examples of the user's real phrasing, so the model mimics it.
        examples = self.store.get_description_history(limit=6)
        examples_block = "\n".join(f'  - "{e}"' for e in examples) if examples else "  (none yet)"

        # The project list the model may choose from.
        proj_lines = []
        for p in projects:
            kw = f" (keywords: {', '.join(p.keywords)})" if p.keywords else ""
            proj_lines.append(f'  - "{p.name}" [ticket {p.ticket_number}]{kw}')
        proj_block = "\n".join(proj_lines) if proj_lines else "  (no projects defined)"

        default_name = self._default_project_name(projects, settings)

        return f"""You convert a person's spoken summary of their workday into timesheet entries.

You MUST return ONLY a JSON object of this exact shape, nothing else:
{{
  "entries": [
    {{"project": "<exact project name from the list, or null>",
      "hours": <integer>,
      "description": "<short entry text>"}}
  ]
}}

PROJECTS the user may reference (choose the project name EXACTLY as written, or null):
{proj_block}

The DEFAULT project is "{default_name}". Use it when:
  - the user does not name a specific project, OR
  - the user says the work was "under the default project" / "support", OR
  - the work is support for another area but booked under the default project.

DESCRIPTION rules (match the user's style exactly):
  - One short sentence, usually fewer than 10 words. Task-focused, not verbose.
  - For support work booked under the DEFAULT project that relates to a specific
    area/product, PREFIX the description with that area and a colon.
    Example: "Agentbot: addressing business team feedback".
  - If the user names a real project AND the work is booked to that project,
    do NOT add a prefix -- just the task.

HOURS rules:
  - Estimate integer hours per task from what the user says
    ("a few hours" ~= 3, "a couple" ~= 2, "most of the day" ~= 5, "quick" ~= 1).
  - The hours across all entries should sum to about the day's remaining target
    hours given by the user message. Adjust proportionally; keep each >= 1.

Examples of the user's past descriptions (mimic this tone and length):
{examples_block}

Return ONLY the JSON object."""

    def _build_user_prompt(self, transcript: str, remaining_hours: int) -> str:
        return (
            f"Remaining target hours to fill today: {remaining_hours}.\n\n"
            f"My spoken summary of the day:\n\"\"\"\n{transcript.strip()}\n\"\"\"\n\n"
            f"Produce the JSON entries now."
        )

    # -------------------------------------------------------------- generate

    def generate(
        self,
        transcript: str,
        target_date: date,
        remaining_hours: int,
        settings: Optional[UserSettings] = None,
    ) -> GenerationResult:
        """Call the LLM and return proposed entries (unsaved)."""
        if not transcript or not transcript.strip():
            raise LLMError("Nothing was said or typed to generate entries from.")

        settings = settings or self.store.get_settings()
        projects = self.store.get_projects(active_only=True)
        llm = self._llm or self._make_llm()

        system = self._build_system_prompt(projects, settings)
        user = self._build_user_prompt(transcript, max(1, remaining_hours))

        raw = llm.chat(
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
            temperature=0.0,
            json_mode=True,
        )

        parsed = self._parse_json(raw)
        entries = self._to_generated_entries(parsed, projects, settings)
        return GenerationResult(entries=entries, raw_response=raw)

    # ----------------------------------------------------------- conversion

    def _to_generated_entries(
        self,
        parsed: dict,
        projects: list[Project],
        settings: UserSettings,
    ) -> list[GeneratedEntry]:
        by_name = {p.name.strip().lower(): p for p in projects}
        default_project = self._default_project(projects, settings)

        out: list[GeneratedEntry] = []
        for raw_entry in parsed.get("entries", []):
            name = (raw_entry.get("project") or "").strip()
            description = (raw_entry.get("description") or "").strip()
            try:
                hours = int(round(float(raw_entry.get("hours", 1))))
            except (TypeError, ValueError):
                hours = 1
            hours = max(1, hours)

            if not description:
                continue

            # Resolve the project. null / unknown -> default project.
            project = by_name.get(name.lower()) if name else None
            if project is None:
                project = default_project

            out.append(GeneratedEntry(
                project_id=project.id if project else None,
                project_name=project.name if project else None,
                ticket_number=project.ticket_number if project else None,
                hours=hours,
                description=description,
                activity_code=settings.default_activity_code,
            ))
        return out

    # -------------------------------------------------------------- helpers

    def _default_project(
        self, projects: list[Project], settings: UserSettings
    ) -> Optional[Project]:
        # Prefer the user's configured default project; else match by name.
        if settings.default_project_id:
            for p in projects:
                if p.id == settings.default_project_id:
                    return p
        for p in projects:
            if p.name.strip().upper() == DEFAULT_PROJECT_NAME:
                return p
        return None

    def _default_project_name(
        self, projects: list[Project], settings: UserSettings
    ) -> str:
        p = self._default_project(projects, settings)
        return p.name if p else DEFAULT_PROJECT_NAME

    @staticmethod
    def _parse_json(raw: str) -> dict:
        """Extract the JSON object from the model output, tolerating extra text."""
        raw = raw.strip()
        # Strip ```json fences if present.
        raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.IGNORECASE | re.MULTILINE)
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            # Fall back to the first {...} block.
            match = re.search(r"\{.*\}", raw, flags=re.DOTALL)
            if match:
                try:
                    return json.loads(match.group(0))
                except json.JSONDecodeError:
                    pass
        raise LLMError("The LLM did not return valid JSON entries.")

    def _make_llm(self) -> LLMClient:
        cfg = load_llm_config(self.store)
        if not cfg.enabled:
            raise LLMError("The AI assistant is disabled. Enable it in Settings.")
        return LLMClient(cfg)


# ---------------------------------------------------------------- config I/O

def load_llm_config(store: Optional[SQLiteStore] = None) -> LLMConfig:
    """Load LLM settings from the key/value app_config table."""
    store = store or get_store()
    get = store.get_config_value
    return LLMConfig(
        enabled=(get("llm_enabled", "0") == "1"),
        base_url=get("llm_base_url", LLMConfig.base_url),
        api_key=get("llm_api_key", LLMConfig.api_key),
        model=get("llm_model", ""),
        verify_ssl=(get("llm_verify_ssl", "0") == "1"),
        timeout_seconds=int(get("llm_timeout", "60") or "60"),
    )


def save_llm_config(config: LLMConfig, store: Optional[SQLiteStore] = None) -> None:
    store = store or get_store()
    put = store.set_config_value
    put("llm_enabled", "1" if config.enabled else "0")
    put("llm_base_url", config.base_url)
    put("llm_api_key", config.api_key)
    put("llm_model", config.model)
    put("llm_verify_ssl", "1" if config.verify_ssl else "0")
    put("llm_timeout", str(config.timeout_seconds))


def get_entry_generation_service() -> EntryGenerationService:
    return EntryGenerationService()
