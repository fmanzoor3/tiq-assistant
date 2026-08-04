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

        known_terms = self._collect_known_terms(projects)
        known_terms_block = (
            "\n".join(f"  - {t}" for t in known_terms) if known_terms else "  (none yet)"
        )

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

TRANSCRIPTION CORRECTION (important -- the summary came from speech-to-text):
  - The text may contain mis-heard acronyms/product names. Correct them to the
    KNOWN TERMS below (and the project list) when clearly intended, and fix any
    wrong hyphenation the transcriber added. e.g. "NGPT"/"and GPT" -> "EnGPT";
    "RAG-Deep research" -> "RAG deep research" (RAG is an acronym; do NOT glue
    it to the next word).
  - Only correct/expand terms you are confident about. NEVER invent a product
    name that isn't in the KNOWN TERMS or project list -- if unsure, keep the
    user's wording as-is.

KNOWN TERMS (correct spellings of products/areas the user works on):
{known_terms_block}

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

    def _build_user_prompt(
        self,
        transcript: str,
        remaining_hours: int,
        target_date: date,
        existing_entries: Optional[list] = None,
        meetings: Optional[list] = None,
        recent_context: Optional[list] = None,
    ) -> str:
        parts = [
            f"Date being logged: {target_date.strftime('%A, %d %B %Y')}.",
            f"Remaining target hours to fill for this day: {remaining_hours}.",
        ]

        if existing_entries:
            lines = "\n".join(
                f"  - {e.hours}h [{e.project_name or 'no project'}] {e.description}"
                for e in existing_entries
            )
            parts.append(
                "Entries ALREADY logged for this day (do NOT duplicate these; "
                "only fill the remaining hours):\n" + lines
            )

        if meetings:
            lines = "\n".join(
                f"  - {m.display_time} {m.subject} ({m.display_duration})"
                for m in meetings
            )
            parts.append(
                "Meetings on this day (already being added separately; you may "
                "reference them but do not re-create meeting entries):\n" + lines
            )

        if recent_context:
            lines = "\n".join(
                f"  - {d.strftime('%d %b')}: [{name or 'no project'}] {desc}"
                for (d, name, desc) in recent_context
            )
            parts.append(
                "For context, my entries over the previous few days (to keep "
                "phrasing and ongoing tasks consistent):\n" + lines
            )

        parts.append(
            f"My spoken summary of what I did on this day:\n"
            f"\"\"\"\n{transcript.strip()}\n\"\"\"\n"
        )
        parts.append("Produce the JSON entries now.")
        return "\n\n".join(parts)

    def gather_day_context(self, target_date: date, recent_days: int = 5) -> dict:
        """Collect context for gap-filling a given day.

        Returns dict with: existing_entries, meetings, recent_context
        (list of (date, project_name, description) from the previous N days).
        """
        from datetime import timedelta

        existing = self.store.get_entries(start_date=target_date, end_date=target_date)

        try:
            meetings = self.store.get_meetings_for_date(target_date)
        except Exception:
            meetings = []

        recent: list = []
        start = target_date - timedelta(days=recent_days)
        end = target_date - timedelta(days=1)
        if end >= start:
            for e in self.store.get_entries(start_date=start, end_date=end):
                recent.append((e.entry_date, e.project_name, e.description))
        # Keep the context compact.
        recent = recent[-15:]

        return {
            "existing_entries": existing,
            "meetings": meetings,
            "recent_context": recent,
        }

    # -------------------------------------------------------------- generate

    def generate(
        self,
        transcript: str,
        target_date: date,
        remaining_hours: int,
        settings: Optional[UserSettings] = None,
        existing_entries: Optional[list] = None,
        meetings: Optional[list] = None,
        recent_context: Optional[list] = None,
    ) -> GenerationResult:
        """Call the LLM and return proposed entries (unsaved).

        Optional context (existing entries, meetings, recent-days entries) helps
        the model fill only the gap and match ongoing phrasing.
        """
        if not transcript or not transcript.strip():
            raise LLMError("Nothing was said or typed to generate entries from.")

        settings = settings or self.store.get_settings()
        projects = self.store.get_projects(active_only=True)
        llm = self._llm or self._make_llm()

        system = self._build_system_prompt(projects, settings)
        user = self._build_user_prompt(
            transcript, max(1, remaining_hours), target_date,
            existing_entries=existing_entries,
            meetings=meetings,
            recent_context=recent_context,
        )

        messages = [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ]

        raw = llm.chat(
            messages=messages,
            temperature=0.0,
            json_mode=True,
            # Entries are short; cap output so the model can't ramble and slow
            # the response. A day's worth of entries fits comfortably in this.
            max_tokens=800,
        )

        try:
            parsed = self._parse_json(raw)
        except LLMError:
            # Retry once without JSON mode (some servers don't honour it), with a
            # blunt instruction to emit only the JSON object.
            retry_messages = messages + [
                {"role": "assistant", "content": raw},
                {"role": "user", "content":
                    'Return ONLY the JSON object {"entries":[...]} with no other '
                    'text, no markdown fences, and no reasoning.'},
            ]
            raw = llm.chat(
                messages=retry_messages,
                temperature=0.0,
                json_mode=False,
                max_tokens=800,
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

    def _collect_known_terms(self, projects: list[Project]) -> list[str]:
        """Gather correct spellings of products/areas the user works on.

        Sources: project names, project keywords, and the "Area:" prefixes the
        user has used in past descriptions (e.g. "Agentbot", "EnGPT").
        These become the LLM's correction targets for mis-heard speech.
        """
        terms: list[str] = []

        # User-supplied custom terms first (highest priority).
        try:
            custom = load_llm_config(self.store).custom_terms
            terms.extend(t.strip() for t in custom.split(",") if t.strip())
        except Exception:
            pass

        for p in projects:
            terms.append(p.name)
            terms.extend(p.keywords or [])

        # Prefixes from history: text before the first colon in past descriptions.
        try:
            for desc in self.store.get_description_history(limit=100):
                if ":" in desc:
                    prefix = desc.split(":", 1)[0].strip()
                    # Only short, name-like prefixes (avoid whole sentences).
                    if prefix and len(prefix) <= 30 and len(prefix.split()) <= 4:
                        terms.append(prefix)
        except Exception:
            pass

        # De-duplicate case-insensitively, preserve first-seen spelling.
        seen = set()
        uniq = []
        for t in terms:
            t = (t or "").strip()
            if t and t.lower() not in seen:
                seen.add(t.lower())
                uniq.append(t)
        return uniq[:50]

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
        """Extract entries JSON from model output, tolerating many shapes.

        Handles: plain JSON object, ```json fenced blocks, a leading <think>
        reasoning block, a bare JSON array (no {"entries": ...} wrapper), and
        extra prose around the JSON. Always returns a dict with an "entries" key.
        """
        original = raw or ""
        raw = original.strip()

        # Drop any reasoning block the server may have inlined.
        raw = re.sub(r"<think>.*?</think>", "", raw, flags=re.DOTALL | re.IGNORECASE).strip()
        # Strip ``` / ```json fences.
        raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.IGNORECASE | re.MULTILINE).strip()

        def _normalize(obj):
            # Accept {"entries": [...]}, a bare list, or a single entry dict.
            if isinstance(obj, dict):
                if "entries" in obj and isinstance(obj["entries"], list):
                    return obj
                # A single entry object -> wrap it.
                if any(k in obj for k in ("description", "project", "hours")):
                    return {"entries": [obj]}
                return {"entries": []}
            if isinstance(obj, list):
                return {"entries": obj}
            return None

        # 1) Try the whole (cleaned) string.
        for candidate in (raw,):
            try:
                norm = _normalize(json.loads(candidate))
                if norm is not None:
                    return norm
            except json.JSONDecodeError:
                pass

        # 2) Try the first {...} object block, then the first [...] array block.
        for pattern in (r"\{.*\}", r"\[.*\]"):
            match = re.search(pattern, raw, flags=re.DOTALL)
            if match:
                try:
                    norm = _normalize(json.loads(match.group(0)))
                    if norm is not None:
                        return norm
                except json.JSONDecodeError:
                    continue

        # Nothing parseable -- include a snippet of what came back for debugging.
        snippet = original.strip().replace("\n", " ")[:300] or "(empty response)"
        raise LLMError(
            "The AI did not return valid entries. It replied:\n\n" + snippet
        )

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
        timeout_seconds=int(get("llm_timeout", "120") or "120"),
        disable_thinking=(get("llm_disable_thinking", "1") == "1"),
        whisper_model=get("whisper_model", "base") or "base",
        custom_terms=get("custom_terms", ""),
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
    put("llm_disable_thinking", "1" if config.disable_thinking else "0")
    put("whisper_model", config.whisper_model)
    put("custom_terms", config.custom_terms)


def get_entry_generation_service() -> EntryGenerationService:
    return EntryGenerationService()
