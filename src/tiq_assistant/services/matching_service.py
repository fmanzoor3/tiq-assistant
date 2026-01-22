"""Service for matching calendar events to projects."""

import re
from typing import Optional

from tiq_assistant.core.models import CalendarEvent, Project, MatchResult
from tiq_assistant.storage.sqlite_store import SQLiteStore, get_store


class MatchingService:
    """Service for matching calendar events to projects."""

    # Pattern to match JIRA-style keys (e.g., PEMP-948, TIQ-123)
    JIRA_KEY_PATTERN = re.compile(r'\b([A-Z]{2,10}-\d{1,6})\b')

    # Pattern to extract ticket from JIRA URLs
    JIRA_URL_PATTERN = re.compile(r'/browse/([A-Z]{2,10}-\d{1,6})')

    def __init__(self, store: Optional[SQLiteStore] = None):
        self.store = store or get_store()

    def match_event(self, event: CalendarEvent) -> MatchResult:
        """
        Match a calendar event to a project.

        Matching strategy (in order of priority):
        1. Extract JIRA key from subject and look up project
        2. Extract JIRA key from description URL and look up
        3. Match keywords from subject against project keywords
        """
        # Strategy 1: Extract JIRA key from subject
        jira_keys = self._extract_jira_keys(event.subject)
        for key in jira_keys:
            project = self.store.find_project_by_jira_key(key)
            if project:
                return MatchResult(
                    project_id=project.id,
                    project_name=project.name,
                    ticket_numeric_id=project.ticket_number,
                    ticket_jira_key=project.jira_key,
                    confidence=1.0,
                    match_source="subject",
                    matched_text=key,
                )

        # Strategy 2: Extract JIRA key from description URL
        if event.description:
            url_keys = self._extract_jira_keys_from_urls(event.description)
            for key in url_keys:
                project = self.store.find_project_by_jira_key(key)
                if project:
                    return MatchResult(
                        project_id=project.id,
                        project_name=project.name,
                        ticket_numeric_id=project.ticket_number,
                        ticket_jira_key=project.jira_key,
                        confidence=1.0,
                        match_source="description_url",
                        matched_text=key,
                    )

            # Also check plain JIRA keys in description
            desc_keys = self._extract_jira_keys(event.description)
            for key in desc_keys:
                if key not in jira_keys:  # Skip already checked
                    project = self.store.find_project_by_jira_key(key)
                    if project:
                        return MatchResult(
                            project_id=project.id,
                            project_name=project.name,
                            ticket_numeric_id=project.ticket_number,
                            ticket_jira_key=project.jira_key,
                            confidence=0.9,
                            match_source="description",
                            matched_text=key,
                        )

        # Strategy 3: Keyword matching
        result = self._match_by_keywords(event.subject)
        if result and result.confidence > 0:
            return result

        # No match found
        return MatchResult(confidence=0.0, match_source="none")

    def _extract_jira_keys(self, text: str) -> list[str]:
        """Extract JIRA-style keys from text."""
        if not text:
            return []
        return self.JIRA_KEY_PATTERN.findall(text.upper())

    def _extract_jira_keys_from_urls(self, text: str) -> list[str]:
        """Extract JIRA keys from URLs in text."""
        if not text:
            return []
        return self.JIRA_URL_PATTERN.findall(text.upper())

    def _match_by_keywords(self, text: str) -> Optional[MatchResult]:
        """Match text against project keywords."""
        if not text:
            return None

        text_lower = text.lower()
        projects = self.store.get_projects(active_only=True)

        best_match: Optional[MatchResult] = None
        best_score = 0.0

        for project in projects:
            for keyword in project.keywords:
                if keyword.lower() in text_lower:
                    score = len(keyword) / len(text)  # Longer matches score higher
                    if score > best_score:
                        best_score = score
                        best_match = MatchResult(
                            project_id=project.id,
                            project_name=project.name,
                            ticket_numeric_id=project.ticket_number,
                            ticket_jira_key=project.jira_key,
                            confidence=min(0.8, score + 0.3),  # Cap at 0.8 for keyword match
                            match_source="keyword",
                            matched_text=keyword,
                        )

        return best_match

    def match_events(self, events: list[CalendarEvent]) -> list[CalendarEvent]:
        """Match multiple events and update them with results."""
        for event in events:
            result = self.match_event(event)
            event.matched_project_id = result.project_id
            event.matched_jira_key = result.ticket_jira_key
            event.match_confidence = result.confidence
            event.match_source = result.match_source
        return events

    def get_unmatched_jira_keys(self, events: list[CalendarEvent]) -> list[str]:
        """Get list of JIRA keys found in events but not in database."""
        unmatched = set()

        for event in events:
            # Extract all JIRA keys
            keys = self._extract_jira_keys(event.subject)
            if event.description:
                keys.extend(self._extract_jira_keys(event.description))
                keys.extend(self._extract_jira_keys_from_urls(event.description))

            for key in keys:
                # Check if in database
                if not self.store.find_project_by_jira_key(key):
                    unmatched.add(key)

        return sorted(unmatched)


def get_matching_service() -> MatchingService:
    """Get a matching service instance."""
    return MatchingService()
