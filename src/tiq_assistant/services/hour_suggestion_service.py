"""Service for smart hour suggestions in time entry popups."""

from datetime import date, datetime, time
from typing import Optional
from decimal import Decimal
import math

from tiq_assistant.core.models import ScheduleConfig, OutlookMeeting
from tiq_assistant.storage.sqlite_store import SQLiteStore, get_store
from tiq_assistant.desktop.windows.day_entry_dialog import SessionType


class HourSuggestionService:
    """
    Provides smart hour suggestions based on:
    - Session target hours (morning=3h, afternoon=5h by default)
    - Already logged entries for the session
    - Detected meetings duration
    - Previous usage patterns
    """

    def __init__(self, store: Optional[SQLiteStore] = None):
        self.store = store or get_store()

    def get_session_info(
        self,
        target_date: date,
        session: SessionType,
        config: Optional[ScheduleConfig] = None
    ) -> dict:
        """
        Get complete information about a session including hours logged,
        meetings detected, and remaining hours.

        Returns:
            dict with keys: target_hours, logged_hours, meeting_hours,
                          remaining_hours, entries, meetings
        """
        if config is None:
            config = self.store.get_schedule_config()

        # Get target hours for this session
        if session == SessionType.MORNING:
            target_hours = config.morning_hours_target
            session_start = self._parse_time(config.workday_start)
            session_end = self._parse_time(config.lunch_start)
        else:
            target_hours = config.afternoon_hours_target
            session_start = self._parse_time(config.lunch_end)
            session_end = self._parse_time(config.workday_end)

        # Get logged entries for this session
        entries = self.store.get_entries(start_date=target_date, end_date=target_date)
        session_entries = self._filter_entries_by_session(
            entries, session, session_start, session_end
        )
        logged_hours = sum(e.hours for e in session_entries)

        # Get meetings for this session
        meetings = self.store.get_meetings_for_date(target_date)
        session_meetings = self._filter_meetings_by_session(
            meetings, session_start, session_end
        )

        # Calculate meeting hours (not yet imported)
        meeting_hours = sum(
            self._round_hours(m.duration_hours)
            for m in session_meetings
            if not m.is_imported
        )

        # Calculate remaining
        remaining_hours = max(0, target_hours - logged_hours)

        return {
            "target_hours": target_hours,
            "logged_hours": logged_hours,
            "meeting_hours": meeting_hours,
            "remaining_hours": remaining_hours,
            "entries": session_entries,
            "meetings": session_meetings,
            "session_start": session_start,
            "session_end": session_end,
        }

    def suggest_hours(
        self,
        target_date: date,
        session: SessionType,
        config: Optional[ScheduleConfig] = None
    ) -> int:
        """
        Suggest the number of hours for a new entry.

        Strategy:
        1. Calculate remaining hours to fill the target
        2. Account for any detected meetings (not yet imported)
        3. Return at least 1 hour

        Args:
            target_date: The date for the entry
            session: Morning or afternoon session
            config: Optional schedule config (fetched if not provided)

        Returns:
            Suggested hours (integer, minimum 1)
        """
        info = self.get_session_info(target_date, session, config)

        # Remaining hours minus pending meeting hours
        available = info["remaining_hours"] - info["meeting_hours"]

        # Return at least 1, at most the remaining hours
        return max(1, min(available, info["remaining_hours"]))

    def get_day_summary(
        self,
        target_date: date,
        config: Optional[ScheduleConfig] = None
    ) -> dict:
        """
        Get a complete summary of the day's time entries.

        Returns:
            dict with morning_info, afternoon_info, total_hours, total_target
        """
        if config is None:
            config = self.store.get_schedule_config()

        morning_info = self.get_session_info(target_date, SessionType.MORNING, config)
        afternoon_info = self.get_session_info(target_date, SessionType.AFTERNOON, config)

        total_hours = morning_info["logged_hours"] + afternoon_info["logged_hours"]
        total_target = morning_info["target_hours"] + afternoon_info["target_hours"]

        return {
            "date": target_date,
            "morning": morning_info,
            "afternoon": afternoon_info,
            "total_hours": total_hours,
            "total_target": total_target,
            "is_complete": total_hours >= total_target,
        }

    def _parse_time(self, time_str: str) -> time:
        """Parse a time string like '09:30' into a time object."""
        parts = time_str.split(":")
        return time(int(parts[0]), int(parts[1]))

    def _filter_entries_by_session(
        self,
        entries: list,
        session: SessionType,
        session_start: time,
        session_end: time
    ) -> list:
        """Return only the entries that belong to ``session``.

        Timesheet entries don't store a time of day, so we attribute them as
        follows:

        - Calendar-sourced entries: use the linked meeting's start time when it
          can be found, so they land in the correct session.
        - Manual entries (and calendar entries whose meeting can't be resolved):
          attributed to the *afternoon* session, which is the point at which the
          full day is normally reconciled. This guarantees every entry is
          counted in exactly one session, so morning + afternoon never
          double-counts (the previous implementation returned all entries for
          both sessions, inflating logged hours).
        """
        from tiq_assistant.core.models import EntrySource

        result = []
        for entry in entries:
            entry_session = self._infer_entry_session(
                entry, session_start, session_end
            )
            if entry_session == session:
                result.append(entry)
        return result

    def _infer_entry_session(
        self,
        entry,
        morning_start: time,
        morning_end: time,
    ) -> SessionType:
        """Best-effort attribution of an entry to a single session."""
        from tiq_assistant.core.models import EntrySource

        # Try to place calendar entries by their source meeting's start time.
        source_event_id = getattr(entry, "source_event_id", None)
        if getattr(entry, "source", None) == EntrySource.CALENDAR and source_event_id:
            meeting = self._find_meeting(entry.entry_date, source_event_id)
            if meeting is not None:
                lunch_start = self._parse_time(
                    self.store.get_schedule_config().lunch_start
                )
                if meeting.start_datetime.time() < lunch_start:
                    return SessionType.MORNING
                return SessionType.AFTERNOON

        # Manual / unresolved entries default to the afternoon (reconciliation).
        return SessionType.AFTERNOON

    def _find_meeting(self, target_date: date, meeting_id: str):
        """Find a cached meeting by id on a given date, or None."""
        try:
            for m in self.store.get_meetings_for_date(target_date):
                if m.id == meeting_id:
                    return m
        except Exception:
            pass
        return None

    def _filter_meetings_by_session(
        self,
        meetings: list[OutlookMeeting],
        session_start: time,
        session_end: time
    ) -> list[OutlookMeeting]:
        """Filter meetings that fall within the session time range."""
        filtered = []
        for meeting in meetings:
            meeting_time = meeting.start_datetime.time()
            if session_start <= meeting_time < session_end:
                filtered.append(meeting)
        return filtered

    def _round_hours(self, decimal_hours: Decimal) -> int:
        """Round decimal hours to nearest integer (minimum 1)."""
        return max(1, round(float(decimal_hours)))


def get_hour_suggestion_service() -> HourSuggestionService:
    """Get an hour suggestion service instance."""
    return HourSuggestionService()
