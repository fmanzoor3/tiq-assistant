"""Service for smart hour suggestions in time entry popups."""

from datetime import date, datetime, time
from typing import Optional
from decimal import Decimal
import math

from tiq_assistant.core.models import SessionType, ScheduleConfig, OutlookMeeting
from tiq_assistant.storage.sqlite_store import SQLiteStore, get_store


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
        """
        Filter entries by session. Since entries don't have time,
        we use the source to determine session if from calendar,
        otherwise we distribute based on creation time.
        """
        # For now, return all entries for the date and let the UI handle it
        # In a more sophisticated implementation, we could track session per entry
        return entries

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
