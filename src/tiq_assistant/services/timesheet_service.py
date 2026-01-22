"""Service for managing timesheet entries."""

from datetime import date
from decimal import Decimal
from typing import Optional
import math

from tiq_assistant.core.models import (
    CalendarEvent, TimesheetEntry, UserSettings,
    ActivityCode, EntryStatus, EntrySource
)
from tiq_assistant.storage.sqlite_store import SQLiteStore, get_store
from tiq_assistant.services.matching_service import MatchingService, get_matching_service


class TimesheetService:
    """Service for creating and managing timesheet entries."""

    def __init__(
        self,
        store: Optional[SQLiteStore] = None,
        matching_service: Optional[MatchingService] = None,
    ):
        self.store = store or get_store()
        self.matching_service = matching_service or get_matching_service()

    def generate_entries_from_events(
        self,
        events: list[CalendarEvent],
        settings: Optional[UserSettings] = None,
    ) -> list[TimesheetEntry]:
        """Generate timesheet entries from calendar events."""
        if settings is None:
            settings = self.store.get_settings()

        # Match events to projects/tickets
        self.matching_service.match_events(events)

        entries = []
        for event in events:
            # Skip if below minimum duration
            min_hours = settings.min_meeting_duration_minutes / 60
            if float(event.duration_hours) < min_hours:
                continue

            entry = self._create_entry_from_event(event, settings)
            entries.append(entry)

        return entries

    def _create_entry_from_event(
        self,
        event: CalendarEvent,
        settings: UserSettings,
    ) -> TimesheetEntry:
        """Create a timesheet entry from a calendar event."""
        # Round hours to nearest integer (as per timesheet format)
        hours = self._round_hours(float(event.duration_hours))

        # Get project/ticket info
        project_name = None
        ticket_number = None
        if event.matched_project_id:
            project = self.store.get_project(event.matched_project_id)
            if project:
                project_name = project.name
                # Find ticket
                if event.matched_ticket_id:
                    ticket = project.find_ticket_by_jira_key(event.matched_jira_key or "")
                    if ticket:
                        ticket_number = ticket.numeric_id

        # Determine activity code (meetings default to TPLNT)
        activity_code = settings.meeting_activity_code

        return TimesheetEntry(
            consultant_id=settings.consultant_id,
            entry_date=event.start_date,
            hours=hours,
            ticket_number=ticket_number,
            project_name=project_name,
            activity_code=activity_code,
            location=settings.default_location,
            description=event.to_timesheet_description(),
            status=EntryStatus.DRAFT,
            source=EntrySource.CALENDAR,
            source_event_id=event.id,
            source_jira_key=event.matched_jira_key,
        )

    def _round_hours(self, hours: float) -> int:
        """Round hours to nearest integer (minimum 1)."""
        rounded = round(hours)
        return max(1, rounded)

    def create_manual_entry(
        self,
        entry_date: date,
        hours: int,
        description: str,
        project_name: Optional[str] = None,
        ticket_number: Optional[str] = None,
        activity_code: Optional[ActivityCode] = None,
        location: Optional[str] = None,
        settings: Optional[UserSettings] = None,
    ) -> TimesheetEntry:
        """Create a manual timesheet entry."""
        if settings is None:
            settings = self.store.get_settings()

        return TimesheetEntry(
            consultant_id=settings.consultant_id,
            entry_date=entry_date,
            hours=hours,
            ticket_number=ticket_number,
            project_name=project_name,
            activity_code=activity_code or settings.default_activity_code,
            location=location or settings.default_location,
            description=description,
            status=EntryStatus.DRAFT,
            source=EntrySource.MANUAL,
        )

    def save_entry(self, entry: TimesheetEntry) -> TimesheetEntry:
        """Save a timesheet entry."""
        return self.store.save_entry(entry)

    def save_entries(self, entries: list[TimesheetEntry]) -> list[TimesheetEntry]:
        """Save multiple timesheet entries."""
        return [self.store.save_entry(entry) for entry in entries]

    def get_entries(
        self,
        start_date: Optional[date] = None,
        end_date: Optional[date] = None,
        status: Optional[EntryStatus] = None,
    ) -> list[TimesheetEntry]:
        """Get timesheet entries with optional filters."""
        return self.store.get_entries(start_date, end_date, status)

    def get_entry(self, entry_id: str) -> Optional[TimesheetEntry]:
        """Get a single entry by ID."""
        return self.store.get_entry(entry_id)

    def update_entry(self, entry: TimesheetEntry) -> TimesheetEntry:
        """Update a timesheet entry."""
        return self.store.save_entry(entry)

    def delete_entry(self, entry_id: str) -> None:
        """Delete a timesheet entry."""
        self.store.delete_entry(entry_id)

    def approve_entries(self, entry_ids: list[str]) -> None:
        """Mark entries as approved."""
        for entry_id in entry_ids:
            entry = self.store.get_entry(entry_id)
            if entry:
                entry.status = EntryStatus.APPROVED
                self.store.save_entry(entry)

    def get_entries_for_export(
        self,
        start_date: Optional[date] = None,
        end_date: Optional[date] = None,
    ) -> list[TimesheetEntry]:
        """Get entries ready for export (approved or draft)."""
        entries = self.store.get_entries(start_date, end_date)
        # Filter to only draft or approved entries
        return [
            e for e in entries
            if e.status in (EntryStatus.DRAFT, EntryStatus.APPROVED)
        ]

    def get_daily_summary(self, target_date: date) -> dict:
        """Get summary of hours for a specific date."""
        entries = self.store.get_entries(start_date=target_date, end_date=target_date)
        total_hours = sum(e.hours for e in entries)

        return {
            "date": target_date,
            "entries": entries,
            "total_hours": total_hours,
            "entry_count": len(entries),
        }

    def aggregate_entries(self, entries: list[TimesheetEntry]) -> list[TimesheetEntry]:
        """
        Aggregate entries by date + project + ticket.
        Combines hours for similar entries on the same day.
        """
        aggregated = {}

        for entry in entries:
            # Create a key for aggregation
            key = (
                entry.entry_date,
                entry.project_name or "",
                entry.ticket_number or "",
                entry.activity_code,
            )

            if key in aggregated:
                # Add hours to existing entry
                existing = aggregated[key]
                existing.hours += entry.hours
                # Append descriptions if different
                if entry.description not in existing.description:
                    existing.description = f"{existing.description}; {entry.description}"
            else:
                # Create new aggregated entry
                aggregated[key] = TimesheetEntry(
                    consultant_id=entry.consultant_id,
                    entry_date=entry.entry_date,
                    hours=entry.hours,
                    ticket_number=entry.ticket_number,
                    project_name=entry.project_name,
                    activity_code=entry.activity_code,
                    location=entry.location,
                    description=entry.description,
                    status=EntryStatus.DRAFT,
                    source=entry.source,
                )

        return list(aggregated.values())


def get_timesheet_service() -> TimesheetService:
    """Get a timesheet service instance."""
    return TimesheetService()
