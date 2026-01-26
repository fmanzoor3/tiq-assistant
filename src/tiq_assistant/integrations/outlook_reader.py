"""Outlook calendar reader using Windows COM automation."""

from datetime import datetime, date, timedelta
from typing import Optional
import logging

from tiq_assistant.core.models import OutlookMeeting, CalendarEvent
from tiq_assistant.core.exceptions import TIQAssistantError

logger = logging.getLogger(__name__)


class OutlookNotAvailableError(TIQAssistantError):
    """Raised when Outlook is not available or COM access is blocked."""
    pass


class OutlookReader:
    """
    Read calendar events from Outlook via COM automation.

    This uses the Windows COM interface to communicate directly with
    Outlook desktop. Requires Outlook to be installed.
    """

    def __init__(self):
        self._outlook = None
        self._namespace = None
        self._calendar = None
        self._available: Optional[bool] = None

    def is_available(self) -> bool:
        """Check if Outlook COM is available."""
        if self._available is not None:
            return self._available

        try:
            self._connect()
            self._available = True
            return True
        except OutlookNotAvailableError:
            self._available = False
            return False

    def _connect(self) -> None:
        """Connect to Outlook via COM."""
        if self._outlook is not None:
            return

        try:
            import win32com.client
            import pythoncom

            # Initialize COM for this thread
            pythoncom.CoInitialize()

            self._outlook = win32com.client.Dispatch("Outlook.Application")
            self._namespace = self._outlook.GetNamespace("MAPI")
            # 9 = olFolderCalendar
            self._calendar = self._namespace.GetDefaultFolder(9)

            logger.info("Connected to Outlook successfully")

        except ImportError:
            raise OutlookNotAvailableError(
                "pywin32 is not installed. Install with: pip install pywin32"
            )
        except Exception as e:
            error_msg = str(e).lower()
            if "class not registered" in error_msg:
                raise OutlookNotAvailableError(
                    "Outlook is not installed on this computer."
                )
            elif "operation aborted" in error_msg or "access denied" in error_msg:
                raise OutlookNotAvailableError(
                    "Access to Outlook was denied. You may need to allow "
                    "programmatic access in Outlook Trust Center settings."
                )
            else:
                raise OutlookNotAvailableError(
                    f"Failed to connect to Outlook: {e}"
                )

    def get_meetings_for_date(self, target_date: date) -> list[OutlookMeeting]:
        """
        Get all calendar events for a specific date.

        Args:
            target_date: The date to fetch meetings for

        Returns:
            List of OutlookMeeting objects
        """
        self._connect()

        meetings = []

        try:
            # Get calendar items
            items = self._calendar.Items

            # IMPORTANT: IncludeRecurrences must be set BEFORE Sort for recurring events
            items.IncludeRecurrences = True
            items.Sort("[Start]")

            # Format dates for Outlook filter (use MM/DD/YYYY format)
            end_date = target_date + timedelta(days=1)
            start_str = target_date.strftime("%m/%d/%Y")
            end_str = end_date.strftime("%m/%d/%Y")

            # Restrict to the specific date
            restriction = f"[Start] >= '{start_str}' AND [Start] < '{end_str}'"
            filtered_items = items.Restrict(restriction)

            for item in filtered_items:
                try:
                    meeting = self._parse_calendar_item(item)
                    if meeting:
                        meetings.append(meeting)
                except Exception as e:
                    logger.warning(f"Failed to parse calendar item: {e}")
                    continue

            logger.info(f"Found {len(meetings)} meetings for {target_date}")
            return meetings

        except Exception as e:
            logger.error(f"Error fetching meetings: {e}")
            return []

    def get_meetings_for_date_range(
        self,
        start_date: date,
        end_date: date
    ) -> list[OutlookMeeting]:
        """
        Get all calendar events for a date range.

        Args:
            start_date: The start date (inclusive)
            end_date: The end date (inclusive)

        Returns:
            List of OutlookMeeting objects
        """
        self._connect()

        meetings = []

        try:
            # Get calendar items
            items = self._calendar.Items

            # IMPORTANT: IncludeRecurrences must be set BEFORE Sort for recurring events
            items.IncludeRecurrences = True
            items.Sort("[Start]")

            # Create datetime bounds for filtering
            start_datetime = datetime.combine(start_date, datetime.min.time())
            end_datetime = datetime.combine(end_date + timedelta(days=1), datetime.min.time())

            # Use MM/DD/YYYY format for Outlook Restrict filter
            start_str = start_date.strftime("%m/%d/%Y")
            end_str = (end_date + timedelta(days=1)).strftime("%m/%d/%Y")

            restriction = f"[Start] >= '{start_str}' AND [Start] < '{end_str}'"
            logger.info(f"Outlook restriction: {restriction}")

            filtered_items = items.Restrict(restriction)

            # When IncludeRecurrences=True, Count returns INT_MAX and normal iteration
            # doesn't work. We need to use Find/FindNext pattern instead.
            item_count = 0
            skipped_count = 0

            # Use GetFirst/GetNext pattern which works with recurring items
            item = filtered_items.GetFirst()
            while item is not None:
                item_count += 1
                try:
                    # Manual date filtering as extra safety
                    item_start = datetime(
                        item.Start.year, item.Start.month, item.Start.day,
                        item.Start.hour, item.Start.minute, item.Start.second
                    )

                    # Stop if we've gone past the end date (items are sorted)
                    if item_start >= end_datetime:
                        break

                    if item_start < start_datetime:
                        skipped_count += 1
                        item = filtered_items.GetNext()
                        continue

                    meeting = self._parse_calendar_item(item)
                    if meeting:
                        meetings.append(meeting)
                except Exception as e:
                    logger.warning(f"Failed to parse calendar item: {e}")

                item = filtered_items.GetNext()

                # Safety limit to prevent infinite loops
                if item_count > 1000:
                    logger.warning("Reached 1000 items limit, stopping")
                    break

            logger.info(f"Found {len(meetings)} meetings from {start_date} to {end_date} (checked {item_count} items, skipped {skipped_count})")
            return meetings

        except Exception as e:
            logger.error(f"Error fetching meetings: {e}", exc_info=True)
            return []

    def get_meetings_for_session(
        self,
        target_date: date,
        session: str,
        morning_end: str = "12:15",
        afternoon_start: str = "13:30"
    ) -> list[OutlookMeeting]:
        """
        Get meetings for a specific session (morning or afternoon).

        Args:
            target_date: The date to fetch meetings for
            session: Either 'morning' or 'afternoon'
            morning_end: End time for morning session (default 12:15)
            afternoon_start: Start time for afternoon session (default 13:30)

        Returns:
            List of OutlookMeeting objects for the session
        """
        all_meetings = self.get_meetings_for_date(target_date)

        # Parse session boundaries
        morning_end_time = datetime.strptime(morning_end, "%H:%M").time()
        afternoon_start_time = datetime.strptime(afternoon_start, "%H:%M").time()

        filtered = []
        for meeting in all_meetings:
            meeting_time = meeting.start_datetime.time()

            if session == "morning":
                # Morning: meetings that start before lunch
                if meeting_time < morning_end_time:
                    filtered.append(meeting)
            else:
                # Afternoon: meetings that start after lunch
                if meeting_time >= afternoon_start_time:
                    filtered.append(meeting)

        return filtered

    def _parse_calendar_item(self, item) -> Optional[OutlookMeeting]:
        """Parse an Outlook calendar item into an OutlookMeeting."""
        try:
            subject = item.Subject or "(No Subject)"

            # Get start and end times
            start_dt = datetime(
                item.Start.year, item.Start.month, item.Start.day,
                item.Start.hour, item.Start.minute, item.Start.second
            )
            end_dt = datetime(
                item.End.year, item.End.month, item.End.day,
                item.End.hour, item.End.minute, item.End.second
            )

            # Check if all-day event
            is_all_day = getattr(item, 'AllDayEvent', False)

            # Skip all-day events (they're usually holidays/PTO, not meetings)
            if is_all_day:
                logger.debug(f"Skipping all-day event: {subject}")
                return None

            # Calculate duration - be more lenient (5 min to 10 hours)
            duration_minutes = (end_dt - start_dt).total_seconds() / 60
            if duration_minutes < 5 or duration_minutes > 600:
                logger.debug(f"Skipping event with duration {duration_minutes} min: {subject}")
                return None

            # Check if it's a Teams meeting
            location = getattr(item, 'Location', '') or ''
            is_teams = (
                'teams' in location.lower() or
                'microsoft teams meeting' in (getattr(item, 'Body', '') or '').lower() or
                getattr(item, 'IsOnlineMeeting', False)
            )

            # Check if recurring
            is_recurring = getattr(item, 'IsRecurring', False)

            # Get organizer
            organizer = None
            try:
                organizer = item.Organizer
            except Exception:
                pass

            # Get body (may be large, truncate)
            body = None
            try:
                raw_body = getattr(item, 'Body', '') or ''
                body = raw_body[:2000] if raw_body else None
            except Exception:
                pass

            return OutlookMeeting(
                subject=subject,
                start_datetime=start_dt,
                end_datetime=end_dt,
                is_teams_meeting=is_teams,
                is_recurring=is_recurring,
                organizer=organizer,
                location=location if location else None,
                body=body,
            )

        except Exception as e:
            logger.warning(f"Error parsing calendar item: {e}")
            return None

    def to_calendar_event(self, meeting: OutlookMeeting) -> CalendarEvent:
        """Convert an OutlookMeeting to a CalendarEvent for matching."""
        from decimal import Decimal

        return CalendarEvent(
            subject=meeting.subject,
            start_date=meeting.start_datetime.date(),
            start_time=meeting.start_datetime.strftime("%H:%M:%S"),
            end_date=meeting.end_datetime.date(),
            end_time=meeting.end_datetime.strftime("%H:%M:%S"),
            duration_hours=meeting.duration_hours,
            is_all_day=False,
            is_canceled=False,
            organizer=meeting.organizer,
            location=meeting.location,
            description=meeting.body,
        )


# Singleton instance
_reader: Optional[OutlookReader] = None


def get_outlook_reader() -> OutlookReader:
    """Get the global OutlookReader instance."""
    global _reader
    if _reader is None:
        _reader = OutlookReader()
    return _reader
