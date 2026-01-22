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
            # Format dates for Outlook filter
            start_str = target_date.strftime("%m/%d/%Y 12:00 AM")
            end_date = target_date + timedelta(days=1)
            end_str = end_date.strftime("%m/%d/%Y 12:00 AM")

            # Get calendar items
            items = self._calendar.Items
            items.Sort("[Start]")
            items.IncludeRecurrences = True

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
            # Format dates for Outlook filter
            start_str = start_date.strftime("%m/%d/%Y 12:00 AM")
            # Add one day to end_date to make it inclusive
            end_date_plus = end_date + timedelta(days=1)
            end_str = end_date_plus.strftime("%m/%d/%Y 12:00 AM")

            # Get calendar items
            items = self._calendar.Items
            items.Sort("[Start]")
            items.IncludeRecurrences = True

            # Restrict to the date range
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

            logger.info(f"Found {len(meetings)} meetings from {start_date} to {end_date}")
            return meetings

        except Exception as e:
            logger.error(f"Error fetching meetings: {e}")
            return []

    def get_meetings_for_session(
        self,
        target_date: date,
        session: str,
        morning_end: str = "12:30",
        afternoon_start: str = "13:30"
    ) -> list[OutlookMeeting]:
        """
        Get meetings for a specific session (morning or afternoon).

        Args:
            target_date: The date to fetch meetings for
            session: Either 'morning' or 'afternoon'
            morning_end: End time for morning session (default 12:30)
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

            # Skip all-day events or very short meetings
            duration_minutes = (end_dt - start_dt).total_seconds() / 60
            if duration_minutes < 15 or duration_minutes > 480:  # 15 min to 8 hours
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
