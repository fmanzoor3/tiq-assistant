"""Parser for Outlook calendar Excel exports."""

from datetime import datetime, date, timedelta, time as dt_time
from decimal import Decimal
from pathlib import Path
from typing import Optional
import pandas as pd

from tiq_assistant.core.models import CalendarEvent
from tiq_assistant.core.exceptions import ParsingError


class OutlookParser:
    """Parse Outlook calendar exports (Excel format)."""

    # Column name mappings for different Outlook versions
    COLUMN_MAPPINGS = {
        "subject": ["Subject", "Konu"],
        "start_date": ["Start Date", "Başlangıç Tarihi"],
        "start_time": ["Start Time", "Başlangıç Saati"],
        "end_date": ["End Date", "Bitiş Tarihi"],
        "end_time": ["End Time", "Bitiş Saati"],
        "all_day": ["All day event", "Tüm gün etkinliği"],
        "organizer": ["Meeting Organizer", "Toplantı Düzenleyicisi"],
        "required_attendees": ["Required Attendees", "Gerekli Katılımcılar"],
        "optional_attendees": ["Optional Attendees", "İsteğe Bağlı Katılımcılar"],
        "location": ["Location", "Konum"],
        "description": ["Description", "Açıklama"],
    }

    def __init__(self, file_path: Path | str):
        self.file_path = Path(file_path)
        if not self.file_path.exists():
            raise ParsingError(f"File not found: {self.file_path}")

    def parse(self, skip_canceled: bool = True) -> list[CalendarEvent]:
        """Parse the Excel file and return calendar events."""
        try:
            df = pd.read_excel(self.file_path)
        except Exception as e:
            raise ParsingError(f"Failed to read Excel file: {e}")

        # Map column names
        column_map = self._detect_columns(df.columns.tolist())

        events = []
        for idx, row in df.iterrows():
            try:
                event = self._parse_row(row, column_map)
                if event:
                    # Skip canceled meetings if requested
                    if skip_canceled and event.is_canceled:
                        continue
                    events.append(event)
            except Exception as e:
                # Log but continue processing other rows
                print(f"Warning: Failed to parse row {idx}: {e}")
                continue

        return events

    def _detect_columns(self, headers: list[str]) -> dict[str, Optional[str]]:
        """Auto-detect column names from headers."""
        column_map = {}
        headers_lower = {h.lower().strip(): h for h in headers}

        for field, possible_names in self.COLUMN_MAPPINGS.items():
            column_map[field] = None
            for name in possible_names:
                if name.lower() in headers_lower:
                    column_map[field] = headers_lower[name.lower()]
                    break

        return column_map

    def _parse_row(self, row: pd.Series, column_map: dict[str, Optional[str]]) -> Optional[CalendarEvent]:
        """Parse a single row into a CalendarEvent."""
        def get_value(field: str, default=None):
            col = column_map.get(field)
            if col and col in row.index:
                val = row[col]
                if pd.notna(val):
                    return val
            return default

        subject = get_value("subject", "")
        if not subject:
            return None

        # Parse dates
        start_date = self._parse_date(get_value("start_date"))
        end_date = self._parse_date(get_value("end_date"))
        if not start_date:
            return None

        # Parse times
        start_time = self._parse_time(get_value("start_time"))
        end_time = self._parse_time(get_value("end_time"))

        # Calculate duration
        duration_hours = self._calculate_duration(
            start_date, start_time, end_date or start_date, end_time
        )

        # Check if all-day event
        is_all_day = get_value("all_day", False)
        if isinstance(is_all_day, str):
            is_all_day = is_all_day.lower() in ("true", "yes", "1")
        elif isinstance(is_all_day, bool):
            pass  # Already boolean
        else:
            is_all_day = bool(is_all_day) if pd.notna(is_all_day) else False

        # Parse attendees
        attendees = self._parse_attendees(get_value("required_attendees", ""))
        optional = self._parse_attendees(get_value("optional_attendees", ""))
        attendees.extend(optional)

        # Check if canceled
        is_canceled = subject.startswith("Canceled:") or subject.startswith("İptal:")

        # Clean description (remove _x000D_ artifacts)
        description = get_value("description", "")
        if description:
            description = description.replace("_x000D_", "\n").strip()

        return CalendarEvent(
            subject=subject,
            start_date=start_date,
            start_time=start_time or "00:00:00",
            end_date=end_date or start_date,
            end_time=end_time or "23:59:00",
            duration_hours=duration_hours,
            is_all_day=is_all_day,
            is_canceled=is_canceled,
            organizer=get_value("organizer"),
            attendees=attendees,
            location=get_value("location"),
            description=description,
        )

    def _parse_date(self, value) -> Optional[date]:
        """Parse a date value."""
        if value is None:
            return None
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, date):
            return value
        if isinstance(value, str):
            # Try common formats
            for fmt in ["%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%m/%d/%Y"]:
                try:
                    return datetime.strptime(value.strip(), fmt).date()
                except ValueError:
                    continue
        if pd.isna(value):
            return None
        return None

    def _parse_time(self, value) -> Optional[str]:
        """Parse a time value and return as string HH:MM:SS."""
        if value is None or pd.isna(value):
            return None
        if isinstance(value, dt_time):
            return value.strftime("%H:%M:%S")
        if isinstance(value, datetime):
            return value.strftime("%H:%M:%S")
        if isinstance(value, str):
            value = value.strip()
            # Handle various time formats
            for fmt in ["%H:%M:%S", "%H:%M", "%I:%M:%S %p", "%I:%M %p"]:
                try:
                    return datetime.strptime(value, fmt).strftime("%H:%M:%S")
                except ValueError:
                    continue
            return value if ":" in value else None
        return None

    def _calculate_duration(
        self,
        start_date: date,
        start_time: Optional[str],
        end_date: date,
        end_time: Optional[str],
    ) -> Decimal:
        """Calculate duration in hours."""
        # Parse times
        if start_time:
            parts = start_time.split(":")
            start_dt = datetime(
                start_date.year, start_date.month, start_date.day,
                int(parts[0]), int(parts[1]), int(parts[2]) if len(parts) > 2 else 0
            )
        else:
            start_dt = datetime(start_date.year, start_date.month, start_date.day, 0, 0, 0)

        if end_time:
            parts = end_time.split(":")
            end_dt = datetime(
                end_date.year, end_date.month, end_date.day,
                int(parts[0]), int(parts[1]), int(parts[2]) if len(parts) > 2 else 0
            )
        else:
            end_dt = datetime(end_date.year, end_date.month, end_date.day, 23, 59, 0)

        # Calculate difference
        delta = end_dt - start_dt
        hours = delta.total_seconds() / 3600

        # Round to 2 decimal places
        return Decimal(str(round(hours, 2)))

    def _parse_attendees(self, value: str) -> list[str]:
        """Parse attendees string (semicolon-separated)."""
        if not value or pd.isna(value):
            return []
        # Split by semicolon and clean up
        attendees = [a.strip() for a in str(value).split(";") if a.strip()]
        # Filter out email-only entries if there's a name
        cleaned = []
        for attendee in attendees:
            # Skip if it looks like just an email
            if "@" in attendee and not any(c.isupper() for c in attendee.split("@")[0]):
                continue
            cleaned.append(attendee)
        return cleaned


def parse_outlook_calendar(file_path: Path | str, skip_canceled: bool = True) -> list[CalendarEvent]:
    """Convenience function to parse an Outlook calendar export."""
    parser = OutlookParser(file_path)
    return parser.parse(skip_canceled=skip_canceled)
