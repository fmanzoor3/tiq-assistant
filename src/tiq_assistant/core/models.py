"""Core data models for TIQ Assistant."""

from datetime import date, datetime
from decimal import Decimal
from enum import Enum
from typing import Optional
from pydantic import BaseModel, Field, field_validator
import uuid


class ActivityCode(str, Enum):
    """Valid activity codes for timesheet entries."""
    GLST = "GLST"      # General work/development
    TPLNT = "TPLNT"    # Meetings (Toplantı)
    IZIN = "IZIN"      # Leave (İzin)
    TATIL = "TATIL"    # Holiday (Tatil)
    POC = "POC"        # Proof of Concept


class EntryStatus(str, Enum):
    """Status of a timesheet entry."""
    DRAFT = "draft"
    PENDING_REVIEW = "pending_review"
    APPROVED = "approved"
    EXPORTED = "exported"


class EntrySource(str, Enum):
    """Source of a timesheet entry."""
    MANUAL = "manual"
    CALENDAR = "calendar"


def generate_id() -> str:
    """Generate a unique ID."""
    return str(uuid.uuid4())


class Project(BaseModel):
    """A project with its required Ticket No (numeric ID)."""
    id: str = Field(default_factory=generate_id)
    name: str  # e.g., "BI BÜYÜK VERI PLATFORM SUPPORT"
    ticket_number: str  # Required numeric ID, e.g., "2019135" - used in timesheet
    jira_key: Optional[str] = None  # Optional JIRA key, e.g., "PEMP-948"
    keywords: list[str] = Field(default_factory=list)  # For matching calendar events
    default_activity_code: ActivityCode = ActivityCode.GLST
    default_location: str = "ANKARA"
    is_active: bool = True
    created_at: datetime = Field(default_factory=datetime.now)
    updated_at: datetime = Field(default_factory=datetime.now)

    @field_validator('jira_key')
    @classmethod
    def normalize_jira_key(cls, v: Optional[str]) -> Optional[str]:
        if v is None or v.strip() == "":
            return None
        return v.strip().upper()


class CalendarEvent(BaseModel):
    """A parsed calendar event from Outlook export."""
    id: str = Field(default_factory=generate_id)
    subject: str
    start_date: date
    start_time: str  # Time as string, e.g., "10:15:00"
    end_date: date
    end_time: str
    duration_hours: Decimal
    is_all_day: bool = False
    is_canceled: bool = False
    organizer: Optional[str] = None
    attendees: list[str] = Field(default_factory=list)
    location: Optional[str] = None
    description: Optional[str] = None

    # Matching results
    matched_project_id: Optional[str] = None
    matched_ticket_id: Optional[str] = None
    matched_jira_key: Optional[str] = None
    match_confidence: float = 0.0
    match_source: Optional[str] = None  # "subject", "description", "keyword"

    @property
    def display_duration(self) -> str:
        """Format duration for display."""
        hours = float(self.duration_hours)
        if hours == int(hours):
            return f"{int(hours)}h"
        return f"{hours:.2f}h"

    def to_timesheet_description(self) -> str:
        """Generate a description for timesheet entry."""
        # Clean up subject - remove common prefixes
        desc = self.subject
        for prefix in ["FW: ", "RE: ", "Canceled: "]:
            if desc.startswith(prefix):
                desc = desc[len(prefix):]
        return desc


class TimesheetEntry(BaseModel):
    """A single timesheet entry."""
    id: str = Field(default_factory=generate_id)

    # Required fields for export
    consultant_id: str
    entry_date: date
    hours: int  # Integer hours as per timesheet format
    ticket_number: Optional[str] = None  # Numeric ID
    project_name: Optional[str] = None
    activity_code: ActivityCode = ActivityCode.GLST
    location: str = "ANKARA"
    description: str

    # Metadata
    status: EntryStatus = EntryStatus.DRAFT
    source: EntrySource = EntrySource.MANUAL
    source_event_id: Optional[str] = None
    source_jira_key: Optional[str] = None

    # Audit
    created_at: datetime = Field(default_factory=datetime.now)
    updated_at: datetime = Field(default_factory=datetime.now)
    exported_at: Optional[datetime] = None

    @field_validator('hours')
    @classmethod
    def validate_hours(cls, v: int) -> int:
        if v < 1:
            raise ValueError("Hours must be at least 1")
        if v > 24:
            raise ValueError("Hours cannot exceed 24")
        return v

    def to_export_row(self) -> dict:
        """Convert to a row for Excel export."""
        return {
            "Consultant ID": self.consultant_id,
            "Date": self.entry_date.strftime("%d.%m.%Y"),
            "Workhour": self.hours,
            "Ticket No": self.ticket_number or "",
            "Project": self.project_name or "",
            "Activity No": self.activity_code.value,
            "Location": self.location,
            "": "",  # Empty column
            "Activity": self.description,
        }


class UserSettings(BaseModel):
    """User-specific settings."""
    consultant_id: str = "FMANZOOR"
    default_location: str = "ANKARA"
    default_activity_code: ActivityCode = ActivityCode.GLST
    meeting_activity_code: ActivityCode = ActivityCode.TPLNT

    # Matching settings
    min_match_confidence: float = 0.5
    skip_canceled_meetings: bool = True
    min_meeting_duration_minutes: int = 15


class MatchResult(BaseModel):
    """Result of matching a calendar event to a project/ticket."""
    project_id: Optional[str] = None
    project_name: Optional[str] = None
    ticket_id: Optional[str] = None
    ticket_jira_key: Optional[str] = None
    ticket_numeric_id: Optional[str] = None
    confidence: float = 0.0
    match_source: str = "none"  # "jira_key", "keyword", "description_url"
    matched_text: Optional[str] = None


class SessionType(str, Enum):
    """Time entry session type."""
    MORNING = "morning"
    AFTERNOON = "afternoon"


class OutlookMeeting(BaseModel):
    """A meeting fetched from Outlook via COM automation."""
    id: str = Field(default_factory=generate_id)
    subject: str
    start_datetime: datetime
    end_datetime: datetime
    is_teams_meeting: bool = False
    is_recurring: bool = False
    organizer: Optional[str] = None
    location: Optional[str] = None
    body: Optional[str] = None

    # Matching (reuse existing MatchResult logic)
    matched_project_id: Optional[str] = None
    matched_jira_key: Optional[str] = None
    match_confidence: float = 0.0

    # Tracking
    is_imported: bool = False
    imported_entry_id: Optional[str] = None
    fetched_at: datetime = Field(default_factory=datetime.now)

    @property
    def duration_hours(self) -> Decimal:
        """Calculate duration in hours."""
        delta = self.end_datetime - self.start_datetime
        return Decimal(str(round(delta.total_seconds() / 3600, 2)))

    @property
    def duration_minutes(self) -> int:
        """Calculate duration in minutes."""
        delta = self.end_datetime - self.start_datetime
        return int(delta.total_seconds() / 60)

    @property
    def display_time(self) -> str:
        """Format time for display (e.g., '10:00')."""
        return self.start_datetime.strftime("%H:%M")

    @property
    def display_duration(self) -> str:
        """Format duration for display (e.g., '1h 30m')."""
        minutes = self.duration_minutes
        if minutes < 60:
            return f"{minutes}m"
        hours = minutes // 60
        remaining_mins = minutes % 60
        if remaining_mins == 0:
            return f"{hours}h"
        return f"{hours}h {remaining_mins}m"


class ScheduleConfig(BaseModel):
    """User-configurable schedule settings for the desktop app."""
    morning_popup_enabled: bool = True
    morning_popup_time: str = "12:15"
    morning_hours_target: int = 3
    afternoon_popup_enabled: bool = True
    afternoon_popup_time: str = "18:15"
    afternoon_hours_target: int = 5
    workday_start: str = "09:30"
    lunch_start: str = "12:15"
    lunch_end: str = "13:30"
    workday_end: str = "18:15"
    auto_start_with_windows: bool = True


class RecentProject(BaseModel):
    """Tracks recently used projects for quick selection."""
    project_id: str
    project_name: str
    ticket_number: str
    last_used_at: datetime = Field(default_factory=datetime.now)
    use_count: int = 1
