"""SQLite storage implementation for TIQ Assistant."""

import json
import sqlite3
from datetime import datetime, date
from pathlib import Path
from typing import Optional

from tiq_assistant.core.models import (
    Project, TimesheetEntry, UserSettings,
    ActivityCode, EntryStatus, EntrySource,
    ScheduleConfig, RecentProject, OutlookMeeting
)
from tiq_assistant.core.exceptions import StorageError, ProjectNotFoundError
from tiq_assistant.config import settings


class SQLiteStore:
    """SQLite-based storage for TIQ Assistant data."""

    def __init__(self, db_path: Optional[Path] = None):
        self.db_path = db_path or settings.database_path
        self._ensure_db_exists()

    def _ensure_db_exists(self) -> None:
        """Create database and tables if they don't exist."""
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._create_tables()

    def _get_connection(self) -> sqlite3.Connection:
        """Get a database connection."""
        conn = sqlite3.connect(str(self.db_path))
        conn.row_factory = sqlite3.Row
        return conn

    def _create_tables(self) -> None:
        """Create all required tables."""
        with self._get_connection() as conn:
            conn.executescript("""
                -- Projects table (simplified - each project has one ticket_number)
                CREATE TABLE IF NOT EXISTS projects (
                    id TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    ticket_number TEXT NOT NULL,
                    jira_key TEXT,
                    keywords TEXT DEFAULT '[]',
                    default_activity_code TEXT DEFAULT 'GLST',
                    default_location TEXT DEFAULT 'ANKARA',
                    is_active INTEGER DEFAULT 1,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );

                -- Timesheet entries table
                CREATE TABLE IF NOT EXISTS timesheet_entries (
                    id TEXT PRIMARY KEY,
                    consultant_id TEXT NOT NULL,
                    entry_date TEXT NOT NULL,
                    hours INTEGER NOT NULL,
                    ticket_number TEXT,
                    project_name TEXT,
                    activity_code TEXT NOT NULL,
                    location TEXT NOT NULL,
                    description TEXT NOT NULL,
                    status TEXT DEFAULT 'draft',
                    source TEXT DEFAULT 'manual',
                    source_event_id TEXT,
                    source_jira_key TEXT,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    exported_at TEXT
                );

                -- User settings table
                CREATE TABLE IF NOT EXISTS user_settings (
                    id INTEGER PRIMARY KEY CHECK (id = 1),
                    consultant_id TEXT NOT NULL,
                    default_location TEXT NOT NULL,
                    default_activity_code TEXT NOT NULL,
                    meeting_activity_code TEXT NOT NULL,
                    min_match_confidence REAL DEFAULT 0.5,
                    skip_canceled_meetings INTEGER DEFAULT 1,
                    min_meeting_duration_minutes INTEGER DEFAULT 15
                );

                -- Schedule configuration for desktop app
                CREATE TABLE IF NOT EXISTS schedule_config (
                    id INTEGER PRIMARY KEY CHECK (id = 1),
                    morning_popup_enabled INTEGER DEFAULT 1,
                    morning_popup_time TEXT DEFAULT '12:15',
                    morning_hours_target INTEGER DEFAULT 3,
                    afternoon_popup_enabled INTEGER DEFAULT 1,
                    afternoon_popup_time TEXT DEFAULT '18:15',
                    afternoon_hours_target INTEGER DEFAULT 5,
                    workday_start TEXT DEFAULT '09:30',
                    lunch_start TEXT DEFAULT '12:15',
                    lunch_end TEXT DEFAULT '13:30',
                    workday_end TEXT DEFAULT '18:15',
                    auto_start_with_windows INTEGER DEFAULT 1
                );

                -- Recent projects for quick selection
                CREATE TABLE IF NOT EXISTS recent_projects (
                    project_id TEXT PRIMARY KEY,
                    project_name TEXT NOT NULL,
                    ticket_number TEXT NOT NULL,
                    last_used_at TEXT NOT NULL,
                    use_count INTEGER DEFAULT 1,
                    FOREIGN KEY (project_id) REFERENCES projects(id)
                );

                -- Cached Outlook meetings
                CREATE TABLE IF NOT EXISTS outlook_meetings (
                    id TEXT PRIMARY KEY,
                    subject TEXT NOT NULL,
                    start_datetime TEXT NOT NULL,
                    end_datetime TEXT NOT NULL,
                    is_teams_meeting INTEGER DEFAULT 0,
                    is_recurring INTEGER DEFAULT 0,
                    organizer TEXT,
                    location TEXT,
                    body TEXT,
                    matched_project_id TEXT,
                    matched_jira_key TEXT,
                    match_confidence REAL DEFAULT 0.0,
                    is_imported INTEGER DEFAULT 0,
                    imported_entry_id TEXT,
                    fetched_at TEXT NOT NULL,
                    FOREIGN KEY (matched_project_id) REFERENCES projects(id),
                    FOREIGN KEY (imported_entry_id) REFERENCES timesheet_entries(id)
                );

                -- Custom holidays (from uploaded PDF/JPG)
                CREATE TABLE IF NOT EXISTS holidays (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    holiday_date TEXT NOT NULL UNIQUE,
                    name TEXT NOT NULL,
                    holiday_type TEXT DEFAULT 'full_day',
                    source_file TEXT,
                    created_at TEXT NOT NULL
                );

                -- Indexes
                CREATE INDEX IF NOT EXISTS idx_entries_date ON timesheet_entries(entry_date);
                CREATE INDEX IF NOT EXISTS idx_entries_status ON timesheet_entries(status);
                CREATE INDEX IF NOT EXISTS idx_projects_jira_key ON projects(jira_key);
                CREATE INDEX IF NOT EXISTS idx_recent_projects_last_used ON recent_projects(last_used_at);
                CREATE INDEX IF NOT EXISTS idx_meetings_start ON outlook_meetings(start_datetime);
                CREATE INDEX IF NOT EXISTS idx_holidays_date ON holidays(holiday_date);
            """)
            conn.commit()

    # ==================== Projects ====================

    def save_project(self, project: Project) -> Project:
        """Save or update a project."""
        project.updated_at = datetime.now()
        with self._get_connection() as conn:
            conn.execute("""
                INSERT OR REPLACE INTO projects
                (id, name, ticket_number, jira_key, keywords, default_activity_code,
                 default_location, is_active, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                project.id,
                project.name,
                project.ticket_number,
                project.jira_key,
                json.dumps(project.keywords),
                project.default_activity_code.value,
                project.default_location,
                1 if project.is_active else 0,
                project.created_at.isoformat(),
                project.updated_at.isoformat(),
            ))
            conn.commit()
        return project

    def get_projects(self, active_only: bool = True) -> list[Project]:
        """Get all projects."""
        with self._get_connection() as conn:
            query = "SELECT * FROM projects"
            if active_only:
                query += " WHERE is_active = 1"
            query += " ORDER BY name"

            projects = []
            for row in conn.execute(query).fetchall():
                projects.append(self._row_to_project(row))
            return projects

    def get_project(self, project_id: str) -> Optional[Project]:
        """Get a project by ID."""
        with self._get_connection() as conn:
            row = conn.execute(
                "SELECT * FROM projects WHERE id = ?", (project_id,)
            ).fetchone()
            if not row:
                return None
            return self._row_to_project(row)

    def _row_to_project(self, row: sqlite3.Row) -> Project:
        """Convert a database row to a Project."""
        return Project(
            id=row["id"],
            name=row["name"],
            ticket_number=row["ticket_number"],
            jira_key=row["jira_key"],
            keywords=json.loads(row["keywords"]),
            default_activity_code=ActivityCode(row["default_activity_code"]),
            default_location=row["default_location"],
            is_active=bool(row["is_active"]),
            created_at=datetime.fromisoformat(row["created_at"]),
            updated_at=datetime.fromisoformat(row["updated_at"]),
        )

    def delete_project(self, project_id: str) -> None:
        """Soft-delete a project."""
        with self._get_connection() as conn:
            conn.execute(
                "UPDATE projects SET is_active = 0, updated_at = ? WHERE id = ?",
                (datetime.now().isoformat(), project_id)
            )
            conn.commit()

    def find_project_by_jira_key(self, jira_key: str) -> Optional[Project]:
        """Find a project by its JIRA key."""
        jira_key = jira_key.upper()
        with self._get_connection() as conn:
            row = conn.execute(
                "SELECT * FROM projects WHERE jira_key = ? AND is_active = 1",
                (jira_key,)
            ).fetchone()
            if not row:
                return None
            return self._row_to_project(row)

    def find_project_by_keyword(self, text: str) -> Optional[Project]:
        """Find a project by keyword match in text."""
        text_lower = text.lower()
        projects = self.get_projects(active_only=True)

        for project in projects:
            for keyword in project.keywords:
                if keyword.lower() in text_lower:
                    return project
        return None

    # ==================== Timesheet Entries ====================

    def save_entry(self, entry: TimesheetEntry) -> TimesheetEntry:
        """Save or update a timesheet entry."""
        entry.updated_at = datetime.now()
        with self._get_connection() as conn:
            conn.execute("""
                INSERT OR REPLACE INTO timesheet_entries
                (id, consultant_id, entry_date, hours, ticket_number, project_name,
                 activity_code, location, description, status, source, source_event_id,
                 source_jira_key, created_at, updated_at, exported_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                entry.id,
                entry.consultant_id,
                entry.entry_date.isoformat(),
                entry.hours,
                entry.ticket_number,
                entry.project_name,
                entry.activity_code.value,
                entry.location,
                entry.description,
                entry.status.value,
                entry.source.value,
                entry.source_event_id,
                entry.source_jira_key,
                entry.created_at.isoformat(),
                entry.updated_at.isoformat(),
                entry.exported_at.isoformat() if entry.exported_at else None,
            ))
            conn.commit()
        return entry

    def get_entries(
        self,
        start_date: Optional[date] = None,
        end_date: Optional[date] = None,
        status: Optional[EntryStatus] = None,
    ) -> list[TimesheetEntry]:
        """Get timesheet entries with optional filters."""
        with self._get_connection() as conn:
            query = "SELECT * FROM timesheet_entries WHERE 1=1"
            params = []

            if start_date:
                query += " AND entry_date >= ?"
                params.append(start_date.isoformat())
            if end_date:
                query += " AND entry_date <= ?"
                params.append(end_date.isoformat())
            if status:
                query += " AND status = ?"
                params.append(status.value)

            query += " ORDER BY entry_date, created_at"

            entries = []
            for row in conn.execute(query, params).fetchall():
                entries.append(self._row_to_entry(row))
            return entries

    def get_entry(self, entry_id: str) -> Optional[TimesheetEntry]:
        """Get an entry by ID."""
        with self._get_connection() as conn:
            row = conn.execute(
                "SELECT * FROM timesheet_entries WHERE id = ?", (entry_id,)
            ).fetchone()
            return self._row_to_entry(row) if row else None

    def delete_entry(self, entry_id: str) -> None:
        """Delete a timesheet entry."""
        with self._get_connection() as conn:
            conn.execute("DELETE FROM timesheet_entries WHERE id = ?", (entry_id,))
            conn.commit()

    def _row_to_entry(self, row: sqlite3.Row) -> TimesheetEntry:
        """Convert a database row to a TimesheetEntry."""
        return TimesheetEntry(
            id=row["id"],
            consultant_id=row["consultant_id"],
            entry_date=date.fromisoformat(row["entry_date"]),
            hours=row["hours"],
            ticket_number=row["ticket_number"],
            project_name=row["project_name"],
            activity_code=ActivityCode(row["activity_code"]),
            location=row["location"],
            description=row["description"],
            status=EntryStatus(row["status"]),
            source=EntrySource(row["source"]),
            source_event_id=row["source_event_id"],
            source_jira_key=row["source_jira_key"],
            created_at=datetime.fromisoformat(row["created_at"]),
            updated_at=datetime.fromisoformat(row["updated_at"]),
            exported_at=datetime.fromisoformat(row["exported_at"]) if row["exported_at"] else None,
        )

    def mark_entries_exported(self, entry_ids: list[str]) -> None:
        """Mark entries as exported."""
        now = datetime.now().isoformat()
        with self._get_connection() as conn:
            for entry_id in entry_ids:
                conn.execute("""
                    UPDATE timesheet_entries
                    SET status = ?, exported_at = ?, updated_at = ?
                    WHERE id = ?
                """, (EntryStatus.EXPORTED.value, now, now, entry_id))
            conn.commit()

    # ==================== User Settings ====================

    def get_settings(self) -> UserSettings:
        """Get user settings."""
        with self._get_connection() as conn:
            row = conn.execute("SELECT * FROM user_settings WHERE id = 1").fetchone()
            if not row:
                # Return defaults
                return UserSettings()
            return UserSettings(
                consultant_id=row["consultant_id"],
                default_location=row["default_location"],
                default_activity_code=ActivityCode(row["default_activity_code"]),
                meeting_activity_code=ActivityCode(row["meeting_activity_code"]),
                min_match_confidence=row["min_match_confidence"],
                skip_canceled_meetings=bool(row["skip_canceled_meetings"]),
                min_meeting_duration_minutes=row["min_meeting_duration_minutes"],
            )

    def save_settings(self, user_settings: UserSettings) -> None:
        """Save user settings."""
        with self._get_connection() as conn:
            conn.execute("""
                INSERT OR REPLACE INTO user_settings
                (id, consultant_id, default_location, default_activity_code,
                 meeting_activity_code, min_match_confidence, skip_canceled_meetings,
                 min_meeting_duration_minutes)
                VALUES (1, ?, ?, ?, ?, ?, ?, ?)
            """, (
                user_settings.consultant_id,
                user_settings.default_location,
                user_settings.default_activity_code.value,
                user_settings.meeting_activity_code.value,
                user_settings.min_match_confidence,
                1 if user_settings.skip_canceled_meetings else 0,
                user_settings.min_meeting_duration_minutes,
            ))
            conn.commit()

    # ==================== Schedule Config ====================

    def get_schedule_config(self) -> ScheduleConfig:
        """Get schedule configuration."""
        with self._get_connection() as conn:
            row = conn.execute("SELECT * FROM schedule_config WHERE id = 1").fetchone()
            if not row:
                return ScheduleConfig()
            return ScheduleConfig(
                morning_popup_enabled=bool(row["morning_popup_enabled"]),
                morning_popup_time=row["morning_popup_time"],
                morning_hours_target=row["morning_hours_target"],
                afternoon_popup_enabled=bool(row["afternoon_popup_enabled"]),
                afternoon_popup_time=row["afternoon_popup_time"],
                afternoon_hours_target=row["afternoon_hours_target"],
                workday_start=row["workday_start"],
                lunch_start=row["lunch_start"],
                lunch_end=row["lunch_end"],
                workday_end=row["workday_end"],
                auto_start_with_windows=bool(row["auto_start_with_windows"]),
            )

    def save_schedule_config(self, config: ScheduleConfig) -> None:
        """Save schedule configuration."""
        with self._get_connection() as conn:
            conn.execute("""
                INSERT OR REPLACE INTO schedule_config
                (id, morning_popup_enabled, morning_popup_time, morning_hours_target,
                 afternoon_popup_enabled, afternoon_popup_time, afternoon_hours_target,
                 workday_start, lunch_start, lunch_end, workday_end, auto_start_with_windows)
                VALUES (1, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                1 if config.morning_popup_enabled else 0,
                config.morning_popup_time,
                config.morning_hours_target,
                1 if config.afternoon_popup_enabled else 0,
                config.afternoon_popup_time,
                config.afternoon_hours_target,
                config.workday_start,
                config.lunch_start,
                config.lunch_end,
                config.workday_end,
                1 if config.auto_start_with_windows else 0,
            ))
            conn.commit()

    # ==================== Recent Projects ====================

    def get_recent_projects(self, limit: int = 10) -> list[RecentProject]:
        """Get recently used projects, sorted by last used."""
        with self._get_connection() as conn:
            rows = conn.execute("""
                SELECT * FROM recent_projects
                ORDER BY last_used_at DESC
                LIMIT ?
            """, (limit,)).fetchall()
            return [
                RecentProject(
                    project_id=row["project_id"],
                    project_name=row["project_name"],
                    ticket_number=row["ticket_number"],
                    last_used_at=datetime.fromisoformat(row["last_used_at"]),
                    use_count=row["use_count"],
                )
                for row in rows
            ]

    def update_recent_project(self, project: Project) -> None:
        """Update or add a project to recent projects."""
        now = datetime.now().isoformat()
        with self._get_connection() as conn:
            existing = conn.execute(
                "SELECT use_count FROM recent_projects WHERE project_id = ?",
                (project.id,)
            ).fetchone()

            if existing:
                conn.execute("""
                    UPDATE recent_projects
                    SET last_used_at = ?, use_count = use_count + 1,
                        project_name = ?, ticket_number = ?
                    WHERE project_id = ?
                """, (now, project.name, project.ticket_number, project.id))
            else:
                conn.execute("""
                    INSERT INTO recent_projects
                    (project_id, project_name, ticket_number, last_used_at, use_count)
                    VALUES (?, ?, ?, ?, 1)
                """, (project.id, project.name, project.ticket_number, now))
            conn.commit()

    # ==================== Outlook Meetings ====================

    def save_outlook_meeting(self, meeting: OutlookMeeting) -> OutlookMeeting:
        """Save or update an Outlook meeting."""
        with self._get_connection() as conn:
            conn.execute("""
                INSERT OR REPLACE INTO outlook_meetings
                (id, subject, start_datetime, end_datetime, is_teams_meeting, is_recurring,
                 organizer, location, body, matched_project_id, matched_jira_key,
                 match_confidence, is_imported, imported_entry_id, fetched_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                meeting.id,
                meeting.subject,
                meeting.start_datetime.isoformat(),
                meeting.end_datetime.isoformat(),
                1 if meeting.is_teams_meeting else 0,
                1 if meeting.is_recurring else 0,
                meeting.organizer,
                meeting.location,
                meeting.body,
                meeting.matched_project_id,
                meeting.matched_jira_key,
                meeting.match_confidence,
                1 if meeting.is_imported else 0,
                meeting.imported_entry_id,
                meeting.fetched_at.isoformat(),
            ))
            conn.commit()
        return meeting

    def get_meetings_for_date(self, target_date: date) -> list[OutlookMeeting]:
        """Get all meetings for a specific date."""
        start_of_day = datetime.combine(target_date, datetime.min.time())
        end_of_day = datetime.combine(target_date, datetime.max.time())

        with self._get_connection() as conn:
            rows = conn.execute("""
                SELECT * FROM outlook_meetings
                WHERE start_datetime >= ? AND start_datetime <= ?
                ORDER BY start_datetime
            """, (start_of_day.isoformat(), end_of_day.isoformat())).fetchall()
            return [self._row_to_meeting(row) for row in rows]

    def _row_to_meeting(self, row: sqlite3.Row) -> OutlookMeeting:
        """Convert a database row to an OutlookMeeting."""
        return OutlookMeeting(
            id=row["id"],
            subject=row["subject"],
            start_datetime=datetime.fromisoformat(row["start_datetime"]),
            end_datetime=datetime.fromisoformat(row["end_datetime"]),
            is_teams_meeting=bool(row["is_teams_meeting"]),
            is_recurring=bool(row["is_recurring"]),
            organizer=row["organizer"],
            location=row["location"],
            body=row["body"],
            matched_project_id=row["matched_project_id"],
            matched_jira_key=row["matched_jira_key"],
            match_confidence=row["match_confidence"],
            is_imported=bool(row["is_imported"]),
            imported_entry_id=row["imported_entry_id"],
            fetched_at=datetime.fromisoformat(row["fetched_at"]),
        )

    def mark_meeting_imported(self, meeting_id: str, entry_id: str) -> None:
        """Mark a meeting as imported with link to the created entry."""
        with self._get_connection() as conn:
            conn.execute("""
                UPDATE outlook_meetings
                SET is_imported = 1, imported_entry_id = ?
                WHERE id = ?
            """, (entry_id, meeting_id))
            conn.commit()

    def clear_old_meetings(self, days_to_keep: int = 7) -> None:
        """Clear meetings older than specified days."""
        cutoff = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        cutoff = cutoff.replace(day=cutoff.day - days_to_keep)

        with self._get_connection() as conn:
            conn.execute("""
                DELETE FROM outlook_meetings
                WHERE start_datetime < ?
            """, (cutoff.isoformat(),))
            conn.commit()

    # ==================== Holidays ====================

    def save_holiday(self, holiday_date: date, name: str, holiday_type: str = "full_day",
                     source_file: str = None) -> None:
        """Save a holiday to the database."""
        with self._get_connection() as conn:
            conn.execute("""
                INSERT OR REPLACE INTO holidays
                (holiday_date, name, holiday_type, source_file, created_at)
                VALUES (?, ?, ?, ?, ?)
            """, (
                holiday_date.isoformat(),
                name,
                holiday_type,
                source_file,
                datetime.now().isoformat(),
            ))
            conn.commit()

    def save_holidays_batch(self, holidays: list[tuple[date, str, str]], source_file: str = None) -> int:
        """Save multiple holidays at once. Returns count of saved holidays."""
        now = datetime.now().isoformat()
        count = 0
        with self._get_connection() as conn:
            for holiday_date, name, holiday_type in holidays:
                try:
                    conn.execute("""
                        INSERT OR REPLACE INTO holidays
                        (holiday_date, name, holiday_type, source_file, created_at)
                        VALUES (?, ?, ?, ?, ?)
                    """, (holiday_date.isoformat(), name, holiday_type, source_file, now))
                    count += 1
                except Exception:
                    pass  # Skip duplicates or errors
            conn.commit()
        return count

    def get_holidays(self, year: int = None) -> list[dict]:
        """Get all holidays, optionally filtered by year."""
        with self._get_connection() as conn:
            if year:
                rows = conn.execute("""
                    SELECT id, holiday_date, name, holiday_type, source_file
                    FROM holidays
                    WHERE holiday_date LIKE ?
                    ORDER BY holiday_date
                """, (f"{year}-%",)).fetchall()
            else:
                rows = conn.execute("""
                    SELECT id, holiday_date, name, holiday_type, source_file
                    FROM holidays
                    ORDER BY holiday_date
                """).fetchall()

            return [
                {
                    "id": row["id"],
                    "holiday_date": date.fromisoformat(row["holiday_date"]),
                    "name": row["name"],
                    "holiday_type": row["holiday_type"],
                    "source_file": row["source_file"],
                }
                for row in rows
            ]

    def delete_holidays_by_source(self, source_file: str) -> int:
        """Delete all holidays from a specific source file. Returns count deleted."""
        with self._get_connection() as conn:
            cursor = conn.execute(
                "DELETE FROM holidays WHERE source_file = ?",
                (source_file,)
            )
            conn.commit()
            return cursor.rowcount

    def delete_holiday(self, holiday_id: int) -> bool:
        """Delete a single holiday by ID. Returns True if deleted."""
        with self._get_connection() as conn:
            cursor = conn.execute(
                "DELETE FROM holidays WHERE id = ?",
                (holiday_id,)
            )
            conn.commit()
            return cursor.rowcount > 0

    def clear_all_holidays(self) -> int:
        """Clear all custom holidays. Returns count deleted."""
        with self._get_connection() as conn:
            cursor = conn.execute("DELETE FROM holidays")
            conn.commit()
            return cursor.rowcount


# Global store instance
_store: Optional[SQLiteStore] = None


def get_store() -> SQLiteStore:
    """Get the global store instance."""
    global _store
    if _store is None:
        _store = SQLiteStore()
    return _store
