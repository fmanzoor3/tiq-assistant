"""Main desktop application entry point for TIQ Assistant."""

import sys
import logging
from typing import Optional
from datetime import date

from PyQt6.QtWidgets import QApplication, QMessageBox
from PyQt6.QtCore import Qt

from tiq_assistant.storage.sqlite_store import get_store
from tiq_assistant.desktop.tray import TrayIconManager
from tiq_assistant.desktop.scheduler import SchedulerManager
from tiq_assistant.desktop.windows.day_entry_dialog import DayEntryDialog, SessionType
from tiq_assistant.desktop.icon import create_app_icon

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class TIQDesktopApp:
    """
    Main desktop application class.

    Coordinates the system tray, scheduler, and popup windows.
    """

    def __init__(self):
        self._app: Optional[QApplication] = None
        self._tray_manager: Optional[TrayIconManager] = None
        self._scheduler: Optional[SchedulerManager] = None
        self._current_popup = None
        self._main_window = None
        self._missed_days: list = []

    def run(self) -> int:
        """
        Run the desktop application.

        Returns:
            Exit code (0 for success)
        """
        # Set Windows App User Model ID for proper taskbar grouping/icon
        # This must be done BEFORE creating QApplication
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('TIQAssistant.Desktop.1.0')
        except Exception:
            pass  # Not on Windows or API not available

        # Create Qt application
        self._app = QApplication(sys.argv)
        self._app.setQuitOnLastWindowClosed(False)  # Keep running in tray
        self._app.setApplicationName("TIQ Assistant")

        # Set application-wide icon for taskbar
        self._app.setWindowIcon(create_app_icon())

        # Initialize storage (creates tables if needed)
        store = get_store()
        logger.info("Database initialized")

        # Reconcile Windows auto-start with the saved preference on every launch.
        # This is what actually makes "start on login" work: it (re)creates the
        # Startup shortcut if the user enabled it, or removes it if disabled,
        # instead of relying on the Settings dialog having been saved.
        try:
            from tiq_assistant.desktop import autostart
            config = store.get_schedule_config()
            autostart.sync(config.auto_start_with_windows)
        except Exception as e:
            logger.warning(f"Auto-start sync failed: {e}")

        # Set up system tray
        self._tray_manager = TrayIconManager()
        if not self._tray_manager.setup(self._app):
            QMessageBox.critical(
                None,
                "System Tray Error",
                "System tray is not available on this system. "
                "TIQ Assistant requires system tray support."
            )
            return 1

        # Connect tray signals
        self._connect_tray_signals()

        # Show tray icon
        self._tray_manager.show()
        logger.info("System tray icon shown")

        # Set up scheduler
        self._scheduler = SchedulerManager()
        self._connect_scheduler_signals()
        self._scheduler.start()
        logger.info("Scheduler started")

        # Show startup notification (reflect the user's actual configured times).
        config = store.get_schedule_config()
        self._tray_manager.show_notification(
            "TIQ Assistant Running",
            f"Time tracking is active. Popups will appear at "
            f"{config.morning_popup_time} and {config.afternoon_popup_time}.",
        )

        # Surface any recent unfilled workdays so missed days are easy to catch.
        self._check_missed_days()

        # Run the application event loop
        return self._app.exec()

    def _check_missed_days(self, lookback_days: int = 14) -> None:
        """Notify the user about recent workdays with no entries.

        Scans the last ``lookback_days`` days for workdays (excluding weekends,
        full-day holidays, and skipped days) that have zero timesheet entries.
        Clicking the notification opens the oldest such day pre-loaded for entry.
        """
        try:
            from datetime import timedelta
            from tiq_assistant.core.holidays import get_holiday_service

            store = get_store()
            holiday_service = get_holiday_service()
            today = date.today()

            missed: list[date] = []
            for offset in range(1, lookback_days + 1):
                day = today - timedelta(days=offset)

                if not holiday_service.is_workday(day):
                    continue

                is_skipped, _ = store.is_day_skipped(day)
                if is_skipped:
                    continue

                entries = store.get_entries(start_date=day, end_date=day)
                if not entries:
                    missed.append(day)

            if not missed:
                return

            missed.sort()  # Oldest first.
            self._missed_days = missed

            count = len(missed)
            oldest = missed[0].strftime("%d %b")
            newest = missed[-1].strftime("%d %b")
            span = oldest if count == 1 else f"{oldest} – {newest}"

            if self._tray_manager:
                self._tray_manager.show_notification(
                    f"{count} day(s) need timesheet entries",
                    f"Unfilled workdays: {span}. "
                    f"Open the dashboard to fill them in.",
                )
            logger.info(f"Found {count} missed workday(s): {missed}")

        except Exception as e:
            logger.warning(f"Missed-day check failed: {e}")

    def _connect_tray_signals(self) -> None:
        """Connect tray icon signals to handlers."""
        if self._tray_manager is None:
            return

        self._tray_manager.morning_entry_requested.connect(
            lambda: self._show_time_entry_popup(SessionType.MORNING)
        )
        self._tray_manager.afternoon_entry_requested.connect(
            lambda: self._show_time_entry_popup(SessionType.AFTERNOON)
        )
        self._tray_manager.voice_entry_requested.connect(self._show_voice_entry)
        self._tray_manager.sync_requested.connect(self._sync_outlook)
        self._tray_manager.settings_requested.connect(self._show_settings)
        self._tray_manager.dashboard_requested.connect(self._show_dashboard)
        self._tray_manager.quit_requested.connect(self._quit)

    def _connect_scheduler_signals(self) -> None:
        """Connect scheduler signals to handlers."""
        if self._scheduler is None:
            return

        self._scheduler.morning_popup_due.connect(
            lambda: self._show_time_entry_popup(SessionType.MORNING, from_schedule=True)
        )
        self._scheduler.afternoon_popup_due.connect(
            lambda: self._show_time_entry_popup(SessionType.AFTERNOON, from_schedule=True)
        )

    def _ai_enabled(self) -> bool:
        """Whether the AI assistant is turned on in settings."""
        try:
            from tiq_assistant.services.entry_generation_service import load_llm_config
            return load_llm_config(get_store()).enabled
        except Exception:
            return False

    def _show_voice_entry(self) -> None:
        """Open the end-of-day 'What did you do today?' voice/AI dialog.

        Auto-pulls today's meetings into the timesheet first, then opens the
        voice dialog to fill the remaining hours from a spoken/typed summary.
        """
        from tiq_assistant.desktop.windows.voice_entry_dialog import VoiceEntryDialog
        from tiq_assistant.services.hour_suggestion_service import get_hour_suggestion_service
        from tiq_assistant.core.models import SessionType as CoreSession

        today = date.today()

        # Auto-add today's meetings as entries (so they're captured alongside).
        try:
            self._auto_add_todays_meetings(today)
        except Exception as e:
            logger.warning(f"Could not auto-add meetings: {e}")

        # Work out how many hours still need filling for the day.
        store = get_store()
        config = store.get_schedule_config()
        target = config.morning_hours_target + config.afternoon_hours_target
        entries = store.get_entries(start_date=today, end_date=today)
        filled = sum(e.hours for e in entries)
        remaining = max(1, target - filled)

        dialog = VoiceEntryDialog(target_date=today, remaining_hours=remaining)
        dialog.exec()

    def _auto_add_todays_meetings(self, day: date) -> None:
        """Save today's un-imported Outlook meetings as draft entries."""
        meetings = self._get_today_meetings()
        if not meetings:
            return
        store = get_store()
        settings = store.get_settings()
        from tiq_assistant.core.models import (
            TimesheetEntry, ActivityCode, EntryStatus, EntrySource,
        )
        existing = store.get_entries(start_date=day, end_date=day)
        existing_event_ids = {e.source_event_id for e in existing if e.source_event_id}

        for m in meetings:
            if m.id in existing_event_ids:
                continue
            project = store.get_project(m.matched_project_id) if m.matched_project_id else None
            entry = TimesheetEntry(
                consultant_id=settings.consultant_id,
                entry_date=day,
                hours=max(1, round(float(m.duration_hours))),
                ticket_number=project.ticket_number if project else None,
                project_name=project.name if project else None,
                activity_code=settings.meeting_activity_code,
                location=settings.default_location,
                description=m.subject,
                status=EntryStatus.DRAFT,
                source=EntrySource.CALENDAR,
                source_event_id=m.id,
                source_jira_key=m.matched_jira_key,
            )
            store.save_entry(entry)

    def _show_time_entry_popup(
        self,
        session: SessionType,
        from_schedule: bool = False
    ) -> None:
        """Show the time entry popup for the given session."""
        logger.info(f"Showing {session.value} time entry popup")

        # End-of-day (afternoon) popup uses the voice/AI dialog when enabled.
        if session == SessionType.AFTERNOON and self._ai_enabled():
            if from_schedule and self._tray_manager:
                self._tray_manager.show_popup_reminder(session)
            self._show_voice_entry()
            return

        # Close existing popup if any
        if self._current_popup is not None:
            self._current_popup.close()

        # Fetch today's Outlook meetings for the dialog
        outlook_meetings = self._get_today_meetings()

        # Create the dialog
        self._current_popup = DayEntryDialog(
            target_date=date.today(),
            session=session,
            outlook_meetings=outlook_meetings,
        )

        # Show notification if from schedule
        if from_schedule and self._tray_manager:
            self._tray_manager.show_popup_reminder(session)

        # Show the popup as a modal dialog
        result = self._current_popup.exec()

        # Handle snooze request (dialog returns 2 for snooze)
        if self._current_popup.get_snooze_requested():
            if self._scheduler:
                if session == SessionType.MORNING:
                    self._scheduler.snooze_morning()
                else:
                    self._scheduler.snooze_afternoon()
                if self._tray_manager:
                    self._tray_manager.show_notification(
                        "Snoozed",
                        "Reminder snoozed for 15 minutes."
                    )

        self._current_popup = None

    def _get_today_meetings(self) -> list:
        """Fetch today's meetings from Outlook."""
        try:
            from tiq_assistant.integrations.outlook_reader import (
                get_outlook_reader, OutlookNotAvailableError
            )
            from tiq_assistant.services.matching_service import get_matching_service

            reader = get_outlook_reader()

            if not reader.is_available():
                return []

            meetings = reader.get_meetings_for_date(date.today())

            # Match meetings to projects
            matching_service = get_matching_service()
            for meeting in meetings:
                event = reader.to_calendar_event(meeting)
                result = matching_service.match_event(event)
                meeting.matched_project_id = result.project_id
                meeting.matched_jira_key = result.ticket_jira_key
                meeting.match_confidence = result.confidence

            return meetings

        except Exception as e:
            logger.warning(f"Failed to fetch Outlook meetings: {e}")
            return []

    def _sync_outlook(self) -> None:
        """Sync calendar from Outlook."""
        logger.info("Syncing Outlook calendar")

        try:
            from tiq_assistant.integrations.outlook_reader import (
                get_outlook_reader, OutlookNotAvailableError
            )
            from tiq_assistant.services.matching_service import get_matching_service

            reader = get_outlook_reader()

            if not reader.is_available():
                QMessageBox.warning(
                    None,
                    "Outlook Not Available",
                    "Could not connect to Outlook. Make sure Outlook is installed "
                    "and running on this computer."
                )
                return

            # Get today's meetings
            meetings = reader.get_meetings_for_date(date.today())
            logger.info(f"Found {len(meetings)} meetings")

            # Match meetings to projects
            matching_service = get_matching_service()
            store = get_store()

            for meeting in meetings:
                # Convert to CalendarEvent for matching
                event = reader.to_calendar_event(meeting)
                result = matching_service.match_event(event)

                # Update meeting with match results
                meeting.matched_project_id = result.project_id
                meeting.matched_jira_key = result.ticket_jira_key
                meeting.match_confidence = result.confidence

                # Save to database
                store.save_outlook_meeting(meeting)

            # Show notification
            if self._tray_manager:
                self._tray_manager.show_notification(
                    "Outlook Sync Complete",
                    f"Found {len(meetings)} meetings for today."
                )

        except Exception as e:
            logger.error(f"Error syncing Outlook: {e}")
            QMessageBox.warning(
                None,
                "Sync Error",
                f"Failed to sync Outlook calendar: {e}"
            )

    def _show_settings(self) -> None:
        """Show the settings dialog."""
        logger.info("Opening settings dialog")

        from tiq_assistant.desktop.windows.settings_dialog import SettingsDialog

        dialog = SettingsDialog(scheduler=self._scheduler)
        dialog.exec()

    def _show_dashboard(self) -> None:
        """Show the main dashboard window."""
        logger.info("Opening main window")

        from tiq_assistant.desktop.windows.main_window import MainWindow

        # Create main window if it doesn't exist
        if self._main_window is None:
            self._main_window = MainWindow()

        # Show and bring to front
        self._main_window.show()
        self._main_window.raise_()
        self._main_window.activateWindow()

    def _on_entries_saved(self, count: int) -> None:
        """Handle entries saved event."""
        if self._tray_manager:
            self._tray_manager.show_notification(
                "Entries Saved",
                f"Saved {count} timesheet {'entry' if count == 1 else 'entries'}."
            )

    def _export_today(self) -> None:
        """Export today's entries to Excel."""
        logger.info("Exporting today's entries")

        try:
            from tiq_assistant.exporters.excel_exporter import ExcelExporter
            from pathlib import Path
            import os

            store = get_store()

            # Get today's entries
            today = date.today()
            entries = store.get_entries(start_date=today, end_date=today)

            if not entries:
                QMessageBox.information(
                    None,
                    "No Entries",
                    "No entries to export for today."
                )
                return

            # Create export directory
            export_dir = Path(os.path.expanduser("~/Documents/TIQ Timesheets"))
            export_dir.mkdir(parents=True, exist_ok=True)

            # Generate filename
            month_str = today.strftime("%Y-%m")
            export_path = export_dir / f"Timesheet_{month_str}.xlsx"

            # Export
            exporter = ExcelExporter()
            if export_path.exists():
                exporter.append_to_existing(entries, export_path)
            else:
                exporter.export_to_new_file(entries, export_path)

            # Mark as exported
            store.mark_entries_exported([e.id for e in entries])

            QMessageBox.information(
                None,
                "Export Complete",
                f"Exported {len(entries)} entries to:\n{export_path}"
            )

        except Exception as e:
            logger.error(f"Export error: {e}")
            QMessageBox.warning(
                None,
                "Export Error",
                f"Failed to export entries: {e}"
            )

    def _quit(self) -> None:
        """Quit the application."""
        logger.info("Quitting application")

        # Stop scheduler
        if self._scheduler:
            self._scheduler.stop()

        # Hide tray icon
        if self._tray_manager:
            self._tray_manager.hide()

        # Quit application
        if self._app:
            self._app.quit()


def main() -> int:
    """Main entry point for the desktop application."""
    app = TIQDesktopApp()
    return app.run()


if __name__ == "__main__":
    sys.exit(main())
