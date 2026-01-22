"""Main desktop application entry point for TIQ Assistant."""

import sys
import logging
from typing import Optional
from datetime import date

from PyQt6.QtWidgets import QApplication, QMessageBox
from PyQt6.QtCore import Qt

from tiq_assistant.core.models import SessionType
from tiq_assistant.storage.sqlite_store import get_store
from tiq_assistant.desktop.tray import TrayIconManager
from tiq_assistant.desktop.scheduler import SchedulerManager

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

    def run(self) -> int:
        """
        Run the desktop application.

        Returns:
            Exit code (0 for success)
        """
        # Create Qt application
        self._app = QApplication(sys.argv)
        self._app.setQuitOnLastWindowClosed(False)  # Keep running in tray
        self._app.setApplicationName("TIQ Assistant")

        # Initialize storage (creates tables if needed)
        store = get_store()
        logger.info("Database initialized")

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

        # Show startup notification
        self._tray_manager.show_notification(
            "TIQ Assistant Running",
            "Time tracking is active. Popups will appear at 12:30 and 18:30.",
        )

        # Run the application event loop
        return self._app.exec()

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

    def _show_time_entry_popup(
        self,
        session: SessionType,
        from_schedule: bool = False
    ) -> None:
        """Show the time entry popup for the given session."""
        logger.info(f"Showing {session.value} time entry popup")

        # Import here to avoid circular imports
        from tiq_assistant.desktop.windows.time_entry_popup import TimeEntryPopup

        # Close existing popup if any
        if self._current_popup is not None:
            self._current_popup.close()

        # Create and show new popup
        self._current_popup = TimeEntryPopup(
            session=session,
            target_date=date.today(),
            scheduler=self._scheduler,
        )

        # Connect popup signals
        self._current_popup.entries_saved.connect(self._on_entries_saved)
        self._current_popup.export_requested.connect(self._export_today)

        # Show notification if from schedule
        if from_schedule and self._tray_manager:
            self._tray_manager.show_popup_reminder(session)

        # Show the popup
        self._current_popup.show()
        self._current_popup.raise_()
        self._current_popup.activateWindow()

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
