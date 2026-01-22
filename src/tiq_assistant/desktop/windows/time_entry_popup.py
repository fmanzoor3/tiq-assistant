"""Time entry popup dialog for quick time logging."""

from datetime import date
from typing import Optional, List
from decimal import Decimal

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QComboBox, QLineEdit, QSpinBox, QListWidget, QListWidgetItem,
    QGroupBox, QCheckBox, QWidget, QFrame, QMessageBox, QSizePolicy
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QFont

from tiq_assistant.core.models import (
    SessionType, Project, TimesheetEntry, ActivityCode,
    EntryStatus, EntrySource, OutlookMeeting
)
from tiq_assistant.storage.sqlite_store import get_store
from tiq_assistant.services.hour_suggestion_service import get_hour_suggestion_service
from tiq_assistant.desktop.scheduler import SchedulerManager


class TimeEntryPopup(QDialog):
    """
    Popup dialog for quick time entry.

    Features:
    - Quick project selection from recent/all projects
    - Hour suggestion based on remaining time
    - Display of detected Outlook meetings
    - One-click meeting import
    - Session summary and day review (afternoon only)
    """

    entries_saved = pyqtSignal(int)  # Emitted with count when entries are saved
    export_requested = pyqtSignal()  # Emitted when user clicks export

    def __init__(
        self,
        session: SessionType,
        target_date: date,
        scheduler: Optional[SchedulerManager] = None,
        parent: Optional[QWidget] = None
    ):
        super().__init__(parent)

        self.session = session
        self.target_date = target_date
        self._scheduler = scheduler
        self._store = get_store()
        self._hour_service = get_hour_suggestion_service()
        self._pending_entries: List[dict] = []

        self._setup_ui()
        self._load_data()

    def _setup_ui(self) -> None:
        """Set up the dialog UI."""
        # Window properties
        session_name = "Morning" if self.session == SessionType.MORNING else "Afternoon"
        self.setWindowTitle(f"{session_name} Time Entry - {self.target_date.strftime('%b %d, %Y')}")
        self.setMinimumWidth(500)
        self.setMinimumHeight(600)
        self.setWindowFlags(
            Qt.WindowType.Dialog |
            Qt.WindowType.WindowStaysOnTopHint
        )

        # Main layout
        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        # Header with session info
        self._create_header(layout)

        # Entry form
        self._create_entry_form(layout)

        # Pending entries list
        self._create_entries_list(layout)

        # Meetings section
        self._create_meetings_section(layout)

        # Day summary (afternoon only)
        if self.session == SessionType.AFTERNOON:
            self._create_day_summary(layout)

        # Action buttons
        self._create_action_buttons(layout)

    def _create_header(self, layout: QVBoxLayout) -> None:
        """Create the header showing session info."""
        header = QFrame()
        header.setFrameShape(QFrame.Shape.StyledPanel)
        header_layout = QHBoxLayout(header)

        # Session label
        session_name = "Morning" if self.session == SessionType.MORNING else "Afternoon"
        title = QLabel(f"{session_name} Time Entry")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title.setFont(title_font)
        header_layout.addWidget(title)

        header_layout.addStretch()

        # Hours summary
        self._hours_label = QLabel("Loading...")
        header_layout.addWidget(self._hours_label)

        layout.addWidget(header)

    def _create_entry_form(self, layout: QVBoxLayout) -> None:
        """Create the entry form for adding new entries."""
        form_group = QGroupBox("Add Entry")
        form_layout = QVBoxLayout(form_group)

        # Project selector
        project_layout = QHBoxLayout()
        project_layout.addWidget(QLabel("Project:"))
        self._project_combo = QComboBox()
        self._project_combo.setMinimumWidth(300)
        self._project_combo.currentIndexChanged.connect(self._on_project_changed)
        project_layout.addWidget(self._project_combo)
        project_layout.addStretch()
        form_layout.addLayout(project_layout)

        # Activity description
        activity_layout = QHBoxLayout()
        activity_layout.addWidget(QLabel("Activity:"))
        self._activity_input = QLineEdit()
        self._activity_input.setPlaceholderText("What did you work on?")
        activity_layout.addWidget(self._activity_input)
        form_layout.addLayout(activity_layout)

        # Hours selector
        hours_layout = QHBoxLayout()
        hours_layout.addWidget(QLabel("Hours:"))
        self._hours_spin = QSpinBox()
        self._hours_spin.setRange(1, 8)
        self._hours_spin.setValue(1)
        hours_layout.addWidget(self._hours_spin)

        # Quick hour buttons
        for h in [1, 2, 3, 4]:
            btn = QPushButton(f"{h}h")
            btn.setFixedWidth(40)
            btn.clicked.connect(lambda checked, hours=h: self._hours_spin.setValue(hours))
            hours_layout.addWidget(btn)

        hours_layout.addStretch()

        # Add button
        self._add_btn = QPushButton("Add Entry")
        self._add_btn.clicked.connect(self._add_entry)
        hours_layout.addWidget(self._add_btn)

        form_layout.addLayout(hours_layout)

        layout.addWidget(form_group)

    def _create_entries_list(self, layout: QVBoxLayout) -> None:
        """Create the list of pending entries for this session."""
        entries_group = QGroupBox("Session Entries")
        entries_layout = QVBoxLayout(entries_group)

        self._entries_list = QListWidget()
        self._entries_list.setMinimumHeight(100)
        entries_layout.addWidget(self._entries_list)

        # Remove button
        remove_btn = QPushButton("Remove Selected")
        remove_btn.clicked.connect(self._remove_selected_entry)
        entries_layout.addWidget(remove_btn)

        layout.addWidget(entries_group)

    def _create_meetings_section(self, layout: QVBoxLayout) -> None:
        """Create the meetings section showing detected Outlook meetings."""
        meetings_group = QGroupBox("Detected Meetings")
        meetings_layout = QVBoxLayout(meetings_group)

        self._meetings_list = QListWidget()
        self._meetings_list.setMinimumHeight(100)
        meetings_layout.addWidget(self._meetings_list)

        # Import button
        import_layout = QHBoxLayout()
        import_btn = QPushButton("Import Selected Meetings")
        import_btn.clicked.connect(self._import_selected_meetings)
        import_layout.addWidget(import_btn)
        import_layout.addStretch()
        meetings_layout.addLayout(import_layout)

        layout.addWidget(meetings_group)

    def _create_day_summary(self, layout: QVBoxLayout) -> None:
        """Create the day summary section (afternoon only)."""
        summary_group = QGroupBox("Day Summary")
        summary_layout = QVBoxLayout(summary_group)

        self._day_summary_label = QLabel("Loading...")
        summary_layout.addWidget(self._day_summary_label)

        layout.addWidget(summary_group)

    def _create_action_buttons(self, layout: QVBoxLayout) -> None:
        """Create the action buttons at the bottom."""
        button_layout = QHBoxLayout()

        # Snooze button
        snooze_btn = QPushButton("Remind in 15min")
        snooze_btn.clicked.connect(self._snooze)
        button_layout.addWidget(snooze_btn)

        button_layout.addStretch()

        # Export button (afternoon only)
        if self.session == SessionType.AFTERNOON:
            export_btn = QPushButton("Export to Excel")
            export_btn.clicked.connect(self._export)
            button_layout.addWidget(export_btn)

        # Save button
        save_btn = QPushButton("Save && Close")
        save_btn.setDefault(True)
        save_btn.clicked.connect(self._save_and_close)
        button_layout.addWidget(save_btn)

        layout.addLayout(button_layout)

    def _load_data(self) -> None:
        """Load data and populate the UI."""
        # Load projects
        self._load_projects()

        # Load session info
        self._update_session_info()

        # Load meetings
        self._load_meetings()

        # Load existing entries
        self._load_existing_entries()

        # Update day summary if afternoon
        if self.session == SessionType.AFTERNOON:
            self._update_day_summary()

    def _load_projects(self) -> None:
        """Load projects into the combo box."""
        self._project_combo.clear()

        # Add placeholder
        self._project_combo.addItem("-- Select Project --", None)

        # Get recent projects first
        recent = self._store.get_recent_projects(limit=5)
        if recent:
            for rp in recent:
                display = f"[Recent] {rp.project_name}"
                self._project_combo.addItem(display, rp.project_id)

            # Add separator
            self._project_combo.insertSeparator(len(recent) + 1)

        # Get all projects
        projects = self._store.get_projects(active_only=True)
        for project in projects:
            # Skip if already in recent
            recent_ids = [r.project_id for r in recent]
            if project.id not in recent_ids:
                self._project_combo.addItem(project.name, project.id)

    def _update_session_info(self) -> None:
        """Update the session hours info."""
        info = self._hour_service.get_session_info(
            self.target_date,
            self.session
        )

        # Update header label
        self._hours_label.setText(
            f"Target: {info['target_hours']}h | "
            f"Logged: {info['logged_hours']}h | "
            f"Remaining: {info['remaining_hours']}h"
        )

        # Update suggested hours
        suggested = self._hour_service.suggest_hours(self.target_date, self.session)
        self._hours_spin.setValue(suggested)

    def _load_meetings(self) -> None:
        """Load meetings from the database."""
        self._meetings_list.clear()

        meetings = self._store.get_meetings_for_date(self.target_date)

        # Filter by session
        config = self._store.get_schedule_config()
        if self.session == SessionType.MORNING:
            cutoff_hour = int(config.lunch_start.split(":")[0])
        else:
            cutoff_hour = int(config.lunch_end.split(":")[0])

        for meeting in meetings:
            meeting_hour = meeting.start_datetime.hour

            # Check if in current session
            if self.session == SessionType.MORNING and meeting_hour >= cutoff_hour:
                continue
            if self.session == SessionType.AFTERNOON and meeting_hour < cutoff_hour:
                continue

            # Create list item
            item = QListWidgetItem()
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Unchecked if meeting.is_imported else Qt.CheckState.Checked)

            # Format display
            display = f"{meeting.display_time} - {meeting.subject} ({meeting.display_duration})"
            if meeting.matched_jira_key:
                display += f" [{meeting.matched_jira_key}]"
            if meeting.is_imported:
                display += " [Imported]"
                item.setCheckState(Qt.CheckState.Unchecked)

            item.setText(display)
            item.setData(Qt.ItemDataRole.UserRole, meeting.id)

            self._meetings_list.addItem(item)

    def _load_existing_entries(self) -> None:
        """Load existing entries for this date/session."""
        entries = self._store.get_entries(
            start_date=self.target_date,
            end_date=self.target_date
        )

        for entry in entries:
            self._add_entry_to_list(entry)

    def _update_day_summary(self) -> None:
        """Update the day summary (afternoon only)."""
        summary = self._hour_service.get_day_summary(self.target_date)

        text = (
            f"Morning: {summary['morning']['logged_hours']}h / {summary['morning']['target_hours']}h\n"
            f"Afternoon: {summary['afternoon']['logged_hours']}h / {summary['afternoon']['target_hours']}h\n"
            f"Total: {summary['total_hours']}h / {summary['total_target']}h"
        )

        if summary['is_complete']:
            text += "\n\n✓ Day complete!"

        self._day_summary_label.setText(text)

    def _on_project_changed(self, index: int) -> None:
        """Handle project selection change."""
        project_id = self._project_combo.currentData()
        if project_id:
            project = self._store.get_project(project_id)
            if project and project.jira_key:
                # Pre-fill activity with JIRA key
                current = self._activity_input.text()
                if not current:
                    self._activity_input.setText(f"{project.jira_key}: ")

    def _add_entry(self) -> None:
        """Add a new entry to the pending list."""
        project_id = self._project_combo.currentData()
        if not project_id:
            QMessageBox.warning(self, "Error", "Please select a project.")
            return

        activity = self._activity_input.text().strip()
        if not activity:
            QMessageBox.warning(self, "Error", "Please enter an activity description.")
            return

        hours = self._hours_spin.value()

        project = self._store.get_project(project_id)
        if not project:
            return

        # Create entry
        settings = self._store.get_settings()
        entry = TimesheetEntry(
            consultant_id=settings.consultant_id,
            entry_date=self.target_date,
            hours=hours,
            ticket_number=project.ticket_number,
            project_name=project.name,
            activity_code=project.default_activity_code,
            location=settings.default_location,
            description=activity,
            status=EntryStatus.DRAFT,
            source=EntrySource.MANUAL,
        )

        # Save to database
        self._store.save_entry(entry)

        # Update recent projects
        self._store.update_recent_project(project)

        # Add to list
        self._add_entry_to_list(entry)

        # Clear form
        self._activity_input.clear()
        self._project_combo.setCurrentIndex(0)

        # Update session info
        self._update_session_info()

        # Update day summary
        if self.session == SessionType.AFTERNOON:
            self._update_day_summary()

    def _add_entry_to_list(self, entry: TimesheetEntry) -> None:
        """Add an entry to the entries list widget."""
        item = QListWidgetItem()
        display = f"{entry.project_name} | {entry.hours}h | {entry.description[:50]}"
        item.setText(display)
        item.setData(Qt.ItemDataRole.UserRole, entry.id)
        self._entries_list.addItem(item)

    def _remove_selected_entry(self) -> None:
        """Remove the selected entry from the list."""
        current = self._entries_list.currentItem()
        if not current:
            return

        entry_id = current.data(Qt.ItemDataRole.UserRole)
        if entry_id:
            self._store.delete_entry(entry_id)

        self._entries_list.takeItem(self._entries_list.row(current))
        self._update_session_info()

        if self.session == SessionType.AFTERNOON:
            self._update_day_summary()

    def _import_selected_meetings(self) -> None:
        """Import selected meetings as entries."""
        settings = self._store.get_settings()
        imported_count = 0

        for i in range(self._meetings_list.count()):
            item = self._meetings_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                meeting_id = item.data(Qt.ItemDataRole.UserRole)
                meetings = self._store.get_meetings_for_date(self.target_date)
                meeting = next((m for m in meetings if m.id == meeting_id), None)

                if meeting and not meeting.is_imported:
                    # Get project if matched
                    project = None
                    if meeting.matched_project_id:
                        project = self._store.get_project(meeting.matched_project_id)

                    # Create entry
                    hours = max(1, round(float(meeting.duration_hours)))
                    entry = TimesheetEntry(
                        consultant_id=settings.consultant_id,
                        entry_date=self.target_date,
                        hours=hours,
                        ticket_number=project.ticket_number if project else None,
                        project_name=project.name if project else "Unknown Project",
                        activity_code=ActivityCode.TPLNT,  # Meetings
                        location=settings.default_location,
                        description=meeting.subject,
                        status=EntryStatus.DRAFT,
                        source=EntrySource.CALENDAR,
                    )

                    self._store.save_entry(entry)
                    self._store.mark_meeting_imported(meeting_id, entry.id)
                    self._add_entry_to_list(entry)
                    imported_count += 1

        # Refresh meetings list
        self._load_meetings()
        self._update_session_info()

        if self.session == SessionType.AFTERNOON:
            self._update_day_summary()

        if imported_count > 0:
            QMessageBox.information(
                self,
                "Import Complete",
                f"Imported {imported_count} meeting(s) as entries."
            )

    def _snooze(self) -> None:
        """Snooze the popup for 15 minutes."""
        if self._scheduler:
            if self.session == SessionType.MORNING:
                self._scheduler.snooze_morning()
            else:
                self._scheduler.snooze_afternoon()

        self.close()

    def _export(self) -> None:
        """Trigger export."""
        self.export_requested.emit()

    def _save_and_close(self) -> None:
        """Save all entries and close."""
        # Count entries
        count = self._entries_list.count()
        if count > 0:
            self.entries_saved.emit(count)

        self.close()
