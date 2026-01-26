"""Dialog for entering timesheet entries for a specific day."""

from datetime import date, datetime, time, timedelta
from enum import Enum
from typing import Optional

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView, QSpinBox,
    QComboBox, QLineEdit, QGroupBox, QCheckBox, QMessageBox,
    QAbstractItemView, QFrame
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QBrush

from tiq_assistant.core.models import (
    Project, TimesheetEntry, ActivityCode, EntryStatus, EntrySource, OutlookMeeting
)
from tiq_assistant.storage.sqlite_store import get_store
from tiq_assistant.services.matching_service import MatchingService
from tiq_assistant.integrations.outlook_reader import get_outlook_reader, OutlookNotAvailableError
from tiq_assistant.desktop.icon import create_app_icon


class SessionType(Enum):
    """Type of entry session."""
    FULL_DAY = "full_day"
    MORNING = "morning"
    AFTERNOON = "afternoon"


class DayEntryDialog(QDialog):
    """
    Dialog for entering timesheet entries for a specific day.

    Can be used for:
    - Manual day selection from workday overview
    - 12:15 morning popup (morning session)
    - 18:15 afternoon popup (afternoon session)
    """

    # Color scheme matching main window
    COLORS = {
        'primary': '#0078D4',
        'success': '#107C10',
        'success_light': '#DFF6DD',
        'warning': '#FFB900',
        'warning_light': '#FFF4CE',
        'danger': '#D13438',
        'danger_light': '#FDE7E9',
        'gray': '#E1E1E1',
        'gray_light': '#F5F5F5',
        'text': '#323130',
        'text_secondary': '#605E5C',
    }

    # Session time boundaries
    MORNING_END = time(12, 15)
    AFTERNOON_START = time(13, 30)

    def __init__(
        self,
        target_date: date,
        session: SessionType = SessionType.FULL_DAY,
        outlook_meetings: Optional[list[OutlookMeeting]] = None,
        parent=None
    ):
        super().__init__(parent)

        self._target_date = target_date
        self._session = session
        self._store = get_store()
        self._matching_service = MatchingService(self._store)

        # Filter meetings for this day and session
        self._meetings = self._filter_meetings(outlook_meetings or [])

        # Calculate target hours based on session
        self._target_hours = self._get_target_hours()

        self._setup_ui()
        self._load_data()

    def _get_target_hours(self) -> int:
        """Get target hours based on session type."""
        if self._session == SessionType.MORNING:
            return 3
        elif self._session == SessionType.AFTERNOON:
            return 5
        else:
            return 8

    def _filter_meetings(self, all_meetings: list[OutlookMeeting]) -> list[OutlookMeeting]:
        """Filter meetings for this day and session."""
        day_meetings = [
            m for m in all_meetings
            if m.start_datetime.date() == self._target_date
        ]

        if self._session == SessionType.MORNING:
            return [
                m for m in day_meetings
                if m.start_datetime.time() < self.MORNING_END
            ]
        elif self._session == SessionType.AFTERNOON:
            return [
                m for m in day_meetings
                if m.start_datetime.time() >= self.AFTERNOON_START
            ]
        else:
            return day_meetings

    def _setup_ui(self) -> None:
        """Set up the dialog UI."""
        # Window properties
        session_label = {
            SessionType.FULL_DAY: "",
            SessionType.MORNING: " - Morning",
            SessionType.AFTERNOON: " - Afternoon",
        }[self._session]

        day_name = self._target_date.strftime("%A")
        date_str = self._target_date.strftime("%d %B %Y")
        self.setWindowTitle(f"Time Entry: {day_name}, {date_str}{session_label}")
        self.setWindowIcon(create_app_icon())
        self.setMinimumSize(700, 500)
        self.setModal(True)

        # Apply stylesheet
        self.setStyleSheet(f"""
            QDialog {{
                background-color: white;
                color: {self.COLORS['text']};
            }}
            QWidget {{
                background-color: white;
                color: {self.COLORS['text']};
            }}
            QLabel {{
                background-color: transparent;
                color: {self.COLORS['text']};
            }}
            QGroupBox {{
                font-weight: bold;
                color: {self.COLORS['text']};
                background-color: white;
                border: 1px solid {self.COLORS['gray']};
                border-radius: 4px;
                margin-top: 12px;
                padding-top: 8px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
                color: {self.COLORS['text']};
            }}
            QGroupBox QLabel {{
                background-color: transparent;
                color: {self.COLORS['text']};
                font-weight: normal;
            }}
            QTableWidget {{
                border: 1px solid {self.COLORS['gray']};
                gridline-color: {self.COLORS['gray']};
                background-color: white;
                color: {self.COLORS['text']};
            }}
            QTableWidget::item {{
                padding: 4px;
                color: {self.COLORS['text']};
            }}
            QHeaderView::section {{
                background-color: {self.COLORS['gray_light']};
                color: {self.COLORS['text']};
                padding: 6px;
                border: none;
                border-right: 1px solid {self.COLORS['gray']};
                border-bottom: 1px solid {self.COLORS['gray']};
                font-weight: bold;
            }}
            QComboBox {{
                color: {self.COLORS['text']};
                background-color: white;
                padding: 4px 8px;
                border: 1px solid {self.COLORS['gray']};
                border-radius: 4px;
            }}
            QComboBox QAbstractItemView {{
                color: {self.COLORS['text']};
                background-color: white;
                selection-background-color: {self.COLORS['primary']};
                selection-color: white;
            }}
            QLineEdit {{
                color: {self.COLORS['text']};
                background-color: white;
                padding: 4px 8px;
                border: 1px solid {self.COLORS['gray']};
                border-radius: 4px;
            }}
            QSpinBox {{
                color: {self.COLORS['text']};
                background-color: white;
                padding: 4px 8px;
                border: 1px solid {self.COLORS['gray']};
                border-radius: 4px;
            }}
            QCheckBox {{
                color: {self.COLORS['text']};
                background-color: transparent;
            }}
            QPushButton {{
                color: {self.COLORS['text']};
                background-color: white;
                padding: 6px 12px;
                border: 1px solid {self.COLORS['gray']};
                border-radius: 4px;
            }}
            QPushButton:hover {{
                background-color: {self.COLORS['gray_light']};
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        # Header with date and progress
        self._create_header(layout)

        # Entries section
        self._create_entries_section(layout)

        # Meetings section (if we have meetings)
        self._create_meetings_section(layout)

        # Add entry form
        self._create_add_entry_form(layout)

        # Bottom buttons
        self._create_bottom_buttons(layout)

    def _create_header(self, layout: QVBoxLayout) -> None:
        """Create the header with date and progress info."""
        header_layout = QHBoxLayout()

        # Date and day
        day_name = self._target_date.strftime("%A")
        date_str = self._target_date.strftime("%d.%m.%Y")
        date_label = QLabel(f"<b>{day_name}, {date_str}</b>")
        date_label.setStyleSheet("font-size: 16px;")
        header_layout.addWidget(date_label)

        # Session badge
        if self._session != SessionType.FULL_DAY:
            session_text = "Morning" if self._session == SessionType.MORNING else "Afternoon"
            session_color = self.COLORS['primary']
            badge = QLabel(session_text)
            badge.setStyleSheet(f"""
                background-color: {session_color};
                color: white;
                padding: 4px 12px;
                border-radius: 4px;
                font-weight: bold;
            """)
            header_layout.addWidget(badge)

        header_layout.addStretch()

        # Progress indicator
        self._progress_label = QLabel()
        self._progress_label.setStyleSheet(f"color: {self.COLORS['text_secondary']};")
        header_layout.addWidget(self._progress_label)

        layout.addLayout(header_layout)

        # Separator
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet(f"background-color: {self.COLORS['gray']};")
        layout.addWidget(line)

    def _create_entries_section(self, layout: QVBoxLayout) -> None:
        """Create the entries table section."""
        group = QGroupBox("Entries")
        group_layout = QVBoxLayout(group)

        self._entries_table = QTableWidget()
        self._entries_table.setColumnCount(6)
        self._entries_table.setHorizontalHeaderLabels([
            "Project", "Ticket", "Hours", "Activity", "Description", "Actions"
        ])
        self._entries_table.horizontalHeader().setSectionResizeMode(
            4, QHeaderView.ResizeMode.Stretch
        )
        self._entries_table.verticalHeader().setVisible(False)
        self._entries_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._entries_table.setMaximumHeight(150)
        group_layout.addWidget(self._entries_table)

        layout.addWidget(group)

    def _create_meetings_section(self, layout: QVBoxLayout) -> None:
        """Create the Outlook meetings section."""
        self._meetings_group = QGroupBox(f"Outlook Meetings ({len(self._meetings)})")
        group_layout = QVBoxLayout(self._meetings_group)

        if not self._meetings:
            no_meetings_label = QLabel("No meetings for this day/session")
            no_meetings_label.setStyleSheet(f"color: {self.COLORS['text_secondary']}; font-style: italic;")
            group_layout.addWidget(no_meetings_label)
        else:
            self._meetings_table = QTableWidget()
            self._meetings_table.setColumnCount(7)
            self._meetings_table.setHorizontalHeaderLabels([
                "Select", "Time", "Subject", "Hours", "Activity", "Project", "Add"
            ])
            self._meetings_table.horizontalHeader().setSectionResizeMode(
                2, QHeaderView.ResizeMode.Stretch
            )
            self._meetings_table.verticalHeader().setVisible(False)
            self._meetings_table.verticalHeader().setDefaultSectionSize(36)
            self._meetings_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
            self._meetings_table.setMaximumHeight(180)
            group_layout.addWidget(self._meetings_table)

            # Add Selected button
            btn_layout = QHBoxLayout()
            add_selected_btn = QPushButton("Add Selected Meetings")
            add_selected_btn.setStyleSheet(f"""
                background-color: {self.COLORS['primary']};
                color: white;
                border: none;
                padding: 8px 16px;
                font-weight: bold;
            """)
            add_selected_btn.clicked.connect(self._add_selected_meetings)
            btn_layout.addWidget(add_selected_btn)
            btn_layout.addStretch()
            group_layout.addLayout(btn_layout)

        layout.addWidget(self._meetings_group)

    def _create_add_entry_form(self, layout: QVBoxLayout) -> None:
        """Create the add entry form."""
        group = QGroupBox("Add Entry")
        form_layout = QHBoxLayout(group)

        # Project dropdown
        form_layout.addWidget(QLabel("Project:"))
        self._project_combo = QComboBox()
        self._project_combo.setMinimumWidth(150)
        form_layout.addWidget(self._project_combo)

        # Hours
        form_layout.addWidget(QLabel("Hours:"))
        self._hours_spin = QSpinBox()
        self._hours_spin.setRange(1, 8)
        self._hours_spin.setValue(1)
        form_layout.addWidget(self._hours_spin)

        # Activity
        form_layout.addWidget(QLabel("Activity:"))
        self._activity_combo = QComboBox()
        for code in ActivityCode:
            self._activity_combo.addItem(code.value, code)
        form_layout.addWidget(self._activity_combo)

        # Description
        form_layout.addWidget(QLabel("Description:"))
        self._description_edit = QLineEdit()
        self._description_edit.setPlaceholderText("Description...")
        self._description_edit.setMinimumWidth(200)
        form_layout.addWidget(self._description_edit)

        # Add button
        add_btn = QPushButton("Add")
        add_btn.setStyleSheet(f"""
            background-color: {self.COLORS['primary']};
            color: white;
            border: none;
            padding: 8px 16px;
            font-weight: bold;
        """)
        add_btn.clicked.connect(self._add_manual_entry)
        form_layout.addWidget(add_btn)

        layout.addWidget(group)

    def _create_bottom_buttons(self, layout: QVBoxLayout) -> None:
        """Create the bottom action buttons."""
        btn_layout = QHBoxLayout()

        # Snooze button (only for scheduled popups)
        if self._session != SessionType.FULL_DAY:
            snooze_btn = QPushButton("Remind in 15 min")
            snooze_btn.setStyleSheet(f"""
                background-color: {self.COLORS['warning']};
                color: white;
                border: none;
                padding: 10px 20px;
                font-weight: bold;
            """)
            snooze_btn.clicked.connect(self._snooze)
            btn_layout.addWidget(snooze_btn)

        btn_layout.addStretch()

        # Close button
        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(f"""
            background-color: {self.COLORS['gray']};
            color: {self.COLORS['text']};
            border: none;
            padding: 10px 20px;
            font-weight: bold;
        """)
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)

    def _load_data(self) -> None:
        """Load data into the dialog."""
        self._load_projects()
        self._refresh_entries()
        self._populate_meetings_table()
        self._update_progress()

    def _load_projects(self) -> None:
        """Load projects into the dropdown."""
        self._project_combo.clear()
        self._project_combo.addItem("-- Select Project --", None)

        projects = self._store.get_projects()
        recent = self._store.get_recent_projects(limit=5)
        recent_ids = {rp.project_id for rp in recent}

        # Add recent projects first
        if recent:
            for rp in recent:
                self._project_combo.addItem(f"★ {rp.project_name}", rp.project_id)

        # Add all other projects
        for project in projects:
            if project.id not in recent_ids:
                self._project_combo.addItem(project.name, project.id)

    def _refresh_entries(self) -> None:
        """Refresh the entries table."""
        entries = self._store.get_entries(start_date=self._target_date, end_date=self._target_date)

        # Filter by session if needed
        if self._session == SessionType.MORNING:
            # For morning session, we might want to show only morning entries
            # But for simplicity, show all entries for the day
            pass

        self._entries_table.setRowCount(len(entries))

        for i, entry in enumerate(entries):
            self._entries_table.setItem(i, 0, QTableWidgetItem(entry.project_name or "-"))
            self._entries_table.setItem(i, 1, QTableWidgetItem(entry.ticket_number or "-"))
            self._entries_table.setItem(i, 2, QTableWidgetItem(str(entry.hours)))
            self._entries_table.setItem(i, 3, QTableWidgetItem(entry.activity_code.value))

            desc = entry.description
            if len(desc) > 40:
                desc = desc[:37] + "..."
            self._entries_table.setItem(i, 4, QTableWidgetItem(desc))

            # Delete button
            delete_btn = QPushButton("Delete")
            delete_btn.setStyleSheet(f"""
                background-color: {self.COLORS['danger']};
                color: white;
                border: none;
                padding: 4px 8px;
            """)
            delete_btn.clicked.connect(lambda checked, eid=entry.id: self._delete_entry(eid))
            self._entries_table.setCellWidget(i, 5, delete_btn)

        self._update_progress()

    def _populate_meetings_table(self) -> None:
        """Populate the meetings table."""
        if not self._meetings or not hasattr(self, '_meetings_table'):
            return

        projects = self._store.get_projects()

        self._meetings_table.setRowCount(len(self._meetings))

        for i, meeting in enumerate(self._meetings):
            is_matched = meeting.match_confidence is not None and meeting.match_confidence > 0

            # Checkbox
            checkbox = QCheckBox()
            checkbox.setChecked(is_matched)
            self._meetings_table.setCellWidget(i, 0, checkbox)

            # Time
            time_str = meeting.start_datetime.strftime("%H:%M")
            self._meetings_table.setItem(i, 1, QTableWidgetItem(time_str))

            # Subject
            subject = meeting.subject
            if len(subject) > 35:
                subject = subject[:32] + "..."
            self._meetings_table.setItem(i, 2, QTableWidgetItem(subject))

            # Hours spinner
            hours_spin = QSpinBox()
            hours_spin.setRange(1, 8)
            hours_spin.setValue(max(1, round(meeting.duration_hours)))
            self._meetings_table.setCellWidget(i, 3, hours_spin)

            # Activity dropdown
            activity_combo = QComboBox()
            for code in ActivityCode:
                activity_combo.addItem(code.value, code)
            # Default to TPLNT for meetings
            tplnt_idx = [j for j, code in enumerate(ActivityCode) if code == ActivityCode.TPLNT]
            if tplnt_idx:
                activity_combo.setCurrentIndex(tplnt_idx[0])
            self._meetings_table.setCellWidget(i, 4, activity_combo)

            # Project dropdown
            project_combo = QComboBox()
            project_combo.addItem("-- Select --", None)
            selected_idx = 0
            for j, project in enumerate(projects):
                project_combo.addItem(project.name, project.id)
                if meeting.matched_project_id and project.id == meeting.matched_project_id:
                    selected_idx = j + 1
            project_combo.setCurrentIndex(selected_idx)
            self._meetings_table.setCellWidget(i, 5, project_combo)

            # Add single button
            add_btn = QPushButton("Add")
            add_btn.setStyleSheet(f"""
                background-color: {self.COLORS['primary']};
                color: white;
                border: none;
                padding: 4px 8px;
            """)
            add_btn.clicked.connect(lambda checked, idx=i: self._add_single_meeting(idx))
            self._meetings_table.setCellWidget(i, 6, add_btn)

            # Highlight matched rows
            if is_matched:
                for col in range(self._meetings_table.columnCount()):
                    item = self._meetings_table.item(i, col)
                    if item:
                        item.setBackground(QBrush(QColor(self.COLORS['success_light'])))

    def _update_progress(self) -> None:
        """Update the progress indicator."""
        entries = self._store.get_entries(start_date=self._target_date, end_date=self._target_date)
        filled_hours = sum(e.hours for e in entries)
        remaining = max(0, self._target_hours - filled_hours)

        if filled_hours >= self._target_hours:
            self._progress_label.setText(f"✓ {filled_hours}h / {self._target_hours}h - Complete!")
            self._progress_label.setStyleSheet(f"color: {self.COLORS['success']}; font-weight: bold;")
        else:
            self._progress_label.setText(f"{filled_hours}h / {self._target_hours}h ({remaining}h remaining)")
            self._progress_label.setStyleSheet(f"color: {self.COLORS['text_secondary']};")

    def _add_manual_entry(self) -> None:
        """Add a manual entry."""
        project_id = self._project_combo.currentData()
        description = self._description_edit.text().strip()

        if not description:
            QMessageBox.warning(self, "Error", "Description is required.")
            return

        project = self._store.get_project(project_id) if project_id else None
        settings = self._store.get_settings()

        entry = TimesheetEntry(
            consultant_id=settings.consultant_id,
            entry_date=self._target_date,
            hours=self._hours_spin.value(),
            ticket_number=project.ticket_number if project else None,
            project_name=project.name if project else None,
            activity_code=self._activity_combo.currentData(),
            location=settings.default_location,
            description=description,
            status=EntryStatus.DRAFT,
            source=EntrySource.MANUAL,
        )

        self._store.save_entry(entry)

        if project:
            self._store.update_recent_project(project)

        self._description_edit.clear()
        self._refresh_entries()

    def _add_single_meeting(self, row: int) -> None:
        """Add a single meeting as an entry."""
        if row >= len(self._meetings):
            return

        meeting = self._meetings[row]

        # Get values from widgets
        hours_spin = self._meetings_table.cellWidget(row, 3)
        hours = hours_spin.value() if hours_spin else max(1, round(meeting.duration_hours))

        activity_combo = self._meetings_table.cellWidget(row, 4)
        activity_code = activity_combo.currentData() if activity_combo else ActivityCode.TPLNT

        project_combo = self._meetings_table.cellWidget(row, 5)
        project_id = project_combo.currentData() if project_combo else None

        settings = self._store.get_settings()
        project = self._store.get_project(project_id) if project_id else None

        entry = TimesheetEntry(
            consultant_id=settings.consultant_id,
            entry_date=self._target_date,
            hours=hours,
            ticket_number=project.ticket_number if project else None,
            project_name=project.name if project else None,
            activity_code=activity_code,
            location=settings.default_location,
            description=meeting.subject,
            status=EntryStatus.DRAFT,
            source=EntrySource.CALENDAR,
            source_event_id=meeting.id,
            source_jira_key=meeting.matched_jira_key,
        )

        self._store.save_entry(entry)

        if project:
            self._store.update_recent_project(project)

        self._refresh_entries()

        # Disable the row after adding
        for col in range(self._meetings_table.columnCount()):
            widget = self._meetings_table.cellWidget(row, col)
            if widget:
                widget.setEnabled(False)
            item = self._meetings_table.item(row, col)
            if item:
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEnabled)

    def _add_selected_meetings(self) -> None:
        """Add all selected meetings as entries."""
        added_count = 0

        for row in range(self._meetings_table.rowCount()):
            checkbox = self._meetings_table.cellWidget(row, 0)
            if checkbox and checkbox.isChecked():
                # Check if already disabled (already added)
                add_btn = self._meetings_table.cellWidget(row, 6)
                if add_btn and add_btn.isEnabled():
                    self._add_single_meeting(row)
                    added_count += 1

        if added_count > 0:
            QMessageBox.information(
                self, "Meetings Added",
                f"Added {added_count} meeting(s) as timesheet entries."
            )

    def _delete_entry(self, entry_id: str) -> None:
        """Delete an entry."""
        reply = QMessageBox.question(
            self, "Confirm Delete",
            "Are you sure you want to delete this entry?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self._store.delete_entry(entry_id)
            self._refresh_entries()

    def _snooze(self) -> None:
        """Snooze the popup for 15 minutes."""
        # This will be handled by the caller
        self.done(2)  # Custom return code for snooze

    def get_snooze_requested(self) -> bool:
        """Check if snooze was requested."""
        return self.result() == 2
