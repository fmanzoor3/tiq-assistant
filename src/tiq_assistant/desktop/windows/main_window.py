"""Main window for TIQ Assistant desktop app with all functionality."""

from datetime import date, timedelta
from typing import Optional
from pathlib import Path
import tempfile

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QTabWidget,
    QLabel, QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QLineEdit, QSpinBox, QComboBox, QFormLayout, QGroupBox,
    QMessageBox, QFileDialog, QDateEdit, QTextEdit, QCheckBox,
    QSplitter, QFrame, QScrollArea, QSizePolicy
)
from PyQt6.QtCore import Qt, QDate
from PyQt6.QtGui import QFont

from tiq_assistant.core.models import (
    Project, TimesheetEntry, ActivityCode, EntryStatus, EntrySource, OutlookMeeting
)
from tiq_assistant.storage.sqlite_store import get_store
from tiq_assistant.services.matching_service import get_matching_service
from tiq_assistant.services.timesheet_service import get_timesheet_service
from tiq_assistant.integrations.outlook_reader import get_outlook_reader, OutlookNotAvailableError
from tiq_assistant.exporters.excel_exporter import (
    ExcelExporter, get_monthly_export_path
)


class MainWindow(QMainWindow):
    """Main application window with all TIQ Assistant functionality."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)

        self._store = get_store()
        self._matching_service = get_matching_service()
        self._timesheet_service = get_timesheet_service()
        self._outlook_meetings: list[OutlookMeeting] = []

        self._setup_ui()
        self._load_data()

    def _setup_ui(self) -> None:
        """Set up the main window UI."""
        self.setWindowTitle("TIQ Assistant")
        self.setMinimumSize(900, 700)

        # Central widget with tabs
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # Tab widget
        self._tabs = QTabWidget()
        layout.addWidget(self._tabs)

        # Create tabs
        self._tabs.addTab(self._create_dashboard_tab(), "Dashboard")
        self._tabs.addTab(self._create_projects_tab(), "Projects")
        self._tabs.addTab(self._create_timesheet_tab(), "Timesheet")
        self._tabs.addTab(self._create_import_tab(), "Calendar Import")
        self._tabs.addTab(self._create_settings_tab(), "Settings")

    # ==================== DASHBOARD TAB ====================

    def _create_dashboard_tab(self) -> QWidget:
        """Create the dashboard tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Stats section
        stats_layout = QHBoxLayout()

        # Active Projects
        self._projects_stat = self._create_stat_card("Active Projects", "0")
        stats_layout.addWidget(self._projects_stat)

        # Hours This Week
        self._hours_stat = self._create_stat_card("Hours This Week", "0")
        stats_layout.addWidget(self._hours_stat)

        # Draft Entries
        self._drafts_stat = self._create_stat_card("Draft Entries", "0")
        stats_layout.addWidget(self._drafts_stat)

        layout.addLayout(stats_layout)

        # Recent entries section
        layout.addWidget(QLabel("Recent Entries (Last 7 Days)"))

        self._recent_table = QTableWidget()
        self._recent_table.setColumnCount(6)
        self._recent_table.setHorizontalHeaderLabels([
            "Date", "Project", "Hours", "Activity", "Description", "Status"
        ])
        self._recent_table.horizontalHeader().setSectionResizeMode(
            4, QHeaderView.ResizeMode.Stretch
        )
        layout.addWidget(self._recent_table)

        # Refresh button
        refresh_btn = QPushButton("Refresh")
        refresh_btn.clicked.connect(self._refresh_dashboard)
        layout.addWidget(refresh_btn)

        return widget

    def _create_stat_card(self, title: str, value: str) -> QFrame:
        """Create a stat card widget."""
        frame = QFrame()
        frame.setFrameShape(QFrame.Shape.StyledPanel)
        frame.setMinimumHeight(80)

        layout = QVBoxLayout(frame)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        value_label = QLabel(value)
        value_label.setObjectName(f"stat_value_{title.replace(' ', '_')}")
        value_font = QFont()
        value_font.setPointSize(24)
        value_font.setBold(True)
        value_label.setFont(value_font)
        value_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title_label = QLabel(title)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        layout.addWidget(value_label)
        layout.addWidget(title_label)

        return frame

    def _refresh_dashboard(self) -> None:
        """Refresh dashboard data."""
        # Update stats
        projects = self._store.get_projects()
        today = date.today()
        week_start = today - timedelta(days=today.weekday())
        week_entries = self._store.get_entries(start_date=week_start, end_date=today)

        # Find and update stat labels
        self._projects_stat.findChild(
            QLabel, "stat_value_Active_Projects"
        ).setText(str(len(projects)))

        week_hours = sum(e.hours for e in week_entries)
        self._hours_stat.findChild(
            QLabel, "stat_value_Hours_This_Week"
        ).setText(str(week_hours))

        draft_count = len([e for e in week_entries if e.status == EntryStatus.DRAFT])
        self._drafts_stat.findChild(
            QLabel, "stat_value_Draft_Entries"
        ).setText(str(draft_count))

        # Update recent entries table
        recent = self._store.get_entries(
            start_date=today - timedelta(days=7),
            end_date=today
        )

        self._recent_table.setRowCount(len(recent))
        for i, entry in enumerate(recent):
            self._recent_table.setItem(i, 0, QTableWidgetItem(
                entry.entry_date.strftime("%d.%m.%Y")
            ))
            self._recent_table.setItem(i, 1, QTableWidgetItem(
                entry.project_name or "-"
            ))
            self._recent_table.setItem(i, 2, QTableWidgetItem(str(entry.hours)))
            self._recent_table.setItem(i, 3, QTableWidgetItem(
                entry.activity_code.value
            ))
            self._recent_table.setItem(i, 4, QTableWidgetItem(
                entry.description[:50] + "..." if len(entry.description) > 50 else entry.description
            ))
            self._recent_table.setItem(i, 5, QTableWidgetItem(entry.status.value))

    # ==================== PROJECTS TAB ====================

    def _create_projects_tab(self) -> QWidget:
        """Create the projects management tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Add project form
        form_group = QGroupBox("Add New Project")
        form_layout = QFormLayout(form_group)

        self._project_name_input = QLineEdit()
        self._project_name_input.setPlaceholderText("BI BÜYÜK VERI PLATFORM SUPPORT")
        form_layout.addRow("Project Name *:", self._project_name_input)

        self._ticket_number_input = QLineEdit()
        self._ticket_number_input.setPlaceholderText("2019135")
        form_layout.addRow("Ticket No *:", self._ticket_number_input)

        self._jira_key_input = QLineEdit()
        self._jira_key_input.setPlaceholderText("PEMP-948 (optional)")
        form_layout.addRow("JIRA Key:", self._jira_key_input)

        self._keywords_input = QLineEdit()
        self._keywords_input.setPlaceholderText("Agent Bot, big data (comma-separated)")
        form_layout.addRow("Keywords:", self._keywords_input)

        add_project_btn = QPushButton("Add Project")
        add_project_btn.clicked.connect(self._add_project)
        form_layout.addRow("", add_project_btn)

        layout.addWidget(form_group)

        # Projects table
        layout.addWidget(QLabel("Existing Projects"))

        self._projects_table = QTableWidget()
        self._projects_table.setColumnCount(6)
        self._projects_table.setHorizontalHeaderLabels([
            "Name", "Ticket No", "JIRA Key", "Keywords", "Location", "Actions"
        ])
        self._projects_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch
        )
        layout.addWidget(self._projects_table)

        return widget

    def _add_project(self) -> None:
        """Add a new project."""
        name = self._project_name_input.text().strip()
        ticket = self._ticket_number_input.text().strip()

        if not name or not ticket:
            QMessageBox.warning(self, "Error", "Project Name and Ticket No are required.")
            return

        jira_key = self._jira_key_input.text().strip() or None
        keywords_text = self._keywords_input.text().strip()
        keywords = [k.strip() for k in keywords_text.split(",") if k.strip()]

        project = Project(
            name=name,
            ticket_number=ticket,
            jira_key=jira_key,
            keywords=keywords,
        )

        self._store.save_project(project)

        # Clear form
        self._project_name_input.clear()
        self._ticket_number_input.clear()
        self._jira_key_input.clear()
        self._keywords_input.clear()

        # Refresh table
        self._refresh_projects()

        QMessageBox.information(self, "Success", f"Project '{name}' added!")

    def _refresh_projects(self) -> None:
        """Refresh the projects table."""
        projects = self._store.get_projects()

        self._projects_table.setRowCount(len(projects))
        for i, project in enumerate(projects):
            self._projects_table.setItem(i, 0, QTableWidgetItem(project.name))
            self._projects_table.setItem(i, 1, QTableWidgetItem(project.ticket_number))
            self._projects_table.setItem(i, 2, QTableWidgetItem(project.jira_key or "-"))
            self._projects_table.setItem(i, 3, QTableWidgetItem(
                ", ".join(project.keywords) if project.keywords else "-"
            ))
            self._projects_table.setItem(i, 4, QTableWidgetItem(project.default_location))

            # Delete button
            delete_btn = QPushButton("Delete")
            delete_btn.clicked.connect(lambda checked, pid=project.id: self._delete_project(pid))
            self._projects_table.setCellWidget(i, 5, delete_btn)

    def _delete_project(self, project_id: str) -> None:
        """Delete a project."""
        reply = QMessageBox.question(
            self, "Confirm Delete",
            "Are you sure you want to delete this project?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self._store.delete_project(project_id)
            self._refresh_projects()

    # ==================== TIMESHEET TAB ====================

    def _create_timesheet_tab(self) -> QWidget:
        """Create the timesheet management tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Date range selector
        range_layout = QHBoxLayout()
        range_layout.addWidget(QLabel("From:"))

        self._start_date = QDateEdit()
        self._start_date.setDate(QDate.currentDate().addDays(-7))
        self._start_date.setCalendarPopup(True)
        range_layout.addWidget(self._start_date)

        range_layout.addWidget(QLabel("To:"))

        self._end_date = QDateEdit()
        self._end_date.setDate(QDate.currentDate())
        self._end_date.setCalendarPopup(True)
        range_layout.addWidget(self._end_date)

        load_btn = QPushButton("Load Entries")
        load_btn.clicked.connect(self._refresh_timesheet)
        range_layout.addWidget(load_btn)

        range_layout.addStretch()
        layout.addLayout(range_layout)

        # Entries table
        self._entries_table = QTableWidget()
        self._entries_table.setColumnCount(8)
        self._entries_table.setHorizontalHeaderLabels([
            "Date", "Project", "Ticket", "Hours", "Activity", "Description", "Status", "Actions"
        ])
        self._entries_table.horizontalHeader().setSectionResizeMode(
            5, QHeaderView.ResizeMode.Stretch
        )
        layout.addWidget(self._entries_table)

        # Add entry form
        add_group = QGroupBox("Add Manual Entry")
        add_layout = QFormLayout(add_group)

        self._entry_date = QDateEdit()
        self._entry_date.setDate(QDate.currentDate())
        self._entry_date.setCalendarPopup(True)
        add_layout.addRow("Date:", self._entry_date)

        self._entry_project = QComboBox()
        add_layout.addRow("Project:", self._entry_project)

        self._entry_hours = QSpinBox()
        self._entry_hours.setRange(1, 24)
        self._entry_hours.setValue(1)
        add_layout.addRow("Hours:", self._entry_hours)

        self._entry_activity = QComboBox()
        for code in ActivityCode:
            self._entry_activity.addItem(code.value, code)
        add_layout.addRow("Activity Code:", self._entry_activity)

        self._entry_description = QLineEdit()
        self._entry_description.setPlaceholderText("What did you work on?")
        add_layout.addRow("Description:", self._entry_description)

        add_entry_btn = QPushButton("Add Entry")
        add_entry_btn.clicked.connect(self._add_manual_entry)
        add_layout.addRow("", add_entry_btn)

        layout.addWidget(add_group)

        # Export section
        export_layout = QHBoxLayout()
        export_layout.addStretch()

        export_btn = QPushButton("Export to Excel")
        export_btn.clicked.connect(self._export_entries)
        export_layout.addWidget(export_btn)

        layout.addLayout(export_layout)

        return widget

    def _refresh_timesheet(self) -> None:
        """Refresh the timesheet entries table."""
        start = self._start_date.date().toPyDate()
        end = self._end_date.date().toPyDate()

        entries = self._store.get_entries(start_date=start, end_date=end)

        self._entries_table.setRowCount(len(entries))
        for i, entry in enumerate(entries):
            self._entries_table.setItem(i, 0, QTableWidgetItem(
                entry.entry_date.strftime("%d.%m.%Y")
            ))
            self._entries_table.setItem(i, 1, QTableWidgetItem(entry.project_name or "-"))
            self._entries_table.setItem(i, 2, QTableWidgetItem(entry.ticket_number or "-"))
            self._entries_table.setItem(i, 3, QTableWidgetItem(str(entry.hours)))
            self._entries_table.setItem(i, 4, QTableWidgetItem(entry.activity_code.value))
            self._entries_table.setItem(i, 5, QTableWidgetItem(
                entry.description[:40] + "..." if len(entry.description) > 40 else entry.description
            ))
            self._entries_table.setItem(i, 6, QTableWidgetItem(entry.status.value))

            # Delete button
            delete_btn = QPushButton("Delete")
            delete_btn.clicked.connect(lambda checked, eid=entry.id: self._delete_entry(eid))
            self._entries_table.setCellWidget(i, 7, delete_btn)

        # Also refresh project dropdown for add form
        self._entry_project.clear()
        self._entry_project.addItem("-- Select Project --", None)
        for project in self._store.get_projects():
            self._entry_project.addItem(project.name, project.id)

    def _add_manual_entry(self) -> None:
        """Add a manual timesheet entry."""
        project_id = self._entry_project.currentData()
        description = self._entry_description.text().strip()

        if not description:
            QMessageBox.warning(self, "Error", "Description is required.")
            return

        project = self._store.get_project(project_id) if project_id else None
        settings = self._store.get_settings()

        entry = TimesheetEntry(
            consultant_id=settings.consultant_id,
            entry_date=self._entry_date.date().toPyDate(),
            hours=self._entry_hours.value(),
            ticket_number=project.ticket_number if project else None,
            project_name=project.name if project else None,
            activity_code=self._entry_activity.currentData(),
            location=settings.default_location,
            description=description,
            status=EntryStatus.DRAFT,
            source=EntrySource.MANUAL,
        )

        self._store.save_entry(entry)

        # Update recent project
        if project:
            self._store.update_recent_project(project)

        # Clear and refresh
        self._entry_description.clear()
        self._refresh_timesheet()

        QMessageBox.information(self, "Success", "Entry added!")

    def _delete_entry(self, entry_id: str) -> None:
        """Delete a timesheet entry."""
        reply = QMessageBox.question(
            self, "Confirm Delete",
            "Are you sure you want to delete this entry?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self._store.delete_entry(entry_id)
            self._refresh_timesheet()

    def _export_entries(self) -> None:
        """Export entries to Excel."""
        start = self._start_date.date().toPyDate()
        end = self._end_date.date().toPyDate()

        entries = self._store.get_entries(start_date=start, end_date=end)

        if not entries:
            QMessageBox.information(self, "No Entries", "No entries to export.")
            return

        # Get export path
        export_path = get_monthly_export_path()

        exporter = ExcelExporter()
        if export_path.exists():
            exporter.append_to_existing(entries, export_path)
        else:
            exporter.export_to_new_file(entries, export_path)

        # Mark as exported
        self._store.mark_entries_exported([e.id for e in entries])

        self._refresh_timesheet()

        QMessageBox.information(
            self, "Export Complete",
            f"Exported {len(entries)} entries to:\n{export_path}"
        )

    # ==================== CALENDAR IMPORT TAB ====================

    def _create_import_tab(self) -> QWidget:
        """Create the calendar import tab with Outlook COM integration."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Date range section
        date_group = QGroupBox("Select Date Range")
        date_layout = QHBoxLayout(date_group)

        date_layout.addWidget(QLabel("From:"))
        self._import_start_date = QDateEdit()
        self._import_start_date.setCalendarPopup(True)
        self._import_start_date.setDate(QDate.currentDate().addDays(-7))
        date_layout.addWidget(self._import_start_date)

        date_layout.addWidget(QLabel("To:"))
        self._import_end_date = QDateEdit()
        self._import_end_date.setCalendarPopup(True)
        self._import_end_date.setDate(QDate.currentDate())
        date_layout.addWidget(self._import_end_date)

        # Quick date buttons
        today_btn = QPushButton("Today")
        today_btn.clicked.connect(self._set_import_today)
        date_layout.addWidget(today_btn)

        week_btn = QPushButton("This Week")
        week_btn.clicked.connect(self._set_import_this_week)
        date_layout.addWidget(week_btn)

        month_btn = QPushButton("This Month")
        month_btn.clicked.connect(self._set_import_this_month)
        date_layout.addWidget(month_btn)

        date_layout.addStretch()
        layout.addWidget(date_group)

        # Fetch button
        fetch_btn = QPushButton("Fetch Meetings from Outlook")
        fetch_btn.setStyleSheet("font-weight: bold; padding: 10px;")
        fetch_btn.clicked.connect(self._fetch_outlook_meetings)
        layout.addWidget(fetch_btn)

        # Status label
        self._import_status = QLabel("Select a date range and click 'Fetch Meetings from Outlook'")
        self._import_status.setStyleSheet("color: gray; font-style: italic;")
        layout.addWidget(self._import_status)

        # Events table
        layout.addWidget(QLabel("Fetched Meetings"))

        self._events_table = QTableWidget()
        self._events_table.setColumnCount(8)
        self._events_table.setHorizontalHeaderLabels([
            "Select", "Date", "Time", "Subject", "Hours", "Project", "Description", "Add"
        ])
        self._events_table.horizontalHeader().setSectionResizeMode(
            3, QHeaderView.ResizeMode.Stretch
        )
        self._events_table.horizontalHeader().setSectionResizeMode(
            6, QHeaderView.ResizeMode.Stretch
        )
        layout.addWidget(self._events_table)

        # Bottom buttons
        btn_layout = QHBoxLayout()

        add_selected_btn = QPushButton("Add Selected Meetings")
        add_selected_btn.clicked.connect(self._add_selected_meetings)
        btn_layout.addWidget(add_selected_btn)

        select_all_btn = QPushButton("Select All")
        select_all_btn.clicked.connect(self._select_all_meetings)
        btn_layout.addWidget(select_all_btn)

        deselect_all_btn = QPushButton("Deselect All")
        deselect_all_btn.clicked.connect(self._deselect_all_meetings)
        btn_layout.addWidget(deselect_all_btn)

        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        return widget

    def _set_import_today(self) -> None:
        """Set import date range to today."""
        today = QDate.currentDate()
        self._import_start_date.setDate(today)
        self._import_end_date.setDate(today)

    def _set_import_this_week(self) -> None:
        """Set import date range to this week (Mon-Fri)."""
        today = QDate.currentDate()
        # Go back to Monday
        days_since_monday = today.dayOfWeek() - 1
        monday = today.addDays(-days_since_monday)
        friday = monday.addDays(4)
        self._import_start_date.setDate(monday)
        self._import_end_date.setDate(friday)

    def _set_import_this_month(self) -> None:
        """Set import date range to this month."""
        today = QDate.currentDate()
        first_of_month = QDate(today.year(), today.month(), 1)
        self._import_start_date.setDate(first_of_month)
        self._import_end_date.setDate(today)

    def _fetch_outlook_meetings(self) -> None:
        """Fetch meetings from Outlook for the selected date range."""
        try:
            reader = get_outlook_reader()

            if not reader.is_available():
                QMessageBox.warning(
                    self, "Outlook Not Available",
                    "Could not connect to Outlook. Make sure Outlook desktop "
                    "(not the web version) is installed and has been opened at least once."
                )
                return

            start_date = self._import_start_date.date().toPyDate()
            end_date = self._import_end_date.date().toPyDate()

            if start_date > end_date:
                QMessageBox.warning(self, "Invalid Range", "Start date must be before end date.")
                return

            self._import_status.setText("Fetching meetings from Outlook...")
            self._import_status.setStyleSheet("color: blue;")

            # Fetch meetings
            meetings = reader.get_meetings_for_date_range(start_date, end_date)
            self._outlook_meetings = meetings

            # Match meetings to projects
            for meeting in meetings:
                event = reader.to_calendar_event(meeting)
                result = self._matching_service.match_event(event)
                meeting.matched_project_id = result.project_id
                meeting.matched_jira_key = result.ticket_jira_key
                meeting.match_confidence = result.confidence

            # Populate table
            self._populate_meetings_table()

            self._import_status.setText(
                f"Found {len(meetings)} meetings. "
                f"{len([m for m in meetings if m.match_confidence and m.match_confidence > 0])} matched to projects."
            )
            self._import_status.setStyleSheet("color: green;")

        except OutlookNotAvailableError as e:
            QMessageBox.warning(self, "Outlook Error", str(e))
            self._import_status.setText("Failed to connect to Outlook")
            self._import_status.setStyleSheet("color: red;")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to fetch meetings: {e}")
            self._import_status.setText(f"Error: {e}")
            self._import_status.setStyleSheet("color: red;")

    def _populate_meetings_table(self) -> None:
        """Populate the meetings table with fetched Outlook meetings."""
        self._events_table.setRowCount(len(self._outlook_meetings))

        projects = self._store.get_projects()
        project_map = {p.id: p for p in projects}

        for i, meeting in enumerate(self._outlook_meetings):
            # Checkbox - pre-select matched meetings
            checkbox = QCheckBox()
            checkbox.setChecked(meeting.match_confidence is not None and meeting.match_confidence > 0)
            self._events_table.setCellWidget(i, 0, checkbox)

            # Date
            self._events_table.setItem(i, 1, QTableWidgetItem(
                meeting.start_datetime.strftime("%d.%m.%Y")
            ))

            # Time
            self._events_table.setItem(i, 2, QTableWidgetItem(
                meeting.display_time
            ))

            # Subject (truncate if long)
            subject = meeting.subject
            if len(subject) > 40:
                subject = subject[:37] + "..."
            self._events_table.setItem(i, 3, QTableWidgetItem(subject))

            # Hours (editable spinner)
            hours_spin = QSpinBox()
            hours_spin.setRange(1, 8)
            hours_spin.setValue(max(1, round(meeting.duration_hours)))
            self._events_table.setCellWidget(i, 4, hours_spin)

            # Project dropdown (editable)
            project_combo = QComboBox()
            project_combo.addItem("-- Select --", None)
            selected_idx = 0
            for j, project in enumerate(projects):
                project_combo.addItem(project.name, project.id)
                if meeting.matched_project_id and project.id == meeting.matched_project_id:
                    selected_idx = j + 1
            project_combo.setCurrentIndex(selected_idx)
            self._events_table.setCellWidget(i, 5, project_combo)

            # Description (editable)
            desc_edit = QLineEdit()
            desc_edit.setText(meeting.subject)
            self._events_table.setCellWidget(i, 6, desc_edit)

            # Add single button
            add_btn = QPushButton("Add")
            add_btn.clicked.connect(lambda checked, idx=i: self._add_single_meeting(idx))
            self._events_table.setCellWidget(i, 7, add_btn)

    def _add_single_meeting(self, row: int) -> None:
        """Add a single meeting as a timesheet entry."""
        if row >= len(self._outlook_meetings):
            return

        meeting = self._outlook_meetings[row]

        # Get values from widgets
        hours_spin = self._events_table.cellWidget(row, 4)
        hours = hours_spin.value() if hours_spin else max(1, round(meeting.duration_hours))

        project_combo = self._events_table.cellWidget(row, 5)
        project_id = project_combo.currentData() if project_combo else None

        desc_edit = self._events_table.cellWidget(row, 6)
        description = desc_edit.text() if desc_edit else meeting.subject

        settings = self._store.get_settings()
        project = self._store.get_project(project_id) if project_id else None

        entry = TimesheetEntry(
            consultant_id=settings.consultant_id,
            entry_date=meeting.start_datetime.date(),
            hours=hours,
            ticket_number=project.ticket_number if project else None,
            project_name=project.name if project else None,
            activity_code=settings.meeting_activity_code,
            location=settings.default_location,
            description=description,
            status=EntryStatus.DRAFT,
            source=EntrySource.CALENDAR,
            source_event_id=meeting.id,
            source_jira_key=meeting.matched_jira_key,
        )

        self._store.save_entry(entry)

        # Disable the row
        checkbox = self._events_table.cellWidget(row, 0)
        if checkbox:
            checkbox.setEnabled(False)
            checkbox.setChecked(False)
        add_btn = self._events_table.cellWidget(row, 7)
        if add_btn:
            add_btn.setEnabled(False)

        QMessageBox.information(self, "Added", f"Added entry for: {meeting.subject[:40]}")

    def _add_selected_meetings(self) -> None:
        """Add all selected meetings as timesheet entries."""
        count = 0
        settings = self._store.get_settings()

        for i in range(self._events_table.rowCount()):
            checkbox = self._events_table.cellWidget(i, 0)
            if checkbox and checkbox.isChecked() and checkbox.isEnabled():
                if i < len(self._outlook_meetings):
                    meeting = self._outlook_meetings[i]

                    # Get values from widgets
                    hours_spin = self._events_table.cellWidget(i, 4)
                    hours = hours_spin.value() if hours_spin else max(1, round(meeting.duration_hours))

                    project_combo = self._events_table.cellWidget(i, 5)
                    project_id = project_combo.currentData() if project_combo else None

                    desc_edit = self._events_table.cellWidget(i, 6)
                    description = desc_edit.text() if desc_edit else meeting.subject

                    project = self._store.get_project(project_id) if project_id else None

                    entry = TimesheetEntry(
                        consultant_id=settings.consultant_id,
                        entry_date=meeting.start_datetime.date(),
                        hours=hours,
                        ticket_number=project.ticket_number if project else None,
                        project_name=project.name if project else None,
                        activity_code=settings.meeting_activity_code,
                        location=settings.default_location,
                        description=description,
                        status=EntryStatus.DRAFT,
                        source=EntrySource.CALENDAR,
                        source_event_id=meeting.id,
                        source_jira_key=meeting.matched_jira_key,
                    )

                    self._store.save_entry(entry)
                    checkbox.setEnabled(False)
                    checkbox.setChecked(False)
                    add_btn = self._events_table.cellWidget(i, 7)
                    if add_btn:
                        add_btn.setEnabled(False)
                    count += 1

        if count > 0:
            QMessageBox.information(self, "Added", f"Added {count} timesheet entries!")
        else:
            QMessageBox.information(self, "No Selection", "No meetings were selected.")

    def _select_all_meetings(self) -> None:
        """Select all meetings in the table."""
        for i in range(self._events_table.rowCount()):
            checkbox = self._events_table.cellWidget(i, 0)
            if checkbox and checkbox.isEnabled():
                checkbox.setChecked(True)

    def _deselect_all_meetings(self) -> None:
        """Deselect all meetings in the table."""
        for i in range(self._events_table.rowCount()):
            checkbox = self._events_table.cellWidget(i, 0)
            if checkbox and checkbox.isEnabled():
                checkbox.setChecked(False)

    # ==================== SETTINGS TAB ====================

    def _create_settings_tab(self) -> QWidget:
        """Create the settings tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # User settings
        user_group = QGroupBox("User Settings")
        user_layout = QFormLayout(user_group)

        self._consultant_id_input = QLineEdit()
        user_layout.addRow("Consultant ID:", self._consultant_id_input)

        self._location_input = QLineEdit()
        user_layout.addRow("Default Location:", self._location_input)

        layout.addWidget(user_group)

        # Activity codes
        activity_group = QGroupBox("Activity Codes")
        activity_layout = QFormLayout(activity_group)

        self._default_activity = QComboBox()
        for code in ActivityCode:
            self._default_activity.addItem(code.value, code)
        activity_layout.addRow("Default Activity:", self._default_activity)

        self._meeting_activity = QComboBox()
        for code in ActivityCode:
            self._meeting_activity.addItem(code.value, code)
        activity_layout.addRow("Meeting Activity:", self._meeting_activity)

        layout.addWidget(activity_group)

        # Matching settings
        matching_group = QGroupBox("Matching Settings")
        matching_layout = QFormLayout(matching_group)

        self._skip_canceled = QCheckBox("Skip Canceled Meetings")
        matching_layout.addRow("", self._skip_canceled)

        self._min_duration = QSpinBox()
        self._min_duration.setRange(5, 60)
        self._min_duration.setSuffix(" minutes")
        matching_layout.addRow("Min Meeting Duration:", self._min_duration)

        layout.addWidget(matching_group)

        # Save button
        save_btn = QPushButton("Save Settings")
        save_btn.clicked.connect(self._save_settings)
        layout.addWidget(save_btn)

        layout.addStretch()

        return widget

    def _load_settings(self) -> None:
        """Load settings into the form."""
        settings = self._store.get_settings()

        self._consultant_id_input.setText(settings.consultant_id)
        self._location_input.setText(settings.default_location)

        # Set activity code combo boxes
        for i in range(self._default_activity.count()):
            if self._default_activity.itemData(i) == settings.default_activity_code:
                self._default_activity.setCurrentIndex(i)
                break

        for i in range(self._meeting_activity.count()):
            if self._meeting_activity.itemData(i) == settings.meeting_activity_code:
                self._meeting_activity.setCurrentIndex(i)
                break

        self._skip_canceled.setChecked(settings.skip_canceled_meetings)
        self._min_duration.setValue(settings.min_meeting_duration_minutes)

    def _save_settings(self) -> None:
        """Save settings."""
        from tiq_assistant.core.models import UserSettings

        settings = UserSettings(
            consultant_id=self._consultant_id_input.text().strip() or "FMANZOOR",
            default_location=self._location_input.text().strip() or "ANKARA",
            default_activity_code=self._default_activity.currentData(),
            meeting_activity_code=self._meeting_activity.currentData(),
            skip_canceled_meetings=self._skip_canceled.isChecked(),
            min_meeting_duration_minutes=self._min_duration.value(),
        )

        self._store.save_settings(settings)
        QMessageBox.information(self, "Saved", "Settings saved!")

    # ==================== DATA LOADING ====================

    def _load_data(self) -> None:
        """Load initial data."""
        self._refresh_dashboard()
        self._refresh_projects()
        self._refresh_timesheet()
        self._load_settings()
