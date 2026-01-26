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
    QSplitter, QFrame, QScrollArea, QSizePolicy, QAbstractItemView
)
from PyQt6.QtCore import Qt, QDate
from PyQt6.QtGui import QFont, QColor, QBrush

from tiq_assistant.core.models import (
    Project, TimesheetEntry, ActivityCode, EntryStatus, EntrySource, OutlookMeeting
)
from tiq_assistant.core.holidays import get_holiday_service, HolidayType
from tiq_assistant.storage.sqlite_store import get_store
from tiq_assistant.services.matching_service import get_matching_service
from tiq_assistant.services.timesheet_service import get_timesheet_service
from tiq_assistant.integrations.outlook_reader import get_outlook_reader, OutlookNotAvailableError
from tiq_assistant.exporters.excel_exporter import (
    ExcelExporter, get_monthly_export_path
)


class MainWindow(QMainWindow):
    """Main application window with all TIQ Assistant functionality."""

    # Color scheme
    COLORS = {
        'primary': '#0078D4',        # Microsoft blue
        'primary_hover': '#106EBE',
        'success': '#107C10',        # Green
        'success_light': '#DFF6DD',
        'warning': '#FFB900',        # Yellow/amber
        'warning_light': '#FFF4CE',
        'danger': '#D13438',         # Red
        'danger_light': '#FDE7E9',
        'gray_light': '#F3F3F3',
        'gray': '#E1E1E1',
        'text': '#323130',
        'text_secondary': '#605E5C',
    }

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)

        self._store = get_store()
        self._matching_service = get_matching_service()
        self._timesheet_service = get_timesheet_service()
        self._outlook_meetings: list[OutlookMeeting] = []

        self._setup_ui()
        self._apply_styles()
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

    def _apply_styles(self) -> None:
        """Apply global stylesheet to the application."""
        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: white;
                color: {self.COLORS['text']};
            }}
            QWidget {{
                background-color: white;
                color: {self.COLORS['text']};
            }}
            QLabel {{
                color: {self.COLORS['text']};
                background-color: transparent;
            }}
            QTabWidget::pane {{
                border: 1px solid {self.COLORS['gray']};
                background-color: white;
            }}
            QTabBar::tab {{
                padding: 8px 16px;
                margin-right: 2px;
                background-color: {self.COLORS['gray_light']};
                color: {self.COLORS['text']};
                border: 1px solid {self.COLORS['gray']};
                border-bottom: none;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }}
            QTabBar::tab:selected {{
                background-color: white;
                color: {self.COLORS['text']};
                border-bottom: 2px solid {self.COLORS['primary']};
            }}
            QTabBar::tab:hover:!selected {{
                background-color: {self.COLORS['gray']};
            }}
            QGroupBox {{
                font-weight: bold;
                color: {self.COLORS['text']};
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
            QTableWidget {{
                border: 1px solid {self.COLORS['gray']};
                gridline-color: {self.COLORS['gray']};
                background-color: white;
                color: {self.COLORS['text']};
                selection-background-color: {self.COLORS['primary']};
                selection-color: white;
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
            QPushButton {{
                padding: 6px 12px;
                border: 1px solid {self.COLORS['gray']};
                border-radius: 4px;
                background-color: white;
                color: {self.COLORS['text']};
            }}
            QPushButton:hover {{
                background-color: {self.COLORS['gray_light']};
            }}
            QPushButton:pressed {{
                background-color: {self.COLORS['gray']};
            }}
            QPushButton[primary="true"] {{
                background-color: {self.COLORS['primary']};
                color: white;
                border: none;
            }}
            QPushButton[primary="true"]:hover {{
                background-color: {self.COLORS['primary_hover']};
            }}
            QPushButton[danger="true"] {{
                background-color: {self.COLORS['danger']};
                color: white;
                border: none;
            }}
            QPushButton[danger="true"]:hover {{
                background-color: #C50F1F;
            }}
            QLineEdit, QSpinBox, QComboBox, QDateEdit {{
                padding: 6px;
                border: 1px solid {self.COLORS['gray']};
                border-radius: 4px;
                background-color: white;
                color: {self.COLORS['text']};
            }}
            QLineEdit:focus, QSpinBox:focus, QComboBox:focus, QDateEdit:focus {{
                border-color: {self.COLORS['primary']};
            }}
            QComboBox QAbstractItemView {{
                background-color: white;
                color: {self.COLORS['text']};
                selection-background-color: {self.COLORS['primary']};
                selection-color: white;
            }}
            QCheckBox {{
                color: {self.COLORS['text']};
                background-color: transparent;
            }}
            QFrame {{
                background-color: white;
                color: {self.COLORS['text']};
            }}
        """)

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
        self._style_table(self._recent_table)
        layout.addWidget(self._recent_table)

        # Refresh button
        refresh_btn = self._create_primary_button("Refresh")
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
            # Status badge
            self._recent_table.setCellWidget(i, 5, self._create_status_badge(entry.status))

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

        add_project_btn = self._create_primary_button("Add Project")
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
        self._style_table(self._projects_table)
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
            delete_btn = self._create_danger_button("Delete")
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

        # Month selector
        month_layout = QHBoxLayout()
        month_layout.addWidget(QLabel("Month:"))

        self._timesheet_month = QComboBox()
        self._populate_month_selector(self._timesheet_month)
        self._timesheet_month.currentIndexChanged.connect(self._refresh_timesheet)
        month_layout.addWidget(self._timesheet_month)

        month_layout.addStretch()

        # Summary label
        self._timesheet_summary = QLabel("")
        self._timesheet_summary.setStyleSheet("font-weight: bold; margin-left: 20px;")
        month_layout.addWidget(self._timesheet_summary)

        layout.addLayout(month_layout)

        # Workday overview section
        workday_group = QGroupBox("Workday Overview")
        workday_layout = QVBoxLayout(workday_group)

        # Progress summary
        self._workday_progress = QLabel("")
        self._workday_progress.setStyleSheet(f"font-size: 13px; color: {self.COLORS['text']};")
        workday_layout.addWidget(self._workday_progress)

        # Workday table showing each day
        self._workday_table = QTableWidget()
        self._workday_table.setColumnCount(6)
        self._workday_table.setHorizontalHeaderLabels([
            "Date", "Day", "Expected", "Filled", "Remaining", "Status"
        ])
        self._workday_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self._workday_table.setMaximumHeight(200)
        # Disable built-in selection to manage it manually with colors
        self._workday_table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self._workday_table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self._workday_table.cellClicked.connect(self._on_workday_clicked)
        self._style_table(self._workday_table)
        workday_layout.addWidget(self._workday_table)

        # Track selected row index and row statuses for coloring
        self._selected_workday_row: int = -1
        self._workday_row_colors: dict[int, str] = {}  # row -> base background color

        layout.addWidget(workday_group)

        # Entries section (shown when a day is selected)
        self._entries_group = QGroupBox("Entries for Selected Day")
        entries_layout = QVBoxLayout(self._entries_group)

        self._selected_day_label = QLabel("Click a day above to view/add entries")
        self._selected_day_label.setStyleSheet(f"color: {self.COLORS['text_secondary']}; font-style: italic;")
        entries_layout.addWidget(self._selected_day_label)

        self._entries_table = QTableWidget()
        self._entries_table.setColumnCount(7)
        self._entries_table.setHorizontalHeaderLabels([
            "Project", "Ticket", "Hours", "Activity", "Description", "Status", "Actions"
        ])
        self._entries_table.horizontalHeader().setSectionResizeMode(
            4, QHeaderView.ResizeMode.Stretch  # Description column
        )
        self._entries_table.setMaximumHeight(150)
        self._style_table(self._entries_table)
        self._entries_table.setVisible(False)  # Hidden until a day is selected
        entries_layout.addWidget(self._entries_table)

        # Add entry form (inline, shown when day selected)
        self._add_entry_widget = QWidget()
        add_layout = QHBoxLayout(self._add_entry_widget)
        add_layout.setContentsMargins(0, 5, 0, 0)

        self._entry_date = QDateEdit()
        self._entry_date.setDate(QDate.currentDate())
        self._entry_date.setCalendarPopup(True)
        self._entry_date.setVisible(False)  # Hidden, set programmatically

        self._entry_project = QComboBox()
        self._entry_project.setMinimumWidth(200)
        add_layout.addWidget(QLabel("Project:"))
        add_layout.addWidget(self._entry_project)

        self._entry_hours = QSpinBox()
        self._entry_hours.setRange(1, 8)
        self._entry_hours.setValue(1)
        add_layout.addWidget(QLabel("Hours:"))
        add_layout.addWidget(self._entry_hours)

        self._entry_activity = QComboBox()
        for code in ActivityCode:
            self._entry_activity.addItem(code.value, code)
        add_layout.addWidget(QLabel("Activity:"))
        add_layout.addWidget(self._entry_activity)

        self._entry_description = QLineEdit()
        self._entry_description.setPlaceholderText("Description...")
        self._entry_description.setMinimumWidth(150)
        add_layout.addWidget(self._entry_description)

        add_entry_btn = self._create_primary_button("Add")
        add_entry_btn.clicked.connect(self._add_manual_entry)
        add_layout.addWidget(add_entry_btn)

        self._add_entry_widget.setVisible(False)  # Hidden until a day is selected
        entries_layout.addWidget(self._add_entry_widget)

        layout.addWidget(self._entries_group)

        # Export section
        export_layout = QHBoxLayout()
        export_layout.addStretch()

        export_btn = self._create_primary_button("Export Month to Excel")
        export_btn.clicked.connect(self._export_entries)
        export_layout.addWidget(export_btn)

        layout.addLayout(export_layout)

        # Track currently selected date
        self._selected_workday: date | None = None

        return widget

    def _refresh_timesheet(self) -> None:
        """Refresh the workday overview for the selected month."""
        # Get date range from month selector
        month_data = self._timesheet_month.currentData()
        if month_data:
            start, end = month_data
        else:
            # Fallback to current month
            today = date.today()
            start = date(today.year, today.month, 1)
            if today.month == 12:
                end = date(today.year + 1, 1, 1) - timedelta(days=1)
            else:
                end = date(today.year, today.month + 1, 1) - timedelta(days=1)

        entries = self._store.get_entries(start_date=start, end_date=end)

        # Update summary
        total_hours = sum(e.hours for e in entries)
        self._timesheet_summary.setText(f"Total: {len(entries)} entries, {total_hours} hours")

        # Update workday overview
        self._refresh_workday_overview(start, entries)

        # Clear selected day when month changes
        self._selected_workday = None
        self._selected_workday_row = -1
        self._entries_table.setRowCount(0)
        self._entries_table.setVisible(False)
        self._add_entry_widget.setVisible(False)
        self._selected_day_label.setText("Click a day above to view/add entries")
        self._selected_day_label.setStyleSheet(f"color: {self.COLORS['text_secondary']}; font-style: italic;")
        self._selected_day_label.setVisible(True)

        # Refresh project dropdown for add form
        self._entry_project.clear()
        self._entry_project.addItem("-- Select Project --", None)
        for project in self._store.get_projects():
            self._entry_project.addItem(project.name, project.id)

    def _refresh_workday_overview(self, month_start: date, entries: list) -> None:
        """Refresh the workday overview table showing expected vs filled hours."""
        holiday_service = get_holiday_service()

        # Get workdays for the month
        workdays = holiday_service.get_workdays_in_month(month_start.year, month_start.month)

        # Calculate filled hours per day
        hours_by_date: dict[date, int] = {}
        for entry in entries:
            if entry.entry_date not in hours_by_date:
                hours_by_date[entry.entry_date] = 0
            hours_by_date[entry.entry_date] += entry.hours

        # Day names in Turkish
        day_names = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]

        # Stats for progress summary
        total_expected = 0
        total_filled = 0
        days_complete = 0
        days_incomplete = 0

        self._workday_table.setRowCount(len(workdays))
        self._workday_row_colors.clear()

        for i, (work_date, expected_hours) in enumerate(workdays):
            filled_hours = hours_by_date.get(work_date, 0)
            remaining = max(0, expected_hours - filled_hours)

            total_expected += expected_hours
            total_filled += filled_hours

            if filled_hours >= expected_hours:
                days_complete += 1
            else:
                days_incomplete += 1

            # Date
            date_item = QTableWidgetItem(work_date.strftime("%d.%m.%Y"))
            date_item.setData(Qt.ItemDataRole.UserRole, work_date)
            self._workday_table.setItem(i, 0, date_item)

            # Day name
            day_name = day_names[work_date.weekday()]
            day_item = QTableWidgetItem(day_name)
            self._workday_table.setItem(i, 1, day_item)

            # Check if it's a holiday (half-day)
            holiday = holiday_service.get_holiday(work_date)
            if holiday and holiday.holiday_type == HolidayType.HALF_DAY:
                day_item.setText(f"{day_name} (Yarım gün)")

            # Expected hours
            expected_item = QTableWidgetItem(f"{expected_hours}h")
            self._workday_table.setItem(i, 2, expected_item)

            # Filled hours
            filled_item = QTableWidgetItem(f"{filled_hours}h")
            self._workday_table.setItem(i, 3, filled_item)

            # Remaining hours
            remaining_item = QTableWidgetItem(f"{remaining}h" if remaining > 0 else "-")
            self._workday_table.setItem(i, 4, remaining_item)

            # Determine status and row color
            if filled_hours >= expected_hours:
                status_item = QTableWidgetItem("✓ Complete")
                row_color = "#C8E6C9"  # Stronger green
                self._workday_row_colors[i] = row_color
            elif filled_hours > 0:
                status_item = QTableWidgetItem(f"Partial ({remaining}h left)")
                row_color = "#FFF9C4"  # Stronger yellow
                self._workday_row_colors[i] = row_color
            elif work_date < date.today():
                status_item = QTableWidgetItem("Missing")
                row_color = "#FFCDD2"  # Stronger red
                self._workday_row_colors[i] = row_color
            else:
                status_item = QTableWidgetItem("Pending")
                row_color = "#FFFFFF"  # White for pending
                self._workday_row_colors[i] = row_color

            self._workday_table.setItem(i, 5, status_item)

            # Apply row background color
            self._set_row_background(self._workday_table, i, row_color)

        # Update progress summary
        remaining_total = max(0, total_expected - total_filled)
        progress_text = (
            f"Progress: {total_filled}h / {total_expected}h expected  |  "
            f"{days_complete} days complete, {days_incomplete} remaining  |  "
            f"{remaining_total}h left to fill"
        )
        self._workday_progress.setText(progress_text)

    def _on_workday_clicked(self, row: int, col: int) -> None:
        """Handle workday row click - show entries for that day."""
        # Get the date from the first column
        date_item = self._workday_table.item(row, 0)
        if not date_item:
            return

        selected_date = date_item.data(Qt.ItemDataRole.UserRole)
        if not selected_date:
            return

        # Update visual selection indicator
        self._highlight_selected_workday(row)

        self._selected_workday = selected_date
        self._selected_workday_row = row
        self._entry_date.setDate(QDate(selected_date.year, selected_date.month, selected_date.day))

        # Get entries for this date
        entries = self._store.get_entries(start_date=selected_date, end_date=selected_date)

        # Calculate suggested hours based on remaining
        holiday_service = get_holiday_service()
        expected = holiday_service.get_expected_hours(selected_date)
        filled = sum(e.hours for e in entries)
        remaining = max(1, expected - filled)

        self._entry_hours.setValue(min(remaining, 8))

        # Update the selected day label
        day_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        day_name = day_names[selected_date.weekday()]
        status = "Complete" if filled >= expected else f"{remaining}h remaining"
        self._selected_day_label.setText(
            f"{selected_date.strftime('%d.%m.%Y')} ({day_name}) - {filled}h / {expected}h ({status})"
        )
        self._selected_day_label.setStyleSheet(f"font-weight: bold; color: {self.COLORS['text']};")

        # Show entries table and add form
        self._entries_table.setVisible(True)
        self._add_entry_widget.setVisible(True)

        # Populate entries table for this day
        self._entries_table.setRowCount(len(entries))
        for i, entry in enumerate(entries):
            self._entries_table.setItem(i, 0, QTableWidgetItem(entry.project_name or "-"))
            self._entries_table.setItem(i, 1, QTableWidgetItem(entry.ticket_number or "-"))
            self._entries_table.setItem(i, 2, QTableWidgetItem(str(entry.hours)))
            self._entries_table.setItem(i, 3, QTableWidgetItem(entry.activity_code.value))
            self._entries_table.setItem(i, 4, QTableWidgetItem(
                entry.description[:50] + "..." if len(entry.description) > 50 else entry.description
            ))
            # Status badge
            self._entries_table.setCellWidget(i, 5, self._create_status_badge(entry.status))

            # Delete button
            delete_btn = self._create_danger_button("Delete")
            delete_btn.clicked.connect(lambda checked, eid=entry.id: self._delete_entry(eid))
            self._entries_table.setCellWidget(i, 6, delete_btn)

    def _highlight_selected_workday(self, selected_row: int) -> None:
        """Highlight the selected row with visual indicators."""
        # Selection highlight color (blue tint)
        selection_color = "#BBDEFB"  # Light blue for selection

        for row in range(self._workday_table.rowCount()):
            date_item = self._workday_table.item(row, 0)
            if not date_item:
                continue

            work_date = date_item.data(Qt.ItemDataRole.UserRole)
            if not work_date:
                continue

            is_selected = (row == selected_row)

            # Update date text with arrow indicator
            if is_selected:
                date_item.setText(f"► {work_date.strftime('%d.%m.%Y')}")
            else:
                date_item.setText(work_date.strftime("%d.%m.%Y"))

            # Set font weight
            font = date_item.font()
            font.setBold(is_selected)

            # Apply to all columns
            for col in range(self._workday_table.columnCount()):
                item = self._workday_table.item(row, col)
                if item:
                    item.setFont(font)

            # Apply background color - use selection color for selected row,
            # otherwise use the stored status color
            if is_selected:
                self._set_row_background(self._workday_table, row, selection_color)
            else:
                base_color = self._workday_row_colors.get(row, "#FFFFFF")
                self._set_row_background(self._workday_table, row, base_color)

    def _add_manual_entry(self) -> None:
        """Add a manual timesheet entry."""
        if self._selected_workday is None:
            QMessageBox.warning(self, "Error", "Please select a day first.")
            return

        project_id = self._entry_project.currentData()
        description = self._entry_description.text().strip()

        if not description:
            QMessageBox.warning(self, "Error", "Description is required.")
            return

        project = self._store.get_project(project_id) if project_id else None
        settings = self._store.get_settings()

        entry = TimesheetEntry(
            consultant_id=settings.consultant_id,
            entry_date=self._selected_workday,
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

        # Clear description and refresh
        self._entry_description.clear()

        # Refresh the workday overview (to update filled hours)
        self._refresh_timesheet_keeping_selection()

    def _refresh_timesheet_keeping_selection(self) -> None:
        """Refresh timesheet but keep the currently selected day."""
        saved_date = self._selected_workday

        # Get date range from month selector
        month_data = self._timesheet_month.currentData()
        if month_data:
            start, end = month_data
        else:
            today = date.today()
            start = date(today.year, today.month, 1)
            if today.month == 12:
                end = date(today.year + 1, 1, 1) - timedelta(days=1)
            else:
                end = date(today.year, today.month + 1, 1) - timedelta(days=1)

        entries = self._store.get_entries(start_date=start, end_date=end)

        # Update summary
        total_hours = sum(e.hours for e in entries)
        self._timesheet_summary.setText(f"Total: {len(entries)} entries, {total_hours} hours")

        # Update workday overview
        self._refresh_workday_overview(start, entries)

        # Restore selected day and refresh its entries
        if saved_date:
            self._selected_workday = saved_date
            # Find the row for the saved date
            for row in range(self._workday_table.rowCount()):
                item = self._workday_table.item(row, 0)
                if item and item.data(Qt.ItemDataRole.UserRole) == saved_date:
                    self._selected_workday_row = row
                    # Manually trigger click handler to update display
                    self._on_workday_clicked(row, 0)
                    break

    def _delete_entry(self, entry_id: str) -> None:
        """Delete a timesheet entry."""
        reply = QMessageBox.question(
            self, "Confirm Delete",
            "Are you sure you want to delete this entry?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self._store.delete_entry(entry_id)
            self._refresh_timesheet_keeping_selection()

    def _export_entries(self) -> None:
        """Export entries to Excel."""
        # Get date range from month selector
        month_data = self._timesheet_month.currentData()
        if month_data:
            start, end = month_data
        else:
            today = date.today()
            start = date(today.year, today.month, 1)
            if today.month == 12:
                end = date(today.year + 1, 1, 1) - timedelta(days=1)
            else:
                end = date(today.year, today.month + 1, 1) - timedelta(days=1)

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

        # Month selector section
        date_group = QGroupBox("Select Period")
        date_layout = QHBoxLayout(date_group)

        date_layout.addWidget(QLabel("Month:"))
        self._import_month = QComboBox()
        self._populate_month_selector(self._import_month, include_custom=True)
        self._import_month.currentIndexChanged.connect(self._on_import_month_changed)
        date_layout.addWidget(self._import_month)

        # Custom date range (initially hidden)
        self._custom_range_widget = QWidget()
        custom_layout = QHBoxLayout(self._custom_range_widget)
        custom_layout.setContentsMargins(0, 0, 0, 0)

        custom_layout.addWidget(QLabel("From:"))
        self._import_start_date = QDateEdit()
        self._import_start_date.setCalendarPopup(True)
        self._import_start_date.setDate(QDate.currentDate().addDays(-7))
        custom_layout.addWidget(self._import_start_date)

        custom_layout.addWidget(QLabel("To:"))
        self._import_end_date = QDateEdit()
        self._import_end_date.setCalendarPopup(True)
        self._import_end_date.setDate(QDate.currentDate())
        custom_layout.addWidget(self._import_end_date)

        self._custom_range_widget.setVisible(False)
        date_layout.addWidget(self._custom_range_widget)

        date_layout.addStretch()
        layout.addWidget(date_group)

        # Fetch button
        fetch_btn = self._create_primary_button("Fetch Meetings from Outlook")
        fetch_btn.clicked.connect(self._fetch_outlook_meetings)
        layout.addWidget(fetch_btn)

        # Status label
        self._import_status = QLabel("Select a date range and click 'Fetch Meetings from Outlook'")
        self._import_status.setStyleSheet("color: gray; font-style: italic;")
        layout.addWidget(self._import_status)

        # Events table
        layout.addWidget(QLabel("Fetched Meetings"))

        self._events_table = QTableWidget()
        self._events_table.setColumnCount(9)
        self._events_table.setHorizontalHeaderLabels([
            "Select", "Date", "Time", "Subject", "Hours", "Activity", "Project", "Description", "Add"
        ])
        self._events_table.horizontalHeader().setSectionResizeMode(
            3, QHeaderView.ResizeMode.Stretch
        )
        self._events_table.horizontalHeader().setSectionResizeMode(
            7, QHeaderView.ResizeMode.Stretch  # Description column
        )
        self._events_table.verticalHeader().setVisible(False)
        layout.addWidget(self._events_table)

        # Bottom buttons
        btn_layout = QHBoxLayout()

        add_selected_btn = self._create_primary_button("Add Selected Meetings")
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

    def _on_import_month_changed(self) -> None:
        """Handle month selector change - show/hide custom date range."""
        current_text = self._import_month.currentText()
        is_custom = current_text == "Custom Range..."
        self._custom_range_widget.setVisible(is_custom)

    def _get_import_date_range(self) -> tuple[date, date]:
        """Get the date range for calendar import based on selector."""
        current_text = self._import_month.currentText()

        if current_text == "Custom Range...":
            return (
                self._import_start_date.date().toPyDate(),
                self._import_end_date.date().toPyDate()
            )

        month_data = self._import_month.currentData()
        if month_data:
            return month_data

        # Fallback to current month
        today = date.today()
        start = date(today.year, today.month, 1)
        if today.month == 12:
            end = date(today.year + 1, 1, 1) - timedelta(days=1)
        else:
            end = date(today.year, today.month + 1, 1) - timedelta(days=1)
        return start, end

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

            start_date, end_date = self._get_import_date_range()

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
            is_matched = meeting.match_confidence is not None and meeting.match_confidence > 0

            # Checkbox - pre-select matched meetings
            checkbox = QCheckBox()
            checkbox.setChecked(is_matched)
            self._events_table.setCellWidget(i, 0, checkbox)

            # Date
            date_item = QTableWidgetItem(meeting.start_datetime.strftime("%d.%m.%Y"))
            self._events_table.setItem(i, 1, date_item)

            # Time
            time_item = QTableWidgetItem(meeting.display_time)
            self._events_table.setItem(i, 2, time_item)

            # Subject (truncate if long)
            subject = meeting.subject
            if len(subject) > 40:
                subject = subject[:37] + "..."
            subject_item = QTableWidgetItem(subject)
            self._events_table.setItem(i, 3, subject_item)

            # Hours (editable spinner)
            hours_spin = QSpinBox()
            hours_spin.setRange(1, 8)
            hours_spin.setValue(max(1, round(meeting.duration_hours)))
            self._events_table.setCellWidget(i, 4, hours_spin)

            # Activity code dropdown (default to TPLNT for meetings)
            activity_combo = QComboBox()
            for code in ActivityCode:
                activity_combo.addItem(code.value, code)
            # Default to TPLNT (Meeting) for calendar events
            tplnt_idx = [i for i, code in enumerate(ActivityCode) if code == ActivityCode.TPLNT]
            if tplnt_idx:
                activity_combo.setCurrentIndex(tplnt_idx[0])
            self._events_table.setCellWidget(i, 5, activity_combo)

            # Project dropdown (editable)
            project_combo = QComboBox()
            project_combo.addItem("-- Select --", None)
            selected_idx = 0
            for j, project in enumerate(projects):
                project_combo.addItem(project.name, project.id)
                if meeting.matched_project_id and project.id == meeting.matched_project_id:
                    selected_idx = j + 1
            project_combo.setCurrentIndex(selected_idx)
            self._events_table.setCellWidget(i, 6, project_combo)

            # Description (editable)
            desc_edit = QLineEdit()
            desc_edit.setText(meeting.subject)
            self._events_table.setCellWidget(i, 7, desc_edit)

            # Add single button
            add_btn = self._create_primary_button("Add")
            add_btn.setStyleSheet(f"""
                background-color: {self.COLORS['primary']};
                color: white;
                border: none;
                padding: 4px 8px;
            """)
            add_btn.clicked.connect(lambda checked, idx=i: self._add_single_meeting(idx))
            self._events_table.setCellWidget(i, 8, add_btn)

            # Color-code matched meetings with light green background
            if is_matched:
                self._set_row_background(self._events_table, i, self.COLORS['success_light'])

    def _add_single_meeting(self, row: int) -> None:
        """Add a single meeting as a timesheet entry."""
        if row >= len(self._outlook_meetings):
            return

        meeting = self._outlook_meetings[row]

        # Get values from widgets
        hours_spin = self._events_table.cellWidget(row, 4)
        hours = hours_spin.value() if hours_spin else max(1, round(meeting.duration_hours))

        activity_combo = self._events_table.cellWidget(row, 5)
        activity_code = activity_combo.currentData() if activity_combo else ActivityCode.TPLNT

        project_combo = self._events_table.cellWidget(row, 6)
        project_id = project_combo.currentData() if project_combo else None

        desc_edit = self._events_table.cellWidget(row, 7)
        description = desc_edit.text() if desc_edit else meeting.subject

        settings = self._store.get_settings()
        project = self._store.get_project(project_id) if project_id else None

        entry = TimesheetEntry(
            consultant_id=settings.consultant_id,
            entry_date=meeting.start_datetime.date(),
            hours=hours,
            ticket_number=project.ticket_number if project else None,
            project_name=project.name if project else None,
            activity_code=activity_code,
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
        add_btn = self._events_table.cellWidget(row, 8)
        if add_btn:
            add_btn.setEnabled(False)

        # Auto-refresh timesheet
        self._refresh_timesheet()

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

                    activity_combo = self._events_table.cellWidget(i, 5)
                    activity_code = activity_combo.currentData() if activity_combo else ActivityCode.TPLNT

                    project_combo = self._events_table.cellWidget(i, 6)
                    project_id = project_combo.currentData() if project_combo else None

                    desc_edit = self._events_table.cellWidget(i, 7)
                    description = desc_edit.text() if desc_edit else meeting.subject

                    project = self._store.get_project(project_id) if project_id else None

                    entry = TimesheetEntry(
                        consultant_id=settings.consultant_id,
                        entry_date=meeting.start_datetime.date(),
                        hours=hours,
                        ticket_number=project.ticket_number if project else None,
                        project_name=project.name if project else None,
                        activity_code=activity_code,
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
                    add_btn = self._events_table.cellWidget(i, 8)
                    if add_btn:
                        add_btn.setEnabled(False)
                    count += 1

        if count > 0:
            # Auto-refresh timesheet
            self._refresh_timesheet()
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
        save_btn = self._create_primary_button("Save Settings")
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

    # ==================== HELPERS ====================

    def _create_status_badge(self, status: EntryStatus) -> QLabel:
        """Create a colored status badge label."""
        label = QLabel(status.value)
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        if status == EntryStatus.EXPORTED:
            label.setStyleSheet(f"""
                background-color: {self.COLORS['success_light']};
                color: {self.COLORS['success']};
                padding: 2px 8px;
                border-radius: 10px;
                font-weight: bold;
                font-size: 11px;
            """)
        else:  # DRAFT
            label.setStyleSheet(f"""
                background-color: {self.COLORS['warning_light']};
                color: #9D5D00;
                padding: 2px 8px;
                border-radius: 10px;
                font-weight: bold;
                font-size: 11px;
            """)
        return label

    def _style_table(self, table: QTableWidget) -> None:
        """Apply consistent styling to a table widget."""
        table.setAlternatingRowColors(True)
        table.setStyleSheet(f"""
            QTableWidget {{
                alternate-background-color: {self.COLORS['gray_light']};
            }}
        """)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.verticalHeader().setVisible(False)

    def _create_primary_button(self, text: str) -> QPushButton:
        """Create a primary styled button."""
        btn = QPushButton(text)
        btn.setProperty("primary", True)
        btn.setStyleSheet(f"""
            background-color: {self.COLORS['primary']};
            color: white;
            border: none;
            padding: 8px 16px;
            font-weight: bold;
        """)
        return btn

    def _create_danger_button(self, text: str) -> QPushButton:
        """Create a danger/delete styled button."""
        btn = QPushButton(text)
        btn.setProperty("danger", True)
        btn.setStyleSheet(f"""
            background-color: {self.COLORS['danger']};
            color: white;
            border: none;
            padding: 4px 8px;
        """)
        return btn

    def _set_row_background(self, table: QTableWidget, row: int, color: str) -> None:
        """Set background color for all cells in a row."""
        brush = QBrush(QColor(color))
        for col in range(table.columnCount()):
            item = table.item(row, col)
            if item:
                item.setBackground(brush)

    def _populate_month_selector(self, combo: QComboBox, include_custom: bool = False) -> None:
        """
        Populate a month selector combo box with the last 12 months.

        Args:
            combo: The QComboBox to populate
            include_custom: Whether to include a "Custom Range..." option
        """
        combo.clear()
        today = date.today()

        # Add months from current month going back 12 months
        for i in range(12):
            # Calculate month
            year = today.year
            month = today.month - i
            while month <= 0:
                month += 12
                year -= 1

            # Calculate date range for this month
            first_day = date(year, month, 1)
            if month == 12:
                last_day = date(year + 1, 1, 1) - timedelta(days=1)
            else:
                last_day = date(year, month + 1, 1) - timedelta(days=1)

            # Format display name
            month_name = first_day.strftime("%B %Y")  # e.g., "January 2026"

            combo.addItem(month_name, (first_day, last_day))

        if include_custom:
            combo.addItem("Custom Range...", None)

    # ==================== DATA LOADING ====================

    def _load_data(self) -> None:
        """Load initial data."""
        self._refresh_dashboard()
        self._refresh_projects()
        self._refresh_timesheet()
        self._load_settings()
