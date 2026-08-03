"""Main window for TIQ Assistant desktop app with all functionality."""

from datetime import date, timedelta
from pathlib import Path
from typing import Optional

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QTabWidget,
    QLabel, QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QLineEdit, QSpinBox, QComboBox, QFormLayout, QGroupBox,
    QMessageBox, QFileDialog, QCheckBox, QAbstractItemView, QApplication,
    QScrollArea
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QBrush

from tiq_assistant.core.models import (
    Project, ActivityCode, EntryStatus, OutlookMeeting
)
from tiq_assistant.core.holidays import get_holiday_service, HolidayType
from tiq_assistant.storage.sqlite_store import get_store
from tiq_assistant.services.matching_service import get_matching_service
from tiq_assistant.services.timesheet_service import get_timesheet_service
from tiq_assistant.integrations.outlook_reader import get_outlook_reader, OutlookNotAvailableError
from tiq_assistant.exporters.excel_exporter import (
    ExcelExporter, get_monthly_export_path
)
from tiq_assistant.desktop.windows.day_entry_dialog import DayEntryDialog, SessionType
from tiq_assistant.desktop.icon import create_app_icon


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
        self.setWindowIcon(create_app_icon())
        self.setMinimumSize(900, 700)

        # Central widget with tabs
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # Tab widget
        self._tabs = QTabWidget()
        layout.addWidget(self._tabs)

        # Create tabs - Timesheet first
        self._tabs.addTab(self._create_timesheet_tab(), "Timesheet")
        self._tabs.addTab(self._create_projects_tab(), "Projects")
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
                selection-background-color: transparent;
                selection-color: {self.COLORS['text']};
            }}
            QTableWidget::item {{
                padding: 4px;
                color: {self.COLORS['text']};
            }}
            QTableWidget::item:hover {{
                background-color: rgba(0, 0, 0, 0.04);
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
            QScrollArea {{
                background-color: white;
                border: none;
            }}
            QScrollArea > QWidget > QWidget {{
                background-color: white;
            }}
        """)

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
        self._projects_table.setColumnCount(7)
        self._projects_table.setHorizontalHeaderLabels([
            "Name", "Ticket No", "JIRA Key", "Keywords", "Location", "Save", "Delete"
        ])
        self._projects_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch
        )
        # Cells are edited inline; a per-row "Save" button persists the changes.
        self._projects_table.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked
            | QAbstractItemView.EditTrigger.EditKeyPressed
        )
        self._style_table(self._projects_table)
        layout.addWidget(self._projects_table)

        hint = QLabel(
            "Double-click a cell to edit, then click Save on that row. "
            "Renaming a project can also update its existing timesheet entries."
        )
        hint.setStyleSheet(f"color: {self.COLORS['text_secondary']}; font-style: italic;")
        hint.setWordWrap(True)
        layout.addWidget(hint)

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

        # Guard against edit signals firing while we repopulate.
        self._projects_table.blockSignals(True)
        self._projects_table.setRowCount(len(projects))
        for i, project in enumerate(projects):
            name_item = QTableWidgetItem(project.name)
            # Stash the project id on the row's first cell so edits can be saved.
            name_item.setData(Qt.ItemDataRole.UserRole, project.id)
            self._projects_table.setItem(i, 0, name_item)
            self._projects_table.setItem(i, 1, QTableWidgetItem(project.ticket_number))
            self._projects_table.setItem(i, 2, QTableWidgetItem(project.jira_key or ""))
            self._projects_table.setItem(i, 3, QTableWidgetItem(
                ", ".join(project.keywords) if project.keywords else ""
            ))
            self._projects_table.setItem(i, 4, QTableWidgetItem(project.default_location))

            # Save button (persists inline edits for this row)
            save_btn = self._create_primary_button("Save")
            save_btn.clicked.connect(lambda checked, pid=project.id: self._save_project_row(pid))
            self._projects_table.setCellWidget(i, 5, save_btn)

            # Delete button
            delete_btn = self._create_danger_button("Delete")
            delete_btn.clicked.connect(lambda checked, pid=project.id: self._delete_project(pid))
            self._projects_table.setCellWidget(i, 6, delete_btn)
        self._projects_table.blockSignals(False)

    def _find_project_row(self, project_id: str) -> int:
        """Return the table row hosting the given project id, or -1."""
        for row in range(self._projects_table.rowCount()):
            item = self._projects_table.item(row, 0)
            if item and item.data(Qt.ItemDataRole.UserRole) == project_id:
                return row
        return -1

    def _save_project_row(self, project_id: str) -> None:
        """Persist inline edits for a project row, and offer to update entries."""
        row = self._find_project_row(project_id)
        if row < 0:
            return

        project = self._store.get_project(project_id)
        if project is None:
            QMessageBox.warning(self, "Error", "Project no longer exists.")
            self._refresh_projects()
            return

        def cell(col: int) -> str:
            item = self._projects_table.item(row, col)
            return item.text().strip() if item else ""

        new_name = cell(0)
        new_ticket = cell(1)
        new_jira = cell(2)
        new_keywords_text = cell(3)
        new_location = cell(4)

        if not new_name or not new_ticket:
            QMessageBox.warning(self, "Error", "Project Name and Ticket No are required.")
            return

        old_name = project.name

        # Apply edits to the model.
        project.name = new_name
        project.ticket_number = new_ticket
        project.jira_key = new_jira or None
        project.keywords = [k.strip() for k in new_keywords_text.split(",") if k.strip()]
        project.default_location = new_location or project.default_location

        self._store.save_project(project)

        # If the name changed, offer to update existing timesheet entries, which
        # store project_name as a copied string (so they don't auto-rename).
        updated_entries = 0
        if new_name != old_name:
            affected = [
                e for e in self._store.get_entries()
                if e.project_name == old_name
            ]
            if affected:
                reply = QMessageBox.question(
                    self, "Update existing entries?",
                    f"{len(affected)} existing timesheet entr"
                    f"{'y' if len(affected) == 1 else 'ies'} use the old name "
                    f"'{old_name}'.\n\nUpdate them to '{new_name}' as well?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                if reply == QMessageBox.StandardButton.Yes:
                    for e in affected:
                        e.project_name = new_name
                        e.ticket_number = new_ticket
                        self._store.save_entry(e)
                    updated_entries = len(affected)

        # Refresh dependent views so the new name shows everywhere.
        self._refresh_projects()
        self._load_settings()   # default-project dropdown
        self._refresh_timesheet()

        msg = f"Project updated to '{new_name}'."
        if updated_entries:
            msg += f"\nAlso updated {updated_entries} timesheet entr" \
                   f"{'y' if updated_entries == 1 else 'ies'}."
        QMessageBox.information(self, "Saved", msg)

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

        # Month selector and actions
        month_layout = QHBoxLayout()
        month_layout.addWidget(QLabel("Month:"))

        self._timesheet_month = QComboBox()
        self._populate_month_selector(self._timesheet_month)
        self._timesheet_month.currentIndexChanged.connect(self._refresh_timesheet)
        month_layout.addWidget(self._timesheet_month)

        # Fetch from Outlook button
        self._fetch_btn = self._create_primary_button("📅 Fetch from Outlook")
        self._fetch_btn.clicked.connect(self._fetch_outlook_for_month)
        month_layout.addWidget(self._fetch_btn)

        month_layout.addStretch()

        # Summary label
        self._timesheet_summary = QLabel("")
        self._timesheet_summary.setStyleSheet("font-weight: bold; margin-left: 20px;")
        month_layout.addWidget(self._timesheet_summary)

        layout.addLayout(month_layout)

        # Outlook fetch status
        self._outlook_status = QLabel("")
        self._outlook_status.setStyleSheet(f"color: {self.COLORS['text_secondary']}; font-style: italic;")
        layout.addWidget(self._outlook_status)

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
        # Disable built-in selection to manage it manually with colors
        self._workday_table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self._workday_table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self._workday_table.cellClicked.connect(self._on_workday_clicked)
        self._style_table(self._workday_table)
        workday_layout.addWidget(self._workday_table)

        # Track selected row index and row statuses for coloring
        self._selected_workday_row: int = -1
        self._workday_row_colors: dict[int, str] = {}  # row -> base background color

        # Add workday group with stretch factor so it expands to fill space
        layout.addWidget(workday_group, 1)

        # Initialize outlook meetings list (for use in day entry dialog)
        self._outlook_meetings = []

        # Tip for user
        tip_label = QLabel("Click on a day to add/edit entries. Fetch from Outlook first to import meetings.")
        tip_label.setStyleSheet(f"color: {self.COLORS['text_secondary']}; font-style: italic; margin-top: 8px;")
        layout.addWidget(tip_label)

        # Export section
        export_layout = QHBoxLayout()
        export_layout.addStretch()

        export_btn = self._create_primary_button("Export Month to Excel")
        export_btn.clicked.connect(self._export_entries)
        export_layout.addWidget(export_btn)

        layout.addLayout(export_layout)

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
        self._selected_workday_row = -1

    def _refresh_workday_overview(self, month_start: date, entries: list) -> None:
        """Refresh the workday overview table showing expected vs filled hours."""
        holiday_service = get_holiday_service()

        # Get workdays for the month
        workdays = holiday_service.get_workdays_in_month(month_start.year, month_start.month)

        # Get skipped days for the month
        if workdays:
            first_date = workdays[0][0]
            last_date = workdays[-1][0]
            skipped_days = self._store.get_skipped_days(first_date, last_date)
        else:
            skipped_days = {}

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
            is_skipped = work_date in skipped_days

            # Don't count skipped days in expected hours
            if not is_skipped:
                total_expected += expected_hours
                total_filled += filled_hours

            if is_skipped or filled_hours >= expected_hours:
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

            # Determine status and row color (softer, muted colors)
            if is_skipped:
                skip_reason = skipped_days[work_date]
                status_item = QTableWidgetItem(f"⊘ {skip_reason}")
                row_color = "#E5E7EB"  # Gray for skipped
            elif filled_hours >= expected_hours:
                status_item = QTableWidgetItem("✓ Complete")
                row_color = "#E8F5E9"  # Soft mint green
            elif filled_hours > 0:
                status_item = QTableWidgetItem(f"Partial ({remaining}h left)")
                row_color = "#FFF8E1"  # Soft cream/pale yellow
            elif work_date < date.today():
                status_item = QTableWidgetItem("Missing")
                row_color = "#FFEBEE"  # Soft blush pink
            else:
                status_item = QTableWidgetItem("Pending")
                row_color = "#FAFAFA"  # Very light gray

            self._workday_table.setItem(i, 5, status_item)
            self._workday_row_colors[i] = row_color

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
        """Handle workday row click - open day entry dialog."""
        # Get the date from the first column
        date_item = self._workday_table.item(row, 0)
        if not date_item:
            return

        selected_date = date_item.data(Qt.ItemDataRole.UserRole)
        if not selected_date:
            return

        # Open the day entry dialog
        self._open_day_entry_dialog(selected_date, SessionType.FULL_DAY)

    def _open_day_entry_dialog(
        self,
        target_date: date,
        session: SessionType = SessionType.FULL_DAY
    ) -> None:
        """Open the day entry dialog for the specified date and session."""
        dialog = DayEntryDialog(
            target_date=target_date,
            session=session,
            outlook_meetings=self._outlook_meetings,
            parent=self
        )

        dialog.exec()

        # Refresh the timesheet after dialog closes
        self._refresh_timesheet()

    def _fetch_outlook_for_month(self) -> None:
        """Fetch meetings from Outlook for the selected month."""
        # Show loading state
        self._fetch_btn.setEnabled(False)
        self._fetch_btn.setText("⏳ Fetching...")
        self._outlook_status.setText("Connecting to Outlook...")
        self._outlook_status.setStyleSheet(f"color: {self.COLORS['primary']};")
        # Force UI update before blocking operation
        QApplication.processEvents()

        try:
            reader = get_outlook_reader()

            if not reader.is_available():
                QMessageBox.warning(
                    self, "Outlook Not Available",
                    "Could not connect to Outlook. Make sure Outlook desktop "
                    "(not the web version) is installed and has been opened at least once."
                )
                return

            # Get date range from timesheet month selector
            month_data = self._timesheet_month.currentData()
            if month_data:
                start_date, end_date = month_data
            else:
                today = date.today()
                start_date = date(today.year, today.month, 1)
                if today.month == 12:
                    end_date = date(today.year + 1, 1, 1) - timedelta(days=1)
                else:
                    end_date = date(today.year, today.month + 1, 1) - timedelta(days=1)

            self._outlook_status.setText("Fetching meetings from Outlook...")
            QApplication.processEvents()

            # Fetch meetings
            meetings = reader.get_meetings_for_date_range(start_date, end_date)
            self._outlook_meetings = meetings

            self._outlook_status.setText("Matching meetings to projects...")
            QApplication.processEvents()

            # Match meetings to projects
            for meeting in meetings:
                event = reader.to_calendar_event(meeting)
                result = self._matching_service.match_event(event)
                meeting.matched_project_id = result.project_id
                meeting.matched_jira_key = result.ticket_jira_key
                meeting.match_confidence = result.confidence

            matched_count = len([m for m in meetings if m.match_confidence and m.match_confidence > 0])
            self._outlook_status.setText(
                f"Found {len(meetings)} meetings ({matched_count} matched). Click a day to add entries."
            )
            self._outlook_status.setStyleSheet(f"color: {self.COLORS['success']};")

        except OutlookNotAvailableError as e:
            QMessageBox.warning(self, "Outlook Error", str(e))
            self._outlook_status.setText("Failed to connect to Outlook")
            self._outlook_status.setStyleSheet(f"color: {self.COLORS['danger']};")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to fetch meetings: {e}")
            self._outlook_status.setText(f"Error: {e}")
            self._outlook_status.setStyleSheet(f"color: {self.COLORS['danger']};")
        finally:
            # Restore button state
            self._fetch_btn.setEnabled(True)
            self._fetch_btn.setText("📅 Fetch from Outlook")

    def _export_entries(self) -> None:
        """Export entries to Excel."""
        from datetime import datetime

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

        # Get export path with auto-incrementing version
        target_datetime = datetime(start.year, start.month, 1)
        export_path = get_monthly_export_path(target_date=target_datetime)

        exporter = ExcelExporter()
        exporter.export_to_new_file(entries, export_path)

        # Mark as exported
        self._store.mark_entries_exported([e.id for e in entries])

        self._refresh_timesheet()

        QMessageBox.information(
            self, "Export Complete",
            f"Exported {len(entries)} entries to:\n{export_path}"
        )

    # ==================== SETTINGS TAB ====================

    def _create_settings_tab(self) -> QWidget:
        """Create the settings tab.

        The tab has grown several sections (User, Activity, Matching, AI, Holiday)
        so its content is placed inside a scroll area -- otherwise the lower
        sections get cut off on smaller windows.
        """
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setSpacing(14)

        # User settings
        user_group = QGroupBox("User Settings")
        user_layout = QFormLayout(user_group)

        self._consultant_id_input = QLineEdit()
        user_layout.addRow("Consultant ID:", self._consultant_id_input)

        self._location_input = QLineEdit()
        user_layout.addRow("Default Location:", self._location_input)

        # Default project dropdown
        self._default_project_combo = QComboBox()
        self._default_project_combo.addItem("-- None --", None)
        user_layout.addRow("Default Project:", self._default_project_combo)

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

        # AI assistant (local LLM) — OFF by default. When enabled, the app can
        # contact the configured internal LLM endpoint to draft entries from a
        # spoken/typed summary. This is the only feature that makes a network
        # call, hence the explicit opt-in.
        ai_group = QGroupBox("AI Assistant (local LLM) — optional")
        ai_layout = QFormLayout(ai_group)

        self._ai_enabled = QCheckBox("Enable voice / AI entry drafting")
        ai_layout.addRow("", self._ai_enabled)

        self._ai_base_url = QLineEdit()
        self._ai_base_url.setPlaceholderText("https://.../v1")
        ai_layout.addRow("Endpoint URL:", self._ai_base_url)

        self._ai_model = QLineEdit()
        self._ai_model.setPlaceholderText("(leave blank to auto-detect)")
        ai_layout.addRow("Model:", self._ai_model)

        self._ai_verify_ssl = QCheckBox("Verify SSL certificate")
        self._ai_verify_ssl.setToolTip(
            "Internal servers often use a self-signed cert; leave unchecked if so."
        )
        ai_layout.addRow("", self._ai_verify_ssl)

        # Whisper (speech-to-text) model: a size name that downloads, or a local
        # folder path for machines that block the Hugging Face download.
        whisper_row = QHBoxLayout()
        self._ai_whisper = QLineEdit()
        self._ai_whisper.setPlaceholderText('"base" (downloads) or a local model folder path')
        whisper_row.addWidget(self._ai_whisper, 1)
        browse_whisper = QPushButton("Browse…")
        browse_whisper.clicked.connect(self._browse_whisper_model)
        whisper_row.addWidget(browse_whisper)
        whisper_widget = QWidget()
        whisper_widget.setLayout(whisper_row)
        ai_layout.addRow("Whisper model:", whisper_widget)

        test_btn = QPushButton("Test connection")
        test_btn.clicked.connect(self._test_llm_connection)
        ai_layout.addRow("", test_btn)

        note = QLabel(
            "When enabled, the app contacts the endpoint above to draft entries. "
            "It stays fully offline when disabled."
        )
        note.setWordWrap(True)
        note.setStyleSheet(f"color: {self.COLORS['text_secondary']}; font-style: italic;")
        ai_layout.addRow("", note)

        layout.addWidget(ai_group)

        # Holidays section — opens a dedicated, spacious manager dialog so it's
        # not squashed into this tab.
        holidays_group = QGroupBox("Holiday Calendar")
        holidays_layout = QVBoxLayout(holidays_group)

        instructions = QLabel(
            "Manage national holidays and half-days used in workday calculations. "
            "You can add, edit dates/types, upload a calendar file, or load defaults."
        )
        instructions.setWordWrap(True)
        instructions.setStyleSheet(f"color: {self.COLORS['text_secondary']}; font-style: italic;")
        holidays_layout.addWidget(instructions)

        self._holidays_summary = QLabel("")
        self._holidays_summary.setStyleSheet(f"color: {self.COLORS['text_secondary']};")
        holidays_layout.addWidget(self._holidays_summary)

        manage_btn = self._create_primary_button("📅 Manage Holiday Calendar…")
        manage_btn.clicked.connect(self._open_holiday_manager)
        holidays_layout.addWidget(manage_btn)

        layout.addWidget(holidays_group)

        # Save button
        save_btn = self._create_primary_button("Save Settings")
        save_btn.clicked.connect(self._save_settings)
        layout.addWidget(save_btn)

        layout.addStretch()

        # Wrap the content in a scroll area so every section is reachable even
        # on smaller windows.
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setWidget(widget)
        return scroll

    def _load_settings(self) -> None:
        """Load settings into the form."""
        settings = self._store.get_settings()

        self._consultant_id_input.setText(settings.consultant_id)
        self._location_input.setText(settings.default_location)

        # Populate default project dropdown
        self._default_project_combo.clear()
        self._default_project_combo.addItem("-- None --", None)
        projects = self._store.get_projects()
        selected_idx = 0
        for i, project in enumerate(projects):
            self._default_project_combo.addItem(project.name, project.id)
            if settings.default_project_id and project.id == settings.default_project_id:
                selected_idx = i + 1  # +1 because of "-- None --" option
        self._default_project_combo.setCurrentIndex(selected_idx)

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

        # AI assistant config
        from tiq_assistant.services.entry_generation_service import load_llm_config
        cfg = load_llm_config(self._store)
        self._ai_enabled.setChecked(cfg.enabled)
        self._ai_base_url.setText(cfg.base_url)
        self._ai_model.setText(cfg.model)
        self._ai_verify_ssl.setChecked(cfg.verify_ssl)
        self._ai_whisper.setText(cfg.whisper_model)

    def _save_settings(self) -> None:
        """Save settings."""
        from tiq_assistant.core.models import UserSettings

        settings = UserSettings(
            consultant_id=self._consultant_id_input.text().strip() or "FMANZOOR",
            default_location=self._location_input.text().strip() or "ANKARA",
            default_activity_code=self._default_activity.currentData(),
            meeting_activity_code=self._meeting_activity.currentData(),
            default_project_id=self._default_project_combo.currentData(),
            skip_canceled_meetings=self._skip_canceled.isChecked(),
            min_meeting_duration_minutes=self._min_duration.value(),
        )

        self._store.save_settings(settings)

        # Persist AI assistant config.
        from tiq_assistant.services.entry_generation_service import (
            load_llm_config, save_llm_config,
        )
        cfg = load_llm_config(self._store)
        cfg.enabled = self._ai_enabled.isChecked()
        cfg.base_url = self._ai_base_url.text().strip() or cfg.base_url
        cfg.model = self._ai_model.text().strip()
        cfg.verify_ssl = self._ai_verify_ssl.isChecked()
        cfg.whisper_model = self._ai_whisper.text().strip() or "base"
        save_llm_config(cfg, self._store)

        QMessageBox.information(self, "Saved", "Settings saved!")

    def _browse_whisper_model(self) -> None:
        """Pick a local folder containing a downloaded faster-whisper model."""
        folder = QFileDialog.getExistingDirectory(
            self, "Select local Whisper model folder", ""
        )
        if folder:
            self._ai_whisper.setText(folder)

    def _test_llm_connection(self) -> None:
        """Test the configured LLM endpoint and report the resolved model."""
        from tiq_assistant.integrations.llm_client import LLMClient, LLMConfig, LLMError

        cfg = LLMConfig(
            enabled=True,
            base_url=self._ai_base_url.text().strip(),
            model=self._ai_model.text().strip(),
            verify_ssl=self._ai_verify_ssl.isChecked(),
        )
        if not cfg.base_url:
            QMessageBox.warning(self, "No endpoint", "Enter the endpoint URL first.")
            return

        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            model = LLMClient(cfg).test_connection()
            QMessageBox.information(
                self, "Connection OK",
                f"Connected successfully.\nModel: {model}"
            )
        except LLMError as e:
            QMessageBox.warning(self, "Connection failed", str(e))
        finally:
            QApplication.restoreOverrideCursor()

    # ==================== HOLIDAY MANAGEMENT ====================

    def _open_holiday_manager(self) -> None:
        """Open the dedicated holiday calendar manager dialog."""
        from tiq_assistant.desktop.windows.holiday_dialog import HolidayDialog

        dialog = HolidayDialog(parent=self)
        dialog.exec()

        # Holidays may have changed -> refresh dependent views.
        get_holiday_service().reload_from_database()
        self._refresh_holidays_summary()
        self._refresh_timesheet()

    def _refresh_holidays_summary(self) -> None:
        """Update the one-line holiday summary shown on the Settings tab."""
        holidays = self._store.get_holidays()
        full = sum(1 for h in holidays if h["holiday_type"] == "full_day")
        half = sum(1 for h in holidays if h["holiday_type"] == "half_day")
        if holidays:
            self._holidays_summary.setText(
                f"{len(holidays)} holidays configured ({full} full-day, {half} half-day)."
            )
        else:
            self._holidays_summary.setText("No custom holidays configured (using built-in defaults).")

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
        self._refresh_timesheet()
        self._refresh_projects()
        self._load_settings()
        self._refresh_holidays_summary()
