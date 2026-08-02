"""Dedicated dialog for managing the holiday calendar.

Moved out of the cramped Settings tab into its own spacious window. Supports:
- inline editing of a holiday's date, name and type (per-row Save),
- adding a single holiday manually,
- uploading a holiday file / loading built-in defaults,
- deleting a single holiday or clearing all custom holidays.

Every change reloads the shared HolidayService so workday calculations update
immediately.
"""

from datetime import date
from pathlib import Path
from typing import Optional

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView, QComboBox,
    QLineEdit, QGroupBox, QWidget, QAbstractItemView, QMessageBox,
    QFileDialog, QDateEdit, QSpinBox,
)
from PyQt6.QtCore import Qt, QDate

from tiq_assistant.core.holidays import get_holiday_service
from tiq_assistant.storage.sqlite_store import get_store
from tiq_assistant.core.exceptions import StorageError
from tiq_assistant.desktop.icon import create_app_icon


FULL_DAY = "full_day"
HALF_DAY = "half_day"


class HolidayDialog(QDialog):
    """Spacious holiday-calendar manager with inline editing."""

    COLORS = {
        'primary': '#0078D4',
        'success': '#107C10',
        'danger': '#D13438',
        'warning_light': '#FFF4CE',
        'gray': '#E1E1E1',
        'gray_light': '#F5F5F5',
        'text': '#323130',
        'text_secondary': '#605E5C',
    }

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._store = get_store()
        self._setup_ui()
        self._refresh_table()

    # ------------------------------------------------------------------ UI

    def _setup_ui(self) -> None:
        self.setWindowTitle("Holiday Calendar")
        self.setWindowIcon(create_app_icon())
        self.setModal(True)
        self.setMinimumSize(760, 620)
        self.resize(900, 720)
        self.setStyleSheet(f"""
            QDialog {{ background-color: white; color: {self.COLORS['text']}; }}
            QLabel {{ color: {self.COLORS['text']}; }}
            QGroupBox {{
                font-weight: bold; color: {self.COLORS['text']};
                border: 1px solid {self.COLORS['gray']}; border-radius: 4px;
                margin-top: 12px; padding-top: 10px;
            }}
            QGroupBox::title {{ subcontrol-origin: margin; left: 10px; padding: 0 5px; }}
            QTableWidget {{
                border: 1px solid {self.COLORS['gray']};
                gridline-color: {self.COLORS['gray']};
                background-color: white; color: {self.COLORS['text']};
            }}
            QHeaderView::section {{
                background-color: {self.COLORS['gray_light']};
                color: {self.COLORS['text']}; padding: 8px; border: none;
                border-right: 1px solid {self.COLORS['gray']};
                border-bottom: 1px solid {self.COLORS['gray']};
                font-weight: bold;
            }}
            QComboBox, QLineEdit, QDateEdit, QSpinBox {{
                color: {self.COLORS['text']}; background-color: white;
                padding: 5px 8px; border: 1px solid {self.COLORS['gray']};
                border-radius: 4px;
            }}
            QComboBox QAbstractItemView {{
                background-color: white; color: {self.COLORS['text']};
                selection-background-color: {self.COLORS['primary']};
                selection-color: white;
            }}
            QPushButton {{
                color: {self.COLORS['text']}; background-color: white;
                padding: 6px 12px; border: 1px solid {self.COLORS['gray']};
                border-radius: 4px;
            }}
            QPushButton:hover {{ background-color: {self.COLORS['gray_light']}; }}
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        intro = QLabel(
            "Edit any holiday's date, name or type directly in the table, then "
            "click Save on that row. Half-day holidays expect 4 working hours; "
            "full-day holidays expect none."
        )
        intro.setWordWrap(True)
        intro.setStyleSheet(f"color: {self.COLORS['text_secondary']};")
        layout.addWidget(intro)

        # --- Add a single holiday ---
        add_group = QGroupBox("Add Holiday")
        add_row = QHBoxLayout(add_group)

        add_row.addWidget(QLabel("Date:"))
        self._add_date = QDateEdit()
        self._add_date.setCalendarPopup(True)
        self._add_date.setDisplayFormat("dd.MM.yyyy")
        self._add_date.setDate(QDate.currentDate())
        add_row.addWidget(self._add_date)

        add_row.addWidget(QLabel("Name:"))
        self._add_name = QLineEdit()
        self._add_name.setPlaceholderText("e.g. Cumhuriyet Bayramı")
        add_row.addWidget(self._add_name, 1)

        add_row.addWidget(QLabel("Type:"))
        self._add_type = QComboBox()
        self._add_type.addItem("Full Day", FULL_DAY)
        self._add_type.addItem("Half Day", HALF_DAY)
        add_row.addWidget(self._add_type)

        add_btn = self._primary_button("Add")
        add_btn.clicked.connect(self._add_holiday)
        add_row.addWidget(add_btn)

        layout.addWidget(add_group)

        # --- The table ---
        self._table = QTableWidget()
        self._table.setColumnCount(5)
        self._table.setHorizontalHeaderLabels(
            ["Date (dd.mm.yyyy)", "Name", "Type", "Save", "Delete"]
        )
        self._table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.Stretch
        )
        self._table.setColumnWidth(0, 150)
        self._table.setColumnWidth(2, 120)
        self._table.setColumnWidth(3, 80)
        self._table.setColumnWidth(4, 90)
        self._table.verticalHeader().setVisible(False)
        self._table.verticalHeader().setDefaultSectionSize(38)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked
            | QAbstractItemView.EditTrigger.EditKeyPressed
        )
        layout.addWidget(self._table, 1)  # stretch to fill

        self._status = QLabel("")
        self._status.setStyleSheet(f"color: {self.COLORS['text_secondary']};")
        layout.addWidget(self._status)

        # --- File / defaults / clear controls ---
        controls = QHBoxLayout()

        upload_btn = QPushButton("📁 Upload File…")
        upload_btn.clicked.connect(self._upload_file)
        controls.addWidget(upload_btn)

        controls.addWidget(QLabel("Year:"))
        self._year = QSpinBox()
        self._year.setRange(2024, 2030)
        self._year.setValue(date.today().year)
        controls.addWidget(self._year)

        defaults_btn = QPushButton("Load Defaults")
        defaults_btn.clicked.connect(self._load_defaults)
        controls.addWidget(defaults_btn)

        controls.addStretch()

        clear_btn = self._danger_button("Clear All")
        clear_btn.clicked.connect(self._clear_all)
        controls.addWidget(clear_btn)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        controls.addWidget(close_btn)

        layout.addLayout(controls)

    def _primary_button(self, text: str) -> QPushButton:
        btn = QPushButton(text)
        btn.setStyleSheet(
            f"background-color: {self.COLORS['primary']}; color: white; "
            f"border: none; padding: 6px 14px; font-weight: bold;"
        )
        return btn

    def _danger_button(self, text: str) -> QPushButton:
        btn = QPushButton(text)
        btn.setStyleSheet(
            f"background-color: {self.COLORS['danger']}; color: white; "
            f"border: none; padding: 6px 12px;"
        )
        return btn

    # ------------------------------------------------------------- table

    def _refresh_table(self) -> None:
        holidays = self._store.get_holidays()
        # Sort by date so misplaced entries are easy to spot.
        holidays.sort(key=lambda h: h["holiday_date"])

        self._table.blockSignals(True)
        self._table.setRowCount(len(holidays))
        for i, h in enumerate(holidays):
            hid = h["id"]

            date_item = QTableWidgetItem(h["holiday_date"].strftime("%d.%m.%Y"))
            date_item.setData(Qt.ItemDataRole.UserRole, hid)
            self._table.setItem(i, 0, date_item)

            self._table.setItem(i, 1, QTableWidgetItem(h["name"]))

            type_combo = QComboBox()
            type_combo.addItem("Full Day", FULL_DAY)
            type_combo.addItem("Half Day", HALF_DAY)
            type_combo.setCurrentIndex(1 if h["holiday_type"] == HALF_DAY else 0)
            self._table.setCellWidget(i, 2, type_combo)

            save_btn = self._primary_button("Save")
            save_btn.clicked.connect(lambda _c, x=hid: self._save_row(x))
            self._table.setCellWidget(i, 3, save_btn)

            del_btn = self._danger_button("Delete")
            del_btn.clicked.connect(lambda _c, x=hid: self._delete_row(x))
            self._table.setCellWidget(i, 4, del_btn)

            if h["holiday_type"] == HALF_DAY:
                for col in (0, 1):
                    cell = self._table.item(i, col)
                    if cell:
                        from PyQt6.QtGui import QColor, QBrush
                        cell.setBackground(QBrush(QColor(self.COLORS['warning_light'])))
        self._table.blockSignals(False)

        full = sum(1 for h in holidays if h["holiday_type"] == FULL_DAY)
        half = sum(1 for h in holidays if h["holiday_type"] == HALF_DAY)
        self._status.setText(
            f"{len(holidays)} holidays ({full} full-day, {half} half-day)"
        )

    def _row_for_id(self, holiday_id: int) -> int:
        for row in range(self._table.rowCount()):
            item = self._table.item(row, 0)
            if item and item.data(Qt.ItemDataRole.UserRole) == holiday_id:
                return row
        return -1

    def _save_row(self, holiday_id: int) -> None:
        row = self._row_for_id(holiday_id)
        if row < 0:
            return

        date_text = self._table.item(row, 0).text().strip()
        name = self._table.item(row, 1).text().strip()
        type_combo = self._table.cellWidget(row, 2)
        holiday_type = type_combo.currentData() if type_combo else FULL_DAY

        parsed = self._parse_date(date_text)
        if parsed is None:
            QMessageBox.warning(
                self, "Invalid date",
                f"'{date_text}' is not a valid date. Use dd.mm.yyyy (e.g. 19.03.2026)."
            )
            return
        if not name:
            QMessageBox.warning(self, "Name required", "Holiday name cannot be empty.")
            return

        try:
            self._store.update_holiday(holiday_id, parsed, name, holiday_type)
        except StorageError as e:
            QMessageBox.warning(self, "Cannot save", str(e))
            return

        self._reload_service()
        self._refresh_table()
        self._status.setText(f"Saved: {name} ({parsed.strftime('%d.%m.%Y')})")

    def _delete_row(self, holiday_id: int) -> None:
        reply = QMessageBox.question(
            self, "Delete holiday",
            "Delete this holiday?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        self._store.delete_holiday(holiday_id)
        self._reload_service()
        self._refresh_table()

    def _add_holiday(self) -> None:
        qd = self._add_date.date()
        d = date(qd.year(), qd.month(), qd.day())
        name = self._add_name.text().strip()
        holiday_type = self._add_type.currentData()

        if not name:
            QMessageBox.warning(self, "Name required", "Enter a holiday name.")
            return
        try:
            self._store.add_holiday(d, name, holiday_type)
        except StorageError as e:
            QMessageBox.warning(self, "Cannot add", str(e))
            return

        self._add_name.clear()
        self._reload_service()
        self._refresh_table()

    # ------------------------------------------------------- files/defaults

    def _upload_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Holiday Calendar", "",
            "All Supported Files (*.pdf *.jpg *.jpeg *.png);;"
            "PDF Files (*.pdf);;Images (*.jpg *.jpeg *.png)"
        )
        if not file_path:
            return

        from tiq_assistant.services.holiday_parser import parse_holiday_file

        result = parse_holiday_file(Path(file_path), self._year.value())
        if result.errors and not result.holidays:
            QMessageBox.warning(
                self, "Parse Error",
                "Could not extract holidays:\n" + "\n".join(result.errors)
            )
            return
        if result.holidays:
            tuples = [(h[0], h[1], h[2]) for h in result.holidays]
            count = self._store.save_holidays_batch(tuples, result.source_file)
            self._reload_service()
            self._refresh_table()
            QMessageBox.information(
                self, "Holidays Loaded",
                f"Loaded {count} holidays from {result.source_file}."
            )

    def _load_defaults(self) -> None:
        from tiq_assistant.services.holiday_parser import get_default_holidays_for_year

        year = self._year.value()
        holidays = get_default_holidays_for_year(year)
        if not holidays:
            QMessageBox.warning(
                self, "No Defaults",
                f"No default holidays available for {year}."
            )
            return
        count = self._store.save_holidays_batch(holidays, f"defaults_{year}")
        self._reload_service()
        self._refresh_table()
        QMessageBox.information(
            self, "Defaults Loaded", f"Loaded {count} default holidays for {year}."
        )

    def _clear_all(self) -> None:
        reply = QMessageBox.question(
            self, "Clear all",
            "Clear all custom holidays and revert to built-in defaults?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        count = self._store.clear_all_holidays()
        self._reload_service()
        self._refresh_table()
        QMessageBox.information(self, "Cleared", f"Cleared {count} custom holidays.")

    # ------------------------------------------------------------- helpers

    @staticmethod
    def _parse_date(text: str) -> Optional[date]:
        """Parse dd.mm.yyyy (also tolerates dd/mm/yyyy and yyyy-mm-dd)."""
        from datetime import datetime
        for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                continue
        return None

    @staticmethod
    def _reload_service() -> None:
        try:
            get_holiday_service().reload_from_database()
        except Exception:
            pass
