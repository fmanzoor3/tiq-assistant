"""Review dialog for distributing top-up hours across specific entries.

When a day is short of its target hours, "Auto-fill Day" opens this dialog so
the user can distribute the remaining hours across one or more lines. Each line
has an editable project, hours, and description. Descriptions are pre-filled
from the user's own past entries (see ``get_description_history``) so the result
is specific rather than a generic "General work" placeholder -- but every field
stays editable and nothing is saved until the totals match and the user confirms.
"""

from datetime import date
from typing import Optional

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView, QSpinBox,
    QComboBox, QGroupBox, QWidget, QAbstractItemView, QMessageBox,
)
from PyQt6.QtCore import Qt

from tiq_assistant.core.models import (
    TimesheetEntry, ActivityCode, EntryStatus, EntrySource,
)
from tiq_assistant.storage.sqlite_store import get_store


class TopUpDialog(QDialog):
    """Distribute remaining hours across editable, specific entries."""

    COLORS = {
        'primary': '#0078D4',
        'success': '#107C10',
        'danger': '#D13438',
        'warning': '#FFB900',
        'gray': '#E1E1E1',
        'gray_light': '#F5F5F5',
        'text': '#323130',
        'text_secondary': '#605E5C',
    }

    MAX_ENTRY_HOURS = 12

    def __init__(
        self,
        target_date: date,
        remaining_hours: int,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)

        self._store = get_store()
        self._target_date = target_date
        self._remaining = remaining_hours

        self._projects = self._store.get_projects()
        self._settings = self._store.get_settings()

        # Cache description history per project name so combos populate fast.
        self._history_cache: dict[str, list[str]] = {}

        self._saved_count = 0

        self._setup_ui()
        self._add_initial_line()
        self._update_total()

    # ------------------------------------------------------------------ UI

    def _setup_ui(self) -> None:
        self.setWindowTitle("Distribute Remaining Hours")
        self.setMinimumSize(720, 380)
        self.setModal(True)
        self.setStyleSheet(f"""
            QDialog {{ background-color: white; color: {self.COLORS['text']}; }}
            QLabel {{ color: {self.COLORS['text']}; }}
            QTableWidget {{
                border: 1px solid {self.COLORS['gray']};
                gridline-color: {self.COLORS['gray']};
                background-color: white;
                color: {self.COLORS['text']};
            }}
            QHeaderView::section {{
                background-color: {self.COLORS['gray_light']};
                color: {self.COLORS['text']};
                padding: 6px; border: none;
                border-right: 1px solid {self.COLORS['gray']};
                border-bottom: 1px solid {self.COLORS['gray']};
                font-weight: bold;
            }}
            QComboBox, QSpinBox {{
                color: {self.COLORS['text']}; background-color: white;
                padding: 3px 6px; border: 1px solid {self.COLORS['gray']};
                border-radius: 4px;
            }}
            QComboBox QAbstractItemView {{
                color: {self.COLORS['text']}; background-color: white;
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

        # Header
        date_str = self._target_date.strftime("%A, %d %B %Y")
        header = QLabel(
            f"<b>{date_str}</b><br>"
            f"Distribute the remaining <b>{self._remaining}h</b> across one or "
            f"more entries. Descriptions are pre-filled from your history — "
            f"edit anything before confirming."
        )
        header.setWordWrap(True)
        layout.addWidget(header)

        # Lines table
        group = QGroupBox("Entries")
        group_layout = QVBoxLayout(group)

        self._table = QTableWidget()
        self._table.setColumnCount(4)
        self._table.setHorizontalHeaderLabels(
            ["Project", "Hours", "Description", ""]
        )
        self._table.horizontalHeader().setSectionResizeMode(
            2, QHeaderView.ResizeMode.Stretch
        )
        self._table.setColumnWidth(1, 70)
        self._table.setColumnWidth(3, 40)
        self._table.verticalHeader().setVisible(False)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        group_layout.addWidget(self._table)

        # Add-line + total row
        controls = QHBoxLayout()
        add_line_btn = QPushButton("+ Add line")
        add_line_btn.clicked.connect(self._add_line)
        controls.addWidget(add_line_btn)

        distribute_btn = QPushButton("Auto-distribute evenly")
        distribute_btn.setToolTip("Spread the remaining hours evenly across the current lines.")
        distribute_btn.clicked.connect(self._distribute_evenly)
        controls.addWidget(distribute_btn)

        controls.addStretch()
        self._total_label = QLabel()
        controls.addWidget(self._total_label)
        group_layout.addLayout(controls)

        layout.addWidget(group)

        # Bottom buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        self._confirm_btn = QPushButton("Confirm")
        self._confirm_btn.setStyleSheet(f"""
            background-color: {self.COLORS['success']};
            color: white; border: none; padding: 8px 20px; font-weight: bold;
        """)
        self._confirm_btn.setDefault(True)
        self._confirm_btn.clicked.connect(self._on_confirm)
        btn_layout.addWidget(self._confirm_btn)

        layout.addLayout(btn_layout)

    # -------------------------------------------------------------- history

    def _history_for_project(self, project_name: Optional[str]) -> list[str]:
        """Return description suggestions for a project (cached)."""
        key = project_name or ""
        if key not in self._history_cache:
            # Project-specific history first, then a general fallback so the
            # combo is never empty even for a brand-new project.
            suggestions = self._store.get_description_history(project_name=project_name)
            if not suggestions:
                suggestions = self._store.get_description_history(project_name=None)
            self._history_cache[key] = suggestions
        return self._history_cache[key]

    # ---------------------------------------------------------------- lines

    def _default_project_index(self, combo: QComboBox) -> int:
        """Index of the user's default project in a project combo, else 0."""
        if self._settings.default_project_id:
            idx = combo.findData(self._settings.default_project_id)
            if idx >= 0:
                return idx
        return 0

    def _make_project_combo(self) -> QComboBox:
        combo = QComboBox()
        combo.addItem("-- None --", None)
        for project in self._projects:
            combo.addItem(project.name, project.id)
        combo.setCurrentIndex(self._default_project_index(combo))
        return combo

    def _make_description_combo(self, project_name: Optional[str]) -> QComboBox:
        combo = QComboBox()
        combo.setEditable(True)  # user can type a fresh description
        combo.addItem("")
        for desc in self._history_for_project(project_name):
            combo.addItem(desc)
        # Default to the top (most recent) suggestion if we have one.
        history = self._history_for_project(project_name)
        if history:
            combo.setCurrentText(history[0])
        else:
            combo.setCurrentText("")
        return combo

    def _add_initial_line(self) -> None:
        """Start with a single line holding all remaining hours."""
        self._add_line(initial_hours=self._remaining)

    def _add_line(self, initial_hours: int = 1) -> None:
        row = self._table.rowCount()
        self._table.insertRow(row)

        # Project
        project_combo = self._make_project_combo()
        project_combo.currentIndexChanged.connect(
            lambda _idx, r=row: self._on_project_changed(r)
        )
        self._table.setCellWidget(row, 0, project_combo)

        # Hours
        hours_spin = QSpinBox()
        hours_spin.setRange(1, self.MAX_ENTRY_HOURS)
        hours_spin.setValue(max(1, min(initial_hours, self.MAX_ENTRY_HOURS)))
        hours_spin.valueChanged.connect(self._update_total)
        self._table.setCellWidget(row, 1, hours_spin)

        # Description (pre-filled from history for the selected project)
        selected_name = project_combo.currentText() if project_combo.currentData() else None
        desc_combo = self._make_description_combo(selected_name)
        self._table.setCellWidget(row, 2, desc_combo)

        # Remove button
        remove_btn = QPushButton("✕")
        remove_btn.setToolTip("Remove this line")
        remove_btn.clicked.connect(lambda _checked, w=remove_btn: self._remove_line(w))
        self._table.setCellWidget(row, 3, remove_btn)

        self._update_total()

    def _remove_line(self, button: QPushButton) -> None:
        # Find the row that currently hosts this button (rows shift on delete).
        for row in range(self._table.rowCount()):
            if self._table.cellWidget(row, 3) is button:
                self._table.removeRow(row)
                break
        self._update_total()

    def _on_project_changed(self, row: int) -> None:
        """Refresh the description suggestions when the project changes."""
        project_combo = self._table.cellWidget(row, 0)
        if not isinstance(project_combo, QComboBox):
            return
        project_name = project_combo.currentText() if project_combo.currentData() else None

        old_desc = self._table.cellWidget(row, 2)
        typed = old_desc.currentText().strip() if isinstance(old_desc, QComboBox) else ""

        new_desc = self._make_description_combo(project_name)
        # Preserve anything the user already typed.
        if typed:
            new_desc.setCurrentText(typed)
        self._table.setCellWidget(row, 2, new_desc)

    # ---------------------------------------------------------------- total

    def _current_total(self) -> int:
        total = 0
        for row in range(self._table.rowCount()):
            spin = self._table.cellWidget(row, 1)
            if isinstance(spin, QSpinBox):
                total += spin.value()
        return total

    def _update_total(self) -> None:
        total = self._current_total()
        ok = (total == self._remaining) and self._table.rowCount() > 0

        color = self.COLORS['success'] if ok else self.COLORS['danger']
        self._total_label.setText(
            f"<span style='color:{color}; font-weight:bold;'>"
            f"Total {total} / {self._remaining}h</span>"
        )
        self._confirm_btn.setEnabled(ok)

    def _distribute_evenly(self) -> None:
        """Spread remaining hours evenly across existing lines."""
        rows = self._table.rowCount()
        if rows == 0:
            return
        base = self._remaining // rows
        extra = self._remaining % rows
        for row in range(rows):
            spin = self._table.cellWidget(row, 1)
            if isinstance(spin, QSpinBox):
                value = base + (1 if row < extra else 0)
                spin.setValue(max(1, min(value, self.MAX_ENTRY_HOURS)))
        self._update_total()

    # -------------------------------------------------------------- confirm

    def _on_confirm(self) -> None:
        if self._current_total() != self._remaining:
            QMessageBox.warning(
                self, "Hours don't match",
                f"The lines total {self._current_total()}h but "
                f"{self._remaining}h remain. Adjust the hours to match."
            )
            return

        # Validate descriptions are present.
        for row in range(self._table.rowCount()):
            desc_combo = self._table.cellWidget(row, 2)
            desc = desc_combo.currentText().strip() if isinstance(desc_combo, QComboBox) else ""
            if not desc:
                QMessageBox.warning(
                    self, "Description required",
                    f"Line {row + 1} needs a description."
                )
                return

        # All good -- persist the entries as drafts.
        saved = 0
        for row in range(self._table.rowCount()):
            project_combo = self._table.cellWidget(row, 0)
            hours_spin = self._table.cellWidget(row, 1)
            desc_combo = self._table.cellWidget(row, 2)

            project_id = project_combo.currentData()
            project = self._store.get_project(project_id) if project_id else None
            hours = hours_spin.value()
            description = desc_combo.currentText().strip()

            entry = TimesheetEntry(
                consultant_id=self._settings.consultant_id,
                entry_date=self._target_date,
                hours=hours,
                ticket_number=project.ticket_number if project else None,
                project_name=project.name if project else None,
                activity_code=self._settings.default_activity_code,
                location=self._settings.default_location,
                description=description,
                status=EntryStatus.DRAFT,
                source=EntrySource.MANUAL,
            )
            self._store.save_entry(entry)
            if project:
                self._store.update_recent_project(project)
            saved += 1

        self._saved_count = saved
        self.accept()

    def get_saved_count(self) -> int:
        """Number of entries saved (0 if cancelled)."""
        return self._saved_count
