"""Settings dialog for TIQ Assistant desktop app."""

from typing import Optional

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QLineEdit, QSpinBox, QCheckBox, QGroupBox, QWidget,
    QFormLayout, QMessageBox, QTimeEdit, QTabWidget, QComboBox
)
from PyQt6.QtCore import QTime

from tiq_assistant.core.models import ScheduleConfig, UserSettings
from tiq_assistant.storage.sqlite_store import get_store
from tiq_assistant.desktop.scheduler import SchedulerManager


class SettingsDialog(QDialog):
    """Settings dialog for configuring the desktop app."""

    def __init__(
        self,
        scheduler: Optional[SchedulerManager] = None,
        parent: Optional[QWidget] = None
    ):
        super().__init__(parent)

        self._scheduler = scheduler
        self._store = get_store()

        self._setup_ui()
        self._load_settings()

    def _setup_ui(self) -> None:
        """Set up the dialog UI."""
        self.setWindowTitle("TIQ Assistant Settings")
        self.setMinimumWidth(450)

        layout = QVBoxLayout(self)

        # Tab widget for different settings categories
        tabs = QTabWidget()
        layout.addWidget(tabs)

        # Schedule tab
        schedule_tab = self._create_schedule_tab()
        tabs.addTab(schedule_tab, "Schedule")

        # User tab
        user_tab = self._create_user_tab()
        tabs.addTab(user_tab, "User")

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)

        save_btn = QPushButton("Save")
        save_btn.setDefault(True)
        save_btn.clicked.connect(self._save_settings)
        button_layout.addWidget(save_btn)

        layout.addLayout(button_layout)

    def _create_schedule_tab(self) -> QWidget:
        """Create the schedule settings tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Morning popup settings
        morning_group = QGroupBox("Morning Popup")
        morning_layout = QFormLayout(morning_group)

        self._morning_enabled = QCheckBox("Enabled")
        morning_layout.addRow("", self._morning_enabled)

        self._morning_time = QTimeEdit()
        self._morning_time.setDisplayFormat("HH:mm")
        morning_layout.addRow("Time:", self._morning_time)

        self._morning_hours = QSpinBox()
        self._morning_hours.setRange(1, 8)
        morning_layout.addRow("Target Hours:", self._morning_hours)

        layout.addWidget(morning_group)

        # Afternoon popup settings
        afternoon_group = QGroupBox("Afternoon Popup")
        afternoon_layout = QFormLayout(afternoon_group)

        self._afternoon_enabled = QCheckBox("Enabled")
        afternoon_layout.addRow("", self._afternoon_enabled)

        self._afternoon_time = QTimeEdit()
        self._afternoon_time.setDisplayFormat("HH:mm")
        afternoon_layout.addRow("Time:", self._afternoon_time)

        self._afternoon_hours = QSpinBox()
        self._afternoon_hours.setRange(1, 10)
        afternoon_layout.addRow("Target Hours:", self._afternoon_hours)

        layout.addWidget(afternoon_group)

        # Workday settings
        workday_group = QGroupBox("Workday")
        workday_layout = QFormLayout(workday_group)

        self._workday_start = QTimeEdit()
        self._workday_start.setDisplayFormat("HH:mm")
        workday_layout.addRow("Start:", self._workday_start)

        self._lunch_start = QTimeEdit()
        self._lunch_start.setDisplayFormat("HH:mm")
        workday_layout.addRow("Lunch Start:", self._lunch_start)

        self._lunch_end = QTimeEdit()
        self._lunch_end.setDisplayFormat("HH:mm")
        workday_layout.addRow("Lunch End:", self._lunch_end)

        self._workday_end = QTimeEdit()
        self._workday_end.setDisplayFormat("HH:mm")
        workday_layout.addRow("End:", self._workday_end)

        layout.addWidget(workday_group)

        # Auto-start
        self._auto_start = QCheckBox("Start with Windows")
        layout.addWidget(self._auto_start)

        layout.addStretch()

        return widget

    def _create_user_tab(self) -> QWidget:
        """Create the user settings tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # User info
        user_group = QGroupBox("User Information")
        user_layout = QFormLayout(user_group)

        self._consultant_id = QLineEdit()
        user_layout.addRow("Consultant ID:", self._consultant_id)

        self._default_location = QLineEdit()
        user_layout.addRow("Default Location:", self._default_location)

        # Default project used for quick entry and auto-fill top-ups.
        self._default_project = QComboBox()
        self._default_project.addItem("-- None --", None)
        for project in self._store.get_projects():
            self._default_project.addItem(project.name, project.id)
        user_layout.addRow("Default Project:", self._default_project)

        layout.addWidget(user_group)

        # Matching settings
        matching_group = QGroupBox("Matching")
        matching_layout = QFormLayout(matching_group)

        self._skip_canceled = QCheckBox("Skip canceled meetings")
        matching_layout.addRow("", self._skip_canceled)

        self._min_duration = QSpinBox()
        self._min_duration.setRange(5, 60)
        self._min_duration.setSuffix(" minutes")
        matching_layout.addRow("Min Meeting Duration:", self._min_duration)

        layout.addWidget(matching_group)

        layout.addStretch()

        return widget

    def _load_settings(self) -> None:
        """Load current settings into the form."""
        # Load schedule config
        config = self._store.get_schedule_config()

        self._morning_enabled.setChecked(config.morning_popup_enabled)
        self._morning_time.setTime(self._parse_time(config.morning_popup_time))
        self._morning_hours.setValue(config.morning_hours_target)

        self._afternoon_enabled.setChecked(config.afternoon_popup_enabled)
        self._afternoon_time.setTime(self._parse_time(config.afternoon_popup_time))
        self._afternoon_hours.setValue(config.afternoon_hours_target)

        self._workday_start.setTime(self._parse_time(config.workday_start))
        self._lunch_start.setTime(self._parse_time(config.lunch_start))
        self._lunch_end.setTime(self._parse_time(config.lunch_end))
        self._workday_end.setTime(self._parse_time(config.workday_end))

        self._auto_start.setChecked(config.auto_start_with_windows)

        # Load user settings
        settings = self._store.get_settings()

        self._consultant_id.setText(settings.consultant_id)
        self._default_location.setText(settings.default_location)
        self._skip_canceled.setChecked(settings.skip_canceled_meetings)
        self._min_duration.setValue(settings.min_meeting_duration_minutes)

        # Select the saved default project, if any.
        if settings.default_project_id:
            idx = self._default_project.findData(settings.default_project_id)
            if idx >= 0:
                self._default_project.setCurrentIndex(idx)

    def _save_settings(self) -> None:
        """Save settings to the database."""
        # Build schedule config
        config = ScheduleConfig(
            morning_popup_enabled=self._morning_enabled.isChecked(),
            morning_popup_time=self._morning_time.time().toString("HH:mm"),
            morning_hours_target=self._morning_hours.value(),
            afternoon_popup_enabled=self._afternoon_enabled.isChecked(),
            afternoon_popup_time=self._afternoon_time.time().toString("HH:mm"),
            afternoon_hours_target=self._afternoon_hours.value(),
            workday_start=self._workday_start.time().toString("HH:mm"),
            lunch_start=self._lunch_start.time().toString("HH:mm"),
            lunch_end=self._lunch_end.time().toString("HH:mm"),
            workday_end=self._workday_end.time().toString("HH:mm"),
            auto_start_with_windows=self._auto_start.isChecked(),
        )

        # Build user settings
        settings = self._store.get_settings()
        settings.consultant_id = self._consultant_id.text().strip() or "FMANZOOR"
        settings.default_location = self._default_location.text().strip() or "ANKARA"
        settings.default_project_id = self._default_project.currentData()
        settings.skip_canceled_meetings = self._skip_canceled.isChecked()
        settings.min_meeting_duration_minutes = self._min_duration.value()

        # Save to database
        self._store.save_schedule_config(config)
        self._store.save_settings(settings)

        # Update scheduler if running
        if self._scheduler and self._scheduler.is_running:
            self._scheduler.reschedule(config)

        # Handle auto-start and report if it couldn't be applied.
        auto_ok = self._update_auto_start(config.auto_start_with_windows)
        if config.auto_start_with_windows and not auto_ok:
            QMessageBox.warning(
                self, "Auto-start",
                "Settings were saved, but the 'Start with Windows' shortcut "
                "could not be created. You can still start the app manually."
            )
        else:
            QMessageBox.information(self, "Settings Saved", "Settings have been saved.")
        self.accept()

    def _parse_time(self, time_str: str) -> QTime:
        """Parse a time string like '12:30' into a QTime."""
        parts = time_str.split(":")
        return QTime(int(parts[0]), int(parts[1]))

    def _update_auto_start(self, enabled: bool) -> bool:
        """Update Windows auto-start via a reliable Startup-folder shortcut.

        Delegates to :mod:`tiq_assistant.desktop.autostart`, which creates a
        ``pythonw`` shortcut (no console window) with an explicit working
        directory so it works regardless of how the app was installed.

        Returns True if the desired state was achieved.
        """
        try:
            from tiq_assistant.desktop import autostart
            return autostart.sync(enabled)
        except Exception as e:
            print(f"Could not update auto-start setting: {e}")
            return False
