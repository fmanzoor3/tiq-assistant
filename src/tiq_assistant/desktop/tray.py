"""System tray functionality for TIQ Assistant desktop app."""

import sys
from typing import Optional, Callable
from PyQt6.QtWidgets import (
    QSystemTrayIcon, QMenu, QApplication, QMessageBox
)
from PyQt6.QtGui import QIcon, QAction
from PyQt6.QtCore import QObject, pyqtSignal

from tiq_assistant.desktop.windows.day_entry_dialog import SessionType
from tiq_assistant.desktop.icon import create_app_icon


class TrayIconManager(QObject):
    """
    Manages the system tray icon and context menu.

    Signals:
        morning_entry_requested: Emitted when user clicks "Morning Entry"
        afternoon_entry_requested: Emitted when user clicks "Afternoon Entry"
        sync_requested: Emitted when user clicks "Sync Outlook"
        settings_requested: Emitted when user clicks "Settings"
        quit_requested: Emitted when user clicks "Exit"
    """

    morning_entry_requested = pyqtSignal()
    afternoon_entry_requested = pyqtSignal()
    sync_requested = pyqtSignal()
    settings_requested = pyqtSignal()
    dashboard_requested = pyqtSignal()
    quit_requested = pyqtSignal()

    def __init__(self, parent: Optional[QObject] = None):
        super().__init__(parent)

        self._tray_icon: Optional[QSystemTrayIcon] = None
        self._menu: Optional[QMenu] = None

    def setup(self, app: QApplication) -> bool:
        """
        Set up the system tray icon.

        Returns:
            True if system tray is available, False otherwise
        """
        if not QSystemTrayIcon.isSystemTrayAvailable():
            return False

        # Create tray icon
        self._tray_icon = QSystemTrayIcon(parent=None)

        # Set icon (use a default icon for now)
        icon = self._create_default_icon()
        self._tray_icon.setIcon(icon)

        # Create context menu
        self._menu = self._create_menu()
        self._tray_icon.setContextMenu(self._menu)

        # Set tooltip
        self._tray_icon.setToolTip("TIQ Assistant - Timesheet Helper")

        # Connect double-click to show dashboard
        self._tray_icon.activated.connect(self._on_activated)

        return True

    def show(self) -> None:
        """Show the tray icon."""
        if self._tray_icon:
            self._tray_icon.show()

    def hide(self) -> None:
        """Hide the tray icon."""
        if self._tray_icon:
            self._tray_icon.hide()

    def show_notification(
        self,
        title: str,
        message: str,
        icon_type: QSystemTrayIcon.MessageIcon = QSystemTrayIcon.MessageIcon.Information,
        duration_ms: int = 5000
    ) -> None:
        """Show a system tray notification."""
        if self._tray_icon:
            self._tray_icon.showMessage(title, message, icon_type, duration_ms)

    def show_popup_reminder(self, session: SessionType) -> None:
        """Show a notification reminding user to log time."""
        session_name = "morning" if session == SessionType.MORNING else "afternoon"
        self.show_notification(
            f"Time to log {session_name} hours!",
            f"Click to open the {session_name} time entry popup.",
            QSystemTrayIcon.MessageIcon.Information
        )

    def _create_default_icon(self) -> QIcon:
        """Create the app icon using the shared function."""
        return create_app_icon()

    def _create_menu(self) -> QMenu:
        """Create the context menu."""
        menu = QMenu()

        # Dashboard action
        dashboard_action = QAction("Show Dashboard", menu)
        dashboard_action.triggered.connect(self.dashboard_requested.emit)
        menu.addAction(dashboard_action)

        menu.addSeparator()

        # Time entry actions
        morning_action = QAction("Morning Entry (12:15)", menu)
        morning_action.triggered.connect(self.morning_entry_requested.emit)
        menu.addAction(morning_action)

        afternoon_action = QAction("Afternoon Entry (18:15)", menu)
        afternoon_action.triggered.connect(self.afternoon_entry_requested.emit)
        menu.addAction(afternoon_action)

        menu.addSeparator()

        # Sync action
        sync_action = QAction("Sync Outlook Calendar", menu)
        sync_action.triggered.connect(self.sync_requested.emit)
        menu.addAction(sync_action)

        menu.addSeparator()

        # Settings action
        settings_action = QAction("Settings", menu)
        settings_action.triggered.connect(self.settings_requested.emit)
        menu.addAction(settings_action)

        menu.addSeparator()

        # Exit action
        exit_action = QAction("Exit", menu)
        exit_action.triggered.connect(self.quit_requested.emit)
        menu.addAction(exit_action)

        return menu

    def _on_activated(self, reason: QSystemTrayIcon.ActivationReason) -> None:
        """Handle tray icon activation (click, double-click, etc.)."""
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self.dashboard_requested.emit()
        elif reason == QSystemTrayIcon.ActivationReason.Trigger:
            # Single click - could show menu or do nothing
            pass
