"""Reliable Windows auto-start (launch on login) for TIQ Assistant.

This module creates/removes a shortcut in the user's Startup folder so the
app launches automatically when the user logs in.

Why a Startup-folder shortcut instead of a registry Run key:
- It uses ``pythonw.exe`` so no console window flashes on login.
- It records an explicit working directory, so ``python -m tiq_assistant``
  resolves the package correctly regardless of how it was installed
  (editable install, global install, or just the source folder on disk).
- It is self-detecting: it figures out the right interpreter and launch
  command at write time, so it keeps working if the folder is moved and the
  shortcut is re-created.
- It is visible and user-removable (shell:startup), which matches the app's
  "no hidden persistence" security posture.

All functions are safe to call on non-Windows platforms (they no-op / return
False) and never raise -- auto-start is a convenience, not a hard dependency.
"""

from __future__ import annotations

import logging
import os
import sys
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

SHORTCUT_NAME = "TIQ Assistant.lnk"


def _is_windows() -> bool:
    return sys.platform.startswith("win")


def get_startup_dir() -> Optional[Path]:
    """Return the current user's Startup folder, or None if unavailable."""
    if not _is_windows():
        return None

    appdata = os.environ.get("APPDATA")
    if not appdata:
        return None

    startup = Path(appdata) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"
    return startup


def _get_shortcut_path() -> Optional[Path]:
    startup = get_startup_dir()
    if startup is None:
        return None
    return startup / SHORTCUT_NAME


def _resolve_launcher() -> tuple[str, str, str]:
    """Work out how to launch the app on this machine.

    Returns a tuple of (target, arguments, working_directory) suitable for a
    Windows shortcut.

    Strategy:
    - If running as a frozen executable (PyInstaller etc.), launch it directly.
    - Otherwise prefer ``pythonw.exe`` (no console) next to the current
      interpreter and run ``-m tiq_assistant`` so the package's own
      ``__main__`` selects desktop mode.
    - The working directory is set to the project root (two levels above the
      package) so an on-disk-only checkout still imports correctly.
    """
    # Frozen / compiled executable
    if getattr(sys, "frozen", False):
        exe = sys.executable
        return exe, "", str(Path(exe).parent)

    # Prefer the windowed interpreter to avoid a console window at login.
    interpreter = Path(sys.executable)
    pythonw = interpreter.with_name("pythonw.exe")
    target = str(pythonw if pythonw.exists() else interpreter)

    # Project root = .../src/tiq_assistant/desktop/autostart.py -> up 4 = repo root
    # (src/tiq_assistant/desktop -> src/tiq_assistant -> src -> repo root)
    package_root = Path(__file__).resolve().parents[1]  # .../src/tiq_assistant
    src_root = package_root.parent                       # .../src
    project_root = src_root.parent                       # repo root

    # Use the package entry point; -m works from the project root even for a
    # source-only checkout because src/ is added, but to be safe we cd to the
    # directory that contains the importable package.
    working_dir = str(src_root if (src_root / "tiq_assistant").exists() else project_root)
    arguments = "-m tiq_assistant"

    return target, arguments, working_dir


def is_enabled() -> bool:
    """Return True if the auto-start shortcut currently exists."""
    shortcut = _get_shortcut_path()
    return bool(shortcut and shortcut.exists())


def enable() -> bool:
    """Create the Startup-folder shortcut. Returns True on success."""
    if not _is_windows():
        logger.info("Auto-start not supported on this platform; skipping.")
        return False

    shortcut = _get_shortcut_path()
    if shortcut is None:
        logger.warning("Could not resolve Startup folder; auto-start not enabled.")
        return False

    target, arguments, working_dir = _resolve_launcher()

    try:
        shortcut.parent.mkdir(parents=True, exist_ok=True)
        _create_shortcut(shortcut, target, arguments, working_dir)
        logger.info("Auto-start enabled: %s -> %s %s", shortcut, target, arguments)
        return True
    except Exception as e:  # noqa: BLE001 - never fatal
        logger.warning("Failed to enable auto-start: %s", e)
        return False


def disable() -> bool:
    """Remove the Startup-folder shortcut. Returns True on success (or absent)."""
    shortcut = _get_shortcut_path()
    if shortcut is None:
        return False
    try:
        if shortcut.exists():
            shortcut.unlink()
            logger.info("Auto-start disabled (shortcut removed).")
        return True
    except Exception as e:  # noqa: BLE001
        logger.warning("Failed to disable auto-start: %s", e)
        return False


def sync(enabled: bool) -> bool:
    """Ensure the on-disk state matches ``enabled``.

    Call this on every app launch (with the stored config value) so auto-start
    is repaired if the shortcut was lost, or refreshed if the app moved.

    Returns True if the desired state was achieved.
    """
    if enabled:
        # Re-create even if present, so a moved folder gets a corrected target.
        return enable()
    return disable()


def _create_shortcut(
    shortcut_path: Path,
    target: str,
    arguments: str,
    working_dir: str,
) -> None:
    """Create a .lnk file, preferring pywin32, falling back to PowerShell."""
    # Preferred: pywin32 (already a dependency of this project).
    try:
        import pythoncom  # type: ignore
        from win32com.client import Dispatch  # type: ignore

        shell = Dispatch("WScript.Shell")
        link = shell.CreateShortcut(str(shortcut_path))
        link.TargetPath = target
        link.Arguments = arguments
        link.WorkingDirectory = working_dir
        link.WindowStyle = 7  # minimized
        link.Description = "TIQ Assistant - Timesheet Helper"
        link.Save()
        return
    except Exception as e:  # noqa: BLE001
        logger.debug("pywin32 shortcut creation failed (%s); trying PowerShell.", e)

    # Fallback: drive PowerShell's WScript.Shell (works even without pywin32).
    import subprocess

    ps_target = target.replace("'", "''")
    ps_args = arguments.replace("'", "''")
    ps_wd = working_dir.replace("'", "''")
    ps_path = str(shortcut_path).replace("'", "''")

    script = (
        "$s = (New-Object -ComObject WScript.Shell).CreateShortcut('{path}');"
        "$s.TargetPath = '{target}';"
        "$s.Arguments = '{args}';"
        "$s.WorkingDirectory = '{wd}';"
        "$s.WindowStyle = 7;"
        "$s.Description = 'TIQ Assistant - Timesheet Helper';"
        "$s.Save()"
    ).format(path=ps_path, target=ps_target, args=ps_args, wd=ps_wd)

    subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
        check=True,
        capture_output=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
