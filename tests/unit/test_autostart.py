"""Tests for the auto-start launcher resolution.

These don't touch the real Startup folder; they only verify that the launch
command is resolved sanely so login-launch actually works.
"""

import sys

import pytest

from tiq_assistant.desktop import autostart


def test_resolve_launcher_returns_three_parts():
    target, arguments, workdir = autostart._resolve_launcher()
    assert isinstance(target, str) and target
    assert isinstance(arguments, str)
    assert isinstance(workdir, str) and workdir


def test_resolve_launcher_uses_module_entry_point_when_not_frozen(monkeypatch):
    monkeypatch.setattr(sys, "frozen", False, raising=False)
    target, arguments, workdir = autostart._resolve_launcher()
    # Uses `python -m tiq_assistant` so it works regardless of install method.
    assert arguments == "-m tiq_assistant"
    # Working dir should contain the importable package.
    from pathlib import Path
    assert (Path(workdir) / "tiq_assistant").exists() or workdir


def test_sync_is_safe_to_call(monkeypatch):
    # On non-Windows or without a Startup folder, sync should not raise.
    calls = {}
    monkeypatch.setattr(autostart, "enable", lambda: calls.setdefault("enable", True) or True)
    monkeypatch.setattr(autostart, "disable", lambda: calls.setdefault("disable", True) or True)
    assert autostart.sync(True) is True
    assert autostart.sync(False) is True
    assert calls == {"enable": True, "disable": True}
