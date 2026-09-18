"""Resolves where config.json/counts.json live: next to the running app, not cwd.

A double-clicked .exe (frozen by PyInstaller) can have an unpredictable current
working directory, so both the frozen executable and the plain `python app.py`
script resolve their "living" data files relative to their own location instead.
"""
from __future__ import annotations

import os
import sys


def app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    # This file lives at survivor_app/ui/paths.py; app.py is two levels up, at
    # the repository root.
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def config_path() -> str:
    return os.path.join(app_dir(), "config.json")


def counts_path() -> str:
    return os.path.join(app_dir(), "counts.json")


def bundled_example_config_path() -> str:
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "resources", "example_config.json")
