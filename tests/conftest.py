"""Shared pytest fixtures for the helper.py test suite.

Ensures:
  * matplotlib uses the non-interactive Agg backend.
  * Any figure created during a test is closed after the test.
  * ``helper._DRIVE_BASE`` / ``helper._LOCAL_BASE`` point to a tmp dir,
    so ``speichere_fortschritt`` writes to an isolated location.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

import warnings

import matplotlib
matplotlib.use("Agg")  # set before helper.py imports pyplot
import matplotlib.pyplot as plt
import pytest

# plt.show() auf Agg ist ein No-Op + Warning; da helper.py show() nach jeder
# Visualisierung aufruft, filtern wir diese speziell heraus.
warnings.filterwarnings(
    "ignore",
    message="FigureCanvasAgg is non-interactive",
    category=UserWarning,
)

# Repo root onto sys.path so `import helper` works without installation.
REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import helper  # noqa: E402


def pytest_configure(config):
    config.addinivalue_line(
        "filterwarnings",
        "ignore:FigureCanvasAgg is non-interactive:UserWarning",
    )


@pytest.fixture
def drive_base(tmp_path, monkeypatch):
    """Redirect persistence (drive + local) to an isolated tmp dir."""
    base = tmp_path / "DMAIC_Daten"
    base.mkdir()
    monkeypatch.setattr(helper, "_DRIVE_BASE", str(base), raising=True)
    monkeypatch.setattr(helper, "_LOCAL_BASE", str(base), raising=True)
    return base


@pytest.fixture(autouse=True)
def close_figures():
    """Close all matplotlib figures after each test to avoid state leaks."""
    yield
    plt.close("all")
