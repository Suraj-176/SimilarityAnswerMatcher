"""Service layer for launching comparison workflows."""

from __future__ import annotations

import runpy
from pathlib import Path


def get_reference_app_path() -> Path:
    """Path to the validated reference app implementation."""
    return Path(__file__).resolve().parents[1] / "similarity_app.py"


def run_reference_comparison_app() -> None:
    """Execute the reference Streamlit app to preserve exact behavior."""
    app_path = get_reference_app_path()
    if not app_path.exists():
        raise FileNotFoundError(f"Reference app not found: {app_path}")
    runpy.run_path(str(app_path), run_name="__main__")
