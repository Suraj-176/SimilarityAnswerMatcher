"""Reusable UI composition helpers."""

from __future__ import annotations

from auth import check_authentication


def require_authenticated_session() -> bool:
    """Render auth flow gate and return access decision.

    Sidebar user info is rendered by the reference app to avoid duplicate
    Streamlit widget IDs when app.py wraps similarity_app.py.
    """
    return bool(check_authentication())
