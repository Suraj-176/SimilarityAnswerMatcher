"""UI chrome helpers for app-level page setup."""

from __future__ import annotations

import streamlit as st


def configure_page() -> None:
    """Configure Streamlit page metadata once per session."""
    if st.session_state.get("_target_page_configured", False):
        return
    st.set_page_config(page_title="Similarity Answer Matcher", layout="wide")
    st.session_state["_target_page_configured"] = True
