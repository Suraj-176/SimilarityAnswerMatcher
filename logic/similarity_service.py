"""Core similarity helpers extracted from the reference app."""

from __future__ import annotations

import streamlit as st

from ai_providers import get_provider
from model_configs import get_model_parameters

CROSS_ENCODER_MODELS = [
    "cross-encoder/stsb-roberta-base",
    "cross-encoder/ms-marco-MiniLM-L6-v2",
]


@st.cache_resource
def load_main_model(selected_model: str):
    """Load sentence-transformers model for local similarity mode."""
    try:
        import importlib

        st_mod = importlib.import_module("sentence_transformers")
        sentence_transformer = getattr(st_mod, "SentenceTransformer")
    except Exception as exc:  # pragma: no cover - runtime environment dependent
        raise ImportError(
            "The 'sentence_transformers' package is required for Local Model mode. "
            "Please install dependencies and restart the app."
        ) from exc
    return sentence_transformer(selected_model)


@st.cache_resource
def load_cross_encoder_model(selected_model: str):
    """Load optional cross-encoder model; return None when unavailable."""
    try:
        import importlib

        ce_mod = importlib.import_module("sentence_transformers.cross_encoder")
        cross_encoder = getattr(ce_mod, "CrossEncoder")
    except Exception:  # pragma: no cover - runtime environment dependent
        st.warning(
            "Cross-encoder support is not available because the required package "
            "could not be imported. Falling back to embedding-only similarity."
        )
        return None
    try:
        return cross_encoder(selected_model)
    except Exception as exc:  # pragma: no cover - model download/runtime dependent
        st.warning(f"Could not load cross-encoder model: {exc}")
        return None


def get_provider_parameters(provider_name: str) -> dict:
    """Return active model parameter values from session state with defaults."""
    model_params = get_model_parameters(provider_name)
    return {
        "temperature": st.session_state.get(
            f"{provider_name}_temperature",
            model_params.get("temperature", {}).get("default", 0.0),
        ),
        "top_p": st.session_state.get(
            f"{provider_name}_top_p",
            model_params.get("top_p", {}).get("default", 1.0),
        ),
        "max_tokens": st.session_state.get(
            f"{provider_name}_max_tokens",
            model_params.get("max_tokens", {}).get("default", 20),
        ),
    }


def get_ai_similarity(
    provider_name: str,
    answer1: str,
    answer2: str,
    api_key: str,
    system_prompt: str | None = None,
    user_template: str | None = None,
    temperature: float = 0.0,
    top_p: float = 1.0,
    max_tokens: int = 20,
):
    """Call configured AI provider and return score + explanation."""
    if user_template:
        try:
            prompt = user_template.format(answer1=answer1, answer2=answer2)
        except Exception:
            prompt = f"Compare the following two texts. Text 1: {answer1} Text 2: {answer2}"
    else:
        prompt = (
            "Compare the following two texts and provide a similarity score as a "
            f"percentage. Text 1: {answer1} Text 2: {answer2}"
        )

    sys_content = system_prompt or (
        "You are a helpful assistant. Provide the similarity score by comparing "
        "text 1 and text 2. Respond with only the similarity score as a plain "
        "number (no explanation)."
    )

    try:
        provider = get_provider(provider_name, api_key)
        score, explanation = provider.get_similarity(
            answer1, answer2, sys_content, prompt, temperature, top_p, max_tokens
        )
        return score if score is not None else 0, explanation or ""
    except Exception as exc:  # pragma: no cover - provider/network dependent
        return None, f"Provider error: {exc}"


def get_gpt4o_similarity(
    answer1: str,
    answer2: str,
    api_key: str,
    system_prompt: str | None = None,
    user_template: str | None = None,
    temperature: float = 0.0,
    top_p: float = 1.0,
    max_tokens: int = 20,
):
    """Backwards-compatible helper for legacy Azure-only calls."""
    return get_ai_similarity(
        "Azure OpenAI GPT-4o",
        answer1,
        answer2,
        api_key,
        system_prompt=system_prompt,
        user_template=user_template,
        temperature=temperature,
        top_p=top_p,
        max_tokens=max_tokens,
    )
