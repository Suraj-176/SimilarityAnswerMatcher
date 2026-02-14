"""
Model-specific configurations for advanced settings.
Each model can have different parameters and constraints.
"""

MODEL_CONFIGS = {
    "Azure OpenAI GPT-4o": {
        "name": "Azure OpenAI GPT-4o",
        "parameters": {
            "temperature": {
                "label": "Temperature",
                "min": 0.0,
                "max": 2.0,
                "step": 0.01,
                "default": 0.0,
                "help": "Controls randomness. 0 = deterministic, higher = more random. Range: 0-2"
            },
            "top_p": {
                "label": "Top-p (Nucleus Sampling)",
                "min": 0.0,
                "max": 1.0,
                "step": 0.01,
                "default": 1.0,
                "help": "Nucleus sampling parameter. Range: 0-1"
            },
            "max_tokens": {
                "label": "Max Tokens",
                "min": 1,
                "max": 4096,
                "step": 1,
                "default": 20,
                "help": "Maximum response length in tokens. Range: 1-4096"
            },
            "frequency_penalty": {
                "label": "Frequency Penalty",
                "min": -2.0,
                "max": 2.0,
                "step": 0.1,
                "default": 0.0,
                "help": "Reduces frequency of repeated tokens. Range: -2 to 2"
            },
            "presence_penalty": {
                "label": "Presence Penalty",
                "min": -2.0,
                "max": 2.0,
                "step": 0.1,
                "default": 0.0,
                "help": "Encourages model to talk about new topics. Range: -2 to 2"
            },
        }
    },
    "Groq": {
        "name": "Groq (Mixtral-8x7b)",
        "parameters": {
            "temperature": {
                "label": "Temperature",
                "min": 0.1,  # Groq requires min 0.1
                "max": 2.0,
                "step": 0.01,
                "default": 0.1,
                "help": "Controls randomness. Groq requires minimum 0.1 (deterministic). Range: 0.1-2"
            },
            "top_p": {
                "label": "Top-p (Nucleus Sampling)",
                "min": 0.0,
                "max": 1.0,
                "step": 0.01,
                "default": 1.0,
                "help": "Nucleus sampling parameter. Range: 0-1"
            },
            "max_tokens": {
                "label": "Max Tokens",
                "min": 1,
                "max": 32768,
                "step": 1,
                "default": 20,
                "help": "Maximum response length in tokens. Groq supports up to 32768"
            },
        }
    },
    "Gemini": {
        "name": "Google Gemini Pro",
        "parameters": {
            "temperature": {
                "label": "Temperature",
                "min": 0.0,
                "max": 2.0,
                "step": 0.01,
                "default": 0.0,
                "help": "Controls randomness. 0 = deterministic. Range: 0-2"
            },
            "top_p": {
                "label": "Top-p (Nucleus Sampling)",
                "min": 0.0,
                "max": 1.0,
                "step": 0.01,
                "default": 0.95,
                "help": "Nucleus sampling parameter. Range: 0-1"
            },
            "top_k": {
                "label": "Top-k",
                "min": 1,
                "max": 100,
                "step": 1,
                "default": 40,
                "help": "Limits to top k most likely tokens. Range: 1-100"
            },
            "max_tokens": {
                "label": "Max Tokens",
                "min": 1,
                "max": 8192,
                "step": 1,
                "default": 20,
                "help": "Maximum response length in tokens. Range: 1-8192"
            },
        }
    },
    "Grok": {
        "name": "xAI Grok Beta",
        "parameters": {
            "temperature": {
                "label": "Temperature",
                "min": 0.0,
                "max": 2.0,
                "step": 0.01,
                "default": 0.5,
                "help": "Controls randomness. 0 = deterministic. Range: 0-2"
            },
            "top_p": {
                "label": "Top-p (Nucleus Sampling)",
                "min": 0.0,
                "max": 1.0,
                "step": 0.01,
                "default": 1.0,
                "help": "Nucleus sampling parameter. Range: 0-1"
            },
            "max_tokens": {
                "label": "Max Tokens",
                "min": 1,
                "max": 32768,
                "step": 1,
                "default": 20,
                "help": "Maximum response length in tokens. Range: 1-32768"
            },
        }
    },
    "Claude 3 Opus": {
        "name": "Anthropic Claude 3 Opus",
        "parameters": {
            "temperature": {
                "label": "Temperature",
                "min": 0.0,
                "max": 1.0,
                "step": 0.01,
                "default": 0.0,
                "help": "Controls randomness. 0 = deterministic. Range: 0-1"
            },
            "max_tokens": {
                "label": "Max Tokens",
                "min": 1,
                "max": 4096,
                "step": 1,
                "default": 20,
                "help": "Maximum response length in tokens. Range: 1-4096"
            },
            "top_p": {
                "label": "Top-p (Nucleus Sampling)",
                "min": 0.0,
                "max": 1.0,
                "step": 0.01,
                "default": 1.0,
                "help": "Nucleus sampling parameter. Range: 0-1"
            },
            "top_k": {
                "label": "Top-k",
                "min": 0,
                "max": 500,
                "step": 1,
                "default": 0,
                "help": "Limits to top k most likely tokens. 0 = disabled. Range: 0-500"
            },
        }
    },
}

# Cloud-provider aliases/new entries used by the current app provider registry.
MODEL_CONFIGS["OpenAI GPT-4o"] = {
    "name": "OpenAI GPT-4o",
    "parameters": {
        "temperature": {
            "label": "Temperature",
            "min": 0.0,
            "max": 2.0,
            "step": 0.01,
            "default": 0.0,
            "help": "Controls randomness. 0 = deterministic. Range: 0-2",
        },
        "top_p": {
            "label": "Top-p (Nucleus Sampling)",
            "min": 0.0,
            "max": 1.0,
            "step": 0.01,
            "default": 1.0,
            "help": "Nucleus sampling parameter. Range: 0-1",
        },
        "max_tokens": {
            "label": "Max Tokens",
            "min": 1,
            "max": 4096,
            "step": 1,
            "default": 20,
            "help": "Maximum response length in tokens. Range: 1-4096",
        },
    },
}

MODEL_CONFIGS["OpenAI GPT-4o-mini"] = {
    "name": "OpenAI GPT-4o-mini",
    "parameters": {
        "temperature": {
            "label": "Temperature",
            "min": 0.0,
            "max": 2.0,
            "step": 0.01,
            "default": 0.0,
            "help": "Controls randomness. 0 = deterministic. Range: 0-2",
        },
        "top_p": {
            "label": "Top-p (Nucleus Sampling)",
            "min": 0.0,
            "max": 1.0,
            "step": 0.01,
            "default": 1.0,
            "help": "Nucleus sampling parameter. Range: 0-1",
        },
        "max_tokens": {
            "label": "Max Tokens",
            "min": 1,
            "max": 4096,
            "step": 1,
            "default": 20,
            "help": "Maximum response length in tokens. Range: 1-4096",
        },
    },
}

MODEL_CONFIGS["OpenRouter"] = {
    "name": "OpenRouter",
    "parameters": {
        "temperature": {
            "label": "Temperature",
            "min": 0.0,
            "max": 2.0,
            "step": 0.01,
            "default": 0.0,
            "help": "Controls randomness. 0 = deterministic. Range: 0-2",
        },
        "top_p": {
            "label": "Top-p (Nucleus Sampling)",
            "min": 0.0,
            "max": 1.0,
            "step": 0.01,
            "default": 1.0,
            "help": "Nucleus sampling parameter. Range: 0-1",
        },
        "max_tokens": {
            "label": "Max Tokens",
            "min": 1,
            "max": 4096,
            "step": 1,
            "default": 20,
            "help": "Maximum response length in tokens. Range: 1-4096",
        },
    },
}

# Friendly names aligned with provider labels in the UI.
MODEL_CONFIGS["Google Gemini"] = MODEL_CONFIGS["Gemini"]
MODEL_CONFIGS["Anthropic Claude"] = MODEL_CONFIGS["Claude 3 Opus"]
MODEL_CONFIGS["xAI Grok"] = MODEL_CONFIGS["Grok"]

# Provider-specific deployment/model choices for Step 2 selection.
PROVIDER_MODEL_OPTIONS = {
    "Azure OpenAI GPT-4o": ["gpt-4o", "gpt-4o-mini", "gpt-4.1", "gpt-4.1-mini"],
    "OpenAI GPT-4o": ["gpt-4o", "gpt-4o-mini", "gpt-4.1", "gpt-4.1-mini"],
    "OpenAI GPT-4o-mini": ["gpt-4o-mini", "gpt-4o", "gpt-4.1-mini"],
    "Google Gemini": ["gemini-2.0-flash", "gemini-1.5-pro", "gemini-1.5-flash", "gemini-pro"],
    "Anthropic Claude": ["claude-3-5-sonnet-latest", "claude-3-7-sonnet-latest", "claude-3-opus-20240229"],
    "Groq": ["llama-3.3-70b-versatile", "mixtral-8x7b-32768", "llama-3.1-8b-instant"],
    "OpenRouter": ["openai/gpt-4o-mini", "openai/gpt-4o", "anthropic/claude-3.5-sonnet"],
    "xAI Grok": ["grok-beta", "grok-2-latest", "grok-2-mini-latest"],
}


def get_provider_model_options(provider_name: str) -> list:
    """Return model/deployment options for a provider."""
    return PROVIDER_MODEL_OPTIONS.get(provider_name, [])


def get_default_provider_model(provider_name: str) -> str:
    """Return default model/deployment for a provider."""
    opts = get_provider_model_options(provider_name)
    return opts[0] if opts else ""


def get_model_config(model_name: str) -> dict:
    """Get configuration for a specific model."""
    return MODEL_CONFIGS.get(model_name, MODEL_CONFIGS["Azure OpenAI GPT-4o"])


def get_model_parameters(model_name: str) -> dict:
    """Get parameters configuration for a specific model."""
    config = get_model_config(model_name)
    return config.get("parameters", {})


def get_parameter_info(model_name: str, param_name: str) -> dict:
    """Get specific parameter info for a model."""
    params = get_model_parameters(model_name)
    return params.get(param_name, {})
