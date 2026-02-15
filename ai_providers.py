"""
AI Provider abstraction layer for prompt-based similarity matching.
Supports: Azure OpenAI, OpenAI, Google Gemini, Anthropic Claude, Groq,
OpenRouter, and xAI Grok.
"""

import logging
import re
from abc import ABC, abstractmethod
from typing import Optional, Tuple

import requests

logger = logging.getLogger(__name__)


class AIProvider(ABC):
    """Base class for all AI providers."""

    def __init__(self, api_key: str):
        self.api_key = api_key

    @abstractmethod
    def get_similarity(
        self,
        answer1: str,
        answer2: str,
        system_prompt: str,
        user_template: str,
        temperature: float = 0.0,
        top_p: float = 1.0,
        max_tokens: int = 20,
        model_name: Optional[str] = None,
    ) -> Tuple[Optional[float], str]:
        """
        Get similarity score between two texts.
        Returns: (score_float_or_0, explanation_string)
        """

    @staticmethod
    def extract_score(content: str) -> float:
        """Extract numerical score from response."""
        try:
            content_str = str(content).strip()
            logger.debug("Extracting score from: %s", content_str[:100])

            # First preference: explicit percentages like "87%" or "87.5%".
            pct_matches = re.findall(r"(\d{1,3}(?:\.\d+)?)\s*%", content_str)
            for token in reversed(pct_matches):
                score = float(token)
                if 0 <= score <= 100:
                    logger.debug("Extracted percent score: %s", score)
                    return score

            # Fallback: plain numbers.
            # Use the last valid number, because models often include context before final answer.
            num_matches = re.findall(r"-?\d+(?:\.\d+)?", content_str)
            for token in reversed(num_matches):
                score = float(token)
                if score < 0:
                    continue
                # Many models return a ratio in [0, 1]. Convert to percentage.
                if 0 <= score <= 1 and "." in token:
                    converted = round(score * 100, 2)
                    logger.debug("Extracted ratio score %s -> %s%%", score, converted)
                    return converted
                if 0 <= score <= 100:
                    logger.debug("Extracted numeric score: %s", score)
                    return score

            # Fallback: extract digits
            digits = "".join(filter(str.isdigit, content_str))
            if digits:
                score = float(digits[:3] or 0)
                logger.debug("Extracted score from digits: %s", score)
                return score

            logger.warning("Could not extract score from: %s", content_str[:100])
            return 0.0
        except Exception as e:
            logger.error("Error extracting score: %s", e)
            return 0.0


class AzureOpenAIProvider(AIProvider):
    """Azure OpenAI GPT provider."""

    def get_similarity(
        self,
        answer1: str,
        answer2: str,
        system_prompt: str,
        user_template: str,
        temperature: float = 0.0,
        top_p: float = 1.0,
        max_tokens: int = 20,
        model_name: Optional[str] = None,
    ) -> Tuple[Optional[float], str]:
        deployment_name = (model_name or "gpt-4o").strip() or "gpt-4o"
        url = (
            "https://f2fdevopenai.openai.azure.com/openai/deployments/"
            f"{deployment_name}/chat/completions?api-version=2024-05-01-preview"
        )
        headers = {
            "Content-Type": "application/json",
            "api-key": self.api_key,
        }

        prompt = user_template.format(answer1=answer1, answer2=answer2)
        data = {
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt},
            ],
            "temperature": float(temperature),
            "top_p": float(top_p),
            "max_tokens": int(max_tokens),
        }

        try:
            response = requests.post(url, headers=headers, json=data, timeout=30)
            response.raise_for_status()
            result = response.json()
            content = result["choices"][0]["message"]["content"]
            score = self.extract_score(content)
            return score, content.strip()
        except Exception as e:
            return None, f"API error: {e}"


class OpenAIGPT4oProvider(AIProvider):
    """OpenAI chat-completions provider."""

    MODEL_NAME = "gpt-4o"

    def get_similarity(
        self,
        answer1: str,
        answer2: str,
        system_prompt: str,
        user_template: str,
        temperature: float = 0.0,
        top_p: float = 1.0,
        max_tokens: int = 20,
        model_name: Optional[str] = None,
    ) -> Tuple[Optional[float], str]:
        url = "https://api.openai.com/v1/chat/completions"
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        prompt = user_template.format(answer1=answer1, answer2=answer2)
        data = {
            "model": model_name or self.MODEL_NAME,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt},
            ],
            "temperature": float(temperature),
            "top_p": float(top_p),
            "max_tokens": int(max_tokens),
        }
        try:
            response = requests.post(url, headers=headers, json=data, timeout=30)
            response.raise_for_status()
            result = response.json()
            content = result["choices"][0]["message"]["content"]
            score = self.extract_score(content)
            return score, content.strip()
        except Exception as e:
            return None, f"OpenAI API error: {e}"


class OpenAIGPT4oMiniProvider(OpenAIGPT4oProvider):
    """OpenAI GPT-4o-mini provider."""

    MODEL_NAME = "gpt-4o-mini"


class GroqProvider(AIProvider):
    """Groq API provider."""

    def get_similarity(
        self,
        answer1: str,
        answer2: str,
        system_prompt: str,
        user_template: str,
        temperature: float = 0.0,
        top_p: float = 1.0,
        max_tokens: int = 20,
        model_name: Optional[str] = None,
    ) -> Tuple[Optional[float], str]:
        url = "https://api.groq.com/openai/v1/chat/completions"
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }

        prompt = user_template.format(answer1=answer1, answer2=answer2)
        data = {
            "model": model_name or "mixtral-8x7b-32768",
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt},
            ],
            "temperature": float(max(0.1, temperature)),
            "max_tokens": int(max_tokens),
        }

        # Groq doesn't support top_p when temperature is very low
        if temperature > 0:
            data["top_p"] = float(top_p)

        try:
            response = requests.post(url, headers=headers, json=data, timeout=30)
            response.raise_for_status()
            result = response.json()

            if "choices" not in result or not result["choices"]:
                error_msg = f"Groq: No choices in response: {result}"
                logger.error(error_msg)
                return None, error_msg

            content = result["choices"][0]["message"]["content"]
            logger.debug("Groq raw response: %s", content)
            score = self.extract_score(content)
            return score, content.strip()
        except requests.exceptions.RequestException as e:
            error_msg = f"Groq API error: {str(e)}"
            logger.error(error_msg)
            return None, error_msg
        except (KeyError, IndexError) as e:
            error_msg = f"Groq response parsing error: {str(e)}"
            logger.error(error_msg)
            return None, error_msg
        except Exception as e:
            error_msg = f"Groq unexpected error: {str(e)}"
            logger.error(error_msg)
            return None, error_msg


class GeminiProvider(AIProvider):
    """Google Gemini API provider."""

    def get_similarity(
        self,
        answer1: str,
        answer2: str,
        system_prompt: str,
        user_template: str,
        temperature: float = 0.0,
        top_p: float = 1.0,
        max_tokens: int = 20,
        model_name: Optional[str] = None,
    ) -> Tuple[Optional[float], str]:
        gemini_model = (model_name or "gemini-pro").strip() or "gemini-pro"
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{gemini_model}:generateContent"
        headers = {"Content-Type": "application/json"}

        prompt = user_template.format(answer1=answer1, answer2=answer2)
        full_prompt = f"{system_prompt}\n\n{prompt}"

        data = {
            "contents": [{"parts": [{"text": full_prompt}]}],
            "generationConfig": {
                "temperature": float(temperature),
                "topP": float(top_p),
                "maxOutputTokens": int(max_tokens),
            },
        }

        try:
            response = requests.post(
                f"{url}?key={self.api_key}",
                headers=headers,
                json=data,
                timeout=30,
            )
            response.raise_for_status()
            result = response.json()

            logger.debug("Gemini raw response: %s", result)

            if "error" in result:
                error_msg = f"Gemini API error: {result['error'].get('message', 'Unknown error')}"
                logger.error(error_msg)
                return None, error_msg

            if "candidates" not in result or not result["candidates"]:
                error_msg = f"Gemini: No candidates in response: {result}"
                logger.error(error_msg)
                return None, error_msg

            candidate = result["candidates"][0]
            if "content" not in candidate or "parts" not in candidate["content"] or not candidate["content"]["parts"]:
                error_msg = f"Gemini: Invalid candidate structure: {candidate}"
                logger.error(error_msg)
                return None, error_msg

            content = candidate["content"]["parts"][0]["text"]
            logger.debug("Gemini parsed content: %s", content)
            score = self.extract_score(content)
            return score, content.strip()
        except requests.exceptions.RequestException as e:
            error_msg = f"Gemini API error: {str(e)}"
            logger.error(error_msg)
            return None, error_msg
        except (KeyError, IndexError) as e:
            error_msg = f"Gemini response parsing error: {str(e)}"
            logger.error(error_msg)
            return None, error_msg
        except Exception as e:
            error_msg = f"Gemini unexpected error: {str(e)}"
            logger.error(error_msg)
            return None, error_msg


class GrokProvider(AIProvider):
    """Grok API provider (via xAI)."""

    def get_similarity(
        self,
        answer1: str,
        answer2: str,
        system_prompt: str,
        user_template: str,
        temperature: float = 0.0,
        top_p: float = 1.0,
        max_tokens: int = 20,
        model_name: Optional[str] = None,
    ) -> Tuple[Optional[float], str]:
        url = "https://api.x.ai/v1/chat/completions"
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }

        prompt = user_template.format(answer1=answer1, answer2=answer2)
        data = {
            "model": model_name or "grok-beta",
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt},
            ],
            "temperature": float(temperature),
            "top_p": float(top_p),
            "max_tokens": int(max_tokens),
        }

        try:
            response = requests.post(url, headers=headers, json=data, timeout=30)
            response.raise_for_status()
            result = response.json()
            content = result["choices"][0]["message"]["content"]
            score = self.extract_score(content)
            return score, content.strip()
        except Exception as e:
            return None, f"API error: {e}"


class ClaudeProvider(AIProvider):
    """Anthropic Claude API provider."""

    def get_similarity(
        self,
        answer1: str,
        answer2: str,
        system_prompt: str,
        user_template: str,
        temperature: float = 0.0,
        top_p: float = 1.0,
        max_tokens: int = 20,
        model_name: Optional[str] = None,
    ) -> Tuple[Optional[float], str]:
        url = "https://api.anthropic.com/v1/messages"
        headers = {
            "x-api-key": self.api_key,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        }

        prompt = user_template.format(answer1=answer1, answer2=answer2)
        data = {
            "model": model_name or "claude-3-opus-20240229",
            "max_tokens": int(max_tokens),
            "system": system_prompt,
            "temperature": float(temperature),
            "top_p": float(top_p),
            "messages": [{"role": "user", "content": prompt}],
        }

        try:
            response = requests.post(url, headers=headers, json=data, timeout=30)
            response.raise_for_status()
            result = response.json()
            content = result["content"][0]["text"]
            score = self.extract_score(content)
            return score, content.strip()
        except Exception as e:
            return None, f"API error: {e}"


class OpenRouterProvider(AIProvider):
    """OpenRouter provider (OpenAI-compatible API)."""

    def get_similarity(
        self,
        answer1: str,
        answer2: str,
        system_prompt: str,
        user_template: str,
        temperature: float = 0.0,
        top_p: float = 1.0,
        max_tokens: int = 20,
        model_name: Optional[str] = None,
    ) -> Tuple[Optional[float], str]:
        url = "https://openrouter.ai/api/v1/chat/completions"
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        prompt = user_template.format(answer1=answer1, answer2=answer2)
        data = {
            "model": model_name or "openai/gpt-4o-mini",
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt},
            ],
            "temperature": float(temperature),
            "top_p": float(top_p),
            "max_tokens": int(max_tokens),
        }
        try:
            response = requests.post(url, headers=headers, json=data, timeout=30)
            response.raise_for_status()
            result = response.json()
            content = result["choices"][0]["message"]["content"]
            score = self.extract_score(content)
            return score, content.strip()
        except Exception as e:
            return None, f"OpenRouter API error: {e}"


# Provider registry
PROVIDERS = {
    "Azure OpenAI GPT-4o": AzureOpenAIProvider,
    "OpenAI GPT-4o": OpenAIGPT4oProvider,
    "OpenAI GPT-4o-mini": OpenAIGPT4oMiniProvider,
    "Google Gemini": GeminiProvider,
    "Anthropic Claude": ClaudeProvider,
    "Groq": GroqProvider,
    "OpenRouter": OpenRouterProvider,
    "xAI Grok": GrokProvider,
}


def get_provider(provider_name: str, api_key: str) -> AIProvider:
    """Factory function to get provider instance."""
    aliases = {
        "Gemini": "Google Gemini",
        "Google Gemini (API)": "Google Gemini",
        "Grok": "xAI Grok",
        "xAI Grok (API)": "xAI Grok",
        "Claude 3 Opus": "Anthropic Claude",
        "Anthropic Claude (API)": "Anthropic Claude",
        "Groq (API)": "Groq",
        "OpenRouter (API)": "OpenRouter",
        "OpenAI GPT-4o (API)": "OpenAI GPT-4o",
        "Azure OpenAI GPT-4o (API)": "Azure OpenAI GPT-4o",
        "OpenAI GPT-4o-mini": "OpenAI GPT-4o-mini",
    }
    provider_name = aliases.get(provider_name, provider_name)
    if provider_name not in PROVIDERS:
        raise ValueError(f"Unknown provider: {provider_name}")
    return PROVIDERS[provider_name](api_key)
