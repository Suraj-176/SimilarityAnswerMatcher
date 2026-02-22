"""
AI Provider abstraction layer for prompt-based similarity matching.
Supports: Azure OpenAI, OpenAI, Google Gemini, Anthropic Claude, Groq,
OpenRouter, and xAI Grok.
"""

import logging
import re
import time
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

    @staticmethod
    def _sanitize_error_text(text: str) -> str:
        """Redact API keys/tokens from error strings before surfacing to UI/logs."""
        sanitized = str(text or "")
        sanitized = re.sub(r"([?&]key=)[^&\s]+", r"\1[REDACTED]", sanitized, flags=re.IGNORECASE)
        sanitized = re.sub(r"(Bearer\s+)[A-Za-z0-9._\-]+", r"\1[REDACTED]", sanitized, flags=re.IGNORECASE)
        sanitized = re.sub(r"\bsk-[A-Za-z0-9_\-]+\b", "[REDACTED]", sanitized)
        return sanitized

    @classmethod
    def _extract_error_detail(cls, response: Optional[requests.Response]) -> str:
        if response is None:
            return ""
        try:
            payload = response.json()
            if isinstance(payload, dict):
                err = payload.get("error")
                if isinstance(err, dict):
                    code = err.get("code")
                    msg = err.get("message") or err.get("msg") or err.get("detail") or err.get("type")
                    if code and msg:
                        return cls._sanitize_error_text(f"{code}: {msg}")
                    if msg:
                        return cls._sanitize_error_text(str(msg))
                    return cls._sanitize_error_text(str(err))
                if isinstance(err, str):
                    return cls._sanitize_error_text(err)
                for key in ("message", "detail"):
                    if key in payload and payload[key]:
                        return cls._sanitize_error_text(str(payload[key]))
                return cls._sanitize_error_text(str(payload))
            return cls._sanitize_error_text(str(payload))
        except Exception:
            body = getattr(response, "text", "") or ""
            return cls._sanitize_error_text(body[:500])

    @classmethod
    def _format_requests_error(cls, provider_name: str, exc: requests.exceptions.RequestException) -> str:
        if isinstance(exc, requests.exceptions.HTTPError):
            response = exc.response
            status = response.status_code if response is not None else "unknown"
            reason = response.reason if response is not None else ""
            detail = cls._extract_error_detail(response)

            hint = ""
            if status == 400:
                hint = "Bad request. Check model name and provider configuration."
            elif status == 401:
                hint = "Authentication failed. Check API key."
            elif status == 403:
                hint = "Access denied. Check model access/permissions."
            elif status == 404:
                hint = "Model/endpoint not found."
            elif status == 429:
                hint = "Rate limit or quota exceeded."

            parts = [f"{provider_name} API error ({status}{(' ' + reason) if reason else ''})."]
            if hint:
                parts.append(hint)
            if detail:
                parts.append(f"Details: {detail}")
            return " ".join(parts)

        return f"{provider_name} API error: {cls._sanitize_error_text(str(exc))}"


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
        except requests.exceptions.RequestException as e:
            return None, self._format_requests_error("Azure OpenAI", e)
        except Exception as e:
            return None, f"Azure OpenAI unexpected error: {self._sanitize_error_text(str(e))}"


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
        except requests.exceptions.RequestException as e:
            return None, self._format_requests_error("OpenAI", e)
        except Exception as e:
            return None, f"OpenAI unexpected error: {self._sanitize_error_text(str(e))}"


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

        max_retries = 3
        for attempt in range(max_retries):
            try:
                response = requests.post(url, headers=headers, json=data, timeout=30)
                response.raise_for_status()
                result = response.json()

                if "choices" not in result or not result["choices"]:
                    error_msg = f"Groq response missing choices. Details: {self._sanitize_error_text(str(result))}"
                    logger.error(error_msg)
                    return None, error_msg

                content = result["choices"][0]["message"]["content"]
                logger.debug("Groq raw response: %s", content)
                score = self.extract_score(content)
                return score, content.strip()
            except requests.exceptions.HTTPError as e:
                status = e.response.status_code if e.response is not None else None
                if status == 429 and attempt < max_retries - 1:
                    wait_seconds = 2 ** attempt
                    logger.warning(
                        "Groq rate-limited (429). Retrying in %ss (attempt %s/%s).",
                        wait_seconds,
                        attempt + 1,
                        max_retries,
                    )
                    time.sleep(wait_seconds)
                    continue
                error_msg = self._format_requests_error("Groq", e)
                logger.error(error_msg)
                return None, error_msg
            except requests.exceptions.RequestException as e:
                error_msg = self._format_requests_error("Groq", e)
                logger.error(error_msg)
                return None, error_msg
            except (KeyError, IndexError) as e:
                error_msg = f"Groq response parsing error: {self._sanitize_error_text(str(e))}"
                logger.error(error_msg)
                return None, error_msg
            except Exception as e:
                error_msg = f"Groq unexpected error: {self._sanitize_error_text(str(e))}"
                logger.error(error_msg)
                return None, error_msg

        return None, "Groq API error: retries exhausted."


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
                err_obj = result.get("error", {})
                err_code = err_obj.get("code") or err_obj.get("status")
                err_msg = err_obj.get("message", "Unknown error")
                if err_code:
                    error_msg = f"Google Gemini API error ({err_code}). Details: {self._sanitize_error_text(err_msg)}"
                else:
                    error_msg = f"Google Gemini API error. Details: {self._sanitize_error_text(err_msg)}"
                logger.error(error_msg)
                return None, error_msg

            if "candidates" not in result or not result["candidates"]:
                error_msg = f"Google Gemini response missing candidates. Details: {self._sanitize_error_text(str(result))}"
                logger.error(error_msg)
                return None, error_msg

            candidate = result["candidates"][0]
            if "content" not in candidate or "parts" not in candidate["content"] or not candidate["content"]["parts"]:
                error_msg = f"Google Gemini response parsing error. Details: {self._sanitize_error_text(str(candidate))}"
                logger.error(error_msg)
                return None, error_msg

            content = candidate["content"]["parts"][0]["text"]
            logger.debug("Gemini parsed content: %s", content)
            score = self.extract_score(content)
            return score, content.strip()
        except requests.exceptions.RequestException as e:
            error_msg = self._format_requests_error("Google Gemini", e)
            logger.error(error_msg)
            return None, error_msg
        except (KeyError, IndexError) as e:
            error_msg = f"Google Gemini response parsing error: {self._sanitize_error_text(str(e))}"
            logger.error(error_msg)
            return None, error_msg
        except Exception as e:
            error_msg = f"Google Gemini unexpected error: {self._sanitize_error_text(str(e))}"
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
        except requests.exceptions.RequestException as e:
            return None, self._format_requests_error("xAI Grok", e)
        except Exception as e:
            return None, f"xAI Grok unexpected error: {self._sanitize_error_text(str(e))}"


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
        except requests.exceptions.RequestException as e:
            return None, self._format_requests_error("Anthropic Claude", e)
        except Exception as e:
            return None, f"Anthropic Claude unexpected error: {self._sanitize_error_text(str(e))}"


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
        except requests.exceptions.RequestException as e:
            return None, self._format_requests_error("OpenRouter", e)
        except Exception as e:
            return None, f"OpenRouter unexpected error: {self._sanitize_error_text(str(e))}"


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
