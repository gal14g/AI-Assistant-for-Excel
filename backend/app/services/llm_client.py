"""
Unified LLM client using the OpenAI SDK.

Supports any OpenAI-compatible provider by configuring base_url:
  - OpenAI:    (default base_url — no prefix needed)
  - Anthropic: https://api.anthropic.com/v1/  (prefix: anthropic/)
  - Cohere:    https://api.cohere.ai/compatibility/v1  (prefix: cohere/)
  - Gemini:    https://generativelanguage.googleapis.com/v1beta/openai/  (prefix: gemini/)
  - Ollama:    http://localhost:11434/v1  (prefix: ollama/)
  - Azure:     LLM_BASE_URL=https://<resource>.openai.azure.com/

Provider is auto-detected from LLM_MODEL prefix or LLM_BASE_URL.
"""

from __future__ import annotations

import logging
from typing import Any, AsyncIterator

from openai import AsyncOpenAI, OpenAI

from ..config import settings

logger = logging.getLogger(__name__)

# ── Provider base URL mapping ────────────────────────────────────────────────
# If the user sets LLM_BASE_URL, it takes priority. Otherwise we auto-detect
# from the model prefix.

_PROVIDER_BASE_URLS: dict[str, str] = {
    "anthropic": "https://api.anthropic.com/v1/",
    "cohere": "https://api.cohere.ai/compatibility/v1",
    "gemini": "https://generativelanguage.googleapis.com/v1beta/openai/",
    "ollama": "http://localhost:11434/v1",
}

# Domains that need /v1 appended for OpenAI-compatible API
_OLLAMA_HOSTS = {"11434", "ollama.com", "ollama.ai"}


def _resolve_base_url() -> str | None:
    """Determine the base URL for the OpenAI client."""
    if settings.llm_base_url:
        url = settings.llm_base_url.rstrip("/")
        # Ollama (local or cloud): ensure /v1 suffix for OpenAI compatibility
        if any(host in url for host in _OLLAMA_HOSTS) and not url.endswith("/v1"):
            return url + "/v1"
        return url

    model_lower = settings.llm_model.lower()
    for prefix, url in _PROVIDER_BASE_URLS.items():
        if model_lower.startswith(prefix + "/"):
            return url

    # OpenAI models: no base_url needed (SDK default)
    return None


def _resolve_model_name() -> str:
    """
    Strip provider prefix from model name for providers that don't need it.
    e.g. "anthropic/claude-sonnet-4-20250514" -> "claude-sonnet-4-20250514"
         "gemini/gemini-2.0-flash"            -> "gemini-2.0-flash"
         "ollama/qwen2.5:14b"                 -> "qwen2.5:14b"
         "gpt-4o"                             -> "gpt-4o"
    """
    model = settings.llm_model
    for prefix in _PROVIDER_BASE_URLS:
        if model.lower().startswith(prefix + "/"):
            return model[len(prefix) + 1:]
    return model


def _resolve_api_key() -> str:
    """Return the API key, or a dummy for providers that don't need one."""
    if settings.llm_api_key:
        return settings.llm_api_key
    # Local Ollama doesn't require an API key but the SDK needs a non-empty string
    is_local_ollama = (
        settings.llm_model.lower().startswith("ollama/")
        and not settings.llm_base_url
    ) or (settings.llm_base_url and "11434" in settings.llm_base_url)
    if is_local_ollama:
        return "ollama"
    return settings.llm_api_key or "no-key-set"


def get_async_client() -> AsyncOpenAI:
    """Create an AsyncOpenAI client configured for the current provider."""
    base_url = _resolve_base_url()
    kwargs: dict[str, Any] = {
        "api_key": _resolve_api_key(),
        "timeout": 60.0,
    }
    if base_url:
        kwargs["base_url"] = base_url
    return AsyncOpenAI(**kwargs)


def get_sync_client() -> OpenAI:
    """Create a sync OpenAI client (for connectivity checks)."""
    base_url = _resolve_base_url()
    kwargs: dict[str, Any] = {
        "api_key": _resolve_api_key(),
        "timeout": 60.0,
    }
    if base_url:
        kwargs["base_url"] = base_url
    return OpenAI(**kwargs)


def get_model_name() -> str:
    """Return the model name to pass to the API."""
    return _resolve_model_name()


def build_completion_kwargs() -> dict[str, Any]:
    """
    Build common keyword arguments for chat completion calls.
    Returns model, max_tokens, temperature, and any provider-specific extras.
    """
    kwargs: dict[str, Any] = {
        "model": get_model_name(),
        "max_tokens": settings.llm_max_tokens,
        "temperature": settings.llm_temperature,
    }
    if settings.llm_json_mode:
        kwargs["response_format"] = {"type": "json_object"}
    # Qwen3 thinking mode produces thousands of reasoning tokens — disable it
    if "qwen3" in settings.llm_model.lower():
        kwargs["extra_body"] = {"think": False}
    return kwargs


async def acompletion(messages: list[dict], **overrides: Any) -> str:
    """
    Send a chat completion request and return the response text.

    This is the main function all services should call.
    Accepts the same kwargs as OpenAI's chat.completions.create().
    """
    client = get_async_client()
    kwargs = build_completion_kwargs()
    kwargs["messages"] = messages
    kwargs.update(overrides)

    response = await client.chat.completions.create(**kwargs)
    return response.choices[0].message.content or ""


async def acompletion_stream(messages: list[dict], **overrides: Any) -> AsyncIterator[str]:
    """
    Stream chat completion tokens. Yields text chunks as they arrive.
    response_format / json_mode is disabled — streaming and JSON mode are
    incompatible on most providers.
    """
    client = get_async_client()
    kwargs = build_completion_kwargs()
    # response_format (json_object) IS compatible with streaming on OpenAI.
    # Keeping it prevents the model from outputting plain text instead of JSON.
    kwargs["messages"] = messages
    kwargs["stream"] = True
    kwargs.update(overrides)

    stream = await client.chat.completions.create(**kwargs)
    async for chunk in stream:
        delta = chunk.choices[0].delta.content  # type: ignore[index]
        if delta:
            yield delta


def completion_sync(messages: list[dict], **overrides: Any) -> str:
    """Synchronous version for scripts and tests."""
    client = get_sync_client()
    kwargs = build_completion_kwargs()
    kwargs["messages"] = messages
    kwargs.update(overrides)

    response = client.chat.completions.create(**kwargs)
    return response.choices[0].message.content or ""
