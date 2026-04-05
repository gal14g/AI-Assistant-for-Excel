#!/usr/bin/env python3
"""
LLM Connectivity Check
======================
Sends a minimal completion request to the configured LLM endpoint and verifies
that a valid response is returned. Used in the GitLab CI/CD pipeline to catch
provider outages, wrong API keys, or misconfigured base URLs before deploying.

Exit codes:
  0  — LLM responded successfully
  1  — LLM unreachable, wrong API key, or invalid response

Usage (locally):
  python backend/scripts/check_llm_connectivity.py

Environment variables (same as the application):
  LLM_MODEL      — Model string (e.g. gpt-4o, gemini/gemini-2.0-flash, ollama/qwen3:27b)
  LLM_API_KEY    — API key if required   (default: empty)
  LLM_BASE_URL   — Base URL for local/proxy endpoints (default: empty)
"""

from __future__ import annotations

import os
import sys
import time


def check() -> None:
    try:
        from openai import OpenAI, AuthenticationError, APIConnectionError
    except ImportError:
        print("openai is not installed. Run: pip install openai", file=sys.stderr)
        sys.exit(1)

    model = os.environ.get("LLM_MODEL", "gpt-4o")
    api_key = os.environ.get("LLM_API_KEY", "")
    base_url = os.environ.get("LLM_BASE_URL", "")

    print("Checking connectivity to LLM provider...")
    print(f"    Model:    {model}")
    print(f"    Base URL: {base_url or '(provider default)'}")
    print(f"    API key:  {'(set)' if api_key else '(not set)'}")
    print()

    # Auto-detect provider base URL from model prefix
    resolved_base_url = base_url
    resolved_model = model
    if not resolved_base_url:
        if model.lower().startswith("gemini/"):
            resolved_base_url = "https://generativelanguage.googleapis.com/v1beta/openai/"
            resolved_model = model[7:]  # strip "gemini/"
        elif model.lower().startswith("ollama/"):
            resolved_base_url = "http://localhost:11434/v1"
            resolved_model = model[7:]  # strip "ollama/"

    # Ensure Ollama gets /v1 suffix
    if resolved_base_url and "11434" in resolved_base_url and not resolved_base_url.endswith("/v1"):
        resolved_base_url = resolved_base_url.rstrip("/") + "/v1"

    client_kwargs: dict = {
        "api_key": api_key or "ollama",
        "timeout": 30.0,
    }
    if resolved_base_url:
        client_kwargs["base_url"] = resolved_base_url

    client = OpenAI(**client_kwargs)

    start = time.monotonic()
    try:
        response = client.chat.completions.create(
            model=resolved_model,
            messages=[
                {"role": "user", "content": "Reply with exactly the word PONG and nothing else."}
            ],
            max_tokens=10,
            temperature=0,
        )
    except AuthenticationError as exc:
        print(f"Authentication failed — check LLM_API_KEY.\n    {exc}", file=sys.stderr)
        sys.exit(1)
    except APIConnectionError as exc:
        print(f"Cannot reach the LLM endpoint.\n    {exc}", file=sys.stderr)
        sys.exit(1)
    except Exception as exc:
        print(f"Unexpected error: {type(exc).__name__}: {exc}", file=sys.stderr)
        sys.exit(1)

    elapsed = (time.monotonic() - start) * 1000
    content = (response.choices[0].message.content or "").strip()

    if not content:
        print("LLM returned an empty response.", file=sys.stderr)
        sys.exit(1)

    print(f"LLM responded in {elapsed:.0f} ms")
    print(f"    Response: {content[:120]}")


if __name__ == "__main__":
    check()
