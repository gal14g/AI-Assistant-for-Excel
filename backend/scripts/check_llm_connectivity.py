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
  LLM_MODEL      — LiteLLM model string  (default: ollama/qwen3:27b)
  LLM_API_KEY    — API key if required   (default: empty)
  LLM_BASE_URL   — Base URL for local/proxy endpoints (default: empty)
"""

from __future__ import annotations

import os
import sys
import time


def check() -> None:
    try:
        import litellm
    except ImportError:
        print("❌  litellm is not installed. Run: pip install litellm", file=sys.stderr)
        sys.exit(1)

    model       = os.environ.get("LLM_MODEL",    "ollama/qwen3:27b")
    api_key     = os.environ.get("LLM_API_KEY",   "")
    base_url    = os.environ.get("LLM_BASE_URL",  "")

    print(f"🔌  Checking connectivity to LLM provider...")
    print(f"    Model:    {model}")
    print(f"    Base URL: {base_url or '(provider default)'}")
    print(f"    API key:  {'(set)' if api_key else '(not set)'}")
    print()

    kwargs: dict = {
        "model": model,
        "messages": [
            {"role": "user", "content": "Reply with exactly the word PONG and nothing else."}
        ],
        "max_tokens": 10,
        "temperature": 0,
    }
    if api_key:
        kwargs["api_key"] = api_key
    if base_url:
        kwargs["api_base"] = base_url

    start = time.monotonic()
    try:
        litellm.success_callback = []
        response = litellm.completion(**kwargs)
    except litellm.exceptions.AuthenticationError as exc:
        print(f"❌  Authentication failed — check LLM_API_KEY.\n    {exc}", file=sys.stderr)
        sys.exit(1)
    except litellm.exceptions.APIConnectionError as exc:
        print(f"❌  Cannot reach the LLM endpoint.\n    {exc}", file=sys.stderr)
        sys.exit(1)
    except Exception as exc:  # noqa: BLE001
        print(f"❌  Unexpected error: {type(exc).__name__}: {exc}", file=sys.stderr)
        sys.exit(1)

    elapsed = (time.monotonic() - start) * 1000
    content = (response.choices[0].message.content or "").strip()

    if not content:
        print("❌  LLM returned an empty response.", file=sys.stderr)
        sys.exit(1)

    # Accept any non-empty response — "PONG" is ideal but not enforced
    print(f"✅  LLM responded in {elapsed:.0f} ms")
    print(f"    Response: {content[:120]}")


if __name__ == "__main__":
    check()
