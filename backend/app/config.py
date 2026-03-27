"""
Application configuration loaded from environment variables.

LLM provider is fully generic via LiteLLM.  Set LLM_MODEL to any model string
that LiteLLM understands:

  Provider        | Example LLM_MODEL value
  --------------- | --------------------------------------------------
  OpenAI          | gpt-4o  /  gpt-4-turbo  /  gpt-3.5-turbo
  Anthropic       | claude-sonnet-4-20250514  /  claude-opus-4-20250514
  Google Gemini   | gemini/gemini-1.5-pro
  Ollama (local)  | ollama/llama3  /  ollama/mistral
  Azure OpenAI    | azure/<your-deployment-name>
  AWS Bedrock     | bedrock/anthropic.claude-3-sonnet-20240229-v1:0
  LiteLLM proxy   | openai/<model>  (set LLM_BASE_URL to proxy URL)
  Any OpenAI-compat endpoint  | openai/<model>  + LLM_BASE_URL

For providers that need an API key, set LLM_API_KEY.
For local/self-hosted models (Ollama, custom endpoints), leave it empty or omit it.
For Azure, also set LLM_API_BASE, LLM_API_VERSION.
"""

from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    """Application settings with defaults for local development."""

    # ── LLM provider ──────────────────────────────────────────────────────────
    # LiteLLM model string (see table above).  Defaults to Claude Sonnet.
    llm_model: str = "claude-sonnet-4-20250514"

    # API key for the chosen provider.  Leave empty for local models (Ollama).
    llm_api_key: str = ""

    # Optional: override the API base URL.
    #   Ollama local:  http://localhost:11434
    #   LiteLLM proxy: http://my-proxy:4000
    #   Azure:         https://<resource>.openai.azure.com/
    llm_base_url: str = ""

    # Optional: API version (Azure only)
    llm_api_version: str = ""

    # Generation parameters
    llm_max_tokens: int = 4096
    llm_temperature: float = 0.1  # Low temperature for deterministic plans

    # ── Server ────────────────────────────────────────────────────────────────
    host: str = "0.0.0.0"
    port: int = 8000
    cors_origins: list[str] = ["https://localhost:3000", "https://localhost:3001"]
    debug: bool = True

    model_config = {"env_file": ".env", "env_file_encoding": "utf-8", "extra": "ignore"}


settings = Settings()
