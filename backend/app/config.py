"""
Application configuration loaded from environment variables.

All settings are read from the single .env file at the project root.
LLM provider is configured via LLM_MODEL + LLM_BASE_URL — uses the OpenAI SDK
with auto-detected base URLs for Gemini, Ollama, and other compatible providers.
"""

from pathlib import Path

from pydantic_settings import BaseSettings

# Resolve the project root .env regardless of the working directory.
# config.py lives at backend/app/config.py → parents[2] = project root.
_ENV_FILE = Path(__file__).resolve().parents[2] / ".env"


class Settings(BaseSettings):
    """Application settings with defaults for local development."""

    # ── LLM provider ──────────────────────────────────────────────────────────
    llm_model: str = "gpt-4o"
    llm_api_key: str = ""
    llm_base_url: str = ""
    llm_api_version: str = ""

    # Generation parameters
    llm_max_tokens: int = 4096
    llm_temperature: float = 0.1

    # Force JSON output mode (response_format: json_object).
    # Recommended for Qwen, Cohere, and smaller Ollama models.
    llm_json_mode: bool = False

    # ── Embedding / Capability Search ────────────────────────────────────────
    embedding_model: str = "all-MiniLM-L6-v2"
    chroma_persist_dir: str = ""          # auto-resolved to backend/data/chroma if empty
    capability_top_k: int = 25

    # ── Few-shot examples ────────────────────────────────────────────────────
    few_shot_top_k: int = 5               # how many dynamic examples to retrieve per query

    # ── Feedback DB ───────────────────────────────────────────────────────────
    feedback_db_path: str = ""            # auto-resolved to backend/data/feedback.db if empty

    # ── Server ────────────────────────────────────────────────────────────────
    host: str = "0.0.0.0"
    port: int = 8000
    cors_origins: list[str] = ["https://localhost:3000", "https://localhost:3001"]
    debug: bool = True

    # ── Deployment mode ───────────────────────────────────────────────────────
    openshift: bool = False
    serve_static: bool = False
    static_dir: str = "./static"

    model_config = {
        "env_file": str(_ENV_FILE),
        "env_file_encoding": "utf-8",
        "extra": "ignore",
    }


settings = Settings()
