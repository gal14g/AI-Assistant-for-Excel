"""
Shared ChromaDB client and embedding function.

Both capability_store and example_store use the same ChromaDB persistence
directory and embedding model. This module initialises them once and provides
a shared accessor so the embedding model is only loaded into memory once.
"""

from __future__ import annotations

from pathlib import Path

from ..config import settings

_client = None
_embedding_fn = None


def _default_chroma_dir() -> Path:
    """backend/data/chroma/ — next to the app/ package."""
    return Path(__file__).resolve().parents[2] / "data" / "chroma"


def get_chroma_client():
    """Return (or create) the shared PersistentClient."""
    global _client  # noqa: PLW0603

    if _client is not None:
        return _client

    from chromadb import PersistentClient
    from chromadb.config import Settings as ChromaSettings

    persist_dir = Path(settings.chroma_persist_dir or str(_default_chroma_dir()))
    persist_dir.mkdir(parents=True, exist_ok=True)

    # Disable anonymous telemetry (posthog.com) — required for air-gapped deployments
    # and cleaner logs everywhere else. Env var ANONYMIZED_TELEMETRY=False also works.
    _client = PersistentClient(
        path=str(persist_dir),
        settings=ChromaSettings(anonymized_telemetry=False),
    )
    return _client


def _resolve_embedding_model() -> str:
    """
    Resolve the embedding model name/path.

    If a bundled local model exists at backend/models/<name>/, use its absolute
    path (no network needed — ideal for air-gapped deployments and CI).
    Otherwise return the bare name so sentence-transformers fetches it from HF.
    """
    name = settings.embedding_model
    # Already an absolute/relative path that exists
    if Path(name).exists():
        return str(Path(name).resolve())
    # Check bundled location: backend/models/<name>/
    bundled = Path(__file__).resolve().parents[2] / "models" / name
    if bundled.exists():
        return str(bundled)
    return name


def get_embedding_fn():
    """Return (or create) the shared SentenceTransformerEmbeddingFunction."""
    global _embedding_fn  # noqa: PLW0603

    if _embedding_fn is not None:
        return _embedding_fn

    from chromadb.utils.embedding_functions import SentenceTransformerEmbeddingFunction

    _embedding_fn = SentenceTransformerEmbeddingFunction(
        model_name=_resolve_embedding_model(),
    )
    return _embedding_fn
