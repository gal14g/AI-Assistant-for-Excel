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

    persist_dir = Path(settings.chroma_persist_dir or str(_default_chroma_dir()))
    persist_dir.mkdir(parents=True, exist_ok=True)

    _client = PersistentClient(path=str(persist_dir))
    return _client


def get_embedding_fn():
    """Return (or create) the shared SentenceTransformerEmbeddingFunction."""
    global _embedding_fn  # noqa: PLW0603

    if _embedding_fn is not None:
        return _embedding_fn

    from chromadb.utils.embedding_functions import SentenceTransformerEmbeddingFunction

    _embedding_fn = SentenceTransformerEmbeddingFunction(
        model_name=settings.embedding_model,
    )
    return _embedding_fn
