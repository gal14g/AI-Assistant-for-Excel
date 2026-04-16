"""
Store-agnostic sentence-transformers embedding function.

The original embedding pipeline lived inside `services/chroma_client.py`
and only the Chroma vector store called into it. Now both `ChromaVectorStore`
and `PgVectorStore` embed documents identically — this module owns the
single SentenceTransformer model instance.

Model resolution matches the old `_resolve_embedding_model` logic: prefer a
bundled copy at `backend/models/<name>/` (crucial for air-gapped deploys),
otherwise fall through to whatever name is set so HF Hub can fetch it.
"""

from __future__ import annotations

import logging
from functools import lru_cache
from pathlib import Path

from app.config import settings

log = logging.getLogger(__name__)


def resolve_model_path() -> str:
    """
    Resolve `settings.embedding_model` to either a filesystem path (when the
    model is bundled under `backend/models/`) or a bare model name (fetched
    from HF Hub at load time).
    """
    name = settings.embedding_model
    if Path(name).exists():
        return str(Path(name).resolve())
    bundled = Path(__file__).resolve().parents[2] / "models" / name
    if bundled.exists():
        return str(bundled)
    return name


@lru_cache(maxsize=1)
def _get_model():
    """Load (and cache) the sentence-transformers model."""
    from sentence_transformers import SentenceTransformer

    path = resolve_model_path()
    log.info("Loading embedding model from %s", path)
    return SentenceTransformer(path)


def embed(texts: list[str]) -> list[list[float]]:
    """
    Encode a batch of strings into 384-dim vectors. Returns plain Python
    lists (not numpy arrays) so consumers don't need numpy installed.

    Empty input → empty output. Safe to call before the model is loaded;
    the first call pays the load cost and caches the instance.
    """
    if not texts:
        return []
    model = _get_model()
    vectors = model.encode(texts, convert_to_numpy=True, show_progress_bar=False)
    return [v.tolist() for v in vectors]


def get_chroma_embedding_function():
    """
    Return a ChromaDB-compatible embedding function. ChromaDB requires an
    object with a specific `__call__(input)` signature plus an
    `embedding_function_name` attribute in newer versions, so we defer to
    the upstream wrapper which handles this.
    """
    from chromadb.utils.embedding_functions import SentenceTransformerEmbeddingFunction

    return SentenceTransformerEmbeddingFunction(model_name=resolve_model_path())


def embedding_dimensions() -> int:
    """
    Dimensionality of the active embedding model. pgvector needs this at
    table-creation time; everything else can infer from the first vector.
    """
    return _get_model().get_sentence_embedding_dimension()
