"""
Persistence backend factory.

Reads `settings.database_url` and `settings.vector_store_url` to pick the
concrete implementation. Both factories cache their result for the process
lifetime — the first call decides which backend is active.

URL scheme matrix:

| Setting            | Empty / prefix        | Backend              |
|--------------------|-----------------------|----------------------|
| database_url       | "" or `sqlite://`     | SqliteRepositories   |
| database_url       | `postgresql://`       | PostgresRepositories |
| vector_store_url   | "" or `chroma://`     | ChromaVectorStore    |
| vector_store_url   | `pgvector://` / `postgresql://` | PgVectorStore |

Switching backends is a deployment-time concern: change the env var, restart.
Data portability is handled separately by `backend/scripts/migrate_db.py`.
"""

from __future__ import annotations

import logging
from typing import Optional

from app.config import settings
from app.persistence.base import Repositories, VectorStore

log = logging.getLogger(__name__)


_repositories: Optional[Repositories] = None
_vector_store: Optional[VectorStore] = None


def _resolve_repo_backend(url: str) -> str:
    """Return a normalised backend name for the given database URL."""
    u = (url or "").strip().lower()
    if not u or u.startswith("sqlite"):
        return "sqlite"
    if u.startswith(("postgres://", "postgresql://", "postgresql+asyncpg://")):
        return "postgres"
    # Unknown scheme → fall back to SQLite rather than crash.
    log.warning("Unknown database_url scheme %r — falling back to SQLite.", url)
    return "sqlite"


def _resolve_vector_backend(url: str) -> str:
    u = (url or "").strip().lower()
    if not u or u.startswith("chroma"):
        return "chroma"
    if u.startswith(("pgvector://", "postgres://", "postgresql://")):
        return "pgvector"
    log.warning("Unknown vector_store_url scheme %r — falling back to ChromaDB.", url)
    return "chroma"


def get_repositories() -> Repositories:
    """
    Return the process-wide `Repositories` bundle. Instantiation is lazy —
    the first call reads config and picks a backend; subsequent calls return
    the cached instance.

    Callers must `await repos.initialize()` once during app startup.
    """
    global _repositories  # noqa: PLW0603

    if _repositories is not None:
        return _repositories

    backend = _resolve_repo_backend(settings.database_url)
    if backend == "postgres":
        from app.persistence.postgres.repositories import PostgresRepositories

        _repositories = PostgresRepositories(settings.database_url)
    else:
        from app.persistence.sqlite.repositories import SqliteRepositories

        _repositories = SqliteRepositories(settings.database_url or settings.feedback_db_path)

    log.info("Persistence: using %s repositories.", backend)
    return _repositories


def get_vector_store() -> VectorStore:
    """Return the process-wide `VectorStore`. Same lazy-init semantics."""
    global _vector_store  # noqa: PLW0603

    if _vector_store is not None:
        return _vector_store

    backend = _resolve_vector_backend(settings.vector_store_url)
    if backend == "pgvector":
        from app.persistence.vector_pgvector import PgVectorStore

        # Reuse database_url if vector_store_url omits the connection info.
        url = settings.vector_store_url
        if url.startswith("pgvector://"):
            url = "postgresql://" + url[len("pgvector://") :]
        if not url or url == "postgresql://":
            url = settings.database_url
        _vector_store = PgVectorStore(url)
    else:
        from app.persistence.vector_chroma import ChromaVectorStore

        _vector_store = ChromaVectorStore(settings.vector_store_url or settings.chroma_persist_dir)

    log.info("Persistence: using %s vector store.", backend)
    return _vector_store


def reset_for_testing() -> None:
    """Clear cached factories — used by tests to switch backends mid-run."""
    global _repositories, _vector_store  # noqa: PLW0603
    _repositories = None
    _vector_store = None
