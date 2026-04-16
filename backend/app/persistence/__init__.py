"""
Swappable persistence layer.

Two orthogonal stores live here:

* **Repositories** (relational) — interactions, choices, few-shot examples,
  conversations, and conversation messages. `DATABASE_URL` picks the backend:

    empty / `sqlite://...`     → `SqliteRepositories` (default; backend/data/feedback.db)
    `postgresql://...`         → `PostgresRepositories` (asyncpg + connection pool)

* **VectorStore** (embeddings) — capability and few-shot example embeddings.
  `VECTOR_STORE_URL` picks the backend:

    empty / `chroma://...`     → `ChromaVectorStore` (default; backend/data/chroma/)
    `pgvector://...` / `postgresql://...` → `PgVectorStore` (pgvector extension)

Both backends pass the **same** parameterised test suite — see
`backend/tests/persistence/`. Switching backends is a deployment-time
decision (edit env vars / configmap); no code changes are needed.

The factory functions (`get_repositories`, `get_vector_store`) resolve the
selected backend on first access and cache the instance for the process
lifetime.
"""

from app.persistence.factory import get_repositories, get_vector_store

__all__ = ["get_repositories", "get_vector_store"]
