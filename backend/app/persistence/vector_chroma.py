"""
ChromaDB-backed vector store.

Thin adapter around `chromadb.PersistentClient` that implements the
`VectorStore` protocol defined in `persistence/base.py`. Collection-first
API: every method takes a collection name, so both `capabilities` and
`few_shot_examples` live in the same persistent directory.

Embedding: the shared SentenceTransformer-based function from
`persistence/embedding.py` (bundled model at `backend/models/<name>/`).
ChromaDB needs its own wrapper object to call the model at insert/query
time, which is provided by `embedding.get_chroma_embedding_function()`.

Persist dir resolution mirrors the pre-refactor `chroma_client.py`:
- `""` or `chroma://` → `backend/data/chroma/`
- `chroma:///abs/path` / `chroma://./rel/path` → that path
- bare path → that path
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Optional

from app.persistence.base import VectorStore

log = logging.getLogger(__name__)


def _default_chroma_dir() -> Path:
    # backend/app/persistence/vector_chroma.py → parents[2] = backend/
    return Path(__file__).resolve().parents[2] / "data" / "chroma"


def _normalise_chroma_url(url: str) -> Path:
    if not url:
        return _default_chroma_dir()
    if url.startswith("chroma:///"):
        return Path(url[len("chroma:///") :])
    if url.startswith("chroma://"):
        rest = url[len("chroma://") :]
        return _default_chroma_dir() if not rest else Path(rest)
    return Path(url)


class ChromaVectorStore(VectorStore):
    """Default vector store backend — persists to `backend/data/chroma/`."""

    def __init__(self, url: str = "") -> None:
        self._persist_dir = _normalise_chroma_url(url)
        self._client = None  # chromadb.PersistentClient
        self._embedding_fn = None
        self._collections: dict[str, Any] = {}

    # ── Lifecycle ────────────────────────────────────────────────────────────

    def initialize(self) -> None:
        """Create the persistent client + cached embedding function."""
        from chromadb import PersistentClient
        from chromadb.config import Settings as ChromaSettings

        from app.persistence.embedding import get_chroma_embedding_function

        self._persist_dir.mkdir(parents=True, exist_ok=True)
        # anonymized_telemetry=False → no posthog.com calls (air-gapped safe)
        self._client = PersistentClient(
            path=str(self._persist_dir),
            settings=ChromaSettings(anonymized_telemetry=False),
        )
        self._embedding_fn = get_chroma_embedding_function()
        log.info("ChromaDB vector store ready at %s", self._persist_dir)

    def _coll(self, name: str):
        """Lazy get-or-create so consumers don't have to pre-declare collections."""
        if self._client is None:
            raise RuntimeError("ChromaVectorStore not initialised")
        if name not in self._collections:
            self._collections[name] = self._client.get_or_create_collection(
                name=name,
                embedding_function=self._embedding_fn,
            )
        return self._collections[name]

    # ── Contract ─────────────────────────────────────────────────────────────

    def upsert(
        self,
        collection: str,
        ids: list[str],
        documents: list[str],
        metadatas: list[dict[str, Any]],
    ) -> None:
        if not ids:
            return
        coll = self._coll(collection)
        # Chroma doesn't have a true upsert in older versions; fall back to
        # delete-then-add for existing ids. Newer clients expose .upsert().
        upsert = getattr(coll, "upsert", None)
        if callable(upsert):
            upsert(ids=ids, documents=documents, metadatas=metadatas)
        else:
            try:
                coll.delete(ids=ids)
            except Exception:  # pragma: no cover — "not found" is fine
                pass
            coll.add(ids=ids, documents=documents, metadatas=metadatas)

    def query(
        self,
        collection: str,
        text: str,
        top_k: int,
        where: Optional[dict[str, Any]] = None,
    ) -> list[dict[str, Any]]:
        coll = self._coll(collection)
        n = min(top_k, coll.count() or 1)
        if n <= 0:
            return []
        kwargs: dict[str, Any] = {"query_texts": [text], "n_results": n}
        if where:
            kwargs["where"] = where
        raw = coll.query(**kwargs)

        out: list[dict[str, Any]] = []
        ids = (raw.get("ids") or [[]])[0]
        docs = (raw.get("documents") or [[]])[0]
        metas = (raw.get("metadatas") or [[]])[0]
        dists = (raw.get("distances") or [[]])[0]
        for i, doc_id in enumerate(ids):
            out.append(
                {
                    "id": doc_id,
                    "document": docs[i] if i < len(docs) else "",
                    "metadata": metas[i] if i < len(metas) else {},
                    "distance": dists[i] if i < len(dists) else None,
                }
            )
        return out

    def get_by_ids(self, collection: str, ids: list[str]) -> list[dict[str, Any]]:
        if not ids:
            return []
        coll = self._coll(collection)
        raw = coll.get(ids=ids)
        out: list[dict[str, Any]] = []
        got_ids = raw.get("ids") or []
        got_docs = raw.get("documents") or []
        got_metas = raw.get("metadatas") or []
        id_to_idx = {i: idx for idx, i in enumerate(got_ids)}
        # Preserve input order for relevance ranking.
        for doc_id in ids:
            idx = id_to_idx.get(doc_id)
            if idx is None:
                continue
            out.append(
                {
                    "id": doc_id,
                    "document": got_docs[idx] if idx < len(got_docs) else "",
                    "metadata": got_metas[idx] if idx < len(got_metas) else {},
                    "distance": None,
                }
            )
        return out

    def delete(self, collection: str, ids: Optional[list[str]] = None) -> None:
        if self._client is None:
            return
        if ids is None:
            # Whole-collection delete → drop it and forget the cache entry.
            try:
                self._client.delete_collection(collection)
            except Exception:  # pragma: no cover
                pass
            self._collections.pop(collection, None)
            return
        if not ids:
            return
        self._coll(collection).delete(ids=ids)

    def count(self, collection: str) -> int:
        return self._coll(collection).count()

    def recreate(self, collection: str) -> None:
        """Drop and recreate — used when seed data has changed between releases."""
        self.delete(collection, ids=None)
        self._coll(collection)  # force recreate
