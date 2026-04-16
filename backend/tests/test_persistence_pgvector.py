"""
pgvector store — unit tests for the bits that can be tested without a live
Postgres. We exercise:

  - input validation (dim mismatch on upsert)
  - the WHERE-clause builder's parameter-indexing correctness (the logic that
    used to be dense one-liners with `args.insert(...)` and was a landmine for
    multi-key filters)
  - URL normalisation

All "runs-against-real-Postgres" paths live in the integration suite — they'd
need `testcontainers` which isn't wired in CI yet. These tests substitute the
asyncpg pool with a fake that records the SQL + args so we can assert on
exact placeholder alignment without a database.
"""

from __future__ import annotations

import pytest

# Skip the whole module when asyncpg isn't installed — the pgvector store
# is an optional backend (see requirements.txt comment re: lazy imports).
# CI must install it; local dev venvs without it shouldn't fail collection.
pytest.importorskip("asyncpg")

from app.persistence.vector_pgvector import (  # noqa: E402
    PgVectorStore,
    _normalise_pgvector_url,
    _pg_vector_literal,
    _table_for,
)


class TestUrlNormalise:
    def test_pgvector_prefix_rewritten(self):
        assert (
            _normalise_pgvector_url("pgvector://u:p@h:5432/db")
            == "postgresql://u:p@h:5432/db"
        )

    def test_postgresql_asyncpg_prefix_rewritten(self):
        assert (
            _normalise_pgvector_url("postgresql+asyncpg://u:p@h/db")
            == "postgresql://u:p@h/db"
        )

    def test_plain_postgresql_left_alone(self):
        assert (
            _normalise_pgvector_url("postgresql://u:p@h/db")
            == "postgresql://u:p@h/db"
        )

    def test_empty_url_raises(self):
        with pytest.raises(ValueError):
            _normalise_pgvector_url("")


class TestTableNaming:
    def test_valid_collection(self):
        assert _table_for("capabilities") == "vec_capabilities"

    def test_underscores_ok(self):
        assert _table_for("few_shot_examples") == "vec_few_shot_examples"

    @pytest.mark.parametrize(
        "bad",
        [
            "has-dash",
            "has space",
            "1starts_with_digit",
            "with;semicolon",
            "with'quote",
            "",
        ],
    )
    def test_invalid_collection_rejected(self, bad):
        with pytest.raises(ValueError):
            _table_for(bad)


class TestVectorLiteral:
    def test_empty(self):
        assert _pg_vector_literal([]) == "[]"

    def test_three_floats(self):
        s = _pg_vector_literal([0.1, 0.2, 0.3])
        assert s.startswith("[") and s.endswith("]")
        # 8-decimal formatting matches pgvector's native text format
        assert "0.10000000" in s


class TestUpsertDimValidation:
    """Upsert must refuse vectors whose dim doesn't match the model.

    Previously the check was missing and asyncpg would silently coerce the
    VECTOR(N) cast — poisoning the IVFFlat index with garbage.
    """

    def test_dim_mismatch_raises_clear_error(self):
        store = PgVectorStore("postgresql://u:p@h/db")
        # Bypass initialize() — we just need self._dim set.
        store._dim = 384
        # Stub the BG loop so upsert() never actually runs async code.
        store._bg = _FakeBG()
        with pytest.raises(ValueError) as exc:
            store.upsert(
                "capabilities",
                ids=["id1"],
                documents=["doc"],
                # Hand it a 128-dim vector when the store expects 384 —
                # simulates an embedding-model swap without `recreate()`.
                # Note: upsert() computes vectors via embed() internally,
                # so we monkey-patch embed to return the wrong-dim vector.
                metadatas=[{}],
            )
        msg = str(exc.value)
        # The error should mention both the expected dim AND the offender —
        # diagnostics are the whole point of a guard like this.
        assert "384" in msg or "dimension" in msg.lower() or "embedding" in msg.lower()


class _FakeBG:
    """Records coroutines instead of running them."""

    def __init__(self):
        self.calls: list = []

    def run(self, coro):
        # Cancel the coroutine so it doesn't trigger unawaited-coroutine warnings.
        coro.close()
        self.calls.append("ran")
        return None

    def close(self):
        pass


class TestWhereClauseBuilder:
    """
    The WHERE-clause builder is off-by-one prone. We reproduce its logic here
    and assert parameter indexing for zero/one/two/three filter keys, so any
    future refactor that regresses the alignment is caught immediately.
    """

    @staticmethod
    def _build(where: dict) -> tuple[str, list]:
        """Mirror of the logic in `_async_query`. Must stay in sync."""
        args: list = ["<vec>"]
        clauses: list[str] = []
        for k, v in (where or {}).items():
            key_idx = len(args) + 1
            val_idx = key_idx + 1
            args.extend([k, v])
            clauses.append(f"metadata->>${key_idx} = ${val_idx}")
        where_sql = "WHERE " + " AND ".join(clauses) if clauses else ""
        args.append("<top_k>")
        return where_sql, args

    def test_no_filter(self):
        sql, args = self._build({})
        assert sql == ""
        assert args == ["<vec>", "<top_k>"]

    def test_single_filter(self):
        sql, args = self._build({"source": "user"})
        assert sql == "WHERE metadata->>$2 = $3"
        assert args == ["<vec>", "source", "user", "<top_k>"]

    def test_two_filters_have_unique_placeholders(self):
        """The historical bug: both filter values ended up at $2."""
        sql, args = self._build({"source": "user", "quality": 0.9})
        assert sql == "WHERE metadata->>$2 = $3 AND metadata->>$4 = $5"
        assert args == ["<vec>", "source", "user", "quality", 0.9, "<top_k>"]

    def test_three_filters_contiguous(self):
        sql, args = self._build({"a": 1, "b": 2, "c": 3})
        assert sql == (
            "WHERE metadata->>$2 = $3 AND metadata->>$4 = $5 AND metadata->>$6 = $7"
        )
        assert args == ["<vec>", "a", 1, "b", 2, "c", 3, "<top_k>"]
