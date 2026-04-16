"""
Abstract interfaces for the swappable persistence layer.

Four repository contracts + one vector store contract. Concrete
implementations (SQLite, Postgres, ChromaDB, pgvector) must preserve
behavioural parity — the same parametrised test suite runs against every
implementation.

Method shapes mirror the original `backend/app/db.py` and
`services/chroma_client.py` call sites verbatim so existing routers and
services don't need to change.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import TYPE_CHECKING, Any, Optional

if TYPE_CHECKING:
    from app.models.chat import ChatResponse, RangeTokenRef as RangeToken


# ── Repositories ────────────────────────────────────────────────────────────


class InteractionRepository(ABC):
    """Chat interaction logging + user choice tracking."""

    @abstractmethod
    async def log_interaction(
        self,
        *,
        interaction_id: str,
        user_message: str,
        active_sheet: Optional[str],
        workbook_name: Optional[str],
        range_tokens: Optional[list["RangeToken"]],
        response: "ChatResponse",
        model_used: str,
        latency_ms: int,
    ) -> None:
        """Record a chat interaction (request + response) to the DB."""

    @abstractmethod
    async def log_choice(
        self,
        *,
        interaction_id: str,
        chosen_plan_id: Optional[str],
        action: str,
    ) -> None:
        """Record the user's accept/dismiss/custom choice for an interaction."""

    @abstractmethod
    async def get_interaction(self, interaction_id: str) -> Optional[dict]:
        """Fetch a single interaction's id, user_message, plans_json."""


class ConversationRepository(ABC):
    """Multi-turn conversation storage."""

    @abstractmethod
    async def create(self, title: str) -> str:
        """Create a new conversation and return its id."""

    @abstractmethod
    async def touch(self, conversation_id: str) -> None:
        """Update the conversation's `updated_at` timestamp."""

    @abstractmethod
    async def rename(self, conversation_id: str, title: str) -> bool:
        """Rename a conversation. Returns True if the row existed."""

    @abstractmethod
    async def delete(self, conversation_id: str) -> bool:
        """Delete a conversation and its messages. Returns True on success."""

    @abstractmethod
    async def list(self, limit: int = 100) -> list[dict]:
        """List conversations, newest-first, with message counts."""

    @abstractmethod
    async def load(self, conversation_id: str) -> Optional[dict]:
        """Fetch a conversation + all its messages, or None if not found."""

    @abstractmethod
    async def append_message(
        self,
        *,
        conversation_id: str,
        message_id: str,
        role: str,
        content: str,
        range_tokens: Optional[object] = None,
        plan: Optional[object] = None,
    ) -> None:
        """Append a user/assistant/system message to a conversation."""

    @abstractmethod
    async def update_message_execution(
        self,
        *,
        conversation_id: str,
        message_id: str,
        execution: Optional[object],
        progress_log: Optional[object],
    ) -> bool:
        """Attach execution state + progress log to an existing message."""

    @abstractmethod
    async def pop_last_exchange(self, conversation_id: str) -> int:
        """Remove the last 2 messages (user + assistant). Returns rows deleted."""


class FewShotRepository(ABC):
    """Few-shot example storage (user message + assistant JSON response)."""

    @abstractmethod
    async def insert(
        self,
        *,
        example_id: str,
        user_message: str,
        assistant_response: str,
        source: str = "seed",
        interaction_id: Optional[str] = None,
    ) -> None:
        """Insert an example. Idempotent — silently skips on ID collision."""

    @abstractmethod
    async def get_by_ids(self, ids: list[str]) -> list[dict]:
        """Fetch examples by ID, preserving input order for relevance ranking."""


class Repositories(ABC):
    """
    Bundle of all three repositories, plus lifecycle management. Concrete
    implementations expose the three repos as attributes:

        repos.interactions : InteractionRepository
        repos.conversations : ConversationRepository
        repos.few_shot : FewShotRepository
    """

    interactions: InteractionRepository
    conversations: ConversationRepository
    few_shot: FewShotRepository

    @abstractmethod
    async def initialize(self) -> None:
        """Open the connection (pool) and create tables if missing."""

    @abstractmethod
    async def close(self) -> None:
        """Release the connection (pool)."""


# ── Vector store ────────────────────────────────────────────────────────────


class VectorStore(ABC):
    """
    Collection-oriented vector store abstraction. Both `ChromaVectorStore`
    and `PgVectorStore` implement this.

    Collections used by Excel Copilot:
      - `capabilities` — action descriptions + example queries (read-heavy)
      - `few_shot_examples` — user message embeddings for few-shot retrieval
    """

    @abstractmethod
    def initialize(self) -> None:
        """Prepare the underlying store (create tables/collections)."""

    @abstractmethod
    def upsert(
        self,
        collection: str,
        ids: list[str],
        documents: list[str],
        metadatas: list[dict[str, Any]],
    ) -> None:
        """Insert or replace documents. `documents` are embedded on the way in."""

    @abstractmethod
    def query(
        self,
        collection: str,
        text: str,
        top_k: int,
        where: Optional[dict[str, Any]] = None,
    ) -> list[dict[str, Any]]:
        """
        Return up to `top_k` most-similar documents.

        Each result dict contains:
          - "id": the document ID
          - "document": the stored text
          - "metadata": the stored metadata dict
          - "distance": cosine distance (lower = more similar)
        """

    @abstractmethod
    def get_by_ids(self, collection: str, ids: list[str]) -> list[dict[str, Any]]:
        """Fetch specific documents by ID."""

    @abstractmethod
    def delete(self, collection: str, ids: Optional[list[str]] = None) -> None:
        """Delete specific IDs, or the entire collection if `ids` is None."""

    @abstractmethod
    def count(self, collection: str) -> int:
        """Number of documents in a collection."""

    @abstractmethod
    def recreate(self, collection: str) -> None:
        """Drop and recreate a collection (for a full re-index)."""
