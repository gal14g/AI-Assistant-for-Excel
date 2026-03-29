"""
ToolOutput: unified result type returned by every analytical tool.

Every tool in the tools module returns a ToolOutput instance.  The orchestrator
stores these in the ExecutionContext and reads them for downstream tool inputs.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class ToolOutput:
    """Unified result wrapper for all analytical tool calls."""

    success: bool
    tool_name: str
    data: Any = None
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    metadata: dict = field(default_factory=dict)

    # ── Convenience helpers ───────────────────────────────────────────────────

    def get(self, key: str, default: Any = None) -> Any:
        """Retrieve a nested key from self.data if it is a dict."""
        if isinstance(self.data, dict):
            return self.data.get(key, default)
        return default

    @classmethod
    def ok(
        cls,
        tool_name: str,
        data: Any = None,
        warnings: list[str] | None = None,
        metadata: dict | None = None,
    ) -> "ToolOutput":
        """Factory for a successful result."""
        return cls(
            success=True,
            tool_name=tool_name,
            data=data,
            warnings=warnings or [],
            errors=[],
            metadata=metadata or {},
        )

    @classmethod
    def fail(
        cls,
        tool_name: str,
        errors: list[str],
        warnings: list[str] | None = None,
        metadata: dict | None = None,
    ) -> "ToolOutput":
        """Factory for a failed result."""
        return cls(
            success=False,
            tool_name=tool_name,
            data=None,
            warnings=warnings or [],
            errors=errors,
            metadata=metadata or {},
        )
