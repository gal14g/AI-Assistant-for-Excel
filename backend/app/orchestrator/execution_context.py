"""
ExecutionContext stores intermediate results between tool executions.
It acts as the shared state bus for a single analytical pipeline run.
"""
from __future__ import annotations

import logging
from typing import Any
from datetime import datetime, timezone

from ..models.tool_output import ToolOutput

logger = logging.getLogger(__name__)


class ExecutionContext:
    """
    Shared mutable state for one orchestrated pipeline run.

    Tools write their ToolOutput here after execution; downstream tools
    can retrieve prior results by name.  The context also maintains an
    ordered execution log and the sequence of tool names that have run.
    """

    def __init__(self) -> None:
        # Ordered dict: tool_name → ToolOutput
        self._results: dict[str, ToolOutput] = {}
        # Ordered list of tool names in the sequence they were stored
        self.tool_sequence: list[str] = []
        # Human-readable log entries with timestamps
        self.execution_log: list[str] = []

    # ── Core storage ──────────────────────────────────────────────────────────

    def store(self, tool_name: str, output: ToolOutput) -> None:
        """
        Store *output* under *tool_name*.

        If a result for the same tool_name already exists it is overwritten
        (supports re-runs of individual steps).  The tool_sequence always
        reflects insertion order; overwritten tools keep their original
        position but a new entry is appended to the log.
        """
        if tool_name not in self._results:
            self.tool_sequence.append(tool_name)
        self._results[tool_name] = output
        status = "ok" if output.success else "FAILED"
        self.log(f"[{status}] {tool_name}")

    def get(self, tool_name: str) -> ToolOutput | None:
        """Return the stored ToolOutput for *tool_name*, or None."""
        return self._results.get(tool_name)

    def get_data(self, tool_name: str, key: str | None = None) -> Any:
        """
        Retrieve data from a stored result.

        Parameters
        ----------
        tool_name:
            The tool whose result you want.
        key:
            Optional nested key.  If *key* is provided and the result's
            ``data`` is a dict, returns ``data[key]``.  If the key does not
            exist, returns None.

        Returns
        -------
        Any
            The full ``data`` value (if *key* is None) or the nested value.
            Returns None if the tool has no stored result.
        """
        output = self._results.get(tool_name)
        if output is None:
            return None
        if key is None:
            return output.data
        if isinstance(output.data, dict):
            return output.data.get(key)
        return None

    def has(self, tool_name: str) -> bool:
        """Return True if a result for *tool_name* has been stored."""
        return tool_name in self._results

    def last_result(self) -> ToolOutput | None:
        """Return the most recently stored ToolOutput, or None."""
        if not self.tool_sequence:
            return None
        last_name = self.tool_sequence[-1]
        return self._results.get(last_name)

    # ── Aggregated accessors ──────────────────────────────────────────────────

    def all_warnings(self) -> list[str]:
        """Collect and return warnings from every stored result."""
        warnings: list[str] = []
        for name in self.tool_sequence:
            result = self._results.get(name)
            if result and result.warnings:
                for w in result.warnings:
                    warnings.append(f"[{name}] {w}")
        return warnings

    def all_errors(self) -> list[str]:
        """Collect and return errors from every stored result."""
        errors: list[str] = []
        for name in self.tool_sequence:
            result = self._results.get(name)
            if result and result.errors:
                for e in result.errors:
                    errors.append(f"[{name}] {e}")
        return errors

    # ── Logging ───────────────────────────────────────────────────────────────

    def log(self, message: str) -> None:
        """
        Append *message* to execution_log with an ISO-8601 UTC timestamp prefix.
        """
        ts = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + "Z"
        entry = f"{ts} {message}"
        self.execution_log.append(entry)
        logger.debug("ExecutionContext: %s", message)

    # ── Summary ───────────────────────────────────────────────────────────────

    def to_summary(self) -> dict:
        """
        Return a serialisable summary of the pipeline run.

        Keys
        ----
        tool_sequence : list[str]
            Ordered list of tool names executed.
        warnings : list[str]
            All warnings across all tools (prefixed with tool name).
        errors : list[str]
            All errors across all tools (prefixed with tool name).
        execution_log : list[str]
            Timestamped log entries.
        """
        return {
            "tool_sequence": list(self.tool_sequence),
            "warnings": self.all_warnings(),
            "errors": self.all_errors(),
            "execution_log": list(self.execution_log),
        }

    # ── Dunder ────────────────────────────────────────────────────────────────

    def __repr__(self) -> str:  # pragma: no cover
        return (
            f"ExecutionContext(tools={self.tool_sequence}, "
            f"warnings={len(self.all_warnings())}, "
            f"errors={len(self.all_errors())})"
        )
