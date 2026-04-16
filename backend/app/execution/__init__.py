"""
Python-side plan execution — parallel to the frontend `engine/` module.

Used by the MCP server (`backend/mcp_server.py`) so desktop chat clients can
drive Excel through xlwings without going through Office.js. All 76
capabilities that the Office.js executor supports are mirrored here as
Python handlers operating on `xlwings.Book` / `xlwings.Range`.

Top-level exports:
- `PlanExecutor`        abstract base contract
- `XlwingsExecutor`     concrete desktop-Excel implementation
- `ExecutorContext`     per-plan state (workbook handle, bindings, snapshots)

The module is *lazy-import safe* for headless servers: nothing at import
time touches xlwings, pywin32, or appscript. Actual COM/AppleScript usage
only happens when `XlwingsExecutor` is instantiated.
"""

from app.execution.base import (
    PlanExecutor,
    ExecutorContext,
    StepResult,
    ProgressCallback,
)

__all__ = [
    "PlanExecutor",
    "ExecutorContext",
    "StepResult",
    "ProgressCallback",
]
