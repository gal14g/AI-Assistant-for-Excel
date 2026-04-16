"""
MCP (Model Context Protocol) server for Excel Copilot.

Exposes Excel Copilot's planner + xlwings execution bridge to any
MCP-compatible desktop chat client (Claude Desktop, Cursor, Windsurf,
Zed, …) via stdio transport.

Registered as the `excel-copilot-mcp` console script (see pyproject.toml),
so the usual MCP client config works out of the box:

    {
      "mcpServers": {
        "excel-copilot": {
          "command": "excel-copilot-mcp"
        }
      }
    }

Tool surface:
    list_open_workbooks       — enumerate open Excel windows + books
    get_active_workbook       — currently frontmost workbook
    get_workbook_snapshot     — per-sheet metadata (headers, dtypes, used range)
    get_selection             — current cell/range selection on a book
    resolve_range_token       — expand `[[Sheet!A1:C10]]` into an address
    search_capabilities       — natural-language → top-k action names
    generate_plan             — user message + snapshot → ExecutionPlan JSON
    validate_plan             — schema + binding + cycle check
    execute_plan              — run the plan via xlwings (with snapshots)
    undo_last                 — pop the latest snapshot off the stack
    analyze_sheet             — deterministic analytical pipeline (no Excel needed)
    list/load/rename/delete_conversation — shared multi-turn history

The server is deliberately thin: every tool delegates to an existing
backend service (or to xlwings via `XlwingsExecutor`). MCP is just a new
edge, not a parallel stack.
"""

from __future__ import annotations

import asyncio
import json
import logging
import sys
from typing import Any, Optional

from app.config import settings

log = logging.getLogger("mcp")


def _bail(message: str) -> None:
    """Log an error + exit 2 so MCP clients show a useful failure."""
    log.error(message)
    sys.stderr.write(f"[excel-copilot-mcp] {message}\n")
    sys.exit(2)


# ── Lazy imports (keep boot time fast; fail-fast on missing deps) ──────────


def _import_mcp():
    try:
        from mcp.server import Server
        from mcp.server.stdio import stdio_server
        import mcp.types as types
    except ImportError:
        _bail(
            "`mcp` package is not installed. Run "
            "`pip install mcp>=1.0.0` and retry."
        )
    return Server, stdio_server, types  # type: ignore[return-value]


def _import_executor():
    try:
        from app.execution.xlwings_bridge import XlwingsExecutor
        # Force capability registration.
        import app.execution.capabilities  # noqa: F401
    except Exception as exc:  # noqa: BLE001
        _bail(
            "Failed to import xlwings bridge. Ensure `xlwings` is installed "
            f"and Excel is available on this machine. Original error: {exc}"
        )
    return XlwingsExecutor


# ── Tool handlers (each returns a JSON-serialisable dict) ──────────────────


async def _tool_list_open_workbooks(args: dict[str, Any]) -> dict[str, Any]:
    import xlwings as xw

    if not xw.apps:
        return {"workbooks": [], "note": "No running Excel instance."}
    books = []
    for app in xw.apps:
        for b in app.books:
            books.append({
                "name": b.name,
                "fullName": b.fullname,
                "sheets": [s.name for s in b.sheets],
                "isActive": b == app.books.active,
            })
    return {"workbooks": books}


async def _tool_get_active_workbook(args: dict[str, Any]) -> dict[str, Any]:
    import xlwings as xw

    if not xw.apps:
        return {"error": "No running Excel instance."}
    app = xw.apps.active
    book = app.books.active
    return {
        "name": book.name,
        "fullName": book.fullname,
        "activeSheet": book.sheets.active.name,
        "sheets": [s.name for s in book.sheets],
    }


async def _tool_get_workbook_snapshot(args: dict[str, Any]) -> dict[str, Any]:
    """Per-sheet metadata — mirrors the frontend's workbookSnapshot.ts."""
    import xlwings as xw

    target = args.get("workbook_name")
    if not xw.apps:
        return {"error": "No running Excel instance."}
    app = xw.apps.active
    book = app.books[target] if target else app.books.active

    sheets_data = []
    for sheet in book.sheets:
        used = sheet.used_range
        headers = used.rows[0].value if used.count > 0 else []
        if isinstance(headers, (str, int, float)) or headers is None:
            headers = [headers]
        # First 20 rows of data (excluding header).
        preview_rows = min(20, max(0, used.rows.count - 1))
        preview = used.rows[1 : 1 + preview_rows].value if preview_rows else []
        if preview and not isinstance(preview[0], list):
            preview = [preview]
        sheets_data.append({
            "name": sheet.name,
            "rowCount": used.rows.count,
            "columnCount": used.columns.count,
            "usedRange": used.address,
            "headers": headers,
            "preview": preview,
        })
    return {"workbook": book.name, "sheets": sheets_data}


async def _tool_get_selection(args: dict[str, Any]) -> dict[str, Any]:
    import xlwings as xw

    target = args.get("workbook_name")
    if not xw.apps:
        return {"error": "No running Excel instance."}
    app = xw.apps.active
    book = app.books[target] if target else app.books.active
    sel = book.app.selection
    return {
        "workbook": book.name,
        "sheet": sel.sheet.name,
        "address": sel.address,
        "value": sel.value if sel.count <= 100 else "<large range — omitted>",
    }


async def _tool_resolve_range_token(args: dict[str, Any]) -> dict[str, Any]:
    from app.execution.range_utils import resolve_range
    import xlwings as xw

    token = args.get("token")
    if not token:
        return {"error": "Missing 'token'."}
    if not xw.apps:
        return {"error": "No running Excel instance."}
    target = args.get("workbook_name")
    app = xw.apps.active
    book = app.books[target] if target else app.books.active
    try:
        rng = resolve_range(book, token)
        return {
            "workbook": rng.sheet.book.name,
            "sheet": rng.sheet.name,
            "address": rng.address,
        }
    except Exception as exc:  # noqa: BLE001
        return {"error": f"Could not resolve {token!r}: {exc}"}


async def _tool_search_capabilities(args: dict[str, Any]) -> dict[str, Any]:
    from app.services.capability_store import search_capabilities

    query = args.get("query") or ""
    top_k = int(args.get("top_k") or 10)
    results = await search_capabilities(query, top_k)
    return {"results": results}


async def _tool_generate_plan(args: dict[str, Any]) -> dict[str, Any]:
    from app.services.chat_service import chat

    user_message = args.get("user_message") or ""
    workbook_snapshot = args.get("workbook_snapshot") or {}
    history = args.get("conversation_history") or []
    plan = await chat(user_message, workbook_snapshot, history)
    return plan


async def _tool_validate_plan(args: dict[str, Any]) -> dict[str, Any]:
    from app.services.validator import validate_plan
    from app.models.plan import ExecutionPlan

    raw = args.get("plan")
    if not raw:
        return {"error": "Missing 'plan'."}
    plan = ExecutionPlan.model_validate(raw) if not isinstance(raw, ExecutionPlan) else raw
    issues = validate_plan(plan)
    return {"valid": not issues, "issues": issues}


async def _tool_execute_plan(args: dict[str, Any]) -> dict[str, Any]:
    from app.models.plan import ExecutionPlan

    XlwingsExecutor = _import_executor()
    raw = args.get("plan")
    target = args.get("target_workbook")
    dry_run = bool(args.get("dry_run", False))
    if not raw:
        return {"error": "Missing 'plan'."}
    plan = ExecutionPlan.model_validate(raw)
    executor = XlwingsExecutor()
    return await executor.execute_plan(plan, target_workbook=target, dry_run=dry_run)


async def _tool_undo_last(args: dict[str, Any]) -> dict[str, Any]:
    XlwingsExecutor = _import_executor()
    executor = XlwingsExecutor()
    return await executor.undo_last(args.get("workbook_name"))


async def _tool_analyze_sheet(args: dict[str, Any]) -> dict[str, Any]:
    # Re-use the /api/analyze orchestrator for the deterministic analytical
    # pipeline. It doesn't need Excel to be running.
    from app.routers.analyze import analyze_impl  # type: ignore[attr-defined]

    return await analyze_impl(args)


# ── Server wiring ──────────────────────────────────────────────────────────


def _build_server():
    Server, stdio_server, types = _import_mcp()

    server = Server("excel-copilot")

    TOOLS = [
        ("list_open_workbooks", "Enumerate every workbook currently open in Excel.", _tool_list_open_workbooks, {}),
        ("get_active_workbook", "Return name + path of the frontmost workbook.", _tool_get_active_workbook, {}),
        (
            "get_workbook_snapshot",
            "Per-sheet metadata (name, row/col count, headers, first ~20 rows).",
            _tool_get_workbook_snapshot,
            {"workbook_name": {"type": "string", "description": "Optional: target book name. Defaults to active."}},
        ),
        (
            "get_selection",
            "Current cell/range selection in a workbook — resolves user's click+copy into an address.",
            _tool_get_selection,
            {"workbook_name": {"type": "string"}},
        ),
        (
            "resolve_range_token",
            "Expand a `[[Sheet!Range]]` or `[[Book.xlsx]Sheet!Range]]` token into a concrete address.",
            _tool_resolve_range_token,
            {"token": {"type": "string"}, "workbook_name": {"type": "string"}},
        ),
        (
            "search_capabilities",
            "Natural-language → top-k action names via vector search.",
            _tool_search_capabilities,
            {"query": {"type": "string"}, "top_k": {"type": "integer", "default": 10}},
        ),
        (
            "generate_plan",
            "LLM-driven: produce an ExecutionPlan for a user message, grounded in a workbook snapshot.",
            _tool_generate_plan,
            {
                "user_message": {"type": "string"},
                "workbook_snapshot": {"type": "object"},
                "conversation_history": {"type": "array", "items": {"type": "object"}},
            },
        ),
        (
            "validate_plan",
            "Schema + binding + cycle check on an ExecutionPlan.",
            _tool_validate_plan,
            {"plan": {"type": "object"}},
        ),
        (
            "execute_plan",
            "Run a validated ExecutionPlan via xlwings. Set dry_run=true to preview.",
            _tool_execute_plan,
            {
                "plan": {"type": "object"},
                "target_workbook": {"type": "string"},
                "dry_run": {"type": "boolean", "default": False},
            },
        ),
        (
            "undo_last",
            "Pop the most recent execution snapshot off the stack and restore.",
            _tool_undo_last,
            {"workbook_name": {"type": "string"}},
        ),
        (
            "analyze_sheet",
            "Deterministic analytical pipeline — runs on sheet dumps without needing Excel open.",
            _tool_analyze_sheet,
            {"sheets": {"type": "array"}, "user_message": {"type": "string"}},
        ),
    ]

    @server.list_tools()
    async def _list_tools():  # noqa: ANN202
        return [
            types.Tool(
                name=name,
                description=desc,
                inputSchema={
                    "type": "object",
                    "properties": schema_props,
                    "additionalProperties": True,
                },
            )
            for name, desc, _fn, schema_props in TOOLS
        ]

    @server.call_tool()
    async def _call_tool(name: str, arguments: dict[str, Any]):  # noqa: ANN202
        tool = next((t for t in TOOLS if t[0] == name), None)
        if tool is None:
            return [types.TextContent(type="text", text=f"Unknown tool: {name}")]
        _, _, fn, _ = tool
        try:
            result = await fn(arguments or {})
        except Exception as exc:  # noqa: BLE001
            log.exception("Tool %s crashed", name)
            return [types.TextContent(type="text", text=f"Tool {name} errored: {exc}")]
        return [types.TextContent(type="text", text=json.dumps(result, default=str, indent=2))]

    return server, stdio_server


async def _serve() -> None:
    server, stdio_server = _build_server()
    async with stdio_server() as (read_stream, write_stream):
        await server.run(read_stream, write_stream, server.create_initialization_options())


def main() -> None:
    """Entry point for the `excel-copilot-mcp` console script."""
    logging.basicConfig(level=logging.INFO if not settings.debug else logging.DEBUG)
    if settings.mcp_mode == "disabled":
        log.info("mcp_mode=disabled — server starting anyway (CLI override).")
    asyncio.run(_serve())


if __name__ == "__main__":
    main()
