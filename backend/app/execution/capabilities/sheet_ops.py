"""
Sheet operations: addSheet, renameSheet, deleteSheet, copySheet, protectSheet.

Ported from `frontend/src/engine/capabilities/sheetOps.ts`. Multiple
registrations happen from this module — each action is its own StepAction
in the plan schema but they share enough logic to keep in one file.
"""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry


def _add_sheet(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    name = params.get("name") or params.get("sheetName")
    after = params.get("after")
    book = ctx.workbook_handle

    if not name:
        return {"status": "error", "message": "addSheet requires 'name'."}
    if ctx.dry_run:
        return {"status": "preview", "message": f"Would add sheet {name!r}."}

    # Idempotent — if it exists, return success without touching it.
    try:
        existing = book.sheets[name]
        return {
            "status": "success",
            "message": f"Sheet {name!r} already exists — reused.",
            "outputs": {"sheet": existing.name},
        }
    except Exception:
        pass

    sheet = book.sheets.add(name=name, after=after) if after else book.sheets.add(name=name)
    return {
        "status": "success",
        "message": f"Added sheet {sheet.name!r}.",
        "outputs": {"sheet": sheet.name},
    }


def _rename_sheet(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    old = params.get("sheetName") or params.get("oldName")
    new = params.get("newName")
    if not old or not new:
        return {"status": "error", "message": "renameSheet requires 'sheetName' and 'newName'."}
    book = ctx.workbook_handle
    if ctx.dry_run:
        return {"status": "preview", "message": f"Would rename {old!r} → {new!r}."}
    try:
        book.sheets[old].name = new
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Rename failed: {exc}", "error": str(exc)}
    return {
        "status": "success",
        "message": f"Renamed {old!r} → {new!r}.",
        "outputs": {"sheet": new},
    }


def _delete_sheet(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    name = params.get("sheetName") or params.get("name")
    if not name:
        return {"status": "error", "message": "deleteSheet requires 'sheetName'."}
    if ctx.dry_run:
        return {"status": "preview", "message": f"Would delete sheet {name!r}."}
    try:
        ctx.workbook_handle.sheets[name].delete()
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Delete failed: {exc}", "error": str(exc)}
    return {"status": "success", "message": f"Deleted sheet {name!r}."}


def _copy_sheet(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    source = params.get("sourceSheet") or params.get("sheetName")
    new_name = params.get("newName")
    if not source:
        return {"status": "error", "message": "copySheet requires 'sourceSheet'."}
    if ctx.dry_run:
        return {"status": "preview", "message": f"Would copy sheet {source!r}."}
    try:
        src = ctx.workbook_handle.sheets[source]
        copied = src.copy(before=None, after=src)
        if new_name:
            copied.name = new_name
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Copy failed: {exc}", "error": str(exc)}
    return {
        "status": "success",
        "message": f"Copied sheet {source!r}{f' as {new_name!r}' if new_name else ''}.",
        "outputs": {"sheet": copied.name},
    }


def _protect_sheet(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    name = params.get("sheetName") or params.get("name")
    password = params.get("password") or ""
    protect = bool(params.get("protect", True))
    book = ctx.workbook_handle
    sheet = book.sheets[name] if name else book.sheets.active
    if ctx.dry_run:
        return {"status": "preview", "message": f"Would {'protect' if protect else 'unprotect'} {sheet.name!r}."}
    try:
        if protect:
            sheet.api.Protect(Password=password) if password else sheet.api.Protect()
        else:
            sheet.api.Unprotect(Password=password) if password else sheet.api.Unprotect()
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"Protect failed: {exc}", "error": str(exc)}
    return {"status": "success", "message": f"{'Protected' if protect else 'Unprotected'} {sheet.name!r}."}


registry.register("addSheet", _add_sheet, mutates=True)
registry.register("renameSheet", _rename_sheet, mutates=True)
registry.register("deleteSheet", _delete_sheet, mutates=True)
registry.register("copySheet", _copy_sheet, mutates=True)
registry.register("protectSheet", _protect_sheet, mutates=False, affects_formatting=True)
