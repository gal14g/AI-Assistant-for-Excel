"""
insertPicture — insert a base64-encoded image into a worksheet.

The Office.js path used Shapes.addImage(base64); the COM equivalent
(Shapes.AddPicture) needs a path on disk, so we decode the base64 into a
temp file, insert it, and best-effort delete the temp file afterwards.
"""

from __future__ import annotations

import base64
import os
import tempfile
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry


def _strip_data_uri_prefix(b64: str) -> str:
    """Accept `data:image/png;base64,XXXX` as well as the bare payload."""
    if b64.startswith("data:"):
        comma = b64.find(",")
        if comma != -1:
            return b64[comma + 1 :]
    return b64


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    image_b64 = params.get("imageBase64")
    sheet_name = params.get("sheetName")
    left = params.get("left")
    top = params.get("top")
    width = params.get("width")
    height = params.get("height")
    alt_text = params.get("altText")

    if not image_b64:
        return {"status": "error", "message": "insertPicture requires 'imageBase64'."}

    if ctx.dry_run:
        on_sheet = f" on {sheet_name!r}" if sheet_name else ""
        return {
            "status": "preview",
            "message": f"Would insert picture{on_sheet}.",
        }

    book = ctx.workbook_handle
    try:
        if sheet_name:
            try:
                sheet = book.sheets[sheet_name]
            except Exception:
                return {
                    "status": "error",
                    "message": f"Sheet {sheet_name!r} not found. Please check the sheet name.",
                }
        else:
            sheet = book.sheets.active

        # Decode base64 → temp file → AddPicture(path).
        payload = _strip_data_uri_prefix(image_b64)
        try:
            image_bytes = base64.b64decode(payload)
        except Exception as exc:  # noqa: BLE001
            return {
                "status": "error",
                "message": f"Could not decode imageBase64: {exc}",
                "error": str(exc),
            }

        tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        tmp_path = tmp.name
        try:
            tmp.write(image_bytes)
            tmp.close()

            # AddPicture(Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
            # LinkToFile=False, SaveWithDocument=True → embed bytes in workbook.
            # Width/Height -1 keeps the native image dimension when unspecified.
            shape = sheet.api.Shapes.AddPicture(
                tmp_path,
                False,
                True,
                float(left) if left is not None else 0.0,
                float(top) if top is not None else 0.0,
                float(width) if width is not None else -1.0,
                float(height) if height is not None else -1.0,
            )

            if alt_text is not None:
                try:
                    shape.AlternativeText = str(alt_text)
                except Exception:  # noqa: BLE001
                    pass
        finally:
            # Best-effort cleanup; Excel copies the bytes into the workbook
            # when SaveWithDocument=True, so the file is safe to remove.
            try:
                os.unlink(tmp_path)
            except Exception:  # noqa: BLE001
                pass
    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": f"Insert picture failed: {exc}",
            "error": str(exc),
        }

    on_sheet = f" on {sheet.name!r}" if sheet_name else ""
    return {
        "status": "success",
        "message": f"Inserted picture{on_sheet}.",
        "outputs": {"sheet": sheet.name},
    }


registry.register("insertPicture", handler, mutates=True)
