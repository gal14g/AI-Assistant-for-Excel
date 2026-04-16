"""groupRows — group or ungroup rows/columns for outline collapsing."""

from __future__ import annotations

import re
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


_ROW_RANGE_RE = re.compile(r"^\d+:\d+$")


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    operation = params.get("operation", "group")
    if not address:
        return {"status": "error", "message": "groupRows requires 'range'."}
    if operation not in ("group", "ungroup"):
        return {"status": "error", "message": "groupRows 'operation' must be 'group' or 'ungroup'."}

    verb = "Group" if operation == "group" else "Ungroup"

    if ctx.dry_run:
        return {"status": "preview", "message": f"Would {verb.lower()} {address}."}

    try:
        rng = resolve_range(ctx.workbook_handle, address)

        # Detect row-vs-column range from the raw address body ("3:8" → rows, "B:E" → cols).
        ref = address.split("!", 1)[1] if "!" in address else address
        ref = ref.strip().strip("[]")  # tolerate [[A:A]] token wrappers
        is_row_range = bool(_ROW_RANGE_RE.match(ref))

        # Excel COM: Rows(...).Group() or Columns(...).Group() work directly;
        # a Range that spans entire rows/columns also accepts Group() which
        # defaults based on the range's shape.
        target = rng.api.EntireRow if is_row_range else rng.api.EntireColumn
        if operation == "group":
            target.Group()
        else:
            target.Ungroup()

    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"{verb} failed: {exc}", "error": str(exc)}

    return {
        "status": "success",
        "message": f"{verb}ed {address}.",
        "outputs": {"range": address, "operation": operation},
    }


registry.register("groupRows", handler, mutates=False, affects_formatting=True)
