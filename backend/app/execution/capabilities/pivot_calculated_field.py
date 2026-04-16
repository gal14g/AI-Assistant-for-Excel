"""pivotCalculatedField — add a calculated field with a custom formula to an existing PivotTable."""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    pivot_name = params.get("pivotName")
    sheet_name = params.get("sheetName")
    field_name = params.get("fieldName")
    formula = params.get("formula")

    if not pivot_name or not field_name or not formula:
        return {
            "status": "error",
            "message": "pivotCalculatedField requires 'pivotName', 'fieldName', and 'formula'.",
        }

    book = ctx.workbook_handle

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would add calculated field {field_name!r} to pivot {pivot_name!r}.",
        }

    try:
        sheet = book.sheets[sheet_name] if sheet_name else book.sheets.active
        pivot_table = sheet.api.PivotTables(pivot_name)

        # CalculatedFields() returns the collection; .Add(Name, Formula) appends a new one.
        # Formula must begin with "=" — normalize to match Excel's expectations.
        formula_str = formula if formula.startswith("=") else f"={formula}"
        pivot_table.CalculatedFields().Add(field_name, formula_str)

    except Exception as exc:  # noqa: BLE001
        message = str(exc)
        lower = message.lower()
        if (
            "calculatedfields" in lower
            or "not a function" in lower
            or "unknown name" in lower
            or "does not support" in lower
        ):
            return {
                "status": "error",
                "message": (
                    "Calculated fields are not available on this PivotTable "
                    "(it may be an OLAP/Data Model pivot). Add the field manually "
                    "via the PivotTable Fields pane."
                ),
                "error": message,
            }
        return {
            "status": "error",
            "message": f"Failed to add calculated field {field_name!r}: {message}",
            "error": message,
        }

    return {
        "status": "success",
        "message": f"Added calculated field {field_name!r} to pivot {pivot_name!r}.",
        "outputs": {"pivotName": pivot_name, "fieldName": field_name},
    }


registry.register("pivotCalculatedField", handler, mutates=True, affects_formatting=False)
