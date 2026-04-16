"""
conditionalFormula — build IF(...) formulas per row using a condition column.

For each data row we build `=IF(<cond>, <trueFormula>, <falseFormula>)` where
`trueFormula` and `falseFormula` are templates with `{row}` placeholders
substituted for the actual Excel row number. Mirrors the TS handler exactly.
"""

from __future__ import annotations

from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import resolve_range


def _column_number_to_letter(col: int) -> str:
    """1 → A, 2 → B, 27 → AA."""
    letter = ""
    while col > 0:
        mod = (col - 1) % 26
        letter = chr(65 + mod) + letter
        col = (col - 1) // 26
    return letter


def _condition_expr(condition: str, cell_ref: str, value: Any) -> str:
    """Build the Excel boolean expression for the condition."""
    if condition == "blank":
        return f'{cell_ref}=""'
    if condition == "notBlank":
        return f'{cell_ref}<>""'
    if condition == "equals":
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return f"{cell_ref}={value}"
        return f'{cell_ref}="{value}"'
    if condition == "notEquals":
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return f"{cell_ref}<>{value}"
        return f'{cell_ref}<>"{value}"'
    if condition == "contains":
        return f'ISNUMBER(SEARCH("{value}",{cell_ref}))'
    if condition == "greaterThan":
        return f"{cell_ref}>{value}"
    if condition == "lessThan":
        return f"{cell_ref}<{value}"
    # Fallback — treat as string equality.
    return f'{cell_ref}="{value}"'


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    address = params.get("range")
    condition_column = params.get("conditionColumn")
    condition = params.get("condition")
    condition_value = params.get("conditionValue")
    true_formula = params.get("trueFormula")
    false_formula = params.get("falseFormula")
    output_address = params.get("outputRange")
    has_headers = params.get("hasHeaders", True)

    if (
        not address
        or condition_column is None
        or not condition
        or true_formula is None
        or false_formula is None
        or not output_address
    ):
        return {
            "status": "error",
            "message": (
                "conditionalFormula requires 'range', 'conditionColumn', 'condition', "
                "'trueFormula', 'falseFormula', and 'outputRange'."
            ),
        }

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": f"Would create conditional formulas based on {condition!r} in {address}.",
        }

    try:
        source_rng = resolve_range(ctx.workbook_handle, address)
        total_rows = source_rng.shape[0]
        start_excel_row = source_rng.row  # 1-based from xlwings

        data_start_row = start_excel_row + 1 if has_headers else start_excel_row
        data_row_count = total_rows - 1 if has_headers else total_rows

        if data_row_count <= 0:
            return {
                "status": "success",
                "message": "No data rows to process.",
                "outputs": {},
            }

        col_letter = _column_number_to_letter(int(condition_column))

        formulas: list[list[str]] = []
        for i in range(data_row_count):
            row_num = data_start_row + i
            cell_ref = f"{col_letter}{row_num}"
            true_expr = str(true_formula).replace("{row}", str(row_num))
            false_expr = str(false_formula).replace("{row}", str(row_num))
            # Strip leading "=" from embedded sub-formulas (IF builds one formula).
            if true_expr.startswith("="):
                true_expr = true_expr[1:]
            if false_expr.startswith("="):
                false_expr = false_expr[1:]

            cond_expr = _condition_expr(condition, cell_ref, condition_value)
            formulas.append([f"=IF({cond_expr},{true_expr},{false_expr})"])

        out_rng = resolve_range(ctx.workbook_handle, output_address)

        if has_headers:
            # Write "Result" header in the first cell, formulas below.
            header_cell = out_rng.sheet.range((out_rng.row, out_rng.column))
            header_cell.value = "Result"
            data_start = out_rng.sheet.range((out_rng.row + 1, out_rng.column))
            target = data_start.resize(len(formulas), 1)
            target.formula = formulas
        else:
            target = out_rng.resize(len(formulas), 1)
            target.formula = formulas
    except Exception as exc:  # noqa: BLE001
        return {
            "status": "error",
            "message": f"Conditional formula failed: {exc}",
            "error": str(exc),
        }

    return {
        "status": "success",
        "message": f"Created {len(formulas)} conditional formulas.",
        "outputs": {"range": output_address},
    }


registry.register("conditionalFormula", handler, mutates=True)
