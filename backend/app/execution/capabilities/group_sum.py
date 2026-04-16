"""groupSum — aggregate values by group using SUMIF formulas or computed values.

Two strategies (selected by ``preferFormula`` — default True):
  * Formula strategy: write SUMIF formulas that auto-update when source data
    changes. Group keys are written as values, the sum column as formulas.
  * Computed strategy: read data, build a {group: sum} dict in Python, write
    the fully materialized 2-column table.

When ``includeHeaders`` is true (default), the first row of the source range
is treated as headers and an ``Sum of <name>`` header is written to the output.
"""

from __future__ import annotations

import re
from typing import Any

from app.execution.base import ExecutorContext
from app.execution.capability_registry import registry
from app.execution.range_utils import normalize_token, resolve_range


_SHEET_SAFE_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")


def _strip_workbook_qualifier(address: str) -> str:
    """Drop the leading ``[Book.xlsx]`` segment if present."""
    addr = normalize_token(address)
    return re.sub(r"^\[.*?\]", "", addr)


def _quote_sheet_in_ref(ref: str) -> str:
    """Mirror the TS ``quoteSheetInRef`` helper used inside formula strings."""
    bang = ref.rfind("!")
    if bang == -1:
        return ref
    sheet_part = ref[:bang]
    cell_part = ref[bang + 1 :]
    if sheet_part.startswith("'") and sheet_part.endswith("'"):
        return ref
    if _SHEET_SAFE_RE.match(sheet_part):
        return ref
    escaped = sheet_part.replace("'", "''")
    return f"'{escaped}'!{cell_part}"


def _offset_column_letter(col: str, offset: int) -> str:
    num = 0
    for ch in col:
        num = num * 26 + (ord(ch.upper()) - 64)
    num += offset
    result = ""
    while num > 0:
        num, rem = divmod(num - 1, 26)
        result = chr(65 + rem) + result
    return result or "A"


def _column_for_range(range_address: str, col_offset: int) -> str:
    """Return a sheet-qualified ``A:A`` column reference for SUMIF."""
    parts = range_address.split("!")
    ref = parts[1] if len(parts) > 1 else parts[0]
    m = re.search(r"[A-Z]+", ref, re.IGNORECASE)
    col = m.group(0).upper() if m else "A"
    offset_col = _offset_column_letter(col, col_offset)
    prefix = parts[0] + "!" if len(parts) > 1 else ""
    return _quote_sheet_in_ref(f"{prefix}{offset_col}:{offset_col}")


def _normalize_to_2d(raw: Any, shape: tuple[int, int]) -> list[list[Any]]:
    if raw is None:
        return []
    if not isinstance(raw, list):
        return [[raw]]
    if raw and not isinstance(raw[0], list):
        rows_, _ = shape
        return [list(raw)] if rows_ == 1 else [[v] for v in raw]
    return [list(r) for r in raw]


def _to_float(x: Any) -> float:
    try:
        if x is None or x == "":
            return 0.0
        return float(x)
    except (TypeError, ValueError):
        return 0.0


def _formula_group_sum(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    data_range = params["dataRange"]
    group_by = int(params["groupByColumn"])
    sum_col = int(params["sumColumn"])
    output_range = params["outputRange"]
    include_headers = bool(params.get("includeHeaders", True))

    data_rng = resolve_range(ctx.workbook_handle, data_range)
    data = _normalize_to_2d(data_rng.value, data_rng.shape)

    start_row = 1 if include_headers else 0
    group_idx = group_by - 1

    # Extract unique group keys (preserve first-seen ordering).
    unique_keys: list[Any] = []
    seen: set[str] = set()
    for i in range(start_row, len(data)):
        if group_idx >= len(data[i]):
            continue
        key_display = data[i][group_idx]
        key_str = str(key_display)
        if key_str in seen:
            continue
        seen.add(key_str)
        unique_keys.append(key_display)

    data_range_for_formula = _strip_workbook_qualifier(data_range)
    criteria_col = _column_for_range(data_range_for_formula, group_idx)
    sum_col_ref = _column_for_range(data_range_for_formula, sum_col - 1)

    # Write optional header row.
    out_rng = resolve_range(ctx.workbook_handle, output_range)
    sheet = out_rng.sheet
    start_out_row = out_rng.row
    start_out_col = out_rng.column

    def _col_letter(col_num: int) -> str:
        letters = ""
        n = col_num
        while n > 0:
            n, rem = divmod(n - 1, 26)
            letters = chr(65 + rem) + letters
        return letters

    key_col_letter = _col_letter(start_out_col)
    val_col_letter = _col_letter(start_out_col + 1)

    row_cursor = start_out_row
    if include_headers and data:
        header_group = str(data[0][group_idx]) if group_idx < len(data[0]) else ""
        header_sum = str(data[0][sum_col - 1]) if (sum_col - 1) < len(data[0]) else ""
        sheet.range(f"{key_col_letter}{row_cursor}").value = header_group
        sheet.range(f"{val_col_letter}{row_cursor}").value = f"Sum of {header_sum}"
        row_cursor += 1

    # Write group keys as values, SUMIF formulas next to them.
    if unique_keys:
        keys_block = [[k] for k in unique_keys]
        formulas_block = [
            [
                f"=SUMIF({criteria_col},"
                + (f'"{k}"' if isinstance(k, str) else str(k))
                + f",{sum_col_ref})"
            ]
            for k in unique_keys
        ]
        n = len(unique_keys)
        sheet.range(f"{key_col_letter}{row_cursor}").resize(n, 1).value = keys_block
        sheet.range(f"{val_col_letter}{row_cursor}").resize(n, 1).formula = formulas_block

    return {
        "status": "success",
        "message": f"Created {len(unique_keys)} SUMIF formulas in {output_range}.",
        "outputs": {"outputRange": output_range, "groupCount": len(unique_keys)},
    }


def _computed_group_sum(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    data_range = params["dataRange"]
    group_by = int(params["groupByColumn"])
    sum_col = int(params["sumColumn"])
    output_range = params["outputRange"]
    include_headers = bool(params.get("includeHeaders", True))

    data_rng = resolve_range(ctx.workbook_handle, data_range)
    data = _normalize_to_2d(data_rng.value, data_rng.shape)

    start_row = 1 if include_headers else 0
    group_idx = group_by - 1
    sum_idx = sum_col - 1

    groups: dict[str, list[Any]] = {}
    for i in range(start_row, len(data)):
        if group_idx >= len(data[i]):
            continue
        display = data[i][group_idx]
        key = str(display)
        if key not in groups:
            groups[key] = [display, 0.0]
        value = _to_float(data[i][sum_idx] if sum_idx < len(data[i]) else 0)
        groups[key][1] += value

    output: list[list[Any]] = []
    if include_headers and data:
        header_group = str(data[0][group_idx]) if group_idx < len(data[0]) else ""
        header_sum = str(data[0][sum_idx]) if sum_idx < len(data[0]) else ""
        output.append([header_group, f"Sum of {header_sum}"])
    for display, total in groups.values():
        output.append([display, total])

    out_rng = resolve_range(ctx.workbook_handle, output_range)
    rows = len(output)
    cols = 2 if output else 0
    if rows and cols:
        out_rng.resize(rows, cols).value = output

    return {
        "status": "success",
        "message": f"Computed {len(groups)} group sums, wrote to {output_range}.",
        "outputs": {"outputRange": output_range, "groupCount": len(groups)},
    }


def handler(ctx: ExecutorContext, params: dict[str, Any]) -> dict[str, Any]:
    required = ("dataRange", "groupByColumn", "sumColumn", "outputRange")
    missing = [k for k in required if params.get(k) in (None, "")]
    if missing:
        return {
            "status": "error",
            "message": f"groupSum missing parameters: {', '.join(missing)}.",
        }

    if ctx.dry_run:
        return {
            "status": "preview",
            "message": (
                f"Would compute grouped sums from {params['dataRange']}, "
                f"output to {params['outputRange']}."
            ),
        }

    prefer_formula = bool(params.get("preferFormula", True))

    try:
        if prefer_formula:
            return _formula_group_sum(ctx, params)
        return _computed_group_sum(ctx, params)
    except Exception as exc:  # noqa: BLE001
        return {"status": "error", "message": f"groupSum failed: {exc}", "error": str(exc)}


registry.register("groupSum", handler, mutates=True)
