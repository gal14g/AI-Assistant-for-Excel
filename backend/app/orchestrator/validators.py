"""
Pre-execution validation of AnalyticalPlan.
Validates that the plan is safe to execute before any tool runs.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from ..models.analytical_plan import AnalyticalPlan, IntentType, OperationType, SheetData

# Matches Excel range/column addresses like "C:D", "A:A", "B", "C2:D50"
# These should not be validated against header names — they are references
# that will be resolved by the orchestrator at execution time.
_RANGE_ADDR_RE = re.compile(r"^[A-Za-z]{1,3}(\d+)?(?::[A-Za-z]{1,3}(\d+)?)?$")


def _is_range_ref(col: str) -> bool:
    """True when *col* looks like an Excel range address, not a header name."""
    return bool(_RANGE_ADDR_RE.match(col.strip()))


# ── Result type ───────────────────────────────────────────────────────────────


@dataclass
class ValidationResult:
    """Outcome of pre-execution plan validation."""

    valid: bool
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


# ── Helpers ───────────────────────────────────────────────────────────────────


def _get_sheet_name(parameters: dict, key: str) -> str | None:
    """
    Extract a sheet name from *parameters* by *key*.

    Returns the string value or None if the key is absent / not a string.
    """
    val = parameters.get(key)
    if val is None:
        return None
    return str(val) if isinstance(val, str) else str(val)


def _resolve_columns(parameters: dict, key: str) -> list[str]:
    """
    Extract a column list from *parameters* by *key*.

    Handles two formats:
    - A flat list:   ["col_a", "col_b"]
    - A split dict:  {"left": ["col_a"], "right": ["col_b"]}

    In the dict case both sides are concatenated so existence can be checked
    against the appropriate sheet in the caller.
    """
    val = parameters.get(key)
    if val is None:
        return []
    if isinstance(val, list):
        return [str(c) for c in val]
    if isinstance(val, dict):
        left = val.get("left", [])
        right = val.get("right", [])
        if isinstance(left, list) and isinstance(right, list):
            return [str(c) for c in left] + [str(c) for c in right]
        # Fall back: return all dict values that are lists
        result: list[str] = []
        for v in val.values():
            if isinstance(v, list):
                result.extend(str(c) for c in v)
        return result
    return []


# ── Intent-tool compatibility map ─────────────────────────────────────────────

# Maps intent → set of OperationType values that are expected to satisfy it.
# At least one of the required ops must appear in the tool chain.
_INTENT_REQUIRED_OPS: dict[str, set[str]] = {
    IntentType.match_rows.value: {
        OperationType.run_exact_match.value,
        OperationType.run_fuzzy_match.value,
        OperationType.run_semantic_match.value,
        OperationType.run_hybrid_match.value,
    },
    IntentType.find_duplicates.value: {
        OperationType.find_duplicates.value,
    },
    IntentType.aggregate.value: {
        OperationType.aggregate_values.value,
    },
    IntentType.group_and_summarize.value: {
        OperationType.aggregate_values.value,
    },
    IntentType.filter_rows.value: {
        OperationType.filter_rows.value,
    },
    IntentType.compare_sheets.value: {
        OperationType.compare_sheets.value,
    },
    IntentType.profile_sheet.value: {
        OperationType.profile_columns.value,
    },
    IntentType.clean_data.value: {
        OperationType.clean_columns.value,
    },
    IntentType.semantic_lookup.value: {
        OperationType.run_semantic_match.value,
        OperationType.run_hybrid_match.value,
    },
    # answer_question and ask_clarification are intentionally permissive
    IntentType.answer_question.value: set(),
    IntentType.ask_clarification.value: set(),
    "explore": set(),
    "unknown": set(),
}

# Threshold-related parameter keys
_THRESHOLD_PARAMS = {"fuzzy_threshold", "final_threshold", "confidence", "semantic_threshold"}


# ── Main validator ─────────────────────────────────────────────────────────────


def validate_plan(
    plan: AnalyticalPlan,
    sheets: dict[str, SheetData],
) -> ValidationResult:
    """
    Validate *plan* against *sheets* before any tool execution.

    Checks (in order):
    a. Tool chain non-empty
    b. Sheets referenced in parameters exist in *sheets*
    c. Columns referenced in parameters exist in the respective sheets
    d. Intent-tool compatibility
    e. Threshold ranges (0–1)
    f. Clarification consistency

    Returns a ValidationResult.  If any errors are found, ``valid`` is False.
    """
    errors: list[str] = []
    warnings: list[str] = []
    params = plan.parameters

    # ── a. Tool chain non-empty ───────────────────────────────────────────────
    if not plan.selected_tool_chain:
        errors.append("selected_tool_chain is empty — no operations to execute.")
        # Cannot validate further without a tool chain
        return ValidationResult(valid=False, errors=errors, warnings=warnings)

    chain_values = {op.value if isinstance(op, OperationType) else str(op) for op in plan.selected_tool_chain}

    # ── b. Sheet existence ────────────────────────────────────────────────────
    sheet_param_keys = ("left_sheet", "right_sheet", "sheet_name")
    for key in sheet_param_keys:
        sheet_name = _get_sheet_name(params, key)
        if sheet_name is not None and sheet_name not in sheets:
            errors.append(
                f"Sheet '{sheet_name}' (from parameter '{key}') not found. "
                f"Available sheets: {list(sheets.keys())}"
            )

    # ── c. Column existence ───────────────────────────────────────────────────
    # These checks are best-effort: we skip if the sheet itself is missing (already reported above).
    _check_columns(params, sheets, errors, warnings)

    # ── d. Intent-tool compatibility ──────────────────────────────────────────
    intent_val = plan.intent.value if hasattr(plan.intent, "value") else str(plan.intent)
    required_ops = _INTENT_REQUIRED_OPS.get(intent_val)

    if required_ops is not None and len(required_ops) > 0:
        if not required_ops.intersection(chain_values):
            warnings.append(
                f"Intent '{intent_val}' typically requires one of "
                f"{sorted(required_ops)}, but none are in the tool chain. "
                f"Tool chain: {sorted(chain_values)}"
            )

    # ── e. Threshold ranges ───────────────────────────────────────────────────
    for key in _THRESHOLD_PARAMS:
        val = params.get(key)
        if val is not None:
            try:
                fval = float(val)
            except (TypeError, ValueError):
                errors.append(f"Parameter '{key}' must be a number, got: {val!r}")
                continue
            if not (0.0 <= fval <= 1.0):
                errors.append(
                    f"Parameter '{key}' must be in range [0, 1], got: {fval}"
                )

    # ── f. Clarification consistency ──────────────────────────────────────────
    if plan.needs_clarification and plan.clarification_question is None:
        errors.append(
            "needs_clarification is True but clarification_question is None. "
            "Provide a clarification_question when needs_clarification=True."
        )

    return ValidationResult(
        valid=len(errors) == 0,
        errors=errors,
        warnings=warnings,
    )


# ── Column existence checker (internal) ──────────────────────────────────────


_COLUMN_KEYS_SIMPLE = ("candidate_columns", "columns", "key_columns", "group_by")
_COLUMN_KEYS_PAIRED = {
    "left_columns": "left_sheet",
    "right_columns": "right_sheet",
}


def _check_columns(
    params: dict,
    sheets: dict[str, SheetData],
    errors: list[str],
    warnings: list[str],
) -> None:
    """
    Check that referenced column names exist in the appropriate sheets.

    Simple column keys (candidate_columns, columns, key_columns, group_by) are
    resolved against 'sheet_name'.  Paired keys (left_columns → left_sheet,
    right_columns → right_sheet) are resolved against their paired sheet.
    """
    # Simple: check against sheet_name
    sheet_name = _get_sheet_name(params, "sheet_name")
    sheet = sheets.get(sheet_name) if sheet_name else None

    for key in _COLUMN_KEYS_SIMPLE:
        cols = _resolve_columns(params, key)
        if not cols:
            continue
        if sheet is None:
            # Sheet missing → already reported or not specified; skip column check
            continue
        available = set(sheet.header_row if hasattr(sheet, "header_row") else sheet.headers or [])
        missing = [c for c in cols if c not in available and not _is_range_ref(c)]
        if missing:
            errors.append(
                f"Column(s) {missing} (from parameter '{key}') not found in "
                f"sheet '{sheet_name}'. Available: {sorted(available)}"
            )

    # Paired: left_columns → left_sheet, right_columns → right_sheet
    for col_key, sheet_key in _COLUMN_KEYS_PAIRED.items():
        cols_val = params.get(col_key)
        if cols_val is None:
            continue

        paired_sheet_name = _get_sheet_name(params, sheet_key)
        if paired_sheet_name is None or paired_sheet_name not in sheets:
            # Sheet already reported missing; skip
            continue

        paired_sheet = sheets[paired_sheet_name]
        available = set(
            paired_sheet.header_row if hasattr(paired_sheet, "header_row")
            else (paired_sheet.headers or [])
        )

        if isinstance(cols_val, dict):
            # {"left": [...], "right": [...]} — check relevant side only
            side = "left" if col_key == "left_columns" else "right"
            cols_list = cols_val.get(side, [])
        elif isinstance(cols_val, list):
            cols_list = cols_val
        else:
            cols_list = []

        missing = [str(c) for c in cols_list if str(c) not in available and not _is_range_ref(str(c))]
        if missing:
            errors.append(
                f"Column(s) {missing} (from parameter '{col_key}') not found in "
                f"sheet '{paired_sheet_name}'. Available: {sorted(available)}"
            )
