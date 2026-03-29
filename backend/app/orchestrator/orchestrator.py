"""
Orchestrator: executes AnalyticalPlan tool chains deterministically.

Responsibilities:
- Pre-validate the plan
- Execute tools in declared sequence
- Pass intermediate results via ExecutionContext
- Stop on error, collect warnings
- Handle clarification requests
- Never call the LLM — only deterministic Python tools
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import Any, Callable

from ..models.analytical_plan import AnalyticalPlan, OperationType, SheetData, StrategyType
from ..models.tool_output import ToolOutput
from .execution_context import ExecutionContext
from .validators import validate_plan, ValidationResult
from .. import tools as tool_module

logger = logging.getLogger(__name__)


# ── Result type ───────────────────────────────────────────────────────────────


@dataclass
class OrchestratorResult:
    """The final outcome of an orchestrated analytical pipeline run."""

    success: bool
    needs_clarification: bool = False
    clarification_question: str | None = None
    context: ExecutionContext | None = None
    errors: list[str] = field(default_factory=list)
    # The most relevant tool's data (last successful result's data dict)
    final_data: dict[str, Any] = field(default_factory=dict)


# ── Orchestrator ──────────────────────────────────────────────────────────────


class Orchestrator:
    """
    Executes an AnalyticalPlan by dispatching operations to the tools layer.

    The orchestrator is stateless between runs: each call to ``execute``
    creates a fresh ExecutionContext.  Tool outputs are threaded through the
    context so downstream tools can consume upstream results.

    Parameters
    ----------
    sheets:
        The available sheet data for this session, keyed by sheet name.
    """

    def __init__(self, sheets: dict[str, SheetData]) -> None:
        self.sheets = sheets
        self._registry: dict[str, Callable[..., ToolOutput]] = self._build_tool_registry()

    # ── Public API ─────────────────────────────────────────────────────────────

    def execute(self, plan: AnalyticalPlan) -> OrchestratorResult:
        """
        Execute *plan* and return an OrchestratorResult.

        Steps:
        1. Pre-validate the plan.
        2. If validation fails → return failure with errors.
        3. If needs_clarification → return success with clarification flag.
        4. For each operation in the tool chain:
           a. Dispatch to the appropriate tool.
           b. Store result in context.
           c. If the tool fails → stop and return failure.
        5. Return success with populated context and final_data.
        """
        context = ExecutionContext()
        context.log(f"Starting pipeline for intent='{plan.intent}'")

        # Step 1–2: Validate
        validation: ValidationResult = validate_plan(plan, self.sheets)
        if not validation.valid:
            context.log(f"Validation failed: {validation.errors}")
            return OrchestratorResult(
                success=False,
                errors=validation.errors,
                context=context,
            )

        # Propagate validation warnings into the context log
        for w in validation.warnings:
            context.log(f"[validation warning] {w}")

        # Step 3: Clarification gate
        if plan.needs_clarification:
            context.log("Plan requires clarification before execution.")
            return OrchestratorResult(
                success=True,
                needs_clarification=True,
                clarification_question=plan.clarification_question,
                context=context,
            )

        # Step 4: Execute tool chain
        for operation in plan.selected_tool_chain:
            op_name = operation.value if isinstance(operation, OperationType) else str(operation)
            context.log(f"Dispatching: {op_name}")

            try:
                result = self._dispatch_tool(operation, plan, context)
            except Exception as exc:
                logger.exception("Unexpected error in tool '%s'", op_name)
                error_msg = f"Tool '{op_name}' raised an unexpected error: {exc}"
                context.log(f"[ERROR] {error_msg}")
                return OrchestratorResult(
                    success=False,
                    errors=[error_msg],
                    context=context,
                )

            context.store(op_name, result)

            if not result.success:
                context.log(f"[STOP] Tool '{op_name}' failed: {result.errors}")
                return OrchestratorResult(
                    success=False,
                    errors=result.errors,
                    context=context,
                )

        # Step 5: Build final result
        context.log("Pipeline complete.")
        last = context.last_result()
        final_data: dict[str, Any] = {}
        if last is not None and isinstance(last.data, dict):
            final_data = last.data
        elif last is not None and last.data is not None:
            final_data = {"result": last.data}

        return OrchestratorResult(
            success=True,
            context=context,
            final_data=final_data,
        )

    # ── Dispatch ───────────────────────────────────────────────────────────────

    def _dispatch_tool(
        self,
        operation: OperationType,
        plan: AnalyticalPlan,
        context: ExecutionContext,
    ) -> ToolOutput:
        """
        Route *operation* to the correct tool function and return its ToolOutput.

        Arguments are resolved from plan.parameters and, where applicable, from
        prior tool outputs stored in *context*.
        """
        op_name = operation.value if isinstance(operation, OperationType) else str(operation)
        params = plan.parameters

        # ── Discovery ─────────────────────────────────────────────────────────

        if op_name == OperationType.list_sheets.value:
            return tool_module.list_sheets(self.sheets)

        if op_name == OperationType.get_sheet_schema.value:
            sheet = self._require_sheet(params.get("sheet_name", ""), op_name)
            if isinstance(sheet, ToolOutput):
                return sheet
            return tool_module.get_sheet_schema(sheet)

        if op_name == OperationType.preview_sheet.value:
            sheet = self._require_sheet(params.get("sheet_name", ""), op_name)
            if isinstance(sheet, ToolOutput):
                return sheet
            n_rows = int(params.get("n_rows", 20))
            return tool_module.preview_sheet(sheet, n_rows=n_rows)

        # ── Profiling ─────────────────────────────────────────────────────────

        if op_name == OperationType.profile_columns.value:
            return self._resolve_profile_inputs(params, op_name)

        if op_name == OperationType.clean_columns.value:
            sheet_name = params.get("sheet_name", "")
            sheet = self._require_sheet(sheet_name, op_name)
            if isinstance(sheet, ToolOutput):
                return sheet
            columns = _coerce_list(params.get("columns") or params.get("key_columns") or [])
            options = params.get("options") or params.get("clean_options") or None
            return tool_module.clean_columns(sheet, columns=columns, options=options)

        # ── Matchability ──────────────────────────────────────────────────────

        if op_name == OperationType.estimate_matchability.value:
            return self._resolve_match_inputs(
                op_name, params, context, tool_module.estimate_matchability
            )

        # ── Match tools ───────────────────────────────────────────────────────

        if op_name in (
            OperationType.run_exact_match.value,
            OperationType.run_fuzzy_match.value,
            OperationType.run_semantic_match.value,
            OperationType.run_hybrid_match.value,
        ):
            return self._resolve_match_inputs(
                op_name, params, context, self._registry[op_name]
            )

        # ── Analysis ──────────────────────────────────────────────────────────

        if op_name == OperationType.find_duplicates.value:
            sheet_name = params.get("sheet_name", "")
            sheet = self._require_sheet(sheet_name, op_name)
            if isinstance(sheet, ToolOutput):
                return sheet
            key_columns = _coerce_list(
                params.get("key_columns") or params.get("columns") or []
            ) or None
            # "mode" maps to "keep" param: "first" / "last" / "all" → default first
            keep = params.get("keep", params.get("mode", "first"))
            if keep not in ("first", "last"):
                keep = "first"
            return tool_module.find_duplicates(sheet, key_columns=key_columns, keep=keep)

        if op_name == OperationType.aggregate_values.value:
            sheet_name = params.get("sheet_name", "")
            sheet = self._require_sheet(sheet_name, op_name)
            if isinstance(sheet, ToolOutput):
                return sheet
            group_by = _coerce_list(
                params.get("group_by_columns") or params.get("group_by") or []
            )
            # Resolve agg_column / agg_function from multiple param shapes
            agg_column = params.get("agg_column", "")
            agg_function = params.get("agg_function", "sum")
            # Also accept a "metrics" list: take the first entry
            metrics = params.get("metrics") or []
            if not agg_column and metrics:
                first = metrics[0] if isinstance(metrics, list) else {}
                agg_column = first.get("column", "")
                agg_function = first.get("function", "sum") or "sum"
            # Fall back to agg_columns dict
            if not agg_column:
                agg_cols = params.get("agg_columns") or {}
                if isinstance(agg_cols, dict) and agg_cols:
                    agg_column, agg_function = next(iter(agg_cols.items()))
            return tool_module.aggregate_values(
                sheet,
                group_by_columns=group_by,
                agg_column=agg_column,
                agg_function=agg_function,
            )

        if op_name == OperationType.filter_rows.value:
            sheet_name = params.get("sheet_name", "")
            sheet = self._require_sheet(sheet_name, op_name)
            if isinstance(sheet, ToolOutput):
                return sheet
            # Support a list of filter dicts OR single column/operator/value
            filters = params.get("filters", [])
            if not filters:
                col = params.get("column", "")
                op = params.get("operator", "eq")
                val = params.get("value", "")
                if col:
                    filters = [{"column": col, "operator": op, "value": val}]
            if not filters:
                return ToolOutput.fail(
                    tool_name=op_name,
                    errors=["filter_rows requires 'filters' list or 'column'/'operator'/'value' params."],
                )
            # Apply filters sequentially (AND logic via successive calls)
            result: ToolOutput = ToolOutput.fail(tool_name=op_name, errors=["no filters applied"])
            current_sheet = sheet
            for f in filters:
                result = tool_module.filter_rows(
                    current_sheet,
                    column=f.get("column", ""),
                    operator=f.get("operator", "eq"),
                    value=f.get("value", ""),
                )
                if not result.success:
                    return result
                # Build a new SheetData from filtered rows for chaining
                if isinstance(result.data, dict) and "rows" in result.data:
                    from ..models.analytical_plan import SheetData as _SD
                    row_data = [[v for v in r.values()] for r in result.data["rows"]]
                    current_sheet = _SD(
                        name=current_sheet.name,
                        data=row_data,
                        headers=current_sheet.header_row,
                    )
            return result

        if op_name == OperationType.compare_sheets.value:
            left_name = params.get("left_sheet", "")
            right_name = params.get("right_sheet", "")
            left = self._require_sheet(left_name, op_name)
            right = self._require_sheet(right_name, op_name)
            if isinstance(left, ToolOutput):
                return left
            if isinstance(right, ToolOutput):
                return right
            # Resolve key columns (accept separate or shared)
            columns = _coerce_list(params.get("columns", []))
            key_col_left = params.get("key_column_left") or (columns[0] if columns else "")
            key_col_right = params.get("key_column_right") or (columns[1] if len(columns) > 1 else key_col_left)
            if not key_col_left:
                return ToolOutput.fail(
                    tool_name=op_name,
                    errors=["compare_sheets requires 'key_column_left' or 'columns' parameter."],
                )
            return tool_module.compare_sheets(
                left, right,
                key_column_left=key_col_left,
                key_column_right=key_col_right,
            )

        if op_name == OperationType.explain_match_result.value:
            # Prefer the last match result from context; fall back to params
            match_data: dict = {}
            for match_op in (
                OperationType.run_hybrid_match.value,
                OperationType.run_fuzzy_match.value,
                OperationType.run_semantic_match.value,
                OperationType.run_exact_match.value,
            ):
                prior = context.get_data(match_op)
                if prior is not None:
                    match_data = prior if isinstance(prior, dict) else {}
                    break
            if not match_data:
                match_data = params.get("match_result", {})
            left_name = params.get("left_sheet", "left")
            right_name = params.get("right_sheet", "right")
            return tool_module.explain_match_result(
                match_data=match_data,
                left_name=left_name,
                right_name=right_name,
            )

        # Unknown operation
        return ToolOutput.fail(
            tool_name=op_name,
            errors=[f"Unknown operation type: '{op_name}'. Not registered in orchestrator."],
        )

    # ── Input resolvers ───────────────────────────────────────────────────────

    def _resolve_profile_inputs(
        self,
        params: dict,
        op_name: str,
    ) -> ToolOutput:
        """
        Resolve inputs for profile_columns.

        Reads sheet_name and the columns list from *params*.
        If columns is absent, profiles all columns in the sheet.
        """
        sheet_name = params.get("sheet_name", "")
        sheet = self._require_sheet(sheet_name, op_name)
        if isinstance(sheet, ToolOutput):
            return sheet

        columns = _coerce_list(
            params.get("columns")
            or params.get("candidate_columns")
            or params.get("key_columns")
            or []
        )
        if not columns:
            # Default: profile every column
            columns = list(sheet.header_row)

        return tool_module.profile_columns(sheet, columns=columns)

    def _resolve_match_inputs(
        self,
        op_name: str,
        params: dict,
        context: ExecutionContext,
        tool_fn: Callable[..., ToolOutput],
    ) -> ToolOutput:
        """
        Resolve inputs for any match or estimate_matchability tool.

        Column lists are taken from params.  If estimate_matchability was
        already run (stored in context), its recommended_strategy may influence
        which columns/config to pass.
        """
        left_name = params.get("left_sheet", "")
        right_name = params.get("right_sheet", "")
        left = self._require_sheet(left_name, op_name)
        right = self._require_sheet(right_name, op_name)
        if isinstance(left, ToolOutput):
            return left
        if isinstance(right, ToolOutput):
            return right

        # Resolve column lists from multiple possible key names
        left_columns = _coerce_list(
            params.get("left_columns")
            or params.get("candidate_columns")
            or []
        )
        right_columns = _coerce_list(
            params.get("right_columns")
            or params.get("candidate_columns")
            or left_columns  # fall back to same column names if symmetric
        )

        if not left_columns:
            return ToolOutput.fail(
                tool_name=op_name,
                errors=["No left_columns specified. Provide 'left_columns' in parameters."],
            )
        if not right_columns:
            return ToolOutput.fail(
                tool_name=op_name,
                errors=["No right_columns specified. Provide 'right_columns' in parameters."],
            )

        # Pull profile data from context if available (produced by a prior profile_columns step)
        left_profiles: list[dict] = []
        right_profiles: list[dict] = []
        prior_profiles = context.get_data(OperationType.profile_columns.value)
        if isinstance(prior_profiles, list):
            # profile_columns returns profiles for ONE sheet; we use them for both sides if symmetric
            left_profiles = prior_profiles
            right_profiles = prior_profiles

        # estimate_matchability uses list-based signatures
        if op_name == OperationType.estimate_matchability.value:
            return tool_fn(
                left_profiles=left_profiles,
                right_profiles=right_profiles,
                left_columns=left_columns,
                right_columns=right_columns,
            )

        # run_exact_match / run_fuzzy_match take a single key string each
        left_key = left_columns[0] if left_columns else ""
        right_key = right_columns[0] if right_columns else ""

        if op_name == OperationType.run_exact_match.value:
            return tool_fn(
                left=left,
                right=right,
                left_key=left_key,
                right_key=right_key,
            )

        if op_name == OperationType.run_fuzzy_match.value:
            # Tool threshold is 0-100; accept either scale
            raw_thresh = float(params.get("fuzzy_threshold", 0.82))
            threshold = raw_thresh * 100 if raw_thresh <= 1.0 else raw_thresh
            return tool_fn(
                left=left,
                right=right,
                left_key=left_key,
                right_key=right_key,
                threshold=threshold,
            )

        if op_name == OperationType.run_semantic_match.value:
            return tool_fn(
                left=left,
                right=right,
                left_key=left_key,
                right_key=right_key,
            )

        if op_name == OperationType.run_hybrid_match.value:
            # Hybrid: use first left/right pair as exact keys, second pair (or same) as fuzzy keys
            fuzzy_left_key = left_columns[1] if len(left_columns) > 1 else left_key
            fuzzy_right_key = right_columns[1] if len(right_columns) > 1 else right_key
            raw_thresh = float(params.get("final_threshold", 0.75))
            threshold = raw_thresh * 100 if raw_thresh <= 1.0 else raw_thresh
            return tool_fn(
                left=left,
                right=right,
                exact_key_left=left_key,
                exact_key_right=right_key,
                fuzzy_key_left=fuzzy_left_key,
                fuzzy_key_right=fuzzy_right_key,
                threshold=threshold,
            )

        return ToolOutput.fail(
            tool_name=op_name,
            errors=[f"Unhandled match operation: {op_name}"],
        )

    # ── Registry builder ──────────────────────────────────────────────────────

    def _build_tool_registry(self) -> dict[str, Callable[..., ToolOutput]]:
        """
        Return a mapping of OperationType string values to tool callables.

        Not all tools are called via the registry (some have bespoke dispatch
        logic above), but the registry provides a convenient lookup for match
        tools and future extensions.
        """
        return {
            OperationType.list_sheets.value: tool_module.list_sheets,
            OperationType.get_sheet_schema.value: tool_module.get_sheet_schema,
            OperationType.preview_sheet.value: tool_module.preview_sheet,
            OperationType.profile_columns.value: tool_module.profile_columns,
            OperationType.clean_columns.value: tool_module.clean_columns,
            OperationType.estimate_matchability.value: tool_module.estimate_matchability,
            OperationType.run_exact_match.value: tool_module.run_exact_match,
            OperationType.run_fuzzy_match.value: tool_module.run_fuzzy_match,
            OperationType.run_semantic_match.value: tool_module.run_semantic_match,
            OperationType.run_hybrid_match.value: tool_module.run_hybrid_match,
            OperationType.find_duplicates.value: tool_module.find_duplicates,
            OperationType.aggregate_values.value: tool_module.aggregate_values,
            OperationType.filter_rows.value: tool_module.filter_rows,
            OperationType.compare_sheets.value: tool_module.compare_sheets,
            OperationType.explain_match_result.value: tool_module.explain_match_result,
        }

    # ── Internal helpers ──────────────────────────────────────────────────────

    def _require_sheet(self, sheet_name: str, op_name: str) -> SheetData | ToolOutput:
        """
        Resolve *sheet_name* to a SheetData or return a failure ToolOutput.
        """
        if not sheet_name:
            return ToolOutput.fail(
                tool_name=op_name,
                errors=[f"Operation '{op_name}' requires 'sheet_name' in parameters."],
            )
        sheet = self.sheets.get(sheet_name)
        if sheet is None:
            return ToolOutput.fail(
                tool_name=op_name,
                errors=[
                    f"Sheet '{sheet_name}' not found. "
                    f"Available: {list(self.sheets.keys())}"
                ],
            )
        return sheet


# ── Module-level helpers ──────────────────────────────────────────────────────


def _coerce_list(value: Any) -> list[str]:
    """
    Normalise *value* to a list of strings.

    Handles:
    - None / empty → []
    - list → each element cast to str
    - str → [str]
    - dict with "left"/"right" keys → flattened list (useful when params encode
      both sides together)
    """
    if value is None:
        return []
    if isinstance(value, list):
        return [str(v) for v in value]
    if isinstance(value, str):
        return [value] if value else []
    if isinstance(value, dict):
        result: list[str] = []
        for v in value.values():
            if isinstance(v, list):
                result.extend(str(i) for i in v)
        return result
    return []


def _build_hybrid_config_from_estimate(
    estimate_data: dict,
    left_columns: list[str],
    right_columns: list[str],
) -> dict:
    """
    Build a hybrid_config dict from the output of estimate_matchability.

    Uses the per-column strategy recommendations to assign each column pair
    a strategy.  Weights are proportional to the column scores.
    """
    col_strategies: dict[str, str] = estimate_data.get("column_strategies", {})
    col_scores: dict[str, float] = estimate_data.get("column_scores", {})

    total_score = sum(col_scores.values()) or 1.0

    columns_config = []
    for lc, rc in zip(left_columns, right_columns):
        strategy = col_strategies.get(lc, StrategyType.fuzzy.value)
        weight = col_scores.get(lc, 1.0) / total_score
        columns_config.append({
            "left": lc,
            "right": rc,
            "strategy": strategy,
            "weight": round(weight, 4),
        })

    return {"columns": columns_config}
