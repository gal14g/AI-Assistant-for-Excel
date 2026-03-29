"""
ExplanationService: Layer 4 — generates human-readable summaries of
analytical pipeline results.

Explains what was done, why that strategy was chosen, key numbers,
warnings, and what to do next.
"""
from __future__ import annotations

from typing import Any

from ..models.analytical_plan import AnalyticalPlan, IntentType
from ..orchestrator.execution_context import ExecutionContext


class ExplanationService:
    """Generates the final natural-language response for an analytical run."""

    def generate(
        self,
        plan: AnalyticalPlan,
        context: ExecutionContext,
        final_data: dict[str, Any] | None = None,
    ) -> str:
        """
        Build the complete explanation string shown to the user.
        """
        parts: list[str] = []

        # 1. What was done
        parts.append(self._intent_summary(plan))

        # 2. Why this strategy
        if plan.reasoning_summary:
            parts.append(self._strategy_explanation(
                plan.parameters.get("strategy", ""),
                plan.reasoning_summary,
            ))

        # 3. Key numbers from the result
        data = final_data or {}
        intent = plan.intent.value if hasattr(plan.intent, "value") else str(plan.intent)

        if intent == IntentType.match_rows.value:
            parts.append(self._format_match_result(data))
        elif intent in (IntentType.find_duplicates.value,):
            parts.append(self._format_duplicate_result(data))
        elif intent in (IntentType.aggregate.value, IntentType.group_and_summarize.value):
            parts.append(self._format_aggregation_result(data))
        elif intent == IntentType.profile_sheet.value:
            raw_profiles = data if isinstance(data, list) else data.get("profiles", [])
            if raw_profiles:
                parts.append(self._format_profile_result(raw_profiles))
        elif intent == IntentType.compare_sheets.value:
            parts.append(self._format_compare_result(data))
        else:
            # Generic: surface any "summary" or "rows" key
            if "summary" in data:
                parts.append(str(data["summary"]))
            elif "rows" in data:
                parts.append(f"Returned {len(data['rows'])} rows.")

        # 4. Warnings
        all_warnings = context.all_warnings()
        if all_warnings:
            parts.append(
                "**Note:** " + "; ".join(all_warnings[:3])
                + (" (and more)" if len(all_warnings) > 3 else "")
            )

        # 5. Confidence caveat
        if plan.confidence < 0.7:
            parts.append(
                f"Confidence: {plan.confidence:.0%}. "
                "I may have misunderstood part of your request — please review the results."
            )

        return "\n\n".join(p for p in parts if p.strip())

    # ── Intent summaries ──────────────────────────────────────────────────────

    def _intent_summary(self, plan: AnalyticalPlan) -> str:
        intent = plan.intent.value if hasattr(plan.intent, "value") else str(plan.intent)
        params = plan.parameters

        mapping = {
            IntentType.match_rows.value: (
                lambda: f"Matched rows between "
                        f"**{params.get('left_sheet', 'Sheet1')}** and "
                        f"**{params.get('right_sheet', 'Sheet2')}**."
            ),
            IntentType.find_duplicates.value: (
                lambda: f"Found duplicate rows in "
                        f"**{params.get('sheet_name', params.get('primary_sheet', 'the sheet'))}**."
            ),
            IntentType.aggregate.value: (
                lambda: f"Aggregated data in "
                        f"**{params.get('sheet_name', params.get('primary_sheet', 'the sheet'))}** "
                        f"grouped by {params.get('group_by', [])}."
            ),
            IntentType.group_and_summarize.value: (
                lambda: f"Grouped and summarised data in "
                        f"**{params.get('sheet_name', params.get('primary_sheet', 'the sheet'))}**."
            ),
            IntentType.compare_sheets.value: (
                lambda: f"Compared **{params.get('left_sheet', 'Sheet1')}** and "
                        f"**{params.get('right_sheet', 'Sheet2')}**."
            ),
            IntentType.clean_data.value: (
                lambda: f"Cleaned columns in "
                        f"**{params.get('sheet_name', params.get('primary_sheet', 'the sheet'))}**."
            ),
            IntentType.filter_rows.value: (
                lambda: f"Filtered rows in "
                        f"**{params.get('sheet_name', params.get('primary_sheet', 'the sheet'))}**."
            ),
            IntentType.profile_sheet.value: (
                lambda: f"Profiled columns in "
                        f"**{params.get('sheet_name', params.get('primary_sheet', 'the sheet'))}**."
            ),
        }
        fn = mapping.get(intent)
        return fn() if fn else f"Completed {intent.replace('_', ' ')} operation."

    # ── Result formatters ─────────────────────────────────────────────────────

    def _format_match_result(self, data: dict) -> str:
        if not data:
            return ""
        mc = data.get("match_count", 0)
        ul = data.get("unmatched_left", data.get("unmatched_left_count", 0))
        ur = data.get("unmatched_right", data.get("unmatched_right_count", 0))
        strategy = data.get("strategy_used", data.get("strategy", "unknown"))
        score_avg = data.get("average_score", None)

        lines = [f"**{mc} rows matched** using {strategy} strategy."]
        if ul:
            lines.append(f"{ul} rows in the left sheet had no match.")
        if ur:
            lines.append(f"{ur} rows in the right sheet had no match.")
        if score_avg is not None:
            lines.append(f"Average match score: {score_avg:.1%}.")
        total = mc + ul
        if total:
            pct = mc / total * 100
            quality = (
                "Excellent" if pct >= 90 else
                "Good" if pct >= 70 else
                "Fair" if pct >= 50 else "Poor"
            )
            lines.append(f"Match quality: **{quality}** ({pct:.0f}% of left rows matched).")
        return "\n".join(lines)

    def _format_duplicate_result(self, data: dict) -> str:
        if not data:
            return ""
        dup_count = data.get("duplicate_count", 0)
        unique_count = data.get("unique_count", 0)
        total = data.get("total_rows", dup_count + unique_count)
        lines = [f"**{dup_count} duplicate rows** found out of {total} total rows."]
        if unique_count:
            lines.append(f"{unique_count} rows are unique.")
        groups = data.get("group_count", data.get("duplicate_groups_count", None))
        if groups:
            lines.append(f"Organised into {groups} duplicate groups.")
        return "\n".join(lines)

    def _format_aggregation_result(self, data: dict) -> str:
        if not data:
            return ""
        groups = data.get("group_count", data.get("groups", []))
        if isinstance(groups, list):
            count = len(groups)
        else:
            count = int(groups)
        total = data.get("total_rows", "?")
        return f"Aggregated **{total} rows** into **{count} groups**."

    def _format_profile_result(self, profiles: list) -> str:
        if not profiles:
            return ""
        lines = [f"Profiled **{len(profiles)} column(s)**:"]
        for p in profiles[:8]:
            name = p.get("name", "?")
            dtype = p.get("dtype", "?")
            ur = float(p.get("uniqueness_ratio", 0))
            nr = float(p.get("null_rate", 0))
            warns = p.get("warnings", [])
            w_str = f" ⚠ {warns[0]}" if warns else ""
            lines.append(
                f"  • **{name}** — {dtype}, {ur:.0%} unique, {nr:.0%} null{w_str}"
            )
        return "\n".join(lines)

    def _format_compare_result(self, data: dict) -> str:
        if not data:
            return ""
        ol = data.get("only_left_count", 0)
        orr = data.get("only_right_count", 0)
        both = data.get("in_both_count", 0)
        lines = []
        if both:
            lines.append(f"**{both} keys** appear in both sheets.")
        if ol:
            lines.append(f"**{ol} keys** exist only in the left sheet.")
        if orr:
            lines.append(f"**{orr} keys** exist only in the right sheet.")
        return "\n".join(lines) if lines else "Comparison complete."

    def _strategy_explanation(self, strategy: str, reasoning: str) -> str:
        strategy_labels = {
            "exact": "exact string matching",
            "fuzzy": "fuzzy (approximate) string matching",
            "semantic": "semantic similarity (embedding-based)",
            "hybrid": "hybrid matching (exact + fuzzy, weighted)",
        }
        label = strategy_labels.get(strategy, strategy) if strategy else None
        if label:
            return f"Strategy: **{label}**. {reasoning}"
        return reasoning
