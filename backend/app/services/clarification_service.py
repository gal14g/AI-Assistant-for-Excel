"""
Clarification policy: determines when the assistant should ask the user
for more information before proceeding with an analytical pipeline.

The golden rule: ask ONLY when truly necessary — never ask for information
that can be derived from the data or the user's message.
"""
from __future__ import annotations

from ..models.analytical_plan import AnalyticalPlan


class ClarificationPolicy:
    """
    Evaluates whether clarification is needed and generates a focused question.
    """

    def should_ask(
        self,
        plan: AnalyticalPlan,
        column_sufficiency: dict,
        user_message: str = "",
    ) -> bool:
        """
        Return True only if asking for clarification is warranted.

        Triggered when:
        - The planner already flagged needs_clarification=True, OR
        - Column sufficiency score is very low (< 0.4) AND the user did not
          explicitly name columns in their message.

        Never triggered if:
        - The user said "match by <column>" — they already told us.
        - The columns are sufficient (score >= 0.4) even if not perfect.
        """
        # Planner already decided clarification is needed
        if plan.needs_clarification:
            return True

        score = float(column_sufficiency.get("score", 1.0))
        if score >= 0.4:
            return False

        # Low score — but only ask if the user didn't already specify columns
        user_lower = user_message.lower()
        explicit_column_keywords = (
            "match by", "using column", "use column", "based on",
            "by column", "on column", "match on", "key column",
        )
        user_specified_columns = any(kw in user_lower for kw in explicit_column_keywords)
        return not user_specified_columns

    def generate_question(
        self,
        plan: AnalyticalPlan,
        column_sufficiency: dict,
        available_columns: list[str],
    ) -> str:
        """
        Generate a specific, helpful clarification question.

        Uses the plan's clarification_question if already set; otherwise
        constructs one from the sufficiency analysis.
        """
        if plan.clarification_question:
            return plan.clarification_question

        weak = column_sufficiency.get("weak_columns", [])
        recommendation = column_sufficiency.get("recommendation", "")

        if weak and available_columns:
            other_cols = [c for c in available_columns if c not in weak]
            suggestion = f" (e.g. {other_cols[0]!r})" if other_cols else ""
            return (
                f"The columns {weak} have low uniqueness or many missing values, "
                f"which may lead to unreliable matches. "
                f"Could you specify an additional distinguishing column{suggestion}?"
            )

        if recommendation:
            return recommendation

        if available_columns:
            cols_str = ", ".join(f"'{c}'" for c in available_columns[:6])
            return (
                f"Which column(s) should I use for matching? "
                f"Available columns: {cols_str}."
            )

        return (
            "I wasn't sure which columns to use for this operation. "
            "Could you specify the column names you'd like to match on?"
        )
