"""
Matching service: column sufficiency evaluation and strategy selection.

These are pure Python decision functions — no I/O, no LLM calls.
"""
from __future__ import annotations

from ..models.analytical_plan import ColumnProfile, StrategyType


# ── Column sufficiency ────────────────────────────────────────────────────────


class ColumnSufficiencyEvaluator:
    """
    Evaluates whether a set of column profiles provides sufficient signal
    for reliable row matching or deduplication.
    """

    def evaluate(self, profiles: list[dict]) -> dict:
        """
        Return a sufficiency assessment dict:

        Keys
        ----
        sufficient : bool
        score : float  (0–1 composite score)
        weak_columns : list[str]  (high null or very low uniqueness)
        strong_columns : list[str]  (high uniqueness, low null rate)
        recommendation : str  (human-readable summary)
        """
        if not profiles:
            return {
                "sufficient": False,
                "score": 0.0,
                "weak_columns": [],
                "strong_columns": [],
                "recommendation": "No columns provided for analysis.",
            }

        weak: list[str] = []
        strong: list[str] = []
        scores: list[float] = []

        for p in profiles:
            name = p.get("name", "?")
            null_rate = float(p.get("null_rate", 0))
            ur = float(p.get("uniqueness_ratio", 0))

            col_score = ur * (1 - null_rate)

            if null_rate > 0.3 or ur < 0.15:
                weak.append(name)
            if ur > 0.7 and null_rate < 0.15:
                strong.append(name)

            scores.append(col_score)

        composite = sum(scores) / len(scores) if scores else 0.0

        # Sufficient if:
        #   - any column is strong (high uniqueness + low nulls), OR
        #   - combined score of 2+ columns pushes composite above 0.45
        any_strong = bool(strong)
        combined_ok = len(scores) >= 2 and composite >= 0.45
        sufficient = any_strong or combined_ok

        if sufficient and strong:
            recommendation = (
                f"Strong columns for matching: {strong}. "
                "Proceed with the current column selection."
            )
        elif sufficient:
            recommendation = (
                "Column combination provides adequate signal for matching, "
                "though individual columns are moderately distinctive."
            )
        elif weak and not strong:
            recommendation = (
                f"Weak columns detected: {weak}. "
                "Consider adding a more distinctive column (e.g. an ID or email) "
                "to improve match accuracy."
            )
        else:
            recommendation = (
                "Column sufficiency is borderline. "
                "Adding one more high-uniqueness column would improve results."
            )

        return {
            "sufficient": sufficient,
            "score": round(composite, 4),
            "weak_columns": weak,
            "strong_columns": strong,
            "recommendation": recommendation,
        }


# ── Strategy selection ────────────────────────────────────────────────────────


def select_strategy(
    profiles: list[dict],
    override: str | None = None,
) -> StrategyType:
    """
    Choose the best matching strategy based on column profiles.

    Rules (applied in order):
    1. If *override* is a valid StrategyType, use it.
    2. If all columns are ID-like (uniqueness > 0.9 or dtype == 'id') → exact.
    3. If any column is long-text (avg_text_length > 50) → hybrid or semantic.
    4. If any column has moderate uniqueness (text + ur < 0.9) → fuzzy.
    5. Default → hybrid.
    """
    if override:
        try:
            return StrategyType(override)
        except ValueError:
            pass

    if not profiles:
        return StrategyType.hybrid

    has_text_long = any(
        float(p.get("avg_text_length") or 0) > 50 for p in profiles
    )
    all_id_like = all(
        float(p.get("uniqueness_ratio", 0)) > 0.9 or p.get("dtype") == "id"
        for p in profiles
    )
    any_text = any(p.get("dtype") in ("text", "mixed") for p in profiles)

    if all_id_like:
        return StrategyType.exact
    if has_text_long:
        # Pure semantic if ALL columns are long-text; hybrid if mixed
        all_long_text = all(float(p.get("avg_text_length") or 0) > 50 for p in profiles)
        return StrategyType.semantic if all_long_text else StrategyType.hybrid
    if any_text:
        return StrategyType.fuzzy
    return StrategyType.hybrid


# ── Hybrid config builder ─────────────────────────────────────────────────────


def build_hybrid_config(
    left_profiles: list[dict],
    right_profiles: list[dict],
    left_columns: list[str],
    right_columns: list[str],
) -> dict:
    """
    Build a hybrid_config dict for run_hybrid_match based on column profiles.

    Assigns each column pair a strategy and weight proportional to its
    matchability score (uniqueness × (1 - null_rate)).
    """
    profile_map_left = {p["name"]: p for p in left_profiles if "name" in p}

    columns_config: list[dict] = []
    total_score = 0.0
    raw: list[tuple[str, str, str, float]] = []

    for lc, rc in zip(left_columns, right_columns):
        p = profile_map_left.get(lc, {})
        ur = float(p.get("uniqueness_ratio", 0.5))
        nr = float(p.get("null_rate", 0))
        avg_len = float(p.get("avg_text_length") or 0)
        dtype = p.get("dtype", "text")

        score = ur * (1 - nr)
        total_score += score

        if ur > 0.9 or dtype == "id":
            strategy = "exact"
        elif avg_len > 50:
            strategy = "semantic"
        else:
            strategy = "fuzzy"

        raw.append((lc, rc, strategy, score))

    denom = total_score or 1.0
    for lc, rc, strategy, score in raw:
        columns_config.append({
            "left": lc,
            "right": rc,
            "strategy": strategy,
            "weight": round(score / denom, 4),
        })

    return {"columns": columns_config}
