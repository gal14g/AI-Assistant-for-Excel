"""
Matching pipeline tools for Excel Copilot.

All functions are synchronous, deterministic, and side-effect-free.
They accept SheetData / ColumnProfile inputs and return ToolOutput.

Supported matching strategies:
  - exact      : pandas merge (case-insensitive, multi-column)
  - fuzzy      : rapidfuzz WRatio / token_sort_ratio
  - semantic   : cosine similarity over pluggable EmbeddingProvider
  - hybrid     : weighted combination of exact + fuzzy + semantic

Public API
----------
estimate_matchability(left_profiles, right_profiles, left_columns, right_columns)
run_exact_match(left_sheet, right_sheet, left_columns, right_columns)
run_fuzzy_match(left_sheet, right_sheet, left_columns, right_columns, threshold, column_weights)
run_semantic_match(left_sheet, right_sheet, left_columns, right_columns, top_k, embedding_provider)
run_hybrid_match(left_sheet, right_sheet, config, embedding_provider)
"""
from __future__ import annotations

import logging
import math
from typing import Any, Optional, Protocol

import pandas as pd

from ..models.analytical_plan import (
    MatchabilityEstimate,
    MatchResult,
    SheetData,
    StrategyType,
)
from ..models.tool_output import ToolOutput

logger = logging.getLogger(__name__)

_TOOL_MATCHABILITY = "estimate_matchability"
_TOOL_EXACT = "run_exact_match"
_TOOL_FUZZY = "run_fuzzy_match"
_TOOL_HYBRID = "run_hybrid_match"
_TOOL_SEMANTIC = "run_semantic_match"


# ── EmbeddingProvider protocol ────────────────────────────────────────────────


class EmbeddingProvider(Protocol):
    """Minimal interface required by run_semantic_match / run_hybrid_match."""

    def encode(self, texts: list[str]) -> list[list[float]]:
        """Return one embedding vector per input text."""
        ...


# ── Internal helpers ──────────────────────────────────────────────────────────


def _cosine_similarity(a: list[float], b: list[float]) -> float:
    dot = sum(x * y for x, y in zip(a, b))
    norm_a = math.sqrt(sum(x * x for x in a))
    norm_b = math.sqrt(sum(x * x for x in b))
    if norm_a == 0.0 or norm_b == 0.0:
        return 0.0
    return dot / (norm_a * norm_b)


def _combine_columns(df: pd.DataFrame, columns: list[str]) -> list[str]:
    """Join column values for each row into a single string with ' | ' separator."""
    if len(columns) == 1:
        return df[columns[0]].fillna("").astype(str).tolist()
    parts = [df[c].fillna("").astype(str) for c in columns]
    combined = parts[0]
    for part in parts[1:]:
        combined = combined + " | " + part
    return combined.tolist()


def _profile_by_name(profiles: list[dict[str, Any]], name: str) -> dict[str, Any]:
    for p in profiles:
        if p.get("name") == name:
            return p
    return {}


# ---------------------------------------------------------------------------
# estimate_matchability
# ---------------------------------------------------------------------------

def estimate_matchability(
    left_profiles: list[dict[str, Any]],
    right_profiles: list[dict[str, Any]],
    left_columns: list[str],
    right_columns: list[str],
) -> ToolOutput:
    """
    Score how matchable each candidate column pair is between two sheets,
    using ColumnProfile data rather than raw SheetData.

    Returns a MatchabilityEstimate with:
      - overall_confidence: 0-1
      - recommended_strategy: exact | fuzzy | hybrid | semantic
      - column_scores: {col: score}
      - column_strategies: {col: strategy}
    """
    if len(left_columns) != len(right_columns):
        return ToolOutput.fail(
            tool_name=_TOOL_MATCHABILITY,
            errors=["left_columns and right_columns must have the same length."],
        )

    warnings: list[str] = []
    column_scores: dict[str, float] = {}
    column_strategies: dict[str, StrategyType] = {}

    # Build quick lookup: name → profile dict
    left_profile_map = {p.get("name", ""): p for p in left_profiles}
    right_profile_map = {p.get("name", ""): p for p in right_profiles}

    resolvable: list[tuple[str, str]] = []
    for lc, rc in zip(left_columns, right_columns):
        if lc not in left_profile_map:
            warnings.append(f"Left column '{lc}' has no profile — skipped.")
            continue
        if rc not in right_profile_map:
            warnings.append(f"Right column '{rc}' has no profile — skipped.")
            continue
        resolvable.append((lc, rc))

    if not resolvable:
        return ToolOutput.fail(
            tool_name=_TOOL_MATCHABILITY,
            errors=["No resolvable column pairs found."],
            warnings=warnings,
        )

    for left_col, right_col in resolvable:
        lp = left_profile_map[left_col]
        rp = right_profile_map[right_col]

        # Use profile attributes to estimate matchability
        l_unique = lp.get("uniqueness_ratio", 0.0)
        r_unique = rp.get("uniqueness_ratio", 0.0)
        avg_unique = (l_unique + r_unique) / 2.0

        l_null = lp.get("null_rate", 0.0)
        r_null = rp.get("null_rate", 0.0)
        avg_null = (l_null + r_null) / 2.0

        l_dtype = lp.get("dtype", "unknown")
        r_dtype = rp.get("dtype", "unknown")

        # Heuristic: same dtype → higher overlap estimate
        dtype_match = 0.6 if l_dtype == r_dtype else 0.3

        score = avg_unique * 0.4 + dtype_match * 0.4 + (1.0 - avg_null) * 0.2
        score = max(0.0, min(1.0, score))
        column_scores[left_col] = round(score, 4)

        # Recommend strategy
        if l_dtype in ("id", "numeric", "date") and r_dtype == l_dtype and score > 0.5:
            column_strategies[left_col] = StrategyType.exact
        elif score >= 0.4:
            column_strategies[left_col] = StrategyType.fuzzy
        elif score >= 0.2:
            column_strategies[left_col] = StrategyType.hybrid
        else:
            column_strategies[left_col] = StrategyType.semantic

        if avg_null > 0.5:
            warnings.append(
                f"Column '{left_col}' has high null rate ({avg_null:.0%}) — match reliability may be low."
            )

    if column_scores:
        overall = sum(column_scores.values()) / len(column_scores)
    else:
        overall = 0.0

    # Recommend overall strategy by majority vote
    strategy_votes: dict[StrategyType, int] = {}
    for s in column_strategies.values():
        strategy_votes[s] = strategy_votes.get(s, 0) + 1
    recommended = max(strategy_votes, key=strategy_votes.get) if strategy_votes else StrategyType.exact  # noqa: E501

    estimate = MatchabilityEstimate(
        overall_confidence=round(overall, 4),
        recommended_strategy=recommended,
        column_scores=column_scores,
        column_strategies=column_strategies,
        warnings=warnings,
        needs_more_columns=overall < 0.4,
        suggested_additional_columns=[],
    )

    return ToolOutput.ok(
        tool_name=_TOOL_MATCHABILITY,
        data=estimate.model_dump(),
        warnings=warnings,
    )


# ── _estimate_matchability_from_sheets (internal helper) ─────────────────────
# Used by other matching tools when only SheetData is available.

def _estimate_from_sheets(
    left_df: pd.DataFrame,
    right_df: pd.DataFrame,
    left_columns: list[str],
    right_columns: list[str],
    warnings: list[str],
) -> tuple[dict[str, float], dict[str, StrategyType]]:
    """
    Compute per-column matchability scores and strategies from raw DataFrames.

    Returns (column_scores, column_strategies) — both keyed by left column name.
    Appends any diagnostic messages to the caller-supplied *warnings* list.
    """
    column_scores: dict[str, float] = {}
    column_strategies: dict[str, StrategyType] = {}

    for left_col, right_col in zip(left_columns, right_columns):
        if left_col not in left_df.columns:
            warnings.append(f"Left column '{left_col}' not found — skipped.")
            continue
        if right_col not in right_df.columns:
            warnings.append(f"Right column '{right_col}' not found — skipped.")
            continue

        lseries = left_df[left_col].dropna().astype(str)
        rseries = right_df[right_col].dropna().astype(str)

        if lseries.empty or rseries.empty:
            column_scores[left_col] = 0.0
            column_strategies[left_col] = StrategyType.exact
            warnings.append(
                f"Column '{left_col}' has all-null values in one sheet — score set to 0."
            )
            continue

        # Uniqueness in each sheet
        l_unique = lseries.nunique() / len(lseries) if len(lseries) > 0 else 0.0
        r_unique = rseries.nunique() / len(rseries) if len(rseries) > 0 else 0.0
        avg_unique = (l_unique + r_unique) / 2.0

        # Overlap ratio — how many left values appear in right
        right_set = set(rseries.str.lower().str.strip())
        left_vals = lseries.str.lower().str.strip()
        exact_overlap = float(left_vals.isin(right_set).mean())

        # Null rate
        l_null = 1.0 - len(lseries) / max(len(left_df), 1)
        r_null = 1.0 - len(rseries) / max(len(right_df), 1)
        avg_null = (l_null + r_null) / 2.0

        # Composite score
        score = avg_unique * 0.4 + exact_overlap * 0.4 + (1.0 - avg_null) * 0.2
        score = max(0.0, min(1.0, score))
        column_scores[left_col] = round(score, 4)

        # Recommend strategy based on uniqueness + overlap
        if avg_unique > 0.9 and exact_overlap > 0.5:
            column_strategies[left_col] = StrategyType.exact
        elif exact_overlap < 0.3 and avg_unique > 0.5:
            column_strategies[left_col] = StrategyType.fuzzy
        elif exact_overlap > 0.3:
            column_strategies[left_col] = StrategyType.hybrid
        else:
            column_strategies[left_col] = StrategyType.semantic

        if avg_null > 0.5:
            warnings.append(
                f"Column '{left_col}' has high null rate ({avg_null:.0%}) — match reliability may be low."
            )

    return column_scores, column_strategies


# ---------------------------------------------------------------------------
# run_exact_match
# ---------------------------------------------------------------------------

def run_exact_match(
    left: SheetData,
    right: SheetData,
    left_key: str,
    right_key: str,
    return_columns: list[str] | None = None,
) -> ToolOutput:
    """
    Perform an exact (case-insensitive, strip-normalised) match between two
    sheets on the specified key columns.

    Returns a ToolOutput whose .data contains:
      - matched: list of dicts with left row + matched right columns
      - unmatched_left: list of left-row dicts with no match
      - unmatched_right: list of right-row dicts with no match
      - match_count, unmatched_left_count, unmatched_right_count
    """
    try:
        import pandas as pd
    except ImportError:
        return ToolOutput.fail(
            tool_name=_TOOL_EXACT,
            errors=["pandas is required for run_exact_match."],
        )

    left_df = left.to_dataframe()
    right_df = right.to_dataframe()

    if left_key not in left_df.columns:
        return ToolOutput.fail(
            tool_name=_TOOL_EXACT,
            errors=[f"Key column '{left_key}' not found in sheet '{left.name}'."],
        )
    if right_key not in right_df.columns:
        return ToolOutput.fail(
            tool_name=_TOOL_EXACT,
            errors=[f"Key column '{right_key}' not found in sheet '{right.name}'."],
        )

    # Normalise keys for comparison
    left_df = left_df.copy()
    right_df = right_df.copy()
    left_df["_norm_key"] = left_df[left_key].astype(str).str.lower().str.strip()
    right_df["_norm_key"] = right_df[right_key].astype(str).str.lower().str.strip()

    merged = left_df.merge(
        right_df,
        on="_norm_key",
        how="left",
        suffixes=("_left", "_right"),
    )

    # Determine which right columns to return
    if return_columns:
        right_cols_to_keep = [
            c + "_right" if c + "_right" in merged.columns else c
            for c in return_columns
            if c in right_df.columns or c + "_right" in merged.columns
        ]
    else:
        right_cols_to_keep = [
            c for c in merged.columns
            if c.endswith("_right") or (c in right_df.columns and c not in left_df.columns and c != "_norm_key")
        ]

    # Build result sets
    matched_mask = merged["_norm_key"].isin(right_df["_norm_key"])
    matched = merged[matched_mask].drop(columns=["_norm_key"]).where(
        pd.notnull(merged[matched_mask]), other=None
    ).to_dict(orient="records")

    unmatched_left = left_df[~left_df["_norm_key"].isin(right_df["_norm_key"])].drop(
        columns=["_norm_key"]
    ).where(pd.notnull(left_df[~left_df["_norm_key"].isin(right_df["_norm_key"])]), other=None).to_dict(
        orient="records"
    )

    matched_right_keys = set(
        merged[matched_mask]["_norm_key"].tolist()
    )
    unmatched_right = right_df[~right_df["_norm_key"].isin(matched_right_keys)].drop(
        columns=["_norm_key"]
    ).where(pd.notnull(right_df[~right_df["_norm_key"].isin(matched_right_keys)]), other=None).to_dict(
        orient="records"
    )

    match_result = MatchResult(
        left_indices=list(range(len(matched))),
        right_indices=list(range(len(matched))),
        scores=[1.0] * len(matched),
        match_count=len(matched),
        unmatched_left=len(unmatched_left),
        unmatched_right=len(unmatched_right),
        strategy_used=StrategyType.exact,
    )

    return ToolOutput.ok(
        tool_name=_TOOL_EXACT,
        data={
            "matched": matched,
            "unmatched_left": unmatched_left,
            "unmatched_right": unmatched_right,
            "match_count": len(matched),
            "unmatched_left_count": len(unmatched_left),
            "unmatched_right_count": len(unmatched_right),
            "match_result": match_result.model_dump(),
        },
        metadata={"strategy": "exact", "left_key": left_key, "right_key": right_key},
    )


# ---------------------------------------------------------------------------
# run_fuzzy_match
# ---------------------------------------------------------------------------

def run_fuzzy_match(
    left: SheetData,
    right: SheetData,
    left_key: str,
    right_key: str,
    threshold: float = 80.0,
    return_columns: list[str] | None = None,
) -> ToolOutput:
    """
    Perform a fuzzy token-sort-ratio match using rapidfuzz.

    threshold: 0-100, minimum score to count as a match (default 80).
    """
    try:
        from rapidfuzz import process as rf_process, fuzz as rf_fuzz
    except ImportError:
        return ToolOutput.fail(
            tool_name=_TOOL_FUZZY,
            errors=[
                "rapidfuzz is not installed. "
                "Install it with: pip install rapidfuzz"
            ],
        )

    try:
        import pandas as pd
    except ImportError:
        return ToolOutput.fail(
            tool_name=_TOOL_FUZZY,
            errors=["pandas is required for run_fuzzy_match."],
        )

    left_df = left.to_dataframe()
    right_df = right.to_dataframe()

    if left_key not in left_df.columns:
        return ToolOutput.fail(
            tool_name=_TOOL_FUZZY,
            errors=[f"Key column '{left_key}' not found in sheet '{left.name}'."],
        )
    if right_key not in right_df.columns:
        return ToolOutput.fail(
            tool_name=_TOOL_FUZZY,
            errors=[f"Key column '{right_key}' not found in sheet '{right.name}'."],
        )

    right_keys = right_df[right_key].astype(str).str.strip().tolist()
    right_keys_norm = [k.lower() for k in right_keys]

    matched: list[dict[str, Any]] = []
    unmatched_left: list[dict[str, Any]] = []
    matched_right_indices: set[int] = set()
    scores: list[float] = []

    for _, left_row in left_df.iterrows():
        query = str(left_row[left_key]).strip()
        if not query or query.lower() in {"nan", "none", ""}:
            unmatched_left.append(left_row.where(pd.notnull(left_row), other=None).to_dict())
            continue

        result = rf_process.extractOne(
            query.lower(),
            right_keys_norm,
            scorer=rf_fuzz.token_sort_ratio,
            score_cutoff=threshold,
        )

        if result is not None:
            best_val, best_score, best_idx = result
            right_row = right_df.iloc[best_idx]
            combined = {**left_row.where(pd.notnull(left_row), other=None).to_dict()}
            if return_columns:
                for rc in return_columns:
                    if rc in right_df.columns:
                        combined[f"{rc}_matched"] = right_row[rc]
            else:
                for rc in right_df.columns:
                    if rc != right_key:
                        combined[f"{rc}_matched"] = right_row.get(rc)
            combined["_match_score"] = round(best_score, 2)
            combined["_matched_value"] = right_row[right_key]
            matched.append(combined)
            matched_right_indices.add(best_idx)
            scores.append(best_score)
        else:
            unmatched_left.append(left_row.where(pd.notnull(left_row), other=None).to_dict())

    unmatched_right = [
        right_df.iloc[i].where(pd.notnull(right_df.iloc[i]), other=None).to_dict()
        for i in range(len(right_df))
        if i not in matched_right_indices
    ]

    match_result = MatchResult(
        left_indices=list(range(len(matched))),
        right_indices=list(matched_right_indices),
        scores=scores,
        match_count=len(matched),
        unmatched_left=len(unmatched_left),
        unmatched_right=len(unmatched_right),
        strategy_used=StrategyType.fuzzy,
    )

    return ToolOutput.ok(
        tool_name=_TOOL_FUZZY,
        data={
            "matched": matched,
            "unmatched_left": unmatched_left,
            "unmatched_right": unmatched_right,
            "match_count": len(matched),
            "unmatched_left_count": len(unmatched_left),
            "unmatched_right_count": len(unmatched_right),
            "match_result": match_result.model_dump(),
        },
        metadata={
            "strategy": "fuzzy",
            "threshold": threshold,
            "left_key": left_key,
            "right_key": right_key,
        },
    )


# ---------------------------------------------------------------------------
# run_hybrid_match
# ---------------------------------------------------------------------------

def run_hybrid_match(
    left: SheetData,
    right: SheetData,
    exact_key_left: str,
    exact_key_right: str,
    fuzzy_key_left: str,
    fuzzy_key_right: str,
    threshold: float = 75.0,
    return_columns: list[str] | None = None,
) -> ToolOutput:
    """
    Two-pass match: first attempt exact on the ID/key columns, then fuzzy
    on the name/text columns for any rows that did not match exactly.
    """
    # Pass 1: exact match
    exact_result = run_exact_match(left, right, exact_key_left, exact_key_right, return_columns)
    if not exact_result.success:
        return ToolOutput.fail(
            tool_name=_TOOL_HYBRID,
            errors=["Exact pass failed: " + "; ".join(exact_result.errors)],
        )

    exact_data = exact_result.data
    matched = list(exact_data.get("matched", []))
    unmatched_left_dicts = exact_data.get("unmatched_left", [])

    if not unmatched_left_dicts:
        # All rows matched exactly — no need for a fuzzy pass
        result = dict(exact_data)
        result["strategy"] = "hybrid_exact_only"
        return ToolOutput.ok(tool_name=_TOOL_HYBRID, data=result)

    # Reconstruct SheetData for unmatched left rows
    if unmatched_left_dicts:
        left_headers = list(unmatched_left_dicts[0].keys())
        left_rows_data = [left_headers] + [
            [row.get(h) for h in left_headers] for row in unmatched_left_dicts
        ]
        unmatched_left_sheet = SheetData(
            name=left.name + "_unmatched",
            data=left_rows_data,
            headers=left_headers,
        )

        fuzzy_result = run_fuzzy_match(
            unmatched_left_sheet,
            right,
            fuzzy_key_left,
            fuzzy_key_right,
            threshold=threshold,
            return_columns=return_columns,
        )

        if fuzzy_result.success:
            fuzzy_data = fuzzy_result.data
            matched += fuzzy_data.get("matched", [])
            final_unmatched_left = fuzzy_data.get("unmatched_left", [])
            final_unmatched_right = exact_data.get("unmatched_right", [])
        else:
            final_unmatched_left = unmatched_left_dicts
            final_unmatched_right = exact_data.get("unmatched_right", [])
    else:
        final_unmatched_left = []
        final_unmatched_right = exact_data.get("unmatched_right", [])

    match_result = MatchResult(
        left_indices=list(range(len(matched))),
        right_indices=list(range(len(matched))),
        scores=[1.0] * len(matched),
        match_count=len(matched),
        unmatched_left=len(final_unmatched_left),
        unmatched_right=len(final_unmatched_right),
        strategy_used=StrategyType.hybrid,
    )

    return ToolOutput.ok(
        tool_name=_TOOL_HYBRID,
        data={
            "matched": matched,
            "unmatched_left": final_unmatched_left,
            "unmatched_right": final_unmatched_right,
            "match_count": len(matched),
            "unmatched_left_count": len(final_unmatched_left),
            "unmatched_right_count": len(final_unmatched_right),
            "match_result": match_result.model_dump(),
            "strategy": "hybrid",
        },
        metadata={
            "exact_key_left": exact_key_left,
            "exact_key_right": exact_key_right,
            "fuzzy_key_left": fuzzy_key_left,
            "fuzzy_key_right": fuzzy_key_right,
            "threshold": threshold,
        },
    )


# ---------------------------------------------------------------------------
# run_semantic_match (stub — requires embeddings)
# ---------------------------------------------------------------------------

def run_semantic_match(
    left: SheetData,
    right: SheetData,
    left_key: str,
    right_key: str,
    threshold: float = 0.75,
    return_columns: list[str] | None = None,
) -> ToolOutput:
    """
    Semantic matching using sentence embeddings (requires sentence-transformers).

    Falls back to fuzzy matching when sentence-transformers is not installed.
    """
    try:
        from sentence_transformers import SentenceTransformer  # noqa: F401
    except ImportError:
        # Graceful fallback to fuzzy
        return run_fuzzy_match(
            left,
            right,
            left_key,
            right_key,
            threshold=int(threshold * 100),
            return_columns=return_columns,
        )

    # Semantic path — full implementation
    try:
        import pandas as pd
        from sentence_transformers import SentenceTransformer
        from sklearn.metrics.pairwise import cosine_similarity
        import numpy as np
    except ImportError as exc:
        return ToolOutput.fail(
            tool_name=_TOOL_SEMANTIC,
            errors=[f"Required package missing for semantic match: {exc}"],
        )

    left_df = left.to_dataframe()
    right_df = right.to_dataframe()

    if left_key not in left_df.columns:
        return ToolOutput.fail(
            tool_name=_TOOL_SEMANTIC,
            errors=[f"Key column '{left_key}' not found in sheet '{left.name}'."],
        )
    if right_key not in right_df.columns:
        return ToolOutput.fail(
            tool_name=_TOOL_SEMANTIC,
            errors=[f"Key column '{right_key}' not found in sheet '{right.name}'."],
        )

    model = SentenceTransformer("all-MiniLM-L6-v2")
    left_keys = left_df[left_key].astype(str).str.strip().tolist()
    right_keys = right_df[right_key].astype(str).str.strip().tolist()

    left_emb = model.encode(left_keys, convert_to_numpy=True)
    right_emb = model.encode(right_keys, convert_to_numpy=True)
    sim_matrix = cosine_similarity(left_emb, right_emb)

    matched: list[dict[str, Any]] = []
    unmatched_left: list[dict[str, Any]] = []
    matched_right_indices: set[int] = set()
    scores: list[float] = []

    for li, left_row in left_df.iterrows():
        best_ri = int(np.argmax(sim_matrix[li]))  # type: ignore[index]
        best_score = float(sim_matrix[li][best_ri])  # type: ignore[index]

        if best_score >= threshold:
            right_row = right_df.iloc[best_ri]
            combined = {**left_row.where(pd.notnull(left_row), other=None).to_dict()}
            if return_columns:
                for rc in return_columns:
                    if rc in right_df.columns:
                        combined[f"{rc}_matched"] = right_row[rc]
            else:
                for rc in right_df.columns:
                    if rc != right_key:
                        combined[f"{rc}_matched"] = right_row.get(rc)
            combined["_match_score"] = round(best_score, 4)
            combined["_matched_value"] = right_row[right_key]
            matched.append(combined)
            matched_right_indices.add(best_ri)
            scores.append(best_score)
        else:
            unmatched_left.append(left_row.where(pd.notnull(left_row), other=None).to_dict())

    unmatched_right = [
        right_df.iloc[i].where(pd.notnull(right_df.iloc[i]), other=None).to_dict()
        for i in range(len(right_df))
        if i not in matched_right_indices
    ]

    match_result = MatchResult(
        left_indices=list(range(len(matched))),
        right_indices=list(matched_right_indices),
        scores=scores,
        match_count=len(matched),
        unmatched_left=len(unmatched_left),
        unmatched_right=len(unmatched_right),
        strategy_used=StrategyType.semantic,
    )

    return ToolOutput.ok(
        tool_name=_TOOL_SEMANTIC,
        data={
            "matched": matched,
            "unmatched_left": unmatched_left,
            "unmatched_right": unmatched_right,
            "match_count": len(matched),
            "unmatched_left_count": len(unmatched_left),
            "unmatched_right_count": len(unmatched_right),
            "match_result": match_result.model_dump(),
        },
        metadata={
            "strategy": "semantic",
            "threshold": threshold,
            "left_key": left_key,
            "right_key": right_key,
        },
    )
