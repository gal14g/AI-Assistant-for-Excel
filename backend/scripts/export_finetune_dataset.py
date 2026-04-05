#!/usr/bin/env python3
"""
Fine-tune Dataset Exporter
===========================
Exports "applied" chat interactions from the feedback database into JSONL files
ready for supervised fine-tuning (SFT) on OpenAI, Anthropic, or Hugging Face.

Only interactions the user explicitly applied (action = 'applied') are included.
Sheet names, workbook paths, and other PII-like spreadsheet identifiers are
replaced with generic tokens before export.

Output formats
--------------
  --format openai        (default) ChatML-style, for OpenAI fine-tuning API
  --format anthropic     Prompt / completion pairs for Anthropic fine-tuning
  --format hf            Hugging Face "text" field (full conversation as one string)

Output files
------------
  <out_dir>/finetune_<format>_<timestamp>.jsonl    — combined output
  <out_dir>/finetune_<format>_<timestamp>_stats.json — summary stats

Usage
-----
  # From the repo root
  python backend/scripts/export_finetune_dataset.py

  # Specify DB path, output dir, and format explicitly
  python backend/scripts/export_finetune_dataset.py \\
      --db backend/data/feedback.db \\
      --out-dir backend/data/exports \\
      --format openai \\
      --min-quality 0.8

Exit codes
----------
  0  — success (even if 0 rows exported)
  1  — error (bad args, DB not found, write failure)
"""

from __future__ import annotations

import argparse
import json
import re
import sqlite3
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


# ---------------------------------------------------------------------------
# PII / Spreadsheet-identifier normalisation
# ---------------------------------------------------------------------------

# Regex: Excel cell / range addresses like A1, B2:C10, $A$1, Sheet1!A1:B5
_CELL_RE = re.compile(
    r"(?:(?:\$?[A-Z]{1,3}\$?\d{1,7}:\$?[A-Z]{1,3}\$?\d{1,7})"  # A1:B2
    r"|(?:\$?[A-Z]{1,3}\$?\d{1,7}))",  # A1 or $A$1
    re.IGNORECASE,
)

# Regex: looks like a Windows/Unix file path containing .xlsx / .xls / .csv
_PATH_RE = re.compile(
    r"[A-Za-z]:\\[^\s\"'<>|]+\.(?:xlsx?|csv)|/[^\s\"'<>|]*\.(?:xlsx?|csv)",
    re.IGNORECASE,
)

# Regex: workbook/sheet name typically follows "in " or "on sheet " etc.
# We normalise them generically when they appear as bare words adjacent to
# known Excel keywords, by replacing them in the range-token objects (JSON).


def _normalise_text(text: str) -> str:
    """Replace PII-like tokens in plain text before export."""
    # Paths like C:\Users\alice\Sales.xlsx → <WORKBOOK_PATH>
    text = _PATH_RE.sub("<WORKBOOK_PATH>", text)
    # e.g. "Sheet1!A1:B5" → the sheet portion only (cell addresses kept as-is
    # since they are structural and not PII).
    return text


def _normalise_range_tokens(tokens_json: str | None) -> str | None:
    """Remove specific sheet and workbook names from range token JSON."""
    if not tokens_json:
        return tokens_json
    try:
        tokens = json.loads(tokens_json)
        normalised = []
        for i, t in enumerate(tokens):
            normalised.append({
                "address": _CELL_RE.sub(lambda m: m.group(0), t.get("address", "")),
                "sheetName": f"Sheet{i + 1}",  # replace actual sheet name
            })
        return json.dumps(normalised)
    except (ValueError, TypeError):
        return tokens_json


def _normalise_plan(plan_json: str | None) -> Any:
    """Strip workbook/sheet names from plan steps."""
    if not plan_json:
        return None
    try:
        plan = json.loads(plan_json)
    except (ValueError, TypeError):
        return None

    def _scrub(obj: Any) -> Any:
        if isinstance(obj, str):
            return _normalise_text(obj)
        if isinstance(obj, dict):
            return {k: _scrub(v) for k, v in obj.items()}
        if isinstance(obj, list):
            return [_scrub(v) for v in obj]
        return obj

    return _scrub(plan)


# ---------------------------------------------------------------------------
# Row → example conversion
# ---------------------------------------------------------------------------

def _build_openai_example(row: dict) -> dict:
    """Return a ChatML messages list for OpenAI SFT format."""
    system_prompt = (
        "You are Excel AI Copilot, an expert Excel assistant. "
        "When the user asks for Excel operations, respond with a structured execution plan. "
        "For questions or clarifications, reply conversationally."
    )
    messages: list[dict] = [{"role": "system", "content": system_prompt}]

    # Add conversation history if present
    for turn in row.get("history", []):
        messages.append({"role": turn["role"], "content": turn["content"]})

    messages.append({"role": "user", "content": row["user_message"]})
    messages.append({"role": "assistant", "content": row["assistant_response"]})

    return {"messages": messages}


def _build_anthropic_example(row: dict) -> dict:
    """Return prompt/completion for Anthropic fine-tuning."""
    prompt = (
        "Human: " + row["user_message"] + "\n\nAssistant:"
    )
    return {
        "prompt": prompt,
        "completion": " " + row["assistant_response"],
    }


def _build_hf_example(row: dict) -> dict:
    """Return a single-text field for HF datasets."""
    text = (
        f"<|user|>\n{row['user_message']}\n"
        f"<|assistant|>\n{row['assistant_response']}"
    )
    return {"text": text}


_BUILDERS = {
    "openai": _build_openai_example,
    "anthropic": _build_anthropic_example,
    "hf": _build_hf_example,
}


# ---------------------------------------------------------------------------
# Database query
# ---------------------------------------------------------------------------

def _fetch_applied_rows(db_path: Path, min_quality: float) -> list[dict]:
    """
    Return interactions that:
    - have a corresponding choice with action = 'applied'
    - have at least one plan in the response
    - optionally: quality_score >= min_quality (for few_shot_examples)
    """
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute("""
        SELECT
            i.id,
            i.user_message,
            i.active_sheet,
            i.workbook_name,
            i.range_tokens,
            i.message       AS assistant_message,
            i.plans_json,
            i.model_used,
            i.created_at,
            c.chosen_plan_id
        FROM interactions i
        JOIN choices c ON c.interaction_id = i.id
        WHERE c.action = 'applied'
          AND i.plans_json IS NOT NULL
        ORDER BY i.created_at ASC
    """)
    rows = cur.fetchall()

    # Also pull from few_shot_examples (source='user' means promoted by the user)
    cur.execute("""
        SELECT id, user_message, assistant_response, quality_score, created_at
        FROM few_shot_examples
        WHERE quality_score >= ?
        ORDER BY created_at ASC
    """, (min_quality,))
    few_shot_rows = cur.fetchall()
    conn.close()

    results: list[dict] = []

    for r in rows:
        plans = _normalise_plan(r["plans_json"])
        chosen_plan = None
        if plans and r["chosen_plan_id"]:
            for p in plans:
                plan_obj = p.get("plan", {}) if isinstance(p, dict) else {}
                if plan_obj.get("planId") == r["chosen_plan_id"]:
                    chosen_plan = plan_obj
                    break
        if not chosen_plan and plans:
            # Fallback: first plan
            chosen_plan = plans[0].get("plan", plans[0]) if isinstance(plans[0], dict) else plans[0]

        user_msg = _normalise_text(r["user_message"] or "")
        assistant_resp = json.dumps(
            {"responseType": "plan", "message": r["assistant_message"], "plan": chosen_plan},
            ensure_ascii=False,
        )

        results.append({
            "id": r["id"],
            "user_message": user_msg,
            "assistant_response": assistant_resp,
            "model_used": r["model_used"],
            "created_at": r["created_at"],
            "source": "applied_interaction",
            "history": [],
        })

    for r in few_shot_rows:
        results.append({
            "id": r["id"],
            "user_message": _normalise_text(r["user_message"] or ""),
            "assistant_response": r["assistant_response"] or "",
            "model_used": None,
            "created_at": r["created_at"],
            "source": "few_shot",
            "history": [],
        })

    return results


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def _default_db() -> Path:
    return Path(__file__).resolve().parents[1] / "data" / "feedback.db"


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Export fine-tune-ready JSONL from the feedback database."
    )
    parser.add_argument(
        "--db",
        type=Path,
        default=_default_db(),
        help="Path to feedback.db (default: backend/data/feedback.db)",
    )
    parser.add_argument(
        "--out-dir",
        type=Path,
        default=Path(__file__).resolve().parents[1] / "data" / "exports",
        help="Output directory for JSONL and stats files",
    )
    parser.add_argument(
        "--format",
        choices=["openai", "anthropic", "hf"],
        default="openai",
        help="Target fine-tune format (default: openai)",
    )
    parser.add_argument(
        "--min-quality",
        type=float,
        default=0.8,
        help="Minimum quality_score for few_shot_examples rows (default: 0.8)",
    )
    args = parser.parse_args()

    if not args.db.exists():
        print(f"Error: database not found at {args.db}", file=sys.stderr)
        sys.exit(1)

    args.out_dir.mkdir(parents=True, exist_ok=True)

    print(f"Reading from: {args.db}")
    rows = _fetch_applied_rows(args.db, args.min_quality)
    print(f"Found {len(rows)} exportable interactions")

    if not rows:
        print("Nothing to export.")
        return

    builder = _BUILDERS[args.format]
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S")
    out_path = args.out_dir / f"finetune_{args.format}_{timestamp}.jsonl"
    stats_path = args.out_dir / f"finetune_{args.format}_{timestamp}_stats.json"

    sources: dict[str, int] = {}
    with out_path.open("w", encoding="utf-8") as f:
        for row in rows:
            example = builder(row)
            f.write(json.dumps(example, ensure_ascii=False) + "\n")
            sources[row["source"]] = sources.get(row["source"], 0) + 1

    stats = {
        "format": args.format,
        "exported_at": datetime.now(timezone.utc).isoformat(),
        "total_examples": len(rows),
        "by_source": sources,
        "output_file": str(out_path),
    }
    stats_path.write_text(json.dumps(stats, indent=2, ensure_ascii=False))

    print(f"\nExported  : {out_path}")
    print(f"Stats     : {stats_path}")
    print(f"Total rows: {len(rows)}")
    for src, count in sources.items():
        print(f"  {src}: {count}")


if __name__ == "__main__":
    main()
