"""
Download and bundle the sentence-transformers embedding model into
`backend/models/<model_name>/` so enclosed / air-gapped deployments work
without ever hitting HuggingFace Hub at runtime.

Usage:
    python backend/scripts/download_embedding_model.py
    python backend/scripts/download_embedding_model.py --model some-other-model
    python backend/scripts/download_embedding_model.py --force    # re-download

Idempotent: skips the download if the target folder already has a config.json
unless --force is passed.

Notes:
- The `sentence-transformers/` HuggingFace repo prefix is applied automatically
  when no org is provided.
- Requires network access the first time it is run. Recommend committing the
  resulting folder to the repo via Git LFS so subsequent clones on enclosed
  networks already have it.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


DEFAULT_MODEL = "paraphrase-multilingual-MiniLM-L12-v2"
BACKEND_ROOT = Path(__file__).resolve().parents[1]
MODELS_DIR = BACKEND_ROOT / "models"


def download(model_name: str, force: bool = False) -> Path:
    """Download `model_name` into `backend/models/<model_name>/`."""
    target = MODELS_DIR / model_name
    target.mkdir(parents=True, exist_ok=True)

    # Skip if it already looks populated
    if not force and (target / "config.json").exists():
        print(f"[skip] {target} already contains config.json. Use --force to re-download.")
        return target

    try:
        from huggingface_hub import snapshot_download
    except ImportError:
        sys.exit(
            "huggingface_hub is required: pip install huggingface_hub"
        )

    repo_id = model_name if "/" in model_name else f"sentence-transformers/{model_name}"
    print(f"[download] {repo_id} → {target}")
    # Note: `local_dir_use_symlinks` was deprecated in huggingface_hub 0.23.
    # We pass only the supported args to stay forward-compatible.
    snapshot_download(
        repo_id=repo_id,
        local_dir=str(target),
    )
    print(f"[done] Model bundled at {target}")
    return target


def verify(path: Path) -> None:
    """Sanity check that the downloaded model loads via sentence-transformers."""
    try:
        from sentence_transformers import SentenceTransformer
    except ImportError:
        print("[warn] sentence-transformers not installed — skipping verification.")
        return

    print(f"[verify] Loading {path} with sentence-transformers...")
    model = SentenceTransformer(str(path))
    vec = model.encode("hello")
    print(f"[verify] OK — embedding dim = {len(vec)}")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--model", default=DEFAULT_MODEL,
                    help=f"Model name (default: {DEFAULT_MODEL})")
    ap.add_argument("--force", action="store_true",
                    help="Re-download even if the folder already looks populated")
    ap.add_argument("--no-verify", action="store_true",
                    help="Skip the sentence-transformers load test at the end")
    args = ap.parse_args()

    path = download(args.model, force=args.force)

    if not args.no_verify:
        verify(path)


if __name__ == "__main__":
    main()
