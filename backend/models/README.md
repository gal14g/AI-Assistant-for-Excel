# Bundled embedding models

This folder contains sentence-transformers models that ship with the repo so
**enclosed / air-gapped deployments work without internet access**.

The default model (referenced by `app.config.settings.embedding_model` and by
`openshift/configmap.yaml`) is:

```
paraphrase-multilingual-MiniLM-L12-v2     (~420 MB, supports Hebrew + English)
```

## How the runtime picks up the bundled model

`backend/app/services/chroma_client.py::_resolve_embedding_model()` looks in
this order:
1. If the configured value is an absolute/relative path that exists → use it.
2. If `backend/models/<configured_name>/` exists → use the bundled copy.
3. Otherwise → fall back to the bare name (downloads from HuggingFace Hub).

When `HF_HUB_OFFLINE=1` is set (the default in `openshift/configmap.yaml`),
step 3 fails — so step 2 *must* succeed, which means the folder for your
configured model must be present here.

## Populating this folder

### Option A — Clone with Git LFS (recommended)
The `.gitattributes` at the repo root marks `*.safetensors` / `*.bin` under
`backend/models/**` as LFS tracked. After `git clone`, run once:

```bash
git lfs install      # one-time per machine
git lfs pull         # fetch the large blobs
```

### Option B — Download from HuggingFace
If you are a repo maintainer and want to add a new bundled model (or refresh
the existing one), run:

```bash
python backend/scripts/download_embedding_model.py
# or a non-default model:
python backend/scripts/download_embedding_model.py --model some-other-model
```

Then commit the resulting folder. LFS will take care of the large binary
files automatically (because of `.gitattributes`).

### Option C — Runtime fallback
The Dockerfile includes a build-time safety net: if `backend/models/<name>/`
is missing from the build context, it runs the download script. Requires
network access during `docker build` (but not at runtime).

## Why not just download at runtime?

Because **enclosed networks** — the primary deployment target — do not have
HuggingFace Hub access. The model must already be on the container filesystem
when the pod boots, otherwise `SentenceTransformer(...)` hangs waiting for a
connection that will never come.
