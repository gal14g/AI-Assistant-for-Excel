# Air-gapped / enclosed-network deployment

Everything that needs to change to run Excel AI Copilot in a network with
zero outbound internet access.

---

## What phones home (and the fix for each)

| # | Component | Default behavior | Fix |
|---|---|---|---|
| 1 | **Office.js** | Loaded from `appsforoffice.microsoft.com` CDN in `taskpane.html` + `commands.html` | Self-host it; see §1 below |
| 2 | **LLM API** | Configurable — most defaults hit internet | Use Ollama or an internal OpenAI-compatible gateway; see §2 |
| 3 | **ChromaDB telemetry** | Sends anonymous events to `posthog.com` | Already disabled in `chroma_client.py` + `ANONYMIZED_TELEMETRY=False` env var |
| 4 | **HuggingFace Hub** | `sentence-transformers` may check model version | Already disabled via `HF_HUB_OFFLINE=1` + `TRANSFORMERS_OFFLINE=1` |
| 5 | **Docker base images** | `node:20-alpine` + `python:3.11-slim` pulled from Docker Hub | Build image on an internet-connected host, then transfer; see §5 |
| 6 | **npm + pip** | Pull packages at build time | Same — happens at build, not runtime |

Items 3 and 4 are **already handled** in this repo (code + env vars set in
`Dockerfile` and `openshift/configmap.yaml`). Items 1, 2, 5 need your action.

---

## 1. Vendoring Office.js

The build is controlled by the `OFFICE_JS_SRC` env var (webpack) / build-arg
(Docker). Default is Microsoft's CDN. For enclosed networks:

```bash
# On an internet-connected host: download office.js once (~2 MB)
curl -o frontend/public/assets/office.js \
  https://appsforoffice.microsoft.com/lib/1/hosted/office.js

# Commit it (or ship it alongside the repo)
git add frontend/public/assets/office.js
git commit -m "Vendor office.js for air-gapped deployments"
```

Then build with the local path:

```bash
docker build \
  --build-arg FRONTEND_URL="https://excel-copilot.apps.cluster.example.com" \
  --build-arg OFFICE_JS_SRC=/assets/office.js \
  -t your-registry/excel-copilot:v1.1.0 .
```

That's it — webpack's `HtmlWebpackPlugin` injects the value into
`taskpane.html` and `commands.html` at build time. No runtime changes needed.

**Refresh cadence:** Microsoft updates `office.js` periodically with bugfixes.
Re-download every 3-6 months and rebuild. Nothing auto-updates in this mode.

---

## 2. LLM provider inside the enclosed network

The project uses the **OpenAI Python SDK directly** (no LiteLLM dependency).
It works with any OpenAI-compatible endpoint via `LLM_BASE_URL`.

### Option A: Ollama (self-hosted, zero external deps)

Run an Ollama pod/container inside the cluster:

```yaml
# ConfigMap values
LLM_MODEL: "ollama/qwen2.5:14b-instruct"
LLM_BASE_URL: "http://ollama.your-namespace.svc.cluster.local:11434"
LLM_JSON_MODE: "true"   # recommended for smaller models
LLM_API_KEY: ""         # not needed for Ollama
```

Model files must be pre-pulled into the Ollama container during its own build
(same air-gap problem — pull on an internet host, bake into image, transfer).

### Option B: Internal OpenAI-compatible gateway

Many enterprises run an internal proxy (Azure OpenAI, vLLM, TGI, LocalAI):

```yaml
LLM_MODEL: "gpt-4o"                                  # whatever your gateway exposes
LLM_BASE_URL: "https://llm-gateway.corp.internal/v1"
LLM_API_KEY: "<gateway-token>"                       # in the Secret
```

### Option C: Allowlist specific external LLM hosts

If your network allows targeted egress (e.g. only `api.openai.com` whitelisted
at the egress firewall), no code changes needed — set `LLM_MODEL` and
`LLM_API_KEY` as usual.

---

## 3. ChromaDB telemetry — already disabled

Two layers:

1. Code: `backend/app/services/chroma_client.py` passes
   `Settings(anonymized_telemetry=False)` to `PersistentClient`.
2. Env var: `ANONYMIZED_TELEMETRY=False` set in `Dockerfile` + `configmap.yaml`.

Nothing to do. This is on by default.

---

## 4. HuggingFace offline mode — already set

The embedding model (`all-MiniLM-L6-v2`, 87 MB) is **bundled** in
`backend/models/` and copied into the image. At runtime:

- `HF_HUB_OFFLINE=1` — no HuggingFace Hub calls
- `TRANSFORMERS_OFFLINE=1` — no `transformers` update checks

Both are set in `Dockerfile` + `configmap.yaml`. `chroma_client.py` loads the
local path directly, so `sentence-transformers` never attempts a download.

---

## 5. Getting the image into the enclosed network

**You cannot just push the repo to GitHub and clone inside the enclosed
network.** The build requires internet access (npm registry, PyPI, Docker
Hub base images). You need one of:

### Option A: Build outside, transfer the image (most common)

```bash
# On internet-connected host
docker build \
  --build-arg FRONTEND_URL="https://excel-copilot.apps.enclosed.example.com" \
  --build-arg OFFICE_JS_SRC=/assets/office.js \
  -t excel-copilot:v1.1.0 .

# Save to a tar file
docker save excel-copilot:v1.1.0 | gzip > excel-copilot-v1.1.0.tar.gz

# Transfer tar file across the air gap (USB, secure file transfer, etc.)

# On enclosed host: load and tag for internal registry
docker load < excel-copilot-v1.1.0.tar.gz
docker tag excel-copilot:v1.1.0 internal-registry.corp/excel-copilot:v1.1.0
docker push internal-registry.corp/excel-copilot:v1.1.0
```

### Option B: OpenShift internal build (BuildConfig)

If your cluster has an internal mirror of Docker Hub + npm + PyPI (e.g. Nexus,
Artifactory), OpenShift can build from source:

```bash
oc new-build --binary --name=excel-copilot --strategy=docker
oc start-build excel-copilot --from-dir=. --follow
```

This requires mirror configs in your cluster. Most enclosed networks already
have these.

### Option C: Mirror the GitHub repo internally

Mirror the repo to an internal Git server. This gets you the **source** across
the gap, but you still need one of A or B to actually build the image —
mirroring source doesn't solve the npm/pip/base-image problem.

---

## So: push to GitHub → clone in enclosed network → done?

**No.** Source code is only one piece. The enclosed-network build needs:

| Thing | How to get it across |
|---|---|
| Source code | Git mirror OR tarball |
| `office.js` | Commit it to the repo (`frontend/public/assets/office.js`) before transfer |
| Node base image + npm packages | Internal Docker Hub mirror + npm registry mirror, OR pre-built image |
| Python base image + pip packages | Internal Docker Hub mirror + PyPI mirror, OR pre-built image |
| Embedding model | Already in repo (`backend/models/all-MiniLM-L6-v2/`) |
| LLM weights (if using Ollama) | Pre-pull into Ollama image outside, transfer |
| LLM API keys (if using external gateway) | Injected at runtime via Secret |

The simplest path for a one-time deployment is **build image outside, ship tar
file across, `docker load` + push to internal registry, then follow
DEPLOYMENT.md** starting at step 3.

---

## Summary checklist

- [ ] Download `office.js` once, commit to `frontend/public/assets/office.js`
- [ ] Pick LLM strategy (Ollama in-cluster / internal gateway / allowlisted egress)
- [ ] Build image with `--build-arg OFFICE_JS_SRC=/assets/office.js`
- [ ] Transfer image into enclosed network (`docker save` / `docker load`)
- [ ] Push to internal registry
- [ ] Update `openshift/configmap.yaml` `LLM_BASE_URL` for your LLM choice
- [ ] Update `openshift/deployment.yaml` `image:` to internal registry path
- [ ] Create Secret with `LLM_API_KEY` (if needed)
- [ ] `oc apply -f openshift/`
- [ ] Verify `/ready` returns 200
- [ ] Test `manifest.xml` loads over the Route
- [ ] Sideload in Excel

Air-gap-specific env vars already set in configmap: `ANONYMIZED_TELEMETRY`,
`HF_HUB_OFFLINE`, `TRANSFORMERS_OFFLINE`.
