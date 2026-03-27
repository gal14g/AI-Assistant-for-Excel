# Excel AI Copilot

A Microsoft Excel Office Add-in that lets you control your spreadsheet with natural language. Powered by any LLM via LiteLLM (OpenAI, Anthropic Claude, Ollama, Azure, etc.).

---

## Table of Contents

1. [What it does](#what-it-does)
2. [Project structure](#project-structure)
3. [Local development](#local-development)
4. [Making changes](#making-changes)
5. [Deploy to OpenShift](#deploy-to-openshift)
6. [Installing the add-in in Excel](#installing-the-add-in-in-excel)
7. [Configuration reference](#configuration-reference)

---

## What it does

- Chat with an AI assistant directly inside Excel
- The AI understands your spreadsheet context (selected ranges, sheet names)
- Executes multi-step operations: match records, create charts/pivots, sort, clean text, apply conditional formatting, and more
- XLOOKUP with automatic fallback to VLOOKUP for older Excel versions (2016/2019)
- Full undo support after every AI operation

---

## Project structure

```
excel-ai-copilot/
├── backend/              Python FastAPI backend
│   ├── app/
│   │   ├── routers/      API endpoints (/api/chat, /api/plan)
│   │   ├── services/     LLM integration, chat service, planner
│   │   ├── models/       Pydantic request/response models
│   │   └── config.py     All settings (env vars)
│   ├── main.py           FastAPI app entry point
│   └── requirements.txt
├── frontend/             React + TypeScript Office Add-in
│   ├── src/
│   │   ├── engine/       Execution engine, capabilities, snapshot/rollback
│   │   ├── services/     API client
│   │   └── taskpane/     UI components, hooks
│   ├── manifest.xml      Office Add-in manifest (loaded by Excel)
│   └── webpack.config.js
├── Dockerfile            Multi-stage build (frontend + backend in one image)
├── docker-compose.yml    Local Docker testing
├── .env.example          All environment variables documented
└── README.md             This file
```

---

## Local development

This is the workflow for **day-to-day coding**. No Docker needed.

### Prerequisites

- Python 3.11+
- Node.js 20+
- An LLM API key (Anthropic, OpenAI, etc.) — or Ollama running locally

### 1. Clone and install

```bash
git clone https://github.com/gal14g/excel-ai-copilot.git
cd excel-ai-copilot
```

**Backend:**
```bash
cd backend
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

**Frontend:**
```bash
cd frontend
npm install
```

### 2. Configure your LLM key

Create `backend/.env`:
```env
LLM_MODEL=claude-sonnet-4-20250514
LLM_API_KEY=your-api-key-here
```

See [Configuration reference](#configuration-reference) for all options (OpenAI, Ollama, Azure, etc.).

### 3. Start both servers

**Terminal 1 — Backend:**
```bash
cd backend
source venv/bin/activate
uvicorn main:app --reload --port 8000
```

**Terminal 2 — Frontend:**
```bash
cd frontend
npm run dev
```

The frontend dev server starts at `https://localhost:3000`.
The backend API runs at `http://localhost:8000`.
Webpack proxies `/api` calls from the frontend to the backend automatically.

### 4. Load the add-in in Excel

In Excel: **Insert → Add-ins → Upload My Add-in** → enter:
```
https://localhost:3000/manifest.xml
```

Accept the self-signed certificate warning. The **AI Copilot** tab appears in the Excel ribbon.

> **Note:** Excel requires HTTPS for add-ins. The webpack dev server provides a self-signed HTTPS certificate automatically.

---

## Making changes

### Backend changes
The backend reloads automatically (`--reload` flag). Just save the file — no restart needed.

### Frontend changes
Webpack hot-reloads automatically. Save the file and the add-in updates within a few seconds.

### Do I need to rebuild Docker every time I make a change?

**No.** Docker is only for deploying to OpenShift.

Your local development loop never involves Docker:
```
Edit code → save → see changes live in Excel immediately
```

You only rebuild the Docker image when you are ready to push a new version to OpenShift. And thanks to layer caching, rebuilds are fast after the first time — Docker only rebuilds the layers that actually changed.

---

## Deploy to OpenShift

### How it works

The Docker image bundles the React frontend and FastAPI backend into a single container. OpenShift provides HTTPS automatically via a Route — no certificate configuration is needed inside the container.

```
User's Excel
    │  HTTPS
    ▼
OpenShift Route  (TLS terminated here — free HTTPS)
    │  HTTP
    ▼
Container :8080
    ├── GET /api/*    →  FastAPI (AI chat, plan generation)
    ├── GET /health   →  health check
    └── GET /*        →  React build (taskpane.html, manifest.xml, assets)
```

Multiple users can use the add-in simultaneously. FastAPI is fully async — while one user's LLM request is in-flight, other requests are handled concurrently. No message queue needed.

---

### Prerequisites

- Docker installed and running on your machine
- `oc` CLI installed ([download here](https://mirror.openshift.com/pub/openshift-v4/clients/ocp/latest/))
- Logged in to your OpenShift cluster: `oc login https://your-cluster.example.com`
- Access to a container image registry (OpenShift built-in, Quay.io, Docker Hub, etc.)

---

### Step 1 — Get your public URL

You need the public HTTPS URL your add-in will live at **before building**, because it gets baked into `manifest.xml`.

The URL pattern is: `https://<app-name>.apps.<cluster-domain>`

**If you don't know it yet:** create a placeholder Route first, get the URL, then build. Or ask your cluster admin.

Example URL: `https://excel-copilot.apps.my-cluster.example.com`

---

### Step 2 — Build the Docker image

Run this from the repo root, substituting your real URL:

```bash
docker build \
  --build-arg FRONTEND_URL=https://excel-copilot.apps.my-cluster.example.com \
  -t excel-ai-copilot:latest \
  .
```

**First build:** 5–10 minutes (downloads Node and Python base images, installs all dependencies).
**Subsequent builds:** 1–3 minutes (Docker caches layers — only changed code rebuilds).

---

### Step 3 — Push the image to a registry

Pick whichever registry you have access to:

**Option A — OpenShift's built-in registry:**
```bash
# Log in to the registry
oc registry login

# Find your registry address
oc get route default-route -n openshift-image-registry

# Tag and push  (replace <registry-address> and <your-namespace>)
docker tag excel-ai-copilot:latest \
  <registry-address>/<your-namespace>/excel-ai-copilot:latest

docker push <registry-address>/<your-namespace>/excel-ai-copilot:latest
```

**Option B — Quay.io:**
```bash
docker login quay.io
docker tag excel-ai-copilot:latest quay.io/<your-username>/excel-ai-copilot:latest
docker push quay.io/<your-username>/excel-ai-copilot:latest
```

**Option C — Docker Hub:**
```bash
docker login
docker tag excel-ai-copilot:latest <your-dockerhub-username>/excel-ai-copilot:latest
docker push <your-dockerhub-username>/excel-ai-copilot:latest
```

---

### Step 4 — Create a Secret for your API key

Never put secrets in the image or in YAML files. Store them in an OpenShift Secret:

```bash
oc create secret generic excel-copilot-secrets \
  --from-literal=LLM_API_KEY=your-api-key-here \
  --from-literal=LLM_MODEL=claude-sonnet-4-20250514
```

---

### Step 5 — Deploy the container

**Option A — Web console (recommended for first deploy):**
1. Open the OpenShift web console → your project
2. Click **+Add → Container Image**
3. Enter your image reference (from step 3)
4. Set the container port to **8080**
5. Under **Environment Variables**, add:
   - `OPENSHIFT` = `true`
   - Add `LLM_API_KEY` from your Secret (select "From Secret", choose `excel-copilot-secrets`)
   - Add `LLM_MODEL` from your Secret
6. Click **Create** — OpenShift creates a Deployment, Service, and Route automatically
7. Click on the Route → **Edit** → set **TLS Termination** to **Edge**

**Option B — `oc` CLI:**
```bash
# Deploy from image
oc new-app --image=quay.io/<your-username>/excel-ai-copilot:latest \
  --name=excel-copilot

# Set the OPENSHIFT flag and inject the secret
oc set env deployment/excel-copilot OPENSHIFT=true
oc set env deployment/excel-copilot --from=secret/excel-copilot-secrets

# Expose with HTTPS edge TLS
oc create route edge excel-copilot \
  --service=excel-copilot \
  --port=8080

# Get the assigned URL
oc get route excel-copilot
```

---

### Step 6 — Verify the deployment

```bash
# Should return: {"status":"ok","version":"1.0.0","mode":"openshift"}
curl https://excel-copilot.apps.my-cluster.example.com/health

# Should return the manifest XML
curl https://excel-copilot.apps.my-cluster.example.com/manifest.xml
```

---

### Deploying updates (after code changes)

```bash
# 1. Rebuild  (fast — only changed layers rebuild)
docker build \
  --build-arg FRONTEND_URL=https://excel-copilot.apps.my-cluster.example.com \
  -t excel-ai-copilot:latest \
  .

# 2. Push the new image
docker push quay.io/<your-username>/excel-ai-copilot:latest

# 3. Restart the pods — OpenShift pulls the new image with zero downtime
oc rollout restart deployment/excel-copilot
```

---

## Installing the add-in in Excel

### For yourself (sideloading)

1. Open Excel
2. Go to **Insert → Add-ins → Upload My Add-in**
3. Enter the manifest URL:
   - **Local dev:** `https://localhost:3000/manifest.xml`
   - **OpenShift:** `https://excel-copilot.apps.my-cluster.example.com/manifest.xml`
4. Click **Upload**
5. The **AI Copilot** tab appears in the Excel ribbon

### For your whole organisation (Microsoft 365 Admin Center)

No action needed from individual users — the add-in appears in Excel automatically.

1. Log in to [admin.microsoft.com](https://admin.microsoft.com) as a Microsoft 365 admin
2. Go to **Settings → Integrated apps → Upload custom apps**
3. Select **Provide link to manifest** and enter your OpenShift manifest URL
4. Assign to specific users, groups, or the entire organisation
5. Done — the add-in rolls out to all assigned users within 24 hours

---

## Configuration reference

All settings are environment variables. Copy `.env.example` to `backend/.env` for local development.

| Variable | Local default | Production default | Description |
|---|---|---|---|
| `LLM_MODEL` | `claude-sonnet-4-20250514` | same | LiteLLM model string |
| `LLM_API_KEY` | _(empty)_ | _(from Secret)_ | API key for your LLM provider |
| `LLM_BASE_URL` | _(empty)_ | _(empty)_ | Custom API base URL (Ollama, Azure, proxy) |
| `LLM_API_VERSION` | _(empty)_ | _(empty)_ | Azure OpenAI API version only |
| `LLM_MAX_TOKENS` | `4096` | `4096` | Max tokens per LLM response |
| `LLM_TEMPERATURE` | `0.1` | `0.1` | Lower = more deterministic plans |
| `PORT` | `8000` | `8080` | Port the backend listens on |
| `DEBUG` | `true` | `false` | Enables uvicorn auto-reload |
| `OPENSHIFT` | `false` | `true` | Master production switch |
| `SERVE_STATIC` | `false` | `true` | Serve built frontend from FastAPI |
| `STATIC_DIR` | `./static` | `./static` | Path to built frontend files |
| `FRONTEND_URL` | `https://localhost:3000` | your OpenShift URL | **Build-time only** — baked into `manifest.xml` |

### LLM provider examples

```env
# Anthropic Claude (recommended)
LLM_MODEL=claude-sonnet-4-20250514
LLM_API_KEY=sk-ant-...

# OpenAI
LLM_MODEL=gpt-4o
LLM_API_KEY=sk-...

# Ollama (local, no key needed)
LLM_MODEL=ollama/llama3
LLM_BASE_URL=http://localhost:11434

# Azure OpenAI
LLM_MODEL=azure/my-deployment-name
LLM_API_KEY=...
LLM_BASE_URL=https://my-resource.openai.azure.com/
LLM_API_VERSION=2024-02-01

# Google Gemini
LLM_MODEL=gemini/gemini-1.5-pro
LLM_API_KEY=...
```
