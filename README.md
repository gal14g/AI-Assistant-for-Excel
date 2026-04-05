# Excel AI Copilot

A Microsoft Excel Office Add-in that lets you control your spreadsheet with natural language. Powered by any OpenAI-compatible LLM provider (OpenAI, Google Gemini, Azure, Ollama, and more).

---

## Table of Contents

1. [What it does](#what-it-does)
2. [Architecture overview](#architecture-overview)
3. [Project structure](#project-structure)
4. [Local development](#local-development)
5. [Making changes](#making-changes)
6. [Switching LLM providers](#switching-llm-providers)
7. [Deploy to OpenShift](#deploy-to-openshift)
8. [Installing the add-in in Excel](#installing-the-add-in-in-excel)
9. [CI/CD pipeline](#cicd-pipeline)
10. [Security](#security)
11. [Migrating to a production database](#migrating-to-a-production-database)
12. [Configuration reference](#configuration-reference)

---

## What it does

- Chat with an AI assistant directly inside Excel's task pane
- The AI understands your spreadsheet context (selected ranges, sheet names, workbook name)
- Returns **2-3 alternative plan options** for each request so you can pick the best approach
- Executes multi-step operations: match records, create charts/pivots, sort, clean text, apply conditional formatting, find & replace, format cells, and 30+ more actions
- XLOOKUP with automatic fallback to VLOOKUP for older Excel versions (2016/2019)
- Full undo support after every AI operation
- **Self-improving**: every plan you approve is stored and used as a few-shot example for future queries via vector similarity search (ChromaDB + sentence-transformers)
- All interactions are logged to a local feedback database for analysis and fine-tuning
- Hebrew and English support with full RTL

---

## Architecture overview

```
User's Excel
    |
    | Office.js (HTTPS)
    v
Frontend (React + TypeScript)
    |  taskpane UI, execution engine, 34+ capability handlers
    |  POST /api/chat
    v
Backend (FastAPI + OpenAI SDK)
    |
    |-- Vector search (ChromaDB + all-MiniLM-L6-v2)
    |   |-- Capability store: finds the ~10 most relevant actions for each query
    |   |-- Example store: retrieves the ~5 most relevant few-shot examples
    |
    |-- LLM call (any OpenAI-compatible provider)
    |   Returns 2-3 alternative plan options as structured JSON
    |
    |-- Feedback DB (SQLite)
    |   Logs every interaction + user's choice (applied / dismissed)
    |   Applied plans are promoted into the example store automatically
    |
    v
Response: plan options -> user picks one -> frontend executes via Office.js
```

When deployed to OpenShift, a single container serves both the API and the frontend:

```
User's Excel
    |  HTTPS
    v
OpenShift Route  (TLS terminated here — free HTTPS)
    |  HTTP
    v
Container :8080
    |-- GET /api/*      ->  FastAPI (AI chat, plan generation, feedback)
    |-- GET /health     ->  health check
    |-- GET /ready      ->  readiness probe (LLM config + vector store)
    |-- GET /manifest.xml -> Office add-in manifest (users load this URL in Excel)
    +-- GET /*          ->  React build (taskpane.html, assets)
```

---

## Project structure

```
excel-ai-copilot/
|
|-- .env.example              All environment variables documented
|-- .env                      Your local config (git-ignored)
|-- .gitignore
|-- .gitlab-ci.yml            GitLab CI/CD pipeline (6-stage)
|-- Dockerfile                Multi-stage build (frontend + backend in one image)
|-- docker-compose.yml        Local Docker testing
|-- README.md                 This file
|-- SECURITY_CHECKLIST.md     Pre-deployment security hardening guide
|
|-- openshift/                OpenShift/Kubernetes deployment manifests
|   |-- deploy.sh             Quick deploy script (one command)
|   |-- deployment.yaml       Pod/replica config with probes
|   |-- service.yaml          ClusterIP networking
|   |-- route.yaml            HTTPS TLS termination + HSTS
|   |-- pvc.yaml              Persistent storage for data
|   |-- configmap.yaml        Non-secret config
|   |-- secret.yaml           Secret template (LLM_API_KEY)
|
|-- backend/                  Python FastAPI backend
|   |-- main.py               App entry point, startup hooks, CORS, rate limiting
|   |-- requirements.txt      Python dependencies
|   |-- app/
|   |   |-- config.py         Settings loaded from .env (Pydantic BaseSettings)
|   |   |-- db.py             SQLite feedback database (aiosqlite)
|   |   |-- routers/          API endpoint definitions
|   |   |-- services/         Business logic (chat, planner, vector search)
|   |   |-- models/           Pydantic request/response schemas
|   |   |-- prompts/          LLM system prompt templates
|   |-- data/                 Runtime data (git-ignored, auto-created)
|   |-- tests/
|
|-- frontend/                 React + TypeScript Office Add-in
    |-- manifest.xml          Office add-in manifest (loaded by Excel)
    |-- webpack.config.js     Build config + HTTPS + API proxy
    |-- package.json
    |-- src/
        |-- engine/           Execution engine (34+ capability handlers)
        |-- services/         API client
        |-- taskpane/         UI components + hooks
        |-- commands/         Office ribbon command handlers
```

---

## Local development

### Prerequisites

- Python 3.11+
- Node.js 20+
- An LLM API key (Anthropic, OpenAI, etc.) or Ollama running locally

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

### 2. Configure your LLM

Create `.env` in the project root (copy from `.env.example`):
```env
LLM_MODEL=gpt-4o
LLM_API_KEY=your-api-key-here
```

See [Switching LLM providers](#switching-llm-providers) for all options.

### 3. Start both servers

**Terminal 1 -- Backend:**
```bash
cd backend
source venv/bin/activate
uvicorn main:app --reload --port 8000
```

On first startup, the backend will:
1. Download the `all-MiniLM-L6-v2` embedding model (~80MB, cached after first time)
2. Index all 34 capabilities into ChromaDB (~2 seconds)
3. Seed curated few-shot examples into the example store
4. Create the SQLite feedback database

**Terminal 2 -- Frontend:**
```bash
cd frontend
npm run dev
```

The frontend dev server starts at `https://localhost:3000`.
The backend API runs at `http://localhost:8000`.
Webpack proxies `/api` calls from the frontend to the backend automatically.

### 4. Load the add-in in Excel

In Excel: **Insert > Add-ins > Upload My Add-in** > enter:
```
https://localhost:3000/manifest.xml
```

Accept the self-signed certificate warning. The **AI Copilot** tab appears in the Excel ribbon.

> **Note:** Excel requires HTTPS for add-ins. The webpack dev server provides a self-signed HTTPS certificate automatically.

---

## Making changes

### Backend changes
The backend reloads automatically (`--reload` flag). Just save the file -- no restart needed.

**Exception:** If you add a new capability, you must restart the backend so ChromaDB re-indexes.

### Frontend changes
Webpack hot-reloads automatically. Save the file and the add-in updates within a few seconds.

### Adding a new capability

1. **Backend types** (`backend/app/models/plan.py`): Add to `StepAction` enum, create a Pydantic params model, add to `ACTION_PARAM_MODELS`
2. **Backend description** (`backend/app/services/planner.py`): Add to `CAPABILITY_DESCRIPTIONS` dict
3. **Backend examples** (`backend/app/services/capability_store.py`): Add example user queries to `CAPABILITY_EXAMPLES`
4. **Frontend types** (`frontend/src/engine/types.ts`): Add to `StepAction` union, create params interface, add to `StepParams` union
5. **Frontend handler** (`frontend/src/engine/capabilities/`): Create a new `.ts` file with the handler
6. **Frontend registration** (`frontend/src/engine/capabilities/index.ts`): Import the new handler
7. **Frontend validation** (`frontend/src/engine/validator.ts`): Add required field checks
8. Restart backend (so ChromaDB re-indexes)

---

## Switching LLM providers

All LLM configuration is in `.env` (project root). The backend uses the **OpenAI Python SDK** — any provider that exposes an OpenAI-compatible API works out of the box.

The provider is auto-detected from the `LLM_MODEL` prefix or `LLM_BASE_URL`. No code changes needed.

### OpenAI (recommended)

```env
LLM_MODEL=gpt-4o
LLM_API_KEY=sk-...
```

### Google Gemini

Gemini exposes an [OpenAI-compatible API](https://ai.google.dev/gemini-api/docs/openai). The `gemini/` prefix auto-routes to the correct base URL.

```env
LLM_MODEL=gemini/gemini-2.0-flash
LLM_API_KEY=AIza...
```

Other Gemini models: `gemini/gemini-2.5-pro`, `gemini/gemini-2.5-flash`.

### Ollama (free, local, no API key)

First install and run a model:
```bash
ollama pull qwen2.5:14b
ollama serve
```

Then configure:
```env
LLM_MODEL=ollama/qwen2.5:14b
LLM_JSON_MODE=true
```

The `ollama/` prefix auto-routes to `http://localhost:11434/v1`. Override with `LLM_BASE_URL` if Ollama runs elsewhere.

**Best local models for this project** (structured JSON generation):

| Model | VRAM/RAM | Quality | Speed |
|---|---|---|---|
| `qwen2.5:14b` | ~9GB | Best JSON at this size | Fast |
| `llama3.1:8b` | ~5GB | Decent | Very fast |
| `mistral-small:22b` | ~13GB | Strong instruction following | Medium |
| `qwen2.5:32b` | ~20GB | Excellent JSON | Slower |

> **Tip:** Set `LLM_JSON_MODE=true` for Ollama models to force valid JSON output.

### Azure OpenAI

```env
LLM_MODEL=gpt-4o
LLM_API_KEY=your-azure-key
LLM_BASE_URL=https://my-resource.openai.azure.com/
LLM_API_VERSION=2024-02-01
```

### Anthropic Claude

Anthropic provides a native [OpenAI-compatible endpoint](https://docs.anthropic.com/en/api/openai-sdk). The `anthropic/` prefix auto-routes to the correct base URL.

```env
LLM_MODEL=anthropic/claude-sonnet-4-20250514
LLM_API_KEY=sk-ant-...
```

Other Claude models: `anthropic/claude-opus-4-6`, `anthropic/claude-haiku-4-5-20251001`.

### Any OpenAI-compatible endpoint

```env
LLM_MODEL=my-model-name
LLM_BASE_URL=http://my-proxy:4000/v1
LLM_API_KEY=key-if-required
```

### Switching at runtime

Change the values in `.env` and restart the backend. No code changes needed. No rebuild needed.

---

## Deploy to OpenShift

### Quick deploy (one command)

```bash
# Build and push the Docker image first
docker build --build-arg FRONTEND_URL=https://excel-copilot.apps.my-cluster.example.com -t excel-ai-copilot .
docker push <registry>/excel-ai-copilot:latest

# Deploy to OpenShift
./openshift/deploy.sh "your-llm-api-key" "<registry>/excel-ai-copilot:latest"
```

The script will:
1. Create the secret with your LLM API key
2. Apply all Kubernetes manifests (PVC, ConfigMap, Deployment, Service, Route)
3. Set the image and wait for rollout
4. Print the app URL and manifest URL for Excel installation

### Step-by-step deployment

#### Step 1 -- Build the Docker image

```bash
docker build \
  --build-arg FRONTEND_URL=https://excel-copilot.apps.my-cluster.example.com \
  -t excel-ai-copilot:latest \
  .
```

`FRONTEND_URL` is baked into `manifest.xml` at build time. It must match the Route URL where users will access the add-in.

#### Step 2 -- Push the image

```bash
# GitLab Registry (CI does this automatically):
docker push registry.gitlab.com/your-group/excel-copilot:latest

# Or Quay.io:
docker push quay.io/your-user/excel-ai-copilot:latest

# Or OpenShift built-in registry:
oc registry login
REGISTRY=$(oc get route default-route -n openshift-image-registry -o jsonpath='{.spec.host}')
docker tag excel-ai-copilot:latest $REGISTRY/<namespace>/excel-ai-copilot:latest
docker push $REGISTRY/<namespace>/excel-ai-copilot:latest
```

#### Step 3 -- Configure

Edit `openshift/configmap.yaml`:
- Set `LLM_MODEL` to your chosen provider
- Set `CORS_ORIGINS` to your Route URL

#### Step 4 -- Create secrets and deploy

```bash
oc login https://your-cluster.example.com
oc project my-namespace

# Create secret
oc create secret generic excel-copilot-secrets \
  --from-literal=LLM_API_KEY=your-api-key-here

# Apply all manifests
oc apply -f openshift/

# Set your image
oc set image deployment/excel-copilot excel-copilot=<registry>/excel-ai-copilot:latest

# Wait for rollout
oc rollout status deployment/excel-copilot --timeout=180s
```

#### Step 5 -- Verify

```bash
# Get the route URL
oc get route excel-copilot

# Test endpoints
curl https://excel-copilot.apps.my-cluster.example.com/health
curl https://excel-copilot.apps.my-cluster.example.com/ready
```

### Deploying updates

```bash
docker build --build-arg FRONTEND_URL=https://... -t excel-ai-copilot:latest .
docker push <registry>/excel-ai-copilot:latest
oc rollout restart deployment/excel-copilot
```

Or use GitLab CI -- push to `main` and click the manual "Deploy" button in the pipeline.

---

## Installing the add-in in Excel

Once deployed, the server serves the `manifest.xml` file that Excel needs to load the add-in.

### For yourself (sideloading)

1. Open Excel (desktop or web)
2. **Insert > Add-ins > Upload My Add-in**
3. Enter the manifest URL:
   - Local dev: `https://localhost:3000/manifest.xml`
   - Production: `https://excel-copilot.apps.my-cluster.example.com/manifest.xml`
4. The **AI Copilot** tab appears in the Excel ribbon
5. Click "Open Copilot" to open the task pane

> **Note:** Sideloaded add-ins persist per-device. If you clear your Office cache or switch devices, you'll need to sideload again.

### Making the ribbon permanent (org-wide deployment via Microsoft 365 Admin Center)

To make the add-in appear automatically for all users in your organization:

1. Log in to [admin.microsoft.com](https://admin.microsoft.com)
2. Go to **Settings > Integrated apps > Upload custom apps**
3. Select **Provide link to manifest file**
4. Enter your manifest URL: `https://excel-copilot.apps.my-cluster.example.com/manifest.xml`
5. Click **Next** and assign to:
   - **Entire organization** — everyone gets it
   - **Specific users/groups** — targeted rollout
6. Click **Deploy**

The add-in will appear in all assigned users' Excel ribbon within **24 hours** (typically faster). Users do **not** need to install anything — it just appears automatically.

### Making it permanent for yourself (without admin access)

If you don't have Microsoft 365 admin access:

1. Sideload the add-in as described above
2. Once loaded, **right-click the "AI Copilot" ribbon tab** > **Pin to ribbon** (Excel desktop)
3. For Excel on the web: the sideloaded add-in persists in your browser profile

> **Tip:** On Excel desktop (Windows), sideloaded add-ins persist in `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`. They survive Excel restarts but not Office reinstalls.

### SharePoint catalog (alternative to Admin Center)

For organizations that prefer SharePoint:

1. Create an [App Catalog](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog) site collection
2. Upload `manifest.xml` to the catalog
3. Users find the add-in via **Insert > Add-ins > My Organization**

---

## CI/CD pipeline

The GitLab CI/CD pipeline runs automatically on pushes to `main` and merge requests.

### Stages

| Stage | Jobs | Purpose |
|---|---|---|
| **lint** | `lint:frontend`, `lint:backend` | ESLint, TypeScript, Ruff, MyPy |
| **build** | `build:frontend` | Webpack production build |
| **test** | `test:frontend`, `test:backend` | Jest, pytest |
| **security** | `security:frontend`, `security:backend` | npm audit, pip-audit |
| **docker** | `docker:build` | Build + push Docker image to GitLab registry |
| **deploy** | `deploy:openshift` | Manual trigger — applies manifests and updates image |

### Required CI/CD variables

Set these in **GitLab > Settings > CI/CD > Variables**:

| Variable | Required | Description |
|---|---|---|
| `LLM_MODEL` | Yes | Model string (e.g. `gpt-4o`, `gemini/gemini-2.0-flash`) |
| `LLM_API_KEY` | Yes | API key (mark as **masked** + **protected**) |
| `FRONTEND_URL` | Yes | Public URL of the deployed add-in |
| `OPENSHIFT_SERVER` | For deploy | Cluster API URL |
| `OPENSHIFT_TOKEN` | For deploy | Service account token |
| `OPENSHIFT_NS` | For deploy | Target namespace |

### Workflow

1. Push code to a branch
2. Create a merge request — lint, build, test, security stages run automatically
3. Merge to `main` — docker stage builds and pushes the image
4. Click **Deploy** in the pipeline UI — deploys to OpenShift

---

## Security

The application implements the following security measures:

- **CORS**: Explicit origins only (no wildcard)
- **Rate limiting**: 15 req/min on chat, 30 req/min on feedback (per IP)
- **Security headers**: CSP, HSTS, X-Frame-Options, X-Content-Type-Options
- **Input validation**: Pydantic Field constraints on all request models
- **Error sanitization**: Generic errors to clients, full stack traces logged server-side
- **Role restriction**: Only `user`/`assistant` roles accepted in conversation history
- **Debug mode**: Disabled by default in production (no `/docs` or `/redoc`)
- **TLS**: Edge termination via OpenShift Route with HTTPS redirect
- **Secrets**: LLM API key injected via OpenShift Secrets, never in images

See [SECURITY_CHECKLIST.md](SECURITY_CHECKLIST.md) for the full checklist.

---

## Migrating to a production database

The project uses **SQLite** + **ChromaDB** (file-based) by default. This works for single-instance deployments. For multi-replica setups, migrate to **PostgreSQL + pgvector**.

### When to migrate

- Running multiple backend replicas
- Need shared state across instances
- Want proper backup/restore and monitoring

### PostgreSQL + pgvector migration

1. Replace `aiosqlite` with `asyncpg` + `sqlalchemy[asyncio]` in `requirements.txt`
2. Update `backend/app/db.py` to use SQLAlchemy async engine
3. Replace ChromaDB with pgvector in `backend/app/services/chroma_client.py`
4. Set `DATABASE_URL=postgresql+asyncpg://user:pass@host/db` in environment

The table schemas remain identical. The `sentence-transformers` model stays the same — only the storage backend changes.

### Files to modify

| File | Change |
|---|---|
| `backend/app/db.py` | SQLite -> asyncpg/SQLAlchemy |
| `backend/app/services/chroma_client.py` | ChromaDB -> pgvector |
| `backend/app/services/capability_store.py` | Collection calls -> SQL |
| `backend/app/services/example_store.py` | Collection calls -> SQL |
| `backend/app/config.py` | Add `database_url` setting |

---

## Configuration reference

All settings are environment variables. Copy `.env.example` to `.env` in the project root.

### Core settings

| Variable | Default | Description |
|---|---|---|
| `LLM_MODEL` | `gpt-4o` | Model string (e.g. `gpt-4o`, `gemini/gemini-2.0-flash`, `ollama/qwen2.5:14b`) |
| `LLM_API_KEY` | _(empty)_ | API key for your LLM provider |
| `LLM_BASE_URL` | _(empty)_ | Custom API base URL (Ollama, Azure, proxy) |
| `LLM_API_VERSION` | _(empty)_ | Azure OpenAI API version only |
| `LLM_MAX_TOKENS` | `4096` | Max tokens per LLM response |
| `LLM_TEMPERATURE` | `0.1` | Lower = more deterministic plans |
| `LLM_JSON_MODE` | `false` | Force JSON output (recommended for Ollama/Qwen) |

### Embedding / vector search

| Variable | Default | Description |
|---|---|---|
| `EMBEDDING_MODEL` | `all-MiniLM-L6-v2` | Sentence-transformers model for embeddings |
| `CHROMA_PERSIST_DIR` | `backend/data/chroma/` | ChromaDB storage directory |
| `CAPABILITY_TOP_K` | `10` | How many capabilities to include per query |
| `FEW_SHOT_TOP_K` | `5` | How many few-shot examples to retrieve per query |

### Server

| Variable | Default | Description |
|---|---|---|
| `HOST` | `0.0.0.0` | Bind address |
| `PORT` | `8000` | Port (8080 in Docker/OpenShift) |
| `DEBUG` | `true` | Enables auto-reload + API docs |
| `CORS_ORIGINS` | `["https://localhost:3000"]` | Allowed CORS origins (explicit list, no wildcards) |

### Deployment

| Variable | Default | Description |
|---|---|---|
| `OPENSHIFT` | `false` | Production mode (static serving, security headers) |
| `SERVE_STATIC` | `false` | Serve built frontend from FastAPI |
| `STATIC_DIR` | `./static` | Path to built frontend files |
| `FRONTEND_URL` | `https://localhost:3000` | **Build-time only** -- baked into manifest.xml |
