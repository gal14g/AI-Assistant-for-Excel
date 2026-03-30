# Excel AI Copilot

A Microsoft Excel Office Add-in that lets you control your spreadsheet with natural language. Powered by any LLM via LiteLLM (OpenAI, Anthropic Claude, Ollama, Azure, Google Gemini, AWS Bedrock, and more).

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
9. [Migrating to a production database](#migrating-to-a-production-database)
10. [Configuration reference](#configuration-reference)

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

---

## Architecture overview

```
User's Excel
    |
    | Office.js (HTTPS)
    v
Frontend (React + TypeScript)
    |  taskpane UI, execution engine, capability handlers
    |  POST /api/chat
    v
Backend (FastAPI + LiteLLM)
    |
    |-- Vector search (ChromaDB + all-MiniLM-L6-v2)
    |   |-- Capability store: finds the ~10 most relevant actions for each query
    |   |-- Example store: retrieves the ~5 most relevant few-shot examples
    |
    |-- LLM call (any provider via LiteLLM)
    |   Returns 2-3 alternative plan options as structured JSON
    |
    |-- Feedback DB (SQLite)
    |   Logs every interaction + user's choice (applied / dismissed)
    |   Applied plans are promoted into the example store automatically
    |
    v
Response: plan options -> user picks one -> frontend executes via Office.js
```

---

## Project structure

```
excel-ai-copilot/
|
|-- .env.example              All environment variables documented
|-- .env                      Your local config (git-ignored)
|-- .gitignore
|-- .gitlab-ci.yml            GitLab CI/CD pipeline
|-- Dockerfile                Multi-stage build (frontend + backend in one image)
|-- docker-compose.yml        Local Docker testing
|-- README.md                 This file
|
|-- backend/                  Python FastAPI backend
|   |-- main.py               App entry point, startup hooks, CORS, static serving
|   |-- requirements.txt      Python dependencies
|   |-- app/
|   |   |-- config.py         Settings loaded from .env (Pydantic BaseSettings)
|   |   |-- db.py             SQLite feedback database (aiosqlite)
|   |   |
|   |   |-- routers/          API endpoint definitions
|   |   |   |-- chat.py         POST /api/chat — main conversational endpoint
|   |   |   |-- plan.py         POST /api/plan — standalone plan generation
|   |   |   |-- feedback.py     POST /api/feedback — logs user apply/dismiss
|   |   |   |-- analyze.py      POST /api/analyze — analytical pipeline (data analysis)
|   |   |
|   |   |-- services/         Business logic
|   |   |   |-- chat_service.py       Chat pipeline: prompt building, LLM call, response parsing
|   |   |   |-- planner.py            Plan generation, capability descriptions, JSON extraction
|   |   |   |-- validator.py          Server-side plan validation
|   |   |   |-- capability_store.py   ChromaDB vector search for relevant actions
|   |   |   |-- example_store.py      ChromaDB vector search for few-shot examples
|   |   |   |-- chroma_client.py      Shared ChromaDB client + embedding function
|   |   |   |-- matching_service.py   Column matching utilities
|   |   |   |-- explanation_service.py  Results explanation generation
|   |   |   |-- clarification_service.py  Clarification request handling
|   |   |
|   |   |-- models/           Pydantic request/response schemas
|   |   |   |-- chat.py         ChatRequest, ChatResponse, PlanOption
|   |   |   |-- plan.py         ExecutionPlan, PlanStep, StepAction enum, param models
|   |   |   |-- request.py      PlanRequest, PlanResponse, ValidationResponse
|   |   |   |-- analytical_plan.py  Analytical pipeline models
|   |   |   |-- tool_output.py  Tool output models for analysis
|   |   |
|   |   |-- prompts/          LLM system prompt templates
|   |   |   |-- planner_system.txt  Planner system prompt with {CAPABILITIES} placeholder
|   |   |
|   |   |-- orchestrator/     Analytical pipeline orchestration
|   |   |-- planner/          Analytical planner (separate from chat planner)
|   |   |-- tools/            Data analysis tools (matching, aggregation, cleaning, comparison)
|   |
|   |-- data/                 Runtime data (git-ignored)
|       |-- chroma/           ChromaDB vector embeddings (auto-created on first start)
|       |-- feedback.db       SQLite feedback database (auto-created on first start)
|
|-- frontend/                 React + TypeScript Office Add-in
    |-- manifest.xml          Office Add-in manifest (loaded by Excel)
    |-- webpack.config.js     Dev server config with HTTPS + API proxy
    |-- package.json
    |-- tsconfig.json
    |-- src/
        |-- index.tsx         App entry point
        |
        |-- engine/           Execution engine (runs plans inside Excel)
        |   |-- types.ts        All TypeScript types (ExecutionPlan, StepAction, params, etc.)
        |   |-- executor.ts     Executes plans step-by-step via Office.js
        |   |-- validator.ts    Client-side plan validation (param checks per action)
        |   |-- snapshot.ts     Snapshot/rollback for undo support
        |   |-- capabilityRegistry.ts  Registry pattern for capability handlers
        |   |-- capabilities/   One file per Excel action (34 total)
        |       |-- index.ts      Imports all capability handlers to register them
        |       |-- rangeUtils.ts Shared range resolution helper
        |       |-- readRange.ts, writeValues.ts, writeFormula.ts
        |       |-- matchRecords.ts, groupSum.ts
        |       |-- createTable.ts, applyFilter.ts, sortRange.ts
        |       |-- createPivot.ts, createChart.ts
        |       |-- conditionalFormat.ts, cleanupText.ts, removeDuplicates.ts
        |       |-- freezePanes.ts, findReplace.ts, validation.ts
        |       |-- sheetOps.ts (add/rename/delete/copy/protect sheet)
        |       |-- autoFitColumns.ts, mergeCells.ts, setNumberFormat.ts
        |       |-- insertDeleteRows.ts, addSparkline.ts
        |       |-- formatCells.ts, clearRange.ts, hideShow.ts
        |       |-- addComment.ts, addHyperlink.ts, groupRows.ts
        |       |-- setRowColSize.ts, copyPasteRange.ts
        |
        |-- services/         API client
        |   |-- api.ts          fetch wrappers for /api/chat, /api/plan, /api/feedback
        |
        |-- taskpane/         UI layer
        |   |-- App.tsx         Root component
        |   |-- components/
        |   |   |-- ChatPanel.tsx         Main chat interface
        |   |   |-- ChatInput.tsx         Message input with range token support
        |   |   |-- MessageBubble.tsx     Chat message rendering
        |   |   |-- PlanOptionsPanel.tsx  Tab-based multi-option plan selector
        |   |   |-- PlanPreview.tsx       Single plan card (steps, validation, actions)
        |   |   |-- ExecutionTimeline.tsx Step-by-step execution progress
        |   |   |-- SuggestedPrompts.tsx  Quick-start prompt suggestions
        |   |   |-- RangeToken.tsx        Inline range reference display
        |   |-- hooks/
        |       |-- useChat.ts            Chat state, message history, plan options
        |       |-- usePlanExecution.ts   Plan execution, preview, undo
        |       |-- useSelectionTracker.ts  Tracks Excel selection changes
        |
        |-- shared/
        |   |-- planSchema.ts   Shared plan schema utilities
        |
        |-- commands/
            |-- commands.ts     Office ribbon command handlers
```

### Key directories explained

| Directory | Purpose |
|---|---|
| `backend/app/services/` | Core business logic. `chat_service.py` orchestrates the entire chat flow: vector search for capabilities, vector search for few-shot examples, LLM call, response parsing, DB logging. |
| `backend/app/models/` | Pydantic models that define the API contract. `plan.py` is the most important — it defines all 34 `StepAction` types and their parameter schemas. |
| `backend/data/` | Auto-created at runtime. Contains ChromaDB embeddings and the SQLite feedback database. Git-ignored. Deleted safely — regenerated on next startup. |
| `frontend/src/engine/` | The execution engine that runs inside Excel. Each capability in `capabilities/` is a self-contained handler that translates a plan step into Office.js API calls. |
| `frontend/src/taskpane/` | React UI. `ChatPanel.tsx` is the main component. `PlanOptionsPanel.tsx` shows multiple plan options as tabs. |

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
LLM_MODEL=claude-sonnet-4-20250514
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
3. Seed 15 curated few-shot examples into the example store
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

**Exception:** If you add a new capability, you must restart the backend so ChromaDB re-indexes. The store detects the document count mismatch and re-indexes automatically.

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

### Do I need to rebuild Docker every time?

**No.** Docker is only for deploying to OpenShift. Your local development loop is:
```
Edit code -> save -> see changes live in Excel immediately
```

---

## Switching LLM providers

All LLM configuration is in `.env` (project root). The backend uses [LiteLLM](https://docs.litellm.ai/docs/providers) which supports 100+ providers with a unified API.

### Anthropic Claude (recommended for quality)

```env
LLM_MODEL=claude-sonnet-4-20250514
LLM_API_KEY=sk-ant-...
```

### OpenAI GPT-4o

```env
LLM_MODEL=gpt-4o
LLM_API_KEY=sk-...
```

### Ollama (free, local, no API key)

First install and run a model:
```bash
ollama pull qwen2.5:14b
ollama serve
```

Then configure:
```env
LLM_MODEL=ollama/qwen2.5:14b
LLM_BASE_URL=http://localhost:11434
LLM_JSON_MODE=true
```

**Best local models for this project** (structured JSON generation):

| Model | VRAM/RAM | Quality | Speed |
|---|---|---|---|
| `qwen2.5:14b` | ~9GB | Best JSON at this size | Fast |
| `llama3.1:8b` | ~5GB | Decent | Very fast |
| `mistral-small:22b` | ~13GB | Strong instruction following | Medium |
| `qwen2.5:32b` | ~20GB | Excellent JSON | Slower |

> **Tip:** Set `LLM_JSON_MODE=true` for Ollama models. This forces the LLM to output valid JSON, reducing parse failures.

### Azure OpenAI

```env
LLM_MODEL=azure/my-deployment-name
LLM_API_KEY=your-azure-key
LLM_BASE_URL=https://my-resource.openai.azure.com/
LLM_API_VERSION=2024-02-01
```

### Google Gemini

```env
LLM_MODEL=gemini/gemini-1.5-pro
LLM_API_KEY=AIza...
```

### AWS Bedrock

```env
LLM_MODEL=bedrock/anthropic.claude-3-sonnet-20240229-v1:0
# Uses AWS credentials from environment or ~/.aws/credentials
```

### LiteLLM proxy / any OpenAI-compatible endpoint

```env
LLM_MODEL=openai/my-model
LLM_BASE_URL=http://my-litellm-proxy:4000
LLM_API_KEY=proxy-key-if-required
```

### Switching at runtime

Change the values in `.env` and restart the backend. No code changes needed. No rebuild needed.

---

## Deploy to OpenShift

### How it works

The Docker image bundles the React frontend and FastAPI backend into a single container. OpenShift provides HTTPS automatically via a Route.

```
User's Excel
    |  HTTPS
    v
OpenShift Route  (TLS terminated here -- free HTTPS)
    |  HTTP
    v
Container :8080
    |-- GET /api/*    ->  FastAPI (AI chat, plan generation, feedback)
    |-- GET /health   ->  health check
    |-- GET /ready    ->  readiness probe (LLM config + vector store + DB)
    +-- GET /*        ->  React build (taskpane.html, manifest.xml, assets)
```

Multiple users can use the add-in simultaneously. FastAPI is fully async -- while one user's LLM request is in-flight, other requests are handled concurrently.

### Prerequisites

- Docker installed and running
- `oc` CLI installed ([download](https://mirror.openshift.com/pub/openshift-v4/clients/ocp/latest/))
- Logged in to your cluster: `oc login https://your-cluster.example.com`
- Access to a container image registry

### Step 1 -- Get your public URL

The URL pattern is: `https://<app-name>.apps.<cluster-domain>`

Example: `https://excel-copilot.apps.my-cluster.example.com`

### Step 2 -- Build the Docker image

```bash
docker build \
  --build-arg FRONTEND_URL=https://excel-copilot.apps.my-cluster.example.com \
  -t excel-ai-copilot:latest \
  .
```

### Step 3 -- Push the image

**OpenShift built-in registry:**
```bash
oc registry login
REGISTRY=$(oc get route default-route -n openshift-image-registry -o jsonpath='{.spec.host}')
docker tag excel-ai-copilot:latest $REGISTRY/<your-namespace>/excel-ai-copilot:latest
docker push $REGISTRY/<your-namespace>/excel-ai-copilot:latest
```

**Quay.io:**
```bash
docker login quay.io
docker tag excel-ai-copilot:latest quay.io/<user>/excel-ai-copilot:latest
docker push quay.io/<user>/excel-ai-copilot:latest
```

### Step 4 -- Create secrets

```bash
oc create secret generic excel-copilot-secrets \
  --from-literal=LLM_API_KEY=your-api-key-here \
  --from-literal=LLM_MODEL=claude-sonnet-4-20250514
```

### Step 5 -- Deploy

```bash
oc new-app --image=quay.io/<user>/excel-ai-copilot:latest --name=excel-copilot
oc set env deployment/excel-copilot OPENSHIFT=true
oc set env deployment/excel-copilot --from=secret/excel-copilot-secrets
oc create route edge excel-copilot --service=excel-copilot --port=8080
oc get route excel-copilot
```

### Step 6 -- Verify

```bash
curl https://excel-copilot.apps.my-cluster.example.com/health
curl https://excel-copilot.apps.my-cluster.example.com/ready
```

### Deploying updates

```bash
docker build --build-arg FRONTEND_URL=https://... -t excel-ai-copilot:latest .
docker push quay.io/<user>/excel-ai-copilot:latest
oc rollout restart deployment/excel-copilot
```

---

## Installing the add-in in Excel

### For yourself (sideloading)

1. Open Excel
2. **Insert > Add-ins > Upload My Add-in**
3. Enter the manifest URL:
   - Local dev: `https://localhost:3000/manifest.xml`
   - OpenShift: `https://excel-copilot.apps.my-cluster.example.com/manifest.xml`
4. The **AI Copilot** tab appears in the Excel ribbon

### For your organisation (Microsoft 365 Admin Center)

1. Log in to [admin.microsoft.com](https://admin.microsoft.com)
2. **Settings > Integrated apps > Upload custom apps**
3. Select **Provide link to manifest** and enter your manifest URL
4. Assign to users, groups, or the entire organisation
5. The add-in rolls out to assigned users within 24 hours

---

## Migrating to a production database

The project currently uses **SQLite** (via `aiosqlite`) for the feedback database and **ChromaDB** (file-based) for vector embeddings. Both store data in `backend/data/`. This is fine for single-instance deployments but won't work for multi-replica setups.

### When to migrate

- You're running multiple backend replicas (horizontal scaling)
- You need shared state across instances
- You want proper backup/restore, replication, or monitoring

### Option 1: PostgreSQL for feedback + pgvector for embeddings

This is the cleanest migration path -- one database for everything.

**1. Install pgvector extension on your PostgreSQL instance:**
```sql
CREATE EXTENSION IF NOT EXISTS vector;
```

**2. Replace `aiosqlite` with `asyncpg`:**

In `requirements.txt`:
```diff
- aiosqlite>=0.20.0
+ asyncpg>=0.29.0
+ sqlalchemy[asyncio]>=2.0.0
```

**3. Update `backend/app/db.py`:**

Replace the SQLite connection with SQLAlchemy async engine:

```python
from sqlalchemy.ext.asyncio import create_async_engine, AsyncSession
from sqlalchemy.orm import sessionmaker

engine = create_async_engine(settings.database_url)  # e.g. "postgresql+asyncpg://user:pass@host/db"
async_session = sessionmaker(engine, class_=AsyncSession)
```

The table schemas remain identical -- just change the connection layer. All `await _db.execute(...)` calls become `await session.execute(text(...))`.

**4. Update `backend/app/config.py`:**

```python
database_url: str = "sqlite+aiosqlite:///data/feedback.db"  # default for local dev
```

For production, set via environment:
```env
DATABASE_URL=postgresql+asyncpg://user:password@postgres-host:5432/excel_copilot
```

**5. Replace ChromaDB with pgvector:**

In `backend/app/services/chroma_client.py`, replace the ChromaDB client with pgvector queries:

```python
# Instead of ChromaDB collection.query(), use:
# SELECT * FROM embeddings ORDER BY embedding <=> $1 LIMIT $2
```

The `sentence-transformers` model stays the same -- you just store and query embeddings in PostgreSQL instead of ChromaDB.

**6. Tables to create in PostgreSQL:**

```sql
-- Feedback tables (same schema as SQLite)
CREATE TABLE interactions ( ... );  -- same columns
CREATE TABLE choices ( ... );        -- same columns
CREATE TABLE few_shot_examples ( ... ); -- same columns

-- Vector tables (replaces ChromaDB)
CREATE TABLE capability_embeddings (
    id TEXT PRIMARY KEY,
    action TEXT NOT NULL,
    document TEXT NOT NULL,
    embedding vector(384)  -- all-MiniLM-L6-v2 produces 384-dim vectors
);

CREATE TABLE example_embeddings (
    id TEXT PRIMARY KEY,
    sqlite_id TEXT NOT NULL,
    source TEXT NOT NULL,
    document TEXT NOT NULL,
    embedding vector(384)
);

-- Create indexes for fast similarity search
CREATE INDEX ON capability_embeddings USING ivfflat (embedding vector_cosine_ops);
CREATE INDEX ON example_embeddings USING ivfflat (embedding vector_cosine_ops);
```

### Option 2: Keep SQLite + separate vector DB

If you prefer managed vector databases:

- **Pinecone**: Replace ChromaDB calls with `pinecone.Index.query()`
- **Weaviate**: Replace with Weaviate client
- **Qdrant**: Replace with Qdrant client

The interface is the same: embed query -> find top-K -> return IDs -> fetch full data from SQL.

### Files that reference the database

| File | What it does | What to change |
|---|---|---|
| `backend/app/db.py` | SQLite connection, all CRUD operations | Replace with asyncpg/SQLAlchemy |
| `backend/app/services/chroma_client.py` | ChromaDB client + embedding function | Replace with pgvector or managed vector DB |
| `backend/app/services/capability_store.py` | Indexes capabilities, searches by query | Change collection calls to SQL queries |
| `backend/app/services/example_store.py` | Indexes few-shot examples, searches by query | Change collection calls to SQL queries |
| `backend/app/config.py` | `feedback_db_path`, `chroma_persist_dir` | Replace with `database_url` |
| `backend/main.py` | Calls `init_db()`, `init_store()`, `init_example_store()` | Update init functions |

---

## Configuration reference

All settings are environment variables. Copy `.env.example` to `.env` in the project root.

### Core settings

| Variable | Default | Description |
|---|---|---|
| `LLM_MODEL` | `claude-sonnet-4-20250514` | LiteLLM model string ([full list](https://docs.litellm.ai/docs/providers)) |
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

### Feedback database

| Variable | Default | Description |
|---|---|---|
| `FEEDBACK_DB_PATH` | `backend/data/feedback.db` | SQLite database file path |

### Server

| Variable | Default | Description |
|---|---|---|
| `HOST` | `0.0.0.0` | Bind address |
| `PORT` | `8000` | Port (use 8080 for OpenShift) |
| `DEBUG` | `true` | Enables uvicorn auto-reload |
| `CORS_ORIGINS` | `["https://localhost:3000"]` | Allowed CORS origins |

### Deployment

| Variable | Default | Description |
|---|---|---|
| `OPENSHIFT` | `false` | Master production switch (serves static files, relaxes CORS) |
| `SERVE_STATIC` | `false` | Serve built frontend from FastAPI |
| `STATIC_DIR` | `./static` | Path to built frontend files |
| `FRONTEND_URL` | `https://localhost:3000` | **Build-time only** -- baked into manifest.xml |
