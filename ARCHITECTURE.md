# Architecture

This document describes the actual architecture of Excel AI Copilot as it is shipping.
See [README.md](README.md) for install / deploy instructions.

---

## 1. High-level flow

```
Excel (Office.js)
      │
      ▼
Frontend (React + TypeScript, task pane)
  │  builds workbook snapshot, sends chat turn
  │  POST /api/chat/stream   (SSE — primary path)
  │  POST /api/chat          (secondary, non-streaming)
      ▼
Backend (FastAPI)
  ├── Capability store  (ChromaDB + all-MiniLM-L6-v2)   →  top-K relevant actions
  ├── Example store     (ChromaDB + all-MiniLM-L6-v2)   →  top-K few-shot examples
  ├── chat_service      →  single LLM call → routes to message | plans
  ├── Pydantic validator                                →  rejects bad plans
  └── Feedback DB (SQLite)                              →  logs every interaction
      ▼
ChatResponse (message | plan | plans)
      ▼
Frontend executor runs each PlanStep via Office.js capability handlers
```

A second, independent pipeline (`/api/analyze`) runs pandas / rapidfuzz /
sentence-transformers tools server-side for analytical questions. It is
fully implemented and tested but not currently called by the shipped UI.

---

## 2. Single-call chat architecture

`POST /api/chat/stream` is the primary endpoint. One LLM call decides
whether to answer conversationally or produce 1-3 execution plan options.

`ChatResponse.responseType ∈ { "message", "plan", "plans" }`:

| responseType | UI renders                                 |
|--------------|---------------------------------------------|
| `message`    | Plain chat reply (Q&A, clarifications)     |
| `plans`      | 1-3 option cards; user picks one to apply  |
| `plan`       | Single plan (backward-compat path)         |

The LLM is grounded via four injected blocks in the user message:

1. **Clean user text** — workbook-qualified range tokens (`[[WB.xlsx]Sheet!A:A]]`)
   are rewritten to clean form (`[[Sheet!A:A]]`).
2. **Date context** — `Current date: ...` + `User date format: dd/mm/yyyy|mm/dd/yyyy`
   derived from the user's locale.
3. **Sheet metadata** — active sheet, workbook name, used-range end cell.
4. **Workbook snapshot** — per-sheet headers, inferred dtypes, sample rows,
   row/col counts, **anchor cell**, and **used-range address**. This is the
   single largest grounding signal: without it the planner has to guess at
   column names.

**Offset-table awareness:** the snapshot reports `data starts at <cell>` and
sample-row labels use sheet-absolute row numbers (e.g. `row 6:` for an anchor
at C5), so the LLM never assumes A1.

**Retry with failure feedback:** if attempt 1 fails to parse/validate, the
retry prompt injects the exact error (e.g. `"unknown action 'splitColum'"`)
so the LLM can correct its own mistakes.

### Streaming (SSE)

`/api/chat/stream` emits three event types:

| Event   | Purpose                                                |
|---------|--------------------------------------------------------|
| `chunk` | Partial LLM token — rendered live as a "thinking" bar  |
| `reset` | Clear the partial preview; retry attempt is starting   |
| `done`  | Final `ChatResponse` object                            |

---

## 3. The plan schema

Canonical source: `backend/app/models/plan.py` + `frontend/src/engine/types.ts`.
Both sides have 51 `StepAction` entries. Backend `ACTION_PARAM_MODELS` maps
each action to its Pydantic params model.

```
ExecutionPlan
├── planId, createdAt, userRequest, summary
├── confidence (0-1), preserveFormatting (bool), warnings?
└── steps: PlanStep[]
              ├── id, description, action (StepAction)
              ├── params (action-specific, validated per-action)
              └── dependsOn? (step ids)
```

### Capability categories (51 total)

| Category          | Actions |
|-------------------|---------|
| Read / write      | readRange, writeValues, writeFormula, bulkFormula |
| Match / aggregate | matchRecords, groupSum, subtotals, consolidateRanges |
| Tables / pivots   | createTable, createPivot, applyFilter, sortRange, addSlicer |
| Charts / visuals  | createChart, addSparkline, insertPicture, insertShape, insertTextBox |
| Conditional / format | addConditionalFormat, formatCells, setNumberFormat, mergeCells, clearRange |
| Text / patterns   | cleanupText, findReplace, extractPattern, splitColumn, categorize |
| Structural        | addSheet, renameSheet, deleteSheet, copySheet, protectSheet, hideShow |
| Data reshape      | unpivot, crossTabulate, transpose, fillBlanks |
| Validation        | addValidation, removeDuplicates, compareSheets |
| Misc              | freezePanes, autoFitColumns, insertDeleteRows, addComment, addHyperlink, groupRows, setRowColSize, copyPasteRange, pageLayout, namedRange |

### Adding a new capability

1. `backend/app/models/plan.py` — add to `StepAction`, create Pydantic params model, add to `ACTION_PARAM_MODELS`.
2. `backend/app/services/planner.py` — add to `CAPABILITY_DESCRIPTIONS` (this is what the LLM sees).
3. `backend/app/services/capability_store.py` — add example user queries to `CAPABILITY_EXAMPLES`.
4. `frontend/src/engine/types.ts` — add to `StepAction`, `StepParams` union, params interface.
5. `frontend/src/engine/capabilities/<name>.ts` — handler.
6. `frontend/src/engine/capabilities/index.ts` — register.
7. `frontend/src/engine/validator.ts` — add required-field check.
8. Restart backend (ChromaDB re-indexes on startup).

---

## 4. Two-layer validation

| Layer      | Location                              | What it checks                                      |
|------------|---------------------------------------|-----------------------------------------------------|
| Schema     | Pydantic models (`models/plan.py`)    | Action enum, param types, required fields          |
| Business   | `services/validator.py`               | Duplicate IDs, missing deps, dep cycles, ranges    |
| Client     | `frontend/src/engine/validator.ts`    | Same business rules before Office.js execution     |

A plan must pass both Pydantic (automatic, at `ExecutionPlan(**data)`) and
`validate_plan()` (explicit, inside the chat router) before reaching the frontend.

---

## 5. Vector retrieval (grounding)

Two ChromaDB collections backed by the bundled `all-MiniLM-L6-v2` model:

| Collection         | Contents                                                    | Used by                          |
|--------------------|-------------------------------------------------------------|----------------------------------|
| `capabilities`     | 51 actions × several example user queries each              | `search_capabilities(query)` → top-K action names → prompt filter |
| `examples`         | Curated + user-applied assistant responses                  | `search_examples(query)` → few-shot messages injected into prompt |

**Self-improvement loop:** when the user clicks "Apply" on a plan option,
the feedback endpoint promotes that (user_message, assistant_response) pair
into the `examples` collection, so similar future queries get it as a
few-shot example.

Embeddings model is loaded from `backend/models/` at startup — bundled in
the repo (~87 MB), no network calls required. The project runs fully
air-gapped when paired with Ollama.

---

## 6. Conversation persistence

SQLite (file-based) with two tables:

| Table            | Columns                                            |
|------------------|----------------------------------------------------|
| `conversations`  | id, title, created_at, updated_at                  |
| `conv_messages`  | id, conversation_id, role, content, range_tokens, plan, execution, progress_log, timestamp |

Endpoints:

| Method | Path                                          | Purpose                          |
|--------|-----------------------------------------------|----------------------------------|
| GET    | `/api/conversations`                          | List all                         |
| GET    | `/api/conversations/:id`                      | Get one (with all messages)      |
| PATCH  | `/api/conversations/:id`                      | Rename                           |
| DELETE | `/api/conversations/:id`                      | Delete                           |
| DELETE | `/api/conversations/:id/last`                 | Pop last user+assistant exchange (undo support) |
| PATCH  | `/api/conversations/:id/messages/:msg_id`     | Attach execution state + progress log |

Undo: when the user hits "Undo" after applying a plan, the frontend pops the
captured cell snapshots (see §8), restores the sheet, and POSTs to
`/:id/last` so the conversation stays in sync.

---

## 7. Frontend architecture

```
frontend/src/
├── index.tsx                   # Office.onReady entry point
├── taskpane/
│   ├── App.tsx                 # root
│   ├── workbookSnapshot.ts     # builds the snapshot injected into every chat turn
│   ├── components/             # ChatPanel, HistoryDrawer, PlanCard, PresetMenu, ...
│   └── hooks/
│       └── useChat.ts          # chat state, streaming, snapshots, undo
├── engine/
│   ├── types.ts                # StepAction + params (mirror of backend/models/plan.py)
│   ├── validator.ts            # client-side plan validation
│   ├── capabilityRegistry.ts   # action → handler map
│   ├── capabilities/           # 51 Office.js handler files
│   └── executor.ts             # runs plan, calls snapshot.ts before each mutating step
├── services/
│   └── api.ts                  # fetch() wrappers for /api/chat[/stream], /api/feedback, /api/conversations
├── shared/
│   └── snapshot.ts             # LIFO stack (MAX_SNAPSHOTS=20) for undo
└── commands/
    └── commands.ts             # Office ribbon stub (undo lives in chat UI)
```

### Range tokens

Users can splice `[[Sheet1!A1:C20]]` tokens into their message. Ctrl+V with
an active Excel selection inserts a token for the current selection at the
cursor. Tokens preserve Hebrew / non-ASCII sheet names verbatim via
`normalizeAddress` + `quoteSheetInRef`.

### Snapshotting (undo)

`snapshot.ts` holds a LIFO stack of `{ range, values, numberFormats }` triples.
Before every mutating step, `captureSnapshotBatched` calls `getUsedRange(false)`
on the target range(s). Undo pops, restores, and notifies the backend.

### Presets

Stored in browser `localStorage` only — per-user, no server state. Saves
user-message + assistant-response pairs; a heuristic picks the **real
instruction** (not a "yes"/"continue" confirmation) when saving.

---

## 8. Backend architecture

```
backend/
├── main.py                     # FastAPI app, CORS, rate limiting, static serving
├── app/
│   ├── config.py               # Pydantic BaseSettings (.env driven)
│   ├── db.py                   # aiosqlite — feedback + conversations tables
│   ├── models/                 # Pydantic request/response schemas
│   ├── routers/                # chat.py, analyze.py, feedback.py, conversations.py
│   ├── services/
│   │   ├── chat_service.py     # single-LLM-call orchestrator (main path)
│   │   ├── planner.py          # CAPABILITY_DESCRIPTIONS + extract_json
│   │   ├── validator.py        # server-side plan validator (two-layer)
│   │   ├── llm_client.py       # OpenAI SDK wrapper (auto-routes per model prefix)
│   │   ├── capability_store.py # ChromaDB collection + search_capabilities()
│   │   ├── example_store.py    # ChromaDB collection + search_examples()
│   │   ├── clarification_service.py
│   │   ├── explanation_service.py
│   │   └── matching_service.py
│   ├── planner/                # AnalyticalPlanner (second pipeline, /api/analyze)
│   ├── orchestrator/           # tool-chain runner for AnalyticalPlan
│   └── tools/                  # pandas/rapidfuzz/sentence-transformers tools
└── tests/                      # pytest suites (68 tests)
```

### LLM client routing (`llm_client.py`)

The `LLM_MODEL` prefix auto-routes to the right base URL:

| Prefix          | Base URL                                         |
|-----------------|--------------------------------------------------|
| `gemini/`       | `https://generativelanguage.googleapis.com/v1beta/openai/` |
| `anthropic/`    | `https://api.anthropic.com/v1/`                  |
| `cohere/`       | `https://api.cohere.ai/compatibility/v1/`        |
| `ollama/`       | `http://localhost:11434/v1/`                     |
| (no prefix)     | OpenAI default                                   |
| `LLM_BASE_URL` set | overrides everything                          |

All providers use the OpenAI SDK with identical message format.

### Streaming handler

`chat_stream()` is an async generator. `_stream_attempt()` is an inner
async generator that yields SSE chunks and side-effects the result via
`nonlocal`. On failure it emits a `reset` event before retrying with a
stripped-down prompt.

---

## 9. Security

| Layer            | Implementation                                            |
|------------------|-----------------------------------------------------------|
| CORS             | Explicit origins only (no wildcard)                       |
| Rate limiting    | slowapi: 15 req/min on chat, 30 req/min on feedback       |
| Security headers | CSP, HSTS, X-Frame-Options, X-Content-Type-Options        |
| Input validation | Pydantic `Field` constraints on every request model       |
| Error sanitization | Generic errors to clients, full stack traces logged     |
| Role restriction | Only `user`/`assistant` roles accepted in history         |
| Secrets          | LLM API key via OpenShift Secret, never in image          |
| TLS              | Edge termination via OpenShift Route with HTTPS redirect  |

See [SECURITY_CHECKLIST.md](SECURITY_CHECKLIST.md).
