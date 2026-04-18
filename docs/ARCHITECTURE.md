# Architecture — AI Assistant For Excel

End-to-end flow from the moment a user sends a message to the final Excel
mutation. Every box in the diagrams maps to a real file in the repo.

The system has **two execution frontends** sharing the same planner:

- **Office.js add-in** (taskpane) — runs inside Excel, universal (Windows /
  Mac / Web / iPad), single-workbook scope.
- **MCP server** (`backend/mcp_server.py`) — stdio transport, exposes the
  same 94-action catalog to any MCP client (Claude Desktop, Cursor, Zed,
  Windsurf), executes against the desktop Excel via `xlwings`, supports
  cross-workbook operations the add-in can't do.

Every `StepAction` exists in three mirrored layers: TypeScript handler →
Pydantic params + LLM-side description → Python xlwings handler. Full
parity is verified automatically by the test suites.

---

## High-Level Overview

```mermaid
flowchart TB
    subgraph Excel["Excel Client (Office.js add-in)"]
        UI["ChatPanel / ChatInput / MessageBubble"]
        Hook["useChat + usePlanExecution hooks"]
        API["api.ts — sendChatMessageStream()"]
        Exec["executor.ts — executePlan()"]
        Snap["snapshot.ts — captureSnapshot + InverseOp"]
        Handlers["94 Capability Handlers (TypeScript)"]
        Registry["capabilityRegistry.ts"]
        Utils["utils/ — normalizeString,\nparseDateFlexible,\nparseNumberFlexible,\nmergedCells"]
        Fallbacks["fallbacks/ — dynamicArrayRewrite,\nclipFullColumnRefs"]
    end

    subgraph MCP["MCP Client (Claude Desktop / Cursor / …)"]
        MCPClient["MCP stdio client"]
    end

    subgraph Backend["FastAPI Backend (shared planner)"]
        Router["POST /api/chat/stream (SSE)"]
        ChatSvc["chat_service.py"]
        PromptBuilder["system prompt + few-shot"]
        LLM["llm_client.py"]
        Parser["response parser + validator"]
        Planner["planner.py — CAPABILITY_DESCRIPTIONS"]

        subgraph MCPServer["backend/mcp_server.py (stdio)"]
            MCPTools["list_open_workbooks,\nget_workbook_snapshot,\ngenerate_plan, validate_plan,\nexecute_plan, undo_last, …"]
            XlBridge["xlwings_bridge.py\n(94 Python handlers)"]
        end

        Routers["routers/ — chat, conversations, feedback, analyze"]
    end

    subgraph Intelligence["Semantic Intelligence"]
        CapStore["capability_store.py"]
        ExStore["example_store.py"]
        Embed["sentence-transformers\nparaphrase-multilingual-MiniLM-L12-v2\n(bundled, Hebrew-capable)"]
    end

    subgraph Persistence["Persistence (swappable)"]
        Factory["persistence/factory.py"]
        SQLite["sqlite/repositories.py"]
        Postgres["postgres/repositories.py"]
        VChroma["vector_chroma.py"]
        VPG["vector_pgvector.py"]
    end

    subgraph External["LLM Provider (any OpenAI-compatible)"]
        LLMApi["OpenAI / Azure / Anthropic /\nOllama / Cohere / Gemini"]
    end

    UI --> Hook
    Hook --> API
    API -->|"POST /api/chat/stream"| Router
    Router --> ChatSvc

    MCPClient -->|"stdio"| MCPServer
    MCPServer --> ChatSvc
    MCPServer --> XlBridge

    ChatSvc --> CapStore
    ChatSvc --> ExStore
    CapStore --> Embed
    ExStore --> Embed
    CapStore --> Factory
    ExStore --> Factory

    ChatSvc --> PromptBuilder
    PromptBuilder --> Planner
    PromptBuilder --> LLM
    LLM --> LLMApi

    ChatSvc --> Parser
    Parser --> ChatSvc
    ChatSvc -->|"SSE"| API

    API -->|"ChatResponse"| Hook
    Hook -->|"plan preview + approve"| UI
    UI --> Exec
    Exec --> Snap
    Exec --> Registry
    Registry --> Handlers
    Handlers --> Utils
    Handlers --> Fallbacks

    XlBridge -->|"xlwings COM"| DesktopExcel["Desktop Excel"]
    Handlers -->|"Office.js"| Excel

    Factory --> SQLite
    Factory --> Postgres
    Factory --> VChroma
    Factory --> VPG
```

---

## Detailed Request Flow (add-in mode)

```mermaid
sequenceDiagram
    participant U as User
    participant CP as ChatPanel
    participant UC as useChat Hook
    participant API as api.ts
    participant R as /api/chat/stream
    participant CS as chat_service
    participant Cap as capability_store
    participant Ex as example_store
    participant DB as Vector Store<br/>(Chroma / pgvector)
    participant LLM as LLM Provider
    participant PR as _parse_response
    participant Repo as Repositories<br/>(SQLite / Postgres)
    participant EX as executor.ts
    participant OH as Office.js Handlers
    participant XL as Excel Workbook
    participant FB as /api/feedback

    U->>CP: Types message + optional range selection
    CP->>UC: handleSend(text, rangeTokens)

    Note over UC: Builds ChatRequest:<br/>userMessage, rangeTokens,<br/>workbookSnapshot (headers,<br/>dtypes, 10 sample rows,<br/>anchorCell, usedRangeAddress),<br/>conversationHistory, activeSheet, locale

    UC->>API: sendChatMessageStream(request)
    API->>R: POST /api/chat/stream (SSE)

    R->>CS: chat_stream(ChatRequest)

    rect rgb(230, 245, 255)
        Note over CS,DB: Phase 1 — Semantic Search
        CS->>Cap: search_capabilities(userMessage)
        Cap->>DB: query "capabilities" collection
        DB-->>Cap: top-K action names
        Cap-->>CS: relevant_actions[]
    end

    rect rgb(255, 245, 230)
        Note over CS,Ex: Phase 2 — Build Prompt
        CS->>CS: _build_chat_system_prompt(relevant_actions)
        Note over CS: LANGUAGE RULE at top (line 1);<br/>filter CAPABILITY_DESCRIPTIONS to<br/>only relevant actions; 600+ line<br/>system prompt with schema,<br/>rules, dashboard patterns,<br/>OUT-OF-SCOPE section, Hebrew<br/>directional hints, self-join<br/>formula guidance.

        CS->>Ex: search_examples(userMessage, top_k=5)
        Ex->>DB: query "few_shot_examples" (v4 seeds)
        DB-->>Ex: similar examples
        Ex-->>CS: few-shot message pairs

        CS->>CS: _build_user_content(request)
        Note over CS: Injects: date/time, locale-aware<br/>date format, range tokens,<br/>used range end, workbook<br/>snapshot with ⚠ WARNING
    end

    rect rgb(230, 255, 230)
        Note over CS,LLM: Phase 3 — LLM Streaming
        CS->>LLM: acompletion_stream(messages)
        loop Token by token
            LLM-->>CS: delta text
            CS-->>R: SSE event: {type: "chunk", text}
            R-->>API: chunk
            API-->>UC: onChunk → streaming preview
        end
    end

    rect rgb(255, 230, 245)
        Note over CS,PR: Phase 4 — Parse, Validate, Localize
        CS->>PR: _parse_response(full_text)
        Note over PR: 1. extract_json()<br/>2. Extract messageLocalized +<br/>   optionLabelLocalized from parsed<br/>3. Salvage off-schema outputs<br/>   (bare step, tool_calls, aliases)<br/>4. _fill_plan_defaults()<br/>5. Pydantic validate steps +<br/>   bindings + cycle check

        alt Parse succeeds
            PR-->>CS: ChatResponse
        else Parse fails — Retry
            CS->>LLM: Compact retry prompt (20 lines)
            LLM-->>CS: second attempt
            CS-->>R: SSE event: {type: "reset"}
            CS->>PR: _parse_response(retry_text)
            PR-->>CS: ChatResponse
        else Both fail
            CS-->>R: Friendly error message
        end
    end

    CS->>Repo: _persist_conversation_turn()
    Note over Repo: Save user msg + assistant response<br/>+ plan. Persist messageLocalized<br/>when set (reload preserves Hebrew).

    CS-->>R: SSE event: {type: "done", result}
    R-->>API: final SSE
    API-->>UC: ChatResponse

    alt responseType = "message"
        UC->>CP: Display text (messageLocalized ?? message)
    else responseType = "plans"
        UC->>CP: Show plan option cards (2-3)
        U->>CP: Selects preferred plan
    end

    rect rgb(245, 245, 220)
        Note over U,XL: Phase 5 — Plan Execution
        U->>CP: Clicks "Run" on chosen plan
        CP->>EX: executePlan(plan, callbacks)

        EX->>EX: validator.validate(plan)
        EX->>EX: topologicalSort(steps)

        loop For each step (dependency order)
            EX->>EX: captureSnapshot(affected ranges)
            Note over EX: Empty snapshot for structural<br/>ops — handler attaches InverseOp
            EX->>OH: handler(context, params)
            Note over OH: Handler runs:<br/>- normalizeString / parseNumberFlexible /<br/>  parseDateFlexible for dirty data<br/>- ensureUnmerged (5 wired handlers)<br/>- clipFullColumnRefs (writeFormula,<br/>  spillFormula)<br/>- dynamicArrayRewrite on pre-365 Excel<br/>- registerInverseOp for structural ops<br/>- expandSnapshotFootprint for wide writes
            OH->>XL: Office.js API calls
            XL-->>OH: result
            OH-->>EX: StepResult
            EX-->>CP: onStepComplete callback
        end

        alt Step fails
            EX->>EX: rollbackPlan() — restore cell values +<br/>apply inverse ops in reverse order
        end

        EX-->>CP: ExecutionState
    end

    CP->>FB: sendFeedback(interactionId, planId, "applied")
    FB->>Repo: Log feedback
    FB->>Ex: add_user_example() — promote to few-shot pool
    Ex->>DB: Upsert into "few_shot_examples"
```

---

## Detailed Request Flow (MCP mode)

```mermaid
sequenceDiagram
    participant Client as MCP Client<br/>(Claude Desktop / Cursor)
    participant MCP as mcp_server.py (stdio)
    participant CS as chat_service.py<br/>(shared with add-in)
    participant LLM as LLM Provider
    participant XB as xlwings_bridge.py
    participant PyReg as Python capability_registry
    participant PyHandlers as 94 Python handlers
    participant XW as xlwings (COM / AppleScript)
    participant Excel as Desktop Excel

    Client->>MCP: call list_open_workbooks
    MCP->>XW: xw.apps[*].books[*]
    XW-->>MCP: [{name, sheets, isActive}, …]
    MCP-->>Client: workbook inventory

    Client->>MCP: call generate_plan(user_message, snapshot)
    MCP->>CS: chat(request)
    CS->>LLM: acompletion(messages)
    LLM-->>CS: JSON ChatResponse
    CS-->>MCP: ChatResponse
    MCP-->>Client: plan (can reference cross-workbook ranges)

    Client->>MCP: call validate_plan(plan)
    MCP-->>Client: issues[]

    Client->>MCP: call execute_plan(plan, target_workbook)
    MCP->>XB: XlwingsExecutor.execute_plan()

    loop For each step
        XB->>XB: capture snapshot (xlwings values + number_format)
        XB->>PyReg: dispatch action
        PyReg->>PyHandlers: handler(ctx, params)
        PyHandlers->>XW: xw.range / sheet / pivot / chart
        XW->>Excel: COM / AppleScript
        Excel-->>XW: result
        XW-->>PyHandlers: Range / value
        PyHandlers-->>XB: StepResult dict
    end

    XB-->>MCP: ExecutionState
    MCP-->>Client: result

    Client->>MCP: call undo_last
    MCP->>XB: undo_last()
    XB->>XW: restore snapshot
```

---

## Component Map

```mermaid
flowchart LR
    subgraph Frontend["frontend/src/"]
        direction TB
        subgraph Components["taskpane/components/"]
            ChatPanel["ChatPanel.tsx"]
            ChatInput["ChatInput.tsx"]
            MessageBubble["MessageBubble.tsx"]
            PlanPreview["PlanPreview.tsx"]
            PlanOptions["PlanOptionsPanel.tsx"]
            ExecTimeline["ExecutionTimeline.tsx"]
            HistoryDrawer["HistoryDrawer.tsx"]
        end
        subgraph Hooks["taskpane/hooks/"]
            useChat["useChat.ts"]
            usePlanExec["usePlanExecution.ts"]
            useSelection["useSelectionTracker.ts"]
        end
        subgraph Engine["engine/"]
            Executor["executor.ts"]
            Snapshot["snapshot.ts — InverseOp, expandSnapshotFootprint"]
            Validator["validator.ts"]
            CapReg["capabilityRegistry.ts"]
            ApiSupport["apiSupport.ts — ExcelApi version gating"]
            subgraph Caps["capabilities/ (94 actions)"]
                writeValues["writeValues, writeFormula, readRange"]
                matchRecords["matchRecords, fuzzyMatch, lookupAll"]
                createChart["createChart, createPivot, createTable"]
                newOps["lateralSpreadDuplicates,\nextractMatchedToNewRow,\nreorderRows, fillSeries, …\n(see full list in Adding_Actions.md)"]
            end
            subgraph CapUtils["engine/utils/"]
                parseDate["parseDateFlexible.ts"]
                parseNum["parseNumberFlexible.ts"]
                normStr["normalizeString.ts"]
                mergeSafe["mergedCells.ts"]
            end
            subgraph CapFallbacks["capabilities/fallbacks/"]
                clipFull["clipFullColumnRefs.ts"]
                dynArr["dynamicArrayRewrite.ts —\nLET, XLOOKUP, FILTER, SORT,\nUNIQUE, SEQUENCE, INDEX/MATCH"]
            end
        end
        subgraph Services["services/"]
            apiTS["api.ts"]
        end
        subgraph Snapshots["taskpane/workbookSnapshot.ts"]
            Snap["per-sheet dtype + sample-row extraction"]
        end
        subgraph Shared["shared/planSchema.ts"]
            planSchema["JSON schema for pre-validation"]
        end
    end

    subgraph BackendFiles["backend/"]
        direction TB
        mainPy["main.py — FastAPI app entry"]
        mcpPy["mcp_server.py — stdio MCP server"]
        subgraph Routers["app/routers/"]
            chatRouter["chat.py"]
            fbRouter["feedback.py"]
            convRouter["conversations.py"]
            analyzeRouter["analyze.py — deterministic pipeline"]
        end
        subgraph ServicesBE["app/services/"]
            chatService["chat_service.py"]
            llmClient["llm_client.py"]
            capStore["capability_store.py"]
            exStore["example_store.py"]
            planner["planner.py — CAPABILITY_DESCRIPTIONS"]
            validatorBE["validator.py"]
            matchingSvc["matching_service.py"]
            explainSvc["explanation_service.py"]
            clarifySvc["clarification_service.py"]
        end
        subgraph Models["app/models/"]
            planModel["plan.py — StepAction enum +\n94 Pydantic params +\nExecutionPlan/PlanStep"]
            chatModel["chat.py — ChatRequest/Response"]
            requestModel["request.py"]
            anaModel["analytical_plan.py"]
            toolOut["tool_output.py"]
        end
        subgraph Persistence["app/persistence/"]
            factoryPy["factory.py — URL-scheme dispatch"]
            basePy["base.py — abstract Repos + VectorStore"]
            embedPy["embedding.py — model loader"]
            sqliteRepos["sqlite/repositories.py"]
            pgRepos["postgres/repositories.py"]
            vChroma["vector_chroma.py"]
            vPG["vector_pgvector.py"]
        end
        subgraph Execution["app/execution/ (MCP side)"]
            execBase["base.py — PlanExecutor ABC, ExecutorContext"]
            execReg["capability_registry.py"]
            xlBridge["xlwings_bridge.py — XlwingsExecutor"]
            execRange["range_utils.py — parse_address, cross-workbook"]
            execSnap["snapshot.py — xlwings-side snapshots"]
            subgraph PyCaps["capabilities/ (94 Python handlers)"]
                pyHandlers["matching parity with frontend"]
            end
            subgraph ExecUtils["execution/utils/"]
                pyParseDate["parse_date_flexible.py"]
                pyParseNum["parse_number_flexible.py"]
                pyNormStr["normalize_string.py"]
            end
        end
        subgraph Orchestrator["app/orchestrator/"]
            orch["orchestrator.py — deterministic analytical pipeline"]
            execCtx["execution_context.py"]
            orchValid["validators.py"]
        end
        subgraph AnaTools["app/tools/"]
            sheetTool["sheet_tools.py"]
            matchTool["matching_tools.py"]
            cleanTool["cleaning_tools.py"]
            aggTool["aggregation_tools.py"]
            compTool["comparison_tools.py"]
        end
        subgraph Planner["app/planner/"]
            anaPlanner["planner.py — analytical planner"]
        end
    end
```

---

## Data Models

```mermaid
classDiagram
    class ChatRequest {
        +str userMessage
        +RangeToken[] rangeTokens
        +WorkbookSnapshot workbookSnapshot
        +ConvMessage[] conversationHistory
        +str conversationId
        +str activeSheet
        +str workbookName
        +str usedRangeEnd
        +str locale
        +ExecutionContext executionContext
    }

    class WorkbookSnapshot {
        +SheetSnapshot[] sheets
        +bool truncated
    }

    class SheetSnapshot {
        +str sheetName
        +int rowCount
        +int columnCount
        +str[] headers
        +str[] dtypes
        +any[][] sampleRows
        +str anchorCell
        +str usedRangeAddress
    }

    class ChatResponse {
        +str responseType
        +str message
        +str messageLocalized
        +ExecutionPlan plan
        +PlanOption[] plans
        +str interactionId
        +str conversationId
        +str assistantMessageId
    }

    class PlanOption {
        +str optionLabel
        +str optionLabelLocalized
        +ExecutionPlan plan
    }

    class ExecutionPlan {
        +str planId
        +str createdAt
        +str userRequest
        +str summary
        +str summaryLocalized
        +PlanStep[] steps
        +bool preserveFormatting
        +float confidence
        +str[] warnings
    }

    class PlanStep {
        +str id
        +str description
        +str descriptionLocalized
        +StepAction action
        +dict params
        +str[] dependsOn
    }

    class ExecutionState {
        +str planId
        +str status
        +StepResult[] stepResults
        +str startedAt
        +str completedAt
    }

    class StepResult {
        +str stepId
        +str status
        +str message
        +dict outputs
        +int durationMs
        +str error
    }

    class PlanSnapshot {
        +str planId
        +str timestamp
        +CellSnapshot[] cells
        +InverseOp[] inverseOps
    }

    ChatRequest --> WorkbookSnapshot
    WorkbookSnapshot --> SheetSnapshot
    ChatResponse --> ExecutionPlan
    ChatResponse --> PlanOption
    PlanOption --> ExecutionPlan
    ExecutionPlan --> PlanStep
    ExecutionState --> StepResult
```

**Canonical-English + `*Localized` pattern.** `message`, `summary`,
`description`, `optionLabel` are always English for stable logs and LLM
planning; the paired `messageLocalized`, `summaryLocalized`,
`descriptionLocalized`, `optionLabelLocalized` carry faithful translations
when the user writes in a non-English language. UIs prefer the localized
field when present.

---

## Persistence Layer (swappable)

```mermaid
flowchart TB
    Env["DATABASE_URL / VECTOR_STORE_URL\nenvironment variables"]
    Env --> Factory["persistence/factory.py"]

    Factory -->|sqlite:// or empty| SqliteRepos["sqlite/repositories.py"]
    Factory -->|postgresql://| PgRepos["postgres/repositories.py"]
    Factory -->|chroma:// or empty| ChromaVec["vector_chroma.py"]
    Factory -->|pgvector:// or postgresql://| PgVec["vector_pgvector.py"]

    SqliteRepos --> SQLiteFile["feedback.db (single file)"]
    PgRepos --> PostgresDB["PostgreSQL (managed)"]
    ChromaVec --> ChromaDir["./data/chroma/"]
    PgVec --> PostgresVec["pgvector extension"]

    SqliteRepos -.shared abstract.-> RepoBase["Repositories ABC"]
    PgRepos -.-> RepoBase
    ChromaVec -.-> VecBase["VectorStore ABC"]
    PgVec -.-> VecBase

    Migrate["scripts/migrate_db.py"]
    Migrate -->|"export from SQLite"| SqliteRepos
    Migrate -->|"import into Postgres"| PgRepos
```

**Development:** `DATABASE_URL=""` (SQLite) + `VECTOR_STORE_URL=""`
(ChromaDB). Everything lives in `backend/data/`.

**Production:** `DATABASE_URL="postgresql://…"` +
`VECTOR_STORE_URL="pgvector://…"` (often the same instance). Migrate with
`backend/scripts/migrate_db.py`.

---

## Semantic Search Pipeline

```mermaid
flowchart LR
    Input["User message"]
    ST["sentence-transformers\nparaphrase-multilingual-MiniLM-L12-v2\n(384-dim, Hebrew-capable)"]

    subgraph VectorStore["Vector Store (Chroma or pgvector)"]
        CapColl["capabilities collection\n(94 action descriptions)"]
        ExColl["few_shot_examples collection\n(v4 seeds, English + Hebrew)"]
    end

    Input --> ST
    ST -->|embedding| CapColl
    ST -->|embedding| ExColl
    CapColl -->|top-K actions<br/>default 10| Prompt["System Prompt\n(filtered to relevant actions)"]
    ExColl -->|top-K examples<br/>default 5| Prompt
```

**Seed versioning.** `example_store.SEED_VERSION` bumps on every
prompt-format migration. Init routine purges prior-version seeds from the
vector store before re-seeding the current version, so retrieval never
returns stale patterns. Currently on **v4**.

---

## Execution Engine (frontend, Office.js)

```mermaid
flowchart TB
    Plan["ExecutionPlan"] --> Validate["validator.ts"]
    Validate -->|fail| Error["Surface validation errors"]
    Validate -->|pass| Sort["topologicalSort(steps)"]

    Sort --> Loop{"For each step"}

    Loop --> MetaCheck{"meta.mutates?"}
    MetaCheck -->|yes| SnapBranch{"range params?"}
    SnapBranch -->|yes| SnapRanges["captureSnapshotBatched()"]
    SnapBranch -->|no| SnapEmpty["createEmptySnapshot()\n(for structural InverseOp)"]
    MetaCheck -->|no| Dispatch
    SnapRanges --> Dispatch
    SnapEmpty --> Dispatch

    Dispatch["registry.getHandler(action)"]
    Dispatch --> ApiGuard{"API set<br/>supported?"}
    ApiGuard -->|no| Fallback["registry.getFallback(action)"]
    ApiGuard -->|yes| Primary["primary handler"]
    Fallback --> Primary

    Primary --> PreChecks
    subgraph PreChecks["Handler entry guards"]
        MergeCheck["ensureUnmerged()<br/>(5 read-then-write handlers)"]
        NormalizeCheck["normalizeString / parseNumberFlexible /<br/>parseDateFlexible for dirty input"]
        ClipCheck["clipFullColumnRefs()<br/>(writeFormula, spillFormula)"]
        DynArrCheck["dynamicArrayRewrite()<br/>(pre-365 Excel only)"]
    end

    PreChecks --> OfficeAPI["Office.js: context.sync()"]
    OfficeAPI --> Excel["Excel Workbook"]

    OfficeAPI --> InverseReg{"structural op?"}
    InverseReg -->|yes| RegInv["registerInverseOp(...)\n(deleteSheet, restoreTabColor,<br/>deleteRows, ...)"]
    InverseReg -->|no| Result

    OfficeAPI --> FootprintReg{"wide/tall write?"}
    FootprintReg -->|yes| ExpSnap["expandSnapshotFootprint()\n(lateralSpread, extractMatched)"]
    FootprintReg -->|no| Result

    Result["StepResult"]
    Result -->|success| Loop
    Result -->|error| Rollback["rollbackPlan()"]
    Rollback --> RestoreCells["Restore cell values"]
    Rollback --> ApplyInv["Apply InverseOps in reverse"]

    Loop -->|all steps done| Complete["ExecutionState: completed"]
    Complete --> PromptFB["Prompt user for feedback"]
```

---

## Excel 2016 Compatibility Layer

```mermaid
flowchart TB
    Formula["Formula from LLM plan"]
    Formula --> Clip["clipFullColumnRefs()<br/>A:A → A1:A{usedRow}"]
    Clip --> DynCheck{"Uses LET / FILTER /<br/>UNIQUE / XLOOKUP /<br/>SORT / SEQUENCE /<br/>INDEX+MATCH(1,arr,0) ?"}

    DynCheck -->|no| Write["range.formulas = [[f]]"]
    DynCheck -->|yes| ApiCheck{"ExcelApi 1.11+?"}

    ApiCheck -->|yes| Write
    ApiCheck -->|no| Rewrite["dynamicArrayRewrite()"]

    subgraph RewriteSteps["Rewrite pipeline (order matters)"]
        LetInline["LET(name, expr, ..., body)<br/>→ body with every name<br/>textually inlined as (expr)"]
        XlookupRe["XLOOKUP → INDEX/MATCH"]
        IndexMatch["INDEX(r, MATCH(1, arr, 0))<br/>→ LOOKUP(2, 1/arr, r)"]
        SeqRe["SEQUENCE → ROW(INDIRECT)"]
        UniqueRe["UNIQUE → INDEX/MATCH/COUNTIF"]
        SortRe["SORT → INDEX/MATCH/SMALL"]
        FilterRe["FILTER → INDEX/SMALL/IF array"]
    end

    Rewrite --> LetInline
    LetInline --> XlookupRe
    XlookupRe --> IndexMatch
    IndexMatch --> SeqRe
    SeqRe --> UniqueRe
    UniqueRe --> SortRe
    SortRe --> FilterRe
    FilterRe --> Write

    Write --> Verify{"Cell evaluates<br/>to error?"}
    Verify -->|#NAME? / #SPILL! /<br/>#REF! / #VALUE!| Report["StepResult: error<br/>(logs rejected formula)"]
    Verify -->|clean| Success["StepResult: success"]
```

**17 handlers register legacy fallbacks** for the Office.js API methods
they use (createPivot → SUMIFS summary, addSparkline → mini charts,
addSlicer → applyFilter guidance, etc.). Dispatch happens at
`capabilityRegistry.getFallback(action)` when `meta.requiresApiSet` is
greater than the current Excel's support level.

---

## Feedback Loop & Continuous Learning

```mermaid
flowchart TB
    User -->|"clicks Applied ✓"| FeedbackAPI["POST /api/feedback"]
    FeedbackAPI --> Repo["Repositories (SQLite / Postgres)"]
    FeedbackAPI --> Promote["example_store.add_user_example()"]
    Promote --> VecStore["Vector Store (Chroma / pgvector)"]

    subgraph "Next Request"
        NewMsg["New user message"] --> Search["search_examples()"]
        Search --> VecStore
        VecStore -->|"returns promoted<br/>example if similar"| FewShot["Few-shot injection"]
        FewShot -->|"better plan quality"| LLM["LLM"]
    end
```

---

## Deployment Topology

### Add-in mode (OpenShift)

```mermaid
flowchart TB
    subgraph UserMachine["User's Machine"]
        ExcelApp["Excel Desktop / Web / Mac / iPad"]
    end

    subgraph OpenShift["OpenShift Cluster"]
        Route["Route (TLS edge)"]
        Service["Service :8080"]
        Pod["Pod (N replicas)"]

        subgraph Container["Container"]
            Uvicorn["Uvicorn"]
            FastAPI["FastAPI"]
            Static["./static/ (built React app)"]
            Model["multilingual-MiniLM-L12-v2"]
        end

        PVC["PVC (RWX for multi-replica)"]
        Secret["Secret — LLM_API_KEY"]
        ConfigMap["ConfigMap — DATABASE_URL,\nVECTOR_STORE_URL, CORS_ORIGINS, ..."]
    end

    subgraph DB["Optional: External Postgres"]
        Pg["PostgreSQL + pgvector"]
    end

    subgraph LLMApi["LLM Provider"]
        LLMProvider["OpenAI / Azure / Anthropic /\nOllama / Gemini / Cohere"]
    end

    ExcelApp -->|"HTTPS"| Route
    Route --> Service --> Pod --> Container
    FastAPI --> Static
    FastAPI --> Model
    Container --> PVC
    Container --> Secret
    Container --> ConfigMap
    FastAPI --> LLMProvider
    FastAPI -.optional.-> Pg
```

### MCP mode (desktop)

```mermaid
flowchart LR
    Client["MCP Client<br/>(Claude Desktop, Cursor,<br/>Windsurf, Zed)"]
    MCP["excel-copilot-mcp (stdio)"]
    Excel["Desktop Excel<br/>(Windows COM / Mac AppleScript)"]
    LLM["LLM Provider"]

    Client <-->|JSON-RPC over stdio| MCP
    MCP <-->|xlwings| Excel
    MCP --> LLM
```

Installed as a console script via `backend/pyproject.toml`. MCP client
config is a one-line entry: `{"excel-copilot": {"command":
"excel-copilot-mcp"}}`.

---

## Capability-layer invariants (enforced by tests)

| Invariant | Test |
|---|---|
| Every TS `StepAction` has a Pydantic param model | `test_chat_service.py::test_all_actions_have_param_models` |
| Every Pydantic param has a Python xlwings handler | runtime registry scan at startup |
| Every TS handler file is imported in `capabilities/index.ts` | `capability-compliance.test.ts` |
| Handlers that load `.values` use `getUsedRange(false)` (or are exempt) | `capability-compliance.test.ts` |
| Handlers with `requiresApiSet > 1.3` register a fallback | `capabilityRegistry.fallback.test.ts` |
| Every handler uses `options.dryRun` guard | `capability-compliance.test.ts` |
| Handlers with sync-after-writes have try-catch | `capability-compliance.test.ts` |

**Action count verification at test time:** 94 TS handlers = 94 Pydantic
param models = 94 Python xlwings handlers. If someone adds an action to
one side without mirroring, the test suite fails.

---

## Security / Air-gap posture

- Telemetry disabled: `ANONYMIZED_TELEMETRY=False`, `HF_HUB_OFFLINE=1`,
  `TRANSFORMERS_OFFLINE=1` (baked into `Dockerfile` and
  `openshift/configmap.yaml`).
- Embedding model bundled on disk — no HuggingFace calls at runtime.
- LLM endpoint fully configurable — works with on-prem Ollama / internal
  OpenAI-compatible gateways.
- Model weights excluded from git (`*.safetensors` etc. in `.gitignore`);
  `backend/scripts/download_embedding_model.py` fetches on first
  container build.
- See `docs/AIRGAP.md` for the enclosed-network checklist and
  `docs/SECURITY_CHECKLIST.md` for production hardening steps.
