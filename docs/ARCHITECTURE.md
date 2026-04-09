# Architecture — AI Assistant For Excel

End-to-end flow from the moment a user sends a message to the final
Excel mutation. Every box in the diagrams maps to a real file in the repo.

---

## High-Level Overview

```mermaid
flowchart TB
    subgraph Excel["Excel Client (Office.js)"]
        UI["ChatPanel / ChatInput"]
        Hook["useChat hook"]
        API["api.ts — sendChatMessageStream()"]
        Exec["Executor — executePlan()"]
        Snap["Snapshot — captureSnapshot()"]
        Handlers["Capability Handlers (76)"]
        Registry["capabilityRegistry.ts"]
        Validator_FE["Validator (frontend)"]
    end

    subgraph Backend["FastAPI Backend"]
        Router["POST /api/chat/stream"]
        ChatSvc["ChatService.chat_stream()"]
        PromptBuilder["_build_chat_messages()"]
        LLM["llm_client — acompletion_stream()"]
        Parser["_parse_response()"]
        Persist["_persist_conversation_turn()"]
        FBRouter["POST /api/feedback"]
        ConvRouter["GET/PATCH /api/conversations"]
    end

    subgraph Intelligence["Semantic Intelligence"]
        CapStore["CapabilityStore — search_capabilities()"]
        ExStore["ExampleStore — search_examples()"]
        Chroma["ChromaDB (PersistentClient)"]
        Embed["SentenceTransformer\nall-MiniLM-L6-v2"]
    end

    subgraph Storage["Persistent Storage"]
        SQLite["SQLite — conversations, feedback"]
        ChromaDir["data/chroma/ — vector collections"]
        PVC["PVC (OpenShift) or local ./data"]
    end

    subgraph LLMProvider["LLM Provider"]
        OpenAI["OpenAI / Azure"]
        Anthropic["Anthropic"]
        Ollama["Ollama (local or cloud)"]
        Other["Cohere / Gemini / etc."]
    end

    UI -->|"user types message"| Hook
    Hook -->|"build ChatRequest\n(message, rangeTokens,\nworkbookSnapshot)"| API
    API -->|"POST /api/chat/stream\n(SSE connection)"| Router
    Router --> ChatSvc

    ChatSvc --> CapStore
    CapStore --> Chroma
    Chroma --> Embed

    ChatSvc --> PromptBuilder
    PromptBuilder -->|"system prompt +\nfew-shot examples +\nconversation history +\nuser content"| LLM
    PromptBuilder --> ExStore
    ExStore --> Chroma

    LLM -->|"streaming tokens"| ChatSvc
    LLM --> OpenAI & Anthropic & Ollama & Other

    ChatSvc --> Parser
    Parser -->|"ChatResponse\n(message or plans)"| ChatSvc
    ChatSvc --> Persist
    Persist --> SQLite

    ChatSvc -->|"SSE: chunk → done"| API
    API -->|"ChatResponse"| Hook

    Hook -->|"show plan preview"| UI
    UI -->|"user clicks Run"| Exec
    Exec --> Validator_FE
    Exec --> Snap
    Exec -->|"dispatch steps"| Registry
    Registry --> Handlers
    Handlers -->|"Office.js API calls\n(Range.values, Chart, Pivot…)"| Excel

    UI -->|"user clicks 👍 Applied"| FBRouter
    FBRouter -->|"promote to few-shot"| ExStore
    FBRouter --> SQLite

    SQLite --> PVC
    ChromaDir --> PVC
```

---

## Detailed Request Flow

```mermaid
sequenceDiagram
    participant U as User
    participant CP as ChatPanel
    participant UC as useChat Hook
    participant API as api.ts
    participant R as /api/chat/stream
    participant CS as ChatService
    participant Cap as CapabilityStore
    participant Ex as ExampleStore
    participant DB as ChromaDB
    participant LLM as LLM Provider
    participant PR as _parse_response
    participant SQ as SQLite
    participant EX as Executor
    participant OH as Office.js Handlers
    participant XL as Excel Workbook
    participant FB as /api/feedback

    U->>CP: Types message + optional range selection
    CP->>UC: handleSend(text, rangeTokens)

    Note over UC: Builds ChatRequest:<br/>userMessage, rangeTokens,<br/>workbookSnapshot (headers,<br/>dtypes, 5 sample rows),<br/>conversationHistory (last 10),<br/>activeSheet, locale

    UC->>API: sendChatMessageStream(request)
    API->>R: POST /api/chat/stream (SSE)

    R->>CS: chat_stream(ChatRequest)

    rect rgb(230, 245, 255)
        Note over CS,DB: Phase 1 — Semantic Search
        CS->>Cap: search_capabilities(userMessage)
        Cap->>DB: query "capabilities" collection
        DB-->>Cap: top-K action names (default 10)
        Cap-->>CS: relevant_actions[]
    end

    rect rgb(255, 245, 230)
        Note over CS,Ex: Phase 2 — Build Prompt
        CS->>CS: _build_chat_system_prompt(relevant_actions)
        Note over CS: Filter CAPABILITY_DESCRIPTIONS<br/>to only relevant actions.<br/>250+ line system prompt with<br/>response schema, rules, dashboard<br/>building patterns, snapshot warnings.

        CS->>Ex: search_examples(userMessage, top_k=5)
        Ex->>DB: query "few_shot_examples" collection
        DB-->>Ex: similar examples
        Ex-->>CS: few-shot message pairs

        CS->>CS: _build_user_content(request)
        Note over CS: Injects: date/time, locale,<br/>used range end, range tokens,<br/>workbook snapshot with<br/>⚠ WARNING about sample rows
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
        Note over CS,PR: Phase 4 — Parse & Validate
        CS->>PR: _parse_response(full_text)
        Note over PR: 1. extract_json() — strip fences,<br/>   trailing commas, control chars<br/>2. Handle off-schema patterns:<br/>   - alias fields (response/content/text)<br/>   - bare steps → wrap in plan<br/>   - tool_calls format<br/>   - snake_case → camelCase<br/>3. _fill_plan_defaults() — UUID,<br/>   timestamp, confidence<br/>4. Validate via Pydantic models

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

    CS->>SQ: _persist_conversation_turn()
    Note over SQ: Save user msg + assistant response<br/>+ plan to conversations table

    CS-->>R: SSE event: {type: "done", result: ChatResponse}
    R-->>API: final SSE
    API-->>UC: ChatResponse

    alt responseType = "message"
        UC->>CP: Display text response
    else responseType = "plans"
        UC->>CP: Show plan option cards (2-3 options)
        U->>CP: Selects preferred plan
    else responseType = "plan"
        UC->>CP: Show single plan preview
    end

    rect rgb(245, 245, 220)
        Note over U,XL: Phase 5 — Plan Execution
        U->>CP: Clicks "Run" on chosen plan
        CP->>EX: executePlan(plan, callbacks)

        EX->>EX: validatePlan(plan)
        EX->>EX: topologicalSort(steps)

        loop For each step (dependency order)
            EX->>EX: captureSnapshot(affected ranges)
            EX->>OH: handler(context, params)
            OH->>XL: Office.js API calls
            Note over OH,XL: Range.values = [[...]]<br/>Chart.setData(...)<br/>PivotTable.add(...)<br/>ConditionalFormat.add(...)
            XL-->>OH: result
            OH-->>EX: StepResult {status, message, durationMs}
            EX-->>CP: onStepComplete callback
        end

        EX-->>CP: ExecutionState {status: "completed"}
    end

    CP->>FB: sendFeedback(interactionId, planId, "applied")
    FB->>SQ: Log feedback
    FB->>Ex: add_user_example() — promote to few-shot pool
    Ex->>DB: Upsert into "few_shot_examples" collection
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
            PlanPreview["PlanPreview.tsx"]
            PlanOptions["PlanOptionsPanel.tsx"]
            StepCard["StepCard.tsx"]
        end
        subgraph Hooks["taskpane/hooks/"]
            useChat["useChat.ts"]
            usePlanExec["usePlanExecution.ts"]
        end
        subgraph Engine["engine/"]
            Executor["executor.ts"]
            Snapshot["snapshot.ts"]
            Validator["validator.ts"]
            CapReg["capabilityRegistry.ts"]
            subgraph Caps["capabilities/ (76 actions)"]
                writeValues["writeValues.ts"]
                writeFormula["writeFormula.ts"]
                matchRecords["matchRecords.ts"]
                createChart["createChart.ts"]
                createPivot["createPivot.ts"]
                moreActions["… 71 more"]
            end
        end
        subgraph Services["services/"]
            apiTS["api.ts"]
        end
    end

    subgraph BackendFiles["backend/app/"]
        direction TB
        subgraph Routers["routers/"]
            chatRouter["chat.py"]
            fbRouter["feedback.py"]
            convRouter["conversations.py"]
            healthRouter["health.py"]
        end
        subgraph ServicesBE["services/"]
            chatService["chat_service.py"]
            llmClient["llm_client.py"]
            chromaClient["chroma_client.py"]
            capStore["capability_store.py"]
            exStore["example_store.py"]
            validatorBE["validator.py"]
        end
        subgraph Models["models/"]
            chatModel["chat.py — ChatRequest/Response"]
            planModel["plan.py — ExecutionPlan, PlanStep"]
            actionsModel["actions.py — StepAction enum"]
            paramsModel["params/ — 76 param models"]
        end
        subgraph Orchestrator["orchestrator/"]
            orch["orchestrator.py"]
        end
        mainPy["main.py — FastAPI app"]
        dbPy["db.py — SQLite"]
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
        +ExecutionPlan plan
        +PlanOption[] plans
        +str interactionId
        +str conversationId
    }

    class ExecutionPlan {
        +str planId
        +str createdAt
        +str userRequest
        +str summary
        +PlanStep[] steps
        +bool preserveFormatting
        +float confidence
        +str[] warnings
    }

    class PlanStep {
        +str id
        +str description
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
        +int durationMs
        +str error
    }

    ChatRequest --> WorkbookSnapshot
    WorkbookSnapshot --> SheetSnapshot
    ChatResponse --> ExecutionPlan
    ExecutionPlan --> PlanStep
    ExecutionState --> StepResult
```

---

## Semantic Search Pipeline

```mermaid
flowchart LR
    subgraph Input
        MSG["User message"]
    end

    subgraph Embedding
        ST["SentenceTransformer\nall-MiniLM-L6-v2"]
    end

    subgraph ChromaDB
        CapColl["capabilities collection\n(76 action descriptions)"]
        ExColl["few_shot_examples collection\n(27 seed + user-promoted)"]
    end

    subgraph Output
        Actions["Top-K relevant actions\n(default 10)"]
        Examples["Top-K similar examples\n(default 5)"]
    end

    MSG --> ST
    ST -->|"384-dim vector"| CapColl
    ST -->|"384-dim vector"| ExColl
    CapColl -->|"cosine similarity"| Actions
    ExColl -->|"cosine similarity"| Examples

    Actions -->|"filter system prompt\nto relevant actions only"| SysPrompt["System Prompt"]
    Examples -->|"inject as few-shot\nmessage pairs"| SysPrompt
```

---

## Feedback Loop & Continuous Learning

```mermaid
flowchart TB
    User -->|"clicks Applied ✓"| FeedbackAPI["POST /api/feedback"]
    FeedbackAPI --> SQLite["SQLite — log interaction"]
    FeedbackAPI --> Promote["add_user_example()"]
    Promote --> ChromaDB["ChromaDB — few_shot_examples"]

    subgraph "Next Request"
        NewMsg["New user message"] --> Search["search_examples()"]
        Search --> ChromaDB
        ChromaDB -->|"returns promoted\nexample if similar"| FewShot["Few-shot prompt injection"]
        FewShot -->|"better plan quality"| LLM["LLM"]
    end
```

---

## Execution Engine (Frontend)

```mermaid
flowchart TB
    Plan["ExecutionPlan"] --> Validate["validatePlan()"]
    Validate -->|"pass"| Sort["topologicalSort(steps)"]
    Validate -->|"fail"| Error["Show validation errors"]

    Sort --> Loop{"For each step\n(dependency order)"}

    Loop --> Capture["captureSnapshot()\nSave affected cell values"]
    Capture --> Dispatch["capabilityRegistry.get(action)"]
    Dispatch --> Handler["Handler function"]
    Handler -->|"Excel.run(async ctx => ...)"| OfficeJS["Office.js API"]
    OfficeJS -->|"ctx.sync()"| Excel["Excel Workbook"]
    Excel --> Result["StepResult"]
    Result -->|"success"| Loop
    Result -->|"error"| Rollback["rollbackPlan()\nRestore snapshots"]

    Loop -->|"all steps done"| Complete["ExecutionState: completed"]
    Complete --> Feedback["Prompt user for feedback"]
```

---

## Deployment Topology (OpenShift)

```mermaid
flowchart TB
    subgraph UserMachine["User's Machine"]
        ExcelApp["Excel Desktop / Web"]
    end

    subgraph OpenShift["OpenShift Cluster"]
        Route["Route (TLS edge)\nhttps://excel-assistant.apps..."]
        Service["Service :8080"]
        Pod["Pod (1 replica)"]

        subgraph Container["Container"]
            Uvicorn["Uvicorn (1 worker)"]
            FastAPI["FastAPI App"]
            Static["./static/ (built React app)"]
            Model["all-MiniLM-L6-v2\n(bundled ~87MB)"]
        end

        PVC["PVC 2Gi (RWO)"]
        Secret["Secret\nexcel-assistant-secrets"]
        ConfigMap["ConfigMap\nexcel-assistant-config"]
    end

    subgraph External["External (or internal)"]
        LLMApi["LLM API\n(OpenAI / Anthropic / Ollama)"]
    end

    ExcelApp -->|"HTTPS"| Route
    Route -->|"HTTP :8080"| Service
    Service --> Pod
    Pod --> Container
    Uvicorn --> FastAPI
    FastAPI -->|"serves"| Static
    FastAPI -->|"uses"| Model
    Container -->|"mount"| PVC
    Container -->|"env from"| Secret
    Container -->|"env from"| ConfigMap
    FastAPI -->|"API calls"| LLMApi
