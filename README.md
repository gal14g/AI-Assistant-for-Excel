# Excel AI Copilot

A production-oriented Excel Office Add-in that provides a natural-language chat interface for spreadsheet operations. Type commands in plain English, click cells to reference ranges, preview execution plans, and run them with full undo support.

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                     Excel Workbook                          │
│  ┌────────────┐  ┌──────────────────────────────────────┐   │
│  │   Ribbon    │  │          Task Pane (React)           │   │
│  │  Commands   │  │  ┌──────────────────────────────┐    │   │
│  │             │  │  │       Chat Panel              │    │   │
│  │ [Open]      │  │  │  - Message history            │    │   │
│  │ [Undo]      │  │  │  - Range tokens               │    │   │
│  │             │  │  │  - Streaming explanations      │    │   │
│  └──────┬─────┘  │  │  - Plan preview                │    │   │
│         │        │  │  - Execution timeline           │    │   │
│    Shared Runtime │  └──────────┬───────────────────┘    │   │
│         │        │             │                          │   │
│  ┌──────┴────────┴─────────────┴──────────────────────┐   │   │
│  │              Execution Engine                       │   │   │
│  │  ┌────────────┐ ┌──────────┐ ┌──────────────────┐  │   │   │
│  │  │ Capability │ │Validator │ │  Snapshot/Rollback│  │   │   │
│  │  │ Registry   │ │          │ │                   │  │   │   │
│  │  └────────────┘ └──────────┘ └──────────────────┘  │   │   │
│  │  ┌──────────────────────────────────────────────┐  │   │   │
│  │  │  Office.js Capability Handlers (15+)         │  │   │   │
│  │  │  readRange | writeValues | writeFormula | ... │  │   │   │
│  │  └──────────────────────────────────────────────┘  │   │   │
│  └────────────────────────────────────────────────────┘   │   │
└───────────────────────────┬─────────────────────────────────┘
                            │ REST + SSE
                ┌───────────┴───────────┐
                │   FastAPI Backend      │
                │  ┌─────────────────┐   │
                │  │  LLM Planner    │   │
                │  │  (Claude API)   │   │
                │  └────────┬────────┘   │
                │  ┌────────┴────────┐   │
                │  │  Plan Validator │   │
                │  │  (Pydantic)     │   │
                │  └─────────────────┘   │
                └────────────────────────┘
```

### Key Design Principles

1. **LLM produces JSON plans, never executable code.** The planner outputs a strict typed JSON plan. The frontend validates and executes it via safe Office.js wrappers.

2. **Formatting preservation by default.** Write operations use `range.values` (not clipboard/copy), which only sets cell values without touching formatting.

3. **Native formulas preferred.** When possible, the planner uses `writeFormula` with `XLOOKUP`, `SUMIF`, etc. so results auto-update and are auditable.

4. **Snapshot before every write.** The executor captures cell values before each mutating step, enabling per-plan rollback.

5. **Shared runtime.** Task pane, ribbon commands, and future custom functions share one JS context for coordinated state.

## Project Structure

```
excel-ai-copilot/
├── frontend/                    # Office Add-in (React + TypeScript)
│   ├── manifest.xml             # Office Add-in manifest (shared runtime)
│   ├── package.json
│   ├── webpack.config.js
│   ├── public/
│   │   ├── taskpane.html        # Task pane entry point
│   │   └── commands.html        # Ribbon commands entry point
│   └── src/
│       ├── index.tsx             # React bootstrap after Office.onReady
│       ├── commands/commands.ts  # Ribbon button handlers
│       ├── engine/
│       │   ├── types.ts          # All TypeScript type definitions
│       │   ├── capabilityRegistry.ts  # Action → handler registry
│       │   ├── executor.ts       # Plan execution engine
│       │   ├── validator.ts      # Client-side plan validation
│       │   ├── snapshot.ts       # Snapshot/rollback management
│       │   └── capabilities/     # Individual Office.js action handlers
│       │       ├── readRange.ts, writeValues.ts, writeFormula.ts
│       │       ├── matchRecords.ts, groupSum.ts
│       │       ├── createTable.ts, applyFilter.ts, sortRange.ts
│       │       ├── createPivot.ts, createChart.ts
│       │       ├── conditionalFormat.ts, cleanupText.ts
│       │       ├── removeDuplicates.ts, freezePanes.ts
│       │       ├── findReplace.ts, sheetOps.ts, validation.ts
│       │       └── index.ts      # Auto-registers all capabilities
│       ├── services/
│       │   ├── api.ts            # REST client for backend
│       │   └── streaming.ts      # SSE client for streaming updates
│       ├── shared/
│       │   └── planSchema.ts     # Zod runtime schema validation
│       └── taskpane/
│           ├── App.tsx
│           ├── App.css
│           ├── components/
│           │   ├── ChatPanel.tsx
│           │   ├── ChatInput.tsx
│           │   ├── MessageBubble.tsx
│           │   ├── RangeToken.tsx
│           │   ├── PlanPreview.tsx
│           │   └── ExecutionTimeline.tsx
│           └── hooks/
│               ├── useChat.ts
│               ├── useSelectionTracker.ts
│               └── usePlanExecution.ts
├── backend/                     # FastAPI backend
│   ├── main.py                  # FastAPI app entry point
│   ├── requirements.txt
│   ├── .env.example
│   ├── app/
│   │   ├── config.py            # Settings from environment
│   │   ├── models/
│   │   │   ├── plan.py          # Pydantic models for plan schema
│   │   │   └── request.py       # API request/response models
│   │   ├── services/
│   │   │   ├── planner.py       # LLM planner (Claude API)
│   │   │   └── validator.py     # Server-side plan validation
│   │   ├── routers/
│   │   │   ├── plan.py          # REST endpoints
│   │   │   └── stream.py        # SSE streaming endpoints
│   │   └── prompts/
│   │       ├── planner_system.txt   # System prompt for planner
│   │       └── validator_system.txt # System prompt for validator
│   ├── tests/
│   │   ├── test_plan_validation.py
│   │   └── test_planner.py
│   └── examples/
│       ├── match_and_sum_plan.json
│       ├── chart_creation_plan.json
│       └── cleanup_plan.json
└── shared/
    └── plan-schema.json         # JSON Schema (shared contract)
```

## Local Development Setup

### Prerequisites

- Node.js 18+
- Python 3.11+
- Excel (desktop or Microsoft 365 web)
- An Anthropic API key

### 1. Backend Setup

```bash
cd backend

# Create virtual environment
python -m venv venv
source venv/bin/activate  # or `venv\Scripts\activate` on Windows

# Install dependencies
pip install -r requirements.txt

# Configure environment
cp .env.example .env
# Edit .env and set your ANTHROPIC_API_KEY

# Run the server
python -m uvicorn main:app --reload --port 8000
```

The API will be available at `http://localhost:8000`. Check `http://localhost:8000/docs` for the interactive API documentation.

### 2. Frontend Setup

```bash
cd frontend

# Install dependencies
npm install

# Start the dev server (HTTPS on port 3000)
npm run dev
```

### 3. Sideload the Add-in

#### Excel Desktop (Windows/Mac)

```bash
cd frontend
npx office-addin-debugging start manifest.xml
```

Or manually:
1. Open Excel
2. Go to **Insert → My Add-ins → Upload My Add-in**
3. Browse to `frontend/manifest.xml`
4. The "AI Copilot" tab will appear in the ribbon

#### Excel on the Web

1. Go to [office.com](https://www.office.com) and open Excel
2. Go to **Insert → Office Add-ins → Upload My Add-in**
3. Upload `frontend/manifest.xml`

### 4. Running Tests

```bash
# Backend tests
cd backend
pytest tests/ -v

# Frontend tests (when configured)
cd frontend
npm test
```

## Usage

1. Click **"Open Copilot"** in the ribbon to open the task pane
2. Type a natural-language command in the chat input
3. Click cells/ranges in Excel to insert `[[Sheet1!A1:C20]]` tokens
4. The system sends your request to the backend LLM planner
5. A JSON execution plan is returned and displayed for review
6. Click **"Preview"** to dry-run the plan (no changes)
7. Click **"Run Plan"** to execute live
8. Click **"Undo Last"** to rollback if needed

### Example Commands

- "Sum column B grouped by the categories in column A"
- "Match everything in [[Sheet2!B:B]] with records in [[Sheet1!A:A]] and pull the prices"
- "Create a bar chart of sales by region from A1:C20"
- "Clean up names in column A: trim whitespace and proper case"
- "Sort the table by date descending"
- "Add dropdown validation to D2:D100 with options: Low, Medium, High"
- "Freeze the top row and first column"
- "Replace all 'N/A' with empty string in the active sheet"
- "Create a pivot table from A1:E500 with Product as rows and sum of Revenue"

## Supported Operations (15+ capabilities)

| Category | Actions |
|----------|---------|
| Read/Write | `readRange`, `writeValues`, `writeFormula` |
| Lookups | `matchRecords` (XLOOKUP/VLOOKUP) |
| Aggregation | `groupSum` (SUMIF/SUMIFS) |
| Tables | `createTable`, `applyFilter`, `sortRange` |
| Pivots | `createPivot` |
| Charts | `createChart` |
| Formatting | `addConditionalFormat` |
| Cleanup | `cleanupText`, `removeDuplicates`, `findReplace` |
| View | `freezePanes` |
| Validation | `addValidation` |
| Sheets | `addSheet`, `renameSheet`, `deleteSheet`, `copySheet`, `protectSheet` |

## Extension Points

### Adding Custom Functions (v2)

The shared runtime is already configured for custom functions. To add them:

1. Create `frontend/src/functions/functions.ts`
2. Define functions using `CustomFunctions.associate()`
3. Add a `functions.json` metadata file
4. Update the manifest's `<CustomFunctions>` extension point

Because of the shared runtime, custom functions can access the same state as the task pane (e.g., cached data, user preferences).

### Adding New Capabilities

1. Create a new file in `frontend/src/engine/capabilities/`
2. Define the `CapabilityMeta` and handler function
3. Call `registry.register(meta, handler)` at module scope
4. Import the new file in `capabilities/index.ts`
5. Add the action to `StepAction` in `types.ts`
6. Add the Pydantic param model in `backend/app/models/plan.py`
7. Update the planner system prompt

## Office.js Limitations & Notes

- **Range size limits**: Very large ranges (>100k cells) may cause performance issues. Consider chunking.
- **XLOOKUP availability**: Requires Excel 365 or Excel 2021+. The planner should fall back to VLOOKUP for older versions.
- **Dynamic arrays**: Spill behavior (XLOOKUP, FILTER, SORT) requires Excel 365. Older versions need explicit ranges.
- **PivotTable API**: Requires ExcelApi 1.8+. Not available in all Excel versions.
- **Conditional format API**: Requires ExcelApi 1.6+.
- **Freeze panes**: Requires ExcelApi 1.7+.
- **Remove duplicates**: Requires ExcelApi 1.9+.
- **context.sync()**: Must be called to flush operations. Batching multiple operations before sync improves performance.
- **Proxy objects**: Only valid within their `Excel.run()` callback. Do not store them across calls.

## Packaging for Production

### Frontend

```bash
cd frontend
npm run build
# Output in dist/ – host on your CDN/web server
```

Update `manifest.xml` URLs to point to your production server instead of `localhost:3000`.

### Backend

```bash
cd backend
# Use gunicorn or uvicorn for production
uvicorn main:app --host 0.0.0.0 --port 8000 --workers 4
```

### Publishing the Add-in

1. Update the manifest with production URLs
2. Submit to [Microsoft AppSource](https://appsource.microsoft.com/) for public distribution
3. Or deploy via your organization's admin center for internal distribution
