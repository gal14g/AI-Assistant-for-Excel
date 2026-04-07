# Adding new actions / tools to AI Assistant For Excel

Step-by-step guide to adding a new capability. Every action touches
**5 backend files** and **4 frontend files** — the system is fully symmetric.

We'll use a fictional `highlightOutliers` action as the running example.

---

## Overview

```
Backend (validation + LLM grounding)        Frontend (validation + execution)
─────────────────────────────────────        ─────────────────────────────────
1. plan.py        → enum + params model      6. types.ts      → type union + interface
2. planner.py     → capability description   7. validator.ts  → required-field checks
3. capability_store.py → search examples     8. capabilities/highlightOutliers.ts → handler
4. validator.py   → (auto — uses model)      9. capabilities/index.ts → import
5. example_store.py → (optional) few-shot
```

---

## Backend

### Step 1 — Add to the `StepAction` enum

**File:** `backend/app/models/plan.py`

```python
class StepAction(str, Enum):
    # ... existing 51 actions ...
    highlightOutliers = "highlightOutliers"   # ← add at the end
```

### Step 2 — Create the parameter model

**Same file:** `backend/app/models/plan.py`

Add a Pydantic model after the existing param classes (~line 70–490).
Use `Optional` for any field the LLM can omit.

```python
class HighlightOutliersParams(BaseModel):
    range: str                                      # required — data range
    method: str = "iqr"                             # "iqr" | "zscore" | "percentile"
    threshold: Optional[float] = 1.5                # IQR multiplier or z-score cutoff
    highlightColor: Optional[str] = "#FF6B6B"       # fill color for outlier cells
```

Rules:
- Every field the LLM must provide → no default (Pydantic will reject if missing)
- Every field with a sensible default → `Optional[T] = default`
- Use `str` for range addresses, `int` for column indices, `list[list[...]]` for 2D data
- Field names must be **camelCase** (matches frontend TypeScript conventions)

### Step 3 — Register in `ACTION_PARAM_MODELS`

**Same file:** `backend/app/models/plan.py` (bottom of file)

```python
ACTION_PARAM_MODELS: dict[StepAction, type[BaseModel]] = {
    # ... existing 51 entries ...
    StepAction.highlightOutliers: HighlightOutliersParams,
}
```

This is what makes backend validation work. If you skip this, the action's
params won't be validated and any garbage the LLM produces will pass through.

### Step 4 — Add the capability description

**File:** `backend/app/services/planner.py`

```python
CAPABILITY_DESCRIPTIONS: dict[str, str] = {
    # ... existing entries ...
    "highlightOutliers": (
        "Detect and highlight statistical outliers in a numeric column. "
        "Params: range (string, required), method ('iqr'|'zscore'|'percentile', default 'iqr'), "
        "threshold (number, default 1.5), highlightColor (hex string, default '#FF6B6B')"
    ),
}
```

This is the **single source of truth** the LLM reads when deciding which action
to use and what params to pass. Be precise:
- Name every param and its type
- State which are required vs optional (with defaults)
- Use the exact param names that match your Pydantic model

### Step 5 — Add search examples

**File:** `backend/app/services/capability_store.py`

```python
CAPABILITY_EXAMPLES: dict[str, list[str]] = {
    # ... existing entries ...
    "highlightOutliers": [
        "find outliers in column B",
        "highlight unusual values",
        "mark statistical outliers in the sales data",
        "show me which numbers are abnormal",
    ],
}
```

These are embedded into ChromaDB and used for **semantic search** — when the
user says "find weird numbers", the vector search matches it to these examples
and retrieves `highlightOutliers` as a relevant capability. Write 3-5 examples
using natural language a user would actually say.

### Step 5b (optional) — Add a few-shot seed example

**File:** `backend/app/services/example_store.py`

Add to the `SEED_EXAMPLES` list. This gives the LLM a concrete input→output
example to learn from:

```python
{
    "user_message": "highlight the outliers in column C of Sheet1",
    "assistant_response": json.dumps({
        "responseType": "plans",
        "message": "I'll highlight statistical outliers in column C.",
        "plans": [{
            "optionLabel": "Option A: IQR method",
            "plan": {
                "summary": "Highlight outliers using IQR method",
                "steps": [{
                    "id": "step_1",
                    "description": "Highlight outliers in Sheet1!C:C",
                    "action": "highlightOutliers",
                    "params": {
                        "range": "Sheet1!C:C",
                        "method": "iqr",
                        "threshold": 1.5,
                        "highlightColor": "#FF6B6B"
                    }
                }],
                "confidence": 0.9
            }
        }]
    })
}
```

---

## Frontend

### Step 6 — Add the TypeScript type

**File:** `frontend/src/engine/types.ts`

Three changes in this file:

**a)** Add to the `StepAction` union type:

```typescript
export type StepAction =
  | "readRange"
  // ... existing 51 ...
  | "highlightOutliers";
```

**b)** Create the params interface (must match the Pydantic model exactly):

```typescript
export interface HighlightOutliersParams {
  range: string;
  method?: "iqr" | "zscore" | "percentile";
  threshold?: number;
  highlightColor?: string;
}
```

**c)** Add to the `StepParams` union:

```typescript
export type StepParams =
  | ReadRangeParams
  // ... existing ...
  | HighlightOutliersParams;
```

### Step 7 — Add frontend validation

**File:** `frontend/src/engine/validator.ts`

Add a case to the `switch (step.action)` block:

```typescript
case "highlightOutliers":
  requireField(step.id, p, "range", errors);
  break;
```

Only check **required** fields here. Optional fields with defaults don't need
validation — the handler should apply defaults.

### Step 8 — Create the capability handler

**New file:** `frontend/src/engine/capabilities/highlightOutliers.ts`

This is the actual Office.js code that executes the action:

```typescript
/**
 * highlightOutliers – Detect and highlight statistical outliers.
 */

import { CapabilityMeta, HighlightOutliersParams, StepResult, ExecutionOptions } from "../types";
import { registry } from "../capabilityRegistry";
import { resolveRange } from "./rangeUtils";

const meta: CapabilityMeta = {
  action: "highlightOutliers",
  description: "Highlight statistical outliers in a numeric column",
  mutates: false,           // doesn't change cell values
  affectsFormatting: true,  // changes cell formatting
};

async function handler(
  context: Excel.RequestContext,
  params: HighlightOutliersParams,
  options: ExecutionOptions
): Promise<StepResult> {
  const { range: rangeAddr, method = "iqr", threshold = 1.5, highlightColor = "#FF6B6B" } = params;

  if (options.dryRun) {
    return {
      stepId: "",
      status: "success",
      message: `Would highlight outliers in ${rangeAddr} using ${method}`,
    };
  }

  options.onProgress?.(`Analyzing outliers in ${rangeAddr}...`);

  const range = resolveRange(context, rangeAddr);
  range.load("values");
  await context.sync();

  // Extract numeric values and compute bounds
  const values: number[] = [];
  const rows = range.values;
  for (const row of rows) {
    const v = row[0];
    if (typeof v === "number" && !isNaN(v)) values.push(v);
  }

  let lowerBound: number;
  let upperBound: number;

  if (method === "zscore") {
    const mean = values.reduce((a, b) => a + b, 0) / values.length;
    const std = Math.sqrt(values.reduce((s, v) => s + (v - mean) ** 2, 0) / values.length);
    lowerBound = mean - threshold * std;
    upperBound = mean + threshold * std;
  } else {
    // IQR method (default)
    const sorted = [...values].sort((a, b) => a - b);
    const q1 = sorted[Math.floor(sorted.length * 0.25)];
    const q3 = sorted[Math.floor(sorted.length * 0.75)];
    const iqr = q3 - q1;
    lowerBound = q1 - threshold * iqr;
    upperBound = q3 + threshold * iqr;
  }

  // Highlight outlier cells
  let count = 0;
  for (let i = 0; i < rows.length; i++) {
    const v = rows[i][0];
    if (typeof v === "number" && (v < lowerBound || v > upperBound)) {
      const cell = range.getCell(i, 0);
      cell.format.fill.color = highlightColor;
      count++;
    }
  }

  await context.sync();

  return {
    stepId: "",
    status: "success",
    message: `Highlighted ${count} outlier(s) in ${rangeAddr} (${method}, threshold=${threshold})`,
  };
}

registry.register(meta, handler as any);
export { meta };
```

Key patterns:
- Import from `../types` and `../capabilityRegistry`
- Use `resolveRange()` for all range addresses (handles sheet-qualified refs)
- Support `dryRun` — return a preview message without executing
- Call `options.onProgress?.()` for UI feedback
- End with `registry.register(meta, handler as any)` — this is the self-registration
- Set `mutates: true` if the action changes cell **values** (triggers snapshot before execution for undo)
- Set `affectsFormatting: true` if it changes formatting

### Step 9 — Register the import

**File:** `frontend/src/engine/capabilities/index.ts`

```typescript
import "./highlightOutliers";  // ← add at the end
```

This single import triggers the `registry.register()` call at app startup.

---

## Verification checklist

After adding all files:

- [ ] **Backend tests pass:** `cd backend && python -m pytest tests/ -x -q`
- [ ] **Frontend compiles:** `cd frontend && npm run build`
- [ ] **Enum sync check:** count of `StepAction` entries matches on both sides
- [ ] **ACTION_PARAM_MODELS complete:** every enum value has an entry
- [ ] **CAPABILITY_DESCRIPTIONS complete:** every enum value has a description
- [ ] **CAPABILITY_EXAMPLES added:** action is discoverable via semantic search
- [ ] **Manual test:** ask the assistant to perform the new action in Excel

## Quick sync check (run from project root):

```bash
# Backend enum count
grep -c "= \"" backend/app/models/plan.py | head -1

# Frontend type count
grep -c "| \"" frontend/src/engine/types.ts | head -1

# ACTION_PARAM_MODELS count
grep -c "StepAction\." backend/app/models/plan.py

# CAPABILITY_DESCRIPTIONS count
grep -c '":' backend/app/services/planner.py | head -1

# capabilities/index.ts import count
grep -c "import \"./" frontend/src/engine/capabilities/index.ts
```

All five numbers should match.

---

## Common patterns for different action types

### Read-only action (no cell changes)
```typescript
mutates: false, affectsFormatting: false
```
Example: `readRange` — just returns data, no snapshot needed.

### Value-writing action
```typescript
mutates: true, affectsFormatting: false
```
Example: `writeValues`, `writeFormula` — changes cell content, snapshot taken for undo.

### Format-only action
```typescript
mutates: false, affectsFormatting: true
```
Example: `addConditionalFormat`, `setNumberFormat` — changes appearance only.

### Multi-range action
If your action reads from one range and writes to another, use two params:
```python
class MyActionParams(BaseModel):
    sourceRange: str      # where to read
    outputRange: str      # where to write
```

### Action with sheet creation
If your action creates a new sheet, use `dependsOn` in multi-step plans:
```json
{ "id": "step_1", "action": "addSheet", "params": { "sheetName": "Results" } },
{ "id": "step_2", "action": "yourAction", "params": { ... }, "dependsOn": ["step_1"] }
```

---

## File reference (all 9 touch points)

| # | Side | File | What to add |
|---|------|------|------------|
| 1 | Backend | `backend/app/models/plan.py` | `StepAction` enum entry |
| 2 | Backend | `backend/app/models/plan.py` | Pydantic params class |
| 3 | Backend | `backend/app/models/plan.py` | `ACTION_PARAM_MODELS` entry |
| 4 | Backend | `backend/app/services/planner.py` | `CAPABILITY_DESCRIPTIONS` entry |
| 5 | Backend | `backend/app/services/capability_store.py` | `CAPABILITY_EXAMPLES` entries |
| 6 | Frontend | `frontend/src/engine/types.ts` | `StepAction` union + params interface + `StepParams` union |
| 7 | Frontend | `frontend/src/engine/validator.ts` | `case` in switch with `requireField` calls |
| 8 | Frontend | `frontend/src/engine/capabilities/<name>.ts` | New file: handler + `registry.register()` |
| 9 | Frontend | `frontend/src/engine/capabilities/index.ts` | `import "./<name>"` |

Optional: `backend/app/services/example_store.py` — add a few-shot seed example for better LLM accuracy.
