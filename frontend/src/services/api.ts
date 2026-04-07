/**
 * API service – communicates with the FastAPI backend.
 *
 * - POST /api/chat + /api/chat/stream: main chat endpoint (streaming SSE is primary)
 * - POST /api/feedback: record applied/dismissed plan
 * - /api/conversations/*: persistent conversation history
 */

import { ExecutionPlan, PlanOption } from "../engine/types";

export interface SheetSnapshotDTO {
  sheetName: string;
  rowCount: number;
  columnCount: number;
  headers: string[];
  sampleRows: (string | number | boolean | null)[][];
  dtypes: string[];
  anchorCell: string;
  usedRangeAddress: string;
}

export interface WorkbookSnapshotDTO {
  sheets: SheetSnapshotDTO[];
  truncated: boolean;
}

/** Execution context from a failed plan — enables multi-turn refinement. */
export interface ExecutionContextDTO {
  originalPlanId: string;
  originalUserRequest: string;
  stepResults: {
    stepId: string;
    status: "success" | "error" | "skipped" | "preview";
    message: string;
    error?: string;
  }[];
  failedStepId?: string;
  failedStepAction?: string;
  failedStepError?: string;
}

export interface ChatRequest {
  userMessage: string;
  rangeTokens?: { address: string; sheetName: string }[];
  activeSheet?: string;
  workbookName?: string;
  usedRangeEnd?: string;   // e.g. "C15" — last used cell on the active sheet
  locale?: string;         // e.g. "he-IL" — user's locale for date/number formatting
  conversationHistory?: { role: string; content: string }[];
  conversationId?: string;
  userMessageId?: string;
  /** Lightweight snapshot of every visible sheet: headers, row/col counts,
   *  sample rows, inferred dtypes. Lets the planner ground its plan in the
   *  actual data instead of guessing column names. */
  workbookSnapshot?: WorkbookSnapshotDTO;
  /** Multi-turn refinement: execution state from a failed plan. */
  executionContext?: ExecutionContextDTO;
}

export interface ChatResponse {
  responseType: "message" | "plan" | "plans";
  message: string;
  plan?: ExecutionPlan;
  plans?: PlanOption[];
  interactionId?: string;
  conversationId?: string;
  assistantMessageId?: string;
}

// Empty base — all /api calls go through the webpack dev-server proxy
// (https://localhost:3000/api → http://localhost:8000/api).
// In production replace with your real backend origin.
const BASE_URL = "";

/**
 * Send a message to the conversational chat AI.
 * Returns either a plain message or an execution plan, depending on the request.
 */
export async function sendChatMessage(
  request: ChatRequest,
  signal?: AbortSignal
): Promise<ChatResponse> {
  const response = await fetch(`${BASE_URL}/api/chat`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(request),
    signal,
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Chat request failed (${response.status}): ${error}`);
  }

  return response.json();
}

/**
 * Streaming version of sendChatMessage.
 *
 * Connects to POST /api/chat/stream (SSE endpoint).
 * Calls onChunk(text) for each partial token as the LLM generates it.
 * Returns the final ChatResponse once the stream ends.
 */
export async function sendChatMessageStream(
  request: ChatRequest,
  signal: AbortSignal | undefined,
  onChunk: (text: string) => void,
  onReset?: () => void,
): Promise<ChatResponse> {
  const res = await fetch(`${BASE_URL}/api/chat/stream`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(request),
    signal,
  });

  if (!res.ok || !res.body) {
    const err = await res.text().catch(() => String(res.status));
    throw new Error(`Chat stream failed (${res.status}): ${err}`);
  }

  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  let result: ChatResponse | null = null;

  // eslint-disable-next-line no-constant-condition
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;

    buffer += decoder.decode(value, { stream: true });
    // SSE events are separated by double newlines
    const parts = buffer.split("\n\n");
    buffer = parts.pop() ?? "";

    for (const part of parts) {
      const line = part.trim();
      if (!line.startsWith("data: ")) continue;
      try {
        const data = JSON.parse(line.slice(6)) as { type: string; text?: string; result?: ChatResponse };
        if (data.type === "chunk" && data.text) {
          onChunk(data.text);
        } else if (data.type === "reset") {
          onReset?.();
        } else if (data.type === "done" && data.result) {
          result = data.result;
        }
      } catch {
        // partial / malformed SSE line — skip
      }
    }
  }

  if (!result) throw new Error("Stream ended without a result event");
  return result;
}

/**
 * Record the user's choice (applied or dismissed) for an interaction.
 * Fire-and-forget — errors are silently swallowed.
 */
export async function sendFeedback(
  interactionId: string,
  chosenPlanId: string | null,
  action: "applied" | "dismissed",
): Promise<void> {
  try {
    await fetch(`${BASE_URL}/api/feedback`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ interactionId, chosenPlanId, action }),
    });
  } catch {
    // Fire-and-forget: feedback logging should never disrupt UX
  }
}

// ── Conversations (server-persisted chat history) ───────────────────────────

export interface ConversationSummary {
  id: string;
  title: string;
  createdAt: string;
  updatedAt: string;
  messageCount: number;
}

export interface PersistedMessage {
  id: string;
  role: string;
  content: string;
  timestamp: string;
  rangeTokens?: { address: string; sheetName: string }[] | null;
  plan?: ExecutionPlan | null;
  execution?: unknown;
  progressLog?: { stepId: string; message: string; timestamp: string }[] | null;
}

export interface ConversationDetail {
  id: string;
  title: string;
  createdAt: string;
  updatedAt: string;
  messages: PersistedMessage[];
}

export async function listConversations(): Promise<ConversationSummary[]> {
  const res = await fetch(`${BASE_URL}/api/conversations`);
  if (!res.ok) throw new Error(`listConversations failed: ${res.status}`);
  return res.json();
}

export async function getConversation(id: string): Promise<ConversationDetail> {
  const res = await fetch(`${BASE_URL}/api/conversations/${id}`);
  if (!res.ok) throw new Error(`getConversation failed: ${res.status}`);
  return res.json();
}

export async function renameConversation(id: string, title: string): Promise<ConversationSummary> {
  const res = await fetch(`${BASE_URL}/api/conversations/${id}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ title }),
  });
  if (!res.ok) throw new Error(`renameConversation failed: ${res.status}`);
  return res.json();
}

export async function deleteConversation(id: string): Promise<void> {
  const res = await fetch(`${BASE_URL}/api/conversations/${id}`, { method: "DELETE" });
  if (!res.ok) throw new Error(`deleteConversation failed: ${res.status}`);
}

export async function patchMessageExecution(
  conversationId: string,
  messageId: string,
  execution: unknown,
  progressLog: unknown,
): Promise<void> {
  try {
    await fetch(`${BASE_URL}/api/conversations/${conversationId}/messages/${messageId}`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ execution, progressLog }),
    });
  } catch {
    // fire-and-forget
  }
}

export async function popLastExchange(conversationId: string): Promise<number> {
  try {
    const res = await fetch(`${BASE_URL}/api/conversations/${conversationId}/last`, {
      method: "DELETE",
    });
    if (!res.ok) return 0;
    const body = await res.json();
    return body.removed ?? 0;
  } catch {
    return 0;
  }
}

// ── Presets ──────────────────────────────────────────────────────────────────
// ── Presets (stored in browser localStorage — per-user, no server) ───────────

export interface Preset {
  id: string;
  name: string;
  userMessage: string;
  assistantResponse?: string;
  createdAt: string;
}

const PRESETS_KEY = "excel_assistant_presets";

function readPresetsFromStorage(): Preset[] {
  try {
    const raw = localStorage.getItem(PRESETS_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch {
    return [];
  }
}

function writePresetsToStorage(presets: Preset[]): void {
  localStorage.setItem(PRESETS_KEY, JSON.stringify(presets));
}

export async function listPresets(): Promise<Preset[]> {
  return readPresetsFromStorage();
}

export async function getPreset(id: string): Promise<Preset | null> {
  return readPresetsFromStorage().find((p) => p.id === id) ?? null;
}

export async function savePreset(name: string, userMessage: string, assistantResponse: string): Promise<{ id: string }> {
  const presets = readPresetsFromStorage();
  const id = crypto.randomUUID?.() ?? `preset_${Date.now()}`;
  presets.unshift({ id, name, userMessage, assistantResponse, createdAt: new Date().toISOString() });
  writePresetsToStorage(presets);
  return { id };
}

export async function renamePreset(id: string, name: string): Promise<void> {
  const presets = readPresetsFromStorage();
  const preset = presets.find((p) => p.id === id);
  if (preset) {
    preset.name = name;
    writePresetsToStorage(presets);
  }
}

export async function deletePreset(id: string): Promise<void> {
  const presets = readPresetsFromStorage().filter((p) => p.id !== id);
  writePresetsToStorage(presets);
}
