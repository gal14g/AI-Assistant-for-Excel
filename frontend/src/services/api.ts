/**
 * API service – communicates with the FastAPI backend.
 *
 * - POST /api/plan: send user request, receive ExecutionPlan
 * - POST /api/validate: validate a plan server-side
 * - GET /api/capabilities: list available capabilities
 */

import { ExecutionPlan, PlanOption } from "../engine/types";

export interface ChatRequest {
  userMessage: string;
  rangeTokens?: { address: string; sheetName: string }[];
  activeSheet?: string;
  workbookName?: string;
  usedRangeEnd?: string;   // e.g. "C15" — last used cell on the active sheet
  locale?: string;         // e.g. "he-IL" — user's locale for date/number formatting
  conversationHistory?: { role: string; content: string }[];
}

export interface ChatResponse {
  responseType: "message" | "plan" | "plans";
  message: string;
  plan?: ExecutionPlan;
  plans?: PlanOption[];
  interactionId?: string;
}

// Empty base — all /api calls go through the webpack dev-server proxy
// (https://localhost:3000/api → http://localhost:8000/api).
// In production replace with your real backend origin.
const BASE_URL = "";

export interface PlanRequest {
  userMessage: string;
  /** Range tokens extracted from [[...]] markers in the message */
  rangeTokens?: { address: string; sheetName: string }[];
  /** Name of the currently active worksheet, e.g. "Sheet1" */
  activeSheet?: string;
  /** Filename of the open workbook, e.g. "Sales.xlsx" */
  workbookName?: string;
  /** Full file path of the workbook – useful for cross-file references */
  workbookPath?: string;
  /** Previous messages for multi-turn context */
  conversationHistory?: { role: string; content: string }[];
}

export interface PlanResponse {
  plan: ExecutionPlan;
  explanation: string;
  alternatives?: string[];
}

export interface ValidationResponse {
  valid: boolean;
  errors: { message: string; code: string }[];
  warnings: { message: string; code: string }[];
}

/**
 * Request a plan from the backend LLM planner.
 */
export async function requestPlan(request: PlanRequest): Promise<PlanResponse> {
  const response = await fetch(`${BASE_URL}/api/plan`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(request),
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Plan request failed (${response.status}): ${error}`);
  }

  return response.json();
}

/**
 * Validate a plan server-side.
 */
export async function validatePlanRemote(plan: ExecutionPlan): Promise<ValidationResponse> {
  const response = await fetch(`${BASE_URL}/api/validate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(plan),
  });

  if (!response.ok) {
    throw new Error(`Validation request failed: ${response.status}`);
  }

  return response.json();
}

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
 * Fetch the list of available capabilities from the backend.
 */
export async function fetchCapabilities(): Promise<{ action: string; description: string }[]> {
  const response = await fetch(`${BASE_URL}/api/capabilities`);
  if (!response.ok) throw new Error("Failed to fetch capabilities");
  return response.json();
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

// ── Presets ──────────────────────────────────────────────────────────────────
// ── Presets (stored in browser localStorage — per-user, no server) ───────────

export interface Preset {
  id: string;
  name: string;
  userMessage: string;
  assistantResponse?: string;
  createdAt: string;
}

const PRESETS_KEY = "excel_copilot_presets";

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
