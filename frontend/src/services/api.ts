/**
 * API service – communicates with the FastAPI backend.
 *
 * - POST /api/plan: send user request, receive ExecutionPlan
 * - POST /api/validate: validate a plan server-side
 * - GET /api/capabilities: list available capabilities
 */

import { ExecutionPlan } from "../engine/types";

export interface ChatRequest {
  userMessage: string;
  rangeTokens?: { address: string; sheetName: string }[];
  activeSheet?: string;
  workbookName?: string;
  conversationHistory?: { role: string; content: string }[];
}

export interface ChatResponse {
  responseType: "message" | "plan";
  message: string;
  plan?: ExecutionPlan;
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
