/**
 * SSE (Server-Sent Events) streaming service.
 *
 * Used for receiving real-time execution updates from the backend
 * during plan generation (streaming LLM responses) and execution monitoring.
 */

import { PlanRequest } from "./api";

// Empty base — proxied through webpack dev server (same as api.ts).
const BASE_URL = "";

export interface StreamEvent {
  type: "planning" | "explanation" | "plan_ready" | "error" | "done";
  data: string;
}

export type StreamCallback = (event: StreamEvent) => void;

/**
 * Start an SSE stream for plan generation.
 * The backend streams explanation tokens as they arrive from the LLM,
 * then sends the final plan as a single JSON event.
 */
export function streamPlanGeneration(
  request: PlanRequest,
  onEvent: (event: StreamEvent) => void,
  onError?: (error: Error) => void
): AbortController {
  const controller = new AbortController();

  (async () => {
    try {
      const response = await fetch(`${BASE_URL}/api/plan/stream`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(request), // ✅ fixed
        signal: controller.signal,
      });

      if (!response.ok) {
        throw new Error(`Stream request failed: ${response.status}`);
      }

      const reader = response.body?.getReader();
      if (!reader) throw new Error("No response body");

      const decoder = new TextDecoder();
      let buffer = "";

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });

        const lines = buffer.split("\n");
        buffer = lines.pop() ?? "";

        for (const line of lines) {
          if (line.startsWith("data: ")) {
            const data = line.slice(6);
            if (data === "[DONE]") {
              onEvent({ type: "done", data: "" });
              return;
            }
            try {
              const parsed = JSON.parse(data) as StreamEvent;
              onEvent(parsed);
            } catch {
              onEvent({ type: "explanation", data });
            }
          }
        }
      }

      onEvent({ type: "done", data: "" });
    } catch (err) {
      if ((err as Error).name === "AbortError") return;
      onError?.(err as Error);
    }
  })();

  return controller;
}

/**
 * Open an SSE connection for execution monitoring.
 * Receives step-by-step progress updates during plan execution.
 */
export function streamExecutionUpdates(
  planId: string,
  onEvent: StreamCallback,
  onError?: (error: Error) => void
): EventSource {
  const source = new EventSource(`${BASE_URL}/api/execution/${planId}/stream`);

  source.onmessage = (event) => {
    try {
      const parsed = JSON.parse(event.data) as StreamEvent;
      onEvent(parsed);
    } catch {
      onEvent({ type: "explanation", data: event.data });
    }
  };

  source.onerror = () => {
    onError?.(new Error("SSE connection error"));
    source.close();
  };

  return source;
}
