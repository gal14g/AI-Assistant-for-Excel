/**
 * useChat – Manages chat state, message history, and plan lifecycle.
 *
 * All messages go through POST /api/chat, which uses a conversational AI
 * to decide whether to reply with text or generate an execution plan.
 */

import { useState, useCallback } from "react";
import { ChatMessage, ExecutionPlan } from "../../engine/types";
import { sendChatMessage, ChatRequest } from "../../services/api";
import { v4 as uuid } from "uuid";

interface ChatState {
  messages: ChatMessage[];
  isLoading: boolean;
  currentPlan: ExecutionPlan | null;
  streamingText: string;
  error: string | null;
}

interface ChatActions {
  sendMessage: (text: string, rangeTokens?: { address: string; sheetName: string }[]) => Promise<void>;
  clearHistory: () => void;
  setCurrentPlan: (plan: ExecutionPlan | null) => void;
  dismissError: () => void;
}

export function useChat(): ChatState & ChatActions {
  const [messages, setMessages] = useState<ChatMessage[]>([
    {
      id: uuid(),
      role: "system",
      content:
        "Hi! I'm Excel AI Copilot. Ask me anything about Excel, or tell me what you'd like to do with your spreadsheet.",
      timestamp: new Date().toISOString(),
    },
  ]);
  const [isLoading, setIsLoading] = useState(false);
  const [currentPlan, setCurrentPlan] = useState<ExecutionPlan | null>(null);
  const [streamingText] = useState("");
  const [error, setError] = useState<string | null>(null);

  const sendMessage = useCallback(
    async (text: string, rangeTokens?: { address: string; sheetName: string }[]) => {
      const userMsg: ChatMessage = {
        id: uuid(),
        role: "user",
        content: text,
        rangeTokens,
        timestamp: new Date().toISOString(),
      };

      setMessages((prev) => [...prev, userMsg]);
      setIsLoading(true);
      setError(null);

      try {
        // Build conversation history (exclude system messages, keep last 10)
        const history = messages
          .filter((m) => m.role !== "system")
          .slice(-10)
          .map((m) => ({ role: m.role, content: m.content }));

        // Get active sheet + workbook name from Excel context
        let activeSheet = "Sheet1";
        let workbookName = "";
        try {
          await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.load("name");
            context.workbook.load("name");
            await context.sync();
            activeSheet = sheet.name;
            workbookName = context.workbook.name ?? "";
          });
        } catch {
          // Fallback when running outside Excel context
        }

        const request: ChatRequest = {
          userMessage: text,
          rangeTokens,
          activeSheet,
          workbookName: workbookName || undefined,
          conversationHistory: history,
        };

        const response = await sendChatMessage(request);

        const assistantMsg: ChatMessage = {
          id: uuid(),
          role: "assistant",
          content: response.message,
          plan: response.plan,
          timestamp: new Date().toISOString(),
        };

        setMessages((prev) => [...prev, assistantMsg]);

        if (response.responseType === "plan" && response.plan) {
          setCurrentPlan(response.plan);
        }
      } catch (err) {
        const errorMsg = err instanceof Error ? err.message : "Failed to get response";
        setError(errorMsg);

        setMessages((prev) => [
          ...prev,
          {
            id: uuid(),
            role: "assistant",
            content: `Sorry, I encountered an error: ${errorMsg}`,
            timestamp: new Date().toISOString(),
          },
        ]);
      } finally {
        setIsLoading(false);
      }
    },
    [messages]
  );

  const clearHistory = useCallback(() => {
    setMessages([
      {
        id: uuid(),
        role: "system",
        content: "Chat cleared. How can I help you?",
        timestamp: new Date().toISOString(),
      },
    ]);
    setCurrentPlan(null);
    setError(null);
  }, []);

  const dismissError = useCallback(() => setError(null), []);

  return {
    messages,
    isLoading,
    currentPlan,
    streamingText,
    error,
    sendMessage,
    clearHistory,
    setCurrentPlan,
    dismissError,
  };
}
