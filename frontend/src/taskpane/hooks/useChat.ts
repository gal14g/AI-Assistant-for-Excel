/**
 * useChat – Manages chat state, message history, and plan lifecycle.
 *
 * All messages go through POST /api/chat, which uses a conversational AI
 * to decide whether to reply with text or generate an execution plan.
 */

import { useState, useCallback, useRef, useEffect } from "react";
import { ChatMessage, ExecutionPlan, ExecutionState, PlanOption } from "../../engine/types";
import {
  sendChatMessageStream,
  ChatRequest,
  getConversation,
  popLastExchange as apiPopLastExchange,
  deleteConversation as apiDeleteConversation,
} from "../../services/api";
import { v4 as uuid } from "uuid";

const LS_CONV_ID_KEY = "excel_copilot_active_conversation_id";

interface ChatState {
  messages: ChatMessage[];
  isLoading: boolean;
  streamingText: string;
  currentPlan: ExecutionPlan | null;
  currentOptions: PlanOption[] | null;
  interactionId: string | null;
  error: string | null;
  conversationId: string | null;
}

interface ChatActions {
  sendMessage: (text: string, rangeTokens?: { address: string; sheetName: string }[]) => Promise<void>;
  stopMessage: () => void;
  clearHistory: () => void;
  /** Remove the last user + assistant message pair (used by undo). Returns the user message text. */
  removeLastExchange: () => string;
  /** Patch an existing message by id — used to attach execution state/log. */
  updateMessage: (id: string, patch: Partial<ChatMessage>) => void;
  /** ID of the most recent assistant message that carries a plan, or null. */
  getLatestPlanMessageId: () => string | null;
  setCurrentPlan: (plan: ExecutionPlan | null) => void;
  setCurrentOptions: (options: PlanOption[] | null) => void;
  dismissError: () => void;
  /** Load a persisted conversation by ID, replacing current chat state. */
  loadConversation: (id: string) => Promise<void>;
  /** Delete the active conversation on the server and clear local state. */
  deleteCurrentConversation: () => Promise<void>;
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
  const [streamingText, setStreamingText] = useState("");
  const [currentPlan, setCurrentPlan] = useState<ExecutionPlan | null>(null);
  const [currentOptions, setCurrentOptions] = useState<PlanOption[] | null>(null);
  const [interactionId, setInteractionId] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [conversationId, setConversationId] = useState<string | null>(null);
  const abortRef = useRef<AbortController | null>(null);
  const conversationIdRef = useRef<string | null>(null);
  conversationIdRef.current = conversationId;

  // Persist active conversation id across reloads
  useEffect(() => {
    if (conversationId) localStorage.setItem(LS_CONV_ID_KEY, conversationId);
    else localStorage.removeItem(LS_CONV_ID_KEY);
  }, [conversationId]);

  const sendMessage = useCallback(
    async (text: string, rangeTokens?: { address: string; sheetName: string }[]) => {
      const userMessageId = uuid();
      const userMsg: ChatMessage = {
        id: userMessageId,
        role: "user",
        content: text,
        rangeTokens,
        timestamp: new Date().toISOString(),
      };

      // Cancel any in-flight request before starting a new one
      abortRef.current?.abort();
      abortRef.current = new AbortController();

      setMessages((prev) => [...prev, userMsg]);
      setIsLoading(true);
      setError(null);

      try {
        // Build conversation history (exclude system messages, keep last 10)
        const history = messages
          .filter((m) => m.role !== "system")
          .slice(-10)
          .map((m) => ({ role: m.role, content: m.content }));

        // Get active sheet, workbook name, and used range from Excel context
        let activeSheet = "";
        let workbookName = "";
        let usedRangeEnd = "";
        try {
          await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.load("name");
            context.workbook.load("name");
            const usedRange = sheet.getUsedRangeOrNullObject(true);
            usedRange.load("address");
            await context.sync();
            activeSheet = sheet.name;
            workbookName = context.workbook.name ?? "";
            if (!usedRange.isNullObject && usedRange.address) {
              // address is like "Sheet1!A1:C15" — extract just the end cell
              const addr = usedRange.address.split("!").pop() ?? "";
              const endCell = addr.includes(":") ? addr.split(":")[1] : addr;
              usedRangeEnd = endCell;
            }
          });
        } catch {
          // Fallback when running outside Excel context
        }

        const request: ChatRequest = {
          userMessage: text,
          rangeTokens,
          activeSheet,
          workbookName: workbookName || undefined,
          usedRangeEnd: usedRangeEnd || undefined,
          locale: navigator.language || undefined,
          conversationHistory: history,
          conversationId: conversationIdRef.current ?? undefined,
          userMessageId,
        };

        setStreamingText("");
        const response = await sendChatMessageStream(
          request,
          abortRef.current?.signal,
          (chunk) => setStreamingText((prev) => prev + chunk),
        );
        setStreamingText("");

        if (response.conversationId && response.conversationId !== conversationIdRef.current) {
          setConversationId(response.conversationId);
        }

        const assistantMsg: ChatMessage = {
          id: response.assistantMessageId ?? uuid(),
          role: "assistant",
          content: response.message,
          plan: response.plans?.[0]?.plan ?? response.plan,
          timestamp: new Date().toISOString(),
        };

        setMessages((prev) => [...prev, assistantMsg]);
        setInteractionId(response.interactionId ?? null);

        if (response.responseType === "plans" && response.plans?.length) {
          setCurrentOptions(response.plans);
          setCurrentPlan(null);
        } else if (response.responseType === "plan" && response.plan) {
          setCurrentOptions(null);
          setCurrentPlan(response.plan);
        }
      } catch (err) {
        // Ignore aborts — they're intentional (e.g. user clicked "New chat")
        if (err instanceof Error && err.name === "AbortError") return;
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

  const stopMessage = useCallback(() => {
    abortRef.current?.abort();
    abortRef.current = null;
    setIsLoading(false);
    setStreamingText("");
  }, []);

  /** Remove the last user + assistant pair and return the user message text. */
  const removeLastExchange = useCallback((): string => {
    const cid = conversationIdRef.current;
    if (cid) {
      // Fire-and-forget: remove server-side too
      void apiPopLastExchange(cid);
    }
    let removedUserText = "";
    setMessages((prev) => {
      const copy = [...prev];
      // Remove last assistant
      for (let i = copy.length - 1; i >= 0; i--) {
        if (copy[i].role === "assistant") { copy.splice(i, 1); break; }
      }
      // Remove last user and capture its text
      for (let i = copy.length - 1; i >= 0; i--) {
        if (copy[i].role === "user") {
          removedUserText = copy[i].content;
          copy.splice(i, 1);
          break;
        }
      }
      return copy;
    });
    setCurrentPlan(null);
    setCurrentOptions(null);
    setInteractionId(null);
    return removedUserText;
  }, []);

  const clearHistory = useCallback(() => {
    // Cancel any in-flight request immediately
    abortRef.current?.abort();
    abortRef.current = null;
    setIsLoading(false);
    setMessages([
      {
        id: uuid(),
        role: "system",
        content: "Chat cleared. How can I help you?",
        timestamp: new Date().toISOString(),
      },
    ]);
    setCurrentPlan(null);
    setCurrentOptions(null);
    setInteractionId(null);
    setError(null);
    // Starting a fresh chat = new conversation; server will mint a new id on next send.
    setConversationId(null);
  }, []);

  const loadConversation = useCallback(async (id: string) => {
    abortRef.current?.abort();
    abortRef.current = null;
    setIsLoading(false);
    setError(null);
    try {
      const conv = await getConversation(id);
      const loaded: ChatMessage[] = conv.messages.map((m) => ({
        id: m.id,
        role: m.role as ChatMessage["role"],
        content: m.content,
        rangeTokens: m.rangeTokens ?? undefined,
        plan: (m.plan as ExecutionPlan | null) ?? undefined,
        execution: (m.execution as ExecutionState | undefined) ?? undefined,
        progressLog: m.progressLog ?? undefined,
        timestamp: m.timestamp,
      }));
      setMessages(
        loaded.length
          ? loaded
          : [{
              id: uuid(),
              role: "system",
              content: "Conversation restored.",
              timestamp: new Date().toISOString(),
            }],
      );
      setConversationId(id);
      setCurrentPlan(null);
      setCurrentOptions(null);
      setInteractionId(null);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to load conversation");
    }
  }, []);

  const deleteCurrentConversation = useCallback(async () => {
    const cid = conversationIdRef.current;
    if (cid) {
      try { await apiDeleteConversation(cid); } catch { /* ignore */ }
    }
    clearHistory();
  }, [clearHistory]);

  // On mount, try to restore the last active conversation
  useEffect(() => {
    const saved = localStorage.getItem(LS_CONV_ID_KEY);
    if (saved) {
      void loadConversation(saved).catch(() => {
        localStorage.removeItem(LS_CONV_ID_KEY);
      });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const dismissError = useCallback(() => setError(null), []);

  const updateMessage = useCallback((id: string, patch: Partial<ChatMessage>) => {
    setMessages((prev) => prev.map((m) => (m.id === id ? { ...m, ...patch } : m)));
  }, []);

  // Use a ref to messages so the returned getter doesn't go stale in closures.
  const messagesRef = useRef(messages);
  messagesRef.current = messages;
  const getLatestPlanMessageId = useCallback((): string | null => {
    const list = messagesRef.current;
    for (let i = list.length - 1; i >= 0; i--) {
      if (list[i].role === "assistant" && list[i].plan) return list[i].id;
    }
    return null;
  }, []);

  return {
    messages,
    isLoading,
    streamingText,
    currentPlan,
    currentOptions,
    interactionId,
    error,
    conversationId,
    sendMessage,
    stopMessage,
    clearHistory,
    removeLastExchange,
    updateMessage,
    getLatestPlanMessageId,
    setCurrentPlan,
    setCurrentOptions,
    dismissError,
    loadConversation,
    deleteCurrentConversation,
  };
}
