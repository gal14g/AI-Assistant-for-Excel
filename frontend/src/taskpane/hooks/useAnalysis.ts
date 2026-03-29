/**
 * useAnalysis — React hook for the analytical pipeline.
 *
 * Detects analytical intent, reads sheet data via Office.js,
 * calls /api/analyze, and returns structured results.
 *
 * Usage:
 *   const { isAnalyzing, analysisResult, runAnalysis, canAnalyze } = useAnalysis();
 *
 *   // In the chat send handler:
 *   if (canAnalyze(userMessage)) {
 *     const result = await runAnalysis(userMessage, conversationHistory);
 *     // display result.message in the chat
 *   }
 */

import { useState, useCallback, useRef } from "react";
import {
  analyzeRequest,
  readAllSheets,
  isAnalyticalRequest,
  AnalyzeResponse,
} from "../../services/analyzeApi";

// ── Types ─────────────────────────────────────────────────────────────────────

export interface UseAnalysisReturn {
  /** True while the analysis pipeline is running. */
  isAnalyzing: boolean;
  /** The last successful analysis result, or null. */
  analysisResult: AnalyzeResponse | null;
  /** The last error message, or null. */
  analysisError: string | null;
  /**
   * Run the full analytical pipeline for *userMessage*.
   * Reads all sheets from the workbook, sends to /api/analyze, returns result.
   * Returns null on error (error is stored in analysisError).
   */
  runAnalysis: (
    userMessage: string,
    conversationHistory?: Array<{ role: string; content: string }>,
    activeSheet?: string,
  ) => Promise<AnalyzeResponse | null>;
  /** Clear result and error state. */
  clearAnalysis: () => void;
  /** Returns true when the message looks like an analytical request. */
  canAnalyze: (message: string) => boolean;
}

// ── Hook ──────────────────────────────────────────────────────────────────────

export function useAnalysis(): UseAnalysisReturn {
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<AnalyzeResponse | null>(null);
  const [analysisError, setAnalysisError] = useState<string | null>(null);
  const abortRef = useRef<AbortController | null>(null);

  const runAnalysis = useCallback(
    async (
      userMessage: string,
      conversationHistory?: Array<{ role: string; content: string }>,
      activeSheet?: string,
    ): Promise<AnalyzeResponse | null> => {
      // Abort any in-flight request
      if (abortRef.current) {
        abortRef.current.abort();
      }
      abortRef.current = new AbortController();

      setIsAnalyzing(true);
      setAnalysisError(null);

      try {
        // Read sheet data from the workbook
        const sheets = await readAllSheets();

        if (Object.keys(sheets).length === 0) {
          throw new Error(
            "No data found in the workbook. Please make sure at least one sheet has data."
          );
        }

        // Call the analytical pipeline
        const result = await analyzeRequest(
          {
            userMessage,
            sheets,
            activeSheet,
            conversationHistory,
          },
          abortRef.current.signal,
        );

        setAnalysisResult(result);
        return result;
      } catch (err: unknown) {
        if (err instanceof Error && err.name === "AbortError") {
          return null; // Silently ignore aborted requests
        }
        const message =
          err instanceof Error ? err.message : "Analysis failed. Please try again.";
        setAnalysisError(message);
        return null;
      } finally {
        setIsAnalyzing(false);
      }
    },
    [],
  );

  const clearAnalysis = useCallback(() => {
    if (abortRef.current) {
      abortRef.current.abort();
    }
    setAnalysisResult(null);
    setAnalysisError(null);
  }, []);

  const canAnalyze = useCallback((message: string) => {
    return isAnalyticalRequest(message);
  }, []);

  return {
    isAnalyzing,
    analysisResult,
    analysisError,
    runAnalysis,
    clearAnalysis,
    canAnalyze,
  };
}
