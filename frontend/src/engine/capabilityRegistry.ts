/**
 * Capability Registry
 *
 * Central registry mapping action names to their handlers and metadata.
 * This is the single source of truth for what operations the executor can perform.
 * New capabilities are registered here, making the system extensible.
 */

import {
  StepAction,
  CapabilityMeta,
  CapabilityHandler,
} from "./types";

interface RegisteredCapability {
  meta: CapabilityMeta;
  handler: CapabilityHandler;
}

class CapabilityRegistry {
  private capabilities = new Map<StepAction, RegisteredCapability>();

  register(meta: CapabilityMeta, handler: CapabilityHandler): void {
    if (this.capabilities.has(meta.action)) {
      console.warn(`Capability "${meta.action}" is being re-registered.`);
    }
    this.capabilities.set(meta.action, { meta, handler });
  }

  get(action: StepAction): RegisteredCapability | undefined {
    return this.capabilities.get(action);
  }

  has(action: StepAction): boolean {
    return this.capabilities.has(action);
  }

  getHandler(action: StepAction): CapabilityHandler | undefined {
    return this.capabilities.get(action)?.handler;
  }

  getMeta(action: StepAction): CapabilityMeta | undefined {
    return this.capabilities.get(action)?.meta;
  }

  /** Get all registered action names */
  listActions(): StepAction[] {
    return Array.from(this.capabilities.keys());
  }

  /** Get metadata for all capabilities (useful for planner context) */
  listCapabilities(): CapabilityMeta[] {
    return Array.from(this.capabilities.values()).map((c) => c.meta);
  }

  /** Check which actions mutate the workbook (need snapshots) */
  getMutatingActions(): StepAction[] {
    return Array.from(this.capabilities.values())
      .filter((c) => c.meta.mutates)
      .map((c) => c.meta.action);
  }

  /** Check which actions affect formatting */
  getFormattingActions(): StepAction[] {
    return Array.from(this.capabilities.values())
      .filter((c) => c.meta.affectsFormatting)
      .map((c) => c.meta.action);
  }
}

/** Singleton registry instance */
export const registry = new CapabilityRegistry();
