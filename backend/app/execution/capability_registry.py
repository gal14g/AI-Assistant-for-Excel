"""
Python-side capability registry — mirrors `frontend/src/engine/capabilityRegistry.ts`.

Each ported handler calls `registry.register("<action>", handler, mutates=...)`
at import time. The xlwings executor looks up handlers by action name when
iterating through a plan's steps.

Unlike the frontend registry, this one doesn't currently store a separate
fallback handler — xlwings runs on a desktop Excel that always supports the
full Office object model, so API-set gating isn't a concern here. If we ever
want to support headless openpyxl execution as a second backend, the
fallback slot can be added the same way the TS side grew one in Item 3.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Callable, Optional

log = logging.getLogger(__name__)


# Wire-shape for handler functions. Takes an opaque `ctx` (ExecutorContext)
# and a params dict (already Pydantic-validated upstream by ACTION_PARAM_MODELS),
# returns a dict matching `StepResult.to_dict()`.
HandlerFn = Callable[..., dict]


@dataclass
class RegisteredCapability:
    action: str
    handler: HandlerFn
    mutates: bool
    affects_formatting: bool = False


class CapabilityRegistry:
    def __init__(self) -> None:
        self._capabilities: dict[str, RegisteredCapability] = {}

    def register(
        self,
        action: str,
        handler: HandlerFn,
        *,
        mutates: bool,
        affects_formatting: bool = False,
    ) -> None:
        if action in self._capabilities:
            log.warning("Capability %r is being re-registered", action)
        self._capabilities[action] = RegisteredCapability(
            action=action,
            handler=handler,
            mutates=mutates,
            affects_formatting=affects_formatting,
        )

    def get(self, action: str) -> Optional[RegisteredCapability]:
        return self._capabilities.get(action)

    def has(self, action: str) -> bool:
        return action in self._capabilities

    def list_actions(self) -> list[str]:
        return sorted(self._capabilities.keys())

    def mutating_actions(self) -> list[str]:
        return [a for a, c in self._capabilities.items() if c.mutates]


registry = CapabilityRegistry()
