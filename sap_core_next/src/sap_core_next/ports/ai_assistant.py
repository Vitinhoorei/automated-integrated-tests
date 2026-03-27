from __future__ import annotations

from typing import Mapping, Protocol

from sap_core_next.core.models import AutoHealSuggestion


class AIAssistant(Protocol):
    def suggest_fix(self, *, module: str, command: str, status_message: str, parameters: Mapping[str, str], dump_path: str | None = None) -> AutoHealSuggestion:
        ...
