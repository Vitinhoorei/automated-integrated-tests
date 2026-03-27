from __future__ import annotations

from typing import Mapping

from sap_core_next.core.models import AutoHealSuggestion


class NoOpAIAssistant:
    def suggest_fix(self, *, module: str, command: str, status_message: str, parameters: Mapping[str, str], dump_path: str | None = None) -> AutoHealSuggestion:
        return AutoHealSuggestion(should_retry=False, confidence=0, reason="AI disabled")
