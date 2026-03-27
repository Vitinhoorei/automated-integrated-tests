from __future__ import annotations

from typing import Protocol

from sap_core_next.core.models import ScenarioStep


class EvidenceStore(Protocol):
    def build_path(self, *, execution_id: str, step: ScenarioStep, attempt: int) -> str:
        ...
