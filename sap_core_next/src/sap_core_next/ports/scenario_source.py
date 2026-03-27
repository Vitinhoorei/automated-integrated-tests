from __future__ import annotations

from typing import Protocol, List

from sap_core_next.core.models import ScenarioStep


class ScenarioSource(Protocol):
    def read_steps(self) -> List[ScenarioStep]:
        ...
