from __future__ import annotations

from typing import Mapping, Protocol

from sap_core_next.core.context import ExecutionContext
from sap_core_next.core.models import AutoHealSuggestion, ScenarioStep


class ModulePlugin(Protocol):
    name: str

    def supports(self, module: str) -> bool:
        ...

    def normalize_step(self, step: ScenarioStep, ctx: ExecutionContext) -> ScenarioStep:
        ...

    def prepare_parameters(self, step: ScenarioStep, ctx: ExecutionContext) -> dict[str, str]:
        ...

    def apply_suggestion(self, params: dict[str, str], suggestion: AutoHealSuggestion) -> dict[str, str]:
        ...

    def on_success(self, step: ScenarioStep, message: str, params: Mapping[str, str], ctx: ExecutionContext) -> None:
        ...
