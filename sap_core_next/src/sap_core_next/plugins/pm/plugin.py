from __future__ import annotations

import re
from dataclasses import dataclass

from sap_core_next.core.context import ExecutionContext
from sap_core_next.core.models import AutoHealSuggestion, ScenarioStep


@dataclass
class PMPlugin:
    name: str = "pm"

    def supports(self, module: str) -> bool:
        return (module or "").strip().lower() == "pm"

    def normalize_step(self, step: ScenarioStep, ctx: ExecutionContext) -> ScenarioStep:
        step.command = step.command.strip().upper()
        return step

    def prepare_parameters(self, step: ScenarioStep, ctx: ExecutionContext) -> dict[str, str]:
        params = dict(step.parameters)

        # Exemplo representativo de memória de contexto PM (ordem/nota)
        if step.command in {"IW32", "IW41"} and "Ordem" not in params and ctx.get("Ordem"):
            params["Ordem"] = str(ctx.get("Ordem"))
        if step.command == "IW31" and "Nota" not in params and ctx.get("Nota"):
            params["Nota"] = str(ctx.get("Nota"))

        return params

    def apply_suggestion(self, params: dict[str, str], suggestion: AutoHealSuggestion) -> dict[str, str]:
        out = dict(params)
        if not suggestion.suggested_parameter or "=" not in suggestion.suggested_parameter:
            return out
        k, v = suggestion.suggested_parameter.split("=", 1)
        out[k.strip()] = v.strip()
        return out

    def on_success(self, step: ScenarioStep, message: str, params: dict[str, str], ctx: ExecutionContext) -> None:
        # Extração simples de IDs para memória compartilhada
        m = re.search(r"(?:ordem|nota|aviso)\s+(\d+)", message or "", re.IGNORECASE)
        if not m:
            return
        found = m.group(1)
        ctx.set("UltimoID", found)
        if re.search(r"nota|aviso", message or "", re.IGNORECASE):
            ctx.set("Nota", found)
        if re.search(r"ordem", message or "", re.IGNORECASE):
            ctx.set("Ordem", found)
