from __future__ import annotations

from dataclasses import dataclass
from typing import Mapping

from sap_core_next.core.models import GatewayResult, StepStatus


@dataclass
class LegacyBridgeGateway:
    """
    Bridge para reaproveitar executor legado sem editar a base antiga.
    """

    field_map_path: str = "configs/field_map.yaml"

    def __post_init__(self) -> None:
        # Import tardio para manter novo core independente em tempo de import.
        from src.sap_automation import SapAutomation  # type: ignore

        self._legacy = SapAutomation(field_map_path=self.field_map_path)

    def execute(self, command: str, parameters: Mapping[str, str], *, explanation: str = "", mode: str = "real", evidence_path: str = "") -> GatewayResult:
        result = self._legacy.run_tcode(
            command,
            dict(parameters),
            explanation,
            evidence_path=evidence_path,
            mode=mode,
            shared_context={},
        )
        status = StepStatus.PASS if result.status.upper() == "PASS" else StepStatus.FAIL
        return GatewayResult(
            status=status,
            source=result.source,
            message=result.message,
            evidence_path=result.evidence_path,
            raw={"legacy": True},
        )
