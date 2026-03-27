from __future__ import annotations

from typing import Mapping, Protocol

from sap_core_next.core.models import GatewayResult


class SapGateway(Protocol):
    def execute(self, command: str, parameters: Mapping[str, str], *, explanation: str = "", mode: str = "real", evidence_path: str = "") -> GatewayResult:
        ...
