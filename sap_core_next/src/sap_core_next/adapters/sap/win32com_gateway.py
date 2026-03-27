from __future__ import annotations

from dataclasses import dataclass
from typing import Mapping

from sap_core_next.core.models import GatewayResult, StepStatus


@dataclass
class Win32ComSapGateway:
    """
    Adapter SAP real mínimo (esqueleto) atrás do contrato do core.

    Observação: execução de fluxos complexos deve ficar nos plugins/commands.
    """

    def __post_init__(self) -> None:
        import win32com.client  # type: ignore

        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        app = sap_gui_auto.GetScriptingEngine
        connection = app.Children(0)
        self.session = connection.Children(0)

    def execute(self, command: str, parameters: Mapping[str, str], *, explanation: str = "", mode: str = "real", evidence_path: str = "") -> GatewayResult:
        try:
            # Comportamento mínimo genérico: abrir TCODE e executar enter.
            self.session.findById("wnd[0]/tbar[0]/okcd").text = f"/n{command.upper()}"
            self.session.findById("wnd[0]").sendVKey(0)

            # Parametrização genérica simples em formato SAP_ID direto
            for key, value in parameters.items():
                if key.startswith("wnd["):
                    obj = self.session.findById(key)
                    try:
                        obj.text = str(value)
                    except Exception:
                        obj.key = str(value)

            self.session.findById("wnd[0]").sendVKey(0)
            msg = self.session.findById("wnd[0]/sbar").Text or "Executed"
            mtype = (self.session.findById("wnd[0]/sbar").MessageType or "").upper()
            if mtype in {"E", "A", "X"}:
                return GatewayResult(StepStatus.FAIL, "STATUSBAR", msg, evidence_path)
            return GatewayResult(StepStatus.PASS, "OK", msg, evidence_path)
        except Exception as exc:
            return GatewayResult(StepStatus.FAIL, "EXCEPTION", str(exc), evidence_path)
