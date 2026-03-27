from __future__ import annotations

import json
import re
from dataclasses import dataclass
from typing import Mapping

import requests

from sap_core_next.core.models import AutoHealSuggestion


@dataclass
class HttpAIAssistant:
    api_url: str
    api_key: str
    model: str = "gpt-4o-mini"
    timeout_seconds: int = 60

    def suggest_fix(self, *, module: str, command: str, status_message: str, parameters: Mapping[str, str], dump_path: str | None = None) -> AutoHealSuggestion:
        prompt = (
            "Analise falha SAP e retorne JSON com campos should_retry(bool), confidence(int), "
            "suggested_parameter(str opcional), reason(str).\\n"
            f"Modulo={module}\\nComando={command}\\nErro={status_message}\\nParams={dict(parameters)}"
        )
        payload = {"question": prompt, "model": self.model, "temperature": 0}
        headers = {"Content-Type": "application/json", "apiKey": self.api_key}

        try:
            resp = requests.post(self.api_url, json=payload, headers=headers, timeout=self.timeout_seconds)
            if resp.status_code != 200:
                return AutoHealSuggestion(False, 0, reason=f"HTTP {resp.status_code}")

            text = resp.json().get("response", "")
            match = re.search(r"\{.*\}", text, re.DOTALL)
            data = json.loads(match.group(0) if match else "{}")
            return AutoHealSuggestion(
                should_retry=bool(data.get("should_retry", False)),
                confidence=int(data.get("confidence", 0)),
                suggested_parameter=data.get("suggested_parameter"),
                reason=str(data.get("reason", "")),
            )
        except Exception as exc:
            return AutoHealSuggestion(False, 0, reason=f"AI error: {exc}")
