from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import os
import yaml


@dataclass
class EngineSettings:
    evidence_dir: str = "data/evidence"
    ai_enabled: bool = True
    retry_max_attempts: int = 3
    retry_min_confidence: int = 70
    use_legacy_bridge: bool = True


def load_settings(path: str | None = None) -> EngineSettings:
    path = path or os.getenv("SAP_CORE_NEXT_CONFIG", "sap_core_next/config/default.yaml")
    p = Path(path)
    if not p.exists():
        return EngineSettings()

    with p.open("r", encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}

    return EngineSettings(
        evidence_dir=raw.get("evidence", {}).get("base_dir", "data/evidence"),
        ai_enabled=bool(raw.get("ai", {}).get("enabled", True)),
        retry_max_attempts=int(raw.get("retry", {}).get("max_attempts", 3)),
        retry_min_confidence=int(raw.get("retry", {}).get("min_confidence", 70)),
        use_legacy_bridge=bool(raw.get("sap", {}).get("use_legacy_bridge", True)),
    )
