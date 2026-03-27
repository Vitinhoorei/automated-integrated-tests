from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import re

from sap_core_next.core.models import ScenarioStep


def _safe(text: str) -> str:
    return re.sub(r"[^\w\-]+", "_", str(text or "UNK")).strip("_")[:60] or "UNK"


@dataclass
class FileEvidenceStore:
    base_dir: str = "data/evidence"

    def build_path(self, *, execution_id: str, step: ScenarioStep, attempt: int) -> str:
        Path(self.base_dir).mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = "_".join([
            _safe(execution_id),
            _safe(step.module),
            _safe(step.step_id),
            _safe(step.command),
            f"a{attempt}",
            ts,
        ]) + ".png"
        return str(Path(self.base_dir) / filename)
