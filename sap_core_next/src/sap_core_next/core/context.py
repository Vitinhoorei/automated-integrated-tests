from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict


@dataclass
class ExecutionContext:
    execution_id: str
    shared: Dict[str, Any] = field(default_factory=dict)

    def get(self, key: str, default: Any = None) -> Any:
        return self.shared.get(key, default)

    def set(self, key: str, value: Any) -> None:
        self.shared[key] = value
