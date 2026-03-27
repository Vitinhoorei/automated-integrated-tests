from __future__ import annotations

from dataclasses import dataclass


@dataclass
class RetryPolicy:
    max_attempts: int = 3
    min_confidence: int = 70
    retry_on_sources: tuple[str, ...] = ("STATUSBAR", "EXCEPTION", "BUSINESS")

    def can_retry(self, attempt: int, source: str, confidence: int, has_suggestion: bool) -> bool:
        if attempt >= self.max_attempts:
            return False
        if source not in self.retry_on_sources:
            return False
        if not has_suggestion:
            return False
        return confidence >= self.min_confidence


@dataclass
class AIPolicy:
    enabled: bool = True
    allow_in_modes: tuple[str, ...] = ("real", "simulado", "executar")

    def should_use_ai(self, mode: str) -> bool:
        return self.enabled and (mode or "").lower() in {m.lower() for m in self.allow_in_modes}
