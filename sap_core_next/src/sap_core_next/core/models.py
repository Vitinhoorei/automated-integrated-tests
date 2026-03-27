from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, List, Optional


class StepStatus(str, Enum):
    PASS = "PASS"
    FAIL = "FAIL"
    SKIP = "SKIP"


@dataclass
class ScenarioStep:
    step_id: str
    module: str
    command: str
    explanation: str = ""
    parameters: Dict[str, str] = field(default_factory=dict)
    mode: str = "real"
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass
class GatewayResult:
    status: StepStatus
    source: str
    message: str
    evidence_path: str = ""
    raw: Dict[str, Any] = field(default_factory=dict)


@dataclass
class AutoHealSuggestion:
    should_retry: bool
    confidence: int = 0
    suggested_parameter: Optional[str] = None
    reason: str = ""


@dataclass
class StepAttempt:
    attempt: int
    parameters: Dict[str, str]
    status: StepStatus
    source: str
    message: str
    ai_suggestion: Optional[AutoHealSuggestion] = None


@dataclass
class ExecutionRecord:
    step: ScenarioStep
    status: StepStatus
    source: str
    message: str
    evidence_path: str = ""
    attempts: List[StepAttempt] = field(default_factory=list)


@dataclass
class ExecutionReport:
    execution_id: str
    records: List[ExecutionRecord]

    @property
    def total(self) -> int:
        return len(self.records)

    @property
    def passed(self) -> int:
        return sum(1 for r in self.records if r.status == StepStatus.PASS)

    @property
    def failed(self) -> int:
        return sum(1 for r in self.records if r.status == StepStatus.FAIL)

    @property
    def skipped(self) -> int:
        return sum(1 for r in self.records if r.status == StepStatus.SKIP)
