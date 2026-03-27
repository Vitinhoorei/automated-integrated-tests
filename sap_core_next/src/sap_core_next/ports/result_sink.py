from __future__ import annotations

from typing import Protocol

from sap_core_next.core.models import ExecutionReport


class ResultSink(Protocol):
    def write_report(self, report: ExecutionReport) -> None:
        ...
