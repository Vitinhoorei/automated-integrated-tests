from __future__ import annotations

from dataclasses import dataclass
from typing import List
import uuid

from sap_core_next.core.context import ExecutionContext
from sap_core_next.core.logging import get_structured_logger, log_event
from sap_core_next.core.models import (
    AutoHealSuggestion,
    ExecutionRecord,
    ExecutionReport,
    GatewayResult,
    ScenarioStep,
    StepAttempt,
    StepStatus,
)
from sap_core_next.core.policies import AIPolicy, RetryPolicy
from sap_core_next.ports.ai_assistant import AIAssistant
from sap_core_next.ports.evidence_store import EvidenceStore
from sap_core_next.ports.result_sink import ResultSink
from sap_core_next.ports.sap_gateway import SapGateway
from sap_core_next.ports.scenario_source import ScenarioSource
from sap_core_next.registry.plugin_registry import PluginRegistry


@dataclass
class ExecutionEngine:
    scenario_source: ScenarioSource
    result_sink: ResultSink
    sap_gateway: SapGateway
    ai_assistant: AIAssistant
    evidence_store: EvidenceStore
    plugin_registry: PluginRegistry
    retry_policy: RetryPolicy
    ai_policy: AIPolicy

    def run(self, execution_id: str | None = None) -> ExecutionReport:
        execution_id = execution_id or uuid.uuid4().hex[:10]
        ctx = ExecutionContext(execution_id=execution_id)
        logger = get_structured_logger()

        records: List[ExecutionRecord] = []
        steps = self.scenario_source.read_steps()
        log_event(logger, "execution_started", execution_id=execution_id, total_steps=len(steps))

        for step in steps:
            records.append(self._run_step(step, ctx, logger))

        report = ExecutionReport(execution_id=execution_id, records=records)
        self.result_sink.write_report(report)
        log_event(
            logger,
            "execution_finished",
            execution_id=execution_id,
            total=report.total,
            passed=report.passed,
            failed=report.failed,
            skipped=report.skipped,
        )
        return report

    def _run_step(self, step: ScenarioStep, ctx: ExecutionContext, logger) -> ExecutionRecord:
        plugin = self.plugin_registry.resolve(step.module)
        step = plugin.normalize_step(step, ctx)
        params = plugin.prepare_parameters(step, ctx)
        attempts: List[StepAttempt] = []

        log_event(
            logger,
            "step_started",
            execution_id=ctx.execution_id,
            step_id=step.step_id,
            module=step.module,
            command=step.command,
            mode=step.mode,
        )

        for attempt in range(1, self.retry_policy.max_attempts + 1):
            evidence_path = self.evidence_store.build_path(execution_id=ctx.execution_id, step=step, attempt=attempt)
            result = self.sap_gateway.execute(
                step.command,
                params,
                explanation=step.explanation,
                mode=step.mode,
                evidence_path=evidence_path,
            )

            if result.status == StepStatus.PASS:
                plugin.on_success(step, result.message, params, ctx)
                attempts.append(
                    StepAttempt(attempt=attempt, parameters=dict(params), status=result.status, source=result.source, message=result.message)
                )
                log_event(logger, "step_passed", execution_id=ctx.execution_id, step_id=step.step_id, attempt=attempt, source=result.source)
                return ExecutionRecord(
                    step=step,
                    status=StepStatus.PASS,
                    source=result.source,
                    message=result.message,
                    evidence_path=result.evidence_path,
                    attempts=attempts,
                )

            suggestion = self._suggest(step, params, result)
            attempts.append(
                StepAttempt(
                    attempt=attempt,
                    parameters=dict(params),
                    status=result.status,
                    source=result.source,
                    message=result.message,
                    ai_suggestion=suggestion,
                )
            )

            if suggestion and self.retry_policy.can_retry(
                attempt=attempt,
                source=result.source,
                confidence=suggestion.confidence,
                has_suggestion=bool(suggestion.suggested_parameter and suggestion.should_retry),
            ):
                log_event(
                    logger,
                    "step_retrying",
                    execution_id=ctx.execution_id,
                    step_id=step.step_id,
                    attempt=attempt,
                    confidence=suggestion.confidence,
                    suggestion=suggestion.suggested_parameter,
                )
                params = plugin.apply_suggestion(params, suggestion)
                continue

            log_event(logger, "step_failed", execution_id=ctx.execution_id, step_id=step.step_id, attempt=attempt, source=result.source)
            return ExecutionRecord(
                step=step,
                status=StepStatus.FAIL,
                source=result.source,
                message=result.message,
                evidence_path=result.evidence_path,
                attempts=attempts,
            )

        # fallback teórico
        final = attempts[-1]
        return ExecutionRecord(
            step=step,
            status=StepStatus.FAIL,
            source=final.source,
            message=final.message,
            attempts=attempts,
        )

    def _suggest(self, step: ScenarioStep, params: dict[str, str], result: GatewayResult) -> AutoHealSuggestion | None:
        if not self.ai_policy.should_use_ai(step.mode):
            return None
        return self.ai_assistant.suggest_fix(
            module=step.module,
            command=step.command,
            status_message=result.message,
            parameters=params,
            dump_path=result.raw.get("dump_path") if result.raw else None,
        )
