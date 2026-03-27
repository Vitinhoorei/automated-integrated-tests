from __future__ import annotations

import os
import sys
import unittest

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SRC = os.path.join(ROOT, "src")
sys.path.insert(0, SRC)

from sap_core_next.core.engine import ExecutionEngine
from sap_core_next.core.models import AutoHealSuggestion, GatewayResult, ScenarioStep, StepStatus
from sap_core_next.core.policies import AIPolicy, RetryPolicy
from sap_core_next.plugins.pm.plugin import PMPlugin
from sap_core_next.registry.plugin_registry import PluginRegistry
from sap_core_next.services.evidence_service import FileEvidenceStore


class FakeSource:
    def read_steps(self):
        return [
            ScenarioStep(step_id="1", module="PM", command="IW31", parameters={"A": "1"}, mode="real"),
        ]


class FakeSink:
    def __init__(self):
        self.report = None

    def write_report(self, report):
        self.report = report


class FlakyGateway:
    def __init__(self):
        self.calls = 0

    def execute(self, command, parameters, **kwargs):
        self.calls += 1
        if self.calls == 1:
            return GatewayResult(StepStatus.FAIL, "STATUSBAR", "campo obrigatório")
        return GatewayResult(StepStatus.PASS, "OK", "Ordem 123 criada")


class FakeAI:
    def suggest_fix(self, **kwargs):
        return AutoHealSuggestion(True, 95, "CampoX=ABC", "known fix")


class EngineTests(unittest.TestCase):
    def test_retry_and_pass(self):
        source = FakeSource()
        sink = FakeSink()
        gw = FlakyGateway()
        ai = FakeAI()
        evidence = FileEvidenceStore(base_dir="/tmp/sap_core_next_tests")

        registry = PluginRegistry()
        registry.register(PMPlugin())

        engine = ExecutionEngine(
            scenario_source=source,
            result_sink=sink,
            sap_gateway=gw,
            ai_assistant=ai,
            evidence_store=evidence,
            plugin_registry=registry,
            retry_policy=RetryPolicy(max_attempts=3, min_confidence=70),
            ai_policy=AIPolicy(enabled=True),
        )

        report = engine.run(execution_id="exec123")
        self.assertEqual(report.passed, 1)
        self.assertEqual(gw.calls, 2)
        self.assertIsNotNone(sink.report)


if __name__ == "__main__":
    unittest.main()
