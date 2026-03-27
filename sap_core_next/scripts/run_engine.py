from __future__ import annotations

import argparse
import os
import sys

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SRC = os.path.join(ROOT, "src")
LEGACY_ROOT = os.path.abspath(os.path.join(ROOT, ".."))

sys.path.insert(0, SRC)
sys.path.insert(0, LEGACY_ROOT)

from sap_core_next.adapters.ai.noop_ai import NoOpAIAssistant
from sap_core_next.adapters.sap.legacy_bridge_gateway import LegacyBridgeGateway
from sap_core_next.adapters.sap.win32com_gateway import Win32ComSapGateway
from sap_core_next.adapters.spreadsheet.openpyxl_sink import OpenpyxlResultSink
from sap_core_next.adapters.spreadsheet.openpyxl_source import OpenpyxlScenarioSource
from sap_core_next.config.settings import load_settings
from sap_core_next.core.engine import ExecutionEngine
from sap_core_next.core.policies import AIPolicy, RetryPolicy
from sap_core_next.plugins.pm.plugin import PMPlugin
from sap_core_next.registry.plugin_registry import PluginRegistry
from sap_core_next.services.evidence_service import FileEvidenceStore


def main() -> None:
    parser = argparse.ArgumentParser(description="Run sap_core_next engine")
    parser.add_argument("--file", required=True)
    parser.add_argument("--sheet", required=True)
    parser.add_argument("--config", default=os.path.join(ROOT, "config", "default.yaml"))
    args = parser.parse_args()

    settings = load_settings(args.config)
    source = OpenpyxlScenarioSource(args.file, args.sheet)
    sink = OpenpyxlResultSink(args.file, args.sheet)
    evidence = FileEvidenceStore(base_dir=settings.evidence_dir)

    sap = LegacyBridgeGateway() if settings.use_legacy_bridge else Win32ComSapGateway()

    registry = PluginRegistry()
    registry.register(PMPlugin())

    engine = ExecutionEngine(
        scenario_source=source,
        result_sink=sink,
        sap_gateway=sap,
        ai_assistant=NoOpAIAssistant(),
        evidence_store=evidence,
        plugin_registry=registry,
        retry_policy=RetryPolicy(max_attempts=settings.retry_max_attempts, min_confidence=settings.retry_min_confidence),
        ai_policy=AIPolicy(enabled=settings.ai_enabled),
    )
    report = engine.run()
    print(f"Done: total={report.total} pass={report.passed} fail={report.failed} skip={report.skipped}")


if __name__ == "__main__":
    main()
