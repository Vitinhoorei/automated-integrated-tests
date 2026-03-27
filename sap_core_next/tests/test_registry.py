from __future__ import annotations

import os
import sys
import unittest

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SRC = os.path.join(ROOT, "src")
sys.path.insert(0, SRC)

from sap_core_next.plugins.pm.plugin import PMPlugin
from sap_core_next.registry.plugin_registry import PluginRegistry


class RegistryTests(unittest.TestCase):
    def test_register_and_resolve(self):
        reg = PluginRegistry()
        reg.register(PMPlugin())
        plugin = reg.resolve("PM")
        self.assertEqual(plugin.name, "pm")


if __name__ == "__main__":
    unittest.main()
