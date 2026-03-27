from __future__ import annotations

import os
import sys
import unittest

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SRC = os.path.join(ROOT, "src")
sys.path.insert(0, SRC)

from sap_core_next.core.policies import RetryPolicy


class PolicyTests(unittest.TestCase):
    def test_retry_decision(self):
        p = RetryPolicy(max_attempts=3, min_confidence=70)
        self.assertTrue(p.can_retry(1, "STATUSBAR", 90, True))
        self.assertFalse(p.can_retry(3, "STATUSBAR", 90, True))
        self.assertFalse(p.can_retry(1, "OK", 90, True))


if __name__ == "__main__":
    unittest.main()
