#!/usr/bin/env python3
"""Self-test for py/js_port_parity.py's wiring tiers (added 2026-10-21).

Builds a tiny fake aeBibleAddin tree in a temp folder and checks that each Case lands in the
expected tier, including the two situations that broke the older name-matching heuristic:
a pure core that is only passed as an ARGUMENT to a tested/hooked wrapper, and the UI walk
not leaking into audit-spike.ts through wireUpAuditSpike.

Run (WSL, like every Python tool here):  python3 -I py/test_js_port_parity.py
"""
import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

sys.dont_write_bytecode = True  # do not leave py/__pycache__ behind

HERE = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location("js_port_parity", HERE / "js_port_parity.py")
parity = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(parity)

FILES = {
    "taskpane/src/taskpane.ts": """\
import { uiFn } from './x'
import { wireUp } from './audit-spike'
export function boot() {
  wireUp()
  return uiFn()
}
""",
    "taskpane/src/audit-spike.ts": """\
import { wrapSpike, coreHooked } from './x'
export function wireUp() {
  return { hook: () => wrapSpike(1), other: () => coreHooked() }
}
""",
    "taskpane/src/x.ts": """\
// Case 901 - reachable from the UI
export function uiFn() {
  return 1
}

// Case 902 - a pure core only PASSED to a wrapper that the console hooks call
export function coreSpike() {
  return 2
}
export function wrapSpike(c: number) {
  return run(c, coreSpike)
}

// Case 903 - tested, but no UI path and no hook
export function testedOnly() {
  return 3
}

// Case 904 - nothing reaches it
export function nowhere() {
  return 4
}

// Case 906 - called directly by a console hook
export function coreHooked() {
  return 6
}

// Case 905 is only mentioned in prose, with no function below it.
""",
    "taskpane/src/x.test.ts": """\
import { testedOnly, wrapSpike } from './x'
testedOnly()
wrapSpike(1)
""",
}


class WiringTiers(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._tmp = tempfile.TemporaryDirectory()
        root = Path(cls._tmp.name)
        for rel, body in FILES.items():
            path = root / rel
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(body, encoding="utf-8")
        cls.status = parity.load_js_status(root)

    @classmethod
    def tearDownClass(cls):
        cls._tmp.cleanup()

    def tier(self, case):
        return self.status[case]["status"]

    def test_ui_reachable_is_live(self):
        self.assertEqual(self.tier(901), "LIVE-WIRED")

    def test_core_passed_to_a_hooked_wrapper_is_spike_wired(self):
        # the core is never called by name outside the wrapper; the wrapper is called by the hook
        self.assertEqual(self.tier(902), "SPIKE-WIRED")

    def test_ui_walk_does_not_leak_into_the_console_hooks(self):
        # taskpane.ts calls wireUp (in audit-spike.ts); that must not make hooked functions LIVE
        self.assertEqual(self.tier(906), "SPIKE-WIRED")

    def test_tested_but_unwired_is_an_orphan(self):
        self.assertEqual(self.tier(903), "TESTED-ONLY")

    def test_unreferenced_function_is_ported_not_wired(self):
        self.assertEqual(self.tier(904), "PORTED-NOT-WIRED")

    def test_prose_only_mention_stays_mentioned(self):
        self.assertEqual(self.tier(905), "MENTIONED")

    def test_strip_code_ignores_comments_imports_and_urls(self):
        code = parity.strip_code(
            "import { foo } from './a'\n// call foo( here\nconst u = 'https://x.y/z'\n/* foo */\nexport function bar() { return 1 }\n"
        )
        self.assertNotIn("foo", code)
        self.assertIn("https://x.y/z", code)
        self.assertIn("bar", code)


if __name__ == "__main__":
    unittest.main(verbosity=2)
