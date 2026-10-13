"""Opt-in real-browser regression tests: SSR_TEST_BROWSER=chrome (or chromium/msedge)."""
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "ssr-ppt-skill"
BROWSER = os.environ.get("SSR_TEST_BROWSER")


@unittest.skipUnless(BROWSER, "Set SSR_TEST_BROWSER to run actual browser checks")
class HtmlTests(unittest.TestCase):
    def test_broken_inline_image_fails_strict_and_preserves_report(self):
        with tempfile.TemporaryDirectory() as tmp:
            work = Path(tmp)
            css = (SKILL / "assets/player.css").read_text(encoding="utf-8")
            js = (SKILL / "assets/player.js").read_text(encoding="utf-8")
            deck = work / "bad.html"
            deck.write_text('<!doctype html><meta charset="utf-8"><style>'+css+'</style><div class="ssr-player"><main class="ssr-stage"><div id="ssr-deck" data-width="1280" data-height="720"><section class="ssr-slide"><h1>Inline image check</h1><img src="data:image/png;base64,bm90LWFuLWltYWdl" alt="broken fixture"></section></div></main></div><script>'+js+'</script>', encoding="utf-8")
            result = subprocess.run([sys.executable, str(SKILL / "scripts/preview.py"), "html", str(deck), "--out", str(work / "review"), "--browser", BROWSER, "--strict"], capture_output=True, timeout=60)
            self.assertEqual(result.returncode, 1, result.stderr)
            report = json.loads((work / "review/report.json").read_text(encoding="utf-8"))
            checks = {c["name"]: c["status"] for c in report["checks"]}
            self.assertEqual(checks["image_decode"], "failed")
            self.assertEqual(checks["keyboard"], "passed")
            self.assertEqual(report["automatic_status"], "failed")

    def test_reduced_motion_keeps_manual_steps_but_print_reveals(self):
        from playwright.sync_api import sync_playwright
        css = (SKILL / "assets/player.css").read_text(encoding="utf-8")
        js = (SKILL / "assets/player.js").read_text(encoding="utf-8")
        with sync_playwright() as p:
            browser = p.chromium.launch(**({} if BROWSER == "chromium" else {"channel": BROWSER}))
            try:
                page = browser.new_page(reduced_motion="reduce")
                page.set_content('<style>'+css+'</style><div class="ssr-player"><main class="ssr-stage"><div id="ssr-deck"><section class="ssr-slide"><h1>Visible title</h1><p id="step" data-step="1">Manual reveal</p></section></div></main></div><script>'+js+'</script>')
                opacity = lambda: page.locator("#step").evaluate("el => getComputedStyle(el).opacity")
                self.assertEqual(opacity(), "0")
                page.keyboard.press("ArrowRight")
                self.assertEqual(opacity(), "1")
                page.keyboard.press("Home")
                page.emulate_media(media="print")
                self.assertEqual(opacity(), "1")
            finally:
                browser.close()


if __name__ == "__main__":
    unittest.main()
