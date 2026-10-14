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

    def test_layout_lint_flags_defects_and_respects_opt_outs(self):
        css = (SKILL / "assets/player.css").read_text(encoding="utf-8")
        js = (SKILL / "assets/player.js").read_text(encoding="utf-8")
        abs_ = 'style="position:absolute;margin:0;'
        bad = (
            f'<p {abs_}left:80px;top:60px;font-size:28px">重叠的第一段文字</p>'
            f'<p {abs_}left:120px;top:66px;font-size:28px">重叠的第二段文字</p>'
            f'<p {abs_}left:1180px;top:200px;font-size:28px;white-space:nowrap">跑出页面右边的文字</p>'
            f'<p data-bleed {abs_}left:1150px;top:300px;font-size:120px;white-space:nowrap">出血</p>'
            f'<div {abs_}left:80px;top:400px;width:300px;height:30px;overflow:hidden;font-size:24px">被容器裁掉的两行文字被容器裁掉的两行文字</div>'
            f'<p {abs_}left:80px;top:500px;width:170px;font-size:40px">一二三四五</p>'
            f'<p {abs_}left:500px;top:500px;font-size:10px">字号太小</p>'
            f'<p {abs_}left:700px;top:500px;font-size:20px;color:#cccccc">对比度不足</p>'
            f'<p {abs_}left:500px;top:600px;font-size:20px">TODO 补充数据</p>'
            f'<p {abs_}left:40px;top:120px;font-size:200px">“</p>'
        )
        clean = (f'<h1 {abs_}left:80px;top:80px;font-size:48px">干净的页面</h1>'
                 f'<p {abs_}left:80px;top:180px;width:600px;font-size:22px">正文不重叠、不出界，对比度充足。</p>'
                 f'<span {abs_}left:80px;top:260px;font-size:16px;color:#fff;background:#c0392b;padding:4px">徽标</span>')
        with tempfile.TemporaryDirectory() as tmp:
            work = Path(tmp)
            deck = work / "lint.html"
            deck.write_text('<!doctype html><meta charset="utf-8"><style>' + css + '</style><div class="ssr-player">'
                            '<main class="ssr-stage"><div id="ssr-deck" data-width="1280" data-height="720">'
                            f'<section class="ssr-slide">{bad}</section><section class="ssr-slide">{clean}</section>'
                            '</div></main></div><script>' + js + '</script>', encoding="utf-8")
            result = subprocess.run([sys.executable, str(SKILL / "scripts/preview.py"), "html", str(deck), "--out",
                                     str(work / "review"), "--browser", BROWSER, "--no-pdf"], capture_output=True, timeout=90)
            self.assertEqual(result.returncode, 0, result.stderr)
            report = json.loads((work / "review/report.json").read_text(encoding="utf-8"))
            first = {f["kind"] for f in report["layout"][0]["findings"]}
            for kind in ("text_overlap", "text_outside_slide", "text_clipped", "orphan_last_line",
                         "font_too_small", "low_contrast", "placeholder"):
                self.assertIn(kind, first)
            texts = [f["text"] for f in report["layout"][0]["findings"]]
            self.assertNotIn("出血", texts)            # data-bleed opt-out
            self.assertNotIn("“", texts)              # decorative punctuation is not an overlap
            self.assertEqual(report["layout"][1]["findings"], [])
            checks = {c["name"]: c["status"] for c in report["checks"]}
            self.assertEqual(checks["layout"], "failed")
            self.assertTrue((work / "review/layout/slide-01.png").exists())

    def test_visual_first_hints_on_long_text_and_boxed_points(self):
        css = (SKILL / "assets/player.css").read_text(encoding="utf-8")
        js = (SKILL / "assets/player.js").read_text(encoding="utf-8")
        long = "社区活动报名分散在多个街道，现场核对名单耗时较长，部分居民住得较远，建议活动中心统一安排志愿者就近登记，并按服务时长折算积分，同时规定回访时间。"
        point = "<li>志愿者就近登记，按时长折算积分，并记录在案</li>"
        heavy = (f'<h1 style="font-size:40px">{long}</h1><p style="font-size:20px">{long}</p>'
                 f'<ul style="background:#eef;padding:12px;width:600px;font-size:18px">{point * 4}</ul>'
                 f'<p style="position:absolute;top:660px;font-size:14px">来源：{long}</p>')
        light = ('<svg width="800" height="200"><g><rect width="200" height="60"/><text x="10" y="40" font-size="20">发起站点</text></g>'
                 '<text x="300" y="40" font-size="20">协助站点 · 7∶3</text></svg>')
        with tempfile.TemporaryDirectory() as tmp:
            work = Path(tmp)
            deck = work / "text.html"
            deck.write_text('<!doctype html><meta charset="utf-8"><style>' + css + '</style><div class="ssr-player">'
                            '<main class="ssr-stage"><div id="ssr-deck" data-width="1280" data-height="720">'
                            f'<section class="ssr-slide">{heavy}</section><section class="ssr-slide">{light}</section>'
                            '</div></main></div><script>' + js + '</script>', encoding="utf-8")
            base = [sys.executable, str(SKILL / "scripts/preview.py"), "html", str(deck), "--browser", BROWSER, "--no-pdf"]
            result = subprocess.run(base + ["--out", str(work / "hints"), "--visual-first"], capture_output=True, timeout=90)
            self.assertEqual(result.returncode, 0, result.stderr)
            text = json.loads((work / "hints/report.json").read_text(encoding="utf-8"))["text"]
            self.assertEqual(sorted(f["kind"] for f in text["slides"][0]["findings"]), ["long_paragraph", "long_text_block"])
            self.assertEqual(text["slides"][1]["findings"], [])
            self.assertGreater(text["slides"][1]["units"], 0)       # SVG labels count as body text
            plain = subprocess.run(base + ["--out", str(work / "plain")], capture_output=True, timeout=90)
            self.assertEqual(plain.returncode, 0, plain.stderr)
            text = json.loads((work / "plain/report.json").read_text(encoding="utf-8"))["text"]
            self.assertEqual([s["findings"] for s in text["slides"]], [[], []])

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
