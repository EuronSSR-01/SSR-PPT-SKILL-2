#!/usr/bin/env python3
"""Render decks to images for visual review.

  preview.py html  deck.html  --out DIR   per-slide PNGs (fully revealed), console/network/offline checks,
                                          keyboard test, small-viewport shot, print PDF, contact sheets
  preview.py pptx  deck.pptx  --out DIR   LibreOffice -> PDF -> PNGs, native object inventory, fonts,
                                          optional --edit-test on a copy, contact sheets
  preview.py sheet PNGS|DIR   --out FILE  contact sheet plus -gray and -blur variants
  preview.py compare A B ...  --out FILE  images side by side for comparing directions

Dependencies: Pillow; html needs playwright (bundled Chromium or installed Chrome/Edge);
pptx needs LibreOffice and PyMuPDF, and python-pptx for the inventory and --edit-test.
Every image this writes must still be opened and looked at; numbers in report.json are only hints.
"""
import argparse
import importlib
from io import BytesIO
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def need(module, hint):
    try:
        return importlib.import_module(module)
    except ImportError:
        sys.exit(f"missing dependency {module!r}: {hint}")


def pymupdf_or_none():
    for name in ("pymupdf", "fitz"):  # newer releases deprecate the fitz name
        try:
            return importlib.import_module(name)
        except ImportError:
            pass
    return None


# ---------------------------------------------------------------- contact sheets

def _font(size):
    from PIL import ImageFont
    try:
        return ImageFont.load_default(size=size)
    except TypeError:
        return ImageFont.load_default()


def make_sheet(images, out, cols=None, thumb_w=400, labels=None):
    """Grid of thumbnails, plus -gray and -blur variants for hierarchy checks."""
    Image = need("PIL.Image", "python -m pip install pillow")
    from PIL import ImageDraw, ImageFilter
    images = [Path(p) for p in images]
    if not images:
        raise SystemExit("no images for contact sheet")
    cols = cols or min(4, max(1, math.ceil(math.sqrt(len(images) * 1.3))))
    thumbs = []
    for p in images:
        im = Image.open(p).convert("RGB")
        im.thumbnail((thumb_w, thumb_w * 4))
        thumbs.append(im)
    cell_h = max(t.height for t in thumbs)
    pad, label_h = 16, 22
    rows = math.ceil(len(thumbs) / cols)
    sheet = Image.new("RGB", (pad + cols * (thumb_w + pad), pad + rows * (cell_h + label_h + pad)), (38, 38, 40))
    draw = ImageDraw.Draw(sheet)
    font = _font(15)
    for i, t in enumerate(thumbs):
        x = pad + (i % cols) * (thumb_w + pad)
        y = pad + (i // cols) * (cell_h + label_h + pad)
        draw.text((x, y), labels[i] if labels else f"{i + 1:02d}", fill=(220, 220, 220), font=font)
        sheet.paste(t, (x, y + label_h))
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(out)
    sheet.convert("L").save(out.with_name(out.stem + "-gray" + out.suffix))
    sheet.filter(ImageFilter.GaussianBlur(max(2, thumb_w // 45))).save(out.with_name(out.stem + "-blur" + out.suffix))
    return out


def make_compare(images, out, labels=None, height=540):
    Image = need("PIL.Image", "python -m pip install pillow")
    from PIL import ImageDraw
    ims = []
    for p in images:
        im = Image.open(p).convert("RGB")
        ims.append(im.resize((round(im.width * height / im.height), height)))
    pad, label_h = 20, 28
    canvas = Image.new("RGB", (pad + sum(i.width + pad for i in ims), height + label_h + 2 * pad), (38, 38, 40))
    draw = ImageDraw.Draw(canvas)
    font = _font(20)
    x = pad
    for n, im in enumerate(ims):
        draw.text((x, pad), labels[n] if labels else chr(65 + n), fill=(230, 230, 230), font=font)
        canvas.paste(im, (x, pad + label_h))
        x += im.width + pad
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(out)
    return out


# ---------------------------------------------------------------- HTML

CAPTURE_CSS = """
.ssr-ui { display: none !important; }
*, *::before, *::after { animation-duration: 0s !important; animation-delay: 0s !important;
  transition-duration: 0s !important; transition-delay: 0s !important; caret-color: transparent !important; }
"""

WAIT_ASSETS_JS = """async () => {
  await document.fonts.ready;
  const svgImages = Array.from(document.querySelectorAll('svg image')).map(el => {
    const img = new Image(); img.src = el.href.baseVal; return img;
  });
  const images = await Promise.all([...document.images, ...svgImages].map(async (img, index) => {
    let error = null;
    try { await img.decode(); } catch (e) { error = String(e); }
    return {index, source: img.src.startsWith('data:') ? 'inline image' : img.src,
            ok: !error && img.naturalWidth > 0 && img.naturalHeight > 0, error};
  }));
  return {fonts_status: document.fonts.status, images,
          fonts: Array.from(document.fonts).map(f => ({family: f.family, status: f.status}))};
}"""


# Geometry checks on the fully revealed slide. They catch common defects without looking at the image
# and tell agents with vision where to look. Opt-outs: data-bleed (intentional bleed off the slide),
# data-lint-ignore (skip a subtree). Coordinates are canvas px relative to the slide.
LAYOUT_JS = r"""(index) => {
  const slide = document.querySelectorAll('#ssr-deck > .ssr-slide')[index];
  const sr = slide.getBoundingClientRect();
  const k = sr.width / (slide.offsetWidth || sr.width) || 1;
  const rel = r => ({x: Math.round((r.left - sr.left) / k), y: Math.round((r.top - sr.top) / k),
                     w: Math.round(r.width / k), h: Math.round(r.height / k)});
  const out = [];
  const label = el => (el.textContent || '').trim().replace(/\s+/g, ' ').slice(0, 40);
  const add = (kind, severity, el, rect, detail) => out.push({kind, severity, text: label(el), rect: rel(rect), detail});
  const skip = el => el.closest('[data-lint-ignore], .ssr-notes, .ssr-ui, script, style');
  const shown = el => { for (let e = el; e && e !== slide; e = e.parentElement) { const cs = getComputedStyle(e);
      if (cs.display === 'none' || cs.visibility === 'hidden' || parseFloat(cs.opacity) < 0.05) return false; } return true; };
  const ownRects = el => { const rects = [];
    for (const n of el.childNodes) if (n.nodeType === 3 && n.textContent.trim()) {
      const range = document.createRange(); range.selectNodeContents(n);
      for (const r of range.getClientRects()) if (r.width > 0.5 && r.height > 0.5) rects.push(r); }
    return rects; };
  const fontPx = el => { const fs = parseFloat(getComputedStyle(el).fontSize) || 16;
    if (el instanceof SVGElement && el.getScreenCTM) { const m = el.getScreenCTM(); if (m) return fs * Math.hypot(m.a, m.b) / k; }
    return fs; };
  const texts = Array.from(slide.querySelectorAll('*')).filter(el => !skip(el) && shown(el) && ownRects(el).length)
    .map(el => ({el, rects: ownRects(el), fs: fontPx(el)}));

  for (const t of texts) {
    // 1. text outside the slide (the slide clips it)
    if (!t.el.closest('[data-bleed]')) for (const r of t.rects)
      if (r.left < sr.left - 1 || r.top < sr.top - 1 || r.right > sr.right + 1 || r.bottom > sr.bottom + 1) {
        add('text_outside_slide', 'error', t.el, r, 'Text runs past the slide edge; add data-bleed if intended.'); break; }
    // 2. text clipped by an overflow-hidden ancestor
    for (let a = t.el; a && a !== slide; a = a.parentElement) {  // the element itself may clip its own text
      const cs = getComputedStyle(a);
      if (cs.overflowX === 'visible' && cs.overflowY === 'visible') continue;
      const ar = a.getBoundingClientRect();
      const hit = t.rects.find(r => r.right > ar.right + 1 || r.bottom > ar.bottom + 1 || r.left < ar.left - 1 || r.top < ar.top - 1);
      if (hit) { add('text_clipped', 'error', t.el, hit, 'Text exceeds an overflow-hidden container.'); break; }
    }
    // 5. font size
    if (t.fs < 12) add('font_too_small', 'warning', t.el, t.rects[0], `Effective size ${t.fs.toFixed(1)}px (< 12px).`);
  }

  // 3. text overlapping other text (em-box height, so tight line-height does not count)
  const box = (r, fs) => { const inset = Math.max(0, (r.height - fs * k) / 2);
    return {l: r.left, r: r.right, t: r.top + inset, b: r.bottom - inset}; };
  const symbolsOnly = el => /^[\p{P}\p{S}\s]+$/u.test(el.textContent || '');  // big quote marks: tiny ink, huge em box
  for (let i = 0; i < texts.length; i++) for (let j = i + 1; j < texts.length; j++) {
    const a = texts[i], b = texts[j];
    if (a.el.contains(b.el) || b.el.contains(a.el) || symbolsOnly(a.el) || symbolsOnly(b.el)) continue;
    let found = null;
    for (const ra of a.rects) { const A = box(ra, a.fs);
      for (const rb of b.rects) { const B = box(rb, b.fs);
        const ix = Math.min(A.r, B.r) - Math.max(A.l, B.l), iy = Math.min(A.b, B.b) - Math.max(A.t, B.t);
        if (ix > 4 && iy > 4) { found = ra; break; } }
      if (found) break; }
    if (found) add('text_overlap', 'error', a.el, found, `Overlaps “${label(b.el)}”.`);
  }

  // 4. orphan last line: a multi-line block whose last line holds about one character
  const blocks = new Set(texts.map(t => t.el instanceof SVGElement ? null :
    t.el.closest('p, li, h1, h2, h3, h4, h5, h6, td, th, dd, blockquote, figcaption, div')).filter(Boolean));
  for (const el of blocks) {
    if (skip(el) || el.querySelector('p, li, h1, h2, h3, h4, h5, h6, td, div, table, svg')) continue;
    const range = document.createRange(); range.selectNodeContents(el);
    const rs = Array.from(range.getClientRects()).filter(r => r.width > 0.5 && r.height > 0.5);
    if (rs.length < 2) continue;
    const fs = fontPx(el) * k, lines = [];
    for (const r of rs) { const c = (r.top + r.bottom) / 2;
      const line = lines.find(l => Math.abs(l.c - c) < fs * 0.5);
      if (line) { line.l = Math.min(line.l, r.left); line.r = Math.max(line.r, r.right); }
      else lines.push({c, l: r.left, r: r.right, t: r.top, b: r.bottom}); }
    if (lines.length < 2) continue;
    lines.sort((a, b) => a.c - b.c);
    const last = lines[lines.length - 1];
    if (last.r - last.l <= fs * 1.6)
      add('orphan_last_line', 'warning', el, new DOMRect(last.l, last.t, last.r - last.l, last.b - last.t),
          'Last line holds about one character; rewrite, widen or break the line deliberately.');
  }

  // 6. contrast against whatever is painted behind the text
  const parse = c => { const m = /rgba?\(([^)]+)\)/.exec(c || ''); if (!m) return null;
    const p = m[1].split(/[ ,/]+/).filter(Boolean).map(Number); return {r: p[0], g: p[1], b: p[2], a: p.length > 3 ? p[3] : 1}; };
  const lum = c => { const f = v => { v /= 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); };
    return 0.2126 * f(c.r) + 0.7152 * f(c.g) + 0.0722 * f(c.b); };
  const behind = (el, r) => {
    if (!(el instanceof SVGElement)) {  // text paints over its own background first (badges, tags)
      const own = getComputedStyle(el);
      if (own.backgroundImage && own.backgroundImage !== 'none') return null;
      const bg = parse(own.backgroundColor); if (bg && bg.a > 0.9) return bg; }
    const stack = document.elementsFromPoint((r.left + r.right) / 2, (r.top + r.bottom) / 2);
    for (const e of stack) {
      if (e === el || el.contains(e) || !slide.contains(e) && e !== slide) continue;
      if (/^(IMG|VIDEO|CANVAS|IMAGE)$/i.test(e.tagName)) return null;
      const cs = getComputedStyle(e);
      if (e instanceof SVGElement) {
        if (e instanceof SVGGeometryElement) { const f = parse(cs.fill);
          if (f && f.a > 0.9 && parseFloat(cs.fillOpacity || 1) > 0.9) return f;
          if (cs.fill && cs.fill !== 'none' && !f) return null; }
        continue; }
      if (cs.backgroundImage && cs.backgroundImage !== 'none') return null;
      const bg = parse(cs.backgroundColor); if (bg && bg.a > 0.9) return bg;
      if (e === slide) break;
    }
    return null; };
  const lowContrast = new Map();  // one finding per colour pair per slide
  for (const t of texts) {
    const cs = getComputedStyle(t.el);
    const fg = parse(t.el instanceof SVGElement ? cs.fill : cs.color);
    const bg = behind(t.el, t.rects[0]);
    if (!fg || !bg) continue;
    const mix = {r: fg.r * fg.a + bg.r * (1 - fg.a), g: fg.g * fg.a + bg.g * (1 - fg.a), b: fg.b * fg.a + bg.b * (1 - fg.a)};
    const [hi, lo] = [lum(mix), lum(bg)].sort((a, b) => b - a);
    const ratio = (hi + 0.05) / (lo + 0.05);
    const large = t.fs >= 24 || (t.fs >= 18.66 && parseInt(cs.fontWeight, 10) >= 700);
    if (ratio >= (large ? 3 : 4.5)) continue;
    const key = `${Math.round(mix.r)},${Math.round(mix.g)},${Math.round(mix.b)}/${bg.r},${bg.g},${bg.b}/${large}`;
    const seen = lowContrast.get(key);
    if (seen) { seen.count++; continue; }
    lowContrast.set(key, {t, ratio, need: large ? 3 : 4.5, count: 1});
  }
  for (const {t, ratio, need, count} of lowContrast.values())
    add('low_contrast', 'warning', t.el, t.rects[0],
        `Contrast ${ratio.toFixed(2)}:1 (needs ${need}:1)` + (count > 1 ? `; ${count} text items use this colour pair.` : '.'));

  // 7. leftover placeholders
  const placeholder = /lorem ipsum|\bTODO\b|\bTBD\b|X{3,}|占位|待补充|待填写|此处填写|示例文字/i;
  for (const t of texts) if (placeholder.test(t.el.textContent))
    add('placeholder', 'warning', t.el, t.rects[0], 'Looks like placeholder text.');
  return out;
}"""


def draw_layout(shot, findings, out_path, scale):
    """Annotate a slide screenshot with numbered boxes: red = error, orange = warning."""
    Image = need("PIL.Image", "python -m pip install pillow")
    from PIL import ImageDraw
    im = Image.open(shot).convert("RGB")
    draw = ImageDraw.Draw(im)
    font = _font(max(14, int(16 * scale)))
    numbered = sorted(enumerate(findings, 1), key=lambda p: p[1]["severity"] == "error")  # errors drawn on top
    for n, f in numbered:
        r = f["rect"]
        color = (220, 30, 30) if f["severity"] == "error" else (240, 140, 0)
        box = [r["x"] * scale, r["y"] * scale, (r["x"] + r["w"]) * scale, (r["y"] + r["h"]) * scale]
        draw.rectangle(box, outline=color, width=max(2, int(3 * scale)))
        draw.text((box[0] + 2, max(0, box[1] - 20 * scale)), str(n), fill=color, font=font)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    im.save(out_path)


def check(name, status, detail="", required=True, **evidence):
    return {"name": name, "status": status, "required": required, "detail": detail, **evidence}


def checks_status(checks):
    required = [c["status"] for c in checks if c["required"]]
    for status in ("failed", "unsupported", "not_run"):
        if status in required:
            return status
    return "passed" if required and all(s == "passed" for s in required) else "not_run"


def write_report(report, out, strict=False):
    report["automatic_status"] = checks_status(report["checks"])
    report["manual_review"] = {
        "status": "not_run", "required_for_delivery": True,
        "detail": "View every final page; inspect content, visual quality and task-specific edits. "
                  "Automatic checks do not prove target-office interaction or design quality."
    }
    (out / "report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if strict and report["automatic_status"] != "passed":
        raise SystemExit(1)


def launch_browser(p, choice):
    order = {"auto": [None, "chrome", "msedge"], "chromium": [None], "chrome": ["chrome"], "msedge": ["msedge"]}[choice]
    errors = []
    for channel in order:
        try:
            return p.chromium.launch(channel=channel) if channel else p.chromium.launch(), channel or "chromium"
        except Exception as e:  # browser not installed for this channel
            errors.append(f"{channel or 'chromium'}: {str(e).splitlines()[0]}")
    sys.exit("no usable browser:\n  " + "\n  ".join(errors) +
             "\ninstall one with: python -m playwright install chromium")


def cmd_html(args):
    sync_api = need("playwright.sync_api", "python -m pip install playwright && python -m playwright install chromium")
    src = Path(args.input).resolve()
    out = Path(args.out)
    slides_dir = out / "slides"
    if slides_dir.exists():
        shutil.rmtree(slides_dir)
    slides_dir.mkdir(parents=True)
    report = {"file": str(src), "console": [], "page_errors": [], "blocked_requests": [], "failed_requests": []}

    with sync_api.sync_playwright() as p:
        browser, used = launch_browser(p, args.browser)
        report["browser"] = used
        ctx = browser.new_context(viewport={"width": 1280, "height": 720}, device_scale_factor=args.scale)

        def route(r):
            url = r.request.url
            if url.startswith(("file:", "data:", "blob:", "about:")):
                r.continue_()
            else:
                report["blocked_requests"].append(url)
                r.abort()
        ctx.route("**/*", route)
        page = ctx.new_page()
        page.on("console", lambda m: m.type in ("error", "warning") and report["console"].append(f"{m.type}: {m.text}"))
        page.on("pageerror", lambda e: report["page_errors"].append(str(e)))
        page.on("requestfailed", lambda r: report["failed_requests"].append(f"{r.url} {r.failure}"))
        page.goto(src.as_uri(), wait_until="load")
        report["assets"] = page.evaluate(WAIT_ASSETS_JS)
        report["fonts_status"] = report["assets"]["fonts_status"]

        info = page.evaluate("""() => {
          const deck = document.getElementById('ssr-deck');
          return { width: Number(deck?.dataset.width || 1280), height: Number(deck?.dataset.height || 720),
                   slides: document.querySelectorAll('#ssr-deck > .ssr-slide').length,
                   api: typeof window.SSRPlayer === 'object' };
        }""")
        report.update(info)
        W, H, n = info["width"], info["height"], info["slides"]
        if not n:
            sys.exit("no #ssr-deck > .ssr-slide found")

        # Wide and small viewports with the player UI visible.
        page.set_viewport_size({"width": 1440, "height": 900})
        page.mouse.move(700, 450)
        page.wait_for_timeout(300)
        page.screenshot(path=str(out / "ui-wide.png"))
        page.set_viewport_size({"width": 800, "height": 600})
        page.mouse.move(400, 300)
        page.wait_for_timeout(300)
        page.screenshot(path=str(out / "ui-small.png"))

        # Keyboard navigation.
        if info["api"]:
            page.set_viewport_size({"width": W, "height": H})
            page.mouse.click(5, 5)
            checks = {}
            page.keyboard.press("End")
            checks["End -> last"] = page.evaluate("SSRPlayer.index") == n - 1
            page.keyboard.press("Home")
            checks["Home -> first"] = page.evaluate("SSRPlayer.index") == 0
            total = page.evaluate("SSRPlayer.totalSteps") + n + 2
            for _ in range(total):
                page.keyboard.press("ArrowRight")
            checks["ArrowRight x all -> last"] = page.evaluate("SSRPlayer.index") == n - 1
            for _ in range(total):
                page.keyboard.press("ArrowLeft")
            checks["ArrowLeft x all -> first"] = page.evaluate("SSRPlayer.index") == 0
            page.evaluate("location.hash = '#2'")
            page.wait_for_timeout(100)
            checks["hash #2 -> slide 2"] = n < 2 or page.evaluate("SSRPlayer.index") == 1
            report["keyboard"] = checks

        # Per-slide captures, fully revealed, without UI or animation.
        page.set_viewport_size({"width": W, "height": H})
        page.add_style_tag(content=CAPTURE_CSS)
        report["layout"] = []
        for i in range(n):
            if info["api"]:
                page.evaluate("i => SSRPlayer.go(i, { reveal: 'all' })", i)
            else:
                page.keyboard.press("Home")
                for _ in range(i):
                    page.keyboard.press("ArrowRight")
            page.wait_for_timeout(60)
            page.locator("#ssr-deck").screenshot(path=str(slides_dir / f"slide-{i + 1:02d}.png"))
            report["layout"].append({"slide": i + 1, "findings": page.evaluate(LAYOUT_JS, i)})

        if not args.no_pdf:
            pdf_path = out / "print.pdf"
            page.pdf(path=str(pdf_path), prefer_css_page_size=True, print_background=True)
            pm = pymupdf_or_none()
            if pm:
                doc = pm.open(pdf_path)
                r = doc[0].rect
                report["print"] = {"pages": doc.page_count, "expected_pages": n,
                                   "page_pt": [round(r.width, 1), round(r.height, 1)],
                                   "expected_pt": [round(W * 0.75, 1), round(H * 0.75, 1)]}
                doc.close()
            else:
                report["print"] = {"pdf": str(pdf_path), "note": "install PyMuPDF to count pages"}
        browser.close()

    shots = sorted(slides_dir.glob("slide-*.png"))
    make_sheet(shots, out / "sheet.png")
    layout_dir = out / "layout"
    if layout_dir.exists():
        shutil.rmtree(layout_dir)
    for item in report["layout"]:
        if item["findings"]:
            draw_layout(slides_dir / f"slide-{item['slide']:02d}.png", item["findings"],
                        layout_dir / f"slide-{item['slide']:02d}.png", args.scale)
    errors = sum(f["severity"] == "error" for item in report["layout"] for f in item["findings"])
    warnings = sum(f["severity"] == "warning" for item in report["layout"] for f in item["findings"])
    problems = report["console"] + report["page_errors"] + report["failed_requests"]
    print(f"browser: {report['browser']}  slides: {n}  canvas: {W}x{H}  fonts: {report['fonts_status']}")
    print(f"slides -> {slides_dir}")
    print(f"sheet  -> {out / 'sheet.png'} (+ -gray, -blur)   ui -> ui-wide.png, ui-small.png")
    if "keyboard" in report:
        print("keyboard: " + ", ".join(f"{k}: {'ok' if v else 'FAIL'}" for k, v in report["keyboard"].items()))
    if "print" in report:
        print(f"print: {report['print']}")
    print(f"blocked external requests: {len(report['blocked_requests'])}  console/page errors: {len(problems)}")
    for line in (report["blocked_requests"] + problems)[:20]:
        print("  " + line)
    print(f"layout: {errors} errors, {warnings} warnings" + (f" -> annotated pages in {layout_dir}" if errors or warnings else ""))
    for item in report["layout"]:
        for f in item["findings"]:
            print(f"  slide {item['slide']:02d} {f['severity']:<7} {f['kind']:<18} “{f['text']}” {f['detail']}")

    report["checks"] = [
        check("offline_resources", "failed" if report["blocked_requests"] or report["failed_requests"] else "passed"),
        check("browser_errors", "failed" if report["console"] or report["page_errors"] else "passed"),
        check("image_decode", "passed" if all(i["ok"] for i in report["assets"]["images"]) else "failed",
              "HTML img and SVG image elements; inspect CSS backgrounds and crop quality visually too."),
        check("font_load", "passed" if report["fonts_status"] == "loaded" and
              all(f["status"] != "error" for f in report["assets"]["fonts"]) else "failed",
              "Loading only; glyph coverage and fallback still require font checks and visual inspection."),
        check("keyboard", "passed" if report.get("keyboard") and all(report["keyboard"].values())
              else "failed" if report.get("keyboard") else "unsupported",
              "SSRPlayer keyboard checks; use task-specific checks for a different player."),
        check("layout", "failed" if errors else "passed",
              f"{errors} errors (outside slide, clipped, overlapping text), {warnings} warnings (orphan line, "
              "small font, contrast, placeholder). Geometry only: composition, meaning and crops still need review.",
              errors=errors, warnings=warnings),
    ]
    pr = report.get("print", {})
    if args.no_pdf:
        report["checks"].append(check("print", "not_run", "Skipped with --no-pdf; strict mode requires this check."))
    elif "pages" not in pr:
        report["checks"].append(check("print", "unsupported", "PyMuPDF needed for page count and dimensions."))
    else:
        ok = pr["pages"] == pr["expected_pages"] and all(
            abs(a-b) <= 1 for a,b in zip(pr["page_pt"], pr["expected_pt"]))
        report["checks"].append(check("print", "passed" if ok else "failed", **pr))
    write_report(report, out, args.strict)


# ---------------------------------------------------------------- PPTX

def find_soffice(explicit):
    """Path to LibreOffice's soffice, or None if it is not available (rendering is then reported as not_run)."""
    if explicit:
        found = explicit if Path(explicit).is_file() else shutil.which(explicit)
        if not found:
            print(f"warning: LibreOffice not found at {explicit}", file=sys.stderr)
        return found
    for name in ("soffice.com", "soffice", "libreoffice"):
        if shutil.which(name):
            return shutil.which(name)
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.com",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.com",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice", "/usr/lib/libreoffice/program/soffice", "/opt/libreoffice/program/soffice",
    ]
    for c in candidates:
        if Path(c).exists():
            return c
    return None


PLACEHOLDER = re.compile(r"lorem ipsum|\bTODO\b|\bTBD\b|X{3,}|占位|待补充|待填写|此处填写|示例文字", re.I)


def pdf_layout(page, pm):
    """Text geometry on a rendered PPTX page, in points: text off the page, text over other text,
    a last line of about one character, tiny text and placeholders. Backgrounds are not known, so no contrast."""
    flags = getattr(pm, "TEXTFLAGS_DICT", 0) & ~getattr(pm, "TEXT_MEDIABOX_CLIP", 64)  # keep text past the edge
    width, height = page.rect.width, page.rect.height
    findings, spans = [], []

    def add(kind, severity, text, bbox, detail):
        x0, y0, x1, y1 = bbox
        findings.append({"kind": kind, "severity": severity, "text": text.strip()[:40],
                         "rect": {"x": round(x0), "y": round(y0), "w": round(x1 - x0), "h": round(y1 - y0)},
                         "detail": detail})

    for bi, block in enumerate(page.get_text("dict", flags=flags)["blocks"]):
        if block.get("type") != 0:
            continue
        lines = []
        for li, line in enumerate(block["lines"]):
            text = "".join(s["text"] for s in line["spans"])
            if text.strip():
                lines.append((text.strip(), line["bbox"]))
            for s in line["spans"]:
                if s["text"].strip():
                    spans.append((bi, li, s))
        block_text = " ".join(t for t, _ in lines)
        bx0, by0, bx1, by1 = block["bbox"]
        off_page = bx0 < -1 or by0 < -1 or bx1 > width + 1 or by1 > height + 1  # reported below as an error
        if len(lines) >= 2 and len(lines[-1][0]) == 1 and not off_page:
            add("orphan_last_line", "warning", block_text, lines[-1][1],
                "Last line holds one character; rewrite, widen the box or break the line deliberately.")
        if PLACEHOLDER.search(block_text):
            add("placeholder", "warning", block_text, block["bbox"], "Looks like placeholder text.")
        tiny = min((s["size"] for line in block["lines"] for s in line["spans"] if s["text"].strip()), default=99)
        if tiny < 9:
            add("font_too_small", "warning", block_text, block["bbox"], f"{tiny:.1f}pt (< 9pt, about 12px).")

    for bi, li, s in spans:
        x0, y0, x1, y1 = s["bbox"]
        if x0 < -1 or y0 < -1 or x1 > width + 1 or y1 > height + 1:
            add("text_outside_slide", "error", s["text"], s["bbox"], "Text runs past the slide edge.")

    def em_box(s):
        x0, y0, x1, y1 = s["bbox"]
        inset = max(0.0, (y1 - y0 - s["size"]) / 2)
        return x0, y0 + inset, x1, y1 - inset

    # Compare every pair of spans: text extraction may merge overlapping boxes into one block or line.
    # Neighbouring spans and consecutive lines only touch; em-box height ignores line-spacing slack.
    symbols = re.compile(r"^[\W_]+$")
    for i, (bi, li, a) in enumerate(spans):
        if symbols.match(a["text"].strip()):
            continue
        A = em_box(a)
        for bj, lj, b in spans[i + 1:]:
            if symbols.match(b["text"].strip()):
                continue
            B = em_box(b)
            if min(A[2], B[2]) - max(A[0], B[0]) > 2 and min(A[3], B[3]) - max(A[1], B[1]) > 2:
                add("text_overlap", "error", a["text"], a["bbox"], f"Overlaps “{b['text'].strip()[:40]}”.")
                break
    return findings


def pptx_to_pngs(pptx, out_dir, soffice, width, layout=None):
    pm = pymupdf_or_none() or sys.exit("missing dependency PyMuPDF: python -m pip install pymupdf")
    out_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        profile = Path(tmp, "profile").resolve().as_uri()
        cmd = [soffice, f"-env:UserInstallation={profile}", "--headless", "--norestore",
               "--convert-to", "pdf", "--outdir", tmp, str(pptx)]
        subprocess.run(cmd, check=True, capture_output=True, timeout=300)
        pdf = Path(tmp, Path(pptx).stem + ".pdf")
        if not pdf.exists():
            raise RuntimeError(f"LibreOffice did not produce a PDF for {pptx}")
        doc = pm.open(pdf)
        paths = []
        for i, page in enumerate(doc):
            zoom = width / page.rect.width
            pix = page.get_pixmap(matrix=pm.Matrix(zoom, zoom))
            path = out_dir / f"slide-{i + 1:02d}.png"
            pix.save(path)
            paths.append(path)
            if layout is not None:
                layout.append({"slide": i + 1, "scale": zoom, "findings": pdf_layout(page, pm)})
        doc.close()
    return paths


NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def font_inventory(pptx):
    """Typefaces actually set on slides and in the theme (latin vs East Asian)."""
    used = {"latin": set(), "ea": set()}
    themes = {}
    with zipfile.ZipFile(pptx) as z:
        for name in z.namelist():
            if name.startswith("ppt/slides/slide") and name.endswith(".xml"):
                root = ET.fromstring(z.read(name))
                for kind in ("latin", "ea"):
                    for el in root.iter(f"{{{NS['a']}}}{kind}"):
                        if el.get("typeface"):
                            used[kind].add(el.get("typeface"))
            elif name.startswith("ppt/theme/theme") and name.endswith(".xml"):
                root = ET.fromstring(z.read(name))
                theme = {}
                for role in ("majorFont", "minorFont"):
                    el = root.find(f".//a:{role}", NS)
                    if el is not None:
                        theme[role] = {k: (el.find(f"a:{k}", NS).get("typeface") if el.find(f"a:{k}", NS) is not None else None)
                                       for k in ("latin", "ea")}
                themes[name] = theme
    return {"slides_latin": sorted(used["latin"]), "slides_ea": sorted(used["ea"]), "themes": themes}


def inventory(pptx):
    try:
        from pptx import Presentation
        from pptx.enum.shapes import MSO_SHAPE_TYPE
    except ImportError:
        return None
    prs = Presentation(pptx)
    area = prs.slide_width * prs.slide_height
    slides = []

    def walk(shapes):
        for s in shapes:
            if s.shape_type == MSO_SHAPE_TYPE.GROUP:
                yield from walk(s.shapes)
            else:
                yield s
    for idx, slide in enumerate(prs.slides, 1):
        c = {"slide": idx, "text_shapes": 0, "tables": 0, "charts": 0, "pictures": 0, "notes": False}
        big_picture = False
        for s in walk(slide.shapes):
            if s.has_text_frame and s.text_frame.text.strip():
                c["text_shapes"] += 1
            if getattr(s, "has_table", False) and s.has_table:
                c["tables"] += 1
            if getattr(s, "has_chart", False) and s.has_chart:
                c["charts"] += 1
            if s.shape_type == MSO_SHAPE_TYPE.PICTURE:
                c["pictures"] += 1
                if s.width and s.height and s.width * s.height >= 0.85 * area:
                    big_picture = True
        c["notes"] = slide.has_notes_slide and bool(slide.notes_slide.notes_text_frame.text.strip())
        if big_picture and not (c["text_shapes"] or c["tables"] or c["charts"]):
            c["warning"] = "full-slide picture with no native text/table/chart"
        slides.append(c)
    return slides


def walk_shapes(shapes, parent=()):
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    for shape in shapes:
        path = parent + (shape.shape_id,)
        yield path, shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from walk_shapes(shape.shapes, path)


def locate_shape(prs, target):
    if not 1 <= target["slide"] <= len(prs.slides) or not target["path"]:
        raise IndexError(f"invalid slide or empty shape path: {target}")
    shapes = prs.slides[target["slide"] - 1].shapes
    for index, ident in enumerate(target["path"]):
        shape = next((s for s in shapes if s.shape_id == ident), None)
        if shape is None:
            raise KeyError(f"shape not found: {target}")
        if index < len(target["path"]) - 1:
            if not hasattr(shape, "shapes"):
                raise KeyError(f"path crosses a non-group shape: {target}")
            shapes = shape.shapes
    return shape


def default_targets(prs):
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    targets = []
    for number, slide in enumerate(prs.slides, 1):
        text_parents = set()
        for path, shape in walk_shapes(slide.shapes):
            base = {"slide": number, "path": list(path)}
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                targets.append(dict(base, kind="move"))
            if shape.has_text_frame and shape.text_frame.text.strip() and path[:-1] not in text_parents:
                targets.append(dict(base, kind="text"))
                text_parents.add(path[:-1])
            if getattr(shape, "has_table", False):
                targets.append(dict(base, kind="table", cell=[0, 0]))
            if getattr(shape, "has_chart", False):
                targets.append(dict(base, kind="chart"))
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                targets.append(dict(base, kind="picture"))
    return targets


def workbook_cells(chart):
    part = chart.part.chart_workbook.xlsx_part
    if part is None:
        raise NotImplementedError("No embedded workbook; external-linked charts are not edited.")
    return workbook_blob_cells(part.blob)


def workbook_blob_cells(blob):
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(BytesIO(blob)) as z:
        sheets = [n for n in z.namelist() if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]
        if len(sheets) != 1:
            raise NotImplementedError("Only a single embedded worksheet is supported.")
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            for item in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("s:si", ns):
                shared.append("".join(item.itertext()))
        result = {}
        for cell in ET.fromstring(z.read(sheets[0])).findall(".//s:c", ns):
            if cell.find("s:f", ns) is not None:
                raise NotImplementedError("Formula-based workbooks require a dedicated edit test.")
            value = cell.find("s:v", ns)
            if cell.get("t") == "inlineStr":
                result[cell.get("r")] = "".join(cell.find("s:is", ns).itertext())
            elif value is not None and value.text is not None:  # PptxGenJS writes empty <v/> cells
                result[cell.get("r")] = shared[int(value.text)] if cell.get("t") == "s" else float(value.text)
        return result


def chart_snapshot(chart):
    supported = {"BAR_CLUSTERED", "COLUMN_CLUSTERED", "LINE"}
    if chart.chart_type.name not in supported or len(chart.plots) != 1:
        raise NotImplementedError("Automatic editing supports single-plot clustered bar/column and line charts.")
    plot = chart.plots[0]
    if plot.categories.depth != 1:
        raise NotImplementedError("Multi-level categories require a dedicated test.")
    categories = [c.label for c in plot.categories]
    if any(not isinstance(c, str) for c in categories):
        raise NotImplementedError("Date/numeric category axes require a dedicated test.")
    series = [{"name": s.name, "values": list(s.values)} for s in plot.series]
    cells = workbook_cells(chart)
    # replace_data rebuilds the workbook. Only test a known, lossless table layout.
    from pptx.chart.data import CategoryChartData
    data = CategoryChartData()
    data.categories = categories
    for item in series:
        data.add_series(item["name"], item["values"])
    if not equivalent(cells, workbook_blob_cells(data.xlsx_blob)):
        raise NotImplementedError("Nonstandard workbook layout/content requires a dedicated edit test.")
    return {"categories": categories, "series": series, "workbook": cells}


def structure(shape):
    # IDs and endpoint bindings survive a move; this does not prove automatic line routing.
    items = shape.shapes if hasattr(shape, "shapes") else [shape]
    return [{"path": list(path), "type": int(s.shape_type),
             "connections": [dict(el.attrib) for el in s._element.iter()
                             if str(el.tag).split("}")[-1] in ("stCxn", "endCxn")]}
            for path, s in walk_shapes(items)]


def snapshot(shape, target):
    kind = target["kind"]
    if kind == "text":
        return shape.text_frame.text
    if kind == "table":
        return shape.table.cell(*target.get("cell", [0, 0])).text
    if kind == "chart":
        return chart_snapshot(shape.chart)
    if kind == "move":
        return {"left": shape.left, "top": shape.top, "structure": structure(shape)}
    if kind == "picture":
        return [shape.crop_left, shape.crop_right, shape.crop_top, shape.crop_bottom]
    raise ValueError(f"unknown edit kind: {kind}")


def equivalent(actual, expected):
    if isinstance(expected, (int, float)):
        return isinstance(actual, (int, float)) and math.isclose(actual, expected, rel_tol=1e-9, abs_tol=1e-10)
    if isinstance(expected, dict):
        return isinstance(actual, dict) and actual.keys() == expected.keys() and all(
            equivalent(actual[k], v) for k,v in expected.items())
    if isinstance(expected, list):
        return isinstance(actual, list) and len(actual) == len(expected) and all(
            equivalent(a,b) for a,b in zip(actual, expected))
    return actual == expected


def verify_edits(prs, records):
    for record in records:
        if record["status"] in ("unsupported", "failed"):
            continue
        try:
            record["actual"] = snapshot(locate_shape(prs, record["target"]), record["target"])
            record["status"] = "passed" if equivalent(record["actual"], record["expected"]) else "failed"
        except (KeyError, IndexError, NotImplementedError) as error:
            record["status"], record["detail"] = "failed", str(error)
        except Exception as error:  # report the object instead of aborting the whole check
            record["status"], record["detail"] = "failed", f"checker error {type(error).__name__}: {error}"
    return records


def edit_test(pptx, out_dir, soffice=None, width=1600, targets=None):
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.util import Inches
    copy = out_dir / "edit-test.pptx"
    if Path(pptx).resolve() == copy.resolve():
        raise ValueError("Edit-test copy must differ from the source.")
    out_dir.mkdir(parents=True, exist_ok=True)
    prs = Presentation(pptx)
    selected = default_targets(prs) if targets is None else targets
    identities = [(t["slide"], tuple(t["path"]), t["kind"], tuple(t.get("cell", []))) for t in selected]
    if len(set(identities)) != len(identities):
        raise ValueError("Duplicate edit targets; select each behavior on an object once.")
    records = []
    for number, target in enumerate(selected):
        record = check(f"edit:{number+1}:{target['kind']}", "not_run", target=target)
        records.append(record)
        try:
            shape = locate_shape(prs, target)
            before = snapshot(shape, target)
            record["before"] = before
            kind = target["kind"]
            if kind in ("text", "table"):
                frame = shape.text_frame if kind == "text" else shape.table.cell(*target.get("cell", [0, 0])).text_frame
                runs = [r for p in frame.paragraphs for r in p.runs]
                marker = f" [E{number+1}]"
                if runs:
                    runs[0].text += marker
                else:
                    frame.paragraphs[0].add_run().text = marker
                record["expected"] = frame.text
            elif kind == "chart":
                if not before["series"]:
                    raise NotImplementedError("Empty chart.")
                values = before["series"][0]["values"]
                point = next((i for i,v in enumerate(values) if v is not None and math.isfinite(v)), None)
                if point is None:
                    raise NotImplementedError("No finite numeric point to edit.")
                expected = json.loads(json.dumps(before))
                value = round(values[point] + abs(values[point]) * 0.1 + 1, 6)
                expected["series"][0]["values"][point] = value
                expected["workbook"][f"B{point+2}"] = value
                data = CategoryChartData()
                data.categories = before["categories"]
                for series in expected["series"]:
                    data.add_series(series["name"], series["values"])
                shape.chart.replace_data(data)
                record["expected"] = expected
            elif kind == "move":
                shape.left += Inches(0.1)
                record["expected"] = dict(before, left=before["left"] + Inches(0.1))
            elif kind == "picture":
                shape.crop_left = 0.1 if abs(shape.crop_left-0.1) > 0.001 else 0.05
                record["expected"] = [shape.crop_left, *before[1:]]
                record["detail"] = "Crop round-trip only; replacement and visual crop quality require inspection."
        except NotImplementedError as error:
            record["status"], record["detail"] = "unsupported", str(error)
        except (KeyError, IndexError) as error:
            record["status"], record["detail"] = "failed", str(error)
        except Exception as error:  # one malformed object must not abort the whole report
            record["status"], record["detail"] = "failed", f"checker error {type(error).__name__}: {error}"
    prs.save(copy)
    verify_edits(Presentation(copy), records)
    rendered, render_note = [], "not rendered: LibreOffice or PyMuPDF unavailable"
    if soffice and pymupdf_or_none():
        try:
            rendered, render_note = pptx_to_pngs(copy, out_dir / "edit-test", soffice, width), None
        except (OSError, subprocess.SubprocessError, RuntimeError) as error:
            render_note = f"rendering the edited copy failed: {error}"
    return {"copy": str(copy), "checks": records,
            "selection": "explicit" if targets is not None else "one text per page/group, all tables/charts/groups/pictures",
            "rendered": [str(p) for p in rendered], "render_note": render_note,
            "scope": "Library save/reload; manual edits, auto-connect and data-linked labels need task-specific checks."}


def cmd_pptx(args):
    src = Path(args.input).resolve()
    out = Path(args.out)
    slides_dir = out / "slides"
    if slides_dir.exists():
        shutil.rmtree(slides_dir)
    # Without LibreOffice or PyMuPDF, rendering and the layout check are reported as not_run;
    # the inventory, font report and edit test still run.
    soffice = find_soffice(args.soffice)
    can_render = bool(soffice) and pymupdf_or_none() is not None
    layout, pngs, render_note, render_failed = [], [], None, False
    if not can_render:
        render_note = ("LibreOffice not found" if not soffice else "PyMuPDF not installed") + \
            "; rendering and the layout check were not run (inventory, fonts and edit test still run)."
        print("warning: " + render_note, file=sys.stderr)
    else:
        try:
            pngs = pptx_to_pngs(src, slides_dir, soffice, args.width, layout)
        except (OSError, subprocess.SubprocessError, RuntimeError) as error:
            render_note, render_failed, layout = f"rendering failed: {error}", True, []
            print("warning: " + render_note, file=sys.stderr)
    report = {"file": str(src), "renderer": "LibreOffice (approximation of PowerPoint)" if pngs else None,
              "render_note": render_note, "slides": len(pngs), "fonts": font_inventory(src),
              "inventory": inventory(src), "layout": layout}
    layout_dir = out / "layout"
    if layout_dir.exists():
        shutil.rmtree(layout_dir)
    for item in layout:
        if item["findings"]:
            draw_layout(slides_dir / f"slide-{item['slide']:02d}.png", item["findings"],
                        layout_dir / f"slide-{item['slide']:02d}.png", item["scale"])
    layout_errors = sum(f["severity"] == "error" for item in layout for f in item["findings"])
    layout_warnings = sum(f["severity"] == "warning" for item in layout for f in item["findings"])
    if args.edit_test:
        if report["inventory"] is None:
            report["edit_test"] = "skipped: python-pptx not installed"
        else:
            targets = json.loads(Path(args.edit_targets).read_text(encoding="utf-8")) if args.edit_targets else None
            report["edit_test"] = edit_test(src, out, soffice if can_render else None, args.width, targets)
    if pngs:
        make_sheet(pngs, out / "sheet.png")
        print(f"slides: {len(pngs)} -> {slides_dir}")
        print(f"sheet  -> {out / 'sheet.png'} (+ -gray, -blur)")
    else:
        print(f"slides: not rendered ({render_note})")
    print(f"fonts: {report['fonts']}")
    if report["inventory"]:
        for c in report["inventory"]:
            print("  " + json.dumps(c, ensure_ascii=False))
    if "edit_test" in report:
        edit = report["edit_test"]
        if isinstance(edit, dict):
            print(f"edit test: {checks_status(edit['checks'])}, {len(edit['checks'])} targets; details in report.json")
        else:
            print(f"edit test: {edit}")
    if pngs:
        print(f"layout: {layout_errors} errors, {layout_warnings} warnings"
              + (f" -> annotated pages in {layout_dir}" if layout_errors or layout_warnings else ""))
    else:
        print("layout: not run (needs a rendering)")
    for item in layout:
        for f in item["findings"]:
            print(f"  slide {item['slide']:02d} {f['severity']:<7} {f['kind']:<18} “{f['text']}” {f['detail']}")

    expected_pages = len(report["inventory"]) if report["inventory"] is not None else None
    if not can_render:
        render_status = "not_run"
    elif render_failed:
        render_status = "failed"
    else:
        render_status = ("unsupported" if expected_pages is None else
                         "passed" if len(pngs) == expected_pages else "failed")
    report["checks"] = [check("render", render_status, render_note or "Compare rendered pages with source slides.",
                              expected=expected_pages, actual=len(pngs)),
                        check("inventory", "passed" if report["inventory"] is not None else "unsupported")]
    if pngs:
        report["checks"].append(check(
            "layout", "failed" if layout_errors else "passed",
            f"{layout_errors} errors (outside slide, overlapping text), {layout_warnings} warnings "
            "(orphan line, small font, placeholder). Measured on the LibreOffice rendering; "
            "PowerPoint/WPS text metrics can differ. No contrast check for PPTX.",
            errors=layout_errors, warnings=layout_warnings))
    else:
        report["checks"].append(check("layout", "not_run", "Needs a rendering; " + (render_note or "")))
    edit = report.get("edit_test")
    if isinstance(edit, dict) and edit["checks"]:
        report["checks"].extend(edit["checks"])
    else:
        report["checks"].append(check("edit_test", "not_run" if edit is None or isinstance(edit, dict) else "unsupported",
                                      "No edit targets verified. Supply --edit-test and task-specific targets if needed."))
    write_report(report, out, args.strict)


# ---------------------------------------------------------------- CLI

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    h = sub.add_parser("html", help="screenshot and check an HTML deck")
    h.add_argument("input")
    h.add_argument("--out", required=True)
    h.add_argument("--scale", type=float, default=1.25, help="device scale factor (1.25 -> 1600px wide)")
    h.add_argument("--browser", choices=["auto", "chromium", "chrome", "msedge"], default="auto")
    h.add_argument("--no-pdf", action="store_true", help="skip the print-to-PDF check")
    h.add_argument("--strict", action="store_true", help="nonzero exit unless required automatic checks pass")
    h.set_defaults(func=cmd_html)

    x = sub.add_parser("pptx", help="render a PPTX through LibreOffice and inspect native objects")
    x.add_argument("input")
    x.add_argument("--out", required=True)
    x.add_argument("--width", type=int, default=1600)
    x.add_argument("--soffice")
    x.add_argument("--edit-test", action="store_true", help="edit text/table/chart in a copy, reload and render")
    x.add_argument("--edit-targets", help="JSON list of slide (1-based), shape-id path, kind, optional table cell")
    x.add_argument("--strict", action="store_true", help="nonzero exit unless required automatic checks pass")
    x.set_defaults(func=cmd_pptx)

    s = sub.add_parser("sheet", help="contact sheet from PNG files or a directory")
    s.add_argument("inputs", nargs="+")
    s.add_argument("--out", required=True)
    s.add_argument("--cols", type=int)
    s.add_argument("--thumb", type=int, default=400)
    s.set_defaults(func=lambda a: print(make_sheet(
        [q for i in a.inputs for q in (sorted(Path(i).glob("*.png")) if Path(i).is_dir() else [Path(i)])],
        a.out, a.cols, a.thumb)))

    c = sub.add_parser("compare", help="place images side by side")
    c.add_argument("inputs", nargs="+")
    c.add_argument("--out", required=True)
    c.add_argument("--labels", help="comma-separated labels (default A, B, C ...)")
    c.add_argument("--height", type=int, default=540)
    c.set_defaults(func=lambda a: print(make_compare(
        a.inputs, a.out, a.labels.split(",") if a.labels else None, a.height)))

    args = ap.parse_args()
    if getattr(args, "edit_targets", None) and not args.edit_test:
        ap.error("--edit-targets requires --edit-test")
    args.func(args)


if __name__ == "__main__":
    main()
