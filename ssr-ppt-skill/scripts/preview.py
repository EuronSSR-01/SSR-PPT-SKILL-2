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
import json
import math
import os
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
  await Promise.all(Array.from(document.images).map(img => img.decode().catch(() => null)));
  return document.fonts.status;
}"""


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
        report["fonts_status"] = page.evaluate(WAIT_ASSETS_JS)

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
        for i in range(n):
            if info["api"]:
                page.evaluate("i => SSRPlayer.go(i, { reveal: 'all' })", i)
            else:
                page.keyboard.press("Home")
                for _ in range(i):
                    page.keyboard.press("ArrowRight")
            page.wait_for_timeout(60)
            page.locator("#ssr-deck").screenshot(path=str(slides_dir / f"slide-{i + 1:02d}.png"))

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
    (out / "report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
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


# ---------------------------------------------------------------- PPTX

def find_soffice(explicit):
    if explicit:
        return explicit
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
    sys.exit("LibreOffice (soffice) not found; install it or pass --soffice PATH")


def pptx_to_pngs(pptx, out_dir, soffice, width):
    pm = pymupdf_or_none() or sys.exit("missing dependency PyMuPDF: python -m pip install pymupdf")
    out_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        profile = Path(tmp, "profile").resolve().as_uri()
        cmd = [soffice, f"-env:UserInstallation={profile}", "--headless", "--norestore",
               "--convert-to", "pdf", "--outdir", tmp, str(pptx)]
        subprocess.run(cmd, check=True, capture_output=True, timeout=300)
        pdf = Path(tmp, Path(pptx).stem + ".pdf")
        if not pdf.exists():
            sys.exit(f"LibreOffice did not produce a PDF for {pptx}")
        doc = pm.open(pdf)
        paths = []
        for i, page in enumerate(doc):
            zoom = width / page.rect.width
            pix = page.get_pixmap(matrix=pm.Matrix(zoom, zoom))
            path = out_dir / f"slide-{i + 1:02d}.png"
            pix.save(path)
            paths.append(path)
        doc.close()
    return paths


NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def font_inventory(pptx):
    """Typefaces actually set on slides and in the theme (latin vs East Asian)."""
    used = {"latin": set(), "ea": set()}
    theme = {}
    with zipfile.ZipFile(pptx) as z:
        for name in z.namelist():
            if name.startswith("ppt/slides/slide") and name.endswith(".xml"):
                root = ET.fromstring(z.read(name))
                for kind in ("latin", "ea"):
                    for el in root.iter(f"{{{NS['a']}}}{kind}"):
                        if el.get("typeface"):
                            used[kind].add(el.get("typeface"))
            elif name.startswith("ppt/theme/theme") and not theme:
                root = ET.fromstring(z.read(name))
                for role in ("majorFont", "minorFont"):
                    el = root.find(f".//a:{role}", NS)
                    if el is not None:
                        theme[role] = {k: (el.find(f"a:{k}", NS).get("typeface") if el.find(f"a:{k}", NS) is not None else None)
                                       for k in ("latin", "ea")}
    return {"slides_latin": sorted(used["latin"]), "slides_ea": sorted(used["ea"]), "theme": theme}


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


def edit_test(pptx, out_dir, soffice, width):
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    copy = out_dir / "edit-test.pptx"
    shutil.copyfile(pptx, copy)
    prs = Presentation(copy)
    done, touched = {}, set()
    marker = "【编辑测试】"
    for idx, slide in enumerate(prs.slides, 1):
        for s in slide.shapes:
            if "text" not in done and s.has_text_frame:
                runs = [r for p in s.text_frame.paragraphs for r in p.runs if r.text.strip()]
                if runs:
                    runs[0].text += marker
                    done["text"] = idx
                    touched.add(idx)
            if "table" not in done and getattr(s, "has_table", False) and s.has_table:
                cell = s.table.cell(0, 0)
                runs = [r for p in cell.text_frame.paragraphs for r in p.runs]
                if runs:
                    runs[0].text += "✓"
                else:
                    cell.text = "✓"
                done["table"] = idx
                touched.add(idx)
            if "chart" not in done and getattr(s, "has_chart", False) and s.has_chart:
                try:
                    plot = s.chart.plots[0]
                    cd = CategoryChartData()
                    cd.categories = list(plot.categories)
                    for k, ser in enumerate(plot.series):
                        vals = list(ser.values)
                        if k == 0:
                            vals[0] = (vals[0] or 0) * 1.5 + 1
                        cd.add_series(ser.name, vals)
                    s.chart.replace_data(cd)
                    done["chart"] = idx
                    touched.add(idx)
                except Exception as e:  # e.g. XY or multi-level category charts
                    done["chart_skipped"] = f"slide {idx}: {e}"
    prs.save(copy)

    # Re-open and confirm the edits survived a save/load round trip.
    prs = Presentation(copy)
    verified = {}
    if "text" in done:
        verified["text"] = marker in "".join(
            sh.text_frame.text for sh in prs.slides[done["text"] - 1].shapes if sh.has_text_frame)
    if "table" in done:
        verified["table"] = any("✓" in sh.table.cell(0, 0).text for sh in prs.slides[done["table"] - 1].shapes
                                if getattr(sh, "has_table", False) and sh.has_table)
    if "chart" in done:
        verified["chart"] = any(sh.has_chart for sh in prs.slides[done["chart"] - 1].shapes
                                if getattr(sh, "has_chart", False))
    pngs = pptx_to_pngs(copy, out_dir / "edit-test", soffice, width)
    keep = {f"slide-{i:02d}.png" for i in touched}
    for pth in pngs:
        if pth.name not in keep:
            pth.unlink()
    return {"edited": done, "verified_after_reload": verified,
            "rendered": sorted(str(out_dir / "edit-test" / k) for k in keep)}


def cmd_pptx(args):
    src = Path(args.input).resolve()
    out = Path(args.out)
    slides_dir = out / "slides"
    if slides_dir.exists():
        shutil.rmtree(slides_dir)
    soffice = find_soffice(args.soffice)
    pngs = pptx_to_pngs(src, slides_dir, soffice, args.width)
    report = {"file": str(src), "renderer": "LibreOffice (approximation of PowerPoint)",
              "slides": len(pngs), "fonts": font_inventory(src), "inventory": inventory(src)}
    if args.edit_test:
        if report["inventory"] is None:
            report["edit_test"] = "skipped: python-pptx not installed"
        else:
            report["edit_test"] = edit_test(src, out, soffice, args.width)
    make_sheet(pngs, out / "sheet.png")
    (out / "report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"slides: {len(pngs)} -> {slides_dir}")
    print(f"sheet  -> {out / 'sheet.png'} (+ -gray, -blur)")
    print(f"fonts: {report['fonts']}")
    if report["inventory"]:
        for c in report["inventory"]:
            print("  " + json.dumps(c, ensure_ascii=False))
    if "edit_test" in report:
        print(f"edit test: {json.dumps(report['edit_test'], ensure_ascii=False)}")


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
    h.set_defaults(func=cmd_html)

    x = sub.add_parser("pptx", help="render a PPTX through LibreOffice and inspect native objects")
    x.add_argument("input")
    x.add_argument("--out", required=True)
    x.add_argument("--width", type=int, default=1600)
    x.add_argument("--soffice")
    x.add_argument("--edit-test", action="store_true", help="edit text/table/chart in a copy, reload and render")
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
    args.func(args)


if __name__ == "__main__":
    main()
