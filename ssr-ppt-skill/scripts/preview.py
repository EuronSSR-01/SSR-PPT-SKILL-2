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
    rendered = pptx_to_pngs(copy, out_dir / "edit-test", soffice, width) if soffice else []
    return {"copy": str(copy), "checks": records,
            "selection": "explicit" if targets is not None else "one text per page/group, all tables/charts/groups/pictures",
            "rendered": [str(p) for p in rendered],
            "scope": "Library save/reload; manual edits, auto-connect and data-linked labels need task-specific checks."}


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
            targets = json.loads(Path(args.edit_targets).read_text(encoding="utf-8")) if args.edit_targets else None
            report["edit_test"] = edit_test(src, out, soffice, args.width, targets)
    make_sheet(pngs, out / "sheet.png")
    print(f"slides: {len(pngs)} -> {slides_dir}")
    print(f"sheet  -> {out / 'sheet.png'} (+ -gray, -blur)")
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

    expected_pages = len(report["inventory"]) if report["inventory"] is not None else None
    render_status = ("unsupported" if expected_pages is None else
                     "passed" if len(pngs) == expected_pages else "failed")
    report["checks"] = [check("render", render_status, "Compare rendered pages with source slides.",
                              expected=expected_pages, actual=len(pngs)),
                        check("inventory", "passed" if report["inventory"] is not None else "unsupported")]
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
