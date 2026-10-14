#!/usr/bin/env python3
"""Report what this environment can do for SSR-PPT-SKILL and which route to take.

  python doctor.py            human-readable report
  python doctor.py --json     machine-readable report
  python doctor.py --probe    also launch a headless browser through Playwright (slower, more certain)

Standard library only; nothing is installed or changed. Whether the model itself can view images
cannot be detected from a script: the agent must answer that and pick the matching review route.
"""
import argparse
import importlib.metadata
import importlib.util
import json
import os
import platform
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

PACKAGES = {  # import name -> (distribution name, purpose)
    "pptx": ("python-pptx", "build and edit PPTX, edit tests"),
    "playwright": ("playwright", "HTML screenshots, interaction, print and layout checks"),
    "PIL": ("pillow", "contact sheets and image processing"),
    "pymupdf": ("pymupdf", "render the LibreOffice PDF to PNG"),
    "fontTools": ("fonttools", "subset and embed fonts in HTML"),
    "brotli": ("brotli", "WOFF2 output for embedded fonts (WOFF without it)"),
    "lxml": ("lxml", "OOXML edits that python-pptx does not expose"),
}

BROWSERS = {
    "Windows": [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"],
    "Darwin": ["/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
               "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
               "/Applications/Chromium.app/Contents/MacOS/Chromium"],
    "Linux": [],
}
BROWSER_COMMANDS = ["google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "msedge", "microsoft-edge"]

SOFFICE = [r"C:\Program Files\LibreOffice\program\soffice.com", r"C:\Program Files (x86)\LibreOffice\program\soffice.com",
           "/Applications/LibreOffice.app/Contents/MacOS/soffice", "/usr/bin/soffice",
           "/usr/lib/libreoffice/program/soffice", "/opt/libreoffice/program/soffice"]

# Font file name fragments -> family. Only fonts whose licenses allow embedding are marked embeddable.
CJK_FONTS = {
    "notosanssc": ("Noto Sans SC", True), "notoserifsc": ("Noto Serif SC", True),
    "sourcehansans": ("Source Han Sans", True), "sourcehanserif": ("Source Han Serif", True),
    "lxgwwenkai": ("LXGW WenKai", True), "smileysans": ("Smiley Sans", True),
    "harmonyos_sans": ("HarmonyOS Sans", False), "misans": ("MiSans", False),
    "alibabapuhuiti": ("Alibaba PuHuiTi", False), "opposans": ("OPPO Sans", False),
    "msyh": ("Microsoft YaHei (system only)", False), "deng": ("DengXian (system only)", False),
    "pingfang": ("PingFang SC (system only)", False), "simsun": ("SimSun (system only)", False),
}


def font_dirs():
    home = Path.home()
    dirs = [Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts",
            home / "AppData/Local/Microsoft/Windows/Fonts",
            Path("/Library/Fonts"), Path("/System/Library/Fonts"), home / "Library/Fonts",
            Path("/usr/share/fonts"), Path("/usr/local/share/fonts"), home / ".local/share/fonts", home / ".fonts"]
    return [d for d in dirs if d.is_dir()]


def find_fonts():
    found = {}
    for d in font_dirs():
        try:
            files = [p for p in d.rglob("*") if p.suffix.lower() in (".ttf", ".otf", ".ttc", ".otc")]
        except OSError:
            continue
        for p in files:
            key = p.name.lower().replace(" ", "").replace("-", "")
            for frag, (family, embeddable) in CJK_FONTS.items():
                if frag in key and family not in found:
                    found[family] = {"file": str(p), "embeddable_license": embeddable}
    return found


def run(cmd, timeout=20):
    try:
        out = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        return (out.stdout or out.stderr).strip().splitlines()[0] if (out.stdout or out.stderr) else ""
    except (OSError, subprocess.SubprocessError):
        return None


def probe_browser():
    code = ("from playwright.sync_api import sync_playwright\n"
            "with sync_playwright() as p:\n"
            "    for ch in (None, 'chrome', 'msedge'):\n"
            "        try:\n"
            "            b = p.chromium.launch(channel=ch) if ch else p.chromium.launch(); b.close(); print(ch or 'chromium'); break\n"
            "        except Exception: pass\n")
    return run([sys.executable, "-c", code], timeout=90) or None


def collect(probe):
    report = {"platform": f"{platform.system()} {platform.release()}", "python": sys.version.split()[0],
              "packages": {}, "browser": {}, "libreoffice": None, "node": None, "pptxgenjs": False,
              "fonts": {}, "temp_writable": False}
    for mod, (dist, purpose) in PACKAGES.items():
        spec = importlib.util.find_spec(mod) or (importlib.util.find_spec("fitz") if mod == "pymupdf" else None)
        version = None
        if spec:
            try:
                version = importlib.metadata.version(dist)
            except importlib.metadata.PackageNotFoundError:
                version = "?"
        report["packages"][dist] = {"installed": bool(spec), "version": version, "purpose": purpose}

    installed = [p for p in BROWSERS.get(platform.system(), []) if Path(p).exists()]
    installed += [shutil.which(c) for c in BROWSER_COMMANDS if shutil.which(c)]
    pw_cache = [Path(os.environ.get("PLAYWRIGHT_BROWSERS_PATH", "")),
                Path.home() / "AppData/Local/ms-playwright", Path.home() / "Library/Caches/ms-playwright",
                Path.home() / ".cache/ms-playwright"]
    bundled = any(d.is_dir() and any(d.glob("chromium*")) for d in pw_cache if str(d) not in ("", "."))
    report["browser"] = {"installed": installed, "playwright_bundled_chromium": bundled,
                         "probe": probe_browser() if probe and report["packages"]["playwright"]["installed"] else None}

    soffice = next((shutil.which(n) for n in ("soffice.com", "soffice", "libreoffice") if shutil.which(n)), None)
    soffice = soffice or next((p for p in SOFFICE if Path(p).exists()), None)
    report["libreoffice"] = soffice

    node = shutil.which("node")
    if node:
        report["node"] = run([node, "-v"])
        resolved = run([node, "-e", "console.log(require.resolve('pptxgenjs'))"])
        report["pptxgenjs"] = bool(resolved) and "pptxgenjs" in resolved

    report["fonts"] = find_fonts()
    try:
        with tempfile.NamedTemporaryFile(prefix="ssr-doctor-"):
            report["temp_writable"] = True
    except OSError:
        pass
    return report


def routes(r):
    pkg = {k: v["installed"] for k, v in r["packages"].items()}
    has_browser = bool(r["browser"]["probe"] or r["browser"]["installed"] or r["browser"]["playwright_bundled_chromium"])
    html_check = pkg["playwright"] and has_browser
    render = bool(r["libreoffice"]) and pkg["pymupdf"]
    embeddable = [f for f, v in r["fonts"].items() if v["embeddable_license"]]
    out = {
        "pptx_build": "python-pptx (default)" if pkg["python-pptx"] else
                      "PptxGenJS (Node)" if r["pptxgenjs"] else "missing: install python-pptx",
        "html_review": "preview.py html: screenshots, interaction, print and layout checks" if html_check else
                       "missing: playwright and a Chromium-family browser; HTML review limited to structure",
        "pptx_render": "preview.py pptx: LibreOffice -> PDF -> PNG, plus layout check" if render else
                       "missing: LibreOffice and/or PyMuPDF; render and layout check reported as not_run",
        "pptx_edit_test": ("missing: python-pptx" if not pkg["python-pptx"] else
                           "preview.py pptx --edit-test" if render else
                           "preview.py pptx --edit-test (structure, fonts and edit test only; no rendering)"),
        "html_fonts": (f"embed with embed_fonts.py: {', '.join(embeddable)}" if pkg["fonttools"] and embeddable else
                       "system fonts with fallbacks (no embeddable CJK font or fontTools)"),
        "visual_review": "If the model can view images: open the PNGs from preview.py. If not: rely on the "
                         "layout report from preview.py html and state 'not visually reviewed' in the delivery notes.",
        "text_hints": "visual-first decks: add --visual-first to preview.py html/pptx to list long paragraphs and "
                      "boxes of points (hints only; the pptx check needs no rendering)" if pkg["python-pptx"] or html_check else
                      "missing: python-pptx or an HTML check; review text density by reading the planning",
    }
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--probe", action="store_true", help="launch a headless browser to confirm Playwright works")
    args = ap.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    report = collect(args.probe)
    report["routes"] = routes(report)
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return
    print(f"SSR-PPT-SKILL doctor  |  {report['platform']}  |  Python {report['python']}")
    print("\nPython packages:")
    for dist, v in report["packages"].items():
        mark = "ok " if v["installed"] else "-- "
        print(f"  {mark} {dist:<12} {v['version'] or '':<10} {v['purpose']}")
    b = report["browser"]
    print("\nBrowser:", ", ".join(b["installed"]) or "none found",
          "| Playwright Chromium:", "yes" if b["playwright_bundled_chromium"] else "no",
          f"| probe: {b['probe']}" if b["probe"] is not None else "")
    print("LibreOffice:", report["libreoffice"] or "not found")
    print("Node:", report["node"] or "not found", "| pptxgenjs:", "yes" if report["pptxgenjs"] else "no")
    print("CJK fonts:", ", ".join(f"{k}{'' if v['embeddable_license'] else ' (do not embed)'}"
                                  for k, v in report["fonts"].items()) or "none found")
    print("Temp dir writable:", report["temp_writable"])
    print("\nRecommended route:")
    for k, v in report["routes"].items():
        print(f"  {k:<15} {v}")
    missing = [d for d, v in report["packages"].items() if not v["installed"]]
    if missing:
        print("\nInstall missing pieces only when the task needs them and the user allows it:\n"
              f"  python -m pip install {' '.join(missing)}")


if __name__ == "__main__":
    main()
