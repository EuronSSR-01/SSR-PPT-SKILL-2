#!/usr/bin/env python3
"""Subset fonts to the characters used in an HTML deck and inline them as base64 @font-face rules.

Usage:
  python embed_fonts.py index.html --font "Noto Sans SC=fonts/NotoSansSC-VF.ttf" [--font ...] [-o out.html]

Font spec: FAMILY=PATH[#TTC_INDEX][@WEIGHT|@MIN-MAX][:italic]
  - A variable font without @ keeps its full weight range.
  - A variable font with a single @WEIGHT is pinned to a static instance (smaller file).
  - A variable font with @MIN-MAX is limited to that weight range.

Re-running replaces the previous <style id="ssr-embedded-fonts"> block, so run it again after text changes.
Requires fontTools; WOFF2 output also needs brotli (falls back to WOFF otherwise).
"""
import argparse
import base64
import html
import io
import re
import sys
from pathlib import Path

try:
    from fontTools import subset
    from fontTools.ttLib import TTFont
except ImportError:
    sys.exit("embed_fonts.py needs fontTools: python -m pip install fonttools brotli")

STYLE_ID = "ssr-embedded-fonts"
BLOCK_RE = re.compile(r'<style id="ssr-embedded-fonts">.*?</style>\s*', re.S)
# Printable ASCII plus common CJK punctuation, so small text edits after embedding still render.
BASE_CHARS = (
    "".join(chr(c) for c in range(0x20, 0x7F))
    + " ，。、；：？！“”‘’（）《》〈〉【】「」『』—…·￥%‰°±×÷→←↑↓～"
)


def parse_spec(spec):
    family, sep, rest = spec.partition("=")
    family = family.strip()
    if not sep or not family or not rest:
        raise ValueError(f"invalid --font {spec!r}; expected FAMILY=PATH[#INDEX][@WEIGHT][:italic]")
    style = "normal"
    if rest.endswith(":italic"):
        style, rest = "italic", rest[: -len(":italic")]
    weight = None
    m = re.search(r"@(\d+)(?:-(\d+))?$", rest)
    if m:
        weight = (int(m[1]), int(m[2] or m[1]))
        rest = rest[: m.start()]
    index = 0
    m = re.search(r"#(\d+)$", rest)
    if m:
        index = int(m[1])
        rest = rest[: m.start()]
    path = Path(rest)
    if not path.is_file():
        raise ValueError(f"font file not found: {path}")
    return family, path, index, weight, style


def collect_chars(text):
    text = html.unescape(BLOCK_RE.sub("", text))
    chars = set(text) | set(BASE_CHARS)
    return {c for c in chars if ord(c) >= 0x20 and not 0x7F <= ord(c) <= 0x9F}


def build_face(family, path, index, weight, style, codepoints, flavor):
    font = TTFont(path, fontNumber=index, lazy=True)  # load glyph tables on demand; much faster for CJK fonts
    covered = set(font.getBestCmap() or {})

    opts = subset.Options()
    opts.layout_features = ["*"]  # keep kern, palt, vert, etc.
    opts.name_IDs = ["*"]
    opts.name_languages = ["*"]
    opts.hinting = False
    opts.notdef_outline = True
    subsetter = subset.Subsetter(opts)
    subsetter.populate(unicodes=codepoints & covered)
    subsetter.subset(font)

    axes = {a.axisTag: a for a in font["fvar"].axes} if "fvar" in font else {}
    if "wght" in axes and weight:
        from fontTools.varLib import instancer

        if weight[0] == weight[1]:
            limits = {tag: a.defaultValue for tag, a in axes.items()}
            limits["wght"] = weight[0]
        else:
            limits = {"wght": weight}
        font = instancer.instantiateVariableFont(font, limits)
    if weight is None:
        if "wght" in axes:
            weight = (int(axes["wght"].minValue), int(axes["wght"].maxValue))
        else:
            w = font["OS/2"].usWeightClass if "OS/2" in font else 400
            weight = (w, w)

    font.flavor = flavor
    buf = io.BytesIO()
    font.save(buf)
    data = buf.getvalue()
    weight_css = str(weight[0]) if weight[0] == weight[1] else f"{weight[0]} {weight[1]}"
    name = family.replace("\\", "\\\\").replace('"', '\\"')
    css = (
        f'@font-face{{font-family:"{name}";'
        f"src:url(data:font/{flavor};base64,{base64.b64encode(data).decode('ascii')}) format(\"{flavor}\");"
        f"font-weight:{weight_css};font-style:{style};font-display:block}}"
    )
    return css, covered, len(data), weight_css


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("html", type=Path, help="HTML file to scan and update")
    ap.add_argument("--font", action="append", required=True, help="FAMILY=PATH[#INDEX][@WEIGHT][:italic]")
    ap.add_argument("-o", "--output", type=Path, help="output HTML (default: overwrite input)")
    ap.add_argument("--extra-text", type=Path, action="append", default=[], help="file with extra characters to keep")
    ap.add_argument("--format", choices=["woff2", "woff"], help="font format (default: woff2 if brotli is installed)")
    args = ap.parse_args()

    flavor = args.format
    if flavor is None:
        try:
            import brotli  # noqa: F401

            flavor = "woff2"
        except ImportError:
            flavor = "woff"
            print("brotli not installed; using WOFF (larger). python -m pip install brotli", file=sys.stderr)

    try:
        specs = [parse_spec(s) for s in args.font]
    except ValueError as e:
        sys.exit(str(e))

    text = args.html.read_text(encoding="utf-8")
    if not re.search(r"</head>", text, re.I):
        sys.exit("no </head> found; the deck must have a <head> element")
    chars = collect_chars(text)
    for extra in args.extra_text:
        chars |= collect_chars(extra.read_text(encoding="utf-8"))
    codepoints = {ord(c) for c in chars}

    rules, union = [], set()
    for family, path, index, weight, style in specs:
        css, covered, size, weight_css = build_face(family, path, index, weight, style, codepoints, flavor)
        rules.append(css)
        union |= covered
        print(f"{family} [{weight_css} {style}] <- {path.name}: "
              f"{len(codepoints & covered)} chars, {size / 1024:.0f} KB {flavor}")

    missing = sorted(c for c in codepoints - union if not chr(c).isspace())
    if missing:
        sample = "".join(chr(c) for c in missing[:80])
        print(f"WARNING: {len(missing)} characters are not in any embedded font and will use a system "
              f"fallback: {sample}{' ...' if len(missing) > 80 else ''}", file=sys.stderr)

    block = f'<style id="{STYLE_ID}">\n' + "\n".join(rules) + "\n</style>\n"
    text = BLOCK_RE.sub("", text)
    m = re.search(r"</head>", text, re.I)
    text = text[: m.start()] + block + text[m.start():]
    out = args.output or args.html
    out.write_text(text, encoding="utf-8")
    print(f"wrote {out} ({out.stat().st_size / 1024:.0f} KB)")


if __name__ == "__main__":
    main()
