#!/usr/bin/env python3
"""Set a PPTX theme's heading/body fonts for both Latin and East Asian (Simplified Chinese) text.

Generators often write only the Latin theme font (PptxGenJS `pptx.theme`) and leave the East Asian
font empty, so text may use an unintended fallback. For a multi-theme deck, explicitly select
--theme (repeatable) or --all-themes. Do not apply global font changes to a company template
unless the task calls for them. Use -o to preserve the source.

Usage:
  python set_theme_fonts.py deck.pptx --heading "Noto Serif SC" --body "Noto Sans SC" [-o out.pptx]
  python set_theme_fonts.py deck.pptx --heading "Noto Serif SC" --heading-latin "Inter" --body "Noto Sans SC"
"""
import argparse
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.sax.saxutils import quoteattr


def set_block(xml, tag, latin, ea):
    m = re.search(rf"<a:{tag}>.*?</a:{tag}>", xml, re.S)
    if not m:
        return xml, False
    block = m.group(0)
    if latin:
        block = re.sub(r'<a:latin\b[^>]*/>', f"<a:latin typeface={quoteattr(latin)}/>", block, count=1)
    block = re.sub(r'<a:ea\b[^>]*/>', f"<a:ea typeface={quoteattr(ea)}/>", block, count=1)
    if re.search(r'<a:font script="Hans"', block):
        block = re.sub(r'<a:font script="Hans"[^>]*/>', f'<a:font script="Hans" typeface={quoteattr(ea)}/>', block)
    else:
        block = block.replace(f"</a:{tag}>", f'<a:font script="Hans" typeface={quoteattr(ea)}/></a:{tag}>')
    return xml[: m.start()] + block + xml[m.end():], True


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pptx", type=Path)
    ap.add_argument("--heading", required=True, help="East Asian heading font (also Latin unless --heading-latin)")
    ap.add_argument("--body", required=True, help="East Asian body font (also Latin unless --body-latin)")
    ap.add_argument("--heading-latin")
    ap.add_argument("--body-latin")
    scope = ap.add_mutually_exclusive_group()
    scope.add_argument("--theme", action="append", help="Exact theme part, e.g. ppt/theme/theme1.xml; repeatable")
    scope.add_argument("--all-themes", action="store_true", help="Explicitly change every theme")
    ap.add_argument("-o", "--output", type=Path)
    args = ap.parse_args()

    out = args.output or args.pptx
    changed = []
    with zipfile.ZipFile(args.pptx) as zin:
        themes = {n for n in zin.namelist() if re.fullmatch(r"ppt/theme/theme\d+\.xml", n)}
    if not themes:
        sys.exit("no theme found; file left unchanged")
    if args.theme:
        selected = set(args.theme)
        if selected - themes:
            sys.exit(f"unknown theme parts: {sorted(selected - themes)}; file left unchanged")
    elif args.all_themes or len(themes) == 1:
        selected = themes
    else:
        sys.exit("multiple themes: select --theme or explicitly use --all-themes; file left unchanged")
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp) / "out.pptx"
        with zipfile.ZipFile(args.pptx) as zin, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename in selected:
                    xml = data.decode("utf-8")
                    xml, a = set_block(xml, "majorFont", args.heading_latin or args.heading, args.heading)
                    xml, b = set_block(xml, "minorFont", args.body_latin or args.body, args.body)
                    if a or b:
                        changed.append(item.filename)
                    data = xml.encode("utf-8")
                zout.writestr(item, data)
        if not changed:
            sys.exit("no theme font scheme found; file left unchanged")
        shutil.copyfile(tmp_path, out)
    print(f"updated {', '.join(changed)} -> {out}")


if __name__ == "__main__":
    main()
