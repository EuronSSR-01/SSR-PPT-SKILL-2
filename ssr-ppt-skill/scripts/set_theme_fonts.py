#!/usr/bin/env python3
"""Set a PPTX theme's heading/body fonts for both Latin and East Asian (Simplified Chinese) text.

Generators often write only the Latin theme font (PptxGenJS `pptx.theme`) and leave the East Asian
font empty, so chart labels, new text boxes and any text without an explicit `a:ea` fall back to
DengXian / SimSun. This rewrites ppt/theme/theme*.xml in place (or to -o).

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
    ap.add_argument("-o", "--output", type=Path)
    args = ap.parse_args()

    out = args.output or args.pptx
    changed = []
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp) / "out.pptx"
        with zipfile.ZipFile(args.pptx) as zin, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if re.fullmatch(r"ppt/theme/theme\d+\.xml", item.filename):
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
