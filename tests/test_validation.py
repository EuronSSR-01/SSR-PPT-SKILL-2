"""Regression tests for observable edits and template preservation. No Office UI claim."""
import copy
import importlib.util
from io import BytesIO
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zipfile
from lxml import etree as ET
from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "ssr-ppt-skill" / "scripts"
spec = importlib.util.spec_from_file_location("preview", SCRIPTS / "preview.py")
preview = importlib.util.module_from_spec(spec)
spec.loader.exec_module(preview)


def add_chart(slide, kind=XL_CHART_TYPE.COLUMN_CLUSTERED):
    data = CategoryChartData()
    data.categories = ["第一组", "第二组", "第三组"]
    data.add_series("已完成", [6, 4, 3])
    return slide.shapes.add_chart(kind, Inches(1), Inches(2), Inches(6), Inches(3), data)


def multi_master_fixture(path):
    """Two linked masters/themes, notes and hyperlinks, generated without customer data."""
    prs = Presentation()
    for name in ("保留标题", "第二主题"):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
        run = box.text_frame.paragraphs[0].add_run()
        run.text = name
        run.hyperlink.address = "https://example.com/fixture"
        slide.notes_slide.notes_text_frame.text = "备注：固定口径与来源"
    prs.save(path)
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    pns = "http://schemas.openxmlformats.org/presentationml/2006/main"
    rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    relns = "http://schemas.openxmlformats.org/package/2006/relationships"
    ctns = "http://schemas.openxmlformats.org/package/2006/content-types"
    def root(name):
        return ET.fromstring(parts[name])
    def save(name, element):
        parts[name] = ET.tostring(element, xml_declaration=True, encoding="UTF-8", standalone=True)
    # Clone the blank layout and its master, keeping a single valid layout relationship.
    master = root("ppt/slideMasters/slideMaster1.xml")
    rels = root("ppt/slideMasters/_rels/slideMaster1.xml.rels")
    keep = next(r.get("Id") for r in rels if r.get("Target") == "../slideLayouts/slideLayout7.xml")
    for rel in list(rels):
        if rel.get("Type").endswith("/slideLayout"):
            if rel.get("Id") == keep:
                rel.set("Target", "../slideLayouts/slideLayout12.xml")
            else:
                rels.remove(rel)
        elif rel.get("Type").endswith("/theme"):
            rel.set("Target", "../theme/theme2.xml")
    layouts = master.find(f"{{{pns}}}sldLayoutIdLst")
    for item in list(layouts):
        if item.get(f"{{{rns}}}id") != keep:
            layouts.remove(item)
    save("ppt/slideMasters/slideMaster2.xml", master)
    save("ppt/slideMasters/_rels/slideMaster2.xml.rels", rels)
    parts["ppt/theme/theme2.xml"] = parts["ppt/theme/theme1.xml"].replace(b'Calibri', b'Arial')
    parts["ppt/slideLayouts/slideLayout12.xml"] = parts["ppt/slideLayouts/slideLayout7.xml"]
    parts["ppt/slideLayouts/_rels/slideLayout12.xml.rels"] = parts["ppt/slideLayouts/_rels/slideLayout7.xml.rels"].replace(b'slideMaster1.xml', b'slideMaster2.xml')
    parts["ppt/slides/_rels/slide2.xml.rels"] = parts["ppt/slides/_rels/slide2.xml.rels"].replace(b'slideLayout7.xml', b'slideLayout12.xml')
    rels = root("ppt/_rels/presentation.xml.rels")
    ET.SubElement(rels, f"{{{relns}}}Relationship", Id="rIdTemplate2", Type=rns+"/slideMaster", Target="slideMasters/slideMaster2.xml")
    save("ppt/_rels/presentation.xml.rels", rels)
    presentation = root("ppt/presentation.xml")
    item = ET.SubElement(presentation.find(f"{{{pns}}}sldMasterIdLst"), f"{{{pns}}}sldMasterId", id="2147483649")
    item.set(f"{{{rns}}}id", "rIdTemplate2")
    save("ppt/presentation.xml", presentation)
    types = root("[Content_Types].xml")
    for part, suffix in (("slideMasters/slideMaster2.xml", "presentationml.slideMaster+xml"), ("slideLayouts/slideLayout12.xml", "presentationml.slideLayout+xml"), ("theme/theme2.xml", "theme+xml")):
        ET.SubElement(types, f"{{{ctns}}}Override", PartName="/ppt/"+part, ContentType="application/vnd.openxmlformats-officedocument."+suffix)
    save("[Content_Types].xml", types)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)


class EditTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.work = Path(self.tmp.name)
        self.src = self.work / "source.pptx"

    def fixture(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_textbox(0, 0, Inches(5), Inches(1)).text = "页面标题"
        group = slide.shapes.add_group_shape()
        box = group.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1))
        box.text = "组内关键数值 12"
        node = group.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4), Inches(1), Inches(1), Inches(1))
        line = group.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(3), Inches(1), Inches(4), Inches(1))
        line.begin_connect(box, 3)
        line.end_connect(node, 1)
        table = slide.shapes.add_table(2, 2, 0, Inches(3), Inches(3), Inches(2))
        table.table.cell(0, 0).text = "原先已有 ✓"
        image = self.work / "picture.png"
        Image.new("RGB", (100, 60), "#256360").save(image)
        slide.shapes.add_picture(str(image), Inches(6), Inches(1), width=Inches(2))
        prs.save(self.src)
        return group.shape_id, box.shape_id, table.shape_id

    def test_group_table_picture_roundtrip_and_source_unchanged(self):
        gid, tid, tableid = self.fixture()
        original = self.src.read_bytes()
        result = preview.edit_test(self.src, self.work / "out")
        self.assertTrue(result["checks"])
        self.assertTrue(all(c["status"] == "passed" for c in result["checks"]))
        self.assertTrue(any(c["target"]["path"] == [gid, tid] for c in result["checks"]))
        self.assertEqual(self.src.read_bytes(), original)
        # Existing checkmarks and unchanged group text must not mask a failed edit.
        reloaded = Presentation(result["copy"])
        for c in result["checks"]:
            if c["target"]["kind"] == "text" and c["target"]["path"] == [gid, tid]:
                preview.locate_shape(reloaded, c["target"]).text = c["before"]
            if c["target"]["path"] == [tableid]:
                preview.locate_shape(reloaded, c["target"]).table.cell(0, 0).text = c["before"]
        checked = preview.verify_edits(reloaded, copy.deepcopy(result["checks"]))
        self.assertEqual(sum(c["status"] == "failed" for c in checked), 2)

    def test_supported_chart_values_and_workbook(self):
        for kind in (XL_CHART_TYPE.COLUMN_CLUSTERED, XL_CHART_TYPE.BAR_CLUSTERED, XL_CHART_TYPE.LINE):
            with self.subTest(kind=kind):
                prs = Presentation()
                shape = add_chart(prs.slides.add_slide(prs.slide_layouts[6]), kind)
                prs.save(self.src)
                target = {"slide": 1, "path": [shape.shape_id], "kind": "chart"}
                result = preview.edit_test(self.src, self.work / str(kind.value), targets=[target])
                record = result["checks"][0]
                self.assertEqual(record["status"], "passed", record)
                self.assertEqual(record["actual"]["series"][0]["values"][0], 7.6)
                self.assertEqual(record["actual"]["workbook"]["B2"], 7.6)
                unchanged = preview.verify_edits(Presentation(self.src), copy.deepcopy(result["checks"]))
                self.assertEqual(unchanged[0]["status"], "failed")
        # A cached number can hide a formula; do not silently replace its workbook.
        with zipfile.ZipFile(BytesIO(shape.chart.part.chart_workbook.xlsx_part.blob)) as source:
            parts = {n: source.read(n) for n in source.namelist()}
        ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        sheet = ET.fromstring(parts["xl/worksheets/sheet1.xml"])
        cell = sheet.find(f".//{{{ns}}}c[@r='B2']")
        ET.SubElement(cell, f"{{{ns}}}f").text = "2+4"
        parts["xl/worksheets/sheet1.xml"] = ET.tostring(sheet)
        blob = BytesIO()
        with zipfile.ZipFile(blob, "w") as output:
            for name, data in parts.items():
                output.writestr(name, data)
        with self.assertRaises(NotImplementedError):
            preview.workbook_blob_cells(blob.getvalue())

    def test_empty_workbook_value_cell_does_not_crash(self):
        # PptxGenJS writes empty <v/> cells in embedded workbooks; they must read as empty, not crash.
        prs = Presentation()
        shape = add_chart(prs.slides.add_slide(prs.slide_layouts[6]))
        with zipfile.ZipFile(BytesIO(shape.chart.part.chart_workbook.xlsx_part.blob)) as source:
            parts = {n: source.read(n) for n in source.namelist()}
        ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        sheet = ET.fromstring(parts["xl/worksheets/sheet1.xml"])
        row = sheet.find(f".//{{{ns}}}row")
        ET.SubElement(ET.SubElement(row, f"{{{ns}}}c", r="Z1"), f"{{{ns}}}v")
        parts["xl/worksheets/sheet1.xml"] = ET.tostring(sheet)
        blob = BytesIO()
        with zipfile.ZipFile(blob, "w") as output:
            for name, data in parts.items():
                output.writestr(name, data)
        cells = preview.workbook_blob_cells(blob.getvalue())
        self.assertNotIn("Z1", cells)
        self.assertEqual(cells["B2"], 6)

    def test_unsupported_chart_is_not_passed(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        data = XyChartData()
        data.add_series("XY").add_data_point(1, 2)
        slide.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER, 0, 0, Inches(5), Inches(3), data)
        prs.save(self.src)
        result = preview.edit_test(self.src, self.work / "out")
        self.assertEqual(result["checks"][0]["status"], "unsupported")
        self.assertEqual(preview.checks_status(result["checks"]), "unsupported")

    def test_explicit_target_and_invalid_selector(self):
        gid, tid, _ = self.fixture()
        target = {"slide": 1, "path": [gid, tid], "kind": "text"}
        result = preview.edit_test(self.src, self.work / "out", targets=[target])
        self.assertEqual(len(result["checks"]), 1)
        self.assertEqual(result["checks"][0]["status"], "passed")
        target["slide"] = 0
        result = preview.edit_test(self.src, self.work / "bad", targets=[target])
        self.assertEqual(result["checks"][0]["status"], "failed")

    def test_strict_statuses(self):
        for status in ("failed", "not_run", "unsupported"):
            with self.subTest(status=status), self.assertRaises(SystemExit):
                preview.write_report({"checks": [preview.check("required", status)]}, self.work, True)
        report = {"checks": [preview.check("required", "passed"), preview.check("unused", "not_applicable", required=False)]}
        preview.write_report(report, self.work, True)
        self.assertEqual(report["automatic_status"], "passed")
        self.assertEqual(report["manual_review"]["status"], "not_run")

    def test_multi_theme_inventory_and_scoped_font_change(self):
        multi_master_fixture(self.src)
        original = self.src.read_bytes()
        themes = preview.font_inventory(self.src)["themes"]
        self.assertEqual(len(themes), 2)
        out = self.work / "font.pptx"
        args = [sys.executable, str(SCRIPTS / "set_theme_fonts.py"), str(self.src), "--heading", "Test Heading", "--body", "Test Body", "-o", str(out)]
        result = subprocess.run(args, capture_output=True)
        self.assertNotEqual(result.returncode, 0)
        self.assertFalse(out.exists())
        result = subprocess.run(args+["--theme", "ppt/theme/theme1.xml"], capture_output=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        with zipfile.ZipFile(self.src) as a, zipfile.ZipFile(out) as b:
            self.assertNotEqual(a.read("ppt/theme/theme1.xml"), b.read("ppt/theme/theme1.xml"))
            for name in a.namelist():
                if name != "ppt/theme/theme1.xml":
                    self.assertEqual(a.read(name), b.read(name), name)
        self.assertEqual(self.src.read_bytes(), original)

    def test_template_roundtrip_preserves_other_parts(self):
        multi_master_fixture(self.src)
        prs = Presentation(self.src)
        self.assertEqual(len(prs.slide_masters), 2)
        self.assertEqual(prs.slides[1].slide_layout.part.partname, "/ppt/slideLayouts/slideLayout12.xml")
        prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].text = "只改这句话"
        out = self.work / "roundtrip.pptx"
        prs.save(out)
        reloaded = Presentation(out)
        self.assertEqual(reloaded.slides[0].shapes[0].text, "只改这句话")
        self.assertEqual(reloaded.slides[1].shapes[0].text, "第二主题")
        for slide in reloaded.slides:
            self.assertIn("固定口径", slide.notes_slide.notes_text_frame.text)
            self.assertEqual(slide.shapes[0].text_frame.paragraphs[0].runs[0].hyperlink.address, "https://example.com/fixture")
        with zipfile.ZipFile(self.src) as a, zipfile.ZipFile(out) as b:
            for name in a.namelist():
                if name.startswith(("ppt/theme/", "ppt/slideMasters/", "ppt/slideLayouts/", "ppt/notesSlides/")) or name == "ppt/slides/slide2.xml":
                    self.assertEqual(a.read(name), b.read(name), name)


class KitTests(unittest.TestCase):
    def test_pptx_kit_defaults_and_editability(self):
        sys.path.insert(0, str(SCRIPTS))
        try:
            from pptx_kit import Deck
        finally:
            sys.path.remove(str(SCRIPTS))
        with tempfile.TemporaryDirectory() as tmp:
            work = Path(tmp)
            Image.new("RGB", (1600, 600), "#256360").save(work / "wide.png")
            deck = Deck()
            s = deck.slide(bg="FBF9F4")
            deck.text(s, 80, 80, 800, 60, [[("改变", {"color": "E8590C"}), ("城市零售", {})]], size=44, font="Test Serif")
            chart = deck.bar_chart(s, 80, 200, 700, 300, ["甲", "乙", "丙"], [5, 3, 2], highlight={0}, font="Test Sans")
            table = deck.table(s, 80, 520, 600, [["城市", "数值"], ["上海", "86.4"]], align=["left", "right"])
            pic = deck.picture(s, work / "wide.png", 800, 200, 300, 300)
            box = deck.rect(s, 800, 520, 100, 50, fill="E8590C")
            src = deck.save(work / "kit.pptx", theme_fonts=("Test Serif", "Test Sans"))

            prs = Presentation(src)
            slide = prs.slides[0]
            self.assertEqual((prs.slide_width, prs.slide_height), (12192000, 6858000))
            runs = [r for sh in slide.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs for r in p.runs]
            self.assertTrue(runs and all(r._r.find(".//" + qn_ea) is not None for r in runs))
            plot = [sh for sh in slide.shapes if sh.has_chart][0].chart.plots[0]
            self.assertEqual(len(plot.series), 1)
            self.assertEqual(list(plot.series[0].values), [5, 3, 2])  # data untouched by the highlight
            self.assertIn("2D5ABB26", ET.tostring([sh for sh in slide.shapes if sh.has_table][0]._element).decode())
            # 1600x600 into a square keeps the centre 600/1600 of the width: each side loses (1 - 0.375) / 2
            self.assertAlmostEqual(next(sh for sh in slide.shapes if sh.shape_id == pic.shape_id).crop_left, 0.3125, places=4)
            self.assertIsNotNone(next(sh for sh in slide.shapes if sh.shape_id == box.shape_id)._element.find(".//" + qn_effect))
            with zipfile.ZipFile(src) as z:
                self.assertIn(b'<a:ea typeface="Test Sans"/>', z.read("ppt/theme/theme1.xml"))
            result = preview.edit_test(src, work / "edit", targets=[
                {"slide": 1, "path": [chart.shape_id], "kind": "chart"},
                {"slide": 1, "path": [table.shape_id], "kind": "table", "cell": [1, 1]}])
            self.assertEqual([c["status"] for c in result["checks"]], ["passed", "passed"])


qn_ea = "{http://schemas.openxmlformats.org/drawingml/2006/main}ea"
qn_effect = "{http://schemas.openxmlformats.org/drawingml/2006/main}effectLst"


def soffice_or_none():
    try:
        return preview.find_soffice(None)
    except SystemExit:
        return None


@unittest.skipUnless(soffice_or_none() and preview.pymupdf_or_none(), "LibreOffice and PyMuPDF are needed to render")
class PptxLayoutTests(unittest.TestCase):
    def test_layout_lint_on_rendered_pptx(self):
        def box(slide, x, y, w, text, size, wrap=True):
            tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(1))
            tb.text_frame.word_wrap = wrap
            run = tb.text_frame.paragraphs[0].add_run()
            run.text, run.font.size = text, Pt(size)
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
        bad = prs.slides.add_slide(prs.slide_layouts[6])
        box(bad, 0.8, 0.6, 8, "重叠的第一段标题文字", 32)
        box(bad, 1.1, 0.75, 8, "重叠的第二段标题文字", 32)
        box(bad, 11.5, 2.5, 4, "这一行文字跑出了页面右边", 24, wrap=False)
        box(bad, 0.8, 4.0, 2.2, "一二三四五六七八九十十", 28)
        box(bad, 6.0, 4.0, 4, "字号太小的说明文字", 7)
        box(bad, 6.0, 5.0, 4, "TODO 补充数据", 18)
        clean = prs.slides.add_slide(prs.slide_layouts[6])
        box(clean, 0.8, 0.6, 10, "干净的页面", 36)
        box(clean, 0.8, 2.0, 10, "正文不重叠，也不跑出页面。", 20)
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "layout.pptx"
            prs.save(src)
            layout = []
            preview.pptx_to_pngs(src, Path(tmp) / "out", soffice_or_none(), 800, layout)
            kinds = {f["kind"] for f in layout[0]["findings"]}
            for kind in ("text_overlap", "text_outside_slide", "orphan_last_line", "font_too_small", "placeholder"):
                self.assertIn(kind, kinds)
            self.assertEqual(layout[1]["findings"], [])


if __name__ == "__main__":
    unittest.main()
