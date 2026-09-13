# -*- coding: utf-8 -*-
import unittest
import os
import sys
import tempfile
import xml.etree.ElementTree as ET
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE

SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if SKILL_ROOT not in sys.path:
    sys.path.insert(0, SKILL_ROOT)

from scripts.korean_pptx import (
    create_korean_card,
    set_native_bullet,
    set_korean_font_and_lang,
    set_word_break_prevention,
    apply_korean_formatting
)
from scripts.fix_korean_pptx import (
    patch_xml_content,
    verify_xml_content,
    process_pptx,
    get_slide_file_map,
    qualify,
    A_NS
)


class TestKoreanPptxHelper(unittest.TestCase):
    def setUp(self):
        self.prs = Presentation()
        self.slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])

    def test_create_korean_card(self):
        card = create_korean_card(self.slide, Inches(1), Inches(1), Inches(4), Inches(2))
        self.assertTrue(card.has_text_frame)
        self.assertTrue(card.text_frame.word_wrap)
        self.assertEqual(card.text_frame.margin_left, Inches(0.3))

    def test_set_native_bullet(self):
        tb = self.slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        p = tb.text_frame.paragraphs[0]
        p.text = "불릿 항목입니다."
        set_native_bullet(p, char="➔")

        pPr = p._p.get_or_add_pPr()
        tags = [c.tag.split('}')[-1] for c in pPr]
        self.assertIn("buChar", tags)
        if "defRPr" in tags:
            self.assertLess(tags.index("buChar"), tags.index("defRPr"))

    def test_apply_korean_formatting_table(self):
        tbl_shape = self.slide.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(4), Inches(2))
        cell = tbl_shape.table.cell(0, 0)
        cell.text = "표 내부 한국어 텍스트"

        apply_korean_formatting(self.slide, font_name="Pretendard")

        p = cell.text_frame.paragraphs[0]
        self.assertEqual(p._p.get_or_add_pPr().get("eaLnBrk"), "0")
        self.assertEqual(p._p.get_or_add_pPr().get("latinLnBrk"), "0")
        ea = p.runs[0]._r.find(".//" + qualify(A_NS, "ea"))
        self.assertIsNotNone(ea)
        self.assertEqual(ea.get("typeface"), "Pretendard")

    def test_apply_korean_formatting_presentation(self):
        tb = self.slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1))
        tb.text_frame.text = "전체 덱 적용 테스트"

        apply_korean_formatting(self.prs, font_name="Pretendard")
        p = tb.text_frame.paragraphs[0]
        self.assertEqual(p._p.get_or_add_pPr().get("eaLnBrk"), "0")


class TestFixKoreanPptxCli(unittest.TestCase):
    def test_patch_xml_missing_pPr(self):
        xml_str = """<p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:txBody>
    <a:bodyPr/>
    <a:p>
      <a:r>
        <a:t>pPr이 없는 문단</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>"""
        mod, new_xml, stats = patch_xml_content(xml_str.encode("utf-8"))
        self.assertTrue(mod)
        root = ET.fromstring(new_xml)
        p = root.find('.//' + qualify(A_NS, 'p'))
        pPr = p.find(qualify(A_NS, 'pPr'))
        self.assertIsNotNone(pPr)
        self.assertEqual(pPr.get("eaLnBrk"), "0")
        self.assertEqual(pPr.get("latinLnBrk"), "0")

    def test_patch_xml_schema_order_with_hyperlink(self):
        xml_str = """<p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:txBody>
    <a:bodyPr/>
    <a:p>
      <a:pPr/>
      <a:r>
        <a:rPr>
          <a:latin typeface="Pretendard"/>
          <a:hlinkClick r:id="rId2"/>
        </a:rPr>
        <a:t>링크 텍스트</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>"""
        mod, new_xml, stats = patch_xml_content(xml_str.encode("utf-8"))
        self.assertTrue(mod)
        root = ET.fromstring(new_xml)
        rPr = root.find('.//' + qualify(A_NS, 'rPr'))
        tags = [c.tag.split('}')[-1] for c in rPr]
        self.assertEqual(tags, ['latin', 'ea', 'hlinkClick'])

    def test_full_pipeline_process_pptx(self):
        temp_dir = tempfile.mkdtemp()
        pptx_path = os.path.join(temp_dir, 'pipeline_test.pptx')

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
        tb.text_frame.text = "파이프라인 테스트"
        prs.save(pptx_path)

        # 1. 보정 전 verify는 미흡 항목 발견 (코드 2)
        code_before = process_pptx(pptx_path, verify_only=True)
        self.assertEqual(code_before, 2)

        # 2. 보정 실행 (코드 0)
        code_fix = process_pptx(pptx_path)
        self.assertEqual(code_fix, 0)

        # 3. 보정 후 verify는 완벽 적용 (코드 0)
        code_after = process_pptx(pptx_path, verify_only=True)
        self.assertEqual(code_after, 0)


if __name__ == '__main__':
    unittest.main()
