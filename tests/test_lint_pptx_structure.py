# -*- coding: utf-8 -*-
import contextlib
import io
import os
import sys
import tempfile
import unittest
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches

SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if SKILL_ROOT not in sys.path:
    sys.path.insert(0, SKILL_ROOT)

from scripts.korean_pptx import set_native_bullet
from scripts.lint_pptx_structure import (
    detect_fake_table_rows,
    detect_split_bullets,
    main as lint_main,
)


class TestLintPptxStructure(unittest.TestCase):
    def _run_lint_cli(self, pptx_path):
        old_argv = sys.argv
        buf = io.StringIO()
        try:
            sys.argv = ["lint_pptx_structure.py", pptx_path]
            with contextlib.redirect_stdout(buf):
                with self.assertRaises(SystemExit) as cm:
                    lint_main()
            return cm.exception.code, buf.getvalue()
        finally:
            sys.argv = old_argv

    def test_native_table_and_single_bullet_box_pass(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        table = slide.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(8), Inches(1.5)).table
        table.cell(0, 0).text = "구분"
        table.cell(0, 1).text = "내용"

        tb = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(8), Inches(2))
        p1 = tb.text_frame.paragraphs[0]
        p1.text = "첫 번째 항목"
        set_native_bullet(p1)
        p2 = tb.text_frame.add_paragraph()
        p2.text = "두 번째 항목"
        set_native_bullet(p2)

        self.assertEqual(detect_fake_table_rows(slide), [])
        self.assertEqual(detect_split_bullets(slide), [])

        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, "clean.pptx")
            prs.save(path)
            code, out = self._run_lint_cli(path)
        self.assertEqual(code, 0)
        self.assertIn("결함 0건", out)

    def test_detects_fake_table_rows_and_returns_exit_code_2(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        for row_idx in range(2):
            top = Inches(1.0 + row_idx * 0.6)
            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), top, Inches(9.0), Inches(0.45))
            bg.fill.solid()
            bg.fill.fore_color.rgb = RGBColor(240, 240, 240)

            c1 = slide.shapes.add_textbox(Inches(0.6), top, Inches(2.0), Inches(0.4))
            c1.text_frame.text = f"항목 {row_idx + 1}"
            c2 = slide.shapes.add_textbox(Inches(3.0), top, Inches(5.0), Inches(0.4))
            c2.text_frame.text = f"설명 {row_idx + 1}"

        fake_rows = detect_fake_table_rows(slide)
        self.assertEqual(len(fake_rows), 2)

        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, "fake_table.pptx")
            prs.save(path)
            code, out = self._run_lint_cli(path)
        self.assertEqual(code, 2)
        self.assertIn("가짜 표 의심", out)

    def test_detects_split_bullets_and_returns_exit_code_2(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        for idx in range(2):
            tb = slide.shapes.add_textbox(Inches(1.0), Inches(1.0 + idx * 0.4), Inches(6.0), Inches(0.35))
            p = tb.text_frame.paragraphs[0]
            p.text = f"분리된 불릿 {idx + 1}"
            set_native_bullet(p)

        clusters = detect_split_bullets(slide)
        self.assertEqual(len(clusters), 1)
        self.assertEqual(clusters[0][1], 2)

        with tempfile.TemporaryDirectory() as tmpdir:
            path = os.path.join(tmpdir, "split_bullets.pptx")
            prs.save(path)
            code, out = self._run_lint_cli(path)
        self.assertEqual(code, 2)
        self.assertIn("분할 불릿 의심", out)


if __name__ == "__main__":
    unittest.main()
