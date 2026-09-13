#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
korean_pptx.py
python-pptx 기반 프레젠테이션 제작 시 사용하는 한국어 전용 서식 헬퍼 모듈.
"""

from pptx.oxml.xmlchemy import OxmlElement
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
POST_EA_TAGS = [f"{{{A_NS}}}{t}" for t in ["cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"]]


def _insert_ea_to_prop(prop_elm, font_name):
    """prop_elm (rPr 또는 defRPr) 내부에 OpenXML 스키마 순서를 준수하여 <a:ea> 요소를 삽입합니다."""
    ea = prop_elm.find(f"{{{A_NS}}}ea")
    if ea is not None:
        ea.set("typeface", font_name)
        return

    ea_elem = OxmlElement("a:ea")
    ea_elem.set("typeface", font_name)

    latin = prop_elm.find(f"{{{A_NS}}}latin")
    if latin is not None:
        idx = list(prop_elm).index(latin)
        prop_elm.insert(idx + 1, ea_elem)
        return

    for i, child in enumerate(prop_elm):
        if child.tag in POST_EA_TAGS:
            prop_elm.insert(i, ea_elem)
            return

    prop_elm.append(ea_elem)


def set_korean_font_and_lang(target, font_name="Pretendard"):
    """
    Run 또는 Paragraph 객체에 언어(ko-KR)와 폰트(latin + ea: East Asian)를 동시 설정.
    한글이 맑은 고딕 등 시스템 기본 폰트로 강제 폴백되는 현상을 방지합니다.
    """
    try:
        target.font.language_id = MSO_LANGUAGE_ID.KOREAN
    except Exception:
        pass

    target.font.name = font_name

    parent_elm = getattr(target, "_r", None)
    if parent_elm is None:
        parent_elm = getattr(target, "_p", None)
    if parent_elm is None:
        return

    prop_elm = (
        parent_elm.get_or_add_rPr()
        if hasattr(parent_elm, "get_or_add_rPr")
        else parent_elm.get_or_add_pPr().get_or_add_defRPr()
    )

    _insert_ea_to_prop(prop_elm, font_name)


def set_word_break_prevention(paragraph):
    """
    문단에 단어 중간 줄바꿈 방지 속성 (eaLnBrk="0", latinLnBrk="0") 설정.
    한글/영문 단어가 음절 단위로 쪼개지는 현상을 방지합니다.
    """
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("eaLnBrk", "0")
    pPr.set("latinLnBrk", "0")


def set_native_bullet(paragraph, char="•", mar_l=Inches(0.3), indent=-Inches(0.2)):
    """
    OpenXML ISO/IEC 29500-1 스키마 순서를 엄격 준수하여 네이티브 불릿을 적용.
    buChar가 defRPr보다 앞에 오도록 보장하여 PowerPoint 파일 복구 경고를 방지합니다.
    """
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("marL", str(int(mar_l)))
    pPr.set("indent", str(int(indent)))

    # 1. 기존 불릿 관련 태그 제거 (중복 방지)
    ns = f"{{{A_NS}}}"
    for tag_name in ["buChar", "buNone", "buAutoNum", "buBlip", "buClrTx", "buSzTx"]:
        for child in pPr.findall(f"{ns}{tag_name}"):
            pPr.remove(child)

    # 2. 불릿 요소 생성: 텍스트 색상 상속(buClrTx), 크기 상속(buSzTx), 불릿 문자(buChar)
    bu_clr_tx = OxmlElement("a:buClrTx")
    bu_sz_tx = OxmlElement("a:buSzTx")
    bu_char = OxmlElement("a:buChar")
    bu_char.set("char", char)

    # 3. 스키마 순서 준수: [buClrTx, buSzTx, buChar] -> defRPr
    defRPr = pPr.find(f"{ns}defRPr")
    if defRPr is not None:
        idx = list(pPr).index(defRPr)
        pPr.insert(idx, bu_clr_tx)
        pPr.insert(idx + 1, bu_sz_tx)
        pPr.insert(idx + 2, bu_char)
    else:
        pPr.append(bu_clr_tx)
        pPr.append(bu_sz_tx)
        pPr.append(bu_char)


def create_korean_card(
    slide,
    left,
    top,
    width,
    height,
    bg_color=RGBColor(245, 247, 250),
    border_color=RGBColor(220, 225, 230),
    line_width_pt=1,
    margin_inch=0.3,
):
    """
    배경 도형과 텍스트 박스를 일체화한 단일 카드 도형 생성.
    사용자 드래그 시 분리되지 않고 편집성이 우수합니다.
    """
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    shape.line.color.rgb = border_color
    shape.line.width = Pt(line_width_pt)

    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = Inches(margin_inch)
    tf.margin_right = Inches(margin_inch)
    tf.margin_top = Inches(margin_inch)
    tf.margin_bottom = Inches(margin_inch)
    return shape


def _extract_text_frames(container):
    """
    Presentation, Slide, Shape, Table, GroupShape 등 다양한 객체에서
    모든 TextFrame을 재귀적으로 수집합니다.
    """
    tfs = []
    # 1. Presentation 객체 (slides 컬렉션 순회)
    if hasattr(container, "slides"):
        for slide in container.slides:
            tfs.extend(_extract_text_frames(slide))
        return tfs

    # 2. Slide 또는 GroupShape의 shapes 컬렉션
    shapes = []
    if hasattr(container, "shapes"):
        shapes = list(container.shapes)
    elif isinstance(container, (list, tuple)):
        shapes = list(container)
    else:
        shapes = [container]

    for shape in shapes:
        # 이미 TextFrame인 경우
        if hasattr(shape, "paragraphs") and hasattr(shape, "word_wrap"):
            tfs.append(shape)
            continue

        # GroupShape: 내부에 shapes 컬렉션이 있음
        if hasattr(shape, "shapes"):
            tfs.extend(_extract_text_frames(shape))
            continue

        # Table 객체
        if getattr(shape, "has_table", False):
            for cell in shape.table.iter_cells():
                if getattr(cell, "text_frame", None):
                    tfs.append(cell.text_frame)
            continue

        # 일반 도형의 TextFrame
        if getattr(shape, "has_text_frame", False) and shape.text_frame:
            tfs.append(shape.text_frame)

    return tfs


def apply_korean_formatting(target_obj, font_name="Pretendard"):
    """
    프레젠테이션(Presentation), 슬라이드(Slide), 도형(Shape), 표(Table),
    또는 텍스트 프레임(TextFrame)을 전달받아 내부의 모든 텍스트에
    언어 태그(ko-KR), 단어 줄바꿈 방지(eaLnBrk/latinLnBrk="0"), 동아시아 폰트를 일괄 적용합니다.
    표(Table) 및 그룹 도형 내부 텍스트까지 재귀적으로 완벽히 지원합니다.
    """
    tfs = _extract_text_frames(target_obj)
    for tf in tfs:
        tf.word_wrap = True
        for p in tf.paragraphs:
            set_word_break_prevention(p)
            if p.runs:
                for r in p.runs:
                    set_korean_font_and_lang(r, font_name=font_name)
            else:
                set_korean_font_and_lang(p, font_name=font_name)
