#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fix_korean_pptx.py
PPTX 파일의 한국어 서식(언어 태그, 단어 단위 줄바꿈, 동아시아 폰트, 마스터/테마 스타일)을 일괄 보정하는 고속 CLI 도구.
python-pptx와 함께 설치되는 lxml만 사용하며, 네임스페이스 접두사(mc:Ignorable, Requires 등)를 원본 그대로 보존합니다.
"""

import argparse
import os
import posixpath
import re
import shutil
import sys
import tempfile
import zipfile

try:
    from lxml import etree
except ImportError as e:  # pragma: no cover
    raise ImportError("lxml이 필요합니다: pip install lxml (python-pptx 설치 시 함께 설치됩니다)") from e

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

# xml.etree는 직렬화 시 접두사를 ns0, ns1...로 바꿔 mc:Ignorable="p14" 같은
# 접두사 문자열 참조를 깨뜨리므로(PowerPoint 복구 경고) lxml을 사용합니다.
_XML_PARSER = etree.XMLParser(resolve_entities=False, huge_tree=True)

# ISO/IEC 29500-1 CT_TextCharacterProperties 자식 요소 순서:
# ln -> Fill -> Effect -> highlight -> uLn -> uFill -> latin -> ea -> cs -> sym -> hlinkClick -> hlinkMouseOver -> rtl -> extLst
POST_LATIN_TAGS = [
    f"{{{A_NS}}}{t}"
    for t in ["ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"]
]
POST_EA_TAGS = [
    f"{{{A_NS}}}{t}"
    for t in ["cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"]
]

# CT_TextParagraphProperties 타입을 가지는 마스터/레이아웃/도형 스타일 문단 프로퍼티 태그들
LEVEL_PPR_TAGS = {
    f"{{{A_NS}}}defPPr",
    *(f"{{{A_NS}}}lvl{i}pPr" for i in range(1, 10)),
}

# 테마 폰트 참조: latin용 참조를 ea에 그대로 복사하면 안 되고 ea용 참조로 대응시켜야 합니다.
THEME_LATIN_TO_EA = {"+mn-lt": "+mn-ea", "+mj-lt": "+mj-ea"}
THEME_EA_REFS = {"+mn-ea", "+mj-ea"}
THEME_LATIN_REFS = {"+mn-lt", "+mj-lt"}

# 아이콘/심볼 전용 폰트 (latin 폰트를 한글 폰트로 덮어쓰면 안 되는 예외 폰트)
SYMBOL_FONTS = {
    "wingdings", "wingdings 2", "wingdings 3", "webdings", "symbol",
    "segoe ui symbol", "segoe mdl2 assets", "segoe fluent icons", "marlett",
}

# 한글 글리프를 가진 것으로 알려진 폰트 이름 단서 (소문자 부분 일치)
KOREAN_FONT_HINTS = (
    "pretendard", "malgun", "nanum", "gulim", "dotum", "batang", "gungsuh",
    "spoqa", "apple sd gothic", "applegothic", "kopub", "noto sans cjk", "noto serif cjk",
    "wanted sans", "suit", "gmarket", "line seed", "ibm plex sans kr", "source han",
)
HANGUL_RE = re.compile("[ᄀ-ᇿ㄰-㆏가-힣]")


def qualify(ns, tag):
    return f"{{{ns}}}{tag}"


def _parse(xml_bytes):
    return etree.fromstring(xml_bytes, _XML_PARSER)


def is_korean_font(typeface):
    """폰트 이름만으로 한글 지원 여부를 추정합니다 (한글 이름, 'KR' 접미사, 알려진 폰트)."""
    if not typeface:
        return False
    if HANGUL_RE.search(typeface):
        return True
    lower = typeface.lower().strip()
    return lower.endswith(" kr") or any(hint in lower for hint in KOREAN_FONT_HINTS)


def is_symbol_font(typeface):
    if not typeface:
        return False
    return typeface.lower().strip() in SYMBOL_FONTS


def resolve_ea_typeface(latin_typeface, default_font, force_font=False):
    """
    latin 폰트에 대응하는 ea 폰트를 결정합니다.
    - 테마 참조(+mn-lt/+mj-lt) -> +mn-ea/+mj-ea
    - force_font=True(-f 명시)인 경우 -> default_font 강제 적용
    - 한글 지원 폰트 -> latin과 동일하게 동기화
    - 그 외(Calibri, Arial 등 한글 글리프 없음) -> default_font
    """
    if latin_typeface in THEME_LATIN_TO_EA:
        return THEME_LATIN_TO_EA[latin_typeface]
    if latin_typeface in THEME_EA_REFS:
        return latin_typeface
    if force_font:
        return default_font
    if is_korean_font(latin_typeface):
        return latin_typeface
    return default_font


def _insert_latin_element(parent_elm, target_typeface):
    """
    rPr 또는 defRPr 내부에 <a:latin> 요소를 스키마 순서에 맞게 삽입하거나 갱신합니다.
    순서: ... -> uFill -> latin -> ea -> cs -> sym -> hlinkClick -> hlinkMouseOver -> rtl -> extLst
    """
    latin = parent_elm.find(qualify(A_NS, "latin"))
    if latin is not None:
        if latin.get("typeface") != target_typeface:
            latin.set("typeface", target_typeface)
            return True
        return False

    latin_elem = parent_elm.makeelement(qualify(A_NS, "latin"))
    latin_elem.set("typeface", target_typeface)

    for i, child in enumerate(parent_elm):
        if child.tag in POST_LATIN_TAGS:
            parent_elm.insert(i, latin_elem)
            return True

    parent_elm.append(latin_elem)
    return True


def _insert_ea_element(parent_elm, target_typeface, force_font=False):
    """
    OpenXML ISO/IEC 29500-1 DrawingML 스키마 규격에 맞춰
    rPr 또는 defRPr 내부에 <a:ea> 요소를 올바른 순서로 삽입/갱신합니다.
    순서: latin -> ea -> cs -> sym -> hlinkClick -> hlinkMouseOver -> rtl -> extLst
    """
    ea = parent_elm.find(qualify(A_NS, "ea"))
    if ea is not None:
        current = ea.get("typeface")
        # 1. 이전 버전이 잘못 기록한 latin용 테마 참조(+mn-lt 등)를 ea용 참조로 교정
        if current in THEME_LATIN_TO_EA:
            mapped = THEME_LATIN_TO_EA[current]
            if current != mapped:
                ea.set("typeface", mapped)
                return True
            return False
        # 2. 테마 ea 참조(+mn-ea, +mj-ea)는 유지 (테마 자체는 patch_theme_xml에서 보정됨)
        if current in THEME_EA_REFS:
            return False
        # 3. -f 옵션(force_font)이 지정되었거나, 비어있거나, 한글 미지원 폰트(Calibri/Arial 등)인 경우 교정
        if force_font or not current or not is_korean_font(current):
            if current != target_typeface and target_typeface:
                ea.set("typeface", target_typeface)
                return True
        return False

    ea_elem = parent_elm.makeelement(qualify(A_NS, "ea"))
    ea_elem.set("typeface", target_typeface)

    latin = parent_elm.find(qualify(A_NS, "latin"))
    if latin is not None:
        idx = list(parent_elm).index(latin)
        parent_elm.insert(idx + 1, ea_elem)
        return True

    for i, child in enumerate(parent_elm):
        if child.tag in POST_EA_TAGS:
            parent_elm.insert(i, ea_elem)
            return True

    parent_elm.append(ea_elem)
    return True


def _sync_rpr_fonts(rpr_elm, fallback_font, force_font=False):
    """
    rPr / defRPr / endParaRPr 내부의 <a:latin>과 <a:ea> 폰트를 동기화합니다.
    - 테마 폰트(+mn-lt, +mj-lt)는 +mn-ea, +mj-ea로 매핑합니다.
    - latin이 한글 미지원 서양 폰트(Calibri 등)이거나 force_font=True이면 latin도 한글 폰트로 동기화하여
      한 문장 내 영문/숫자와 한글의 폰트 불일치(짝짝이 폰트)를 방지합니다.
    """
    changed = False
    latin = rpr_elm.find(qualify(A_NS, "latin"))
    ea = rpr_elm.find(qualify(A_NS, "ea"))

    latin_face = latin.get("typeface") if latin is not None else None
    ea_face = ea.get("typeface") if ea is not None else None

    if is_symbol_font(latin_face):
        return False

    # 타겟 ea 폰트 결정: latin이 없으면 기존 ea(한글 폰트인 경우)도 참고
    source_face = latin_face
    if not source_face and ea_face and is_korean_font(ea_face):
        source_face = ea_face

    target_ea = resolve_ea_typeface(source_face, fallback_font, force_font=force_font)

    if _insert_ea_element(rpr_elm, target_ea, force_font=force_font):
        changed = True

    # latin 동기화:
    # - latin이 이미 테마 참조(+mn-lt/+mj-lt)이거나 target_ea가 테마 참조(+mn-ea/+mj-ea)이면 건드리지 않음
    # - latin이 한글 미지원 폰트(Calibri 등)이거나, force_font=True이거나,
    #   latin이 없는데 ea가 일반 폰트로 설정된 경우 latin도 동일 폰트로 맞춤
    final_ea_elm = rpr_elm.find(qualify(A_NS, "ea"))
    final_ea_face = final_ea_elm.get("typeface") if final_ea_elm is not None else target_ea

    if final_ea_face and final_ea_face not in THEME_EA_REFS and latin_face not in THEME_LATIN_REFS:
        if latin is not None:
            if force_font or not latin_face or not is_korean_font(latin_face):
                if latin_face != final_ea_face:
                    latin.set("typeface", final_ea_face)
                    changed = True
        else:
            if _insert_latin_element(rpr_elm, final_ea_face):
                changed = True

    return changed


def _get_effective_defrpr_font(pPr, default_font, force_font=False):
    """문단 pPr의 defRPr에 설정된 유효 한글 폰트가 있으면 반환하여 하위 런이 계승할 수 있게 합니다."""
    if pPr is None:
        return default_font
    defRPr = pPr.find(qualify(A_NS, "defRPr"))
    if defRPr is None:
        return default_font
    if force_font:
        return default_font
    ea = defRPr.find(qualify(A_NS, "ea"))
    if ea is not None:
        ea_face = ea.get("typeface")
        if ea_face in THEME_EA_REFS or is_korean_font(ea_face):
            return ea_face
    latin = defRPr.find(qualify(A_NS, "latin"))
    if latin is not None:
        latin_face = latin.get("typeface")
        if latin_face in THEME_LATIN_TO_EA:
            return THEME_LATIN_TO_EA[latin_face]
        if is_korean_font(latin_face):
            return latin_face
    return default_font


def patch_theme_xml(xml_bytes, default_font="Pretendard", force_font=False):
    """
    ppt/theme/theme*.xml 내의 <a:fontScheme>(majorFont, minorFont)을 보정합니다.
    - 기본 상태(ea="", script="Hang"이 '맑은 고딕')에서 +mj-ea/+mn-ea 참조 시 '맑은 고딕'으로
      폴백되는 문제를 방지합니다.
    - force_font=True(-f 옵션 사용 시)이면 테마의 latin/ea/Hang 폰트를 모두 지정 폰트로 교체합니다.
    """
    try:
        root = _parse(xml_bytes)
    except Exception as e:
        return False, xml_bytes, {"error": str(e)}

    modified = False
    font_scheme = root.find(f".//{qualify(A_NS, 'fontScheme')}")
    if font_scheme is None:
        return False, xml_bytes, {}

    for group_tag in ("majorFont", "minorFont"):
        group = font_scheme.find(qualify(A_NS, group_tag))
        if group is None:
            continue

        latin = group.find(qualify(A_NS, "latin"))
        ea = group.find(qualify(A_NS, "ea"))

        latin_face = latin.get("typeface") if latin is not None else ""
        ea_face = ea.get("typeface") if ea is not None else ""

        # 타겟 테마 폰트 결정
        if force_font:
            target_font = default_font
        elif is_korean_font(ea_face):
            target_font = ea_face
        elif is_korean_font(latin_face):
            target_font = latin_face
        else:
            target_font = default_font

        if latin is not None:
            if force_font or not latin_face or not is_korean_font(latin_face):
                if latin.get("typeface") != target_font:
                    latin.set("typeface", target_font)
                    modified = True

        if ea is not None:
            if force_font or not ea_face or not is_korean_font(ea_face):
                if ea.get("typeface") != target_font:
                    ea.set("typeface", target_font)
                    modified = True

        # <a:font script="Hang" typeface="..."/> 동기화
        for font_elm in group.findall(qualify(A_NS, "font")):
            if font_elm.get("script") == "Hang":
                cur_hang = font_elm.get("typeface")
                if force_font or not cur_hang or not is_korean_font(cur_hang) or cur_hang != target_font:
                    if cur_hang != target_font:
                        font_elm.set("typeface", target_font)
                        modified = True

    if modified:
        new_xml = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
        return True, new_xml, {"theme_patched": 1}
    return False, xml_bytes, {"theme_patched": 0}


def patch_xml_content(xml_bytes, default_font="Pretendard", force_font=False):
    """
    슬라이드 또는 마스터/레이아웃 XML 바이트를 파싱하여
    1. pPr 및 lvl1pPr~lvl9pPr/defPPr: eaLnBrk="0", latinLnBrk="0" (pPr 미존재 시 자동 생성 주입)
    2. defRPr: lang="ko-KR" 및 폰트 동기화 (문단/마스터 레벨 서식 상속 보존)
    3. rPr (a:r 및 a:fld): lang="ko-KR", <a:ea typeface="..."/> 스키마 순서 엄격 동기화
    4. endParaRPr: lang="ko-KR" 및 폰트 동기화
    를 수행하고 수정 여부와 수정된 바이트를 반환합니다.
    """
    try:
        root = _parse(xml_bytes)
    except Exception as e:
        return False, xml_bytes, {"error": str(e)}

    stats = {
        "pPr_patched": 0,
        "rPr_lang_patched": 0,
        "rPr_ea_patched": 0,
        "endPara_patched": 0,
        "defRPr_patched": 0,
    }

    modified = False

    # 0. 마스터/레이아웃/도형의 lstStyle 및 txStyles 내부 레벨 프로퍼티(defPPr, lvl1pPr~lvl9pPr) 보정
    for elm in root.iter():
        if elm.tag in LEVEL_PPR_TAGS:
            if elm.get("eaLnBrk") != "0":
                elm.set("eaLnBrk", "0")
                stats["pPr_patched"] += 1
                modified = True
            if elm.get("latinLnBrk") != "0":
                elm.set("latinLnBrk", "0")
                stats["pPr_patched"] += 1
                modified = True

            lvl_defRPr = elm.find(qualify(A_NS, "defRPr"))
            if lvl_defRPr is not None:
                if lvl_defRPr.get("lang") != "ko-KR":
                    lvl_defRPr.set("lang", "ko-KR")
                    stats["defRPr_patched"] += 1
                    modified = True
                if _sync_rpr_fonts(lvl_defRPr, default_font, force_font=force_font):
                    stats["defRPr_patched"] += 1
                    modified = True

    # 1. 모든 문단 (<a:p>) 순회
    for p in root.iter(qualify(A_NS, "p")):
        pPr = p.find(qualify(A_NS, "pPr"))
        if pPr is None:
            # pPr이 없는 문단: 최상단에 pPr을 신규 생성하여 주입
            pPr = p.makeelement(qualify(A_NS, "pPr"))
            pPr.set("eaLnBrk", "0")
            pPr.set("latinLnBrk", "0")
            p.insert(0, pPr)
            stats["pPr_patched"] += 1
            modified = True
        else:
            # eaLnBrk="0" & latinLnBrk="0" 주입
            if pPr.get("eaLnBrk") != "0":
                pPr.set("eaLnBrk", "0")
                stats["pPr_patched"] += 1
                modified = True
            if pPr.get("latinLnBrk") != "0":
                pPr.set("latinLnBrk", "0")
                stats["pPr_patched"] += 1
                modified = True

        # defRPr 처리 (문단 레벨 서식 상속 우선)
        defRPr = pPr.find(qualify(A_NS, "defRPr"))
        if defRPr is not None:
            if defRPr.get("lang") != "ko-KR":
                defRPr.set("lang", "ko-KR")
                stats["defRPr_patched"] += 1
                modified = True
            if _sync_rpr_fonts(defRPr, default_font, force_font=force_font):
                stats["defRPr_patched"] += 1
                modified = True

        # 문단의 defRPr에 설정된 폰트가 있으면 하위 런의 기본값으로 계승
        para_effective_font = _get_effective_defrpr_font(pPr, default_font, force_font=force_font)

        # 2. 문단 내 런 (<a:r>) 및 필드 (<a:fld>) 순회
        for run_like in p.iter(qualify(A_NS, "r"), qualify(A_NS, "fld")):
            rPr = run_like.find(qualify(A_NS, "rPr"))
            if rPr is None:
                rPr = run_like.makeelement(qualify(A_NS, "rPr"))
                run_like.insert(0, rPr)
                modified = True

            # lang="ko-KR"
            if rPr.get("lang") != "ko-KR":
                rPr.set("lang", "ko-KR")
                stats["rPr_lang_patched"] += 1
                modified = True

            # latin 폰트와 ea 폰트 동기화 (문단 defRPr 상속 폰트 존중 및 스키마 순서 준수)
            if _sync_rpr_fonts(rPr, para_effective_font, force_font=force_font):
                stats["rPr_ea_patched"] += 1
                modified = True

        # 3. endParaRPr 처리
        endParaRPr = p.find(qualify(A_NS, "endParaRPr"))
        if endParaRPr is not None:
            if endParaRPr.get("lang") != "ko-KR":
                endParaRPr.set("lang", "ko-KR")
                stats["endPara_patched"] += 1
                modified = True
            if (
                endParaRPr.find(qualify(A_NS, "latin")) is not None
                or endParaRPr.find(qualify(A_NS, "ea")) is not None
                or force_font
            ):
                if _sync_rpr_fonts(endParaRPr, para_effective_font, force_font=force_font):
                    stats["endPara_patched"] += 1
                    modified = True

    if modified:
        new_xml = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
        return True, new_xml, stats
    return False, xml_bytes, stats


def _has_valid_ea(rpr_elm):
    if rpr_elm is None:
        return False
    ea = rpr_elm.find(qualify(A_NS, "ea"))
    if ea is None:
        return False
    face = ea.get("typeface")
    return bool(face and (face in THEME_EA_REFS or is_korean_font(face)))


def verify_xml_content(xml_bytes):
    """
    서식 누락 항목 정밀 진단 (읽기 전용).
    문단 레벨(defRPr)에서 언어와 동아시아 폰트를 상속받는 경우(SKILL.md 제3원칙)도 정상으로 인정합니다.
    """
    try:
        root = _parse(xml_bytes)
    except Exception as e:
        return {"error": str(e)}

    issues = {
        "missing_eaLnBrk": 0,
        "missing_latinLnBrk": 0,
        "missing_ko_lang": 0,
        "missing_ea_font": 0,
    }

    for p in root.iter(qualify(A_NS, "p")):
        pPr = p.find(qualify(A_NS, "pPr"))
        defRPr = pPr.find(qualify(A_NS, "defRPr")) if pPr is not None else None
        defrpr_has_ko = defRPr is not None and defRPr.get("lang") == "ko-KR"
        defrpr_has_ea = _has_valid_ea(defRPr)

        if pPr is None:
            issues["missing_eaLnBrk"] += 1
            issues["missing_latinLnBrk"] += 1
        else:
            if pPr.get("eaLnBrk") != "0":
                issues["missing_eaLnBrk"] += 1
            if pPr.get("latinLnBrk") != "0":
                issues["missing_latinLnBrk"] += 1

        for r in p.iter(qualify(A_NS, "r")):
            rPr = r.find(qualify(A_NS, "rPr"))
            r_lang = rPr.get("lang") if rPr is not None else None
            if r_lang != "ko-KR" and not (r_lang is None and defrpr_has_ko):
                issues["missing_ko_lang"] += 1

            r_has_own_latin = rPr is not None and rPr.find(qualify(A_NS, "latin")) is not None
            if not _has_valid_ea(rPr) and not (not r_has_own_latin and defrpr_has_ea):
                issues["missing_ea_font"] += 1

    return issues


def get_slide_file_map(zin):
    """
    presentation.xml 및 .rels 파일을 파싱하여
    실제 프레젠테이션 슬라이드 표시 순서 (1-based index) -> zip 내부 파일 경로 매핑을 생성합니다.
    슬라이드 순서 재배치나 삭제가 있어도 정확한 순서 매핑을 보장합니다.
    """
    try:
        if "ppt/_rels/presentation.xml.rels" not in zin.namelist():
            return {}
        rels_xml = zin.read("ppt/_rels/presentation.xml.rels")
        rels_root = _parse(rels_xml)
        rel_map = {}
        for rel in rels_root.findall(qualify(PKG_REL_NS, "Relationship")):
            rid = rel.get("Id")
            target = rel.get("Target")
            if not target:
                continue
            if target.startswith("/"):
                path = posixpath.normpath(target.lstrip("/"))
            else:
                path = posixpath.normpath(posixpath.join("ppt", target))
            rel_map[rid] = path

        if "ppt/presentation.xml" not in zin.namelist():
            return {}
        pres_xml = zin.read("ppt/presentation.xml")
        pres_root = _parse(pres_xml)

        slide_map = {}
        sld_id_lst = pres_root.find(qualify(P_NS, "sldIdLst"))
        if sld_id_lst is not None:
            idx = 1
            for sld_id in sld_id_lst.findall(qualify(P_NS, "sldId")):
                rid = sld_id.get(qualify(R_NS_DOC, "id"))
                if rid and rid in rel_map:
                    slide_map[idx] = rel_map[rid]
                    idx += 1
        return slide_map
    except Exception:
        return {}


def process_pptx(
    pptx_path,
    output_path=None,
    target_slide=None,
    default_font="Pretendard",
    verify_only=False,
    include_masters=True,
    force_font=None,
):
    if force_font is None:
        force_font = default_font != "Pretendard"

    if not os.path.exists(pptx_path):
        print(f"[ERROR] 파일이 존재하지 않습니다: {pptx_path}", file=sys.stderr)
        return 1

    if output_path is None:
        output_path = pptx_path

    temp_dir = tempfile.mkdtemp(prefix="korean_pptx_")
    total_stats = {
        "slides_processed": 0,
        "slides_modified": 0,
        "pPr_patched": 0,
        "rPr_lang_patched": 0,
        "rPr_ea_patched": 0,
        "defRPr_patched": 0,
        "endPara_patched": 0,
        "theme_patched": 0,
    }

    issues_found = {}

    try:
        with zipfile.ZipFile(pptx_path, "r") as zin:
            file_list = zin.namelist()
            slide_file_map = get_slide_file_map(zin)

            layout_pattern = re.compile(r"^ppt/slideLayouts/slideLayout\d+\.xml$")
            master_pattern = re.compile(r"^ppt/slideMasters/slideMaster\d+\.xml$")
            theme_pattern = re.compile(r"^ppt/theme/theme\d+\.xml$")
            slide_pattern = re.compile(r"^ppt/slides/slide(\d+)\.xml$")

            target_files = []
            theme_files = set()

            # 1. 슬라이드 선별
            if slide_file_map:
                for idx, path in sorted(slide_file_map.items()):
                    if target_slide is None or target_slide == idx:
                        if path in file_list:
                            target_files.append((path, "slide", idx))
            else:
                # 폴백: 파일명 기준 매칭
                for name in file_list:
                    m = slide_pattern.match(name)
                    if m:
                        s_num = int(m.group(1))
                        if target_slide is None or target_slide == s_num:
                            target_files.append((name, "slide", s_num))

            # 2. 마스터, 레이아웃 및 테마 선별 (전체 모드이거나 force_font인 경우 테마 포함)
            if include_masters and target_slide is None:
                for name in file_list:
                    if layout_pattern.match(name):
                        target_files.append((name, "layout", 0))
                    elif master_pattern.match(name):
                        target_files.append((name, "master", 0))
                    elif theme_pattern.match(name):
                        theme_files.add(name)
            elif force_font and include_masters:
                for name in file_list:
                    if theme_pattern.match(name):
                        theme_files.add(name)

            if not target_files:
                target_desc = f"슬라이드 #{target_slide}" if target_slide is not None else "슬라이드"
                print(f"[!] 지정된 {target_desc} 파일을 찾을 수 없습니다.", file=sys.stderr)
                return 1

            if verify_only:
                has_any_issue = False
                for name, ftype, num in target_files:
                    xml_data = zin.read(name)
                    issues = verify_xml_content(xml_data)
                    total_issue_count = sum(issues.values())
                    if total_issue_count > 0:
                        has_any_issue = True
                        issues_found[name] = issues

                if has_any_issue:
                    print(f"[-] 서식 미흡 항목 발견 ({len(issues_found)}개 파트):")
                    for name, iss in issues_found.items():
                        detail = ", ".join(f"{k}={v}" for k, v in iss.items() if v > 0)
                        print(f"  - {name}: {detail}")
                    return 2
                else:
                    print("[OK] 검증 완료: 모든 문단 및 런의 한국어 서식이 완벽하게 적용되어 있습니다.")
                    return 0

            # 수정 모드: 임시 zip 파일로 복사하며 패치
            temp_zip_path = os.path.join(temp_dir, "patched.zip")
            with zipfile.ZipFile(temp_zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                target_names = {tf[0] for tf in target_files}
                for item in zin.infolist():
                    data = zin.read(item.filename)

                    if item.filename in target_names:
                        total_stats["slides_processed"] += 1
                        mod, new_data, stats = patch_xml_content(
                            data, default_font=default_font, force_font=force_font
                        )
                        if mod:
                            total_stats["slides_modified"] += 1
                            for k in stats:
                                total_stats[k] = total_stats.get(k, 0) + stats[k]
                            zout.writestr(item, new_data)
                        else:
                            zout.writestr(item, data)
                    elif item.filename in theme_files:
                        mod, new_data, t_stats = patch_theme_xml(
                            data, default_font=default_font, force_font=force_font
                        )
                        if mod:
                            total_stats["theme_patched"] += t_stats.get("theme_patched", 1)
                            zout.writestr(item, new_data)
                        else:
                            zout.writestr(item, data)
                    else:
                        zout.writestr(item, data)

        # 원본 또는 출력 경로로 안전하게 이동
        shutil.move(temp_zip_path, output_path)

        # 결과 요약 보고
        mod_count = total_stats["slides_modified"]
        proc_count = total_stats["slides_processed"]
        p_count = total_stats["pPr_patched"]
        lang_count = total_stats["rPr_lang_patched"]
        ea_count = total_stats["rPr_ea_patched"]
        theme_count = total_stats["theme_patched"]

        target_desc = f"슬라이드 #{target_slide}" if target_slide is not None else "전체 슬라이드 및 마스터"
        font_mode_desc = f"강제 지정 typeface='{default_font}'" if force_font else f"기본 typeface='{default_font}', 테마/한글 폰트는 유지"
        print(f"[OK] 한국어 서식 보정 완료 ({target_desc}):")
        print(f"  - 검사 대상: {proc_count}개 파일 중 {mod_count}개 보정 적용 (테마 {theme_count}개 보정)")
        print(f"  - 줄바꿈 방지 설정: {p_count}개 속성 주입 (eaLnBrk/latinLnBrk='0')")
        print(f"  - 한국어 언어 태그: {lang_count}개 런 적용 (lang='ko-KR')")
        print(f"  - 동아시아 폰트 보정: {ea_count}개 런 동기화 ({font_mode_desc})")
        return 0

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def main():
    parser = argparse.ArgumentParser(
        description="PowerPoint (.pptx) 한국어 서식(언어 태그, 줄바꿈 방지, 동아시아/테마 폰트) 고속 보정 도구"
    )
    parser.add_argument("pptx_file", help="대상 .pptx 파일 경로")
    parser.add_argument("-o", "--output", help="출력 .pptx 파일 경로 (기본값: 원본 파일 덮어쓰기)")
    parser.add_argument("-s", "--slide", type=int, help="특정 슬라이드 번호만 지정 보정 (예: 1)")
    parser.add_argument("-f", "--font", default=None, help="동아시아/테마 폰트명 지정 (미지정 시 기본값: Pretendard, 지정 시 테마 폰트까지 일괄 교체)")
    parser.add_argument("--verify", action="store_true", help="수정하지 않고 서식 누락 여부만 검증")
    parser.add_argument("--no-masters", action="store_true", help="슬라이드 마스터/레이아웃/테마 제외하고 슬라이드 본문만 보정")

    args = parser.parse_args()
    force_font = args.font is not None
    default_font = args.font if args.font is not None else "Pretendard"

    # cp949 등 제한된 콘솔 인코딩에서 출력 문자 때문에 (이미 저장된 뒤) 크래시하지 않도록 함
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass

    ret = process_pptx(
        pptx_path=args.pptx_file,
        output_path=args.output,
        target_slide=args.slide,
        default_font=default_font,
        verify_only=args.verify,
        include_masters=not args.no_masters,
        force_font=force_font,
    )
    sys.exit(ret)


if __name__ == "__main__":
    main()
