#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fix_korean_pptx.py
PPTX 파일의 한국어 서식(언어 태그, 단어 단위 줄바꿈, 동아시아 폰트)을 일괄 보정하는 고속 CLI 도구.
파이썬 표준 라이브러리(zipfile, xml.etree.ElementTree)만으로 동작하여 의존성 없이 초고속(0.1초 미만) 처리됩니다.
"""

import argparse
import os
import re
import shutil
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

ET.register_namespace("a", A_NS)
ET.register_namespace("p", P_NS)
ET.register_namespace("r", R_NS_DOC)

POST_EA_TAGS = [f"{{{A_NS}}}{t}" for t in ["cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"]]


def qualify(ns, tag):
    return f"{{{ns}}}{tag}"


def _insert_ea_element(parent_elm, target_typeface):
    """
    OpenXML ISO/IEC 29500-1 DrawingML 스키마 규격에 맞춰
    rPr 또는 defRPr 내부에 <a:ea> 요소를 올바른 순서로 삽입/갱신합니다.
    순서: latin -> ea -> cs -> sym -> hlinkClick -> hlinkMouseOver -> rtl -> extLst
    """
    ea = parent_elm.find(qualify(A_NS, "ea"))
    if ea is not None:
        if not ea.get("typeface") and target_typeface:
            ea.set("typeface", target_typeface)
            return True
        return False

    ea_elem = ET.Element(qualify(A_NS, "ea"))
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


def patch_xml_content(xml_bytes, default_font="Pretendard"):
    """
    슬라이드 또는 마스터/레이아웃 XML 바이트를 파싱하여
    1. pPr: eaLnBrk="0", latinLnBrk="0" (pPr 미존재 시 자동 생성 주입)
    2. rPr: lang="ko-KR", <a:ea typeface="..."/> 스키마 순서 엄격 동기화
    3. defRPr / endParaRPr: lang="ko-KR" 보장
    을 수행하고 수정 여부와 수정된 바이트를 반환합니다.
    """
    try:
        root = ET.fromstring(xml_bytes)
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

    # 1. 모든 문단 (<a:p>) 순회
    for p in root.iter(qualify(A_NS, "p")):
        pPr = p.find(qualify(A_NS, "pPr"))
        if pPr is None:
            # pPr이 없는 문단: 최상단에 pPr을 신규 생성하여 주입
            pPr = ET.Element(qualify(A_NS, "pPr"))
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

        # defRPr 처리
        defRPr = pPr.find(qualify(A_NS, "defRPr"))
        if defRPr is not None:
            if defRPr.get("lang") != "ko-KR":
                defRPr.set("lang", "ko-KR")
                stats["defRPr_patched"] += 1
                modified = True
            # defRPr 내부의 ea 폰트 확인
            latin_def = defRPr.find(qualify(A_NS, "latin"))
            font_to_apply = default_font
            if latin_def is not None and latin_def.get("typeface"):
                font_to_apply = latin_def.get("typeface")
            if _insert_ea_element(defRPr, font_to_apply):
                stats["defRPr_patched"] += 1
                modified = True

        # 2. 문단 내 런 (<a:r>) 순회
        for r in p.iter(qualify(A_NS, "r")):
            rPr = r.find(qualify(A_NS, "rPr"))
            if rPr is None:
                rPr = ET.Element(qualify(A_NS, "rPr"))
                r.insert(0, rPr)
                modified = True

            # lang="ko-KR"
            if rPr.get("lang") != "ko-KR":
                rPr.set("lang", "ko-KR")
                stats["rPr_lang_patched"] += 1
                modified = True

            # latin 폰트와 ea 폰트 동기화 (스키마 순서 준수)
            latin = rPr.find(qualify(A_NS, "latin"))
            target_typeface = default_font
            if latin is not None and latin.get("typeface"):
                target_typeface = latin.get("typeface")

            if _insert_ea_element(rPr, target_typeface):
                stats["rPr_ea_patched"] += 1
                modified = True

        # 3. endParaRPr 처리
        endParaRPr = p.find(qualify(A_NS, "endParaRPr"))
        if endParaRPr is not None:
            if endParaRPr.get("lang") != "ko-KR":
                endParaRPr.set("lang", "ko-KR")
                stats["endPara_patched"] += 1
                modified = True

    if modified:
        new_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        return True, new_xml, stats
    return False, xml_bytes, stats


def verify_xml_content(xml_bytes):
    """서식 누락 항목 정밀 진단 (읽기 전용)"""
    try:
        root = ET.fromstring(xml_bytes)
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
            if rPr is None or rPr.get("lang") != "ko-KR":
                issues["missing_ko_lang"] += 1
            if rPr is None or rPr.find(qualify(A_NS, "ea")) is None:
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
        rels_root = ET.fromstring(rels_xml)
        rel_map = {}
        for rel in rels_root.findall(qualify(PKG_REL_NS, "Relationship")):
            rid = rel.get("Id")
            target = rel.get("Target")
            if target.startswith("/"):
                path = target.lstrip("/")
            else:
                path = f"ppt/{target}"
            rel_map[rid] = path

        if "ppt/presentation.xml" not in zin.namelist():
            return {}
        pres_xml = zin.read("ppt/presentation.xml")
        pres_root = ET.fromstring(pres_xml)

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


def process_pptx(pptx_path, output_path=None, target_slide=None, default_font="Pretendard", verify_only=False, include_masters=True):
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
    }

    issues_found = {}

    try:
        with zipfile.ZipFile(pptx_path, "r") as zin:
            file_list = zin.namelist()
            slide_file_map = get_slide_file_map(zin)

            layout_pattern = re.compile(r"^ppt/slideLayouts/slideLayout\d+\.xml$")
            master_pattern = re.compile(r"^ppt/slideMasters/slideMaster\d+\.xml$")
            slide_pattern = re.compile(r"^ppt/slides/slide(\d+)\.xml$")

            target_files = []

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

            # 2. 마스터 및 레이아웃 선별 (전체 모드인 경우)
            if include_masters and target_slide is None:
                for name in file_list:
                    if layout_pattern.match(name):
                        target_files.append((name, "layout", 0))
                    elif master_pattern.match(name):
                        target_files.append((name, "master", 0))

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
                        print(f"  • {name}: {detail}")
                    return 2
                else:
                    print("[✓] 검증 완료: 모든 문단 및 런의 한국어 서식이 완벽하게 적용되어 있습니다.")
                    return 0

            # 수정 모드: 임시 zip 파일로 복사하며 패치
            temp_zip_path = os.path.join(temp_dir, "patched.zip")
            with zipfile.ZipFile(temp_zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                target_names = {tf[0] for tf in target_files}
                for item in zin.infolist():
                    data = zin.read(item.filename)

                    if item.filename in target_names:
                        total_stats["slides_processed"] += 1
                        mod, new_data, stats = patch_xml_content(data, default_font=default_font)
                        if mod:
                            total_stats["slides_modified"] += 1
                            for k in stats:
                                total_stats[k] = total_stats.get(k, 0) + stats[k]
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

        target_desc = f"슬라이드 #{target_slide}" if target_slide is not None else "전체 슬라이드 및 마스터"
        print(f"[✓] 한국어 서식 보정 완료 ({target_desc}):")
        print(f"  • 검사 대상: {proc_count}개 파일 중 {mod_count}개 보정 적용")
        print(f"  • 줄바꿈 방지 설정: {p_count}개 속성 주입 (eaLnBrk/latinLnBrk='0')")
        print(f"  • 한국어 언어 태그: {lang_count}개 런 적용 (lang='ko-KR')")
        print(f"  • 동아시아 폰트 보정: {ea_count}개 런 동기화 (typeface='{default_font}')")
        return 0

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def main():
    parser = argparse.ArgumentParser(
        description="PowerPoint (.pptx) 한국어 서식(언어 태그, 줄바꿈 방지, 동아시아 폰트) 고속 보정 도구"
    )
    parser.add_argument("pptx_file", help="대상 .pptx 파일 경로")
    parser.add_argument("-o", "--output", help="출력 .pptx 파일 경로 (기본값: 원본 파일 덮어쓰기)")
    parser.add_argument("-s", "--slide", type=int, help="특정 슬라이드 번호만 지정 보정 (예: 1)")
    parser.add_argument("-f", "--font", default="Pretendard", help="기본 동아시아 폰트명 (기본값: Pretendard)")
    parser.add_argument("--verify", action="store_true", help="수정하지 않고 서식 누락 여부만 검증")
    parser.add_argument("--no-masters", action="store_true", help="슬라이드 마스터/레이아웃 제외하고 슬라이드 본문만 보정")

    args = parser.parse_args()

    ret = process_pptx(
        pptx_path=args.pptx_file,
        output_path=args.output,
        target_slide=args.slide,
        default_font=args.font,
        verify_only=args.verify,
        include_masters=not args.no_masters,
    )
    sys.exit(ret)


if __name__ == "__main__":
    main()
