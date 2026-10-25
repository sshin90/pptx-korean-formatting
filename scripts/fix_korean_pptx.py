#!/usr/bin/env python3
"""Conservative Korean DrawingML formatting.

Exit codes: 0 processed / checked scope passed, 1 input or I/O error,
2 verified defects, 3 unresolved inheritance requiring review.
"""
import argparse
import os
import posixpath
import re
import sys
import tempfile
import zipfile
from collections import Counter
from lxml import etree

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
HANGUL_RE = re.compile("[ᄀ-ᇿ㄰-㆏가-힣ꥠ-꥿ힰ-퟿]")
THEME_LATIN_TO_EA = {"+mn-lt": "+mn-ea", "+mj-lt": "+mj-ea"}
THEME_EA_REFS = {"+mn-ea", "+mj-ea"}
SYMBOL_FONTS = {
    "wingdings", "wingdings 2", "wingdings 3", "webdings", "symbol",
    "segoe ui symbol", "segoe mdl2 assets", "segoe fluent icons", "marlett",
}


def qualify(ns, tag):
    return f"{{{ns}}}{tag}"


def a(tag):
    return qualify(A_NS, tag)


LEVEL_PPR_TAGS = {a("defPPr"), *(a(f"lvl{i}pPr") for i in range(1, 10))}
RPR_TAGS = {a(t) for t in ("rPr", "defRPr", "endParaRPr")}
FALSE_VALUES = {"0", "false"}
FONT_ORDER = [a(t) for t in ("latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst")]
BULLET_GROUPS = [
    ("buClrTx", "buClr"), ("buSzTx", "buSzPct", "buSzPts"),
    ("buFontTx", "buFont"), ("buNone", "buAutoNum", "buChar", "buBlip"),
]


def _parse(data):
    root = etree.fromstring(data, etree.XMLParser(resolve_entities=False, no_network=True))
    if root.getroottree().docinfo.doctype:
        raise ValueError("DTD를 포함한 XML은 지원하지 않습니다")
    return root


def _serialize(root):
    return etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


def is_symbol_font(face):
    return bool(face and face.strip().lower() in SYMBOL_FONTS)


def resolve_ea_typeface(latin_typeface, default_font=None, force_font=False):
    """Only map theme tokens; never infer glyph coverage from a font name."""
    if force_font:
        return default_font
    return THEME_LATIN_TO_EA.get(latin_typeface, latin_typeface)


def _face(prop, tag):
    elm = prop.find(a(tag)) if prop is not None else None
    return elm.get("typeface") if elm is not None else None


def _symbol(prop):
    return any(is_symbol_font(_face(prop, tag)) for tag in ("latin", "ea", "sym"))


def _ensure(parent, tag, before=()):
    elm = parent.find(a(tag))
    if elm is None:
        elm = parent.makeelement(a(tag), {})
        for i, child in enumerate(parent):
            if child.tag in {a(t) for t in before}:
                parent.insert(i, elm)
                break
        else:
            parent.append(elm)
    return elm


def _set_face(prop, tag, value):
    successors = [etree.QName(t).localname for t in FONT_ORDER[FONT_ORDER.index(a(tag)) + 1:]]
    _ensure(prop, tag, successors).set("typeface", value)


def _ppr_chain(p):
    """Local paragraph, current list level, list defaults; no cross-part guess."""
    ppr = p.find(a("pPr"))
    chain = [ppr] if ppr is not None else []
    level = ppr.get("lvl", "0") if ppr is not None else "0"
    if not level.isdigit() or not 0 <= int(level) <= 8:
        raise ValueError(f"유효하지 않은 문단 레벨: {level}")
    body = p.getparent()
    style = body.find(a("lstStyle")) if body is not None else None
    if style is not None:
        for tag in (f"lvl{int(level) + 1}pPr", "defPPr"):
            elm = style.find(a(tag))
            if elm is not None:
                chain.append(elm)
    return chain


def _rpr_chain(p, own=None):
    chain = [own] if own is not None else []
    return chain + [d for pr in _ppr_chain(p) if (d := pr.find(a("defRPr"))) is not None]


def _effective_attr(chain, attr):
    return next((pr.get(attr) for pr in chain if pr.get(attr) is not None), None)


def _effective_face(chain, tag="ea"):
    # An explicit empty/invalid value must not fall back to its parent.
    for pr in chain:
        elm = pr.find(a(tag))
        if elm is not None:
            return elm.get("typeface", "")
    return None


def _run_text(run):
    return "".join(t.text or "" for t in run.findall(a("t")))


def _runs(p):
    return [r for r in p if r.tag in {a("r"), a("fld")}]


def _selected(run, all_text):
    return not _symbol(run.find(a("rPr"))) and (all_text or bool(HANGUL_RE.search(_run_text(run))))


def _paragraph_selected(p, all_text):
    return all_text or any(_selected(r, False) for r in _runs(p))


def patch_tree(root, font_name=None, all_text=False):
    """Mutate lxml/python-pptx tree, preserving absent font overrides by default.

    Explicit normalization uses paragraph defaults; only selected run font
    overrides are removed. Bold, colors, hyperlinks, foreign runs survive.
    """
    before = etree.tostring(root)
    counts = Counter()
    if all_text:
        for prop in root.iter():
            if prop.tag in LEVEL_PPR_TAGS:
                prop.set("eaLnBrk", "0")
                prop.set("latinLnBrk", "0")
                default = _ensure(prop, "defRPr", ("extLst",))
                default.set("lang", "ko-KR")
                if font_name and not _symbol(default):
                    for tag in ("latin", "ea"):
                        _set_face(default, tag, font_name)

    for p in root.iter(a("p")):
        if not _paragraph_selected(p, all_text):
            continue
        counts["paragraphs"] += 1
        runs = _runs(p)
        selected = [r for r in runs if _selected(r, all_text)]
        homogeneous = all_text or all(not _run_text(r).strip() or r in selected for r in runs)
        ppr = _ensure(p, "pPr", ("r", "br", "fld", "endParaRPr"))
        for attr in ("eaLnBrk", "latinLnBrk"):
            if _effective_attr(_ppr_chain(p), attr) not in FALSE_VALUES:
                ppr.set(attr, "0")
        body = p.getparent()
        bodypr = body.find(a("bodyPr")) if body is not None else None
        if bodypr is not None:
            bodypr.set("wrap", "square")

        default = ppr.find(a("defRPr"))
        if homogeneous or font_name:
            default = _ensure(ppr, "defRPr", ("extLst",))
        if homogeneous:
            default.set("lang", "ko-KR")
        if font_name:
            for tag in ("latin", "ea"):
                _set_face(default, tag, font_name)
        for run in selected:
            prop = run.find(a("rPr"))
            if homogeneous:
                if prop is not None:
                    prop.attrib.pop("lang", None)
            elif _effective_attr(_rpr_chain(p, prop), "lang") != "ko-KR":
                prop = _ensure(run, "rPr", ("pPr", "t"))
                prop.set("lang", "ko-KR")
            if font_name and prop is not None:
                for tag in ("latin", "ea"):
                    for elm in list(prop.findall(a(tag))):
                        prop.remove(elm)
        end = p.find(a("endParaRPr"))
        if end is not None and homogeneous and not _symbol(end):
            end.attrib.pop("lang", None)
            if font_name:
                for tag in ("latin", "ea"):
                    for elm in list(end.findall(a(tag))):
                        end.remove(elm)
        # Repair explicitly wrong theme tokens, without creating run fonts.
        props = [default, end] + [r.find(a("rPr")) for r in selected]
        for prop in props:
            if prop is not None:
                face = _face(prop, "ea")
                if face in THEME_LATIN_TO_EA:
                    prop.find(a("ea")).set("typeface", THEME_LATIN_TO_EA[face])
    return etree.tostring(root) != before, dict(counts)


def patch_xml_content(xml_bytes, default_font=None, force_font=False, *, all_text=False):
    """default_font is used only with force_font=True; preserve fonts otherwise."""
    try:
        root = _parse(xml_bytes)
        if force_font and (not default_font or not default_font.strip() or default_font.startswith("+")):
            raise ValueError("-f에는 실제 폰트 이름이 필요합니다")
        modified, stats = patch_tree(root, default_font if force_font else None, all_text)
        return modified, _serialize(root) if modified else xml_bytes, stats
    except (etree.Error, ValueError) as exc:
        return False, xml_bytes, {"error": str(exc)}


def patch_theme_xml(xml_bytes, default_font=None, force_font=False):
    """No-op by default; only explicit --update-theme calls this for mutation."""
    try:
        root = _parse(xml_bytes)
        if not force_font:
            return False, xml_bytes, {"theme_patched": 0}
        if not default_font or not default_font.strip() or default_font.startswith("+"):
            raise ValueError("실제 테마 폰트 이름이 필요합니다")
        before = etree.tostring(root)
        for scheme in root.iter(a("fontScheme")):
            for tag in ("majorFont", "minorFont"):
                group = scheme.find(a(tag))
                if group is None:
                    raise ValueError(f"테마에 {tag}가 없습니다")
                for name in ("latin", "ea"):
                    elm = group.find(a(name))
                    if elm is None:
                        raise ValueError(f"테마에 {tag}/{name}가 없습니다")
                    elm.set("typeface", default_font)
                hang = next((e for e in group.findall(a("font")) if e.get("script") == "Hang"), None)
                if hang is None:
                    hang = group.makeelement(a("font"), {"script": "Hang"})
                    ext = group.find(a("extLst"))
                    group.insert(list(group).index(ext) if ext is not None else len(group), hang)
                hang.set("typeface", default_font)
        modified = etree.tostring(root) != before
        return modified, _serialize(root) if modified else xml_bytes, {"theme_patched": int(modified)}
    except (etree.Error, ValueError) as exc:
        return False, xml_bytes, {"error": str(exc)}


def _structural_issues(root, issues):
    """Targeted checks, not a complete ISO schema validator."""
    for prop in root.iter():
        if prop.tag in RPR_TAGS:
            positions = [FONT_ORDER.index(c.tag) for c in prop if c.tag in FONT_ORDER]
            if positions != sorted(set(positions)):
                issues["font_element_order"] += 1
            face = _face(prop, "ea")
            if prop.find(a("ea")) is not None and (not face or not face.strip() or (face.startswith("+") and face not in THEME_EA_REFS)):
                issues["invalid_ea_typeface"] += 1
        if prop.tag == a("pPr") or prop.tag in LEVEL_PPR_TAGS:
            groups = [("lnSpc",), ("spcBef",), ("spcAft",)] + BULLET_GROUPS + [("tabLst",), ("defRPr",), ("extLst",)]
            ranks = {a(t): i for i, group in enumerate(groups) for t in group}
            positions = [ranks[c.tag] for c in prop if c.tag in ranks]
            if positions != sorted(set(positions)):
                issues["paragraph_element_order_or_choice"] += 1
            for attr in ("eaLnBrk", "latinLnBrk"):
                if prop.get(attr) is not None and prop.get(attr) not in {"0", "1", "false", "true"}:
                    issues["invalid_boolean"] += 1


def verify_xml_content(xml_bytes, *, all_text=False, expected_font=None):
    """Local paragraph/list inheritance only; unknown_* never means success.

    Layout/master/theme inheritance, installed fonts, glyph coverage and
    rendering require further review. Font pairs need not be identical.
    """
    try:
        root = _parse(xml_bytes)
        issues = Counter()
        _structural_issues(root, issues)
        if all_text:
            for prop in root.iter():
                if prop.tag in LEVEL_PPR_TAGS:
                    for attr in ("eaLnBrk", "latinLnBrk"):
                        if prop.get(attr) not in FALSE_VALUES:
                            issues[f"style_{attr}"] += 1
                    default = prop.find(a("defRPr"))
                    if default is None or default.get("lang") != "ko-KR":
                        issues["style_language"] += 1
                    if expected_font and not _symbol(default):
                        for tag in ("latin", "ea"):
                            if _face(default, tag) != expected_font:
                                issues[f"style_{tag}_font"] += 1
        for p in root.iter(a("p")):
            if not _paragraph_selected(p, all_text):
                continue
            for attr in ("eaLnBrk", "latinLnBrk"):
                value = _effective_attr(_ppr_chain(p), attr)
                if value is None:
                    issues[f"unknown_{attr}_inheritance"] += 1
                elif value not in FALSE_VALUES:
                    issues[f"enabled_{attr}"] += 1
            body = p.getparent()
            bodypr = body.find(a("bodyPr")) if body is not None else None
            wrap = bodypr.get("wrap") if bodypr is not None else None
            if wrap == "none":
                issues["word_wrap_disabled"] += 1
            elif wrap != "square":
                issues["unknown_word_wrap_inheritance"] += 1
            selected = [r for r in _runs(p) if _selected(r, all_text)]
            props = [r.find(a("rPr")) for r in selected]
            if all_text and not props:
                props.append(None)
            end = p.find(a("endParaRPr"))
            if end is not None and not _symbol(end):
                if all_text or all(not _run_text(r).strip() or r in selected for r in _runs(p)):
                    props.append(end)
            for prop in props:
                chain = _rpr_chain(p, prop)
                lang = _effective_attr(chain, "lang")
                if lang is None:
                    issues["unknown_language_inheritance"] += 1
                elif lang != "ko-KR":
                    issues["non_korean_language"] += 1
                face = _effective_face(chain)
                if face is None or face in THEME_EA_REFS:
                    issues["unknown_ea_font_inheritance"] += 1
                elif not face.strip() or face.startswith("+"):
                    pass  # Explicit invalid EA values are counted structurally.
                elif expected_font and face != expected_font:
                    issues["unexpected_ea_font"] += 1
                if expected_font and _effective_face(chain, "latin") != expected_font:
                    issues["unexpected_latin_font"] += 1
        return dict(issues)
    except (etree.Error, ValueError) as exc:
        return {"error": str(exc)}


def get_slide_file_map(zin):
    """Use relationship order; fail closed on malformed/missing relationships."""
    rels = _parse(zin.read("ppt/_rels/presentation.xml.rels"))
    pres = _parse(zin.read("ppt/presentation.xml"))
    if pres.tag != qualify(P_NS, "presentation"):
        raise ValueError("Transitional PPTX/POTX만 지원합니다")
    rel_map = {}
    for rel in rels.findall(qualify(PKG_REL_NS, "Relationship")):
        if rel.get("TargetMode") == "External" or not rel.get("Type", "").endswith("/slide"):
            continue
        target = rel.get("Target", "")
        rel_map[rel.get("Id")] = posixpath.normpath(target.lstrip("/") if target.startswith("/") else posixpath.join("ppt", target))
    ids = pres.find(qualify(P_NS, "sldIdLst"))
    result = {}
    for i, elm in enumerate(ids if ids is not None else [], 1):
        rid = elm.get(qualify(R_NS_DOC, "id"))
        path = rel_map.get(rid)
        if not path or path not in zin.namelist():
            raise ValueError(f"슬라이드 #{i}의 관계 대상이 없습니다: {rid}")
        result[i] = path
    return result


def process_pptx(pptx_path, output_path=None, target_slide=None, default_font=None,
                 verify_only=False, include_masters=True, force_font=False, *,
                 all_text=False, update_theme=False):
    """Patch to a sibling temporary ZIP and atomically replace only on success."""
    temp_path = None
    try:
        if target_slide is not None and target_slide < 1:
            raise ValueError("슬라이드 번호는 1 이상이어야 합니다")
        if force_font and (not default_font or not default_font.strip() or default_font.startswith("+")):
            raise ValueError("폰트 통일에는 실제 폰트 이름이 필요합니다")
        if update_theme and (not force_font or target_slide is not None or not include_masters):
            raise ValueError("--update-theme은 -f와 함께 전체 문서에만 사용하며 --no-masters와 병용할 수 없습니다")
        if verify_only and output_path is not None:
            raise ValueError("--verify는 출력 파일을 생성하지 않습니다; -o를 제거하세요")
        destination = os.path.abspath(os.fspath(output_path or pptx_path))
        with zipfile.ZipFile(pptx_path) as zin:
            names = zin.namelist()
            if len(names) != len(set(names)):
                raise ValueError("중복 ZIP 파트 이름은 지원하지 않습니다")
            slide_map = get_slide_file_map(zin)
            if target_slide is not None and target_slide not in slide_map:
                raise ValueError(f"슬라이드 #{target_slide}가 없습니다")
            targets = {p for i, p in slide_map.items() if target_slide is None or i == target_slide}
            if include_masters and target_slide is None:
                targets.update(n for n in names if re.fullmatch(r"ppt/(slideMasters|slideLayouts)/[^/]+\.xml", n))
            if not targets:
                raise ValueError("처리할 슬라이드/마스터/레이아웃이 없습니다")
            themes = {n for n in names if re.fullmatch(r"ppt/theme/[^/]+\.xml", n)} if update_theme else set()
            replacements = {}
            diagnostics = {}
            for name in sorted(targets | themes):
                data = zin.read(name)
                if verify_only:
                    if name in themes:
                        modified, _, stats = patch_theme_xml(data, default_font, True)
                        issues = stats if "error" in stats else ({"unexpected_theme_font": 1} if modified else {})
                    else:
                        issues = verify_xml_content(data, all_text=all_text,
                                                    expected_font=default_font if force_font else None)
                    if "error" in issues:
                        raise ValueError(f"{name}: {issues['error']}")
                    if issues:
                        diagnostics[name] = issues
                else:
                    if name in themes:
                        modified, patched, stats = patch_theme_xml(data, default_font, True)
                    else:
                        modified, patched, stats = patch_xml_content(data, default_font, force_font, all_text=all_text)
                    if "error" in stats:
                        raise ValueError(f"{name}: {stats['error']}")
                    if modified:
                        replacements[name] = patched
            if verify_only:
                defects = unknown = 0
                for name, issues in diagnostics.items():
                    print(f"{name}: " + ", ".join(f"{k}={v}" for k, v in issues.items()))
                    defects += sum(v for k, v in issues.items() if not k.startswith("unknown_"))
                    unknown += sum(v for k, v in issues.items() if k.startswith("unknown_"))
                print(f"[검증] 결함 {defects}, 상속 추가 확인 {unknown}. 폰트 설치/글리프/렌더링 및 전체 스키마는 미검증.")
                return 2 if defects else 3 if unknown else 0
            # Validate every XML/rels before saving, including unchanged parts.
            for name in names:
                if name.endswith((".xml", ".rels")):
                    try:
                        _parse(replacements.get(name, zin.read(name)))
                    except (etree.Error, ValueError) as exc:
                        raise ValueError(f"{name}: {exc}") from exc
            fd, temp_path = tempfile.mkstemp(prefix=".korean_pptx_", suffix=".tmp", dir=os.path.dirname(destination))
            os.close(fd)
            with zipfile.ZipFile(temp_path, "w") as zout:
                zout.comment = zin.comment
                for item in zin.infolist():
                    zout.writestr(item, replacements.get(item.filename, zin.read(item.filename)))
        os.replace(temp_path, destination)
        temp_path = None
        print(f"[OK] {len(replacements)}개 파트 보정. 폰트: {'지정 폰트로 통일' if force_font else '기존 상속/폰트 보존'}; 테마: {'변경' if update_theme else '보존'}")
        print("저장은 완료되었으며, --verify와 PowerPoint 화면 확인은 별도입니다.")
        return 0
    except (OSError, ValueError, KeyError, zipfile.BadZipFile, etree.Error) as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 1
    finally:
        if temp_path is not None and os.path.exists(temp_path):
            os.unlink(temp_path)


def main():
    parser = argparse.ArgumentParser(description="한국어 PPTX/POTX 서식 보정: 기본값은 기존 폰트와 테마 보존")
    parser.add_argument("pptx_file")
    parser.add_argument("-o", "--output", help="출력 경로; 생략하면 원본 교체")
    parser.add_argument("-s", "--slide", type=int, help="표시 순서 기준 슬라이드 번호 (1부터)")
    parser.add_argument("-f", "--font", help="대상 문단 폰트 통일; 폰트 설치는 별도 확인")
    parser.add_argument("--all-text", action="store_true", help="외국어/빈 문단/숫자 필드와 공유 목록 스타일까지 한국어 정책 적용; 심볼 run 제외")
    parser.add_argument("--update-theme", action="store_true", help="공용 테마도 변경; -f 필수, 슬라이드 한정 모드 불가")
    parser.add_argument("--verify", action="store_true", help="읽기 전용 제한 범위 검증 (0 정상, 1 오류, 2 결함, 3 상속 미확인)")
    parser.add_argument("--no-masters", action="store_true", help="마스터/레이아웃 제외")
    args = parser.parse_args()
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass
    return process_pptx(args.pptx_file, args.output, args.slide, args.font,
                        args.verify, not args.no_masters, args.font is not None,
                        all_text=args.all_text, update_theme=args.update_theme)


if __name__ == "__main__":
    sys.exit(main())
