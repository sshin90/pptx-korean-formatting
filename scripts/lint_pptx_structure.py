#!/usr/bin/env python3
"""
PPTX 구조 린터:
1) 가짜 표 탐지: 텍스트 없는 넓은 배경 도형 위에 2개 이상의 텍스트 도형이 가로로 나란히 배치된 경우
2) 분할 불릿 탐지: 동일한 X좌표·너비의 불릿 텍스트 상자가 세로로 촘촘히(gap <= 0.15 in) 연속 배치된 경우

종료 코드:
  0: 구조 결함 없음
  2: 가짜 표 또는 분할 불릿 의심 요소 발견
"""
import argparse
import sys
from pptx import Presentation
from pptx.oxml.ns import qn


def _shape_text(s):
    return s.text_frame.text.strip() if s.has_text_frame else ''


def _has_bullet_xml(p):
    pPr = p._p.find(qn('a:pPr'))
    if pPr is None:
        return False
    return any(pPr.find(qn(tag)) is not None for tag in ('a:buChar', 'a:buAutoNum', 'a:buBlip'))


def detect_fake_table_rows(slide):
    """Detect wide empty solid rectangles that host >=2 horizontal text shapes."""
    all_shapes = list(slide.shapes)
    suspect_rows = []
    for bg in all_shapes:
        if bg.has_table:
            continue
        w = bg.width.inches
        h = bg.height.inches
        if w < 8.0 or not (0.22 <= h <= 0.90) or _shape_text(bg):
            continue
        try:
            if bg.fill.type != 1:
                continue
        except Exception:
            continue

        bg_t, bg_h, bg_l, bg_r = bg.top.inches, h, bg.left.inches, bg.left.inches + w
        cells = []
        has_badge = False
        for s in all_shapes:
            if s is bg:
                continue
            st, sh, sl = s.top.inches, s.height.inches, s.left.inches
            sr = sl + s.width.inches
            if abs(st - bg_t) <= 0.12 and sh <= bg_h + 0.08 and sl >= bg_l - 0.05 and sr <= bg_r + 0.05:
                try:
                    if s.fill.type == 1 and s.width.inches < 0.8 and s.height.inches < 0.8:
                        has_badge = True
                except Exception:
                    pass
                if s.has_text_frame and _shape_text(s):
                    cells.append(s)

        if not has_badge and len(cells) >= 2:
            suspect_rows.append((round(bg_t, 2), len(cells)))
    return suspect_rows


def detect_split_bullets(slide):
    """Detect vertically stacked separate bullet text boxes with matching left/width."""
    candidates = []
    for s in slide.shapes:
        if s.has_table or not s.has_text_frame or not _shape_text(s):
            continue
        try:
            if s.fill.type == 1:
                continue
        except Exception:
            pass
        if _has_bullet_xml(s.text_frame.paragraphs[0]):
            candidates.append(s)

    groups = {}
    for s in candidates:
        key = (round(s.left.inches, 1), round(s.width.inches, 1))
        groups.setdefault(key, []).append(s)

    clusters = []
    for key, items in groups.items():
        items.sort(key=lambda s: s.top)
        cur = [items[0]]
        for item in items[1:]:
            prev_s = cur[-1]
            gap = item.top.inches - (prev_s.top.inches + prev_s.height.inches)
            if gap <= 0.15:
                cur.append(item)
            else:
                if len(cur) >= 2:
                    clusters.append((key, len(cur), round(cur[0].top.inches, 2)))
                cur = [item]
        if len(cur) >= 2:
            clusters.append((key, len(cur), round(cur[0].top.inches, 2)))
    return clusters


def main():
    parser = argparse.ArgumentParser(description='Lint PPTX for fake tables and split bullet text boxes.')
    parser.add_argument('pptx_path', help='Path to .pptx file')
    args = parser.parse_args()

    prs = Presentation(args.pptx_path)
    issues = 0

    for si, slide in enumerate(prs.slides, 1):
        fake_rows = detect_fake_table_rows(slide)
        if len(fake_rows) >= 2:
            issues += 1
            print(f"[가짜 표 의심] Slide {si}: 행 배경 사각형 {len(fake_rows)}개 (예: T={fake_rows[0][0]}\", 열={fake_rows[0][1]}개)")

        split_bullets = detect_split_bullets(slide)
        for key, cnt, top in split_bullets:
            issues += 1
            print(f"[분할 불릿 의심] Slide {si}: L={key[0]}\", W={key[1]}\", T={top}\" 위치에 개별 불릿 상자 {cnt}개 분리")

    if issues == 0:
        print(f"[구조 검증 통과] 가짜 표 및 분할 불릿 결함 0건 ({len(prs.slides)} slides checked)")
        sys.exit(0)
    else:
        print(f"[구조 검증 경고] 총 {issues}건의 구조 문제 발견")
        sys.exit(2)


if __name__ == '__main__':
    main()
