---
name: pptx-korean-formatting
description: "Korean-language PowerPoint (.pptx) generation rules that must be applied ON TOP OF the standard pptx/pptxgenjs/python-pptx workflow whenever a slide deck contains Korean text. Trigger this whenever creating or editing a .pptx, .potx, or html2pptx-based deck that includes any Korean (한글) content — even if the user's request doesn't explicitly mention 'Korean' or 'lang'. This covers: (1) forcing Korean language tags (lang=\"ko-KR\" / msoLanguageIDKorean) on every text run/paragraph so PowerPoint's Korean word-wrap engine is used instead of the English one, and (2) pushing font/color formatting up to paragraph, layout, or master level instead of hardcoding it per-run, so downstream edits (e.g., bulk color changes in PowerPoint) aren't blocked. Always consult this skill together with the main pptx skill for any Korean-language deck."
license: MIT
---

# PPTX 한국어 생성 규칙 (Korean PPTX Generation Rules)

이 스킬은 표준 `pptx` 스킬(또는 `python-pptx`, `pptxgenjs`, `html2pptx` 기반 스크립트)과 **함께** 적용하는 보조 규칙이다. 슬라이드에 한글이 포함되는 모든 생성/편집 작업에서 아래 두 가지를 반드시 지킨다.

## 1. 언어 속성 (Language Tag)

**증상:** 언어 속성이 `en-US`로 남아있으면 PowerPoint가 영어 줄바꿈 규칙을 적용해 한글 단어가 음절 단위로 잘못 잘린다 (예: "안녕하세요"가 "안녕하" / "세요"로 분리).

**규칙:**
- 한글이 포함된 모든 텍스트 런(run)·문단(paragraph)에 언어를 한국어로 명시한다.
- 도구별 지정 방법:

| 도구 | 방법 |
|---|---|
| `pptxgenjs` | 텍스트 옵션에 `lang: "ko-KR"` 추가 (`{ text: "...", options: { lang: "ko-KR", ... } }`). 슬라이드 전체 기본값을 정하고 싶으면 공통 옵션 객체를 만들어 모든 `addText` 호출에 spread(`...`)로 병합한다. |
| `python-pptx` | run 객체는 `lang` 속성이 없으므로 XML 레벨에서 처리: `run.font.language_id = MSO_LANGUAGE_ID.KOREAN` (`from pptx.enum.lang import MSO_LANGUAGE_ID`). 또는 `run._r.rPr.set('lang', 'ko-KR')`로 직접 XML 속성을 설정한다. |
| 직접 XML 편집 (`ppt/slides/slideN.xml`) | 모든 `<a:rPr>` / `<a:defRPr>`에 `lang="ko-KR"` 속성을 추가한다. 예: `<a:rPr lang="ko-KR" sz="1800"/>` |
| `html2pptx` / HTML 중간 산출물 | 루트 요소 또는 텍스트 컨테이너에 `lang="ko-KR"`을 지정해 변환기가 속성을 이어받도록 한다. |

- 한글 텍스트가 잘리지 않도록 문단 속성에서 자동 줄바꿈을 활성화한다: `python-pptx`에서는 `text_frame.word_wrap = True`. pptxgenjs는 텍스트박스에 별도 `wrap` 옵션이 기본 True이므로, `wrap: false`로 끄지 않았는지 확인한다.

## 2. 서식 상속 (Formatting Inheritance)

**증상:** 글자 크기·색상을 run마다 하드코딩하면, 사용자가 PowerPoint에서 나중에 "테마 색상 일괄 변경" 같은 편집을 할 때 적용되지 않는다.

**규칙:**
- 폰트 크기·색상은 가능한 한 다음 우선순위로 올려서 지정한다: **슬라이드 마스터/레이아웃 > 문단(paragraph) 레벨 > run 레벨**. run 레벨 하드코딩은 문단 내에서 부분적으로 다른 서식(예: 굵게 처리된 키워드 하나)이 필요한 경우에만 예외적으로 허용한다.
- 도구별 적용 지점:

| 도구 | 문단/마스터 레벨 지정 방법 |
|---|---|
| `pptxgenjs` | 개별 `text` 배열 항목마다 색상을 반복 지정하지 말고, `addText`의 최상위 `options`에 공통 `color`/`fontFace`/`fontSize`를 지정한 뒤, 예외가 필요한 항목에만 개별 override를 준다. |
| `python-pptx` | `paragraph.font.size` / `paragraph.font.color.rgb`를 먼저 설정하고, run은 `run.text`만 채운다. 템플릿 기반 작업에서는 레이아웃의 placeholder 서식을 그대로 두고 텍스트만 교체한다 (`text_frame.text = "..."`로 덮어써서 서식을 날리지 말 것 — 표준 pptx 스킬의 QA 규칙과 동일). |
| 직접 XML 편집 | 색상·폰트는 `<a:pPr>` 안의 `<a:defRPr>`에 지정하고, 개별 `<a:rPr>`에는 꼭 필요한 예외만 남긴다. |

## 체크리스트 (생성/편집 완료 후)

- [ ] 한글이 포함된 모든 run/문단에 `lang="ko-KR"` (또는 해당 도구의 동등 속성)이 적용되었는가?
- [ ] 줄바꿈(word wrap)이 꺼져 있지 않은가?
- [ ] 폰트 크기·색상이 run마다 반복 하드코딩되지 않고 문단/마스터 레벨에서 상속되는가?
- [ ] (템플릿 기반 작업 시) 표준 `pptx` 스킬의 QA 절차(`validate.py`, 시각 QA)를 함께 통과했는가?
