---
name: pptx-korean-formatting
description: "Supplementary Korean-language PowerPoint formatting rules and fast CLI tools. Use whenever a .pptx, .potx, or slide contains Korean text. Enforces ko-KR language metadata, word-boundary wrapping (eaLnBrk=0), East Asian fonts, native hanging bullets, and single-shape cards. Includes an instant one-click CLI fixer."
license: Personal use
metadata:
  compatibility: "Works with python-pptx, pure OpenXML, or any PowerPoint generator."
---

# PPTX 한국어 서식 규칙 (Korean PPTX Formatting)

파워포인트에 한글이 포함될 때 글자 쪼개짐, 폰트 깨짐(굴림/맑은 고딕 강제 폴백), 불릿 들여쓰기 붕괴를 방지하기 위한 핵심 지침과 고속 보정 도구를 제공합니다.

---

## ⚡ 빠른 시작 (Fast-Path: 가장 빠른 적용법)

### 1. [초고속 권장] 편집 완료 후 원클릭 CLI 보정 (0.1초 소요)
일반적인 방식으로 슬라이드 텍스트나 도형, 표를 수정한 뒤 제공된 CLI 도구를 실행하면 모든 한글 서식이 0.1초 만에 자동 보정됩니다.

```bash
# 전체 슬라이드, 표, 마스터 일괄 보정 (in-place)
python <skill-path>/scripts/fix_korean_pptx.py presentation.pptx

# 특정 슬라이드만 고속 보정 (프레젠테이션 표시 순서 1-based 매핑)
python <skill-path>/scripts/fix_korean_pptx.py presentation.pptx --slide 2

# 특정 폰트로 지정 보정
python <skill-path>/scripts/fix_korean_pptx.py presentation.pptx -f "Pretendard"

# 서식 누락 여부 정밀 진단
python <skill-path>/scripts/fix_korean_pptx.py presentation.pptx --verify
```

> **장점**: 에이전트가 긴 헬퍼 함수를 매번 만들거나 복잡한 XML을 파싱할 필요가 없어 **작업 시간이 80% 이상 단축**되며, 표/그룹도형/하이퍼링크 스키마 규격까지 100% 자동 준수됩니다.

### 2. 파이썬(`python-pptx`) 신규 생성 시 모듈 임포트
신규 슬라이드를 직접 생성할 때는 스킬에 내장된 헬퍼 모듈을 임포트하여 사용합니다.

```python
import sys
# sys.path.insert(0, "<skill-path>")  # 필요 시 추가
from scripts.korean_pptx import (
    set_korean_font_and_lang,   # 언어(ko-KR)와 폰트(latin + ea) 동시 지정
    set_word_break_prevention,  # 단어 단위 줄바꿈 (eaLnBrk="0")
    set_native_bullet,          # 스키마 순서가 보장된 네이티브 불릿
    create_korean_card,         # 일체형 단일 카드 도형
    apply_korean_formatting     # prs/slide/table/shape 재귀 일괄 적용
)

# 프레젠테이션 전체(표 및 그룹도형 포함) 원클릭 일괄 적용:
apply_korean_formatting(prs, font_name="Pretendard")
```

---

## 📋 핵심 서식 4대 원칙

### 1. 언어 태그 & 단어 단위 줄바꿈
* **원칙**: 모든 텍스트에 `ko-KR` 언어를 지정하고, 단어 중간 줄바꿈 속성을 끕니다.
* **설정**: 텍스트 상자 `word_wrap = True`를 유지하면서 문단에 `eaLnBrk="0"`, `latinLnBrk="0"` 지정.
* *CLI 스크립트(`fix_korean_pptx.py`)가 pPr이 없는 기본 문단까지 자동으로 신규 생성하여 주입합니다.*

### 2. 동아시아 폰트(East Asian Font) 동기화
* **원칙**: `font.name`만 지정하면 영문 폰트(`<a:latin>`)만 설정되어 한글이 **시스템 기본 폰트(맑은 고딕 등)로 강제 폴백**됩니다.
* **설정**: `<a:latin>`과 `<a:ea>`를 동일한 폰트(예: Pretendard)로 반드시 쌍을 맞추어 지정합니다. 하이퍼링크 등이 있어도 스키마 규격 순서(`latin` → `ea` → `hlinkClick`)를 엄격 준수합니다.

### 3. 네이티브 불릿 및 내어쓰기 (Hanging Indent)
* **원칙**: 텍스트에 유니코드 기호(`•`, `-`)를 하드코딩하면 둘째 줄 들여쓰기가 깨집니다.
* **설정**: PowerPoint 네이티브 불릿(`buChar`, `marL`, `indent`)을 사용합니다.
* **주의**: ISO/IEC 29500-1 스키마 규격상 `<a:pPr>` 내부에서 `<a:buChar>`는 반드시 `<a:defRPr>` **앞에** 위치해야 파일 복구 경고가 발생하지 않습니다 (`set_native_bullet` 함수 사용).

### 4. 카드 UI는 단일 도형으로 통합 (Single Shape)
* **원칙**: 배경 사각형 위에 별도의 텍스트 상자를 얹으면 사용자가 수정할 때 배경과 글자가 분리됩니다.
* **설정**: 단일 사각형(`MSO_SHAPE.RECTANGLE`)에 직접 텍스트를 넣고, `vertical_anchor = MSO_ANCHOR.TOP`과 내부 마진(`margin_left` 등)으로 여백을 조정합니다 (`create_korean_card` 함수 사용).

---

## 🛠️ 작업 유형별 권장 워크플로

| 작업 유형 | 권장 워크플로 | 소요 시간 |
|---|---|---|
| **기존 슬라이드 수정 (1~2장)** | 1. `python-pptx` 등으로 평소처럼 텍스트/표 수정<br>2. `python <skill-path>/scripts/fix_korean_pptx.py deck.pptx --slide N` 실행 | **10~20초** |
| **전체 덱 일괄 정리 (표/도형 포함)** | `python <skill-path>/scripts/fix_korean_pptx.py deck.pptx` 1회 실행 | **1~2초** |
| **신규 슬라이드/카드 생성** | `scripts/korean_pptx.py`의 `create_korean_card`, `apply_korean_formatting` 활용 | **30초** |
| **서식 준수 여부 진단** | `python <skill-path>/scripts/fix_korean_pptx.py deck.pptx --verify` 실행 | **0.1초** |

---

## 🔍 상세 레퍼런스
* **OpenXML 스키마 및 태그 규격**: [`references/openxml_reference.md`](references/openxml_reference.md)
* **PptxGenJS (Node.js/브라우저)**: [`references/pptxgenjs.md`](references/pptxgenjs.md)
