---
name: pptx-korean-formatting
description: "Supplementary Korean-language PowerPoint formatting rules and fast CLI tools. Use whenever a .pptx, .potx, or slide contains Korean text. Enforces ko-KR language metadata, word-boundary wrapping (eaLnBrk=0), East Asian fonts, formatting inheritance (master/paragraph > run), native hanging bullets, and single-shape cards. Includes an instant one-click CLI fixer."
license: CC0-1.0
metadata:
  compatibility: "Works with python-pptx, pure OpenXML, or any PowerPoint generator."
---

# PPTX 한국어 서식 규칙 (Korean PPTX Formatting)

파워포인트에 한글이 포함될 때 글자 쪼개짐, 폰트 깨짐(굴림/맑은 고딕 강제 폴백), 서식 상속 단절, 불릿 들여쓰기 붕괴를 방지하기 위한 핵심 지침과 고속 보정 도구를 제공합니다.

---

## ⚡ 빠른 시작 (Fast-Path: 가장 빠른 적용법)

### 1. [초고속 권장] 편집 완료 후 원클릭 CLI 보정 (0.1초 소요)
일반적인 방식으로 슬라이드 텍스트나 도형, 표를 수정한 뒤 제공된 CLI 도구를 실행하면 모든 한글 서식이 0.1초 만에 자동 보정됩니다.

```bash
# 전체 슬라이드, 표, 마스터 일괄 보정 (in-place)
python <skill-path>/scripts/fix_korean_pptx.py presentation.pptx

# 특정 슬라이드만 고속 보정 (프레젠테이션 표시 순서 1-based 매핑)
python <skill-path>/scripts/fix_korean_pptx.py presentation.pptx --slide 2

# 특정 폰트로 지정 보정 (슬라이드/마스터 및 테마 폰트까지 일괄 교체)
python <skill-path>/scripts/fix_korean_pptx.py presentation.pptx -f "Pretendard"

# 서식 누락 여부 정밀 진단 (문단 레벨 상속 포함 검증)
python <skill-path>/scripts/fix_korean_pptx.py presentation.pptx --verify
```

> **장점**: 에이전트가 긴 헬퍼 함수를 매번 만들거나 복잡한 XML을 파싱할 필요가 없어 **작업 시간이 80% 이상 단축**되며, 마스터(`txStyles`)/레이아웃/테마(`theme1.xml`)/표/그룹도형/하이퍼링크 스키마 규격까지 100% 자동 준수됩니다.

### 2. 파이썬(`python-pptx`) 신규 생성 시 모듈 임포트
신규 슬라이드를 직접 생성할 때는 스킬에 내장된 헬퍼 모듈을 임포트하여 사용합니다.

```python
import sys
# sys.path.insert(0, "<skill-path>")  # 필요 시 추가
from scripts.korean_pptx import (
    set_korean_font_and_lang,   # 언어(ko-KR)와 폰트(latin + ea) 동시 지정 (문단/run 모두 지원)
    set_word_break_prevention,  # 단어 단위 줄바꿈 (eaLnBrk="0")
    set_native_bullet,          # 스키마 순서가 보장된 네이티브 불릿
    create_korean_card,         # 일체형 단일 카드/도형 (사각형, 원형 OVAL 등 지원)
    create_korean_shape,        # 단일 도형 생성 헬퍼 별칭
    apply_korean_formatting     # prs/slide/table/shape 재귀 일괄 적용
)

# 프레젠테이션 전체(표 및 그룹도형 포함) 원클릭 일괄 적용:
apply_korean_formatting(prs, font_name="Pretendard")
```

---

## 📋 핵심 서식 5대 원칙

### 1. 언어 태그 & 단어 단위 줄바꿈 (Language Tag & Word Wrap)
* **원칙**: 모든 텍스트에 `ko-KR` 언어를 지정하고, 단어 중간 줄바꿈 속성을 끕니다. PowerPoint에서 "한글 단어 잘림 허용"을 꺼도 언어가 `en-US`이면 영어 기준 줄바꿈이 적용되어 한글 단어가 깨집니다.
* **설정**: 텍스트 상자 `word_wrap = True`를 유지하면서 문단에 `eaLnBrk="0"`, `latinLnBrk="0"` 지정.
* *CLI 스크립트(`fix_korean_pptx.py`)가 pPr이 없는 기본 문단까지 자동으로 신규 생성하여 주입합니다.*

### 2. 동아시아 폰트(East Asian Font) 동기화
* **원칙**: `font.name`만 지정하면 영문 폰트(`<a:latin>`)만 설정되어 한글이 **시스템 기본 폰트(맑은 고딕 등)로 강제 폴백**됩니다.
* **설정**: `<a:latin>`과 `<a:ea>`를 동일한 폰트(예: Pretendard)로 반드시 쌍을 맞추어 지정합니다. 하이퍼링크 등이 있어도 스키마 규격 순서(`latin` → `ea` → `hlinkClick`)를 엄격 준수합니다.

### 3. 서식 상속 (Formatting Inheritance: 마스터/문단 레벨 우선)
* **증상**: 글자 크기·색상·폰트를 개별 run마다 하드코딩하면, 사용자가 PowerPoint에서 "테마 색상 일괄 변경"이나 슬라이드 마스터 서식 수정을 할 때 적용되지 않고 막힙니다.
* **원칙**: 서식은 가능한 한 높은 수준에서 상속시킵니다: **슬라이드 마스터/레이아웃 > 문단(Paragraph) 레벨 > Run 레벨**.
  * 공통 서식(폰트, 크기, 색상)은 문단 레벨(`paragraph.font.*` 또는 `<a:defRPr>`)이나 레이아웃의 placeholder 서식을 계승합니다.
  * Run 레벨 하드코딩은 문단 내에서 부분적으로 다른 서식(예: 굵게 처리된 특정 키워드 하나)이 필요한 경우에만 예외(Override)로 한정합니다.
  * **템플릿 작업 주의**: placeholder에 텍스트를 채울 때 `text_frame.text = "..."`로 덮어쓰면 기존 placeholder 서식이 완전히 소실되므로, 기존 문단의 서식을 유지하며 텍스트를 교체해야 합니다.
* **도구별 상속 지정 방법**:
  * `python-pptx`: `paragraph.font.size` / `paragraph.font.color.rgb`를 먼저 설정하고, run은 `run.text` 위주로 설정.
  * `pptxgenjs`: 개별 `text` 배열마다 색상을 반복하지 말고, `addText`의 최상위 `options`에 공통 `color`/`fontFace`/`fontSize`를 지정한 뒤 예외 항목에만 개별 override 적용.
  * 직접 XML: `<a:pPr>` 안의 `<a:defRPr>`에 기본 서식을 지정하고 개별 `<a:rPr>`에는 꼭 필요한 예외만 남김.

### 4. 네이티브 불릿 및 내어쓰기 (Hanging Indent)
* **원칙**: 텍스트에 유니코드 기호(`•`, `-`)를 하드코딩하면 둘째 줄 들여쓰기가 깨집니다.
* **설정**: PowerPoint 네이티브 불릿(`buChar`, `marL`, `indent`)을 사용합니다.
* **주의**: ISO/IEC 29500-1 스키마 규격상 `<a:pPr>` 내부에서 `<a:buChar>`는 반드시 `<a:defRPr>` **앞에** 위치해야 파일 복구 경고가 발생하지 않습니다 (`set_native_bullet` 함수 사용).

### 5. 카드 및 도형 UI는 단일 도형으로 통합 (Single Shape)
* **원칙**: 배경 도형 위에 별도의 투명 텍스트 상자를 얹으면 사용자가 수정할 때 배경과 글자가 분리됩니다.
* **설정**: 사각형(`RECTANGLE`), 원형(`OVAL`), 둥근 사각형 등 단일 도형에 직접 텍스트를 넣고 여백과 정렬을 조정합니다 (`create_korean_card` / `create_korean_shape` 활용).
  * 원형/타원형: `shape_type=MSO_SHAPE.OVAL` (자동 세로 중앙 정렬 `MIDDLE` 적용)
  * 직사각형 카드: `shape_type=MSO_SHAPE.RECTANGLE` (세로 상단 `TOP` 및 문단별 독립 서식 적용)
  * 글자색: 기본 도형 스타일(흰 글자·그림자)을 제거하고 테마 텍스트 색을 상속합니다. 어두운 배경이면 `text_color=RGBColor(...)`를 지정합니다.
* **다중 도형 예외 기준**:
  * 단순 서식 차이(예: 큰 제목 + 작은 본문)는 단일 도형의 여러 문단(`paragraphs`)으로 충분히 구현 가능하므로 단일 도형을 유지합니다.
  * 단, **좌우 분할 배치, 상하 양끝 분할, 아이콘/버튼 칩 결합** 등 독립적인 2차원 공간 배치가 필요한 경우에만 도형을 분리한 뒤 그룹화(`GroupShape`)합니다.

---

## 🛠️ 작업 유형별 권장 워크플로

| 작업 유형 | 권장 워크플로 | 소요 시간 |
|---|---|---|
| **기존 슬라이드 수정 (1~2장)** | 1. `python-pptx` 등으로 평소처럼 텍스트/표 수정<br>2. `python <skill-path>/scripts/fix_korean_pptx.py deck.pptx --slide N` 실행 | **10~20초** |
| **전체 덱 일괄 정리 (표/도형 포함)** | `python <skill-path>/scripts/fix_korean_pptx.py deck.pptx` 1회 실행 | **1~2초** |
| **신규 슬라이드/도형 생성** | `scripts/korean_pptx.py`의 `create_korean_shape`, `apply_korean_formatting` 활용 | **30초** |
| **서식 준수 여부 진단** | `python <skill-path>/scripts/fix_korean_pptx.py deck.pptx --verify` 실행 | **0.1초** |

---

## ✅ 작업 완료 후 최종 점검 체크리스트 (QA)

- [ ] 한글이 포함된 모든 문단/런에 `lang="ko-KR"` 및 단어 단위 줄바꿈(`eaLnBrk="0"`, `latinLnBrk="0"`)이 적용되었는가?
- [ ] 영문 폰트(`latin`)와 동아시아 폰트(`ea`)가 동일하게 지정되어 한글 폰트 폴백(맑은 고딕 등)이 방지되었는가?
- [ ] 폰트 크기·색상이 run마다 불필요하게 하드코딩되지 않고 문단/마스터 레벨에서 상속되는가?
- [ ] 문자열 기호(`•`) 대신 네이티브 불릿(`buChar`, `marL`, `indent`)과 내어쓰기를 적용했는가?
- [ ] 카드/원형 UI에서 배경과 텍스트가 분리되지 않고 단일 도형의 `text_frame`으로 통합되었는가? (복합 레이아웃 제외)

---

## 🔍 상세 레퍼런스
* **OpenXML 스키마 및 태그 규격**: [`references/openxml_reference.md`](references/openxml_reference.md)
* **PptxGenJS (Node.js/브라우저)**: [`references/pptxgenjs.md`](references/pptxgenjs.md)
