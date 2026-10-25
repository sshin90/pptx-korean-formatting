---
name: pptx-korean-formatting
description: "Supplementary Korean PowerPoint formatting for creating or editing PPTX/POTX files with Korean text. Preserve existing fonts and inheritance, control Korean word wrapping, create native bullets and editable single-shape cards, and run scoped XML checks. Includes python-pptx helpers and an optional CLI postprocessor."
license: CC0-1.0
metadata:
  compatibility: "Python 3.10+ and lxml for CLI; python-pptx for helpers. Transitional PPTX/POTX."
---

# PPTX 한국어 서식 및 템플릿 설계

기본 PPTX 제작 도구에 보완 적용한다. 한국어 줄바꿈, 동아시아 폰트, 서식 상속, 불릿 들여쓰기, 도형 편집성을 다룬다. 기존 템플릿의 폰트·테마·마스터 레이아웃과 사용자가 선택한 디자인을 먼저 보존한다.

## 작업 선택

| 목적 | 적용 |
|---|---|
| 기존 파일의 한국어 언어·줄바꿈 보정 | CLI 기본 모드. 기존 폰트와 테마는 유지 |
| 지정 폰트로 통일 | `-f "폰트명"`. 한국어가 있는 문단의 기본 폰트 설정 |
| 공용 테마까지 통일 | `-f "폰트명" --update-theme`. 전체 문서에만 사용 |
| 숫자 필드·외국어·빈 문단·공유 목록 스타일도 한국어 정책 적용 | 의도된 전체 통일 작업일 때만 `--all-text` |
| 마스터 기반 신규 슬라이드 제작 | 역할 기반 레이아웃 매핑, 플레이스홀더 주입, 열린 캔버스 배치 |
| 신규 카드·도형·불릿 생성 | Python 헬퍼 또는 제작 라이브러리의 네이티브 기능 |
| 생성된 파일 점검 | `--verify`와 PowerPoint 화면 확인 |

`<skill-path>`는 이 SKILL.md가 있는 실제 폴더로 치환한다. 경로에 공백이 있으면 따옴표로 감싼다.

```bash
# 기본: 한국어 문단 보정, 폰트/테마 보존. 검토용 출력 파일 생성
python "<skill-path>/scripts/fix_korean_pptx.py" input.pptx -o output.pptx

# 표시 순서 기준 2번 슬라이드만 지정 폰트로 통일
python "<skill-path>/scripts/fix_korean_pptx.py" input.pptx --slide 2 -f "Pretendard" -o output.pptx

# 읽기 전용 점검
python "<skill-path>/scripts/fix_korean_pptx.py" output.pptx --verify

# 지정 폰트까지 확인
python "<skill-path>/scripts/fix_korean_pptx.py" output.pptx --verify -f "Pretendard"

# 사용자가 전체 문서의 테마/외국어/공유 스타일 통일을 원할 때
python "<skill-path>/scripts/fix_korean_pptx.py" input.pptx -f "Pretendard" --all-text --update-theme -o output.pptx
```

`-o`를 생략하면 원본을 교체한다. XML 오류나 저장 실패 시 원본/기존 출력은 보존한다. `--slide`는 공유 마스터·레이아웃·테마를 수정하지 않는다. `--update-theme`은 `--slide`, `--no-masters`와 병용할 수 없고 `-f`가 필요하다. 테마 변경은 다른 언어와 테마를 공유하는 요소의 표시에도 영향을 줄 수 있다.

## 핵심 원칙

### 1. 한국어 언어와 줄바꿈

한국어 텍스트에는 `ko-KR`을 적용하고, 단어 단위 줄바꿈이 필요한 문단에 `eaLnBrk="0"`, `latinLnBrk="0"`을 사용한다. 텍스트 상자의 자동 줄바꿈은 Python에서 `word_wrap=True`, XML에서 `a:bodyPr wrap="square"`다.

CLI는 한국어를 포함한 문단과 그 텍스트 프레임에 적용한다. 같은 프레임의 다른 문단도 프레임 줄바꿈 변경의 영향을 받는다. 외국어 전용 문단과 숫자 필드의 언어는 기본 보존하며, `--all-text`는 의도된 예외다. 한 run 안의 한영 혼합 문자열은 run 단위로 처리한다. 긴 어절·상자 높이·폰트 대체에 의한 넘침은 화면에서 확인한다.

### 2. 동아시아 폰트와 상속

`python-pptx`의 `font.name`은 Latin 폰트를 설정하므로, 신규 생성에서 동아시아 폰트를 직접 지정할 때는 `a:ea`도 설정한다. 다만 `a:ea`가 없다는 이유만으로 오류라고 판단하지 않는다. 문단·목록 스타일·레이아웃·마스터·테마에서 상속할 수 있다.

기존 폰트 이름만으로 한글 지원 여부를 추정해 교체하지 않는다. 의도된 영문/한글 폰트 조합은 허용한다. 폰트 설치·글리프·임베딩 여부는 별도 확인하며, 맑은 고딕도 정상적인 선택이다.

공통 서식은 가능한 한 상위 스타일에서 상속한다. 개별 run에는 부분 강조 등 필요한 예외만 둔다. 테마 변경을 따라야 하는 색상은 고정 RGB 대신 테마 색 참조를 사용한다. 문단에 고정 폰트나 RGB를 지정해도 그 속성은 마스터/테마의 변경을 덮어쓴다.

CLI 기본 모드는 없는 run 폰트를 생성하지 않는다. `-f`는 문단 기본 Latin/EA 폰트를 지정하고 대상 한국어 run의 폰트 override를 제거한다. 굵기·색상·하이퍼링크와 외국어 run의 명시적 폰트는 유지한다. 혼합 문단에서 폰트가 없는 영문 run도 새 문단 기본 폰트를 상속한다. 심볼 폰트가 직접 지정된 run은 제외한다. 기존의 run 폰트를 임의로 상속형으로 전환하지 않으므로, 기본 보정만으로 과거의 직접 서식이 제거되지는 않는다.

`text_frame.text = "..."`는 로컬 문단/run 내용을 교체해 그 서식을 잃을 수 있다. placeholder의 레이아웃·마스터 관계까지 사라지는 것은 아니다. 부분 수정에는 기존 run의 텍스트 교체 등 보존 가능한 방법을 선택한다.

### 3. 네이티브 불릿

문자열 앞에 `•`를 붙이기보다 네이티브 불릿과 `marL`/`indent`를 사용한다. `a:pPr` 안에서 불릿 요소는 `tabLst`/`defRPr`보다 앞에 있어야 하며, 불릿 종류 choice는 하나만 둔다. Python에서는 `set_native_bullet()`을 사용한다.

### 4. 편집 가능한 단일 도형

단순 카드·원형·배지는 도형 자체의 텍스트 프레임을 사용한다. 제목/본문은 문단별 서식으로 구분한다. 좌우 분할, 상하 양끝 배치, 독립 아이콘·버튼처럼 별도 위치 제어가 필요하면 여러 도형을 그룹화한다. 이는 편집성 기본 원칙이며 사용자의 배치 요구를 제한하지 않는다.

### 5. 슬라이드 마스터와 플레이스홀더 우선 상속 (Master & Placeholder Inheritance)

프레젠테이션의 일관성을 위해 공통 요소는 개별 슬라이드에서 텍스트 상자(`TextBox`)로 매번 새로 그리지 않고, 마스터 레이아웃의 플레이스홀더(`placeholders`)를 상속받아 텍스트를 주입한다.

- **헤더/푸터 일관성**: 슬라이드 제목, 부제목, 섹션명, 푸터, 슬라이드 번호는 레이아웃에 배치된 플레이스홀더(`idx=100, 101, 102` 등)를 통해 제어한다.
- **역할 기반(Role-based) 레이아웃 매핑**: 레이아웃은 특정 배경 색상에 종속되지 않고, 슬라이드의 **기능적 목적(Role)**에 따라 식별·선택한다.
  - `Cover`: 표지 및 피날레 엔딩
  - `Section / Divider`: 대주제 전환, 간지
  - `Standard Content`: 기본 본문
  - `Split / Comparison`: 2단 비교, 대비
  - `Full Media / Image`: 전면 시각 자료
  - `Break / Wrap-up`: 정리, 휴식 시간

### 6. 정보 관계 중심의 열린 캔버스 배치 (Open-ended Content-First Layout)

특정 템플릿 틀(카드, 표 등) 몇 개를 강제하여 배치의 다양성을 제한하지 않는다. 본문 공간은 **전달하려는 정보의 본질적 관계**에 따라 유연하게 구성한다.

- **박스 만능주의(Box-ism) 탈피**: 모든 내용을 카드나 박스 안에 가둘 필요는 없다. 박스는 성격이 다른 정보를 시각적으로 격리해야 할 필연적 이유가 있을 때만 절제하여 사용한다.
- **여백과 타이포그래피 위계**: 폰트 크기 대비(대형 키워드 32~44pt vs 본문 12~14pt)와 충분한 여백(White Space)만으로도 박스 없이 세련되고 자유로운 레이아웃을 만들 수 있다.
- **정보 관계에 따른 자유로운 공간 분할**:
  - **단일 집중/화두**: 단 하나의 핵심 질문, 강력한 인용구, 대형 지표(Big Stat)를 중앙에 여백과 함께 배치
  - **흐름/순서**: 시간 순서, 업무 프로세스, 타임라인을 가로 또는 세로 연결선으로 배치
  - **대비/비교**: 2분할(좌우/상하), Before & After, 대안 비교
  - **위계/포함**: 상하 트리, 중첩 원, 피라미드 구조
  - **실물/데이터 증거**: 화면 캡처, 제품 사진, 복잡한 표를 전면에 채우고 여백 캡션 활용
  - **구조화 정보**: 절차 안내, 체크리스트 등 다중 요소가 필요할 때 카드/표/다이어그램 활용

### 7. 도형 조형과 곡률 제어 (Shape Geometry & Radius Control)

둥근 모서리(`roundRect`)를 무조건 금지하지 않는다. 템플릿의 전체 디자인 언어와 용도에 맞추어 조형을 결정한다.

- **과도한 기본 곡률 주의**: PowerPoint 기본 `roundRect`는 도형 크기가 커질수록 곡률이 비례하여 타원형처럼 뭉툭해지는 현상(기본 `adj` 약 16.7%)이 있다.
- **곡률 미세 제어**: 라운드 카드를 제작할 때는 모던 UI 기준(반경 4~8px 수준, `adj` 2~4%)으로 미세 조정(`corner_radius_pct=0.03`)하여 단정함을 유지한다.
- **스타일 일관성**: 플랫하고 정갈한 비즈니스 문서에는 각진 직사각형(`MSO_SHAPE.RECTANGLE`)을, 부드러운 서비스 소개에는 제어된 둥근 모서리를 템플릿 전반에 일관되게 적용한다.

## 마스터 템플릿 기반 슬라이드 생성 절차 (Workflow)

신규 프레젠테이션을 제작하거나 기존 템플릿을 기반으로 개편할 때 다음 단계를 준수한다.

1. **템플릿 정제**: 마스터 슬라이드와 레이아웃이 포함된 베이스 파일에서 잔여 슬라이드(`ppt/slides/slide*.xml`)와 관계를 정리하고 깨끗한 빈 템플릿을 준비한다.
2. **역할 기반 레이아웃 선택**: 슬라이드의 목적에 맞는 레이아웃을 선택하여 `prs.slides.add_slide(layout)`을 호출한다.
3. **플레이스홀더 주입**: 슬라이드의 `placeholders`를 순회하며 제목/부제/섹션명 텍스트를 채운다.
4. **정보 관계 기반 본문 배치**: 정보의 성격(단일 강조, 프로세스, 비교, 전면 이미지, 구조화 카드 등)에 맞춰 캔버스를 도화지처럼 자유롭게 설계·배치한다.
5. **한국어 서식 보정 및 검증**: `apply_korean_formatting()` 또는 CLI 도구(`fix_korean_pptx.py -f "폰트명" --verify`)를 통해 폰트 일치, 한국어 줄바꿈, 상속 결함 유무를 최종 확인한다.

## Python 헬퍼

스킬 폴더를 Python 모듈 검색 경로에 추가한 뒤 사용한다. 기존 템플릿에는 `font_name` 생략을 기본으로 한다.

```python
import sys
sys.path.insert(0, "<skill-path>")
from scripts.korean_pptx import (
    apply_korean_formatting, set_korean_font_and_lang,
    set_word_break_prevention, set_native_bullet,
    create_korean_card, create_korean_shape,
)

# 1. 기존 프레젠테이션 한국어 서식 보정
apply_korean_formatting(prs)  # 기존 폰트와 상속 보존
apply_korean_formatting(slide, font_name="Pretendard")  # 폰트 통일 시

# 2. 각진 직사각형 카드 생성 (플랫 스타일)
card = create_korean_card(slide, left, top, width, height, shape_type=MSO_SHAPE.RECTANGLE)

# 3. 미세 곡률 제어 카드 생성 (소프트 스타일, 과도한 라운드 방지)
round_card = create_korean_card(
    slide, left, top, width, height,
    shape_type=MSO_SHAPE.ROUNDED_RECTANGLE,
    corner_radius_pct=0.03  # 3% 미세 라운드 (약 4~8px)
)
```

헬퍼는 Presentation/Slide/Shape/Table/GroupShape/TextFrame을 순회한다. 마스터·레이아웃·테마는 수정하지 않는다. `set_korean_font_and_lang(paragraph, font_name)`은 명시적 문단 폰트 설정이며 기존 run 폰트 예외를 제거하지 않는다. 새 빈 카드에 텍스트를 채운 뒤 언어·폰트·줄바꿈을 적용한다.

## 검증과 완료 기준

`--verify` 종료 코드:

| 코드 | 의미 |
|---|---|
| 0 | 검사한 XML 범위에서 결함·상속 미확인 없음 |
| 1 | XML/ZIP/옵션/입출력 오류 |
| 2 | 명시적 서식 결함 발견 |
| 3 | 결함은 없으나 상위 상속 확인 필요 |

일반 보정 실행의 0은 저장 성공을 의미하며 검증 완료가 아니다. 코드 3을 없애기 위해 불필요하게 폰트를 하드코딩하지 않는다. 해당 템플릿의 마스터·테마를 확인하고 판단을 기록한다.

검증기는 문단과 같은 도형의 목록 스타일을 해석하고, 한국어 run/필드, 줄바꿈, 명시적 EA 토큰, 주요 폰트·불릿 요소 순서를 검사한다. 레이아웃/마스터/테마를 가로지르는 유효 폰트는 추정하지 않고 `unknown_*`로 보고한다. `-f`를 함께 주면 해당 폰트와의 일치도 확인한다. 임의 폰트의 글리프 지원, 전체 ISO 스키마, 화면 넘침, 단일 도형 여부까지 검증하지 않는다.

지원 범위는 Transitional PPTX/POTX의 슬라이드·마스터·레이아웃에 저장된 DrawingML 텍스트이며 표/그룹을 포함한다. 별도 차트·SmartArt·노트·임베디드 객체·프레젠테이션 기본 텍스트 스타일은 보정 대상이 아니다. CLI는 문자열 불릿 변환이나 카드 병합을 하지 않는다.

완료 전:
- 의도한 슬라이드/문단만 변경되었는지 확인한다.
- 한국어 줄바꿈, 장문 넘침, 불릿 내어쓰기, 폰트 표시를 PowerPoint에서 확인한다.
- 상속을 유지해야 하는 곳은 상위 서식을 바꾸어 하위가 따라오는지 확인한다.
- 자동 검사 결과와 앱에서 확인하지 못한 항목을 구분해 보고한다.

## 도구별 레퍼런스

- XML을 직접 편집할 때: [OpenXML 구조](references/openxml_reference.md)
- PptxGenJS로 제작할 때: [PptxGenJS 가이드](references/pptxgenjs.md)
