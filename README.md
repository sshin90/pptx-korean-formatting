# pptx-korean-formatting

PowerPoint(`.pptx`/`.potx`)에 한글이 포함될 때 필요한 언어 태그, 단어 단위 줄바꿈, 동아시아 폰트, 서식 상속, 네이티브 불릿, 편집 가능한 카드 도형 규칙을 제공하는 Agent Skill입니다.

초고속 CLI 후처리기(`scripts/fix_korean_pptx.py`)와 재사용 가능한 파이썬 헬퍼(`scripts/korean_pptx.py`)를 내장하여 슬라이드 편집 및 생성 작업을 최소한의 지연 시간으로 수행할 수 있습니다.

---

## 주요 기능

1. **원클릭 CLI 후처리기 (`scripts/fix_korean_pptx.py`)**:
   * python-pptx 설치 시 함께 설치되는 `lxml`만 사용하며, 네임스페이스 접두사(`mc:Ignorable`, `Requires` 등)를 원본 그대로 보존
   * 슬라이드/마스터(`txStyles`)/레이아웃/테마(`theme1.xml`) XML을 고속(0.1초 미만) 분석하여 `ko-KR`, `eaLnBrk="0"`, `<a:latin>`/`<a:ea>` 폰트 일괄 주입 (`-f` 지정 시 테마 폰트까지 일괄 교체)
   * OpenXML ISO/IEC 29500-1 스키마 순서(`latin` → `ea` → `hlinkClick`) 및 `pPr` 미존재 문단 자동 생성, 문단 레벨(`defRPr`) 서식 상속 엄격 준수
   * `presentation.xml` 기반 실제 슬라이드 순서(1-based) 지정 보정(`--slide`) 및 정밀 서식 진단(`--verify`) 지원
2. **파이썬 헬퍼 모듈 (`scripts/korean_pptx.py`)**:
   * `set_korean_font_and_lang`: 영문(`latin`)과 동아시아(`ea`) 폰트 및 언어 동시 설정 (Run, Paragraph, Font 모두 지원, 스키마 순서 준수)
   * `set_native_bullet`: OpenXML 스키마 규격을 준수한 네이티브 불릿 및 들여쓰기
   * `create_korean_card` / `create_korean_shape`: 배경과 텍스트가 일체화된 단일 카드/도형 생성 (다단 레벨 색상 상속 지원)
   * `apply_korean_formatting`: 슬라이드, 도형뿐만 아니라 **표(Table), 그룹 도형(GroupShape), 전체 프레젠테이션(`prs`)**까지 재귀 일괄 적용
3. **핵심 원칙 가이드 (`SKILL.md`)**:
   * 에이전트 토큰 소모를 줄인 초경량 Fast-Path 중심 문서화

---

## 빠른 사용법

### 1. CLI로 기존 PPTX 일괄 보정
```bash
# 전체 슬라이드, 표, 마스터, 테마 보정 (in-place)
python scripts/fix_korean_pptx.py presentation.pptx

# 특정 폰트로 슬라이드/마스터/테마 폰트까지 일괄 변경
python scripts/fix_korean_pptx.py presentation.pptx -f "Pretendard"

# 특정 슬라이드만 고속 보정 (1-based 인덱스)
python scripts/fix_korean_pptx.py presentation.pptx --slide 2

# 서식 누락 항목 정밀 진단
python scripts/fix_korean_pptx.py presentation.pptx --verify
```

### 2. 파이썬 스크립트에서 모듈 사용
```python
from pptx import Presentation
from pptx.util import Inches
from scripts.korean_pptx import create_korean_card, set_native_bullet, apply_korean_formatting

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])

# 일체형 카드 생성 (Inches 또는 Pt 단위 사용)
card = create_korean_card(
    slide, left=Inches(1.0), top=Inches(1.0), width=Inches(5.0), height=Inches(3.0)
)
card.text_frame.text = "카드 본문 내용입니다."

# 전체 프레젠테이션(표, 도형 포함) 원클릭 서식 일괄 보정
apply_korean_formatting(prs, font_name="Pretendard")
prs.save("output.pptx")
```

---

## 디렉터리 구조

```
pptx-korean-formatting/
├── SKILL.md                          # 에이전트 참조용 핵심 지침 (경량화)
├── README.md                         # 저장소 소개 및 CLI 사용 가이드
├── scripts/
│   ├── __init__.py
│   ├── fix_korean_pptx.py            # 고속 XML 후처리기 (lxml 기반)
│   └── korean_pptx.py                # python-pptx용 헬퍼 모듈
├── references/
│   ├── openxml_reference.md          # OpenXML 스키마 상세 레퍼런스
│   └── pptxgenjs.md                  # PptxGenJS (Node.js) 가이드
└── tests/
    └── test_korean_pptx.py           # 단위 및 회귀 테스트 스위트
```

---

## 라이선스

[CC0 1.0 Universal](LICENSE) — 퍼블릭 도메인, 저작자 표시 의무 없음.
