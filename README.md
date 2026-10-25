# pptx-korean-formatting

한국어 PowerPoint의 언어·줄바꿈·EA 폰트·서식 상속·네이티브 불릿·단일 도형 카드 제작을 보완하는 Agent Skill입니다.

기존 템플릿을 위한 기본 모드는 **폰트와 테마 보존**입니다. 단순 누락 보정, 폰트 통일, 공용 테마 변경을 구분합니다. 세부 제작 지침은 [SKILL.md](SKILL.md)를 참고하세요.

## 요구 환경

- Python 3.10 이상, CLI는 `lxml` 필요
- 생성 헬퍼와 테스트는 `python-pptx` 필요
- 검증 환경: Python 3.12, python-pptx 1.0.2, lxml 6.1.1, PptxGenJS 4.0.1
- Transitional `.pptx`/`.potx`; 폰트는 사용 환경에 별도 설치

## CLI

```bash
python scripts/fix_korean_pptx.py input.pptx -o output.pptx
python scripts/fix_korean_pptx.py input.pptx --slide 2 -f "Pretendard" -o output.pptx
python scripts/fix_korean_pptx.py output.pptx --verify
python scripts/fix_korean_pptx.py output.pptx --verify -f "Pretendard"
```

- `-o` 생략: 원본 교체. XML 오류 또는 저장 실패 시 원본/기존 출력 보존.
- `-f`: 한국어 문단 기본 폰트 통일. 한국어 run의 폰트 override만 제거하고 색상·강조·링크 보존.
- `--slide N`: 표시 순서 기준 N번째 슬라이드만 변경. 공유 파트 불변.
- `--no-masters`: 마스터·레이아웃 제외.
- `--all-text`: 외국어·숫자 필드·빈 문단·공유 목록 스타일도 한국어 정책으로 통일. 심볼 폰트가 직접 지정된 run은 제외.
- `--update-theme`: `-f`와 함께 공용 테마 폰트까지 변경. `--slide`/`--no-masters`와 병용 불가.

전체 문서 통일이 의도된 경우:

```bash
python scripts/fix_korean_pptx.py input.pptx -f "Pretendard" --all-text --update-theme -o output.pptx
```

## Python

```python
from pptx import Presentation
from pptx.util import Inches
from scripts.korean_pptx import create_korean_card, apply_korean_formatting

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
card = create_korean_card(slide, Inches(1), Inches(1), Inches(5), Inches(3))
card.text_frame.text = "한국어 카드 본문"
apply_korean_formatting(prs, font_name="Pretendard")
prs.save("output.pptx")

# 기존 템플릿은 폰트를 생략하여 상속 보존
# apply_korean_formatting(prs)
```

## 검증

```bash
python -X utf8 -B -m unittest discover -s tests -v
```

`--verify`는 읽기 전용이며, 0=검사 범위 정상, 1=처리 오류, 2=서식 결함, 3=상속 미확인입니다. 보정 실행의 0은 저장 성공만 뜻합니다.

문단/도형 목록 스타일까지만 상속을 해석합니다. 상위 마스터·테마의 유효 폰트는 미확인으로 보고하며, 폰트 설치·글리프·전체 스키마·실제 화면은 별도 확인합니다. 차트·SmartArt·노트 등 별도 파트는 보정하지 않습니다. 도구는 문자열 불릿 변환이나 여러 도형의 카드 병합을 하지 않습니다.

## 이전 버전과 다른 점

- 기본 실행은 Pretendard를 자동 주입하거나 테마 폰트를 바꾸지 않습니다.
- 폰트 통일은 CLI의 `-f`, 헬퍼의 명시적 `font_name`으로 선택합니다.
- Python `process_pptx`/`patch_xml_content`에서는 `default_font`만 넘겨도 강제 통일되지 않습니다. `force_font=True`를 함께 지정합니다.
- 테마 변경은 `--update-theme`, 외국어·빈 문단·공유 목록 스타일 변경은 `--all-text`가 필요합니다.
- 상속을 확정하지 못한 검증은 0 대신 3을 반환합니다. 자동화에서 이 상태를 결함과 구분하세요.

[CC0 1.0 Universal](LICENSE)
