# PPTX 한국어 생성 규칙 (Korean PPTX Generation Rules)

> 이 문서는 특정 도구에 종속되지 않는 범용 규칙 파일이다. `python-pptx`, `pptxgenjs`, `html2pptx` 또는 직접 XML 편집으로 PowerPoint(.pptx/.potx)를 생성·수정하는 **모든 AI 코딩 에이전트**(Claude Code, OpenAI Codex, Google Antigravity, Cursor, Windsurf 등)가 한글이 포함된 슬라이드를 다룰 때 적용해야 한다. 각 에이전트에 연결하는 방법은 [README](README.md#설치--연결-방법)를 참고한다 (예: Codex는 이 파일을 `AGENTS.md`에 병합, Antigravity는 `~/.gemini/config/rules/`에 복사하거나 [SKILL.md](SKILL.md)를 스킬로 설치).

이 스킬은 표준 프레젠테이션 스킬과 **함께** 적용하는 보조 규칙이다. 먼저 해당 스킬이 지정한 생성 백엔드와 검증 절차를 따른다. 현재 Codex의 표준 프레젠테이션 작업은 `@oai/artifact-tool`을 사용하므로, 별도 승인이나 기존 프로젝트의 필요가 없는 한 `python-pptx`나 PptxGenJS로 대체하지 않는다. 아래 Python/PptxGenJS 예시는 그 백엔드를 명시적으로 사용하는 기존 워크플로에만 적용한다.

---

## 1. 언어 속성 (Language Tag) & 단어 단위 줄바꿈

**증상:** 언어 속성이 기본값(`en-US`)으로 남아있으면 PowerPoint가 영어 줄바꿈 규칙을 적용하여 한글 단어가 음절 단위로 뚝뚝 끊겨 잘린다 (예: "안녕하세요"가 "안녕하" / "세요"로 분리).

**규칙:**
- 한글이 포함된 모든 텍스트 런(run)과 문단(paragraph)에 언어를 한국어(`ko-KR` / 1042)로 명시한다.
- 텍스트 컨테이너의 자동 줄바꿈(`word_wrap`)을 활성화한다.
- 도구별 지정 방법:

| 도구 | 방법 |
|---|---|
| `python-pptx` (고수준 API) | • 개별 Run: `run.font.language_id = MSO_LANGUAGE_ID.KOREAN`<br>• 문단 전체 상속: `paragraph.font.language_id = MSO_LANGUAGE_ID.KOREAN`<br>(`from pptx.enum.lang import MSO_LANGUAGE_ID` 임포트) |
| `python-pptx` (저수준 XML 직접 제어) | `run._r.get_or_add_rPr().set('lang', 'ko-KR')`<br>*(주의: `run._r.rPr.set(...)`은 `rPr` 미생성 시 `AttributeError: 'NoneType'`을 발생시키므로 반드시 `get_or_add_rPr()` 사용)* |
| `pptxgenjs` | `{ text: "...", options: { lang: "ko-KR", ... } }`<br>공통 옵션 객체를 만들어 모든 `addText` 세그먼트 옵션에 spread(`...`)로 병합한다. |
| 직접 XML 편집 (`ppt/slides/slideN.xml`) | 모든 `<a:rPr>` / `<a:defRPr>`에 `lang="ko-KR"` 속성 추가. 줄바꿈 문단은 `<a:endParaRPr lang="ko-KR"/>` 추가. |
| `html2pptx` / HTML 중간 산출물 | 루트 요소 또는 텍스트 컨테이너에 `lang="ko-KR"` 명시. |

- `python-pptx` 줄바꿈 활성화: `text_frame.word_wrap = True`
- `pptxgenjs` 줄바꿈: 텍스트박스는 기본 활성화되어 있으므로 `wrap: false`로 끄지 않도록 주의.

---

## 2. 한글 폰트(East Asian Font) 명시

**증상:** `python-pptx`에서 `font.name = 'Pretendard'`를 지정하면 영문 폰트(`<a:latin>`)만 설정된다. OpenXML 규격상 동아시아 폰트(`<a:ea>`)가 설정되지 않으면, PowerPoint에서 파일을 열었을 때 영문/숫자만 Pretendard로 나오고 **한글은 시스템 기본 동아시아 폰트(맑은 고딕, 굴림 등)로 강제 폴백**되는 치명적인 서식 불일치가 발생한다.

**규칙:**
- 폰트 지정 시 서양 문자(`<a:latin>`)와 동아시아 문자(`<a:ea>`)를 반드시 둘 다 동일한 폰트로 지정한다.
- `python-pptx`에서는 아래 제공되는 `set_korean_font()` 헬퍼 함수를 사용하여 폰트를 지정한다.

```python
from pptx.oxml.xmlchemy import OxmlElement

def set_korean_font(font_target, font_name="Pretendard"):
    """
    Run 또는 Paragraph 객체에 라틴(latin)과 한글(ea: East Asian) 폰트를 동시에 설정.
    한글이 맑은 고딕 등으로 강제 폴백되는 현상을 완벽 차단한다.
    """
    font_target.font.name = font_name
    
    parent_elm = getattr(font_target, '_r', None)
    if parent_elm is None:
        parent_elm = getattr(font_target, '_p', None)
    if parent_elm is None:
        return
        
    prop_elm = parent_elm.get_or_add_rPr() if hasattr(parent_elm, 'get_or_add_rPr') else parent_elm.get_or_add_pPr().get_or_add_defRPr()
    
    ea = prop_elm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}ea')
    if ea is None:
        ea = OxmlElement('a:ea')
        prop_elm.append(ea)
    ea.set('typeface', font_name)
```

---

## 3. 서식 상속 (Formatting Inheritance)

**증상:** 글자 크기·색상을 run마다 하드코딩하면, PowerPoint에서 사용자가 "테마 색상 일괄 변경" 같은 레이아웃/스타일 일괄 편집을 할 때 적용되지 않고 무시된다.

**규칙:**
- 폰트 크기·색상은 가능한 한 다음 우선순위로 올려서 지정한다:
  **슬라이드 마스터/레이아웃 > 문단(paragraph) 레벨 > run 레벨**.
- run 레벨 하드코딩은 문단 내에서 부분 강조(예: 특정 키워드 볼드 처리, 부분 색상 변경)가 필요한 경우에만 예외적으로 허용한다.

| 도구 | 문단/마스터 레벨 지정 방법 |
|---|---|
| `python-pptx` | 문단 기본 속성인 `paragraph.font.size`, `paragraph.font.color.rgb`, `paragraph.font.language_id`를 먼저 설정하고, 일반 텍스트 run은 `run.text`만 채운다. |
| `pptxgenjs` | 개별 text 조각마다 색상을 반복하지 말고, `addText`의 최상위 공통 옵션에 `color`, `fontFace`, `fontSize`를 두고 예외 run에만 override를 적용한다. |
| 템플릿 기반 작업 | 레이아웃의 placeholder 서식을 그대로 살리고 텍스트만 교체한다 (`text_frame.text = "..."`로 프레임 전체를 덮어써서 서식을 날리지 말 것). |

---

## 4. 글머리 기호 및 내어쓰기 (Native Bullet & Hanging Indent)

**증상:** 목록 형태의 문단 앞에 유니코드 기호(`•`, `·`, `-`)를 문자열로 직접 적으면(`run.text = "• 내용"`), 두 번째 줄로 넘어갈 때 텍스트가 불릿 기호 바로 아래로 떨어져 들여쓰기/내어쓰기(Hanging Indent)가 무너진다.

**규칙:**
- 텍스트 문자열 내에 불릿 기호를 하드코딩하지 않는다.
- PowerPoint **네이티브 불릿(Native Bullet)** 기능과 **내어쓰기 마진(`marL`, `indent`)**을 설정하여, 둘째 줄 이후의 텍스트가 첫 줄 본문 시작 위치에 수직으로 완벽하게 일렬 정렬되도록 한다.
- **[중요 - OpenXML 스키마 순서 준수]:** ISO/IEC 29500-1 규격상 `<a:pPr>` 내부에서 `<a:buChar>`는 반드시 `<a:defRPr>` **앞에** 위치해야 한다. `pPr.append(buChar)`를 무조건 호출하면 이미 존재하는 `defRPr` 뒤에 불릿이 붙어 **PowerPoint 열기 시 "프레젠테이션 복구" 경고가 발생**하므로, 반드시 아래 제공되는 안전 헬퍼 함수 `set_native_bullet()`을 사용한다.
- 불릿 색상과 크기가 본문 텍스트를 상속받도록 `<a:buClrTx/>` 및 `<a:buSzTx/>`를 함께 구성한다.

### python-pptx 안전 구현 헬퍼
```python
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches

def set_native_bullet(paragraph, char='•', mar_l=Inches(0.3), indent=-Inches(0.2)):
    """
    OpenXML ISO/IEC 29500-1 스키마 순서를 엄격 준수하여 네이티브 불릿을 적용.
    buChar가 defRPr보다 앞에 오도록 보장하여 PowerPoint 파일 복구 경고를 방지한다.
    """
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set('marL', str(int(mar_l)))
    pPr.set('indent', str(int(indent)))
    
    # 1. 기존 불릿 관련 태그 제거 (중복 방지)
    ns = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    for tag_name in ['buChar', 'buNone', 'buAutoNum', 'buBlip', 'buClrTx', 'buSzTx']:
        for child in pPr.findall(f'{ns}{tag_name}'):
            pPr.remove(child)
            
    # 2. 불릿 요소 생성: 텍스트 색상 상속(buClrTx), 크기 상속(buSzTx), 불릿 문자(buChar)
    bu_clr_tx = OxmlElement('a:buClrTx')
    bu_sz_tx = OxmlElement('a:buSzTx')
    bu_char = OxmlElement('a:buChar')
    bu_char.set('char', char)
    
    # 3. 스키마 순서: [buClrTx, buSzTx, buChar] -> tabLst -> defRPr
    defRPr = pPr.find(f'{ns}defRPr')
    if defRPr is not None:
        idx = pPr.index(defRPr)
        pPr.insert(idx, bu_clr_tx)
        pPr.insert(idx + 1, bu_sz_tx)
        pPr.insert(idx + 2, bu_char)
    else:
        pPr.append(bu_clr_tx)
        pPr.append(bu_sz_tx)
        pPr.append(bu_char)
```

### pptxgenjs 구현
```javascript
// pptxgenjs는 내장 엔진이 자동으로 Hanging Indent 및 서식을 처리함
slide.addText([
  { text: "첫 번째 항목 내용입니다.", options: { bullet: true, lang: "ko-KR" } },
  { text: "두 번째 항목 내용입니다.", options: { bullet: { code: "2022" }, lang: "ko-KR" } }
], { x: 1.0, y: 1.0, w: 8.0, h: 2.0, fontSize: 14, fontFace: "Pretendard", color: "333333" });
```

---

## 5. 카드 및 컨테이너 단일 도형 통합 (Single Shape for Cards)

**증상:** 배경 박스(사각형 도형) 위에 별도의 텍스트 상자(`textbox`)를 얹어 카드를 구현하면, PowerPoint에서 사용자가 카드를 드래그하거나 크기를 바꿀 때 배경과 글자가 분리되고 더블클릭 수정 시 배경 도형이 먼저 선택되어 편집성이 극도로 저하된다.

**규칙:**
- 배경 사각형과 텍스트 상자를 이중으로 분리하여 생성하지 않는다.
- 반드시 **단일 도형(`add_shape(MSO_SHAPE.RECTANGLE)`)**을 생성하고, 해당 도형 자체의 `text_frame`에 텍스트를 직접 삽입한다.
- 도형 내부 텍스트가 정가운데로 몰리지 않도록 **상단 정렬(`vertical_anchor = MSO_ANCHOR.TOP`)**을 지정하고, 테두리와 텍스트 사이의 여백은 **내부 마진(`margin_left`, `margin_top`, `margin_right`, `margin_bottom`)**으로 정밀 제어한다.

### python-pptx 구현
```python
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

def create_korean_card(slide, left, top, width, height, bg_color=RGBColor(245, 247, 250), border_color=RGBColor(220, 225, 230)):
    """편집 친화적인 일체형 단일 카드 도형 생성"""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    shape.line.color.rgb = border_color
    shape.line.width = Pt(1)

    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = Inches(0.3)
    tf.margin_right = Inches(0.3)
    tf.margin_top = Inches(0.3)
    tf.margin_bottom = Inches(0.3)
    return shape
```

### pptxgenjs 구현
```javascript
// addText에 shape를 지정하면 배경과 텍스트가 하나의 도형으로 생성된다.
// margin 배열은 [left, right, bottom, top] (pt)이며, 단일 숫자 margin: 15도 가능하다.
slide.addText([
  { text: "카드 제목", options: { fontSize: 18, bold: true, color: "111827", breakLine: true } },
  { text: "카드 본문 설명 내용입니다. 단일 도형으로 생성되어 이동 및 편집이 일체화됩니다." }
], {
  shape: pptx.ShapeType.rect,
  x: 1.0, y: 1.5, w: 4.5, h: 3.0,
  fill: { color: "F8F9FA" },
  line: { color: "DCE1E6", width: 1 },
  margin: 15,
  align: "left",
  valign: "top",
  lang: "ko-KR",
  fontFace: "Pretendard",
  fontSize: 13,
  color: "4B5563"
});
```

### PptxGenJS 의존성
- 이 예시는 프로젝트에 `pptxgenjs`가 설치되어 있을 때만 실행된다. 설치 여부와 버전은 `package.json` 또는 패키지 관리자에서 확인한다.
- `lang: "ko-KR"`, `fontFace`, `fontSize`, `color`처럼 모든 텍스트에 공통인 값은 `addText`의 최상위 옵션에 둔다. run 옵션은 제목 강조처럼 문단 안에서 달라져야 하는 값에만 사용한다.

---

## 6. 완결형 파이썬 헬퍼 모음 (Production Helper Module)

파이썬 기반 PPTX 생성 시 아래 헬퍼 코드 블록을 스크립트 상단에 복사하여 표준으로 사용한다:

```python
from pptx.oxml.xmlchemy import OxmlElement
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

def set_korean_font_and_lang(target, font_name="Pretendard"):
    """Run 또는 Paragraph 객체에 언어(ko-KR)와 폰트(latin + ea)를 완전 지정"""
    target.font.language_id = MSO_LANGUAGE_ID.KOREAN
    target.font.name = font_name
    
    parent_elm = getattr(target, '_r', None)
    if parent_elm is None:
        parent_elm = getattr(target, '_p', None)
    if parent_elm is None:
        return
    prop_elm = parent_elm.get_or_add_rPr() if hasattr(parent_elm, 'get_or_add_rPr') else parent_elm.get_or_add_pPr().get_or_add_defRPr()
    
    ea = prop_elm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}ea')
    if ea is None:
        ea = OxmlElement('a:ea')
        prop_elm.append(ea)
    ea.set('typeface', font_name)

def set_native_bullet(paragraph, char='•', mar_l=Inches(0.3), indent=-Inches(0.2)):
    """스키마 순서가 보장된 네이티브 불릿 설정"""
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set('marL', str(int(mar_l)))
    pPr.set('indent', str(int(indent)))
    
    ns = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    for tag_name in ['buChar', 'buNone', 'buAutoNum', 'buBlip', 'buClrTx', 'buSzTx']:
        for child in pPr.findall(f'{ns}{tag_name}'):
            pPr.remove(child)
            
    bu_clr_tx = OxmlElement('a:buClrTx')
    bu_sz_tx = OxmlElement('a:buSzTx')
    bu_char = OxmlElement('a:buChar')
    bu_char.set('char', char)
    
    defRPr = pPr.find(f'{ns}defRPr')
    if defRPr is not None:
        idx = pPr.index(defRPr)
        pPr.insert(idx, bu_clr_tx)
        pPr.insert(idx + 1, bu_sz_tx)
        pPr.insert(idx + 2, bu_char)
    else:
        pPr.append(bu_clr_tx)
        pPr.append(bu_sz_tx)
        pPr.append(bu_char)

def create_korean_card(slide, left, top, width, height, bg_color=RGBColor(245, 247, 250), border_color=RGBColor(220, 225, 230)):
    """단일 도형 카드 생성 헬퍼"""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    shape.line.color.rgb = border_color
    shape.line.width = Pt(1)

    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = Inches(0.3)
    tf.margin_right = Inches(0.3)
    tf.margin_top = Inches(0.3)
    tf.margin_bottom = Inches(0.3)
    return shape
```

---

## 체크리스트 (생성/편집 완료 후)

- [ ] 한글이 포함된 모든 run/문단에 `lang="ko-KR"` (또는 `MSO_LANGUAGE_ID.KOREAN`)이 적용되었는가?
- [ ] 텍스트 프레임의 줄바꿈(`word_wrap = True`)이 켜져 있는가?
- [ ] 한글 폰트 지정 시 `<a:latin>`뿐만 아니라 `<a:ea>`(동아시아 폰트)가 동일하게 지정되어 맑은 고딕 강제 폴백을 방지했는가?
- [ ] 폰트 크기·색상이 run마다 반복 하드코딩되지 않고 문단/마스터 레벨에서 정상 상속되는가?
- [ ] 목록 형태의 문단에 유니코드 기호(`•`)를 텍스트로 직접 치지 않고, 네이티브 불릿(`marL`, `indent`, `<a:buChar>`)을 통한 내어쓰기(Hanging Indent)가 적용되었는가?
- [ ] 불릿 적용 시 `<a:buChar>`가 OpenXML 스키마 규격에 맞게 `<a:defRPr>` 앞에 위치하여 파일 복구 경고를 차단했는가?
- [ ] 카드/박스 UI 구현 시 배경 도형과 텍스트 상자를 이중화하지 않고, 단일 도형(`vertical_anchor=TOP` + 내부 마진)으로 일체화되었는가?
- [ ] (템플릿 기반 작업 시) 표준 `pptx` 스킬의 QA 절차(`validate.py`, 시각 QA)를 함께 통과했는가?
