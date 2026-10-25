# PptxGenJS 한국어 서식 가이드

Node.js 또는 브라우저 환경에서 `pptxgenjs`를 사용할 때 한국어 텍스트 서식을 올바르게 구성하기 위한 참조 문서입니다.

---

## 1. 언어 속성 및 단어 단위 줄바꿈 (CLI 후처리 필수)

PptxGenJS에서는 `addText` 옵션 객체에 언어 코드(`lang: "ko-KR"`)와 폰트(`fontFace: "Pretendard"`)를 지정합니다.

```javascript
// 공통 옵션 객체 정의
const koreanTextOptions = {
  lang: "ko-KR",
  fontFace: "Pretendard",
  fontSize: 14,
  color: "333333"
};

// 텍스트 추가 시 공통 옵션 병합
slide.addText("한국어 프레젠테이션 본문 텍스트입니다.", {
  x: 1.0,
  y: 1.0,
  w: 8.0,
  h: 1.5,
  ...koreanTextOptions
});
```

> **주의 1 (자동 줄바꿈 유지)**: 텍스트 상자의 자동 줄바꿈은 기본 활성화되어 있으므로 `wrap: false`로 끄지 마십시오.
>
> **주의 2 (`eaLnBrk="0"` 후처리 필수)**: PptxGenJS 엔진은 OpenXML `<a:pPr>`에 `eaLnBrk="0"`(단어 단위 줄바꿈) 속성을 생성하는 옵션이 없습니다. 따라서 PptxGenJS로 `.pptx` 파일을 생성·저장한 직후 반드시 CLI 후처리기를 1회 실행해야 한글 단어 쪼개짐과 폰트 폴백이 완벽히 방지됩니다:
> ```bash
> python <skill-path>/scripts/fix_korean_pptx.py output.pptx
> ```

---

## 2. 네이티브 불릿 목록 (서식 상속 준수)

PptxGenJS는 엔진 자체에서 불릿 기호와 Hanging Indent를 처리합니다. 유니코드 기호(`•`)를 텍스트 문자열에 직접 넣지 말고 `bullet` 옵션을 사용하며, 공통 서식(`lang`, `fontFace`, `fontSize`, `color`)은 최상위 옵션에서 상속시킵니다.

```javascript
slide.addText([
  { text: "첫 번째 항목 내용입니다.", options: { bullet: true } },
  { text: "두 번째 항목 내용입니다.", options: { bullet: { code: "2022" } } }
], {
  x: 1.0,
  y: 1.0,
  w: 8.0,
  h: 2.0,
  lang: "ko-KR",
  fontSize: 14,
  fontFace: "Pretendard",
  color: "333333"
});
```

---

## 3. 일체형 단일 카드 도형 (Single Shape Card)

배경 사각형 위에 텍스트 박스를 별도로 얹지 않고, `shape` 옵션을 사용하여 단일 도형으로 생성합니다. 텍스트 배열에서 줄바꿈을 할 때는 `breakLine: true`만 사용하고 문자열 끝에 `\n`을 중복으로 넣지 않도록 주의하십시오.

```javascript
// margin 배열은 [top, right, bottom, left] (pt 단위) 또는 단일 숫자
slide.addText([
  { text: "카드 제목", options: { fontSize: 18, bold: true, color: "111827", breakLine: true } },
  { text: "카드 본문 설명 내용입니다. 단일 도형으로 생성되어 사용자가 이동하거나 편집할 때 일체형으로 유지됩니다." }
], {
  shape: pptx.ShapeType.rect,
  x: 1.0,
  y: 1.5,
  w: 4.5,
  h: 3.0,
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
