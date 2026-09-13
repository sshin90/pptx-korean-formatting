# PptxGenJS 한국어 서식 가이드

Node.js 또는 브라우저 환경에서 `pptxgenjs`를 사용할 때 한국어 텍스트 서식을 올바르게 구성하기 위한 참조 문서입니다.

---

## 1. 언어 속성 및 줄바꿈

PptxGenJS에서는 `addText` 옵션 객체에 언어 코드(`lang: "ko-KR"`)를 지정합니다.

```javascript
// 공통 옵션 객체 정의
const koreanTextOptions = {
  lang: "ko-KR",
  fontFace: "Pretendard",
  fontSize: 14,
  color: "333333"
};

// 텍스트 추가 시 공통 옵션 병합
slide.addText("단어 중간 줄바꿈 없이 깔끔하게 렌더링됩니다.", {
  x: 1.0,
  y: 1.0,
  w: 8.0,
  h: 1.5,
  ...koreanTextOptions
});
```

> **주의**: 텍스트 상자의 자동 줄바꿈은 기본 활성화되어 있으므로 `wrap: false`로 끄지 마십시오.

---

## 2. 네이티브 불릿 목록

PptxGenJS는 엔진 자체에서 불릿 기호와 Hanging Indent를 처리합니다. 유니코드 기호(`•`)를 텍스트 문자열에 직접 넣지 말고 `bullet` 옵션을 사용하십시오.

```javascript
slide.addText([
  { text: "첫 번째 항목 내용입니다.", options: { bullet: true, lang: "ko-KR" } },
  { text: "두 번째 항목 내용입니다.", options: { bullet: { code: "2022" }, lang: "ko-KR" } }
], {
  x: 1.0,
  y: 1.0,
  w: 8.0,
  h: 2.0,
  fontSize: 14,
  fontFace: "Pretendard",
  color: "333333"
});
```

---

## 3. 일체형 단일 카드 도형 (Single Shape Card)

배경 사각형 위에 텍스트 박스를 별도로 얹지 않고, `shape` 옵션을 사용하여 단일 도형으로 생성합니다.

```javascript
// margin 배열은 [top, right, bottom, left] (pt 단위) 또는 단일 숫자
slide.addText([
  { text: "카드 제목\n", options: { fontSize: 18, bold: true, color: "111827", breakLine: true } },
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
