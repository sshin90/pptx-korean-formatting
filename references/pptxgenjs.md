# PptxGenJS 한국어 서식

PptxGenJS 4.0.1 생성 결과를 기준으로 확인했다. 다른 버전에서는 저장된 XML도 확인한다.

## 언어·폰트·줄바꿈

```javascript
const koreanOptions = {
  lang: "ko-KR", fontFace: "Pretendard",
  fontSize: 14, color: "333333"
};
slide.addText("한국어 본문입니다.", {
  x: 1, y: 1, w: 8, h: 1.5, wrap: true, ...koreanOptions
});
```

4.0.1의 일반 텍스트 생성에서 `eaLnBrk="0"`은 지정되지 않는다. 단어 단위 줄바꿈이 필요한 한국어 문단은 저장 완료 후 보정한다.

```javascript
await pptx.writeFile({ fileName: "output.pptx" });
```

```bash
python "<skill-path>/scripts/fix_korean_pptx.py" output.pptx
python "<skill-path>/scripts/fix_korean_pptx.py" output.pptx --verify
```

후처리는 파일 저장을 기다린 다음 실행한다. 기본 보정은 폰트·테마를 유지한다. 임의 폰트 설치나 글리프 지원을 보장하지 않는다.

## 공통 옵션과 파일의 서식 상속은 다르다

최상위 옵션은 생성 코드 중복을 줄인다. 하지만 PptxGenJS 4.0.1은 공통 폰트·크기·색상도 각 run의 `rPr`에 기록할 수 있다. 최상위 옵션 사용만으로 마스터 상속이 보장되지는 않는다.

마스터/테마를 바꿀 때 함께 변해야 하는 요소에는 테마 참조나 실제 상속 구조를 사용하고 생성 XML로 확인한다. 고정 RGB는 문단/run 위치와 관계없이 테마 색 변경을 따르지 않는다. 기본 CLI는 의도된 직접 서식을 제거하지 않는다. 명시적 `-f`도 폰트만 정리하며 색상·크기는 바꾸지 않는다.

## 네이티브 불릿

```javascript
slide.addText([
  { text: "첫 번째 항목입니다.", options: { bullet: true } },
  { text: "두 번째 항목입니다.", options: { bullet: { code: "2022" } } }
], { x: 1, y: 1, w: 8, h: 2, ...koreanOptions });
```

4.0.1에서 두 문단·두 불릿이 생성된다. 긴 항목의 둘째 줄 들여쓰기는 렌더링하여 확인한다. 일반 rich text 줄바꿈에는 `breakLine: true`를 사용하고 동일 경계에 `\n`을 중복하지 않는다.

## 단일 카드 도형

```javascript
slide.addText([
  { text: "카드 제목", options: { fontSize: 18, bold: true, breakLine: true } },
  { text: "본문 설명입니다." }
], {
  shape: pptx.ShapeType.rect,
  x: 1, y: 1.5, w: 4.5, h: 3,
  fill: { color: "F8F9FA" }, line: { color: "DCE1E6", width: 1 },
  margin: 15, align: "left", valign: "top", wrap: true,
  ...koreanOptions
});
```

`margin`은 pt 단위이며 배열은 [top, right, bottom, left]다. 위 예제는 제목과 본문이 하나의 도형에 들어간다. 독립적인 공간 배치가 필요하면 그룹화를 선택한다.

[공식 텍스트 API](https://gitbrent.github.io/PptxGenJS/docs/api-text/)
