# OpenXML 한국어 서식 규격 레퍼런스

PowerPoint 프레젠테이션의 기반 규격인 OpenXML (ISO/IEC 29500-1)에서 한글 타이포그래피와 관련된 태그 구조 및 스키마 요구사항입니다.

---

## 1. 단어 단위 줄바꿈 방지 (`<a:pPr>` 및 `<a:lvl1pPr>~<a:lvl9pPr>`)

한글 단어가 음절 단위로 분리되는 현상을 막기 위해 문단 프로퍼티(`<a:pPr>`)와 마스터/도형 스타일의 레벨별 문단 프로퍼티(`<a:defPPr>`, `<a:lvl1pPr>` ~ `<a:lvl9pPr>`)에 다음 두 속성을 지정합니다.

```xml
<a:pPr eaLnBrk="0" latinLnBrk="0"/>
```

* `eaLnBrk="0"`: 동아시아 문자(한글/한자/가나)의 단어 중간 줄바꿈 방지 (Break words at syllable boundary 비활성화)
* `latinLnBrk="0"`: 서양 문자(영문/숫자)의 단어 중간 줄바꿈 방지
* 텍스트 프레임의 `word_wrap="1"`은 그대로 유지해야 텍스트가 박스 바깥으로 넘치지 않습니다.
* 슬라이드 마스터(`slideMaster1.xml`)의 `<p:txStyles>`(`titleStyle`, `bodyStyle`, `otherStyle`) 내부 `<a:lvl1pPr>` 등에는 기본적으로 `eaLnBrk="1"`이 설정되어 있으므로 함께 `eaLnBrk="0"`으로 보정해야 상속 시 줄바꿈이 깨지지 않습니다.

---

## 2. 동아시아 폰트 명시 및 `<a:rPr>` / `<a:defRPr>` 스키마 순서

OpenXML 규격상 `<a:rPr>` 또는 `<a:defRPr>` 내에 `<a:latin>`만 정의되어 있으면, PowerPoint는 한글 문자를 시스템 기본 동아시아 폰트(맑은 고딕, 굴림 등)로 강제 폴백(Fallback)시킵니다. 반대로 `<a:ea>`만 정의하고 `<a:latin>`을 누락하거나 영문 폰트(`Calibri` 등)로 두면 한 문장 내 영문·숫자와 한글의 폰트가 달라집니다. 따라서 서양 폰트와 동아시아 폰트를 항상 쌍으로 명시해야 합니다.

### `CT_TextCharacterProperties` (`<a:rPr>`, `<a:defRPr>`, `<a:endParaRPr>`) 자식 요소 순서:
```xml
<a:rPr lang="ko-KR">
  <!-- 1. 테두리 및 채우기 (선택) -->
  <a:ln>...</a:ln>
  <a:solidFill>...</a:solidFill>
  <!-- 2. 효과, 하이라이트, 밑줄 (선택) -->
  <a:effectLst>...</a:effectLst>
  <a:highlight>...</a:highlight>
  <a:uLn>...</a:uLn>
  <a:uFill>...</a:uFill>
  <!-- 3. 폰트 지정 (반드시 latin -> ea -> cs -> sym 순서) -->
  <a:latin typeface="Pretendard"/>
  <a:ea typeface="Pretendard"/>
  <a:cs typeface="Pretendard"/>
  <a:sym typeface="..."/>
  <!-- 4. 하이퍼링크 및 확장 (반드시 폰트 뒤에 위치) -->
  <a:hlinkClick r:id="rId2"/>
  <a:hlinkMouseOver r:id="rId3"/>
  <a:rtl val="0"/>
  <a:extLst>...</a:extLst>
</a:rPr>
```

* **테마 폰트 참조 주의**: `<a:latin typeface="+mn-lt"/>`(`+mj-lt`)인 경우 `<a:ea>`에는 `+mn-ea`(`+mj-ea`)를 대응시켜야 하며, `ppt/theme/theme1.xml`의 `<a:fontScheme>` 내 `<a:ea>`와 `<a:font script="Hang"/>`에도 한글 폰트가 지정되어 있어야 `'맑은 고딕'`으로 폴백되지 않습니다.

---

## 3. 네이티브 불릿 및 `<a:pPr>` 스키마 순서 (ISO/IEC 29500-1 규격)

OpenXML 스키마 규격(`CT_TextParagraphProperties`)상 `<a:pPr>` 내부 자식 요소들은 엄격한 순서를 따라야 합니다. 순서가 어긋나거나 동일 그룹(Choice) 요소가 중복되면 PowerPoint를 열 때 **"프레젠테이션 복구" 경고**가 발생합니다.

### 올바른 순서:
```xml
<a:pPr marL="274320" indent="-182880" eaLnBrk="0" latinLnBrk="0">
  <!-- 1. 줄 간격 및 문단 간격 (선택: 반드시 불릿보다 앞에 위치) -->
  <a:lnSpc>...</a:lnSpc>
  <a:spcBef>...</a:spcBef>
  <a:spcAft>...</a:spcAft>
  <!-- 2. 불릿 색상 그룹 (buClrTx 또는 buClr 중 택1) -->
  <a:buClrTx/>
  <!-- 3. 불릿 크기 그룹 (buSzTx, buSzPct, buSzPts 중 택1) -->
  <a:buSzTx/>
  <!-- 4. 불릿 폰트 그룹 (buFontTx 또는 buFont 중 택1, 선택) -->
  <a:buFontTx/>
  <!-- 5. 불릿 종류 그룹 (buNone, buAutoNum, buChar, buBlip 중 택1) -->
  <a:buChar char="•"/>
  <!-- 6. 탭 리스트 (선택) -->
  <a:tabLst>...</a:tabLst>
  <!-- 7. 기본 런 프로퍼티 (반드시 불릿과 탭 리스트 뒤에 위치) -->
  <a:defRPr lang="ko-KR">...</a:defRPr>
  <!-- 8. 확장 리스트 (선택) -->
  <a:extLst>...</a:extLst>
</a:pPr>
```

> **주의**: `defRPr`나 `tabLst` 뒤에 `buChar`를 `append`하면 안 되며, 반드시 `spcBef`/`spcAft` 뒤이자 `tabLst`/`defRPr` 앞에 위치해야 합니다.
