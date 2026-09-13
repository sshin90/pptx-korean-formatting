# OpenXML 한국어 서식 규격 레퍼런스

PowerPoint 프레젠테이션의 기반 규격인 OpenXML (ISO/IEC 29500-1)에서 한글 타이포그래피와 관련된 태그 구조 및 스키마 요구사항입니다.

---

## 1. 단어 단위 줄바꿈 방지 (`<a:pPr>`)

한글 단어가 음절 단위로 분리되는 현상을 막기 위해 문단 프로퍼티(`<a:pPr>`)에 다음 두 속성을 지정합니다.

```xml
<a:pPr eaLnBrk="0" latinLnBrk="0"/>
```

* `eaLnBrk="0"`: 동아시아 문자(한글/한자/가나)의 단어 중간 줄바꿈 방지 (Break words at syllable boundary 비활성화)
* `latinLnBrk="0"`: 서양 문자(영문/숫자)의 단어 중간 줄바꿈 방지
* 텍스트 프레임의 `word_wrap="1"`은 그대로 유지해야 텍스트가 박스 바깥으로 넘치지 않습니다.

---

## 2. 동아시아 폰트 명시 (`<a:rPr>`)

OpenXML 규격상 `<a:rPr>` 내에 `<a:latin>`만 정의되어 있으면, PowerPoint는 한글 문자를 시스템 기본 동아시아 폰트(맑은 고딕, 굴림 등)로 강제 폴백(Fallback)시킵니다. 따라서 서양 폰트와 동아시아 폰트를 항상 쌍으로 명시해야 합니다.

```xml
<a:rPr lang="ko-KR">
  <a:latin typeface="Pretendard"/>
  <a:ea typeface="Pretendard"/>
</a:rPr>
```

---

## 3. 네이티브 불릿 스키마 순서 (ISO/IEC 29500-1 규격)

OpenXML 스키마 규격상 `<a:pPr>` 내부 자식 요소들은 엄격한 순서를 따라야 합니다. 순서가 어긋나면 PowerPoint를 열 때 **"프레젠테이션 복구" 경고**가 발생합니다.

### 올바른 순서:
```
<a:pPr marL="..." indent="...">
  <!-- 1. 불릿 서식 상속 -->
  <a:buClrTx/>
  <a:buSzTx/>
  <!-- 2. 불릿 문자 -->
  <a:buChar char="•"/>
  <!-- 3. 탭 리스트 (선택) -->
  <a:tabLst>...</a:tabLst>
  <!-- 4. 기본 런 프로퍼티 (반드시 불릿 뒤에 위치) -->
  <a:defRPr .../>
</a:pPr>
```

> **주의**: `defRPr` 뒤에 `buChar`를 `append`하면 안 되며, 반드시 `defRPr` 앞에 `insert`해야 합니다.
