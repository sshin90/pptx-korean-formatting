# OpenXML 한국어 서식 참조

이 도구는 Transitional DrawingML 네임스페이스 `http://schemas.openxmlformats.org/drawingml/2006/main`을 처리한다. 아래는 주요 구조 참고이며 전체 ISO 스키마 검증을 대체하지 않는다.

## 줄바꿈

```xml
<p:txBody>
  <a:bodyPr wrap="square"/>
  <a:lstStyle/>
  <a:p>
    <a:pPr eaLnBrk="0" latinLnBrk="0"/>
    <a:r><a:t>한국어 문장</a:t></a:r>
  </a:p>
</p:txBody>
```

예제의 `p`/`a` 접두사는 상위 요소에서 선언되어 있다고 가정한다. Python `word_wrap=True`와 XML `wrap="square"`를 혼동하지 않는다. `eaLnBrk`는 동아시아 단어의 중간 줄바꿈, `latinLnBrk`는 서양 단어의 중간 줄바꿈을 제어한다. Boolean `0`과 `false`는 동일하게 취급한다.

문단 속성은 `a:pPr`, 도형 목록 스타일은 `a:lstStyle/a:lvl1pPr`부터 `lvl9pPr`, 마스터 텍스트 스타일은 `p:txStyles`에 위치할 수 있다. 공유 스타일에 값이 있더라도 문단의 명시적 값이 우선한다. 한국어 정책을 공유 스타일 전체에 적용하는 것은 외국어에도 영향을 줄 수 있다.

## 폰트와 상속

폰트는 run의 `a:rPr`, 문단·목록 스타일의 `a:defRPr`, 입력 끝의 `a:endParaRPr` 등에 존재한다. 낮은 레벨에 같은 폰트 이름을 복사하면 상속이 고정 override로 바뀐다. 속성 부재와 명시적 값은 구분한다.

- `a:latin`: 라틴 문자 폰트
- `a:ea`: 동아시아 폰트
- `+mn-lt`/`+mj-lt`: 본문/제목 라틴 테마 참조
- `+mn-ea`/`+mj-ea`: 본문/제목 동아시아 테마 참조

EA 요소에 잘못 기록된 `+mn-lt`/`+mj-lt`는 대응하는 EA 토큰으로 수정한다. EA 속성이 없을 때에는 상위 상속을 먼저 고려한다. 다른 영문·한글 폰트 조합도 유효하다.

테마의 `fontScheme/majorFont`, `minorFont`는 `latin`, `ea`, `cs`, 스크립트별 `font` 매핑을 가진다. 테마의 빈 `ea typeface`와 정상적인 `font script="Hang"` 조합을 오류로 간주하지 않는다. 기본 CLI는 이 매핑을 변경하지 않는다.

### 문자 속성의 주요 자식 순서

`ln → Fill choice → Effect choice → highlight → 밑줄 속성 → latin → ea → cs → sym → hlinkClick → hlinkMouseOver → rtl → extLst`

각 요소는 선택 사항이다. 폰트를 추가할 때 하이퍼링크 뒤에 append하지 않는다. 기존 잘못된 순서를 무조건 자동 복구하지 않으며 검증기가 주요 폰트 순서 이상을 보고한다.

## 불릿과 문단 속성 순서

```xml
<a:pPr marL="274320" indent="-182880">
  <a:lnSpc><a:spcPct val="100000"/></a:lnSpc>
  <a:spcBef><a:spcPts val="0"/></a:spcBef>
  <a:spcAft><a:spcPts val="600"/></a:spcAft>
  <a:buClrTx/>
  <a:buSzTx/>
  <a:buFontTx/>
  <a:buChar char="•"/>
  <a:tabLst/>
  <a:defRPr lang="ko-KR"/>
</a:pPr>
```

불릿 그룹마다 choice는 하나다: 색상(buClrTx/buClr), 크기(buSzTx/buSzPct/buSzPts), 폰트(buFontTx/buFont), 종류(buNone/buAutoNum/buChar/buBlip). 불릿은 `tabLst`/`defRPr`/`extLst`보다 앞에 둔다. `marL`과 `indent`는 EMU 단위다.

## 검증기의 해석 경계

문단의 직접 속성과 같은 도형의 현재 목록 레벨·기본 목록 스타일까지 읽는다. 레이아웃/마스터/테마를 가로지르는 유효 폰트는 `unknown_*`로 보고한다. 직접 지정된 빈/잘못된 EA 토큰을 정상 상위 폰트로 덮어 해석하지 않는다. 폰트 이름만으로 글리프를 판정하지 않는다.

[python-pptx 텍스트 API](https://python-pptx.readthedocs.io/en/latest/api/text.html) · [Microsoft OpenXML 문단 속성](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.paragraphproperties)
