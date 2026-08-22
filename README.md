# pptx-korean-formatting

*[한국어 설명은 아래를 참고하세요](#한국어)*

[Claude Code Skill](https://docs.claude.com/en/docs/claude-code/skills) that adds Korean-language formatting rules on top of the standard `pptx` skill (or any `python-pptx` / `pptxgenjs` / `html2pptx` based workflow) whenever a generated PowerPoint deck contains Korean (한글) text.

## Why

PowerPoint applies word-wrap rules based on each run/paragraph's language tag. When Korean text is written with the default `en-US` tag, PowerPoint wraps it using English rules and breaks words mid-syllable (e.g. "안녕하세요" gets split into "안녕하" / "세요"). This skill makes sure every Korean text run is tagged `lang="ko-KR"` and that formatting (font size/color) is set at the paragraph/master level instead of being hardcoded per-run, so later bulk edits in PowerPoint (like a theme color change) still work.

## What it does

1. **Language tag enforcement** — forces `lang="ko-KR"` (or the tool-specific equivalent: `MSO_LANGUAGE_ID.KOREAN` for `python-pptx`, `lang: "ko-KR"` for `pptxgenjs`, `lang="ko-KR"` on the root element for `html2pptx`) on every run/paragraph containing Korean text, and confirms word-wrap is not disabled.
2. **Formatting inheritance** — pushes font size/color up to slide master/layout or paragraph level instead of hardcoding on every run, so downstream edits in PowerPoint aren't blocked.

See [SKILL.md](SKILL.md) for the full rule set, per-tool instructions, and a post-generation checklist.

## Installation

This is a [Claude Code Skill](https://docs.claude.com/en/docs/claude-code/skills). To use it:

```bash
git clone https://github.com/sshin90/pptx-korean-formatting.git ~/.claude/skills/pptx-korean-formatting
```

Claude Code will automatically pick it up alongside the built-in `pptx` skill whenever a task involves creating or editing a `.pptx`/`.potx` file with Korean content.

## License

[CC0 1.0 Universal](LICENSE) — public domain, no attribution required.

---

## 한국어

한글이 포함된 PowerPoint(.pptx) 슬라이드를 생성할 때, 표준 `pptx` 스킬(또는 `python-pptx`, `pptxgenjs`, `html2pptx` 기반 워크플로우)에 **추가로** 적용하는 [Claude Code Skill](https://docs.claude.com/en/docs/claude-code/skills)입니다.

### 왜 필요한가

PowerPoint는 각 텍스트 런/문단에 지정된 언어 속성에 따라 줄바꿈 규칙을 적용합니다. 한글 텍스트에 기본값인 `en-US` 태그가 남아있으면 PowerPoint가 영어 줄바꿈 규칙을 적용해 단어를 음절 단위로 잘못 잘라버립니다 (예: "안녕하세요"가 "안녕하" / "세요"로 분리). 이 스킬은 한글이 포함된 모든 텍스트 런에 `lang="ko-KR"`을 강제로 지정하고, 폰트 크기·색상 같은 서식을 run마다 하드코딩하지 않고 문단/마스터 레벨에 지정하도록 해서, 나중에 PowerPoint에서 테마 색상을 일괄 변경하는 등의 편집이 정상적으로 반영되도록 합니다.

### 무엇을 하는가

1. **언어 태그 강제 지정** — 한글이 포함된 모든 run/문단에 `lang="ko-KR"` (또는 도구별 동등 속성: `python-pptx`의 `MSO_LANGUAGE_ID.KOREAN`, `pptxgenjs`의 `lang: "ko-KR"`, `html2pptx`의 루트 요소 `lang="ko-KR"`)을 적용하고, 줄바꿈(word-wrap)이 꺼져있지 않은지 확인합니다.
2. **서식 상속 처리** — 폰트 크기·색상을 run마다 반복 지정하지 않고 슬라이드 마스터/레이아웃 또는 문단 레벨로 올려서, 이후 PowerPoint에서의 편집이 막히지 않도록 합니다.

전체 규칙, 도구별 상세 지침, 생성 완료 후 체크리스트는 [SKILL.md](SKILL.md)를 참고하세요.

### 설치 방법

```bash
git clone https://github.com/sshin90/pptx-korean-formatting.git ~/.claude/skills/pptx-korean-formatting
```

설치 후에는 한글이 포함된 `.pptx`/`.potx` 파일을 생성하거나 편집하는 작업에서 표준 `pptx` 스킬과 함께 Claude Code가 자동으로 이 스킬을 적용합니다.

### 라이선스

[CC0 1.0 Universal](LICENSE) — 퍼블릭 도메인, 저작자 표시 의무 없음.
