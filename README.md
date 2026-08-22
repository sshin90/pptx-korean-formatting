# pptx-korean-formatting

*[한국어 설명은 아래를 참고하세요](#한국어)*

Korean-language formatting rules for AI coding agents that generate PowerPoint (`.pptx`/`.potx`) files with `python-pptx`, `pptxgenjs`, `html2pptx`, or direct XML editing. Works with **any agent** that can read a project-level instructions file — Claude Code, OpenAI Codex, Google Antigravity, Cursor, Windsurf, and others.

## Why

PowerPoint applies word-wrap rules based on each run/paragraph's language tag. When Korean text is written with the default `en-US` tag, PowerPoint wraps it using English rules and breaks words mid-syllable (e.g. "안녕하세요" gets split into "안녕하" / "세요"). These rules make sure every Korean text run is tagged `lang="ko-KR"` and that formatting (font size/color) is set at the paragraph/master level instead of being hardcoded per-run, so later bulk edits in PowerPoint (like a theme color change) still work.

## What it does

1. **Language tag enforcement** — forces `lang="ko-KR"` (or the tool-specific equivalent: `MSO_LANGUAGE_ID.KOREAN` for `python-pptx`, `lang: "ko-KR"` for `pptxgenjs`, `lang="ko-KR"` on the root element for `html2pptx`) on every run/paragraph containing Korean text, and confirms word-wrap is not disabled.
2. **Formatting inheritance** — pushes font size/color up to slide master/layout or paragraph level instead of hardcoding on every run, so downstream edits in PowerPoint aren't blocked.

## Files

| File | Format | Use with |
|---|---|---|
| [AGENTS.md](AGENTS.md) | Plain markdown, no frontmatter | Any agent — OpenAI Codex, Google Antigravity, Cursor, Windsurf, or manually pasted into a system prompt |
| [SKILL.md](SKILL.md) | Claude [Agent Skills](https://docs.claude.com/en/docs/claude-code/skills) format (YAML frontmatter + auto-trigger description) | Claude Code, or any tool that supports the Agent Skills spec |

The rule content is identical in both — pick the file that matches how your agent loads project instructions.

## Installation / wiring it up

**Claude Code** (auto-discovered Agent Skill):
```bash
git clone https://github.com/sshin90/pptx-korean-formatting.git ~/.claude/skills/pptx-korean-formatting
```

**OpenAI Codex CLI** (or any tool reading `AGENTS.md`):
```bash
git clone https://github.com/sshin90/pptx-korean-formatting.git
cat pptx-korean-formatting/AGENTS.md >> AGENTS.md   # merge into your project's AGENTS.md
```

**Google Antigravity, Cursor, Windsurf, or other agentic IDEs**: these tools each have their own convention for project rules (custom instructions, `.cursorrules`, `.windsurfrules`, workspace guidelines, etc.). Copy the contents of [AGENTS.md](AGENTS.md) into whichever rules file your tool reads — check its docs for the exact filename/location.

**Any other agent**: paste [AGENTS.md](AGENTS.md)'s contents into the system prompt, or tell the agent to read this file before generating a Korean-language deck.

## License

[CC0 1.0 Universal](LICENSE) — public domain, no attribution required.

---

## 한국어

`python-pptx`, `pptxgenjs`, `html2pptx`, 또는 직접 XML 편집으로 PowerPoint(.pptx/.potx)를 생성하는 **모든 AI 코딩 에이전트**를 위한 한국어 서식 규칙입니다. 프로젝트 단위 지침 파일을 읽을 수 있는 에이전트라면 어디서든 사용할 수 있습니다 — Claude Code, OpenAI Codex, Google Antigravity, Cursor, Windsurf 등.

### 왜 필요한가

PowerPoint는 각 텍스트 런/문단에 지정된 언어 속성에 따라 줄바꿈 규칙을 적용합니다. 한글 텍스트에 기본값인 `en-US` 태그가 남아있으면 PowerPoint가 영어 줄바꿈 규칙을 적용해 단어를 음절 단위로 잘못 잘라버립니다 (예: "안녕하세요"가 "안녕하" / "세요"로 분리). 이 규칙들은 한글이 포함된 모든 텍스트 런에 `lang="ko-KR"`을 강제로 지정하고, 폰트 크기·색상 같은 서식을 run마다 하드코딩하지 않고 문단/마스터 레벨에 지정하도록 해서, 나중에 PowerPoint에서 테마 색상을 일괄 변경하는 등의 편집이 정상적으로 반영되도록 합니다.

### 무엇을 하는가

1. **언어 태그 강제 지정** — 한글이 포함된 모든 run/문단에 `lang="ko-KR"` (또는 도구별 동등 속성: `python-pptx`의 `MSO_LANGUAGE_ID.KOREAN`, `pptxgenjs`의 `lang: "ko-KR"`, `html2pptx`의 루트 요소 `lang="ko-KR"`)을 적용하고, 줄바꿈(word-wrap)이 꺼져있지 않은지 확인합니다.
2. **서식 상속 처리** — 폰트 크기·색상을 run마다 반복 지정하지 않고 슬라이드 마스터/레이아웃 또는 문단 레벨로 올려서, 이후 PowerPoint에서의 편집이 막히지 않도록 합니다.

### 파일 구성

| 파일 | 형식 | 사용 대상 |
|---|---|---|
| [AGENTS.md](AGENTS.md) | 프론트매터 없는 순수 마크다운 | OpenAI Codex, Google Antigravity, Cursor, Windsurf 등 모든 에이전트, 또는 시스템 프롬프트에 직접 붙여넣기 |
| [SKILL.md](SKILL.md) | Claude [Agent Skills](https://docs.claude.com/en/docs/claude-code/skills) 형식 (YAML 프론트매터 + 자동 트리거 설명 포함) | Claude Code, 또는 Agent Skills 스펙을 지원하는 도구 |

두 파일의 규칙 내용은 동일합니다. 사용 중인 에이전트가 프로젝트 지침을 읽는 방식에 맞는 파일을 고르면 됩니다.

### 설치 / 연결 방법

**Claude Code** (Agent Skill로 자동 인식):
```bash
git clone https://github.com/sshin90/pptx-korean-formatting.git ~/.claude/skills/pptx-korean-formatting
```

**OpenAI Codex CLI** (또는 `AGENTS.md`를 읽는 다른 도구):
```bash
git clone https://github.com/sshin90/pptx-korean-formatting.git
cat pptx-korean-formatting/AGENTS.md >> AGENTS.md   # 프로젝트의 AGENTS.md에 병합
```

**Google Antigravity, Cursor, Windsurf 등 다른 에이전틱 IDE**: 각 도구마다 프로젝트 규칙을 지정하는 자체 방식(커스텀 지침, `.cursorrules`, `.windsurfrules`, 워크스페이스 가이드라인 등)이 있습니다. 사용 중인 도구의 문서에서 정확한 파일명/위치를 확인한 뒤, [AGENTS.md](AGENTS.md)의 내용을 그 파일에 복사해 넣으세요.

**그 외 모든 에이전트**: [AGENTS.md](AGENTS.md)의 내용을 시스템 프롬프트에 붙여넣거나, 한글 슬라이드를 생성하기 전에 이 파일을 읽도록 에이전트에게 지시하면 됩니다.

### 라이선스

[CC0 1.0 Universal](LICENSE) — 퍼블릭 도메인, 저작자 표시 의무 없음.
