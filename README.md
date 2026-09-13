# pptx-korean-formatting

PowerPoint(`.pptx`/`.potx`)에 한글이 포함될 때 필요한 언어 태그, 단어 단위 줄바꿈, 동아시아 폰트, 서식 상속, 네이티브 불릿, 편집 가능한 카드 도형 규칙을 제공하는 보조 스킬입니다.

## 한국어

`python-pptx`, `pptxgenjs`, `html2pptx`, 또는 직접 XML 편집으로 PowerPoint 파일을 생성하는 **모든 AI 코딩 에이전트**에서 사용할 수 있습니다. 프로젝트 단위 또는 전역 지침 파일을 읽을 수 있는 Claude Code, OpenAI Codex, Google Antigravity, Cursor, Windsurf 등을 지원합니다.

### 왜 필요한가

가독성이 중요한 프레젠테이션에서는 한글이나 영문 단어가 음절·문자 중간에서 잘리지 않고 어절 또는 단어 단위로 줄바꿈되는 것이 좋습니다. PowerPoint에서는 자동 줄바꿈과 단어 중간 줄바꿈 허용 여부가 서로 다른 속성으로 관리됩니다.

- `word_wrap = True`는 텍스트가 상자 폭을 넘을 때 다음 줄로 보내는 자동 줄바꿈을 활성화합니다.
- `eaLnBrk="0"`과 `latinLnBrk="0"`은 한글과 영문 단어가 중간에서 잘리지 않도록 합니다.
- `lang="ko-KR"`은 PowerPoint가 한글 텍스트를 올바르게 인식하여 한글 맞춤법 검사와 문자 처리 규칙을 적용하도록 합니다.

따라서 단어 잘림을 막기 위해 자동 줄바꿈 자체를 끄면 안 됩니다. 자동 줄바꿈은 유지하면서 단어 중간 줄바꿈 속성을 문단과 마스터·레이아웃의 기본 문단 속성에 명시해야 합니다. 기존 PPTX를 가져와 다시 내보내는 경우에도 해당 속성이 `1`로 바뀌지 않았는지 XML을 검증해야 합니다.

이 스킬은 한글 폰트가 시스템 기본 폰트로 바뀌는 문제와 run마다 하드코딩된 서식 때문에 후속 편집이 어려워지는 문제도 함께 방지합니다.

### 무엇을 하는가

1. **언어 태그 강제 지정** — 한글이 포함된 모든 run/문단에 `lang="ko-KR"` 또는 도구별 동등 속성(`python-pptx`의 `MSO_LANGUAGE_ID.KOREAN`, `pptxgenjs`의 `lang: "ko-KR"`, `html2pptx`의 루트 요소 `lang="ko-KR"`)을 적용합니다.
2. **단어 중간 줄바꿈 방지** — 자동 줄바꿈은 켠 채로 문단과 마스터·레이아웃의 기본 문단 속성에 `eaLnBrk="0"`, `latinLnBrk="0"`을 지정하고, 가져오기·내보내기 전후에 값이 유지되는지 확인합니다.
3. **한글 폰트(East Asian Font) 명시** — `<a:latin>`과 동아시아 폰트 태그(`<a:ea>`)에 같은 폰트를 지정하여 맑은 고딕 등으로 강제 폴백되는 현상을 방지합니다.
4. **서식 상속 처리** — 폰트 크기·색상을 run마다 반복 지정하지 않고 슬라이드 마스터·레이아웃 또는 문단 레벨에 두어 PowerPoint에서 일괄 편집할 수 있게 합니다.
5. **네이티브 불릿 및 내어쓰기** — 문자열 불릿(`•`) 대신 네이티브 불릿과 내어쓰기(`marL`, `indent`)를 사용합니다. `<a:buChar>`를 `<a:defRPr>` 앞에 배치하는 OpenXML 순서를 지켜 파일 복구 경고를 방지합니다.
6. **카드·컨테이너 단일 도형 통합** — 배경 도형과 텍스트 상자를 분리하지 않고 `vertical_anchor=TOP`과 내부 마진이 적용된 단일 도형으로 만들어 편집성을 높입니다.
7. **완결형 파이썬 헬퍼 제공** — 실무에서 바로 사용할 수 있는 `set_korean_font_and_lang`, `set_native_bullet`, `create_korean_card` 헬퍼를 제공합니다.

### 파일 구성

| 파일 | 형식 | 사용 대상 |
|---|---|---|
| [AGENTS.md](AGENTS.md) | 프론트매터 없는 순수 마크다운 | OpenAI Codex, Google Antigravity, Cursor, Windsurf 등 모든 에이전트 또는 시스템 프롬프트에 직접 붙여넣기 |
| [SKILL.md](SKILL.md) | [Agent Skills](https://docs.claude.com/en/docs/claude-code/skills) 형식(YAML 프론트매터와 자동 트리거 설명 포함) | Claude Code, Google Antigravity 또는 Agent Skills 사양을 지원하는 도구 |

두 파일은 로딩 방식이 다릅니다. 사용 중인 에이전트가 프로젝트 지침을 읽는 방식에 맞는 파일을 선택하고, 최신 Agent Skill 규칙은 [SKILL.md](SKILL.md)를 기준으로 확인하세요.

### 설치 및 연결

**Claude Code** (Agent Skill로 자동 인식):

```bash
git clone https://github.com/sshin90/pptx-korean-formatting.git ~/.claude/skills/pptx-korean-formatting
```

**OpenAI Codex CLI** (또는 `AGENTS.md`를 읽는 다른 도구):

```bash
git clone https://github.com/sshin90/pptx-korean-formatting.git
cat pptx-korean-formatting/AGENTS.md >> AGENTS.md   # 프로젝트의 AGENTS.md에 병합
```

**Google Antigravity**: Claude Code와 동일한 Agent Skills 형식(`SKILL.md`가 포함된 스킬 폴더)과 더 단순한 전역 규칙 디렉터리를 모두 지원합니다.

- 스킬로 설치:

  ```bash
  git clone https://github.com/sshin90/pptx-korean-formatting.git ~/.gemini/config/skills/pptx-korean-formatting
  ```

- 전역 규칙으로 설치(모든 워크스페이스에 적용):

  ```bash
  mkdir -p ~/.gemini/config/rules
  curl -o ~/.gemini/config/rules/pptx-korean-formatting.md \
    https://raw.githubusercontent.com/sshin90/pptx-korean-formatting/master/AGENTS.md
  ```

**Cursor, Windsurf 등 다른 에이전틱 IDE**: 각 도구의 프로젝트 규칙 파일(`.cursorrules`, `.cursor/rules/`, `.windsurfrules`, 워크스페이스 가이드라인 등)에 [AGENTS.md](AGENTS.md)의 내용을 복사하세요. 정확한 파일명과 위치는 해당 도구의 문서를 확인하세요.

**그 외 모든 에이전트**: [AGENTS.md](AGENTS.md)의 내용을 시스템 프롬프트에 붙여넣거나, 한글 슬라이드를 만들기 전에 이 파일을 읽도록 지시하세요.

### 라이선스

[CC0 1.0 Universal](LICENSE) — 퍼블릭 도메인, 저작자 표시 의무 없음.

---

## English

Supplementary Korean-language formatting rules for **any AI coding agent** that generates PowerPoint (`.pptx`/`.potx`) files with `python-pptx`, `pptxgenjs`, `html2pptx`, or direct XML editing. It works with agents that can read project-level or global instruction files, including Claude Code, OpenAI Codex, Google Antigravity, Cursor, and Windsurf.

### Why it is needed

Presentation text is easier to read when Korean and English words wrap as whole words or phrases instead of splitting in the middle of a syllable or character. PowerPoint controls automatic text wrapping and mid-word line breaking with separate properties.

- `word_wrap = True` enables automatic wrapping when text reaches the edge of its container.
- `eaLnBrk="0"` and `latinLnBrk="0"` prevent Korean and Latin words from breaking in the middle.
- `lang="ko-KR"` lets PowerPoint recognize Korean text and apply the appropriate language and typography behavior.

Do not disable automatic wrapping to prevent broken words. Keep automatic wrapping enabled while explicitly disabling mid-word breaks on paragraphs and inherited master/layout paragraph defaults. When importing and re-exporting an existing deck, verify in the resulting XML that these values have not changed to `1`.

The skill also prevents Korean font fallback and preserves downstream editability by keeping shared formatting at the paragraph, layout, or master level instead of hardcoding it on every run.

### What it does

1. **Language tag enforcement** — applies `lang="ko-KR"` or its tool-specific equivalent (`MSO_LANGUAGE_ID.KOREAN` in `python-pptx`, `lang: "ko-KR"` in `pptxgenjs`, or `lang="ko-KR"` on an `html2pptx` root element) to Korean text runs and paragraphs.
2. **Mid-word break prevention** — keeps automatic wrapping enabled while setting `eaLnBrk="0"` and `latinLnBrk="0"` on paragraphs and inherited master/layout defaults, then checks that import/export workflows preserve those values.
3. **East Asian font specification** — sets matching Latin (`<a:latin>`) and East Asian (`<a:ea>`) typefaces so PowerPoint does not fall back to system fonts such as Malgun Gothic or Gulim.
4. **Formatting inheritance** — places shared font size and color at the slide master, layout, or paragraph level instead of hardcoding every run, preserving downstream PowerPoint edits.
5. **Native bullets and hanging indents** — uses native PowerPoint bullets and hanging indentation (`marL`, `indent`) rather than bullet characters embedded in strings. It keeps `<a:buChar>` before `<a:defRPr>` as required by OpenXML to avoid repair warnings.
6. **Single-shape cards and containers** — combines the background and text into one editable shape with `vertical_anchor=TOP` and internal margins.
7. **Production-ready Python helpers** — provides `set_korean_font_and_lang`, `set_native_bullet`, and `create_korean_card` helper functions.

### Files

| File | Format | Use with |
|---|---|---|
| [AGENTS.md](AGENTS.md) | Plain Markdown without frontmatter | OpenAI Codex, Google Antigravity, Cursor, Windsurf, other agents, or a system prompt |
| [SKILL.md](SKILL.md) | [Agent Skills](https://docs.claude.com/en/docs/claude-code/skills) format with YAML frontmatter and an auto-trigger description | Claude Code, Google Antigravity, or any tool that supports the Agent Skills specification |

The files use different loading formats. Choose the one that matches how your agent loads instructions, and treat [SKILL.md](SKILL.md) as the source for the latest Agent Skill rules.

### Installation and setup

**Claude Code** (auto-discovered Agent Skill):

```bash
git clone https://github.com/sshin90/pptx-korean-formatting.git ~/.claude/skills/pptx-korean-formatting
```

**OpenAI Codex CLI** (or another tool that reads `AGENTS.md`):

```bash
git clone https://github.com/sshin90/pptx-korean-formatting.git
cat pptx-korean-formatting/AGENTS.md >> AGENTS.md   # merge into the project's AGENTS.md
```

**Google Antigravity**: supports both the Agent Skills format used by Claude Code (a skill folder containing `SKILL.md`) and a simpler global rules directory.

- Install as a skill:

  ```bash
  git clone https://github.com/sshin90/pptx-korean-formatting.git ~/.gemini/config/skills/pptx-korean-formatting
  ```

- Install as a global rule for all workspaces:

  ```bash
  mkdir -p ~/.gemini/config/rules
  curl -o ~/.gemini/config/rules/pptx-korean-formatting.md \
    https://raw.githubusercontent.com/sshin90/pptx-korean-formatting/master/AGENTS.md
  ```

**Cursor, Windsurf, and other agentic IDEs**: copy [AGENTS.md](AGENTS.md) into the tool's project rules file, such as `.cursorrules`, `.cursor/rules/`, `.windsurfrules`, or its workspace guidelines. Check the tool's documentation for the exact filename and location.

**Any other agent**: paste [AGENTS.md](AGENTS.md) into the system prompt or instruct the agent to read it before generating a Korean-language deck.

### License

[CC0 1.0 Universal](LICENSE) — public domain, no attribution required.
