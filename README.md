# pptx-korean-formatting

PowerPoint(`.pptx`/`.potx`)에 한글이 포함될 때 필요한 언어 태그, 단어 단위 줄바꿈, 동아시아 폰트, 서식 상속, 네이티브 불릿, 편집 가능한 카드 도형 규칙을 제공하는 Agent Skill입니다.

## 한국어

### 배포 원칙

이 저장소는 **Skill-first** 방식으로 배포합니다.

- [SKILL.md](SKILL.md)가 규칙의 유일한 원본입니다.
- Agent Skills를 지원하는 환경에서는 스킬로 설치하는 방법을 권장합니다.
- `AGENTS.md`, 전역 Rule, 시스템 프롬프트에 규칙 본문을 중복 저장하지 마세요.
- [AGENTS.md](AGENTS.md)는 규칙 사본이 아니라 `SKILL.md`를 가리키는 짧은 호환성 안내입니다.
- Skill을 지원하지 않는 환경에서만 `SKILL.md`를 해당 도구의 프로젝트 규칙 파일에 연결하세요. 전역 Rule 등록은 권장하지 않습니다.

이 방식은 관련 PPTX 작업에서만 전문 지침을 불러오고, 평소에는 불필요한 규칙이 컨텍스트를 차지하거나 다른 지침과 충돌하는 일을 줄입니다.

### 무엇을 하는가

1. 한글 텍스트의 언어를 `ko-KR`로 지정합니다.
2. 자동 줄바꿈을 유지하면서 `eaLnBrk="0"`, `latinLnBrk="0"`으로 단어 중간 줄바꿈을 방지합니다.
3. 라틴 문자와 동아시아 문자에 같은 폰트를 지정하여 한글 폰트 폴백을 막습니다.
4. 공통 서식을 문단·레이아웃·마스터 수준에서 상속시킵니다.
5. 문자열 불릿 대신 네이티브 불릿과 내어쓰기를 사용합니다.
6. 카드의 배경과 텍스트를 하나의 편집 가능한 도형으로 만듭니다.
7. `python-pptx`용 검증된 헬퍼와 최종 QA 체크리스트를 제공합니다.

전체 규칙과 구현 예시는 [SKILL.md](SKILL.md)를 참고하세요.

### 권장 설치

저장소를 사용하는 에이전트의 Skill 디렉터리에 복제합니다. 설치 위치는 제품과 버전에 따라 달라질 수 있으므로 해당 에이전트의 최신 문서를 우선하세요.

**Claude Code:**

```bash
git clone https://github.com/sshin90/pptx-korean-formatting.git ~/.claude/skills/pptx-korean-formatting
```

**Google Antigravity:**

```bash
git clone https://github.com/sshin90/pptx-korean-formatting.git ~/.gemini/config/skills/pptx-korean-formatting
```

**공통 Agent Skills 디렉터리를 지원하는 에이전트:**

```bash
git clone https://github.com/sshin90/pptx-korean-formatting.git ~/.agents/skills/pptx-korean-formatting
```

설치 후 에이전트의 Skill 목록에서 `pptx-korean-formatting`이 검색되는지 확인하세요.

### Skill 미지원 환경

Skill을 지원하지 않는 에이전트에서는 [SKILL.md](SKILL.md)를 프로젝트 범위의 규칙 또는 지침 파일에서 참조하세요. 규칙 본문을 복사하면 두 사본이 시간이 지나며 달라질 수 있으므로, 가능한 경우 파일 링크나 명시적인 읽기 지침을 사용하세요.

예시:

```md
한글이 포함된 PPTX 작업을 수행하기 전에
path/to/pptx-korean-formatting/SKILL.md를 읽고 적용한다.
```

전역 Rule은 Skill 탐색이나 프로젝트 범위 참조가 불가능한 경우에만 마지막 호환 수단으로 사용하세요.

### 라이선스

[CC0 1.0 Universal](LICENSE) — 퍼블릭 도메인, 저작자 표시 의무 없음.

---

## English

An Agent Skill for Korean-language PowerPoint (`.pptx`/`.potx`) formatting. It covers language metadata, word-boundary wrapping, East Asian fonts, inherited formatting, native bullets, and editable single-shape cards.

### Distribution policy

This repository uses a **Skill-first** distribution model.

- [SKILL.md](SKILL.md) is the single source of truth.
- Install the repository as an Agent Skill whenever the target agent supports Skills.
- Do not duplicate the full rules in `AGENTS.md`, global rules, or system prompts.
- [AGENTS.md](AGENTS.md) is only a small compatibility pointer to `SKILL.md`.
- For agents without Skill support, reference `SKILL.md` from a project-scoped rules file. Global rule installation is not recommended.

This keeps the guidance out of unrelated tasks and reduces stale copies and instruction conflicts.

### What it does

1. Tags Korean text with `ko-KR` language metadata.
2. Keeps automatic wrapping enabled while setting `eaLnBrk="0"` and `latinLnBrk="0"` to prevent mid-word breaks.
3. Assigns matching Latin and East Asian fonts to prevent Korean font fallback.
4. Preserves formatting inheritance at paragraph, layout, or master level.
5. Uses native bullets with proper hanging indents.
6. Builds cards as single editable shapes containing their own text.
7. Provides tested `python-pptx` helpers and a final QA checklist.

See [SKILL.md](SKILL.md) for the complete rules and implementation examples.

### Recommended installation

Clone the repository into the Skill directory used by your agent. Locations can change between products and versions, so prefer the agent's current documentation.

**Claude Code:**

```bash
git clone https://github.com/sshin90/pptx-korean-formatting.git ~/.claude/skills/pptx-korean-formatting
```

**Google Antigravity:**

```bash
git clone https://github.com/sshin90/pptx-korean-formatting.git ~/.gemini/config/skills/pptx-korean-formatting
```

**Agents supporting the shared Agent Skills directory:**

```bash
git clone https://github.com/sshin90/pptx-korean-formatting.git ~/.agents/skills/pptx-korean-formatting
```

After installation, verify that `pptx-korean-formatting` appears in the agent's Skill list.

### Agents without Skill support

Reference [SKILL.md](SKILL.md) from a project-scoped rules or instruction file. Prefer a file reference or an explicit instruction to read it instead of copying the full body, which can create stale divergent versions.

Example:

```md
Before working on a PPTX containing Korean text, read and apply
path/to/pptx-korean-formatting/SKILL.md.
```

Use a global rule only as a last compatibility option when neither Skill discovery nor a project-scoped reference is available.

### License

[CC0 1.0 Universal](LICENSE) — public domain, no attribution required.
