# pptx-korean-formatting

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

[MIT](LICENSE)
