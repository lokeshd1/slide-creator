# Markdown Slide Deck Claude Plugin

Create polished PowerPoint decks from Markdown, text, or Word documents inside Claude Code.

This plugin ships one Claude Code skill, `/markdown-slide-deck:create-deck`, backed by a self-contained Python script. It reads structured source content, splits it into slides, applies one of four visual themes, and can add simple automatic infographics for metrics, steps, and highlight lists.

## Contents

```text
.claude-plugin/plugin.json
skills/create-deck/SKILL.md
skills/create-deck/scripts/build_deck.py
skills/create-deck/references/markdown-format.md
skills/create-deck/references/themes.md
```

## Requirements

```bash
pip install python-pptx
```

For `.docx` input:

```bash
pip install python-docx
```

## Test Locally

From this plugin directory:

```bash
claude plugin validate .
claude --plugin-dir .
```

After Claude loads the plugin, invoke:

```text
/markdown-slide-deck:create-deck
```

## Direct Script Usage

```bash
python3 skills/create-deck/scripts/build_deck.py input.md --output out/deck.pptx --theme modern
```

Themes:

- `modern`
- `executive`
- `studio`
- `mono`

Disable generated visuals:

```bash
python3 skills/create-deck/scripts/build_deck.py input.md --output out/plain.pptx --no-infographics
```

Inspect parsed deck JSON:

```bash
python3 skills/create-deck/scripts/build_deck.py input.md --inspect --infographics
```

## Package As Zip

From this plugin directory:

```bash
zip -r markdown-slide-deck-claude-plugin.zip .
```

Claude Code can also test the unpacked directory directly with `claude --plugin-dir .`.
