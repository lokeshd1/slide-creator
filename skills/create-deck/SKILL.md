---
description: Create polished themed PowerPoint PPTX decks from Markdown, text, or Word documents. Use when the user asks to turn notes, Markdown files, product updates, reports, strategy docs, meeting notes, or document drafts into editable slides with optional automatic infographics.
---

# Markdown Slide Deck

## Overview

Use this Claude Code skill to convert structured Markdown or documents into themed `.pptx` slide decks. Prefer this skill when the desired output is an editable PowerPoint deck, especially when the source contains headings, bullets, numbered lists, speaker notes, metrics, or status updates.

## Quick Workflow

1. Find or create a source file in Markdown, text, or `.docx` format.
2. Choose a theme: `modern`, `executive`, `studio`, or `mono`.
3. Build the deck with `scripts/build_deck.py`.
4. Verify the output exists and, when possible, open it with `python-pptx` to confirm the slide count.
5. Return the absolute path to the generated `.pptx`.

## Build Command

From this skill folder, run:

```bash
python3 scripts/build_deck.py /path/to/source.md --output /path/to/output.pptx --theme modern
```

Useful options:

```bash
python3 scripts/build_deck.py source.md --output out/deck.pptx --theme studio
python3 scripts/build_deck.py source.md --output out/plain.pptx --no-infographics
python3 scripts/build_deck.py source.md --inspect --infographics
```

The script is bundled with this skill and is self-contained except for Python package dependencies. It requires `python-pptx` to build PPTX files and `python-docx` only when reading `.docx` input.

## Authoring Guidance

Use `---` as the slide break. The first heading on each slide becomes the slide title. `Notes:` or `Speaker notes:` becomes presenter notes.

For detailed Markdown syntax, read `references/markdown-format.md`.

For theme selection and visual guidance, read `references/themes.md`.

## Infographics

Automatic infographics are enabled by default:

- numbers, percentages, currency, and multipliers become metric cards
- numbered lists become process flows
- compact bullet lists become highlight tiles

Disable this with `--no-infographics` when the source is dense, legalistic, or already includes manually designed visuals.

## Quality Bar

Keep slides concise. If a source slide has too many bullets or paragraphs, split it into multiple slides before building. Prefer clear titles that state the message, not just the topic.

After building, verify with:

```bash
python3 -c 'from pptx import Presentation; import sys; p=Presentation(sys.argv[1]); print(len(p.slides))' /path/to/output.pptx
```

## Installation Notes

This folder is part of a Claude Code plugin. Test the plugin from the plugin root with:

```bash
claude plugin validate .
claude --plugin-dir .
```

After loading, use `/markdown-slide-deck:create-deck` or let Claude invoke the skill automatically.
