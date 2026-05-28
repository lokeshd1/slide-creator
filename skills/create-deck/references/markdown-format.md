# Markdown Format

Use normal Markdown with these conventions. Plain `.txt` files are treated as Markdown-like text. `.docx` files are converted through paragraph styles when `python-docx` is installed.

## Slide Breaks

Separate slides with `---` on its own line.

```markdown
# Title Slide

Subtitle text

---

## Status

- Point one
- Point two
```

## Titles

The first heading on each slide becomes the slide title.

## Speaker Notes

Put notes at the end of a slide:

```markdown
Notes:
Mention that metrics are preliminary.
```

`Speaker notes:` is also supported.

## Lists

Bullets become slide bullet text. Numbered lists can become automatic process-flow infographics.

```markdown
1. Decide
2. Build
3. Launch
```

## Images

Use Markdown image syntax:

```markdown
![Architecture](assets/architecture.png)
```

Prefer image paths that resolve from the source file directory.

## Front Matter

Optional front matter can define deck metadata:

```markdown
---
title: Quarterly Product Update
owner: Product Team
---
```

The `title` field becomes the deck title.
