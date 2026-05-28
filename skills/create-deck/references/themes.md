# Themes

Choose one theme per generated deck. The default is `modern`.

## Available Themes

- `modern`: default product/report theme with crisp panels and cool accents.
- `executive`: restrained business theme with warm accents.
- `studio`: more expressive editorial theme with brighter accents.
- `mono`: minimal black-and-white theme for dense work documents.

## Selection Guidance

- Use `modern` for product updates, launch reviews, planning docs, and general work decks.
- Use `executive` for leadership readouts, business reviews, board-style summaries, and strategy decks.
- Use `studio` for research readouts, concept pitches, creative reviews, and more expressive narrative decks.
- Use `mono` when content density matters more than visual color, or when the user wants a restrained document-like deck.

Automatic infographics inherit the selected theme colors.

## Commands

```bash
python3 scripts/build_deck.py source.md --output out/deck.pptx --theme modern
python3 scripts/build_deck.py source.md --output out/deck.pptx --theme executive
python3 scripts/build_deck.py source.md --output out/deck.pptx --theme studio
python3 scripts/build_deck.py source.md --output out/deck.pptx --theme mono
```
