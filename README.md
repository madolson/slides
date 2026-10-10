# slides

Madelyn Olson's talk decks, and the code that generates them.

| Folder | What's in it |
|---|---|
| `slides/` | Generated decks. Don't hand-edit them: a rebuild overwrites the file. |
| `generation/` | Builders, the Valkey template tooling, icons, assets, donor decks, and the slide skills. |
| `descriptions/` | One markdown file per deck: structure, sources, and decisions. |

## Decks

| Deck | Builder | Description |
|---|---|---|
| `Valkey placeholder slides.pptx` | `build_placeholder_deck.py` | `valkey-placeholder-slides.md` |
| `Scaling Up and Out - How Valkey Does Both.pptx` | `build_nerdearla_deck.py` | `scaling-up-and-out-deck.md` |
| `Open Source Maintainers vs The CVE Tsunami.pptx` | `build_cve_tsunami_deck.py` | `cve-tsunami-deck.md` |
| `The Boring Layer.pptx` | `build_boring_layer_deck.py` (base deck: `~/Downloads/oss_elc_EU26.pptx`) | none yet |
| `Percona Caching Deck.pptx` | hand-built | none |
| `Percona Live 2026 Keynote.pptx` | hand-built | none |

The CVE deck is built from the placeholder deck, so rebuild in this order:

```bash
python3 generation/build_placeholder_deck.py
python3 generation/build_cve_tsunami_deck.py
```

Requires `python-pptx`, Pillow, lxml, and Google Chrome (for SVG fallbacks and mock renders).

## Skills

- `generation/slides-skill.md`: the Valkey template, its mechanics, and the content tenets.
- `generation/skills/write-slides/SKILL.md`: how Madelyn writes slides. `~/.claude/skills/write-slides` is a symlink to it.
