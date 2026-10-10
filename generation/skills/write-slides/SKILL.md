---
name: write-slides
description: Write or outline talk slides in madolson's style. Short claim titles, graphics instead of text, points in speaker notes. Use when drafting a slide outline from a talk track, writing slide content or speaker notes, or reviewing a deck's wording.
---

# Writing slides

Madelyn talks over the slides, so the slides only hold what the room has to see. The baseline is `~/git/slides/slides/Scaling Up and Out - How Valkey Does Both.pptx`. Before writing anything, dump it with the `python3 dump_deck.py <deck>` script in that folder.

For template mechanics, plus the full content tenets those mechanics follow, read `~/git/slides/generation/slides-skill.md` ("Writing the Content" and "No Textual Asides"). This skill is the short version for drafting.

## Per slide

Each slide gets these things. Nothing else.

- **Point.** The one thing the slide has to land, in one sentence. It goes in the outline, not on the slide. If you can't write one, or it repeats a neighbour's, cut the slide. Every slide has to move the narrative forward. Fold a slide that only supports another into that slide's notes.
- **Title.** Usually the only sentence on the slide.
- **Graphic.** What the audience looks at: a chart, a diagram, a build, cards, or a single big number.
- **Notes.** Optional, and usually one to three short sentences.

Never write a "Says" line or a script. She says it out loud.

## Titles

- **A slide that argues gets a short claim as its title.** From the baseline: "Latency is a cliff", "Steady state is a lie", "You could overprovision", "Up: about 8x, and then it stops", "The host looks idle. The main thread is full."
- **A reference slide gets the thing's name as its title**, e.g. `CLUSTER MIGRATESLOTS` or `CLUSTER SLOT-STATS`.
- **Every frame of a build repeats the same title.** "Thread scalable command processing" runs across 7 frames.
- **No subtitles.**
- **Never put a comma in a slide title.** A title makes one point, and a comma hedges it or tacks on a second one. "Rust removes a class, not the job" becomes "Rust removes one class of bug". "No reproducer, no triage" becomes "No reproducer means no triage". The talk title is exempt, since it's locked by the CFP. So are quotes, which stay verbatim. The builder should fail the build on a comma in a title or statement.
- **Name the interesting point, not the evidence.** "Same bug, three reporters" describes the slide. "Finding bugs got easy" says why it matters.
- **Section headers name the act.** Use "Act 1: …", "Act 2: …", and so on, on `Valkey Insight`. The outline's `###` headings are the same acts in the same order. No separate list of talk-track items next to the slides.

## Text on the slide

- Labels name things on the graphic (`Shard 1`, `source`, `staging area`). They never explain it.
- Cards get a two or three word heading and at most one short line. For example: "Vendor neutral" / "A Linux Foundation project, BSD-3 licensed, so it stays open source."
- No caption under a chart or diagram telling the room what to conclude. That's her line.
- One number per slide, with its source in a small citation line. One cited benchmark per section is plenty.
- List slides, like the closing "architectural goals", get three rows of a few words each.

## Graphics

- **Prefer real data.** Use a real trace or a real benchmark result over an illustrative shape.
- **Bring one artifact back at each turning point.** In the baseline, the two-box "Vertical / Horizontal" slide returns three times (slides 8, 20, 31) with its state updated.
- **Put the build effort on the one diagram people should remember.** Add one element per frame and never remove one mid-build.
- When a slide needs code, use the real code from the commit, under 10 lines. Never write it from memory. Show the code first, then the diagram that explains it.
- **On rose, everything is black: text and icons.** Use `icon_paths(name, fill)` / `icon(..., on=fill)` so rose fills get the black icon variant. No icon on a blue fill.
- **Arrows are straight or right-angled, never diagonal.** Use `arrow(..., elbow=2)` to land on the top or bottom of a box, and `elbow=3` for box side to box side.
- **Charts:** one chart with one series per group beats small multiples. Put the series that carries the point in `ACCENT_1` and mute the rest. Use real data with the query that produced it in the notes.

## Speaker notes

- Include what the slide can't show: caveats ("Alpha. Do not oversell it."), test conditions ("Clients in the same availability zone as the server."), and why the slide is there.
- Use short declarative sentences or a few bullets, never a script. Leave the notes empty if there's nothing to add.
- Put sources and the exact wording of any claim she must get right here.

## Deck shape

- **Open:** the title slide, then straight to a symptom. Only add an about-me or a "what is Valkey" slide if she asks for one. Some decks deliberately skip both.
- **Close:**
	1. A **Recap** slide: one card per act, with "Act N" as the heading, the act's name in bold, and its one-line takeaway.
	2. One `Valkey Insight Bold` line that tells the room how to think about it going forward, e.g. "Assume bug-finding keeps getting cheaper."
	3. "Getting involved" (valkey.io, Slack, the next event).
	4. "Thank you! Let's connect." with her handles.

  Audience-specific actions go in the recap's notes.
- Valkey leads; AWS is context ("Valkey maintainer" before "Principal engineer at AWS").

## Outline format

The vault talk file (`content/talks/<talk>.md`) is the spec. Write the full slide description there first, and generate the deck from it afterwards. Older decks and builder scripts in `~/git/slides/generation/` are placeholders, not sources. When the outline and a deck disagree, the outline wins.

Open the outline with an `Audience:` line naming the event, the format and length, and who's in the room (for example, "Keynote, 10 minutes, mixed engineers and executives"). If you don't know, ask before drafting. The audience sets the density:
- **Keynote:** short, punchy slides that she sits on. Each one is a claim, one number, or one simple diagram, with almost nothing else on it. Use fewer slides, and keep the notes to sources and caveats.
- **Mixed technical and executive room:** attach a consequence to every insider term, or cut the term.
- **Technical session:** diagrams, code, and builds can carry more detail.

Use one `###` heading per PowerPoint section, in deck order: `Opening`, `Act 1: …`, `Act 2: …`, `Act 3: …`, `Close`. List the slides under each:

```markdown
### Act 1: <name>

- **<title, exactly as on screen>**
	- Point: <the one thing it lands>
	- Layout: <`Valkey *` layout name>. <Bullet slide / two-section / three-section / diagram / big number / insight>. <Build frames, if any.>
	- On screen: <every other piece of text on the slide, quoted. If it isn't listed, it isn't on the slide.>
	- Graphic: <what it looks like. Diagrams are native shapes. Images name their file in `generation/assets/`.>
	- SVG: <icons: existing names from `generation/icons/`, or new ones marked (new)>
	- Notes: <optional>
```

Choosing the layout:
- Use `Valkey Insight` for section headers and one-line statements, where the statement is the title.
- `Valkey Insight Bold` appears once per talk.
- Use `Valkey Big Number` for a single measurement. The number replaces the title.
- Use `Valkey List Rows` for bullet slides, three to five rows.
- Use `Valkey Two Box` / `Valkey Three Box` for two- or three-section comparisons. Their icons are baked in, so cards with custom icons are built with `card()` on `Valkey Title Only`.
- Use `Valkey Title Only` for diagrams, charts, code, and screenshots.

End the outline with a `## SVGs to generate` table: icon name, shape description, and which slides use it. Follow the icon rules in `slides-skill.md`. Screenshots and QR codes are PNGs, not SVGs. Make QR codes with https://madelynolson.com/qr-generator, using the default voxel-hexagon shape with the Valkey logo. Drive it in headless Chrome over CDP: set `#url`, dispatch `input`, read `#canvas.toDataURL()`. Decode the result with CoreImage `CIDetector` to confirm the URL before using it. Screenshot tweets with the `platform.twitter.com/embed/Tweet.html?id=<id>&theme=light` embed in headless Chrome at 2x, then crop.

## Building the deck

Generate the deck with a `slides/build_<talk>_deck.py` script in `~/git/slides`. `build_cve_tsunami_deck.py` is the worked example. Read `slides-skill.md` first.

**Use the placeholder template wherever possible.** The donor is `slides/Valkey placeholder slides.pptx`. In order of preference:

1. **A `Valkey *` layout, filled through its placeholders** with `on_layout()` + `fill()` and `drop_placeholders()` for unused ones. Use `list_rows()` for list layouts. Leave size, position, and colour to the layout.
2. **The template's primitives on `Valkey Title Only`:**
   - `chip()` for labelled boxes.
   - `node()` for servers and connections, with `mark=True` only on a real Valkey instance.
   - `arrow()`, `card()`, `source_note()`, `box_columns()`, `card_geometry()`.
3. **Custom shapes**, only for what neither covers. Examples: a dashed "freed" box, code panels.

Exceptions:
- **Box layouts.** `Valkey Two Box` / `Valkey Three Box` bake in their icons (scaling; performance, security, reliability), and a slide can't change them. When those icons don't fit, build the same design with `card()`.
- **Icons on blue fills.** The icons are Valkey-blue strokes, so they vanish on a blue fill. Don't put one there.

**Set PowerPoint sections to match the acts.** Donor decks carry their own sections (the placeholder deck's donor had "General", "Features", "Template examples"), and python-pptx doesn't know about them. `strip_slides()` clears them. Record each act's first slide index while building, then call `V.set_sections(prs, [(name, start), ...])` last, once the order is final. Section names match the outline's `###` headings exactly. `V.audit_sections()` fails the build if any slide is missing from a section, or a section names a slide that doesn't exist.

**Never resize an inherited placeholder.** Setting only `width` writes a partial `xfrm`, and PowerPoint misplaces the box even though every other tool says it's fine (`slides-skill.md` pitfall 21). Make the text fit the layout instead.

**Verify in this order:**
1. The builder's audits.
2. `mock_render.py`.
3. Open the deck in PowerPoint. The mock can't show layout type sizes, dashes, or arrowheads, or placeholder geometry problems.

PowerPoint holds the file open, so build to `/tmp` while it's open and install once it's closed.

**Hand edits in PowerPoint are real input, and a rebuild erases them.** Before installing any build over a deck that may have been edited, diff the installed deck against a fresh build: slide order, layout, every text frame, notes, picture and chart counts. Then port the edits into the outline and the builder, rebuild, and re-diff until only intended changes remain. Speaker notes are the speaker's voice, so keep their wording. Keep sources in the outline's `Source:` lines, not in the notes.

## Don't invent

Every number, quote, and claim needs a source the user gave or one you checked. Mark anything missing as `TODO: <what's needed>` rather than filling it in.
