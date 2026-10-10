# Slides Skill: Authoring Valkey-Themed Slides

How to build and edit slides in the Valkey template. This file is about the template and the
technique, not about any one deck. Deck-specific notes live next to the deck in `descriptions/*.md`.

All facts below were verified against the four decks in `slides/` (Valkey community deck, The
Vertical Scaling wall, Percona Caching Deck, Percona Live 2026 Keynote). The theme is identical in
all four, so anything in the Template section is safe to rely on.

## Before You Touch a Deck

1. **Check for a lock file.** `ls slides/` and look for `~$<name>.pptx`. That means PowerPoint has
   the file open. Your `prs.save()` will be silently clobbered the next time the user hits save,
   and PowerPoint may write its own stale in-memory copy over the file at any moment. Ask them to
   close it first, and do not use an open deck as a template donor.
2. **These pptx files are almost all untracked in git.** Only `Valkey community deck.pptx` is
   committed. There is no undo for the others. Before any destructive step, copy the file aside, and
   consider suggesting the user commit first.
3. **Dump the deck before editing.** Walk every shape, position, run, chart series and speaker
   note, and read it back. Never edit a slide you have not dumped: run structure and shape names
   are not guessable.
4. **Copy to `/tmp` and test there first** for anything structural (adding, deleting, reordering
   slides, replacing chart data). Save, reopen, re-dump, confirm, then apply to the real file.

Requires `python-pptx` (`pip3 install python-pptx`), tested at 1.0.2, plus Pillow and Google
Chrome for the icon and rendering tooling.

## Tooling in This Repo

| Script | What it does |
|--------|-------------|
| `generation/pptx_svg.py` | `add_svg(slide, svg, png_fallback, left, top, w, h)` to embed an SVG. |
| `generation/pptx_layout.py` | Add slide layouts to the master, which python-pptx cannot do. See Adding a Layout to the Master. |
| `generation/build_placeholder_deck.py` | Re-runnable builder for the placeholder deck. Good worked example of cards, icon embedding, and building a deck from scratch. |
| `generation/build_nerdearla_deck.py` | Re-runnable builder that *edits a donor* rather than starting empty. Worked example of slide cloning, reordering, chart data edits, and rewriting existing runs without losing their formatting. |
| `generation/make_band_art.py` | Derives the title slide's band artwork from a `~/git/valkey-svgs` banner. See Banner Artwork. |
| `generation/icons/` | 18 reusable Valkey-blue outline icons, plus 4x PNG fallbacks in `icons/png/`. |

## Template Facts

### Canvas

16:9, `12192000 x 6858000` EMU (13.333in x 7.5in). 914400 EMU per inch, 12700 per point.

### Theme Colors

Always use `MSO_THEME_COLOR.*`, never hardcoded RGB. Identical across all decks.

| Token | Hex | Use |
|-------|-----|-----|
| `DARK_1` | `#000000` | Primary text |
| `DARK_2` | `#232F3E` | Secondary text, subtitles, column headings |
| `LIGHT_1` | `#FFFFFF` | White; text on dark backgrounds (dividers, title, thank-you) |
| `LIGHT_2` | `#F1F3F3` | Neutral card background |
| `ACCENT_1` | `#6983FF` | **Valkey blue.** Bullets, highlighted numbers, accent bars, links |
| `ACCENT_2` | `#1A2026` | Dark accent |
| `ACCENT_3` | `#632637` | Dark maroon. **Reads brown, not red** — see below |
| `ACCENT_4` | `#30176E` | Purple accent |
| `ACCENT_5` | `#BCB5E7` | Light purple card fill |
| `ACCENT_6` | `#E0A2AF` | Brand rose, Pantone 494 C. **Patched in**, replacing the stock peach `#FBD8BF` |

Hyperlink and followed-hyperlink are both `#6983FF`.

**Official card colours, and where they live in the theme:**

| Colour | Hex | RGB | CMYK | Pantone | Theme slot |
|---|---|---|---|---|---|
| Rose | `#E0A2AF` | 224 162 175 | 4 45 15 0 | 494 C | `ACCENT_6`, patched in |
| Deep purple | `#30176E` | 48 23 110 | 95 100 20 14 | 2685 C | `ACCENT_4`, already correct |
| Card text | `#1A2026` | 26 32 38 | | | `ACCENT_2`, already correct |

**Prefer patching the theme over hardcoding RGB.** The stock `accent6` was peach; rewriting that one
scheme entry makes the rose a real theme colour, so it shows up in PowerPoint's colour picker and
every reference stays a theme reference. python-pptx does not parse the theme part, so edit the blob:

```python
for part in prs.part.package.iter_parts():
    if str(part.partname) == "/ppt/theme/theme1.xml":
        xml = part.blob.decode("utf-8")
        xml, n = re.subn(r'(<a:accent6>\s*<a:srgbClr val=")[0-9A-Fa-f]{6}(")',
                         r"\g<1>E0A2AF\g<2>", xml, count=1)
        part._blob = xml.encode("utf-8")
```

Anchor the regex on the `clrScheme` entry and assert exactly one substitution: the same hex may
appear as an explicit `srgbClr` elsewhere, and those are not yours to change.

On the deep purple, `#1A2026` text is unreadable — pass `LIGHT_1`, and note the blue concept icons
also lose contrast on it. A dark fill needs white icon variants, which do not exist yet.

**The palette has no red.** `ACCENT_3` is the closest and it reads as brown, which is wrong for a
failure or abort state. This is the one documented exception to theme-colours-only: use an explicit
`RGBColor` for failure states. `#D32F2F` reads clearly red next to the Valkey blue; `#C0392B` is a
darker brick alternative. Keep the exception to failure semantics, and do not start hand-picking
RGB for anything the palette already covers.

Supporting both means a small shim, since python-pptx has separate setters:

```python
def _set_colour(target, colour):
    if isinstance(colour, RGBColor):
        target.rgb = colour
    else:
        target.theme_color = colour
```

Route every colour assignment through it: run fonts, shape fills, shape lines, **and connector
lines**. Missing one raises `ValueError: (211, 47, 47) is not a valid MSO_THEME_COLOR_INDEX`, which
is the giveaway that a code path still assumes a theme colour.

**Everything on `ACCENT_6` (rose) is black: text and icons.** Text is `DARK_1`, not `ACCENT_2`, not `DARK_2`, never white. Icons use the black variants in `generation/icons/dark/`, which `make_dark_icons()` generates from the blue originals (stroke `#000000`). A Valkey-blue icon on rose reads as a faint smudge. This applies to chips, cards, layout placeholders and diagram labels. `card()`, the Two Box layout and `icon_paths(name, fill)` enforce it, so pass the fill an icon sits on. Never put a blue icon on a blue fill either.

For two-way comparisons, `ACCENT_5` vs `ACCENT_6` reads clearly and stays on-brand. Reserve
`ACCENT_1` for the thing you want the eye to land on. If everything is blue, nothing is.

### Fonts

**Arial, and nothing else.** One font, in the theme and on every run.

The stock decks ship two theme fonts that are not Arial, and both are a liability:

- **Play** — theme major font, `+mj-lt`. A Google font. The master applies it to titles *and* to
  the body placeholder, so body text is not Century Gothic by default.
- **Century Gothic** — theme minor font, `+mn-lt`.

Neither is reliably installed on a conference laptop, and PowerPoint substitutes **silently**: a
title that looks right in your editor reshapes in the room, and nothing warns you. Arial is on
every machine that will ever open these files. So convert the theme rather than working around it,
and pin Arial on runs too, because an unpinned run is exactly the one that moves.

Three things have to happen or one of them leaks:

1. Rewrite `typeface="Play"` and `typeface="Century Gothic"` to `Arial` across **every** part —
   theme font scheme, both masters, every layout, `presentation.xml`'s default text style, and the
   `buFont` on every bulleted paragraph. python-pptx does not parse the theme part, so patch part
   blobs for unparsed parts and elements for parsed ones. Patching the blob of a part python-pptx
   has already parsed is thrown away on save, because the blob is regenerated from the element.
2. Set `run.font.name` on every run on every slide, including carried-over donor slides.
3. **Audit, and fail the build.** Assert no run is off-font and no banned typeface survives
   anywhere in the package. `retypeface()`, `pin_font()` and `audit_fonts()` in
   `build_nerdearla_deck.py` are the working versions of all three.

Run the font pass **last, over the finished deck**, so a slide added later cannot be missed.

**Set the font on titles too, not just body.** The classic symptom is a deck where most titles look
right and a couple look different: those slides had body runs pinning Arial while their titles
inherited the theme major font. Audit it rather than trusting it — collect `run.font.name` across
every title and assert the set has exactly one member.

### Type Scale

Not from the master, despite where it looks like it lives: `slideMaster1.xml`'s `titleStyle` and
`bodyStyle` are 14pt across all four levels. Every size below comes from an `a:lstStyle` on a
*layout* placeholder, which the slide's placeholder then inherits. That matters when you add a
layout: putting the sizes in its `lstStyle` is not a workaround, it is the same mechanism the stock
layouts use.


| Role | Size | Weight | Color |
|------|------|--------|-------|
| Slide title | 32pt | bold | `DARK_1` |
| Body level 1 | 24pt | regular | `DARK_1` |
| Body level 2 | 20pt | regular | `DARK_1` |
| Body level 3 | 18pt | regular | `DARK_1` |
| Divider / Quote / title-slide title | 44pt | regular | `LIGHT_1` |
| Column heading (`… with Headings` layouts) | 20pt | bold | `DARK_2` |
| Column body (`… with Headings` layouts) | 18pt | regular | inherit |
| Source citation | 10pt | italic | `DARK_1` |

Two things the table does not capture:

- **Leave placeholder titles unstyled** so they inherit 32pt. Most title placeholders across the
  decks set no explicit run size. The overrides that exist are 48pt (opening title slides) and
  28pt.
- **Hand-built title shapes use 28pt bold**, consistently, on every custom slide in every deck.
  When you fake a title with a rectangle on a `Title Only` or `Holding` slide, use 28pt bold at
  the standard title coordinates, not 32pt.

Card text does not have one house size; it scales to the card. Observed: a small 14pt bold eyebrow
label over 24pt bold rows for a full-width two-card comparison, and 18pt bold titles with 13pt
body for denser three-column cards. Pick sizes for the density you have and keep the two cards
identical.

### Standard Geometry

Every slide must line up on this grid. Content is inset 609600 EMU (0.667in) from the left.

```python
TITLE_LEFT   = 609600
TITLE_TOP    = 400019
TITLE_W      = 11062800
TITLE_H      = 443198

CONTENT_LEFT = 609600
CONTENT_TOP  = 1445931
CONTENT_W    = 11062800
CONTENT_H    = 4565111
```

The Valkey logo sits in the master at `(10389732, 5943597)`, size `1434431 x 443198`. Anything
pinned to the bottom of a slide (source lines, footnotes) should align to `y = 5943597` so it sits
on the same baseline as the logo.

**The content box is not symmetric.** Left inset is `609600` but width is `11062800`, so it
overhangs the right margin by `90000` EMU (0.1in). Text does not reveal this, but a grid or a row
of evenly divided cells does: the last column sits closer to the right edge than the first does to
the left. For grids, centre on the slide instead:

```python
grid_w = SLIDE_W - 2 * CONTENT_LEFT      # 10972800, symmetric
cell_w = grid_w // cols
grid_left = CONTENT_LEFT
```

The same applies vertically. `CONTENT_TOP + CONTENT_H` bottoms out at `6011042`, so a block anchored
to `CONTENT_TOP` sits low: its centre lands at `3728486` against a page centre of `3429000`, a third
of an inch off. For a slide that is mostly one block, measure the block and centre it on the page:

```python
block_h = (rows - 1) * row_pitch + row_content_h
grid_top = (SLIDE_H - block_h) // 2
```

Then assert it: the top gap and bottom gap should come out equal. Eyeballing will not catch 0.33in.

## Choosing a Layout

**Never hardcode a layout index.** Indices differ in every deck: `Title Only` is index 6 in the
community deck, 8 in The Vertical Scaling wall, 4 in the Percona Caching deck, 2 in the Percona
keynote. **The placeholder deck has none of the stock layouts at all** — its master is the eleven
Valkey layouts and nothing else. Resolve by name:

```python
layout = prs.slide_layouts.get_by_name("Title + Cont")
if layout is None:
    raise SystemExit("layout missing from this deck")
slide = prs.slides.add_slide(layout)
```

Layout names are consistent where they exist, but not every deck has every layout. Check first.

### Stock Layouts

Present in the community deck, The Vertical Scaling wall, and both Percona decks. **Not in the
placeholder deck**, which replaced them; see The Valkey Layouts.

| Layout | Placeholders | Use it for |
|--------|-------------|------------|
| `Title + Cont` | title(0), body(1) | The default. Title plus bulleted content. |
| `Title Only` | title(0) | Custom slides: diagrams, cards, anything hand-built. |
| `Divider` | title(0), 44pt on dark | Section breaks. |
| `Title + Sub + Cont` | title(0), body(1), sub(2, 11pt bold) | Title needing a kicker line above the content. |
| `Title + Subtitle` | title(0), body(1) | Title with a one-line subtitle, no body. |
| `Title + 2x Cont with Headings` | title(0), heading(1,3), body(2,4) | Two-column comparison with real headings. |
| `Title + 3x Cont with Headings` | title(0), heading(1,3,5), body(2,4,6) | Three-column comparison. |
| `Title + Cont + Right Img Small` | title(0), body(1), picture(2) | Screenshot or product shot on the right. Picture placeholder is full-bleed right: `(7884543, 0)` size `4307456 x 6858000`. |
| `Quote` | title(0, 44pt), body(1,2) | Pull quote with attribution. |
| `Code` | title(0) only | Code slides. Add the code as your own text box; there is no body placeholder. |
| `Full Image` | picture(2), full bleed | Edge-to-edge image. |
| `Video` | media(2), full bleed | Embedded video. |
| `Title + 2x Speaker` | title(0), body(1–5) | Opening slide, up to two presenters. |
| `Thank You 2x` | body(1–4), no title | Closing slide, up to two presenters. |
| `Holding` / `DEFAULT` | none | Blank canvas for fully custom work. |

Prefer the `… with Headings` layouts over hand-built cards when the content is genuinely a
two- or three-column comparison. They inherit the type scale for free.

### The Valkey Layouts

`build_placeholder_deck.py` builds twelve layouts and then **deletes all fifteen stock ones**, so the
placeholder deck's master is only these. They are the designs this repo used to hand-draw onto a
slide every time. All are white-background, all put decoration in the layout and text in
placeholders.

| Layout | Placeholders | Use it for |
|--------|-------------|------------|
| `Valkey Title Only` | title | The foundation, not a design. Every custom slide (diagrams, anything hand-built) sits on this. Replaces the stock `Title Only`. |
| `Valkey Title Band` | title, subtitle, 2x (speaker, role), band image | Title slide, full-height band down the right third. One or two presenters side by side (idx 2/3 and 4/5). Default band art is `assets/band-security-shield-clean.jpg`; the band image picture placeholder (idx 10) overrides it per slide. Build it with `title_band()`, which drops unused placeholders and re-places the white lockup over a custom image. |
| `Valkey Title Top Banner` | title, subtitle, speaker, role | Title slide, navy banner across the top. |
| `Valkey Title Bottom Banner` | title, subtitle, speaker, role | Title slide, centred lockup, speaker in a blue bottom banner. Bookends `Valkey Thank You`. |
| `Valkey Insight` | statement, caption | One assertion on white under an accent rule. No title: the statement is the title. |
| `Valkey Insight Bold` | statement, caption | Same, full-bleed Valkey blue. Once per talk. |
| `Valkey Big Number` | number, label, source | One measurement at 130pt. |
| `Valkey List Rows` | title, 5 rows | The default list treatment. Each row is one placeholder, so unused rows delete cleanly. |
| `Valkey List Rows with Image` | title, 5 rows, picture | Same rows, narrowed, with a **square** picture placeholder on the right. QR codes, logos, screenshots. |
| `Valkey Two Box` | title, 2x (heading, body) | Two-up cards, `ACCENT_5` against the rose, icons baked in. |
| `Valkey Three Box` | title, 3x (heading, body) | Three-up cards, `LIGHT_2`. Three *options* or a comparison. |
| `Valkey Three Features` | title, version, 3x (heading, body) | What is new in a release: a version eyebrow, then three **numbered** features on white. Numerals and rules are static. |
| `Valkey Thank You` | closing, 2x (speaker, 3 contacts) | Closing slide. One or two presenters side by side (idx 2/3-5 and 6/7-9). Build it with `thank_you()`, which drops unused placeholders. |

One layout per design, not one per palette variant. Recolouring a card on a slide is two clicks;
three near-identical gallery entries is clutter. The deep purple is the exception worth knowing
about and it is deliberately absent: it needs `LIGHT_1` text and white icons, so it is not a fill
swap.

### Removing the Stock Layouts

```python
dropped = strip_stock_layouts(prs, "Valkey ")   # after every slide is on one of yours
```

Order is not optional: **build the replacements and every slide first, delete last.**
`SlideLayouts.remove()` raises `ValueError` on a layout that is the basis for a slide, and that
refusal is the safety net. If it fires, a slide is still on a stock layout and the fix is to rebuild
it on one of yours, not to skip the delete. This is the whole reason `Valkey Title Only` exists: 19
diagram slides were on the stock `Title Only`.

Removal drops the master→layout relationship, so each layout part and any media only it referenced
falls out of the package on save. The stock layouts each carry a full-bleed background picture, so
this is most of why the placeholder deck went from 894 KB to 541 KB. It also removes the four shapes
in the deck that overhung the canvas, all of them Google-Slides-exported backgrounds in stock
layouts.

Two things to know afterwards:

- **Pasting a slide in from another Valkey deck drags its layout back in.** PowerPoint brings the
  source layout along, so the master grows a `Title + Cont` again. Re-running the builder is what
  cleans that up; nothing detects it automatically.
- **Assert it, do not assume it.** A removed layout can leave an orphaned part in the zip if the
  relationship graph is wrong. Compare the count of `ppt/slideLayouts/slideLayoutN.xml` entries in
  the package against `len(prs.slide_layouts)`, and check that every slide's
  `slide_layout.name` is one of yours.

## Adding a Layout to the Master

python-pptx cannot: `SlideLayouts` exposes only `get_by_name`, `index` and `remove`. That is the
whole reason this repo used to hand-draw reusable frames onto slides. `generation/pptx_layout.py`
closes the gap by building the part directly.

```python
from pptx_layout import new_layout, drawing_canvas, add_placeholder, audit

layout = new_layout(prs, "Valkey Two Box")     # empty, registered on the master
canvas = drawing_canvas(layout)                # slide-shaped, so the slide helpers work
rect(canvas, x, CONTENT_TOP, w, h, fill=MSO_THEME_COLOR.ACCENT_5)
add_placeholder(layout, "body", 1, x, y, w, h, "Heading one", 18, bold=True)
audit(prs, ["Valkey Two Box"])
```

`drawing_canvas` is the piece that saves the work: `layout.shapes` is a read-only `LayoutShapes`,
but `SlideShapes` only needs an spTree and a parent with a `.part`, so wrapping it lets `rect()`,
`text_box()`, `logo()` and `add_svg()` target a layout unchanged.

Three things have to line up or PowerPoint reports a repair, and it is always one of these:

1. The layout part exists with content type `PML_SLIDE_LAYOUT`.
2. A relationship in **each** direction between layout and master. The layout→master rel is what
   makes `SlideLayout.slide_master` resolve; the master→layout rel is what the id list points at.
3. A `p:sldLayoutId` entry in the master's `sldLayoutIdLst`, with an `id` in
   `2147483648 .. 4294967295`.

`SlideLayoutPart.load`'s signature is `(partname, content_type, package, blob)` — the same argument
order trap as `Part()` in `pptx_svg.py`, and it fails deep inside lxml rather than where you made
the mistake.

### What Belongs in a Layout, and What Cannot

- **Decoration goes in the layout, text goes in placeholders.** Static layout shapes are inherited
  and drawn on the slide but cannot be selected there. Anything a presenter edits has to be a
  placeholder or it is unreachable.
- **The layout's placeholder text is prompt text, not content.** A new slide comes up empty and
  correctly styled; PowerPoint does not copy the layout's words. Set `hasCustomPrompt="1"` and the
  prompt replaces "Click to add text", which is worth doing: "Heading one" beats the generic.
- **Per-run formatting cannot live in a layout.** A blue highlighted number is two runs with
  different colours; `lstStyle` styles a whole level. Keep that in the generator.
- **Static layout shapes draw *beneath* slide shapes.** A placeholder therefore cannot sit on top of
  layout decoration it is meant to sit beside: a full-width row placeholder covers the accent bar at
  its left edge. Inset the placeholder past the static shape instead. In the list layout the bar is
  static at `CONTENT_LEFT` width `BAR_H`, and the row placeholder starts at `CONTENT_LEFT + BAR_H`.
- **Do not fake a one-sided accent with a gradient fill.** A DrawingML outline has no per-side
  control, so a two-stop `a:gradFill` looks like a clever way to get an accent edge inside a single
  shape, and it is wrong twice: PowerPoint then shows the shape as gradient-filled in its own UI,
  which is not what a flat card is, and the hard stop it relies on is fragile (pitfall 12). Use two
  shapes. `solid_fill()` gives the placeholder its flat card, and the accent is its own static shape.
- **What a placeholder's own fill buys you is deletion.** The card travels with the text, so removing
  a row on a slide takes its background with it and leaves only the hairline bar, not an empty card.
- **Sizes go in `a:lstStyle`,** not on the layout's own runs. See Type Scale.
- `pptx_layout.audit()` catches the failures that are silent otherwise: an unregistered layout, a
  layout that does not resolve back to the master, duplicate `(type, idx)` pairs, duplicate shape
  ids, and a layout with no placeholders at all. Run it before saving.

## Writing the Content

Eleven tenets. They are about what makes a *talk* work, and they outrank anything mechanical below
them. Most were reverse-engineered from the best moments in the decks in this repo, so they are
descriptive of what already works, not aspirational.

### Narrative spine

1. **Open on a symptom, not a subject.** Earn the first five minutes with evidence that something
   hurts. The audience needs a reason to care before it can hold mechanism. A talk that opens with
   "what X is" has already lost the room it needed to convince.
2. **Give the talk one specific opponent.** "The main thread is full" is something you can point at
   for thirty minutes; "performance" is not. Talks that wander usually have a topic where they needed
   a problem. One opponent, named concretely, stated early.
3. **Make progress visible.** The audience should be able to tell, at any point, what has been
   settled and what has not. Reuse one artifact at each turning point so it doubles as a progress
   marker: they learn to read it once, and it costs nothing to bring back.

   *How* you show progress is a choice per talk, not a rule. Striking out a limitation you have just
   disproved is one way, and the two-box build in `descriptions/scaling-up-and-out-deck.md` is a worked
   example. Re-running a benchmark, filling in a scorecard, or returning to a diagram with more of it
   answered all do the same job. Pick the one the material suggests.
4. **Close with a recap, then on something a person could repeat.** The recap slide is one card per
   act: its name and its takeaway, so the room sees the shape of the talk once more. It is not the
   ending: nobody remembers a list they have already heard. What works is one line that survives the walk to the next session, and
   what form it takes depends on the talk: a bounded answer to a question the talk raised, a single
   number, a named limitation. Do not force any particular device.

### Per slide

5. **Build, don't reveal.** One element added per frame, **strictly increasing**, same title
   throughout. Never remove an element mid-build: the audience re-parses the whole diagram, and a
   frame that retracts something it just added reads as a mistake. Two identical frames with a
   retraction between them is the specific shape to look for. Duplicating the slide and adding one
   element is the house method for architecture and threading diagrams.

   A frame set where the shape count stays constant and only the *labels* change is a different
   device, and a good one: same diagram, changing state. Do not confuse it with an additive build or
   try to make it increase.
6. **Design for the one diagram they will redraw on a whiteboard.** Audiences forget mechanism and
   remember shape. Pick the single figure that is the takeaway and give it the most build frames.
   Everything else can be a label. This means deliberately unequal effort across sections, which
   fights the instinct to treat them evenly.
7. **Labels on the shape, never captions under it.** If a box needs a sentence beneath it to make
   sense, fix the box. **One idea per slide**: if a slide needs two sentences to summarize, it is two
   slides. A slide is a visual aid, not a document, and the audience is either reading or listening,
   never both. Prefer a plain title over a title plus a subtitle plus a caption. When in doubt, cut.

   This is the rule that gets broken most, so it has its own section: see **No Textual Asides**.

### Honesty, which is where technical credibility comes from

8. **Show how you would know, not just what you fixed.** The slides about *observing* the problem
   carry more weight than the ones about solving it. A talk with mechanism and no measurement reads
   as marketing.
9. **One number per beat, and every number has a source.** The discipline is using *fewer*, not
   citing more. Four speedup figures on one slide compete with each other; one does not. Highlight
   percentages, throughput, latency and counts; never ordinals or version numbers.
10. **Say what you did not fix.** A named limitation buys more trust than any benchmark, and it is
    the honest ending for most infrastructure work.
11. **Setup is not argument.** Method, environment, tooling and configuration are how you got the
    result, not the result. They belong in speaker notes, an appendix, or a link. The test is whether
    a section moves the argument, and effort already spent on it is not an answer to that question.

Two of these are genuinely contested rather than settled, and knowing which matters:

- **#1 fights a mixed audience.** A room that is not all developers may need "what is this thing"
  before it can read a symptom. The resolution is usually a single card slide early, not a section.
- **#6 argues for unequal build effort,** which will feel like neglecting a section. It is.

Tenets 3 and 4 were narrower in an earlier draft, naming the strikethrough build and a closing
question as the way to do each. Both were one deck's device promoted to a rule. If a tenet only
describes how one talk happened to solve something, it is an example, and it belongs in that deck's
notes.

### On a slide that argues, the title carries the argument

"Throughput is a cliff" beats "Throughput": state the claim so the slide has a point even if you say
nothing.

**This is not a rule about every title, and it is not an audit.** It applies to slides that make a
case. It does not apply to **reference slides**, where the subject *is* the name (`CLUSTER
MIGRATESLOTS` beats any sentence about it); **build frames**, which must repeat one title across
every step, because the repetition is what says "same idea, developing"; or **section headers**,
whose job is to say where you are.

Do not go title-hunting across a finished deck. A checker for this flags roughly a third of any real
deck, is almost all false positives, and the fix it suggests is longer titles, which collides with
tenet 8.

### No Textual Asides

Tenet 8, in full, because it is the one that gets broken most and breaking it feels like helping. Do
not add a footnote, a caption, a gloss, a parenthetical, a "clarifying" subtitle, or an explanatory
line under a chart or diagram.

**The test is what the text does, not where it came from.** An aside *comments on* something else on
the slide. A label *names* something on the slide. Provenance decides whether a claim is safe to
make; it never makes an aside belong. A donor deck's own "Needs `cluster-slot-stats-enabled yes`" is
an aside, and being the donor's did not save it.

| Keep: it names something | Cut: it comments on something |
|---|---|
| A label inside or on a shape (`Shard 1`, `source`, `staging area`) | Prose floating under the shape explaining what the shape is doing |
| Axis titles, chart legends, data values, phase chips | A sentence under the chart telling the audience what to conclude |
| A source citation: it points at evidence rather than adding a claim | A footnote supplying a gloss the source itself leaves to the speaker |
| A column heading over a table (`cpu-usec`) | A line under the table saying how to turn the feature on |

Two mechanical checks separate them:

1. **Is anything behind it?** Look for a filled shape at the same left/top/width/height. A chip has
   one; a floating caption does not. Do **not** test this with `other is not shape` —
   `slide.shapes` hands out fresh wrappers, so that is always true (pitfall 16). Compare `_element`.
2. **Does the title already say it?** Two asides once read "Half the machine is doing nothing." and
   "One core, and it is full. That is the ceiling." under a slide titled *The host looks idle. The
   main thread is full.* The same point, three times, in three sizes.

The failure mode is specific: an asterisk, a bare number or an unexplained term looks like an
omission, so the instinct is to supply the missing gloss. It is not an omission. A speaker's deck is
a script for a person who is going to talk, and the gloss is the line they say out loud. Writing it
on the slide takes their line and stakes a claim they did not make, in their name, to a room.

Invented asides, all from one deck, none traceable to anything:

| Invented | Why it is wrong |
|---|---|
| "A ~1 ms floor, with repeated excursions past 15 ms. Nothing changed about the offered load." | A reading of someone else's trace, plus a claim about their workload. |
| "Single node, clients in the same availability zone." | The source records the invocation and nothing about the host. |
| "One key per MIGRATE: ~213K round trips." | The round-trip count was derived by the build script, not by the cited post. |
| "\* Command execution. Everything around it is already threaded." | Closes a loop on an asterisk the source deliberately leaves for the next slide. |

And the redundant kind: a four-box pipeline labelled `GET "user:42"` / `CRC16(key) mod 16384` /
`slot 9158` / `shard 2`, each with a caption beneath it, "the client sends a key", "slots are the
unit of ownership". The box already said it. The caption is the narration.

What to do instead: **speaker notes**, or nothing. If a slide genuinely needs the aside, that is a
question for the speaker, not a gap to fill. When trimming text that came from a source, keep the
source's wording verbatim rather than "improving" it: "\* Latency capped at 5ms" is the source;
"\* Latency capped at 5 ms by the benchmark's timeout" invents a cause.

**Audit for it, and fail the build.** Removing them once is not enough, because a later pass adds
them back for the same reason the first one did. `strip_asides()` in `build_nerdearla_deck.py`
deletes a named list and raises if an entry stops matching; `audit_no_stray_prose()` flags small
free-floating text ending in a full stop that is not a citation or a footnote marker.

### Visual Rules

Not story tenets, but they decide whether a slide looks hand-made or generated.

- **Bullets are short and punchy.** Lead with the number. One line each; four to six per slide.
- **Do not decorate labels.** A short piece of text does not need a coloured box behind it. Use the
  real title placeholder as the header rather than inventing a filled chip above the content.
- **No accent stripe on top of a filled card.** A card's fill already defines it; a thin bar across
  its top edge is pure decoration, and it is the single most recognisable tell of a machine-made
  slide. Same for node cards in diagrams. Leave the card plain.

  Accents that are *not* a line on top of a box are fine and worth keeping: a coloured left edge on a
  list row, a full-width rule marking a banner boundary, a short rule separating a title from a
  subtitle, hairline dividers between rows. The distinction is the horizontal stripe sitting on a
  box's top edge.

  To audit: for each thin shape, check whether a larger shape shares its exact left, top and width.
  If so it is a line on top of a box. Delete those; leave the rest.

### Bullet Style

Valkey bullets are a ► caret in Valkey blue. python-pptx cannot set bullet characters through its
API, so set the paragraph properties XML directly.

```python
from pptx.oxml import parse_xml

BULLET_PPR = '''<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    marL="457200" marR="0" lvl="0" indent="-350520" algn="l" rtl="0">
  <a:lnSpc><a:spcPct val="100000"/></a:lnSpc>
  <a:spcBef><a:spcPts val="{spc}"/></a:spcBef>
  <a:spcAft><a:spcPts val="0"/></a:spcAft>
  <a:buClr><a:schemeClr val="accent1"/></a:buClr>
  <a:buSzPts val="1920"/>
  <a:buFont typeface="Arial"/>
  <a:buChar char="&#9658;"/>
</a:pPr>'''

# spc="0" on the first paragraph, spc="1000" (10pt) on the rest for breathing room.
para._p.insert(0, parse_xml(BULLET_PPR.format(spc=0)))
```

The template's own `marL=457200 indent=-228600` is the inherited default; the caret style uses
`indent=-350520` to fit the wider glyph. Verified against community deck slides 7-8, which are the
reference for this style.

**Know when not to use it.** The caret style is raw text on white. In a deck built from cards and
filled shapes it is the only slide type with no fill and no structure, and it reads as harsh and
out of place. In that context use a row treatment instead: one `LIGHT_2` row per item with a
`73152` EMU `ACCENT_1` left edge, text at 18pt, vertically centred. Same information, same
highlight rule, and it borrows the deck's existing vocabulary. See `slide_bullets_cards()` in
`build_placeholder_deck.py`. Reach for carets when the surrounding slides are also text-led, as in
the community deck.

Those community deck runs pin `latin typeface="Arial"`, overriding the theme font. That is now the
house convention rather than a quirk to avoid: see Fonts.

### Highlighting Numbers

Measurable improvements get `ACCENT_1` + bold; the rest of the line stays `DARK_1` regular. This
is the single most recognizable thing about the deck style.

```python
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.util import Pt

r1 = para.add_run()
r1.text = "20%+ "
r1.font.bold = True
r1.font.size = Pt(24)
r1.font.color.theme_color = MSO_THEME_COLOR.ACCENT_1

r2 = para.add_run()
r2.text = "faster on the same hardware"
r2.font.bold = False
r2.font.size = Pt(24)
r2.font.color.theme_color = MSO_THEME_COLOR.DARK_1
```

Rules:
- Highlight percentages, throughput, latency, counts. **Do not** highlight ordinals ("1st",
  "first") or version numbers.
- Append `+` when the improvement is "at least" that much: `20%+`, `100M+`, `1B+`.
- One highlighted number per bullet. Two competes with itself.

### Source Citations

Put the source in its own text box pinned to the logo baseline. **Never** put it in the body
placeholder; it floats up under the last bullet and looks broken.

```python
from pptx.util import Emu, Pt

box = slide.shapes.add_textbox(Emu(609600), Emu(5943597), Emu(9000000), Emu(300000))
run = box.text_frame.paragraphs[0].add_run()
run.text = "Source: valkey.io/blog/..."
run.font.size = Pt(10)
run.font.italic = True
run.font.color.theme_color = MSO_THEME_COLOR.DARK_1
```

## Custom Slides

For anything no layout covers, start from `Title Only` (or `Holding` / `DEFAULT` for no title at
all) and build with native shapes. Native shapes stay editable and scale cleanly; never render a
diagram to PNG.

**Before you build custom, ask whether it should be a layout instead.** If the design will be used
more than once, hand-drawing it onto each slide is the wrong answer: add it to the master. Custom
slides are for content that only exists once, which in practice means diagrams.

### Shape Rules

- `MSO_SHAPE.RECTANGLE` for cards and callouts. No rounded corners.
  (Older diagram slides use `ROUNDED_RECTANGLE`. Match the slide you are editing, but square is
  the house style for new work.)
- Text-only shapes: `fill.background()`, `line.fill.background()`, all four margins zero,
  `vertical_anchor = MSO_ANCHOR.MIDDLE`.
- Card backgrounds: `fill.solid()` with a theme color, `line.fill.background()` for no border, or
  a hairline with `line.width = Pt(0.75)`.
- Accent bars: `RECTANGLE`, height `73152` EMU, solid accent fill, no line.
- Card padding: `182880` EMU (0.2in) inside the card edge.
- Hand-built titles: 28pt bold at `(609600, 400019)`, size `11062800 x 443198`.
- Remove placeholders you are not using, or PowerPoint shows empty "Click to add title" fields:

```python
spTree = slide.shapes._spTree
for shape in list(slide.shapes):
    try:
        if shape.placeholder_format.idx == 0:
            spTree.remove(shape._element)
    except (ValueError, AttributeError):
        pass
```

### Two-Column Cards

```python
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.util import Emu

gap  = 200000
col_w = (CONTENT_W - gap) // 2
col_h = 2800000                      # adjust to content
lx, rx = CONTENT_LEFT, CONTENT_LEFT + col_w + gap

for x, fill in ((lx, MSO_THEME_COLOR.ACCENT_5), (rx, MSO_THEME_COLOR.ACCENT_6)):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, CONTENT_TOP, col_w, col_h)
    bg.fill.solid()
    bg.fill.fore_color.theme_color = fill
    bg.line.fill.background()

    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, CONTENT_TOP, col_w, 73152)
    bar.fill.solid()
    bar.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
    bar.line.fill.background()
```

Three columns: `card_w = (CONTENT_W - 2 * gap) // 3`, same pattern.

Label cards with what they *are*. "State A" / "State B" tells the audience nothing; "Steady state"
/ "Rebalancing" does.

### Diagrams

- Use `add_shape()` + `add_connector()`. Connectors with `size=(0, h)` are vertical lifelines;
  `size=(w, 0)` are horizontal handoffs.
- For a dashed grouping box around related shapes, put the label in a **separate text box above**
  the children, not in the box itself, or it renders behind them.
- Overlay-based animation (stacking `req` and `retry` boxes at identical coordinates on one slide)
  is fragile and confusing to edit. Prefer one slide per state.

## Charts

Charts carry their own embedded workbook. Edit them through the chart API, not by touching text
runs. `replace_data` is verified working for both category and XY charts.

```python
from pptx.chart.data import CategoryChartData, XyChartData

chart = next(s.chart for s in slide.shapes if getattr(s, "has_chart", False) and s.has_chart)

# Category charts (LINE, LINE_MARKERS, BAR, COLUMN, ...)
cd = CategoryChartData()
cd.categories = ["10K", "20K", "30K"]
cd.add_series("p50", (167.0, 203.0, 228.0))
chart.replace_data(cd)

# XY scatter charts
xd = XyChartData()
s = xd.add_series("P50 Latency")
for x, y in [(0, 25.0), (1, 28.0), (2, 25.0)]:
    s.add_data_point(x, y)
chart.replace_data(xd)
```

**`replace_data` replaces the whole series set.** If you pass one series to a three-series chart,
you now have a one-series chart, and the deleted series' formatting is gone. Always re-add every
series you want to keep. Dump the chart first to get the existing series names and values.

Category labels can be empty strings when the shape of the curve is the point and the absolute
values are not. That is a legitimate choice, but label the axes.

**Blank categories do not survive a naive `replace_data` round trip, and the failure is disguised.**
`list(chart.plots[0].categories)` on a blank category returns something that prints as `'None'`,
because python-pptx's `Category` subclasses `str` and wraps a missing value as `str(None)`. Two
different things go wrong from there:

- Feeding that straight back into `add_series` writes the literal four-character string `None`
  into the embedded workbook, and PowerPoint then draws "None" under every point.
- Chasing it the other way wastes time: a deck that is *already correct* also reads back as
  `'None'`, because `<c:v/>` is a legitimately empty cell. Grep the unzipped
  `ppt/charts/chart*.xml` for the literal text before concluding the deck is broken.

Normalise on the way out and the whole class of problem disappears:

```python
cats = ["" if c is None else c for c in chart.plots[0].categories]
```

## Images and SVGs

Raster (PNG/JPG) for screenshots and product shots; `Title + Cont + Right Img Small` and
`Valkey List Rows with Image` both have a picture placeholder for exactly this. Native shapes for
everything diagram-shaped.

### Picture Placeholders

`add_placeholder(layout, "pic", idx, ...)` works, and a `pic` placeholder **is** cloned onto new
slides: `SlideLayout.iter_cloneable_placeholders()` only skips date, footer and slide number. Fill it
on a slide with `slide.placeholders[idx].insert_picture(path)`.

Two things to know:

- **`insert_picture()` crops to fill.** It does not letterbox. A square placeholder holding a square
  image crops nothing, which is why `Valkey List Rows with Image` is square. For a photo, whose
  aspect ratio has to survive, use `add_picture` on `Valkey Title Only` and derive one dimension from
  the other, as `slide_bio()` does.
- **A pic placeholder counts as a placeholder,** so any audit that selects rows as "every placeholder
  but the title" picks it up. Select by idx range instead. It also has no `text_frame` once the
  picture is in, so an emptiness check over it raises `AttributeError`.

### Brand Logos

Eight official logo SVGs live in `generation/branding/logo svgs/`:
`valkey-{horizontal,stacked}-{black,white,color,color-light}.svg`.

- **horizontal** aspect ratio `187.9 / 63.5` = **2.959**
- **stacked** aspect ratio `778.6 / 669.9` = **1.162**

Always derive width from height using those ratios; never type two independent numbers. Use the
`white` variants on `ACCENT_1` / `ACCENT_2` fills and `color` on white. The `logo()` helper in
`build_placeholder_deck.py` takes a height and does the arithmetic.

### White-Background Title Slides

The slide master's background is already white (`lt1`). The stock title-ish layouts
(`Title + 2x Speaker`, `Divider`, `Thank You 2x`) look dark because the *layout* carries a
full-bleed background **picture**, which is why their placeholder text is `LIGHT_1`. You cannot
make them white without covering that image up.

For a white title slide, build custom on **`Holding`**: zero placeholders, no background picture,
so the master's white shows through and you have a clean canvas. `Title Only` also works if you
want a real title placeholder.

**Every layout inherits the master's shapes** (`showMasterSp` is unset everywhere), so the Valkey
logo at `(10389732, 5943597)` renders bottom-right on every slide. Two consequences:

- A full-bleed fill **covers** it, because slide shapes draw over master shapes. Put a white
  horizontal logo back at the same spot: height `443198`, right edge `11824163`.
- A bottom banner taller than `914403` also covers it, which is usually what you want on a title or
  closing slide. If you add your own logo elsewhere on a slide that still shows the corner mark,
  you get two logos. Check before shipping.

### Banner Artwork

`~/git/valkey-svgs` holds 39 generated banners (`svg/<theme>.svg`, previews in
`images/` and chrome-free copies in `images/plain/`). `generation/make_band_art.py` turns one
into the title slide's right-hand band.

**They are not flourishes, and that is the thing to know before reaching for one.** Each is
a whole 1920x1080 picture: navy-to-purple gradient, vignette, film grain, motif, *its own
copy of the Valkey mark* at the motif's focal point, plus a corner lockup and a caption.

- **The motif cannot be lifted off its background.** Stripping the gradient to put the
  linework on the flat `ACCENT_1` band does not work: the motifs carry blurred dark glow
  layers drawn to bloom against navy, and on `#6983FF` they read as smudges. Tried, and it
  looked like a dirty band. So the gradient comes along and the band stops being Valkey
  blue, which is a design change to the slide rather than a decoration.
- **Strip the banner's own mark** or it collides with the layout's lockup. Match it by the
  path data in `valkey-svgs/assets/Valkey-logo.svg`; there are two groups, the mark and a
  blurred glow copy at the same transform.
- **The layers are separable by position,** because the generator emits them in a fixed
  order: `defs`, sky rect, motif groups, vignette rect, grain rect, then the chrome. Cut
  every sibling after the grain rect and you have the art without the lockup or caption.
- **Rasterise; do not embed the SVG.** These lean on `feGaussianBlur` and `feTurbulence`,
  and PowerPoint renders SVG filters inconsistently or not at all. Nothing in them is
  linework that needs to scale.
- **Crop to the target aspect centred on where the mark was,** so the motif's focal point
  lands on the lockup. The band is 0.6:1 portrait against the banner's 16:9, so this throws
  away two thirds of the width: only radial motifs survive it. `performance`, `release`,
  `clustering`, `blackhole-halo`, `community` and `planet-ring` all do.

### Icon Library

`generation/icons/` holds 18 reusable outline icons, all `96x96` viewBox, `#6983FF` stroke at width 4,
no fills: `community`, `cost-efficiency`, `data-hash`, `data-list`, `durability`, `general-scaling`,
`global-community`, `hardening`, `horizontal-scaling`, `memory-efficiency`, `migration`,
`observability`, `operator`, `performance`, `reliability`, `replication`, `security`,
`vertical-scaling`.

`community` is three linked figures and reads as "a team"; `global-community` is a globe with
scattered points and reads as "spread across the world". Pick on which claim the card is making.

**A dot has to be drawn as a stroke, not a fill.** The validator requires `fill="none"` on every
element, so a point is a tiny circle whose stroke width exceeds its radius: `r="1.8"` with
`stroke-width="7"` renders solid. Keep the points clear of the other strokes by at least the sum of
their half-widths, or they read as attached to the line they touch rather than as separate points.

Keep a validator for these and run it on every addition: single root `<svg>` at `96x96`, every element
`fill="none"` with a `#6983FF` or `#232F3E` stroke, 3-8 elements, geometry inside the viewBox, and none
of `<style>`, `class=`, `style=`, gradients, filters, masks, `clipPath`, `<text>` or `<use>`. **Exclude
the derived `valkey-mark-*.svg`** — those come from the brand pack and legitimately carry a gradient
and a `<style>` block.

A grid of icons should adapt rather than assume: `rows = ceil(n / cols)`, and shrink the icon to
`row_pitch - label_gap - label_h - padding` so an extra row never overlaps the one above.

Sizing that works: **1097280 EMU (1.2in)** for an icon above an 18pt card heading, **914400 EMU
(1.0in)** in a 4-column gallery grid. 0.7in looks like a bullet point next to real text; do not go
smaller than 1in on a card.

To add more, keep the constraints: single root `<svg>` with `width="96" height="96"
viewBox="0 0 96 96"`, presentation attributes only, and no `<style>`, CSS classes, `style=`,
gradients, filters, masks, `clipPath`, `<text>`, or `<use>`. PowerPoint is strict, and those
features render inconsistently or not at all.

### Embedding SVGs

Use `generation/pptx_svg.py`:

```python
from pptx_svg import add_svg
add_svg(slide, "generation/icons/security.svg", "generation/icons/png/security.png",
        left, top, 1097280, 1097280, name="security")
```

python-pptx cannot insert SVGs (PIL cannot decode them), so that helper builds the package parts by
hand. PowerPoint renders the `asvg:svgBlip` extension and falls back to the embedded PNG in older
viewers. Three traps, all of which produce silently wrong output:

1. **`Part`'s signature is `Part(partname, content_type, package, blob)`.** Wrong argument order
   gives a 0-byte part and an "image cannot be displayed" box.
2. **A PNG fallback is required, and it must be correct.** Rasterize with headless Chrome, but
   **always go through an HTML wrapper with an explicitly sized `<img>`.** Pointing Chrome straight
   at a `.svg` fails in two different ways:

   - Illustrator/Inkscape exports (all the brand logos) render as a **fully transparent image**.
     Not a timing issue; `--virtual-time-budget` does not help.
   - An SVG that declares `width`/`height` renders at that natural size inside the window rather
     than scaling to fit, so an oversized window silently corner-anchors the artwork with
     transparent padding around it.

   The wrapper sidesteps both, because CSS sizing always wins:

   ```html
   <html><body style="margin:0">
     <img src="file:///abs/path/icon.svg" style="width:384px;height:384px;display:block">
   </body></html>
   ```

   ```bash
   chrome --headless --disable-gpu --default-background-color=00000000 \
     --window-size=384,384 --screenshot=out.png file:///tmp/wrapper.html
   ```

   Then assert it: the rendered content's bounding box should fill most of the canvas. Under ~60%
   means something went wrong. `rasterize()` in `build_placeholder_deck.py` does all of this.
3. **`shape.image` returns the PNG fallback, not the SVG.** It resolves the `a:blip` `r:embed`. To
   read the actual SVG, follow the `svgBlip` `r:embed` instead.

## Running Header

A running header belongs on **every** slide, including the ones that already have a title. Two
mistakes to avoid, both of which look reasonable in code:

- **Too small.** A 10pt line reads as a footnote. 16pt is visible without competing with a title.
- **Only on untitled slides.** Skipping titled slides means the header lands *only* on the title and
  closing slides, where it duplicates the big title. That is the opposite of a running header.

What works:

- Right-aligned in a band **above** the title: `(609600, 130000)`, size `11062800 x 260000`,
  vertically centred, 16pt regular.
- Right-aligned and above, so it coexists with any slide's own title however long that title runs.
  Do not right-align it *within* the title band: a 44-character title at 32pt is already 11.4in
  wide, and the header needs 1.9in of the 13.3in slide.
- `DARK_2` normally, `LIGHT_1` where the top of the slide is a dark fill.
- Its own text box per slide, so it can be deleted freely.

Add it in a pass over the finished deck, not inside each generator, so a newly added slide cannot
forget it, and run it last so it draws over any background fill. Detect the dark background rather
than hardcoding slide numbers: walk the autoshapes for a solid dark-theme fill covering the header
band.

## The Glyph, and Where to Use It

**The brand pack ships only lockups** (glyph plus the "Valkey" wordmark), and at corner-mark size the
wordmark is illegible. Derive a glyph-only asset instead of shipping an illegible lockup: parse
`valkey-stacked-<colour>.svg`, remove the `<g id="text...">` wordmark group, and tighten the viewBox
to the glyph. The glyph's box is `248.0 0 282.6 326.0`, aspect **0.867** — measured by rasterising
the wordmark-free SVG and taking its bounding box, which is the reliable way to find it.

Where the glyph belongs: the **upper-right corner of boxes that are a real instance or shard**,
`220000` EMU tall, inset `110000` — exactly where those boxes used to carry a blue top bar. It fills
the space the stripe vacated, stops a grid of grey rectangles looking inert, and unlike the stripe it
carries information: a box with the glyph is a server, a box without one is a client, a proxy, or a
content card.

Where it does not belong: **loose in the corner of the slide.** The master already puts a full lockup
bottom-right, so a second mark floating at the top is just another piece of furniture. Brand marks
should be attached to something they identify.

**Cache the image parts.** The same logo placed on 21 boxes would otherwise add 42 media parts to
the package. Key a cache on `(id(package), svg_path)` and reuse the `Part` objects;
`slide_part.relate_to()` returns the existing rId when a slide already references one. In the
placeholder deck this keeps 62 pictures down to 35 media parts.

## Sections

PowerPoint sections (the slide-sorter groups) live in `presentation.xml` as a `p14:sectionLst`
extension listing slide ids. python-pptx neither reads nor updates them, so a deck built by deleting
a donor's slides keeps the donor's sections, pointing at slide ids that no longer exist.
`strip_slides()` in `build_placeholder_deck.py` clears them. `set_sections(prs, [(name, start)])`
writes new ones from 0-based start indexes; call it after the slide order is final, because it
records ids, not positions. `audit_sections()` fails the build when sections and slides disagree.
Section ids are derived from the name, so rebuilds stay byte-stable.

## Speaker Notes

```python
slide.notes_slide.notes_text_frame.text = "..."
```

Notes are a common source of embarrassment: duplicating a slide duplicates its notes, so notes
routinely end up describing a completely different slide. When you copy a slide, either rewrite
the notes or clear them. Dump notes before presenting.

## Structural Edits

Order of operations, always: **delete (highest index first) → add → reorder once at the end.**

```python
# Add, then move into position (0-indexed).
slide = prs.slides.add_slide(layout)          # appends
lst = prs.slides._sldIdLst
lst.insert(target_index, lst[-1])             # insert moves the element
```

Deleting a slide is safe **as long as you drop the relationship and remove the `sldIdLst` entry
together.** python-pptx only serializes parts reachable through the relationship graph, so the
orphaned slide part is dropped on save. Verified: stripping all 31 slides from a deck left zero
`ppt/slides/*` parts and preserved all 17 layouts, the master, and the theme.

```python
rId = prs.slides._sldIdLst[index].rId
prs.part.drop_rel(rId)          # both lines, or you leave a dangling r:id
del prs.slides._sldIdLst[index]
```

Removing the `sldIdLst` entry *without* `drop_rel` is what produces the "needs repair" prompt.
Always reopen the saved file afterwards to confirm.

### Duplicating a Slide

There is no copy-slide API either. Add a slide on the source's layout, empty it, and deep-copy the
source's shape elements in. `see clone_slide()` in `build_nerdearla_deck.py`:

```python
dst = prs.slides.add_slide(src.slide_layout)
for shape in list(dst.shapes):
    dst.shapes._spTree.remove(shape._element)
for shape in src.shapes:
    dst.shapes._spTree.append(copy.deepcopy(shape._element))
```

**This only works for shapes that carry no relationships:** autoshapes, text boxes, connectors.
A picture or a chart references a part through `r:embed`, and a deep copy of the element points at
a relationship the new slide does not have, which is an "image cannot be displayed" box or a
repair prompt. Refuse rather than produce that:

```python
if "r:embed" in shape._element.xml or "r:link" in shape._element.xml:
    raise ValueError(f"{shape.name!r} carries a relationship; cannot clone")
```

Notes are not copied either. Copy them deliberately or clear them: see Speaker Notes.

### Reordering

Compute the whole target order once and rewrite `sldIdLst` in a single pass, rather than moving
slides one at a time and reasoning about how each move shifts the rest. Assert it is a permutation
first; an index used twice silently drops a slide.

```python
def reorder(prs, order):                 # order holds current 0-based indices
    lst = prs.slides._sldIdLst
    assert sorted(order) == list(range(len(lst)))
    ids = list(lst)
    for el in ids:
        lst.remove(el)
    for i in order:
        lst.append(ids[i])
```

### Starting a New Deck From the Template

There is no blank Valkey template file, so clone an existing deck and empty it. Use the deck with
the richest layout set as the donor:

```python
import shutil
shutil.copyfile("generation/donors/The Vertical Scaling wall.pptx", "slides/New deck.pptx")
prs = Presentation("slides/New deck.pptx")
for i in range(len(prs.slides) - 1, -1, -1):     # highest index first
    prs.part.drop_rel(prs.slides._sldIdLst[i].rId)
    del prs.slides._sldIdLst[i]
```

You keep the theme, master, all layouts, and the logo. A handful of media parts referenced by the
master and layouts stay behind; that is correct, not leakage. See
`generation/build_placeholder_deck.py` for the full pattern.

If a deck lacks a layout you need, add it with `generation/pptx_layout.py` rather than hand-drawing the
design onto every slide. See Adding a Layout to the Master.

## Verifying Your Work

There is no LibreOffice or PowerPoint CLI here, but you can still **see your slides**: read the real
geometry, theme fills, text runs and embedded images out of the file, lay them out as HTML, and
screenshot that with headless Chrome. An approximate renderer built that way reliably catches the
things that actually go wrong: text overflowing a card, shapes colliding, icons at the wrong scale,
dead space, bad alignment.

The rest of this section is what such a renderer can and cannot tell you, learned by building one.
It is worth the hour, and the gaps below are worth knowing before you trust any output.

**Draw master and layout content beneath each slide's own shapes** (the corner Valkey logo, the dark
full-bleed backgrounds on the title layouts), or what you see omits inherited artwork and
misrepresents the slide badly.

It is an approximation, not a renderer. Expect these gaps: fonts fall back to a generic
sans so line breaks differ slightly; autofit, **bullet glyphs**, **paragraph `spcBef` spacing**,
rotation, gradients, and shadows are ignored; charts draw as labelled grey boxes.

**No text from a layout or master appears at all.** For inherited shapes the renderer emits only
pictures and fills, never text, so a layout that puts *text* in its static decoration renders blank.
`Valkey Three Features` bakes its "1", "2", "3" numerals in that way, and they are invisible in the
mock while being perfectly present in the file. Check the layout part with python-pptx instead of
believing the render. **Never claim a slide is visually correct on the strength
of an approximate render alone.** It proves layout, not typography. Say which you verified, and check the
things it cannot show (carets, spacing) in the XML instead.

**Text inheritance from a layout is the hardest gap to close,** and it limits what any quick renderer
tells you about a slide built from one. Position and size of the box come through, because those are on the slide's own
placeholder, but nothing from the layout's `lstStyle` does: a 130pt big number and a 13pt card body
render at the same default, and `algn="ctr"` renders left. Renderer gaps rather than deck bugs, and they
cannot be settled that way. To check type scale, add a slide from the layout, fill the placeholders,
and **open it in PowerPoint once.**

**Do resolve a placeholder's *fill* from its layout placeholder, and do draw linear `a:gradFill`.**
Skipping either hid a real bug once: a row placeholder whose card background lives on the layout
rendered as bare text, so a broken fill shipped unseen. If you use a fill mechanism your renderer
does not know about, teach it rather than working around it.

What you *can* establish without opening it: that the mechanism is the one the template already
relies on everywhere. Both size and alignment reach slides through layout `lstStyle` in the stock
layouts — that is where the whole type scale comes from, and `Video` sets `algn="ctr"` there. If
inheritance were broken, every stock layout would be broken too.

A caution learned the hard way, four times: a renderer's own bugs make *correct* decks look broken.
An alignment map keyed on XML values instead of python-pptx enum names; the same for theme colours;
PowerPoint's internal text margins applied as CSS padding without `box-sizing:border-box`, which drew
every filled shape up to 0.2in wider than it really was; and a hardcoded colour scheme that rendered
the patched `accent6` as the stock peach. Read the theme from `theme1.xml`, and draw unmapped theme
colours in magenta with a warning rather than quietly substituting black.

**The renderer cannot show you a run PowerPoint refuses to draw.** It reads the text out of the
file, so a schema violation that makes PowerPoint discard a run is invisible to it: the render shows
the text, PowerPoint shows a gap. `endParaRPr` ordering is the known case (pitfall 20) and it cost a
deck. When a user says text is missing and every tool here says it is present, believe the user and
go read the raw `a:p` child order.

**When the render looks wrong, confirm against the XML before "fixing" the deck.** Measuring the
rendered PNG (bounding boxes of a known colour) against the EMU geometry in the file is the fastest
way to tell a renderer bug from a real one.

Alongside it, always:

- Re-dump every slide you touched and read it back.
- Check bounds arithmetic: `left + width <= 12192000`, `top + height <= 6858000`.
- Count runs before and after. A paragraph that had two runs and now has one lost its formatting.
- Reopen the saved file with python-pptx. If it raises, PowerPoint will complain too.

For a pixel-accurate check, LibreOffice (`brew install --cask libreoffice`) would allow
`soffice --headless --convert-to png`. Worth suggesting when typography matters.

## Pitfalls

1. **Formatting is per-run, not per-paragraph.** Never assign a whole line to `runs[0].text`. A
   bullet with a highlighted number is two runs; collapsing them kills the color and bold. Dump
   the run structure first, then edit runs in place or rebuild all of them.

   The mirror of this, and a trap worth naming: **multiple runs in one paragraph are often the
   point, not a mistake.** The Vertical Scaling wall's two-box slides read as though two lines had
   been jammed together — "Limited with IO threading Provides about 8x throughput scaling" — and
   that is a deliberate build step: the first run carries `strike="sngStrike"` and the second is its
   replacement. Splitting them into paragraphs destroys the build and looks like a tidy-up.
   **Dump `rPr` before you decide a run structure is wrong.** `strike`, `u`, `baseline` and
   `highlight` are all invisible in a text dump and all load-bearing.
2. **New placeholders have zero runs.** Use `para.add_run()`; do not assume `para.runs[0]` exists.
   The same helper therefore needs two paths: clone-and-retext for a donor shape that has
   formatting worth keeping, and plain `tf.text = ...` plus a font pin for a fresh one.
3. **Layout indices are per-deck.** Resolve by name, every time.
4. **`replace_data` drops series you omit.** Re-supply the full set.
5. **Delete from the highest index down,** and always pair `drop_rel` with the `sldIdLst` deletion.
6. **Group shapes hide their children.** `shape.shape_type == 6` means recurse into
   `shape.shapes`; top-level iteration will not find the text inside.
7. **Theme colors, not RGB.** `MSO_THEME_COLOR.ACCENT_1`, not `RGBColor(0x69, 0x83, 0xFF)`.
8. **Duplicated slides carry stale notes and duplicate shape names.** Both bite later.
9. **Check for `~$` lock files before saving.**
10. **Off-template layouts (`DEFAULT`, `Blank`) lose the footer and logo.** If a slide looks
    naked compared to its neighbors, check what layout it is on.
11. **Never derive a shape id from anything but the shape tree.** `add_svg` used to set
    `id = media_number + 1000`, which collides with the ids python-pptx hands out as soon as
    pictures and shapes interleave: it left duplicate ids on 12 of the 31 slides in the placeholder
    deck. PowerPoint tolerated it silently, which is why it survived. Allocate from
    `max(existing) + 1` over the tree you are appending to, and assert uniqueness after building.
12. **Gradient stop positions must be strictly increasing.** `a:gsLst` positions are required to be
    distinct; duplicates are undefined and PowerPoint resolves them by dropping a stop, turning a
    hard edge into a full-width wash. Nothing in this repo generates a gradient any more, for the
    reason above, but `pptx_layout.audit()` still checks it: a deck edited in PowerPoint can bring
    one in.
13. **The mock renderer is not a source of truth about the theme.** It used to hardcode the colour
    scheme, so a deck whose `accent6` had been patched to the brand rose rendered in the old stock
    peach. It reads `theme1.xml` now. When a colour looks wrong in the mock, check the theme blob
    before changing the deck.
14. **`shape.image` raises `ValueError("no embedded image")` on an SVG picture.** Anything that
    walks every shape in a deck built by this repo has to catch it, because `add_svg` pictures
    resolve their raster through `svgBlip`. `dump_deck.py` used to die on slide 1 of every deck.
15. **`a:ln`'s `w` attribute is optional.** Absent means the DrawingML default of 0, a hairline.
    Reading it as `int(ln.get("w"))` crashes on any shape styled through python-pptx's colour API
    without an explicit width, which is most of them.
16. **`slide.shapes.title` returns a fresh wrapper object on every call.** `sh is not slide.shapes.title`
    is therefore always true, so an audit written that way reports the title as colliding with
    itself on every slide. Compare `sh._element is title._element`.
17. **Strikethrough, and fill alpha, have no python-pptx API.** Both are raw XML. Strikethrough is
    an attribute on the run: `run._r.get_or_add_rPr().set("strike", "sngStrike")`. Alpha is a child
    of the *colour* element and only attaches to `<a:srgbClr>`, so a theme colour cannot carry it:
    setting `theme_color` writes `<a:schemeClr>`, and appending `<a:alpha>` there is ignored. Use
    explicit RGB when you need translucency, and note the mock renderer draws it fully opaque.
18. **`<a:spAutoFit/>` on a donor text box makes your explicit height a lie.** The box was sized to
    its old content; re-texting it with something longer clips instead of wrapping. Replace the
    autofit element with `<a:noAutofit/>` rather than deleting it, because absent means "inherit".
19. **A pasted-in slide can carry a second, invisible-looking title.** Seven slides in The Vertical
    Scaling wall hold a leftover `Google Shape;204;p8` text box the same size as the title
    placeholder, offset 152400 EMU, so two different strings draw on top of each other. It is
    invisible in the PowerPoint outline view and obvious in a render. Audit for it: any text shape
    other than the title whose vertical span intersects the title band.
20. **`a:endParaRPr` must be the last child of `a:p`, and PowerPoint is the only thing that
    tells you.** The schema sequence is `pPr?, EG_TextRun*, endParaRPr?`. Appending a run to a
    paragraph that already carries an `endParaRPr` puts the run *after* it, and PowerPoint
    enforces the order by **silently dropping that run**. Nothing else notices: python-pptx reads
    the text back correctly, `dump_deck.py` prints it, and the mock renderer draws it. The slide
    is correct in every tool you have and blank in the room.

    This bit the Nerdearla deck on 9 paragraphs across 7 slides at once — two slide titles
    vanished, the conference line on the title slide vanished, and both strikethrough two-box
    builds lost their cards. It is a one-line fix in any helper that builds runs by hand:

    ```python
    end = para._p.find(qn("a:endParaRPr"))
    if end is None:
        para._p.append(new_r)
    else:
        end.addprevious(new_r)          # never append past endParaRPr
    ```

    `para.add_run()` gets this right on its own; the trap is only in code that deep-copies a run
    element to preserve formatting and then appends it. Audit for it rather than trusting it,
    because no render will show you: for every `a:p`, assert the index of `endParaRPr` is greater
    than the index of the last `r` / `br` / `fld`. `audit_paragraph_order()` in
    `build_nerdearla_deck.py` does exactly that and fails the build.

21. **Never set only `width` or `height` on an inherited placeholder.** A slide placeholder that
    inherits its geometry has no `a:xfrm`. Setting `ph.width` makes python-pptx create one holding
    only `<a:ext cx="..." cy="0"/>`, with no `a:off`. PowerPoint then puts the box at the slide
    origin with zero height. The mock renderer and python-pptx still report the layout's position,
    so the slide looks fine everywhere except in PowerPoint. This broke the CVE Tsunami title slide.
    Either leave the layout geometry alone and make the text fit, or set all four of `left`, `top`,
    `width`, `height`. `build_cve_tsunami_deck.py`'s `audit()` fails the build on a placeholder
    with a partial `xfrm`.

## Deck-Specific Notes

Per-deck structure, update cadence, and content sources live alongside each deck:

- `descriptions/valkey-placeholder-slides.md` — the placeholder/starter slide library
- `generation/donors/vertical-scaling-wall-deck.md` — The Vertical Scaling wall
- `descriptions/scaling-up-and-out-deck.md` — Nerdearla 2026, generated from the above
- `generation/donors/valkey-community-deck.md` — the reusable community deck

When you learn something new about a deck, write it there. When you learn something about the
template or the tooling, write it here.
