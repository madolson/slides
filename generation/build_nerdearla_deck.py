#!/usr/bin/env python3
"""Build the Nerdearla Argentina 2026 deck from The Vertical Scaling wall.

    python3 generation/build_nerdearla_deck.py

Re-runnable: always starts from a fresh copy of the donor, so editing this script
and re-running is the only supported way to change the output. Never hand-edit the
generated pptx and then re-run; your changes are gone.

The deck is The Vertical Scaling wall plus the edits made in the P99 CONF 2026
deck, retitled for Nerdearla's "Scaling Up and Out: How Valkey Does Both". The
P99 deck is on a different template at a 1.5x canvas (20 x 11.25in against the
Valkey 13.333 x 7.5in), so nothing is copied from it: the new slides are rebuilt
at Valkey-template geometry and the changed text and chart data are transcribed.

Requires python-pptx, Pillow, and Google Chrome (for SVG icon PNG fallbacks).
"""
import copy
import json
import os
import re
import shutil
import sys

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_placeholder_deck import (BAR_H, CARD_H, CONTENT_LEFT, CONTENT_TOP,
                                    CONTENT_W, FONT, ICONS, PAD, SLIDE_H,
                                    SLIDE_W, TITLE_FONT, TITLE_H, TITLE_TOP,
                                    LAYOUT_LIST_ROWS_IMAGE,
                                    LIST_ROWS, LIST_ROW_GAP, LIST_ROW_H, arrow,
                                    box_columns, build_layouts, card, chip,
                                    list_row_y, patch_theme, rect, source_note,
                                    strip_stock_layouts, text_box)
from build_placeholder_deck import PNGS as ICON_PNGS
from pptx_svg import add_svg

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)                      # repo root
SLIDES = os.path.join(ROOT, "slides")             # generated decks live here
ASSETS = os.path.join(HERE, "assets")
DONOR = os.path.join(HERE, "donors", "The Vertical Scaling wall.pptx")
OUT = os.path.join(SLIDES, "Scaling Up and Out - How Valkey Does Both.pptx")

TALK_TITLE_L1 = "Scaling up and out,"
TALK_TITLE_L2 = "and how Valkey does both"
SPEAKER = "Madelyn Olson"
ROLE = "Valkey Engineer"

# This deck is entirely on the Valkey master that build_placeholder_deck.py defines.
# build_layouts() adds all eleven layouts, retarget_layouts() moves the carried-over
# donor slides onto `Valkey Title Only`, and the stock layouts and the unused second
# master are then deleted. No slide is left on a donor layout.
LAYOUT_TITLE_ONLY = "Valkey Title Only"
LAYOUT_INSIGHT = "Valkey Insight"
LAYOUT_LIST_ROWS = "Valkey List Rows"
LAYOUT_THREE_BOX = "Valkey Three Box"

# Marks a slide whose content is deliberately unfinished. Grep the output for this
# before presenting.
TODO = "TODO"

# P99's wording for the capped-latency footnote, verbatim. Both latency-cliff charts
# carry it. Do not extend it: the reason for the cap is not in any source here.
CAP_NOTE = "* Latency capped at 5ms"

# The bio slide's speaker note: the talk's opening lines, spoken in Spanish before the
# switch to English.
SPANISH_OPENER = (
    "Buenas tardes. Me llamo Madelyn Olson. Soy de Estados Unidos.\n\n"
    "Perdón por mi español. No lo hablo desde la secundaria, y estoy muy fuera "
    "de práctica.\n\n"
    "Estoy muy emocionada de estar acá en Nerdearla. Me gusta mucho el trekking, "
    "la naturaleza y el buen vino. Argentina es perfecta para mí. Ya probé el "
    "Malbec. Es muy, muy bueno.\n\n"
    "El resto de la charla va a ser en inglés.")

# ACCENT_1 as explicit hex. A translucent fill needs an <a:alpha> child, and that
# has to hang off an <a:srgbClr>; python-pptx writes a <a:schemeClr> for a theme
# colour, which cannot carry it. Same value as the theme's accent1.
VALKEY_BLUE = RGBColor(0x69, 0x83, 0xFF)


# --- locating and editing what is already on a slide ----------------------------


def shapes_flat(slide):
    """Every shape on the slide, recursing into groups."""
    def walk(shapes):
        for sh in shapes:
            yield sh
            if sh.shape_type == 6:      # GROUP: top-level iteration misses the text
                yield from walk(sh.shapes)
    return list(walk(slide.shapes))


def by_text(slide, text, exact=True):
    """The one shape whose text matches. Raises unless exactly one matches.

    Shape *names* are duplicated all over these decks (three `Shape 3` on one
    slide), so text is the only stable handle.
    """
    hits = [sh for sh in shapes_flat(slide)
            if sh.has_text_frame
            and (sh.text_frame.text.strip() == text if exact
                 else text in sh.text_frame.text)]
    if len(hits) != 1:
        raise LookupError(f"{len(hits)} shapes match {text!r}; expected 1")
    return hits[0]


def set_runs(shape, *lines):
    """Rewrite a text frame to exactly `lines` paragraphs.

    Each line is a list of (text, bold) or (text, bold, strike), or None to leave
    that paragraph alone. Formatting is per-run, so this clones the first paragraph
    that has runs and re-texts the clone rather than assigning to `runs[0].text`
    blind: a paragraph with a highlighted number is two runs, and collapsing them
    kills the colour. Paragraphs are added or removed so the count matches.

    Multiple runs in one paragraph are not always an accident. The donor's two-box
    cards put a struck-through phrase and its replacement in one paragraph on
    purpose, which is the whole point of those slides.
    """
    tf = shape.text_frame
    seed = next((p for p in tf.paragraphs if p.runs), None)
    if seed is None:
        raise LookupError("no paragraph with runs to copy formatting from")
    para_tmpl = copy.deepcopy(seed._p)
    run_tmpl = copy.deepcopy(seed.runs[0]._r)

    body = tf._txBody
    existing = list(tf.paragraphs)
    for i in range(len(lines) - len(existing)):
        body.append(copy.deepcopy(para_tmpl))
    for para in existing[len(lines):]:
        body.remove(para._p)

    for para, runs in zip(tf.paragraphs, lines):
        if runs is None:
            continue
        for r in list(para.runs):
            para._p.remove(r._r)
        # `a:endParaRPr` must be the LAST child of `a:p`: the schema sequence is
        # pPr?, EG_TextRun*, endParaRPr?. Appending a run to a paragraph that already
        # carries one puts the run after it, and PowerPoint enforces the order by
        # silently dropping the run. python-pptx does not validate, so the text reads
        # back fine from the file and simply never appears on screen. Insert ahead of
        # it instead.
        end = para._p.find(qn("a:endParaRPr"))
        for spec in runs:
            text, bold = spec[0], spec[1]
            strike = spec[2] if len(spec) > 2 else None
            new_r = copy.deepcopy(run_tmpl)
            if end is None:
                para._p.append(new_r)
            else:
                end.addprevious(new_r)
            run = para.runs[-1]
            run.text = text
            if bold is not None:
                run.font.bold = bold
            # python-pptx has no strikethrough API; it is an rPr attribute.
            rPr = run._r.get_or_add_rPr()
            if strike:
                rPr.set("strike", "sngStrike")
            else:
                rPr.attrib.pop("strike", None)


def unset_autofit(shape, wrap=True):
    """Turn off text autofit so an explicit box size is honoured.

    A donor text box carrying `<a:spAutoFit/>` sizes itself to one line of its old
    content. Re-texting it with something longer then clips, because the stored
    height is what renders. Replace the autofit element with `noAutofit` rather than
    just deleting it: absent means "inherit", which can still be spAutoFit.
    """
    bodyPr = shape.text_frame._txBody.bodyPr
    for tag in ("a:spAutoFit", "a:normAutofit", "a:noAutofit"):
        for el in bodyPr.findall(qn(tag)):
            bodyPr.remove(el)
    bodyPr.append(bodyPr.makeelement(qn("a:noAutofit"), {}))
    bodyPr.set("wrap", "square" if wrap else "none")


def retext(shape, text, bold=None):
    """Single-run convenience wrapper for set_runs."""
    set_runs(shape, [(text, bold)])


def set_title(slide, text):
    """Set the title placeholder's text, on a donor slide or a fresh one.

    A fresh placeholder has zero runs, so there is no formatting to preserve and
    nothing to clone: set the text and pin the font. Leave the size unset so it
    inherits 32pt from the layout.
    """
    title = slide.shapes.title
    if title is None:
        raise LookupError("slide has no title placeholder")
    if title.text_frame.paragraphs[0].runs:
        retext(title, text)
        return
    title.text_frame.text = text
    for run in title.text_frame.paragraphs[0].runs:
        run.font.name = TITLE_FONT


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def clear_notes(slide):
    if slide.has_notes_slide:
        slide.notes_slide.notes_text_frame.text = ""


def chart_of(slide):
    return next(sh.chart for sh in slide.shapes
                if getattr(sh, "has_chart", False) and sh.has_chart)


def categories_of(chart):
    """The chart's category labels, with blanks preserved as blanks.

    These decks deliberately use blank categories where the shape of the curve is
    the point, and those have to survive a replace_data() round trip. Two traps:
    python-pptx's Category subclasses str and wraps a missing value as str(None),
    so a blank reads back as the four-character string "None"; and feeding a real
    None to add_series() writes "None" into the workbook for real, where it does
    render under every point. Normalising to "" avoids both.
    """
    return ["" if c is None else c for c in chart.plots[0].categories]


def rename_series(slide, old, new):
    """Rename one series, keeping every series and its values.

    replace_data() replaces the whole series set, so the full set has to be
    re-supplied or the omitted series (and their formatting) are dropped.
    """
    chart = chart_of(slide)
    plot = chart.plots[0]
    cats = categories_of(chart)
    series = [(s.name, list(s.values)) for s in plot.series]
    if old not in [n for n, _ in series]:
        raise LookupError(f"no series named {old!r}: have {[n for n, _ in series]}")
    cd = CategoryChartData()
    cd.categories = cats
    for name, values in series:
        cd.add_series(new if name == old else name, values)
    chart.replace_data(cd)


def replace_series(slide, spec):
    """Rewrite every series from `spec`, a list of (name, values)."""
    chart = chart_of(slide)
    cd = CategoryChartData()
    cd.categories = categories_of(chart)
    for name, values in spec:
        cd.add_series(name, values)
    chart.replace_data(cd)


def footnote(slide, text):
    """A small note under a chart, pinned to the logo baseline like a source line."""
    return text_box(slide, CONTENT_LEFT, 5943597, CONTENT_W, 300000,
                    [(text, 12, False, MSO_THEME_COLOR.DARK_2)])


# --- fonts ----------------------------------------------------------------------
#
# The donor's theme major font is Play and its bullets set buFont="Play". Play is a
# Google font: it is not on a conference laptop, and PowerPoint substitutes silently,
# so a title that looked right here reshapes in the room. Arial everywhere removes
# the risk and matches what every run in this deck already pins.


def retypeface(prs, old, new):
    """Rewrite every `typeface="old"` in the package to `new`.

    Covers what python-pptx will not reach: the theme's font scheme, the slide and
    notes masters, all layouts, `presentation.xml`'s default text style, and the
    `buFont` on every bulleted paragraph. python-pptx does not parse the theme part
    at all, so this edits part blobs and elements directly.
    """
    hits = 0
    for part in prs.part.package.iter_parts():
        name = str(part.partname)
        if not name.endswith(".xml") or "/ppt/" not in name:
            continue
        # A part python-pptx has parsed exposes an element, and its blob is
        # regenerated from that element on save: patching the blob would be thrown
        # away. Patch whichever representation is the live one.
        element = getattr(part, "_element", None)
        if element is not None:
            for el in element.iter():
                if el.get("typeface") == old:
                    el.set("typeface", new)
                    hits += 1
            continue
        xml = part.blob.decode("utf-8")
        patched, n = re.subn(f'typeface="{re.escape(old)}"', f'typeface="{new}"', xml)
        if n:
            part._blob = patched.encode("utf-8")
            hits += n
    print(f"typeface {old!r} -> {new!r}: {hits} references")
    return hits


def pin_font(prs, font=FONT):
    """Name `font` on every run on every slide, so nothing inherits the theme font.

    Runs that inherit are the ones that move when the theme font is substituted. The
    deck's own generators already pin Arial; the 28 carried-over donor slides do not.
    """
    pinned = 0
    for slide in prs.slides:
        for sh in shapes_flat(slide):
            if not sh.has_text_frame:
                continue
            for para in sh.text_frame.paragraphs:
                for run in para.runs:
                    if run.font.name != font:
                        run.font.name = font
                        pinned += 1
    print(f"pinned {font} on {pinned} runs")
    return pinned


def audit_fonts(prs, font=FONT, banned=("Play",)):
    """Fail the build if any run is off-font or any banned typeface survives."""
    stray = set()
    for i, slide in enumerate(prs.slides, 1):
        for sh in shapes_flat(slide):
            for para in sh.text_frame.paragraphs if sh.has_text_frame else []:
                for run in para.runs:
                    if run.font.name not in (None, font):
                        stray.add((i, run.font.name))
    if stray:
        raise SystemExit(f"runs not in {font}: {sorted(stray)}")
    for part in prs.part.package.iter_parts():
        if not str(part.partname).endswith(".xml"):
            continue
        blob = part.blob.decode("utf-8", "replace")
        for bad in banned:
            if f'typeface="{bad}"' in blob:
                raise SystemExit(f"{bad!r} still referenced in {part.partname}")


# Free-floating prose the donor drew onto its slides: a gloss under a diagram, a
# config instruction under a table. Every one of these is the speaker's line written
# on the slide, and two of the three pairs restate the slide's own title.
#
# These are the donor's, not this script's, and that is not a defence. Provenance
# decides whether a *claim* is safe to make; it does not make an aside belong on a
# slide. Anything genuinely needed goes in the speaker notes.
DONOR_ASIDES = (
    # Slide 17. The title already reads "The host looks idle. The main thread is full."
    "Half the machine is doing nothing.",
    "One core, and it is full. That is the ceiling.",
    # Slide 32, under the two CLUSTER SLOT-STATS tables.
    "On by default. Every slot looks the same.",
    "Needs cluster-slot-stats-enabled yes.",
)


def strip_asides(prs, texts=DONOR_ASIDES):
    """Delete every shape whose whole text is one of `texts`.

    Exact whole-shape match, so this cannot eat a label that merely contains the
    phrase. Each entry must hit at least once: a miss means the donor changed and the
    aside is still on a slide under different wording, which is exactly the failure
    this is here to prevent, so it raises rather than passing quietly.
    """
    removed = {}
    for i, slide in enumerate(prs.slides, 1):
        for sh in list(slide.shapes):
            if not sh.has_text_frame:
                continue
            body = sh.text_frame.text.strip()
            if body in texts:
                slide.shapes._spTree.remove(sh._element)
                removed.setdefault(body, []).append(i)
    missing = [t for t in texts if t not in removed]
    if missing:
        raise SystemExit(f"aside not found, donor may have changed: {missing}")
    total = sum(len(v) for v in removed.values())
    print(f"stripped {total} donor asides from slides "
          f"{sorted({i for v in removed.values() for i in v})}")
    return removed


def audit_no_stray_prose(prs, keep_prefixes=("* ",), max_pt=12):
    """Flag small free-floating text that is a sentence rather than a label.

    A heuristic, not a proof, and it is deliberately loud: anything at or below
    `max_pt` that ends in a full stop and is not a citation or a footnote marker is
    probably an aside that crept back. Source lines are 10pt italic and exempt;
    `CAP_NOTE` starts with "* " and is exempt.
    """
    suspects = []
    for i, slide in enumerate(prs.slides, 1):
        for sh in shapes_flat(slide):
            if not sh.has_text_frame:
                continue
            body = sh.text_frame.text.strip()
            if not body.endswith(".") or body.startswith(keep_prefixes):
                continue
            sizes = [r.font.size.pt for p in sh.text_frame.paragraphs
                     for r in p.runs if r.font.size]
            if not sizes or max(sizes) > max_pt:
                continue
            italic = any(r.font.italic for p in sh.text_frame.paragraphs
                         for r in p.runs)
            if italic:                      # source citations are 10pt italic
                continue
            suspects.append((i, body))
    if suspects:
        print("  note: small trailing-period text, check these are not asides:")
        for i, body in suspects:
            print(f"    slide {i}: {body!r}")
    return suspects


def flatten(shape):
    """Strip the theme shape style, so a shape renders as exactly its fill and line.

    `add_shape()` emits a `<p:style>` block referencing the theme's line, fill, **and
    effect**. Setting an explicit fill and `line.noFill` overrides the first two and
    leaves `<a:effectRef idx="2">` in place, which is a shadow. On a coloured card that
    is invisible. On a **white** rectangle used to mask something, the shadow is the
    only thing you see, and it reads as a dangling square.

    Nothing that relies on theme styling should go through this; it is for shapes whose
    appearance is fully specified.
    """
    style = shape._element.find(qn("p:style"))
    if style is not None:
        shape._element.remove(style)
    return shape


def audit_list_rows(prs, *layouts):
    """On every List Rows slide, assert items, row placeholders and bars all agree.

    The invariant `list_rows()` establishes, re-checked on the finished deck. An orphan
    bar with no row, or a row with no bar, means a slide was built by hand instead of
    through the helper.

    Rows are selected by idx rather than "everything but the title", because the image
    variant's picture placeholder is idx `LIST_ROWS + 1` and would otherwise count as a
    row. It also has no text frame once a picture is in it, so the emptiness check below
    raises `AttributeError` rather than reporting anything useful.
    """
    for i, slide in enumerate(prs.slides, 1):
        if slide.slide_layout.name not in layouts:
            continue
        rows = [p for p in slide.placeholders
                if 1 <= p.placeholder_format.idx <= LIST_ROWS]
        bars = [sh for sh in slide.shapes
                if sh.shape_type == 1 and sh.width == BAR_H and sh.height == LIST_ROW_H]
        if len(rows) != len(bars):
            raise SystemExit(f"slide {i}: {len(rows)} rows but {len(bars)} accent bars")
        want = {list_row_y(n) for n in range(len(rows))}
        if {sh.top for sh in bars} != want:
            raise SystemExit(f"slide {i}: bars at {sorted(sh.top for sh in bars)}, "
                             f"expected {sorted(want)}")
        empty = [p.placeholder_format.idx for p in rows if not p.text_frame.text.strip()]
        if empty:
            raise SystemExit(f"slide {i}: empty row placeholders {empty}")
    return True


def audit_paragraph_order(prs):
    """Fail the build on any `a:p` whose `a:endParaRPr` precedes a run.

    The schema sequence is pPr?, EG_TextRun*, endParaRPr?. PowerPoint enforces it by
    dropping the misplaced run; python-pptx reads the text back happily, and so does
    the mock renderer, so the slide looks correct everywhere except on screen. That
    combination is why this shipped once. It is cheap to check, so check it.
    """
    RUNS = ("r", "br", "fld")
    bad = []
    for i, slide in enumerate(prs.slides, 1):
        for sh in shapes_flat(slide):
            if not sh.has_text_frame:
                continue
            for para in sh.text_frame._txBody.findall(qn("a:p")):
                kids = [k.tag.split("}")[1] for k in para]
                if "endParaRPr" not in kids:
                    continue
                last_run = max((j for j, k in enumerate(kids) if k in RUNS), default=-1)
                if kids.index("endParaRPr") < last_run:
                    bad.append((i, sh.name, " ".join(kids)))
    if bad:
        for slide_no, name, kids in bad:
            print(f"  ! slide {slide_no} {name!r}: {kids}")
        raise SystemExit(f"{len(bad)} paragraphs have endParaRPr before a run; "
                         "PowerPoint will drop those runs")


# --- structural edits -----------------------------------------------------------


def clone_slide(prs, index):
    """Append a deep copy of slide `index`. Shapes only: no charts, no pictures.

    python-pptx has no slide-copy API. Deep-copying the shape tree works for
    autoshapes, text boxes and connectors, which carry no relationships. Anything
    with an `r:embed` or `r:id` needs its part copied too, so this refuses rather
    than silently producing a slide with a broken image.
    """
    src = prs.slides[index]
    dst = prs.slides.add_slide(src.slide_layout)
    for shape in list(dst.shapes):
        dst.shapes._spTree.remove(shape._element)
    for shape in src.shapes:
        xml = shape._element.xml
        if "r:embed" in xml or "r:link" in xml:
            raise ValueError(f"shape {shape.name!r} carries a relationship; "
                             "clone_slide cannot copy it")
        dst.shapes._spTree.append(copy.deepcopy(shape._element))
    if src.has_notes_slide and src.notes_slide.notes_text_frame.text.strip():
        notes(dst, src.notes_slide.notes_text_frame.text)
    return dst


def hide_slide(slide):
    """Mark a slide hidden: kept in the file, skipped when presenting.

    `show="0"` on `<p:sld>` is PowerPoint's own hidden-slide flag, which is the right
    tool for a slide worth keeping but not worth presenting. Deleting it would lose the
    work; leaving it visible would spend time on it. python-pptx has no API for the
    attribute, so set it directly.
    """
    slide._element.set("show", "0")
    return slide


def is_hidden(slide):
    return slide._element.get("show") == "0"


def reorder_and_prune(prs, order):
    """Set the slide order to `order`; drop every slide not listed.

    `order` holds current 0-based indices, each at most once. Doing both in one pass
    is what keeps the index arithmetic honest: dropping first renumbers everything
    after the hole, and reordering first means the drop list no longer means what it
    said. Anything omitted is deleted, so a typo loses a slide rather than silently
    duplicating one; hence the duplicate check.
    """
    lst = prs.slides._sldIdLst
    if len(set(order)) != len(order):
        raise ValueError("order lists the same slide twice")
    unknown = [i for i in order if not 0 <= i < len(lst)]
    if unknown:
        raise ValueError(f"order references slides that do not exist: {unknown}")

    ids = list(lst)
    dropped = [i for i in range(len(ids)) if i not in set(order)]
    for i in dropped:
        prs.part.drop_rel(ids[i].rId)     # pair with the sldIdLst removal below
    for el in ids:
        lst.remove(el)
    for i in order:
        lst.append(ids[i])
    return dropped


# --- slides built on the Valkey layouts -----------------------------------------
#
# These four replace hand-built donor slides. A design that exists as a layout should
# be used as a layout: the decoration is inherited, the geometry cannot drift, and the
# slide holds nothing but its text. That is the whole reason the layouts exist.


def on_layout(prs, layout_name):
    layout = prs.slide_layouts.get_by_name(layout_name)
    if layout is None:
        raise SystemExit(f"layout {layout_name!r} not in this deck")
    return prs.slides.add_slide(layout)


def fill(slide, idx, *lines, size=None, bold=None, colour=None):
    """Fill placeholder `idx` with one paragraph per line.

    A new slide's placeholders come up empty, so there is no formatting to preserve:
    set the text and let the layout's lstStyle supply size, weight and colour. Pass
    `size`/`bold`/`colour` only to override the layout on purpose, and say why.
    """
    ph = next((p for p in slide.placeholders if p.placeholder_format.idx == idx), None)
    if ph is None:
        raise LookupError(f"no placeholder idx={idx} on this slide")
    tf = ph.text_frame
    tf.text = lines[0]
    for line in lines[1:]:
        tf.add_paragraph().text = line
    for para in tf.paragraphs:
        for run in para.runs:
            run.font.name = FONT
            if size is not None:
                run.font.size = Pt(size)
            if bold is not None:
                run.font.bold = bold
            if colour is not None:
                run.font.color.theme_color = colour
    return ph


def drop_placeholders(slide, *idxs):
    """Remove placeholders that have no content, so PowerPoint shows no empty prompts."""
    spTree = slide.shapes._spTree
    for sh in list(slide.shapes):
        try:
            if sh.placeholder_format.idx in idxs:
                spTree.remove(sh._element)
        except (ValueError, AttributeError):
            pass


def slide_title(prs):
    """Title slide on `Valkey Title Band`, replacing the hand-built Holding slide.

    No conference-and-city line: the audience is at the conference, in the city.

    The subtitle placeholder is deleted rather than filled. The layout offers one, but
    nothing in any source supplies a subtitle, and writing one would be inventing copy
    for the speaker. An empty optional placeholder goes.
    """
    # One presenter, and the clustering band this deck shipped with rather than the
    # template's default art.
    from build_placeholder_deck import title_band
    return title_band(prs, [TALK_TITLE_L1, TALK_TITLE_L2], [(SPEAKER, ROLE)],
                      art=os.path.join(ASSETS, "band-clustering.jpg"))


def linked_subhead(slide, text, url, top=900000):
    """A subheader under the title whose text is a live hyperlink.

    A URL belongs here rather than in a 10pt footer when it is the thing you want the
    audience to act on: a footer reads as provenance, a subheader reads as an
    invitation. The colour is left unset so it inherits the theme's `hlink`, which is
    already Valkey blue, and so it keeps looking like a link if the theme changes.
    """
    box = slide.shapes.add_textbox(Emu(CONTENT_LEFT), Emu(top),
                                   Emu(CONTENT_W), Emu(340000))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    run = tf.paragraphs[0].add_run()
    run.text = text
    run.font.size = Pt(16)
    run.font.name = FONT
    run.hyperlink.address = url
    return box


def list_rows(slide, items):
    """Fill a `Valkey List Rows` slide with exactly `items`, bars included.

    This is the mechanism that keeps the row count honest. Three things have to agree:
    the number of items, the number of filled row placeholders, and the number of accent
    bars. Getting that wrong is what produced two phantom stubs on the goals slide
    through three separate attempted fixes, so it is enforced here rather than
    remembered:

    - More items than the layout has rows raises, rather than silently dropping one.
    - Unused row placeholders are deleted, which takes their grey cards with them.
    - Exactly one bar is drawn per used row, on the slide, so there is nothing left over
      to mask. Masking was the wrong instinct twice: a mask is a shape, and a shape
      carries its own theme effect.

    `audit_list_rows()` re-checks the invariant on the finished deck.
    """
    if not 1 <= len(items) <= LIST_ROWS:
        raise SystemExit(f"{len(items)} items; the layout has {LIST_ROWS} rows")
    for i, text in enumerate(items):
        fill(slide, i + 1, text)
        rect(slide, CONTENT_LEFT, list_row_y(i), BAR_H, LIST_ROW_H,
             fill=MSO_THEME_COLOR.ACCENT_1)
    drop_placeholders(slide, *range(len(items) + 1, LIST_ROWS + 1))
    return slide


def slide_goals(prs):
    """Valkey's architectural goals, on `Valkey List Rows`."""
    slide = on_layout(prs, "Valkey List Rows")
    set_title(slide, "Valkey\u2019s architectural goals")
    list_rows(slide, ["Efficiency trumps headline numbers",
                      "Make scaling operations seamless",
                      "Keep the command execution simple"])
    return slide


def slide_get_involved(prs):
    """How to get involved. Replaces the single-threading question as the closing beat.

    Three rows, in the order a newcomer needs them: the site, then the room where the
    work happens, then the one item with a deadline. Matrix, GitHub Discussions and the
    issue tracker are all real and all on valkey.io/connect, but five pointers is a
    directory, not an invitation; the site is the directory.

    Every item is verbatim from valkey.io/connect. No community-call day or time: that
    page lists none, and inventing a standing meeting would send people to an empty room.
    """
    slide = on_layout(prs, LAYOUT_LIST_ROWS_IMAGE)
    set_title(slide, "Getting involved")
    list_rows(slide, [
        "Start at valkey.io",
        "Talk to contributors on Slack: valkey.io/slack",
        "ValkeyConf 2026, 5 October, Prague",
    ])
    # The QR carries the same three pointers the rows do, for the half of the room that
    # will photograph a slide and never type a URL. No caption: a QR code beside a list
    # of links does not need to be told it is a QR code.
    slide.placeholders[LIST_ROWS + 1].insert_picture(
        os.path.join(ASSETS, "valkey-qr.png"))
    source_note(slide, "valkey.io/connect")
    notes(slide, "Point people at the site first; it links everything else, including "
                 "GitHub Discussions, Matrix and the issue tracker. Slack is where "
                 "contributors actually talk. ValkeyConf is the only item with a date.")
    return slide


def slide_thanks(prs):
    """Closing slide on `Valkey Thank You`, which bookends the title layout."""
    from build_placeholder_deck import thank_you
    return thank_you(prs, "Thank you! Let’s connect.",
                     [(SPEAKER, ["@reconditerose (X and Bluesky)", "madelynolson.com"])])


def retarget_layouts(prs, target="Valkey Title Only"):
    """Move every slide still on a stock layout onto `target`.

    Only safe because the audit says so: of the 25 slides on stock layouts, 24 use the
    title placeholder and nothing else, with their content as loose shapes and charts.
    `Valkey Title Only` has the same single title placeholder at the same geometry
    (TITLE_LEFT/TITLE_TOP), so nothing re-inherits differently. A slide using a *body*
    placeholder would lose its type scale, so this refuses those rather than quietly
    reflowing them.

    python-pptx has no API for changing a slide's layout: the link is a relationship
    on the slide part, so swap the relationship.
    """
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    layout = prs.slide_layouts.get_by_name(target)
    if layout is None:
        raise SystemExit(f"layout {target!r} missing")
    moved = []
    for i, slide in enumerate(prs.slides, 1):
        if slide.slide_layout.name.startswith("Valkey "):
            continue
        body = [p.placeholder_format.idx for p in slide.placeholders
                if p.placeholder_format.idx != 0]
        if body:
            raise SystemExit(f"slide {i} uses placeholders {body}; it needs a real "
                             f"layout, not {target}")
        for rId, rel in list(slide.part.rels.items()):
            if rel.reltype == RT.SLIDE_LAYOUT:
                slide.part.drop_rel(rId)
        slide.part.relate_to(layout.part, RT.SLIDE_LAYOUT)
        moved.append(i)
    print(f"retargeted {len(moved)} slides to {target!r}: {moved}")
    return moved


def drop_second_master(prs):
    """Delete the unused second slide master and its 14 layouts.

    The donor carried two masters: the Valkey one, and a Google-Slides import holding
    duplicate layouts (`Title + 2x Speaker`, `Divider`, `Thank You 2x`, another
    `Holding`) with full-bleed dark background pictures. No slide in this deck uses
    any of them. Dropping the master's relationship drops its layouts and their media
    with it, because python-pptx only serialises parts reachable from the graph.
    """
    masters = prs.slide_masters
    if len(masters) < 2:
        return 0
    used = {s.slide_layout.slide_master._element for s in prs.slides}
    doomed = [m for m in masters if m._element not in used]
    if len(doomed) != len(masters) - 1:
        raise SystemExit(f"expected exactly one unused master, found {len(doomed)}")
    n = 0
    for master in doomed:
        count = len(master.slide_layouts)
        for rId, rel in list(prs.part.rels.items()):
            if rel.reltype.endswith("slideMaster") and rel.target_part is master.part:
                prs.part.drop_rel(rId)
                n += 1
        lst = prs.part._element.find(qn("p:sldMasterIdLst"))
        for el in list(lst):
            if el.get(qn("r:id")) not in {r.rId for r in prs.part.rels.values()}:
                lst.remove(el)
        print(f"dropped unused slide master and its {count} layouts")
    return n


# --- new slides -----------------------------------------------------------------


def slide_bio(prs):
    """Speaker bio. Ported from P99 slide 2, rebuilt at Valkey geometry."""
    slide = prs.slides.add_slide(prs.slide_layouts.get_by_name(LAYOUT_TITLE_ONLY))
    set_title(slide, "Madelyn Olson (she/her)")

    # Landscape 4:3 photo. Height matches what the old portrait crop occupied, so the
    # text column beside it keeps its vertical rhythm; the width follows from 4:3.
    photo_h = 3290700
    photo_w = int(photo_h * 4 / 3)
    slide.shapes.add_picture(os.path.join(ASSETS, "madelyn-winery.jpg"),
                             Emu(CONTENT_LEFT), Emu(CONTENT_TOP),
                             Emu(photo_w), Emu(photo_h))

    tx = CONTENT_LEFT + photo_w + 500000
    tw = SLIDE_W - CONTENT_LEFT - tx
    text_box(slide, tx, CONTENT_TOP, tw, 400110,
             [("Valkey maintainer", 24, True, MSO_THEME_COLOR.DARK_1)])
    text_box(slide, tx, CONTENT_TOP + 460000, tw, 400110,
             [("Principal engineer at AWS", 24, False, MSO_THEME_COLOR.DARK_2)])
    rect(slide, tx, CONTENT_TOP + 1050000, 1371600, 54864, fill=MSO_THEME_COLOR.ACCENT_1)

    items = ["Original co-creator of the Valkey fork",
             "I love hearing diverse opinions about performance",
             "I enjoy hiking, nature, and good wine"]
    for i, item in enumerate(items):
        y = CONTENT_TOP + 1350000 + i * 620000
        rect(slide, tx, y, tw, 480000, fill=MSO_THEME_COLOR.LIGHT_2)
        rect(slide, tx, y, BAR_H, 480000, fill=MSO_THEME_COLOR.ACCENT_1)
        text_box(slide, tx + BAR_H + PAD, y, tw - BAR_H - 2 * PAD, 480000,
                 [(item, 18, False, MSO_THEME_COLOR.DARK_1)], anchor=MSO_ANCHOR.MIDDLE)

    # The opener, in Spanish, verbatim as the speaker wrote it. Do not "correct" the
    # accents or the phrasing: it is a script a person is going to read out loud.
    notes(slide, SPANISH_OPENER)
    return slide


TITLE_BOTTOM, LOGO_BASELINE = TITLE_TOP + TITLE_H, 5943597
# Enough for the icon, heading and a two-line body. CARD_H (3.9in) is sized for a
# denser card than these carry, and leaves a third of each one empty.
SHORT_CARD_H = 2600000


def centred_top(block_h):
    """Top edge that centres `block_h` between the title and the logo baseline."""
    return (TITLE_BOTTOM + LOGO_BASELINE - block_h) // 2


def card_with_mark(slide, x, y, w, h, mark, heading, body):
    """A LIGHT_2 card whose top slot holds an icon, a brand SVG, or a raster logo.

    `card()` in the placeholder builder only takes an icon name. A logo is not an
    icon: it has its own aspect ratio and must never be squared off, so this sizes
    to a fixed height and derives the width. Everything below the mark matches
    card_geometry() so these cards line up with icon cards on other slides.
    """
    kind = mark[0]
    rect(slide, x, y, w, h, fill=MSO_THEME_COLOR.LIGHT_2)
    icon_x, icon_y = x + PAD, y + PAD
    slot = 1097280                      # CARD_ICON: the height an icon would occupy
    if kind == "icon":
        add_svg(slide, os.path.join(ICONS, f"{mark[1]}.svg"),
                os.path.join(ICON_PNGS, f"{mark[1]}.png"),
                icon_x, icon_y, slot, slot, name=mark[1])
    elif kind == "logo":
        # A mark is taller than wide; a wordmark lockup is wider than tall. Fit the
        # slot's height for the former and leave the latter to the png branch.
        height = slot
        add_svg(slide, os.path.join(ICONS, mark[1]),
                os.path.join(ICON_PNGS, mark[1].replace(".svg", ".png")),
                icon_x, icon_y, int(height * mark[2]), height, name=mark[1])
    elif kind == "png":
        # Wide lockup: fit a band, not a square, so it does not dwarf the icons.
        height = int(slot * 0.55)
        slide.shapes.add_picture(os.path.join(ASSETS, mark[1]),
                                 Emu(icon_x), Emu(icon_y + (slot - height) // 2),
                                 Emu(int(height * mark[2])), Emu(height))
    else:
        raise ValueError(f"unknown mark kind {kind!r}")

    inner_w = w - 2 * PAD
    hy = icon_y + slot + 137160         # CARD_HEAD_GAP
    by = hy + 400110 + 45720            # CARD_HEAD_H + CARD_BODY_GAP
    text_box(slide, x + PAD, hy, inner_w, 400110,
             [(heading, 18, True, MSO_THEME_COLOR.ACCENT_2)])
    text_box(slide, x + PAD, by, inner_w, h - (by - y) - PAD,
             [(body, 13, False, MSO_THEME_COLOR.ACCENT_2)], spacing=1.2)


def slide_intro(prs):
    """Introducing the Valkey project. Ported from Valkey community deck slide 3.

    The community deck lays this out as three loose text columns plus a Linux
    Foundation PNG and a freeform graphic. Same three claims, rebuilt as the card
    treatment this deck already uses, so it does not read as a pasted-in slide.
    Its speaker notes carry the three points verbatim.
    """
    slide = prs.slides.add_slide(prs.slide_layouts.get_by_name(LAYOUT_TITLE_ONLY))
    set_title(slide, "Introducing the Valkey project")
    col_w, xs = box_columns(3)
    # Two of the three claims are about *who owns the project*, and a logo says that
    # faster than an icon does: the Valkey mark for compatibility, the Linux
    # Foundation lockup for vendor neutrality. The third is about people, so it keeps
    # an icon. Both raster logos come from the community deck's own slide 3.
    specs = [
        (("logo", "valkey-mark-color.svg", 282.6 / 326.0), "Drop-in compatible",
         "Fully compatible with Redis OSS 7.2, and always will be."),
        (("png", "linux-foundation.png", 388 / 130), "Vendor neutral",
         "A Linux Foundation project, BSD-3 licensed, so it stays open source."),
        # `community` is three linked figures, which says "a team", not "the world".
        # `global-community` is a globe with scattered points, which is the claim:
        # the committee is spread across companies and countries.
        (("icon", "global-community"), "Built in the open",
         "Nine technical steering committee members, from eight companies."),
    ]
    top = centred_top(SHORT_CARD_H)
    for x, spec in zip(xs, specs):
        card_with_mark(slide, x, top, col_w, SHORT_CARD_H, *spec)
    notes(slide,
          "Three key points about Valkey:\n"
          "1. It's fully compatible with Redis OSS 7.2, and will always be.\n"
          "2. It's backed by the Linux Foundation, so it will remain open source "
          "and vendor neutral.\n"
          "3. It has a diverse backing. There are nine technical steering members, "
          "from eight companies.")
    return slide


def slide_rows(prs, title, rows, note=None, source=None):
    """Title plus one LIGHT_2 row per item with an ACCENT_1 left edge.

    The row treatment rather than caret bullets: this deck is built from filled
    cards, and raw text on white is the one slide type with no structure in it.
    """
    slide = prs.slides.add_slide(prs.slide_layouts.get_by_name(LAYOUT_TITLE_ONLY))
    set_title(slide, title)
    row_h, pitch = 620000, 760000
    block_h = (len(rows) - 1) * pitch + row_h
    top = (SLIDE_H - block_h) // 2
    for i, row in enumerate(rows):
        y = top + i * pitch
        rect(slide, CONTENT_LEFT, y, CONTENT_W, row_h, fill=MSO_THEME_COLOR.LIGHT_2)
        rect(slide, CONTENT_LEFT, y, BAR_H, row_h, fill=MSO_THEME_COLOR.ACCENT_1)
        runs = row if isinstance(row, list) else [(row, 18, False, MSO_THEME_COLOR.DARK_1)]
        text_box(slide, CONTENT_LEFT + BAR_H + PAD, y,
                 CONTENT_W - BAR_H - 2 * PAD, row_h, runs, anchor=MSO_ANCHOR.MIDDLE)
    if note:
        footnote(slide, note)
    if source:
        source_note(slide, source)
    return slide




def slide_customer_latency(prs):
    """A customer latency story. Ported from P99 slide 3.

    Built rather than copied: the P99 chart is on a 1.5x canvas. The trace is the
    real 520-point series, extracted from that deck into assets/.
    """
    slide = prs.slides.add_slide(prs.slide_layouts.get_by_name(LAYOUT_TITLE_ONLY))
    set_title(slide, "A customer latency story")

    with open(os.path.join(ASSETS, "customer-latency-trace.json")) as f:
        trace = json.load(f)

    cd = CategoryChartData()
    cd.categories = [""] * len(trace)          # 520 ticks would be unreadable
    cd.add_series("Latency (ms)", trace)
    frame = slide.shapes.add_chart(XL_CHART_TYPE.LINE, Emu(CONTENT_LEFT),
                                   Emu(CONTENT_TOP), Emu(CONTENT_W), Emu(3700000), cd)
    chart = frame.chart
    chart.has_legend = False                   # one series; the title says which
    series = chart.plots[0].series[0]
    series.format.line.color.theme_color = MSO_THEME_COLOR.ACCENT_1
    series.format.line.width = Pt(1.25)
    chart.category_axis.has_major_gridlines = False
    chart.value_axis.has_major_gridlines = True
    # 520 blank categories give the category axis nothing to label, and an axis title
    # squeezed under them renders tiny and unreadable. Drop the axis furniture and
    # say "time" once, at body size, in a label of its own.
    chart.category_axis.has_title = False
    chart.category_axis.tick_labels.font.size = Pt(1)
    chart.value_axis.has_title = True
    chart.value_axis.axis_title.text_frame.text = "latency (ms)"
    chart.value_axis.tick_labels.font.size = Pt(12)
    text_box(slide, CONTENT_LEFT, CONTENT_TOP + 3700000 + 60000, CONTENT_W, 300000,
             [("time", 14, False, MSO_THEME_COLOR.DARK_2)], align=PP_ALIGN.CENTER)

    # No caption. The trace is the whole point and it speaks for itself; anything
    # written under it would be a claim about the customer's workload that this
    # repo has no evidence for.
    notes(slide, "This is the shape that starts every one of these "
                 "conversations. The average is fine. The customer is not.")
    return slide



FLAME_W, FLAME_H = 2402, 1328        # assets/valkey-flamegraph.png, in pixels


def translucent(shape, rgb, alpha):
    """Give a shape a translucent solid fill.

    python-pptx has no alpha: it has to be an `<a:alpha>` child of the colour
    element, and it only attaches to `<a:srgbClr>`. Setting a theme colour writes
    `<a:schemeClr>` instead, which is why this takes explicit RGB.
    """
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb
    srgb = shape.fill.fore_color._xFill.find(qn("a:srgbClr"))
    srgb.append(srgb.makeelement(qn("a:alpha"), {"val": str(alpha)}))
    return shape


def slide_flamegraph(prs):
    """Where the CPU actually goes. Ported from P99 slide 9.

    A real `valkey-server` flame graph, partitioned into the four phases. This is the
    slide that earns the whole threading section: command execution is the leftmost
    fifth, and the write path plus the kernel network stack is about half the
    profile. A drawn bar chart cannot make that argument, which is why the first
    rebuild of this slide, which dropped the graph, was wrong.

    Band boundaries are P99's, divided by 1.5 for this canvas. Two deliberate
    departures from P99, both so the graph stays readable:

    - The tint is 25%, not 60%. At 60% a wash of blue buries the function names,
      and the function names are the reason to show a flame graph at all.
    - Each label gets a solid chip instead of floating on the tint, so dropping the
      tint does not cost label contrast.

    No title placeholder, as in P99: the bands run the full height of the image, so
    anything in the title band would sit under them. The labels are the title.
    """
    slide = prs.slides.add_slide(prs.slide_layouts.get_by_name(LAYOUT_TITLE_ONLY))
    slide.shapes._spTree.remove(slide.shapes.title._element)

    # Fit the full width: a flame graph is unreadable inset, and the aspect is within
    # 2% of the slide's, so preserving it costs only a sliver top and bottom.
    img_h = int(SLIDE_W * FLAME_H / FLAME_W)
    img_y = (SLIDE_H - img_h) // 2
    slide.shapes.add_picture(os.path.join(ASSETS, "valkey-flamegraph.png"),
                             Emu(0), Emu(img_y), Emu(SLIDE_W), Emu(img_h))

    edges = [0, 2679405, 3785191, 5769935, SLIDE_W]
    labels = ["Parse and execute command", "Epoll", "Read command", "Send response"]
    chip_h = 420000
    for i, label in enumerate(labels):
        x, w = edges[i], edges[i + 1] - edges[i]
        translucent(rect(slide, x, img_y, w, img_h), VALKEY_BLUE, 25000)
        # P99 gives all four bands the same fill and no border, so the partition is
        # invisible: it reads as one blue wash. A hairline divider is the minimum
        # needed for the four regions to read as regions.
        if i:
            arrow(slide, x, img_y, x, img_y + img_h,
                  colour=MSO_THEME_COLOR.LIGHT_1, pt=1.25, head=False)
        rect(slide, x, img_y, w, chip_h, fill=MSO_THEME_COLOR.ACCENT_1)
        text_box(slide, x + 40000, img_y, w - 80000, chip_h,
                 [(label, 13, True, MSO_THEME_COLOR.LIGHT_1)],
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Dark text would be illegible on the tint; and this sits over the image, not on
    # white, so it cannot use source_note().
    text_box(slide, CONTENT_LEFT, img_y + img_h - 380000, CONTENT_W, 300000,
             [("perf flame graph, valkey-server under valkey-benchmark GET.",
               10, False, MSO_THEME_COLOR.ACCENT_2)], italic=True)
    notes(slide, "Single threaded vs multi-threaded phases. Only the leftmost band "
                 "has to be serialised; the rest is I/O and kernel network stack, "
                 "and that is what moved to the IO threads.")
    return slide


def slide_operator(prs):
    """Introducing the Valkey Operator. Three icon cards, the three-box design.

    Closes the "going out" section. Everything before it is a command someone has to
    run in the right order; the operator is who runs them.

    **Not on the `Valkey Three Box` layout, deliberately.** That layout bakes its three
    icons in as static layout shapes (`performance`, `security`, `reliability`), and a
    static layout shape cannot be deleted on a slide, only drawn over. None of those
    three fit these features. So this builds the same design with `card()` and the right
    icons, exactly as `slide_intro()` does for the same reason.

    The three features are chosen to land in *this* talk rather than to summarise the
    project: feature 1 is the slot migration the previous ten slides just explained, now
    driven by a controller. Drawn from the repo README and the v0.6.0 release notes.

    The repo is a linked subheader under the title, not a footer: it is the one thing on
    this slide an audience can act on, and a 10pt footer reads as provenance instead of
    as an invitation. The version and the alpha warning stay in the source line, because
    those *are* provenance.
    """
    slide = prs.slides.add_slide(prs.slide_layouts.get_by_name(LAYOUT_TITLE_ONLY))
    set_title(slide, "Introducing the Valkey Operator")
    col_w, xs = box_columns(3)
    specs = [
        ("horizontal-scaling", "Scaling",
         "Set spec.shards. The operator rebalances slots on the way out and drains "
         "them on the way in."),
        # "Rolling upgrades and failover" is 29 characters and wraps to two lines in
        # a 3.45in card, which runs into the body: card_geometry sizes the heading
        # slot for one line. Keep card headings to about 22 characters.
        ("reliability", "Upgrades and failover",
         "Replicas roll before primaries, with a proactive failover before the "
         "primary goes."),
        ("security", "TLS and ACLs, declared",
         "Certificates and users live in the spec. ACL changes apply without a "
         "restart."),
    ]
    top = centred_top(SHORT_CARD_H)
    for x, spec in zip(xs, specs):
        card(slide, x, top, col_w, SHORT_CARD_H, MSO_THEME_COLOR.LIGHT_2, *spec)
    linked_subhead(slide, "github.com/valkey-io/valkey-operator",
                   "https://github.com/valkey-io/valkey-operator")
    source_note(slide, "v0.6.0, alpha and not production ready")
    notes(slide, "Alpha. Breaking API changes and data loss are both called out in the "
                 "project's own release notes, so do not oversell it. The point is that "
                 "the slot migration primitives from the last section are something a "
                 "controller drives for you, not a runbook you follow.")
    return slide

def slide_migration_benchmark(prs):
    """Legacy vs atomic slot migration. Ported from P99 slide 29.

    Two cards, one per load condition, four bar pairs each. Bars are drawn to the
    same scale within a card so legacy against atomic is a visual comparison, not
    four numbers to read.
    """
    slide = prs.slides.add_slide(prs.slide_layouts.get_by_name(LAYOUT_TITLE_ONLY))
    set_title(slide, "Legacy vs atomic slot migration")

    cards = [
        ("No load", [("3 → 4 shards", 102, 11, "9.5x"),
                     ("4 → 3 shards", 80, 10, "8.4x")]),
        ("Heavy load", [("3 → 4 shards", 147, 31, "4.8x"),
                        ("4 → 3 shards", 125, 27, "4.6x")]),
    ]
    col_w, xs = box_columns(2)
    card_h = 3900000
    worst = max(sec for _, pairs in cards for _, sec, _, _ in pairs)

    for x, (heading, pairs) in zip(xs, cards):
        rect(slide, x, CONTENT_TOP, col_w, card_h, fill=MSO_THEME_COLOR.LIGHT_2)
        text_box(slide, x + PAD, CONTENT_TOP + PAD, col_w - 2 * PAD, 460000,
                 [(heading, 24, True, MSO_THEME_COLOR.ACCENT_2)])
        for j, (label, legacy, atomic, speedup) in enumerate(pairs):
            gy = CONTENT_TOP + 900000 + j * 1500000
            text_box(slide, x + PAD, gy, col_w - 2 * PAD, 300000,
                     [(label, 16, True, MSO_THEME_COLOR.DARK_1)])
            for k, (who, secs, hot) in enumerate((("legacy", legacy, False),
                                                  ("atomic", atomic, True))):
                ry = gy + 400000 + k * 400000
                text_box(slide, x + PAD, ry - 40000, 900000, 340000,
                         [(who, 12, False, MSO_THEME_COLOR.DARK_2)],
                         anchor=MSO_ANCHOR.MIDDLE)
                track_x = x + PAD + 950000
                track_w = col_w - 2 * PAD - 950000 - 1500000
                rect(slide, track_x, ry, track_w, 260000, fill=MSO_THEME_COLOR.ACCENT_5)
                rect(slide, track_x, ry, max(int(track_w * secs / worst), 20000),
                     260000, fill=MSO_THEME_COLOR.ACCENT_1 if hot
                     else MSO_THEME_COLOR.ACCENT_2)
                mins, rem = divmod(secs, 60)
                shown = f"{mins}m {rem:02d}s" if mins else f"{rem}s"
                runs = [(shown, 12, hot, MSO_THEME_COLOR.DARK_1)]
                if hot:
                    runs.append((f"   {speedup} faster", 12, True,
                                 MSO_THEME_COLOR.ACCENT_1))
                text_box(slide, track_x + track_w + 160000, ry - 40000, 1400000, 340000,
                         runs, anchor=MSO_ANCHOR.MIDDLE)
        # No "One key per MIGRATE: ~213K round trips." line here. It was an aside
        # explaining the chart rather than labelling it, and the round-trip count was
        # derived by this script, not taken from the cited blog post.

    source_note(slide, "Valkey 9.0, c4-standard-8, 40 GB of 16 KB values, "
                       "valkey-cli --cluster rebalance defaults. "
                       "Source: valkey.io/blog/atomic-slot-migration")
    return slide


def slide_cluster_routing(prs):
    """How a key finds its shard. Sits between the cluster picture and migration.

    Needed because everything after it is about moving slots, and "slot" has not
    been explained: the previous slide only labels three ranges. Nerdearla is not a
    developer-only audience, so the hash-to-slot-to-shard chain cannot be assumed.
    """
    slide = prs.slides.add_slide(prs.slide_layouts.get_by_name(LAYOUT_TITLE_ONLY))
    set_title(slide, "Every key belongs to exactly one slot")

    # Labels only, no captions. Each caption under these boxes restated the box:
    # "slot 9158" / "slots are the unit of ownership" is the same fact twice, and the
    # second copy is the line the speaker says out loud.
    steps = ['GET "user:42"', "CRC16(key) mod 16384", "slot 9158", "shard 2"]
    box_w, gap = 2400000, 450000
    total = len(steps) * box_w + (len(steps) - 1) * gap
    left = (SLIDE_W - total) // 2
    box_h, y = 800000, 2250000

    for i, label in enumerate(steps):
        x = left + i * (box_w + gap)
        hot = i >= 2                       # slot and shard: the two new words
        rect(slide, x, y, box_w, box_h,
             fill=MSO_THEME_COLOR.ACCENT_1 if hot else MSO_THEME_COLOR.LIGHT_2)
        text_box(slide, x + 60000, y, box_w - 120000, box_h,
                 [(label, 14, True, MSO_THEME_COLOR.LIGHT_1 if hot
                   else MSO_THEME_COLOR.DARK_1)],
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        if i:
            arrow(slide, x - gap, y + box_h // 2, x, y + box_h // 2)

    # The reason any of this matters for the rest of the talk.
    note_y = 4300000
    rect(slide, CONTENT_LEFT, note_y, CONTENT_W, 620000, fill=MSO_THEME_COLOR.LIGHT_2)
    rect(slide, CONTENT_LEFT, note_y, BAR_H, 620000, fill=MSO_THEME_COLOR.ACCENT_1)
    text_box(slide, CONTENT_LEFT + BAR_H + PAD, note_y,
             CONTENT_W - BAR_H - 2 * PAD, 620000,
             [("Scaling out moves slots, not keys. ", 18, True,
               MSO_THEME_COLOR.DARK_1),
              ("A client that asks the wrong shard gets ", 18, False,
               MSO_THEME_COLOR.DARK_1),
              ("MOVED", 18, True, MSO_THEME_COLOR.ACCENT_1),
              (" and retries.", 18, False, MSO_THEME_COLOR.DARK_1)],
             anchor=MSO_ANCHOR.MIDDLE)
    notes(slide, "16384 slots, fixed. The mapping from key to slot never changes; "
                 "only the mapping from slot to shard does. That is what every "
                 "migration slide after this is about.")
    return slide


def slide_getkeysinslot(prs, src_index):
    """Legacy migration frame 1. Ported from P99 slide 22.

    A clone of the MIGRATE frame with the caption and the active rail step changed,
    so the two frames stay pixel-aligned as a build. Cloned after the rail has been
    trimmed, so both frames get the two-step rail.
    """
    slide = clone_slide(prs, src_index)
    set_title(slide, "CLUSTER GETKEYSINSLOT")
    retext(by_text(slide, "one key per MIGRATE, clients follow -ASK"),
           "list every key in the slot")
    trim_legacy_steps(slide, active="GETKEYSINSLOT")
    clear_notes(slide)
    return slide


# --- edits to the slides carried over from the donor -----------------------------



def edit_two_box(slide, title, vertical=None, horizontal=None):
    """Retitle a two-box slide and optionally rewrite a card's body.

    The donor titles these 'Two-box placeholder (neutral)', which is a working
    title left in by accident. Each one now says what its state of the argument is.

    `vertical` / `horizontal` are lists of paragraphs; a paragraph is a string, or a
    list of (text, bold, strike) runs. The struck-through form is load-bearing: these
    four slides are a build, and each step strikes the claim the talk has just
    disproved and writes the replacement beside it. Passing plain strings for those
    lines is what silently turned the build into four near-identical slides.
    """
    set_title(slide, title)
    for want, body in (("Vertical Scaling", vertical), ("Horizontal Scaling", horizontal)):
        if body is None:
            continue
        head = by_text(slide, want)
        target = min(
            (sh for sh in shapes_flat(slide)
             if sh.has_text_frame and sh.top > head.top and sh.left == head.left),
            key=lambda sh: sh.top)
        lines = [[(p, None)] if isinstance(p, str) else p for p in body]
        set_runs(target, *lines)


RAIL_TOP, RAIL_H, RAIL_BOX_W, RAIL_GAP = 5000000, 400000, 2518200, 180000
RAIL_STEPS = ("GETKEYSINSLOT", "MIGRATE")


def drop_section_label(slide):
    """Delete the donor's stray section label from the title band.

    Donor slides 17-23 each carry a leftover `Google Shape;204;p8` at
    (762000, 552419), the same size as the title placeholder and 152400 EMU to its
    right, so "Legacy slot migration" renders on top of "CLUSTER SETSLOT". It cannot
    simply be ignored: left alone it draws over the title.

    An earlier pass promoted it into a running-header band instead of deleting it, on
    the grounds that it is the only on-screen marker of which migration mechanism is
    showing across slides 23-29. That reasoning was wrong on balance. A running header
    on 8 of 36 slides is not a running header, it is an inconsistency, and the skill
    doc's own rule is every slide or none. The legacy-versus-atomic distinction is the
    speaker's to make out loud, or the titles' to carry.

    Returns the deleted label text, or None if this slide has no stray box.
    """
    strays = [sh for sh in shapes_flat(slide)
              if sh._element is not slide.shapes.title._element and sh.has_text_frame
              and sh.top == 552419 and sh.left == 762000]
    if not strays:
        return None
    if len(strays) != 1:
        raise LookupError(f"{len(strays)} stray title-band boxes; expected 1")
    label = strays[0].text_frame.text.strip()
    slide.shapes._spTree.remove(strays[0]._element)
    return label

def trim_legacy_steps(slide, active):
    """Edit 18: the step rail drops the SETSLOT bookkeeping.

    The donor rail is four steps: SETSLOT IMPORTING, SETSLOT MIGRATING,
    GETKEYSINSLOT + MIGRATE, SETSLOT NODE. P99 cut it to the two commands that
    actually move data. Rebuilt rather than edited in place: the donor draws each
    step as a rectangle with a separate text box on top at identical coordinates,
    and re-laying out four of those to make two is more fiddly than drawing two.
    """
    for sh in list(shapes_flat(slide)):
        if sh.top in (RAIL_TOP, 5200000):      # step boxes, their labels, connectors
            slide.shapes._spTree.remove(sh._element)
    left = (SLIDE_W - (2 * RAIL_BOX_W + RAIL_GAP)) // 2
    for i, step in enumerate(RAIL_STEPS):
        x = left + i * (RAIL_BOX_W + RAIL_GAP)
        hot = step == active
        chip(slide, x, RAIL_TOP, RAIL_BOX_W, RAIL_H, step,
             fill=MSO_THEME_COLOR.ACCENT_1 if hot else MSO_THEME_COLOR.LIGHT_2,
             fg=MSO_THEME_COLOR.LIGHT_1 if hot else MSO_THEME_COLOR.DARK_2, size=11)
    arrow(slide, left + RAIL_BOX_W, RAIL_TOP + RAIL_H // 2,
          left + RAIL_BOX_W + RAIL_GAP, RAIL_TOP + RAIL_H // 2,
          colour=MSO_THEME_COLOR.DARK_2, pt=1.0)


def main():
    if not os.path.exists(DONOR):
        raise SystemExit(f"donor missing: {DONOR}")
    for lock in (os.path.join(os.path.dirname(p), f"~${os.path.basename(p)}") for p in (DONOR, OUT)):
        if os.path.exists(lock):
            print(f"note: lock file present, PowerPoint may have it open: {lock}")

    tmp = OUT + ".tmp"
    shutil.copyfile(DONOR, tmp)
    prs = Presentation(tmp)
    s = prs.slides
    if len(s) != 29:
        raise SystemExit(f"donor has {len(s)} slides, expected 29; re-dump it first")

    # --- fold in the Valkey slide template ---
    # The eleven Valkey layouts and the brand-rose accent6, both from
    # build_placeholder_deck.py. Layouts first: every slide this script builds resolves
    # `Valkey Title Only` by name, so they have to exist before the first add_slide.
    build_layouts(prs)
    patch_theme(prs)

    # --- edits to carried-over slides, by donor index (0-based) ---
    # Donor slide 0 (title), 26 (goals), 27 (closing question) and 28 (thank you) are
    # not edited any more: each is rebuilt on the Valkey layout that already carries
    # its design, and the donor originals are dropped in `order`.

    # Edit 10 + 17: footnote on both latency-cliff charts, and the missing setup note.
    # The footnote is P99's own wording, verbatim. An earlier pass had it as "capped at
    # 5 ms by the benchmark's timeout"; the cause was never in the source.
    footnote(s[1], CAP_NOTE)
    notes(s[1], "Clients in the same availability zone as the server.")

    # Edit 7: 300K rps is an xlarge, not a large. The donor mislabels it.
    rename_series(s[3], "Provisioned capacity (r6g.large)",
                  "Provisioned capacity (r6g.xlarge)")
    notes(s[3], "We would like customers to be able to over provision their "
                "clusters. This is what it costs.")

    edit_two_box(s[4], "Two axes, and neither one is optional")

    # The IO-thread builds carry speaker notes copy-pasted from a memory-efficiency
    # slide in another deck: they talk about hash table pointer chasing. Wrong slide,
    # wrong topic, seven times over.
    for i in range(5, 12):
        clear_notes(s[i])
    notes(s[6], "Single threaded vs multi-threaded phases.")

    # Edit 8 + 12 + 10: the run was 8 IO threads, not 9.
    set_title(s[13], "Latency is (still) a cliff though")
    replace_series(s[13], [
        ("p50 (1 IO thread)", [145.4, 425.6, 697.0] + [5000.0] * 21),
        ("p99 (1 IO thread)", [650.2, 858.0, 1557.0] + [5000.0] * 21),
        ("p50 (8 IO threads)",
         [133.4, 181.0, 202.0, 245.3, 280.7, 313.0, 339.7, 374.3, 414.0, 451.7,
          480.7, 512.6, 543.9, 578.0, 616.8, 646.0, 681.5, 716.9, 750.7, 775.0,
          811.9, 853.3, 5000.0, 5000.0]),
        ("p99 (8 IO threads)",
         [589.7, 511.7, 542.8, 852.5, 604.2, 844.9, 960.8, 1066.8, 843.5, 1232.4,
          1214.7, 1556.6, 973.1, 1820.4, 1205.4, 1866.9, 1650.1, 2059.6, 1445.9,
          1658.2, 1450.9, 2334.5, 5000.0, 5000.0]),
    ])
    footnote(s[13], CAP_NOTE)

    # Edit 9: the capacity line now steps between two real instance sizes rather
    # than three invented ones.
    set_title(s[14], "Adapt the requirements to the workload")
    replace_series(s[14], [
        ("Cache throughput",
         [15000.0, 12000.0, 10000.0, 8000.0, 10000.0, 15000.0, 20000.0, 30000.0,
          35000.0, 40000.0, 50000.0, 90000.0, 180000.0, 120000.0, 60000.0, 55000.0,
          70000.0, 110000.0, 150000.0, 160000.0, 140000.0, 100000.0, 60000.0,
          30000.0, 15000.0]),
        ("Adapted capacity",
         [150000.0] * 12 + [300000.0] * 2 + [150000.0] * 3 + [300000.0] * 5
         + [150000.0] * 3),
    ])

    # Edit 11: measured scaling is about 8x, not 10x.
    #
    # The two-box slides are a four-step build, and the mechanism is a strikethrough:
    # each step crosses out the limitation the talk has just disproved and writes the
    # replacement next to it, in the same paragraph. So the vertical card goes from
    # "Limited with IO threading" to "~~Limited with IO threading~~ Provides about 8x
    # throughput scaling", and the horizontal card gets the same treatment later.
    # Frame 1 (donor slide 5) already reads "Limited with IO threading" unstruck, so
    # it needs no rewrite; frames 2 and 4 strike it.
    VERTICAL_8X = [[("Add more cores to handle incoming requests", None)],
                   [("Limited with IO threading ", None, True),
                    ("Provides about 8x throughput scaling", None)]]
    HORIZONTAL_SLOW = ["Add more hosts to parallelize requests",
                       "No cross-shard transactions",
                       "Slow and dangerous to scale"]
    HORIZONTAL_FAST = ["Add more hosts to parallelize requests",
                       "No cross-shard transactions",
                       [("Slow and dangerous to scale", None, True),
                        (" Fast and safe operation", None)]]
    edit_two_box(s[15], "Up: about 8x, and then it stops",
                 vertical=VERTICAL_8X, horizontal=HORIZONTAL_SLOW)

    # Donor defect, not a P99 edit: slides 17-23 each draw a section label on top
    # of their own title. Promote it to a running header over the whole section.
    for i in range(16, 23):
        if drop_section_label(s[i]) is None:
            raise SystemExit(f"donor slide {i + 1} has no stray section label")

    # Edit 18, on both legacy frames. MIGRATE is the active step on each: the
    # GETKEYSINSLOT frame is cloned off s[17] later and overrides it.
    trim_legacy_steps(s[17], active="MIGRATE")
    set_title(s[17], "MIGRATE")
    trim_legacy_steps(s[18], active="MIGRATE")

    # "Out:" undersold it. Atomic migration is measured in both directions on the
    # previous slide (3 -> 4 shards and 4 -> 3, under no load and heavy load), so the
    # claim is horizontal scaling either way, not just scaling out.
    edit_two_box(s[23], "Horizontal scaling, in and out: now fast and safe",
                 vertical=VERTICAL_8X, horizontal=HORIZONTAL_FAST)
    # Donor slide 26 ("Up for the spike, out for the scale") is dropped: it is a
    # holdover, and its 1000x claim was the one number in the deck with no source.

    # Donor slides 26 (goals), 27 (closing question) and 28 (thank you) used to be
    # retexted in place. They are rebuilt on Valkey layouts instead, so there is
    # nothing to edit here and the originals are dropped in `order`.
    #
    # The "core execution" -> "command execution" wording from P99 survives in
    # slide_goals(), and the trailing `*` goes with it: the footnote it pointed at
    # was invented, and the next slide is the gloss.

    # --- new slides, appended then moved into place ---
    n_before = len(s)
    new = {
        "title": slide_title(prs),
        "bio": slide_bio(prs),
        "intro": slide_intro(prs),
        "story": slide_customer_latency(prs),
        "flame": slide_flamegraph(prs),
        "routing": slide_cluster_routing(prs),
        "getkeys": slide_getkeysinslot(prs, 17),
        "migbench": slide_migration_benchmark(prs),
        "operator": slide_operator(prs),
        "goals": slide_goals(prs),
        "involved": slide_get_involved(prs),
        "thanks": slide_thanks(prs),
    }
    idx = {name: n_before + i for i, name in enumerate(new)}

    # Donor indices are 0-based; the comments give the final 1-based position.
    # Dropped donor slides: 0 (title), 25 ("Up for the spike"), 26 (goals),
    # 27 (closing question), 28 (thank you). The four bookend slides are rebuilt on
    # Valkey layouts; 25 is a holdover with an unsourced 1000x claim.
    #
    # "Evaluating Valkey", the valkey-benchmark invocation, is gone: a wall of flags
    # is setup, not argument, and the deck's numbers are cited where they are used.
    order = [
        idx["title"],       # 1  Scaling up and out, and how Valkey does both
        idx["bio"],         # 2  Madelyn Olson (she/her)
        idx["intro"],       # 3  Introducing the Valkey project
        idx["story"],       # 4  A customer latency story
        1,                  # 5  Latency is a cliff
        2,                  # 6  Steady state is a lie
        3,                  # 7  You could overprovision
        4,                  # 8  Two axes, and neither one is optional
        idx["flame"],       # 9  Most of the work is not command execution
        5, 6, 7, 8, 9, 10, 11,   # 10-16 thread scalable command processing, 7 frames
        12,                 # 17 The host looks idle. The main thread is full.
        13,                 # 18 Latency is (still) a cliff though
        14,                 # 19 Adapt the requirements to the workload
        15,                 # 20 Up: about 8x, and then it stops
        16,                 # 21 Clustering
        idx["routing"],     # 22 Every key belongs to exactly one slot
        idx["getkeys"],     # 23 CLUSTER GETKEYSINSLOT
        17,                 # 24 MIGRATE
        18,                 # 25 CLUSTER SETSLOT STABLE
        19, 20, 21, 22,     # 26-29 atomic slot migration, 4 frames
        idx["migbench"],    # 30 Legacy vs atomic slot migration
        23,                 # 31 Out: now fast, and now safe
        24,                 # 32 CLUSTER SLOT-STATS
        idx["operator"],    # 33 Nobody should run these commands by hand
        idx["goals"],       # 34 Valkey's architectural goals
        idx["involved"],    # 35 Getting involved
        idx["thanks"],      # 36 Thank you! Let's connect.
    ]
    dropped = reorder_and_prune(prs, order)
    print(f"dropped {len(dropped)} slides: donor indices {dropped}")

    # --- everything onto the Valkey template ---
    # Order matters: move the slides first, then delete what nothing uses. Both
    # deletions refuse rather than corrupt if a slide is still on a stock layout, and
    # that refusal is the safety net.
    # The flame graph earns its place as evidence but not as airtime: it is one of the
    # denser slides and the phases it shows are restated by the build that follows.
    # Hidden rather than deleted, so it is one click away if a question needs it.
    hide_slide(prs.slides[8])

    strip_asides(prs)
    retarget_layouts(prs)
    drop_second_master(prs)
    gone = strip_stock_layouts(prs, "Valkey ")
    print(f"dropped {len(gone)} stock layouts: {gone}")
    left = sorted({s.slide_layout.name for s in prs.slides})
    if any(not n.startswith("Valkey ") for n in left):
        raise SystemExit(f"slides still on non-Valkey layouts: {left}")

    # --- fonts, last, over the finished deck ---
    # After every slide exists, so a slide added above cannot be missed.
    retypeface(prs, "Play", FONT)
    retypeface(prs, "Century Gothic", FONT)
    pin_font(prs)
    audit_fonts(prs)
    audit_paragraph_order(prs)
    audit_no_stray_prose(prs)
    audit_list_rows(prs, "Valkey List Rows", LAYOUT_LIST_ROWS_IMAGE)

    prs.save(tmp)
    Presentation(tmp)           # reopen: if it raises, PowerPoint would complain too
    os.replace(tmp, OUT)
    print(f"wrote {OUT}")
    verify(OUT)


def verify(path):
    prs = Presentation(path)
    print(f"verify: {len(prs.slides)} slides, {len(prs.slide_layouts)} layouts, "
          f"{prs.slide_width}x{prs.slide_height}")
    todos, fonts = [], set()
    for i, slide in enumerate(prs.slides, 1):
        for sh in shapes_flat(slide):
            if sh.has_text_frame and TODO in sh.text_frame.text:
                todos.append(i)
            for para in sh.text_frame.paragraphs if sh.has_text_frame else []:
                for run in para.runs:
                    fonts.add(run.font.name)
            try:
                if (sh.left is None or sh.top is None
                        or sh.left + sh.width > SLIDE_W or sh.top + sh.height > SLIDE_H
                        or sh.left < 0 or sh.top < 0):
                    print(f"  ! slide {i}: {sh.name!r} out of bounds "
                          f"({sh.left},{sh.top},{sh.width},{sh.height})")
            except TypeError:
                pass
        title = slide.shapes.title
        mark = "  [hidden]" if is_hidden(slide) else ""
        print(f"  {i}: {(title.text if title is not None else '(no title)')!r}{mark}")
    print(f"fonts on runs: {sorted(f for f in fonts if f)}")
    print(f"slides carrying {TODO}: {sorted(set(todos))}")


if __name__ == "__main__":
    main()
