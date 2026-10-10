#!/usr/bin/env python3
"""Derive the title slide's band artwork from a valkey-svgs banner.

Run this once to produce `assets/band-clustering.jpg`; `layout_title_band()` in
`build_placeholder_deck.py` then reads that file. The output is committed, because
`~/git/valkey-svgs` is a separate repo and is not a build dependency of this one. Re-run
it to pick a different motif or to pull in a regenerated banner.

    python3 make_band_art.py                  # clustering, the default
    python3 make_band_art.py performance      # any theme name in valkey-svgs/svg/

**The banners are not flourishes,** and that is the thing to know before reaching for
one. Each is a whole 1920x1080 picture: a navy-to-purple gradient, a vignette, film
grain, a motif, its own copy of the Valkey mark at the motif's focal point, plus a corner
lockup and a caption. Two consequences:

1. **The motif cannot be lifted off its background.** Stripping the gradient and keeping
   the linework, to sit on the flat `ACCENT_1` band, does not work: the motifs carry
   blurred dark glow layers drawn to bloom against navy, and on Valkey blue those read as
   smudges. So the gradient comes along, and the band stops being flat blue. That is a
   real design change to the title slide, not a decoration.
2. **The banner's own mark has to go,** or it collides with the layout's white lockup.

What *is* separable is the layering, because the generator emits it in a fixed order, so
this cuts by position rather than by guessing:

    <defs>                          keep: the motif references its gradients and filters
    <rect fill="url(#sky)"/>        keep: the ground the motif was drawn against
    ...motif groups...              keep
    <rect fill="url(#vignette)"/>   keep
    <rect filter="url(#grain)"/>    keep, and drop every sibling after it: the chrome
    <g>lockup</g> <g>caption</g>    drop

Output is a raster, not an SVG. These banners lean on `feGaussianBlur` and `feTurbulence`,
and PowerPoint renders SVG filters inconsistently or not at all, so `add_svg` would be a
coin flip on the projector. Nothing here is linework that needs to scale.

JPEG, not PNG. The artwork is a smooth gradient with soft glows, which is the case PNG is
worst at: the same image is 1.6 MB as PNG and 139 KB at quality 92, with no banding at
band size. It goes in the *layout*, so the cost is paid once per deck rather than per
slide, but 1.6 MB tripled the placeholder deck on its own.
"""
import os
import re
import subprocess
import sys
import tempfile

from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")
SVGS = os.path.expanduser("~/git/valkey-svgs")
CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

SVG_NS = "http://www.w3.org/2000/svg"
BAND_W, BAND_H = 4114800, 6858000               # from layout_title_band()
OUT_W = 1100                                    # px; 2.4x the band's on-screen size


def _q(tag):
    return f"{{{SVG_NS}}}{tag}"


def _logo_geometry():
    """The mark's path data and viewBox, as the banner generator reads them."""
    src = open(os.path.join(SVGS, "assets/Valkey-logo.svg")).read()
    d = re.search(r'\sd="([^"]+)"', src).group(1)
    vb = [float(v) for v in
          re.search(r'viewBox="([-\d.\s]+)"', src).group(1).split()]
    return d, vb[2], vb[3]


def _mark_centre(root, mark_d, vw, vh):
    """Where the banner puts its own copy of the mark, in viewBox units.

    `mark()` in generate.mjs corners the glyph and scales it, so the centre is the
    translate plus half the scaled box. A blurred glow copy sits at the same transform,
    so the first hit is enough.
    """
    for g in root.iter(_q("g")):
        if not any(p.get("d") == mark_d for p in g.findall(_q("path"))):
            continue
        m = re.match(r"translate\(([-\d.]+) ([-\d.]+)\) scale\(([-\d.]+)\)",
                     g.get("transform", ""))
        if m:
            x, y, s = (float(v) for v in m.groups())
            return x + vw * s / 2, y + vh * s / 2
    return 960.0, 540.0


def band_svg(name):
    """Banner `name`, chrome and mark removed, cropped to the band's aspect ratio."""
    path = os.path.join(SVGS, "svg", f"{name}.svg")
    if not os.path.exists(path):
        raise SystemExit(f"no such banner: {path}")
    root = etree.parse(path).getroot()
    mark_d, vw, vh = _logo_geometry()
    cx, _ = _mark_centre(root, mark_d, vw, vh)

    kids = list(root)
    grain = next((i for i, c in enumerate(kids)
                  if c.tag == _q("rect") and "grain" in (c.get("filter") or "")), None)
    if grain is None:
        raise SystemExit(f"{name}: no grain rect, so the chrome cannot be located")
    for child in kids[grain + 1:]:
        root.remove(child)

    marks = [g for g in root.iter(_q("g"))
             if any(p.get("d") == mark_d for p in g.findall(_q("path")))]
    if not marks:
        raise SystemExit(f"{name}: no mark group found; did the logo path change?")
    for g in marks:
        g.getparent().remove(g)

    # Full artwork height, centred horizontally on where the mark was so the motif's
    # focal point lands on the layout's lockup, and clamped so the crop stays inside.
    h = 1080.0
    w = h * BAND_W / BAND_H
    x = min(max(cx - w / 2, 0.0), 1920.0 - w)
    root.set("viewBox", f"{x:.1f} 0 {w:.1f} {h:.1f}")
    for attr in ("width", "height"):
        root.attrib.pop(attr, None)
    return etree.tostring(root).decode(), len(marks)


def rasterize(svg_text, jpg_path, width=OUT_W, quality=92):
    """Render `svg_text` to an opaque JPEG at the band's aspect ratio.

    Goes through an HTML wrapper with a CSS-sized <img>, for the reason `rasterize()` in
    build_placeholder_deck.py documents: Chrome pointed at a bare .svg renders it at its
    natural size and corner-anchors the result.
    """
    from PIL import Image

    height = round(width * BAND_H / BAND_W)
    with tempfile.TemporaryDirectory() as tmp:
        svg = os.path.join(tmp, "band.svg")
        with open(svg, "w") as fh:
            fh.write(svg_text)
        wrapper = os.path.join(tmp, "w.html")
        with open(wrapper, "w") as fh:
            fh.write(f'<html><body style="margin:0">'
                     f'<img src="file://{svg}" '
                     f'style="width:{width}px;height:{height}px;display:block">'
                     f'</body></html>')
        shot = os.path.join(tmp, "band.png")
        subprocess.run(
            [CHROME, "--headless", "--disable-gpu", "--no-sandbox",
             "--virtual-time-budget=4000", f"--screenshot={shot}",
             f"--window-size={width},{height}", f"file://{wrapper}"],
            check=True, capture_output=True,
        )
        im = Image.open(shot).convert("RGB")
        if im.size != (width, height):
            raise SystemExit(f"{shot}: got {im.size}, wanted {(width, height)}")
        # These banners are edge-to-edge artwork, so a nearly uniform render means the
        # gradient drew and the motif did not. A rich histogram does not prove a motif,
        # but a flat one disproves it.
        if len(im.getcolors(maxcolors=1 << 20) or [1]) < 500:
            raise SystemExit(f"{shot}: too few distinct colours; motif did not render")
        # subsampling=0 keeps the ring's saturated cyan and magenta edges sharp; the
        # default halves chroma resolution and softens exactly those.
        im.save(jpg_path, "JPEG", quality=quality, optimize=True, subsampling=0)
    return im.size


if __name__ == "__main__":
    name = sys.argv[1] if len(sys.argv) > 1 else "clustering"
    svg_text, n_marks = band_svg(name)
    out = os.path.join(ASSETS, f"band-{name}.jpg")
    size = rasterize(svg_text, out)
    print(f"{out}  {size[0]}x{size[1]}  {os.path.getsize(out) // 1024} KB  "
          f"(stripped {n_marks} mark groups and the chrome)")
