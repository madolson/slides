#!/usr/bin/env python3
"""Generate icons/global-community.svg: a wireframe globe with linked nodes.

    python3 generation/make_globe_icon.py

Re-runnable. This icon is generated rather than hand-drawn because the grid lines have
to stop exactly where each dot sits, and those break points are not numbers a person
can type. Edit this script, never the SVG.

Why the grid is split instead of masked: every element in the icon library must be
`fill="none"` (so an icon works on any card colour), and `clipPath` is on the banned
list because PowerPoint renders it inconsistently. So a dot cannot hide what is behind
it. The equator and meridian are instead emitted as several subpaths that stop at each
dot's clearance circle and resume on the far side.

Both intersections are solved algebraically:

- **Equator**: a horizontal line against a circle. One square root.
- **Meridian**: an ellipse against a circle, which reduces to a quartic in `cos t`
  and is solved by its companion matrix via `numpy.roots`.

Multiple subpaths live in a single `d`, so the icon still has 7 elements and passes the
library validator's 3-8 element rule.

`verify()` re-samples the emitted geometry and asserts it never enters a dot. It parses
subpaths properly: splitting on coordinates alone samples straight across the gaps and
reports a correct icon as broken.
"""
import math, numpy as np

CX = CY = 48.0
R_GLOBE = 34.0
A, B = 19.0, 34.0            # meridian ellipse semi-axes
EQ_Y = 48.0
X0, X1 = CX - R_GLOBE, CX + R_GLOBE

# Three nodes, deliberately uneven: a symmetric arrangement reads as decoration rather
# than as a network. The right node sits on the equator (y = CY) so the line runs
# through it like a bead on a wire; the other two stay off both grid lines.
DOTS = [(31.0, 33.0), (64.0, 48.0), (37.0, 65.0)]
DOT_R, DOT_W = 4.0, 3.0
GRID_W, LINK_W, GLOBE_W = 2.0, 2.5, 4.0
# Where the grid stops: exactly on the dot's circle, r = DOT_R. No padding.
#
# This is the true intersection of the grid line with the dot's path, so the line's end
# cap lands on the ring's centreline and the two strokes overlap into a clean T joint.
# Adding the strokes' half-widths and a gap (an earlier version used 7.5) leaves visible
# air between the line and the ring, which reads as a mistake rather than as a join.
# PAD exists only to make the alternative explicit: PAD = DOT_W / 2 would stop the line
# at the ring's outer ink edge, touching without overlapping.
PAD = 0.0
CLEAR = DOT_R + PAD


def merge(iv):
    out = []
    for lo, hi in sorted(iv):
        if out and lo <= out[-1][1]:
            out[-1][1] = max(out[-1][1], hi)
        else:
            out.append([lo, hi])
    return out


def complement(iv, lo, hi):
    keep, x = [], lo
    for a, b in iv:
        if b <= lo or a >= hi:
            continue
        if a > x:
            keep.append((x, min(a, hi)))
        x = max(x, b)
    if x < hi:
        keep.append((x, hi))
    return [(a, b) for a, b in keep if b - a > 1e-9]


# --- equator: line y = EQ_Y meets circle |P - D| = CLEAR exactly ---------------
def equator_path():
    cut = []
    for dx, dy in DOTS:
        disc = CLEAR ** 2 - (EQ_Y - dy) ** 2
        if disc <= 0:
            continue
        h = math.sqrt(disc)                 # exact half-chord
        cut.append([dx - h, dx + h])
    segs = complement(merge(cut), X0, X1)
    return " ".join(f"M {a:.3f} {EQ_Y:.3f} L {b:.3f} {EQ_Y:.3f}" for a, b in segs)


# --- meridian: ellipse (CX + A cos t, CY + B sin t) meets each dot circle ------
def ellipse_dot_ts(dx, dy, r):
    """Parameters t where the ellipse meets the circle of radius r about (dx, dy).

    (A u + p)^2 + (B v + q)^2 = r^2 with u = cos t, v = sin t, u^2 + v^2 = 1.
    Eliminating v gives a quartic in u, solved by its companion matrix. Exact
    algebra; no sampling.
    """
    p, q = CX - dx, CY - dy
    K = p * p + q * q + B * B - r * r
    D = A * A - B * B
    ts = []
    if abs(q) < 1e-12:
        # v drops out: D u^2 + 2 A p u + K = 0, solve directly.
        for u in np.roots([D, 2 * A * p, K]) if D else ([-K / (2 * A * p)] if p else []):
            if abs(np.imag(u)) < 1e-9 and -1 <= np.real(u) <= 1:
                u = float(np.real(u))
                for v in (math.sqrt(max(0.0, 1 - u * u)), -math.sqrt(max(0.0, 1 - u * u))):
                    ts.append(math.atan2(v, u))
    else:
        # 2 B q v = -(D u^2 + 2 A p u + K)  ->  substitute into u^2 + v^2 = 1.
        c2, c1, c0 = D, 2 * A * p, K
        s = 2 * B * q
        # v = -(c2 u^2 + c1 u + c0)/s ; u^2 + v^2 - 1 = 0  ->  quartic
        quart = [c2 * c2,
                 2 * c2 * c1,
                 c1 * c1 + 2 * c2 * c0 + s * s,
                 2 * c1 * c0,
                 c0 * c0 - s * s]
        for u in np.roots(quart):
            if abs(np.imag(u)) > 1e-7:
                continue
            u = float(np.real(u))
            if not -1.0000001 <= u <= 1.0000001:
                continue
            u = max(-1.0, min(1.0, u))
            v = -(c2 * u * u + c1 * u + c0) / s
            if abs(u * u + v * v - 1) > 1e-6:
                continue
            ts.append(math.atan2(v, u))
    return ts


def meridian_path():
    cuts = []
    for dx, dy in DOTS:
        ts = sorted(t % (2 * math.pi) for t in ellipse_dot_ts(dx, dy, CLEAR))
        if len(ts) < 2:
            continue
        # Pair consecutive crossings and keep the spans that lie inside the dot.
        cand = list(zip(ts, ts[1:])) + [(ts[-1], ts[0] + 2 * math.pi)]
        for a, b in cand:
            m = (a + b) / 2
            px, py = CX + A * math.cos(m), CY + B * math.sin(m)
            if math.hypot(px - dx, py - dy) < CLEAR:
                cuts.append([a, b] if b <= 2 * math.pi else [a, 2 * math.pi])
                if b > 2 * math.pi:
                    cuts.append([0.0, b - 2 * math.pi])
    keep = complement(merge(cuts), 0.0, 2 * math.pi)
    if not keep:
        return ""
    out = []
    for a, b in keep:
        # Split spans over 180 deg so large-arc-flag stays 0 and stays unambiguous.
        n = max(1, math.ceil((b - a) / (math.pi * 0.98)))
        step = (b - a) / n
        x, y = CX + A * math.cos(a), CY + B * math.sin(a)
        d = f"M {x:.3f} {y:.3f}"
        for i in range(1, n + 1):
            t = a + step * i
            x, y = CX + A * math.cos(t), CY + B * math.sin(t)
            d += f" A {A} {B} 0 0 1 {x:.3f} {y:.3f}"
        out.append(d)
    return " ".join(out)


# --- links -----------------------------------------------------------------------
# LINK_GAP is where a link stops, measured from the dot centre. DOT_R makes it flush
# with the ring, matching the grid. A larger value leaves air, which is inconsistent
# with a grid that joins the ring, and reads as a mistake rather than as a join.
LINK_GAP = DOT_R


def links_path(gap=LINK_GAP):
    out = []
    for p, q in zip(DOTS, DOTS[1:]):
        dx, dy = q[0] - p[0], q[1] - p[1]
        L = math.hypot(dx, dy)
        ux, uy = dx / L, dy / L
        out.append(f"M {p[0]+ux*gap:.3f} {p[1]+uy*gap:.3f} "
                   f"L {q[0]-ux*gap:.3f} {q[1]-uy*gap:.3f}")
    return " ".join(out)


def el(body, w, tag="path"):
    return (f'  <{tag} {body} fill="none" stroke="#6983FF" stroke-width="{w}" '
            f'stroke-linecap="round" stroke-linejoin="round"/>')


def build():
    els = [el(f'cx="{CX:g}" cy="{CY:g}" r="{R_GLOBE:g}"', GLOBE_W, "circle"),
           el(f'd="{equator_path()}"', GRID_W),
           el(f'd="{meridian_path()}"', GRID_W),
           el(f'd="{links_path()}"', LINK_W)]
    els += [el(f'cx="{x:g}" cy="{y:g}" r="{DOT_R:g}"', DOT_W, "circle") for x, y in DOTS]
    return ('<svg xmlns="http://www.w3.org/2000/svg" width="96" height="96" '
            'viewBox="0 0 96 96">\n' + "\n".join(els) + "\n</svg>\n")


def verify():
    """Re-sample the emitted geometry; assert no grid ink enters a dot's clearance.

    Split on `M` first. A checker that pulls every coordinate pair out of the path and
    joins them in order samples straight through the gaps, and then reports this exact
    icon as broken.
    """
    import re

    def line_pts(d):
        for sub in (s for s in d.split("M") if s.strip()):
            n = [float(v) for v in re.findall(r"-?[\d.]+", sub)]
            for i in range(2001):
                t = i / 2000
                yield n[0] + (n[2] - n[0]) * t, n[1] + (n[3] - n[1]) * t

    def arc_pts(d):
        for sub in (s for s in d.split("M") if s.strip()):
            n = [float(v) for v in re.findall(r"-?[\d.]+", sub)]
            pts = [(n[0], n[1])] + [(n[2:][i + 5], n[2:][i + 6])
                                    for i in range(0, len(n[2:]), 7)]
            for (x0, y0), (x1, y1) in zip(pts, pts[1:]):
                t0 = math.atan2((y0 - CY) / B, (x0 - CX) / A)
                t1 = math.atan2((y1 - CY) / B, (x1 - CX) / A)
                if t1 < t0:
                    t1 += 2 * math.pi
                for i in range(1001):
                    t = t0 + (t1 - t0) * i / 1000
                    yield CX + A * math.cos(t), CY + B * math.sin(t)

    worst = min(math.hypot(x - dx, y - dy)
                for gen in (line_pts(equator_path()), arc_pts(meridian_path()))
                for x, y in gen for dx, dy in DOTS)
    assert worst >= CLEAR - 1e-3, f"grid enters a dot: {worst:.4f} < {CLEAR}"
    print(f"  verified: closest grid ink to a dot centre is {worst:.4f} "
          f"(clearance {CLEAR})")


if __name__ == "__main__":
    import os
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "icons", "global-community.svg")
    svg = build()
    open(out, "w").write(svg)
    n = svg.count("<circle") + svg.count("<path")
    print(f"wrote {out}: {n} elements, "
          f"{equator_path().count('M')} equator + {meridian_path().count('M')} "
          f"meridian subpaths")
    assert 3 <= n <= 8, f"{n} elements, library allows 3-8"
    verify()
