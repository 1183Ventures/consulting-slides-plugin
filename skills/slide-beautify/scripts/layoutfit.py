"""Text boxes sized to their text, and pages that read full and even.

A port of SlideShortcuts' src/templates/layoutfit.js (lavender main 70839fd8),
run on a library page once the deck's words are in it, so the rebuilt page is
fitted and spaced the way our own library pages are. The rule is "fit, then
fill":

1. FIT. Every plain text box (no fill, no line, upright, a rectangle) takes
   the height its text needs at its size, with Resize Shape to Fit Text on.
   The edge its text is anchored to stays, so the text does not move.
2. (No stacking: fitting never moves a text, so a box is never pushed.)
3. ARRANGE, so the content uses the body (from under the title to above the
   source line):
   - a row of cards that ends the page reaches the body's bottom; text stays
     at the top, and what sat at a card's foot stays pinned to its bottom;
     cards in a grid keep their size;
   - three or more like boxes, one per row, that end the page are spread
     down with equal gaps between their texts (never closer than drawn); each
     row's label, rule and number move with it.
4. FILL. While the page has room, its body text (9 to 13.5 pt) steps up half
   a point at a time (never above MAX_FILL) and steps 1 to 3 run again. The
   largest size at which everything still fits is kept: no text past the
   body's bottom, out of its card, onto a rule, into a filled shape or into
   another text; no label outside a column of rows gains a line; lined-up rows
   stay lined up; no column of rows gets closer; no line runs sideways into a
   shape beside it; body text stays half a point under the smallest text the
   fill leaves alone; and a page's measured cap (layoutfit-caps.json) holds.

Never touched: placeholders, turned shapes, text in a coloured shape, titles
and big figures (above 13.5 pt), a page with fewer than two text boxes. A page
with a table, group or picture is fitted only (nothing is spread).

Shapes are the page's descriptions, in inches (pagexml.py). Text is measured
with fixed width tables (assets/font-widths.json, the tables the library's
own tests use), so this and the library agree; tools/test_beautify.py holds
it to the JS on recorded pages (tools/vectors-layoutfit.json). Standard
library only.
"""

import copy
import json
import os
import re
from decimal import Decimal, ROUND_HALF_UP

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "..", "assets")

TOL = 0.02          # inches: edges this close are the same edge
MIN_SLACK = 0.12    # inches: a box this much off its text is fitted
CARD_GEOMS = ("rect", "roundRect")
NEAR = 0.35         # inches: cards this close form a grid
FOOT = 0.35         # inches: a shape this close to a card's bottom is its foot
MAX_FILL = 14
FILL_FROM = 9
FILL_TO = 13.5
MAX_GAP = 0.9       # inches: rows are never spread further apart than this
NARROW = 1.2        # inches: a text box narrower than this is never shortened
SAFETY = 1.06
SAFETY_ADD = 0.04
MIN_LINE = 1.2
SIDE_GAP = 0.08
BODY = {"top": 1.62, "bottom": 6.58}
DEFAULT_INSET = [0.08, 0.05, 0.08, 0.05]
DEFAULT_FONTS = {"heading": "Georgia", "body": "Arial"}
INF = float("inf")


def _load_caps():
    try:
        with open(os.path.join(ASSETS, "layoutfit-caps.json"), encoding="utf-8") as fh:
            return dict(json.load(fh).get("caps") or {})
    except (OSError, ValueError):
        return {}


# The fill's caps the renderer measured (lavender tools/library-overflow.mjs), by page id
CAPS = _load_caps()


class Measure:
    """Widths from the fixed tables; a face not in them is measured as Arial."""

    def __init__(self, path=os.path.join(ASSETS, "font-widths.json")):
        with open(path, encoding="utf-8") as fh:
            t = json.load(fh)
        self.index = {ch: i for i, ch in enumerate(t["chars"])}
        self.n = self.index["n"]
        self.fonts = t["fonts"]

    def face(self, font):
        return self.fonts.get(font) or self.fonts["Arial"]

    def width(self, text, font, size, bold):
        f = self.face(font)
        w = f["bold"] if bold else f["regular"]
        return sum(w[self.index.get(ch, self.n)] for ch in str(text)) / 1000.0 * size

    def line_height(self, font):
        return self.face(font).get("lineHeight") or 1.2


class SlideMeasure(Measure):
    """Widths as a deck of another shape will draw them, in the page's inches:
    each text measured at the size the slide carries (`slide_pt`, the writer's
    own rounding and floor) over the width scale `sx`, widths and lines times
    `margin` for what one model leaves out (insets and indents scale with the
    fonts, sizes round up to half points)."""

    def __init__(self, base, slide_pt, sx, margin=1.0, line_floor=0):
        self.index, self.n, self.fonts = base.index, base.n, base.fonts
        self.base = base
        self.slide_pt, self.sx, self.margin, self.line_floor = slide_pt, sx, margin, line_floor

    def face(self, font):
        return self.base.face(font)

    def width(self, text, font, size, bold):
        return self.base.width(text, font, self.slide_pt(size), bold) / self.sx * self.margin

    def line_height(self, font):
        # the margin covers heights too: Place.pt rounds a size up by as much as a quarter point
        return max(self.base.line_height(font), self.line_floor) * self.margin


class _O:
    """The pass's options: measure, fonts, body box, cap."""

    def __init__(self, measure, fonts, body, max_raise, safety_add=SAFETY_ADD, margin=True):
        self.margin = margin
        self.m = measure
        self.fonts = fonts
        self.body = body
        self.max_raise = max_raise
        self.safety_add = safety_add


def to_fixed(x, digits):
    """JS Number.prototype.toFixed: halves round away from zero, on the exact value."""
    q = Decimal(1).scaleb(-digits)
    return str(Decimal(x).quantize(q, rounding=ROUND_HALF_UP))


def round_half_up(x):
    """JS Math.round: halves go up (toward +infinity)."""
    import math
    return math.floor(x + 0.5)


def _num(x, digits):
    return float(to_fixed(x, digits))


# JS's \s and String.prototype.trim: these characters and no others (Python's
# own \s adds U+001C to U+001F and U+0085 and leaves out U+FEFF)
JS_SPACE = "\t\n\v\f\r \u00a0\u1680\u2000-\u200a\u2028\u2029\u202f\u205f\u3000\ufeff"
_TRIM = re.compile("^[%s]+|[%s]+$" % (JS_SPACE, JS_SPACE))


def js_truthy(v):
    """JS truthiness: an empty {} or [] is true, 0, NaN, "" and null are not."""
    if v is None or v is False:
        return False
    if isinstance(v, (int, float)):
        return v == v and v != 0
    if isinstance(v, str):
        return v != ""
    return True


def js_str(v):
    """JS String(v) for what a run's text can hold."""
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, float):
        if v != v:
            return "NaN"
        if v in (float("inf"), float("-inf")):
            return "Infinity" if v > 0 else "-Infinity"
        if v.is_integer() and abs(v) < 1e21:
            return str(int(v))
        return repr(v)
    return str(v)


# ---- shapes -----------------------------------------------------------------

def bottom(s):
    return s["y"] + s["h"]


def right(s):
    return s["x"] + s["w"]


def _paras(s):
    return (s.get("text") or {}).get("paras") or []


def has_text(s):
    return any(js_truthy(r.get("t")) and _TRIM.sub("", js_str(r.get("t"))) for p in _paras(s) for r in p.get("runs") or [])


def is_text_box(s):
    t = s.get("text") or {}
    return (s.get("kind") == "sp" and (s.get("geom") or "rect") == "rect" and not js_truthy(s.get("rot"))
            and not js_truthy(s.get("fill")) and not js_truthy(s.get("line")) and has_text(s) and not js_truthy(t.get("vert")))


def is_card(s):
    return (s.get("kind") == "sp" and (s.get("geom") or "rect") in CARD_GEOMS and not js_truthy(s.get("rot"))
            and js_truthy(s.get("fill")) and not has_text(s) and s["w"] > 0.6 and s["h"] > 0.6)


def inside(a, c):
    return (a is not c and a["x"] >= c["x"] - TOL and a["y"] >= c["y"] - TOL
            and right(a) <= right(c) + TOL and bottom(a) <= bottom(c) + TOL)


def overlap_x(a, b):
    return min(right(a), right(b)) - max(a["x"], b["x"])


def span_y(s):
    if s.get("kind") == "cxn":
        return [min(s["y1"], s["y2"]), max(s["y1"], s["y2"])]
    return [s["y"], bottom(s)]


def span_x(s):
    if s.get("kind") == "cxn":
        return [min(s["x1"], s["x2"]), max(s["x1"], s["x2"])]
    return [s["x"], right(s)]


def _shift(s, d):
    if s.get("kind") == "cxn":
        s["y1"] += d
        s["y2"] += d
    else:
        s["y"] += d


def _has(items, s):
    return any(x is s for x in items)


def _index(items, s):
    for i, x in enumerate(items):
        if x is s:
            return i
    return -1


# ---- text -------------------------------------------------------------------

def _runs(p):
    return [r for r in p.get("runs") or [] if r.get("t") is not None]


def _size(runs):
    for r in runs:
        if r.get("sz"):
            return r["sz"]
    return 12


def _font(runs, o):
    return o.fonts["heading"] if any(r.get("font") == "head" for r in runs) else o.fonts["body"]


def _bold(runs):
    return len(runs) > 0 and all(r.get("b") for r in runs)


def _indent(p):
    return (p.get("level") or 0) * 0.2 + (0.17 if js_truthy(p.get("bullet")) else 0)


_BREAK = re.compile(r"\r?\n|\v")
_WORDS = re.compile("([%s]+)" % JS_SPACE)
_SPACE = re.compile("^[%s]+$" % JS_SPACE)


# A line fits only when it leaves half a percent of its room or one point,
# whichever is larger (layoutfit.js fitWidth; CEO, 14 Nov 2026)
WRAP_SPARE = 0.005
WRAP_SPARE_PT = 1


def fit_width(width, o=None):
    """The width a line may fill when sizing its box (no margin when o.margin is False)."""
    if o is not None and getattr(o, "margin", True) is False:
        return width
    return width - max(WRAP_SPARE * width, WRAP_SPARE_PT)


def _wrap(text, width, font, size, bold, o, fit=True):
    """Lines a paragraph's text wraps to: greedy, at spaces, as PowerPoint does.
    fit: with the wrap margin (sizing a box); without, the lines as drawn."""
    if fit:
        width = fit_width(width, o)
    lines = 0
    for para in _BREAK.split(text):
        line = 0
        lines += 1
        for word in _WORDS.split(para):
            if not word:
                continue
            w = o.m.width(word, font, size, bold)
            if line + w > width and line > 0 and not _SPACE.match(word):
                lines += 1
                line = w
            else:
                line += w
    return lines


def text_height(s, o, raw=False):
    """The height, in inches, the text of `s` needs at its sizes, with its
    insets; `raw`: the text alone, without the fit's safety room."""
    t = s.get("text") or {}
    ins = t.get("inset") or DEFAULT_INSET
    room = max((s["w"] - ins[0] - ins[2]) * 72, 4)
    pts = 0
    for p in t.get("paras") or []:
        runs = _runs(p)
        size = _size(runs)
        font = _font(runs, o)
        bold = _bold(runs)
        width = max(room - _indent(p) * 72, 4)
        text = "".join(js_str(r["t"]) for r in runs)
        lines = 1 if t.get("wrap") is False else _wrap(text, width, font, size, bold, o)
        spacing = p["spacing"] / 100 if p.get("spacing") else 1
        # never under 1.2 of the size, where a text sits too (layoutfit.js)
        pts += lines * size * max(o.m.line_height(font), MIN_LINE) * spacing
        pts += (p.get("after") or 0) + (p.get("before") or 0)
    if raw:
        return pts / 72 + ins[1] + ins[3]
    return (pts / 72) * SAFETY + o.safety_add + ins[1] + ins[3]


def lines_of(s, o):
    """Lines the text wraps to, all paragraphs (for the fill's guard)."""
    t = s.get("text") or {}
    ins = t.get("inset") or DEFAULT_INSET
    room = max((s["w"] - ins[0] - ins[2]) * 72, 4)
    lines = 0
    for p in t.get("paras") or []:
        runs = _runs(p)
        if t.get("wrap") is False:
            lines += 1
            continue
        width = max(room - _indent(p) * 72, 4)
        lines += _wrap("".join(js_str(r["t"]) for r in runs), width, _font(runs, o), _size(runs), _bold(runs), o, fit=False)
    return lines


def text_top(s, o):
    """Where the text starts in its box, by its anchor."""
    need = text_height(s, o, True)
    a = (s.get("text") or {}).get("anchor") or "t"
    if a == "ctr":
        return s["y"] + (s["h"] - need) / 2
    if a == "b":
        return s["y"] + s["h"] - need
    return s["y"]


def text_bottom(s, o):
    return text_top(s, o) + text_height(s, o, True)


def _ink_spans(s, o):
    """One-line paragraphs' ink, [left, right] in inches."""
    t = s.get("text") or {}
    ins = t.get("inset") or DEFAULT_INSET
    inner = s["w"] - ins[0] - ins[2]
    out = []
    for p in t.get("paras") or []:
        runs = _runs(p)
        if not runs:
            continue
        w = o.m.width("".join(js_str(r["t"]) for r in runs), _font(runs, o), _size(runs),
                      all(r.get("b") for r in runs)) / 72 + _indent(p)
        if w > inner + 0.01 and t.get("wrap") is not False:
            continue
        x0 = s["x"] + ins[0]
        x1 = s["x"] + s["w"] - ins[2]
        align = p.get("align")
        left = x1 - w if align == "r" else (x0 + x1) / 2 - w / 2 if align == "ctr" else x0
        out.append((left, left + w))
    return out


def sideways_depths(shapes, o):
    """Each one-line paragraph within its gap of a filled shape beside it,
    keyed (text, paragraph, shape), with how far it reaches into that gap."""
    solids = [(f, k) for k, f in enumerate(shapes) if f.get("kind") == "sp" and js_truthy(f.get("fill"))
              and not js_truthy(f.get("rot")) and not is_text_box(f)]
    out = {}
    for i, s in enumerate(shapes):
        if not is_text_box(s):
            continue
        top = text_top(s, o)
        bot = top + text_height(s, o, True)
        spans = _ink_spans(s, o)
        for f, k in solids:
            if inside(s, f) or f["y"] >= bot - 0.02 or bottom(f) <= top + 0.02:
                continue
            for j, (left, rgt) in enumerate(spans):
                gap = max(SIDE_GAP, 0.03 * (rgt - left))
                depth = None
                if f["x"] >= left and rgt > f["x"] - gap:
                    depth = rgt - (f["x"] - gap)
                elif right(f) <= rgt and left < right(f) + gap:
                    depth = right(f) + gap - left
                if depth is not None:
                    out[(i, j, k)] = depth
    return out


def sideways_hits(shapes, o):
    """One-line paragraphs whose ink comes within SIDE_GAP (3% of a long
    line's width) of a filled shape beside them."""
    solids = [f for f in shapes if f.get("kind") == "sp" and js_truthy(f.get("fill")) and not js_truthy(f.get("rot")) and not is_text_box(f)]
    n = 0
    for s in shapes:
        if not is_text_box(s):
            continue
        top = text_top(s, o)
        bot = top + text_height(s, o, True)
        spans = None
        for f in solids:
            if inside(s, f) or f["y"] >= bot - 0.02 or bottom(f) <= top + 0.02:
                continue
            if spans is None:
                spans = _ink_spans(s, o)
            for left, rgt in spans:
                gap = max(SIDE_GAP, 0.03 * (rgt - left))
                if f["x"] >= left and rgt > f["x"] - gap:
                    n += 1
                elif right(f) <= rgt and left < right(f) + gap:
                    n += 1
    return n


def _in_row_column(s, shapes):
    return sum(1 for b in shapes if is_text_box(b) and abs(b["x"] - s["x"]) < TOL and abs(b["w"] - s["w"]) < TOL) >= 3


# ---- 1. fit -----------------------------------------------------------------

def _fit_boxes(shapes, o, changed):
    for s in shapes:
        if not is_text_box(s):
            continue
        need = text_height(s, o)
        if abs(s["h"] - need) < MIN_SLACK and need <= s["h"]:
            continue
        # a narrow label keeps a spare line: its line breaks are the estimate's least sure
        if s["w"] < NARROW:
            paras = _paras(s)
            sz = _size(paras[0].get("runs") or []) if paras else 12
            need = min(max(need, s["h"]), need + (sz * MIN_LINE) / 72)
        if abs(s["h"] - need) < MIN_SLACK and need <= s["h"]:
            continue
        a = s["text"].get("anchor") or "t"
        old = s["h"]
        if a == "ctr":
            s["y"] += (s["h"] - need) / 2
        elif a == "b":
            s["y"] += s["h"] - need
        s["h"] = need
        if s["text"].get("wrap") is not False:
            s["text"]["autofit"] = "resize"
        changed.append({"name": s.get("name") or "text box", "what": "height to its text",
                        "from": _num(old, 3), "to": _num(need, 3)})


# ---- 3. arrange -------------------------------------------------------------

def _in_grid(c, cards):
    for d in cards:
        if d is c:
            continue
        across = overlap_x(c, d) > 0
        below = d["y"] >= bottom(c) - TOL and d["y"] - bottom(c) < NEAR
        above = c["y"] >= bottom(d) - TOL and c["y"] - bottom(d) < NEAR
        if across and (below or above):
            return True
        down = min(bottom(c), bottom(d)) - max(c["y"], d["y"]) > 0
        beside = ((d["x"] >= right(c) - TOL and d["x"] - right(c) < NEAR)
                  or (c["x"] >= right(d) - TOL and c["x"] - right(d) < NEAR))
        if down and beside and abs(d["h"] - c["h"]) > TOL:
            return True
    return False


def _ends_page(shapes, y, x0, x1, except_ids):
    """Nothing else of the page below `y` across [x0, x1]: the block ends the page."""
    for s in shapes:
        if id(s) in except_ids or s.get("kind") == "ph":
            continue
        sy = span_y(s)[0]
        sx0, sx1 = span_x(s)
        if sy >= y - TOL and sx1 > x0 + TOL and sx0 < x1 - TOL:
            return False
    return True


def _stretch_cards(shapes, o, changed):
    if not o.body:
        return
    every = [s for s in shapes if is_card(s)]
    cards = [c for c in every if not _in_grid(c, every)]
    done = set()
    for c in cards:
        if id(c) in done:
            continue
        row = [d for d in cards if abs(d["y"] - c["y"]) < TOL and abs(d["h"] - c["h"]) < TOL]
        done.update(id(d) for d in row)
        # a card alone is stretched only when it is a panel, not a node in a diagram
        if len(row) < 2 and c["h"] < 0.5 * (o.body["bottom"] - o.body["top"]):
            continue
        row_bottom = bottom(c)

        def kid(s, card):
            if s is card:
                return False
            if s.get("kind") == "cxn":
                sy, sx = span_y(s), span_x(s)
                return (sy[0] >= card["y"] - TOL and sy[1] <= bottom(card) + TOL
                        and sx[0] >= card["x"] - TOL and sx[1] <= right(card) + TOL)
            return inside(s, card)

        kids = [s for s in shapes if any(kid(s, card) for card in row)]
        kid_ids = {id(s) for s in kids}
        left = min(x["x"] for x in row)
        right_edge = max(right(x) for x in row)
        row_ids = {id(x) for x in row}
        # what sits between the row's cards (an arrow) belongs to the row too
        between = [s for s in shapes if s.get("kind") == "sp" and id(s) not in row_ids and id(s) not in kid_ids
                   and s["y"] >= c["y"] - TOL and bottom(s) <= row_bottom + TOL
                   and s["x"] >= left - TOL and right(s) <= right_edge + TOL]
        except_ids = row_ids | kid_ids | {id(s) for s in between}
        if not _ends_page(shapes, row_bottom - TOL, left, right_edge, except_ids):
            continue
        # a bottom edge the page lines up with shapes beside the row stays
        lined = any(id(s) not in except_ids and s.get("kind") == "sp" and abs(bottom(s) - row_bottom) < 0.06
                    and (right(s) <= left + TOL or s["x"] >= right_edge - TOL) for s in shapes)
        if lined:
            continue
        d = o.body["bottom"] - row_bottom
        if abs(d) < TOL:
            continue
        for card in row:
            # the card's foot: what sits under its one clear gap (0.3" or more)
            # and ends near its bottom; a card without such a gap has no foot
            own = [s for s in kids if ((span_x(s)[0] >= card["x"] - TOL and span_x(s)[1] <= right(card) + TOL)
                                       if s.get("kind") == "cxn" else inside(s, card))]
            spans = sorted(([span_y(s)[0], span_y(s)[1], s] for s in own), key=lambda v: v[0])
            reach = spans[0][1] if spans else 0
            cut = None
            widest = 0.3
            for t0, t1, _ in spans[1:]:
                if t0 - reach >= widest:
                    widest = t0 - reach
                    cut = t0
                reach = max(reach, t1)
            feet = [] if cut is None else [s for s in own if span_y(s)[0] >= cut - TOL and row_bottom - span_y(s)[1] < FOOT]
            for s in feet:
                _shift(s, d)
            card["h"] += d
            changed.append({"name": card.get("name") or "card", "what": "card reaches the body's bottom",
                            "from": _num(card["h"] - d, 3), "to": _num(card["h"], 3)})
        for s in between:
            _shift(s, d / 2)


def _distribute_rows(shapes, o, changed):
    boxes = [s for s in shapes if is_text_box(s)]
    cols = {}
    for s in boxes:
        runs = (s["text"]["paras"][0].get("runs") or []) if s["text"].get("paras") else []
        k = to_fixed(s["x"], 2) + "|" + to_fixed(s["w"], 2) + "|" + repr(float(_size(runs)))
        cols.setdefault(k, []).append(s)
    cards = [s for s in shapes if is_card(s)]
    moved = set()
    # the column with the most text first: its rows carry the other columns
    order = sorted(cols.values(), key=lambda lst: (-len(lst), -sum(s["h"] for s in lst)))
    for lst in order:
        if len(lst) < 3 or any(id(s) in moved for s in lst):
            continue
        lst.sort(key=lambda s: s["y"])
        pitches = [lst[i + 1]["origY"] - lst[i]["origY"] for i in range(len(lst) - 1)]
        if max(pitches) - min(pitches) > 0.05 or min(pitches) < 0.2:
            continue
        if any(sum(1 for s in lst if inside(s, c)) > 1 for c in cards):
            continue
        hs = [text_height(s, o, True) for s in lst]
        tops = [text_top(s, o) for s in lst]
        last = lst[-1]
        p0 = (last["origY"] - lst[0]["origY"]) / (len(lst) - 1)
        band_top = lst[0]["origY"] - p0 / 2
        band_bottom = last["origY"] + p0 / 2 + (bottom(last) - last["y"])
        lst_ids = {id(s) for s in lst}
        # the rows' companions: shapes inside the rows' band, none spanning rows
        members = []
        for s in shapes:
            if s.get("kind") == "ph" or id(s) in lst_ids:
                continue
            sy, sb = span_y(s)
            if sy >= band_top - 0.3 and sb <= band_bottom + 0.1 and sb - sy <= 1.2 * p0:
                members.append(s)
        groups = {}
        for s in members:
            sx0, sx1 = span_x(s)
            groups.setdefault(s.get("kind") + "|" + to_fixed(sx0, 1) + "|" + to_fixed(sx1, 1), []).append(s)

        def top_of(s):
            if s.get("kind") == "cxn":
                return span_y(s)[0]
            return s["origY"] if s.get("origY") is not None else s["y"]

        row_index = {}
        orderly = True
        for g in groups.values():
            # rules between the rows (one fewer than the rows) go with the row under them
            between = len(g) == len(lst) - 1 and all(s.get("kind") == "cxn" for s in g)
            if len(g) != len(lst) and not between:
                orderly = False
                break
            g.sort(key=top_of)
            first = 1 if between else 0
            offs = [top_of(s) - lst[j + first]["origY"] for j, s in enumerate(g)]
            if max(offs) - min(offs) > 0.06:
                orderly = False
                break
            for j, s in enumerate(g):
                row_index[id(s)] = j + first
        if not orderly:
            continue
        # rows tied to a graphic beside them stay as drawn
        member_ids = {id(s) for s in members}
        band_lo = lst[0]["origY"]
        band_hi = bottom(last)
        tied = False
        for s in shapes:
            if (s.get("kind") == "ph" and s.get("type") != "pic") or id(s) in lst_ids or id(s) in member_ids:
                continue
            sy, sb = span_y(s)
            if min(sb, band_hi) - max(sy, band_lo) > 1.2 * p0:
                tied = True
                break
        if tied:
            continue
        last_bottom = tops[-1] + hs[-1]
        if not o.body:
            continue
        ends = _ends_page(shapes,
                          max([last_bottom] + [span_y(s)[1] for s in members]) - TOL,
                          min([s["x"] for s in lst] + [span_x(s)[0] for s in members]),
                          max([right(s) for s in lst] + [span_x(s)[1] for s in members]),
                          member_ids | lst_ids)
        if not ends:
            continue
        # one pitch for every row, no wider than MAX_GAP between texts
        pitch = min(p0 + (o.body["bottom"] - 0.05 - last_bottom) / (len(lst) - 1), p0 + MAX_GAP, max(hs) + MAX_GAP)
        # a spread only ever opens rows up: never closer than drawn
        step = pitch - p0
        if step < 0.03:
            continue
        for s in lst + members:
            row = _index(lst, s) if id(s) in lst_ids else row_index.get(id(s))
            if row:
                _shift(s, row * step)
            moved.add(id(s))
        moved.update(lst_ids)
        changed.append({"name": lst[0].get("name") or "rows", "what": "rows spread to the body's bottom",
                        "from": _num(p0, 3), "to": _num(pitch, 3)})


def _arrange(shapes, o, changed, spread=True):
    top = [s for s in shapes if s.get("kind") in ("sp", "cxn")]
    for s in top:
        if s.get("kind") == "sp":
            s["origY"] = s["y"]
    _fit_boxes(top, o, changed)
    if spread:
        _stretch_cards(top, o, changed)
        _distribute_rows(top, o, changed)
    for s in top:
        s.pop("origY", None)


# ---- 4. fill ----------------------------------------------------------------

def _faults(shapes, o):
    """Each failure of the fill's check, as (text, what it meets, which shape)
    by index in `shapes` (_fits stops at the first)."""
    index = {id(x): k for k, x in enumerate(shapes)}
    boxes = [s for s in shapes if is_text_box(s)]
    cards = [s for s in shapes if is_card(s)]
    for s in boxes:
        i = index[id(s)]
        ts = text_top(s, o)
        tb = ts + text_height(s, o)
        if o.body and s["y"] < o.body["bottom"] and tb > o.body["bottom"] + 0.05:
            yield (i, "bottom", None)
        probe = dict(s, h=max(s["h"], 0.01))
        card = next((c for c in cards if inside(probe, c) or (
            s["x"] >= c["x"] - TOL and right(s) <= right(c) + TOL and s["y"] >= c["y"] - TOL and s["y"] < bottom(c))), None)
        if card is not None and tb > bottom(card) + 0.02:
            yield (i, "card", index[id(card)])
        for L in shapes:
            if L.get("kind") != "cxn" or abs(L["y1"] - L["y2"]) > 0.02:
                continue
            lx = {"x": min(L["x1"], L["x2"]), "w": abs(L["x2"] - L["x1"])}
            lx["y"], lx["h"] = 0, 0
            if overlap_x(s, lx) >= 0.2 and ts + 0.02 < L["y1"] < tb - 0.02:
                yield (i, "rule", index[id(L)])
        for f in shapes:
            if f.get("kind") != "sp" or not js_truthy(f.get("fill")) or js_truthy(f.get("rot")) or f is s or inside(s, f) or overlap_x(s, f) < 0.2:
                continue
            if ts + 0.02 < f["y"] < tb - 0.02 and not (s["y"] >= f["y"] - 0.02 and s["y"] < bottom(f)):
                yield (i, "fill", index[id(f)])
        for b in boxes:
            if b is s or overlap_x(s, b) < 0.2:
                continue
            bt = text_top(b, o)
            if min(tb, bt + text_height(b, o, True)) - max(ts, bt) > 0.02:
                yield (i, "text", index[id(b)])


def _fits(shapes, o):
    return next(_faults(shapes, o), None) is None


def _raisable(s):
    return is_text_box(s) and s["w"] >= NARROW


def _lead(s):
    paras = (s.get("text") or {}).get("paras") or []
    runs = (paras[0].get("runs") or []) if paras else []
    return runs[0] if runs else None


def _color_key(c):
    return "" if not js_truthy(c) else c if isinstance(c, str) else "[object Object]" if isinstance(c, dict) else js_str(c)


def _held_rows(shapes, seed=frozenset()):
    """One row, one size: texts drawn side by side at one size, weight and
    colour end at one size. Where the raise must skip one of them (a narrow
    label, text in a coloured shape), or one must stay (a seed: it would get
    closer to a shape beside it), the others are held too; a seed keeps its
    whole kind; a text on a grouped table keeps the table's size; and the hold
    runs down a column. -> (indices of the shapes to hold, the held sizes)"""
    items = []

    def walk(lst, top):
        for i, s in enumerate(lst):
            if js_truthy(s.get("children")):
                walk(s["children"], False)
            else:
                items.append((s, i if top else -1))

    walk(shapes, True)
    rows = {}
    for s, i in items:
        r = _lead(s)
        if not r or not has_text(s):
            continue
        sz = r.get("sz") or 12
        if sz < FILL_FROM or sz > FILL_TO:
            continue
        # a shape without a place (a placeholder) is NaN in JS: they share one key
        y = s.get("y")
        k = ("NaN" if not isinstance(y, (int, float)) else round_half_up(y * 50), sz, js_truthy(r.get("b")), _color_key(r.get("color")))
        rows.setdefault(k, []).append((s, i))
    held, sizes = set(), []
    for k, row in rows.items():
        up = [i for s, i in row if i >= 0 and _raisable(s)]
        if not up or (len(up) == len(row) and not any(i in seed for i in up)):
            continue
        held.update(up)
        sizes.append(k[1])

    def kind_of(s):
        r = _lead(s)
        return (r.get("sz") or 12, js_truthy(r.get("b")), _color_key(r.get("color"))) if r else None

    seed_kinds = {kind_of(shapes[i]) for i in seed} - {None}
    by_kind = []
    for s, i in items:
        if i < 0 or not _raisable(s) or kind_of(s) not in seed_kinds:
            continue
        if i not in held and i not in seed:
            by_kind.append((s, i))
        held.add(i)
        if kind_of(s)[0] not in sizes:
            sizes.append(kind_of(s)[0])
    held.update(seed)
    # a text held only for its kind is let go where it heads text of its size that is free to grow
    for s, i in by_kind:
        sz = kind_of(s)[0]
        if any(j >= 0 and b is not s and _raisable(b) and j not in held and kind_of(b) and kind_of(b)[0] == sz
               and abs(b["x"] - s["x"]) < TOL and b["y"] >= bottom(s) - 0.05 and b["y"] - bottom(s) <= 0.8
               for b, j in items):
            held.discard(i)
    # a text on a table drawn as a group is held with the table
    plain = [(s, i) for s, i in items if i >= 0 and _raisable(s) and _lead(s)]
    tables = []
    for g in shapes:
        if not js_truthy(g.get("children")):
            continue
        cells = []

        def walk2(lst):
            for c in lst:
                if js_truthy(c.get("children")):
                    walk2(c["children"])
                else:
                    cells.append(c)

        walk2(g["children"])
        bx = [c for c in cells if c.get("kind") == "sp"]
        if not bx:
            continue
        tables.append({"x0": min(c["x"] for c in bx), "y0": min(c["y"] for c in bx),
                       "x1": max(c["x"] + c["w"] for c in bx), "y1": max(c["y"] + c["h"] for c in bx),
                       "sizes": {((_lead(c) or {}).get("sz") or 12) for c in bx if has_text(c)}})
    for s, i in plain:
        sz = _lead(s).get("sz") or 12
        if sz < FILL_FROM or sz > FILL_TO or i in held:
            continue
        cx = s["x"] + min(s["w"], 0.6) / 2
        cy = s["y"] + s["h"] / 2
        if any(sz in g["sizes"] and g["x0"] < cx < g["x1"] and g["y0"] < cy < g["y1"] for g in tables):
            held.add(i)
            if sz not in sizes:
                sizes.append(sz)
    # the hold runs down a column: same size and colour, a shared left edge or centre
    grew = True
    while grew:
        grew = False
        for s, i in plain:
            if i in held:
                continue
            rt = _lead(s)
            sz = rt.get("sz") or 12
            if sz < FILL_FROM or sz > FILL_TO:
                continue
            for h, j in plain:
                if j in held and (_lead(h).get("sz") or 12) == sz and _color_key(_lead(h).get("color")) == _color_key(rt.get("color")) \
                        and (abs(h["x"] - s["x"]) < TOL or abs(h["x"] + h["w"] / 2 - s["x"] - s["w"] / 2) < TOL):
                    held.add(i)
                    grew = True
                    break
    return held, sizes


def _raise(shapes, d, held=frozenset()):
    for i, s in enumerate(shapes):
        # labels in diagrams and charts keep their size; a held row keeps its size
        if not _raisable(s) or i in held:
            continue
        for p in _paras(s):
            for r in p.get("runs") or []:
                sz = r.get("sz") or 12
                if FILL_FROM <= sz <= FILL_TO:
                    r["sz"] = min(MAX_FILL, sz + d)


def _run_sizes(shapes):
    return [r.get("sz") or 12 for s in shapes if is_text_box(s) for p in _paras(s) for r in p.get("runs") or []]


def fit_layout(shapes, fonts=None, measure=None, body=BODY, page_id=None, max_raise=None, safety_add=SAFETY_ADD,
               label_half_step=False):
    """layoutfit.js fitLayout: the margin sizes boxes, never the raise. The pass
    runs as drawn (no margin) to learn the raise, then with the margin, capped
    at it. -> {"changed": [...], "size": points added}"""
    kw = dict(fonts=fonts, measure=measure, body=body, page_id=page_id, safety_add=safety_add, label_half_step=label_half_step)
    as_drawn = _fit_core(copy.deepcopy(shapes), max_raise=max_raise, margin=False, **kw)
    if isinstance(max_raise, (int, float)) and not isinstance(max_raise, bool):
        given = max_raise
    elif page_id is not None and isinstance(CAPS.get(page_id), (int, float)):
        given = CAPS[page_id]
    else:
        given = None
    if given is not None and given < 0:
        return _fit_core(shapes, max_raise=max_raise, margin=True, **kw)
    cap = min(given, as_drawn["size"]) if given is not None else as_drawn["size"]
    return _fit_core(shapes, max_raise=cap, margin=True, **kw)


def _fit_core(shapes, fonts=None, measure=None, body=BODY, page_id=None, max_raise=None, safety_add=SAFETY_ADD,
              label_half_step=False, margin=True):
    """Fits, arranges and fills a page's shapes in place (layoutfit.js).

    page_id: the library page's id, for its measured cap (layoutfit-caps.json);
    max_raise: points the fill may add at most (overrides the cap); below 0
    leaves the page as drawn.
    -> {"changed": [{name, what, from, to}], "size": points added}"""
    # safety_add: the flat room over each text's estimate, in the page's inches (the
    # library's 0.04; the beautifier passes 0.04 slide inches in page inches)
    o = _O(measure or Measure(), fonts or DEFAULT_FONTS, body, None, safety_add, margin)
    if isinstance(max_raise, (int, float)) and not isinstance(max_raise, bool):
        o.max_raise = max_raise
    elif page_id is not None and isinstance(CAPS.get(page_id), (int, float)):
        o.max_raise = CAPS[page_id]
    changed = []
    if o.max_raise is not None and o.max_raise < 0:
        return {"changed": changed, "size": 0}
    if sum(1 for s in shapes if is_text_box(s)) < 2:
        return {"changed": changed, "size": 0}
    base = copy.deepcopy(shapes)
    # tables, groups and pictures are never moved: a page with one is only fitted
    fixed = any(s.get("kind") in ("tbl", "grp", "pic") for s in shapes)
    _arrange(shapes, o, changed, not fixed)
    if not _fits(shapes, o):
        plain = copy.deepcopy(base)
        ch = []
        _arrange(plain, o, ch, False)
        shapes[:] = plain
        return {"changed": ch, "size": 0}
    if not o.body:
        return {"changed": changed, "size": 0}
    def sizes_of(held):
        zs = [r.get("sz") or 12 for i, s in enumerate(shapes) if is_text_box(s) and i not in held
              for p in _paras(s) for r in p.get("runs") or []]
        return [z for z in zs if FILL_FROM <= z <= FILL_TO]

    if not sizes_of(_held_rows(base)[0]):
        return {"changed": changed, "size": 0}
    # body text stays half a point under the smallest text the fill leaves alone
    kept = [z for z in _run_sizes(shapes) if FILL_TO < z < 20]
    ceiling = min(MAX_FILL, min(kept) - 0.5 if kept else MAX_FILL)
    # labels outside a column of rows may not gain a line
    base_lines = [lines_of(s, o) if is_text_box(s) and not _in_row_column(s, base) else None for s in base]
    # rows lined up as drawn stay lined up
    aligned = []
    for i, a in enumerate(base):
        if not is_text_box(a):
            continue
        for j in range(i + 1, len(base)):
            b = base[j]
            if not is_text_box(b) or overlap_x(a, b) > 0:
                continue
            if abs(text_top(a, o) - text_top(b, o)) <= 0.02:
                aligned.append((i, j))
    # columns of three or more boxes keep at least the span their texts had
    col_idx = {}
    for i, s in enumerate(base):
        if is_text_box(s):
            col_idx.setdefault(to_fixed(s["x"], 2) + "|" + to_fixed(s["w"], 2), []).append(i)

    def span_of(lst, idx):
        tops = [text_top(lst[i], o) for i in idx]
        return max(tops) - min(tops)

    columns = [(idx, span_of(base, idx)) for idx in col_idx.values() if len(idx) >= 3]
    base_sideways = sideways_depths(base, o)
    best = None
    # a text that would get closer to a shape beside it stays at its size, with
    # its kind (a seed), and the fill runs again: the rest of the page still fills
    seed = set()
    for _round in range(6):
        held, held_sizes = _held_rows(base, seed)
        sizes = sizes_of(held)
        if not sizes:
            break
        best = None
        closer = None
        cap = o.max_raise if o.max_raise is not None else INF
        # OPTION (off by default; not in layoutfit.js): a smaller label keeps at
        # least half of its drawn step under the text the fill leaves alone
        # above it (11 under 14 stops at 12.5), so the label's role stays seen
        if label_half_step and kept:
            for t in sizes:
                above = [k for k in kept if k > t]
                if above:
                    cap = min(cap, (min(above) - t) / 2)
        # text drawn smaller than a held row grows at most to that row's size
        for h in held_sizes:
            for t in sizes:
                if t < h:
                    cap = min(cap, h - t)
        d = 0.5
        while max(sizes) + d <= ceiling + 1e-9 and d <= cap + 1e-9:
            trial = copy.deepcopy(base)
            _raise(trial, d, held)
            if any(base_lines[i] is not None and lines_of(x, o) > base_lines[i] for i, x in enumerate(trial)):
                break
            ch = []
            _arrange(trial, o, ch, not fixed)
            if not _fits(trial, o):
                break
            if any(span_of(trial, idx) < span - 0.05 for idx, span in columns):
                break
            # nor may a line grow into a shape beside it, nor a near miss get closer
            side = sideways_depths(trial, o)
            nearer = [k[0] for k, dep in side.items() if k not in base_sideways or dep > base_sideways[k] + 0.005]
            if nearer:
                closer = nearer
                break
            if any(abs(text_top(trial[i], o) - text_top(trial[j], o)) > 0.05 for i, j in aligned):
                break
            best = (trial, d, ch)
            d += 0.5
        if not closer or all(i in seed for i in closer):
            break
        seed.update(closer)
    if not best:
        return {"changed": changed, "size": 0}
    shapes[:] = best[0]
    best[2].append({"name": "page", "what": "body text up to fill the page", "from": 0, "to": best[1]})
    return {"changed": best[2], "size": best[1]}


def page_fits(shapes, fonts=None, measure=None, body=BODY, safety_add=SAFETY_ADD):
    """No text past the body's bottom, out of its card, onto a rule, into a
    filled shape or into another text (the fill's own check)."""
    return _fits(shapes, _O(measure or Measure(), fonts or DEFAULT_FONTS, body, None, safety_add))


def fit_faults(shapes, fonts=None, measure=None, body=BODY, safety_add=SAFETY_ADD):
    """Each failure of the fill's check, as (text, what it meets, which shape) by
    index in `shapes` (a library page can fail it by design: a label drawn on a
    rule)."""
    return set(_faults(shapes, _O(measure or Measure(), fonts or DEFAULT_FONTS, body, None, safety_add)))
