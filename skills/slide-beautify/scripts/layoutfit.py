"""Text boxes sized to their text, and cards and rows drawn tight and even.

A port of SlideShortcuts' templates/layoutfit.js, run on a library page once
the deck's words are in it, so the rebuilt page is fitted and evenly spaced
the way our own library pages are:

1. FIT. Every plain text box (no fill, no line, upright, a rectangle) takes
   the height its text needs, with Resize Shape to Fit Text left on. The edge
   its text is anchored to stays, so the text does not move.
2. STACK. In a column of boxes, one that now runs into the box under it
   pushes it down (and what is under that), keeping the page's gap.
3. CARDS. A filled card with no text of its own ends one margin under the
   last thing in it; cards side by side keep one height; a card in a grid
   keeps its size.
4. DISTRIBUTE. Three or more like boxes, one per row, get equal gaps between
   their texts (the first and last stay); each row's label, rule and number
   move with it.

Shapes are the page's descriptions, in inches (pagexml.py). Text is measured
with fixed width tables (assets/font-widths.json, the same tables the
library's own tests use), so this and the library agree. Standard library only.
"""

import json
import math
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
TOL = 0.02
MIN_SLACK = 0.12
BIG_GAP = 0.45
CARD_GEOMS = ("rect", "roundRect")
SAFETY = 1.06
SAFETY_ADD = 0.04
MIN_LINE = 1.2
NEAR = 0.35
FOOT = 0.35
MAX_FILL = 14
FILL_FROM = 9
FILL_TO = 13.5
MAX_GAP = 0.9
NARROW = 1.2
BODY = {"top": 1.62, "bottom": 6.58}
DEFAULT_INSET = [0.08, 0.05, 0.08, 0.05]


class Measure:
    """Widths from the fixed tables; a face not in them is measured as Arial."""

    def __init__(self, path=os.path.join(HERE, "..", "assets", "font-widths.json")):
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
        return self.face(font).get("lineHeight", 1.2)


def bottom(s):
    return s["y"] + s["h"]


def right(s):
    return s["x"] + s["w"]


def has_text(s):
    t = s.get("text") or {}
    return any(str(r.get("t") or "").strip() for p in t.get("paras") or [] for r in p.get("runs") or [])


def is_text_box(s):
    return (s.get("kind") == "sp" and (s.get("geom") or "rect") == "rect" and not s.get("rot")
            and not s.get("fill") and not s.get("line") and has_text(s) and not (s.get("text") or {}).get("vert"))


def is_card(s):
    return (s.get("kind") == "sp" and (s.get("geom") or "rect") in CARD_GEOMS and not s.get("rot")
            and bool(s.get("fill")) and not has_text(s) and s["w"] > 0.6 and s["h"] > 0.6)


def inside(a, c):
    return (a is not c and a["x"] >= c["x"] - TOL and a["y"] >= c["y"] - TOL
            and right(a) <= right(c) + TOL and bottom(a) <= bottom(c) + TOL)


def overlap_x(a, b):
    return min(right(a), right(b)) - max(a["x"], b["x"])


def cxn_span_y(s):
    return min(s["y1"], s["y2"]), max(s["y1"], s["y2"])


def text_height(s, m, fonts, raw=False):
    """Inches the text of `s` needs at its sizes, with its insets."""
    t = s.get("text") or {}
    ins = t.get("inset") or DEFAULT_INSET
    room = max((s["w"] - ins[0] - ins[2]) * 72, 4)
    pts = 0.0
    for p in t.get("paras") or []:
        runs = [r for r in p.get("runs") or [] if r.get("t") is not None]
        size = next((r["sz"] for r in runs if r.get("sz")), 12)
        font = fonts["heading"] if any(r.get("font") == "head" for r in runs) else fonts["body"]
        bold = bool(runs) and all(r.get("b") for r in runs)
        indent = ((p.get("level") or 0) * 0.2 + (0.17 if p.get("bullet") else 0)) * 72
        width = max(room - indent, 4)
        text = "".join(str(r["t"]) for r in runs)
        lines = 0
        if t.get("wrap") is False:
            lines = 1
        else:
            for para in re.split(r"\r?\n|\v", text):
                line = 0.0
                lines += 1
                for word in re.split(r"(\s+)", para):
                    if not word:
                        continue
                    w = m.width(word, font, size, bold)
                    if line + w > width and line > 0 and not word.isspace():
                        lines += 1
                        line = w
                    else:
                        line += w
        spacing = (p.get("spacing") or 100) / 100.0
        pts += lines * size * max(m.line_height(font), 0 if raw else MIN_LINE) * spacing
        pts += (p.get("after") or 0) + (p.get("before") or 0)
    if raw:
        return pts / 72 + ins[1] + ins[3]
    return (pts / 72) * SAFETY + SAFETY_ADD + ins[1] + ins[3]


def text_top(s, m, fonts):
    need = text_height(s, m, fonts, True)
    a = (s.get("text") or {}).get("anchor") or "t"
    if a == "ctr":
        return s["y"] + (s["h"] - need) / 2
    if a == "b":
        return s["y"] + s["h"] - need
    return s["y"]


def _fit_boxes(shapes, m, fonts, changed):
    for s in shapes:
        if not is_text_box(s):
            continue
        need = text_height(s, m, fonts)
        if abs(s["h"] - need) < MIN_SLACK and need <= s["h"]:
            continue
        if s["w"] < NARROW:
            r = next((x for x in (s["text"]["paras"][0].get("runs") or []) if x.get("sz")), {}) if s["text"].get("paras") else {}
            need = min(max(need, s["h"]), need + (r.get("sz") or 12) * MIN_LINE / 72)
        if abs(s["h"] - need) < MIN_SLACK and need <= s["h"]:
            continue
        a = s["text"].get("anchor") or "t"
        frm = s["h"]
        if a == "ctr":
            s["y"] += (s["h"] - need) / 2
        elif a == "b":
            s["y"] += s["h"] - need
        s["h"] = need
        if s["text"].get("wrap") is not False:
            s["text"]["autofit"] = "resize"
        changed.append({"name": s.get("name") or "text box", "what": "height to its text", "from": round(frm, 3), "to": round(need, 3)})


def _stack(shapes, m, fonts, changed, orig_y, orig_top, orig_bottom):
    boxes = sorted([s for s in shapes if is_text_box(s)], key=lambda s: s["y"])
    for i, a in enumerate(boxes):
        for b in boxes[i + 1:]:
            if overlap_x(a, b) < 0.6 * min(a["w"], b["w"]):
                continue
            gap = orig_top[id(b)] - orig_bottom[id(a)]
            if gap < 0:
                break
            at = text_top(b, m, fonts)
            a_bottom = text_top(a, m, fonts) + text_height(a, m, fonts, True)
            if a_bottom <= at:
                break
            want = a_bottom + min(gap, 0.04)
            if at < want - TOL and orig_y[id(b)] >= orig_y[id(a)]:
                d = want - at
                top = b["y"]
                for s in shapes:
                    if s is not a and s.get("kind") == "sp" and s["y"] >= top - TOL and overlap_x(s, b) > 0:
                        s["y"] += d
                changed.append({"name": b.get("name") or "text box", "what": "moved down, clear of the box above", "from": 0, "to": round(d, 3)})
            break


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
        beside = (d["x"] >= right(c) - TOL and d["x"] - right(c) < NEAR) or (c["x"] >= right(d) - TOL and c["x"] - right(d) < NEAR)
        if down and beside and abs(d["h"] - c["h"]) > TOL:
            return True
    return False


def _kids(c, shapes):
    out = []
    for s in shapes:
        if s.get("kind") == "cxn":
            y0, y1 = cxn_span_y(s)
            if y0 >= c["y"] - TOL and y1 <= bottom(c) + TOL and min(s["x1"], s["x2"]) >= c["x"] - TOL and max(s["x1"], s["x2"]) <= right(c) + TOL:
                out.append(s)
        elif s.get("kind") == "sp" and inside(s, c):
            out.append(s)
    return out


def _ends_page(shapes, y, x0, x1, except_ids):
    for s in shapes:
        if id(s) in except_ids or s.get("kind") == "ph":
            continue
        sy = cxn_span_y(s)[0] if s.get("kind") == "cxn" else s["y"]
        sx0, sx1 = (min(s["x1"], s["x2"]), max(s["x1"], s["x2"])) if s.get("kind") == "cxn" else (s["x"], right(s))
        if sy >= y - TOL and sx1 > x0 + TOL and sx0 < x1 - TOL:
            return False
    return True


def _shift(s, d):
    if s.get("kind") == "cxn":
        s["y1"] += d
        s["y2"] += d
    else:
        s["y"] += d


def _span(s):
    if s.get("kind") == "cxn":
        return cxn_span_y(s), (min(s["x1"], s["x2"]), max(s["x1"], s["x2"]))
    return (s["y"], bottom(s)), (s["x"], right(s))


def _stretch_cards(shapes, body, changed):
    if not body:
        return
    every = [s for s in shapes if is_card(s)]
    cards = [c for c in every if not _in_grid(c, every)]
    done = set()
    for c in cards:
        if id(c) in done:
            continue
        row = [d for d in cards if abs(d["y"] - c["y"]) < TOL and abs(d["h"] - c["h"]) < TOL]
        done.update(id(d) for d in row)
        if len(row) < 2 and c["h"] < 0.5 * (body["bottom"] - body["top"]):
            continue
        row_bottom = bottom(c)
        kids = []
        for card in row:
            kids += [k for k in _kids(card, shapes) if all(k is not x for x in kids)]
        left = min(x["x"] for x in row)
        right_edge = max(right(x) for x in row)
        member = set(id(x) for x in row) | set(id(k) for k in kids)
        between = [s for s in shapes if s.get("kind") == "sp" and id(s) not in member and s["y"] >= c["y"] - TOL
                   and bottom(s) <= row_bottom + TOL and s["x"] >= left - TOL and right(s) <= right_edge + TOL]
        except_ids = member | set(id(s) for s in between)
        if not _ends_page(shapes, row_bottom - TOL, left, right_edge, except_ids):
            continue
        if any(id(s) not in except_ids and s.get("kind") == "sp" and abs(bottom(s) - row_bottom) < 0.06
               and (right(s) <= left + TOL or s["x"] >= right_edge - TOL) for s in shapes):
            continue
        d = body["bottom"] - row_bottom
        if abs(d) < TOL:
            continue
        for card in row:
            own = _kids(card, shapes)
            spans = sorted([(_span(s)[0][0], _span(s)[0][1]) for s in own])
            reach = spans[0][1] if spans else 0
            cut_at, widest = None, 0.3
            for t0, t1 in spans[1:]:
                if t0 - reach >= widest:
                    widest, cut_at = t0 - reach, t0
                reach = max(reach, t1)
            if cut_at is not None:
                for s in own:
                    (y0, y1), _ = _span(s)
                    if y0 >= cut_at - TOL and row_bottom - y1 < FOOT:
                        _shift(s, d)
            card["h"] += d
            changed.append({"name": card.get("name") or "card", "what": "card reaches the body's bottom", "from": round(card["h"] - d, 3), "to": round(card["h"], 3)})
        for s in between:
            _shift(s, d / 2)


def _distribute(shapes, m, fonts, body, changed, orig_y):
    cols = {}
    for s in shapes:
        if not is_text_box(s):
            continue
        size = next((r["sz"] for r in s["text"]["paras"][0].get("runs") or [] if r.get("sz")), 12)
        cols.setdefault("%.2f|%.2f|%s" % (s["x"], s["w"], size), []).append(s)
    cards = [s for s in shapes if is_card(s)]
    moved = set()
    order = sorted(cols.values(), key=lambda l: (-len(l), -sum(s["h"] for s in l)))
    oy = lambda s: orig_y.get(id(s), s.get("y"))
    for lst in order:
        if len(lst) < 3 or any(id(s) in moved for s in lst):
            continue
        lst.sort(key=lambda s: s["y"])
        tops0 = [oy(s) for s in lst]
        pitches = [t - tops0[i] for i, t in enumerate(tops0[1:])]
        if max(pitches) - min(pitches) > 0.05 or min(pitches) < 0.2:
            continue
        if any(sum(1 for s in lst if inside(s, c)) > 1 for c in cards):
            continue
        hs = [text_height(s, m, fonts, True) for s in lst]
        tts = [text_top(s, m, fonts) for s in lst]
        p0 = (tops0[-1] - tops0[0]) / (len(lst) - 1)
        band_top = tops0[0] - p0 / 2
        band_bottom = tops0[-1] + p0 / 2 + (bottom(lst[-1]) - lst[-1]["y"])
        in_list = set(id(s) for s in lst)
        members = []
        for s in shapes:
            if s.get("kind") == "ph" or id(s) in in_list:
                continue
            (sy, sb), _ = _span(s)
            if sy >= band_top - 0.3 and sb <= band_bottom + 0.1 and sb - sy <= 1.2 * p0:
                members.append(s)
        groups = {}
        for s in members:
            _, (x0, x1) = _span(s)
            groups.setdefault("%s|%.1f|%.1f" % (s.get("kind"), x0, x1), []).append(s)
        top_of = lambda s: _span(s)[0][0] if s.get("kind") == "cxn" else oy(s)
        row_index = {}
        orderly = True
        for g in groups.values():
            between = len(g) == len(lst) - 1 and all(s.get("kind") == "cxn" for s in g)
            if len(g) != len(lst) and not between:
                orderly = False
                break
            g.sort(key=top_of)
            first = 1 if between else 0
            offs = [top_of(s) - tops0[j + first] for j, s in enumerate(g)]
            if max(offs) - min(offs) > 0.06:
                orderly = False
                break
            for j, s in enumerate(g):
                row_index[id(s)] = j + first
        if not orderly:
            continue
        band_lo, band_hi = tops0[0], bottom(lst[-1])
        member_ids = set(id(s) for s in members)
        tied = False
        for s in shapes:
            if (s.get("kind") == "ph" and s.get("type") != "pic") or id(s) in in_list or id(s) in member_ids:
                continue
            (sy, sb), _ = _span(s)
            if min(sb, band_hi) - max(sy, band_lo) > 1.2 * p0:
                tied = True
                break
        if tied:
            continue
        last_bottom = tts[-1] + hs[-1]
        low = max([last_bottom] + [_span(s)[0][1] for s in members])
        x0 = min([s["x"] for s in lst] + [_span(s)[1][0] for s in members])
        x1 = max([right(s) for s in lst] + [_span(s)[1][1] for s in members])
        if not (body and _ends_page(shapes, low - TOL, x0, x1, in_list | member_ids)):
            continue
        pitch = min(p0 + (body["bottom"] - 0.05 - last_bottom) / (len(lst) - 1), p0 + MAX_GAP)
        step = pitch - p0
        if abs(step) < 0.03:
            continue
        for i, s in enumerate(lst):
            if i:
                _shift(s, i * step)
            moved.add(id(s))
        for s in members:
            r = row_index.get(id(s), 0)
            if r:
                _shift(s, r * step)
            moved.add(id(s))
        changed.append({"name": lst[0].get("name") or "rows", "what": "rows spread to the body's bottom", "from": round(p0, 3), "to": round(pitch, 3)})


def _arrange(shapes, m, fonts, body, changed, spread=True):
    top = [s for s in shapes if s.get("kind") in ("sp", "cxn")]
    orig_y, orig_top, orig_bottom = {}, {}, {}
    for s in top:
        if s.get("kind") != "sp":
            continue
        orig_y[id(s)] = s["y"]
        if is_text_box(s):
            orig_top[id(s)] = text_top(s, m, fonts)
            orig_bottom[id(s)] = orig_top[id(s)] + text_height(s, m, fonts, True)
    _fit_boxes(top, m, fonts, changed)
    _stack(top, m, fonts, changed, orig_y, orig_top, orig_bottom)
    if spread:
        _stretch_cards(top, body, changed)
        _distribute(top, m, fonts, body, changed, orig_y)


def _fits(shapes, m, fonts, body):
    boxes = [s for s in shapes if is_text_box(s)]
    cards = [s for s in shapes if is_card(s)]
    for s in boxes:
        ts = text_top(s, m, fonts)
        tb = ts + text_height(s, m, fonts)
        if body and s["y"] < body["bottom"] and tb > body["bottom"] + 0.05:
            return False
        for c in cards:
            if s["x"] >= c["x"] - TOL and right(s) <= right(c) + TOL and c["y"] - TOL <= s["y"] < bottom(c):
                if tb > bottom(c) + 0.02:
                    return False
                break
        for b in boxes:
            if b is s or overlap_x(s, b) < 0.2:
                continue
            bt = text_top(b, m, fonts)
            if min(tb, bt + text_height(b, m, fonts, True)) - max(ts, bt) > 0.02:
                return False
    return True


def _raise(shapes, d):
    for s in shapes:
        if not is_text_box(s) or s["w"] < NARROW:
            continue
        for p in s["text"].get("paras") or []:
            for r in p.get("runs") or []:
                sz = r.get("sz") or 12
                if FILL_FROM <= sz <= FILL_TO:
                    r["sz"] = min(MAX_FILL, sz + d)


def fit_layout(shapes, fonts=None, measure=None, body=BODY):
    """Fits, arranges and fills a page's shapes in place (layoutfit.js).
    -> (changes, points added to the body text)"""
    import copy
    m = measure or Measure()
    fonts = fonts or {"heading": "Georgia", "body": "Arial"}
    changed = []
    if sum(1 for s in shapes if is_text_box(s)) < 2:
        return changed, 0
    base = copy.deepcopy(shapes)
    fixed = any(s.get("kind") in ("tbl", "grp", "pic") for s in shapes)
    _arrange(shapes, m, fonts, body, changed, spread=not fixed)
    if not _fits(shapes, m, fonts, body):
        plain = copy.deepcopy(base)
        ch = []
        _arrange(plain, m, fonts, body, ch, spread=False)
        shapes[:] = plain
        return ch, 0
    if not body:
        return changed, 0
    sizes = [r.get("sz") or 12 for s in shapes if is_text_box(s) for p in s["text"].get("paras") or [] for r in p.get("runs") or []]
    sizes = [z for z in sizes if FILL_FROM <= z <= FILL_TO]
    if not sizes:
        return changed, 0
    best = None
    d = 0.5
    while max(sizes) + d <= MAX_FILL + 1e-9:
        trial = copy.deepcopy(base)
        _raise(trial, d)
        ch = []
        _arrange(trial, m, fonts, body, ch, spread=not fixed)
        if not _fits(trial, m, fonts, body):
            break
        best = (trial, d, ch)
        d += 0.5
    if not best:
        return changed, 0
    shapes[:] = best[0]
    best[2].append({"name": "page", "what": "body text up to fill the page", "from": 0, "to": best[1]})
    return best[2], best[1]
