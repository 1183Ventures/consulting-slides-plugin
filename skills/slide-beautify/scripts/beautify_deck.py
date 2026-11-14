#!/usr/bin/env python3
"""Rebuild slides on library pages, in the deck's own theme.

Usage:
    python3 beautify_deck.py DECK.pptx PLAN.json OUT.pptx [--report REPORT.json] [--keep-originals]

PLAN.json says which slides to rebuild and how:

    {"pages": [
      {"slide": 3, "page": "executive-summary",
       "title": "Three moves lift operating profit by a fifth",
       "slots": {"s1": "Context", "s2": "The market grew 6% a year ...",
                 "s4": ["First point", "**Lead-in:** the rest of the line"]},
       "keep": ["s9"], "clear": ["s10"],
       "source": "Source: Company data, 2025"}
    ]}

- "page" is a page id from assets/pages.json (read_deck.py --pages lists them).
- "slots" fills a page's text slots. A value is a string, or a list of
  paragraphs; a paragraph is a string or {"text": ..., "level": 1}.
  "**Lead-in:** rest" makes the lead-in bold.
- Every slot whose role is not "fixed" must be filled, kept ("keep": the
  page's own sample text stays, for generic headers like "Pros") or cleared
  ("clear": an empty box, removed when it has no fill and no line).
- "source" puts a source line where the page has one, or at the bottom left.

The deck is never written to. OUT.pptx is the deck with each planned slide
replaced by its rebuilt page in the same place (its speaker notes kept), or
followed by the hidden original with --keep-originals. Other slides are
untouched. The report lists every change, every text that had to shrink and
anything that may still overflow.

Needs python-pptx (pip install python-pptx).
"""

import argparse
import copy
import json
import math
import os
import re
import sys

try:
    from lxml import etree
    from pptx import Presentation
    from pptx.enum.shapes import PP_PLACEHOLDER
except ImportError:
    sys.exit("This needs python-pptx: pip install python-pptx")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pagexml  # noqa: E402
import layoutfit  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_PAGES = os.path.join(HERE, "..", "assets", "pages.json")
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
LEAD = re.compile(r"^\*\*(.+?)\*\*\s*(.*)$", re.S)
TITLE_TYPES = (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)
KEEP_PH = (PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER)
MIN_SIZE = {"heading": 10, "body": 9, "label": 8, "fixed": 8}
CHAR_W = 0.52   # average character width, in ems, for the fit estimate
LINE_H = 1.2
SLIDE_MARGIN = 1.03   # width room for decks not at the library's scale (Reviewer, 13 Nov 2026)
# a theme face without a width table: the widest common face measured is 15%
# wider than Arial, and lines are set up to 1.5 of the size (Poppins by its own
# file; Poppins has a table of its own, at its Windows figure of 1.76)
UNKNOWN_FACE = {"width": 1.15, "line": 1.5}


class PlanError(Exception):
    pass


def load_pages(path):
    with open(path, encoding="utf-8") as fh:
        return {p["id"]: p for p in json.load(fh)["pages"]}


def title_only_layout(prs):
    """The deck's own title-only layout, else the layout with a title and
    the fewest other placeholders."""
    best, best_score = None, None
    for layout in prs.slide_layouts:
        types = [ph.placeholder_format.type for ph in layout.placeholders]
        if not any(t in TITLE_TYPES for t in types):
            continue
        others = [t for t in types if t not in TITLE_TYPES and t not in KEEP_PH]
        score = (len(others), 0 if "title only" in layout.name.lower() else 1)
        if best_score is None or score < best_score:
            best, best_score = layout, score
    if best is None:
        raise PlanError("The deck has no layout with a title placeholder")
    return best


def paragraphs_of(value):
    if isinstance(value, str):
        return [{"text": value}]
    out = []
    for p in value or []:
        out.append({"text": p} if isinstance(p, str) else dict(p))
    return out


def template_paras(shape):
    paras = [p for p in (shape.get("text") or {}).get("paras", []) if any(r.get("t") for r in p.get("runs", []))]
    return paras or (shape.get("text") or {}).get("paras") or [{"runs": [{}]}]


def fill_text(shape, value):
    """The shape's text with `value` in it, in the sample's own formatting."""
    shape = copy.deepcopy(shape)
    temps = template_paras(shape)
    paras = []
    for i, p in enumerate(paragraphs_of(value)):
        t = copy.deepcopy(temps[min(i, len(temps) - 1)])
        base = dict(next((r for r in t.get("runs", []) if r.get("t")), t["runs"][0] if t.get("runs") else {}))
        base.pop("t", None)
        text = str(p.get("text", ""))
        m = LEAD.match(text)
        if m:
            t["runs"] = [dict(base, t=m.group(1).rstrip() + " ", b=True), dict(base, t=m.group(2), b=False)]
        else:
            t["runs"] = [dict(base, t=text)]
        if p.get("level"):
            t["level"] = int(p["level"])
        paras.append(t)
    shape.setdefault("text", {})["paras"] = paras or [{"runs": []}]
    return shape


def needed_height(shape, pt, place):
    """Points of height the text needs at `pt` points on the slide, and the
    points the box has (a rough, safe estimate)."""
    t = shape.get("text") or {}
    ins = t.get("inset") or pagexml.DEFAULT_INSET
    width = place.width_pt(shape["w"]) - (ins[0] + ins[2]) * 72 * place.s
    total = 0.0
    for p in t.get("paras", []):
        text = "".join(r.get("t", "") for r in p.get("runs", []))
        indent = ((p.get("level") or 0) * pagexml.LEVEL_STEP + (pagexml.BULLET_INDENT if p.get("bullet") else 0)) * 72 * place.s
        per_line = max(1, int(max(width - indent, 8) / (pt * CHAR_W)))
        lines = 1 if t.get("wrap") is False else max(1, math.ceil(len(text) / per_line))
        total += lines * pt * LINE_H + (p.get("after") or 0) + (p.get("before") or 0)
    return total, max(place.height_pt(shape["h"]) - (ins[1] + ins[3]) * 72 * place.s, 4)


def fitted_size(shape, role, place):
    """(points on the slide, the page's own size, fits) for this text."""
    size = next((r.get("sz") for p in (shape.get("text") or {}).get("paras", []) for r in p.get("runs", []) if r.get("sz")), 12)
    start = place.pt(size)
    floor = min(start, MIN_SIZE.get(role, 9))
    s = start
    while s > floor:
        need, have = needed_height(shape, s, place)
        if need <= have:
            return s, size, True
        s -= 0.5
    need, have = needed_height(shape, floor, place)
    return floor, size, need <= have


def check_entry(entry, page, n_slides):
    if not isinstance(entry.get("slide"), int) or not 1 <= entry["slide"] <= n_slides:
        raise PlanError("Slide %r is not in the deck (1 to %d)" % (entry.get("slide"), n_slides))
    slots = {s["id"]: s for s in page["slots"]}
    filled = set(entry.get("slots") or {})
    keep = set(entry.get("keep") or [])
    clear = set(entry.get("clear") or [])
    unknown = sorted((filled | keep | clear) - set(slots))
    if unknown:
        raise PlanError("Slide %d: page %s has no slot %s" % (entry["slide"], page["id"], ", ".join(unknown)))
    open_slots = [sid for sid, s in slots.items()
                  if s["role"] not in ("fixed", "title", "source") and sid not in filled | keep | clear]
    if open_slots:
        raise PlanError("Slide %d: fill, keep or clear every slot of %s; still open: %s"
                        % (entry["slide"], page["id"], ", ".join(open_slots)))


def new_shape_id(slide):
    ids = [int(x) for x in slide.shapes._spTree.xpath(".//p:cNvPr/@id") if str(x).isdigit()]
    return max(ids + [1]) + 1


def deck_fonts(prs):
    """The theme's heading and body faces -> (faces to measure with, faces the
    width tables lack). A face the tables lack is measured as the library's
    Georgia or Arial, with the room UNKNOWN_FACE gives."""
    fonts, unknown = dict(layoutfit.DEFAULT_FONTS), {}
    try:
        part = prs.slide_master.part.part_related_by(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        root = etree.fromstring(part.blob)
    except (KeyError, ValueError, etree.XMLSyntaxError):
        return fonts, unknown
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    known = layoutfit.Measure().fonts
    for key, tag in (("heading", "majorFont"), ("body", "minorFont")):
        face = root.xpath(".//a:fontScheme/a:%s/a:latin/@typeface" % tag, namespaces=ns)
        if not face or not face[0]:
            continue
        if face[0] in known:
            fonts[key] = face[0]
        else:
            unknown[key] = face[0]
    return fonts, unknown


def lay_out(shapes, page, place, deck_faces, report, label_half_step=False, content_top=None, page_index=None):
    """layoutfit.py on the filled page, in the page's own inches: text boxes
    fitted to their text, rows and cards spread to the body's bottom, body
    text raised while the page has room (the library's own pass)."""
    fonts, unknown = deck_faces
    # text measured as the slide will carry it; a deck not drawn at the
    # library's own scale (4:3, a band squeezed under a low title) keeps 3%
    # more room, which the Reviewer measured covers what one model leaves out
    plain = all(abs(v - 1) < 0.005 for v in (place.sx, place.sy, place.k))
    # (a line that fills its box to within half a percent or a point wraps: layoutfit's own rule now)
    margin, floor, max_raise = (1.0 if plain else SLIDE_MARGIN), 0, None
    # a face the tables lack, used by this page's text: fitted with room for a
    # face up to 15% wider and taller lines, and its text never raised
    used = {("heading" if r.get("font") == "head" else "body")
            for s in shapes if layoutfit.is_text_box(s)
            for p in s["text"].get("paras", []) for r in p.get("runs", [])}
    missing = sorted({unknown[k] for k in used if k in unknown})
    if missing:
        margin, floor, max_raise = margin * UNKNOWN_FACE["width"], UNKNOWN_FACE["line"], 0
        report["warnings"].append("The deck's font %s has no width table: the page was fitted with extra room and "
                                  "its text was not enlarged; check it in PowerPoint" % " and ".join(missing))
    measure = layoutfit.SlideMeasure(layoutfit.Measure(), place.pt, place.sx, margin, floor)
    drawn = copy.deepcopy(shapes)
    # the flat 0.04 in of room counts in slide inches, so every slide size keeps the same spare
    add = layoutfit.SAFETY_ADD / (place.sy * place.k)
    r = layoutfit.fit_layout(shapes, fonts=fonts, measure=measure, page_id=page["id"], max_raise=max_raise, safety_add=add,
                             label_half_step=label_half_step)
    fitted = sum(1 for c in r["changed"] if c["what"] == "height to its text")
    report["layout"] = (["%d text boxes fitted to their text" % fitted] if fitted else []) + \
        [c["what"] for c in r["changed"] if c["what"] not in ("height to its text", "body text up to fill the page")]
    report["text_up"] = r["size"]
    r["measure"] = measure
    up = rises(drawn, shapes, measure, fonts, content_top)
    # the page check fails a library page that draws a label on a rule or a shape by design, so a
    # page is warned only when it has more failures than the library page laid out with its own words
    # failures compared by which text and what it meets, in the library page's own numbering: a
    # page warns on any the library page does not have, so clearing a slot frees nothing
    own = copy.deepcopy(page["shapes"])
    layoutfit.fit_layout(own, fonts=fonts, measure=measure, page_id=page["id"], max_raise=max_raise, safety_add=add)
    base = layoutfit.fit_faults(own, fonts=fonts, measure=measure, safety_add=add)
    where = page_index if page_index is not None else list(range(len(shapes)))
    mine = {(where[i], kind, None if j is None else where[j])
            for i, kind, j in layoutfit.fit_faults(shapes, fonts=fonts, measure=measure, safety_add=add)}
    down = bool(mine - base) or not words_fit(shapes, measure, fonts)
    r["fits"] = not (up or down)
    if down:
        report["warnings"].append("Texts on this page may run into each other or past the bottom: shorten the longest")
    if up:
        report["warnings"].append("A text on this page grows up into the title area: shorten it")
    return r


def rises(drawn, shapes, measure, fonts, content_top=None):
    """A text that starts above both the body's top and the box the page drew
    for it (a box anchored at its foot, its text growing up into the title band). The page
    check otherwise looks only down (Reviewer, 14 Nov 2026: big-numbers at
    1.4 times its words)."""
    o = layoutfit._O(measure, fonts, layoutfit.BODY, None)
    for before, after in zip(drawn, shapes):
        if not (layoutfit.is_text_box(before) and layoutfit.is_text_box(after)):
            continue
        # the box as the page drew it: long words can already overflow it upward
        # nor above where the page's content starts (under the deck's title once placed)
        limit = min(before["y"], layoutfit.BODY["top"] if content_top is None else max(content_top, 0))
        if layoutfit.text_top(after, o) < limit - 0.02:
            return True
    return False


def words_fit(shapes, measure, fonts):
    """No single word wider than its text box (a long figure, a name)."""
    for s in shapes:
        if not layoutfit.is_text_box(s) or s["text"].get("wrap") is False:
            continue
        ins = s["text"].get("inset") or layoutfit.DEFAULT_INSET
        room = (s["w"] - ins[0] - ins[2]) * 72
        for p in s["text"].get("paras", []):
            runs = [r for r in p.get("runs", []) if r.get("t") is not None]
            if not runs:
                continue
            face = fonts["heading"] if any(r.get("font") == "head" for r in runs) else fonts["body"]
            size = next((r["sz"] for r in runs if r.get("sz")), 12)
            bold = all(r.get("b") for r in runs)
            width = room - ((p.get("level") or 0) * 0.2 + (0.17 if p.get("bullet") else 0)) * 72
            for word in "".join(str(r["t"]) for r in runs).split():
                if measure.width(word, face, size, bold) > width + 0.5:
                    return False
    return True


A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
HOUSE_TITLE, HOUSE_TITLE_MAX = 24, 32   # pt on a 13.333 in slide (lavender decktitle.js, #376)
LIB_TITLE_BOX = (0.5, 0.56, 12.333, 0.9)  # in, the library's title frame (frame.js), scaled to the slide


def _title_ph_xml(owner):
    """The title placeholder's element in a layout or master, or None."""
    for sp in owner.element.iter(P + "sp"):
        ph = sp.find(".//" + P + "nvPr/" + P + "ph")
        if ph is not None and ph.get("type") in ("title", "ctrTitle"):
            return sp
    return None


def _box_of(sp):
    if sp is None:
        return None
    off, ext = sp.find(".//" + A + "xfrm/" + A + "off"), sp.find(".//" + A + "xfrm/" + A + "ext")
    if off is None or ext is None or int(ext.get("cx", 0)) <= 0 or int(ext.get("cy", 0)) <= 0:
        return None
    return int(off.get("x", 0)), int(off.get("y", 0)), int(ext.get("cx")), int(ext.get("cy"))


def _lvl1(sp, attr, child=None):
    if sp is None:
        return None
    el = sp.find(".//" + A + "lstStyle/" + A + "lvl1pPr")
    if el is None:
        return None
    if child:
        el = el.find(A + child)
    return None if el is None else el.get(attr)


def _title_style(master, attr, child=None):
    el = master.element.find(".//" + P + "txStyles/" + P + "titleStyle/" + A + "lvl1pPr")
    if el is None:
        return None
    if child:
        el = el.find(A + child)
    return None if el is None else el.get(attr)


STD_TITLE_FRAME = (0.5, 0.56, 12.33, 0.9)   # in: the library's own top title frame


def page_title_frame(page):
    """The page's own title frame (library inches) when it is not the standard
    top one: the side-title page's column. None for a standard page."""
    ph = next((s for s in page["shapes"] if s.get("kind") == "ph" and s.get("type") == "title"), None)
    if not ph or "w" not in ph:
        return None
    box = (ph.get("x", 0), ph.get("y", 0), ph["w"], ph.get("h", 0))
    return None if all(abs(a - b) < 0.02 for a, b in zip(box, STD_TITLE_FRAME)) else box


def place_title(prs, slide, title_ph, text, report, frame=None):
    """The title, as the file will draw it: box, size and alignment resolved
    through slide, layout and master and written on the slide itself, never
    left to PowerPoint's autofit (Windows render, 14 Nov 2026: a blank
    deck's 44 pt titles ran into the page). The product's rule (lavender
    decktitle.js, #376): a title style above 32 pt (scaled to the slide's
    width) gets 24 pt, in a box two lines tall at the top of the layout's
    title box. The title is measured with the width tables and steps down
    half a point at a time if it does not fit; a box with no width is
    replaced by the library's title frame. A page with its own title frame
    (the side-title column) has its title there, top-anchored and left-aligned.
    -> (the title box's bottom, the bottom of its text), EMU."""
    layout, master = slide.slide_layout, slide.slide_layout.slide_master
    lay_sp, mas_sp = _title_ph_xml(layout), _title_ph_xml(master)
    E = 914400
    k = prs.slide_width / (13.333 * E)
    box = _box_of(lay_sp) or _box_of(mas_sp)
    if frame is not None:
        # the page's own title frame (the side-title column), scaled to the slide
        fx, fy, fw, fh = frame
        sx, sy = prs.slide_width / (13.333 * E), prs.slide_height / (7.5 * E)
        box = (round(fx * sx * E), round(fy * sy * E), round(fw * sx * E), round(fh * sy * E))
    if box is None:
        x, y, w, h = LIB_TITLE_BOX
        box = (round(x * k * E), round(y * prs.slide_height / (7.5 * E) * E), round(w * k * E), round(h * prs.slide_height / (7.5 * E) * E))
        report["warnings"].append("The deck's title layout has no usable title box; the title was placed in the library's frame")
    sz = _lvl1(lay_sp, "sz", "defRPr") or _lvl1(mas_sp, "sz", "defRPr") or _title_style(master, "sz", "defRPr")
    size = int(sz) / 100 if sz else 44.0
    algn = _lvl1(lay_sp, "algn") or _lvl1(mas_sp, "algn") or _title_style(master, "algn") or "l"
    anchor = None
    resolved_anchor = next((sp.find(".//" + A + "bodyPr").get("anchor") for sp in (lay_sp, mas_sp)
                            if sp is not None and sp.find(".//" + A + "bodyPr") is not None
                            and sp.find(".//" + A + "bodyPr").get("anchor")), "t")
    x, y, w, h = box
    if size > HOUSE_TITLE_MAX * k:
        size = round(HOUSE_TITLE * k * 2) / 2
        if frame is None:
            h = min(h, round((2 * size * 1.2 / 72 + 0.1) * E))
        anchor = "t"
    if frame is not None:
        # the side-title column heads a left-aligned list: its title is left-aligned
        # there, also in a deck whose titles are centred (CEO, 14 Nov 2026)
        anchor = "t"
        algn = "l"
    # measured with the width tables in the deck's heading face
    fonts, unknown = deck_fonts(prs)
    face = fonts["heading"]
    margin = UNKNOWN_FACE["width"] if "heading" in unknown else 1.0
    m = layoutfit.Measure()
    room_w = (w / E - 0.2) * 72
    floor = max(round(size * 0.75 * 2) / 2, 12.0)

    def need(pt):
        lines = layoutfit._wrap(text, room_w / margin, face, pt, False, _MeasureRef(m)) if text else 1
        return lines * pt * max(1.2, m.line_height(face)) / 72 + 0.1

    fitted = size
    while need(fitted) > h / E and fitted > floor:
        fitted -= 0.5
    if need(fitted) > h / E:
        report["warnings"].append("The title is long for its box: shorten it")
    if fitted < size:
        report["shrunk"].append("title %g pt to %g pt" % (size, fitted))
    title_ph.left, title_ph.top, title_ph.width, title_ph.height = x, y, w, h
    tf = title_ph.text_frame
    tf.text = text
    body = tf._txBody.find(A + "bodyPr")
    for child in list(body):
        if child.tag in (A + "normAutofit", A + "spAutoFit", A + "noAutofit"):
            body.remove(child)
    etree.SubElement(body, A + "noAutofit")
    if anchor:
        body.set("anchor", anchor)
    for para in tf.paragraphs:
        para._p.get_or_add_pPr().set("algn", algn)
        for run in para.runs:
            run._r.get_or_add_rPr().set("sz", str(int(round(fitted * 100))))
    # the box's bottom, and the bottom of the title's text where its anchor puts it
    text_h = round(need(fitted) * E)
    a = anchor or resolved_anchor
    text_bottom = y + text_h if a == "t" else y + h if a == "b" else y + (h + text_h) // 2
    return y + h, text_bottom


class _MeasureRef:
    """layoutfit._wrap reads widths through `.m`."""

    def __init__(self, m):
        self.m = m


def build_slide(prs, layout, entry, page, report, fit=True, fit_first=False, dump=None, label_half_step=False):
    slide = prs.slides.add_slide(layout)
    title_ph = None
    for ph in list(slide.placeholders):
        t = ph.placeholder_format.type
        if t in TITLE_TYPES and title_ph is None:
            title_ph = ph
        elif t not in KEEP_PH:
            ph.element.getparent().remove(ph.element)
    title_bottom = title_text = None
    frame = page_title_frame(page)
    if title_ph is not None:
        box_bottom, title_text = place_title(prs, slide, title_ph, entry.get("title") or "", report, frame)
        # a title in the page's own side frame is beside the content, not above it
        title_bottom = None if frame is not None else box_bottom
        if frame is not None:
            title_text = None
    elif entry.get("title"):
        report["warnings"].append("The layout has no title placeholder; the title was not placed")

    content = [s for s in page["shapes"] if s["kind"] != "ph"]
    top = min([s["y"] for s in content if "y" in s] + [min(s["y1"], s["y2"]) for s in content if "y1" in s] or [1.6])
    place = pagexml.Place(prs.slide_width, prs.slide_height, top, title_bottom, title_text)
    if place.want is not None and place.y(top) < place.want - 0.03 * 914400 * place.sy:
        report["warnings"].append("The title runs into the page under it: shorten the title")

    slots_by_shape = {s["shape"]: s for s in page["slots"]}
    values = entry.get("slots") or {}
    keep = set(entry.get("keep") or [])
    clear = set(entry.get("clear") or [])

    # first pass: text in, and the size each text needs
    drafts, sizes = [], {}
    for i, s in enumerate(page["shapes"]):
        slot = slots_by_shape.get(i)
        if s["kind"] == "ph":
            drafts.append((i, s, slot, None))
            continue
        if slot is None or s["kind"] != "sp":
            drafts.append((i, s, None, None))
            continue
        sid = slot["id"]
        if sid in clear:
            if not s.get("fill") and not s.get("line"):
                report["cleared"].append(sid)
                continue
            s = fill_text(s, "")
            report["cleared"].append(sid)
        elif sid in values:
            s = fill_text(s, values[sid])
            n = sum(len(p["text"]) for p in paragraphs_of(values[sid]))
            if slot.get("max_chars") and n > 1.5 * slot["max_chars"]:
                report["warnings"].append("%s is long for its box (%d characters, about %d fit)" % (sid, n, slot["max_chars"]))
            report["filled"].append(sid)
        elif sid in keep:
            report["kept"].append(sid)
        size, orig, ok = fitted_size(s, slot["role"], place)
        key = (slot["role"], orig)
        sizes[key] = min(sizes.get(key, place.pt(orig)), size)
        drafts.append((i, s, slot, (key, ok)))

    # one size per role on the page, written into the page in its own points
    scale = min(place.s, place.sy * place.k)
    slots_of = [slot for _, _, slot, _ in drafts]
    page_index = [i for i, _, _, _ in drafts]      # each drawn shape's number in the library page
    drawn_at = [place.pt(f[0][1]) if f is not None else None for _, _, _, f in drafts]

    def sized(step, keep=frozenset()):
        """The page with each text at its shrink-first size; with a step (fit
        first), plain text boxes start from the page's own size and give up
        half a point a step, never below their shrink-first size (those in
        `keep` stay at their shrink-first size)."""
        out = []
        for i, s, slot, fitted in drafts:
            s = copy.deepcopy(s)
            if fitted is not None:
                own, size = place.pt(fitted[0][1]), sizes[fitted[0]]
                if step is not None and layoutfit.is_text_box(s) and len(out) not in keep:
                    size = max(size, own - 0.5 * step)
                if size < own:
                    for p in (s.get("text") or {}).get("paras", []):
                        for r in p.get("runs", []):
                            r["sz"] = size / scale
            out.append(s)
        return out

    # then the library's layout pass: boxes fitted, the page spread and filled
    shapes = sized(None)
    if fit:
        faces = deck_fonts(prs)
        if fit_first:
            # the pass at the page's own sizes first; shrink only what it cannot hold
            steps = max([int(round((d - sizes[f[0]]) * 2)) for d, (_, _, _, f) in zip(drawn_at, drafts) if f is not None] + [0])
            def finals(lst):
                return [[r.get("sz") or 12 for p in (x.get("text") or {}).get("paras", []) for r in p.get("runs", [])] for x in lst]
            today = sized(None)
            judge = lay_out(today, page, place, faces, {"warnings": []}, label_half_step, top, page_index)["measure"]

            o_judge = layoutfit._O(judge, faces[0], layoutfit.BODY, None)

            def lines(lst):
                return [layoutfit.lines_of(x, o_judge) if layoutfit.is_text_box(x) else 0 for x in lst]

            def big(x):
                return max((r.get("sz") or 12 for p in x["text"].get("paras", []) for r in p.get("runs", [])), default=0) >= 18 / scale

            def on_fill(i, lst):
                # a text box whose middle sits on a filled shape (a card, a band): its box is fixed
                x = lst[i]
                cx, cy = x["x"] + x["w"] / 2, x["y"] + x["h"] / 2
                return any(f is not x and f.get("kind") == "sp" and layoutfit.js_truthy(f.get("fill"))
                           and not layoutfit.is_text_box(f) and f["x"] <= cx <= f["x"] + f["w"] and f["y"] <= cy <= f["y"] + f["h"]
                           for f in lst)

            # row peers as drawn: one top, size, weight and colour (layoutfit's own row key)
            def row_key(x):
                r = ((x.get("text") or {}).get("paras") or [{}])[0].get("runs") or [{}]
                return (round(x.get("y", 0) * 50), r[0].get("sz") or 12, bool(r[0].get("b")), r[0].get("color") or "")
            # a kind as the shrink step sizes it: role, and size, weight and colour as drawn
            def kind_key(j):
                x, slot = drafts[j][1], drafts[j][2]
                r = ((x.get("text") or {}).get("paras") or [{}])[0].get("runs") or [{}]
                return ((slot or {}).get("role"), r[0].get("sz") or 12, bool(r[0].get("b")), r[0].get("color") or "")

            drawn = [s for _, s, _, _ in drafts]
            guarded = [i for i, x in enumerate(today) if layoutfit.is_text_box(x) and (big(x) or on_fill(i, today))]
            base_lines = lines(today)

            def with_peers(idx):
                # a text kept at today's size keeps its row and its whole kind with it (Reviewer, 14 Nov 2026)
                rows, kinds = {row_key(drawn[i]) for i in idx}, {kind_key(i) for i in idx}
                return set(idx) | {j for j, x in enumerate(drawn) if layoutfit.is_text_box(x) and (row_key(x) in rows or kind_key(j) in kinds)}

            # a big text that already wraps today never grows (Reviewer: "$24M $24M" rose into the title)
            wraps = with_peers({i for i in guarded if big(today[i]) and base_lines[i] > 1}) if guarded else set()
            for step in range(steps + 1):
                # a guarded text that would gain a line stays at today's size, with its row and kind
                keep = set(wraps)
                for _ in range(6):
                    trial = sized(step, keep)
                    ok = lay_out(trial, page, place, faces, {"warnings": []}, label_half_step, top, page_index)["fits"]
                    grew = {i for i in guarded if lines(trial)[i] > base_lines[i]} - keep
                    if not grew:
                        break
                    keep |= with_peers(grew)
                if not ok and step < steps:
                    continue
                # kept only where it is a plain gain: no text smaller than today, no guarded text on a new line, one larger
                pairs = [(a, b) for xa, xb in zip(finals(today), finals(trial)) for a, b in zip(xa, xb)]
                more = any(lines(trial)[i] > base_lines[i] for i in guarded)
                if ok and not more and all(b >= a - 1e-9 for a, b in pairs) and any(b > a + 1e-9 for a, b in pairs):
                    shapes = sized(step, keep)   # fresh: the pass runs once on what is drawn
                break
        lay_out(shapes, page, place, faces, report, label_half_step, top, page_index)
        if dump is not None:
            dump.append({"id": page["id"], "title": entry.get("title") or "", "shapes": copy.deepcopy(shapes)})
    # what shrank, as the slide carries it once the pass is done (the fill
    # may have given some of it back)
    for s, slot, own in zip(shapes, slots_of, drawn_at):
        if own is None:
            continue
        final = next((r.get("sz") for p in (s.get("text") or {}).get("paras", []) for r in p.get("runs", []) if r.get("sz")), None)
        if final is not None and place.pt(final) < own:
            report["shrunk"].append("%s %g pt to %g pt" % (slot["id"], own, place.pt(final)))

    # draw
    tree = slide.shapes._spTree
    sid_next = new_shape_id(slide)
    for s, slot in zip(shapes, slots_of):
        if s["kind"] == "ph":
            if s.get("type") == "source" and entry.get("source"):
                src = fill_text(s, entry["source"])
                tree.append(etree.fromstring(pagexml.source_xml(src, sid_next, place)))
                sid_next += 1
                report["source"] = "placed"
            continue
        if s["kind"] == "cxn":
            tree.append(etree.fromstring(pagexml.cxn_xml(s, sid_next, place)))
        elif s["kind"] == "sp":
            # text in a coloured shape keeps its box: check it still holds its text
            if slot is not None and not (fit and layoutfit.is_text_box(s)):
                size = next((r.get("sz") for p in (s.get("text") or {}).get("paras", [])
                             for r in p.get("runs", []) if r.get("sz")), 12)
                need, have = needed_height(s, place.pt(size), place)
                if need > have * 1.05:
                    report["warnings"].append("%s may still overflow its box: shorten it" % slot["id"])
            name = "Beautify %s %s" % (page["id"], slot["id"]) if slot else None
            tree.append(etree.fromstring(pagexml.sp_xml(s, sid_next, place, name=name)))
        else:
            continue
        sid_next += 1

    if entry.get("source") and report.get("source") != "placed":
        box = {"kind": "ph", "type": "source", "x": 0.5, "y": 6.76, "w": 11.43, "h": 0.4,
               "text": {"paras": [{"runs": [{"t": entry["source"]}]}]}}
        tree.append(etree.fromstring(pagexml.source_xml(box, sid_next, place)))
        report["source"] = "placed at the bottom left"
    return slide


def copy_notes(src, dst):
    if src.has_notes_slide and src.notes_slide.notes_text_frame is not None:
        text = src.notes_slide.notes_text_frame.text
        if text.strip():
            dst.notes_slide.notes_text_frame.text = text
            return True
    return False


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("deck")
    ap.add_argument("plan")
    ap.add_argument("out")
    ap.add_argument("--report")
    ap.add_argument("--keep-originals", action="store_true")
    ap.add_argument("--no-fit", action="store_true", help="draw the pages as the library drew them, without the layout pass")
    ap.add_argument("--fit-first", action="store_true", help="EXPERIMENT (off by default): run the layout pass at the page's own sizes and shrink only what it cannot hold")
    ap.add_argument("--label-half-step", action="store_true", help="OPTION (off by default): a smaller label keeps at least half of its drawn step under the larger text above it")
    ap.add_argument("--dump-shapes", help="write each rebuilt page's final shapes (page inches) to this JSON file, for the gallery renderer")
    ap.add_argument("--pages", default=DEFAULT_PAGES, help="the page catalogue (default: the skill's sample)")
    a = ap.parse_args()
    if os.path.abspath(a.deck) == os.path.abspath(a.out):
        sys.exit("Write to a new file: the deck itself is never changed")

    pages = load_pages(a.pages)
    with open(a.plan, encoding="utf-8") as fh:
        plan = json.load(fh)
    prs = Presentation(a.deck)
    originals = list(prs.slides)
    entries = sorted(plan.get("pages") or [], key=lambda e: e.get("slide", 0))
    seen = set()
    try:
        for e in entries:
            if e.get("page") not in pages:
                raise PlanError("Slide %s: no page %r in the catalogue" % (e.get("slide"), e.get("page")))
            if e["slide"] in seen:
                raise PlanError("Slide %d is in the plan twice" % e["slide"])
            seen.add(e["slide"])
            check_entry(e, pages[e["page"]], len(originals))
        layout = title_only_layout(prs)
    except PlanError as err:
        sys.exit("Plan: %s" % err)

    sld_ids = prs.slides._sldIdLst
    reports, built = [], []
    dumped = [] if a.dump_shapes else None
    # every new slide first: part names are handed out by count, so no
    # slide may be dropped until all are made
    for e in entries:
        rep = {"slide": e["slide"], "page": e["page"], "filled": [], "kept": [], "cleared": [],
               "shrunk": [], "layout": [], "text_up": 0, "warnings": [], "source": None}
        orig = originals[e["slide"] - 1]
        new = build_slide(prs, layout, e, pages[e["page"]], rep, fit=not a.no_fit, fit_first=a.fit_first, dump=dumped, label_half_step=a.label_half_step)
        rep["notes"] = "kept" if copy_notes(orig, new) else None
        built.append((orig, sld_ids[-1], rep))
        reports.append(rep)
    for orig, new_el, rep in built:
        sld_ids.remove(new_el)
        orig_el = next(el for el in sld_ids if prs.part.related_part(el.rId) is orig.part)
        sld_ids.insert(list(sld_ids).index(orig_el), new_el)
        if a.keep_originals:
            orig._element.set("show", "0")
            rep["original"] = "kept after it, hidden"
        else:
            prs.part.drop_rel(orig_el.rId)
            sld_ids.remove(orig_el)
            rep["original"] = "replaced"

    prs.save(a.out)
    if a.dump_shapes:
        with open(a.dump_shapes, "w", encoding="utf-8") as fh:
            json.dump({"slides": dumped}, fh, ensure_ascii=True)
    out = {"deck": os.path.basename(a.deck), "out": os.path.basename(a.out), "rebuilt": reports,
           "unchanged": [n for n in range(1, len(originals) + 1) if n not in seen]}
    text = json.dumps(out, ensure_ascii=False, indent=1)
    if a.report:
        with open(a.report, "w", encoding="utf-8") as fh:
            fh.write(text + "\n")
    print(text)


if __name__ == "__main__":
    main()
