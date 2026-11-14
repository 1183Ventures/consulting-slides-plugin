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
# wider than Arial (Poppins), and some set lines up to 1.35 of the size
UNKNOWN_FACE = {"width": 1.15, "line": 1.35}


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


def lay_out(shapes, page, place, deck_faces, report):
    """layoutfit.py on the filled page, in the page's own inches: text boxes
    fitted to their text, rows and cards spread to the body's bottom, body
    text raised while the page has room (the library's own pass)."""
    fonts, unknown = deck_faces
    # text measured as the slide will carry it; a deck not drawn at the
    # library's own scale (4:3, a band squeezed under a low title) keeps 3%
    # more room, which the Reviewer measured covers what one model leaves out
    plain = all(abs(v - 1) < 0.005 for v in (place.sx, place.sy, place.k))
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
    r = layoutfit.fit_layout(shapes, fonts=fonts, measure=measure, page_id=page["id"], max_raise=max_raise)
    fitted = sum(1 for c in r["changed"] if c["what"] == "height to its text")
    report["layout"] = (["%d text boxes fitted to their text" % fitted] if fitted else []) + \
        [c["what"] for c in r["changed"] if c["what"] not in ("height to its text", "body text up to fill the page")]
    report["text_up"] = r["size"]
    if not layoutfit.page_fits(shapes, fonts=fonts, measure=measure):
        report["warnings"].append("Texts on this page may run into each other or past the bottom: shorten the longest")
    return r


def build_slide(prs, layout, entry, page, report, fit=True):
    slide = prs.slides.add_slide(layout)
    title_ph = None
    for ph in list(slide.placeholders):
        t = ph.placeholder_format.type
        if t in TITLE_TYPES and title_ph is None:
            title_ph = ph
        elif t not in KEEP_PH:
            ph.element.getparent().remove(ph.element)
    if title_ph is not None:
        title_ph.text_frame.text = entry.get("title") or ""
    elif entry.get("title"):
        report["warnings"].append("The layout has no title placeholder; the title was not placed")

    content = [s for s in page["shapes"] if s["kind"] != "ph"]
    top = min([s["y"] for s in content if "y" in s] + [min(s["y1"], s["y2"]) for s in content if "y1" in s] or [1.6])
    title_bottom = (title_ph.top + title_ph.height) if title_ph is not None and title_ph.height else None
    place = pagexml.Place(prs.slide_width, prs.slide_height, top, title_bottom)

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
    shapes, slots_of, drawn_at = [], [], []
    for i, s, slot, fitted in drafts:
        s = copy.deepcopy(s)
        drawn_at.append(place.pt(fitted[0][1]) if fitted is not None else None)
        if fitted is not None:
            size = sizes[fitted[0]]
            if size < place.pt(fitted[0][1]):
                for p in (s.get("text") or {}).get("paras", []):
                    for r in p.get("runs", []):
                        r["sz"] = size / scale
        shapes.append(s)
        slots_of.append(slot)
    # then the library's layout pass: boxes fitted, the page spread and filled
    if fit:
        lay_out(shapes, page, place, deck_fonts(prs), report)
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
    # every new slide first: part names are handed out by count, so no
    # slide may be dropped until all are made
    for e in entries:
        rep = {"slide": e["slide"], "page": e["page"], "filled": [], "kept": [], "cleared": [],
               "shrunk": [], "layout": [], "text_up": 0, "warnings": [], "source": None}
        orig = originals[e["slide"] - 1]
        new = build_slide(prs, layout, e, pages[e["page"]], rep, fit=not a.no_fit)
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
    out = {"deck": os.path.basename(a.deck), "out": os.path.basename(a.out), "rebuilt": reports,
           "unchanged": [n for n in range(1, len(originals) + 1) if n not in seen]}
    text = json.dumps(out, ensure_ascii=False, indent=1)
    if a.report:
        with open(a.report, "w", encoding="utf-8") as fh:
            fh.write(text + "\n")
    print(text)


if __name__ == "__main__":
    main()
