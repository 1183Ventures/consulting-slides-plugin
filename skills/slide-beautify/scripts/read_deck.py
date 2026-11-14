#!/usr/bin/env python3
"""List what is on each slide of a deck, for the slide-beautify skill.

Usage:
    python3 read_deck.py DECK.pptx [--slides 2,4-6] [--pages]

Prints JSON: the deck's slide size and theme (fonts and colours), and for
each slide its title, its text boxes (paragraphs with their levels and any
bold lead-in), its tables, what else is on it (pictures, charts, SmartArt,
groups, media), its source line and its speaker notes, and whether the
beautify step can rebuild it ("rebuild": true) or must leave it alone, with
the reason.

--pages also prints the page catalogue (assets/pages.json without the
drawing data): each page's id, what it fits, and its text slots.

Reads only the file you pass. Needs python-pptx (pip install python-pptx).
"""

import argparse
import json
import os
import re
import sys

try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
    from pptx.util import Emu
except ImportError:
    sys.exit("This needs python-pptx: pip install python-pptx")

HERE = os.path.dirname(os.path.abspath(__file__))
PAGES = os.path.join(HERE, "..", "assets", "pages.json")
SOURCE = re.compile(r"^\s*(sources?|note|notes)\s*[:.]", re.I)
NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
TITLES = (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)
SKIP_PH = (PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.HEADER)


def slide_list(spec, count):
    if not spec:
        return list(range(1, count + 1))
    out = []
    for part in spec.split(","):
        if "-" in part:
            a, b = part.split("-", 1)
            out += list(range(int(a), int(b) + 1))
        elif part.strip():
            out.append(int(part))
    return [n for n in out if 1 <= n <= count]


def para_info(p):
    text = "".join(r.text for r in p.runs).strip()
    if not text:
        return None
    info = {"text": text, "level": p.level}
    runs = [r for r in p.runs if r.text.strip()]
    if len(runs) > 1 and runs[0].font.bold and not all(r.font.bold for r in runs):
        info["lead"] = runs[0].text.strip()
    return info


def inches(v):
    return round(Emu(v).inches, 2)


def theme_of(prs):
    master = prs.slide_masters[0]
    for rel in master.part.rels.values():
        if rel.reltype.endswith("/theme"):
            from lxml import etree
            root = etree.fromstring(rel.target_part.blob)
            out = {"colors": {}, "fonts": {}}
            scheme = root.find(".//%sclrScheme" % NS_A)
            if scheme is not None:
                for el in scheme:
                    tag = el.tag.replace(NS_A, "")
                    c = el[0] if len(el) else None
                    if c is not None:
                        out["colors"][tag] = c.get("val") or c.get("lastClr")
            for kind in ("majorFont", "minorFont"):
                f = root.find(".//%s%s/%slatin" % (NS_A, kind, NS_A))
                if f is not None:
                    out["fonts"]["heading" if kind == "majorFont" else "body"] = f.get("typeface")
            return out
    return {}


def read_slide(slide, n):
    out = {"slide": n, "layout": slide.slide_layout.name, "hidden": slide._element.get("show") == "0",
           "title": None, "texts": [], "tables": [], "other": [], "source": None, "notes": None}
    blockers = []

    def walk(shapes, depth=0):
        for sh in shapes:
            st = sh.shape_type
            if sh.is_placeholder:
                pt = sh.placeholder_format.type
                if pt in SKIP_PH:
                    continue
                if pt in TITLES and sh.has_text_frame and out["title"] is None:
                    out["title"] = sh.text_frame.text.strip() or None
                    continue
            if st == MSO_SHAPE_TYPE.GROUP:
                out["other"].append({"kind": "group", "name": sh.name})
                blockers.append("a group")
                continue
            if sh.has_text_frame and not (st == MSO_SHAPE_TYPE.PLACEHOLDER and not sh.text_frame.text.strip()):
                paras = [x for x in (para_info(p) for p in sh.text_frame.paragraphs) if x]
                if not paras:
                    continue
                if len(paras) == 1 and SOURCE.match(paras[0]["text"]):
                    out["source"] = paras[0]["text"]
                    continue
                out["texts"].append({"name": sh.name, "left": inches(sh.left or 0), "top": inches(sh.top or 0),
                                     "width": inches(sh.width or 0), "height": inches(sh.height or 0), "paras": paras})
                continue
            if getattr(sh, "has_table", False) and sh.has_table:
                rows = [[c.text.strip() for c in r.cells] for r in sh.table.rows]
                out["tables"].append({"name": sh.name, "rows": rows})
                blockers.append("a table (the sample has no table pages)")
                continue
            if getattr(sh, "has_chart", False) and sh.has_chart:
                out["other"].append({"kind": "chart", "name": sh.name})
                blockers.append("a chart")
            elif st == MSO_SHAPE_TYPE.PICTURE or (sh.is_placeholder and sh.placeholder_format.type == PP_PLACEHOLDER.PICTURE):
                out["other"].append({"kind": "picture", "name": sh.name})
                blockers.append("a picture")
            elif st in (MSO_SHAPE_TYPE.MEDIA, MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT, MSO_SHAPE_TYPE.LINKED_OLE_OBJECT):
                out["other"].append({"kind": "media or object", "name": sh.name})
                blockers.append("media or an embedded object")
            elif sh.element.tag.endswith("graphicFrame"):
                out["other"].append({"kind": "smartart or diagram", "name": sh.name})
                blockers.append("SmartArt or a diagram")
            elif st in (MSO_SHAPE_TYPE.LINE, MSO_SHAPE_TYPE.FREEFORM) or sh.element.tag.endswith("cxnSp"):
                out["other"].append({"kind": "line", "name": sh.name})
            else:
                out["other"].append({"kind": "shape", "name": sh.name})

    walk(slide.shapes)
    if slide.has_notes_slide:
        notes = slide.notes_slide.notes_text_frame.text.strip() if slide.notes_slide.notes_text_frame else ""
        out["notes"] = notes or None
    if blockers:
        out["rebuild"] = False
        out["why_not"] = "has " + ", ".join(sorted(set(blockers)))
    elif not out["texts"]:
        out["rebuild"] = False
        out["why_not"] = "no text besides the title"
    else:
        out["rebuild"] = True
    return out


def catalogue():
    with open(PAGES, encoding="utf-8") as fh:
        pages = json.load(fh)["pages"]
    return [{"id": p["id"], "name": p["name"], "fits": p["fits"],
             "slots": [{k: s[k] for k in ("id", "role", "sample", "max_chars", "bullets", "paras") if k in s}
                       for s in p["slots"]]} for p in pages]


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("deck")
    ap.add_argument("--slides")
    ap.add_argument("--pages", action="store_true")
    a = ap.parse_args()
    prs = Presentation(a.deck)
    nums = slide_list(a.slides, len(prs.slides))
    out = {"file": os.path.basename(a.deck), "slide_size_in": [inches(prs.slide_width), inches(prs.slide_height)],
           "theme": theme_of(prs), "slides": [read_slide(prs.slides[n - 1], n) for n in nums]}
    if a.pages:
        out["pages"] = catalogue()
    json.dump(out, sys.stdout, ensure_ascii=False, indent=1)
    sys.stdout.write("\n")


if __name__ == "__main__":
    main()
