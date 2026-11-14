#!/usr/bin/env python3
"""Does every text box the layout pass fitted hold its text, on the slide as saved?

    python3 tools/check_slide_fit.py --run WORKDIR      # build decks, beautify them, check
    python3 tools/check_slide_fit.py OUT.pptx [...]     # check decks already beautified

Works from the saved deck alone, not from the pass's own model: for every
shape written with Resize Shape to Fit Text (the boxes the pass fitted), the
text is measured at the size, box width, insets, indents and spacing written
into the slide, in the deck's theme face, with the skill's width tables, and
its height is set against the box's. A face the tables lack is reported, not
measured. --run builds four decks (16:9 and 4:3, each also with a title that
ends low), fills every catalogue page with its own sample words at half,
full and 1.8 times their length, and runs beautify_deck.py on each.
"""

import json
import os
import subprocess
import sys

from lxml import etree
from pptx import Presentation
from pptx.util import Inches

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(HERE, "..", "skills", "slide-beautify", "scripts")
# the beautifier to run (another checkout, for a before and after); measured with this one's tables
RUN_SCRIPTS = os.environ.get("CHECK_SCRIPTS", SCRIPTS)
sys.path.insert(0, SCRIPTS)
import layoutfit  # noqa: E402

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
EMU = 914400.0
CHECK = "check_slide_fit"


def theme_faces(prs):
    part = prs.slide_master.part.part_related_by(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
    root = etree.fromstring(part.blob)
    get = lambda tag: (root.xpath(".//a:fontScheme/a:%s/a:latin/@typeface" % tag, namespaces=NS) or ["Arial"])[0]
    return {"+mj-lt": get("majorFont"), "+mn-lt": get("minorFont")}


def need_and_lines(sp, faces, m):
    """(inches the text needs, lines, faces not in the tables) for one shape."""
    ext = sp.find(".//a:xfrm/a:ext", NS)
    w = int(ext.get("cx")) / EMU
    body = sp.find(".//p:txBody/a:bodyPr", NS)
    ins = [int(body.get(k, d)) / EMU for k, d in (("lIns", 91440), ("tIns", 45720), ("rIns", 91440), ("bIns", 45720))]
    nowrap = body.get("wrap") == "none"
    pts, lines, missing = 0.0, 0, set()
    for p in sp.findall(".//p:txBody/a:p", NS):
        ppr = p.find("a:pPr", NS)
        mar = int(ppr.get("marL", 0)) / EMU if ppr is not None else 0
        spacing = 1.0
        before = after = 0.0
        if ppr is not None:
            pct = ppr.find("a:lnSpc/a:spcPct", NS)
            spacing = int(pct.get("val")) / 100000 if pct is not None else 1.0
            b = ppr.find("a:spcBef/a:spcPts", NS)
            a = ppr.find("a:spcAft/a:spcPts", NS)
            before = int(b.get("val")) / 100 if b is not None else 0
            after = int(a.get("val")) / 100 if a is not None else 0
        runs = p.findall("a:r", NS)
        text = "".join((r.findtext("a:t", default="", namespaces=NS) or "") for r in runs)
        rpr = [r.find("a:rPr", NS) for r in runs]
        sizes = [int(x.get("sz")) / 100 for x in rpr if x is not None and x.get("sz")]
        end = p.find("a:endParaRPr", NS)
        size = sizes[0] if sizes else (int(end.get("sz")) / 100 if end is not None and end.get("sz") else 18)
        refs = [x.find("a:latin", NS).get("typeface") if x is not None and x.find("a:latin", NS) is not None else "+mn-lt" for x in rpr]
        face = faces.get(refs[0] if refs else "+mn-lt", refs[0] if refs else "Arial")
        if face not in m.fonts:
            missing.add(face)
            face = "Arial"
        bold = bool(rpr) and all(x is not None and x.get("b") == "1" for x in rpr)
        width = max((w - ins[0] - ins[2] - mar) * 72, 4)
        n = 1 if nowrap else layoutfit._wrap(text, width, face, size, bold, _M(m))
        lines += n
        pts += n * size * m.line_height(face) * spacing + before + after
    return pts / 72 + ins[1] + ins[3], lines, missing


class _M:
    def __init__(self, m):
        self.m = m


def check(path, m):
    prs = Presentation(path)
    faces = theme_faces(prs)
    out = {"boxes": 0, "taller": 0, "worst": -1.0, "missing": set(), "where": []}
    for n, slide in enumerate(prs.slides, 1):
        for sp in slide.shapes._spTree.iter("{%s}sp" % NS["p"]):
            if sp.find(".//p:txBody/a:bodyPr/a:spAutoFit", NS) is None:
                continue
            h = int(sp.find(".//a:xfrm/a:ext", NS).get("cy")) / EMU
            need, lines, missing = need_and_lines(sp, faces, m)
            out["missing"] |= missing
            out["boxes"] += 1
            over = need - h
            if over > 0.01:
                out["taller"] += 1
                out["where"].append("slide %d %s: %d lines, %.3f in over" % (n, sp.find(".//p:cNvPr", NS).get("name"), lines, over))
            out["worst"] = max(out["worst"], over)
    return out


def set_theme_faces(prs, heading, body):
    part = prs.slide_master.part.part_related_by(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
    root = etree.fromstring(part.blob)
    for tag, face in (("majorFont", heading), ("minorFont", body)):
        for el in root.xpath(".//a:fontScheme/a:%s/a:latin" % tag, namespaces=NS):
            el.set("typeface", face)
    part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def make_deck(path, wide, low_title, n, faces=None):
    prs = Presentation()
    if faces:
        set_theme_faces(prs, *faces)
    if wide:
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    lay = prs.slide_layouts[5]
    if low_title:
        for ph in lay.placeholders:
            if "TITLE" in str(ph.placeholder_format.type):
                ph.top, ph.height = Inches(0.9), Inches(1.4)
    for i in range(n):
        prs.slides.add_slide(lay).shapes.title.text = "Slide %d" % (i + 1)
    prs.save(path)


def stretch(text, f):
    w = text.split(" ")
    if f < 1:
        return " ".join(w[:max(1, round(len(w) * f))])
    if f > 1:
        return text + " " + " ".join(w[:max(1, round(len(w) * (f - 1)))])
    return text


def sample_of(page, slot):
    s = page["shapes"][slot["shape"]]
    paras = [("".join(str(r.get("t") or "") for r in p.get("runs", []))) for p in (s.get("text") or {}).get("paras", [])]
    paras = [p for p in paras if p.strip()] or ["Text"]
    return paras


def plan_for(pages, f):
    out = []
    for i, p in enumerate(pages):
        slots = {}
        for s in p["slots"]:
            if s["role"] in ("fixed", "title", "source"):
                continue
            paras = [stretch(x, f) for x in sample_of(p, s)]
            slots[s["id"]] = paras if len(paras) > 1 else paras[0]
        out.append({"slide": i + 1, "page": p["id"], "title": "A title for " + p["id"], "slots": slots, "source": "Source: Company data"})
    return {"pages": out}


def main():
    m = layoutfit.Measure()
    if sys.argv[1:2] == ["--run"]:
        work = sys.argv[2]
        os.makedirs(work, exist_ok=True)
        with open(os.path.join(SCRIPTS, "..", "assets", "pages.json"), encoding="utf-8") as fh:
            pages = json.load(fh)["pages"]
        paths = []
        shapes = (("16x9", True, False), ("4x3", False, False), ("16x9-low-title", True, True), ("4x3-low-title", False, True))
        themes = (("calibri", None), ("arial", ("Georgia", "Arial")), ("gothic", ("Century Gothic", "Century Gothic")))
        for (sname, wide, low), (tname, faces) in [(a, b) for b in themes for a in shapes]:
            name = "%s-%s" % (tname, sname)
            deck = os.path.join(work, name + ".pptx")
            make_deck(deck, wide, low, len(pages), faces)
            for f in (0.5, 1.0, 1.8):
                plan = os.path.join(work, "%s-%s.json" % (name, f))
                with open(plan, "w", encoding="utf-8") as fh:
                    json.dump(plan_for(pages, f), fh)
                out = os.path.join(work, "%s-%s-out.pptx" % (name, f))
                r = subprocess.run([sys.executable, os.path.join(RUN_SCRIPTS, "beautify_deck.py"), deck, plan, out], capture_output=True, text=True)
                if r.returncode:
                    sys.exit(r.stderr)
                paths.append(out)
    else:
        paths = sys.argv[1:]
    total = 0
    for path in paths:
        r = check(path, m)
        total += r["taller"]
        print("%-36s boxes %4d | text taller than its box %3d | worst %+.3f in%s" % (
            os.path.basename(path), r["boxes"], r["taller"], r["worst"],
            " | faces without tables: " + ", ".join(sorted(r["missing"])) if r["missing"] else ""))
        for w in r["where"][:5]:
            print("    " + w)
    return total


if __name__ == "__main__":
    sys.exit(1 if main() else 0)
