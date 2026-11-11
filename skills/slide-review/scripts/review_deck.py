#!/usr/bin/env python3
"""Review a PowerPoint deck the way a manager does before it goes out.

Usage:
    python3 review_deck.py DECK.pptx [--json]

It prints, slide by slide, what a careful reviewer would catch:

Numbers and sources
- a number in the title that is nowhere on the page (or that the page
  shows differently, such as "Title says 12%, page shows 11%");
- a chart without a unit or a period anywhere the reader can see them;
- a page with figures but no source line ("Source: ...");
- a table whose total row does not add up.

Leftovers
- template text and unfilled placeholders ("Click to add text", "[X%]",
  "TBD", "Lorem ipsum");
- empty placeholders;
- objects that sit wholly off the slide;
- hidden slides, speaker notes and comments, which travel with the file;
- reviewer sticky notes and Draft stamps still on the page;
- fonts that are not the theme's.

Layout (alignment, spacing, sizes) is not checked here: a box set 2 pt in on
purpose and one that is 2 pt off look the same in the file. Judge layout by
eye, on the rendered slide.

What it does and doesn't do:
- Reads only the file you pass on the command line.
- Uses only the Python standard library (3.8 or later).
- Writes nothing to disk and makes no network calls.
- Works on .pptx, .pptm, .potx and Google Slides exports to .pptx. Old binary
  .ppt files are not supported: save them as .pptx first.
- Does not read text inside pictures or SmartArt. For PowerPoint's own charts
  it reads the chart's titles, axis titles, series names and category labels,
  not the data values, so a title figure that is only in a chart's data is
  not reported as missing.
"""

import argparse
import json
import math
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
R_ID = "{%s}id" % NS["r"]
OFFICE_DOC_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
)
MAX_PART_BYTES = 50 * 1024 * 1024  # refuse absurdly large XML parts
EMU_PER_PT = 12700.0

TITLE_TYPES = {"title", "ctrTitle"}
COVER_TYPES = {"ctrTitle", "subTitle"}
FURNITURE_TYPES = {"dt", "ftr", "sldNum", "hdr"}
CHART_GROUP = "SlideShortcuts chart"
CHECK_NOTE_NAME = "SlideShortcuts page check"
STICKY_FILLS = {"FFFF66", "FFFF00"}


class DeckError(Exception):
    """A problem with the input file that the user can act on."""


# =============================================================================
# Reading numbers and judging a page
#
# The rules and patterns below follow SlideShortcuts' Check Pages, so this
# script and that tool agree on what passes. They match the way JavaScript
# reads text: \w, \d and \b are ASCII, and "space" is JavaScript's \s.
# =============================================================================

S = "[\t\n\x0b\x0c\r    -     　﻿]"
FLAGS = re.ASCII

SCALES = {"k": 1e3, "K": 1e3, "m": 1e6, "M": 1e6, "mn": 1e6, "bn": 1e9, "B": 1e9, "tn": 1e12, "T": 1e12}

# A number as written: an optional currency, digits with thousands commas,
# decimals, and an optional unit or scale right after it. Not a number inside
# a word or a code (Q3, FY24, B2B, COVID-19).
NUMBER = re.compile(
    "(?<![\\w.,$€£¥-])([$€£¥])?" + S + "?(\\d{1,3}(?:,\\d{3})+|\\d+)(?:\\.(\\d+))?"
    "(?:" + S + "?(%|pp|bps|x|×|bn|mn|tn|k|m|K|M|B|T)(?![\\w]))?(?![\\w])",
    FLAGS,
)

# A unit on the page: a currency, a percent, a scale word, a unit word, or a
# number with a scale letter
UNIT = re.compile(
    "%|[$€£¥₹]|(?:^|[^\\w])#|\\b(?:USD|EUR|GBP|JPY|CHF|CAD|AUD|INR|CNY|RMB|units?|number of|no\\. of|count|"
    "people|persons|FTEs?|headcount|hours?|hrs|days|weeks|users|customers|clients|employees|patients|stores|sites|"
    "orders|visits|tons?|tonnes?|kg|km|miles|bps|pp|percent(?:age)?|index|ratio|multiple|score|NPS|thousands?|"
    "millions?|billions?|trillions?|mn|bn)\\b|\\d" + S + "?(?:k|K|M|mn|bn|B)\\b",
    FLAGS | re.IGNORECASE,
)

# A period on the page: a year, a fiscal year, a quarter or half, a month,
# a span such as YTD, or words such as "monthly" or "as of"
PERIOD = re.compile(
    "\\b(?:19|20)\\d{2}\\b|\\b(?:FY|CY)" + S + "?'?\\d{2}(?:\\d{2})?\\b|\\bQ[1-4]\\b|\\bH[12]\\b|"
    "\\b(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|June?|July?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|"
    "Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\\b|\\bMay" + S + "+(?:19|20)?\\d{2}\\b|"
    "\\b(?:YTD|LTM|TTM|MTD|QTD|YoY|MoM|QoQ|CAGR)\\b|'\\d{2}\\b|"
    "\\b(?:weekly|monthly|quarterly|annual(?:ly)?|daily|per (?:year|month|week|day|quarter)|p\\.a\\.|today|as of|"
    "last (?:year|quarter|month|week)|this (?:year|quarter|month|week))\\b",
    FLAGS | re.IGNORECASE,
)

# A source line: a paragraph, or a sentence after a note, that starts
# "Source:" or "Sources:", a footnote number before it allowed
SOURCE = re.compile(
    "(?:^|[\r\n\x0b]|[;.]" + S + ")" + S + "*(?:\\d+[.)]?" + S + "*)?(?:data" + S + "+)?sources?" + S + "*(?:\\(s\\))?"
    + S + "*[:\\-–]",
    FLAGS | re.IGNORECASE,
)

TOTAL_ROW = re.compile("^" + S + "*(?:grand" + S + "+)?total\\b", FLAGS | re.IGNORECASE)
SUBTOTAL_ROW = re.compile("\\b(?:sub-?" + S + "?total|total)\\b", FLAGS | re.IGNORECASE)
PAGE_NUMBER = re.compile("^" + S + "*\\d{1,3}" + S + "*\\Z", FLAGS)


def js_trim(text):
    return re.sub("^" + S + "+|" + S + "+$", "", text)


def numbers_in(text):
    """The numbers in `text`, as written. A plain four-digit number from 1900
    to 2100 is a year: a period, not a figure to find."""
    out = []
    for m in NUMBER.finditer(str(text or "")):
        digits = m.group(2).replace(",", "")
        decimals = len(m.group(3)) if m.group(3) else 0
        try:
            value = float(digits + ("." + m.group(3) if m.group(3) else ""))
        except ValueError:
            continue
        if math.isinf(value):
            continue
        currency = m.group(1) or ""
        unit = m.group(4) or ""
        year = (not currency and not unit and not decimals and "," not in m.group(2)
                and 1900 <= value <= 2100)
        out.append({"raw": js_trim(m.group(0)), "value": value, "decimals": decimals,
                    "currency": currency, "unit": unit, "year": year})
    return out


def telling(n):
    """Whether a title's number is worth looking for on the page: one with a
    currency, a unit or decimals, or a whole number of 10 or more."""
    if n["year"]:
        return False
    return bool(n["currency"] or n["unit"] or n["decimals"] or n["value"] >= 10)


def kind_of(n):
    if n["unit"] in ("%", "pp", "bps"):
        return "share"
    if n["currency"]:
        return "money"
    if n["unit"] in ("x", "×"):
        return "times"
    return "plain"


def same(a, b):
    """Whether `a` (from the title) and `b` (from the page) are the same
    figure, the less precise one rounded: 12% and 12.3%, $4.2M and 4.2 in a
    chart in $M, $1.2B and 1,200 in a table in $M."""
    def half(n, scale):
        return 0.5 * math.pow(10, -n["decimals"]) * scale
    if abs(a["value"] - b["value"]) <= max(half(a, 1), half(b, 1)) + 1e-9:
        return True
    sa = SCALES.get(a["unit"], 1)
    sb = SCALES.get(b["unit"], 1)
    if sa == 1 and sb == 1:
        return False
    for k in range(-4, 5):
        f = math.pow(1000, k)
        tol = max(half(a, sa), half(b, sb * f))
        if abs(a["value"] * sa - b["value"] * sb * f) <= tol * (1 + 1e-9):
            return True
    return False


def line(text):
    """A finding's line, kept under 40 characters."""
    s = str(text)
    return s if len(s) <= 40 else s[:39] + "…"


def totals_off(values):
    """Whether a table's total row does not add up (one finding per table)."""
    rows = [["" if c is None else str(c) for c in (r or [])] for r in (values or [])]
    for r in range(2, len(rows)):
        if not TOTAL_ROW.search(rows[r][0] if rows[r] else ""):
            continue
        for c in range(1, len(rows[r])):
            found = [n for n in numbers_in(rows[r][c]) if not n["year"]]
            if not found:
                continue
            total = found[0]
            # A rate or a multiple in the total row (a margin, a growth rate)
            # is not a sum; a share column totals 100
            if ((total["unit"] == "%" and abs(total["value"] - 100) > 0.5)
                    or total["unit"] in ("x", "×", "pp", "bps")):
                continue
            parts = []
            decimals = total["decimals"]
            for q in range(r):
                label = rows[q][0] if rows[q] else ""
                if SUBTOTAL_ROW.search(label):
                    continue
                cell = rows[q][c] if c < len(rows[q]) else ""
                nums = numbers_in(cell)
                if not nums or nums[0]["year"]:
                    continue
                parts.append(nums[0]["value"])
                decimals = max(decimals, nums[0]["decimals"])
            if len(parts) < 2:
                continue
            tol = 0.5 * math.pow(10, -decimals) * (len(parts) + 1) + 1e-9
            if abs(sum(parts) - total["value"]) > tol:
                return True
    return False


def judge_page(page):
    """What is wrong with one page.

    page = {title, texts, tables, charts, nativeCharts, isTitleSlide,
            chartTexts (optional)}
    -> the reasons, each a short line; none when the page passes.
    """
    reasons = []
    title = str(page.get("title") or "")
    texts = [str(t) for t in (page.get("texts") or [])]
    chart_texts = [str(t) for t in (page.get("chartTexts") or [])]
    table_texts = ["\n".join("\t".join("" if c is None else str(c) for c in (r or [])) for r in (t or []))
                   for t in (page.get("tables") or [])]
    body = texts + table_texts + chart_texts
    native = page.get("nativeCharts") or 0
    charts = (page.get("charts") or 0) + native

    # A number in the title is on the page
    found = [n for t in body for n in numbers_in(t) if not n["year"]]
    missing = [] if page.get("isTitleSlide") else [
        n for n in numbers_in(title) if telling(n) and not any(same(n, p) for p in found)
    ]
    for n in missing[:2]:
        near = sorted(
            [p for p in found if kind_of(p) == kind_of(n) and p["value"] > 0
             and abs(p["value"] - n["value"]) / max(n["value"], p["value"]) <= 0.5],
            key=lambda p: abs(p["value"] - n["value"]),
        )
        if near:
            reasons.append(line("Title says " + n["raw"] + ", page shows " + near[0]["raw"]))
        elif not native:
            reasons.append(line("Title's " + n["raw"] + " is not on the page"))

    # Every chart has a unit and a period, somewhere a reader sees them
    if charts > 0:
        everything = "\n".join([title] + body)
        unit = bool(UNIT.search(everything))
        period = bool(PERIOD.search(everything))
        if not unit and not period:
            reasons.append("Chart needs a unit and period")
        elif not unit:
            reasons.append("Chart needs a unit")
        elif not period:
            reasons.append("Chart needs a period")

    # A table's total row adds up
    if any(totals_off(t) for t in (page.get("tables") or [])):
        reasons.append("Table total does not add up")

    # A page with figures has a source line. A text that is only a small
    # whole number is a page number typed by hand, not a figure.
    figures = charts > 0 or len(page.get("tables") or []) > 0 or any(
        not PAGE_NUMBER.match(t) and any(telling(n) for n in numbers_in(t))
        for t in texts + table_texts
    )
    if figures and not page.get("isTitleSlide") and not any(SOURCE.search(t) for t in texts):
        reasons.append("No source line")

    return reasons


# =============================================================================
# Reading the package
# =============================================================================

def local(tag):
    return tag.rsplit("}", 1)[-1]


def read_xml(package, name):
    """Parse one XML part of the package, or return None if it is missing."""
    try:
        info = package.getinfo(name)
    except KeyError:
        return None
    if info.file_size > MAX_PART_BYTES:
        raise DeckError("part %s is too large to read safely" % name)
    data = package.read(info)
    if b"<!DOCTYPE" in data or b"<!ENTITY" in data:
        # Office files never declare a DOCTYPE; refuse anything that does.
        raise DeckError("part %s has a DOCTYPE, which Office files never use" % name)
    return ET.fromstring(data)


def resolve(base_dir, target):
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base_dir, target))


def read_rels(package, part_name):
    """Map relationship ids to (type, package path) for one part."""
    folder, file_name = posixpath.split(part_name)
    root = read_xml(package, posixpath.join(folder, "_rels", file_name + ".rels"))
    rels = {}
    if root is None:
        return rels
    for rel in root.findall("rel:Relationship", NS):
        if rel.get("TargetMode") == "External":
            continue
        rels[rel.get("Id")] = (rel.get("Type", ""), resolve(folder, rel.get("Target", "")))
    return rels


def find_presentation_part(package):
    root = read_xml(package, "_rels/.rels")
    if root is not None:
        for rel in root.findall("rel:Relationship", NS):
            if rel.get("Type") == OFFICE_DOC_TYPE:
                return resolve("", rel.get("Target", ""))
    return "ppt/presentation.xml"


def text_of(element):
    """A text body's text the way PowerPoint gives it: paragraphs on their
    own lines, a line break inside a paragraph as a vertical tab."""
    paras = []
    for para in element.iter("{%s}p" % NS["a"]):
        parts = []
        for node in para:
            kind = local(node.tag)
            if kind in ("r", "fld"):
                t = node.find("a:t", NS)
                if t is not None and t.text:
                    parts.append(t.text)
            elif kind == "br":
                parts.append("\x0b")
        paras.append("".join(parts))
    return "\n".join(paras)


def fonts_of(element):
    """The Latin typefaces set directly on runs that hold text ("+mj-lt" and
    "+mn-lt" are the theme's own, so they never count)."""
    faces = set()
    for run in element.iter("{%s}r" % NS["a"]):
        t = run.find("a:t", NS)
        node = run.find("a:rPr/a:latin", NS)
        if t is None or not (t.text or "").strip() or node is None:
            continue
        face = node.get("typeface", "")
        if face and not face.startswith("+"):
            faces.add(face)
    return faces


def box_of(xfrm):
    """(left, top, width, height) in points from an a:xfrm or p:xfrm."""
    if xfrm is None:
        return None
    off = xfrm.find("a:off", NS)
    ext = xfrm.find("a:ext", NS)
    if off is None or ext is None:
        return None
    try:
        return (int(off.get("x", "0")) / EMU_PER_PT, int(off.get("y", "0")) / EMU_PER_PT,
                int(ext.get("cx", "0")) / EMU_PER_PT, int(ext.get("cy", "0")) / EMU_PER_PT)
    except ValueError:
        return None


def rotation_of(xfrm):
    try:
        return int(xfrm.get("rot", "0")) if xfrm is not None else 0
    except ValueError:
        return 0


def solid_fill(shape):
    fill = shape.find("p:spPr/a:solidFill/a:srgbClr", NS)
    return fill.get("val", "").upper() if fill is not None else ""


def chart_texts(package, part_name):
    """A PowerPoint chart's visible labels: titles, axis titles, series names,
    category labels and number formats (the data values are not read)."""
    root = read_xml(package, part_name)
    if root is None:
        return []
    out = []
    for title in root.iter("{%s}title" % NS["c"]):
        text = " ".join(t.text for t in title.iter("{%s}t" % NS["a"]) if t.text)
        if text:
            out.append(text)
    for tag in ("tx", "cat"):
        for node in root.iter("{%s}%s" % (NS["c"], tag)):
            for v in node.iter("{%s}v" % NS["c"]):
                if v.text:
                    out.append(v.text)
    for fmt in root.iter("{%s}formatCode" % NS["c"]):
        if fmt.text and re.search(r"%|\$|€|£|¥", fmt.text):
            out.append("%" if "%" in fmt.text else "$")
    return out


def walk(tree, slide_rels, package, items, top=True, in_chart=False):
    """Collect every shape on a slide with what the checks need."""
    for shape in tree:
        kind = local(shape.tag)
        nv = None
        for path in ("p:nvSpPr", "p:nvGrpSpPr", "p:nvGraphicFramePr", "p:nvPicPr", "p:nvCxnSpPr"):
            nv = shape.find(path, NS)
            if nv is not None:
                break
        if nv is None:
            continue
        c_nv = nv.find("p:cNvPr", NS)
        name = c_nv.get("name", "") if c_nv is not None else ""
        ph = nv.find("p:nvPr/p:ph", NS)
        ph_type = (ph.get("type", "obj") if ph is not None else None)
        item = {"kind": kind, "name": name, "ph": ph_type, "top": top, "inChart": in_chart,
                "box": None, "rot": 0, "text": None, "fonts": set(), "fill": "", "prst": ""}
        if kind == "sp":
            xfrm = shape.find("p:spPr/a:xfrm", NS)
            item["box"], item["rot"] = box_of(xfrm), rotation_of(xfrm)
            geom = shape.find("p:spPr/a:prstGeom", NS)
            item["prst"] = geom.get("prst", "") if geom is not None else ""
            item["fill"] = solid_fill(shape)
            body = shape.find("p:txBody", NS)
            if body is not None:
                item["text"] = text_of(body)
                item["fonts"] = fonts_of(body)
            items.append(item)
        elif kind == "grpSp":
            xfrm = shape.find("p:grpSpPr/a:xfrm", NS)
            item["box"], item["rot"] = box_of(xfrm), rotation_of(xfrm)
            chart = in_chart or name == CHART_GROUP
            item["isChart"] = top and name == CHART_GROUP
            items.append(item)
            walk(shape, slide_rels, package, items, top=False, in_chart=chart)
        elif kind == "graphicFrame":
            xfrm = shape.find("p:xfrm", NS)
            item["box"], item["rot"] = box_of(xfrm), rotation_of(xfrm)
            data = shape.find("a:graphic/a:graphicData", NS)
            uri = data.get("uri", "") if data is not None else ""
            if uri.endswith("/table"):
                rows = []
                for tr in shape.iter("{%s}tr" % NS["a"]):
                    rows.append([text_of(tc) for tc in tr.findall("a:tc", NS)])
                item["table"] = rows
                item["fonts"] = fonts_of(shape)
            elif uri.endswith("/chart"):
                item["nativeChart"] = True
                ref = data.find("c:chart", NS)
                rel = slide_rels.get(ref.get(R_ID)) if ref is not None else None
                item["chartTexts"] = chart_texts(package, rel[1]) if rel else []
            elif "diagram" in uri:
                item["smartArt"] = True
            items.append(item)
        elif kind == "pic":
            xfrm = shape.find("p:spPr/a:xfrm", NS)
            item["box"], item["rot"] = box_of(xfrm), rotation_of(xfrm)
            items.append(item)
        elif kind == "cxnSp":
            xfrm = shape.find("p:spPr/a:xfrm", NS)
            item["box"], item["rot"] = box_of(xfrm), rotation_of(xfrm)
            items.append(item)


def theme_fonts(package, slide_part, slide_rels):
    """The theme's heading and body typefaces for the slide's master."""
    layout = next((p for t, p in slide_rels.values() if t.endswith("/slideLayout")), None)
    if not layout:
        return set()
    master = next((p for t, p in read_rels(package, layout).values() if t.endswith("/slideMaster")), None)
    if not master:
        return set()
    theme = next((p for t, p in read_rels(package, master).values() if t.endswith("/theme")), None)
    root = read_xml(package, theme) if theme else None
    if root is None:
        return set()
    faces = set()
    for kind in ("majorFont", "minorFont"):
        node = root.find(".//a:fontScheme/a:%s/a:latin" % kind, NS)
        if node is not None and node.get("typeface"):
            faces.add(node.get("typeface"))
    return faces


def layout_type(package, slide_rels):
    layout = next((p for t, p in slide_rels.values() if t.endswith("/slideLayout")), None)
    root = read_xml(package, layout) if layout else None
    return root.get("type", "") if root is not None else ""


def notes_text(package, slide_rels):
    notes = next((p for t, p in slide_rels.values() if t.endswith("/notesSlide")), None)
    root = read_xml(package, notes) if notes else None
    if root is None:
        return ""
    texts = []
    for shape in root.iter("{%s}sp" % NS["p"]):
        ph = shape.find("p:nvSpPr/p:nvPr/p:ph", NS)
        if ph is not None and ph.get("type") in ("sldNum", "sldImg", "hdr", "ftr", "dt"):
            continue
        body = shape.find("p:txBody", NS)
        if body is not None:
            texts.append(text_of(body))
    return js_trim("\n".join(texts))


def comment_count(package, slide_rels):
    count = 0
    for rel_type, path in slide_rels.values():
        if rel_type.endswith("/comments"):
            root = read_xml(package, path)
            if root is not None:
                count += sum(1 for node in root if local(node.tag) in ("cm",))
    return count


# =============================================================================
# Leftovers
# =============================================================================

TEMPLATE_TEXT = re.compile(
    r"click to (?:add|edit)|lorem ipsum|dolor sit amet|insert (?:text|title|chart|picture|logo)|"
    r"(?:your|add) (?:text|title) here|\bTBD\b|\bTBC\b|\bTODO\b|\bXX+\b|\[(?!\d+\])[^\]\n]{1,30}\]",
    re.IGNORECASE,
)
DRAFT_STAMP = re.compile(r"^\s*(?:draft|draft for discussion|preliminary draft|wip|work in progress)\s*$", re.IGNORECASE)


def leftovers_on(slide, items, slide_w, slide_h, theme_faces):
    found = []
    for it in items:
        text = js_trim(it["text"] or "") if it["text"] is not None else ""
        if it.get("sticky"):
            continue
        # Template text and placeholders still to fill
        if text:
            for m in TEMPLATE_TEXT.finditer(text):
                found.append("Fill in: \"%s\"" % m.group(0)[:30])
                break
            if DRAFT_STAMP.match(text):
                found.append("Draft stamp: \"%s\"" % text[:30])
        # Empty placeholders show their prompt text while editing
        if it["ph"] and it["kind"] == "sp" and it["ph"] not in FURNITURE_TYPES and not text:
            found.append("Empty placeholder (%s)" % (it["name"] or it["ph"]))
        # Wholly off the slide
        if it["top"] and it["box"]:
            x, y, w, h = it["box"]
            if x >= slide_w or y >= slide_h or x + w <= 0 or y + h <= 0:
                found.append("Off the slide: %s" % (it["name"] or it["kind"]))
    for it in items:
        if it.get("sticky"):
            found.append("Sticky note left on: \"%s\"" % js_trim(it["text"] or "")[:30])
    if theme_faces:
        off = sorted(set(f for it in items if not it["inChart"] for f in it["fonts"]) - theme_faces)
        if off:
            found.append("Fonts off the theme: %s" % ", ".join(off[:4]))
    return found


# =============================================================================
# The deck
# =============================================================================

def review(path):
    try:
        package = zipfile.ZipFile(path)
    except FileNotFoundError:
        raise DeckError("file not found: %s" % path)
    except zipfile.BadZipFile:
        raise DeckError(
            "%s is not a .pptx file. Old .ppt files are not supported: save as .pptx first." % path
        )
    with package:
        pres_name = find_presentation_part(package)
        pres = read_xml(package, pres_name)
        if pres is None:
            raise DeckError("no presentation part found in %s" % path)
        size = pres.find("p:sldSz", NS)
        slide_w = int(size.get("cx", "12192000")) / EMU_PER_PT if size is not None else 960.0
        slide_h = int(size.get("cy", "6858000")) / EMU_PER_PT if size is not None else 540.0
        rels = read_rels(package, pres_name)
        slides = []
        for number, sld_id in enumerate(pres.findall("p:sldIdLst/p:sldId", NS), start=1):
            rel = rels.get(sld_id.get(R_ID))
            if rel is None:
                slides.append({"slide": number, "error": "slide is listed but its part is missing"})
                continue
            try:
                slides.append(review_slide(package, rel[1], number, slide_w, slide_h))
            except (DeckError, ET.ParseError) as exc:
                slides.append({"slide": number, "error": "could not read this slide (%s)" % exc})
        return slides


def review_slide(package, part, number, slide_w, slide_h):
    root = read_xml(package, part)
    if root is None:
        raise DeckError("slide part %s is missing" % part)
    slide_rels = read_rels(package, part)
    items = []
    tree = root.find("p:cSld/p:spTree", NS)
    if tree is not None:
        walk(tree, slide_rels, package, items)

    # A reviewer's sticky (a yellow note) or a note Check Pages left is not
    # the page's content
    for it in items:
        if it["top"] and it["kind"] == "sp" and (it["fill"] in STICKY_FILLS or it["name"] == CHECK_NOTE_NAME):
            it["sticky"] = True

    content = [it for it in items if not it.get("sticky") and it["ph"] not in FURNITURE_TYPES]
    texts = [it for it in content if it["text"] is not None and js_trim(it["text"])]
    title = next((it for it in texts if it["top"] and it["ph"] in TITLE_TYPES), None)
    if title is None:
        cands = [it for it in texts if it["top"] and not it["inChart"] and it["box"]
                 and it["box"][1] < slide_h * 0.2 and it["box"][3] < slide_h * 0.25]
        title = min(cands, key=lambda it: it["box"][1]) if cands else None
    cover = any(it["ph"] in COVER_TYPES for it in items) or layout_type(package, slide_rels) == "title"

    page = {
        "title": title["text"] if title else "",
        "texts": [it["text"] for it in texts if it is not title],
        "tables": [it["table"] for it in content if it.get("table")],
        "charts": sum(1 for it in content if it.get("isChart")),
        "nativeCharts": sum(1 for it in content if it.get("nativeChart")),
        "chartTexts": [t for it in content for t in it.get("chartTexts", [])],
        "isTitleSlide": cover,
    }
    notes = notes_text(package, slide_rels)
    return {
        "slide": number,
        "title": js_trim(page["title"]).replace("\n", " ").replace("\x0b", " "),
        "hidden": root.get("show") == "0",
        "cover": cover,
        "checks": judge_page(page),
        "leftovers": leftovers_on(number, items, slide_w, slide_h, theme_fonts(package, part, slide_rels)),
        "notes": bool(notes),
        "comments": comment_count(package, slide_rels),
        "unread": sorted(set(
            ["SmartArt"] * any(it.get("smartArt") for it in items)
            + ["pictures"] * any(it["kind"] == "pic" for it in items)
        )),
    }


def print_text(path, slides):
    ok = [s for s in slides if "error" not in s]
    hidden = [s["slide"] for s in ok if s["hidden"]]
    notes = [s["slide"] for s in ok if s["notes"]]
    comments = [s["slide"] for s in ok if s["comments"]]
    flagged = [s for s in ok if s["checks"] or s["leftovers"]]
    print("Deck: %s (%d slides)" % (path, len(slides)))
    print("Pages to fix: %d of %d" % (len(flagged), len(slides)))
    print("")
    for s in slides:
        if "error" in s:
            print("Slide %d: [%s]" % (s["slide"], s["error"]))
            continue
        if not (s["checks"] or s["leftovers"]):
            continue
        tags = ["hidden"] if s["hidden"] else []
        print("Slide %d%s: %s" % (s["slide"], " [" + ", ".join(tags) + "]" if tags else "",
                                   s["title"] or "(no title)"))
        for r in s["checks"]:
            print("    check: %s" % r)
        for r in s["leftovers"]:
            print("    leftover: %s" % r)
    print("")
    print("Travels with the file:")
    print("    hidden slides: %s" % (", ".join(map(str, hidden)) or "none"))
    print("    speaker notes on slides: %s" % (", ".join(map(str, notes)) or "none"))
    print("    comments on slides: %s" % (", ".join(map(str, comments)) or "none"))
    unread = sorted(set((s["slide"], u) for s in ok for u in s["unread"]))
    if unread:
        print("Not read (check by eye): %s" % ", ".join("slide %d %s" % u for u in unread))


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Review a deck before it goes out: numbers, sources and leftovers."
    )
    parser.add_argument("deck", help="path to a .pptx file")
    parser.add_argument("--json", action="store_true", help="print JSON instead of text")
    args = parser.parse_args(argv)

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    try:
        slides = review(args.deck)
    except DeckError as exc:
        print("Error: %s" % exc, file=sys.stderr)
        return 2

    if args.json:
        print(json.dumps({"deck": args.deck, "slides": slides}, ensure_ascii=False, indent=2))
    else:
        print_text(args.deck, slides)
    return 0


if __name__ == "__main__":
    sys.exit(main())
