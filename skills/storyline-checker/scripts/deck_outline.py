#!/usr/bin/env python3
"""Print the outline of a PowerPoint deck: each slide's title, in order,
with a short preview of the slide's other text.

Usage:
    python3 deck_outline.py DECK.pptx [--body-chars N] [--json]

What it does and doesn't do:
- Reads only the file you pass on the command line.
- Uses only the Python standard library (3.8 or later).
- Writes nothing to disk and makes no network calls.
- Works on .pptx, .pptm, .potx and Google Slides exports to .pptx.
  Old binary .ppt files are not supported: save them as .pptx first.
"""

import argparse
import json
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
R_ID = "{%s}id" % NS["r"]
OFFICE_DOC_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
)
TITLE_TYPES = {"title", "ctrTitle"}
# Reviewer comments, to-dos and footers are not headlines.
NOTE_START = re.compile(r"^\s*(@|WIP\b|TBD\b|TODO\b|Notes?\s*:|Sources?\s*:)", re.I)
SKIP_TYPES = {"sldNum", "dt", "ftr", "hdr"}
MAX_PART_BYTES = 50 * 1024 * 1024  # refuse absurdly large XML parts


class DeckError(Exception):
    """A problem with the input file that the user can act on."""


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
    """Resolve a relationship target against the folder of the part that owns it."""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base_dir, target))


def read_rels(package, part_name):
    """Map relationship ids to package paths for one part."""
    folder, file_name = posixpath.split(part_name)
    rels_name = posixpath.join(folder, "_rels", file_name + ".rels")
    root = read_xml(package, rels_name)
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


def paragraphs(element):
    """Return the non-empty paragraphs of text inside an element."""
    result = []
    for para in element.iter("{%s}p" % NS["a"]):
        parts = []
        for node in para:
            kind = local(node.tag)
            if kind in ("r", "fld"):
                text_node = node.find("a:t", NS)
                if text_node is not None and text_node.text:
                    parts.append(text_node.text)
            elif kind == "br":
                parts.append(" ")
        text = " ".join("".join(parts).split())
        if text:
            result.append(text)
    return result


def placeholder_type(shape):
    ph = shape.find("p:nvSpPr/p:nvPr/p:ph", NS)
    if ph is None:
        return None
    return ph.get("type", "obj")


def top_offset(shape):
    off = shape.find("p:spPr/a:xfrm/a:off", NS)
    if off is None:
        return None
    try:
        return int(off.get("y", ""))
    except ValueError:
        return None


def walk_shapes(tree, found):
    """Collect title candidates, body text and object markers from a shape tree."""
    for shape in tree:
        kind = local(shape.tag)
        if kind == "sp":
            ph_type = placeholder_type(shape)
            if ph_type in SKIP_TYPES:
                continue
            text = paragraphs(shape)
            if ph_type in TITLE_TYPES:
                found["titles"].append(" ".join(text))
            elif text:
                found["texts"].append((top_offset(shape), ph_type, text))
        elif kind == "grpSp":
            walk_shapes(shape, found)
        elif kind == "graphicFrame":
            data = shape.find("a:graphic/a:graphicData", NS)
            uri = data.get("uri", "") if data is not None else ""
            if uri.endswith("/table"):
                rows = []
                for row in shape.iter("{%s}tr" % NS["a"]):
                    cells = [" ".join(paragraphs(cell)) for cell in row.findall("a:tc", NS)]
                    rows.append(" / ".join(c for c in cells if c))
                found["texts"].append((top_offset(shape), "table", [r for r in rows if r]))
                found["objects"].append("table")
            elif "chart" in uri:
                found["objects"].append("chart")
            elif "diagram" in uri:
                found["objects"].append("SmartArt")
            else:
                found["objects"].append("object")
        elif kind == "pic":
            found["objects"].append("picture")


def read_slide(package, part_name, slide_height):
    root = read_xml(package, part_name)
    if root is None:
        raise DeckError("slide part %s is missing" % part_name)
    found = {"titles": [], "texts": [], "objects": []}
    tree = root.find("p:cSld/p:spTree", NS)
    if tree is not None:
        walk_shapes(tree, found)

    title = next((t for t in found["titles"] if t), "")
    title_source = "title placeholder" if title else ""
    texts = list(found["texts"])

    if not title:
        # Hand-built slides often use a plain text box as the title. Take the
        # highest short text box in the top third of the slide as a guess.
        candidates = [
            item for item in texts
            if item[0] is not None
            and item[1] in (None, "obj", "body")
            and item[0] <= slide_height * 0.35
            and len(" ".join(item[2])) <= 200
        ]
        if candidates:
            best = min(candidates, key=lambda item: item[0])
            title = " ".join(best[2])
            title_source = "top text box (guessed)"
            texts.remove(best)

    # Many consulting templates keep a short label in the title box (a section
    # or topic name) and write the takeaway in a text box just below it. When
    # the title is that short, the highest sentence-length text box in the top
    # third of the slide is reported as the headline.
    headline = ""
    if title and len(title.split()) <= 8:
        below = [
            item for item in texts
            if item[0] is not None
            and item[1] in (None, "obj", "body", "subTitle")
            and item[0] <= slide_height * 0.35
            and len(" ".join(item[2]).split()) >= 8
            and not NOTE_START.match(" ".join(item[2]))
        ]
        if below:
            best = min(below, key=lambda item: item[0])
            headline = " ".join(best[2])
            texts.remove(best)

    ordered = sorted(texts, key=lambda item: (item[0] is None, item[0] or 0))
    body = ["; ".join(item[2]) for item in ordered]
    return {
        "hidden": root.get("show") == "0",
        "title": title,
        "headline": headline,
        "title_source": title_source,
        "text": body,
        "objects": sorted(set(found["objects"])),
    }


def outline(path):
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
        slide_height = int(size.get("cy", "6858000")) if size is not None else 6858000
        rels = read_rels(package, pres_name)
        slides = []
        for number, sld_id in enumerate(pres.findall("p:sldIdLst/p:sldId", NS), start=1):
            rel = rels.get(sld_id.get(R_ID))
            if rel is None:
                slides.append({"slide": number, "error": "slide is listed but its part is missing"})
                continue
            try:
                info = read_slide(package, rel[1], slide_height)
            except (DeckError, ET.ParseError) as exc:
                slides.append({"slide": number, "error": "could not read this slide (%s)" % exc})
                continue
            info["slide"] = number
            slides.append(info)
        return slides


def clip(text, limit):
    if limit <= 0 or len(text) <= limit:
        return text
    return text[: max(0, limit - 3)].rstrip() + "..."


def print_text(path, slides, body_chars):
    hidden = sum(1 for s in slides if s.get("hidden"))
    print("Deck: %s (%d slides, %d hidden)" % (path, len(slides), hidden))
    print("")
    for s in slides:
        if "error" in s:
            print("Slide %d: [%s]" % (s["slide"], s["error"]))
            continue
        flags = []
        if s["hidden"]:
            flags.append("hidden")
        if s["title_source"] == "top text box (guessed)":
            flags.append("no title placeholder; title guessed from top text box")
        label = "Slide %d" % s["slide"]
        if flags:
            label += " [" + "; ".join(flags) + "]"
        print("%s: %s" % (label, s["title"] or "(no title)"))
        if s.get("headline"):
            print("    headline: %s" % s["headline"])
        body = clip(" | ".join(s["text"]), body_chars)
        if body:
            print("    text: %s" % body)
        if s["objects"]:
            print("    contains: %s" % ", ".join(s["objects"]))


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Print each slide's title in order, with a preview of its other text."
    )
    parser.add_argument("deck", help="path to a .pptx file")
    parser.add_argument(
        "--body-chars", type=int, default=300,
        help="characters of other slide text to show per slide (0 for all, default 300)",
    )
    parser.add_argument("--json", action="store_true", help="print JSON instead of text")
    args = parser.parse_args(argv)

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    try:
        slides = outline(args.deck)
    except DeckError as exc:
        print("Error: %s" % exc, file=sys.stderr)
        return 2

    if args.json:
        print(json.dumps({"deck": args.deck, "slides": slides}, ensure_ascii=False, indent=2))
    else:
        print_text(args.deck, slides, args.body_chars)
    return 0


if __name__ == "__main__":
    sys.exit(main())
