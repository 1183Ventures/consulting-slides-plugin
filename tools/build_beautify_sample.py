#!/usr/bin/env python3
"""Build the slide-beautify skill's page sample from the SlideShortcuts library.

Usage:
    python3 tools/build_beautify_sample.py <library dir> [--out skills/slide-beautify/assets/pages.json]

<library dir> holds index.json and pages/*.json: the SlideShortcuts slide
library as the Google Slides add-on serves it (lavender: Website/gslides/library).
Only the pages listed in SAMPLE are copied, each with its text slots worked out.
Davon picks the final 20 before anything ships (CEO, 13 Nov 2026).

Not shipped with the skill: a maintainer's tool.
"""

import argparse
import glob
import json
import math
import os
import re

# page id -> what the page is for, in a line Claude can match a slide against
SAMPLE = {
    "executive-summary": "Executive summary: 5 rows, each a short heading and one or two sentences",
    "situation-complication-resolution": "Situation, complication, resolution: three labelled blocks telling a story",
    "key-takeaways": "Key takeaways: 3 numbered points, each a bold line and a sentence",
    "key-question": "The key question and the answer, with the reasons under it",
    "decisions": "Decisions needed: a list of decisions, each with options or a recommendation",
    "three-pillars": "Three parallel ideas side by side (pillars, levers, themes), each a heading and text",
    "four-columns": "Four parallel ideas in columns, each a heading and short bullets",
    "two-options": "Two options compared side by side",
    "from-to": "From today to the future: pairs of 'from' and 'to' statements",
    "numbered-list": "A numbered list of points, each a bold lead-in and a sentence",
    "big-numbers": "Three or four big figures, each with a label and a line of context",
    "labelled-rows": "Rows of text, each with a label on the left",
    "checklist": "A checklist of items, with ticks",
    "side-title": "Text with a side panel stating the message",
    "matrix-2x2": "A 2x2 matrix: four quadrants, each a heading and a sentence, with two axes",
    "swot": "Strengths, weaknesses, opportunities, threats",
    "chevron-process": "A process of steps left to right (chevrons), each with a short description",
    "vertical-steps": "Steps top to bottom, each a heading and a description",
    "phases-activities": "Phases across the top, the activities of each phase under it",
    "lib-pros-cons": "Pros and cons of one option",
    "issue-tree": "An issue tree: a question broken into sub-questions",
    "plan-30-60-90": "A 30-60-90 day plan: three columns of actions",
    "milestone-timeline": "Milestones along a timeline",
    "lib-messages-key-messages-4": "Four key messages, each with a short explanation",
    "agenda": "An agenda: numbered sections",
    "kpi-tiles": "KPI tiles: a few metrics, each with a value and a label",
}

NUMBERING = re.compile(r"^(\d{1,2}|0\d|[A-Z]|[ivxIVX]{1,4})[.)]?$")


def run_text(shape):
    return "\n".join("".join(r.get("t", "") for r in p.get("runs", [])) for p in (shape.get("text") or {}).get("paras", []))


def first_run(shape):
    for p in (shape.get("text") or {}).get("paras", []):
        for r in p.get("runs", []):
            if r.get("t"):
                return r
    return {}


def capacity(shape, size):
    t = shape.get("text") or {}
    ins = t.get("inset") or [0.08, 0.05, 0.08, 0.05]
    w = max(shape["w"] - ins[0] - ins[2], 0.1) * 72
    h = max(shape["h"] - ins[1] - ins[3], 0.1) * 72
    per_line = max(1, int(w / (size * 0.5)))
    lines = 1 if t.get("wrap") is False else max(1, int(h / (size * 1.2)))
    return per_line, lines


def slots_of(page):
    slots = []
    n = 0
    for i, s in enumerate(page["shapes"]):
        if s["kind"] == "ph":
            if s.get("type") in ("title", "source"):
                slots.append({"id": s["type"], "shape": i, "role": s["type"], "sample": run_text(s)})
            continue
        if s["kind"] != "sp":
            continue
        text = run_text(s).strip()
        if not text:
            continue
        n += 1
        r = first_run(s)
        size = r.get("sz") or 12
        paras = [p for p in s["text"]["paras"] if any(x.get("t") for x in p.get("runs", []))]
        if NUMBERING.match(text):
            role = "fixed"
        elif r.get("b") or size >= 16:
            role = "heading"
        elif size <= 11 and len(text) <= 30:
            role = "label"
        else:
            role = "body"
        per_line, lines = capacity(s, size)
        slot = {"id": "s%d" % n, "shape": i, "role": role, "sample": text,
                "size": size, "chars_per_line": per_line, "lines": lines,
                "max_chars": per_line * lines}
        if any(p.get("bullet") for p in paras):
            slot["bullets"] = True
        if len(paras) > 1:
            slot["paras"] = len(paras)
        slots.append(slot)
    return slots


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("library")
    ap.add_argument("--out", default=os.path.join(os.path.dirname(__file__), "..", "skills", "slide-beautify", "assets", "pages.json"))
    a = ap.parse_args()
    found = {}
    for f in sorted(glob.glob(os.path.join(a.library, "pages", "*.json"))):
        d = json.load(open(f, encoding="utf-8"))
        for p in d["pages"]:
            if p["id"] in SAMPLE:
                found[p["id"]] = (d["section"], p)
    missing = [k for k in SAMPLE if k not in found]
    if missing:
        raise SystemExit("Not in the library: " + ", ".join(missing))
    pages = []
    for pid, fits in SAMPLE.items():
        section, p = found[pid]
        if any(s["kind"] in ("grp", "tbl", "pic") for s in p["shapes"]):
            raise SystemExit(pid + ": groups, tables and pictures are not in the sample")
        pages.append({"id": pid, "name": p["name"], "section": section, "fits": fits,
                      "slots": slots_of(p), "shapes": p["shapes"]})
    out = {"version": 1, "canvas": {"w": 13.333, "h": 7.5}, "source": "SlideShortcuts slide library",
           "pages": pages}
    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    with open(a.out, "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, separators=(",", ":"))
    print("%d pages, %d slots -> %s" % (len(pages), sum(len(p["slots"]) for p in pages), a.out))


if __name__ == "__main__":
    main()
