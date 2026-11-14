#!/usr/bin/env python3
"""Tests for the slide-beautify scripts: python3 tools/test_beautify.py

Builds small decks with python-pptx (16:9 and 4:3), runs read_deck.py and
beautify_deck.py on them, and checks the output: slides in place, words
carried over, notes kept, the original untouched, bad plans refused.
"""

import hashlib
import json
import os
import subprocess
import sys
import tempfile
import unittest

from pptx import Presentation
from pptx.util import Inches

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "skills", "slide-beautify", "scripts")
READ = os.path.join(ROOT, "read_deck.py")
BEAUTIFY = os.path.join(ROOT, "beautify_deck.py")

SUMMARY = [("Context", "The market grew 6% a year since 2021, but our share fell two points"),
           ("What we found", "Most of the lost share sits in two segments where rivals cut prices"),
           ("What it means", "Winning back half the lost share is worth about $40M a year"),
           ("What we recommend", "Reprice the core range and double coverage in the two segments"),
           ("What it takes", "A 12-week pilot in one region and a team of six")]


def make_deck(path, wide=True):
    prs = Presentation()
    if wide:
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s = prs.slides.add_slide(prs.slide_layouts[0])
    s.shapes.title.text = "Project Atlas"
    s = prs.slides.add_slide(prs.slide_layouts[1])
    s.shapes.title.text = "Summary"
    tf = s.placeholders[1].text_frame
    tf.text = "%s: %s" % SUMMARY[0]
    for k, v in SUMMARY[1:]:
        tf.add_paragraph().text = "%s: %s" % (k, v)
    s.notes_slide.notes_text_frame.text = "Speak to the pilot timing."
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "Closing"
    prs.save(path)


def plan_for(slide=2):
    slots = {}
    for i, (k, v) in enumerate(SUMMARY):
        slots["s%d" % (2 * i + 1)] = k
        slots["s%d" % (2 * i + 2)] = v
    return {"pages": [{"slide": slide, "page": "executive-summary", "title": "Summary",
                       "slots": slots, "source": "Source: Company data"}]}


def run(*args):
    return subprocess.run([sys.executable] + list(args), capture_output=True, text=True)


def digest(path):
    return hashlib.sha256(open(path, "rb").read()).hexdigest()


class Beautify(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()
        self.deck = os.path.join(self.dir, "deck.pptx")
        self.out = os.path.join(self.dir, "out.pptx")
        self.plan = os.path.join(self.dir, "plan.json")

    def write_plan(self, plan):
        with open(self.plan, "w") as fh:
            json.dump(plan, fh)

    def check_output(self, wide):
        make_deck(self.deck, wide)
        before = digest(self.deck)
        self.write_plan(plan_for())
        r = run(BEAUTIFY, self.deck, self.plan, self.out)
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertEqual(digest(self.deck), before, "the deck itself must not change")
        report = json.loads(r.stdout)
        self.assertEqual(report["unchanged"], [1, 3])
        self.assertFalse(report["rebuilt"][0]["warnings"])
        prs = Presentation(self.out)
        self.assertEqual(len(prs.slides), 3)
        new = prs.slides[1]
        text = " ".join(sh.text_frame.text for sh in new.shapes if sh.has_text_frame)
        for k, v in SUMMARY:
            self.assertIn(k, text)
            self.assertIn(v, text)
        self.assertIn("Source: Company data", text)
        self.assertEqual(new.notes_slide.notes_text_frame.text, "Speak to the pilot timing.")
        self.assertEqual(prs.slides[2].shapes.title.text, "Closing")
        w, h = prs.slide_width, prs.slide_height
        for sh in new.shapes:
            if sh.width:
                self.assertGreaterEqual(sh.left, 0)
                self.assertLessEqual(sh.left + sh.width, w + 12700)
                self.assertLessEqual(sh.top + sh.height, h + 12700)

    def test_wide_deck(self):
        self.check_output(True)

    def test_four_by_three_deck(self):
        self.check_output(False)

    def test_read_deck(self):
        make_deck(self.deck)
        r = run(READ, self.deck, "--pages")
        self.assertEqual(r.returncode, 0, r.stderr)
        d = json.loads(r.stdout)
        self.assertEqual([s["rebuild"] for s in d["slides"]], [False, True, False])
        self.assertTrue(any(p["id"] == "executive-summary" for p in d["pages"]))

    def test_keep_originals(self):
        make_deck(self.deck)
        self.write_plan(plan_for())
        r = run(BEAUTIFY, self.deck, self.plan, self.out, "--keep-originals")
        self.assertEqual(r.returncode, 0, r.stderr)
        prs = Presentation(self.out)
        self.assertEqual(len(prs.slides), 4)
        self.assertEqual(prs.slides[2]._element.get("show"), "0")

    def test_open_slot_refused(self):
        make_deck(self.deck)
        plan = plan_for()
        del plan["pages"][0]["slots"]["s4"]
        self.write_plan(plan)
        r = run(BEAUTIFY, self.deck, self.plan, self.out)
        self.assertNotEqual(r.returncode, 0)
        self.assertIn("still open: s4", r.stderr)
        self.assertFalse(os.path.exists(self.out))

    def test_unknown_page_refused(self):
        make_deck(self.deck)
        plan = plan_for()
        plan["pages"][0]["page"] = "no-such-page"
        self.write_plan(plan)
        r = run(BEAUTIFY, self.deck, self.plan, self.out)
        self.assertNotEqual(r.returncode, 0)

    def test_never_overwrites(self):
        make_deck(self.deck)
        self.write_plan(plan_for())
        r = run(BEAUTIFY, self.deck, self.plan, self.deck)
        self.assertNotEqual(r.returncode, 0)

    def test_long_text_shrinks_evenly(self):
        make_deck(self.deck)
        plan = plan_for()
        plan["pages"][0]["slots"]["s2"] = SUMMARY[0][1] + ". " + SUMMARY[0][1] * 4
        self.write_plan(plan)
        r = run(BEAUTIFY, self.deck, self.plan, self.out)
        self.assertEqual(r.returncode, 0, r.stderr)
        shrunk = json.loads(r.stdout)["rebuilt"][0]["shrunk"]
        sizes = {line.split(" to ")[1] for line in shrunk if line.startswith(("s2 ", "s4 ", "s6 "))}
        self.assertEqual(len(sizes), 1, shrunk)


if __name__ == "__main__":
    unittest.main()
