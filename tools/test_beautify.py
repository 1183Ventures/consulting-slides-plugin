#!/usr/bin/env python3
"""Tests for the slide-beautify scripts: python3 tools/test_beautify.py

Builds small decks with python-pptx (16:9 and 4:3), runs read_deck.py and
beautify_deck.py on them, and checks the output: slides in place, words
carried over, notes kept, the original untouched, bad plans refused. And
holds layoutfit.py to layoutfit.js on recorded library pages
(tools/vectors-layoutfit.json, from lavender tools/layoutfit-vectors.mjs):
every field of every shape within 0.0002 in or pt, and the same points added.
"""

import hashlib
import json
import os
import subprocess
import sys
import tempfile
import copy
import unittest

from pptx import Presentation
from pptx.util import Inches

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "skills", "slide-beautify", "scripts")
READ = os.path.join(ROOT, "read_deck.py")
BEAUTIFY = os.path.join(ROOT, "beautify_deck.py")
VECTORS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vectors-layoutfit.json")
EDGE_VECTORS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vectors-layoutfit-edge.json")
sys.path.insert(0, ROOT)
import layoutfit  # noqa: E402

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

    def test_layout_pass_runs_and_can_be_turned_off(self):
        for wide in (True, False):
            make_deck(self.deck, wide=wide)
            self.write_plan(plan_for())
            r = run(BEAUTIFY, self.deck, self.plan, self.out)
            self.assertEqual(r.returncode, 0, r.stderr)
            rep = json.loads(r.stdout)["rebuilt"][0]
            self.assertTrue(any("fitted to their text" in x for x in rep["layout"]), rep)
            self.assertGreaterEqual(rep["text_up"], 0)
            prs = Presentation(self.out)
            for sh in prs.slides[1].shapes:
                self.assertLessEqual(sh.left + sh.width, prs.slide_width + 9525)
                self.assertLessEqual(sh.top + sh.height, prs.slide_height + 9525)
            r = run(BEAUTIFY, self.deck, self.plan, self.out, "--no-fit")
            rep = json.loads(r.stdout)["rebuilt"][0]
            self.assertEqual((rep["layout"], rep["text_up"]), ([], 0))

    def test_fitted_boxes_keep_resize_to_fit_text(self):
        make_deck(self.deck)
        self.write_plan(plan_for())
        counts = []
        for extra in ((), ("--no-fit",)):
            r = run(BEAUTIFY, self.deck, self.plan, self.out, *extra)
            self.assertEqual(r.returncode, 0, r.stderr)
            xml = b"".join(Presentation(self.out).slides[1].part.blob.split(b"\n"))
            counts.append(xml.count(b"<a:spAutoFit/>"))
        self.assertGreater(counts[0], 0)
        self.assertEqual(counts[1], 0)

    def test_a_face_without_tables_is_fitted_with_room_and_not_raised(self):
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        import check_slide_fit
        for faces, unknown in ((("Segoe UI", "Segoe UI"), True), (("Poppins", "Poppins"), False), (("Aptos Display", "Aptos"), False)):
            prs = Presentation()
            prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
            check_slide_fit.set_theme_faces(prs, *faces)
            for _ in range(3):
                prs.slides.add_slide(prs.slide_layouts[5]).shapes.title.text = "T"
            prs.save(self.deck)
            self.write_plan(plan_for())
            r = run(BEAUTIFY, self.deck, self.plan, self.out)
            self.assertEqual(r.returncode, 0, r.stderr)
            rep = json.loads(r.stdout)["rebuilt"][0]
            said = any("Segoe UI has no width table" in w for w in rep["warnings"])
            self.assertEqual(said, unknown, rep["warnings"])
            if unknown:
                self.assertEqual(rep["text_up"], 0)

    def test_every_fitted_box_holds_its_text_on_smaller_slides(self):
        # the saved slide, measured as drawn (size, width, insets), in Arial
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        import check_slide_fit
        with open(os.path.join(ROOT, "..", "assets", "pages.json"), encoding="utf-8") as fh:
            pages = json.load(fh)["pages"]
        with open(self.plan, "w", encoding="utf-8") as fh:
            json.dump(check_slide_fit.plan_for(pages, 1.0), fh)
        # 4:3, 10 x 5.625 (On-screen Show 16:9, Google Slides' default) and A4
        for size in (False, "screen", "a4"):
            check_slide_fit.make_deck(self.deck, size, False, len(pages), ("Georgia", "Arial"))
            r = run(BEAUTIFY, self.deck, self.plan, self.out)
            self.assertEqual(r.returncode, 0, r.stderr)
            res = check_slide_fit.check(self.out, layoutfit.Measure())
            self.assertGreater(res["boxes"], 100)
            self.assertEqual(res["taller"], 0, (size, res["where"]))

    def test_fit_first_experiment_only_ever_gains(self):
        # --fit-first (off by default): no text ends smaller than the default run, every fitted box holds its text
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        import check_slide_fit
        with open(os.path.join(ROOT, "..", "assets", "pages.json"), encoding="utf-8") as fh:
            pages = json.load(fh)["pages"]
        check_slide_fit.make_deck(self.deck, True, False, len(pages), ("Georgia", "Arial"))
        with open(self.plan, "w", encoding="utf-8") as fh:
            json.dump(check_slide_fit.plan_for(pages, 1.8), fh)
        dumps = {}
        for name, extra in (("default", ()), ("first", ("--fit-first",))):
            path = os.path.join(self.dir, name + ".json")
            r = run(BEAUTIFY, self.deck, self.plan, self.out, "--dump-shapes", path, *extra)
            self.assertEqual(r.returncode, 0, r.stderr)
            with open(path, encoding="utf-8") as fh:
                dumps[name] = json.load(fh)["slides"]
            if extra:
                res = check_slide_fit.check(self.out, layoutfit.Measure())
                self.assertEqual(res["taller"], 0, res["where"])
        gained = 0
        for a, b in zip(dumps["default"], dumps["first"]):
            for sa, sb in zip(a["shapes"], b["shapes"]):
                for pa, pb in zip((sa.get("text") or {}).get("paras", []), (sb.get("text") or {}).get("paras", [])):
                    for ra, rb in zip(pa.get("runs", []), pb.get("runs", [])):
                        self.assertGreaterEqual(rb.get("sz") or 12, (ra.get("sz") or 12) - 1e-9, a["id"])
                        gained += (rb.get("sz") or 12) > (ra.get("sz") or 12) + 1e-9
        self.assertGreater(gained, 0)

    def test_a_text_growing_into_the_title_is_warned(self):
        # big-numbers at 1.4 times its words: a figure anchored at its foot grows up past the body's top
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        import check_slide_fit
        with open(os.path.join(ROOT, "..", "assets", "pages.json"), encoding="utf-8") as fh:
            big = [p for p in json.load(fh)["pages"] if p["id"] == "big-numbers"]
        for f, warned in ((1.0, False), (1.4, True)):
            check_slide_fit.make_deck(self.deck, True, False, 1, None)
            with open(self.plan, "w", encoding="utf-8") as fh:
                json.dump(check_slide_fit.plan_for(big, f), fh)
            r = run(BEAUTIFY, self.deck, self.plan, self.out)
            rep = json.loads(r.stdout)["rebuilt"][0]
            self.assertEqual(any("title area" in w for w in rep["warnings"]), warned, (f, rep["warnings"]))

    def test_fit_first_keeps_kinds_together(self):
        # a text kept at today's size keeps its whole kind (Reviewer, 14 Nov 2026: kpi-tiles' figures at
        # 35 and 40 pt, two-options); 2cc6320 split both on these decks
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        import check_slide_fit
        with open(os.path.join(ROOT, "..", "assets", "pages.json"), encoding="utf-8") as fh:
            pages = {p["id"]: p for p in json.load(fh)["pages"]}
        picks = [pages["kpi-tiles"], pages["two-options"], pages["big-numbers"]]
        for faces in (None, ("Georgia", "Arial")):
            check_slide_fit.make_deck(self.deck, True, False, len(picks), faces)
            with open(self.plan, "w", encoding="utf-8") as fh:
                json.dump(check_slide_fit.plan_for(picks, 1.4), fh)
            dump = os.path.join(self.dir, "ff.json")
            r = run(BEAUTIFY, self.deck, self.plan, self.out, "--fit-first", "--dump-shapes", dump)
            self.assertEqual(r.returncode, 0, r.stderr)
            with open(dump, encoding="utf-8") as fh:
                slides = json.load(fh)["slides"]
            for sl in slides:
                kinds = {}
                for d, a in zip(pages[sl["id"]]["shapes"], sl["shapes"]):
                    if not layoutfit.is_text_box(d):
                        continue
                    r0 = ((d.get("text") or {}).get("paras") or [{}])[0].get("runs") or [{}]
                    key = (r0[0].get("sz") or 12, bool(r0[0].get("b")), r0[0].get("color") or "")
                    kinds.setdefault(key, set()).add(round((a["text"]["paras"][0].get("runs") or [{}])[0].get("sz") or 12, 2))
                for key, sizes in kinds.items():
                    self.assertEqual(len(sizes), 1, (faces, sl["id"], key, sizes))
            # a big text that already wraps never grows: big-numbers warns no more than it does by default
            rep = json.loads(r.stdout)["rebuilt"]
            r2 = run(BEAUTIFY, self.deck, self.plan, self.out)
            base = json.loads(r2.stdout)["rebuilt"]
            for a, b in zip(base, rep):
                self.assertLessEqual(len(b["warnings"]), len(a["warnings"]), (a["page"], b["warnings"]))

    def test_a_design_the_check_fails_is_not_warned_at_its_own_words(self):
        # a library page drawing a label across a rule fails the fill's check by design: no warning with
        # its own words, a warning once longer words make it worse
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        import build_beautify_sample as BS
        box = lambda y, t: {"kind": "sp", "geom": "rect", "x": 1, "y": y, "w": 5, "h": 0.4, "fill": None, "line": None,
                            "text": {"anchor": "t", "inset": [0, 0, 0, 0], "paras": [{"runs": [{"sz": 12, "t": t}]}]}}
        page = {"id": "design-on-a-rule", "name": "Label on a rule", "section": "Test", "fits": "a test",
                "shapes": [{"kind": "ph", "type": "title", "text": {"paras": [{"runs": [{"t": "A title"}]}]}},
                           box(2.0, "A label that a rule runs through as drawn and wraps to two lines here"),
                           {"kind": "cxn", "x1": 1, "y1": 2.2, "x2": 6, "y2": 2.2, "color": "a4", "weight": 0.75},
                           box(2.7, "A second text under it"), box(3.4, "A third text under that one")]}
        page["slots"] = BS.slots_of(page)
        cat = os.path.join(self.dir, "cat.json")
        with open(cat, "w", encoding="utf-8") as fh:
            json.dump({"version": 1, "canvas": {"w": 13.333, "h": 7.5}, "pages": [page]}, fh)
        own = {sl["id"]: sl["sample"] for sl in page["slots"] if sl["role"] not in ("title", "source", "fixed")}
        for words, warned in ((own, False), ({k: (v + " ") * 6 for k, v in own.items()}, True)):
            prs = Presentation()
            prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
            prs.slides.add_slide(prs.slide_layouts[5]).shapes.title.text = "T"
            prs.save(self.deck)
            self.write_plan({"pages": [{"slide": 1, "page": page["id"], "title": "A title", "slots": words}]})
            r = run(BEAUTIFY, self.deck, self.plan, self.out, "--pages", cat)
            self.assertEqual(r.returncode, 0, r.stderr)
            rep = json.loads(r.stdout)["rebuilt"][0]
            self.assertEqual(any("run into" in w for w in rep["warnings"]), warned, (warned, rep["warnings"]))

    def _round_one_deck(self, path, n=3):
        # a deck as round 1 of the Windows render had it: a 10 x 5.625 slide still typed screen4x3, its
        # title layout moved down with no width written, the blank Office master's 44 pt title
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(10), Inches(5.625)
        lay = prs.slide_layouts[5]
        for ph in lay.placeholders:
            if "TITLE" in str(ph.placeholder_format.type):
                ph.top, ph.height = Inches(0.9), Inches(1.4)
        for i in range(n):
            prs.slides.add_slide(lay).shapes.title.text = "A long working title that will not fit at forty-four points %d" % i
        prs.save(path)

    def test_the_file_check_fails_round_one_titles_and_the_writer_fixes_them(self):
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        import check_slide_fit
        self._round_one_deck(self.deck)
        faults = check_slide_fit.title_faults(Presentation(self.deck), check_slide_fit.theme_faces(Presentation(self.deck)), layoutfit.Measure())
        self.assertTrue(any("screen4x3" in f for f in faults), faults)
        self.assertTrue(any("no width" in f for f in faults), faults)
        # the beautifier writes each title's box, size and alignment itself
        self.write_plan(plan_for())
        r = run(BEAUTIFY, self.deck, self.plan, self.out)
        self.assertEqual(r.returncode, 0, r.stderr)
        out = Presentation(self.out)
        title = out.slides[1].shapes.title
        self.assertGreater(title.width, 0)
        faults = [f for f in check_slide_fit.title_faults(out, check_slide_fit.theme_faces(out), layoutfit.Measure()) if "slide 2:" in f]
        self.assertEqual(faults, [])
        # with the master's title box empty too, the library's frame stands in, and the report says so
        prs = Presentation(self.deck)
        for sp in prs.slide_master.element.iter("{http://schemas.openxmlformats.org/presentationml/2006/main}sp"):
            ph = sp.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}ph")
            if ph is not None and ph.get("type") == "title":
                ext = sp.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}ext")
                ext.set("cx", "0")
        prs.save(self.deck)
        r = run(BEAUTIFY, self.deck, self.plan, self.out)
        rep = json.loads(r.stdout)["rebuilt"][0]
        self.assertTrue(any("library's frame" in w for w in rep["warnings"]), rep["warnings"])
        out = Presentation(self.out)
        self.assertEqual([f for f in check_slide_fit.title_faults(out, check_slide_fit.theme_faces(out), layoutfit.Measure()) if "slide 2:" in f], [])

    def test_a_blank_office_deck_gets_the_house_title_size_written(self):
        make_deck(self.deck)                     # python-pptx's blank Office master: 44 pt titles
        self.write_plan(plan_for())
        r = run(BEAUTIFY, self.deck, self.plan, self.out)
        self.assertEqual(r.returncode, 0, r.stderr)
        title = Presentation(self.out).slides[1].shapes.title
        p0 = title.text_frame.paragraphs[0]
        self.assertEqual(p0.runs[0].font.size.pt, 24.0)
        self.assertIsNotNone(p0._p.pPr.get("algn"))
        body = title.text_frame._txBody.find("{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr")
        self.assertIsNotNone(body.find("{http://schemas.openxmlformats.org/drawingml/2006/main}noAutofit"))
        self.assertLessEqual(title.height, Inches(2 * 24 * 1.2 / 72 + 0.1) + 1)

    def test_long_text_shrinks_evenly(self):
        make_deck(self.deck)
        plan = plan_for()
        plan["pages"][0]["slots"]["s2"] = SUMMARY[0][1] + ". " + SUMMARY[0][1] * 4
        self.write_plan(plan)
        # the shrink step alone: one size for the kind of text
        r = run(BEAUTIFY, self.deck, self.plan, self.out, "--no-fit")
        self.assertEqual(r.returncode, 0, r.stderr)
        shrunk = json.loads(r.stdout)["rebuilt"][0]["shrunk"]
        sizes = {line.split(" to ")[1] for line in shrunk if line.startswith(("s2 ", "s4 ", "s6 "))}
        self.assertEqual(len(sizes), 1, shrunk)
        # with the pass, every shrunk line tells the size the slide carries
        # (the fill may give some back), and a text not listed kept its size
        for extra in ((), ("--no-fit",)):
            r = run(BEAUTIFY, self.deck, self.plan, self.out, *extra)
            rep = json.loads(r.stdout)["rebuilt"][0]
            said = {line.split(" ")[0]: float(line.split(" to ")[1].split(" ")[0]) for line in rep["shrunk"]}
            slide = Presentation(self.out).slides[1]
            for sh in slide.shapes:
                if not sh.name.startswith("Beautify executive-summary s") or not sh.has_text_frame:
                    continue
                sid = sh.name.split()[-1]
                sz = next((rn.font.size.pt for p in sh.text_frame.paragraphs for rn in p.runs if rn.font.size), None)
                if sid in said:
                    self.assertEqual(sz, said[sid], (extra, sid, rep["shrunk"]))



def rounded(x):
    if isinstance(x, bool) or x is None or isinstance(x, str):
        return x
    if isinstance(x, (int, float)):
        return round(x, 4)
    if isinstance(x, list):
        return [rounded(v) for v in x]
    return {k: rounded(v) for k, v in x.items()}


_MISSING = object()


def first_difference(a, b, path=""):
    """None when a and b agree (numbers within 0.0002), else where they part.
    A key one side lacks is a difference, even against a null."""
    if a is _MISSING or b is _MISSING:
        return None if a is b else "%s: present on one side only" % path
    if isinstance(a, dict) and isinstance(b, dict):
        for k in sorted(set(a) | set(b)):
            d = first_difference(a.get(k, _MISSING), b.get(k, _MISSING), path + "." + k)
            if d:
                return d
        return None
    if isinstance(a, list) and isinstance(b, list):
        if len(a) != len(b):
            return "%s: %d items vs %d" % (path, len(a), len(b))
        for i, (x, y) in enumerate(zip(a, b)):
            d = first_difference(x, y, "%s[%d]" % (path, i))
            if d:
                return d
        return None
    num = (int, float)
    if isinstance(a, num) and isinstance(b, num) and not isinstance(a, bool) and not isinstance(b, bool):
        return None if abs(a - b) <= 2e-4 else "%s: %r vs %r" % (path, a, b)
    return None if a == b else "%s: %r vs %r" % (path, a, b)


class LayoutFitParity(unittest.TestCase):
    """layoutfit.py does to each recorded page what layoutfit.js did."""

    @classmethod
    def setUpClass(cls):
        with open(VECTORS, encoding="utf-8") as fh:
            cls.vectors = json.load(fh)["vectors"]
        cls.measure = layoutfit.Measure()

    def run_page(self, v, **kw):
        shapes = copy.deepcopy(v["input"])
        r = layoutfit.fit_layout(shapes, fonts=v["fonts"], measure=self.measure, body=v["body"], **kw)
        return shapes, r

    def test_every_recorded_page_matches(self):
        self.assertGreaterEqual(len(self.vectors), 60)
        for v in self.vectors:
            shapes, r = self.run_page(v, page_id=v["id"])
            with self.subTest(page=v["id"]):
                self.assertIsNone(first_difference(rounded(shapes), v["expect"]))
                self.assertEqual(r["size"], v["size"])

    def test_values_read_as_js_reads_them(self):
        # odd whitespace, an empty {} fill, numeric run text, an empty [] line
        with open(EDGE_VECTORS, encoding="utf-8") as fh:
            edge = json.load(fh)["vectors"]
        self.assertEqual(len(edge), 16)
        for v in edge:
            shapes, r = self.run_page(v, page_id=v["page"])
            with self.subTest(case=v["id"]):
                self.assertIsNone(first_difference(rounded(shapes), v["expect"]))
                self.assertEqual(r["size"], v["size"])

    def test_a_row_with_a_skipped_text_is_held(self):
        # as layoutfit.test.mjs: one row, one size, and smaller text stops at the held row's size
        def box(x, y, w, sz, t):
            return {"kind": "sp", "geom": "rect", "x": x, "y": y, "w": w, "h": 0.3, "fill": None, "line": None, "name": t,
                    "text": {"anchor": "t", "inset": [0, 0, 0, 0], "paras": [{"runs": [{"sz": sz, "t": t}]}]}}
        shapes = [box(0.5, 1.8, 0.5, 12, "#"), box(1.2, 1.8, 3, 12, "Header one"), box(4.4, 1.8, 3, 12, "Header two"),
                  box(0.5, 2.6, 6, 11, "Body text that may grow"), box(0.5, 3.4, 6, 11, "More body text that may grow")]
        r = layoutfit.fit_layout(shapes, measure=self.measure)
        size = {s["name"]: s["text"]["paras"][0]["runs"][0]["sz"] for s in shapes}
        self.assertEqual((size["#"], size["Header one"], size["Header two"]), (12, 12, 12))
        self.assertTrue(0 < r["size"] <= 1, r["size"])
        self.assertLessEqual(size["Body text that may grow"], 12)

    def test_content_clears_the_titles_text(self):
        import pagexml
        E = 914400
        W, H = int(13.333 * E), int(7.5 * E)
        # a one-line title's text ends at 1.01 in: a page starting at 1.3 in keeps its scale
        self.assertEqual(pagexml.Place(W, H, 1.3, int(1.01 * E)).k, 1.0)
        # a two-line title's text ends at 1.41 in: the page moves under it, 0.12 in clear
        two = pagexml.Place(W, H, 1.3, int(1.41 * E))
        self.assertLess(two.k, 1.0)
        self.assertAlmostEqual(two.dy / E, 1.41 + 0.12 - 1.3, places=3)
        # a page starting at 1.62 in under text ending at 1.7 in keeps 0.12 in clear, as before
        low = pagexml.Place(W, H, 1.62, int(1.7 * E))
        self.assertAlmostEqual(low.dy / E, 1.7 + 0.12 - 1.62, places=3)

    def test_rounded_shapes_keep_their_corners(self):
        # a pill (adj 50000) and a chevron's point reach the deck, resized or not (CEO, 14 Nov 2026)
        import pagexml
        E = 914400
        place = pagexml.Place(int(13.333 * E), int(7.5 * E))
        pill = {"kind": "sp", "geom": "roundRect", "x": 1, "y": 1, "w": 1.85, "h": 0.38, "fill": "a3", "adj": {"adj": 50000}}
        self.assertIn('<a:gd name="adj" fmla="val 50000"/>', pagexml.sp_xml(pill, 5, place))
        taller = dict(pill, h=0.6)
        self.assertIn('<a:gd name="adj" fmla="val 50000"/>', pagexml.sp_xml(taller, 6, place))
        self.assertIn("<a:avLst></a:avLst>", pagexml.sp_xml(dict(pill, adj=None), 7, place))
        # every shape the sample draws with adjust values carries them into a beautified deck
        with open(os.path.join(ROOT, "..", "assets", "pages.json"), encoding="utf-8") as fh:
            pages = json.load(fh)["pages"]
        self.assertGreaterEqual(sum(1 for p in pages for s in p["shapes"] if s.get("adj")), 20)
        self.assertEqual(next(s for s in next(p for p in pages if p["id"] == "two-options")["shapes"]
                              if s.get("geom") == "roundRect")["adj"], {"adj": 50000})

    def test_sample_roles(self):
        with open(os.path.join(ROOT, "..", "assets", "pages.json"), encoding="utf-8") as fh:
            pages = {p["id"]: p for p in json.load(fh)["pages"]}
        role = lambda pid, sample: next(s["role"] for s in pages[pid]["slots"] if s.get("sample") == sample)
        # a wide figure is the user's to replace; badges and numeral columns stay fixed
        self.assertEqual(role("kpi-tiles", "64"), "heading")
        self.assertEqual(role("key-takeaways", "1"), "fixed")
        self.assertEqual(role("decisions", "1"), "fixed")
        self.assertEqual(role("swot", "S"), "fixed")
        # texts drawn alike in one row share a role
        scr = pages["situation-complication-resolution"]
        leads = [s["role"] for s in scr["slots"] if s.get("sample", "").startswith(("What everyone", "What changed", "What we propose"))]
        self.assertEqual(len(leads), 3)
        self.assertEqual(len(set(leads)), 1, leads)

    def test_spacing_rounds_down_to_half_points_and_keeps_the_librarys(self):
        import pagexml
        E = 914400
        lib = pagexml.Place(int(13.333 * E), int(7.5 * E))
        self.assertEqual([lib.spacing(v) for v in (4, 8, 12)], [4, 8, 12])
        screen = pagexml.Place(int(10 * E), int(5.625 * E))
        self.assertEqual(screen.spacing(12), 9.0)
        self.assertEqual(screen.spacing(10), 7.5)

    def test_label_half_step_is_an_option(self):
        # off by default (the JS parity holds); on, a smaller label keeps half its step under the kept text above
        v = next(x for x in self.vectors if x["id"] == "situation-complication-resolution") if any(
            x["id"] == "situation-complication-resolution" for x in self.vectors) else None
        shapes = [{"kind": "sp", "geom": "rect", "x": 0.5, "y": 2 + i * 0.6, "w": 6, "h": 0.4, "fill": None, "line": None,
                   "text": {"anchor": "t", "inset": [0, 0, 0, 0], "paras": [{"runs": [{"sz": sz, "t": t}]}]}}
                  for i, (sz, t) in enumerate(((14, "The body text the fill leaves alone"), (11, "A small label under it"),
                                               (14, "More body text"), (11, "Another small label")))]
        plain, half = copy.deepcopy(shapes), copy.deepcopy(shapes)
        r0 = layoutfit.fit_layout(plain, measure=self.measure)
        r1 = layoutfit.fit_layout(half, measure=self.measure, label_half_step=True)
        self.assertEqual(r0["size"], 2.5)            # 11 to 13.5, half a point under 14
        self.assertEqual(r1["size"], 1.5)            # 11 to 12.5: half of the 3 pt step kept
        if v is not None:
            a, b = copy.deepcopy(v["input"]), copy.deepcopy(v["input"])
            layoutfit.fit_layout(a, fonts=v["fonts"], measure=self.measure, body=v["body"], page_id=v["id"])
            self.assertIsNone(first_difference(rounded(a), v["expect"]))

    def test_comparer_tells_missing_from_null(self):
        self.assertIsNotNone(first_difference({"a": None}, {}))
        self.assertIsNone(first_difference({"a": 1.00001}, {"a": 1.0}))

    def test_the_row_rules_match_the_js(self):
        # the 12 pages the Reviewer's row findings change (column hold, table hold, near misses, kinds)
        with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "vectors-layoutfit-rows.json"), encoding="utf-8") as fh:
            rows = json.load(fh)["vectors"]
        self.assertEqual(len(rows), 12)
        for v in rows:
            shapes, r = self.run_page(v, page_id=v["id"])
            with self.subTest(page=v["id"]):
                self.assertIsNone(first_difference(rounded(shapes), v["expect"]))
                self.assertEqual(r["size"], v["size"])

    def test_caps_hold(self):
        capped = [v for v in self.vectors if v["id"] in layoutfit.CAPS]
        self.assertGreaterEqual(len(capped), 7)
        for v in capped:
            cap = layoutfit.CAPS[v["id"]]
            _, r = self.run_page(v, page_id=v["id"])
            self.assertLessEqual(r["size"], max(cap, 0))
        # a page below zero is left exactly as drawn
        v = next(v for v in capped if layoutfit.CAPS[v["id"]] < 0)
        shapes, r = self.run_page(v, page_id=v["id"])
        self.assertEqual(r, {"changed": [], "size": 0})
        self.assertEqual(shapes, v["input"])

    def test_fill_touches_only_body_text(self):
        # only runs drawn at 9 to 13.5 pt change; each grows by the page's step, never past 14 pt
        for v in self.vectors:
            shapes, r = self.run_page(v, page_id=v["id"])
            if not r["size"] or len(shapes) != len(v["input"]):
                continue
            for before, after in zip(v["input"], shapes):
                for pb, pa in zip((before.get("text") or {}).get("paras", []), (after.get("text") or {}).get("paras", [])):
                    for rb, ra in zip(pb.get("runs", []), pa.get("runs", [])):
                        was, now = rb.get("sz") or 12, ra.get("sz") or 12
                        if now == was:
                            continue
                        with self.subTest(page=v["id"]):
                            self.assertTrue(layoutfit.FILL_FROM <= was <= layoutfit.FILL_TO)
                            self.assertEqual(now, min(layoutfit.MAX_FILL, was + r["size"]))

    def test_to_fixed_rounds_like_js(self):
        self.assertEqual(layoutfit.to_fixed(0.125, 2), "0.13")
        self.assertEqual(layoutfit.to_fixed(1.005, 2), "1.00")
        self.assertEqual(layoutfit.to_fixed(-0.125, 2), "-0.13")
        self.assertEqual(layoutfit.to_fixed(2.5, 0), "3")


LIBRARY = os.path.join(ROOT, "library.py")


class FakeGate:
    """The gate's protocol (lavender netlify/lib/beautify-gate.mjs), in memory."""

    def __init__(self, mode="on"):
        import http.server
        import threading
        self.mode, self.codes, self.tokens, self.seen = mode, {}, set(), []
        gate = self

        class H(http.server.BaseHTTPRequestHandler):
            def log_message(self, *a):
                pass

            def do_POST(self):
                body = json.loads(self.rfile.read(int(self.headers.get("Content-Length") or 0)) or b"{}")
                gate.seen.append((body, self.headers.get("Authorization")))
                status, out = gate.answer(body, self.headers.get("Authorization") or "")
                data = json.dumps(out).encode()
                self.send_response(status)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)

        self.server = http.server.HTTPServer(("127.0.0.1", 0), H)
        threading.Thread(target=self.server.serve_forever, daemon=True).start()
        self.url = "http://127.0.0.1:%d/" % self.server.server_port

    def answer(self, b, auth):
        if self.mode != "on":
            return 503, {"ok": False, "error": "not_open"}
        a = b.get("action")
        if a == "start":
            self.codes["K7QD-M2PX"] = {"poll": "secret", "state": "pending"}
            return 200, {"ok": True, "code": "K7QD-M2PX", "poll": "secret", "url": "https://x/activate/?code=K7QD-M2PX", "expires_in": 600, "interval": 1}
        if a == "poll":
            c = self.codes.get(b.get("code"))
            if not c or c["poll"] != b.get("poll"):
                return 404, {"ok": False, "error": "code_expired"}
            if c["state"] == "pending":
                return 200, {"ok": True, "state": "pending"}
            self.tokens.add("tok-1")
            return 200, {"ok": True, "state": "approved", "token": "tok-1", "expires_in": 43200}
        if a == "token":
            if b.get("key") != "GOOD-KEY":
                return 403, {"ok": False, "error": "license_required"}
            self.tokens.add("tok-2")
            return 200, {"ok": True, "token": "tok-2", "expires_in": 43200}
        if a == "pages":
            if auth.replace("Bearer ", "") not in self.tokens:
                return 401, {"ok": False, "error": "sign_in_required"}
            return 200, {"ok": True, "version": 1, "canvas": {"w": 13.333, "h": 7.5}, "pages": [{"id": "p1"}, {"id": "p2"}], "remaining": 19}
        return 400, {"ok": False, "error": "unknown_action"}


class Library(unittest.TestCase):
    def setUp(self):
        self.home = tempfile.mkdtemp()

    def lib(self, gate, *args, extra=None):
        env = dict(os.environ, SLIDESHORTCUTS_API=gate.url, SLIDESHORTCUTS_HOME=self.home, **(extra or {}))
        return subprocess.run([sys.executable, LIBRARY] + list(args), capture_output=True, text=True, env=env, timeout=60)

    def test_a_closed_gate_says_so(self):
        r = self.lib(FakeGate("off"), "start")
        self.assertEqual(r.returncode, 2)
        self.assertIn("not open yet", r.stdout)

    def test_sign_in_then_pages(self):
        gate = FakeGate()
        r = self.lib(gate, "start")
        self.assertEqual(r.returncode, 0, r.stdout)
        self.assertIn("K7QD-M2PX", r.stdout)
        r = self.lib(gate, "pages", "--out", os.path.join(self.home, "x.json"))
        self.assertIn("Sign in first", r.stdout)
        gate.codes["K7QD-M2PX"]["state"] = "approved"     # the person signs in on the web page
        r = self.lib(gate, "wait", "--limit", "20")
        self.assertEqual(r.returncode, 0, r.stdout)
        token_file = os.path.join(self.home, "beautify-token.json")
        self.assertEqual(oct(os.stat(token_file).st_mode & 0o777), "0o600")
        out = os.path.join(self.home, "pages-full.json")
        r = self.lib(gate, "pages", "--out", out)
        self.assertEqual(r.returncode, 0, r.stdout)
        with open(out, encoding="utf-8") as fh:
            self.assertEqual(len(json.load(fh)["pages"]), 2)
        self.assertEqual(gate.seen[-1][1], "Bearer tok-1")
        self.assertIn("Signed in until", self.lib(gate, "status").stdout)

    def test_key_from_the_environment_never_lands_on_disk(self):
        gate = FakeGate()
        r = self.lib(gate, "key", extra={"SLIDESHORTCUTS_EMAIL": "a@b.co", "SLIDESHORTCUTS_KEY": "GOOD-KEY"})
        self.assertEqual(r.returncode, 0, r.stdout)
        self.assertNotIn("GOOD-KEY", r.stdout)
        for name in os.listdir(self.home):
            with open(os.path.join(self.home, name), encoding="utf-8") as fh:
                text = fh.read()
            self.assertNotIn("GOOD-KEY", text)
            self.assertNotIn("a@b.co", text)
        r = self.lib(gate, "key", extra={"SLIDESHORTCUTS_EMAIL": "a@b.co", "SLIDESHORTCUTS_KEY": "BAD"})
        self.assertEqual(r.returncode, 2)
        self.assertIn("not accepted", r.stdout)


if __name__ == "__main__":
    unittest.main()
