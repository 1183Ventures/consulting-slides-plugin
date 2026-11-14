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

    def test_a_title_where_the_librarys_ends_squeezes_nothing(self):
        import pagexml
        E = 914400
        lib_title = int(1.46 * E)
        # a page whose column heads start at 1.3 in (above the library's title box end) keeps its scale
        self.assertEqual(pagexml.Place(int(13.333 * E), int(7.5 * E), 1.3, lib_title).k, 1.0)
        # a page starting at 1.62 in still keeps 0.12 in clear of a lower title, as before
        low = pagexml.Place(int(13.333 * E), int(7.5 * E), 1.62, int(1.7 * E))
        self.assertLess(low.k, 1.0)
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

    def test_comparer_tells_missing_from_null(self):
        self.assertIsNotNone(first_difference({"a": None}, {}))
        self.assertIsNone(first_difference({"a": 1.00001}, {"a": 1.0}))

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


if __name__ == "__main__":
    unittest.main()
