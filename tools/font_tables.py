#!/usr/bin/env python3
"""Width tables for the layout pass, read from font files (standard library only).

    python3 tools/font_tables.py --check            # rebuild Arial, Georgia, Verdana, Tahoma and compare
    python3 tools/font_tables.py --write            # add the faces in FACES to assets/font-widths.json

A table is, for each character of the file's "chars", its advance width in
thousandths of the em (rounded), regular and bold, and the face's line
spacing (hhea ascender - descender + line gap, over the em, 4 places), the
way the library's own tables were made. --check proves this reader against
the four faces already in the file before any new face goes in: every width
within one thousandth of the em (the old tables came from a browser's
measure, which rounds a hair differently) and the same line spacing. A
character the font lacks (a check mark, an arrow) takes a full em, as in the
old tables: another face draws it, and a full em is the safe side.

Only metrics are kept (no outlines, no font data), from fonts installed on
this Mac: the Microsoft faces that ship inside PowerPoint and the system's
own. Nothing is downloaded.
"""

import json
import os
import struct
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
TABLES = os.path.join(HERE, "..", "skills", "slide-beautify", "assets", "font-widths.json")
SUPP = "/System/Library/Fonts/Supplemental"
PPT = "/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts"

KNOWN = {  # the faces already in the file, to prove the reader
    "Arial": (SUPP + "/Arial.ttf", SUPP + "/Arial Bold.ttf"),
    "Georgia": (SUPP + "/Georgia.ttf", SUPP + "/Georgia Bold.ttf"),
    "Verdana": (SUPP + "/Verdana.ttf", SUPP + "/Verdana Bold.ttf"),
    "Tahoma": (SUPP + "/Tahoma.ttf", SUPP + "/Tahoma Bold.ttf"),
}
FACES = {  # common Office theme faces (Reviewer, 13 Nov 2026: decks in other faces)
    "Calibri": (PPT + "/Calibri.ttf", PPT + "/Calibrib.ttf"),
    "Calibri Light": (PPT + "/calibril.ttf", PPT + "/Calibrib.ttf"),
    "Aptos": (PPT + "/Aptos.ttf", PPT + "/Aptos-Bold.ttf"),
    "Cambria": (PPT + "/Cambria.ttc", PPT + "/Cambriab.ttf"),
    "Century Gothic": (PPT + "/Century Gothic.ttf", PPT + "/Century Gothic Bold.ttf"),
    "Franklin Gothic Book": (PPT + "/Franklin Gothic Book.ttf", PPT + "/Franklin Gothic Demi.ttf"),
    "Tw Cen MT": (PPT + "/Tw Cen MT.ttf", PPT + "/Tw Cen MT Bold.ttf"),
    "Times New Roman": (SUPP + "/Times New Roman.ttf", SUPP + "/Times New Roman Bold.ttf"),
}


class Font:
    """The few tables a width needs: head, hhea, hmtx, cmap."""

    def __init__(self, path, index=0):
        with open(path, "rb") as fh:
            self.data = fh.read()
        base = 0
        if self.data[:4] == b"ttcf":
            base = struct.unpack(">I", self.data[12 + 4 * index:16 + 4 * index])[0]
        num = struct.unpack(">H", self.data[base + 4:base + 6])[0]
        self.tables = {}
        for i in range(num):
            rec = self.data[base + 12 + 16 * i:base + 28 + 16 * i]
            tag, _, off, length = struct.unpack(">4sIII", rec)
            self.tables[tag.decode("latin-1")] = (off, length)
        head = self.table("head")
        self.upm = struct.unpack(">H", head[18:20])[0]
        hhea = self.table("hhea")
        self.ascender, self.descender, self.line_gap = struct.unpack(">hhh", hhea[4:10])
        n_hm = struct.unpack(">H", hhea[34:36])[0]
        hmtx = self.table("hmtx")
        self.advances = [struct.unpack(">H", hmtx[4 * i:4 * i + 2])[0] for i in range(n_hm)]
        self.cmap = self._cmap()

    def table(self, tag):
        off, length = self.tables[tag]
        return self.data[off:off + length]

    def _cmap(self):
        t = self.table("cmap")
        n = struct.unpack(">H", t[2:4])[0]
        subs = []
        for i in range(n):
            pid, eid, off = struct.unpack(">HHI", t[4 + 8 * i:12 + 8 * i])
            fmt = struct.unpack(">H", t[off:off + 2])[0]
            subs.append(((pid, eid) in ((3, 10), (0, 4), (0, 6)), fmt in (12,), (pid, eid) in ((3, 1), (0, 3)), off, fmt))
        out = {}
        # the Unicode BMP table first, then a full one if there is one
        for _, _, _, off, fmt in sorted(subs, key=lambda s: (s[0], s[1], s[2])):
            if fmt == 4:
                out.update(self._fmt4(t, off))
            elif fmt == 12:
                out.update(self._fmt12(t, off))
        return out

    @staticmethod
    def _fmt4(t, off):
        seg2 = struct.unpack(">H", t[off + 6:off + 8])[0]
        segs = seg2 // 2
        ends = struct.unpack(">%dH" % segs, t[off + 14:off + 14 + seg2])
        p = off + 16 + seg2
        starts = struct.unpack(">%dH" % segs, t[p:p + seg2])
        deltas = struct.unpack(">%dh" % segs, t[p + seg2:p + 2 * seg2])
        ro_at = p + 2 * seg2
        ros = struct.unpack(">%dH" % segs, t[ro_at:ro_at + seg2])
        out = {}
        for i in range(segs):
            for c in range(starts[i], ends[i] + 1):
                if c == 0xFFFF:
                    continue
                if ros[i] == 0:
                    g = (c + deltas[i]) & 0xFFFF
                else:
                    at = ro_at + 2 * i + ros[i] + 2 * (c - starts[i])
                    g = struct.unpack(">H", t[at:at + 2])[0]
                    if g:
                        g = (g + deltas[i]) & 0xFFFF
                if g:
                    out[c] = g
        return out

    @staticmethod
    def _fmt12(t, off):
        n = struct.unpack(">I", t[off + 12:off + 16])[0]
        out = {}
        for i in range(n):
            s, e, g = struct.unpack(">III", t[off + 16 + 12 * i:off + 28 + 12 * i])
            for c in range(s, e + 1):
                out[c] = g + (c - s)
        return out

    def advance(self, ch):
        g = self.cmap.get(ord(ch))
        if g is None:
            return None
        return self.advances[min(g, len(self.advances) - 1)]

    def widths(self, chars):
        return [1000 if self.advance(ch) is None else round(self.advance(ch) * 1000 / self.upm) for ch in chars]

    def line_height(self):
        return round((self.ascender - self.descender + self.line_gap) / self.upm, 4)


def table_for(regular, bold, chars):
    r, b = Font(regular), Font(bold)
    return {"lineHeight": r.line_height(), "regular": r.widths(chars), "bold": b.widths(chars)}


def main():
    with open(TABLES, encoding="utf-8") as fh:
        tables = json.load(fh)
    chars = tables["chars"]
    bad = 0
    for face, (reg, bold) in KNOWN.items():
        got = table_for(reg, bold, chars)
        want = tables["fonts"][face]
        fr, fb = Font(reg), Font(bold)
        drawn = [i for i in range(len(chars)) if fr.advance(chars[i]) is not None and fb.advance(chars[i]) is not None]
        diffs = [i for i in drawn if abs(got["regular"][i] - want["regular"][i]) > 1 or abs(got["bold"][i] - want["bold"][i]) > 1]
        print("%-16s %d characters the font lacks (taken by another face): %s" % (face, len(chars) - len(drawn), "".join(c for i, c in enumerate(chars) if i not in drawn)))
        lh = abs(got["lineHeight"] - want["lineHeight"]) < 1e-4
        print("%-16s %d of %d characters it draws more than 1/1000 em apart, line height %s" % (face, len(diffs), len(drawn), "same" if lh else "%s vs %s" % (got["lineHeight"], want["lineHeight"])))
        if diffs:
            print("   ", "".join(chars[i] for i in diffs[:20]))
        bad += bool(diffs) or not lh
    if "--write" in sys.argv:
        if bad:
            sys.exit("The reader does not reproduce the existing tables: nothing written")
        for face, (reg, bold) in FACES.items():
            if not (os.path.exists(reg) and os.path.exists(bold)):
                print("skip", face, "(font files not on this Mac)")
                continue
            tables["fonts"][face] = table_for(reg, bold, chars)
            print("added", face, tables["fonts"][face]["lineHeight"])
        with open(TABLES, "w", encoding="utf-8") as fh:
            json.dump(tables, fh, ensure_ascii=True, separators=(",", ":"))  # escaped: no literal dash characters in the file
    return bad


if __name__ == "__main__":
    sys.exit(1 if main() and "--check" in sys.argv else 0)
