#!/usr/bin/env python3
"""A rough HTML preview of a deck's slides, for checking beautify output by eye.

Usage:
    python3 tools/preview_deck.py DECK.pptx OUT.html [--slides 2,3]

Draws each shape's box, fill, outline and text from the file, with the
deck's theme colours resolved (tints and shades included). Not a renderer:
fonts are approximated and autofit is ignored. A maintainer's tool, not
shipped with the skill.
"""

import argparse
import colorsys
import html
import sys

from lxml import etree
from pptx import Presentation

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
EMU_PX = 96 / 914400


def theme(prs):
    for rel in prs.slide_masters[0].part.rels.values():
        if rel.reltype.endswith("/theme"):
            root = etree.fromstring(rel.target_part.blob)
            out = {}
            for el in root.find(".//a:clrScheme", NS):
                c = el[0]
                sys_clr = etree.QName(c).localname == "sysClr"
                out[etree.QName(el).localname] = (c.get("lastClr") if sys_clr else c.get("val")) or "000000"
            out.update({"tx1": out.get("dk1"), "bg1": out.get("lt1"), "tx2": out.get("dk2"), "bg2": out.get("lt2")})
            fonts = {k: root.find(".//a:%s/a:latin" % k, NS).get("typeface") for k in ("majorFont", "minorFont")}
            return out, fonts
    return {}, {}


def color(el, th):
    if el is None:
        return None
    c = el.find("a:schemeClr", NS)
    hexv = None
    if c is not None:
        hexv = th.get(c.get("val"), "000000")
    else:
        c = el.find("a:srgbClr", NS)
        if c is not None:
            hexv = c.get("val")
    if not hexv:
        return None
    r, g, b = (int(hexv[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    mod = c.find("a:lumMod", NS)
    off = c.find("a:lumOff", NS)
    if mod is not None:
        l *= int(mod.get("val")) / 100000
    if off is not None:
        l += int(off.get("val")) / 100000
    r, g, b = colorsys.hls_to_rgb(h, min(1, l), s)
    alpha = c.find("a:alpha", NS)
    a = int(alpha.get("val")) / 100000 if alpha is not None else 1
    return "rgba(%d,%d,%d,%.2f)" % (r * 255, g * 255, b * 255, a)


def shape_html(sh, th, fonts):
    el = sh.element
    x, y = (sh.left or 0) * EMU_PX, (sh.top or 0) * EMU_PX
    w, h = (sh.width or 0) * EMU_PX, (sh.height or 0) * EMU_PX
    sppr = el.find("p:spPr", NS)
    style = ["left:%.1fpx" % x, "top:%.1fpx" % y, "width:%.1fpx" % w, "height:%.1fpx" % h]
    if el.tag.endswith("cxnSp"):
        ln = sppr.find("a:ln", NS)
        col = color(ln.find("a:solidFill", NS), th) if ln is not None else "#000"
        flip_v = sppr.find("a:xfrm", NS).get("flipV") == "1"
        x1, y1, x2, y2 = (0, h, w, 0) if flip_v else (0, 0, w, h)
        return ('<svg style="position:absolute;left:%.1fpx;top:%.1fpx;overflow:visible" width="%.1f" height="%.1f">'
                '<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="%s" stroke-width="1"/></svg>') % (x, y, max(w, 1), max(h, 1), x1, y1, x2, y2, col)
    if sppr is not None:
        fill = sppr.find("a:solidFill", NS)
        if fill is not None:
            style.append("background:%s" % color(fill, th))
        ln = sppr.find("a:ln", NS)
        if ln is not None and ln.find("a:solidFill", NS) is not None:
            style.append("outline:1px solid %s" % color(ln.find("a:solidFill", NS), th))
        geom = sppr.find("a:prstGeom", NS)
        if geom is not None and geom.get("prst") == "ellipse":
            style.append("border-radius:50%")
        if geom is not None and geom.get("prst") == "roundRect":
            style.append("border-radius:8px")
    body = ""
    if sh.has_text_frame:
        bp = el.find(".//a:bodyPr", NS)
        anchor = {"ctr": "center", "b": "flex-end"}.get(bp.get("anchor") if bp is not None else "t", "flex-start")
        style += ["display:flex", "flex-direction:column", "justify-content:%s" % anchor]
        paras = []
        for p in el.findall(".//a:p", NS):
            ppr = p.find("a:pPr", NS)
            algn = {"ctr": "center", "r": "right"}.get(ppr.get("algn") if ppr is not None else "l", "left")
            bullet = "• " if ppr is not None and ppr.find("a:buChar", NS) is not None else ""
            runs = []
            for r in p.findall("a:r", NS):
                rpr = r.find("a:rPr", NS)
                t = html.escape(r.findtext("a:t", default="", namespaces=NS))
                sz = int(rpr.get("sz")) / 100 if rpr is not None and rpr.get("sz") else 18
                col = color(rpr.find("a:solidFill", NS), th) if rpr is not None else None
                lat = rpr.find("a:latin", NS) if rpr is not None else None
                font = fonts.get("majorFont") if lat is not None and lat.get("typeface") == "+mj-lt" else fonts.get("minorFont")
                st = "font-size:%.1fpx;font-family:'%s'" % (sz * 96 / 72, font)
                if rpr is not None and rpr.get("b") == "1":
                    st += ";font-weight:700"
                if col:
                    st += ";color:%s" % col
                runs.append('<span style="%s">%s</span>' % (st, t))
            if not runs and sh.is_placeholder:
                continue
            paras.append('<div style="text-align:%s;line-height:1.15">%s%s</div>' % (algn, bullet, "".join(runs) or "&nbsp;"))
        body = "".join(paras)
        if sh.is_placeholder and not body:
            return ""
        if sh.is_placeholder:
            style.append("font-size:%.1fpx;font-family:'%s'" % (32 * 96 / 72, fonts.get("majorFont")))
    return '<div class="sh" style="%s">%s</div>' % (";".join(style), body)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("out")
    ap.add_argument("--slides")
    a = ap.parse_args()
    prs = Presentation(a.deck)
    th, fonts = theme(prs)
    want = set(int(x) for x in a.slides.split(",")) if a.slides else None
    W, H = prs.slide_width * EMU_PX, prs.slide_height * EMU_PX
    parts = ["<!doctype html><meta charset=utf-8><title>Preview</title><style>body{background:#888;margin:0;padding:16px;font-family:sans-serif}"
             ".slide{position:relative;background:#fff;margin:0 auto 24px;overflow:hidden}.sh{position:absolute;box-sizing:border-box;overflow:visible;padding:2px 4px}"
             "h2{color:#fff;font-size:14px;margin:0 0 6px}</style>"]
    for i, slide in enumerate(prs.slides, 1):
        if want and i not in want:
            continue
        parts.append('<h2>Slide %d: %s</h2><div class="slide" style="width:%.0fpx;height:%.0fpx">' % (i, html.escape(slide.slide_layout.name), W, H))
        for sh in slide.shapes:
            parts.append(shape_html(sh, th, fonts))
        parts.append("</div>")
    open(a.out, "w", encoding="utf-8").write("".join(parts))
    print(a.out)


if __name__ == "__main__":
    sys.exit(main())
