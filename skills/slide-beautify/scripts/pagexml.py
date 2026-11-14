"""Library pages as PresentationML shapes.

A library page is a list of plain shape descriptions in inches and points,
the same descriptions SlideShortcuts draws its templates from:

  {"kind": "sp", "geom", "x", "y", "w", "h", "fill", "line", "text", "textBox", "rot", "alpha"}
  {"kind": "cxn", "x1", "y1", "x2", "y2", "color", "weight", "dash", "head", "tail"}
  {"kind": "ph", "type": "title" | "source", "x", "y", "w", "h", "text"}

Colours are theme colours by name, so the page takes the deck's own theme:
a1 to a6 (accent1 to accent6), tx1, tx2, bg1, bg2. "a1/20" is a tint (20%
of the colour, the rest white), "a1*70" a shade (70% of its lightness), and
"#RRGGBB" a fixed colour kept for meanings a brand must not change (status
red, amber and green). Fonts are the theme's: a run with "font": "head" uses
the heading font, every other run the body font.

Text: {"paras": [para], "anchor", "inset": [l, t, r, b], "wrap", "vert"};
a para {"runs": [run], "align", "bullet", "bulletColor", "level", "before",
"after", "spacing"}; a run {"t", "b", "i", "sz", "color", "font", "caps"}.

Standard library only. `Place` maps the page's 13.333 x 7.5 inch canvas onto
the deck's slide.
"""

import re
from xml.sax.saxutils import escape as _esc

EMU = 914400
PT = 12700
CANVAS_W = 13.333
CANVAS_H = 7.5

SCHEME = {
    "a1": "accent1", "a2": "accent2", "a3": "accent3", "a4": "accent4",
    "a5": "accent5", "a6": "accent6", "tx1": "tx1", "tx2": "tx2",
    "bg1": "bg1", "bg2": "bg2",
}
DASHES = {"dash": "dash", "dot": "sysDot", "dashDot": "dashDot", "longDash": "lgDash"}
FONT_REF = {"head": "+mj-lt", "body": "+mn-lt"}
BULLET_INDENT = 0.17
LEVEL_STEP = 0.2
DEFAULT_INSET = [0.08, 0.05, 0.08, 0.05]

_BAD = re.compile("[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]")


def esc(text):
    return _esc(_BAD.sub("", str(text)), {'"': "&quot;"})


def parse_color(spec):
    """"a1/20" -> ("accent1", tint 20, None); "#C0504D" -> srgb."""
    if spec is None or spec == "none":
        return None
    s = str(spec).strip()
    m = re.fullmatch(r"#([0-9A-Fa-f]{6})", s)
    if m:
        return {"srgb": m.group(1).upper()}
    m = re.fullmatch(r"([a-z]+\d?)(?:([/*])(\d{1,3}))?", s)
    if not m or m.group(1) not in SCHEME:
        raise ValueError("Unknown colour " + s)
    out = {"scheme": SCHEME[m.group(1)]}
    if m.group(2) == "/":
        out["tint"] = min(100, int(m.group(3)))
    if m.group(2) == "*":
        out["shade"] = min(100, int(m.group(3)))
    return out


def color_xml(spec, alpha=None):
    c = parse_color(spec)
    if not c:
        return ""
    mods = []
    if c.get("tint") is not None and c["tint"] < 100:
        mods.append('<a:lumMod val="%d"/><a:lumOff val="%d"/>' % (c["tint"] * 1000, (100 - c["tint"]) * 1000))
    if c.get("shade") is not None and c["shade"] < 100:
        mods.append('<a:lumMod val="%d"/>' % (c["shade"] * 1000))
    if alpha is not None and alpha < 100:
        mods.append('<a:alpha val="%d"/>' % round(alpha * 1000))
    if "srgb" in c:
        return '<a:srgbClr val="%s">%s</a:srgbClr>' % (c["srgb"], "".join(mods))
    return '<a:schemeClr val="%s">%s</a:schemeClr>' % (c["scheme"], "".join(mods))


def fill_xml(fill, alpha=None):
    if fill is None or fill == "none":
        return "<a:noFill/>"
    return "<a:solidFill>%s</a:solidFill>" % color_xml(fill, alpha)


def line_xml(line, scale=1.0):
    if not line:
        return "<a:ln><a:noFill/></a:ln>"
    w = round((0.75 if line.get("weight") is None else line["weight"]) * PT * max(scale, 0.6))
    parts = ['<a:ln w="%d">' % w, fill_xml(line.get("color") or "tx1")]
    if line.get("dash") in DASHES:
        parts.append('<a:prstDash val="%s"/>' % DASHES[line["dash"]])
    if line.get("head"):
        parts.append('<a:headEnd type="%s" w="med" len="med"/>' % line["head"])
    if line.get("tail"):
        parts.append('<a:tailEnd type="%s" w="med" len="med"/>' % line["tail"])
    parts.append("</a:ln>")
    return "".join(parts)


ROUND = ("ellipse", "triangle", "diamond", "chevron", "homePlate", "rightArrow")


LIB_TITLE_BOTTOM = 1.46   # inches: the library's title box ends here (frame.js, 16:9: 0.56 + 0.9)


class Place:
    """The page's canvas (13.333 x 7.5 in) on the deck's slide.

    Across and down are scaled on their own (a 4:3 slide is narrower, not
    shorter); text, line weights and insets by the mean of the two, and
    round shapes keep their proportions. The content band moves below the
    deck's own title when that title ends lower than the page's.
    """

    def __init__(self, slide_w_emu, slide_h_emu, content_top=None, title_bottom_emu=None):
        self.sx = slide_w_emu / (CANVAS_W * EMU)
        self.sy = slide_h_emu / (CANVAS_H * EMU)
        self.s = (self.sx * self.sy) ** 0.5
        self.top = content_top or 0.0
        self.k = 1.0   # how much the content band is squeezed down
        self.dy = 0.0
        if content_top is not None and title_bottom_emu is not None:
            # the content keeps 0.12 in clear of the deck's title box, or the
            # clearance the page itself keeps under the library's own title
            # box (LIB_TITLE_BOTTOM) when that is less: a page whose column
            # heads start above it is never squeezed under a title that ends
            # where the library's does
            clear = min(0.12, content_top - LIB_TITLE_BOTTOM)
            want = title_bottom_emu + clear * EMU * self.sy
            have = content_top * EMU * self.sy
            if want > have:
                bottom = (CANVAS_H - 0.35) * EMU * self.sy
                self.k = max(0.75, (bottom - want) / max(bottom - have, 1))
                self.dy = want - have

    def x(self, inches):
        return round(inches * EMU * self.sx)

    def y(self, inches, content=True):
        if not content:
            return round(inches * EMU * self.sy)
        base = self.top * EMU * self.sy
        return round(base + self.dy + (inches - self.top) * EMU * self.sy * self.k)

    def dx(self, inches):
        return round(max(inches, 0) * EMU * self.sx)

    def dh(self, inches):
        return round(max(inches, 0) * EMU * self.sy * self.k)

    def d(self, inches):
        return round(max(inches, 0) * EMU * self.s)

    def pt(self, size):
        return max(7.0, round(size * min(self.s, self.sy * self.k) * 2) / 2)

    def spacing(self, points):
        """Paragraph spacing on the slide: scaled as the text is (no rounding, no floor)."""
        return points * min(self.s, self.sy * self.k)

    def box(self, x, y, w, h, geom=None):
        """x, y, width, height in EMU; round shapes keep their proportions."""
        X, Y, W, H = self.x(x), self.y(y), self.dx(w), self.dh(h)
        if geom in ROUND and w > 0 and h > 0:
            f = min(W / (w * EMU), H / (h * EMU))
            W2, H2 = round(w * EMU * f), round(h * EMU * f)
            X, Y, W, H = X + (W - W2) // 2, Y + (H - H2) // 2, W2, H2
        return X, Y, W, H

    def width_pt(self, inches):
        return max(inches, 0) * 72 * self.sx

    def height_pt(self, inches):
        return max(inches, 0) * 72 * self.sy * self.k


def _rpr(run, tag, place, size_override=None):
    attrs = ['lang="en-US"']
    sz = size_override if size_override is not None else run.get("sz")
    if sz:
        attrs.append('sz="%d"' % round(place.pt(sz) * 100 if size_override is None else sz * 100))
    if run.get("b") is not None:
        attrs.append('b="%d"' % (1 if run["b"] else 0))
    if run.get("i") is not None:
        attrs.append('i="%d"' % (1 if run["i"] else 0))
    if run.get("caps"):
        attrs.append('cap="all"')
    if run.get("baseline"):
        attrs.append('baseline="%s"' % run["baseline"])
    attrs.append('dirty="0"')
    inner = []
    if run.get("color"):
        inner.append(fill_xml(run["color"]))
    if run.get("font") in FONT_REF:
        inner.append('<a:latin typeface="%s"/>' % FONT_REF[run["font"]])
    return "<a:%s %s>%s</a:%s>" % (tag, " ".join(attrs), "".join(inner), tag)


def para_xml(p, place, size_override=None):
    level = p.get("level") or 0
    attrs = []
    if p.get("bullet"):
        attrs += ['marL="%d"' % place.d(level * LEVEL_STEP + BULLET_INDENT), 'indent="%d"' % -place.d(BULLET_INDENT)]
    elif level:
        attrs += ['marL="%d"' % place.d(level * LEVEL_STEP), 'indent="0"']
    else:
        attrs += ['marL="0"', 'indent="0"']
    if level:
        attrs.append('lvl="%d"' % level)
    attrs.append('algn="%s"' % (p.get("align") or "l"))
    inner = []
    if p.get("spacing"):
        inner.append('<a:lnSpc><a:spcPct val="%d"/></a:lnSpc>' % round(p["spacing"] * 1000))
    # space before and after scales with the text, as box heights do (Reviewer, 14 Nov 2026:
    # unscaled spacing left fitted boxes short on slides under 7.5 in tall)
    inner.append('<a:spcBef><a:spcPts val="%d"/></a:spcBef>' % round(place.spacing(p.get("before") or 0) * 100))
    inner.append('<a:spcAft><a:spcPts val="%d"/></a:spcAft>' % round(place.spacing(p.get("after") or 0) * 100))
    if p.get("bullet"):
        ch = "•" if p["bullet"] is True else p["bullet"]
        if p.get("bulletColor"):
            inner.append("<a:buClr>%s</a:buClr>" % color_xml(p["bulletColor"]))
        inner.append('<a:buFont typeface="Arial"/><a:buChar char="%s"/>' % esc(ch))
    else:
        inner.append("<a:buNone/>")
    runs = [r for r in (p.get("runs") or []) if r and r.get("t") not in (None, "")]
    body = []
    for r in runs:
        t = str(r["t"])
        space = ' xml:space="preserve"' if re.match(r"^\s|.*\s$", t) else ""
        body.append("<a:r>%s<a:t%s>%s</a:t></a:r>" % (_rpr(r, "rPr", place, size_override), space, esc(t)))
    last = runs[-1] if runs else ((p.get("runs") or [{}])[0] or {})
    end = _rpr({"sz": last.get("sz"), "b": last.get("b")}, "endParaRPr", place, size_override)
    return '<a:p><a:pPr %s>%s</a:pPr>%s%s</a:p>' % (" ".join(attrs), "".join(inner), "".join(body), end)


def body_pr_xml(text, place):
    ins = text.get("inset") or DEFAULT_INSET
    attrs = [
        'wrap="%s"' % ("none" if text.get("wrap") is False else "square"),
        'lIns="%d"' % place.d(ins[0]), 'tIns="%d"' % place.d(ins[1]),
        'rIns="%d"' % place.d(ins[2]), 'bIns="%d"' % place.d(ins[3]),
        'anchor="%s"' % (text.get("anchor") or "t"), 'rtlCol="0"',
    ]
    if text.get("vert"):
        attrs.append('vert="%s"' % text["vert"])
    if text.get("anchorCtr"):
        attrs.append('anchorCtr="1"')
    # as the library writes it (ooxml.js bodyPrXml): Resize Shape to Fit Text
    # for a box the layout pass fitted, Shrink Text on Overflow where asked
    fit = {"shrink": "<a:normAutofit/>", "resize": "<a:spAutoFit/>"}.get(text.get("autofit"), "<a:noAutofit/>")
    return "<a:bodyPr %s>%s</a:bodyPr>" % (" ".join(attrs), fit)


def text_body_xml(text, place, size_override=None):
    text = text or {"paras": [{"runs": []}]}
    paras = text.get("paras") or [{"runs": []}]
    return "<p:txBody>%s<a:lstStyle/>%s</p:txBody>" % (
        body_pr_xml(text, place), "".join(para_xml(p, place, size_override) for p in paras))


def _xfrm(place, x, y, w, h, rot=0, flip_h=False, flip_v=False, geom=None):
    extra = ""
    if rot:
        extra += ' rot="%d"' % round(rot * 60000)
    if flip_h:
        extra += ' flipH="1"'
    if flip_v:
        extra += ' flipV="1"'
    X, Y, W, H = place.box(x, y, w, h, geom)
    return '<a:xfrm%s><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>' % (extra, X, Y, W, H)


def adj_xml(adj):
    """A preset's adjust values (a pill's corner, a chevron's point), as
    ooxml.js geomXml writes them. They are relative to the shape's short side,
    so a box the pass resizes keeps its look."""
    return "".join('<a:gd name="%s" fmla="val %d"/>' % (esc(k), round(v)) for k, v in (adj or {}).items())


def sp_xml(s, sid, place, name=None, size_override=None):
    nm = esc(name or s.get("name") or "Shape %d" % sid)
    tx = ' txBox="1"' if s.get("textBox") else ""
    return (
        '<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr%s/><p:nvPr/></p:nvSpPr>'
        '<p:spPr>%s<a:prstGeom prst="%s"><a:avLst>%s</a:avLst></a:prstGeom>%s%s</p:spPr>%s</p:sp>'
    ) % (sid, nm, tx, _xfrm(place, s["x"], s["y"], s["w"], s["h"], s.get("rot") or 0, geom=s.get("geom")),
         s.get("geom") or "rect", adj_xml(s.get("adj")), fill_xml(s.get("fill"), s.get("alpha")),
         line_xml(s.get("line"), place.s), text_body_xml(s.get("text"), place, size_override))


def cxn_xml(s, sid, place):
    x, y = min(s["x1"], s["x2"]), min(s["y1"], s["y2"])
    line = {k: s.get(k) for k in ("color", "weight", "dash", "head", "tail")}
    line["color"] = line["color"] or "tx1"
    return (
        '<p:cxnSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<p:nvCxnSpPr><p:cNvPr id="%d" name="Connector %d"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>'
        '<p:spPr>%s<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/>%s</p:spPr></p:cxnSp>'
    ) % (sid, sid, _xfrm(place, x, y, abs(s["x2"] - s["x1"]), abs(s["y2"] - s["y1"]), 0,
                          s["x2"] < s["x1"], s["y2"] < s["y1"]), line_xml(line, place.s))


def source_xml(s, sid, place):
    """The page's source line, as a plain text box in the deck's body font."""
    box = dict(s, kind="sp", geom="rect", textBox=True)
    text = dict(box.get("text") or {})
    text.setdefault("anchor", "b")
    text.setdefault("inset", [0, 0, 0, 0])
    paras = []
    for p in text.get("paras") or [{"runs": []}]:
        runs = [dict({"sz": 10, "color": "tx1"}, **r) for r in p.get("runs") or []]
        paras.append(dict(p, runs=runs))
    text["paras"] = paras
    box["text"] = text
    return sp_xml(box, sid, place, name="Source")
