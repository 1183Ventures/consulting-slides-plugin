#!/usr/bin/env bash
# Scaffold for the storyline-from-pptx eval case.
# Writes deck.pptx into the current (empty) eval workspace: a minimal
# PowerPoint package with seven slides, built with python3's standard
# library. It reads nothing else and makes no network calls.
# It runs only when you pass --scaffold to claude plugin eval.
set -euo pipefail

python3 - <<'PY'
import zipfile
from xml.sax.saxutils import escape

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
SLIDE_TYPE = R + "/slide"
DOC_TYPE = R + "/officeDocument"

# (title placeholder type, title, body paragraphs, table rows, hidden)
SLIDES = [
    ("ctrTitle", "Project Atlas: steering committee update",
     ["Operations cost review, October 2026"], None, False),
    ("title", "Agenda",
     ["Where costs went", "Why freight grew", "What we recommend"], None, False),
    ("title", "Background",
     ["Cost base of $480M in 2025 across manufacturing, freight and overhead"], None, False),
    ("title", "Operating costs rose 12% in 2025, and freight drove two thirds of the increase",
     ["Increase in operating cost by driver, 2024 to 2025"],
     [["Driver", "Increase"], ["Freight", "$38M"], ["Labor", "$12M"], ["Materials", "$7M"]],
     False),
    ("title", "Freight analysis",
     ["Spot-market share of loads rose from 15% to 41%",
      "Average spot lane rate is 22% above contract rate"], None, False),
    ("title", "Old appendix: 2023 cost baseline",
     ["Superseded by the 2025 baseline"], None, True),
    ("title", "Next steps",
     ["Agree on contract strategy", "Launch carrier RFP"], None, False),
]


def run(text):
    return '<a:r><a:rPr lang="en-US"/><a:t>%s</a:t></a:r>' % escape(text)


def para(text):
    return "<a:p>%s</a:p>" % run(text)


def shape(shape_id, name, ph, paragraphs, y):
    ph_xml = '<p:ph type="%s"/>' % ph if ph in ("title", "ctrTitle", "subTitle") else '<p:ph idx="1"/>'
    return (
        "<p:sp><p:nvSpPr>"
        '<p:cNvPr id="%d" name="%s"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        "<p:nvPr>%s</p:nvPr></p:nvSpPr>"
        '<p:spPr><a:xfrm><a:off x="457200" y="%d"/><a:ext cx="8229600" cy="1143000"/></a:xfrm></p:spPr>'
        "<p:txBody><a:bodyPr/><a:lstStyle/>%s</p:txBody></p:sp>"
    ) % (shape_id, name, ph_xml, y, "".join(para(t) for t in paragraphs))


def table(rows, y):
    cells = "".join(
        "<a:tr h=\"370840\">%s</a:tr>" % "".join(
            "<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>%s</a:txBody><a:tcPr/></a:tc>" % para(c)
            for c in row
        )
        for row in rows
    )
    return (
        "<p:graphicFrame><p:nvGraphicFramePr>"
        '<p:cNvPr id="9" name="Table 8"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/>'
        '</p:nvGraphicFramePr><p:xfrm><a:off x="457200" y="%d"/><a:ext cx="6096000" cy="1483360"/></p:xfrm>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
        '<a:tbl><a:tblPr firstRow="1"/><a:tblGrid><a:gridCol w="3048000"/><a:gridCol w="3048000"/></a:tblGrid>%s</a:tbl>'
        "</a:graphicData></a:graphic></p:graphicFrame>"
    ) % (y, cells)


def slide_xml(ph, title, body, rows, hidden):
    sub = "subTitle" if ph == "ctrTitle" else "body"
    parts = [shape(2, "Title 1", ph, [title], 365125), shape(3, "Content 2", sub, body, 1825625)]
    if rows:
        parts.append(table(rows, 3200400))
    show = ' show="0"' if hidden else ""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sld xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"%s><p:cSld><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
        "%s</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
    ) % (A, R, P, show, "".join(parts))


content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
    + "".join(
        '<Override PartName="/ppt/slides/slide%d.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' % n
        for n in range(1, len(SLIDES) + 1)
    )
    + "</Types>"
)
root_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="%s"><Relationship Id="rId1" Type="%s" Target="ppt/presentation.xml"/></Relationships>'
) % (REL, DOC_TYPE)
presentation = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<p:presentation xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"><p:sldIdLst>%s</p:sldIdLst>'
    '<p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>'
) % (A, R, P, "".join('<p:sldId id="%d" r:id="rId%d"/>' % (255 + n, n) for n in range(1, len(SLIDES) + 1)))
pres_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="%s">%s</Relationships>'
) % (REL, "".join(
    '<Relationship Id="rId%d" Type="%s" Target="slides/slide%d.xml"/>' % (n, SLIDE_TYPE, n)
    for n in range(1, len(SLIDES) + 1)
))

with zipfile.ZipFile("deck.pptx", "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", root_rels)
    z.writestr("ppt/presentation.xml", presentation)
    z.writestr("ppt/_rels/presentation.xml.rels", pres_rels)
    for n, spec in enumerate(SLIDES, start=1):
        z.writestr("ppt/slides/slide%d.xml" % n, slide_xml(*spec))
PY
