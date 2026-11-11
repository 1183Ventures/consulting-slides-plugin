#!/usr/bin/env bash
# Scaffold for the review-from-pptx eval case.
# Writes deck.pptx into the current (empty) eval workspace: a minimal
# PowerPoint package with six slides and one speaker note, built with python3's standard
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
NOTES_TYPE = R + "/notesSlide"

# (title placeholder type, title, body paragraphs, table rows, hidden, notes)
SLIDES = [
    ("ctrTitle", "Project Atlas: steering committee pre-read",
     ["Operations cost review, October 2026"], None, False, None),
    ("title", "Agenda",
     ["09:00 Welcome and objectives", "09:30 Where the cost growth came from", "10:45 What we recommend"],
     None, False, None),
    ("title", "Freight drove 70% of the $52M cost increase in 2025",
     ["Increase in operating cost by driver ($M)"],
     [["Driver", "Increase"], ["Freight", "38"], ["Labor", "12"], ["Materials", "7"], ["Total", "52"]],
     False, "Partner asked whether 52 is the right total. Check with finance before Friday."),
    ("title", "Moving to advice hubs saves [$X]M a year",
     ["Click to add text"], None, False, None),
    ("title", "We recommend a carrier RFP in Q1 2027",
     ["Launch the RFP in January", "Award contracts by March"], None, False, None),
    ("title", "Old appendix: 2023 cost baseline",
     ["Superseded by the 2025 baseline"], None, True, None),
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


def slide_xml(ph, title, body, rows, hidden, notes):
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
    + "".join(
        '<Override PartName="/ppt/notesSlides/notesSlide%d.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>' % n
        for n, spec in enumerate(SLIDES, start=1) if spec[5]
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
        if spec[5]:
            z.writestr("ppt/slides/_rels/slide%d.xml.rels" % n, (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="%s"><Relationship Id="rId1" Type="%s" Target="../notesSlides/notesSlide%d.xml"/></Relationships>'
            ) % (REL, NOTES_TYPE, n))
            z.writestr("ppt/notesSlides/notesSlide%d.xml" % n, (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<p:notes xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"><p:cSld><p:spTree>'
                '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
                '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder 1"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>'
                '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>%s</p:txBody></p:sp>'
                '</p:spTree></p:cSld></p:notes>'
            ) % (A, R, P, para(spec[5])))
PY
