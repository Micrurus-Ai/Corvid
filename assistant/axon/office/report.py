"""Code-generated branded PDF reports/proposals (ReportLab). Real data charts, brand colors, and an
optional logo — no image model and no Office dependency, so output is accurate and repeatable."""
import os

from axon.util import _result
from axon.office.brand import get_brand

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
                                Image as RLImage, ListFlowable, ListItem, KeepTogether)
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.linecharts import HorizontalLineChart
from reportlab.graphics.charts.piecharts import Pie


def _hex(c, default="222222"):
    try:
        return colors.HexColor("#" + str(c).lstrip("#"))
    except Exception:
        return colors.HexColor("#" + default)


def _downloads():
    d = os.path.join(os.environ.get("USERPROFILE", os.path.expanduser("~")), "Downloads")
    return d if os.path.isdir(d) else os.path.expanduser("~")


def _palette(b):
    return [_hex(b.get("primary"), "1F3A5F"), _hex(b.get("accent"), "E07A2F"),
            colors.HexColor("#4C8CBF"), colors.HexColor("#7BB662"),
            colors.HexColor("#E0B94E"), colors.HexColor("#9B6FB0")]


def _to_floats(row):
    out = []
    for v in row:
        try:
            out.append(float(v))
        except Exception:
            out.append(0.0)
    return out


def _chart(spec, b):
    """Build a real data chart (bar/line/pie) coloured with the brand palette."""
    ctype = (spec.get("type") or "bar").lower()
    cats = [str(c) for c in (spec.get("categories") or [])]
    series = spec.get("series") or None          # {"Revenue":[...], "Cost":[...]}
    values = spec.get("values") or None          # [..] single series / pie
    pal = _palette(b)
    d = Drawing(460, 240)

    if ctype == "pie":
        vals = _to_floats(values or (list(series.values())[0] if series else []))
        pie = Pie()
        pie.x, pie.y, pie.width, pie.height = 150, 30, 170, 170
        pie.data = vals or [1]
        if cats:
            pie.labels = cats
        for i in range(len(pie.data)):
            pie.slices[i].fillColor = pal[i % len(pal)]
        pie.slices.strokeWidth = 0.5
        d.add(pie)
        return d

    data = [_to_floats(r) for r in (list(series.values()) if series else [values or []])]
    if ctype == "line":
        lc = HorizontalLineChart()
        lc.x, lc.y, lc.width, lc.height = 45, 35, 390, 165
        lc.data = data
        if cats:
            lc.categoryAxis.categoryNames = cats
        for i in range(len(lc.data)):
            lc.lines[i].strokeColor = pal[i % len(pal)]
            lc.lines[i].strokeWidth = 2
        d.add(lc)
    else:  # bar
        bc = VerticalBarChart()
        bc.x, bc.y, bc.width, bc.height = 45, 35, 390, 165
        bc.data = data
        bc.valueAxis.valueMin = 0
        if cats:
            bc.categoryAxis.categoryNames = cats
        for i in range(len(bc.data)):
            bc.bars[i].fillColor = pal[i % len(pal)]
        d.add(bc)
    return d


def _decorate(canvas, doc, b, header, logo):
    """Brand band + optional logo on top, page number + footer at the bottom, every page."""
    w, h = A4
    canvas.saveState()
    canvas.setFillColor(_hex(b.get("primary"), "1F3A5F"))
    canvas.rect(0, h - 16 * mm, w, 16 * mm, fill=1, stroke=0)
    canvas.setFillColor(_hex(b.get("light"), "FFFFFF"))
    canvas.setFont("Helvetica-Bold", 11)
    canvas.drawString(15 * mm, h - 10.5 * mm, (header or "")[:80])
    if logo and os.path.isfile(logo):
        try:
            canvas.drawImage(logo, w - 45 * mm, h - 15 * mm, width=30 * mm, height=12 * mm,
                             preserveAspectRatio=True, mask="auto")
        except Exception:
            pass
    canvas.setFillColor(colors.grey)
    canvas.setFont("Helvetica", 8)
    canvas.drawRightString(w - 15 * mm, 10 * mm, "Page %d" % doc.page)
    foot = b.get("footer") or ""
    if foot:
        canvas.drawString(15 * mm, 10 * mm, foot[:90])
    canvas.restoreState()


def make_report(args):
    """Create a branded PDF report/proposal from structured content (code, not an image model).

    args: title, subtitle?, sections[], logo? (else brand logo), filename?
    Each section: {heading?, text?, bullets?[], table?{headers[],rows[[]]}, chart?{type,title,categories,series|values}}
    """
    title = (args.get("title") or "Report").strip()
    subtitle = (args.get("subtitle") or "").strip()
    sections = args.get("sections") or []
    b = get_brand()
    logo = args.get("logo") or b.get("logo") or ""        # optional — only drawn if it exists

    filename = (args.get("filename") or (title + ".pdf")).strip()
    if not filename.lower().endswith(".pdf"):
        filename += ".pdf"
    filename = "".join(ch for ch in filename if ch not in '<>:"/\\|?*')
    path = os.path.join(_downloads(), filename)

    ss = getSampleStyleSheet()
    h_title = ParagraphStyle("AxTitle", parent=ss["Title"], textColor=_hex(b.get("primary"), "1F3A5F"),
                             fontSize=24, spaceAfter=4, alignment=TA_LEFT)
    h_sub = ParagraphStyle("AxSub", parent=ss["Normal"], textColor=colors.grey, fontSize=12, spaceAfter=16)
    h_head = ParagraphStyle("AxHead", parent=ss["Heading2"], textColor=_hex(b.get("primary"), "1F3A5F"),
                            fontSize=14, spaceBefore=14, spaceAfter=6)
    h_body = ParagraphStyle("AxBody", parent=ss["Normal"], textColor=_hex(b.get("text"), "222222"),
                            fontSize=10.5, leading=15, spaceAfter=6)
    h_chart = ParagraphStyle("AxChart", parent=h_body, textColor=colors.grey, fontSize=10,
                             spaceBefore=6, spaceAfter=2)

    story = [Paragraph(title, h_title)]
    if subtitle:
        story.append(Paragraph(subtitle, h_sub))
    else:
        story.append(Spacer(1, 8))

    for sec in sections:
        blk = []
        if sec.get("heading"):
            blk.append(Paragraph(str(sec["heading"]), h_head))
        for para in ([sec["text"]] if isinstance(sec.get("text"), str) else (sec.get("text") or [])):
            if para:
                blk.append(Paragraph(str(para), h_body))
        if sec.get("bullets"):
            blk.append(ListFlowable(
                [ListItem(Paragraph(str(x), h_body), leftIndent=8) for x in sec["bullets"]],
                bulletType="bullet", start="•", bulletColor=_hex(b.get("accent"), "E07A2F")))
        tbl = sec.get("table")
        if tbl and tbl.get("rows"):
            rows = ([tbl["headers"]] if tbl.get("headers") else []) + tbl["rows"]
            rows = [[str(c) for c in r] for r in rows]
            t = Table(rows, hAlign="LEFT")
            style = [("FONTSIZE", (0, 0), (-1, -1), 9.5),
                     ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D0D3DA")),
                     ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                     ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5)]
            if tbl.get("headers"):
                style += [("BACKGROUND", (0, 0), (-1, 0), _hex(b.get("primary"), "1F3A5F")),
                          ("TEXTCOLOR", (0, 0), (-1, 0), _hex(b.get("light"), "FFFFFF")),
                          ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold")]
            t.setStyle(TableStyle(style))
            blk.append(Spacer(1, 4))
            blk.append(t)
        ch = sec.get("chart")
        if ch and (ch.get("series") or ch.get("values")):
            if ch.get("title"):
                blk.append(Paragraph(str(ch["title"]), h_chart))
            try:
                blk.append(_chart(ch, b))
            except Exception:
                pass
        if blk:
            story.append(KeepTogether(blk) if len(blk) <= 4 else blk[0])
            for extra in (blk[1:] if len(blk) > 4 else []):
                story.append(extra)

    try:
        doc = SimpleDocTemplate(path, pagesize=A4, topMargin=22 * mm, bottomMargin=18 * mm,
                                leftMargin=15 * mm, rightMargin=15 * mm, title=title)
        deco = lambda c, d: _decorate(c, d, b, title, logo)
        doc.build(story, onFirstPage=deco, onLaterPages=deco)
    except Exception as e:
        return _result("Couldn't build the PDF: " + str(e), True)

    try:
        os.startfile(path)   # open it so the user sees the result
    except Exception:
        pass
    return _result("Created a branded PDF report: " + path +
                   ("" if logo else "  (no logo used — set one with set_brand, or pass logo=path, if you want it)"))
