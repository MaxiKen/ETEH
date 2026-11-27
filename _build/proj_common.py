"""Shared helpers for student projects and instructor capstones: PDF briefs and KPI cards."""
from __future__ import annotations

import math
from pathlib import Path

from xlsxwriter.utility import xl_rowcol_to_cell

from common import FONT_DIR, IPR, SPR

SKEY = IPR / "Student_Project_Keys"
SKEY.mkdir(parents=True, exist_ok=True)

_FONTS = False


def _fonts():
    global _FONTS
    if _FONTS:
        return
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    pdfmetrics.registerFont(TTFont("DejaVu", str(FONT_DIR / "DejaVuSans.ttf")))
    pdfmetrics.registerFont(TTFont("DejaVu-Bold", str(FONT_DIR / "DejaVuSans-Bold.ttf")))
    pdfmetrics.registerFont(TTFont("DejaVu-Oblique", str(FONT_DIR / "DejaVuSans-Oblique.ttf")))
    _FONTS = True


def brief_pdf(path: Path, title: str, subtitle: str, audience: str, blocks: list):
    """blocks: list of ("h", text) | ("p", text) | ("b", [bullets]) | ("t", [[row], ...], col_widths_mm or None) | ("note", text)"""
    _fonts()
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import ListFlowable, ListItem, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    NAVY, TEAL, GOLD = colors.HexColor("#0B2545"), colors.HexColor("#13807A"), colors.HexColor("#E0A526")
    T = ParagraphStyle("t", fontName="DejaVu-Bold", fontSize=18, leading=22, textColor=NAVY)
    ST = ParagraphStyle("st", fontName="DejaVu-Oblique", fontSize=10, leading=13, textColor=colors.HexColor("#555555"))
    H = ParagraphStyle("h", fontName="DejaVu-Bold", fontSize=12, leading=15, textColor=TEAL, spaceBefore=8, spaceAfter=3)
    P = ParagraphStyle("p", fontName="DejaVu", fontSize=9.5, leading=13)
    N = ParagraphStyle("n", fontName="DejaVu-Oblique", fontSize=8.5, leading=11, textColor=colors.HexColor("#555555"), backColor=colors.HexColor("#FFF4C2"),
                       borderPadding=4)
    C = ParagraphStyle("c", fontName="DejaVu", fontSize=8.5, leading=11)
    CB = ParagraphStyle("cb", fontName="DejaVu-Bold", fontSize=8.5, leading=11, textColor=colors.white)

    def footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("DejaVu", 7.5)
        canvas.setFillColor(colors.HexColor("#777777"))
        canvas.drawString(15 * mm, 10 * mm, f"ETEH Ltd · Eaglet2EagleHub Limited · Four-Week Excel Data Analytics Bootcamp · {audience}")
        canvas.drawRightString(195 * mm, 10 * mm, f"Page {doc.page}")
        canvas.setStrokeColor(GOLD)
        canvas.setLineWidth(2)
        canvas.line(15 * mm, 287 * mm, 195 * mm, 287 * mm)
        canvas.restoreState()

    doc = SimpleDocTemplate(str(path), pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm, topMargin=16 * mm, bottomMargin=16 * mm,
                            title=title, author="ETEH Ltd")
    story = [Paragraph(title, T), Spacer(1, 2 * mm), Paragraph(subtitle, ST), Spacer(1, 4 * mm)]
    for blk in blocks:
        kind = blk[0]
        if kind == "h":
            story.append(Paragraph(blk[1], H))
        elif kind == "p":
            story.append(Paragraph(blk[1], P))
            story.append(Spacer(1, 1.5 * mm))
        elif kind == "note":
            story.append(Spacer(1, 1 * mm))
            story.append(Paragraph(blk[1], N))
            story.append(Spacer(1, 3 * mm))
        elif kind == "b":
            story.append(ListFlowable([ListItem(Paragraph(x, P), leftIndent=10) for x in blk[1]], bulletType="bullet", start="•", leftIndent=10,
                                      bulletFontName="DejaVu"))
            story.append(Spacer(1, 1.5 * mm))
        elif kind == "t":
            rows = blk[1]
            widths = blk[2] if len(blk) > 2 and blk[2] else None
            data = [[Paragraph(str(c), CB) for c in rows[0]]] + [[Paragraph(str(c), C) for c in r] for r in rows[1:]]
            t = Table(data, colWidths=[w * mm for w in widths] if widths else None, repeatRows=1)
            t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), NAVY), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                                   ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#99A")),
                                   ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F3F7FA")])]))
            story.append(t)
            story.append(Spacer(1, 3 * mm))
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return path


def kpi_cards(s, cards, c0=1, width=3):
    """Row of KPI tiles. cards: (label, formula, cached, kind, sub). Each tile spans `width` columns and 3 rows."""
    b = s.b
    wb = b.wb
    lab = wb.add_format({"font_name": "Calibri", "font_size": 9, "bold": True, "font_color": "white", "bg_color": "#0B2545", "align": "center", "valign": "vcenter"})
    sub = wb.add_format({"font_name": "Calibri", "font_size": 8, "italic": True, "font_color": "#555555", "bg_color": "#EAF2F8", "align": "center",
                         "valign": "vcenter", "text_wrap": True})
    r = s.r
    s.ws.set_row(r + 1, 34)
    s.ws.set_row(r + 2, 26)
    out = []
    for i, (label, fx, val, kind, subtext) in enumerate(cards):
        a = c0 + i * width
        z = a + width - 1
        s.ws.merge_range(r, a, r, z, label, lab)
        big = b.nf(kind, style="plain", font_size=20, bold=True, font_color="#0B2545", align="center", valign="vcenter", bg_color="#EAF2F8")
        s.ws.merge_range(r + 1, a, r + 1, z, "", big)
        v = val.item() if hasattr(val, "item") else val
        if isinstance(fx, str) and fx.startswith("="):
            s.ws.write_formula(r + 1, a, fx, big, v)
        else:
            s.ws.write(r + 1, a, v, big)
        s.ws.merge_range(r + 2, a, r + 2, z, subtext, sub)
        out.append(xl_rowcol_to_cell(r + 1, a))
    s.r = r + 4
    return out
