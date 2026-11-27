"""Layout engine for the beginner-friendly slide revision (06_revise_slides.py).

Every shape added through a Canvas is also recorded in a display list, so the same geometry can be
rendered to a PNG preview with PIL (Carlito = metric-compatible with Calibri) for visual QA.
Text is measured with the real font metrics, cards grow to fit their text, and anything that would
run past the safe area is reported in WARN.
"""
import copy
import re
from pathlib import Path

from lxml import etree
from PIL import Image, ImageDraw, ImageFont
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

EMU = 914400
FONT_DIR = Path(__file__).resolve().parent / "fonts"
WARN = []

# ETEH palette (taken from the original decks)
NAVY, ORANGE, GREEN, RED = "00213F", "FF6B00", "0F6B3A", "B42B1E"
INK, GREY, BORDER, CREAM, WHITE = "1B2B3A", "5B6B7A", "DDD8CC", "FAF8F4", "FFFFFF"
AMBER, BAND = "B7791F", "F4F1EA"
LINE = 1.22  # single line spacing as a multiple of the font size

_FONTS = {}


def font(size, bold=False, scale=1.0):
    key = (round(size * scale, 2), bold)
    if key not in _FONTS:
        w = "700" if bold else "400"
        px = max(1, int(round(size * scale)))
        _FONTS[key] = (ImageFont.truetype(str(FONT_DIR / f"carlito-latin-{w}-normal.woff"), px),
                       ImageFont.truetype(str(FONT_DIR / f"carlito-latin-ext-{w}-normal.woff"), px))
    return _FONTS[key]


def text_width_pt(s, size, bold=False):
    """Width of a string in points, using 100x scaled metrics for precision."""
    lat, ext = font(size, bold, 100 / size)
    w = 0.0
    for ch in s:
        f = lat if ord(ch) < 0x250 or ch in "–—‘’“”•…›‹€" else ext
        w += f.getlength(ch)
    return w / 100.0


# ---------------------------------------------------------------- rich text: **bold** segments
def segments(text):
    out, bold = [], False
    for part in re.split(r"(\*\*)", text):
        if part == "**":
            bold = not bold
        elif part:
            out.append((part, bold))
    return out


def wrap(text, size, width_in, bold=False):
    """Word-wrap one paragraph. Returns list of lines; each line is a list of (text, bold) runs."""
    words = []
    for seg, b in segments(text):
        for i, wd in enumerate(re.split(r"( )", seg)):
            if wd == "":
                continue
            words.append((wd, b or bold))
    lines, cur, cur_w = [], [], 0.0
    maxw = width_in * 72
    for wd, b in words:
        if wd == " ":
            if cur:
                cur.append((" ", b))
                cur_w += text_width_pt(" ", size, b)
            continue
        ww = text_width_pt(wd, size, b)
        if cur and cur_w + ww > maxw:
            while cur and cur[-1][0] == " ":
                cur.pop()
            lines.append(cur)
            cur, cur_w = [], 0.0
        cur.append((wd, b))
        cur_w += ww
    while cur and cur[-1][0] == " ":
        cur.pop()
    lines.append(cur)
    return lines


def measure(paras, size, width_in, bold=False, gap=0.0):
    """Height in inches of a list of paragraphs (gap = extra space after each paragraph, in inches)."""
    n = sum(len(wrap(p, size, width_in, bold)) for p in paras)
    return n * size * LINE / 72 + gap * max(0, len(paras) - 1), n


# ---------------------------------------------------------------- canvas
def _rgb(h):
    return RGBColor.from_string(h)


class Canvas:
    def __init__(self, slide, where=""):
        self.slide = slide
        self.ops = []
        self.where = where

    # -- primitives -------------------------------------------------------
    def rect(self, x, y, w, h, fill=WHITE, line=None, rounded=False, name=None):
        shp = self.slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
                                          Emu(int(x * EMU)), Emu(int(y * EMU)), Emu(int(w * EMU)), Emu(int(h * EMU)))
        if rounded:
            shp.adjustments[0] = 0.25
        shp.shadow.inherit = False
        if fill:
            shp.fill.solid()
            shp.fill.fore_color.rgb = _rgb(fill)
        else:
            shp.fill.background()
        if line:
            shp.line.color.rgb = _rgb(line)
            shp.line.width = Pt(0.75)
        else:
            shp.line.fill.background()
        if name:
            shp.name = name
        shp.text_frame.text = ""
        self.ops.append(("rect", x, y, w, h, fill, line, rounded))
        return shp

    def text(self, x, y, w, h, paras, size, color=INK, bold=False, align="l", anchor="t", gap=0.0, name=None):
        tb = self.slide.shapes.add_textbox(Emu(int(x * EMU)), Emu(int(y * EMU)), Emu(int(w * EMU)), Emu(int(h * EMU)))
        if name:
            tb.name = name
        tf = tb.text_frame
        tf.word_wrap = True
        tf.auto_size = None
        bp = tf._txBody.find(qn("a:bodyPr"))
        for k in ("lIns", "rIns", "tIns", "bIns"):
            bp.set(k, "0")
        tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
        for i, ptxt in enumerate(paras):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
            if gap and i < len(paras) - 1:
                p.space_after = Pt(gap * 72)
            for seg, b in segments(ptxt):
                r = p.add_run()
                r.text = seg
                f = r.font
                f.name = "Calibri"
                f.size = Pt(size)
                f.bold = bool(b or bold)
                f.color.rgb = _rgb(color)
        self.ops.append(("text", x, y, w, h, list(paras), size, color, bold, align, anchor, gap))
        need, _ = measure(paras, size, w, bold, gap)
        if need > h + 0.01:
            WARN.append(f"CLIP {self.where}: text needs {need:.2f}in, box {h:.2f}in: {paras[0][:60]}")
        return tb

    def check(self, x, y, w, h, color=WHITE):
        """Font-independent tick mark (freeform polyline)."""
        fb = self.slide.shapes.build_freeform(Emu(int(x * EMU)), Emu(int((y + h * 0.55) * EMU)), scale=1.0)
        fb.add_line_segments([(Emu(int((x + w * 0.38) * EMU)), Emu(int((y + h) * EMU))),
                              (Emu(int((x + w) * EMU)), Emu(int(y * EMU)))], close=False)
        shp = fb.convert_to_shape()
        shp.fill.background()
        shp.line.color.rgb = _rgb(color)
        shp.line.width = Pt(2.25)
        shp.shadow.inherit = False
        shp.name = "Tick"
        self.ops.append(("tick", x, y, w, h, color))
        return shp

    # -- table --------------------------------------------------------------
    def table(self, x, y, w, rows, size=12, widths=None, name="Example Table"):
        ncol = max(len(r) for r in rows)
        rows = [list(r) + [""] * (ncol - len(r)) for r in rows]
        if widths is None:
            weight = []
            for c in range(ncol):
                lens = [text_width_pt(str(r[c]).replace("**", ""), size, i == 0) for i, r in enumerate(rows)]
                weight.append(max(55.0, min(max(lens), 260.0)) + 0.35 * sum(lens) / len(lens))
            tot = sum(weight)
            widths = [w * k / tot for k in weight]
        else:
            tot = sum(widths)
            widths = [w * k / tot for k in widths]
        pad_x, pad_y = 0.08, 0.045
        heights = []
        for i, r in enumerate(rows):
            n = max(len(wrap(str(cell), size, widths[c] - 2 * pad_x, i == 0)) for c, cell in enumerate(r))
            heights.append(n * size * LINE / 72 + 2 * pad_y + 0.02)
            if n > 3:
                WARN.append(f"TABLE {self.where}: row {i} wraps to {n} lines: {r}")
        H = sum(heights)
        gf = self.slide.shapes.add_table(len(rows), ncol, Emu(int(x * EMU)), Emu(int(y * EMU)),
                                         Emu(int(w * EMU)), Emu(int(H * EMU)))
        gf.name = name
        tbl = gf.table
        tblPr = tbl._tbl.tblPr
        sid = tblPr.find(qn("a:tableStyleId"))
        if sid is not None:
            tblPr.remove(sid)
        tblPr.set("firstRow", "1")
        tblPr.set("bandRow", "0")
        for c, cw in enumerate(widths):
            tbl.columns[c].width = Emu(int(cw * EMU))
        for i, r in enumerate(rows):
            tbl.rows[i].height = Emu(int(heights[i] * EMU))
            for c, val in enumerate(r):
                cell = tbl.cell(i, c)
                cell.margin_left = cell.margin_right = Emu(int(pad_x * EMU))
                cell.margin_top = cell.margin_bottom = Emu(int(pad_y * EMU))
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf = cell.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                for seg, b in segments(str(val)):
                    run = p.add_run()
                    run.text = seg
                    run.font.name = "Calibri"
                    run.font.size = Pt(size)
                    run.font.bold = bool(b or i == 0)
                    run.font.color.rgb = _rgb(WHITE if i == 0 else INK)
                fill = NAVY if i == 0 else (WHITE if i % 2 else BAND)
                cell.fill.solid()
                cell.fill.fore_color.rgb = _rgb(fill)
                tcPr = cell._tc.get_or_add_tcPr()
                for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
                    ln = etree.SubElement(tcPr, qn(tag), w="6350")
                    sf = etree.SubElement(ln, qn("a:solidFill"))
                    etree.SubElement(sf, qn("a:srgbClr"), val=BORDER)
                # schema order: borders must precede the fill
                fill_el = tcPr.find(qn("a:solidFill"))
                if fill_el is not None:
                    tcPr.remove(fill_el)
                    tcPr.append(fill_el)
        self.ops.append(("table", x, y, w, rows, widths, heights, size))
        return H

    # -- composite: a card with accent bar, label and body -------------------
    def card_height(self, w, paras, size, label=True, gap=0.06):
        h, _ = measure(paras, size, w - 0.50, False, gap)
        return (0.40 if label else 0.14) + h + 0.14

    def card(self, x, y, w, h, accent, label, paras, size, gap=0.06, body_color=INK):
        self.rect(x, y, w, h, WHITE, BORDER)
        self.rect(x, y, 0.09, h, accent)
        ty = y + 0.12
        if label:
            self.text(x + 0.30, ty, w - 0.45, 0.26, [label], 11.5, accent, bold=True)
            ty += 0.28
        bh, _ = measure(paras, size, w - 0.50, False, gap)
        self.text(x + 0.30, ty, w - 0.50, max(bh, 0.2), paras, size, body_color, gap=gap)


# ---------------------------------------------------------------- copying template elements
def copy_shape(src_shape, dst_slide):
    el = copy.deepcopy(src_shape._element)
    dst_slide.shapes._spTree.append(el)
    return el


def copy_picture(src_shape, src_slide, dst_slide):
    el = copy.deepcopy(src_shape._element)
    blip = el.find(".//" + qn("a:blip"))
    rid = blip.get(qn("r:embed"))
    part = src_slide.part.related_part(rid)
    new_rid = dst_slide.part.relate_to(part, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
    blip.set(qn("r:embed"), new_rid)
    dst_slide.shapes._spTree.append(el)
    return el


def set_runs_text(shape, text, size=None):
    """Replace the text of a copied text shape, keeping its first run's formatting."""
    tf = shape.text_frame
    p0 = tf.paragraphs[0]
    runs = p0.runs
    for p in tf.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
        if size:
            runs[0].font.size = Pt(size)
    else:
        p0.text = text


# ---------------------------------------------------------------- preview renderer
def render(ops, path, extra=(), dpi=110):
    W, H = int(13.333 * dpi), int(7.5 * dpi)
    im = Image.new("RGB", (W, H), "#" + CREAM)
    d = ImageDraw.Draw(im)
    s = dpi  # pixels per inch
    pt = dpi / 72

    def col(h):
        return "#" + h

    def draw_lines(x, y, w, paras, size, color, bold, align, gap=0.0):
        yy = y
        for ptxt in paras:
            for line in wrap(ptxt, size, w, bold):
                lw = sum(text_width_pt(t, size, b) for t, b in line) * pt
                xx = x * s if align == "l" else (x + w) * s - lw if align == "r" else x * s + (w * s - lw) / 2
                for t, b in line:
                    lat, ext = font(size, b, pt)
                    for ch in t:
                        f = lat if ord(ch) < 0x250 or ch in "–—‘’“”•…›‹€" else ext
                        d.text((xx, yy * s), ch, font=f, fill=col(color))
                        xx += f.getlength(ch)
                yy += size * LINE / 72
            yy += gap
    for op in list(extra) + list(ops):
        k = op[0]
        if k == "rect":
            _, x, y, w, h, fill, line, rounded = op
            box = [x * s, y * s, (x + w) * s, (y + h) * s]
            if rounded:
                d.rounded_rectangle(box, radius=min(w, h) * s * 0.25, fill=col(fill) if fill else None,
                                    outline=col(line) if line else None)
            else:
                d.rectangle(box, fill=col(fill) if fill else None, outline=col(line) if line else None)
        elif k == "text":
            _, x, y, w, h, paras, size, color, bold, align, anchor, gap = op
            th, _ = measure(paras, size, w, bold, gap)
            yy = y if anchor == "t" else y + (h - th) / 2 if anchor == "m" else y + h - th
            draw_lines(x, yy, w, paras, size, color, bold, align, gap)
        elif k == "tick":
            _, x, y, w, h, color = op
            d.line([(x * s, (y + h * 0.55) * s), ((x + w * 0.38) * s, (y + h) * s), ((x + w) * s, y * s)],
                   fill=col(color), width=max(2, int(2.25 * pt)))
        elif k == "table":
            _, x, y, w, rows, widths, heights, size = op
            yy = y
            for i, r in enumerate(rows):
                xx = x
                for c, val in enumerate(r):
                    fill = NAVY if i == 0 else (WHITE if i % 2 else BAND)
                    d.rectangle([xx * s, yy * s, (xx + widths[c]) * s, (yy + heights[i]) * s], fill=col(fill), outline=col(BORDER))
                    th, _ = measure([str(val)], size, widths[c] - 0.16, i == 0)
                    draw_lines(xx + 0.08, yy + (heights[i] - th) / 2, widths[c] - 0.16, [str(val)], size,
                               WHITE if i == 0 else INK, i == 0, "l")
                    xx += widths[c]
                yy += heights[i]
    im.save(path)
