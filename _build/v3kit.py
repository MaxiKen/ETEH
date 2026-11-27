"""Edition-3 slide layout engine: one adaptive slide per mini-section.

A topic is a dict (see T() below). Only the blocks a topic actually has are drawn, in this order:
  DEFINITION (standard definition)  ->  IN SIMPLE TERMS / ANALOGY (side by side when both)  ->
  body blocks (any of: diagram, example table from our data, comparison grid, formulas, key points, steps)
  ->  WATCH OUT strip.
Heights are measured with the real font metrics (slidekit); anything that would pass the safe area is
reported as OVERFLOW so the copy can be tightened.
"""
from __future__ import annotations

import slidekit as K
from slidekit import AMBER, BORDER, GREEN, GREY, INK, NAVY, ORANGE, RED, WHITE, measure, text_width_pt, wrap

TEAL = "13807A"
X0, W = 0.70, 11.93
TOP, BOTTOM = 1.60, 6.86
GAP = 0.16


def T(tag, title, **kw):
    kw.update(tag=tag, title=title)
    return kw


# ---------------------------------------------------------------- table geometry (mirrors Canvas.table)
def col_widths(rows, w, size, widths=None):
    ncol = max(len(r) for r in rows)
    if widths is None:
        weight = []
        for c in range(ncol):
            lens = [text_width_pt(str(r[c]).replace("**", ""), size, i == 0) for i, r in enumerate(rows) if c < len(r)]
            weight.append(max(55.0, min(max(lens), 260.0)) + 0.35 * sum(lens) / len(lens))
        widths = weight
    tot = sum(widths)
    return [w * k / tot for k in widths]


def check_tokens(rows, w, size, widths, where):
    """Warn when a word/formula cannot fit its column (it would be clipped, not wrapped)."""
    ws = col_widths(rows, w, size, widths)
    for i, r in enumerate(rows):
        for c, cell in enumerate(r):
            for tok in str(cell).replace("**", "").split():
                if text_width_pt(tok, size, i == 0) / 72 > ws[c] - 0.16:
                    K.WARN.append(f"TOKEN {where}: '{tok}' too wide for column {c} ({ws[c]:.2f}in)")


def table_height(rows, w, size, widths=None):
    ws = col_widths(rows, w, size, widths)
    h = 0
    for i, r in enumerate(rows):
        n = max(len(wrap(str(cell), size, ws[c] - 0.16, i == 0)) for c, cell in enumerate(r))
        h += n * size * K.LINE / 72 + 0.09 + 0.02
    return h


# ---------------------------------------------------------------- blocks
class Block:
    """kind: text | table | diagram"""

    def __init__(self, kind, label, color, body=None, rows=None, size=13, note="", widths=None, draw=None, h=None,
                 numbered=False):
        self.kind, self.label, self.color = kind, label, color
        self.body, self.rows, self.size, self.note = body, rows, size, note
        self.widths, self.drawfn, self.fixed_h, self.numbered = widths, draw, h, numbered

    def paras(self):
        if self.numbered:
            return [f"**{i}.**  {x}" for i, x in enumerate(self.body, 1)]
        if isinstance(self.body, list):
            return ["**•**  " + x for x in self.body]
        return [self.body]

    def height(self, w):
        if self.kind == "text":
            h, _ = measure(self.paras(), self.size, w - 0.50, False, 0.07)
            return 0.40 + h + 0.16
        if self.kind == "table":
            h = 0.32 + table_height(self.rows, w, self.size, self.widths)
            if self.note:
                h += 0.28
            return h
        if self.kind == "code":
            return 0.32 + code_h(self.body, self.size) + (0.28 if self.note else 0)
        return self.fixed_h(w) if callable(self.fixed_h) else self.fixed_h

    def draw(self, cv, x, y, w, h):
        if self.kind == "text":
            cv.rect(x, y, w, h, WHITE, BORDER)
            cv.rect(x, y, 0.09, h, self.color)
            cv.text(x + 0.30, y + 0.12, w - 0.45, 0.26, [self.label], 11.5, self.color, bold=True)
            bh, _ = measure(self.paras(), self.size, w - 0.50, False, 0.07)
            cv.text(x + 0.30, y + 0.40, w - 0.50, bh, self.paras(), self.size, INK, gap=0.07)
        elif self.kind == "table":
            check_tokens(self.rows, w, self.size, self.widths, getattr(cv, "where", ""))
            cv.text(x, y, w, 0.26, [self.label], 11.5, self.color, bold=True)
            th = cv.table(x, y + 0.32, w, self.rows, self.size, widths=self.widths)
            if self.note:
                cv.text(x, y + 0.32 + th + 0.05, w, 0.22, [self.note], 10.5, GREY)
        elif self.kind == "code":
            cv.text(x, y, w, 0.26, [self.label], 11.5, self.color, bold=True)
            ch = code_h(self.body, self.size)
            code_box(cv, x, y + 0.32, w, ch, self.body, self.size)
            if self.note:
                cv.text(x, y + 0.32 + ch + 0.05, w, 0.22, [self.note], 10.5, GREY)
        else:
            self.drawfn(cv, x, y, w, h)


def blocks_of(t, D, k=1.0):
    """Body blocks of topic t, in order. D = data module (w1_data). k scales body font sizes."""
    out = {}
    r = lambda v: round(v * k * 2) / 2  # noqa: E731
    if t.get("diagram"):
        name, args = t["diagram"]
        fn, hfn = DIAGRAMS[name]
        lab = t.get("diagram_label")
        off = 0.32 if lab else 0.0

        def _draw(cv, x, y, w, h, a=args, lab=lab, off=off):
            if lab:
                cv.text(x, y, w, 0.26, [lab.upper()], 11.5, NAVY, bold=True)
            fn(cv, x, y + off, w, h - off, **a)
        out["diagram"] = Block("diagram", "", NAVY, draw=_draw, h=lambda w, a=args, off=off: hfn(w, **a) + off)
    if t.get("ex"):
        e = t["ex"]
        rows, total = D.excerpt(e["t"], cols=e.get("cols"), n=e.get("n", 5), rows=e.get("rows"), sort=e.get("sort"),
                                ascending=e.get("asc", False))
        if e.get("total_row"):
            rows.append(e["total_row"])
        note = e.get("note") or (f"Showing {len(rows) - 1} of {total:,} rows" if total > len(rows) - 1 else "")
        out["ex"] = Block("table", f"EXAMPLE FROM OUR DATA  ·  table {e['t']}", ORANGE, rows=rows, size=r(e.get("size", 11.5)),
                          note=note, widths=e.get("widths"))
    if t.get("grid"):
        g = t["grid"]
        out["grid"] = Block("table", g.get("label", "COMPARISON").upper(), TEAL, rows=g["rows"], size=r(g.get("size", 12)),
                            note=g.get("note", ""), widths=g.get("widths"))
    if t.get("fx"):
        f = t["fx"]
        head = t.get("fx_head") or (["Formula", "Result", "What it does"] if len(f[0]) == 3 else ["Formula", "What it does"])
        rows = [head] + [list(r) for r in f]
        out["fx"] = Block("table", t.get("fx_label", "FORMULAS"), TEAL, rows=rows, size=r(12), note=t.get("fx_note", ""),
                          widths=t.get("fx_widths"))
    if t.get("code"):
        c = t["code"]
        lines = c if isinstance(c, list) else c.strip("\n").split("\n")
        out["code"] = Block("code", t.get("code_label", "SQL"), NAVY, body=lines, size=round(t.get("code_size", 12) * min(k, 1.0) * 2) / 2,
                            note=t.get("code_note", ""))
    if t.get("points"):
        out["points"] = Block("text", t.get("points_label", "KEY POINTS"), NAVY, body=t["points"], size=r(13))
    if t.get("steps"):
        out["steps"] = Block("text", t.get("steps_label", "HOW TO DO IT IN EXCEL"), GREEN, body=t["steps"], size=r(13),
                             numbered=True)
    order = t.get("order") or ["diagram", "code", "ex", "grid", "fx", "points", "steps"]
    return [out[k] for k in order if k in out]


def card_h(w, text, size):
    h, _ = measure([text], size, w - 0.50, False, 0.06)
    return 0.40 + h + 0.14


def card(cv, x, y, w, h, color, label, text, size):
    cv.rect(x, y, w, h, WHITE, BORDER)
    cv.rect(x, y, 0.09, h, color)
    cv.text(x + 0.30, y + 0.12, w - 0.45, 0.26, [label], 11.5, color, bold=True)
    bh, _ = measure([text], size, w - 0.50, False, 0.06)
    cv.text(x + 0.30, y + 0.40, w - 0.50, bh, [text], size, INK)


class Dry:
    """Canvas stand-in used to measure a layout without drawing it."""
    ops = []

    def rect(self, *a, **k):
        return None

    def text(self, *a, **k):
        return None

    def check(self, *a, **k):
        return None

    def table(self, x, y, w, rows, size=12, widths=None, name=None):
        return table_height(rows, w, size, widths)

    def code(self, *a, **k):
        return None


# ---------------------------------------------------------------- code panel (SQL / DAX / M), monospace
CODE_BG, CODE_INK, CODE_W = "1E2A38", "F3F6FA", 0.55     # Consolas advance width = 0.55 em
CODE_PAD = 0.14


def code_h(lines, size):
    return len(lines) * size * 1.22 / 72 + 2 * CODE_PAD


def code_box(cv, x, y, w, h, lines, size):
    need = max((len(l_) for l_ in lines), default=0) * size * CODE_W / 72 + 2 * CODE_PAD
    if need > w + 0.005:
        K.WARN.append(f"CODE {getattr(cv, 'where', '')}: line needs {need:.2f}in, panel {w:.2f}in: {max(lines, key=len)[:50]}")
    cv.rect(x, y, w, h, CODE_BG, None)
    if hasattr(cv, "code"):
        cv.code(x + CODE_PAD, y + CODE_PAD, w - 2 * CODE_PAD, h - 2 * CODE_PAD, lines, size)


def _canvas_code(self, x, y, w, h, lines, size):
    from pptx.util import Emu as _E, Pt as _P
    from slidekit import EMU as _EMU, _rgb
    from pptx.oxml.ns import qn as _qn
    tb = self.slide.shapes.add_textbox(_E(int(x * _EMU)), _E(int(y * _EMU)), _E(int(w * _EMU)), _E(int(h * _EMU)))
    tb.name = "Code"
    tf = tb.text_frame
    tf.word_wrap = False
    tf.auto_size = None
    bp = tf._txBody.find(_qn("a:bodyPr"))
    for k_ in ("lIns", "rIns", "tIns", "bIns"):
        bp.set(k_, "0")
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = 1.0
        r_ = p.add_run()
        r_.text = ln if ln else " "
        r_.font.name, r_.font.size, r_.font.bold = "Consolas", _P(size), False
        r_.font.color.rgb = _rgb(CODE_INK)
    self.ops.append(("code", x, y, w, h, list(lines), size, CODE_INK))
    return tb


K.Canvas.code = _canvas_code


SCALES = (1.25, 1.18, 1.1, 1.0, 0.94, 0.88)


def layout_topic(cv, t, D, where):
    """Pick the largest body scale that fits, then draw."""
    keep = len(K.WARN)
    chosen = SCALES[-1]
    for k in SCALES:
        end = _layout(Dry(), t, D, where, k)
        if end <= BOTTOM + 0.005:
            chosen = k
            break
    del K.WARN[keep:]
    return _layout(cv, t, D, where, chosen)


def _layout(cv, t, D, where, k):
    y = TOP
    if t.get("defn"):
        h = card_h(W, t["defn"], 15)
        card(cv, X0, y, W, h, NAVY, t.get("defn_label", "DEFINITION"), t["defn"], 15)
        y += h + GAP
    pair = [(GREEN, "IN SIMPLE TERMS", t.get("simple")), (AMBER, "ANALOGY", t.get("like"))]
    pair = [p for p in pair if p[2]]
    if len(pair) == 2:
        wc = (W - 0.25) / 2
        h = max(card_h(wc, p[2], 13.5) for p in pair)
        for i, (c, lab, txt) in enumerate(pair):
            card(cv, X0 + i * (wc + 0.25), y, wc, h, c, lab, txt, 13.5)
        y += h + GAP
    elif pair:
        c, lab, txt = pair[0]
        h = card_h(W, txt, 13.5)
        card(cv, X0, y, W, h, c, lab, txt, 13.5)
        y += h + GAP
    bl = blocks_of(t, D, k)
    # diagrams (and everything when stack=True) take a full-width row; others pair up side by side
    rowsets, cur = [], []
    for b_ in bl:
        if b_.kind == "diagram" or t.get("stack"):
            if cur:
                rowsets.append(cur)
                cur = []
            rowsets.append([b_])
        else:
            cur.append(b_)
            if len(cur) == 2:
                rowsets.append(cur)
                cur = []
    if cur:
        rowsets.append(cur)
    if len(bl) > 3:
        K.WARN.append(f"BLOCKS {where}: {len(bl)} body blocks (max 3)")
    for rs in rowsets:
        if len(rs) == 1:
            h = rs[0].height(W)
            rs[0].draw(cv, X0, y, W, h)
        elif len(rs) == 2:
            split = t.get("split", 0.5)
            w1 = (W - 0.30) * split
            w2 = W - 0.30 - w1
            h = max(rs[0].height(w1), rs[1].height(w2))
            for b_, x_, w_ in ((rs[0], X0, w1), (rs[1], X0 + w1 + 0.30, w2)):
                hh = h if b_.kind == "text" else b_.height(w_)
                b_.draw(cv, x_, y, w_, hh)
        else:
            continue
        y += h + GAP
    if t.get("watch"):
        txt = f"**WATCH OUT**   {t['watch']}"
        hh, _ = measure([txt], 13, W - 0.45, False)
        h = hh + 0.22
        cv.rect(X0, y, W, h, "FBEDEC", None)
        cv.rect(X0, y, 0.09, h, RED)
        cv.text(X0 + 0.30, y + 0.11, W - 0.45, hh, [txt], 13, "7B1F28")
        y += h + GAP
    if y - GAP > BOTTOM + 0.005:
        K.WARN.append(f"OVERFLOW {where}: content ends at {y - GAP:.2f}in (limit {BOTTOM})")
    return y - GAP


# ---------------------------------------------------------------- diagrams
def _box(cv, x, y, w, h, fill, txt, size, color=WHITE, bold=True, line=None):
    cv.rect(x, y, w, h, fill, line)
    lines = sum(len(wrap(p, size, w - 0.16, bold)) for p in txt)
    th = lines * size * K.LINE / 72
    cv.text(x + 0.08, y + (h - th) / 2, w - 0.16, th + 0.02, txt, size, color, bold=bold, align="c")


def flow_h(w, steps, per_row=None, sub=None, **_):
    per = per_row or len(steps)
    rows = -(-len(steps) // per)
    return rows * (1.05 if sub else 0.78) + (rows - 1) * 0.22


def flow(cv, x, y, w, h, steps, per_row=None, highlight=None, sub=None):
    """Process boxes left->right with arrows; optional per-step sub-caption (sub list)."""
    per = per_row or len(steps)
    aw = 0.34
    bw = (w - (per - 1) * aw) / per
    bh = 0.62
    for i, s in enumerate(steps):
        r, c = divmod(i, per)
        bx = x + c * (bw + aw)
        by = y + r * ((1.05 if sub else 0.78) + 0.22)
        hl = highlight is not None and (i in highlight if isinstance(highlight, (list, tuple, set)) else i == highlight)
        _box(cv, bx, by, bw, bh, ORANGE if hl else NAVY, [s], 12)
        if sub:
            cv.text(bx, by + bh + 0.05, bw, 0.36, [sub[i]], 10.5, GREY, align="c")
        if c < per - 1 and i < len(steps) - 1:
            cv.text(bx + bw, by + bh / 2 - 0.13, aw, 0.26, ["→"], 14, GREY, bold=True, align="c")


def pyramid_h(w, levels, notes=None, **_):
    return len(levels) * 0.56 + 0.05


def pyramid(cv, x, y, w, h, levels, notes=None, highlight=None):
    """levels bottom->top; drawn top (narrow) to bottom (wide) with notes on the right."""
    n = len(levels)
    pw = w * (0.48 if notes else 1.0)
    cols = [NAVY, TEAL, GREEN, ORANGE]
    for k in range(n):
        lvl = n - 1 - k                 # top first
        bw = pw * (0.40 + 0.60 * (k + 1) / n)
        bx = x + (pw - bw) / 2
        by = y + k * 0.56
        _box(cv, bx, by, bw, 0.48, cols[lvl % 4] if highlight is None or lvl == highlight else "9AA5B1", [levels[lvl]], 12.5)
        if notes:
            cv.text(x + pw + 0.25, by + 0.10, w - pw - 0.25, 0.40, [notes[lvl]], 12, INK)


def excel_h(w, **_):
    return 3.25


def excel(cv, x, y, w, h, highlight=(), labels=None, active="C4"):
    """Schematic Excel window. highlight: parts to outline in orange: title, ribbon, qat, namebox, formulabar,
    colhdr, rowhdr, grid, active, tabs, status, views, file."""
    hl = set(highlight)
    labels = labels or {}
    o = ORANGE
    lg = "F3F2EF"
    ww = w if not labels else w * 0.66
    # title bar
    cv.rect(x, y, ww, 0.30, "1D6F42" if "title" not in hl else o)
    cv.text(x + 0.10, y + 0.06, ww - 0.2, 0.2, ["AutoSave  ·  ETEH_Week1_Section_Examples.xlsx  ·  Excel"], 10, WHITE, align="c")
    if "qat" in hl:
        cv.rect(x + 0.05, y + 0.03, 1.05, 0.24, None, o)
    # tabs row
    tabs = ["File", "Home", "Insert", "Page Layout", "Formulas", "Data", "Review", "View"]
    ty = y + 0.30
    cv.rect(x, ty, ww, 0.26, "FFFFFF", BORDER)
    tx = x + 0.08
    for i, tname in enumerate(tabs):
        tw = text_width_pt(tname, 10) / 72 + 0.18
        if tx + tw > x + ww:
            break
        col = o if (tname == "File" and "file" in hl) else ("1D6F42" if tname == "Home" else INK)
        cv.text(tx, ty + 0.05, tw, 0.18, [tname], 10, col, bold=(tname == "Home"))
        tx += tw
    # ribbon groups
    ry = ty + 0.26
    cv.rect(x, ry, ww, 0.50, lg, o if "ribbon" in hl else BORDER)
    groups = ["Clipboard", "Font", "Alignment", "Number", "Styles", "Cells", "Editing"]
    gw = (ww - 0.1) / len(groups)
    for i, g in enumerate(groups):
        cv.rect(x + 0.05 + i * gw, ry + 0.06, gw - 0.06, 0.28, WHITE, BORDER)
        cv.text(x + 0.05 + i * gw, ry + 0.35, gw - 0.06, 0.14, [g], 8.5, GREY, align="c")
    # name box + formula bar
    fy = ry + 0.56
    cv.rect(x + 0.05, fy, 0.75, 0.24, WHITE, o if "namebox" in hl else BORDER)
    cv.text(x + 0.10, fy + 0.04, 0.65, 0.18, [active], 10, INK)
    cv.text(x + 0.85, fy + 0.03, 0.30, 0.2, ["fx"], 10, GREY, bold=True)
    cv.rect(x + 1.15, fy, ww - 1.20, 0.24, WHITE, o if "formulabar" in hl else BORDER)
    cv.text(x + 1.22, fy + 0.04, ww - 1.4, 0.18, ["=SUM(tStores25[Sales])"], 10, INK)
    # grid
    gy = fy + 0.32
    ncol, nrow = 8, 5
    cw = (ww - 0.40) / ncol
    rh = 0.22
    cv.rect(x, gy, ww, rh, lg, o if "colhdr" in hl else BORDER)
    for c in range(ncol):
        cv.text(x + 0.40 + c * cw, gy + 0.03, cw, 0.16, [chr(65 + c)], 9.5, GREY, align="c")
    cv.rect(x, gy + rh, 0.40, nrow * rh, lg, o if "rowhdr" in hl else BORDER)
    cv.rect(x + 0.40, gy + rh, ww - 0.40, nrow * rh, WHITE, o if "grid" in hl else BORDER)
    for r in range(nrow):
        cv.text(x, gy + rh + r * rh + 0.03, 0.40, 0.16, [str(r + 1)], 9.5, GREY, align="c")
    ac, ar = ord(active[0]) - 65, int(active[1:]) - 1
    cv.rect(x + 0.40 + ac * cw, gy + rh + ar * rh, cw, rh, None, o if "active" in hl else "1D6F42")
    # sheet tabs + status bar
    sy = gy + rh + nrow * rh
    cv.rect(x, sy, ww, 0.24, lg, o if "tabs" in hl else BORDER)
    for i, sname in enumerate(["00_Cover", "1.1 Intro", "1.2 Types", "+"]):
        cv.text(x + 0.10 + i * 0.95, sy + 0.04, 0.9, 0.18, [sname], 9.5, "1D6F42" if i == 1 else GREY, bold=(i == 1))
    cv.rect(x, sy + 0.24, ww, 0.22, "1D6F42" if "status" not in hl else o)
    cv.text(x + 0.10, sy + 0.27, ww * 0.6, 0.16, ["Ready     Average: 61,101,271   Count: 12   Sum: 733,215,255"], 9, WHITE)
    cv.text(x + ww - 1.6, sy + 0.27, 1.5, 0.16, ["Normal  Layout  Break   100%"], 9, WHITE, align="r")
    if labels:
        lx = x + ww + 0.25
        ly = y + 0.05
        for k, v in labels.items():
            hh, _ = measure([f"**{k}**  {v}"], 12, w - ww - 0.25)
            cv.text(lx, ly, w - ww - 0.25, hh, [f"**{k}**  {v}"], 12, INK)
            ly += hh + 0.08


def ladder_h(w, steps, **_):
    return 1.45


def ladder(cv, x, y, w, h, steps, highlight=None):
    """Rising staircase of analytics types: steps = [(name, question)]."""
    n = len(steps)
    sw = (w - (n - 1) * 0.15) / n
    cols = [NAVY, TEAL, GREEN, ORANGE]
    for i, (name, q) in enumerate(steps):
        bh = 0.75 + i * 0.23
        bx = x + i * (sw + 0.15)
        by = y + (1.45 - bh)
        fill = cols[i % 4] if highlight is None or i == highlight else "9AA5B1"
        _box(cv, bx, by, sw, bh, fill, [name, q], 12)


DIAGRAMS = {"flow": (flow, flow_h), "pyramid": (pyramid, pyramid_h), "excel": (excel, excel_h), "ladder": (ladder, ladder_h)}


def pqeditor_h(w, **_):
    return 3.30


def pqeditor(cv, x, y, w, h, highlight=(), labels=None, query="qSalesClean",
             cols=("Order_ID", "Order_Date", "Store_Name", "Sales"), types=("ABC", "Date", "ABC", "1.2"),
             steps=("Source", "Promoted Headers", "Changed Type", "Trimmed Text", "Removed Duplicates"),
             formula="= Table.Distinct(#\"Trimmed Text\")"):
    """Schematic Power Query Editor. highlight: ribbon, queries, formulabar, preview, settings, steps, status."""
    hl = set(highlight)
    labels = labels or {}
    o = ORANGE
    lg = "F3F2EF"
    ww = w if not labels else w * 0.64
    cv.rect(x, y, ww, 0.28, "1D6F42")
    cv.text(x + 0.10, y + 0.05, ww - 0.2, 0.2, [f"{query} - Power Query Editor"], 10, WHITE, align="c")
    ty = y + 0.28
    cv.rect(x, ty, ww, 0.24, WHITE, BORDER)
    tx = x + 0.08
    for tname in ["File", "Home", "Transform", "Add Column", "View", "Tools", "Help"]:
        tw = text_width_pt(tname, 9.5) / 72 + 0.16
        if tx + tw > x + ww:
            break
        cv.text(tx, ty + 0.04, tw, 0.18, [tname], 9.5, "1D6F42" if tname == "Home" else INK, bold=(tname == "Home"))
        tx += tw
    ry = ty + 0.24
    cv.rect(x, ry, ww, 0.42, lg, o if "ribbon" in hl else BORDER)
    groups = ["Close & Load", "Refresh", "Manage Columns", "Reduce Rows", "Sort", "Transform", "Combine"]
    gw = (ww - 0.1) / len(groups)
    for i, g in enumerate(groups):
        cv.rect(x + 0.05 + i * gw, ry + 0.05, gw - 0.06, 0.22, WHITE, BORDER)
        cv.text(x + 0.05 + i * gw, ry + 0.28, gw - 0.06, 0.13, [g], 8, GREY, align="c")
    by = ry + 0.48
    bh = h - (by - y) - 0.26
    qw, sw = ww * 0.20, ww * 0.25
    # queries pane
    cv.rect(x, by, qw, bh, WHITE, o if "queries" in hl else BORDER)
    cv.text(x + 0.06, by + 0.05, qw - 0.1, 0.18, ["Queries [3]"], 9, NAVY, bold=True)
    for i, q in enumerate([query, "fxPath", "qZoneSummary"]):
        cv.text(x + 0.10, by + 0.30 + i * 0.22, qw - 0.12, 0.18, [q], 9, INK, bold=(i == 0))
    # formula bar + preview
    px, pw = x + qw + 0.06, ww - qw - sw - 0.12
    cv.rect(px, by, pw, 0.24, WHITE, o if "formulabar" in hl else BORDER)
    cv.text(px + 0.06, by + 0.04, pw - 0.1, 0.18, [formula], 9, INK)
    gy = by + 0.30
    gh = bh - 0.30
    cv.rect(px, gy, pw, gh, WHITE, o if "preview" in hl else BORDER)
    cw = pw / len(cols)
    for i, (c, t_) in enumerate(zip(cols, types)):
        cv.rect(px + i * cw, gy, cw, 0.36, lg, BORDER)
        cv.text(px + i * cw + 0.04, gy + 0.02, cw - 0.06, 0.16, [t_ + "  " + c], 8.5, INK, bold=True)
        cv.rect(px + i * cw + 0.04, gy + 0.22, cw - 0.10, 0.07, "2E7D32")
    for r in range(4):
        cv.rect(px, gy + 0.36 + r * 0.22, pw, 0.22, None, BORDER)
    # settings pane
    sx = px + pw + 0.06
    cv.rect(sx, by, sw, bh, WHITE, o if "settings" in hl else BORDER)
    cv.text(sx + 0.06, by + 0.05, sw - 0.1, 0.18, ["Query Settings"], 9, NAVY, bold=True)
    cv.text(sx + 0.06, by + 0.27, sw - 0.1, 0.16, ["PROPERTIES  ·  Name: " + query], 8, GREY)
    cv.text(sx + 0.06, by + 0.50, sw - 0.1, 0.16, ["APPLIED STEPS"], 8.5, NAVY, bold=True)
    if "steps" in hl:
        cv.rect(sx + 0.03, by + 0.47, sw - 0.06, 0.24 + 0.19 * len(steps), None, o)
    for i, st in enumerate(steps):
        cv.text(sx + 0.12, by + 0.70 + i * 0.19, sw - 0.15, 0.16, [st], 8.5, INK, bold=(i == len(steps) - 1))
    # status bar
    sy = by + bh + 0.04
    cv.rect(x, sy, ww, 0.20, lg, o if "status" in hl else BORDER)
    cv.text(x + 0.08, sy + 0.03, ww - 0.2, 0.15, [f"{len(cols)} COLUMNS, 999+ ROWS     Column profiling based on top 1000 rows"], 8.5, GREY)
    if labels:
        lx = x + ww + 0.25
        ly = y + 0.05
        for k, v in labels.items():
            hh, _ = measure([f"**{k}**  {v}"], 12, w - ww - 0.25)
            cv.text(lx, ly, w - ww - 0.25, hh, [f"**{k}**  {v}"], 12, INK)
            ly += hh + 0.08


def pivot_h(w, **_):
    return 2.75


def pivot(cv, x, y, w, h, filters=(), columns=(), rows=(), values=(), table=None, size=11):
    """PivotTable Fields: the four areas (left) and, optionally, the resulting report (right)."""
    lw = w * (0.42 if table else 1.0)
    bw = (lw - 0.15) / 2
    bh = (h - 0.15) / 2
    areas = [("FILTERS", filters, NAVY), ("COLUMNS", columns, TEAL), ("ROWS", rows, GREEN), ("VALUES", values, ORANGE)]
    for i, (lab, items, col) in enumerate(areas):
        r_, c_ = divmod(i, 2)
        bx, by = x + c_ * (bw + 0.15), y + r_ * (bh + 0.15)
        cv.rect(bx, by, bw, bh, WHITE, BORDER)
        cv.rect(bx, by, bw, 0.30, col)
        cv.text(bx + 0.08, by + 0.05, bw - 0.1, 0.2, [lab], 11, WHITE, bold=True)
        for k, it in enumerate(items):
            cv.rect(bx + 0.10, by + 0.40 + k * 0.34, bw - 0.20, 0.28, "F3F2EF", BORDER)
            cv.text(bx + 0.16, by + 0.44 + k * 0.34, bw - 0.3, 0.2, [it], 10.5, INK)
    if table:
        tx = x + lw + 0.30
        cv.table(tx, y, w - lw - 0.30, table, size)


DIAGRAMS.update({"pqeditor": (pqeditor, pqeditor_h), "pivot": (pivot, pivot_h)})
