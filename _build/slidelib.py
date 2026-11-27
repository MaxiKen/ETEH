"""Small python-pptx helpers used by 05_update_slides.py.

Text replacement keeps the formatting of the first run in a text frame (font, size, colour) and
supports **bold** spans and \\n paragraph breaks. Every replacement is checked against an estimate
of the text box capacity so new copy never spills out of the fixed-size boxes of the template.
"""
import copy
import math
import re

from pptx.oxml.ns import qn

EMU = 914400
WARN = []


def shapes_by_name(slide):
    d = {}
    for sh in slide.shapes:
        d.setdefault(sh.name, sh)
    return d


def _first_rpr(tf):
    for p in tf.paragraphs:
        for r in p.runs:
            rpr = r._r.find(qn("a:rPr"))
            return copy.deepcopy(rpr) if rpr is not None else None
    return None


def _font_size(tf, default=14.0):
    for p in tf.paragraphs:
        for r in p.runs:
            if r.font.size:
                return r.font.size.pt
    return default


def capacity(shape, size, wfactor=0.52):
    w = shape.width / EMU - 0.2
    h = shape.height / EMU
    lines = max(1, round(h / (size * 1.2 / 72)))
    cpl = max(10, int(w * 72 / (size * wfactor)))
    return lines, cpl


def needed_lines(text, cpl):
    n = 0
    for line in text.split("\n"):
        words, cur, k = line.split(" "), 0, 1
        for wd in words:
            add = len(wd) + (1 if cur else 0)
            if cur + add > cpl:
                k += 1
                cur = len(wd)
            else:
                cur += add
        n += k
    return n


def set_text(tf, text, where="", shape=None, check=True):
    """Replace all text in a text frame, keeping the first run's formatting."""
    rpr = _first_rpr(tf)
    size = _font_size(tf)
    plain = text.replace("**", "")
    if check and shape is not None:
        lines, cpl = capacity(shape, size)
        need = needed_lines(plain, cpl)
        if need > lines:
            WARN.append(f"OVERFLOW {where}: needs {need} lines, box holds {lines} ({cpl} cpl, {size}pt): {plain[:70]}...")
    paras = tf.paragraphs
    p0 = paras[0]._p
    for p in paras[1:]:
        p._p.getparent().remove(p._p)
    template = copy.deepcopy(p0)
    for child in list(template):
        if child.tag in (qn("a:r"), qn("a:br"), qn("a:fld")):
            template.remove(child)
    for child in list(p0):
        if child.tag in (qn("a:r"), qn("a:br"), qn("a:fld")):
            p0.remove(child)
    lines = text.split("\n")
    targets = [p0]
    for _ in lines[1:]:
        np_ = copy.deepcopy(template)
        targets[-1].addnext(np_)
        targets.append(np_)
    for p, line in zip(targets, lines):
        end = p.find(qn("a:endParaRPr"))
        for i, part in enumerate(line.split("**")):
            if not part:
                continue
            r = p.makeelement(qn("a:r"), {})
            if rpr is not None:
                rp = copy.deepcopy(rpr)
                if i % 2 == 1:
                    rp.set("b", "1")
                r.append(rp)
            t = r.makeelement(qn("a:t"), {})
            t.text = part
            r.append(t)
            if end is not None:
                end.addprevious(r)
            else:
                p.append(r)


def put(slide, name, text, where=""):
    sh = shapes_by_name(slide).get(name)
    if sh is None:
        raise KeyError(f"{where}: shape {name!r} not found")
    set_text(sh.text_frame, text, f"{where}/{name}", sh)
    return sh


def get_text(slide, name):
    sh = shapes_by_name(slide).get(name)
    return sh.text_frame.text if sh is not None and sh.has_text_frame else None


def set_table(slide, name, rows, widths=None, where=""):
    """Resize a table to len(rows) x len(rows[0]) (copying existing row/cell formatting) and fill it."""
    gf = shapes_by_name(slide).get(name)
    if gf is None or not gf.has_table:
        raise KeyError(f"{where}: table {name!r} not found")
    tbl = gf.table._tbl
    nr, nc = len(rows), len(rows[0])
    # rows
    while len(tbl.tr_lst) < nr:
        trs = tbl.tr_lst
        src = trs[-2] if len(trs) >= 3 else trs[-1]
        new = copy.deepcopy(src)
        trs[-1].addnext(new)
    while len(tbl.tr_lst) > nr:
        tbl.remove(tbl.tr_lst[-1])
    # columns
    grid = tbl.tblGrid
    cols = grid.findall(qn("a:gridCol"))
    total = sum(int(c.get("w")) for c in cols)
    while len(cols) < nc:
        grid.append(copy.deepcopy(cols[-1]))
        for tr in tbl.tr_lst:
            tr.append(copy.deepcopy(tr.tc_lst[-1]))
        cols = grid.findall(qn("a:gridCol"))
    while len(cols) > nc:
        grid.remove(cols[-1])
        for tr in tbl.tr_lst:
            tr.remove(tr.tc_lst[-1])
        cols = grid.findall(qn("a:gridCol"))
    if widths is None:
        if nc == len(cols) and nc == len(rows[0]) and len(set(int(c.get("w")) for c in cols)) > 1 and nc <= 3:
            widths = [int(c.get("w")) for c in cols]
        else:
            widths = [1] * nc
    s = sum(widths)
    acc = 0
    for i, c in enumerate(cols):
        w = total - acc if i == nc - 1 else int(total * widths[i] / s)
        c.set("w", str(w))
        acc += w
    t = gf.table
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = t.cell(r, c)
            set_text(cell.text_frame, str(val), f"{where}/{name}[{r},{c}]", check=False)
            colw = int(cols[c].get("w")) / EMU - 0.15
            size = _font_size(cell.text_frame, 11.5)
            cpl = int(colw * 72 / (size * 0.52))
            if needed_lines(str(val), cpl) > 2:
                WARN.append(f"TABLE WRAP {where}/{name}[{r},{c}] > 2 lines: {val}")
    gf.height = sum(int(tr.get("h")) for tr in tbl.tr_lst)


def delete_shape(sh):
    sh._element.getparent().remove(sh._element)


def remove_slides(prs, first, last):
    """Remove slides first..last (1-based, inclusive)."""
    sldIdLst = prs.slides._sldIdLst
    ids = list(sldIdLst)
    for sid in ids[first - 1:last]:
        prs.part.drop_rel(sid.rId)
        sldIdLst.remove(sid)


def replace_in_runs(tf, func):
    """Apply func(str)->str to every run (and table cell run) of a text frame. Returns number of changes."""
    n = 0
    for p in tf.paragraphs:
        for r in p.runs:
            new = func(r.text)
            if new != r.text:
                r.text = new
                n += 1
    return n


def iter_text_frames(slide):
    for sh in slide.shapes:
        if sh.has_text_frame:
            yield sh, sh.text_frame
        if getattr(sh, "has_table", False) and sh.has_table:
            for row in sh.table.rows:
                for cell in row.cells:
                    yield sh, cell.text_frame


def para_replace(tf, old, new):
    """Replace `old` with `new` in a text frame, working at paragraph level when the phrase spans runs.
    Formatting of the run where the match starts is kept. Returns number of replacements."""
    n = 0
    for p in tf.paragraphs:
        runs = p.runs
        if not runs:
            continue
        full = "".join(r.text for r in runs)
        if old not in full:
            continue
        # fast path: fully inside a run
        done = False
        for r in runs:
            if old in r.text:
                c = r.text.count(old)
                r.text = r.text.replace(old, new)
                n += c
                done = True
        full2 = "".join(r.text for r in runs)
        if old not in full2:
            continue
        # spanning runs: rebuild using char ownership
        while old in full2:
            i = full2.index(old)
            j = i + len(old)
            pos = 0
            first = True
            for r in runs:
                a, b = pos, pos + len(r.text)
                pos = b
                if b <= i or a >= j:
                    continue
                s0, s1 = max(a, i) - a, min(b, j) - a
                if first:
                    r.text = r.text[:s0] + new + r.text[s1:]
                    first = False
                else:
                    r.text = r.text[:s0] + r.text[s1:]
            n += 1
            full2 = "".join(r.text for r in runs)
    return n
