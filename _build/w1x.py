"""Edition-3 workbook blocks for the Week 1 examples workbook (built on xlhelp).

A module sheet is a tutor walk-through: the module's dataset first, then every mini-section as
    band (tag + title)  ->  what it is  ->  Steps  ->  Worked solution (live formulas)  ->  Your turn
and an Answer key at the bottom. 'Your turn' rows have a yellow answer cell and a live check cell
that compares the learner's entry with the worked answer in the key.

Data sheets hold one Excel table each, starting in cell A1 (one header row, nothing else on the sheet),
so cell addresses quoted on the slides (e.g. tStores25 row 2 = Lagos - Ikeja, column E = Sales) are true.
"""
from __future__ import annotations

import math
import re

import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell

from xlhelp import LAST_COL, Sheet

COVERED: list[str] = []


def expand(tag: str) -> list[str]:
    """'1.6.12-13' -> ['1.6.12', '1.6.13'];  '1.1.4' -> ['1.1.4']"""
    m = re.fullmatch(r"(\d+\.\d+)\.(\d+)(?:-(\d+))?", tag)
    if not m:
        return []
    a, b = int(m.group(2)), int(m.group(3) or m.group(2))
    return [f"{m.group(1)}.{k}" for k in range(a, b + 1)]


def _lines(text, chars):
    return sum(max(1, math.ceil(len(p) / chars)) for p in str(text).split("\n"))


def _py(v):
    if isinstance(v, (np.floating, np.integer)):
        return v.item()
    return v


class DataSheet(Sheet):
    """A sheet that holds exactly one Excel table starting in A1."""

    def __init__(self, book, name, title, topics):   # noqa: D107  (no Sheet.__init__: nothing above the header row)
        self.b = book
        self.F = book.F
        self.ws = book.wb.add_worksheet(name)
        self.name = name
        self.r = 0
        book.toc.append((name, title, topics))

    def put(self, df, tname, fmts=None, calc=None, widths=None):
        t = self.table(df, tname, fmts=fmts, c0=0, style="Table Style Light 9", calc=calc, widths=widths)
        self.ws.freeze_panes(1, 0)
        self.ws.repeat_rows(0)
        self.ws.set_landscape()
        self.ws.fit_to_pages(1, 0)
        return t


class ModuleSheet:
    def __init__(self, book, name, title, subtitle, topics):
        self.b = book
        self.s = book.sheet(name, title, subtitle, topics)
        self.ws = self.s.ws
        self.F = book.F
        self.tag = ""
        self.n = 0
        self.checks = []        # (row, col, kind, payload)
        self.key_num = []       # (n, tag, short, formula, cached, fmt)
        self.key_txt = []       # (n, tag, answer text)
        f = book.wb.add_format
        if not hasattr(book, "_w1fmt"):
            base = {"font_name": "Calibri", "font_size": 10, "valign": "top"}
            book._w1fmt = {
                "num": f({**base, "bold": True, "font_color": "white", "bg_color": "#13807A", "align": "center"}),
                "step": f({**base, "text_wrap": True, "border": 1, "border_color": "#D5DEE6"}),
                "exid": f({**base, "bold": True, "font_color": "#0B2545", "bg_color": "#E0A526", "align": "center"}),
                "task": f({**base, "text_wrap": True, "bg_color": "#FFF4C2", "border": 1, "border_color": "#E0A526"}),
                "in": f({**base, "bg_color": "#FFFFFF", "border": 2, "border_color": "#E0A526", "text_wrap": True}),
                "chk": f({**base, "bold": True, "font_color": "#1B5E20", "text_wrap": True}),
                "see": f({**base, "italic": True, "text_wrap": True, "font_color": "#555555", "border": 1,
                          "border_color": "#D5DEE6"}),
                "ds": f({**base, "text_wrap": True, "bg_color": "#F2F2F2", "border": 1, "border_color": "#BBBBBB"}),
            }
        self.X = book._w1fmt

    # ------------------------------------------------------------ passthrough
    def __getattr__(self, k):
        return getattr(self.s, k)

    @property
    def r(self):
        return self.s.r

    @r.setter
    def r(self, v):
        self.s.r = v

    # ------------------------------------------------------------ structure
    def dataset(self, items):
        """items: (table or range name, where it is, what one row is / why we use it)."""
        self.s.band("DATA", "This module's dataset")
        self.s.grid([list(i) for i in items], header=["Table / range", "Where", "What it holds and how we use it"],
                    merge_last=True)
        self.s.r += 1

    def topic(self, tag, title, what=None):
        COVERED.extend(expand(tag))
        self.tag = tag
        self.s.band(tag, title)
        if what:
            self.s.para(what, "text")

    def h(self, text):
        self.s.h2(text)

    def steps(self, items, label="Steps"):
        self.s.h2(label)
        for i, t in enumerate(items, 1):
            self.ws.write(self.r, 1, i, self.X["num"])
            self.ws.merge_range(self.r, 2, self.r, LAST_COL, t, self.X["step"])
            self.ws.set_row(self.r, 14 * _lines(t, 140) + 2)
            self.r += 1
        self.r += 1

    def worked(self, rows, label="Worked solution", **kw):
        self.s.h2(label)
        return self.s.calc(rows, **kw)

    def note(self, text):
        self.s.para(text, "note")

    # ------------------------------------------------------------ exercises
    def turn(self, label="Your turn"):
        self.s.h2(label + "  (type in the white cell with the gold border; the check updates as you type)")

    def _task(self, task, short=None):
        self.n += 1
        r = self.r
        self.ws.write(r, 1, f"Ex {self.n}", self.X["exid"])
        self.ws.merge_range(r, 2, r, 7, task, self.X["task"])
        self.ws.set_row(r, 14 * max(_lines(task, 80), 1) + 4)
        return r

    def ex(self, task, formula, cached, fmt="gen", tol=None, short=None):
        """Numeric answer worked out by `formula` in the key; the learner types a number."""
        r = self._task(task)
        self.ws.write_blank(r, 8, None, self.b.nf(fmt, style="plain", border=2, border_color="#E0A526"))
        self.ws.merge_range(r, 9, r, LAST_COL, "", self.X["chk"])
        if tol is None:
            tol = 0.0006 if fmt.startswith("pct") else (0.06 if fmt in ("dec1", "dec2") else 0.5)
        self.checks.append((r, "num", tol))
        self.key_num.append((self.n, self.tag, short or task, formula, cached, fmt))
        self.r += 1

    def ex_text(self, task, answer, formula=None, short=None, alts=()):
        """Text answer. With `formula`, the key works it out live; otherwise the answer is a fixed word."""
        r = self._task(task)
        self.ws.write_blank(r, 8, None, self.X["in"])
        self.ws.merge_range(r, 9, r, LAST_COL, "", self.X["chk"])
        if formula:
            self.checks.append((r, "tcell", alts))
            self.key_num.append((self.n, self.tag, short or task, formula, answer, "gen"))
        else:
            self.checks.append((r, "tlit", (answer,) + tuple(alts)))
            self.key_txt.append((self.n, self.tag, answer))
        self.r += 1

    def ex_open(self, task, model):
        """Written answer: the key shows a model answer."""
        r = self._task(task)
        self.ws.merge_range(r, 8, r, LAST_COL, "", self.X["in"])
        self.ws.set_row(r, max(14 * _lines(task, 80) + 4, 44))
        self.key_txt.append((self.n, self.tag, "Model answer: " + model))
        self.r += 1

    def ex_formula_live(self, task, check_formula, cached_msg, short=None, key=None):
        """The learner acts in the workbook (or writes a formula); `check_formula` inspects the result itself."""
        r = self._task(task)
        self.ws.write_blank(r, 8, None, self.X["in"])
        self.ws.merge_range(r, 9, r, LAST_COL, "", self.X["chk"])
        self.ws.write_formula(r, 9, check_formula, self.X["chk"], cached_msg)
        if key:
            self.key_txt.append((self.n, self.tag, key))
        self.r += 1

    def do(self, items, label="Your turn: do it in Excel"):
        """Procedural checklist: (instruction, what you should see)."""
        self.s.h2(label)
        hdr = self.F["hdr"]
        self.ws.write(self.r, 1, "Task", hdr)
        self.ws.merge_range(self.r, 2, self.r, 7, "Do this", hdr)
        self.ws.write(self.r, 8, "Done?", hdr)
        self.ws.merge_range(self.r, 9, self.r, LAST_COL, "What you should see", hdr)
        self.r += 1
        for ins, see in items:
            self.n += 1
            self.ws.write(self.r, 1, f"Ex {self.n}", self.X["exid"])
            self.ws.merge_range(self.r, 2, self.r, 7, ins, self.X["task"])
            self.ws.write_blank(self.r, 8, None, self.X["in"])
            self.ws.data_validation(self.r, 8, self.r, 8, {"validate": "list", "source": ["Done", "Not yet"]})
            self.ws.merge_range(self.r, 9, self.r, LAST_COL, see, self.X["see"])
            self.ws.set_row(self.r, 14 * max(_lines(ins, 80), _lines(see, 58)) + 4)
            self.r += 1
        self.r += 1

    # ------------------------------------------------------------ answer key
    def finish(self):
        s = self.s
        s.band("KEY", "Answer key: the worked answer to every 'Your turn' exercise on this sheet")
        s.para("Each answer below is a live formula on the module's data, so you can click it and read how it was worked out. "
               "The check cells beside the exercises compare your entry with these answers.", "note")
        cells = {}
        if self.key_num:
            rows = [(f"Ex {n}  ·  {tag}  ·  {short}", fx, cached, fmt, "") for n, tag, short, fx, cached, fmt in self.key_num]
            out = s.calc(rows, header=("Exercise", "Answer", "Formula (live text)", "Notes"))
            for (n, *_), (label, *_r) in zip(self.key_num, rows):
                cells[n] = out[label]
        if self.key_txt:
            s.grid([[f"Ex {n}", tag, txt] for n, tag, txt in self.key_txt], header=["Exercise", "Topic", "Answer"],
                   merge_last=True)
        # write the live checks now that the key cells are known
        num_iter = iter([n for n, *_ in self.key_num])
        for r, kind, payload in self.checks:
            inp = xl_rowcol_to_cell(r, 8)
            if kind == "num":
                key = cells[next(num_iter)]
                ka = re.sub(r"([A-Z]+)(\d+)", r"$\1$\2", key)
                fx = (f'=IF({inp}="","",IF(ISNUMBER({inp}),IF(ABS({inp}-{ka})<={payload},"✓ Correct",'
                      f'"✗ Not yet: check the steps"),"✗ Type a number"))')
            elif kind == "tcell":
                key = cells[next(num_iter)]
                ka = re.sub(r"([A-Z]+)(\d+)", r"$\1$\2", key)
                alts = "".join(f',LOWER(TRIM({inp}))="{a.lower()}"' for a in payload)
                fx = f'=IF({inp}="","",IF(OR(LOWER(TRIM({inp}))=LOWER({ka}){alts}),"✓ Correct","✗ Not yet"))'
            else:
                conds = ",".join(f'LOWER(TRIM({inp}))="{a.lower()}"' for a in payload)
                fx = f'=IF({inp}="","",IF(OR({conds}),"✓ Correct","✗ Not yet"))'
            self.ws.write_formula(r, 9, fx, self.X["chk"], "")
        return cells
