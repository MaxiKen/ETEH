"""Shared toolkit for the weekly assignment workbooks.

Every assignment is produced twice from ONE builder function:
  * student version  -> 02_Assignments/ETEH_WeekN_Assignment.xlsx  (yellow answer cells left blank)
  * answer key       -> 03_Instructor_Projects/Assignment_Answer_Keys/ETEH_WeekN_Assignment_KEY.xlsx
                        (same layout; every answer is a live formula with its cached value + a marking guide)
Because both files come from the same code, question numbers, cell positions and marks can never drift apart.
"""
from __future__ import annotations

import math

import numpy as np
import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell

from common import ASG, IPR
from xlhelp import Book, LAST_COL

KEYDIR = IPR / "Assignment_Answer_Keys"
KEYDIR.mkdir(parents=True, exist_ok=True)


def _py(v):
    if isinstance(v, (pd.Timestamp, np.datetime64)) or (hasattr(v, "isoformat") and not isinstance(v, str)):
        if pd.isna(v):
            return ""
        return (pd.Timestamp(v) - pd.Timestamp("1899-12-30")) / pd.Timedelta(days=1)
    if isinstance(v, (np.floating, np.integer)):
        v = v.item()
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return ""
    return v


class Asg:
    def __init__(self, week: int, title: str, key: bool, subtitle: str):
        self.key = key
        self.week = week
        fn = f"ETEH_Week{week}_Assignment" + ("_ANSWER_KEY" if key else "") + ".xlsx"
        self.path = (KEYDIR if key else ASG) / fn
        label = f"Week {week} Assignment" + ("  ·  ANSWER KEY (instructor only)" if key else "")
        self.b = Book(str(self.path), title, label)
        self.subtitle = subtitle
        self.marks: list[tuple[str, str, str, int]] = []   # sheet, task, modules, marks
        self._qcount = {}
        F = self.b.F
        wb = self.b.wb
        self.qtag = wb.add_format({"font_name": "Calibri", "font_size": 10, "bold": True, "font_color": "#0B2545",
                                   "bg_color": "#E0A526", "align": "center", "valign": "top", "border": 1})
        self.qtext = wb.add_format({"font_name": "Calibri", "font_size": 10, "text_wrap": True, "valign": "top", "border": 1,
                                    "border_color": "#BBBBBB"})
        self.guide = wb.add_format({"font_name": "Calibri", "font_size": 9, "text_wrap": True, "valign": "top", "border": 1,
                                    "border_color": "#2E7D32", "font_color": "#1B5E20", "bg_color": "#F4FBF3"})
        self.work = wb.add_format({"font_name": "Calibri", "font_size": 9, "text_wrap": True, "valign": "top", "border": 1,
                                   "border_color": "#BBBBBB", "italic": True, "font_color": "#777777"})
        self.longans = wb.add_format({"font_name": "Calibri", "font_size": 10, "text_wrap": True, "valign": "top", "border": 1,
                                      "border_color": "#E0A526", "bg_color": "#FFF4C2"})
        self.longkey = wb.add_format({"font_name": "Calibri", "font_size": 10, "text_wrap": True, "valign": "top", "border": 1,
                                      "border_color": "#2E7D32", "bg_color": "#E3F4E1"})
        self.F = F

    # ------------------------------------------------------------------ sheets
    def task(self, name, title, modules, brief, marks_note=""):
        s = self.b.sheet(name, title, self.subtitle, modules)
        s.ws.set_column(8, 8, 17)     # answer column I
        s.ws.set_column(9, 9, 7)      # marks column J
        s.scenario(brief)
        s._task_name = title
        s._modules = modules
        s._marks = 0
        self._qcount[name] = 0
        if marks_note:
            s.para(marks_note, "note")
        return s

    def end_task(self, s):
        self.marks.append((s.name, s._task_name, s._modules, s._marks))

    # ------------------------------------------------------------------ question rows
    def qhead(self, s):
        ws = s.ws
        ws.write(s.r, 1, "Q", self.F["hdr"])
        ws.merge_range(s.r, 2, s.r, 7, "Question", self.F["hdr"])
        ws.write(s.r, 8, "Your answer" if not self.key else "Answer (live)", self.F["hdr"])
        ws.write(s.r, 9, "Marks", self.F["hdr"])
        ws.merge_range(s.r, 10, s.r, LAST_COL, "Working / notes" if not self.key else "Marking guide", self.F["hdr"])
        s.r += 1

    def q(self, s, qid, text, answer, cached=None, kind="gen", marks=2, guide="", choices=None, header=False):
        """One question row. answer: '=formula' | literal value | text. cached: value stored with a formula."""
        if header:
            self.qhead(s)
        ws = s.ws
        ws.write(s.r, 1, qid, self.qtag)
        ws.merge_range(s.r, 2, s.r, 7, text, self.qtext)
        if self.key:
            fmt = s.b.nf(kind or "gen", style="result", text_wrap=True)
            if isinstance(answer, str) and answer.startswith("="):
                ws.write_formula(s.r, 8, answer, fmt, _py(cached) if cached is not None else "")
                g = f"{answer}   ·   {guide}" if guide else answer
            else:
                s._w(s.r, 8, answer, fmt)
                g = guide
            s.mtext(s.r, 10, LAST_COL, g, self.guide)
        else:
            ws.write_blank(s.r, 8, None, s.b.nf(kind or "gen", style="input", text_wrap=True))
            ws.merge_range(s.r, 10, s.r, LAST_COL, "", self.work)
        if choices:
            ws.data_validation(s.r, 8, s.r, 8, {"validate": "list", "source": choices, "error_type": "warning"})
        ws.write_number(s.r, 9, marks, s.b.nf("int", style="cell", align="center"))
        guide_len = len(str(answer)) + len(guide) if self.key else 0
        lines = max(math.ceil(len(text) / 80), math.ceil(guide_len / 46), 1)
        ws.set_row(s.r, 14 * lines + 3)
        addr = xl_rowcol_to_cell(s.r, 8)
        s.r += 1
        s._marks += marks
        return addr

    def long(self, s, qid, text, model, marks=4, guide="", rows=4):
        """Free-text question: question row then a tall answer box (student) / model answer (key)."""
        ws = s.ws
        ws.write(s.r, 1, qid, self.qtag)
        ws.merge_range(s.r, 2, s.r, 8, text, self.qtext)
        ws.write_number(s.r, 9, marks, s.b.nf("int", style="cell", align="center"))
        if self.key and guide:
            s.mtext(s.r, 10, LAST_COL, guide, self.guide)
        else:
            ws.merge_range(s.r, 10, s.r, LAST_COL, "", self.work if not self.key else self.guide)
        ws.set_row(s.r, 14 * max(math.ceil(len(text) / 95), math.ceil(len(guide) / 46) if self.key else 1, 1) + 3)
        s.r += 1
        if self.key:
            s.mtext(s.r, 2, LAST_COL, model, self.longkey)
            ws.set_row(s.r, max(14 * rows, 14 * (sum(math.ceil(len(p) / 140) for p in model.split("\n")) + 1)))
        else:
            ws.merge_range(s.r, 2, s.r, LAST_COL, "", self.longans)
            ws.set_row(s.r, 14 * rows)
        s.r += 2
        s._marks += marks

    # ------------------------------------------------------------------ build grids
    def fill(self, s, df, answers: dict, fmts=None, marks=0, guide=""):
        """Grid where df columns are given and `answers` columns are for the learner.
        answers = {col: (formula_template using {r} for the Excel row number, cached list, kind)}"""
        ws = s.ws
        fmts = fmts or {}
        cols = list(df.columns) + list(answers)
        for j, c in enumerate(cols):
            ws.write(s.r, 1 + j, c, self.F["hdr"] if c in df.columns else s.b.nf("gen", style="input", bold=True, text_wrap=True))
        s.r += 1
        top = s.r
        for i, rec in enumerate(df.itertuples(index=False)):
            for j, v in enumerate(rec):
                k = fmts.get(df.columns[j])
                s._w(s.r, 1 + j, v, s.b.nf(k) if k else self.F["cell"])
            for j, (c, (tpl, cached, kind)) in enumerate(answers.items(), start=len(df.columns)):
                if self.key and tpl is None:
                    s._w(s.r, 1 + j, cached[i], s.b.nf(kind or "gen", style="result"))
                elif self.key:
                    ws.write_formula(s.r, 1 + j, tpl.format(r=s.r + 1), s.b.nf(kind or "gen", style="result"), _py(cached[i]))
                else:
                    ws.write_blank(s.r, 1 + j, None, s.b.nf(kind or "gen", style="input"))
            s.r += 1
        if marks:
            ws.write(s.r, 1, "Marks", self.F["lab"])
            ws.write_number(s.r, 2, marks, s.b.nf("int", style="cell", align="center"))
            if self.key and guide:
                s.mtext(s.r, 3, LAST_COL, guide, self.guide)
            elif guide:
                ws.merge_range(s.r, 3, s.r, LAST_COL, "", self.work)
            s._marks += marks
            s.r += 1
        s.r += 1
        return top

    def build_table(self, s, df, name, calc: dict, fmts=None, marks=0, guide="", widths=None, style="Table Style Medium 2"):
        """Excel Table whose calculated columns the learner must create. Student: empty columns; key: formulas."""
        if self.key:
            t = s.table(df, name, fmts=fmts, calc=calc, widths=widths, style=style)
        else:
            d2 = df.copy()
            for c in calc:
                d2[c] = None
            f2 = dict(fmts or {})
            f2.update({c: calc[c][1] for c in calc if calc[c][1]})
            t = s.table(d2, name, fmts=f2, widths=widths, style=style)
            for c in calc:
                j = t.ci(c)
                s.ws.conditional_format(t.r0 + 1, j, t.r0 + t.n, j, {"type": "no_errors", "format": self.b.wb.add_format({"bg_color": "#FFF4C2"})})
        if marks:
            s._marks += marks
            s.ws.write(s.r - 1, 1, "Marks", self.F["lab"])
            s.ws.write_number(s.r - 1, 2, marks, s.b.nf("int", style="cell", align="center"))
            if self.key and guide:
                s.mtext(s.r - 1, 3, LAST_COL, guide, self.guide)
            s.r += 1
        return t

    # ------------------------------------------------------------------ cover + marks
    def finish(self, heading, intro, datasets, files=None, rubric=None):
        b = self.b
        total = sum(m[3] for m in self.marks)
        ms = b.sheet("99_Marks", "Marking scheme and score sheet", self.subtitle, "All tasks")
        ms.scenario("Marks per task. The instructor enters the awarded marks in the yellow column; percentages and the grade update automatically. "
                    "Pass mark: 50%. Distinction: 70% and above.")
        ms.ws.write(ms.r, 1, "Sheet", self.F["hdr"]); ms.ws.merge_range(ms.r, 2, ms.r, 6, "Task", self.F["hdr"])
        ms.ws.merge_range(ms.r, 7, ms.r, 8, "Modules", self.F["hdr"]); ms.ws.write(ms.r, 9, "Available", self.F["hdr"])
        ms.ws.write(ms.r, 10, "Awarded", self.F["hdr"]); ms.ws.write(ms.r, 11, "%", self.F["hdr"])
        ms.r += 1
        top = ms.r
        for sh, t, mod, mk in self.marks:
            ms.ws.write_url(ms.r, 1, f"internal:'{sh}'!A1", self.F["link"], sh)
            ms.ws.merge_range(ms.r, 2, ms.r, 6, t, self.F["cell"])
            ms.ws.merge_range(ms.r, 7, ms.r, 8, mod, self.F["cell"])
            ms.ws.write_number(ms.r, 9, mk, b.nf("int"))
            ms.ws.write_blank(ms.r, 10, None, b.nf("int", style="input"))
            ms.ws.write_formula(ms.r, 11, f'=IF(K{ms.r + 1}="","",K{ms.r + 1}/J{ms.r + 1})', b.nf("pct0"), "")
            ms.r += 1
        bot = ms.r
        ms.ws.merge_range(ms.r, 2, ms.r, 8, "TOTAL", self.F["lab"])
        ms.ws.write_formula(ms.r, 9, f"=SUM(J{top + 1}:J{bot})", b.nf("int", style="result"), total)
        ms.ws.write_formula(ms.r, 10, f'=IF(COUNT(K{top + 1}:K{bot})=0,"",SUM(K{top + 1}:K{bot}))', b.nf("int", style="result"), "")
        ms.ws.write_formula(ms.r, 11, f'=IF(K{ms.r + 1}="","",K{ms.r + 1}/J{ms.r + 1})', b.nf("pct0", style="result"), "")
        ms.r += 1
        ms.ws.merge_range(ms.r, 2, ms.r, 8, "Grade", self.F["lab"])
        g = ms.r
        ms.ws.write_formula(ms.r, 11, f'=IF(L{g}="","",IF(L{g}>=0.7,"Distinction",IF(L{g}>=0.5,"Pass","Resubmit")))', b.nf("gen", style="result"), "")
        ms.r += 2
        if rubric:
            ms.h2("General marking rules")
            ms.grid(rubric, header=["Rule", "Detail"], merge_last=True)
        self.marks_total = total
        extra = [("Total marks", f"{total}"), ("Submission", "Save as Surname_Firstname_Week%d.xlsx. Do not rename sheets, move answer cells or delete questions." % self.week),
                 ("How to answer", "Type in the YELLOW cells only. Where a question says 'formula', the cell must contain a live formula - typed numbers score zero."),
                 ("Integrity", "Individual work. You may use the slides, the Section Examples workbook and Microsoft documentation. If you use an AI assistant, say which one and for what in the Working / notes cell beside the answer.")]
        if self.key:
            extra.insert(0, ("ANSWER KEY", "Instructor copy - do not distribute. Green cells hold live formulas with cached results; tolerances are in each marking guide."))
        if files:
            extra.append(("Data files", files))
        b.cover(heading, intro, datasets, extra_rows=extra)
        b.close()
        return self.path
