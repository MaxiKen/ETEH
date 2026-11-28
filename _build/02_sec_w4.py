"""Week 4 Examples workbook, edition 3 (modules 4.1 - 4.14).

One sheet per module, opened by the tutor after the module's slides. Each sheet starts with the module's dataset, then covers
every mini-section (139 in total) with Steps, a Worked solution (live formulas) and 'Your turn' exercises with live checks, and
ends with an Answer key. Shared tables sit on Data_ sheets, each starting in A1, with the table names printed on the slides
(w4_data.TABLES) plus tCustRev (net revenue per customer, the Pareto source). Every slide figure is a live formula here.
The edition-2 builder is kept in rev/02_sec_w4_edition2.py.

    python3 _build/02_sec_w4.py
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "v3"))
from common import SEC  # noqa: E402
import w1x  # noqa: E402
from xlhelp import Book  # noqa: E402
import wb4_common as C  # noqa: E402

PARTS = []
for _name in ("wb4_a", "wb4_b", "wb4_c"):
    try:
        PARTS.append(__import__(_name))
    except ModuleNotFoundError as e:
        if e.name != _name:
            raise

OUT = SEC / "ETEH_Week4_Section_Examples.xlsx"
EXPECT = {"4.1": 14, "4.2": 9, "4.3": 10, "4.4": 12, "4.5": 13, "4.6": 12, "4.7": 9, "4.8": 10, "4.9": 10, "4.10": 8, "4.11": 10,
          "4.12": 8, "4.13": 9, "4.14": 5}


def main():
    D = C.data()
    b = Book(str(OUT), "ETEH Week 4 - Examples Workbook", "Week 4 · Statistics, Forecasting, What-If, Reporting, AI and Ethics")
    b.wb.strings_to_formulas = False
    for part in PARTS:
        part.build(b, D)
    C.data_sheets(b, D)
    want = [f"{m}.{k}" for m, c in EXPECT.items() for k in range(1, c + 1)]
    missing = [t for t in want if t not in w1x.COVERED]
    extra = [t for t in w1x.COVERED if t not in want]
    if len(PARTS) == 3:
        assert not missing and not extra, (missing, extra)
    b.cover(
        "Week 4 Examples Workbook",
        ["How to use it: after each module's slides, your tutor opens that module's sheet and works through it with you. "
         "Every sheet follows the same order: the module's dataset, then each topic (gold tag = slide topic number) with Steps, "
         "a Worked solution and 'Your turn' exercises, then an Answer key at the bottom.",
         "Your turn: type your answer in the white cell with the gold border. The check cell beside it turns to '✓ Correct' "
         "when your answer matches the worked answer. 'Do it in Excel' tasks have a Done? drop-down and a 'What you should see' note.",
         "Statistics (4.1 - 4.4) use worksheet functions you can click and read; the Analysis ToolPak tasks reproduce the same numbers. "
         "LINEST is read with INDEX(LINEST(...), row, column): coefficients run right to left, the intercept last.",
         "Forecasting (4.5) works on tMonthly: Sales is column E, rows 2-49 of Data_Monthly, the cells quoted on the slides. "
         "What-if (4.6) has a live Atlantic Pearl profit model with yellow input cells for Goal Seek, data tables, scenarios and Solver.",
         f"All {len(want)} Week 4 topics are covered. Every number is a live formula on the bootcamp data, so you can click any "
         "result and read how it was worked out."],
        [("ETEH Talent Salary Survey", f"tSal: {len(D.S):,} respondents. Descriptive statistics, correlation, regression and AI prompts."),
         ("SuperMart Online / Stores", f"tOrders ({len(D.O):,} order lines), tCustRev, tDash25 ({len(D.D25):,} FY2025 lines), tStores, tMonthly (48 months)."),
         ("Mercyland Specialist Hospital", f"tLOS ({len(D.LOS):,} admissions in 2023) and tCensus23 (365 days, 200 beds): report writing and ethics."),
         ("Kingsway University", f"tCV: {len(D.CV):,} students. Distributions and correlation."),
         ("Atlantic Pearl Hotel", f"tHotel: {len(D.H):,} reservations with Is_Cancelled. Simpson's paradox, what-if and Solver.")],
        extra_rows=[("Answer cells", "White cell with a gold border: type here. The check cell beside it updates as you type."),
                    ("Assignment", "The Week 4 assignment is in 02_Assignments/ETEH_Week4_Assignment.xlsx.")])
    b.close()
    print(f"wrote {OUT.name}: {len(w1x.COVERED)} topics covered; missing {len(missing)}")


if __name__ == "__main__":
    main()
