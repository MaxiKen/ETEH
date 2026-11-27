"""Week 1 Examples workbook, edition 3 (modules 1.1 - 1.9).

One sheet per module, opened by the tutor after the module's slides. Each sheet starts with the module's dataset,
then covers every mini-section (111 in total) with Steps, a Worked solution (live formulas) and 'Your turn'
exercises with live checks, and ends with an Answer key. Shared tables sit on Data_ sheets, each starting in A1.
Table names are the ones printed on the slides (see w1_data.py).

    python3 _build/02_sec_w1.py
"""
from __future__ import annotations

import os
import re
import sys
from types import SimpleNamespace

import pandas as pd

sys.path.insert(0, os.path.dirname(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "v3"))
from common import SEC  # noqa: E402
import w1_data as W  # noqa: E402
import w1x  # noqa: E402
from w1x import DataSheet  # noqa: E402
from xlhelp import Book  # noqa: E402
import wb_a  # noqa: E402
import wb_b  # noqa: E402
import wb_c  # noqa: E402

OUT = SEC / "ETEH_Week1_Section_Examples.xlsx"
EXPECT = {"1.1": 14, "1.2": 6, "1.3": 12, "1.4": 10, "1.5": 17, "1.6": 14, "1.7": 14, "1.8": 12, "1.9": 12}


def data():
    s25 = W.s25()
    d = SimpleNamespace()
    d.ST = W.table("tStores25")
    d.SALES = W.table("tSales")
    d.IK = W.table("tIkejaLines")
    d.Y = W.table("tYears")
    d.DB = W.table("tDiscBand")
    d.CAT = W.table("tCat25")
    d.STU = W.table("tStudents")
    d.FAC = W.table("tFaculty")
    d.DEPT = W.table("tDept23")
    d.HSEG = W.table("tHotelSeg")
    d.SEN = W.table("tSeniority")
    d.PAY = W.table("tPayMix")
    d.COUR = W.table("tCouriers")
    q = s25.groupby(s25.Order_Date.dt.quarter).Sales.sum()
    d.Q25 = pd.DataFrame({"Quarter": [f"Q{i}" for i in q.index], "Sales": q.values})
    f = s25[s25.Category == "Furniture"]
    d.FURN = f.groupby("Discount").agg(Lines=("Row_ID", "count"), Sales=("Sales", "sum"), Profit=("Profit", "sum")).reset_index()
    dirty = W.table("tDirtySales").copy()
    iso = re.compile(r"\d{4}-\d\d-\d\d$")
    dirty["Order_Date"] = [pd.Timestamp(v) if isinstance(v, str) and iso.match(v) else v for v in dirty.Order_Date]
    d.DIRTY = dirty
    d.s25 = s25
    return d


def data_sheets(b, D):
    st = DataSheet(b, "Data_Stores25", "tStores25: SuperMart FY2025 results, one row per store (A1)", "Used in 1.1, 1.3, 1.5 - 1.8")
    base = D.ST.drop(columns=["Margin", "Attainment"])
    st.put(base, "tStores25", fmts={"Sales": "naira", "Profit": "naira", "Target": "naira", "Orders": "int"},
           calc={"Margin": ("=[@Profit]/[@Sales]", "pct", D.ST.Margin.tolist()),
                 "Attainment": ("=[@Sales]/[@Target]", "pct", D.ST.Attainment.tolist())},
           widths={"Store_ID": 9, "Store_Name": 27, "Zone": 14, "Store_Manager": 18, "Sales": 14, "Profit": 13,
                   "Orders": 8, "Target": 14, "Margin": 9, "Attainment": 11})
    sa = DataSheet(b, "Data_Sales", "tSales: SuperMart sales lines 2022-2025, one row per product line (A1)",
                   f"{len(D.SALES):,} rows · used in 1.2, 1.3, 1.5, 1.8")
    sa.put(D.SALES, "tSales", fmts={"Row_ID": "gen", "Order_Date": "date", "Sales": "naira", "Quantity": "int",
                                    "Discount": "pct0", "Profit": "naira"},
           widths={"Row_ID": 8, "Order_ID": 16, "Order_Date": 11, "Customer_Name": 20, "Store_Name": 26, "Category": 15,
                   "Sub_Category": 13, "Sales": 12, "Quantity": 9, "Discount": 9, "Profit": 12})
    du = DataSheet(b, "Data_DirtySales", "tDirtySales: the raw FY2025 export, exactly as received (A1)",
                   f"{len(D.DIRTY):,} rows, deliberately untidy · used in 1.3, 1.6, 1.8, 1.9")
    du.put(D.DIRTY, "tDirtySales", fmts={"Order_Date": "date", "Row_ID": "gen"},
           widths={"Row_ID": 8, "Order_ID": 16, "Order_Date": 12, "Store_Name": 26, "Zone": 16, "Category": 15,
                   "Sales": 12, "Discount": 9, "Profit": 10})
    su = DataSheet(b, "Data_Students", "tStudents: Crestview University students, 2024/2025 first semester (A1)",
                   f"{len(D.STU):,} rows · used in 1.2, 1.7")
    su.put(D.STU, "tStudents", fmts={"Level": "gen"},
           widths={"Student_ID": 18, "Faculty": 17, "Department": 24, "Level": 7, "Gender": 8, "Attendance_Pct": 14,
                   "Hours_Studied_Per_Week": 21, "Exam_Score": 11, "Grade": 7, "At_Risk": 8})


def main():
    D = data()
    b = Book(str(OUT), "ETEH Week 1 - Examples Workbook", "Week 1 · Data Analytics Foundations & Core Microsoft Excel Mastery")
    # Week 1 text often starts with '=' (formulas shown as text): store plain strings as text, never as formulas.
    b.wb.strings_to_formulas = False
    for _ws in b.wb.worksheets():
        _ws.strings_to_formulas = False
    wb_a.build(b, D)
    wb_b.build(b, D)
    wb_c.build(b, D)
    data_sheets(b, D)
    # every mini-section is covered exactly where the slides teach it
    want = [f"{m}.{k}" for m, c in EXPECT.items() for k in range(1, c + 1)]
    missing = [t for t in want if t not in w1x.COVERED]
    extra = [t for t in w1x.COVERED if t not in want]
    assert not missing and not extra, (missing, extra)
    b.cover(
        "Week 1 Examples Workbook",
        ["How to use it: after each module's slides, your tutor opens that module's sheet and works through it with you. "
         "Every sheet follows the same order: the module's dataset, then each topic (gold tag = slide topic number) with Steps, "
         "a Worked solution and 'Your turn' exercises, then an Answer key at the bottom.",
         "Your turn: type your answer in the white cell with the gold border. The check cell beside it turns to '✓ Correct' "
         "when your answer matches the worked answer. 'Do it in Excel' tasks have a Done? drop-down and a 'What you should see' note.",
         "Data sheets (Data_...) hold one Excel table each, starting in cell A1, exactly as the slides describe them. "
         "Treat them as raw data: do your practice on the module sheets, or on a copy of a data sheet (right-click the tab › Move or Copy).",
         f"All {len(want)} Week 1 topics are covered. Every number is a live formula on the bootcamp data, so you can click any "
         "result and read how it was worked out."],
        [("SuperMart Retail", "tStores25, tSales, tIkejaLines, tYears, tQuarter25, tDiscBand, tFurn25, tCat25 and tDirtySales: "
                              "12 Nigerian stores, 9,994 sales lines 2022-2025 and the raw FY2025 export (3,387 rows)."),
         ("Crestview University", "tStudents and tFaculty: 6,607 students, 2024/2025 first semester."),
         ("Lifecare General Hospital", "tDept23: 10,918 admissions in 2023, summarised by department."),
         ("Eko Suites Hotel", "tHotelSeg: 36,275 bookings 2023-2024, summarised by market segment."),
         ("ETEH Talent Salary Survey", "tSeniority: 1,781 professionals, summarised by seniority."),
         ("SuperMart Online", "tPayMix and tCouriers: 11,000 orders and their deliveries, summarised."),
         ("Practice data", "A few small tables are marked 'illustrative practice data'; they are made up for a skill and "
                           "are not from a bootcamp dataset.")],
        extra_rows=[("Answer cells", "White cell with a gold border: type here. The check cell beside it updates as you type."),
                    ("Assignment", "The Week 1 assignment is in 02_Assignments/ETEH_Week1_Assignment.xlsx.")])
    b.close()
    print(f"wrote {OUT.name}: {len(w1x.COVERED)} topics covered")


if __name__ == "__main__":
    main()
