"""Capstone (Volume 5) shared data for the edition-3 deck: the named tables of the capstone section-examples workbook and
the instructor capstone solutions, excerpted on slides by their Excel table names."""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "v3"))
import w4_data  # noqa: E402  (shared cell formatting)
import w5_fx as F  # noqa: E402

TABLES = {
    "tSM": lambda: F.SM[["Order_ID", "Order_Date", "Store_Name", "Category", "Sales", "Discount", "Profit"]],
    "tBasket": lambda: F.O25.reset_index(),
    "tSKU": lambda: F.SKU,
    "tCV": lambda: F.CV[["Matric_No", "Faculty", "Level", "Final_Score", "Lecture_Attendance_Pct"]],
    "tAtRiskTop20": lambda: F.LST.sort_values(["Final_Score", "Lecture_Attendance_Pct"], kind="stable").head(20),
    "tAdm23": lambda: F.ADM[["episode_id", "admit_date", "ward", "department", "admission_route", "los_days", "readmitted_30d"]]
    .rename(columns={"episode_id": "Episode_ID", "admit_date": "Admit_Date", "ward": "Ward", "department": "Department",
                     "admission_route": "Admission_Route", "los_days": "LOS_Days", "readmitted_30d": "Readmitted_30d"}),
    "tCensus": lambda: F.CEN,
}


def table(name):
    return TABLES[name]().copy()


def excerpt(name, cols=None, n=5, rows=None, sort=None, ascending=False):
    df = table(name)
    if rows is not None:
        df = rows(df)
    if sort:
        df = df.sort_values(sort, ascending=ascending)
    cols = cols or list(df.columns)
    total = len(df)
    recs = df.head(n)[cols].to_dict("records")
    return [list(cols)] + [[w4_data.fmt(c, r[c]) for c in cols] for r in recs], total
