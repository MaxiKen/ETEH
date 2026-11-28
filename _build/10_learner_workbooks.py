"""Edition 3.1 learner workbooks: one Section Examples workbook per week, in its own week folder.

    python3 _build/10_learner_workbooks.py [1 2 3 4 5]      (5 = the Capstone week)

Each workbook: Start_Here › one sheet per module (questions in teaching order) › Data_<Name> sheets › Solutions.
Files that an exercise links to (CSV, TXT, database ...) are copied into an Ex<module>_Files folder next to the
workbook, and each sheet names its data as File › Sheet › Table (or the folder that holds the file).
"""
from __future__ import annotations

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path[:0] = [HERE, os.path.join(HERE, "v3"), os.path.join(HERE, "learn")]

import lkit  # noqa: E402
import paths  # noqa: E402

WEEKS = {
    1: ("Week 1 · Data Analytics Foundations & Core Microsoft Excel", "w1"),
    2: ("Week 2 · External Data, Power Query, Functions & PivotTables", "w2"),
    3: ("Week 3 · Databases, SQL, Data Model, DAX & Dashboards", "w3"),
    4: ("Week 4 · Statistics, Forecasting, Reporting, AI & Ethics", "w4"),
    5: ("Week 5 · Capstone Projects, Portfolio & Career", "w5"),
}


def build_week(n):
    title, modname = WEEKS[n]
    folder = paths.sec(n)
    folder.mkdir(parents=True, exist_ok=True)
    fname = "ETEH_Capstone_Section_Examples.xlsx" if n == 5 else f"ETEH_Week{n}_Section_Examples.xlsx"
    b = lkit.LBook(folder / fname, fname, n, title, week_label="Week 5 (Capstone)" if n == 5 else None)
    mod = __import__(modname)
    mod.build(b)
    b.close()
    q = sum(m.nq for m in b.modules)
    print(f"wrote {paths.rel(folder / fname)}: {len(b.modules)} module sheets, {q} questions, {len(b.datas)} data sheets")
    return b


if __name__ == "__main__":
    for w in (sys.argv[1:] or ["1", "2", "3", "4", "5"]):
        build_week(int(w))
