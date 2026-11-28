"""Week 3 Examples workbook, edition 3 (modules 3.1 - 3.10, with 3.6 in two parts).

One sheet per module, opened by the tutor after the module's slides. Each sheet starts with the module's dataset, then covers
every mini-section (158 in total) with Steps, a Worked solution (live formulas) and 'Your turn' exercises with live checks, and
ends with an Answer key. Shared tables sit on Data_ sheets, each starting in A1. Table names are the ones printed on the slides
(w3_data.TABLES) plus tLines, the v_order_lines view imported through ODBC. SQL and DAX shown on a sheet are the slide copy.
Also writes Week3_Data_Files/ETEH_Week3_SQL_Examples.sql: every query shown on the slides, with the database it runs on.
The edition-2 builder is kept in rev/02_sec_w3_edition2.py.

    python3 _build/02_sec_w3.py
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "v3"))
from common import SEC  # noqa: E402
import w1x  # noqa: E402
from xlhelp import Book  # noqa: E402
import w3_data as W  # noqa: E402
import wb3_common as C  # noqa: E402

PARTS = []
for _name in ("wb3_a", "wb3_b", "wb3_c"):
    try:
        PARTS.append(__import__(_name))
    except ModuleNotFoundError as e:
        if e.name != _name:
            raise

OUT = SEC / "ETEH_Week3_Section_Examples.xlsx"
SQL_OUT = SEC / "Week3_Data_Files" / "ETEH_Week3_SQL_Examples.sql"
EXPECT = {"3.1": 15, "3.2": 17, "3.3": 9, "3.4": 16, "3.5": 7, "3.6": 41, "3.7": 26, "3.8": 19, "3.9": 3, "3.10": 5}


def _key(tag):
    return tuple(int(x) for x in tag.split("."))


def write_sql():
    """Every SQL panel on the slides, in topic order, with the database it runs on (taken from the slide copy itself)."""
    entries = []
    for tag in sorted(C.DEF, key=_key):
        t = C.DEF[tag]
        label = t.get("code_label", "")
        if not t.get("code") or not ("SQL" in label.upper() and "SHELL" not in label.upper()):
            continue
        db = t.get("code_db")
        where = (f"-- database: {W.DBNAME[db]}" if db else
                 "-- database: a staging copy - load Mercyland_Episodes_2023_DIRTY.csv as stg_admissions_2023 (3.10.2), then\n"
                 "--           ATTACH DATABASE 'mercyland_hospital.db' AS prod;")
        entries.append(f"-- {tag} {t['title']}\n{where}\n{t['code'].strip().rstrip(';')};\n")
    head = ("-- ETEH Week 3 - SQL examples (SQLite), edition 3: every query shown on the Week 3 slides.\n"
            "-- Run each against the database named above it (folder Week3_Data_Files) in DB Browser for SQLite or sqlite3.\n"
            "-- The result printed on the slide is what the query returns on these files.\n\n")
    SQL_OUT.write_text(head + "\n".join(entries), encoding="utf-8")
    return len(entries)


def main():
    D = C.data()
    b = Book(str(OUT), "ETEH Week 3 - Examples Workbook", "Week 3 · Databases, SQL, the Data Model, DAX and Dashboards")
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
        "Week 3 Examples Workbook",
        ["How to use it: after each module's slides, your tutor opens that module's sheet and works through it with you. "
         "Every sheet follows the same order: the module's dataset, then each topic (gold tag = slide topic number) with Steps, "
         "a Worked solution and 'Your turn' exercises, then an Answer key at the bottom.",
         "Your turn: type your answer in the white cell with the gold border. The check cell beside it turns to '✓ Correct' "
         "when your answer matches the worked answer. 'Do it in Excel' tasks have a Done? drop-down and a 'What you should see' note.",
         "SQL work happens in DB Browser for SQLite (or sqlite3) on the files in 01_Section_Examples/Week3_Data_Files; every query "
         "on the slides is in ETEH_Week3_SQL_Examples.sql there. The exercises ask for the number your query returns; the answer "
         "key works the same number out live from the tables here, so SQL, DAX and worksheet formulas can be compared.",
         "DAX work happens in Power Pivot on a copy of this workbook (add tTransactions, tProducts, tCustomers, tSellers and "
         "tDimDate to the Data Model, 3.6.3). Each DAX result has a worksheet 'twin' formula here (3.6.39).",
         f"All {len(want)} Week 3 topics are covered. Every number is a live formula on the bootcamp data, so you can click any "
         "result and read how it was worked out."],
        [("SuperMart Online", "supermart_online.db: tTransactions (11,000 order lines 2020-2025), tProducts, tCustomers, tSellers, "
                              "tDeliveries, tStates, and tLines (the v_order_lines view). Used in 3.1 - 3.6."),
         ("SuperMart Stores", "tDash: 3,312 FY2025 sales lines and tDashStores (12 stores with targets), the dashboard data of 3.7 - 3.9."),
         ("Mercyland Specialist Hospital", "tAdm23 (12,626 admissions in 2023), tCensus (daily beds) and tMsDirty, the raw 2023 extract of 3.10."),
         ("Kingsway University", "tCV: 5,840 students, used in the KPI, chart and blueprint examples.")],
        extra_rows=[("Answer cells", "White cell with a gold border: type here. The check cell beside it updates as you type."),
                    ("Date table", "tDimDate on Data_DimDate is the marked date table for the time-intelligence measures."),
                    ("Assignment", "The Week 3 assignment is in 02_Assignments/ETEH_Week3_Assignment.xlsx.")])
    b.close()
    n_sql = write_sql()
    print(f"wrote {OUT.name}: {len(w1x.COVERED)} topics covered; missing {len(missing)}; {n_sql} SQL examples -> {SQL_OUT.name}")


if __name__ == "__main__":
    main()
