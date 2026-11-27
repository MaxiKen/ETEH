"""Week 2 Examples workbook, edition 3 (modules 2.1 - 2.11, with 2.8 in two parts).

One sheet per module, opened by the tutor after the module's slides. Each sheet starts with the module's dataset, then covers
every mini-section (175 in total) with Steps, a Worked solution (live formulas) and 'Your turn' exercises with live checks, and
ends with an Answer key. Shared tables sit on Data_ sheets, each starting in A1, so cell addresses quoted on the slides are true.
Table names are the ones printed on the slides (see w2_data.py). The edition-2 builder is kept in rev/02_sec_w2_edition2.py.

    python3 _build/02_sec_w2.py
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "v3"))
from common import SEC  # noqa: E402
import w1x  # noqa: E402
from w1x import DataSheet  # noqa: E402
from xlhelp import Book  # noqa: E402
import wb2_common as C  # noqa: E402
import wb2_a  # noqa: E402

PARTS = [wb2_a]
try:
    import wb2_b  # noqa: E402
    PARTS.append(wb2_b)
    import wb2_c  # noqa: E402
    PARTS.append(wb2_c)
    import wb2_d  # noqa: E402
    PARTS.append(wb2_d)
except ImportError:
    pass

OUT = SEC / "ETEH_Week2_Section_Examples.xlsx"
EXPECT = {"2.1": 8, "2.2": 10, "2.3": 8, "2.4": 5, "2.5": 12, "2.6": 20, "2.7": 6, "2.8": 65, "2.9": 11, "2.10": 19, "2.11": 11}
SALES_FMT = {"Row_ID": "gen", "Order_Date": "date", "Ship_Date": "date", "Sales": "naira", "Quantity": "int", "Discount": "pct0",
             "Profit": "naira", "Year": "gen"}
SALES_W = {"Row_ID": 8, "Order_ID": 16, "Order_Date": 11, "Ship_Date": 11, "Ship_Mode": 14, "Customer_ID": 11, "Customer_Name": 19,
           "Segment": 11, "Store_ID": 8, "Store_Name": 26, "City": 13, "State": 11, "Zone": 13, "Product_ID": 16, "Category": 15,
           "Sub_Category": 13, "Product_Name": 40, "Sales": 12, "Quantity": 8, "Discount": 8, "Profit": 12}


def _w(df, widths):
    return {k: v for k, v in widths.items() if k in df.columns}


def data_sheets(b, D):
    sa = DataSheet(b, "Data_Sales25", "tSales25: SuperMart FY2025 sales lines after cleaning, one row per product line (A1)",
                   f"{len(D.S):,} rows · used in 2.2 - 2.8, 2.10")
    sa.put(D.S, "tSales25", fmts=SALES_FMT, widths=_w(D.S, SALES_W))
    st = DataSheet(b, "Data_Stores", "tStores: the 12 SuperMart stores with FY2025 target and sales (A1)", "Used in 2.1, 2.2, 2.8")
    st.put(D.ST, "tStores", fmts={"Opened": "date", "Floor_Area_sqm": "int", "Target_2025": "naira", "Sales_2025": "naira"},
           widths={"Store_ID": 9, "Store_Name": 27, "City": 14, "State": 11, "Zone": 14, "Opened": 11, "Floor_Area_sqm": 14,
                   "Store_Manager": 18, "Target_2025": 14, "Sales_2025": 14})
    du = DataSheet(b, "Data_Dirty25", "tDirty25: SuperMart_Sales_2025_DIRTY.csv exactly as Excel opens it - do not clean here (A1)",
                   f"{len(D.DR):,} rows incl. 15 blank rows and 60 duplicates · used in 2.1, 2.2, 2.5 - 2.8")
    du.put(D.DR, "tDirty25", fmts={"Order_Date": "date", "Ship_Date": "date", "Row_ID": "gen"},
           widths=_w(D.DR, {**SALES_W, "Store_Location": 20, "Column1": 9, "Order_Date": 12}))
    sm = DataSheet(b, "Data_SM", "tSM: every SuperMart sales line 2022-2025 (clean), one row per product line (A1)",
                   f"{len(D.SM):,} rows · the source of the PivotTables in 2.10 and 2.11")
    tsm = sm.put(D.SM, "tSM", fmts=SALES_FMT, widths=_w(D.SM, SALES_W))
    ad = DataSheet(b, "Data_Adm23", "tAdm23: Lifecare admissions 2023 from the pipe-delimited file, 14 of 22 columns (A1)",
                   f"{len(D.A):,} rows · used in 2.1, 2.3, 2.8, 2.11")
    ad.put(D.A, "tAdm23", fmts={"Admission_Date": "date", "Discharge_Date": "date", "Billing_Amount": "naira", "Age": "gen",
                                "Length_of_Stay": "gen"},
           widths={"Admission_ID": 15, "Patient_Name": 22, "Department": 20, "Admission_Type": 13, "Admission_Date": 12,
                   "Discharge_Date": 12, "Ward": 14, "Doctor_Name": 20, "HMO": 16, "Billing_Amount": 14})
    return tsm


def main():
    D = C.data()
    b = Book(str(OUT), "ETEH Week 2 - Examples Workbook",
             "Week 2 · Data Import, Power Query Cleaning & PivotTable Mastery")
    b.wb.strings_to_formulas = False
    for _ws in b.wb.worksheets():
        _ws.strings_to_formulas = False
    for part in PARTS:
        part.build(b, D)
    tsm = data_sheets(b, D)
    for part in PARTS:
        if hasattr(part, "after_data"):
            part.after_data(b, D, tsm)
    want = [f"{m}.{k}" for m, c in EXPECT.items() for k in range(1, c + 1)]
    missing = [t for t in want if t not in w1x.COVERED]
    extra = [t for t in w1x.COVERED if t not in want]
    if len(PARTS) == 4:
        assert not missing and not extra, (missing, extra)
    b.cover(
        "Week 2 Examples Workbook",
        ["How to use it: after each module's slides, your tutor opens that module's sheet and works through it with you. "
         "Every sheet follows the same order: the module's dataset, then each topic (gold tag = slide topic number) with Steps, "
         "a Worked solution and 'Your turn' exercises, then an Answer key at the bottom.",
         "Your turn: type your answer in the white cell with the gold border. The check cell beside it turns to '✓ Correct' "
         "when your answer matches the worked answer. 'Do it in Excel' tasks have a Done? drop-down and a 'What you should see' note.",
         "Power Query work happens in the Power Query Editor on the practice files in 01_Section_Examples/Week2_Data_Files. "
         "The exercises then ask for a count or total from your query; the answer key works it out live from the tables here.",
         "Data sheets (Data_...) hold one Excel table each, starting in cell A1, exactly as the slides describe them. "
         "Data_Dirty25 is the raw file: practise cleaning on a copy (right-click the tab › Move or Copy).",
         f"All {len(want)} Week 2 topics are covered. Every number is a live formula on the bootcamp data, so you can click any "
         "result and read how it was worked out."],
        [("SuperMart Retail", "tSales25 (3,312 clean FY2025 lines), tDirty25 (the raw FY2025 export, 3,387 rows), tSM (9,994 lines "
                              "2022-2025), tStores (12 stores) and the monthly files, web page and PDF bulletin."),
         ("Lifecare General Hospital", "tAdm23: 10,918 admissions in 2023 (pipe-delimited file), plus the dirty 2023 extract used in 2.9."),
         ("Crestview University", "6,607 students, used in the M examples of 2.9."),
         ("SuperMart Online", "SuperMart_Online_RAW_Export.xlsx (5 sheets) and the sellers JSON, used in 2.1 and 2.3.")],
        extra_rows=[("Answer cells", "White cell with a gold border: type here. The check cell beside it updates as you type."),
                    ("Folder path", "Set your own folder once in the DataFolder cell on sheet 2.3 Connect & Combine."),
                    ("Assignment", "The Week 2 assignment is in 02_Assignments/ETEH_Week2_Assignment.xlsx.")])
    b.close()
    print(f"wrote {OUT.name}: {len(w1x.COVERED)} topics covered; missing {len(missing)}")


if __name__ == "__main__":
    main()
