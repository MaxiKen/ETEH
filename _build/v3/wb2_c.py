"""Week 2 workbook, module 2.8 in two parts (edition 3): the 65 analytical Excel functions, one topic per function.

Definitions are taken from the slide copy (w2_m8.py) so slide and workbook say the same thing; every worked result is a live formula
on tSales25, tStores, tDirty25 or tAdm23, and every function has at least one 'Your turn' exercise with a live answer key.
"""
from __future__ import annotations

import datetime as dt

import numpy as np
import pandas as pd

import w2_m8 as M8
from w1x import ModuleSheet
from wb2_common import N, serial
from w2_fx import DF

TODAY = dt.date(2026, 11, 27)
DEF = {t["tag"]: t for t in M8.TOPICS_A + M8.TOPICS_B}


def fn(m, tag, worked, ex, steps=None, note=None):
    t = DEF[tag]
    m.topic(tag, t["title"], t["defn"] + "  " + t["simple"])
    if steps:
        m.steps(steps)
    m.worked(worked)
    if note:
        m.note(note)
    m.turn()
    for e in ex:
        if e[0] == "n":
            _, task, fx, val, kind = e[:5]
            m.ex(task, fx, val, kind, short=e[5] if len(e) > 5 else None)
        elif e[0] == "t":
            _, task, ans = e[:3]
            fx = e[3] if len(e) > 3 else None
            if fx and ans in ("TRUE", "FALSE"):
                ans = ans == "TRUE"
            m.ex_text(task, ans, formula=fx, short=e[4] if len(e) > 4 else None, alts=e[5] if len(e) > 5 else ())
        else:
            m.ex_open(e[1], e[2])


def _att(S, ST, sid):
    x = S[S.Store_ID == sid]
    tgt = float(ST.set_index("Store_ID").loc[sid, "Target_2025"])
    return float(x.Sales.sum()), tgt, float(x.Profit.sum()) / float(x.Sales.sum())


def _sid(sid):
    return (f'SUMIFS(tSales25[Sales],tSales25[Store_ID],"{sid}")', f'XLOOKUP("{sid}",tStores[Store_ID],tStores[Target_2025])',
            f'SUMIFS(tSales25[Profit],tSales25[Store_ID],"{sid}")/SUMIFS(tSales25[Sales],tSales25[Store_ID],"{sid}")')


# =====================================================================================================
def m28a(b, D):
    S, ST, DR, A = D.S, D.ST, D.DR, D.A
    m = ModuleSheet(b, "2.8A Functions 1-33", "2.8  The 65 Analytical Excel Functions - Part A (1-33)",
                    "Tutor walk-through after Module 2.8A: aggregation, counting, statistics, rounding, logic and joining text, each on "
                    "the FY2025 sales lines", "2.8.1 - 2.8.33")
    m.dataset([("tSales25", "Sheet Data_Sales25 (starts in A1)", "3,312 clean FY2025 sales lines. Most examples use it."),
               ("tStores", "Sheet Data_Stores (starts in A1)", "12 stores with Target_2025 and Sales_2025."),
               ("tDirty25", "Sheet Data_Dirty25 (starts in A1)", "The raw file, used where a function measures a data problem."),
               ("tAdm23", "Sheet Data_Adm23 (starts in A1)", "Mercyland 2023 admissions, for a second dataset.")])
    TOT = int(S.Sales.sum())
    num = [v for v in DR.Sales if isinstance(v, (int, float)) and v == v]
    zsum = lambda z: int(S[S.Zone == z].Sales.sum())  # noqa: E731

    fn(m, "2.8.1", [("FY2025 sales", "=SUM(tSales25[Sales])", TOT, N, ""),
                    ("Numbers summed", "=COUNT(tSales25[Sales])", len(S), "int", ""),
                    ("Non-empty cells", "=COUNTA(tSales25[Sales])", len(S), "int", "Equal to COUNT: no text hiding in the column.")],
       [("n", "What was the total FY2025 profit?", "=SUM(tSales25[Profit])", int(S.Profit.sum()), N, "Total profit"),
        ("n", "SUM the raw tDirty25[Sales] column. What do you get? (It is wrong: 2.7.1 explains why.)", "=SUM(tDirty25[Sales])", int(sum(num)), N,
         "SUM of the raw Sales")])
    fn(m, "2.8.2", [("Technology sales", '=SUMIF(tSales25[Category],"Technology",tSales25[Sales])', int(S[S.Category == "Technology"].Sales.sum()), N, ""),
                    ("Total of all losses", '=SUMIF(tSales25[Profit],"<0")', int(S[S.Profit < 0].Profit.sum()), N, "No sum_range: the Profit cells themselves.")],
       [("n", "What were the FY2025 sales of store ST04?", '=SUMIF(tSales25[Store_ID],"ST04",tSales25[Sales])', int(S[S.Store_ID == "ST04"].Sales.sum()), N,
         "Sales of ST04"),
        ("n", "What is the total profit of all profitable lines?", '=SUMIF(tSales25[Profit],">0")', int(S[S.Profit > 0].Profit.sum()), N, "Profit of profitable lines")])
    fn(m, "2.8.3", [("Technology in the South-West", '=SUMIFS(tSales25[Sales],tSales25[Category],"Technology",tSales25[Zone],"South-West")',
                     int(S[(S.Category == "Technology") & (S.Zone == "South-West")].Sales.sum()), N, "Sum range first.")],
       [("n", "What were Furniture sales in the South-East?", '=SUMIFS(tSales25[Sales],tSales25[Category],"Furniture",tSales25[Zone],"South-East")',
         int(S[(S.Category == "Furniture") & (S.Zone == "South-East")].Sales.sum()), N, "Furniture, South-East"),
        ("n", "What profit did Consumer-segment lines make at ST01?", '=SUMIFS(tSales25[Profit],tSales25[Segment],"Consumer",tSales25[Store_ID],"ST01")',
         int(S[(S.Segment == "Consumer") & (S.Store_ID == "ST01")].Profit.sum()), N, "Consumer profit at ST01")])
    fn(m, "2.8.4", [("Average line", "=AVERAGE(tSales25[Sales])", float(S.Sales.mean()), N, ""),
                    ("Same, by hand", "=SUM(tSales25[Sales])/COUNT(tSales25[Sales])", float(S.Sales.mean()), N, "")],
       [("n", "What is the average profit per line?", "=AVERAGE(tSales25[Profit])", float(S.Profit.mean()), N, "Average profit per line")])
    c1 = S[(S.Segment == "Corporate") & (S.Store_ID == "ST01")]
    tn = S[(S.Category == "Technology") & (S.Zone == "North-Central")]
    fn(m, "2.8.5", [("Corporate line at Ikeja", '=AVERAGEIFS(tSales25[Sales],tSales25[Segment],"Corporate",tSales25[Store_ID],"ST01")',
                     float(c1.Sales.mean()), N, ""),
                    ("Lines averaged", '=COUNTIFS(tSales25[Segment],"Corporate",tSales25[Store_ID],"ST01")', len(c1), "int", "Check the count.")],
       [("n", "What is the average Technology line in North-Central?", '=AVERAGEIFS(tSales25[Sales],tSales25[Category],"Technology",tSales25[Zone],"North-Central")',
         float(tn.Sales.mean()), N, "Average Technology line, North-Central")])
    fn(m, "2.8.6", [("Numeric Quantity cells", "=COUNT(tSales25[Quantity])", len(S), "int", ""),
                    ("Dirty file: numeric Sales", "=COUNT(tDirty25[Sales])", len(num), "int", "Text Sales not counted.")],
       [("n", "How many Discount cells in tDirty25 hold numbers?", "=COUNT(tDirty25[Discount])", sum(1 for v in DR.Discount if isinstance(v, float)),
         "int", "Numeric Discount cells")])
    fn(m, "2.8.7", [("Sales lines", "=COUNTA(tSales25[Order_ID])", len(S), "int", ""),
                    ("Distinct orders", "=COUNTA(UNIQUE(tSales25[Order_ID]))", int(S.Order_ID.nunique()), "int", "")],
       [("n", "How many Customer_Name cells in tDirty25 are not empty? ('N/A' counts.)", "=COUNTA(tDirty25[Customer_Name])",
         sum(1 for v in DR.Customer_Name if v not in (None, "")), "int", "Non-empty names")])
    fn(m, "2.8.8", [("Empty Quantity cells", "=COUNTBLANK(tDirty25[Quantity])", sum(1 for v in DR.Quantity if v is None), "int", ""),
                    ("Blank spacer rows", "=COUNTBLANK(tDirty25[Order_ID])", DF["blank_rows"], "int", ""),
                    ("Real lines missing Quantity", "=COUNTBLANK(tDirty25[Quantity])-COUNTBLANK(tDirty25[Order_ID])", DF["qty_blank"], "int", "")],
       [("n", "How many Store_Name cells in tDirty25 are blank?", "=COUNTBLANK(tDirty25[Store_Name])", sum(1 for v in DR.Store_Name if not v),
         "int", "Blank Store_Name")])
    loss = int((S.Profit < 0).sum())
    fn(m, "2.8.9", [("Loss-making lines", '=COUNTIF(tSales25[Profit],"<0")', loss, "int", ""),
                    ("Share of lines", '=COUNTIF(tSales25[Profit],"<0")/ROWS(tSales25)', loss / len(S), "pct", "")],
       [("n", "How many lines shipped Same Day?", '=COUNTIF(tSales25[Ship_Mode],"Same Day")', int((S.Ship_Mode == "Same Day").sum()), "int",
         "Same Day lines"),
        ("n", "How many lines mention 'chair' anywhere in Product_Name? (Wildcards.)", '=COUNTIF(tSales25[Product_Name],"*chair*")',
         int(S.Product_Name.str.lower().str.contains("chair").sum()), "int", "Chair lines")])
    fl = S[(S.Category == "Furniture") & (S.Profit < 0)]
    tsw = S[(S.Category == "Technology") & (S.Zone == "South-West") & (S.Discount == 0)]
    fn(m, "2.8.10", [("Loss-making Furniture lines", '=COUNTIFS(tSales25[Category],"Furniture",tSales25[Profit],"<0")', len(fl), "int", ""),
                     ("… at 30% discount or more", '=COUNTIFS(tSales25[Category],"Furniture",tSales25[Profit],"<0",tSales25[Discount],">=0.3")',
                      int((fl.Discount >= 0.3).sum()), "int", "")],
       [("n", "How many South-West Technology lines had no discount?", '=COUNTIFS(tSales25[Category],"Technology",tSales25[Zone],"South-West",tSales25[Discount],0)',
         len(tsw), "int", "SW Technology, no discount")])
    fn(m, "2.8.11", [("Whole year (no filter on Data_Sales25)", "=SUBTOTAL(109,tSales25[Sales])", TOT, N, "Changes when you filter Data_Sales25."),
                     ("SUM ignores filters", "=SUM(tSales25[Sales])", TOT, N, "")],
       [("n", "Filter Data_Sales25 to Zone = South-South. What does =SUBTOTAL(109,tSales25[Sales]) show? (Clear the filter afterwards.)",
         '=SUMIFS(tSales25[Sales],tSales25[Zone],"South-South")', zsum("South-South"), N, "SUBTOTAL with South-South filter")],
       steps=["Type =SUBTOTAL(109,tSales25[Sales]) in an empty cell here.", "Go to Data_Sales25 and filter Zone to one zone.",
              "Come back: the SUBTOTAL shows the visible rows only; SUM still shows the year.", "Clear the filter."])
    se = S[S.Zone == "South-East"].Sales.nlargest(1).iloc[0]
    nc2 = S[S.Zone == "North-Central"].Sales.nlargest(2).iloc[1]
    fn(m, "2.8.12", [("Largest South-East line", '=AGGREGATE(14,6,tSales25[Sales]/(tSales25[Zone]="South-East"),1)', int(se), N,
                      "Other zones divide by 0; option 6 skips those errors.")],
       [("n", "What is the 2nd largest North-Central line?", '=AGGREGATE(14,6,tSales25[Sales]/(tSales25[Zone]="North-Central"),2)', int(nc2), N,
         "2nd largest North-Central line")])
    fn(m, "2.8.13", [("Middle line", "=MEDIAN(tSales25[Sales])", float(S.Sales.median()), N, ""),
                     ("Mean line", "=AVERAGE(tSales25[Sales])", float(S.Sales.mean()), N, "Mean far above median: skewed.")],
       [("n", "What is the median profit per line?", "=MEDIAN(tSales25[Profit])", float(S.Profit.median()), N, "Median profit")])
    mode = S.Quantity.value_counts()
    los = A.LOS_Days.value_counts()
    fn(m, "2.8.14", [("Most common quantity", "=MODE.SNGL(tSales25[Quantity])", int(mode.index[0]), "int", ""),
                     ("Lines with that quantity", f"=COUNTIF(tSales25[Quantity],{int(mode.index[0])})", int(mode.iloc[0]), "int", "")],
       [("n", "What is the most common length of stay in tAdm23 (days)?", "=MODE.SNGL(tAdm23[LOS_Days])", int(los.index[0]), "int",
         "Most common length of stay")])
    mx = S.loc[S.Sales.idxmax()]
    fn(m, "2.8.15", [("Largest line", "=MAX(tSales25[Sales])", int(mx.Sales), N, f"{mx.Product_Name}, {mx.Store_Name}."),
                     ("Latest order date", "=MAX(tSales25[Order_Date])", serial(S.Order_Date.max()), "date", "")],
       [("n", "What is the largest profit on a single line?", "=MAX(tSales25[Profit])", int(S.Profit.max()), N, "Largest profit")])
    fn(m, "2.8.16", [("Worst single loss", "=MIN(tSales25[Profit])", int(S.Profit.min()), N, ""),
                     ("First order date", "=MIN(tSales25[Order_Date])", serial(S.Order_Date.min()), "date", "")],
       [("n", "What is the smallest Sales value on any line?", "=MIN(tSales25[Sales])", int(S.Sales.min()), N, "Smallest line")])
    lg = S.Sales.nlargest(5).tolist()
    fn(m, "2.8.17", [(f"Rank {k}", f"=LARGE(tSales25[Sales],{k})", int(v), N, "") for k, v in enumerate(lg[:3], 1)],
       [("n", "What is the 5th largest sales line?", "=LARGE(tSales25[Sales],5)", int(lg[4]), N, "5th largest line")])
    sm_ = S.Profit.nsmallest(5).tolist()
    fn(m, "2.8.18", [(f"Rank {k} (worst)", f"=SMALL(tSales25[Profit],{k})", int(v), N, "") for k, v in enumerate(sm_[:3], 1)],
       [("n", "What is the 5th worst profit?", "=SMALL(tSales25[Profit],5)", int(sm_[4]), N, "5th smallest profit")])
    avg = float(S.Sales.mean())
    ap = float(S.Profit.mean())
    fn(m, "2.8.19", [("Unrounded", "=AVERAGE(tSales25[Sales])", avg, "#,##0.0000", ""),
                     ("Whole naira", "=ROUND(AVERAGE(tSales25[Sales]),0)", round(avg), N, ""),
                     ("Nearest ₦1,000", "=ROUND(AVERAGE(tSales25[Sales]),-3)", round(avg, -3), N, "")],
       [("n", "Round the average profit per line to the nearest ₦100.", "=ROUND(AVERAGE(tSales25[Profit]),-2)", round(ap, -2), N,
         "Average profit to ₦100")])
    tu = int(S[S.Category == "Technology"].Quantity.sum())
    fu = int(S[S.Category == "Furniture"].Quantity.sum())
    fn(m, "2.8.20", [("Technology units", '=SUMIFS(tSales25[Quantity],tSales25[Category],"Technology")', tu, "int", ""),
                     ("Cartons of 12", '=ROUNDUP(SUMIFS(tSales25[Quantity],tSales25[Category],"Technology")/12,0)', int(np.ceil(tu / 12)), "int",
                      f"{tu / 12:.2f} rounded up.")],
       [("n", "Furniture ships in packs of 6. How many packs cover all FY2025 Furniture units?",
         '=ROUNDUP(SUMIFS(tSales25[Quantity],tSales25[Category],"Furniture")/6,0)', int(np.ceil(fu / 6)), "int", "Furniture packs of 6")])
    mg = float(S.Profit.sum()) / float(S.Sales.sum())
    aq = float(S.Quantity.mean())
    fn(m, "2.8.21", [("Exact margin", "=SUM(tSales25[Profit])/SUM(tSales25[Sales])", mg, "0.00000", ""),
                     ("Never rounds up", "=ROUNDDOWN(SUM(tSales25[Profit])/SUM(tSales25[Sales]),3)", np.floor(mg * 1000) / 1000, "0.000", ""),
                     ("Whole units only", "=ROUNDDOWN(19.99,0)", 19, "int", "")],
       [("n", "Round the average Quantity per line DOWN to 1 decimal place.", "=ROUNDDOWN(AVERAGE(tSales25[Quantity]),1)",
         np.floor(aq * 10) / 10, "dec1", "Average quantity rounded down")])
    ls = int(S[S.Profit < 0].Profit.sum())
    fn(m, "2.8.22", [("Worst loss", "=MIN(tSales25[Profit])", int(S.Profit.min()), N, ""),
                     ("As a positive amount", "=ABS(MIN(tSales25[Profit]))", abs(int(S.Profit.min())), N, "")],
       [("n", "Show the total of all losses as a positive amount.", '=ABS(SUMIF(tSales25[Profit],"<0"))', abs(ls), N, "Losses as a positive amount")])

    tgt = int(ST.Target_2025.sum())
    s12, t12, g12 = _att(S, ST, "ST12")
    q12 = _sid("ST12")
    fn(m, "2.8.23", [("Actual", "=SUM(tSales25[Sales])", TOT, N, ""), ("Target", "=SUM(tStores[Target_2025])", tgt, N, ""),
                     ("The verdict", '=IF(SUM(tSales25[Sales])>=SUM(tStores[Target_2025]),"Met","Missed")', "Met" if TOT >= tgt else "Missed", None, "")],
       [("t", "Did Owerri - New Owerri (ST12) meet its target? Answer Met or Missed.", "Met" if s12 >= t12 else "Missed",
         f'=IF({q12[0]}>={q12[1]},"Met","Missed")', "ST12 verdict")])
    band = lambda v: "Strong" if v >= 0.15 else ("Healthy" if v >= 0.08 else "Weak")  # noqa: E731
    w = m  # IFS uses a margin cell, as the slide does with B2
    fn(w, "2.8.24", [("Chain margin", "=SUM(tSales25[Profit])/SUM(tSales25[Sales])", mg, "pct", "The slide calls this cell B2."),
                     ("Band", '=IFS(SUM(tSales25[Profit])/SUM(tSales25[Sales])>=0.15,"Strong",SUM(tSales25[Profit])/SUM(tSales25[Sales])>=0.08,"Healthy",TRUE,"Weak")',
                      band(mg), None, "First TRUE test wins.")],
       [("t", "Which band does Owerri (ST12) fall in? (Strong / Healthy / Weak)", band(g12),
         f'=IFS({q12[2]}>=0.15,"Strong",{q12[2]}>=0.08,"Healthy",TRUE,"Weak")', "ST12 margin band")])
    s1, t1, g1 = _att(S, ST, "ST01")
    q1 = _sid("ST01")
    s2, t2, g2 = _att(S, ST, "ST02")
    q2 = _sid("ST02")
    fn(m, "2.8.25", [("Ikeja sales >= target", f"={q1[0]}>={q1[1]}", s1 >= t1, None, f"{s1 / t1:.1%} of target."),
                     ("Ikeja margin > 10%", f"={q1[2]}>0.1", g1 > 0.1, None, f"Margin {g1:.1%}."),
                     ("Both", f"=AND({q1[0]}>={q1[1]},{q1[2]}>0.1)", bool(s1 >= t1 and g1 > 0.1), None, "")],
       [("t", "Lagos - Lekki (ST02): is attainment >= 100% AND margin > 10%? Type TRUE or FALSE.", str(bool(s2 >= t2 and g2 > 0.1)).upper(),
         f"=AND({q2[0]}>={q2[1]},{q2[2]}>0.1)", "ST02 AND test")])
    s9, t9, g9 = _att(S, ST, "ST09")
    q9 = _sid("ST09")
    fn(m, "2.8.26", [("Owerri attainment < 85%", f"={q12[0]}/{q12[1]}<0.85", s12 / t12 < 0.85, None, f"{s12 / t12:.1%}."),
                     ("Owerri margin < 5%", f"={q12[2]}<0.05", g12 < 0.05, None, f"{g12:.1%}."),
                     ("Either", f"=OR({q12[0]}/{q12[1]}<0.85,{q12[2]}<0.05)", bool(s12 / t12 < 0.85 or g12 < 0.05), None, "Review the store.")],
       [("t", "Jos - Rayfield (ST09): attainment < 85% OR margin < 5%? Type TRUE or FALSE.", str(bool(s9 / t9 < 0.85 or g9 < 0.05)).upper(),
         f"=OR({q9[0]}/{q9[1]}<0.85,{q9[2]}<0.05)", "ST09 OR test")])
    fn(m, "2.8.27", [("'No losses'", '=COUNTIF(tSales25[Profit],"<0")=0', loss == 0, None, f"{loss} losses."),
                     ("Losses exist", '=NOT(COUNTIF(tSales25[Profit],"<0")=0)', loss != 0, None, "")],
       [("t", "=NOT(COUNTIF(tSales25[Quantity],\"<1\")=0): TRUE or FALSE?", str(bool((S.Quantity < 1).sum() != 0)).upper(),
         '=NOT(COUNTIF(tSales25[Quantity],"<1")=0)', "NOT test on Quantity")])
    q4 = _sid("ST04")
    s4, t4, g4 = _att(S, ST, "ST04")
    fn(m, "2.8.28", [("ST13 margin, protected", f'=IFERROR({_sid("ST13")[2]},"No sales")', "No sales", None,
                      "Without IFERROR: #DIV/0! (0 ÷ 0, no such store).")],
       [("n", "Use the same IFERROR formula for ST04. What margin does it return?", f'=IFERROR({q4[2]},"No sales")', g4, "pct", "ST04 margin")])
    sti = ST.set_index("Store_ID")
    fn(m, "2.8.29", [("ST01 manager", '=IFNA(VLOOKUP("ST01",tStores[[Store_ID]:[Store_Manager]],8,FALSE),"Unknown store")', sti.loc["ST01", "Store_Manager"],
                      None, ""),
                     ("ST13", '=IFNA(VLOOKUP("ST13",tStores[[Store_ID]:[Store_Manager]],8,FALSE),"Unknown store")', "Unknown store", None, "Not found.")],
       [("t", "Who manages ST07? Use VLOOKUP inside IFNA.", sti.loc["ST07", "Store_Manager"],
         '=IFNA(VLOOKUP("ST07",tStores[[Store_ID]:[Store_Manager]],8,FALSE),"Unknown store")', "ST07 manager")])
    sw = {"FUR": "Furniture", "OFF": "Office Supplies", "TEC": "Technology"}
    p1, p3 = S.Product_ID.iloc[0], S.Product_ID.iloc[2]
    sfx = lambda i: (f'=SWITCH(LEFT(INDEX(tSales25[Product_ID],{i}),3),"FUR","Furniture","OFF","Office Supplies","TEC","Technology",'  # noqa: E731
                     f'"Unknown")')
    fn(m, "2.8.30", [(f"Line 1 ({p1})", sfx(1), sw[p1[:3]], None, "")],
       [("t", f"What does the same SWITCH return for line 3 ({p3})?", sw[p3[:3]], sfx(3), "SWITCH on line 3")])
    fn(m, "2.8.31", [("Formatted total", '=TEXT(SUM(tSales25[Sales]),"₦#,##0")', f"₦{TOT:,}", None, ""),
                     ("Live headline", '="FY2025 sales: "&TEXT(SUM(tSales25[Sales]),"₦#,##0")', f"FY2025 sales: ₦{TOT:,}", None, ""),
                     ("Date as text", '=TEXT(DATE(2025,12,31),"dd mmm yyyy")', "31 Dec 2025", None, "")],
       [("t", "On which weekday was the last FY2025 order? Use TEXT with \"dddd\".", S.Order_Date.max().day_name(),
         '=TEXT(MAX(tSales25[Order_Date]),"dddd")', "Weekday of last order")])
    fn(m, "2.8.32", [("Store label (row 1 of tStores)", '=CONCAT(INDEX(tStores[Store_Name],1)," (",INDEX(tStores[Store_ID],1),")")',
                      f"{ST.Store_Name[0]} ({ST.Store_ID[0]})", None, ""),
                     ("Same with &", '=INDEX(tStores[Store_Name],1)&" ("&INDEX(tStores[Store_ID],1)&")"', f"{ST.Store_Name[0]} ({ST.Store_ID[0]})", None, "")],
       [("t", "Build the same label for row 12 of tStores.", f"{ST.Store_Name[11]} ({ST.Store_ID[11]})",
         '=CONCAT(INDEX(tStores[Store_Name],12)," (",INDEX(tStores[Store_ID],12),")")', "Label for row 12")])
    fn(m, "2.8.33", [("South-East stores", '=TEXTJOIN(", ",TRUE,FILTER(tStores[Store_Name],tStores[Zone]="South-East"))',
                      ", ".join(ST[ST.Zone == "South-East"].Store_Name), None, "")],
       [("t", "List the North-Central stores in one cell, separated by \", \".", ", ".join(ST[ST.Zone == "North-Central"].Store_Name),
         '=TEXTJOIN(", ",TRUE,FILTER(tStores[Store_Name],tStores[Zone]="North-Central"))', "North-Central stores")])
    m.finish()


# =====================================================================================================
def m28b(b, D):
    S, ST, DR, A = D.S, D.ST, D.DR, D.A
    m = ModuleSheet(b, "2.8B Functions 34-65", "2.8  The 65 Analytical Excel Functions - Part B (34-65)",
                    "Tutor walk-through after Module 2.8B: cutting and cleaning text, dates and working days, lookups, dynamic arrays and LET",
                    "2.8.34 - 2.8.65")
    m.dataset([("tSales25", "Sheet Data_Sales25 (starts in A1)", "IDs, names and dates of the clean FY2025 lines."),
               ("tStores", "Sheet Data_Stores (starts in A1)", "The lookup table for 2.8.57 - 2.8.64."),
               ("tDirty25", "Sheet Data_Dirty25 (starts in A1)", "Raw text to clean (G31 holds a name with CHAR(160) and a line feed)."),
               ("tAdm23", "Sheet Data_Adm23 (starts in A1)", "Admission and discharge dates for the date functions.")])
    O1, P1, P2 = S.Order_ID.iloc[0], S.Product_ID.iloc[0], S.Product_ID.iloc[1]
    o = lambda i: f"INDEX(tSales25[Order_ID],{i})"  # noqa: E731
    p = lambda i: f"INDEX(tSales25[Product_ID],{i})"  # noqa: E731
    fn(m, "2.8.34", [(f"Category code of line 1 ({P1})", f"=LEFT({p(1)},3)", P1[:3], None, "")],
       [("t", "What are the first 2 characters of every Order_ID? (Line 1.)", O1[:2], f"=LEFT({o(1)},2)", "LEFT 2 of Order_ID")])
    fn(m, "2.8.35", [(f"Order number of line 1 ({O1})", f"=RIGHT({o(1)},6)", O1[-6:], None, "Text."),
                     ("As a number", f"=VALUE(RIGHT({o(1)},6))", int(O1[-6:]), "0", "")],
       [("t", "What are the last 8 characters of line 1's Product_ID?", P1[-8:], f"=RIGHT({p(1)},8)", "RIGHT 8 of Product_ID")])
    fn(m, "2.8.36", [("Year inside the Order_ID", f"=MID({o(1)},4,4)", O1[3:7], None, "S=1, M=2, dash=3, year starts at 4."),
                     ("Sub-category code", f"=MID({p(1)},5,2)", P1[4:6], None, "")],
       [("t", f"What is the sub-category code of line 2 ({P2})?", P2[4:6], f"=MID({p(2)},5,2)", "MID of line 2")])
    zsp = sum(1 for v in DR.Zone if isinstance(v, str) and v != " ".join(v.split()))
    fn(m, "2.8.37", [("Length of a Product_ID", f"=LEN({p(1)})", len(P1), "int", ""),
                     ("Zones with extra spaces", "=SUMPRODUCT(--(LEN(tDirty25[Zone])<>LEN(TRIM(tDirty25[Zone]))))", zsp, "int", "")],
       [("n", "How many characters has the longest store name?", "=MAX(LEN(tStores[Store_Name]))", int(ST.Store_Name.str.len().max()), "int",
         "Longest store name")])
    d2 = P1.find("-", P1.find("-") + 1) + 1
    fn(m, "2.8.38", [("First dash in line 1's Product_ID", f'=FIND("-",{p(1)})', P1.find("-") + 1, "int", ""),
                     ("Second dash", f'=FIND("-",{p(1)},FIND("-",{p(1)})+1)', d2, "int", "Start searching after the first."),
                     ("Case-sensitive: lower-case 'chair' lines", '=SUMPRODUCT(--ISNUMBER(FIND("chair",tSales25[Product_Name])))',
                      int(S.Product_Name.str.contains("chair").sum()), "int", "")],
       [("n", f"Where is the second dash in line 2's Product_ID ({P2})?", f'=FIND("-",{p(2)},FIND("-",{p(2)})+1)',
         P2.find("-", P2.find("-") + 1) + 1, "int", "Second dash, line 2")])
    fn(m, "2.8.39", [("Lines mentioning chair (any case)", '=SUMPRODUCT(--ISNUMBER(SEARCH("chair",tSales25[Product_Name])))',
                      int(S.Product_Name.str.lower().str.contains("chair").sum()), "int", "SEARCH ignores case.")],
       [("n", "How many lines mention 'binder' in any case?", '=SUMPRODUCT(--ISNUMBER(SEARCH("binder",tSales25[Product_Name])))',
         int(S.Product_Name.str.lower().str.contains("binder").sum()), "int", "Binder lines")])
    g31 = DR.Customer_Name[29]
    fn(m, "2.8.40", [("R23 without ₦", '=SUBSTITUTE(Data_Dirty25!R23,"₦","")', DR.Sales[21].replace("₦", ""), None, ""),
                     ("G31: CHAR(160) to a space, then LEN", '=LEN(SUBSTITUTE(Data_Dirty25!G31,CHAR(160)," "))', len(g31), "int",
                      "Same length: one character swapped; the line feed stays."),
                     ("Names with CHAR(160)", "=SUMPRODUCT(--ISNUMBER(SEARCH(CHAR(160),tDirty25[Customer_Name])))", DF["nbsp"], "int", "")],
       [("t", "Remove the commas from R23 as well. What text remains?", DR.Sales[21].replace("₦", "").replace(",", ""),
         '=SUBSTITUTE(SUBSTITUTE(Data_Dirty25!R23,"₦",""),",","")', "R23 without ₦ and commas")])
    fn(m, "2.8.41", [("Masked phone", '=REPLACE("07083197857",5,4,"****")', "0708****857", None, ""),
                     ("Year swapped in line 1's Order_ID", f'=REPLACE({o(1)},4,4,"2026")', O1[:3] + "2026" + O1[7:], None, "")],
       [("t", "Replace the first 3 characters of line 1's Product_ID with 'XXX'.", "XXX" + P1[3:], f'=REPLACE({p(1)},1,3,"XXX")', "REPLACE on line 1")])
    fn(m, "2.8.42", [("L47 trimmed", "=TRIM(Data_Dirty25!L47)", DR.Zone[45].strip(), None, ""),
                     ("Distinct zones after TRIM", '=ROWS(UNIQUE(TRIM(FILTER(tDirty25[Zone],tDirty25[Zone]<>""))))', 4, "int", "12 before.")],
       [("n", "How many characters does TRIM remove from L47?", "=LEN(Data_Dirty25!L47)-LEN(TRIM(Data_Dirty25!L47))",
         len(DR.Zone[45]) - len(" ".join(DR.Zone[45].split())), "int", "Characters TRIM removes from L47")])
    fn(m, "2.8.43", [("LEN of G31", "=LEN(Data_Dirty25!G31)", len(g31), "int", "13 characters + line feed."),
                     ("LEN after CLEAN", "=LEN(CLEAN(Data_Dirty25!G31))", len(g31) - 1, "int", ""),
                     ("Names with a line feed", "=SUMPRODUCT(--ISNUMBER(SEARCH(CHAR(10),tDirty25[Customer_Name])))", DF["lf"], "int", "")],
       [("t", "Fully clean G31 (SUBSTITUTE, CLEAN, TRIM). What is the result?", "Omolara Yusuf",
         '=TRIM(CLEAN(SUBSTITUTE(Data_Dirty25!G31,CHAR(160)," ")))', "G31 fully cleaned")])
    fn(m, "2.8.44", [("Fixed case", '=PROPER("lagos - ikeja")', "Lagos - Ikeja", None, ""),
                     ("Careful with codes", '=PROPER("port harcourt - gra")', "Port Harcourt - Gra", None, "PROPER cannot know GRA is an abbreviation.")],
       [("t", "What does =PROPER(\"ENUGU - INDEPENDENCE LAYOUT\") return?", "Enugu - Independence Layout",
         '=PROPER("ENUGU - INDEPENDENCE LAYOUT")', "PROPER of an upper-case name")])
    fn(m, "2.8.45", [("Case and spaces", '=UPPER(TRIM("office supplies "))', "OFFICE SUPPLIES", None, ""),
                     ("O118 upper-cased", "=UPPER(Data_Dirty25!O118)", DR.Category[116].upper(), None, "Misspellings remain.")],
       [("t", "What does =UPPER(INDEX(tStores[Store_ID],5)&\"-\"&INDEX(tStores[City],5)) return?", f"{ST.Store_ID[4]}-{ST.City[4]}".upper(),
         '=UPPER(INDEX(tStores[Store_ID],5)&"-"&INDEX(tStores[City],5))', "UPPER label")])
    m1 = ST.Store_Manager[0]
    fn(m, "2.8.46", [("Login ID from a name", '=LOWER(SUBSTITUTE(INDEX(tStores[Store_Manager],1)," ","."))', m1.lower().replace(" ", "."), None, ""),
                     ("Case-free comparison", '=LOWER("LAGOS - IKEJA")=LOWER("lagos - ikeja")', True, None, "")],
       [("t", "Build the login ID for the manager in row 4 of tStores.", ST.Store_Manager[3].lower().replace(" ", "."),
         '=LOWER(SUBSTITUTE(INDEX(tStores[Store_Manager],4)," ","."))', "Login ID, row 4")])
    tsv = sum(float(v.replace("₦", "").replace(",", "")) for v in DR.Sales if isinstance(v, str))
    fn(m, "2.8.47", [("R23 as a number", '=VALUE(SUBSTITUTE(SUBSTITUTE(Data_Dirty25!R23,"₦",""),",",""))', 239240, N, ""),
                     ("Text Sales values", "=SUMPRODUCT(--ISTEXT(tDirty25[Sales]))", DF["sales_text"], "int", ""),
                     ("Their total once converted", '=SUMPRODUCT(ISTEXT(tDirty25[Sales])*IFERROR(--SUBSTITUTE(SUBSTITUTE(tDirty25[Sales],"₦",""),",",""),0))',
                      tsv, N, "Hidden from SUM before.")],
       [("n", "Convert T27 ('20%') to a number with VALUE.", '=VALUE(SUBSTITUTE(Data_Dirty25!T27,"%",""))/100', 0.2, "pct0", "T27 as a number")])
    k2 = DR.Store_Location[0]
    fn(m, "2.8.48", [("City from K2", '=TEXTBEFORE(Data_Dirty25!K2,", ")', k2.split(", ")[0], None, ""),
                     ("Category code", f'=TEXTBEFORE({p(1)},"-")', P1.split("-")[0], None, "")],
       [("t", "What is the text before the first space in line 1's Customer_Name?", S.Customer_Name.iloc[0].split(" ")[0],
         '=TEXTBEFORE(INDEX(tSales25[Customer_Name],1)," ")', "First name, line 1")])
    fn(m, "2.8.49", [("State from K2", '=TEXTAFTER(Data_Dirty25!K2,", ")', k2.split(", ")[1], None, ""),
                     ("Item number", f'=TEXTAFTER({p(1)},"-",-1)', P1.split("-")[-1], None, "-1 = the last delimiter."),
                     ("Surname", '=TEXTAFTER("Omolara Yusuf"," ",-1)', "Yusuf", None, "")],
       [("t", "What is the surname (text after the last space) in line 1's Customer_Name?", S.Customer_Name.iloc[0].split(" ")[-1],
         '=TEXTAFTER(INDEX(tSales25[Customer_Name],1)," ",-1)', "Surname, line 1")])
    days = (pd.Timestamp(TODAY) - pd.Timestamp("2025-12-31")).days
    fn(m, "2.8.50", [("Today is in 2025 or later", "=YEAR(TODAY())>=2025", True, None, "TODAY() itself changes daily."),
                     ("Days since FY2025 year end", "=TODAY()-DATE(2025,12,31)", days, "int", f"{days} on {TODAY:%d %b %Y}; more tomorrow.")],
       [("n", "How many days are there from 1 Jan 2025 to 31 Dec 2025 (end minus start)?", "=DATE(2025,12,31)-DATE(2025,1,1)", 364, "int",
         "Days in FY2025 minus 1")])
    fn(m, "2.8.51", [("Date part of NOW equals TODAY", "=INT(NOW())=TODAY()", True, None, "NOW = date + time of day."),
                     ("Time of day as a fraction", "=NOW()-TODAY()<1", True, None, "Always less than one day.")],
       [("t", "Which function returns the date AND the time: TODAY or NOW?", "NOW")],
       note='A dashboard stamp: ="Last refreshed: "&TEXT(NOW(),"dd-mmm-yyyy hh:mm"). It updates whenever the workbook recalculates.')
    fn(m, "2.8.52", [("C24 '14/05/2025' rebuilt", "=DATE(RIGHT(Data_Dirty25!C24,4),MID(Data_Dirty25!C24,4,2),LEFT(Data_Dirty25!C24,2))",
                      serial("2025-05-14"), "date", ""),
                     ("Month 13 rolls over", "=DATE(2025,13,1)", serial("2026-01-01"), "date", ""),
                     ("dd/mm/yyyy text dates to repair", '=SUMPRODUCT(--ISNUMBER(SEARCH("/",tDirty25[Order_Date])))', DF["d_slash"], "int", "")],
       [("n", "What does =DATE(2025,2,29) return? Format it as a date and type the day number of the result.", "=DAY(DATE(2025,2,29))", 1, "int",
         "Day of DATE(2025,2,29)")])
    a0, a1 = A.iloc[0], A.iloc[1]
    los = float((A.Discharge_Date - A.Admit_Date).dt.days.mean())
    dd_ = lambda i: f'DATEDIF(INDEX(tAdm23[Admit_Date],{i}),INDEX(tAdm23[Discharge_Date],{i}),"d")'  # noqa: E731
    fn(m, "2.8.53", [(f"Stay of {a0.Episode_ID}", "=" + dd_(1), (a0.Discharge_Date - a0.Admit_Date).days, "int", ""),
                     ("Average stay, all of tAdm23", "=AVERAGE(tAdm23[Discharge_Date]-tAdm23[Admit_Date])", los, "dec2", "Days.")],
       [("n", f"How long was the stay of {a1.Episode_ID} (row 2 of tAdm23)?", "=" + dd_(2), (a1.Discharge_Date - a1.Admit_Date).days, "int",
         f"Stay of {a1.Episode_ID}")])
    dis = a0.Discharge_Date
    fn(m, "2.8.54", [("End of the discharge month", "=EOMONTH(INDEX(tAdm23[Discharge_Date],1),0)", serial(dis + pd.offsets.MonthEnd(0)), "date", ""),
                     ("End of next month", "=EOMONTH(INDEX(tAdm23[Discharge_Date],1),1)", serial(dis + pd.offsets.MonthEnd(0) + pd.offsets.MonthEnd(1)), "date", ""),
                     ("First of the month", "=EOMONTH(INDEX(tAdm23[Discharge_Date],1),-1)+1", serial(dis - pd.offsets.MonthBegin(1)), "date", "")],
       [("n", "What is the last day of February 2024? Type the day number.", "=DAY(EOMONTH(DATE(2024,2,1),0))", 29, "int", "Last day of Feb 2024")])
    od0, sd0 = S.Order_Date.iloc[0], S.Ship_Date.iloc[0]
    fn(m, "2.8.55", [("Wed 1 Jan to Mon 6 Jan 2025", "=NETWORKDAYS(DATE(2025,1,1),DATE(2025,1,6))", int(np.busday_count("2025-01-01", "2025-01-07")),
                      "int", "Both ends count."),
                     ("Same, 1 Jan as a holiday", "=NETWORKDAYS(DATE(2025,1,1),DATE(2025,1,6),DATE(2025,1,1))",
                      int(np.busday_count("2025-01-01", "2025-01-07", holidays=["2025-01-01"])), "int", "")],
       [("n", "How many working days from line 1's Order_Date to its Ship_Date?",
         "=NETWORKDAYS(INDEX(tSales25[Order_Date],1),INDEX(tSales25[Ship_Date],1))",
         int(np.busday_count(od0.date(), (sd0 + pd.Timedelta(days=1)).date())), "int", "Working days, line 1")])

    def wd(s, n):
        d0 = np.datetime64(s)
        return serial(pd.Timestamp(np.busday_offset(d0, n, roll="backward")))
    fn(m, "2.8.56", [("Follow-up 7 working days after 18 Jan 2023", "=WORKDAY(DATE(2023,1,18),7)", wd("2023-01-18", 7), "date", ""),
                     ("From Saturday 21 Jan 2023", "=WORKDAY(DATE(2023,1,21),7)", wd("2023-01-21", 7), "date", "Counting starts Monday.")],
       [("n", "What date is 10 working days after Friday 19 Dec 2025 (no holidays)?", "=WORKDAY(DATE(2025,12,19),10)", wd("2025-12-19", 10),
         "date", "10 working days after 19 Dec 2025")])
    sti = ST.set_index("Store_ID")
    vl = lambda sid: f'VLOOKUP("{sid}",tStores[[Store_ID]:[Store_Manager]],8,FALSE)'  # noqa: E731
    fn(m, "2.8.57", [("Manager of ST09", "=" + vl("ST09"), sti.loc["ST09", "Store_Manager"], None, ""),
                     ("Manager of ST02", "=" + vl("ST02"), sti.loc["ST02", "Store_Manager"], None, "")],
       [("t", "Who manages ST11?", sti.loc["ST11", "Store_Manager"], "=" + vl("ST11"), "ST11 manager")],
       note="With \"ST13\" VLOOKUP returns #N/A: no such store. Always type FALSE as the last argument.")
    fn(m, "2.8.58", [("Target of ST09", '=XLOOKUP("ST09",tStores[Store_ID],tStores[Target_2025],"not found")', int(sti.loc["ST09", "Target_2025"]), N, ""),
                     ("Attainment of ST09", '=XLOOKUP("ST09",tStores[Store_ID],tStores[Sales_2025])/XLOOKUP("ST09",tStores[Store_ID],tStores[Target_2025])',
                      sti.loc["ST09", "Sales_2025"] / sti.loc["ST09", "Target_2025"], "pct", ""),
                     ("ST13", '=XLOOKUP("ST13",tStores[Store_ID],tStores[Target_2025],"not found")', "not found", None, "")],
       [("n", "What were the FY2025 sales of the store managed by Kemi Oladipo? (XLOOKUP can look left or right.)",
         '=XLOOKUP("Kemi Oladipo",tStores[Store_Manager],tStores[Sales_2025])',
         int(ST.set_index("Store_Manager").loc["Kemi Oladipo", "Sales_2025"]), N, "Sales of Kemi Oladipo's store")])
    fn(m, "2.8.59", [("9th city", "=INDEX(tStores[City],9)", ST.City[8], None, ""),
                     ("INDEX-MATCH on ST09", '=INDEX(tStores[City],MATCH("ST09",tStores[Store_ID],0))', sti.loc["ST09", "City"], None, "")],
       [("t", "Which city is store ST06 in? Use INDEX with MATCH.", sti.loc["ST06", "City"], '=INDEX(tStores[City],MATCH("ST06",tStores[Store_ID],0))',
         "City of ST06")])
    fn(m, "2.8.60", [("Position of ST09", '=MATCH("ST09",tStores[Store_ID],0)', 9, "int", ""),
                     ("ST13 with a message", '=IFNA(MATCH("ST13",tStores[Store_ID],0),"not in list")', "not in list", None, "")],
       [("n", "In which row of tStores is 'Abuja - Garki'?", '=MATCH("Abuja - Garki",tStores[Store_Name],0)',
         int(ST.index[ST.Store_Name == "Abuja - Garki"][0]) + 1, "int", "Position of Abuja - Garki")])

    # dynamic arrays: one real spill each, plus aggregate checks in the worked block
    ss = ST[ST.Zone == "South-South"].sort_values("Sales_2025", ascending=False)
    ssu = ST[ST.Zone == "South-South"]
    fn(m, "2.8.61", [("Stores in South-South", '=ROWS(FILTER(tStores[Store_Name],tStores[Zone]="South-South"))', len(ssu), "int", ""),
                     ("Their sales", '=SUM(FILTER(tStores[Sales_2025],tStores[Zone]="South-South"))', int(ssu.Sales_2025.sum()), N, "")],
       [("n", "How many stores had FY2025 sales above ₦70m? (ROWS of a FILTER.)", "=ROWS(FILTER(tStores[Store_Name],tStores[Sales_2025]>70000000))",
         int((ST.Sales_2025 > 70e6).sum()), "int", "Stores above ₦70m")])
    m.h("The spill: =FILTER(tStores[[Store_Name]:[Store_Name]],tStores[Zone]=\"South-South\") and the Sales_2025 beside it")
    r = m.s.r
    m.ws.write(r, 1, "Store_Name", m.F["hdr"])
    m.ws.write(r, 4, "Sales_2025", m.F["hdr"])
    m.s.spill(r + 1, 1, '=FILTER(tStores[Store_Name],tStores[Zone]="South-South")', [[v] for v in ssu.Store_Name])
    m.s.spill(r + 1, 4, '=FILTER(tStores[Sales_2025],tStores[Zone]="South-South")', [[float(v)] for v in ssu.Sales_2025], ["naira"])
    m.s.r = r + len(ssu) + 3
    cu = S.Customer_ID.value_counts()
    fn(m, "2.8.62", [("Zones", "=ROWS(UNIQUE(tStores[Zone]))", ST.Zone.nunique(), "int", ""),
                     ("Distinct customers", "=ROWS(UNIQUE(tSales25[Customer_ID]))", len(cu), "int", ""),
                     ("Customers seen exactly once", "=ROWS(UNIQUE(tSales25[Customer_ID],,TRUE))", int((cu == 1).sum()), "int", "")],
       [("n", "How many distinct products were sold in FY2025?", "=ROWS(UNIQUE(tSales25[Product_ID]))", int(S.Product_ID.nunique()), "int",
         "Distinct products")])
    rk = ST.sort_values("Sales_2025", ascending=False)
    fn(m, "2.8.63", [("Rank 1", "=INDEX(SORT(tStores[[Store_Name]:[Sales_2025]],9,-1),1,1)", rk.Store_Name.iloc[0], None, ""),
                     ("Rank 12", "=INDEX(SORT(tStores[[Store_Name]:[Sales_2025]],9,-1),12,1)", rk.Store_Name.iloc[-1], None,
                      "Sales_2025 is column 9 counting from Store_Name.")],
       [("t", "Which store ranks 3rd by FY2025 sales?", rk.Store_Name.iloc[2], "=INDEX(SORT(tStores[[Store_Name]:[Sales_2025]],9,-1),3,1)",
         "3rd store by sales")])
    m.h("The spill: =SORT(CHOOSE({1,2},tStores[Store_Name],tStores[Sales_2025]),2,-1)")
    r = m.s.r
    m.ws.write(r, 1, "Store_Name", m.F["hdr"])
    m.ws.write(r, 2, "Sales_2025", m.F["hdr"])
    m.s.spill(r + 1, 1, "=SORT(CHOOSE({1,2},tStores[Store_Name],tStores[Sales_2025]),2,-1)",
              [[n_, float(v)] for n_, v in zip(rk.Store_Name, rk.Sales_2025)], [None, "naira"])
    m.ws.set_column(1, 1, 26)
    m.s.r = r + len(rk) + 3
    att = (ST.Sales_2025 / ST.Target_2025)
    ao = ST.assign(a=att).sort_values("a", ascending=False)
    sb = lambda i: f"=INDEX(SORTBY(tStores[Store_Name],tStores[Sales_2025]/tStores[Target_2025],-1),{i})"  # noqa: E731
    fn(m, "2.8.64", [("1st by attainment", sb(1), ao.Store_Name.iloc[0], None, f"{ao.a.iloc[0]:.1%}."),
                     ("Last by attainment", sb(12), ao.Store_Name.iloc[-1], None, f"{ao.a.iloc[-1]:.1%}.")],
       [("t", "Which store is 2nd by attainment?", ao.Store_Name.iloc[1], sb(2), "2nd by attainment")])
    TOT, TGT = int(S.Sales.sum()), int(ST.Target_2025.sum())
    gap = TOT - TGT
    fn(m, "2.8.65", [("Chain margin", "=LET(s,SUM(tSales25[Sales]),p,SUM(tSales25[Profit]),p/s)", float(S.Profit.sum()) / TOT, "pct", ""),
                     ("Target gap", '=LET(gap,SUM(tSales25[Sales])-SUM(tStores[Target_2025]),IF(gap>=0,"Ahead by ","Behind by ")&TEXT(ABS(gap),"₦#,##0"))',
                      ("Ahead by " if gap >= 0 else "Behind by ") + f"₦{abs(gap):,}", None, "")],
       [("n", "Use LET to work out the share of FY2025 sales made by Technology (t = Technology sales, s = all sales, t/s).",
         '=LET(t,SUMIF(tSales25[Category],"Technology",tSales25[Sales]),s,SUM(tSales25[Sales]),t/s)',
         float(S[S.Category == "Technology"].Sales.sum()) / TOT, "pct", "Technology share with LET")])
    m.finish()


def build(b, D):
    m28a(b, D)
    m28b(b, D)
