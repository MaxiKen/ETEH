"""Shared data and helpers for the Week 4 edition-3 workbook (02_sec_w4.py).

Table names match the slides (w4_data.TABLES): tSal, tOrders, tCV, tLOS, tCensus23, tHotel, tDash25, tMonthly and tStores each sit
on a Data_ sheet starting in A1 (tHotel adds Is_Cancelled, 1 = cancelled, so CORREL and AVERAGE can use it). Every cached value is
worked out here in pandas or taken from the slide modules (w4_m*), so the slide, the worked solution and the answer key always
show the same number; verify_xlsx.py recomputes every formula independently.
"""
from __future__ import annotations

import importlib
from types import SimpleNamespace

import numpy as np
import pandas as pd

import w4_data as W
from w1x import DataSheet

N = "naira"
PARTS_SLIDES = ["w4_m1_2", "w4_m3_4", "w4_m5_6", "w4_m7_10", "w4_m11_14"]
SL = {p: importlib.import_module(p) for p in PARTS_SLIDES}
DEF = {}
for _p in PARTS_SLIDES:
    for _M in SL[_p].MODULES:
        for _t in _M["topics"]:
            DEF[_t["tag"]] = _t


def _obj(df):
    return df.astype(object).where(df.notna(), None)


def data():
    d = SimpleNamespace()
    d.S = W.table("tSal")
    d.O = W.table("tOrders")
    d.O["order_date"] = pd.to_datetime(d.O.order_date)
    d.CV = W.table("tCV")
    d.LOS = W.table("tLOS")
    d.C23 = W.table("tCensus23")
    h = W.table("tHotel")
    h["Is_Cancelled"] = (h.Reservation_Status == "Cancelled").astype(int)
    d.H = h
    d.D25 = W.table("tDash25")
    d.D25["Order_Date"] = pd.to_datetime(d.D25.Order_Date)
    d.MON = W.table("tMonthly")
    d.STO = W.table("tStores")
    d.ROOMS = W.ROOMS
    net = d.O[d.O.order_status.isin(["Delivered", "Completed"])].groupby("customer_id").total_price.sum()
    net = net.reset_index().sort_values(["total_price", "customer_id"], ascending=[False, True])
    cr = pd.DataFrame({"Rank": range(1, len(net) + 1), "customer_id": net.customer_id.values, "Net_Revenue": net.total_price.values})
    cr["Cum_Revenue"] = cr.Net_Revenue.cumsum()
    cr["Cum_Share"] = cr.Cum_Revenue / cr.Net_Revenue.sum()
    d.CR = cr
    return d


def _w(df, widths):
    return {k: v for k, v in widths.items() if k in df.columns}


def data_sheets(b, D):
    specs = [
        ("Data_Sal", "tSal", D.S, "tSal: ETEH Talent Salary Survey, one respondent per row (A1)", f"{len(D.S):,} respondents · 4.1 - 4.4, 4.11 - 4.13",
         {"Annual_Salary_NGN": N, "Age": "int", "Years_of_Experience": "dec1", "Is_Masters": "int", "Is_PhD": "int"},
         {"Respondent_ID": 12, "Education_Level": 15, "Job_Title": 28, "Job_Family": 18, "Seniority": 11, "Annual_Salary_NGN": 15}),
        ("Data_Orders", "tOrders", D.O, "tOrders: SuperMart Online order lines with product category (A1)", f"{len(D.O):,} order lines · 4.1, 4.2",
         {"order_date": "date", "total_price": N, "quantity": "int"}, {"order_id": 19, "order_date": 11, "customer_id": 12, "category": 12,
                                                                      "order_status": 12, "total_price": 13}),
        ("Data_CV", "tCV", D.CV, "tCV: Kingsway University student results, 2024/25 first semester (A1)", f"{len(D.CV):,} students · 4.1 - 4.3",
         {}, {"Matric_No": 17, "Faculty": 16, "Lecture_Attendance_Pct": 12, "Study_Hours_Wk": 10}),
        ("Data_LOS", "tLOS", D.LOS, "tLOS: Mercyland inpatient episodes admitted in 2023 (A1)", f"{len(D.LOS):,} episodes · 4.1, 4.7 - 4.8, 4.13 - 4.14",
         {"Admit_Date": "date", "Bill_Amount_NGN": N, "Age": "int", "LOS_Days": "int"},
         {"Episode_ID": 15, "Admit_Date": 11, "Department": 22, "Admission_Route": 13, "HMO": 18, "Bill_Amount_NGN": 14}),
        ("Data_Census", "tCensus23", D.C23, "tCensus23: Mercyland midnight bed census, 2023 (200 staffed beds) (A1)", "365 days · 4.2, 4.8",
         {"Census_Date": "date", "Occupancy_Pct": "pct"}, {"Census_Date": 12, "Inpatients": 11, "Occupancy_Pct": 12}),
        ("Data_Hotel", "tHotel", D.H, "tHotel: Atlantic Pearl Hotel reservations, Jul 2023 - Dec 2024 (A1)",
         f"{len(D.H):,} reservations · Is_Cancelled: 1 = cancelled, 0 = stayed · 4.3, 4.6, 4.9",
         {"Check_In": "date", "Check_In_Year": "gen", "Room_Rate_NGN": N, "Reservation_Value_NGN": N, "Is_Cancelled": "int"},
         {"Reservation_ID": 12, "Check_In": 11, "Room_Category": 17, "Sales_Channel": 21, "Reservation_Status": 12,
          "Room_Rate_NGN": 13, "Reservation_Value_NGN": 14}),
        ("Data_Dash25", "tDash25", D.D25, "tDash25: SuperMart Stores FY2025 sales lines (A1)", f"{len(D.D25):,} lines · 4.2, 4.7, 4.10, 4.12, 4.14",
         {"Order_Date": "date", "Sales": N, "Profit": N}, {"Order_ID": 16, "Order_Date": 11, "Store_Name": 24, "Zone": 14, "Category": 15}),
        ("Data_Monthly", "tMonthly", D.MON, "tMonthly: SuperMart Stores monthly sales, Jan 2022 - Dec 2025 (A1)",
         "48 months · Sales is column E, rows 2-49 · 4.5", {"Month": "mon", "Sales": N, "Year": "gen", "Month_No": "gen", "t": "gen"},
         {"Month": 10, "Sales": 15}),
        ("Data_CustRev", "tCustRev", D.CR, "tCustRev: SuperMart Online net revenue per customer, largest first (A1)",
         f"{len(D.CR):,} buyers · Delivered + Completed orders · a PivotTable of tOrders pasted as values, sorted · 4.2.4",
         {"Rank": "int", "Net_Revenue": N, "Cum_Revenue": N, "Cum_Share": "pct"}, {"customer_id": 12, "Net_Revenue": 14, "Cum_Revenue": 15}),
        ("Data_Stores", "tStores", D.STO, "tStores: the 12 SuperMart stores with floor area and FY2025 target (A1)", "12 stores · 4.2, 4.7, 4.14",
         {"Annual_Target_2025": N, "Floor_Area_sqm": "int"}, {"Store_Name": 24, "Zone": 14, "Annual_Target_2025": 15}),
    ]
    out = {}
    for sheet, tname, df, title, sub, fmts, widths in specs:
        ds = DataSheet(b, sheet, title, sub)
        out[tname] = ds.put(_obj(df), tname, fmts=fmts, widths=_w(df, widths))
        if tname == "tMonthly":
            # column F, outside the table: one-step simple exponential smoothing (alpha 0.3), the F49 quoted in 4.5.5
            ws = ds.ws
            ws.set_column(5, 5, 16)
            ws.write(0, 5, "SES_Forecast", b.F["hdr"])
            fmt = b.nf(N)
            v = ses_column(df.Sales)
            ws.write_formula(1, 5, "=E2", fmt, v[0])
            for i in range(1, len(df)):
                ws.write_formula(i + 1, 5, f"=0.3*E{i + 1}+0.7*F{i + 1}", fmt, v[i])
            ws.write(len(df) + 2, 5, "F = 0.3 x last month's sales + 0.7 x last month's forecast (4.5.5)", b.F["note"])
    return out


def ses_column(sales, alpha=0.3):
    """One-step-ahead simple exponential smoothing: F2 = E2, F(r) = alpha*E(r-1) + (1-alpha)*F(r-1)."""
    s = list(map(float, sales))
    out = [s[0]]
    for i in range(1, len(s)):
        out.append(alpha * s[i - 1] + (1 - alpha) * out[-1])
    return out


# ---------------------------------------------------------------- one topic
def tp(m, tag, worked=None, ex=(), steps=None, code=None, code_label=None, note=None, do=None, what=None, extra=None):
    """Band + definition (slide copy) + Steps + optional code + Worked solution + Your turn (+ do-it checklist)."""
    t = DEF[tag]
    m.topic(tag, t["title"], what or (t["defn"] + "  " + t.get("simple", "")))
    st = t.get("steps") if steps is True else steps
    if st:
        m.steps(st)
    if code is True:
        code, code_label = t["code"], code_label or t.get("code_label", "Code")
    if code:
        m.h(code_label or "Code")
        m.s.code(code.strip("\n"))
        m.s.r += 1
    if extra:
        extra(m)
    if callable(worked):
        worked = worked(m)
    if worked:
        m.worked(worked)
    if note:
        m.note(note)
    if ex:
        m.turn()
        for e in ex:
            if e[0] == "n":
                _, task, fx, val, kind = e[:5]
                m.ex(task, fx, val, kind, short=e[5] if len(e) > 5 else None, tol=e[6] if len(e) > 6 else None)
            elif e[0] == "t":
                _, task, ans = e[:3]
                fx = e[3] if len(e) > 3 else None
                m.ex_text(task, ans, formula=fx, short=e[4] if len(e) > 4 else None, alts=e[5] if len(e) > 5 else ())
            else:
                m.ex_open(e[1], e[2])
    if do:
        m.do(do)


def f(x):
    """plain Python number for cached values"""
    if isinstance(x, (np.integer,)):
        return int(x)
    if isinstance(x, (np.floating,)):
        return float(x)
    return x


def lit(rows):
    """Small typed grid helper: list of lists -> grid rows (strings)."""
    return [[str(c) for c in r] for r in rows]
