"""Shared data and helpers for the Week 3 edition-3 workbook (02_sec_w3.py).

Table names match the slides (w3_data.TABLES): tTransactions, tProducts, tCustomers, tSellers, tDeliveries, tStates, tDimDate,
tDash, tDashStores, tAdm23, tCensus, tCV and tMsDirty each sit on a Data_ sheet starting in A1. One extra table, tLines, is the
database view v_order_lines imported through ODBC (3.3.6): the order lines with category, customer zone and courier already
joined in SQL, so worksheet formulas (SUMIFS, COUNTIFS) can answer the same questions the SQL and DAX on the slides answer.
Every cached value is worked out in pandas; verify_xlsx.py recomputes every formula independently.
"""
from __future__ import annotations

import importlib
from types import SimpleNamespace

import numpy as np
import pandas as pd

import w3_data as W
from w1x import DataSheet

N = "naira"
PARTS_SLIDES = ["w3_m1_2", "w3_m3_5", "w3_m6a", "w3_m6b", "w3_m7", "w3_m8_10"]
DEF = {}
for _p in PARTS_SLIDES:
    for _M in importlib.import_module(_p).MODULES:
        for _t in _M["topics"]:
            DEF[_t["tag"]] = _t
NETS = ("Delivered", "Completed")


def _obj(df):
    return df.astype(object).where(df.notna(), None)


def data():
    d = SimpleNamespace()
    tx = W.table("tTransactions")
    tx["order_date"] = pd.to_datetime(tx.order_date)
    d.TX = tx
    li = W.q("SELECT transaction_id, order_id, order_date, CAST(order_year AS INTEGER) AS order_year, order_month, customer_id, "
             "membership_tier, customer_zone, product_id, category, subcategory, seller_id, quantity, unit_price, total_price, "
             "commission_amount, order_status, payment_method, courier_name, delivery_status, days_to_deliver "
             "FROM v_order_lines ORDER BY transaction_id", "smo")
    li["order_date"] = pd.to_datetime(li.order_date)
    d.L = li
    d.P = W.table("tProducts")[["product_id", "product_name", "category", "subcategory", "brand_name", "seller_id", "price",
                                "discount_price", "commission_percentage", "stock_quantity", "total_sold", "average_rating"]]
    c = W.table("tCustomers")[["customer_id", "first_name", "last_name", "gender", "city", "state", "zone", "membership_tier",
                               "loyalty_points", "total_orders", "total_spent", "created_at"]]
    c["created_at"] = pd.to_datetime(c.created_at)
    d.C = c
    d.SE = W.table("tSellers")[["seller_id", "store_name", "owner_name", "store_city", "store_state", "store_zone", "rating",
                                "total_sales", "total_revenue", "account_status"]]
    dv = W.table("tDeliveries")[["delivery_id", "order_id", "customer_id", "city", "state", "zone", "courier_name",
                                 "delivery_status", "shipped_date", "delivered_date"]]
    for k in ("shipped_date", "delivered_date"):
        dv[k] = pd.to_datetime(dv[k])
    d.DV = dv
    d.ST = W.table("tStates")
    d.DD = W.table("tDimDate")
    d.DS = W.table("tDash")
    d.DSS = W.table("tDashStores")
    d.A = W.table("tAdm23")
    d.CEN = W.table("tCensus")
    d.CV = W.table("tCV")
    from common import W2D
    d.LD = pd.read_csv(W2D / "Mercyland_Episodes_2023_DIRTY.csv", dtype=str, keep_default_na=False, encoding="utf-8-sig")  # as received
    d.NT = li[li.order_status.isin(NETS)]
    return d


def _w(df, widths):
    return {k: v for k, v in widths.items() if k in df.columns}


def data_sheets(b, D):
    out = {}
    specs = [
        ("Data_Transactions", "tTransactions", D.TX, "tTransactions: the transactions table of supermart_online.db, one row per order line (A1)",
         f"{len(D.TX):,} rows · order_date converted to a date in Power Query · used in 3.1 - 3.6",
         {"order_date": "date", "unit_price": N, "total_price": N, "commission_amount": N, "quantity": "int"},
         {"transaction_id": 13, "order_id": 19, "customer_id": 11, "product_id": 11, "seller_id": 10, "delivery_id": 11,
          "order_status": 11, "payment_method": 14, "payment_status": 12, "order_date": 11, "total_price": 12, "unit_price": 12}),
        ("Data_Lines", "tLines", D.L, "tLines: the view v_order_lines imported through ODBC: order lines with category, zone and courier (A1)",
         f"{len(D.L):,} rows, 21 of the view's columns · the joins were done in SQL · used in 3.2 - 3.7",
         {"order_date": "date", "order_year": "gen", "unit_price": N, "total_price": N, "commission_amount": N, "days_to_deliver": "gen"},
         {"transaction_id": 13, "order_id": 19, "order_date": 11, "customer_zone": 14, "category": 12, "subcategory": 14,
          "courier_name": 18, "delivery_status": 13, "order_status": 11, "payment_method": 14, "total_price": 12}),
        ("Data_Products", "tProducts", D.P, "tProducts: the products table (A1)", f"{len(D.P):,} products · 12 of 15 columns",
         {"price": N, "discount_price": N}, {"product_id": 11, "product_name": 30, "category": 12, "subcategory": 14, "brand_name": 14}),
        ("Data_Customers", "tCustomers", D.C, "tCustomers: the customers table (A1)", f"{len(D.C):,} customers · 12 of 16 columns",
         {"total_spent": N, "created_at": "date"}, {"customer_id": 11, "first_name": 12, "last_name": 12, "state": 12, "zone": 14,
                                                   "membership_tier": 13, "total_spent": 13}),
        ("Data_Sellers", "tSellers", D.SE, "tSellers: the sellers table (A1)", f"{len(D.SE):,} sellers · 10 of 14 columns",
         {"total_revenue": N}, {"store_name": 26, "owner_name": 20, "store_zone": 14, "total_revenue": 14}),
        ("Data_Deliveries", "tDeliveries", D.DV, "tDeliveries: the deliveries table (A1)", f"{len(D.DV):,} deliveries · 10 of 13 columns",
         {"shipped_date": "date", "delivered_date": "date"}, {"delivery_id": 11, "order_id": 19, "courier_name": 18, "zone": 14,
                                                             "delivery_status": 13, "shipped_date": 11, "delivered_date": 11}),
        ("Data_States", "tStates", D.ST, "tStates: the nigeria_states lookup (A1)", "32 states and their zones", {}, {"state": 14, "zone": 14}),
        ("Data_DimDate", "tDimDate", D.DD, "tDimDate: one row per day, 2020-01-01 to 2025-12-31 (A1)", f"{len(D.DD):,} days · the date table of 3.5 and 3.6",
         {"Date": "date", "Year": "gen", "Month_No": "gen"}, {"Date": 11, "Year_Month": 11}),
        ("Data_Dash", "tDash", D.DS, "tDash: SuperMart Stores FY2025 sales lines, the dashboard source (A1)", f"{len(D.DS):,} lines · used in 3.6B, 3.7 - 3.9",
         {"Order_Date": "date", "Sales": N, "Profit": N, "Row_ID": "gen"}, {"Order_ID": 16, "Order_Date": 11, "Store_Name": 26, "Zone": 14,
                                                                          "Category": 15}),
        ("Data_DashStores", "tDashStores", D.DSS, "tDashStores: the 12 stores with FY2025 sales, profit and target (A1)", "12 stores",
         {"Sales": N, "Profit": N, "Target": N, "Attainment": "pct"}, {"Store_Name": 26, "Zone": 14, "Sales": 14, "Target": 14}),
        ("Data_Adm23", "tAdm23", D.A, "tAdm23: Mercyland admissions 2023 (A1)", f"{len(D.A):,} admissions · used in 3.6B, 3.9",
         {"Admit_Date": "date", "Discharge_Date": "date", "Bill_Amount_NGN": N, "Age": "gen", "LOS_Days": "gen"},
         {"Episode_ID": 15, "Department": 20, "Admission_Route": 13, "Admit_Date": 12, "Discharge_Date": 12, "Ward": 14,
          "Consultant_Name": 20, "HMO": 16, "Bill_Amount_NGN": 14}),
        ("Data_Census", "tCensus", D.CEN, "tCensus: Mercyland daily bed census (A1)", f"{len(D.CEN):,} days",
         {"Census_Date": "date", "Occupancy_Pct": "pct"}, {"Census_Date": 12}),
        ("Data_CV", "tCV", D.CV, "tCV: Kingsway University student performance 2024/25 (A1)", f"{len(D.CV):,} students",
         {"Level": "gen"}, {"Matric_No": 11, "Faculty": 16, "Department": 22}),
        ("Data_LcDirty", "tMsDirty", D.LD, "tMsDirty: Mercyland_Episodes_2023_DIRTY.csv exactly as received, every column text (A1)",
         f"{len(D.LD):,} rows · the staging data of 3.10 · do not clean here", {k: "text" for k in D.LD.columns},
         {"Patient_Name": 22, "Diagnosis_Group": 16, "Admit_Date": 12, "Discharge_Date": 12, "Consultant_Name": 20, "Ward": 14}),
    ]
    for sheet, tname, df, title, sub, fmts, widths in specs:
        out[tname] = DataSheet(b, sheet, title, sub).put(_obj(df), tname, fmts=fmts, widths=_w(df, widths))
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
        m.h((code_label or "Code").title() if code_label and code_label.isupper() else (code_label or "Code"))
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
                m.ex(task, fx, val, kind, short=e[5] if len(e) > 5 else None)
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


# ---------------------------------------------------------------- formula snippets
def net(col="total_price", t="tLines", **crit):
    """SUMIFS over Delivered + Completed lines with extra criteria {column: value}."""
    extra = "".join(f',{t}[{k}],{_q(v)}' for k, v in crit.items())
    return (f'SUMIFS({t}[{col}],{t}[order_status],"Delivered"{extra})+'
            f'SUMIFS({t}[{col}],{t}[order_status],"Completed"{extra})')


def netn(t="tLines", **crit):
    extra = "".join(f',{t}[{k}],{_q(v)}' for k, v in crit.items())
    return f'COUNTIFS({t}[order_status],"Delivered"{extra})+COUNTIFS({t}[order_status],"Completed"{extra})'


def _q(v):
    if isinstance(v, str) and not v.startswith("&"):
        return '"' + v + '"'
    return str(v)
