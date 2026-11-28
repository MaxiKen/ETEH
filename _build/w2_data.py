"""Week 2 shared example tables (edition 3).

One place defines every Week 2 example table by its Excel table name. The slide builder (08_v3_week2_deck.py)
prints small excerpts of these tables, and the Week 2 workbook builder (02_sec_w2.py) writes the same tables, so
a table name seen on a slide is the table the tutor opens in the workbook.
All values come from the bootcamp datasets and the Week2_Data_Files practice files (_build/01_build_datasets.py).
"""
from __future__ import annotations

import glob
import json
import os
import re
import sys
from functools import lru_cache
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(__file__))
from common import W2D  # noqa: E402
from wk_common import L, lc_adm, sm_sales  # noqa: E402

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONEY = {"Sales", "Profit", "Target_2025", "Sales_2025", "Bill_Amount_NGN", "Revenue", "total_revenue", "Avg_Bill", "Avg_Sale", "Loss",
         "Furniture", "Office Supplies", "Technology", "Grand Total", "Total_Sales", "Total_Profit", "Sales_Total"}
PCT = {"Margin", "Discount", "Share", "Readmit_Rate", "Attainment", "Pct_of_Total"}
W2 = Path(W2D)


# ---------------------------------------------------------------- base data
@lru_cache(None)
def sales():
    d = sm_sales()
    d["Year"] = d.Order_Date.dt.year
    return d


@lru_cache(None)
def s25():
    d = sales()
    return d[d.Year == 2025].sort_values(["Order_Date", "Row_ID"]).reset_index(drop=True)


@lru_cache(None)
def stores():
    return L("sm_stores")


@lru_cache(None)
def adm23():
    a = pd.read_csv(W2 / "Mercyland_Episodes_2023_pipe.txt", sep="|", parse_dates=["Admit_Date", "Discharge_Date"])
    return a


@lru_cache(None)
def dirty_raw():
    """SuperMart_Sales_2025_DIRTY.csv exactly as text (what Excel / Power Query receive)."""
    return pd.read_csv(W2 / "SuperMart_Sales_2025_DIRTY.csv", dtype=str, keep_default_na=False, encoding="utf-8-sig")


@lru_cache(None)
def lc_dirty_raw():
    return pd.read_csv(W2 / "Mercyland_Episodes_2023_DIRTY.csv", dtype=str, keep_default_na=False)


# ---------------------------------------------------------------- named tables
@lru_cache(None)
def table(name: str) -> pd.DataFrame:
    return _BUILD[name]().reset_index(drop=True)


def _files():
    rows = [
        ("SuperMart_Sales_2025_DIRTY.csv", "CSV (comma)", "UTF-8 with BOM", len(dirty_raw()), "Raw FY2025 sales export with deliberate defects"),
        ("Mercyland_Episodes_2023_pipe.txt", "TXT (pipe |)", "UTF-8", len(adm23()), "Clean 2023 hospital admissions"),
        ("Mercyland_Episodes_2023_DIRTY.csv", "CSV (comma)", "UTF-8", len(lc_dirty_raw()), "Raw 2023 admissions from the old EMR"),
        ("SuperMart_Store_Targets_EU_semicolon.csv", "CSV (semicolon)", "Windows-1252", 12, "Store targets in European number format"),
        ("SuperMart_Monthly_Sales_2025 (folder)", "12 CSV files", "UTF-8", sum(r for _, r, _ in _month_counts()), "One file per month of FY2025"),
        ("SuperMart_Online_RAW_Export.xlsx", "Workbook (5 sheets)", "n/a", 11000, "Online orders, sellers, products, deliveries, customers"),
        ("SuperMart_Online_Sellers.json", "JSON", "UTF-8", 200, "Seller directory from an API"),
        ("SuperMart_Store_Directory.html", "Web page (HTML)", "UTF-8", 12, "Store directory published on the intranet"),
        ("SuperMart_Monthly_Sales_Bulletin_2025.pdf", "PDF (1 table)", "n/a", 13, "Store x month sales in ₦ millions"),
    ]
    return pd.DataFrame(rows, columns=["File", "Format", "Encoding", "Rows", "What it holds"])


def _month_counts():
    out = []
    for f in sorted(glob.glob(str(W2 / "SuperMart_Monthly_Sales_2025" / "*.csv"))):
        d = pd.read_csv(f)
        out.append((Path(f).name, len(d), int(d.Sales.sum())))
    return out


def _month_files():
    return pd.DataFrame(_month_counts(), columns=["File", "Rows", "Sales"])


def _sales25():
    cols = ["Row_ID", "Order_ID", "Order_Date", "Ship_Date", "Ship_Mode", "Customer_ID", "Customer_Name", "Segment", "Store_ID",
            "Store_Name", "Zone", "Product_ID", "Category", "Sub_Category", "Product_Name", "Sales", "Quantity", "Discount", "Profit"]
    return s25()[cols]


def _stores():
    st = stores().copy()
    st["Opened"] = pd.to_datetime(st.Opened)
    st = st.rename(columns={"Annual_Target_2025": "Target_2025"})
    st["Sales_2025"] = st.Store_ID.map(s25().groupby("Store_ID").Sales.sum())
    return st[["Store_ID", "Store_Name", "City", "State", "Zone", "Opened", "Floor_Area_sqm", "Store_Manager", "Target_2025",
               "Sales_2025"]]


def _adm23():
    a = adm23()
    return a[["Episode_ID", "Patient_Name", "Age", "Gender", "Department", "Admission_Route", "Admit_Date", "Discharge_Date",
              "LOS_Days", "Ward", "Consultant_Name", "HMO", "Bill_Amount_NGN", "Readmitted_30d"]]


def _dirty25():
    return dirty_raw()


def _defects25():
    lg = L("dirty_sm_log")
    return lg.rename(columns={"What was injected": "Example", "Rows affected": "Rows"})[["Defect", "Example", "Rows", "Expected fix"]]


def _lc_dirty():
    return lc_dirty_raw()


def _lc_defects():
    lg = L("dirty_lc_log")
    return lg.rename(columns={"What was injected": "Example", "Rows affected": "Rows"})[["Defect", "Example", "Rows", "Expected fix"]]


def _bulletin():
    s = s25().copy()
    s["M"] = s.Order_Date.dt.month
    ct = s.pivot_table(index="Store_Name", columns="M", values="Sales", aggfunc="sum", fill_value=0) / 1e6
    ct = ct.reindex(stores().Store_Name)
    out = pd.DataFrame({"Store": ct.index})
    for m in range(1, 13):
        out[MONTHS[m - 1]] = [float(f"{v:.1f}") for v in ct[m].values]
    out["FY_2025"] = [float(f"{v:.1f}") for v in ct.sum(axis=1).values]
    return out


def _web_stores():
    st = stores()
    return pd.DataFrame({"Store ID": st.Store_ID, "Store": st.Store_Name, "City": st.City, "State": st.State, "Zone": st.Zone,
                         "Opened": pd.to_datetime(st.Opened).dt.strftime("%d %b %Y"),
                         "Floor area (sq m)": st.Floor_Area_sqm.map(lambda v: f"{v:,}"), "Manager": st.Store_Manager})


def _zone_cat():
    s = s25()
    p = s.pivot_table(index="Zone", columns="Category", values="Sales", aggfunc="sum").reindex(
        ["South-West", "South-South", "North-Central", "South-East"])
    p["Grand Total"] = p.sum(axis=1)
    p = p.reset_index()
    tot = {"Zone": "Grand Total", **{c: p[c].sum() for c in p.columns[1:]}}
    return pd.concat([p, pd.DataFrame([tot])], ignore_index=True)


def _month_sales():
    s = s25().copy()
    g = s.groupby(s.Order_Date.dt.month).agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"), Orders=("Order_ID", "nunique"),
                                             Lines=("Row_ID", "count")).reset_index(names="M")
    g.insert(0, "Month", [MONTHS[m - 1] for m in g.M])
    g["Margin"] = g.Profit / g.Sales
    return g.drop(columns="M")


def _cat_quarter():
    s = s25().copy()
    p = s.pivot_table(index="Category", columns=s.Order_Date.dt.quarter, values="Sales", aggfunc="sum")
    p.columns = [f"Q{c}" for c in p.columns]
    return p.reset_index()


def _sellers():
    j = json.loads((W2 / "SuperMart_Online_Sellers.json").read_text(encoding="utf-8"))["sellers"]
    d = pd.DataFrame(j)
    return d[["seller_id", "store_name", "store_city", "store_state", "rating", "total_sales", "total_revenue", "account_status"]]


def _dept23():
    a = adm23()
    g = a.groupby("Department").agg(Admissions=("Episode_ID", "count"), Avg_LOS=("LOS_Days", "mean"),
                                    Readmitted=("Readmitted_30d", lambda v: int((v == "Yes").sum())),
                                    Avg_Bill=("Bill_Amount_NGN", "mean")).reset_index()
    g["Readmit_Rate"] = g.Readmitted / g.Admissions
    return g.sort_values("Admissions", ascending=False)


def _store_perf():
    s = s25()
    g = s.groupby(["Store_ID", "Store_Name", "Zone"]).agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"),
                                                         Orders=("Order_ID", "nunique")).reset_index()
    st = stores().set_index("Store_ID").Annual_Target_2025
    g["Target_2025"] = g.Store_ID.map(st)
    g["Margin"] = g.Profit / g.Sales
    g["Attainment"] = g.Sales / g.Target_2025
    return g


_BUILD = {
    "tFiles": _files, "tMonthFiles": _month_files, "tSales25": _sales25, "tStores": _stores, "tAdm23": _adm23,
    "tDirty25": _dirty25, "tDefects25": _defects25, "tLCDirty": _lc_dirty, "tLCDefects": _lc_defects, "tBulletin": _bulletin,
    "tWebStores": _web_stores, "tZoneCat": _zone_cat, "tMonthSales": _month_sales, "tCatQuarter": _cat_quarter,
    "tSellers": _sellers, "tDept23": _dept23, "tStorePerf": _store_perf,
}


# ---------------------------------------------------------------- slide formatting
def fmt(col, v):
    if v is None or (isinstance(v, float) and np.isnan(v)) or v is pd.NaT:
        return ""
    if isinstance(v, str):
        return v
    if isinstance(v, (pd.Timestamp,)) or hasattr(v, "isoformat"):
        return pd.Timestamp(v).strftime("%Y-%m-%d")
    if (col in PCT or col.endswith("_Rate")) and isinstance(v, (float, np.floating)) and abs(v) <= 1.5:
        return f"{v:.0%}" if col == "Discount" else f"{v:.1%}"
    if col in MONEY or col.endswith("_NGN"):
        return f"-₦{-v:,.0f}" if v < 0 else f"₦{v:,.0f}"
    if col in MONTHS or col == "FY_2025":
        return f"{v:,.1f}"
    if isinstance(v, (float, np.floating)):
        return f"{v:,.1f}" if abs(v) < 1000 else f"{v:,.0f}"
    if isinstance(v, (int, np.integer)):
        return str(v) if col in ("Row_ID", "Age", "Floor_Area_sqm") else f"{v:,}"
    return str(v)


def excerpt(name, cols=None, n=5, rows=None, sort=None, ascending=False):
    df = table(name)
    if rows is not None:
        df = rows(df)
    if sort:
        df = df.sort_values(sort, ascending=ascending)
    cols = cols or list(df.columns)
    total = len(df)
    df = df.head(n)
    recs = df[cols].to_dict("records")
    out = [list(cols)] + [[fmt(c, r[c]) for c in cols] for r in recs]
    return out, total
