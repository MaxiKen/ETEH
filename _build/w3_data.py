"""Week 3 shared data (edition 3).

One place defines every Week 3 example table by its Excel table name. The slide builder (07_v3_week3_deck.py) prints
small excerpts of these tables and runs the SQL it shows against the real SQLite files, and the Week 3 workbook builder
(02_sec_w3.py) writes the same tables, so a table name seen on a slide is the table the tutor opens in the workbook.

Sources (all inside the pack):
  Week3_Data_Files/supermart_online.db   SuperMart Online marketplace: transactions, deliveries, customers, products,
                                          sellers, nigeria_states, view v_order_lines
  Week3_Data_Files/mercyland_hospital.db  Mercyland Hospital: admissions, patients, doctors, wards, bed_census
  _stage pickles from 01_build_datasets.py: SuperMart Stores POS lines (sm_sales), stores, Kingsway students (cv_perf),
                                          Mercyland_Episodes_2023_DIRTY.csv (dirty_lc)
"""
from __future__ import annotations

import os
import sqlite3
import sys
from functools import lru_cache

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(__file__))
from common import W3D  # noqa: E402
from wk_common import L, lc_adm, sm_sales  # noqa: E402

DBS = {"smo": W3D / "supermart_online.db", "lc": W3D / "mercyland_hospital.db"}
DBNAME = {"smo": "supermart_online.db", "lc": "mercyland_hospital.db"}
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
NET = ("Delivered", "Completed")          # Net Revenue = order lines that were delivered or completed (not cancelled/returned)
SQL_LOG: list[str] = []                   # every query shown on a slide or in the workbook, written to ETEH_Week3_SQL_Examples.sql


# ---------------------------------------------------------------- SQL
@lru_cache(None)
def _q(sql: str, db: str) -> pd.DataFrame:
    with sqlite3.connect(DBS[db]) as c:
        return pd.read_sql(sql, c)


def q(sql: str, db: str = "smo") -> pd.DataFrame:
    """Run SQL on a Week 3 database and return the result (a copy, so callers may modify it)."""
    return _q(sql.strip(), db).copy()


def v(sql: str, db: str = "smo"):
    """Single value of a one-cell query."""
    x = _q(sql.strip(), db).iat[0, 0]
    return x.item() if hasattr(x, "item") else x


def log(tag: str, title: str, sql: str, db: str = "smo"):
    entry = f"-- {tag} {title}\n-- database: {DBNAME[db]}\n{sql.strip()};\n"
    if entry not in SQL_LOG:
        SQL_LOG.append(entry)


def cell(x, kind=None):
    """Format one result value for a slide table."""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "NULL"
    if kind == "n":
        return f"₦{float(x):,.0f}"
    if kind == "n2":
        return f"₦{float(x):,.2f}"
    if kind == "m":
        return f"₦{float(x) / 1e6:,.1f}m"
    if kind == "p":
        return f"{float(x):.1%}"
    if kind == "d1":
        return f"{float(x):,.1f}"
    if kind == "d2":
        return f"{float(x):,.2f}"
    if isinstance(x, (int, np.integer)) or (isinstance(x, float) and float(x).is_integer() and kind != "raw"):
        return f"{int(x):,}" if kind != "id" else str(int(x))
    if isinstance(x, float):
        return f"{x:,.2f}"
    return str(x)


def rows(sql: str, db: str = "smo", fmt: dict | None = None, head: list | None = None, n: int | None = None,
         extra: list | None = None):
    """Run SQL and return [header] + formatted rows for a slide grid. fmt maps column -> kind (see cell())."""
    df = q(sql, db)
    fmt = fmt or {}
    if n:
        df = df.head(n)
    out = [head or list(df.columns)]
    for r in df.itertuples(index=False):
        out.append([cell(x, fmt.get(c, "id" if c in PLAIN else None)) for c, x in zip(df.columns, r)])
    return out + (extra or [])


# ---------------------------------------------------------------- SuperMart Online (pandas twins of the SQL)
@lru_cache(None)
def smo(name: str) -> pd.DataFrame:
    return _q(f"SELECT * FROM {name}", "smo").copy()


@lru_cache(None)
def lines() -> pd.DataFrame:
    """transactions + product, customer and seller attributes (the flat order-line view, built in pandas)."""
    t = smo("transactions").copy()
    t["order_date"] = pd.to_datetime(t.order_date)
    p = smo("products")[["product_id", "product_name", "category", "subcategory", "brand_name"]]
    cu = smo("customers")[["customer_id", "membership_tier", "state", "zone"]].rename(columns={"state": "customer_state",
                                                                                            "zone": "customer_zone"})
    se = smo("sellers")[["seller_id", "store_name", "store_state", "store_zone"]]
    d = t.merge(p, on="product_id", how="left").merge(cu, on="customer_id", how="left").merge(se, on="seller_id", how="left")
    d["year"] = d.order_date.dt.year
    d["month"] = d.order_date.dt.month
    d["is_net"] = d.order_status.isin(NET)
    return d


@lru_cache(None)
def net() -> pd.DataFrame:
    return lines()[lines().is_net].copy()


def net_rev(df=None):
    return float((net() if df is None else df).total_price.sum())


@lru_cache(None)
def dimdate() -> pd.DataFrame:
    d = pd.DataFrame({"Date": pd.date_range("2020-01-01", "2025-12-31", freq="D")})
    d["Year"] = d.Date.dt.year
    d["Quarter"] = "Q" + d.Date.dt.quarter.astype(str)
    d["Month_No"] = d.Date.dt.month
    d["Month"] = d.Date.dt.strftime("%b")
    d["Year_Month"] = d.Date.dt.strftime("%Y-%m")
    d["Weekday"] = d.Date.dt.strftime("%a")
    return d


# ---------------------------------------------------------------- SuperMart Stores POS (FY2025 dashboard data)
@lru_cache(None)
def sm() -> pd.DataFrame:
    d = sm_sales()
    d["Year"] = d.Order_Date.dt.year
    return d


@lru_cache(None)
def s25() -> pd.DataFrame:
    d = sm()
    return d[d.Year == 2025].sort_values(["Order_Date", "Row_ID"]).reset_index(drop=True)


@lru_cache(None)
def stores() -> pd.DataFrame:
    return L("sm_stores")


@lru_cache(None)
def dash_stores() -> pd.DataFrame:
    s = s25().groupby(["Store_ID", "Store_Name", "Zone"]).agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"),
                                                               Orders=("Order_ID", "nunique")).reset_index()
    s = s.merge(stores()[["Store_ID", "Annual_Target_2025"]].rename(columns={"Annual_Target_2025": "Target"}), on="Store_ID")
    s["Attainment"] = s.Sales / s.Target
    return s


# ---------------------------------------------------------------- Mercyland / Kingsway
@lru_cache(None)
def lc() -> pd.DataFrame:
    return lc_adm()


@lru_cache(None)
def census() -> pd.DataFrame:
    c = L("lc_census").copy()
    c["Census_Date"] = pd.to_datetime(c.Census_Date)
    return c


@lru_cache(None)
def cv() -> pd.DataFrame:
    return L("cv_perf")


@lru_cache(None)
def dirty_lc() -> pd.DataFrame:
    return L("dirty_lc")


# ---------------------------------------------------------------- named example tables (same names as the workbook)
def _tx():
    t = smo("transactions").copy()
    return t


def _dash():
    d = s25()[["Row_ID", "Order_ID", "Order_Date", "Store_ID", "Store_Name", "Zone", "Category", "Sales", "Quantity", "Profit"]]
    return d.reset_index(drop=True)


def _dash_stores():
    return dash_stores()[["Store_ID", "Store_Name", "Zone", "Sales", "Profit", "Target", "Attainment"]]


def _adm23():
    a = lc()
    a = a[a.Admit_Date.dt.year == 2023]
    cols = ["Episode_ID", "Patient_No", "Age", "Gender", "Department", "Admission_Route", "Admit_Date", "Discharge_Date",
            "LOS_Days", "Ward", "Consultant_Name", "HMO", "Bill_Amount_NGN", "Readmitted_30d"]
    return a[cols].reset_index(drop=True)


def _cv():
    cols = ["Matric_No", "Faculty", "Department", "Level", "Gender", "Study_Hours_Wk", "Lecture_Attendance_Pct", "Sleep_Hrs_Night",
            "Final_Score", "Grade", "At_Risk"]
    return cv()[cols].copy()


TABLES = {
    "tTransactions": _tx, "tDeliveries": lambda: smo("deliveries"), "tCustomers": lambda: smo("customers"),
    "tProducts": lambda: smo("products"), "tSellers": lambda: smo("sellers"), "tStates": lambda: smo("nigeria_states"),
    "tDimDate": dimdate, "tDash": _dash, "tDashStores": _dash_stores, "tAdm23": _adm23, "tCensus": census, "tCV": _cv,
    "tMsDirty": dirty_lc,
}


@lru_cache(None)
def _table(name):
    return TABLES[name]()


def table(name: str) -> pd.DataFrame:
    return _table(name).copy()


MONEY = {"total_price", "unit_price", "commission_amount", "price", "discount_price", "total_spent", "total_revenue", "Sales",
         "Profit", "Target", "Bill_Amount_NGN"}
PCT = {"Attainment", "Occupancy_Pct"}
PLAIN = {"Year", "Month_No", "Level", "year", "month", "Row_ID", "Bed_No", "yr", "order_year", "adm_year"}


def fmt(col, x):
    if isinstance(x, pd.Timestamp):
        return x.strftime("%Y-%m-%d")
    if col in PLAIN and isinstance(x, (int, float, np.integer, np.floating)) and float(x).is_integer():
        return str(int(x))
    if col in MONEY and isinstance(x, (int, float, np.integer, np.floating)):
        return f"₦{float(x):,.0f}"
    if col in PCT and isinstance(x, (float, np.floating)):
        return f"{x:.1%}"
    if isinstance(x, (float, np.floating)) and float(x).is_integer():
        return f"{int(x):,}"
    if isinstance(x, (int, np.integer)):
        return f"{int(x):,}" if col not in ("Level",) else str(x)
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "NULL"
    return str(x)


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
