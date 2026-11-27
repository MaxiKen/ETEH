"""Helpers for Week 3 slide copy: every number on a slide is computed here or by running the SQL shown (never typed)."""
from __future__ import annotations

import numpy as np
import pandas as pd

import w3_data as D
from w3_data import q, v, rows, NET  # noqa: F401  (re-exported for the copy files)


def N(x, d=0):
    x = float(x)
    return f"-₦{-x:,.{d}f}" if x < 0 else f"₦{x:,.{d}f}"


def P(x, d=1):
    return f"{x:.{d}%}"


def I(x):
    return f"{int(round(float(x))):,}"


def Nm(x, d=1):
    return f"₦{float(x) / 1e6:,.1f}m" if d == 1 else f"₦{float(x) / 1e6:,.{d}f}m"


def F1(x):
    return f"{float(x):,.1f}"


# ---------------------------------------------------------------- SuperMart Online
LN = D.lines()                    # every order line with product / customer / seller attributes
NT = D.net()                      # Delivered + Completed lines
TX = D.smo("transactions")
CU, PR, SE, DV = D.smo("customers"), D.smo("products"), D.smo("sellers"), D.smo("deliveries")

GROSS = float(TX.total_price.sum())
NETREV = float(NT.total_price.sum())
COMM = float(TX.commission_amount.sum())
N_LINES, N_NET = len(TX), len(NT)
NO_DELIV = int(TX.delivery_id.isna().sum())
BUYERS = TX.customer_id.nunique()
NET_BUYERS = NT.customer_id.nunique()

YEAR = NT.groupby("year").total_price.sum()                                   # net revenue by year
M25 = NT[NT.year == 2025].groupby("month").total_price.sum().reindex(range(1, 13), fill_value=0)
M24 = NT[NT.year == 2024].groupby("month").total_price.sum().reindex(range(1, 13), fill_value=0)
YTD25 = M25.cumsum()
CAT = NT.groupby("category").total_price.sum().sort_values(ascending=False)
CAT_N = NT.groupby("category").size()
ZONE = NT.groupby("customer_zone").total_price.sum().sort_values(ascending=False)
STORE = NT.groupby("store_name").total_price.sum().sort_values(ascending=False)
AOV = NETREV / N_NET


def month_name(m):
    return D.MONTHS[m - 1]


def mon_full(m):
    return pd.Timestamp(2025, m, 1).strftime("%B")


# ---------------------------------------------------------------- SuperMart Stores FY2025 (dashboard data)
S25 = D.s25()
SM = D.sm()
DS = D.dash_stores()
STORES = D.stores()
FY25 = float(S25.Sales.sum())
FY24 = float(SM[SM.Year == 2024].Sales.sum())
PROFIT25 = float(S25.Profit.sum())
TARGET25 = float(STORES.Annual_Target_2025.sum())
ORDERS25 = S25.Order_ID.nunique()
ZS = S25.groupby("Zone").agg(Sales=("Sales", "sum"), Profit=("Profit", "sum")).sort_values("Sales", ascending=False)
ZS["Margin"] = ZS.Profit / ZS.Sales
ZT = DS.groupby("Zone").agg(Sales=("Sales", "sum"), Target=("Target", "sum"))
ZS["Attainment"] = ZT.Sales / ZT.Target
SMONTH = SM.assign(YM=SM.Order_Date.dt.to_period("M")).groupby("YM").Sales.sum()

# ---------------------------------------------------------------- Lifecare / Crestview
LC = D.lc()
CEN = D.census()
CV = D.cv()
DL = D.dirty_lc()


# ---------------------------------------------------------------- SQL panels: the code shown is the code run
import sqlite3 as _sq
import textwrap as _tw


def SQL(sql, db="smo", fmt=None, head=None, n=None, extra=None, label="Result", widths=None, size=12, note="",
        split=None, code_size=12):
    """code panel + result grid for one query. The grid is produced by running exactly the SQL in the panel."""
    sql = _tw.dedent(sql).strip("\n")
    r = rows(sql, db, fmt, head, n, extra)
    out = dict(code=sql, code_label=f"SQL  ·  {D.DBNAME[db].upper()}", code_db=db, code_size=code_size,
               grid=dict(label=label, rows=r, widths=widths, size=size, note=note))
    if split:
        out["split"] = split
    return out


def CODE(text, label="DAX", size=12, note=""):
    return dict(code=_tw.dedent(text).strip("\n"), code_label=label, code_size=size, code_note=note)


def sandbox(stmts, db="smo"):
    """Run statements on an in-memory COPY of a database (foreign keys on). Returns one result or 'Error: …' per statement."""
    src, mem = _sq.connect(D.DBS[db]), _sq.connect(":memory:")
    src.backup(mem)
    mem.execute("PRAGMA foreign_keys = ON")
    out = []
    for s in stmts:
        try:
            cur = mem.execute(s)
            r = cur.fetchall()
            out.append(r[0][0] if len(r) == 1 and len(r[0]) == 1 else r)
        except _sq.Error as e:
            out.append(f"Error: {e}")
    mem.close()
    return out
