"""Week 4 shared data (edition 3).

One place defines every Week 4 example table by its Excel table name. The slide builder (07_v3_week4_deck.py) prints small
excerpts of these tables and computes every figure from them (v3/w4_fx.py); the Week 4 workbook builder (02_sec_w4.py) writes
the same tables on Data_ sheets, so a table name seen on a slide is the table the tutor opens in the workbook.

Sources (all inside the pack, via the _stage pickles of 01_build_datasets.py):
  tSal       ETEH Talent salary survey, cleaned in Week 2 (1,781 respondents) + 0/1 dummy columns for regression
  tOrders    SuperMart Online order lines (11,000) with product category
  tCV        Kingsway University student results 2024/25 (5,840)
  tLOS       Mercyland admissions 2023, minimised (no names or patient IDs)
  tCensus23  Mercyland daily bed census 2023 (365 days, 200 staffed beds)
  tHotel     Atlantic Pearl bookings 2023-2024
  tDash25    SuperMart Stores FY2025 sales lines (3,312)
  tMonthly   SuperMart Stores monthly sales Jan 2022 - Dec 2025 (48 months): the forecasting series
  tStores    the 12 SuperMart stores with floor area and FY2025 target
"""
from __future__ import annotations

import os
import sys
from functools import lru_cache

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(__file__))
from wk_common import L, hotel, lc_adm, sm_sales, smo  # noqa: E402

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
ROOMS = 320            # Atlantic Pearl modelled room count (peak nights in the data reach 307 rooms in use)


@lru_cache(None)
def _sm():
    d = sm_sales()
    d["Year"] = d.Order_Date.dt.year
    return d


def sm() -> pd.DataFrame:
    return _sm().copy()


@lru_cache(None)
def _sal():
    s = L("sal_clean").copy()
    s["Is_Masters"] = (s.Education_Level == "Master's").astype(int)
    s["Is_PhD"] = (s.Education_Level == "PhD").astype(int)
    return s[["Respondent_ID", "Gender", "Education_Level", "Job_Title", "Job_Family", "Seniority", "Age", "Years_of_Experience",
              "Is_Masters", "Is_PhD", "Annual_Salary_NGN"]].reset_index(drop=True)


@lru_cache(None)
def _orders():
    o = smo()
    tx, pr = o["transactions"], o["products"]
    d = tx.merge(pr[["product_id", "category"]], on="product_id", how="left")
    return d[["order_id", "order_date", "customer_id", "category", "quantity", "total_price", "order_status"]].reset_index(drop=True)


@lru_cache(None)
def _cv():
    c = L("cv_perf")
    return c[["Matric_No", "Faculty", "Gender", "Study_Hours_Wk", "Lecture_Attendance_Pct", "Sleep_Hrs_Night", "Last_Sem_Score",
              "Tutorials_Attended", "Exercise_Hrs_Wk", "Final_Score", "Grade", "At_Risk"]].reset_index(drop=True)


@lru_cache(None)
def _los():
    a = lc_adm()
    a = a[a.Admit_Date.dt.year == 2023]
    return a[["Episode_ID", "Admit_Date", "Department", "Admission_Route", "Age", "Gender", "LOS_Days", "Bill_Amount_NGN",
              "HMO", "Readmitted_30d"]].reset_index(drop=True)


@lru_cache(None)
def _census23():
    c = L("lc_census").copy()
    c["Census_Date"] = pd.to_datetime(c.Census_Date)
    return c[c.Census_Date.dt.year == 2023][["Census_Date", "Inpatients", "Occupancy_Pct"]].reset_index(drop=True)


@lru_cache(None)
def _hotel():
    h = hotel()
    h["Check_In_Year"] = h.Check_In.dt.year
    return h[["Reservation_ID", "Check_In", "Check_In_Year", "Room_Category", "Sales_Channel", "Days_Ahead", "Nights",
              "Guest_Requests", "Returning_Guest", "Room_Rate_NGN", "Reservation_Status", "Reservation_Value_NGN"]].reset_index(drop=True)


@lru_cache(None)
def _dash25():
    s = _sm()
    return s[s.Year == 2025][["Order_ID", "Order_Date", "Store_ID", "Store_Name", "Zone", "Category", "Sales", "Profit"]].reset_index(drop=True)


@lru_cache(None)
def _monthly():
    s = _sm()
    m = s.groupby(s.Order_Date.dt.to_period("M")).Sales.sum()
    d = pd.DataFrame({"Month": [p.to_timestamp() for p in m.index], "Sales": m.values.astype(float)})
    d.insert(1, "t", range(1, len(d) + 1))
    d.insert(2, "Year", d.Month.dt.year)
    d.insert(3, "Month_No", d.Month.dt.month)
    return d


@lru_cache(None)
def _stores():
    st = L("sm_stores")
    tg = st[["Store_ID", "Store_Name", "Zone", "Floor_Area_sqm", "Annual_Target_2025"]].copy()
    return tg.reset_index(drop=True)


TABLES = {"tSal": _sal, "tOrders": _orders, "tCV": _cv, "tLOS": _los, "tCensus23": _census23, "tHotel": _hotel,
          "tDash25": _dash25, "tMonthly": _monthly, "tStores": _stores}


def table(name: str) -> pd.DataFrame:
    return TABLES[name]().copy()


MONEY = {"total_price", "Sales", "Profit", "Bill_Amount_NGN", "Annual_Salary_NGN", "Room_Rate_NGN", "Reservation_Value_NGN",
         "Annual_Target_2025"}
PCT = {"Occupancy_Pct"}
PLAIN = {"Year", "Month_No", "t", "Check_In_Year", "Age", "Is_Masters", "Is_PhD"}


def fmt(col, x):
    if isinstance(x, pd.Timestamp):
        return x.strftime("%b %Y") if col == "Month" else x.strftime("%Y-%m-%d")
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""
    if col in PLAIN and isinstance(x, (int, float, np.integer, np.floating)) and float(x).is_integer():
        return str(int(x))
    if col in MONEY and isinstance(x, (int, float, np.integer, np.floating)):
        return f"₦{float(x):,.0f}"
    if col in PCT and isinstance(x, (float, np.floating)):
        return f"{x:.1%}"
    if isinstance(x, (float, np.floating)):
        return f"{int(x):,}" if float(x).is_integer() else f"{x:,.1f}"
    if isinstance(x, (int, np.integer)):
        return f"{int(x):,}"
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
