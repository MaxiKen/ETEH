"""Week 1 shared example tables (edition 3).

One place defines every example table by its Excel table name. The slide builder (07_v3_week1_deck.py) prints
small excerpts of these tables on the slides, and the Week 1 workbook builder (02_sec_w1.py) writes the
same tables, so a table name seen on a slide is the table the tutor opens in the workbook.
All values come from the bootcamp datasets (see _build/01_build_datasets.py).
"""
from __future__ import annotations

import os
import sys
from functools import lru_cache

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(__file__))
from wk_common import L, hotel, lc_adm, sm_sales, smo  # noqa: E402

MONEY = {"Sales", "Profit", "Target", "Bill_Amount_NGN", "Annual_Salary_NGN", "Avg_Order_Value", "Unit_Price", "Cost",
         "Reservation_Value_NGN", "Room_Rate_NGN", "total_price", "unit_price", "Annual_Target_2025", "Sales_2024",
         "Sales_2025", "Loss", "Average_Bill", "Revenue", "Amount", "Price", "VAT", "Total"}
PCT = {"Margin", "Discount", "Attainment", "Occupancy_Pct", "Share", "Growth", "Cancel_Rate", "Readmission_Rate",
       "Pass_Rate", "At_Risk_Rate", "Margin_2024", "Margin_2025"}


# ---------------------------------------------------------------- base data
@lru_cache(None)
def sales():
    d = sm_sales()
    d["Year"] = d.Order_Date.dt.year
    return d


@lru_cache(None)
def stores():
    return L("sm_stores")


def s25():
    d = sales()
    return d[d.Year == 2025]


# ---------------------------------------------------------------- named tables
@lru_cache(None)
def table(name: str) -> pd.DataFrame:
    """Return the full Excel table `name` as a DataFrame (columns in workbook order)."""
    return _BUILD[name]().reset_index(drop=True)


def _sales_lines():
    cols = ["Row_ID", "Order_ID", "Order_Date", "Customer_Name", "Store_Name", "Category", "Sub_Category", "Sales",
            "Quantity", "Discount", "Profit"]
    return sales()[cols]


def _ikeja_lines():
    d = s25()
    d = d[d.Store_ID == "ST01"].sort_values(["Order_Date", "Row_ID"]).head(12)
    return d[["Order_ID", "Order_Date", "Customer_Name", "Category", "Sub_Category", "Sales", "Quantity", "Profit"]]


def _stores25():
    g = s25().groupby("Store_ID").agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"),
                                      Orders=("Order_ID", "nunique")).reset_index()
    st = stores()[["Store_ID", "Store_Name", "Zone", "Store_Manager", "Annual_Target_2025"]]
    g = st.merge(g, on="Store_ID").rename(columns={"Annual_Target_2025": "Target"})
    g["Margin"] = g.Profit / g.Sales
    g["Attainment"] = g.Sales / g.Target
    return g[["Store_ID", "Store_Name", "Zone", "Store_Manager", "Sales", "Profit", "Orders", "Target", "Margin", "Attainment"]]


def _years():
    g = sales().groupby("Year").agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"),
                                    Orders=("Order_ID", "nunique")).reset_index()
    g["Margin"] = g.Profit / g.Sales
    g["Growth"] = g.Sales.pct_change()
    return g


def _disc_band():
    d = s25().copy()
    d["Discount_Band"] = pd.cut(d.Discount, [-0.01, 0.0, 0.2, 1.0], labels=["No discount", "1-20%", "Above 20%"])
    g = d.groupby("Discount_Band", observed=True).agg(Lines=("Row_ID", "count"), Sales=("Sales", "sum"),
                                                      Profit=("Profit", "sum")).reset_index()
    g["Margin"] = g.Profit / g.Sales
    return g


def _cat25():
    d = s25()
    g = d.groupby("Category").agg(Sales=("Sales", "sum"), Profit=("Profit", "sum")).reset_index()
    p24 = sales()[sales().Year == 2024].groupby("Category").agg(S=("Sales", "sum"), P=("Profit", "sum"))
    g["Margin"] = g.Profit / g.Sales
    g["Margin_2024"] = (p24.P / p24.S).reindex(g.Category).values
    return g


def _students():
    cv = L("cv_perf")
    return cv[["Matric_No", "Faculty", "Department", "Level", "Gender", "Lecture_Attendance_Pct", "Study_Hours_Wk",
               "Final_Score", "Grade", "At_Risk"]]


def _faculty():
    cv = L("cv_perf")
    g = cv.groupby("Faculty").agg(Students=("Matric_No", "count"), Avg_Score=("Final_Score", "mean"),
                                  At_Risk=("At_Risk", lambda v: (v == "Yes").sum())).reset_index()
    g["At_Risk_Rate"] = g.At_Risk / g.Students
    return g


def _admissions23():
    a = lc_adm()
    a = a[a.Admit_Date.dt.year == 2023]
    return a[["Episode_ID", "Patient_Name", "Age", "Department", "Ward", "Admission_Route", "Admit_Date",
              "Discharge_Date", "LOS_Days", "HMO", "Bill_Amount_NGN", "Readmitted_30d"]]


def _dept23():
    a = _admissions23()
    g = a.groupby("Department").agg(Admissions=("Episode_ID", "count"), Avg_LOS=("LOS_Days", "mean"),
                                    Readmitted=("Readmitted_30d", lambda v: (v == "Yes").sum())).reset_index()
    g["Readmission_Rate"] = g.Readmitted / g.Admissions
    return g.sort_values("Readmission_Rate", ascending=False)


def _census23():
    c = L("lc_census").copy()
    c["Census_Date"] = pd.to_datetime(c.Census_Date)
    return c[c.Census_Date.dt.year == 2023]


def _bookings():
    h = hotel()
    return h[["Reservation_ID", "Check_In", "Room_Category", "Sales_Channel", "Days_Ahead", "Nights",
              "Room_Rate_NGN", "Reservation_Status"]]


def _hotel_seg():
    h = hotel()
    g = h.groupby("Sales_Channel").agg(Bookings=("Reservation_ID", "count"),
                                        Cancelled=("Reservation_Status", lambda v: (v == "Cancelled").sum())).reset_index()
    g["Cancel_Rate"] = g.Cancelled / g.Bookings
    return g.sort_values("Bookings", ascending=False)


def _salary():
    return L("sal_clean")[["Respondent_ID", "Gender", "Education_Level", "Job_Title", "Years_of_Experience",
                           "Seniority", "Annual_Salary_NGN"]]


def _seniority():
    s = L("sal_clean")
    order = ["Junior", "Mid", "Senior", "Executive"]
    g = s.groupby("Seniority").agg(People=("Respondent_ID", "count"), Median_Salary=("Annual_Salary_NGN", "median"),
                                   Avg_Experience=("Years_of_Experience", "mean")).reset_index()
    g["_o"] = g.Seniority.map({k: i for i, k in enumerate(order)}).fillna(99)
    return g.sort_values("_o").drop(columns="_o")


def _orders_online():
    t = smo()["transactions"]
    return t[["order_id", "customer_id", "order_date", "quantity", "unit_price", "total_price", "payment_method",
              "order_status"]]


def _pay_mix():
    t = smo()["transactions"]
    g = t.groupby("payment_method").agg(Orders=("order_id", "count"), Revenue=("total_price", "sum")).reset_index()
    g["Share"] = g.Orders / g.Orders.sum()
    return g.sort_values("Orders", ascending=False).rename(columns={"payment_method": "Payment_Method"})


def _couriers():
    dv = smo()["deliveries"]
    g = dv.groupby("courier_name").agg(Deliveries=("delivery_id", "count"),
                                       Delivered=("delivery_status", lambda v: (v == "Delivered").sum())).reset_index()
    g["Delivered_Rate"] = g.Delivered / g.Deliveries
    return g.sort_values("Deliveries", ascending=False).rename(columns={"courier_name": "Courier"})


def _dirty():
    d = L("dirty_sm")
    return d[["Row_ID", "Order_ID", "Order_Date", "Store_Name", "Zone", "Category", "Sales", "Discount", "Profit"]]


_BUILD = {
    "tSales": _sales_lines, "tIkejaLines": _ikeja_lines, "tStores25": _stores25, "tYears": _years,
    "tDiscBand": _disc_band, "tCat25": _cat25, "tStudents": _students, "tFaculty": _faculty,
    "tAdmissions23": _admissions23, "tDept23": _dept23, "tCensus23": _census23, "tBookings": _bookings,
    "tHotelSeg": _hotel_seg, "tSalary": _salary, "tSeniority": _seniority, "tOrdersOnline": _orders_online,
    "tPayMix": _pay_mix, "tCouriers": _couriers, "tDirtySales": _dirty,
}


# ---------------------------------------------------------------- slide formatting
def fmt(col, v):
    if v is None or (isinstance(v, float) and np.isnan(v)) or v is pd.NaT:
        return ""
    if isinstance(v, (pd.Timestamp,)) or (hasattr(v, "isoformat") and not isinstance(v, str)):
        return pd.Timestamp(v).strftime("%Y-%m-%d")
    if col in PCT or col.endswith("_Rate") or col.endswith("_Pct") and isinstance(v, float) and v <= 1.5:
        return f"{v:.1%}" if col != "Discount" else f"{v:.0%}"
    if col in MONEY or col.startswith("Median_Salary") or col.endswith("_NGN"):
        return f"-₦{-v:,.0f}" if v < 0 else f"₦{v:,.0f}"
    if isinstance(v, (float, np.floating)):
        return f"{v:,.1f}" if abs(v) < 1000 else f"{v:,.0f}"
    if isinstance(v, (int, np.integer)):
        return str(v) if col in ("Row_ID", "Level", "Year") else f"{v:,}"
    return str(v)


def excerpt(name, cols=None, n=5, rows=None, sort=None, ascending=False):
    """Formatted rows (header first) for a slide table. rows: callable(df)->df to pick specific rows."""
    df = table(name)
    if rows is not None:
        df = rows(df)
    if sort:
        df = df.sort_values(sort, ascending=ascending)
    cols = cols or list(df.columns)
    total = len(df)
    df = df.head(n)
    recs = df[cols].to_dict("records")          # keeps each column's own type (iterrows would upcast ints)
    out = [list(cols)] + [[fmt(c, r[c]) for c in cols] for r in recs]
    return out, total
