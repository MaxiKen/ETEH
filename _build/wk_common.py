"""Loaders shared by workbook builders."""
from __future__ import annotations

import pandas as pd

from common import STAGE


def L(name):
    return pd.read_pickle(STAGE / f"{name}.pkl")


def sm_sales():
    d = L("sm_sales").copy()
    d["Order_Date"] = pd.to_datetime(d.Order_Date)
    d["Ship_Date"] = pd.to_datetime(d.Ship_Date)
    return d


def lc_adm():
    d = L("lc_adm").copy()
    d["Admit_Date"] = pd.to_datetime(d.Admit_Date)
    d["Discharge_Date"] = pd.to_datetime(d.Discharge_Date)
    return d


def hotel():
    d = L("hotel").copy()
    d["Check_In"] = pd.to_datetime(d.Check_In)
    d["Booked_On"] = pd.to_datetime(d.Booked_On)
    return d


def smo():
    t = {k: L(f"smo_{k}").copy() for k in ["customers", "sellers", "products", "deliveries", "transactions"]}
    t["transactions"]["order_date"] = pd.to_datetime(t["transactions"]["order_date"])
    for c in ("shipped_date", "delivered_date"):
        t["deliveries"][c] = pd.to_datetime(t["deliveries"][c])
    return t


DATASETS_TEXT = {
    "SuperMart Stores": "SuperMart Retail - 12 Nigerian stores, 9,994 sales lines 2022-2025 (built from 'Sample - Superstore.csv': stores re-mapped, dates +8 yrs, values x1,000 = ₦)",
    "SuperMart Online": "SuperMart Online marketplace - 5 related tables, 11,000 orders 2020-2025 (built from 'e-commerce.xlsx'; keys and ₦ prices standardised). SQL: supermart_online.db",
    "Mercyland": "Mercyland Specialist Hospital, Abuja - 58,856 inpatient episodes 2019-2024, 200 staffed beds. SQL: mercyland_hospital.db",
    "Mercyland Screening": "Mercyland community diabetes screening programme - 60,000 adult screenings, 2023-2024",
    "Kingsway": "Kingsway University, Ilorin - 5,840 students, 2024/2025 first semester",
    "Kingsway Welfare": "Kingsway student welfare survey - 9,600 respondents, SENSITIVE",
    "Digital Wellbeing": "Kingsway digital habits survey - 640 students from 5 West African countries",
    "Salary Survey": "ETEH Talent Salary Survey - 1,781 professionals (built from 'Salary_Data.csv')",
    "Atlantic Pearl": "Atlantic Pearl Hotel, Lekki - 38,400 room reservations 2023-2024, 320 rooms",
}
