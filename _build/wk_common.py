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
    "Mercyland": "Mercyland Specialist Hospital - 54,860 admissions 2019-2024, 520 staffed beds (built from 'healthcare_dataset.csv'). SQL: mercyland_hospital.db",
    "Mercyland Screening": "Mercyland diabetes screening clinic - 100,000 screenings (built from 'diabetes_prediction_dataset.csv')",
    "Kingsway": "Kingsway University - 6,607 students, 2024/2025 first semester (built from 'StudentPerformanceFactors.csv')",
    "Kingsway Welfare": "Kingsway wellbeing survey - 27,870 respondents, SENSITIVE (built from 'Student Depression Dataset.csv')",
    "Digital Wellbeing": "Kingsway digital-wellbeing survey - 705 students in 110 countries (built from 'Students Social Media Addiction.csv')",
    "Salary Survey": "ETEH Talent Salary Survey - 1,781 professionals (built from 'Salary_Data.csv')",
    "Atlantic Pearl": "Atlantic Pearl Hotel, Lekki - 36,275 bookings 2023-2024 (built from 'Hotel Reservations.csv')",
}
