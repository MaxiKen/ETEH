"""Reference cleaning of the Week 2 dirty Lifecare 2023 EMR export (mirrors the Power Query steps in the answer key)."""
from __future__ import annotations

import pandas as pd

WARD_MAP = {"ruby ward": "Ruby Ward", "ruby": "Ruby Ward", "sapphire ward": "Sapphire Ward", "saphire ward": "Sapphire Ward", "sapphire": "Sapphire Ward",
            "topaz ward": "Topaz Ward", "topaz": "Topaz Ward", "emerald ward": "Emerald Ward", "emerald": "Emerald Ward"}
GENDER_MAP = {"m": "Male", "male": "Male", "f": "Female", "female": "Female"}
TYPE_MAP = {"emerg": "Emergency", "emergency": "Emergency", "urgent": "Urgent", "elective": "Elective"}


def ward_key(x):
    return " ".join(str(x).split()).lower()


def clean(d: pd.DataFrame):
    c = d.copy()
    c["Patient_Name"] = c.Patient_Name.str.strip().str.title()
    c["Ward"] = c.Ward.map(lambda x: WARD_MAP[ward_key(x)])
    c["Gender"] = c.Gender.map(lambda x: GENDER_MAP[str(x).strip().lower()])
    c["Admission_Type"] = c.Admission_Type.map(lambda x: TYPE_MAP[str(x).strip().lower()])
    c["Billing_Amount"] = c.Billing_Amount.map(lambda x: float(str(x).replace(",", "")) if isinstance(x, str) else float(x))
    c["HMO"] = c.HMO.fillna("Self-pay / Unknown")
    c["Admission_Date"] = pd.to_datetime(c.Admission_Date)
    c["Discharge_Date"] = pd.to_datetime(c.Discharge_Date)
    steps = [("Rows in EMR export", len(c))]
    c = c.drop_duplicates()
    steps.append(("After removing duplicates (AFTER standardising text)", len(c)))
    neg = c[c.Billing_Amount < 0]
    c = c[c.Billing_Amount >= 0].copy()
    steps.append(("After removing negative-bill reversals", len(c)))
    los = (c.Discharge_Date - c.Admission_Date).dt.days
    c["Length_of_Stay"] = los
    c["LOS_Valid"] = (los.notna() & (los >= 0)).map({True: "Yes", False: "No"})
    c["Age_Valid"] = c.Age.between(1, 120).map({True: "Yes", False: "No"})
    c["Quarter"] = "Q" + c.Admission_Date.dt.quarter.astype(str)
    c["Month"] = c.Admission_Date.dt.strftime("%b")
    return c.reset_index(drop=True), steps, neg
