"""Reference cleaning of the Week 2 dirty Mercyland 2023 EMR export (mirrors the Power Query steps in the answer key)."""
from __future__ import annotations

import pandas as pd

WARD_MAP = {"obeche ward": "Obeche Ward", "obeche": "Obeche Ward", "mahogany ward": "Mahogany Ward", "mahogony ward": "Mahogany Ward", "mahogany": "Mahogany Ward",
            "baobab ward": "Baobab Ward", "baobab": "Baobab Ward", "iroko ward": "Iroko Ward", "iroko": "Iroko Ward"}
GENDER_MAP = {"m": "Male", "male": "Male", "f": "Female", "female": "Female"}
TYPE_MAP = {"emerg": "Emergency", "emergency": "Emergency", "urgent": "Urgent", "elective": "Elective"}


def ward_key(x):
    return " ".join(str(x).split()).lower()


def clean(d: pd.DataFrame):
    c = d.copy()
    c["Patient_Name"] = c.Patient_Name.str.strip().str.title()
    c["Ward"] = c.Ward.map(lambda x: WARD_MAP[ward_key(x)])
    c["Gender"] = c.Gender.map(lambda x: GENDER_MAP[str(x).strip().lower()])
    c["Admission_Route"] = c.Admission_Route.map(lambda x: TYPE_MAP[str(x).strip().lower()])
    c["Bill_Amount_NGN"] = c.Bill_Amount_NGN.map(lambda x: float(str(x).replace(",", "")) if isinstance(x, str) else float(x))
    c["HMO"] = c.HMO.fillna("Self-pay / Unknown")
    c["Admit_Date"] = pd.to_datetime(c.Admit_Date)
    c["Discharge_Date"] = pd.to_datetime(c.Discharge_Date)
    steps = [("Rows in EMR export", len(c))]
    c = c.drop_duplicates()
    steps.append(("After removing duplicates (AFTER standardising text)", len(c)))
    neg = c[c.Bill_Amount_NGN < 0]
    c = c[c.Bill_Amount_NGN >= 0].copy()
    steps.append(("After removing negative-bill reversals", len(c)))
    los = (c.Discharge_Date - c.Admit_Date).dt.days
    c["LOS_Days"] = los
    c["LOS_Valid"] = (los.notna() & (los >= 0)).map({True: "Yes", False: "No"})
    c["Age_Valid"] = c.Age.between(1, 120).map({True: "Yes", False: "No"})
    c["Quarter"] = "Q" + c.Admit_Date.dt.quarter.astype(str)
    c["Month"] = c.Admit_Date.dt.strftime("%b")
    return c.reset_index(drop=True), steps, neg
