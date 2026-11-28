"""Headline facts of the edition-3 replacement datasets, computed from the CSVs (used by copy that quotes them)."""
from __future__ import annotations

import pandas as pd

from common import CORE, W2D

ep = pd.read_csv(CORE / "Mercyland_Inpatient_Episodes_2019_2024.csv", parse_dates=["Admit_Date", "Discharge_Date"])
ce = pd.read_csv(CORE / "Mercyland_Daily_Bed_Census.csv", parse_dates=["Census_Date"])
st = pd.read_csv(CORE / "Kingsway_Student_Results.csv")
ho = pd.read_csv(CORE / "AtlanticPearl_Reservations_2023_2024.csv", parse_dates=["Check_In", "Booked_On"])
sc = pd.read_csv(CORE / "Mercyland_Diabetes_Screening.csv")
we = pd.read_csv(CORE / "Kingsway_Student_Welfare_Survey_SENSITIVE.csv")
dg = pd.read_csv(CORE / "Kingsway_Digital_Habits_Survey.csv")
pa = pd.read_csv(CORE / "Mercyland_Patients.csv")
dirty = pd.read_csv(W2D / "Mercyland_Episodes_2023_DIRTY.csv", dtype=str, keep_default_na=False)
e23 = ep[ep.Admit_Date.dt.year == 2023]
c23 = ce[ce.Census_Date.dt.year == 2023]
occm = c23.groupby(c23.Census_Date.dt.month).Occupancy_Pct.mean()
C = (ho.Reservation_Status == "Cancelled")
F = dict(
    EP_ALL=len(ep), EP23=len(e23), PATIENTS=len(pa), ALOS23=e23.LOS_Days.mean(), READM23=(e23.Readmitted_30d == "Yes").mean(),
    OCC23=c23.Occupancy_Pct.mean(), HI95=int((c23.Occupancy_Pct >= 0.95).sum()), PEAKM=int(occm.idxmax()), PEAKOCC=occm.max(),
    BEDS=int(ce.Open_Beds.iloc[0]), DIRTY_ROWS=len(dirty),
    STUD=len(st), ATRISK=int((st.At_Risk == "Yes").sum()), SCORE=st.Final_Score.mean(),
    RES=len(ho), CANC=C.mean(), OTA=C[ho.Sales_Channel == "Online Travel Agent"].mean(),
    CORP=C[ho.Sales_Channel == "Corporate"].mean(), SCREEN=len(sc), WELF=len(we), DIG=len(dg),
)
if __name__ == "__main__":
    for k, v in F.items():
        print(k, v)
