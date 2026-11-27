"""Student Project 2 - Eko Suites Hotel: clean the PMS export, explain cancellations, recommend a deposit policy."""
from __future__ import annotations

import numpy as np
import pandas as pd

from asg_common import Asg
from common import SPR
from proj_common import SKEY, brief_pdf
from wk_common import DATASETS_TEXT, L, hotel

NAME = "Project2_EkoSuites_Hotel"
FOLDER = SPR / NAME
(FOLDER / "Data").mkdir(parents=True, exist_ok=True)
ROOMS = 320

raw = L("dirty_hotel").copy()
raw.to_csv(FOLDER / "Data" / "EkoSuites_PMS_Export_RAW.csv", index=False)
SEG = {"OTA": "Online Travel Agent", "Offline": "Walk-in / Travel Agent", "corporate": "Corporate", "Comp": "Complimentary", "Aviation": "Airline Crew"}

# ---------------- reference clean
c = raw.drop_duplicates().copy()
n_dedup = len(c)
c["Market_Segment"] = c.Market_Segment.replace(SEG)
c["Booking_Status"] = c.Booking_Status.str.title()
inv = c.Arrival_Date.str.startswith("29/02/2023")
c["Booking_Date"] = pd.to_datetime(c.Booking_Date, format="%d/%m/%Y")
for col in ["Adults", "Children", "Weekend_Nights", "Week_Nights", "Lead_Time_Days", "Previous_Cancellations", "Previous_Stays", "Special_Requests"]:
    c[col] = pd.to_numeric(c[col])
c["Avg_Room_Rate_NGN"] = pd.to_numeric(c.Avg_Room_Rate_NGN)
arr = pd.to_datetime(c.Arrival_Date.where(~inv), format="%d/%m/%Y", errors="coerce")
c["Arrival_Date"] = arr.where(~inv, c.Booking_Date + pd.to_timedelta(c.Lead_Time_Days, unit="D"))
c["Total_Nights"] = c.Weekend_Nights + c.Week_Nights
c["Booking_Value_NGN"] = c.Total_Nights * c.Avg_Room_Rate_NGN
ref = hotel().set_index("Booking_ID").sort_index()
chk = c.set_index("Booking_ID").sort_index()
assert len(chk) == len(ref) and (chk.Arrival_Date == ref.Arrival_Date).all() and np.allclose(chk.Booking_Value_NGN, ref.Booking_Value_NGN)
h = c
h["Cancelled"] = (h.Booking_Status == "Cancelled").astype(int)
h["Year"] = h.Arrival_Date.dt.year

n_seg_fix = int(raw.Market_Segment.isin(SEG).sum())
n_status_fix = int(raw.Booking_Status.isin(["HONOURED", "CANCELLED"]).sum())
n_zero = int((h.Avg_Room_Rate_NGN == 0).sum())
canc = h.Cancelled.mean()
seg_rate = h.groupby("Market_Segment").Cancelled.mean()
bands = pd.cut(h.Lead_Time_Days, [-1, 7, 30, 90, 180, 10_000], labels=["0-7", "8-30", "31-90", "91-180", "181+"])
band_rate = h.groupby(bands, observed=True).Cancelled.mean()
h24 = h[(h.Year == 2024)]
hon24 = h24[h24.Booking_Status == "Honoured"]
paid24 = hon24[hon24.Avg_Room_Rate_NGN > 0]
rev24 = hon24.Booking_Value_NGN.sum()
rn24 = hon24.Total_Nights.sum()
adr24 = paid24.Booking_Value_NGN.sum() / paid24.Total_Nights.sum()
occ24 = rn24 / (ROOMS * 366)
revpar24 = rev24 / (ROOMS * 366)
lost24 = h24[h24.Cancelled == 1].Booking_Value_NGN.sum()
r_lead = np.corrcoef(h.Lead_Time_Days, h.Cancelled)[0, 1]
rep = h.groupby("Repeat_Guest").Cancelled.mean()
sp = h[h.Room_Type.isin(["Deluxe", "Standard Twin"])]
sp_all = sp.groupby("Room_Type").Cancelled.mean()
sp_yr = sp.groupby(["Year", "Room_Type"]).Cancelled.mean().unstack()
sp_n = sp.groupby(["Year", "Room_Type"]).size().unstack()
risk = h24[(h24.Market_Segment == "Online Travel Agent") & (h24.Lead_Time_Days > 90) & (h24.Cancelled == 1)]
risk_val = risk.Booking_Value_NGN.sum()
DEP = 0.30


def build(key):
    A = Asg("P2", "Student Project 2 - Eko Suites Hotel: Cancellations, Revenue and a Deposit Policy", key,
            "Student Project 2 · cleaning · KPIs · diagnosis · what-if · Eko Suites Hotel, Victoria Island",
            path=(SKEY / f"{NAME}_KEY.xlsx") if key else (FOLDER / f"{NAME}_Workbook.xlsx"), label="Student Project 2",
            submission="Submit this workbook, your cleaning/analysis workbook (with the Power Query), a dashboard and a 1-page recommendation memo. Names: Surname_Firstname_P2_*.")
    s = A.task("M1 Clean the Export", "Milestone 1 - Clean the property-management-system export", "2.5-2.9",
               "Data/EkoSuites_PMS_Export_RAW.csv is the bookings export as the front office sent it. Clean it with Power Query (recommended) and record each fix. "
               "The answer key reconciles to the official clean file in 00_Core_Datasets.")
    rows = [("Rows in the raw export.", len(raw), "int", 1, ""),
            ("Rows after removing exact duplicate bookings.", n_dedup, "int", 2, "120 re-sent by the channel manager."),
            ("Rows whose Market_Segment had to be recoded (OTA, Offline, corporate, Comp, Aviation).", n_seg_fix, "int", 2,
             "Map: OTA→Online Travel Agent; Offline→Walk-in / Travel Agent; corporate→Corporate; Comp→Complimentary; Aviation→Airline Crew."),
            ("Rows whose Booking_Status was in capitals.", n_status_fix, "int", 1, ""),
            ("How many rows have the impossible arrival date 29/02/2023?", int(inv.sum()), "int", 2, ""),
            ("How should those arrival dates be repaired?", "Booking_Date + Lead_Time_Days", "gen", 3,
             "Recovers the true date exactly (e.g. 20/08/2023 + 34 days = 23/09/2023). Never guess 28/02 or 01/03."),
            ("Bookings with a zero room rate (keep them - complimentary/crew - but exclude from ADR).", n_zero, "int", 2, ""),
            ("Rebuild Total_Nights (= Weekend_Nights + Week_Nights). Total room-nights in the clean file?", int(h.Total_Nights.sum()), "int", 3, ""),
            ("Rebuild Booking_Value_NGN (= Total_Nights × Avg_Room_Rate_NGN). Total value of HONOURED bookings arriving in 2024 (₦)?", rev24, "naira", 4,
             "Control total - must match the official clean file to the naira."),
            ("Clean row count you will analyse.", len(h), "int", 1, "= 00_Core_Datasets/EkoSuites_Hotel_Bookings_2023_2024.csv")]
    for i, (t, v, k, m, gd) in enumerate(rows, 1):
        choices = ["Booking_Date + Lead_Time_Days", "Change to 28/02/2023", "Change to 01/03/2023", "Delete the rows"] if i == 6 else None
        A.q(s, f"1.{i}", t, v, None, k, m, gd, choices=choices, header=(i == 1))
    A.long(s, "1.11", "Paste your cleaning log: one line per defect - defect, rows affected, fix.",
           f"Duplicates 120 → removed; segment codes {n_seg_fix:,} → mapping table; status capitals {n_status_fix} → Text.Proper; dd/mm/yyyy text dates → Change Type with Locale en-GB; "
           f"29/02/2023 ×4 → Booking_Date + Lead_Time; zero rates {n_zero} → kept, flagged, excluded from ADR; Children up to 10 ×3 → queried, kept; "
           "Total_Nights and Booking_Value rebuilt as calculated columns.", marks=4, rows=4, guide="Any 6 defects with counts and fixes.")
    A.end_task(s)

    s = A.task("M2 KPIs", "Milestone 2 - The General Manager's KPI pack", "2.10, 3.6, 4.1",
               f"Use the clean data. Eko Suites has {ROOMS} rooms; 2024 has 366 days. ADR excludes zero-rate stays. Occupancy = honoured room-nights ÷ ({ROOMS} × 366).")
    k = [("Overall cancellation rate (all bookings, both years).", canc, "pct", 2),
         ("Cancellation rate - Online Travel Agent.", seg_rate["Online Travel Agent"], "pct", 2),
         ("Cancellation rate - Corporate.", seg_rate["Corporate"], "pct", 2),
         ("Cancellation rate for bookings made 0-7 days ahead.", band_rate["0-7"], "pct", 2),
         ("Cancellation rate for bookings made 181+ days ahead.", band_rate["181+"], "pct", 2),
         ("2024 ADR - honoured, paid stays (₦).", adr24, "naira", 3),
         ("2024 occupancy.", occ24, "pct", 3),
         ("2024 RevPAR (₦) = honoured room revenue ÷ available room-nights.", revpar24, "naira", 3),
         ("Value of bookings arriving in 2024 that were CANCELLED (₦).", lost24, "naira", 3),
         ("Repeat guests' cancellation rate.", rep.get("Yes", np.nan), "pct", 3)]
    for i, (t, v, kk, m) in enumerate(k, 1):
        A.q(s, f"2.{i}", t, v, None, kk, m, "Accept ±0.1 pt / ±₦1 rounding." if i == 1 else "", header=(i == 1))
    A.end_task(s)

    s = A.task("M3 Diagnose", "Milestone 3 - Why do bookings cancel?", "4.2-4.3",
               "The GM believes Deluxe guests cancel more than Standard Twin guests and wants Deluxe deposits raised. Test that, and measure the lead-time effect.")
    A.q(s, "3.1", "Overall cancellation rate - Deluxe.", sp_all["Deluxe"], None, "pct", 2, header=True)
    A.q(s, "3.2", "Overall cancellation rate - Standard Twin.", sp_all["Standard Twin"], None, "pct", 2)
    A.q(s, "3.3", "2023 arrivals: Deluxe rate.", sp_yr.loc[2023, "Deluxe"], None, "pct", 1)
    A.q(s, "3.4", "2023 arrivals: Standard Twin rate.", sp_yr.loc[2023, "Standard Twin"], None, "pct", 1)
    A.q(s, "3.5", "2024 arrivals: Deluxe rate.", sp_yr.loc[2024, "Deluxe"], None, "pct", 1)
    A.q(s, "3.6", "2024 arrivals: Standard Twin rate.", sp_yr.loc[2024, "Standard Twin"], None, "pct", 1)
    A.long(s, "3.7", "Explain what you see and whether the GM's conclusion holds.",
           f"Simpson's paradox. Overall Deluxe cancels more ({sp_all['Deluxe']:.1%} vs {sp_all['Standard Twin']:.1%}), but within EACH year Standard Twin cancels more "
           f"(2023: {sp_yr.loc[2023, 'Standard Twin']:.1%} vs {sp_yr.loc[2023, 'Deluxe']:.1%}; 2024: {sp_yr.loc[2024, 'Standard Twin']:.1%} vs {sp_yr.loc[2024, 'Deluxe']:.1%}). "
           f"Deluxe bookings are concentrated in 2024 ({int(sp_n.loc[2024, 'Deluxe']):,} of {int(sp_n.Deluxe.sum()):,}), the year with more cancellations overall, which "
           "inflates its pooled rate. Room type is not the driver - do not raise Deluxe deposits on this evidence; look at channel and lead time.",
           marks=5, rows=5, guide="Names Simpson's paradox 2; explains the mix (year weights) 2; correct conclusion 1.")
    A.q(s, "3.8", "Correlation between Lead_Time_Days and Cancelled (1 = cancelled, 0 = honoured).", r_lead, None, "dec3", 3, "Point-biserial r; CORREL works.", header=True)
    A.long(s, "3.9", "In two sentences: what is the strongest cancellation driver you found, with evidence?",
           f"Lead time: bookings made 181+ days ahead cancel {band_rate['181+']:.1%} of the time vs {band_rate['0-7']:.1%} for 0-7 days (r = {r_lead:.2f}). "
           f"Channel compounds it - OTA {seg_rate['Online Travel Agent']:.1%} vs Corporate {seg_rate['Corporate']:.1%}.", marks=4, rows=3)
    A.end_task(s)

    s = A.task("M4 Recommend", "Milestone 4 - Deposit policy what-if, dashboard and memo", "4.6-4.9",
               f"Proposal: a {DEP:.0%} non-refundable deposit on Online Travel Agent bookings made more than 90 days ahead. Size it with 2024 data, then present.")
    A.q(s, "4.1", "2024 arrivals: OTA bookings with lead time > 90 days that were cancelled - how many?", len(risk), None, "int", 3, header=True)
    A.q(s, "4.2", "Their total booking value (₦).", risk_val, None, "naira", 3)
    A.q(s, "4.3", f"Deposit income if {DEP:.0%} had been retained on those cancellations (₦).", risk_val * DEP, None, "naira", 3)
    A.long(s, "4.4", "Name two risks of the policy and how you would monitor them.",
           "• Fewer OTA bookings (guests choose a competitor without deposits) → track OTA booking volume and conversion monthly vs 2024.\n"
           "• OTA contract terms may forbid deposits / channel ranking penalties → check contracts; pilot with one OTA first.\n"
           "• Guest complaints/refund disputes → log disputes; review after one quarter.", marks=4, rows=4)
    A.q(s, "4.5", "Dashboard: KPI cards (occupancy, ADR, RevPAR, cancellation), cancellation by channel and lead-time band, monthly trend, slicers (instructor marks).",
        "Rubric" if key else "", marks=10, header=True)
    A.q(s, "4.6", "One-page memo to the GM: finding, evidence, recommendation, risks (instructor marks).", "Rubric" if key else "", marks=7)
    A.end_task(s)
    return A.finish("Student Project 2 - Eko Suites Hotel",
                    ["Client: General Manager, Eko Suites Hotel, Victoria Island (320 rooms). One in three bookings is cancelled and the GM wants to know why - and what to do.",
                     "Clean the PMS export (M1), build the KPI pack (M2), diagnose the drivers (M3) and size a deposit policy (M4). Duration: Weeks 2-4, about 12-15 hours."],
                    [("Eko Suites", DATASETS_TEXT["Eko Suites"] + " - you receive the raw PMS export (36,395 rows) with deliberate defects.")],
                    files="04_Student_Projects/Project2_EkoSuites_Hotel/Data/EkoSuites_PMS_Export_RAW.csv",
                    rubric=[["Reconciliation", "M2-M4 must be computed on YOUR cleaned table; if M1 control totals differ, explain why."],
                            ["Tolerance", "Rates ±0.1 pt; money ±₦1."], ["Communication", "Memo: BLUF, numbers, action, owner."]])


def main():
    for k in (False, True):
        print("P2", "KEY" if k else "student", "->", build(k))
    brief_pdf(FOLDER / "Project2_Brief.pdf", "Student Project 2 - Eko Suites Hotel: Cancellations, Revenue and a Deposit Policy",
              "Power Query cleaning · KPIs (occupancy, ADR, RevPAR) · Simpson's paradox · what-if · executive memo", "Student Project 2 brief", [
                  ("h", "1. The client and the question"),
                  ("p", f"Eko Suites is a {ROOMS}-room business hotel on Victoria Island, Lagos. About a third of all bookings are cancelled. The General Manager asks: "
                        "<b>Which bookings cancel, why, and would a deposit policy pay for itself?</b> The GM already suspects Deluxe guests - test that belief before accepting it."),
                  ("h", "2. The data"),
                  ("p", "Data/EkoSuites_PMS_Export_RAW.csv - 36,395 booking rows exported from the property-management system (arrivals July 2023 to December 2024). "
                        "It has NOT been cleaned: dates are day-first text, codes are inconsistent, some rows are duplicated, and two derived columns are missing."),
                  ("t", [["Column", "Meaning"], ["Booking_ID", "Booking reference (key)"], ["Booking_Date / Arrival_Date", "dd/mm/yyyy text"],
                         ["Weekend_Nights / Week_Nights", "Nights booked; Total_Nights must be rebuilt"], ["Room_Type, Meal_Plan, Market_Segment", "Booking attributes"],
                         ["Lead_Time_Days", "Days between booking and arrival"], ["Avg_Room_Rate_NGN", "Average nightly rate; 0 = complimentary/crew"],
                         ["Booking_Status", "Honoured or Cancelled"]], [60, 120]),
                  ("h", "3. Milestones"),
                  ("t", [["Milestone", "What you produce", "Marks"], ["M1 Clean", "Repeatable Power Query, cleaning log, control totals", "25"],
                         ["M2 KPIs", "Cancellation (by channel and lead time), ADR, occupancy, RevPAR, revenue lost", "25"],
                         ["M3 Diagnose", "Test the Deluxe belief (by year), quantify the lead-time effect", "20"],
                         ["M4 Recommend", "Deposit what-if, dashboard, one-page memo", "30"]], [35, 120, 25]),
                  ("note", "Definitions: ADR = room revenue ÷ room-nights for honoured, paid stays. Occupancy = honoured room-nights ÷ (320 × days). "
                           "RevPAR = honoured room revenue ÷ available room-nights."),
                  ("h", "4. Deliverables and assessment"),
                  ("b", ["Project2_EkoSuites_Hotel_Workbook.xlsx with every yellow cell answered", "Your cleaning + analysis workbook (Power Query steps visible)",
                         "Dashboard sheet", "One-page memo to the General Manager"]),
                  ("p", "Pass 50%, distinction 70%. Numbers must reconcile from cleaning through to the memo."),
              ])


if __name__ == "__main__":
    main()
