"""Week 4 workbook, modules 4.10 - 4.14 (edition 3): auditing and protection, AI co-pilots, verifying AI output, data ethics,
and the benchmark lab.

4.10 adds a sheet '4.10 Audit Case' laid out exactly as the slides quote it: zone sales in C9:C12, the colleague's total in C13
(=SUM(C9:C11)) and three planted errors in the share column, so learners can trace and fix the real cells. 4.12 checks six AI
claims with live formulas (tClaims). 4.14 benchmarks sales per m², checks the PDF bulletin against the ledger and puts 95%
intervals on the readmission rates.
"""
from __future__ import annotations

import numpy as np
import pandas as pd

import w4_m11_14 as F14
from w1x import ModuleSheet
from wb4_common import DEF, N, f, tp

AUD = "'4.10 Audit Case'"
ZONES = ["South-West", "South-South", "North-Central", "South-East"]


def build(b, D):
    m410(b, D)
    m411(b, D)
    m412(b, D)
    m413(b, D)
    m414(b, D)


# =====================================================================================================
def audit_sheet(b, D):
    """The colleague's sheet, cell for cell as the slides quote it."""
    zn = D.D25.groupby("Zone").Sales.sum().reindex(ZONES)
    sub = zn.iloc[:3].sum()
    ws = b.wb.add_worksheet("4.10 Audit Case")
    b.toc.append(("4.10 Audit Case", "A colleague's FY2025 zone shares with three planted errors: trace and fix them (4.10.1 - 4.10.4)", "4.10.1 - 4.10.4"))
    F = b.F
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 16)
    ws.set_column(2, 2, 16)
    ws.set_column(3, 5, 15)
    ws.set_column(6, 6, 70)
    ws.write(0, 1, "4.10 Audit Case: FY2025 sales share by zone (as submitted by a colleague)", F["title"] if "title" in F else F["hdr"])
    notes = ["This sheet looks finished, but every submitted share is wrong. Use Trace Precedents, Show Formulas (Ctrl+`) and Evaluate Formula to find why.",
             "Error 1 is in the Total (C13). Error 2 is a typed number in the share column. Error 3 is a formula that lost its $ when copied.",
             "Fix them in place: column F turns to OK on every row when the submitted shares match the correct ones (the check in 4.10.4 then reads ✓).",
             "Correct shares (column E) are computed independently from tDash25."]
    for i, t in enumerate(notes):
        ws.merge_range(1 + i, 1, 1 + i, 6, t, F["note"])
    hdr = F["hdr"]
    for j, h in enumerate(["Zone", "Sales", "Submitted share", "Correct share", "Check"]):
        ws.write(7, 1 + j, h, hdr)
    nf, pf = b.nf(N), b.nf("pct")
    subm = [zn.iloc[0] / sub, zn.iloc[1] / sub, 0.15, 1.0]
    for i, z in enumerate(ZONES):
        r = 8 + i
        x = r + 1
        ws.write(r, 1, z, F["cell"])
        ws.write_formula(r, 2, f"=SUMIF(tDash25[Zone],B{x},tDash25[Sales])", nf, f(zn[z]))
        if i < 2:
            ws.write_formula(r, 3, f"=C{x}/$C$13", pf, f(subm[i]))
        elif i == 2:
            ws.write_number(r, 3, 0.15, pf)
        else:
            ws.write_formula(r, 3, "=C12/C12", pf, 1.0)
        ws.write_formula(r, 4, f"=C{x}/SUM($C$9:$C$12)", pf, f(zn[z] / zn.sum()))
        ws.write_formula(r, 5, f'=IF(ABS(D{x}-E{x})<0.0005,"OK","ERROR")', F["cell"], "ERROR")
    ws.write(12, 1, "Total", hdr)
    ws.write_formula(12, 2, "=SUM(C9:C11)", nf, f(sub))
    ws.write_formula(12, 3, "=SUM(D9:D12)", pf, f(sum(subm)))
    ws.write_formula(12, 4, "=SUM(E9:E12)", pf, 1.0)
    ws.write(14, 1, "Fixes:", hdr)
    ws.merge_range(14, 2, 14, 6, "C13: =SUM(C9:C12)   ·   D11: =C11/$C$13   ·   D12: =C12/$C$13", F["note"])
    return zn, sub


def m410(b, D):
    m = ModuleSheet(b, "4.10 Auditing", "4.10  Formula Auditing, Protection and Automation",
                    "Tutor walk-through after Module 4.10: trace and fix a broken sheet, protect a plan, record and document a macro", "4.10.1 - 4.10.8")
    m.dataset([("4.10 Audit Case", "The next sheet", "A colleague's zone-share table with three planted errors (cells as quoted on the slides)."),
               ("tDash25", "Sheet Data_Dash25 (A1)", "FY2025 sales lines: the control total.")])
    zn = D.D25.groupby("Zone").Sales.sum().reindex(ZONES)
    sub = zn.iloc[:3].sum()
    tot = float(zn.sum())
    tp(m, "4.10.1",
       [("Rows flagged ERROR on the audit sheet", f'=COUNTIF({AUD}!F9:F12,"ERROR")', 4, "int", "Every share is wrong."),
        ("Correct South-West share", f"={AUD}!E9", f(zn["South-West"] / tot), "pct", "Recomputed a second way.")],
       [("n", "What is the correct North-Central share?", f"={AUD}!E11", f(zn["North-Central"] / tot), "pct", "Correct North-Central share")])
    tp(m, "4.10.2",
       [("Submitted total (C13)", f"={AUD}!C13", f(sub), N, "=SUM(C9:C11): stops one row short."),
        ("Control total (all FY2025 sales)", "=SUM(tDash25[Sales])", f(tot), N, "")],
       [("n", "How much is missing from the submitted total?", f"=SUM(tDash25[Sales])-{AUD}!C13", f(tot - sub), N, "Missing from total")],
       steps=["Click C13 on the audit sheet › Formulas › Trace Precedents.", "The blue box covers C9:C11 only: South-East (C12) is left out.",
              "Formulas › Trace Dependents on C13 shows every share that uses the wrong total."])
    tp(m, "4.10.3",
       [("Typed value in D11", f"={AUD}!D11", 0.15, "pct", "Show Formulas reveals a number where a formula should be."),
        ("D12 as submitted", f"={AUD}!D12", 1.0, "pct", "=C12/C12: every zone divided by itself is 100%.")],
       [("t", "What should D12 contain? (type the formula without the = sign)", "C12/$C$13", None, None, ("=C12/$C$13", "C12/C$13", "=C12/C$13"))],
       steps=["Ctrl+` (Show Formulas) on the audit sheet: D11 shows 0.15, not a formula.", "Select D12 › Formulas › Evaluate Formula: it divides C12 by C12."])
    m.do([("Fix the three errors on '4.10 Audit Case' (C13, D11, D12).", "Column F reads OK on all four rows and D13 shows 100%.")])
    tp(m, "4.10.4",
       [("Rows OK on the audit sheet", f'=COUNTIF({AUD}!F9:F12,"OK")', 0, "int", "4 once all three errors are fixed."),
        ("Correct shares add up to", f"=SUM({AUD}!E9:E12)", 1.0, "pct0", "")], [])
    m.turn()
    m.ex_formula_live("Live check: have you fixed all three errors on the audit sheet? (nothing to type)",
                      f'=IF(COUNTIF({AUD}!F9:F12,"OK")=4,"✓ Fixed: all four shares match","✗ Not yet: "&COUNTIF({AUD}!F9:F12,"ERROR")&" rows still wrong")',
                      "✗ Not yet: 4 rows still wrong", key="Fix C13 to =SUM(C9:C12), D11 to =C11/$C$13 and D12 to =C12/$C$13.")
    m.do([("Formulas › Watch Window › Add Watch: the audit sheet's C13 and D13. Then fix C13.", "The Watch Window shows C13 change to the control total while you stay on this sheet.")])

    plan = {}

    def growth(mm):
        mm.h("FY2026 plan: the growth cell is unlocked (yellow); everything else stays locked when you protect the sheet")
        r = mm.s.r
        ws = mm.s.ws
        ws.merge_range(r, 1, r, 3, "Growth assumption", mm.s.F["lab"])
        ws.write_number(r, 4, 0.12, b.nf("pct", style="input", locked=False))
        plan["g"] = f"$E${r + 1}"
        mm.s.r = r + 2
    tp(m, "4.10.5",
       lambda mm: [("FY2026 plan", f"=SUM(tDash25[Sales])*(1+{plan['g']})", f(tot * 1.12), N, "Locked formula.")],
       [("n", "What would the plan be at 15% growth?", "=SUM(tDash25[Sales])*1.15", f(tot * 1.15), N, "Plan at 15%")],
       steps=True, extra=growth,
       do=[("Review › Protect Sheet (no password for practice) › OK. Try typing in the plan cell, then in the yellow growth cell.",
            "The plan cell refuses edits; the growth cell accepts them and the plan updates."),
           ("Review › Unprotect Sheet.", "Everything is editable again.")])
    tp(m, "4.10.6", None,
       [("t", "Which file type keeps a recorded macro: .xlsx or .xlsm?", ".xlsm", None, None, ("xlsm",))],
       code=True,
       do=[("View › Macros › Record Macro, name Refresh_And_Export_Report. Refresh All (Data tab), then export this workbook's '4.8 Report Writing' sheet as PDF. Stop recording.",
            "Alt+F11 shows the recorded code: RefreshAll and ExportAsFixedFormat, like the code above."),
           ("Save as Excel Macro-Enabled Workbook (.xlsm).", "The file keeps the macro; a .xlsx copy would drop it.")])
    tp(m, "4.10.7", None,
       [("t", "Which tool runs scripts in Excel on the web: VBA or Office Scripts?", "Office Scripts", None, None, ("office script",)),
        ("t", "Which tool runs a script on a schedule without opening Excel?", "Power Automate")])
    tp(m, "4.10.8", None,
       [("o", "Write the documentation log (name, purpose, owner, test date, what to do if it fails) for the macro of 4.10.6.",
         "Name: Refresh_And_Export_Report. Purpose: monthly refresh and PDF of the Mercyland report. Owner: reporting analyst. "
         "Tested: 2024-01-15. If it fails: run Refresh All by hand, check the data source path, export the PDF manually.")])
    m.finish()
    audit_sheet(b, D)


# =====================================================================================================
def m411(b, D):
    S = D.S
    m = ModuleSheet(b, "4.11 AI Co-pilot", "4.11  AI Co-pilots for Analysts",
                    "Tutor walk-through after Module 4.11: prompts with a schema, AI-drafted formulas and SQL, and checking them against the data",
                    "4.11.1 - 4.11.10")
    m.dataset([("tSal", "Sheet Data_Sal (A1)", "The formula prompts: average salary by gender and seniority."),
               ("tOrders", "Sheet Data_Orders (A1)", "SuperMart Online orders = the Week 3 transactions table: the control total for the AI's SQL."),
               ("tDash25 / tStores", "Sheets Data_Dash25, Data_Stores (A1)", "Checking an AI-written narrative.")])
    fs, fall = F14.FS, F14.FALL
    tp(m, "4.11.1", None,
       [("t", "Can an AI assistant tell you SuperMart's 2025 revenue without seeing your data? (Yes or No)", "No"),
        ("t", "Should you paste patient records into a public AI tool? (Yes or No)", "No")])
    AV = 'AVERAGEIFS(tSal[Annual_Salary_NGN],tSal[Gender],"{g}",tSal[Seniority],"{s}")'
    tp(m, "4.11.2",
       [("The prompt's answer, tested", "=" + AV.format(g="Female", s="Senior"), f(fs), N, "Always run the AI's formula on the real table.")],
       [("n", "Change the prompt to male Senior respondents. What does the formula return?", "=" + AV.format(g="Male", s="Senior"),
         f(S[(S.Gender == "Male") & (S.Seniority == "Senior")].Annual_Salary_NGN.mean()), N, "Male Senior average")])
    tp(m, "4.11.3",
       [("Columns in tSal (what the schema lists)", "=COLUMNS(tSal)", S.shape[1], "int", "Names and types only."),
        ("Rows in tSal (never pasted)", "=ROWS(tSal)", len(S), "int", ""),
        ("Schema line for the prompt", '="tSal("&TEXTJOIN(", ",TRUE,tSal[#Headers])&")"', "tSal(" + ", ".join(S.columns) + ")", "gen",
         "Paste this, not the data.")],
       [("n", "How many columns would the schema of tHotel list?", "=COLUMNS(tHotel)", D.H.shape[1], "int", "Columns in tHotel")])
    tp(m, "4.11.4",
       [("AI formula 1: female Senior average", "=" + AV.format(g="Female", s="Senior"), f(fs), N, ""),
        ("AI formula 2: female average (all levels)", '=AVERAGEIF(tSal[Gender],"Female",tSal[Annual_Salary_NGN])', f(fall), N, "")],
       [("n", "Test an AI formula for female Executive respondents. What does it return?", "=" + AV.format(g="Female", s="Executive"),
         f(S[(S.Gender == "Female") & (S.Seniority == "Executive")].Annual_Salary_NGN.mean()), N, "Female Executive average")])
    air = float(F14.AIR.sum())
    ctrl = float(D.O.total_price.sum())
    tp(m, "4.11.5",
       [("Control total: all SuperMart Online orders", "=SUM(tOrders[total_price])", f(ctrl), N, "Equals SUM(total_price) FROM transactions."),
        ("Sum of the AI query's zone revenues (from the slide)", f"={air:.0f}", f(air), N, "Typed: the query's output."),
        ("AI total ÷ control total", f"={air:.0f}/SUM(tOrders[total_price])", f(air / ctrl), "dec1", "Joining deliveries on customer_id repeats rows.")],
       [("t", "Does the AI query pass the control total check? (pass or fail)", "fail")],
       code=True)
    st = D.STO.set_index("Store_Name").join(D.D25.groupby("Store_Name").Sales.sum())
    st["Att"] = st.Sales / st.Annual_Target_2025
    low = st.Att.idxmin()
    T = {}

    def att(mm):
        df = pd.DataFrame({"Store": st.index})
        calc = {"Attainment": ("=SUMIF(tDash25[Store_Name],[@Store],tDash25[Sales])/XLOOKUP([@Store],tStores[Store_Name],tStores[Annual_Target_2025])", "pct",
                               [f(x) for x in st.Att])}
        mm.h("tAtt: store attainment (the table the AI's formula reads)")
        T["a"] = mm.s.table(df, "tAtt", calc=calc, widths={"Store": 24, "Attainment": 11})
        mm.s.r += 1
    tp(m, "4.11.6",
       [("The formula the AI explained", "=INDEX(tAtt[Store],MATCH(MIN(tAtt[Attainment]),tAtt[Attainment],0))", low, "gen", "Lowest attainment store."),
        ("Step 1: MIN(Attainment)", "=MIN(tAtt[Attainment])", f(st.Att.min()), "pct", ""),
        ("Step 2: MATCH(...,0)", "=MATCH(MIN(tAtt[Attainment]),tAtt[Attainment],0)", int(np.argmin(st.Att.values)) + 1, "int", "Row position in the table.")],
       [("t", "Which store has the highest attainment?", st.Att.idxmax(), "=INDEX(tAtt[Store],MATCH(MAX(tAtt[Attainment]),tAtt[Attainment],0))", "Highest attainment")],
       extra=att)
    tp(m, "4.11.7",
       [("Draft says 7 of 12 missed target", '=COUNTIF(tAtt[Attainment],"<1")', int((st.Att < 1).sum()), "int", "Matches."),
        ("Draft says sales were ₦733.2m", "=ROUND(SUM(tDash25[Sales])/1000000,1)", round(float(D.D25.Sales.sum()) / 1e6, 1), "dec1", "Matches.")],
       [("n", "The draft says the lowest store reached 81% of target. What does the table say?", "=MIN(tAtt[Attainment])", f(st.Att.min()), "pct", "Lowest attainment")])
    tp(m, "4.11.8", None,
       [("t", "Can you paste Patient_No values into a public AI tool? (Yes or No)", "No"),
        ("t", "What can you paste instead: the schema or the rows?", "schema", None, None, ("the schema",))])
    tp(m, "4.11.9",
       [("Does tSal have a column called Salary?", '=IF(ISNUMBER(MATCH("Salary",tSal[#Headers],0)),"Yes","No")', "No", "gen", "An AI that writes tSal[Salary] invented it.")],
       [("t", "An AI formula uses SUMIFSX. What error will Excel show?", "#NAME?", None, None, ("#NAME", "name"))])
    tp(m, "4.11.10",
       [("AI total ÷ control total", f"={air:.0f}/SUM(tOrders[total_price])", f(air / ctrl), "dec1", "Step 4 (verify) fails: discard and log.")],
       [("o", "List the five steps of the AI-assisted workflow.", "1 Schema prompt, 2 Ask, 3 Run, 4 Verify against a control total, 5 Log what was used or discarded.")])
    m.finish()


# =====================================================================================================
def m412(b, D):
    m = ModuleSheet(b, "4.12 AI Verification", "4.12  Verifying AI Output",
                    "Tutor walk-through after Module 4.12: check six AI claims with live formulas, the control total and a sign-off", "4.12.1 - 4.12.8")
    m.dataset([("tClaims", "On this sheet", "Six claims an AI assistant made; Verified is a live formula on the bootcamp data."),
               ("tDash25, tSal, tLOS, tHotel, tCV", "Data_ sheets (A1)", "The tables each claim is checked against.")])
    VER = ["=SUM(tDash25[Sales])", '=(SUMIF(tDash25[Store_ID],"ST01",tDash25[Sales])+SUMIF(tDash25[Store_ID],"ST02",tDash25[Sales]))/SUM(tDash25[Sales])',
           "=MEDIAN(tSal[Annual_Salary_NGN])", '=COUNTIF(tLOS[Readmitted_30d],"Yes")/ROWS(tLOS)', "=AVERAGE(tHotel[Is_Cancelled])",
           "=CORREL(tCV[Lecture_Attendance_Pct],tCV[Final_Score])"]
    cl = F14.CLAIMS
    tol = [c[3] * (abs(c[1]) if abs(c[1]) > 1 else 1) for c in cl]
    ok = [abs(c[2] - c[1]) <= t for c, t in zip(cl, tol)]
    T = {}

    def claims(mm):
        df = pd.DataFrame({"Claim": [c[0] for c in cl], "Claimed": [c[1] for c in cl], "Tolerance": tol, "Verified": [f(c[2]) for c in cl]})
        calc = {"Verdict": ('=IF(ABS([@Verified]-[@Claimed])<=[@Tolerance],"VERIFIED","WRONG")', "gen", ["VERIFIED" if o else "WRONG" for o in ok])}
        mm.h("tClaims: each Verified cell is its own formula on the data (click one to read it)")
        t = mm.s.table(df, "tClaims", fmts={"Claimed": "#,##0.000", "Tolerance": "#,##0.000", "Verified": "#,##0.000"}, calc=calc,
                       widths={"Claim": 40, "Claimed": 14, "Tolerance": 12, "Verified": 14, "Verdict": 11})
        for i, fx in enumerate(VER):
            mm.s.ws.write_formula(t.r0 + 1 + i, t.ci("Verified"), fx, b.nf("#,##0.000"), f(cl[i][2]))
        T["c"] = t
        mm.s.r += 1
    tp(m, "4.12.1",
       [("Claims checked", "=ROWS(tClaims)", len(cl), "int", ""), ("Claims that are WRONG", '=COUNTIF(tClaims[Verdict],"WRONG")', len(cl) - sum(ok), "int",
                                                                    "Right and wrong claims look identical until checked.")],
       [("n", "What is the real median salary?", "=MEDIAN(tSal[Annual_Salary_NGN])", f(cl[2][2]), N, "Real median salary")], extra=claims)
    tp(m, "4.12.2",
       [("Claims VERIFIED", '=COUNTIF(tClaims[Verdict],"VERIFIED")', sum(ok), "int", "Within tolerance of the recomputed value.")],
       [("n", "What is the real Atlantic Pearl cancellation rate?", "=AVERAGE(tHotel[Is_Cancelled])", f(cl[4][2]), "pct", "Real cancellation rate"),
        ("t", "Should a rounded-but-true claim such as '₦733.2m' pass? (Yes or No)", "Yes")],
       steps=["Recompute every number the AI states, on the source table.", "Set a tolerance for rounding.", "Mark each claim VERIFIED or WRONG; correct or remove the wrong ones."])
    zn = D.D25.groupby("Zone").Sales.sum()
    aiz = float(zn.drop("South-East").sum())
    tp(m, "4.12.3",
       [("Control total: FY2025 sales (ledger)", "=SUM(tDash25[Sales])", f(zn.sum()), N, ""),
        ("Sum of the AI zone table (South-East missing)", '=SUM(tDash25[Sales])-SUMIF(tDash25[Zone],"South-East",tDash25[Sales])', f(aiz), N, ""),
        ("Difference", '=SUMIF(tDash25[Zone],"South-East",tDash25[Sales])', f(zn["South-East"]), N, "Not zero: the AI table is incomplete.")],
       [("t", "The AI table's total differs from the ledger. Do you publish it? (Yes or No)", "No")])
    tp(m, "4.12.4", None,
       [("t", "An AI says '22% of bookings cancel'. Name the warning sign: a round number or a precise number?", "round number", None, None, ("round",)),
        ("n", "How far off is 22% from the real rate (real minus claimed, as %)?", "=AVERAGE(tHotel[Is_Cancelled])-0.22", f(cl[4][2] - 0.22), "pct", "Gap from 22%")])
    tp(m, "4.12.5", None,
       [("o", "Write one provenance-log row for the AI's zone revenue SQL.",
         "2026-01-12 10:40 · prompt: zone revenue SQL · check: 56,725 rows, ₦7.19bn vs control ₦1.39bn · outcome: discarded, join rewritten.")])
    tp(m, "4.12.6", None,
       [("t", "The AI invents a column name. Patch or discard?", "discard"), ("t", "A chart title has a typo. Patch or discard?", "patch")])
    tp(m, "4.12.7", None,
       [("o", "Write the sign-off block for the claims check.", "Prepared by A. Analyst; checks: 6 claims recomputed, control total matched; reviewed by B. Reviewer on 2026-01-13.")])
    tp(m, "4.12.8",
       [("Claims recomputed", "=ROWS(tClaims)", len(cl), "int", "Evidence: the Verdict column."),
        ("Control total matches the online orders?", '=IF(ROUND(SUM(tOrders[total_price]),0)=ROUND(SUM(tOrders[total_price]),0),"Yes","No")', "Yes", "gen", "")],
       [("n", "How many claims passed?", '=COUNTIF(tClaims[Verdict],"VERIFIED")', sum(ok), "int", "Claims passed")])
    m.finish()


# =====================================================================================================
def m413(b, D):
    S, L = D.S, D.LOS
    m = ModuleSheet(b, "4.13 Data Ethics", "4.13  Data Ethics, PII and Health Data",
                    "Tutor walk-through after Module 4.13: classify columns, minimise an extract, suppress small cells and test for bias", "4.13.1 - 4.13.9")
    m.dataset([("tLOS", "Sheet Data_LOS (A1)", "Mercyland 2023 admissions: health data (no names or patient numbers in this extract)."),
               ("tSal", "Sheet Data_Sal (A1)", "The pay-gap fairness test.")])
    tp(m, "4.13.1", None,
       [("t", "Who is harmed if the Kingsway wellbeing survey leaks: students or the university's suppliers?", "students"),
        ("t", "Which dataset holds health details: Mercyland admissions or SuperMart orders?", "Mercyland admissions", None, None, ("Mercyland", "admissions"))])
    tp(m, "4.13.2", None,
       [("t", "Is email_address a direct identifier or a quasi-identifier?", "direct", None, None, ("direct identifier",)),
        ("t", "Is date_of_birth a direct identifier or a quasi-identifier?", "quasi", None, None, ("quasi-identifier", "quasi identifier"))])
    tp(m, "4.13.3",
       [("Columns in the full Mercyland extract", f"={len(F14.LC_COLS)}", len(F14.LC_COLS), "int", "Typed: the CSV header in 00_Core_Datasets."),
        ("Columns in tLOS (names and patient numbers removed)", "=COLUMNS(tLOS)", L.shape[1], "int", "")],
       [("t", "Is Diagnosis_Group personal health data? (Yes or No)", "Yes")])
    dep = L.groupby("Department").Readmitted_30d.apply(lambda x: (x == "Yes").mean())
    tp(m, "4.13.4",
       [("Readmission rate, Cardiology (3 columns needed)", '=COUNTIFS(tLOS[Department],"Cardiology",tLOS[Readmitted_30d],"Yes")/COUNTIF(tLOS[Department],"Cardiology")',
         f(dep["Cardiology"]), "pct", "Department, Readmitted_30d and the year: nothing else."),
        ("Columns to drop from the full extract", f"={len(F14.LC_COLS)}-3", len(F14.LC_COLS) - 3, "int", "")],
       [("n", "What is the readmission rate for Paediatrics?", '=COUNTIFS(tLOS[Department],"Paediatrics",tLOS[Readmitted_30d],"Yes")/COUNTIF(tLOS[Department],"Paediatrics")',
         f(dep["Paediatrics"]), "pct", "Paediatrics readmission")])
    old = int((L.Age >= 95).sum())
    oldre = int(((L.Age >= 95) & (L.Readmitted_30d == "Yes")).sum())
    tp(m, "4.13.5",
       [("Admissions aged 95+", '=COUNTIF(tLOS[Age],">=95")', old, "int", "A small group: easy to re-identify."),
        ("Readmissions aged 95+", '=COUNTIFS(tLOS[Age],">=95",tLOS[Readmitted_30d],"Yes")', oldre, "int", "Below 5: suppress or merge into a wider band.")],
       [("n", "How many admissions were aged 80 or over (a safer band)?", '=COUNTIF(tLOS[Age],">=80")', int((L.Age >= 80).sum()), "int", "Admissions 80+"),
        ("t", "A table cell counts 3 patients. Publish it or suppress it?", "suppress")])
    tp(m, "4.13.6", None,
       [("t", "SuperMart wants to email marketing to customers. Is the purchase contract enough, or is consent needed?", "consent", None, None, ("consent needed",))])
    s2 = S[S.Gender.isin(["Male", "Female"])]
    gaps = 1 - s2[s2.Gender == "Female"].groupby("Seniority").Annual_Salary_NGN.mean() / s2[s2.Gender == "Male"].groupby("Seniority").Annual_Salary_NGN.mean()
    G = '1-AVERAGEIFS(tSal[Annual_Salary_NGN],tSal[Gender],"Female"{x})/AVERAGEIFS(tSal[Annual_Salary_NGN],tSal[Gender],"Male"{x})'
    tp(m, "4.13.7",
       [("Raw gap, all levels", "=" + G.format(x=""), f(F14.GAP_ALL), "pct", "Headline figure."),
        ("Gap within Mid level", "=" + G.format(x=',tSal[Seniority],"Mid"'), f(gaps["Mid"]), "pct", "Investigate."),
        ("'Other' gender respondents", '=COUNTIF(tSal[Gender],"Other")', F14.N_OTHER, "int", "Too few to report.")],
       [("n", "What is the gap within the Senior level?", "=" + G.format(x=',tSal[Seniority],"Senior"'), f(gaps["Senior"]), "pct", "Senior gap")])
    tp(m, "4.13.8", None,
       [("t", "Can a training extract with patient data be sent as an email attachment? (Yes or No)", "No")])
    tp(m, "4.13.9", None,
       [("t", "Who approves a new research use of Mercyland data: the analyst or the ethics committee?", "ethics committee", None, None, ("the ethics committee",))])
    m.finish()


# =====================================================================================================
def m414(b, D):
    m = ModuleSheet(b, "4.14 Benchmark Lab", "4.14  Practical Lab: Internal Benchmarks",
                    "Tutor walk-through after Module 4.14: sales per m² benchmark, PDF bulletin check, definitions, intervals and the memo",
                    "4.14.1 - 4.14.5")
    m.dataset([("tDash25 / tStores", "Sheets Data_Dash25, Data_Stores (A1)", "FY2025 sales and floor areas."),
               ("tBulletin", "On this sheet", "SuperMart_Monthly_Sales_Bulletin_2025.pdf (Week2_Data_Files), typed in, against the ledger."),
               ("tLOS, tMonthly, tOrders", "Data_ sheets (A1)", "Readmission intervals and the definitions check."),
               ("tSPM, tCI", "On this sheet", "Benchmark and interval tables: every cell is a formula.")])
    st = D.STO.set_index("Store_Name").join(D.D25.groupby("Store_Name").Sales.sum())
    st["SPM"] = st.Sales / st.Floor_Area_sqm
    med = float(st.SPM.median())
    st = st.sort_values("SPM", ascending=False)
    T = {}

    def spm(mm):
        df = pd.DataFrame({"Store_Name": st.index})
        calc = {"Sales": ("=SUMIF(tDash25[Store_Name],[@Store_Name],tDash25[Sales])", N, [f(x) for x in st.Sales]),
                "Floor_Area": ("=XLOOKUP([@Store_Name],tStores[Store_Name],tStores[Floor_Area_sqm])", "int", [int(x) for x in st.Floor_Area_sqm]),
                "Sales_per_m2": ("=[@Sales]/[@Floor_Area]", N, [f(x) for x in st.SPM]),
                "vs_Median": ("=[@Sales_per_m2]/MEDIAN([Sales_per_m2])-1", "pct", [f(x / med - 1) for x in st.SPM])}
        mm.h("tSPM: FY2025 sales per m², highest first")
        T["s"] = mm.s.table(df, "tSPM", calc=calc, widths={"Store_Name": 24, "Sales": 14, "Floor_Area": 11, "Sales_per_m2": 13, "vs_Median": 11})
        mm.s.r += 1
        mm.s.chart("bar", [{"name": "vs median", "categories": T["s"].ref("Store_Name"), "values": T["s"].ref("vs_Median"),
                            "points": [{"fill": {"color": "#B23A48"}} if v < 0 else {"fill": {"color": "#2A9D8F"}} for v in st.SPM / med - 1]}],
                   f"{st.index[-1]} sells {1 - st.SPM.iloc[-1] / med:.0%} less per m² than the median store", "", "vs chain median", legend=False,
                   y_fmt="0%", rows=16, size=(640, 290))
    tp(m, "4.14.1",
       [("Chain median sales per m²", "=MEDIAN(tSPM[Sales_per_m2])", f(med), N, "The benchmark: robust to the two big Lagos stores."),
        ("Store furthest below the median", "=INDEX(tSPM[Store_Name],MATCH(MIN(tSPM[vs_Median]),tSPM[vs_Median],0))", st.index[-1], "gen", "")],
       [("n", "What are Lagos - Ikeja's sales per m²?", '=XLOOKUP("Lagos - Ikeja",tSPM[Store_Name],tSPM[Sales_per_m2])', f(st.SPM["Lagos - Ikeja"]), N, "Ikeja per m²"),
        ("n", "How many stores are below the median?", '=COUNTIF(tSPM[vs_Median],"<0")', int((st.SPM < med).sum()), "int", "Stores below median")], extra=spm)

    bul = F14.BUL.copy()

    def bulletin(mm):
        df = pd.DataFrame({"Store_Name": bul.index, "Bulletin_FY_m": bul.fy.values, "Months_Added_m": bul.months.values})
        calc = {"Ledger_m": ("=SUMIF(tDash25[Store_Name],[@Store_Name],tDash25[Sales])/1000000", "dec2", [f(x) for x in bul.ledger]),
                "Diff_m": ("=[@Bulletin_FY_m]-[@Ledger_m]", "dec2", [f(a - c) for a, c in zip(bul.fy, bul.ledger)]),
                "Status": ('=IF(ABS([@Diff_m])<=0.05,"OK","CHECK")', "gen", ["OK" if abs(a - c) <= 0.05 else "CHECK" for a, c in zip(bul.fy, bul.ledger)])}
        mm.h("tBulletin: the PDF's FY column and its 12 months added up (typed from the PDF), against the ledger (₦m)")
        T["b"] = mm.s.table(df, "tBulletin", fmts={"Bulletin_FY_m": "dec1", "Months_Added_m": "dec1"}, calc=calc,
                            widths={"Store_Name": 24, "Bulletin_FY_m": 13, "Months_Added_m": 15, "Ledger_m": 11, "Diff_m": 9, "Status": 9})
        mm.s.r += 1
    nok = int(sum(abs(a - c) <= 0.05 for a, c in zip(bul.fy, bul.ledger)))
    tp(m, "4.14.2",
       [("Stores within ±₦0.05m of the ledger", '=COUNTIF(tBulletin[Status],"OK")', nok, "int", "The PDF rounds to ₦0.1m."),
        ("Largest gap between the FY column and the months added", "=MAX(MAX(tBulletin[Months_Added_m]-tBulletin[Bulletin_FY_m]),-MIN(tBulletin[Months_Added_m]-tBulletin[Bulletin_FY_m]))",
         f(bul.mdiff.abs().max()), "dec1", "Rounding each month adds up.")],
       [("n", "What is the PDF's FY figure for Lagos - Ikeja (₦m)?", '=XLOOKUP("Lagos - Ikeja",tBulletin[Store_Name],tBulletin[Bulletin_FY_m])', f(bul.fy["Lagos - Ikeja"]),
         "dec1", "Ikeja in the PDF")],
       do=[("Open Week2_Data_Files/SuperMart_Monthly_Sales_Bulletin_2025.pdf. Data › Get Data › From File › From PDF, load the store table.",
            "Twelve store rows with monthly figures and an FY column, matching tBulletin.")], extra=bulletin)

    O = D.O
    o25 = O[O.order_date.dt.year == 2025]
    gross, net = float(o25.total_price.sum()), float(o25[o25.order_status.isin(["Delivered", "Completed"])].total_price.sum())
    MO = D.MON
    apr = float(MO[(MO.Month >= "2024-04-01") & (MO.Month <= "2025-03-01")].Sales.sum())
    W = 'tOrders[order_date],">="&DATE(2025,1,1),tOrders[order_date],"<="&DATE(2025,12,31)'
    tp(m, "4.14.3",
       [("SuperMart Stores FY2025 (Jan-Dec)", "=SUM(tDash25[Sales])", f(D.D25.Sales.sum()), N, ""),
        ("SuperMart Stores Apr 2024 - Mar 2025", '=SUMIFS(tMonthly[Sales],tMonthly[Month],">="&DATE(2024,4,1),tMonthly[Month],"<="&DATE(2025,3,1))', f(apr), N,
         "Same company, April year-end."),
        ("SuperMart Online 2025 gross (all orders)", f"=SUMIFS(tOrders[total_price],{W})", f(gross), N, "")],
       [("n", "What is SuperMart Online's 2025 net revenue (Delivered + Completed)?",
         f'=SUMIFS(tOrders[total_price],{W},tOrders[order_status],"Delivered")+SUMIFS(tOrders[total_price],{W},tOrders[order_status],"Completed")', f(net), N,
         "Online 2025 net")])

    RE = F14.RE

    def ci(mm):
        df = pd.DataFrame({"Department": RE.index})
        calc = {"Admissions": ("=COUNTIF(tLOS[Department],[@Department])", "int", [int(x) for x in RE.n]),
                "Rate": ('=COUNTIFS(tLOS[Department],[@Department],tLOS[Readmitted_30d],"Yes")/[@Admissions]', "pct", [f(x) for x in RE.p]),
                "MoE_95": ("=1.96*SQRT([@Rate]*(1-[@Rate])/[@Admissions])", "pct", [f(x) for x in RE.moe]),
                "Lower": ("=[@Rate]-[@MoE_95]", "pct", [f(a - c) for a, c in zip(RE.p, RE.moe)]),
                "Upper": ("=[@Rate]+[@MoE_95]", "pct", [f(a + c) for a, c in zip(RE.p, RE.moe)])}
        mm.h("tCI: 2023 readmission rate by department with a 95% interval")
        mm.s.table(df, "tCI", calc=calc, widths={"Department": 24, "Admissions": 11, "Rate": 9, "MoE_95": 9, "Lower": 9, "Upper": 9})
        mm.s.r += 1
    sm = RE.loc["Surgery"]
    im = RE.loc["Internal Medicine"]
    tp(m, "4.14.4",
       [("Cardiology lower bound", '=XLOOKUP("Cardiology",tCI[Department],tCI[Lower])', f(RE.p["Cardiology"] - RE.moe["Cardiology"]), "pct",
         "Above every other upper bound: Cardiology is genuinely highest."),
        ("Do Surgery and Internal Medicine overlap?", '=IF(AND(XLOOKUP("Surgery",tCI[Department],tCI[Lower])<=XLOOKUP("Internal Medicine",tCI[Department],tCI[Upper]),'
         'XLOOKUP("Internal Medicine",tCI[Department],tCI[Lower])<=XLOOKUP("Surgery",tCI[Department],tCI[Upper])),"Overlap","Separate")',
         "Overlap" if (sm.p - sm.moe <= im.p + im.moe and im.p - im.moe <= sm.p + sm.moe) else "Separate", "gen", "Do not rank them.")],
       [("n", "What is the 95% margin of error for Obstetrics & Gynaecology?", '=XLOOKUP("Obstetrics & Gynaecology",tCI[Department],tCI[MoE_95])',
         f(RE.moe["Obstetrics & Gynaecology"]), "pct", "O&G margin")], extra=ci)
    tp(m, "4.14.5",
       [("Lowest sales per m²", "=MIN(tSPM[Sales_per_m2])", f(st.SPM.min()), N, ""), ("Its gap to the median", "=MIN(tSPM[vs_Median])", f(st.SPM.min() / med - 1), "pct", "")],
       [("o", "Write the four-part memo (benchmark, finding, uncertainty, recommendation) in four sentences.",
         f"Benchmark: chain median sales per m², FY2025 ledger reconciled to the PDF within ±₦0.05m. Finding: {st.index[-1]} is lowest at "
         f"₦{st.SPM.min():,.0f} per m², {1 - st.SPM.min() / med:.0%} below the median. Uncertainty: one year of data, store age not adjusted, figures unaudited. "
         "Recommendation: a site review for the bottom three stores before allocating the refurbishment budget.")])
    m.finish()


assert DEF  # slide copy loaded
