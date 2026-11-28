"""Capstone Section Examples (Edition 3.1): Week 5, the project week. Everything here uses Weeks 1-4 skills only.

    P.1 - P.2   the project lifecycle and source governance (ideas)
    P.3         Capstone 1 SuperMart: growth vs margin with SUMIFS on tSM
    P.4         Capstone 2 Kingsway: a fixed progression rule with COUNTIFS on tStudents
    P.5         Capstone 3 Mercyland: occupancy, length of stay and readmission on tAdmissions / tCensus
    P.6 - P.9   QA, the executive defence, portfolio and CV, interviews and the roadmap (ideas)
"""
from __future__ import annotations

import pandas as pd

import w5_data as W

GOOD = '">="&62'
ATT = '">="&65'


def build(b):
    sm = W.table("tSM")
    sm["Order_Date"] = pd.to_datetime(sm.Order_Date)
    sm["Year"] = sm.Order_Date.dt.year
    sm["Quarter"] = "Q" + sm.Order_Date.dt.quarter.astype(str)
    cv = W.table("tCV")
    adm = W.table("tAdm23")
    adm["Admit_Date"] = pd.to_datetime(adm.Admit_Date)
    cen = W.table("tCensus")
    cen = cen[cen.year == 2023].reset_index(drop=True)
    cen["census_date"] = pd.to_datetime(cen.census_date)

    def mod(code, sheet, title, learned):
        return b.module(code, sheet, title, learned)

    p1 = mod("P.1", "P.1 Project Lifecycle", "The Seven-Stage Project Lifecycle",
             "Scope, acquire, clean, model, visualise, verify, hand over; quality gates, time-boxing, stakeholders "
             "and the post-project review.")
    p2 = mod("P.2", "P.2 Source Register", "Source Governance, Citation and Capstone Delivery",
             "The source register, before-and-after states, citation, licence and privacy, reproducibility and the "
             "evidence-to-decision handoff.")
    p3 = mod("P.3", "P.3 Capstone1 SuperMart", "Capstone 1: SuperMart Retail Analytics",
             "Brief and success criteria, data inventory, cleaning, the sales model, basket and margin analysis, slow "
             "movers, the dashboard, the insight memo and verification.")
    p4 = mod("P.4", "P.4 Capstone2 Kingsway", "Capstone 2: Kingsway University Progression",
             "A fixed progression rule, where at-risk students concentrate, the at-risk list, quantifying an "
             "intervention, the dashboard and the Senate paper.")
    p5 = mod("P.5", "P.5 Capstone3 Mercyland", "Capstone 3: Mercyland Hospital Operations",
             "Occupancy measured properly, length of stay, readmissions, one escalation rule, the hospital dashboard "
             "and the clinical board pack.")
    p6 = mod("P.6", "P.6 Workbook QA", "Workbook QA and Stress Testing",
             "The 00_README standard, reconciliation, stress, refresh and edge-case tests, performance and sign-off.")
    p7 = mod("P.7", "P.7 Executive Defence", "The Ten-Minute Executive Defence",
             "Structuring ten minutes, the live demo, panel questions, challenge, plain-language explanations and "
             "closing for a decision.")
    p8 = mod("P.8", "P.8 Portfolio & CV", "Portfolio, Anonymisation and CV",
             "Project write-ups, anonymising artefacts, packaging files, CV bullets, describing tools honestly and LinkedIn.")
    p9 = mod("P.9", "P.9 Interviews & Roadmap", "Technical Interviews and the Roadmap Beyond Excel",
             "Interview formats, the live Excel test, SQL questions, STAR answers, and the roadmap to Power BI, SQL and Python.")

    r_sm = b.data("SM", sm, "One SuperMart sales line, 2022-2025 (Year and Quarter added)",
                  fmts={"Order_Date": "date", "Sales": "naira", "Profit": "naira", "Discount": "pct0"}, used_in="P.3")
    r_st = b.data("Students", cv, "One Kingsway University student in the 2024/25 cohort", used_in="P.4")
    r_ad = b.data("Admissions", adm, "One Mercyland Hospital admission in 2023", fmts={"Admit_Date": "date"}, used_in="P.5")
    r_ce = b.data("Census", cen, "One day of the Mercyland bed census in 2023 (200 open beds)",
                  fmts={"census_date": "date", "occupancy_pct": "pct"}, used_in="P.5")
    p3.set_refs([r_sm])
    p4.set_refs([r_st])
    p5.set_refs([r_ad, r_ce])

    # ================================================================== P.1
    m = p1
    m.part("The seven stages")
    stages = ["Scoping", "Acquisition", "Cleaning", "Modelling", "Visualisation", "Verification", "Handover"]
    m.pick("Agreeing the business question, the audience and what 'done' means.", stages[:4], "Scoping", "Stage 1.")
    m.pick("Checking totals against the source and testing every slicer.", stages[3:], "Verification", "Stage 6.")
    m.pick("Writing the README and showing the owner how to refresh.", stages[3:], "Handover", "Stage 7.")
    m.mc("A quality gate is ...", ["a check you must pass before moving to the next stage", "a password",
                                   "a chart type", "a PivotTable"], "A", "e.g. row counts reconcile before modelling.")
    m.yn("Should you start building charts before the question is agreed?", "No", "Scope first, or you build the wrong thing.")
    m.tf("A post-project review records what went well and what to change next time.", True, "Stage after handover.")

    # ================================================================== P.2
    m = p2
    m.part("Source register")
    m.mc("A source register lists ...", ["every data source: file name, owner, date received, rows, licence",
                                         "only the charts", "the passwords", "the slides"], "A",
         "So anyone can trace a number back to its source.")
    m.yn("Should you keep an untouched copy of the raw file before cleaning?", "Yes", "Preserve the before state.")
    m.mc("Which is a correct citation for a public dataset?", ["'the internet'",
                                                               "Publisher, dataset name, year, URL and date accessed",
                                                               "a screenshot", "nothing"], "B", "Cite so others can find it.")
    m.tf("Reproducible work means another analyst can follow your steps and get the same numbers.", True,
         "Document steps; use Power Query rather than manual edits.")

    # ================================================================== P.3
    m = p3
    s24 = float(sm.loc[sm.Year == 2024, "Sales"].sum())
    s25 = float(sm.loc[sm.Year == 2025, "Sales"].sum())
    p24 = float(sm.loc[sm.Year == 2024, "Profit"].sum())
    p25 = float(sm.loc[sm.Year == 2025, "Profit"].sum())
    m.example("the headline numbers", ["Each measure is one SUMIFS on tSM, filtered by Year.",
                                       "Growth and margin are simple divisions of those totals."],
              [("FY2025 sales", "=SUMIFS(tSM[Sales],tSM[Year],2025)", "=SUMIFS(tSM[Sales],tSM[Year],2025)", s25, "naira"),
               ("FY2025 profit", "=SUMIFS(tSM[Profit],tSM[Year],2025)", "=SUMIFS(tSM[Profit],tSM[Year],2025)", p25, "naira")])
    m.part("Part A · Brief")
    m.mc("The board's question for Capstone 1 is ...", ["Why is margin falling while sales grow?", "How many stores exist?",
                                                       "What colour should the dashboard be?", "Who is the CEO?"], "A",
         "Growth with a shrinking margin.")
    m.part("Part B · Work it out on tSM")
    m.num("What were FY2024 sales?", "=SUMIFS(tSM[Sales],tSM[Year],2024)", s24, "SUMIFS by Year.", fmt="naira")
    m.num("Sales growth from FY2024 to FY2025 (as a %)?",
          "=SUMIFS(tSM[Sales],tSM[Year],2025)/SUMIFS(tSM[Sales],tSM[Year],2024)-1", s25 / s24 - 1, "2025 ÷ 2024 − 1.",
          fmt="pct2", tol=0.0006)
    m.num("Profit margin FY2024 (profit ÷ sales, as a %)?",
          "=SUMIFS(tSM[Profit],tSM[Year],2024)/SUMIFS(tSM[Sales],tSM[Year],2024)", p24 / s24, "Profit ÷ sales.", fmt="pct2",
          tol=0.0006)
    m.num("Profit margin FY2025 (as a %)?",
          "=SUMIFS(tSM[Profit],tSM[Year],2025)/SUMIFS(tSM[Sales],tSM[Year],2025)", p25 / s25, "Profit ÷ sales.", fmt="pct2",
          tol=0.0006)
    m.yn("Did the margin fall from FY2024 to FY2025 even though sales grew?", "Yes",
         f"{p24 / s24:.2%} → {p25 / s25:.2%}.")
    cat = sm[sm.Year == 2025].groupby("Category").Sales.sum().sort_values(ascending=False)
    m.pick("Which category had the highest FY2025 sales?", list(cat.index), cat.index[0],
           "SUMIFS by Category and Year (or a PivotTable).")
    q = sm[sm.Year == 2025].groupby("Quarter").apply(lambda d: d.Profit.sum() / d.Sales.sum())
    m.pick("In which quarter of FY2025 was the margin lowest?", ["Q1", "Q2", "Q3", "Q4"], q.idxmin(),
           "A PivotTable with Quarter in rows, Sum of Profit ÷ Sum of Sales.")
    m.num("How many FY2025 sales lines had a discount of 20% or more?",
          '=COUNTIFS(tSM[Year],2025,tSM[Discount],">=0.2")', int(((sm.Year == 2025) & (sm.Discount >= 0.2)).sum()),
          "COUNTIFS with two conditions.")
    m.mc("Which is a good memo headline?", ["Sales data 2025",
                                            f"Sales grew {s25 / s24 - 1:.1%} but margin fell to {p25 / s25:.1%}: cap discounts",
                                            "Chart 1", "See attached"], "B", "A headline that states the finding and the action.")

    # ================================================================== P.4
    m = p4
    ok = (cv.Final_Score >= 62) & (cv.Lecture_Attendance_Pct >= 65)
    m.info("Rule", "Progressed = Final_Score of 62 or more AND Lecture_Attendance_Pct of 65 or more. "
                   "At risk = everyone else. Target: 90% of the cohort progressed.")
    m.part("Part A · The definition")
    m.mc("Why must the progression rule be fixed BEFORE you calculate?", ["It is faster",
                                                                         "So the result cannot be tuned to look good",
                                                                         "Excel needs it", "It is not needed"], "B",
         "Agree definitions first; document them.")
    m.part("Part B · Work it out on tStudents")
    m.num("How many students are in the cohort?", "=ROWS(tStudents)", len(cv), "One row per student.")
    m.num("How many students progressed?", f"=COUNTIFS(tStudents[Final_Score],{GOOD},tStudents[Lecture_Attendance_Pct],{ATT})",
          int(ok.sum()), "COUNTIFS with both conditions.")
    m.num("Progression % (as a %)?", f"=COUNTIFS(tStudents[Final_Score],{GOOD},tStudents[Lecture_Attendance_Pct],{ATT})"
          "/ROWS(tStudents)", float(ok.mean()), "Progressed ÷ cohort.", fmt="pct")
    m.num("At-risk % (as a %)?", f"=1-COUNTIFS(tStudents[Final_Score],{GOOD},tStudents[Lecture_Attendance_Pct],{ATT})"
          "/ROWS(tStudents)", float(1 - ok.mean()), "1 − progression %.", fmt="pct")
    m.yn("Did the cohort reach the 90% target?", "No", f"{ok.mean():.1%} is below 90%.")
    both = int(((cv.Final_Score < 62) & (cv.Lecture_Attendance_Pct < 65)).sum())
    m.num("How many students are at risk on BOTH score and attendance (highest priority)?",
          '=COUNTIFS(tStudents[Final_Score],"<62",tStudents[Lecture_Attendance_Pct],"<65")', both, "COUNTIFS with two '<' conditions.")
    fr = cv.assign(r=~ok).groupby("Faculty").r.mean().sort_values(ascending=False)
    m.pick("Which faculty has the highest at-risk rate?", list(fr.index), fr.index[0],
           "For each faculty: 1 − progressed ÷ students (a PivotTable or COUNTIFS).")

    # ================================================================== P.5
    m = p5
    rr = float((adm.Readmitted_30d == "Yes").mean())
    m.part("Part A · Ideas")
    m.mc("Bed occupancy % is ...", ["occupied beds ÷ open beds", "admissions ÷ 365", "beds × days", "patients ÷ wards"], "A",
         "Measured each day from the census, then averaged.")
    m.mc("A 30-day readmission is ...", ["a patient admitted again within 30 days of discharge", "any second visit ever",
                                         "a long stay", "a cancelled admission"], "A", "A safety and quality measure.")
    m.yn("Should every report use the same escalation rule (e.g. stay of 14 days or more)?", "Yes",
         "One rule for every report, written in the README.")
    m.part("Part B · Work it out on tAdmissions and tCensus")
    m.num("How many admissions were there in 2023?", "=ROWS(tAdmissions)", len(adm), "One row per admission.")
    m.num("Average bed occupancy in 2023 (as a %)?", "=AVERAGE(tCensus[occupancy_pct])", float(cen.occupancy_pct.mean()),
          "Average of the daily occupancy.", fmt="pct")
    m.num("Average length of stay (days, 1 decimal place)?", "=AVERAGE(tAdmissions[LOS_Days])", float(adm.LOS_Days.mean()),
          "AVERAGE of LOS_Days.", fmt="dec1")
    m.num("30-day readmission rate (as a %)?", '=COUNTIFS(tAdmissions[Readmitted_30d],"Yes")/ROWS(tAdmissions)', rr,
          "Readmitted ÷ admissions.", fmt="pct")
    m.num("How many long stays (LOS_Days of 14 or more)?", '=COUNTIFS(tAdmissions[LOS_Days],">=14")',
          int((adm.LOS_Days >= 14).sum()), "The escalation rule: 14 days or more.")
    dep = adm.assign(r=adm.Readmitted_30d == "Yes").groupby("Department").r.mean().sort_values(ascending=False)
    m.pick("Which department has the highest readmission rate?", list(dep.index), dep.index[0],
           "COUNTIFS by Department and 'Yes', divided by COUNTIFS by Department.")

    # ================================================================== P.6
    m = p6
    m.part("Workbook QA")
    m.mc("A 00_README sheet should hold ...", ["purpose, sources, definitions, refresh steps and owner", "only the logo",
                                               "passwords", "the raw data"], "A", "Anyone can pick up the workbook.")
    m.mc("Reconciliation testing means ...", ["checking totals and row counts against the source", "changing colours",
                                              "deleting old sheets", "adding charts"], "A", "e.g. 3,312 rows and ₦733,215,255.")
    m.yn("Should you test the workbook with messier and larger data before release?", "Yes", "Stress testing.")
    m.tf("After Refresh All, every PivotTable and chart should still show correct totals.", True, "Refresh testing.")

    # ================================================================== P.7
    m = p7
    m.part("The defence")
    m.mc("The first minute of a ten-minute defence should give ...", ["the headline finding and the decision you need",
                                                                     "your data-cleaning steps", "the full SQL",
                                                                     "your CV"], "A", "Answer first, then evidence.")
    m.yn("If a panel member challenges a number, should you show where it comes from (the control total)?", "Yes",
         "Calm, evidence-based answers.")
    m.mc("Explaining a DAX measure to a non-technical panel, you should ...", ["read the formula aloud",
                                                                              "say in plain words what it calculates and why",
                                                                              "skip it", "show the M code"], "B",
         "Plain language.")
    m.tf("A defence should close by asking for a decision.", True, "Closing for a decision.")

    # ================================================================== P.8
    m = p8
    m.part("Portfolio and CV")
    m.mc("Before publishing a capstone in your portfolio you must ...", ["anonymise names, IDs and anything confidential",
                                                                        "add more colours", "delete the README",
                                                                        "rename the sheets"], "A", "Anonymising artefacts.")
    m.mc("Which CV bullet is strongest?", ["Used Excel", "Built a Power Pivot dashboard that showed a 0.7-point margin fall "
                                           "and led to a discount cap", "Good with computers", "Did analysis"], "B",
         "Action + tool + measurable result.")
    m.yn("Should you list a tool on your CV that you have never used?", "No", "Describe tools honestly.")

    # ================================================================== P.9
    m = p9
    m.part("Interviews and the roadmap")
    m.mc("STAR stands for ...", ["Situation, Task, Action, Result", "Start, Test, Analyse, Report", "SQL, Tables, Arrays, Rows",
                                 "Style, Tone, Audience, Response"], "A", "For behavioural questions.")
    m.mc("A common SQL interview question asks you to ...", ["count or total by group with GROUP BY", "draw a chart",
                                                             "write a macro", "format a table"], "A", "Practise Week 3 queries.")
    m.pick("The natural next tool after Excel Power Pivot and DAX", ["Power BI", "Word", "Paint"], "Power BI",
           "Power BI uses the same Power Query and DAX.")
    m.tf("In a live Excel test, explaining your steps out loud helps the interviewer follow your thinking.", True,
         "Explaining your analytical process.")
