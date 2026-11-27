"""Build 04_Student_Projects — project briefs for students."""
import openpyxl
from xl_helpers import *

wb = openpyxl.Workbook()

# ============ 00_README ============
readme_sheet(wb,
    "Student Projects — Capstone Briefs",
    "Four project briefs for the final week. Choose one (or as instructed), work the milestones, and submit the full project pack.",
    ["This workbook contains the project briefs students will work on. Each brief defines the question, the data, the milestones, the deliverables, and the success criteria.",
     "The 03_Instructor_Projects workbook holds the worked solutions — do not open it before completing your own work, as the reconciliation numbers are the test."],
    [["E-commerce (SQLite + CLEAN)", "01_Project_Ecommerce", "Track A", "Full analytics lifecycle", "Recommended for SQL/dashboard focus"],
     ["Student Performance (CLEAN)", "02_Project_StudentPerf", "Track B", "Education analytics", "Recommended for statistics focus"],
     ["Healthcare (CLEAN, anonymised)", "03_Project_Healthcare", "Track C", "Operations analytics", "Recommended for governance/ethics focus"],
     ["Superstore + Hotel (CLEAN)", "04_Project_Forecast", "Track D", "Forecasting & benchmarking", "Recommended for statistics/storytelling focus"]],
    ["1. Read the brief sheet fully before starting; the success criteria are the test you will be marked against.",
     "2. Work the milestones in order — each builds on the previous one.",
     "3. Use only the attached datasets; all numbers must be live formulas, queries or verified values.",
     "4. Submit the full pack: workbook, dashboard, PDF report, deck, QA sign-off, reflection (see each brief's deliverables).",
     "5. Ethics: sensitive datasets (student depression, healthcare) must be used in their anonymised form only."])

# ============ PROJECT A — E-commerce ============
ws = wb.create_sheet("01_Project_Ecommerce")
setup_sheet(ws, "Project Brief A — E-commerce Analytics", "Full analytics lifecycle on the attached E-commerce database.")
r = 4
r = section(ws, r, "The brief", ncols=6)
r = write_table(ws, r, ["Item", "Detail"],
    [["Client", "ETEH Ltd — e-commerce operations team"],
     ["Question", "Where does e-commerce revenue come from, and what holds it back?"],
     ["Data", "E-commerce SQLite database (5 tables: transactions, customers, products, sellers, deliveries) + CLEAN CSVs"],
     ["Your task", "Profile, clean, model, analyse, visualise and report — the complete lifecycle"],
     ["Deliverables", "Analysis workbook · dashboard · 3-page PDF report · 8-slide deck · QA sign-off · reflection"],
     ["Time", "5 working sessions (milestones below)"]],
    widths=[18, 108])
r = section(ws, r, "Milestones (work in this order)", ncols=6)
r = header_row(ws, r, ["Session", "Milestone", "What you must produce", "Done"], widths=[10, 26, 100, 12])
msA = [
 ["1", "Profile & scope", "Data profile: 5 tables, row counts, keys, grain. Scope statement (question, success, constraints).", "☐"],
 ["2", "Clean & model", "Cleaning decisions documented (DIRTY→CLEAN). Star schema with 4 relationships set up and validated.", "☐"],
 ["3", "Analyse", "Verified KPIs: total revenue, lines, cancelled count, revenue by category/method/month, top seller. Reconcile each to SQL.", "☐"],
 ["4", "Visualise", "Dashboard: KPI row, slicers, monthly trend, category bar, payment donut, top-seller bar, detail table.", "☐"],
 ["5", "Report & defend", "3-page PDF report, 8-slide deck, completed QA checklist, 5-line reflection.", "☐"],
]
for m in msA:
    for j, v in enumerate(m, 1):
        c = ws.cell(row=r, column=j, value=v)
        c.font = Font(name="Calibri", size=9); c.border = BORDER; c.alignment = WRAP
    ws.row_dimensions[r].height = 32
    r += 1
r = section(ws, r, "Success criteria (how you are marked)", ncols=6)
r = write_table(ws, r, ["#", "Criterion", "Evidence the assessor will look for"],
    [["1", "Correct numbers", "Every KPI reconciles to a SQL query on the database — total revenue, cancelled count, top category, top seller"],
     ["2", "Reproducible method", "Formulas/queries live in the workbook; a colleague could rebuild the numbers"],
     ["3", "Complete analysis", "Category, payment method, month and seller perspectives all covered"],
     ["4", "Dashboard quality", "One page, one message, slicers work, no hardcoded numbers"],
     ["5", "Story & recommendation", "One quantified recommendation tied to a number in the data"],
     ["6", "Professional pack", "Report and deck labelled, sourced, page-numbered, consistent formatting"]],
    widths=[6, 34, 112])
r = section(ws, r, "Rules & ethics", ncols=6)
for t in ["Attached datasets only — no external data sources",
          "Practice on 500-row extracts; the full database stays with the dataset folder",
          "Every number must be traceable to a query or formula — typed results are not accepted for KPIs",
          "Report cancellations honestly: they are part of the story, not an error to hide"]:
    r = add_check_box(ws, r, t, ncols=6)

# ============ PROJECT B — Student Performance ============
ws = wb.create_sheet("02_Project_StudentPerf")
setup_sheet(ws, "Project Brief B — Student Performance Analytics", "Education analytics on the attached Student Performance dataset.")
r = 4
r = section(ws, r, "The brief", ncols=6)
r = write_table(ws, r, ["Item", "Detail"],
    [["Client", "ETEH Education — academic performance team"],
     ["Question", "Which factors associate with higher exam performance, and where should support go?"],
     ["Data", "StudentPerformanceFactors CLEAN (6,607 records, 27 fields — anonymised)"],
     ["Your task", "Profile the data, run the analyses, build the dashboard and report the findings with one quantified recommendation"],
     ["Deliverables", "Analysis workbook · dashboard · 3-page PDF report · 8-slide deck · QA sign-off · reflection"],
     ["Time", "5 working sessions"]],
    widths=[18, 108])
r = section(ws, r, "Milestones", ncols=6)
r = header_row(ws, r, ["Session", "Milestone", "What you must produce", "Done"], widths=[10, 26, 100, 12])
msB = [
 ["1", "Profile & scope", "Data profile: records, fields, band distribution. Scope statement with success criteria.", "☐"],
 ["2", "Analyse core", "Band counts, average score by band, attendance vs score correlation, city comparison, support crosstab.", "☐"],
 ["3", "Interpret", "Write the findings: what associates with performance, what the crosstabs show, and what does not (correlation ≠ causation).", "☐"],
 ["4", "Visualise", "Dashboard: band distribution, support crosstab, city comparison, attendance correlation."],
 ["5", "Report & defend", "3-page report with one quantified recommendation, 8-slide deck, QA checklist, reflection.", "☐"],
]
for m in msB:
    for j, v in enumerate(m, 1):
        c = ws.cell(row=r, column=j, value=v)
        c.font = Font(name="Calibri", size=9); c.border = BORDER; c.alignment = WRAP
    ws.row_dimensions[r].height = 32
    r += 1
r = section(ws, r, "Success criteria", ncols=6)
r = write_table(ws, r, ["#", "Criterion", "Evidence the assessor will look for"],
    [["1", "Analyses complete", "Band distribution, score by band, correlation, crosstab and city comparison all computed and correct"],
     ["2", "Honest interpretation", "Associations stated as associations; correlation ≠ causation acknowledged"],
     ["3", "Dashboard quality", "One page, one message, live formulas, no hardcoded numbers"],
     ["4", "Recommendation", "One actionable recommendation tied to a number in the data"],
     ["5", "Ethics", "Anonymised data only; no individual student identified; consent/purpose noted"],
     ["6", "Professional pack", "Report and deck labelled, sourced, consistent"]],
    widths=[6, 34, 112])
r = section(ws, r, "Rules & ethics", ncols=6)
for t in ["Anonymised dataset only — never attempt to re-identify students",
          "Report findings as associations; do not claim causation from this data",
          "Attached datasets only",
          "If you use external context to frame findings, label it clearly as context — not data"]:
    r = add_check_box(ws, r, t, ncols=6)

# ============ PROJECT C — Healthcare ============
ws = wb.create_sheet("03_Project_Healthcare")
setup_sheet(ws, "Project Brief C — Healthcare Operations Analytics", "Operations analytics on the anonymised Healthcare dataset.")
r = 4
r = section(ws, r, "The brief", ncols=6)
r = write_table(ws, r, ["Item", "Detail"],
    [["Client", "ETEH Health — hospital operations"],
     ["Question", "Where do admissions and billing concentrate, and what are the operational risks?"],
     ["Data", "healthcare_ANONYMISED CLEAN (55,500 admissions, anonymised)"],
     ["Your task", "Profile, analyse admissions/billing/stay length, build the dashboard and report operational recommendations"],
     ["Deliverables", "Analysis workbook · dashboard · 3-page PDF report · 8-slide deck · QA sign-off · reflection"],
     ["Time", "5 working sessions"]],
    widths=[18, 108])
r = section(ws, r, "Milestones", ncols=6)
r = header_row(ws, r, ["Session", "Milestone", "What you must produce", "Done"], widths=[10, 26, 100, 12])
msC = [
 ["1", "Profile & scope", "Data profile: 55,500 admissions, key fields, data-quality notes (from the cleaning log).", "☐"],
 ["2", "Analyse", "Admissions by hospital/ward, billing distribution (mean vs median), LOS stats, capacity-risk ranking.", "☐"],
 ["3", "Interpret", "Findings: where volume and revenue concentrate, the long-stay tail, and the honest data-quality caveats.", "☐"],
 ["4", "Visualise", "Dashboard: admissions by hospital, billing distribution, LOS stats, capacity-risk ranking."],
 ["5", "Report & defend", "3-page report with operational recommendations, 8-slide deck, QA checklist, reflection.", "☐"],
]
for m in msC:
    for j, v in enumerate(m, 1):
        c = ws.cell(row=r, column=j, value=v)
        c.font = Font(name="Calibri", size=9); c.border = BORDER; c.alignment = WRAP
    ws.row_dimensions[r].height = 32
    r += 1
r = section(ws, r, "Success criteria", ncols=6)
r = write_table(ws, r, ["#", "Criterion", "Evidence the assessor will look for"],
    [["1", "Analyses complete", "Admissions concentration, billing distribution, LOS stats and risk ranking computed and correct"],
     ["2", "Honest reporting", "Data-quality issues stated openly; median used where data is skewed"],
     ["3", "Dashboard quality", "One page, one message, live formulas, no hardcoded numbers"],
     ["4", "Operational insight", "At least one recommendation tied to a number (capacity, billing or LOS)"],
     ["5", "Ethics", "Anonymised data only; anonymisation verified and noted"],
     ["6", "Professional pack", "Report and deck labelled, sourced, consistent"]],
    widths=[6, 34, 112])
r = section(ws, r, "Rules & ethics", ncols=6)
for t in ["Anonymised dataset only — verify no names or identifiable detail remain",
          "Report the median alongside the mean for billing (right-skewed data)",
          "State data-quality caveats in the report — hiding known issues is a finding in itself",
          "Attached datasets only"]:
    r = add_check_box(ws, r, t, ncols=6)

# ============ PROJECT D — Forecast ============
ws = wb.create_sheet("04_Project_Forecast")
setup_sheet(ws, "Project Brief D — Forecasting & Cross-Dataset Benchmark", "Trend forecasting on Superstore with a cancellation benchmark from the Hotel dataset.")
r = 4
r = section(ws, r, "The brief", ncols=6)
r = write_table(ws, r, ["Item", "Detail"],
    [["Client", "ETEH operations"],
     ["Question", "What is the sales trend, how accurate is a linear forecast, and how much demand do we lose to cancellations?"],
     ["Data", "Superstore CLEAN (9,994 order lines) + Hotel Reservations CLEAN (36,275 reservations)"],
     ["Your task", "Build the monthly trend, forecast with error measurement, and benchmark demand loss from cancellations"],
     ["Deliverables", "Analysis workbook · 3-page PDF report · 8-slide deck · QA sign-off · reflection"],
     ["Time", "5 working sessions"]],
    widths=[18, 108])
r = section(ws, r, "Milestones", ncols=6)
r = header_row(ws, r, ["Session", "Milestone", "What you must produce", "Done"], widths=[10, 26, 100, 12])
msD = [
 ["1", "Profile & trend", "Monthly sales pivot from Superstore; profile the hotel cancellation rate.", "☐"],
 ["2", "Forecast", "SLOPE/INTERCEPT trend, FORECAST.LINEAR next month, and a backtest (hold out the last month, compute MAE).", "☐"],
 ["3", "Benchmark", "Cancellation-adjusted demand view: net demand = gross × (1 − cancel rate); state the caveat.", "☐"],
 ["4", "Validate", "Regression cross-check: Sales → Profit (SLOPE/INTERCEPT/RSQ) and interpretation of R².", "☐"],
 ["5", "Report & defend", "3-page report with qualified forecast (value + MAE + lead time), 8-slide deck, QA checklist, reflection.", "☐"],
]
for m in msD:
    for j, v in enumerate(m, 1):
        c = ws.cell(row=r, column=j, value=v)
        c.font = Font(name="Calibri", size=9); c.border = BORDER; c.alignment = WRAP
    ws.row_dimensions[r].height = 32
    r += 1
r = section(ws, r, "Success criteria", ncols=6)
r = write_table(ws, r, ["#", "Criterion", "Evidence the assessor will look for"],
    [["1", "Forecast with honesty", "Forecast value quoted WITH lead time and MAE — never an unqualified number"],
     ["2", "Backtest", "A holdout prediction with its error, showing you measured accuracy"],
     ["3", "Benchmark applied", "Cancellation rate used correctly to adjust demand; caveat stated"],
     ["4", "Regression correctness", "SLOPE/INTERCEPT/RSQ computed; R² interpreted (not just quoted)"],
     ["5", "Story quality", "Conclusion-first deck; one message per slide; sources on every data slide"],
     ["6", "Professional pack", "Report and deck labelled, sourced, consistent"]],
    widths=[6, 34, 112])
r = section(ws, r, "Rules & ethics", ncols=6)
for t in ["Never extrapolate beyond the data range without stating it as a direction, not a promise",
          "Quote the MAE with every forecast",
          "Seasonality claims need at least two full cycles — the data has about one, so treat them as hypotheses",
          "Attached datasets only"]:
    r = add_check_box(ws, r, t, ncols=6)

# ============ Submission ============
ws = wb.create_sheet("05_Submission")
setup_sheet(ws, "Submission Checklist — All Project Tracks", "Complete this before submitting. Every item must be checked.")
r = 4
r = section(ws, r, "The project pack (6 pieces)", ncols=6)
r = write_table(ws, r, ["#", "Piece", "Quality bar", "Done"],
    [["1", "Analysis workbook", "Live formulas/queries; DATA sheets labelled; no hardcoded results", "☐"],
     ["2", "Dashboard", "One page, one message, all KPIs live, slicers tested", "☐"],
     ["3", "PDF report (3 pages)", "Summary, evidence, recommendation; sources and version stated", "☐"],
     ["4", "Deck (8 slides)", "Conclusion-first; every number from the project data; sources present", "☐"],
     ["5", "QA sign-off", "10-point checklist completed; defects fixed; reviewer named", "☐"],
     ["6", "Reflection (5 lines)", "What you learned, what you would redo, next skill", "☐"]],
    widths=[6, 26, 100, 12])
r = section(ws, r, "Final checks", ncols=6)
for t in ["Every KPI reconciles to a query or formula on the attached datasets",
          "No external datasets used anywhere in the deliverables",
          "Sensitive data (Track B/C) used in anonymised form only",
          "Files named consistently: <Track>_<Piece>_<YourName>.xlsx / .pdf / .pptx",
          "Workbook opens without errors; formulas calculate on open"]:
    r = add_check_box(ws, r, t, ncols=6)
r = note(ws, r, "Marking: the 100-mark rubric in the Week 5 Section Examples (10_P.10_Rubric) applies to every track. Reconciliation is binary — a number that disagrees with the dataset fails the correctness criterion.", ncols=6, height=32)

wb.save("04_Student_Projects/Student_Projects_Briefs.xlsx")
print("saved 04_Student_Projects/Student_Projects_Briefs.xlsx", wb.sheetnames)
