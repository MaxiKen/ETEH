"""Build Week 5 Section Examples workbook (Capstone & Career, P.1-P.10)."""
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from xl_helpers import *

CLEAN = "ETEH_Bootcamp/Datasets/CLEAN"
W5 = "01_Section_Examples/Week5_Capstone_Career_Section_Examples.xlsx"
GOLD = PatternFill("solid", fgColor=GOLD)
GREEN = PatternFill("solid", fgColor=GREEN)

ec = pd.read_csv(f"{CLEAN}/ecommerce_transactions_CLEAN.csv")
sp = pd.read_csv(f"{CLEAN}/student_performance_CLEAN.csv")
hc = pd.read_csv(f"{CLEAN}/healthcare_ANONYMISED.csv")
hot = pd.read_csv(f"{CLEAN}/hotel_reservations_CLEAN.csv")
sup = pd.read_csv(f"{CLEAN}/superstore_CLEAN.csv")

ec_rev = ec["total_price"].sum()
ec_cancelled = (ec["order_status"] == "Cancelled").sum()
sp_n = len(sp)
hc_n = len(hc)
hot_cancel_pct = round(hot["Is_Cancelled"].mean() * 100, 2)
sup_rev = sup["Sales"].sum()

wb = openpyxl.Workbook()

# ============ 00_README ============
readme_sheet(wb,
    "Week 5 — Capstone Project & Career (Section Examples)",
    "Modules P.1-P.10. The capstone week: full project lifecycle on the attached datasets, then portfolio and interview preparation.",
    ["Worked examples for every module of Week 5 (Vol 5), in presentation order.",
     "The three capstone tracks map to three attached datasets: the E-commerce track (full analytics lifecycle), the Student-Performance track (education analytics), and the Healthcare track (operations analytics).",
     "Every number quoted here comes from the CLEAN datasets — your project deliverables must reconcile to them."],
    [["E-commerce (CLEAN)", "11_DATA_Ecommerce", "P.3, P.6, P.7, P.8", f"{len(ec):,} lines", "The capstone track: revenue, payments, delivery"],
     ["Student Performance (CLEAN)", "12_DATA_Education", "P.4, P.6", f"{sp_n:,} rows", "The education capstone track"],
     ["Healthcare (CLEAN, anonymised)", "13_DATA_Health", "P.5, P.6", f"{hc_n:,} rows", "The healthcare capstone track"],
     ["Hotel Reservations (CLEAN)", "14_DATA_Hotel", "P.3, P.6", f"{len(hot):,} rows", "Demand/cancellation analysis for QA practice"],
     ["Superstore (CLEAN)", "15_DATA_Superstore", "P.4, P.8", f"{len(sup):,} rows", "Cross-track benchmark for portfolio work"]],
    ["1. Week 5 is project week: each module is a stage of a real deliverable — brief, data, build, QA, defence.",
     "2. P.3-P.5 are the three capstone tracks; pick one (instructor assigns) and follow its sheet.",
     "3. DATA sheets are 500-row extracts; full datasets live in ETEH_Bootcamp/Datasets/CLEAN.",
     "4. The portfolio sheet (P.8) lists exactly what to submit — the rubric (P.10) marks it."])

# ============ P.1 ============
ws = wb.create_sheet("01_P.1_Lifecycle")
setup_sheet(ws, "Module P.1 — The Project Lifecycle", "From brief to sign-off. The 8-stage analytics lifecycle applied to a real project.")
r = 4
r = section(ws, r, "The 8 stages of an analytics project", ncols=6)
r = write_table(ws, r, ["Stage", "What happens", "Deliverable", "E-commerce capstone example"],
    [["1 Scope", "Define the question, success criteria, constraints", "One-sentence brief + acceptance criteria", "“How do payment method and delivery performance drive e-commerce revenue?”"],
     ["2 Data", "Source, profile, quality-check", "Data dictionary + quality report", f"11,000 transaction lines from the SQLite file; key integrity checked"],
     ["3 Clean", "Standardise, dedupe, dirty-fix", "CLEAN extracts + issue log", "DIRTY files from Week 2 → CLEAN (see Datasets folder)"],
     ["4 Model", "Structure tables, relationships, measures", "Model + measures list", "Star schema: transactions → customers/products/sellers/deliveries"],
     ["5 Analyse", "Aggregate, compare, forecast", "Analysis tables + charts", "Revenue by category/method/month; delivery performance"],
     ["6 Visualise", "Dashboard / report assembly", "Dashboard + PDF report", "One-page ops dashboard (Week 3 blueprint)"],
     ["7 Review", "QA, verification, peer review", "QA checklist + sign-off", "Every KPI reconciled to SQL (Week 4.10 audit)"],
     ["8 Deliver", "Publish, present, hand over", "Final deliverable + deck", "Capstone deck + workbook + PDF report"]],
    widths=[10, 40, 36, 60])
r = section(ws, r, "Worked example — scope statement (the most common failure is no scope)", ncols=6)
r = write_table(ws, r, ["Scope element", "Good (specific, measurable)", "Bad (vague)"],
    [["Question", "Which factors separate high-value from low-value e-commerce lines?", "“Analyse the e-commerce data”"],
     ["Success", "Dashboard with 4 KPIs, each reconciled to SQL within rounding", "“Make a nice dashboard”"],
     ["Constraints", "One workbook, attached datasets only, 500-row practice extracts", "“Use whatever data you find”"],
     ["Out of scope", "No external data, no live systems, no customer contact", "(unstated — scope creep begins here)"],
     ["Deliverables", "Workbook, PDF report, 10-slide deck, QA sign-off", "“Something impressive”"]],
    widths=[18, 66, 46])

# ============ P.2 ============
ws = wb.create_sheet("02_P.2_Governance")
setup_sheet(ws, "Module P.2 — Data Governance", "Who owns data, how versions and access are controlled. Applied to the bootcamp datasets.")
r = 4
r = section(ws, r, "Governance controls in a small analytics team", ncols=6)
r = write_table(ws, r, ["Control", "What it means", "Applied to this bootcamp", "Evidence"],
    [["Ownership", "A named person answers for the data", "Dataset owner signs the dictionary", "DATASET_GUIDE.pdf"],
     ["Versioning", "Files get versions, never overwritten silently", "CLEAN vs DIRTY naming; v1 in folder names", "Datasets/CLEAN, DIRTY_FOR_WEEK2"],
     ["Access", "Who may read/modify/publish", "Raw files read-only for students; extracts for practice", "500-row extract cap"],
     ["Quality gates", "Data may not enter analysis until it passes checks", "Cleaning verified against issue log", "DIRTY_ISSUES_LOG.csv"],
     ["Lineage", "Every number traces to a source", "SQL answers + formulas kept as evidence", "Week 3 ANSWERS sheet"],
     ["Retention", "Data kept only as long as needed", "Teaching extracts, not full files, in workbooks", "500-row caps"],
     ["Change control", "Modifications documented and approved", "Cleaning scripts kept, not just results", "build_*.py in the repo"]],
    widths=[16, 40, 50, 40])
r = section(ws, r, "Governance in practice — the decision log", ncols=6)
r = write_table(ws, r, ["Decision", "Decision", "Rationale", "Owner"],
    [["Use CLEAN files for analysis", "Approved", "Standardised, dedupe-checked, issue-logged", "Dataset owner"],
     ["Students practise on 500-row extracts", "Approved", "Practice files stay small; full data stays with instructor", "Course lead"],
     ["Depression file used anonymised only", "Mandatory", "Sensitive data — PRIVATE framework (Week 4.13)", "Ethics lead"],
     ["Raw DIRTY files regenerated by script", "Approved", "Reproducible dirt, documented issues", "Dataset owner"],
     ["Publish the cleaned dataset dictionary", "Approved", "Transparency for every field", "Dataset owner"],
     ["No external datasets in deliverables", "Mandatory", "Attached datasets only — course standard", "Course lead"]],
    widths=[40, 16, 60, 20])

# ============ P.3 ============
ws = wb.create_sheet("03_P.3_CapstoneEcom")
setup_sheet(ws, "Module P.3 — Capstone Track A: E-commerce Analytics", "The full lifecycle on the attached E-commerce dataset. Track A brief + worked milestones.")
r = 4
r = section(ws, r, "Project brief (Track A)", ncols=6)
r = write_table(ws, r, ["Item", "Detail"],
    [["Client", "ETEH Ltd — internal e-commerce operations team"],
     ["Question", "Where does e-commerce revenue come from, and what holds it back?"],
     ["Data", "E-commerce SQLite database + CLEAN CSVs (transactions, customers, products, sellers, deliveries)"],
     ["Deliverables", "Analysis workbook, one-page dashboard, 3-page PDF report, 8-slide deck, QA sign-off"],
     ["Success", "Every KPI reconciled to SQL; one recommendation tied to a number; deck reviewed by peer"],
     ["Timeline", "5 working sessions (see milestone table below)"],
     ["Constraints", "Attached datasets only; practice on 500-row extracts; no external data"]],
    widths=[18, 104])
r = section(ws, r, "Milestone plan with real targets (computed from the file)", ncols=6)
r = write_table(ws, r, ["Session", "Milestone", "Target / real numbers", "Evidence to keep"],
    [["1", "Profile & scope", f"{len(ec):,} lines; keys verified; cancellations identified", "Data profile + quality notes"],
     ["2", "Clean & model", "DIRTY→CLEAN reconciliation; star schema set", "Cleaning log + model"],
     ["3", "Analyse", f"Total revenue ₦{ec_rev:,.0f}; revenue by category/method/month computed", "SQL queries + answers"],
     ["4", "Visualise", "Dashboard: 4 KPIs, trend, top sellers, payment mix", "Dashboard + blueprint"],
     ["5", "Report & defend", "PDF report + deck; peer QA on every KPI", "Signed QA checklist"]],
    widths=[10, 24, 66, 40])
r = section(ws, r, "Worked analysis — the numbers your project must reproduce", ncols=6)
r = write_table(ws, r, ["Finding", "Value (from the file)", "How to reproduce"],
    [["Total revenue (all lines)", f"₦{ec_rev:,.2f}", "SELECT SUM(total_price) FROM transactions"],
     ["Cancelled lines", f"{ec_cancelled:,} ({ec_cancelled/len(ec)*100:.1f}%)", "COUNT WHERE order_status='Cancelled'"],
     ["Average line value", f"₦{ec['total_price'].mean():,.2f}", "AVG(total_price)"],
     ["Largest category revenue", "See Week 3 ANSWERS (category pivot)", "SUM by category, ORDER BY DESC LIMIT 1"],
     ["Payment-method revenue mix", "See Week 3 ANSWERS (method pivot)", "SUM by payment_method"],
     ["Monthly trend", "See Week 3 monthly table", "SUM by month"],
     ["Top seller", "See Week 3 ANSWERS", "JOIN sellers, GROUP BY, ORDER BY DESC LIMIT 1"]],
    widths=[30, 40, 64])
r = note(ws, r, "Presentation note: the capstone deck uses ONLY the attached e-commerce data for its numbers — inline examples may reference the other datasets, but every project figure comes from this track's file.", ncols=6, height=30)

# ============ P.4 ============
ws = wb.create_sheet("04_P.4_CapstoneEdu")
setup_sheet(ws, "Module P.4 — Capstone Track B: Student Performance Analytics", "Education analytics track on the Student Performance dataset.")
r = 4
r = section(ws, r, "Project brief (Track B)", ncols=6)
r = write_table(ws, r, ["Item", "Detail"],
    [["Client", "ETEH Education — academic performance team"],
     ["Question", "Which factors associate with higher exam performance, and where should support go?"],
     ["Data", "StudentPerformanceFactors CLEAN (6,607 records, 27 fields)"],
     ["Deliverables", "Analysis workbook, dashboard, 3-page PDF report, 8-slide deck"],
     ["Success", "Findings backed by the data; no claims without a comparison; recommendation actionable"],
     ["Ethics", "Anonymised records; no individual student identified (Week 4.13)"]],
    widths=[18, 104])
r = section(ws, r, "Worked analysis seeds (compute from the extract)", ncols=6)
r = write_table(ws, r, ["Analysis", "What to build", "Reference result"],
    [["Performance bands", "Count students by Exam_Band", "High/Medium/Low split — see Week 1 ANSWERS"],
     ["Parental support vs band", "Crosstab: Parental_Support × Exam_Band", "Higher support → more High bands (verify direction)"],
     ["Attendance vs score", "Correlation: Attendance_Rate vs Exam_Score", "Compute r on the full file"],
     ["City comparison", "Average Exam_Score by City", "Top and bottom cities with n per city"],
     ["Support recommendation", "Target the band with the weakest attendance/support mix", "Tie the recommendation to the computed gap"]],
    widths=[26, 56, 64])
r = note(ws, r, f"Cross-track practice: repeat the same analysis shape on the Superstore extract (Segment × Region × Profit) to show the method transfers — that is a portfolio-grade insight.", ncols=6, height=30)

# ============ P.5 ============
ws = wb.create_sheet("05_P.5_CapstoneHealth")
setup_sheet(ws, "Module P.5 — Capstone Track C: Healthcare Operations Analytics", "Hospital admissions track on the Healthcare dataset (anonymised).")
r = 4
r = section(ws, r, "Project brief (Track C)", ncols=6)
r = write_table(ws, r, ["Item", "Detail"],
    [["Client", "ETEH Health — hospital operations"],
     ["Question", "Where do admissions and billing concentrate, and what are the operational risks?"],
     ["Data", "healthcare_ANONYMISED CLEAN (55,500 admissions, anonymised)"],
     ["Deliverables", "Analysis workbook, dashboard, 3-page PDF report, 8-slide deck"],
     ["Success", "Billing and occupancy findings reconciled to the file; data-quality issues flagged honestly"],
     ["Ethics", "Anonymised only — no patient names, no identifiable detail"]],
    widths=[18, 104])
r = section(ws, r, "Worked analysis seeds", ncols=6)
r = write_table(ws, r, ["Analysis", "What to build", "Reference result"],
    [["Admissions by department/ward", "COUNT admissions, SUM billing by ward", "Concentration check — verify on the extract"],
     ["Billing distribution", "Distribution of Billing_Amount_Abs; mean vs median", "Right-skewed: quote median with mean"],
     ["Stay length", "Length of Stay stats: mean, median, P90", "Operational planning numbers"],
     ["Data-quality flags", "Negative billing rows (pre-cleaning), missing fields", "From the cleaning log — report transparently"],
     ["Operational risk", "Ward with high admissions + long stay = capacity risk", "Rank wards by admissions × LOS"]],
    widths=[26, 56, 64])
r = note(ws, r, "Hotel-track cross-check: use the hotel cancellation rate (" + f"{hot_cancel_pct}%" + ") as a benchmark for the 'demand you do not get' narrative — a capacity-planning insight across two datasets.", ncols=6, height=32)

# ============ P.6 ============
ws = wb.create_sheet("06_P.6_QualityAssurance")
setup_sheet(ws, "Module P.6 — Quality Assurance", "The QA gate before any capstone is submitted. The checklist that protects the numbers.")
r = 4
r = section(ws, r, "Capstone QA checklist (all tracks)", ncols=6)
r = write_table(ws, r, ["#", "Check", "Pass condition", "Evidence"],
    [["1", "Scope honoured", "Deliverables match the brief; nothing extra, nothing missing", "Brief sheet"],
     ["2", "Data source", "Only attached datasets; version stated", "Dataset version in title block"],
     ["3", "Cleanliness", "No uncleaned DIRTY values in analysis", "Cleaning log reference"],
     ["4", "Reconciliation", "Every KPI equals a SQL/Excel-verified number", "ANSWERS cross-check"],
     ["5", "Completeness", "All requested charts/tables present and labelled", "Dashboard + report"],
     ["6", "Accuracy", "No hardcoded numbers; formulas/measures live", "Inspect workbook"],
     ["7", "Ethics", "Sensitive datasets anonymised; consent noted", "Ethics note"],
     ["8", "Readability", "One message per slide/page; sources on every data page", "Peer review"],
     ["9", "Reproducibility", "A colleague could rebuild from the workbook + dictionary", "Handover test"],
     ["10", "Sign-off", "Analyst and reviewer named, dated", "Sign-off block"]],
    widths=[6, 24, 56, 40])
r = section(ws, r, "Worked QA — finding and fixing three real defects", ncols=6)
r = write_table(ws, r, ["Defect found", "Why it fails QA", "Fix"],
    [["Dashboard KPI total is a typed number, not a formula", "Not reproducible; a refresh changes nothing", "Replace with SUM over the model; reconcile to SQL"],
     ["Report says “most bookings are cancelled” while the data shows 20.6% cancelled", "Claim contradicts the data (inverted)", "Re-read the cancel rate; rewrite the headline"],
     ["Slides use a figure from an external example", "Not from the attached datasets — scope violation", "Rebuild the figure from the project dataset"]],
    widths=[50, 50, 50])

# ============ P.7 ============
ws = wb.create_sheet("07_P.7_Defence")
setup_sheet(ws, "Module P.7 — Defending Your Work", "Presenting and defending the capstone: structure, common questions, honest answers.")
r = 4
r = section(ws, r, "The defence structure (10 minutes)", ncols=6)
r = write_table(ws, r, ["Minute", "Content", "Rule"],
    [["0-2", "The headline finding + why it matters", "State the conclusion, not the process"],
     ["2-5", "The evidence: 2 charts + the numbers behind them", "Every number traceable to the dataset"],
     ["5-7", "The method in one breath", "Source, cleaning, one formula/measure example"],
     ["7-9", "Limitations & data quality", "Be honest — it builds credibility"],
     ["9-10", "The recommendation with a number", "“Do X, which moves Y by Z”"]],
    widths=[10, 56, 60])
r = section(ws, r, "Likely examiner questions and strong answers", ncols=6)
r = write_table(ws, r, ["Question", "Strong answer pattern", "Dataset-grounded example"],
    [["“Why this metric?”", "Metric → decision it informs", "Revenue by payment method → where to focus payment experience"],
     ["“How do we know the number is right?”", "Reconciliation story", "Dashboard KPI = SQL query on the same file, zero mismatch"],
     ["“What about cancelled orders?”", "State the rate and its treatment", f"{ec_cancelled:,} cancelled lines excluded from delivered revenue, shown separately"],
     ["“What would change your conclusion?”", "Name the driver that would flip it", "A category overtaking the current leader by >15%"],
     ["“Why not more data?”", "Scope discipline + governance", "Attached datasets only; governance keeps the deliverable auditable"],
     ["“Is correlation causation here?”", "Name the trap and the safeguard", "qty vs total_price is arithmetic, not causation (Week 4.3)"]],
    widths=[34, 44, 66])
r = section(ws, r, "Defence drills", ncols=6)
for t in ["Rehearse the 10-minute structure out loud until it fits the time exactly",
          "Prepare one reconciliation number you can quote instantly (total revenue)",
          "Write down two limitations BEFORE the defence — do not improvise them",
          "Practise answering a wrong-but-confident question: acknowledge, check the data, correct"]:
    r = add_check_box(ws, r, t, ncols=6)

# ============ P.8 ============
ws = wb.create_sheet("08_P.8_Portfolio")
setup_sheet(ws, "Module P.8 — Building Your Portfolio", "What to submit, in what order, with what evidence. The portfolio checklist.")
r = 4
r = section(ws, r, "The portfolio pack (this course's standard)", ncols=6)
r = write_table(ws, r, ["Piece", "What it shows", "Quality bar", "Source material"],
    [["1 Workbook", "Cleaned data + live formulas/measures + DATA sheets", "No hardcoded results; ANSWERS sheet present", "Section Examples + Assignments"],
     ["2 Dashboard", "One page, one message, all KPIs reconciled", "Blueprint-compliant (Week 3.9)", "Capstone track"],
     ["3 PDF report", "3 pages: summary, evidence, recommendation", "Sources and version stated", "Capstone track"],
     ["4 Deck", "8-10 slides, conclusion-first", "Every number from the project data", "Capstone track"],
     ["5 QA sign-off", "Checklist completed, defects fixed", "All 10 checks passed", "Week P.6"],
     ["6 Reflection", "What you learned, what you would redo", "Specific to this project", "Self-authored"]],
    widths=[20, 44, 44, 36])
r = section(ws, r, "Portfolio review criteria (what a reviewer looks for)", ncols=6)
r = write_table(ws, r, ["Criterion", "Evidence of quality", "Red flag"],
    [["Data discipline", "Cleaning documented; issues logged; version stated", "“I just used the data as-is”"],
     ["Reproducibility", "Formulas/queries live; another person can rebuild", "Typed numbers in dashboards"],
     ["Analytical thinking", "Comparisons, context, honest limitations", "Charts with no interpretation"],
     ["Storytelling", "Conclusion-first, one message per slide", "Process-heavy, conclusion-free decks"],
     ["Ethics", "Anonymisation and consent noted", "Sensitive data used carelessly"],
     ["Professional polish", "Consistent formatting, sources, page numbers", "Inconsistent fonts/colours"]],
    widths=[20, 60, 48])
r = section(ws, r, "Submission checklist", ncols=6)
for t in ["All six pieces present and named consistently",
          "Every workbook opens without errors; formulas calculate",
          "Dashboard and report use only the assigned track's dataset",
          "QA sign-off complete with reviewer named",
          "Files in the 04_Student_Projects folder structure (or as instructed)"]:
    r = add_check_box(ws, r, t, ncols=6)

# ============ P.9 ============
ws = wb.create_sheet("09_P.9_Interviews")
setup_sheet(ws, "Module P.9 — Interview Readiness", "The analytics interview: technical + behavioural, grounded in what you did this course.")
r = 4
r = section(ws, r, "Technical questions you can answer from this bootcamp", ncols=6)
r = write_table(ws, r, ["Question", "Answer shape", "Where you practised it"],
    [["Explain a join and when it loses rows", "INNER vs LEFT; non-matching keys", "Week 3.2 / 3.4"],
     ["How would you clean a messy CSV?", "Profile → issue log → standardise → verify", "Week 2.4-2.5"],
     ["Correlation vs causation?", "r measures co-movement, not mechanism", "Week 4.3"],
     ["How do you validate a dashboard?", "Reconcile every KPI to source; refresh test", "Week 4.10 / P.6"],
     ["What is a measure vs a calculated column?", "Computed at query time vs load time", "Week 3.6"],
     ["How do you handle bad data quality?", "Flag, log, fix, re-verify — never silently", "Week 2 + P.2"]],
    widths=[36, 56, 36])
r = section(ws, r, "Behavioural questions — STAR format with this course's evidence", ncols=6)
r = write_table(ws, r, ["Question", "Situation/Task", "Action (use these)", "Result (quantify)"],
    [["Tell me about a project", "Capstone track (pick yours)", "Scoped the brief, cleaned DIRTY→CLEAN, built the model, QA'd every KPI", f"Reconciled all KPIs to SQL; delivered workbook, dashboard, report, deck"],
     ["Tell me about a data problem", "DIRTY files with 16 logged issues", "Profiled, logged issues, standardised, verified against the CLEAN truth", "Zero residual issues; cleaning fully scripted"],
     ["Tell me about working to a deadline", "5-session capstone timeline", "Milestone plan, evidence kept at each stage", "All five milestones delivered on the plan"],
     ["Tell me about ethics", "Sensitive student/health data", "PRIVATE framework; anonymised extracts; purpose limitation", "Deliverables pass the ethics gate"]],
    widths=[28, 36, 56, 40])
r = section(ws, r, "Interview drills", ncols=6)
for t in ["Answer every technical question with the dataset you actually used — specifics beat generalities",
          "Keep one reconciliation story ready (dashboard KPI vs SQL)",
          "Prepare one failure story with what you learned — interviewers trust calibrated honesty",
          "Know your numbers cold: total revenue, cancellations, rates, top category"]:
    r = add_check_box(ws, r, t, ncols=6)

# ============ P.10 ============
ws = wb.create_sheet("10_P.10_Rubric")
setup_sheet(ws, "Module P.10 — Capstone Rubric", "The marking rubric. What earns each grade, stated before the work is done.")
r = 4
r = section(ws, r, "Capstone rubric (100 marks)", ncols=6)
r = header_row(ws, r, ["Criterion", "Weight", "Excellent (full marks)", "Pass", "Fail (no marks)"], widths=[26, 10, 44, 40, 40])
rub = [
 ["Data handling (cleaning, quality, governance)", "20", "All issues logged and fixed; cleaning scripted and reproducible; version and source stated", "Most issues fixed; cleaning documented informally", "Data used as-is; no quality checks"],
 ["Analysis correctness", "25", "All KPIs reconciled to SQL/formulas; no hardcoded numbers; edge cases (cancellations, outliers) handled", "KPIs mostly correct; minor hardcoding or edge-case misses", "Numbers unverified or wrong"],
 ["Modelling & tooling", "15", "Star schema/relationships correct; measures or formulas live and reusable", "Tools used correctly; minor structural issues", "Copy-paste only; no modelling"],
 ["Visualisation & reporting", "15", "Dashboard one page/one message; charts labelled; PDF and deck polished and sourced", "Deliverables complete, some labelling/polish gaps", "Charts missing or misleading"],
 ["Storytelling & insight", "15", "Conclusion-first; recommendation tied to a number; limitations honest", "Findings presented with some interpretation", "Numbers listed with no narrative"],
 ["Ethics & professionalism", "10", "Sensitive data anonymised; consent and purpose stated; QA sign-off complete", "Ethics requirements met, QA mostly complete", "Ethics or QA ignored"],
]
for row in rub:
    for j, v in enumerate(row, 1):
        c = ws.cell(row=r, column=j, value=v)
        c.font = Font(name="Calibri", size=9)
        c.border = BORDER
        c.alignment = WRAP
    ws.row_dimensions[r].height = 44
    r += 1
r = section(ws, r, "How the marks map", ncols=6)
r = write_table(ws, r, ["Grade", "Marks", "Meaning", "Capstone outcome"],
    [["A", "90-100", "Professional-grade: auditable, insightful, polished", "Portfolio-ready; recommend for interview"],
     ["B", "75-89", "Solid: all requirements met, some polish missing", "Pass with minor revisions"],
     ["C", "60-74", "Adequate: core requirements met, gaps in verification", "Pass with revisions required"],
     ["D / F", "<60", "Requirements missing: numbers unverified or deliverables incomplete", "Resubmit"]],
    widths=[10, 12, 66, 44])
r = note(ws, r, "Rubric transparency: students see this rubric before starting the capstone — there are no hidden criteria. Markers use the same table; the QA checklist (P.6) is the evidence trail behind every criterion.", ncols=6, height=32)

# ============ DATA sheets ============
data_sheet(wb, "11_DATA_Ecommerce", ec, "E-commerce — transactions (CLEAN, extract)",
           "Attached dataset. Track A capstone practice data.", tab_color=TEAL, max_rows=500, table_name="CapEcom")
data_sheet(wb, "12_DATA_Education", sp, "Student Performance (CLEAN, extract)",
           "Attached dataset. Track B capstone practice data.", tab_color=TEAL, max_rows=500, table_name="CapEdu")
data_sheet(wb, "13_DATA_Health", hc, "Healthcare (CLEAN, anonymised, extract)",
           "Attached dataset. Track C capstone practice data.", tab_color=TEAL, max_rows=500, table_name="CapHealth")
data_sheet(wb, "14_DATA_Hotel", hot, "Hotel Reservations (CLEAN, extract)",
           "Attached dataset. Cancellation benchmark for capacity narratives.", tab_color=TEAL, max_rows=500, table_name="CapHotel")
data_sheet(wb, "15_DATA_Superstore", sup, "Superstore (CLEAN, extract)",
           "Attached dataset. Cross-track benchmark for portfolio work.", tab_color=TEAL, max_rows=500, table_name="CapSuper")

# ============ ANSWERS ============
ws = wb.create_sheet("16_ANSWERS")
setup_sheet(ws, "Week 5 — Answers & Marking Guide (Instructor)", "Capstone reference numbers and rubric application notes.")
r = 4
r = section(ws, r, "Capstone reference numbers", ncols=6)
r = write_table(ws, r, ["Track", "Key numbers (from the files)"],
    [["A — E-commerce", f"{len(ec):,} lines · revenue ₦{ec_rev:,.2f} · cancelled {ec_cancelled:,} · average line ₦{ec['total_price'].mean():,.2f}"],
     ["B — Student Performance", f"{sp_n:,} records · Exam bands High/Medium/Low · attendance/support crosstabs (compute per class)"],
     ["C — Healthcare", f"{hc_n:,} admissions · anonymised · billing distribution right-skewed · LOS P90 from file"],
     ["Cross-track", f"Hotel cancel rate {hot_cancel_pct}% · Superstore total sales ${sup_rev:,.0f}"]],
    widths=[26, 110])
r = section(ws, r, "Marking notes per criterion", ncols=6)
for t in ["Data handling: accept the scripted cleaning or an equally reproducible manual log with the DIRTY_ISSUES_LOG as the source of truth.",
          "Analysis: the reconciliation test is binary — a KPI that disagrees with the dataset fails the criterion regardless of presentation.",
          "Visualisation: mark the message, not the decoration — a plain but correct chart beats a decorated wrong one.",
          "Storytelling: require at least one recommendation with a quantified target.",
          "Ethics: any use of the depression/healthcare files without anonymisation notes fails the ethics criterion outright."]:
    r = note(ws, r, t, ncols=6, height=26)

wb.save(W5)
print("saved", W5, "sheets=", wb.sheetnames, "size=", f"{__import__('os').path.getsize(W5)/1024:.0f} KB")
