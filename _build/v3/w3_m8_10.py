"""Week 3 edition-3 slide copy: Modules 3.8 (dashboard development), 3.9 (three blueprints) and 3.10 (file extracts into
SQL, DAX and dashboards). Dashboard figures: SuperMart Stores FY2025 (tDash, tDashStores). 3.10 loads
Week2_Data_Files/Lifecare_Admissions_2023_DIRTY.csv into an in-memory SQLite staging table and runs the SQL shown."""
import sqlite3

import pandas as pd

import w3_data as D
from common import SEC, W2D
from v3kit import T
from w3_fx import (CEN, CODE, CV, DS, FY24, FY25, I, LC, N, Nm, P, PROFIT25, S25, SQL, TARGET25, ZS, q, v)

ALL_ATT = FY25 / TARGET25
DSA = DS.sort_values("Attainment", ascending=False)
ON = DS[DS.Attainment >= 1]
BELOW = DS[DS.Attainment < 1]
SE_Z = ZS.loc["South-East"]
QTR = S25.groupby(S25.Order_Date.dt.quarter).Sales.sum()
LOW = DSA.iloc[-1]
LOWCAT = S25[S25.Store_ID == LOW.Store_ID].groupby("Category").Sales.sum().sort_values(ascending=False)
LOW_LINES = int((S25.Store_ID == LOW.Store_ID).sum())
LOSS = int((S25.Profit < 0).sum())


def store(r):
    return f"{r.Store_ID} {r.Store_Name}"


M38 = dict(
    id="3.8", title="Dashboard Development",
    subtitle="From audience and questions to a one-screen, sliceable, accessible dashboard that reconciles.",
    objectives=["Define a dashboard and its design principles.",
                "Plan from audience, questions and decisions.",
                "Lay out a grid with a colour system and type hierarchy.",
                "Build KPI cards, choose visuals and connect slicers and timelines.",
                "Assemble, print, check accessibility and performance, and QA before release."],
    topics=[
        T("3.8.1", "What Is a Dashboard?",
          defn="A dashboard is a single screen that monitors a small set of agreed KPIs for one audience, using live, refreshable data.",
          simple="No scrolling and no explanation needed. A report explains; a dashboard monitors.",
          like="Like a car's dashboard: speed, fuel and warnings at a glance, not the engine manual.",
          grid=dict(label="KPI cards, SuperMart Stores FY2025 (all zones)", widths=[2, 1.8], rows=[
              ["KPI card", "Value"], ["Net sales", N(FY25)], ["Profit", N(PROFIT25)], ["Margin", P(PROFIT25 / FY25)],
              ["Target attainment", P(ALL_ATT)]]),
          watch="A dashboard with twenty charts is a report squeezed onto one page; it monitors nothing."),

        T("3.8.2", "Dashboard Design Principles",
          defn="Dashboard design principles: one screen, one audience, one purpose; the most important KPI largest and top-left; "
               "consistent design; colour only for meaning; every number reconciles.",
          simple="Each principle can be checked on the finished sheet.",
          grid=dict(label="How the 3.8 dashboard meets them", widths=[1.6, 3.2], rows=[
              ["Principle", "Evidence"],
              ["One audience", "Regional managers: 'How is my zone doing?'"],
              ["Hierarchy", "KPI row on top, then store bar, then monthly line"],
              ["Colour for meaning", "Red only for stores below target"],
              ["Reconciles", f"Net sales (all zones) = {N(FY25)} POS total"]]),
          watch="A beautiful dashboard whose total does not match the ledger will not be trusted twice."),

        T("3.8.3", "Planning: Audience, Questions and Decisions",
          defn="Dashboard planning records who will use it, the questions they must answer and the decision each answer drives.",
          simple="List the measures needed, then build nothing that does not serve a decision.",
          grid=dict(label="Plan for the regional managers", widths=[1.8, 1.6, 1.6], rows=[
              ["Question", "Decision", "Visual"],
              ["Is my zone on target?", "Act on stores behind", "Attainment card"],
              ["Which stores lag?", "Visit or support", "Store bar chart"],
              ["When do sales peak?", "Plan stock and staff", "Monthly line"],
              ["Are we profitable?", "Review pricing", "Margin card"]]),
          watch="Starting from the data ('what can we chart?') instead of the decisions gives a dashboard nobody uses."),

        T("3.8.4", "Grid and Layout",
          defn="A dashboard grid aligns every element to common edges and equal gaps, in the order readers scan.",
          simple="Readers scan in a Z: top-left, across, down, across. KPIs on top, detail below.",
          grid=dict(label="Layout of the 3.8 dashboard", widths=[1.4, 2.4, 0.8], rows=[
              ["Area", "Content", "Read"],
              ["Top", "Zone selector + 4 KPI cards", "First"],
              ["Middle left", "Store sales bar chart", "Second"],
              ["Middle right", "Monthly sales line", "Third"],
              ["Bottom", "QA checklist and source note", "Last"]]),
          watch="Charts placed 'roughly' by eye look untidy; use Format › Align and equal column widths."),

        T("3.8.5", "Colour System",
          defn="A colour system assigns each colour one meaning: one main colour for data, grey for context, one accent, and "
               "red/amber/green only for status against a target.",
          simple="Colour should say something; otherwise leave it neutral.",
          grid=dict(label="Store attainment, FY2025", widths=[0.9, 1.4, 2.6], rows=[
              ["Colour", "Meaning", "Stores"],
              ["Navy", "At or above target", f"{len(ON)} (best: {DSA.iloc[0].Store_ID} {DSA.iloc[0].Store_Name} {P(DSA.iloc[0].Attainment)})"],
              ["Red", "Below target", f"{len(BELOW)} (lowest: {LOW.Store_ID} {LOW.Store_Name} {P(LOW.Attainment)})"]]),
          watch="About 1 man in 12 cannot tell red from green: add a label or symbol as well as colour (3.8.17)."),

        T("3.8.6", "Typography and Label Hierarchy",
          defn="Type hierarchy uses a few consistent sizes so that size signals importance.",
          simple="Three sizes are enough: large for KPI values, medium for card labels and chart titles, small for notes.",
          grid=dict(label="Three sizes", widths=[2, 1, 1.8], rows=[
              ["Element", "Size", "Example"],
              ["KPI value", "Large", Nm(FY25)],
              ["Card label / chart title", "Medium", "Net sales"],
              ["Note / source", "Small", "FY2025, POS data"]]),
          watch="Five font sizes and three fonts make every element shout equally."),

        T("3.8.7", "KPI Cards",
          defn="A KPI card is a tile showing a measure's name, current value, a comparison (target or last year) and a status.",
          simple="Feed it from a formula or GETPIVOTDATA so it updates with the slicers.",
          grid=dict(label="Cards for two selections", widths=[1.8, 1.5, 1.5], rows=[
              ["KPI card", "All zones", "South-East"],
              ["Net sales", N(FY25), N(SE_Z.Sales)],
              ["Profit", N(PROFIT25), N(SE_Z.Profit)],
              ["Margin", P(PROFIT25 / FY25), P(SE_Z.Margin)],
              ["Target attainment", P(ALL_ATT), P(SE_Z.Attainment)]]),
          watch="A card with a value but no comparison cannot say whether the number is good or bad."),

        T("3.8.8", "Selecting the Right Chart per KPI",
          defn="Each KPI question has a natural visual: total = card, trend = line, ranking = sorted bar, share = stacked bar, "
               "actual against target = bar with a target marker.",
          simple="Pick the visual from the question the KPI answers.",
          grid=dict(label="Question and visual", widths=[2, 1.8, 1.6], rows=[
              ["Question", "Visual", "Not this"],
              ["How much in total?", "KPI card", "Pie"],
              ["Which store is behind?", "Sorted bar of attainment", "Column, diagonal names"],
              ["How did sales move?", "Line by month", "Pie per month"],
              ["Profit against sales by zone?", "Combo chart", "Two pies"]]),
          watch="The same KPI shown in two different chart types on one dashboard makes readers look for a difference."),

        T("3.8.9", "Slicers",
          defn="A slicer is an on-sheet panel of buttons that filters one or more PivotTables, PivotCharts or tables.",
          simple="Easier to see and use than a drop-down filter. Set its columns, size and style to fit the grid.",
          grid=dict(label="Zone slicer, FY2025", widths=[1.6, 1.6, 1.2], rows=[
              ["Button", "Net sales shown", "Attainment"]] + [
              [z, N(r.Sales), P(r.Attainment)] for z, r in ZS.iterrows()]),
          steps=["Click inside the PivotTable", "PivotTable Analyze › Insert Slicer › Zone",
                 "Slicer tab › Columns: 4; height to fit the row"],
          order=["grid", "steps"], split=0.55,
          watch="A slicer with no visible 'filtered' state confuses users; keep the Clear Filter button visible."),

        T("3.8.10", "Connecting One Slicer to Multiple PivotTables",
          defn="Report Connections links one slicer to several PivotTables that share the same pivot cache or Data Model.",
          simple="Right-click the slicer › Report Connections › tick every PivotTable. Every visual must respond.",
          grid=dict(label="South-East selected", widths=[1.6, 1.6, 1.6], rows=[
              ["Visual", "Before connecting", "After connecting"],
              ["KPI cards", Nm(SE_Z.Sales), Nm(SE_Z.Sales)],
              ["Store bar chart", f"All {len(DS)} stores", f"{int((DS.Zone == 'South-East').sum())} South-East stores"],
              ["Monthly line", "South-East only", "South-East only"]]),
          watch="PivotTables built on different caches cannot share a slicer; build them all from one table or the Data Model."),

        T("3.8.11", "Timelines",
          defn="A timeline is a slicer for dates that filters by a dragged range at year, quarter, month or day level.",
          simple="It needs a real date column, ideally from the date table.",
          grid=dict(label="FY2025 by quarter", widths=[1, 1.8, 1.2], rows=[
              ["Quarter", "Net sales", "Share"]] + [[f"Q{k}", N(x), P(x / FY25)] for k, x in QTR.items()]),
          steps=["PivotTable Analyze › Insert Timeline", "Tick Order_Date (or tDimDate[Date])", "Switch the level to QUARTERS"],
          order=["grid", "steps"], split=0.55,
          watch="Text dates do not appear in Insert Timeline; convert them to real dates first."),

        T("3.8.12", "Report Filters versus Slicers",
          defn="A report filter sits above one PivotTable as a drop-down; a slicer is a visible panel that can control many PivotTables.",
          simple="On dashboards use slicers; keep report filters for analysis sheets.",
          grid=dict(label="Compared", widths=[1.4, 1.2, 1.5, 1.4], rows=[
              ["Control", "Visible?", "Scope", "Use"],
              ["Report filter", "Easy to miss", "One PivotTable", "Analysis sheets"],
              ["Slicer", "Yes", "Many PivotTables", "Dashboards"]]),
          watch="A forgotten report filter makes a pivot show part of the data while everyone assumes it shows all of it."),

        T("3.8.13", "Interaction Design and Drill-Down",
          defn="Interaction design decides how users explore: slicers filter, double-clicking a value drills through to its "
               "rows, and a clear button resets the view.",
          simple=f"Double-click {LOW.Store_Name}'s value in the store pivot to see the {I(LOW_LINES)} lines behind it.",
          grid=dict(label=f"{LOW.Store_ID} {LOW.Store_Name} by category", widths=[2, 1.6], rows=[
              ["Category", "FY2025 sales"]] + [[c, N(x)] for c, x in LOWCAT.items()] + [
              [f"All {I(LOW_LINES)} lines", N(LOWCAT.sum())]]),
          watch="Drill-through creates a new sheet each time; delete those sheets before sharing the file."),

        T("3.8.14", "Assembling the Dashboard Sheet",
          defn="Assembly places the finished visuals on their own sheet, aligned to a grid and protected from accidental moves.",
          simple="Gridlines and headings off, even column widths, charts aligned with Format › Align, sheet protected.",
          grid=dict(label="Assembly steps", widths=[1.4, 2, 1.5], rows=[
              ["Step", "Setting", "Benefit"],
              ["Gridlines off", "View tab", "Clean background"],
              ["Align charts", "Format › Align", "Neat grid"],
              ["Protect sheet", "Allow 'Use PivotTable & PivotChart'", "Layout can't move"]]),
          watch="Protecting without allowing pivot use stops the slicers working."),

        T("3.8.15", "Printing, PDF Export and Page Setup",
          defn="Page setup controls how a sheet prints: print area, orientation, scaling and margins.",
          simple="A dashboard should print on one landscape page. Check the PDF: fonts, slicers and charts must be readable.",
          grid=dict(label="Page setup", widths=[1.4, 1.8, 1.6], rows=[
              ["Setting", "Value", "Why"],
              ["Print area", "Dashboard range only", "No stray data"],
              ["Orientation", "Landscape", "Matches the screen"],
              ["Scaling", "Fit to 1 page", "Nothing split"]]),
          watch="'Fit to 1 page' on a huge sheet prints text too small to read; design for the page."),

        T("3.8.16", "Web and Mobile Viewing",
          defn="Responsive dashboard design keeps a dashboard usable in a browser or on a phone, as well as in desktop Excel.",
          simple="Design for the smallest screen: fewer visuals, bigger text, nothing that needs hovering.",
          grid=dict(label="Phone issues and fixes", widths=[2, 2.4], rows=[
              ["Issue", "Fix"],
              ["Small fonts", "Larger KPI values"],
              ["Too many visuals", "Cards first, charts below"],
              ["Hover-only details", "Show values as data labels"]]),
          watch="Some features (certain slicer styles, macros) do not work in Excel for the web; test there."),

        T("3.8.17", "Accessibility",
          defn="An accessible dashboard can be read by people with low vision or colour-vision deficiency: strong contrast "
               "(4.5:1 for text), meaning not carried by colour alone, alt text on charts.",
          simple="Add the status as text next to the colour.",
          grid=dict(label="Status as text and colour", widths=[2.4, 1, 1.2], rows=[
              ["Store", "Attainment", "Status"]] + [
              [store(r), P(r.Attainment), "On target" if r.Attainment >= 1 else "Below"]
              for r in [DSA.iloc[0], DSA.iloc[-2], DSA.iloc[-1]]]),
          watch="Red/green dots with no text are invisible to many colour-blind readers."),

        T("3.8.18", "Dashboard Performance",
          defn="Dashboard performance is how fast it calculates and refreshes; slowness usually comes from many pivot caches, "
               "volatile functions or heavy queries.",
          simple="Use one source table, non-volatile functions (SUMIFS, INDEX) and only the columns you need.",
          grid=dict(label="Causes and fixes", widths=[1.8, 1.8, 1.4], rows=[
              ["Cause of slowness", "Fix", "Sheet 3.8"],
              ["Many caches", "One shared source", "One table: tDash"],
              ["Volatile functions", "INDEX / SUMIFS", "None used"],
              ["Unused columns", "Remove in Power Query", f"{D.table('tDash').shape[1]} columns kept"]]),
          watch="OFFSET, INDIRECT and TODAY recalculate on every change; a few thousand of them freeze a dashboard."),

        T("3.8.19", "Dashboard QA Checklist",
          defn="Dashboard QA checks before release that every KPI reconciles, every slicer drives every visual, no errors "
               "show, the default view is useful and the refresh date is visible.",
          simple="Write the checks down and record the results.",
          grid=dict(label="QA on the 3.8 dashboard", widths=[2, 2.2, 0.7], rows=[
              ["Check", "Result", "Pass"],
              ["Net sales (all) = POS control", f"{N(FY25)} = {N(S25.Sales.sum())}", "Yes"],
              ["4 zones add up to the total", N(ZS.Sales.sum()), "Yes"],
              ["Errors visible", "None", "Yes"],
              ["Default view", "Zone = All", "Yes"]]),
          watch="'It looked fine' is not QA. A check has a stated expected value and a recorded result."),
    ],
    takeaways=["A dashboard monitors a few KPIs for one audience on one screen.",
               "Plan from questions and decisions; lay out on a grid.",
               "Colour for meaning, three type sizes, KPIs with comparisons.",
               "Slicers connected to every visual; timelines for dates.",
               "Accessible, fast, printable, and QA'd against control totals."],
)

# ---------------------------------------------------------------- 3.9
LC23 = LC[LC.Admission_Date.dt.year == 2023]
C23 = CEN[CEN.Census_Date.dt.year == 2023]
AR = int((CV.At_Risk == "Yes").sum())
FAC = CV.groupby("Faculty").Exam_Score.mean().sort_values(ascending=False)

M39 = dict(
    id="3.9", title="Three Dashboard Blueprints",
    subtitle="KPI specifications with current values for SuperMart, Crestview University and Lifecare Hospital.",
    objectives=["Write a dashboard blueprint: audience, KPIs, definitions, visuals and slicers.",
                "Specify a retail sales dashboard.",
                "Specify a university progression dashboard with restricted detail.",
                "Specify a hospital operations dashboard with threshold bands."],
    topics=[
        T("3.9.1", "Blueprint 1: SuperMart Retail Sales Dashboard",
          defn="A dashboard blueprint is the specification for a dashboard: audience, KPIs and their definitions, the visual "
               "for each, and the slicers.",
          simple="Audience: regional managers. Slicers: Zone and Month. KPI cards on top, store attainment bar, zone combo, monthly line.",
          grid=dict(label="KPI specification, FY2025", widths=[2, 1.5, 2.2], rows=[
              ["KPI", "Current value", "Visual"],
              ["Net sales", N(FY25), "KPI card"],
              ["Growth against FY2024", P(FY25 / FY24 - 1), "KPI card with arrow"],
              ["Target attainment", P(ALL_ATT), "Card + bar by store"],
              ["Margin %", P(PROFIT25 / FY25), "Card + combo by zone"],
              ["Loss-making lines", I(LOSS), "Drill-down table"]]),
          watch="Every KPI needs a written definition and source before anyone builds a chart of it."),

        T("3.9.2", "Blueprint 2: Crestview University Progression Dashboard",
          defn="The university blueprint serves the senate and faculty deans, with slicers for Faculty, Department and Level.",
          simple="KPI cards, a bar chart by faculty and a restricted list of at-risk students for advisers only.",
          grid=dict(label="KPI specification, 2024/25", widths=[2, 1.5, 2.2], rows=[
              ["KPI", "Current value", "Visual"],
              ["Students enrolled", I(len(CV)), "KPI card"],
              ["Average exam score", f"{CV.Exam_Score.mean():.1f}", f"Card + bar by faculty ({FAC.index[0]} highest)"],
              ["Average attendance", P(CV.Attendance_Pct.mean() / 100), "Card"],
              ["Grade A share (70+)", P((CV.Exam_Score >= 70).mean()), "Stacked bar by faculty"],
              ["At-risk students", I(AR), "Card + restricted list"]]),
          watch="Named student lists are personal data: restrict them to the people who need them."),

        T("3.9.3", "Blueprint 3: Lifecare Hospital Operations Dashboard",
          defn="The hospital blueprint supports daily operations, with slicers for Department, Month and HMO.",
          simple="An occupancy line with 85% and 95% bands, ALOS by department and average bill by HMO.",
          grid=dict(label="KPI specification, 2023", widths=[2, 1.5, 2.2], rows=[
              ["KPI", "Current value", "Visual"],
              ["Admissions", I(len(LC23)), "Card + monthly line"],
              ["Average occupancy", P(C23.Occupancy_Pct.mean()), "Line with 85% / 95% bands"],
              ["ALOS", f"{LC23.Length_of_Stay.mean():.1f} days", "Bar by department"],
              ["30-day readmission rate", P((LC23.Readmit_30d == 'Yes').mean()), "Card with status text"],
              ["Average bill", N(LC23.Billing_Amount.mean()), "Bar by HMO"]]),
          watch="Occupancy needs the daily census table; it cannot be derived reliably from admissions alone."),
    ],
    takeaways=["A blueprint fixes audience, KPIs, definitions, visuals and slicers before building.",
               "Retail: sales, growth, attainment, margin, losses.",
               "University: enrolment, score, attendance, grade A, at risk (restricted).",
               "Hospital: admissions, occupancy bands, ALOS, readmissions, bill."],
)

# ---------------------------------------------------------------- 3.10
RAW = pd.read_csv(W2D / "Lifecare_Admissions_2023_DIRTY.csv", dtype=str, keep_default_na=False, encoding="utf-8-sig")


def _staging():
    c = sqlite3.connect(":memory:")
    RAW.to_sql("stg_admissions_2023", c, index=False, dtype={k: "TEXT" for k in RAW.columns})
    c.execute(f"ATTACH DATABASE '{D.DBS['lc']}' AS prod")
    return c


STG = _staging()


def sq(sql):
    return pd.read_sql(sql, STG)


STG_ROWS = sq("SELECT COUNT(*) AS n FROM stg_admissions_2023").n[0]
COMMA = sq("SELECT COUNT(*) AS n FROM stg_admissions_2023 WHERE Billing_Amount LIKE '%,%'").n[0]
EX = sq("SELECT Billing_Amount AS b, CAST(Billing_Amount AS REAL) AS c FROM stg_admissions_2023 "
        "WHERE Billing_Amount LIKE '%,%' LIMIT 1").iloc[0]
RULES_SQL = """
SELECT 'Doctor not in doctors' AS rule, COUNT(*) AS n
FROM stg_admissions_2023 s
LEFT JOIN prod.doctors d ON d.doctor_name = s.Doctor_Name
WHERE d.doctor_id IS NULL
UNION ALL
SELECT 'Discharge before admission', COUNT(*)
FROM stg_admissions_2023
WHERE Discharge_Date <> '' AND Discharge_Date < Admission_Date
UNION ALL
SELECT 'Bill below zero', COUNT(*)
FROM stg_admissions_2023
WHERE CAST(REPLACE(Billing_Amount, ',', '') AS REAL) < 0"""
RULES = sq(RULES_SQL)
DUPS = int(RAW.duplicated().sum())
PROD = {r: v(s, "lc") for r, s in [
    ("Discharge before admission", "SELECT COUNT(*) FROM admissions WHERE discharge_date < admission_date"),
    ("Bill below zero", "SELECT COUNT(*) FROM admissions WHERE billing_amount < 0"),
    ("Doctor not in doctors", "SELECT COUNT(*) FROM admissions a LEFT JOIN doctors d ON d.doctor_id = a.doctor_id "
                              "WHERE d.doctor_id IS NULL")]}
PROFILE = {c: (int((RAW[c].str.strip() == "").sum()), RAW[c].replace("", pd.NA).nunique())   # distinct as received
           for c in ("Discharge_Date", "HMO", "Ward", "Gender")}
EMERG = (LC[LC.Admission_Type == "Emergency"].Readmit_30d == "Yes").mean()
ONC = (LC[LC.Department == "Oncology"].Readmit_30d == "Yes").mean()
SRC = pd.read_csv(SEC / "00_Core_Datasets" / "Lifecare_Admissions_2019_2024.csv", usecols=["Billing_Amount"])
SQLREC = q("SELECT COUNT(*) AS n, ROUND(SUM(billing_amount), 2) AS b FROM admissions", "lc").iloc[0]

M310 = dict(
    id="3.10", title="File Extracts into SQL, DAX and Dashboards",
    subtitle="The whole week on one extract: profile it, stage it, report exceptions in SQL, measure it in DAX, reconcile it.",
    objectives=["Profile a file extract before loading it.",
                "Load it unchanged into a staging table and check the row count.",
                "Find orphans and rule breaks with SQL joins and an exception report.",
                "Define explicit DAX measures and test them under filters.",
                "Reconcile file, SQL and dashboard before publishing."],
    topics=[
        T("3.10.1", "Beginner: profile the extract before loading",
          defn="Data profiling summarises each column of a dataset (rows, blanks, distinct values, ranges) to find problems "
               "before the data is used.",
          simple="Lifecare_Admissions_2023_DIRTY.csv (Week2_Data_Files). Problems found now are cheap; problems found on a dashboard are not.",
          grid=dict(label=f"Profile: {I(len(RAW))} rows, {RAW.shape[1]} columns", widths=[1.6, 0.8, 0.9, 2.2], rows=[
              ["Column", "Blank", "Distinct", "Action"],
              ["Discharge_Date", I(PROFILE["Discharge_Date"][0]), I(PROFILE["Discharge_Date"][1]), "Flag; don't guess dates"],
              ["HMO", I(PROFILE["HMO"][0]), I(PROFILE["HMO"][1]), "Label 'Unknown'"],
              ["Ward", I(PROFILE["Ward"][0]), I(PROFILE["Ward"][1]), "Clean to 4 wards"],
              ["Gender", I(PROFILE["Gender"][0]), I(PROFILE["Gender"][1]), "Map to Female / Male"]]),
          watch="Never 'fix' a blank discharge date by guessing; flag it and ask the source system."),

        T("3.10.2", "Beginner: ingest the extract into a staging table",
          defn="A staging table holds data exactly as received, every column as TEXT, before any cleaning or type conversion.",
          simple="Load first, check the row count matches the file, then clean in the next step.",
          **CODE(""".mode csv
.import --skip 1 Lifecare_Admissions_2023_DIRTY.csv stg_admissions_2023
SELECT COUNT(*) FROM stg_admissions_2023;""", label="SQLITE3 SHELL", size=11),
          grid=dict(label="Checks", widths=[2.4, 1.6], size=11.5, rows=[
              ["Check", "Result"],
              ["CSV data rows / staging rows", f"{I(len(RAW))} / {I(STG_ROWS)}"],
              ["Bills written with commas", f"{I(COMMA)}, e.g. '{EX.b}'"],
              [f"CAST('{EX.b}' AS REAL)", f"{EX.c:g} (wrong): remove commas first"]]),
          stack=True,
          watch="Converting types during the import hides the bad values; stage as text, then convert with checks."),

        T("3.10.3", "Intermediate: SQL joins and an exception report",
          defn="An exception report lists, for each business rule, how many rows break it, so problems can be fixed at source.",
          simple="A LEFT JOIN … WHERE … IS NULL finds orphans; UNION ALL stacks one count per rule into a single report.",
          **CODE(RULES_SQL.strip("\n"), label="SQL  ·  STAGING + PROD.DOCTORS", size=10.5),
          grid=dict(label="Exceptions", widths=[2.2, 1, 1], size=11.5, rows=[
              ["Rule", "Staging", "Production"]] + [[r.rule, I(r.n), I(PROD[r.rule])] for r in RULES.itertuples()] + [
              ["Exact duplicate rows", I(DUPS), "0"]]),
          split=0.6,
          watch="Run the same report on production: a rule that fails there too is a system problem, not an extract problem."),

        T("3.10.4", "Intermediate: DAX measures and filter context",
          defn="Explicit measures on the hospital model turn the cleaned admissions into KPIs that respond correctly to filters.",
          simple="Admissions = COUNTROWS; ALOS = AVERAGE of length of stay; Readmission % = DIVIDE of two counts.",
          **CODE("""
Admissions := COUNTROWS ( Admissions )
ALOS := AVERAGE ( Admissions[length_of_stay] )
Readmission % :=
DIVIDE (
    CALCULATE ( [Admissions],
        Admissions[readmit_30d] = "Yes" ),
    [Admissions] )""", label="DAX MEASURES", size=11.5),
          grid=dict(label="2019–2024 results", widths=[2.2, 1.2], rows=[
              ["Measure / filter", "Result"],
              ["[Admissions]", I(len(LC))],
              ["[ALOS]", f"{LC.Length_of_Stay.mean():.1f} days"],
              ["[Readmission %]", P((LC.Readmit_30d == 'Yes').mean())],
              ["… Admission type = Emergency", P(EMERG)],
              ["… Department = Oncology", P(ONC)]]),
          split=0.52,
          watch="The filtered results only make sense if the department and type columns were cleaned in 3.10.1."),

        T("3.10.5", "Advanced practical: reconcile a governed dashboard",
          defn="Reconciliation confirms that the source file, the database query and the dashboard report the same totals "
               "under the same filters.",
          simple="Publish only when all three agree, and record the check.",
          grid=dict(label="Reconciliation, all admissions 2019–2024", widths=[1.4, 1.9, 1.9, 0.8], size=11.5, rows=[
              ["Control", "Source CSV", "SQL", "Result"],
              ["Admissions", I(len(SRC)), I(SQLREC.n), "PASS" if len(SRC) == SQLREC.n else "FAIL"],
              ["Billing (₦)", f"{SRC.Billing_Amount.sum():,.2f}", f"{SQLREC.b:,.2f}",
               "PASS" if abs(SRC.Billing_Amount.sum() - SQLREC.b) < 0.5 else "FAIL"]]),
          points=["Source: 00_Core_Datasets › Lifecare_Admissions_2019_2024.csv",
                  "SQL: COUNT(*) and SUM(billing_amount) on lifecare_hospital.db",
                  "Dashboard: the Admissions card and a billing measure, no slicers"],
          points_label="WHERE EACH NUMBER COMES FROM", stack=True,
          watch="A difference of one row is still a FAIL: find it before anyone sees the dashboard."),
    ],
    takeaways=["Profile first: blanks, distinct values, odd formats.",
               "Stage as text; the row count must match the file.",
               "Exception reports with LEFT JOIN … IS NULL and UNION ALL.",
               "Explicit DAX measures, tested under filters.",
               "File = SQL = dashboard before publishing."],
)

MODULES = [M38, M39, M310]
