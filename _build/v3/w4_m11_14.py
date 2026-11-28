"""Week 4 edition-3 slide copy: Modules 4.11 (AI co-pilots), 4.12 (verifying AI output), 4.13 (data ethics, PII and PHI) and
4.14 (internal benchmarks, PDF checks and reconciliation).
Data: tSal, tDash25, tStores, tLOS, tHotel, tCV, the Week 3 SuperMart Online database, the Kingsway wellbeing survey
(aggregates only), the Week 2 dirty extract and the full SuperMart sales file. Every number is computed here."""
import sqlite3
from pathlib import Path

import numpy as np
import pandas as pd

from v3kit import T
from w4_fx import CV, D25, HOT, I, LOS, N, Nm, P, SAL, STO

PACK = Path(__file__).resolve().parents[2] / "ETEH_Bootcamp_Pack" / "01_Section_Examples"
DB = PACK / "Week3_Data_Files" / "supermart_online.db"

# ---------------------------------------------------------------- SuperMart Online (Week 3 database)
AI_SQL = ("SELECT c.zone, SUM(t.total_price) AS revenue\nFROM transactions t\nJOIN customers c ON c.customer_id = t.customer_id\n"
          "JOIN deliveries d ON d.customer_id = t.customer_id\nGROUP BY c.zone;")
OK_SQL = ("SELECT c.zone, SUM(t.total_price) AS revenue\nFROM transactions t\nJOIN customers c ON c.customer_id = t.customer_id\n"
          "GROUP BY c.zone;")
with sqlite3.connect(DB) as _con:
    SCHEMA = {t: [r[1] for r in _con.execute(f"PRAGMA table_info({t})")]
              for (t,) in _con.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")}
    NROWS = {t: _con.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0] for t in SCHEMA}
    AIR = pd.read_sql_query(AI_SQL, _con).set_index("zone").revenue
    OKR = pd.read_sql_query(OK_SQL, _con).set_index("zone").revenue.sort_values(ascending=False)
    CTRL = _con.execute("SELECT SUM(total_price) FROM transactions").fetchone()[0]
    AIROWS = _con.execute("SELECT COUNT(*) FROM transactions t JOIN deliveries d ON d.customer_id = t.customer_id").fetchone()[0]
    TX = pd.read_sql_query("SELECT order_date, order_status, total_price FROM transactions", _con, parse_dates=["order_date"])
X25 = TX[TX.order_date.dt.year == 2025]
GROSS25 = X25.total_price.sum()
NET25 = X25[X25.order_status.isin(["Delivered", "Completed"])].total_price.sum()

FS = SAL[(SAL.Gender == "Female") & (SAL.Seniority == "Senior")].Annual_Salary_NGN.mean()
FALL = SAL[SAL.Gender == "Female"].Annual_Salary_NGN.mean()

# ---------------------------------------------------------------- SuperMart stores FY2025
ST = (D25.groupby(["Store_ID", "Store_Name"]).agg(Sales=("Sales", "sum"), Profit=("Profit", "sum")).reset_index()
      .merge(STO[["Store_ID", "Floor_Area_sqm", "Annual_Target_2025"]]))
ST["att"] = ST.Sales / ST.Annual_Target_2025
ST["spm"] = ST.Sales / ST.Floor_Area_sqm
MED_SPM = ST.spm.median()
ST["vs"] = ST.spm / MED_SPM - 1
STS = ST.sort_values("spm", ascending=False)
TOT25 = ST.Sales.sum()
ZN = D25.groupby("Zone").Sales.sum()
AIZ = ZN.drop("South-East").sum()
LAGOS = ST[ST.Store_ID.isin(["ST01", "ST02"])].Sales.sum() / TOT25
WORST = ST.sort_values("att").iloc[0]

_m = D25.assign(M=pd.to_datetime(D25.Order_Date).dt.month).pivot_table(index="Store_Name", columns="M", values="Sales",
                                                                         aggfunc="sum", fill_value=0) / 1e6
BUL = pd.DataFrame({"fy": _m.sum(axis=1).round(1), "months": _m.round(1).sum(axis=1).round(1),
                    "ledger": ST.set_index("Store_Name").Sales / 1e6})
BUL["diff"] = BUL.fy - BUL.ledger
BUL["mdiff"] = BUL.months - BUL.fy
BIGM = BUL.mdiff.abs().idxmax()

_sm = pd.read_csv(PACK / "00_Core_Datasets" / "SuperMart_Stores_Sales_2022_2025.csv", usecols=["Order_Date", "Sales"],
                  parse_dates=["Order_Date"])
APRFY = _sm[(_sm.Order_Date >= "2024-04-01") & (_sm.Order_Date <= "2025-03-31")].Sales.sum()
_dirty = pd.read_csv(PACK / "Week2_Data_Files" / "SuperMart_Sales_2025_DIRTY.csv", dtype=str, keep_default_na=False)
QMISS = int(pd.to_numeric(_dirty.Quantity, errors="coerce").isna().sum())
NDIRTY = len(_dirty)

# ---------------------------------------------------------------- Mercyland, Atlantic Pearl, Kingsway
RE = LOS.assign(k=(LOS.Readmitted_30d == "Yes").astype(int)).groupby("Department").agg(n=("k", "size"), k=("k", "sum"))
RE["p"] = RE.k / RE.n
RE["moe"] = 1.96 * np.sqrt(RE.p * (1 - RE.p) / RE.n)
RE = RE.sort_values("p", ascending=False)
TOPD = RE.index[0]
SEPD = [d for d in RE.index[1:] if RE.p[d] + RE.moe[d] < RE.p[TOPD] - RE.moe[TOPD]]
RE_ALL = (LOS.Readmitted_30d == "Yes").mean()
LC_COLS = pd.read_csv(PACK / "00_Core_Datasets" / "Mercyland_Inpatient_Episodes_2019_2024.csv", nrows=1).columns.tolist()
NEED = ["Department", "Readmitted_30d", "Admit_Date"]
CR = (HOT.Reservation_Status == "Cancelled").mean()
R_ATT = CV.Lecture_Attendance_Pct.corr(CV.Final_Score)

_wb = pd.read_csv(PACK / "00_Core_Datasets" / "Kingsway_Student_Welfare_Survey_SENSITIVE.csv")
WB_N = len(_wb)
WBP = (_wb.groupby("Programme").agg(n=("Respondent_ID", "count"), k=("PHQ9_Positive", lambda v: (v == "Yes").sum()))
       .sort_values("n"))
UNIQ1 = int((_wb.groupby(["Gender", "Age", "State_of_Origin", "Programme"]).size() == 1).sum())
UNIQ2 = int((_wb.assign(B=pd.cut(_wb.Age, [0, 20, 25, 30, 100]).astype(str)).groupby(["Gender", "B", "Programme"]).size() == 1).sum())
del _wb

_s2 = SAL[SAL.Gender.isin(["Male", "Female"])]
GAP_ALL = 1 - _s2[_s2.Gender == "Female"].Annual_Salary_NGN.mean() / _s2[_s2.Gender == "Male"].Annual_Salary_NGN.mean()
_g = _s2.groupby(["Seniority", "Gender"]).Annual_Salary_NGN.mean().unstack()
GAPS = 1 - _g.Female / _g.Male
N_OTHER = int((SAL.Gender == "Other").sum())

CLAIMS = [("FY2025 store sales were ₦733.2m", 733.2e6, TOT25, 0.001, "m"),
          ("Lagos stores made 30% of FY2025 sales", 0.30, LAGOS, 0.005, "p"),
          ("Median salary is ₦12.5m", 12.5e6, SAL.Annual_Salary_NGN.median(), 0.001, "m"),
          ("2023 readmission rate was 10.3%", 0.103, RE_ALL, 0.0005, "p"),
          ("Atlantic Pearl cancellation rate is 22%", 0.22, CR, 0.005, "p"),
          ("Attendance-score correlation is 0.81", 0.81, R_ATT, 0.005, "r")]


def _ok(c):
    _, cl, v, tol, _k = c
    return abs(v - cl) <= tol * (abs(cl) if abs(cl) > 1 else 1)


def _f(v, k):
    return Nm(v) if k == "m" else P(v) if k == "p" else f"{v:.2f}"


NOK = sum(_ok(c) for c in CLAIMS)
BN = lambda v: f"₦{v / 1e9:,.2f}bn"  # noqa: E731


M411 = dict(
    id="4.11", title="AI Co-pilots for Analysts",
    subtitle="Using large language models productively and safely: prompt design, schema prompts, formula drafting, and the "
             "limits that must never be crossed.",
    objectives=["Explain what an AI co-pilot can and cannot do for an analyst.",
                "Write structured prompts and schema prompts.",
                "Draft formulas and SQL with AI, then test them.",
                "Use AI to explain code and draft narrative.",
                "Know which data must never be pasted, and follow a safe AI workflow."],
    topics=[
        T("4.11.1", "What an AI Co-pilot Can and Cannot Do",
          defn="An AI co-pilot is a large language model that drafts text, formulas and code from a prompt; it predicts likely "
               "words, it does not look up or calculate facts about your data.",
          simple="Use it to draft and explain; never as the source of a number.",
          grid=dict(label="Good use of AI, or not?", widths=[2.4, 0.8, 1.8], rows=[
              ["Task", "Use AI?", "Why"],
              ["Draft an AVERAGEIFS formula", "Yes", "You can test it on tSal"],
              ["Explain a long formula", "Yes", "You can check each part"],
              ["'What was our 2025 revenue?'", "No", "It has not seen your data"],
              ["Summarise patient records", "No", "Personal health data"]]),
          watch="An AI answer sounds equally confident whether it is right or wrong."),

        T("4.11.2", "Prompt Structure for Analytical Tasks",
          defn="A structured prompt states the role, context, task, output format and constraints.",
          simple="The more specific the prompt, the more testable the answer.",
          grid=dict(label="Prompt: average salary of female Senior respondents", widths=[1.2, 3.6], size=11.5, rows=[
              ["Part", "What we wrote"],
              ["Role", "You are an Excel analyst helping a Nigerian retail chain."],
              ["Context", "Table tSal has columns Seniority, Gender, Annual_Salary_NGN."],
              ["Task", "Write one formula returning the average salary of female Senior respondents."],
              ["Format / limits", "Formula only, then one sentence. Excel 365, structured references."]]),
          watch="'Help me with salaries' gets a generic essay; the structured prompt gets a formula you can test."),

        T("4.11.3", "Schema Prompts",
          defn="A schema prompt gives the AI the table and column names, types, keys and business rules, but no data rows.",
          simple="The AI writes correct column names without seeing any personal data.",
          grid=dict(label="Schema for supermart_online.db (Week 3)", widths=[1.3, 0.9, 1.3], rows=[
              ["Table", "Columns", "Rows (not pasted)"]] + [
              [t, str(len(SCHEMA[t])), I(NROWS[t])] for t in ["transactions", "customers", "deliveries", "products", "sellers"]]),
          watch="Add the business rules too: 'net revenue = Delivered + Completed orders only'."),

        T("4.11.4", "Generating Excel Formulas with AI",
          defn="AI can draft formulas quickly, but each must be tested on data where you know the answer.",
          simple="Test with a filter or a PivotTable: same number, or reject.",
          fx=[("=AVERAGEIFS(tSal[Annual_Salary_NGN],tSal[Gender],\"Female\",tSal[Seniority],\"Senior\")", N(FS), "Correct"),
              ("=AVERAGEIF(tSal[Gender],\"Female\",tSal[Annual_Salary_NGN])", N(FALL), "Drops Seniority")],
          fx_widths=[4.2, 1.1, 1.2], stack=True,
          watch="Both answers look reasonable; only the check shows the second one answers a different question."),

        T("4.11.5", "Generating SQL with AI",
          defn="AI-written SQL often runs without error but joins or filters wrongly, so the result is silently wrong.",
          simple="Check row counts and a control total after every join.",
          grid=dict(label="AI query joined deliveries on customer_id", widths=[1.3, 1.2, 1.2, 0.8], rows=[
              ["Zone", "Correct", "AI", "Times"]] + [[z, Nm(v), Nm(AIR[z]), f"{AIR[z] / v:.1f}x"] for z, v in OKR.head(4).items()]),
          code=AI_SQL, code_label="AI SQL (wrong join)", code_size=10.5, split=0.55,
          watch=f"The AI total is {BN(AIR.sum())}, {AIR.sum() / CTRL:.1f} times the true {BN(CTRL)}: each order repeats once per "
                "delivery that customer ever had."),

        T("4.11.6", "Explaining Code and Formulas with AI",
          defn="AI can explain a formula or query step by step, which helps you understand inherited work.",
          simple="Paste the formula, not the data; then confirm the explanation with the result.",
          grid=dict(label="Explain =INDEX(Store,MATCH(MIN(Attainment),Attainment,0))", widths=[1.7, 3.1], rows=[
              ["Part", "What it does"],
              ["MIN(Attainment)", "Finds the lowest attainment"],
              ["MATCH(...,0)", "Finds its row position, exact match"],
              ["INDEX(Store,...)", f"Returns that store's name: {WORST.Store_Name}"]]),
          watch="If the explanation and the result disagree, trust the result and question the explanation."),

        T("4.11.7", "Using AI for Narrative and Communication",
          defn="AI can draft summaries, emails and slide text from aggregated results you provide.",
          simple="Give it the totals, not the records; then check every number in the draft.",
          grid=dict(label="Checking an AI draft against the store table", widths=[2.8, 1.8], rows=[
              ["Sentence in the draft", "Check"],
              [f"{int((ST.att < 1).sum())} of 12 stores missed target", f"Matches: {int((ST.att < 1).sum())}"],
              [f"Sales were {Nm(TOT25)}", f"Matches: {Nm(TOT25)}"],
              [f"{WORST.Store_Name} is lowest", f"Matches: {P(WORST.att, 0)}"]]),
          watch="AI drafts often 'round' numbers into new ones; check each figure, not just the tone."),

        T("4.11.8", "Prohibited AI Use: Confidential and Personal Data",
          defn="Never paste personal, health, financial or confidential data into an AI tool that your organisation has not "
               "approved for that data.",
          simple="Structure and totals are usually safe; individual records never are.",
          grid=dict(label="Which data can be pasted?", widths=[2.4, 0.8, 1.6], rows=[
              ["Data", "Paste?", "Instead"],
              ["Patient names, Patient_No", "Never", "Schema only"],
              ["Customer email, phone", "Never", "Schema only"],
              ["Table and column names", "Yes", "-"],
              ["12 store totals", "If approved", "Aggregates only"]]),
          watch="The Nigeria Data Protection Act 2023 applies to what you paste, wherever the AI server is."),

        T("4.11.9", "AI Limitations Every Analyst Must Know",
          defn="AI models invent functions, columns and statistics, miss edge cases, and do not know your data or recent changes.",
          simple="Each limitation is easy to catch if you test.",
          grid=dict(label="Common AI mistakes", widths=[1.5, 1.6, 1.7], rows=[
              ["Limitation", "Example", "How to catch it"],
              ["Invented function", "SUMIFSX", "#NAME? error"],
              ["Invented column", "tSal[Salary]", "#REF! or wrong result"],
              ["Wrong join", "4.11.5", "Control total"],
              ["Made-up statistic", "'22% cancel'", f"Recompute: {P(CR)}"]]),
          watch="'The AI said so' is never evidence."),

        T("4.11.10", "Building an AI-Assisted Workflow",
          defn="An AI-assisted workflow fixes the steps: schema prompt, ask, run on a copy, verify, log.",
          simple="The verification step is not optional.",
          grid=dict(label="The workflow on the zone revenue query", widths=[1.5, 3.3], rows=[
              ["Step", "What happened"],
              ["1 Schema prompt", f"{len(SCHEMA)} tables, no rows"],
              ["2 Ask, 3 Run", f"Zone revenue SQL; {I(AIROWS)} joined rows"],
              ["4 Verify", f"{BN(AIR.sum())} vs control total {BN(CTRL)}: fail"],
              ["5 Log", "Discarded; join rewritten; check logged"]]),
          watch="The workflow caught the error before anyone saw the wrong numbers."),
    ],
    takeaways=["AI drafts and explains; it is never the source of a number.",
               "Structured prompts and schema prompts get testable answers.",
               "Test every AI formula and query against a known answer.",
               "Never paste personal or confidential data.",
               "Follow the workflow: schema, ask, run, verify, log."],
)


M412 = dict(
    id="4.12", title="Verifying AI Output and Handling Hallucinations",
    subtitle="A verification protocol that makes AI output safe to use: independent checks, provenance, disclosure and the "
             "discipline of the control total.",
    objectives=["Define a hallucination and why it happens.",
                "Apply a verification protocol to AI claims.",
                "Use control totals to catch silent errors.",
                "Spot plausible but wrong output and decide when to discard it.",
                "Keep provenance, sign off and use a checklist."],
    topics=[
        T("4.12.1", "What a Hallucination Is",
          defn="A hallucination is AI output that is fluent and confident but false: an invented number, source, function or fact.",
          simple="The model predicts plausible text; plausible is not the same as true.",
          grid=dict(label="Six claims an AI assistant made about our data", widths=[2.6, 1.0, 1.0], size=11, rows=[
              ["AI claim", "Real value", "Verdict"]] + [[c[0], _f(c[2], c[4]), "Right" if _ok(c) else "WRONG"] for c in CLAIMS]),
          watch="Right and wrong claims look identical until you check them."),

        T("4.12.2", "Verification Protocol for AI Output",
          defn="A verification protocol recomputes every AI claim from the source data with a formula and compares within a "
               "stated tolerance.",
          simple="Claimed value, verified value, tolerance, verdict: four columns.",
          fx=[("=IF(ABS([@Verified]-[@Claimed])<=[@Tolerance],\"VERIFIED\",\"WRONG\")", f"{NOK} of {len(CLAIMS)} pass", "Verdict per claim")],
          fx_widths=[4.2, 1.1, 1.2],
          steps=["List each claim and its value", "Recompute it from the table", "Set a tolerance (rounding)", "Flag WRONG; never use it"],
          stack=True,
          watch="Rounded-but-true claims should pass; use a tolerance, not exact equality."),

        T("4.12.3", "The Control Total Discipline",
          defn="A control total is a figure taken from the source (row count, grand total) that any derived table must add back to.",
          simple="Take it from the ledger, never from the summary you are checking.",
          grid=dict(label="An AI-built zone summary vs the ledger", widths=[2.6, 1.8], rows=[
              ["Measure", "Value"],
              ["Control total: FY2025 sales (ledger)", N(TOT25)],
              ["Sum of the AI zone table", N(AIZ)],
              ["Difference", N(TOT25 - AIZ)]]),
          watch="The difference is exactly South-East's sales: the AI silently dropped a zone. Discard and redo."),

        T("4.12.4", "Spotting Plausible but Wrong Output",
          defn="Warning signs of a hallucination: suspiciously round numbers, figures with no source, columns that do not exist, "
               "and results that are too neat.",
          simple="Treat a warning sign as a reason to check, not proof of error.",
          grid=dict(label="Warning signs in the AI's claims", widths=[1.6, 1.8, 1.2], rows=[
              ["Warning sign", "Seen in", "Truth"],
              ["Round number", "'22% cancellations'", P(CR)],
              ["Strong, neat link", "'r = 0.81'", f"{R_ATT:.2f}"],
              ["Close but off", "'Median ₦12.5m'", Nm(SAL.Annual_Salary_NGN.median())]]),
          watch="Being 'about right' is the most dangerous kind of wrong: it passes a quick glance."),

        T("4.12.5", "Provenance and Audit Trails for AI-Assisted Work",
          defn="Provenance records where each output came from: the prompt, the tool, the check performed and the outcome.",
          simple="A provenance log sheet with one row per AI interaction.",
          grid=dict(label="Provenance log", widths=[1.1, 1.3, 1.6, 0.9], size=11, rows=[
              ["When", "Prompt", "Check", "Outcome"],
              ["2026-01-12 10:40", "Zone revenue SQL", f"{I(AIROWS)} rows; {BN(AIR.sum())}", "Discarded"],
              ["2026-01-12 11:05", "Explain RANK.EQ ties", "Matched Microsoft docs", "Used"]]),
          watch="Disclose AI use in the report's method section; hiding it undermines trust when it is discovered."),

        T("4.12.6", "When to Discard AI Output Entirely",
          defn="Discard AI output when it invents things, fails verification repeatedly, or you cannot explain it; rewrite it "
               "yourself.",
          simple="Patching a fundamentally wrong answer takes longer than starting again.",
          grid=dict(label="Patch or discard?", widths=[2.6, 1.2], rows=[
              ["Condition", "Action"],
              ["Invents columns or functions", "Discard"],
              ["Fails the control total twice", "Discard"],
              ["You cannot explain how it works", "Discard"],
              ["Small formatting issue", "Patch"]]),
          watch=f"The rewritten zone query matches the {N(CTRL)} control total."),

        T("4.12.7", "Responsibility and Sign-Off",
          defn="The analyst who signs the work is responsible for every number in it, however it was produced.",
          simple="Sign-off means 'I checked this', not 'I produced this'.",
          grid=dict(label="Sign-off block for the AI-assisted claims check", widths=[1.5, 3.3], rows=[
              ["Item", "Entry"],
              ["Prepared by", "A. Analyst"],
              ["Checks", f"{len(CLAIMS)} claims recomputed; control total matched"],
              ["Reviewed by / date", "B. Reviewer / 2026-01-13"]]),
          watch="'The AI made the mistake' is not a defence; the signature is yours."),

        T("4.12.8", "A Verification Checklist for AI-Assisted Analysis",
          defn="A verification checklist lists the checks every AI-assisted analysis passes before release.",
          simple="Tick with evidence, not from memory.",
          grid=dict(label="Final checklist for the zone revenue analysis", widths=[1.6, 2.4, 0.7], rows=[
              ["Check", "Evidence", "Pass?"],
              ["Schema only", "No data rows pasted", "Yes"],
              ["Claims recomputed", "tClaims verdict column", "Yes"],
              ["Control total", f"Matches {Nm(CTRL)}", "Yes"],
              ["Provenance logged", "Log sheet rows", "Yes"]]),
          watch="One 'No' stops the release until it is fixed."),
    ],
    takeaways=["Hallucinations are fluent, confident and false.",
               "Recompute every claim with a formula and a tolerance.",
               "Control totals from the source catch silent errors.",
               "Discard output you cannot verify or explain.",
               "Log provenance; your signature means you checked."],
)


M413 = dict(
    id="4.13", title="Data Ethics, PII and PHI",
    subtitle="Consent, minimisation, anonymisation, bias, retention and governance — the obligations that travel with every "
             "dataset you touch.",
    objectives=["Explain why data ethics is part of an analyst's job.",
                "Classify personal (PII) and health (PHI) data.",
                "Minimise, anonymise and aggregate data with small-cell suppression.",
                "Check purpose, legal basis, bias and fairness.",
                "Apply retention, secure handling and governance."],
    topics=[
        T("4.13.1", "Why Data Ethics Belongs in the Bootcamp",
          defn="Data ethics is the set of principles for collecting, using and sharing data fairly, lawfully and safely.",
          simple="Every bootcamp dataset describes real kinds of people; mistakes harm them.",
          grid=dict(label="Who is in our datasets?", widths=[1.8, 1.3, 1.7], rows=[
              ["Dataset", "Who", "Harm if mishandled"],
              ["Mercyland admissions", "Patients", "Health details exposed"],
              ["Kingsway wellbeing", "Students", "Mental-health stigma"],
              ["SuperMart Online", "Customers", "Fraud, spam"],
              ["Salary survey", "Employees", "Pay revealed"]]),
          watch="Nigeria Data Protection Act 2023: fines and liability apply to organisations and their staff."),

        T("4.13.2", "What Counts as Personal Data (PII)",
          defn="Personal data (PII) is any information that identifies a person directly (name, email) or indirectly in "
               "combination (date of birth plus city).",
          simple="Direct identifiers: remove. Quasi-identifiers: generalise.",
          grid=dict(label="SuperMart Online customers table", widths=[2.2, 1.3, 1.6], size=11, rows=[
              ["Columns", "Type", "Rule"],
              ["first_name, last_name, email, phone_number", "Direct", "Never leave the database"],
              ["date_of_birth, city", "Quasi", "Band age; use zone"],
              ["customer_id", "Pseudonym", "Analyse by this only"]]),
          watch="Removing names is not enough: date of birth plus city can single out one customer."),

        T("4.13.3", "What Counts as Health Data (PHI)",
          defn="Health data (PHI) is information about a person's physical or mental health or care; it is special-category data "
               "with stricter rules.",
          simple="Health data never travels with names attached.",
          grid=dict(label=f"Mercyland admissions ({len(LC_COLS)} columns)", widths=[2.3, 1.2, 1.5], size=11, rows=[
              ["Columns", "Class", "Rule"],
              ["Patient_Name, Patient_No", "Direct (PHI)", "Remove before extract"],
              ["Diagnosis_Group, Medication, Lab_Result", "Health", "Aggregate; min 5 per cell"],
              ["Age, Gender, Admit_Date", "Quasi", "Band or use year"]]),
          watch="Mental-health screening answers (Kingsway) are among the most sensitive data in the course."),

        T("4.13.4", "Data Minimisation",
          defn="Data minimisation means collecting and extracting only the columns and rows needed for the stated question.",
          simple="Start from the question, list the columns it needs, drop the rest.",
          grid=dict(label="'Which departments have the highest readmission rate?'", widths=[2.6, 1.4], rows=[
              ["Measure", "Value"],
              ["Columns in the admissions extract", str(len(LC_COLS))],
              ["Columns needed", f"{len(NEED)}: Department, Readmitted_30d, year"],
              ["Columns to drop", str(len(LC_COLS) - len(NEED))]]),
          watch=f"{len(NEED)} columns answer the question; the other {len(LC_COLS) - len(NEED)} only add risk."),

        T("4.13.5", "Anonymisation and Aggregation",
          defn="Anonymisation removes or generalises identifiers so people cannot be re-identified; aggregation publishes counts, "
               "with small cells (under 5) suppressed.",
          simple="Band exact values, drop rare fields, and replace counts under 5 with '<5'.",
          grid=dict(label=f"Kingsway wellbeing survey ({I(WB_N)} responses, aggregates only)", widths=[3.2, 1.2], rows=[
              ["Check", "Result"],
              [f"Smallest programme ({WBP.index[0]})", f"{WBP.n.iloc[0]} respondents"],
              ["Unique on gender + exact age + state + programme", I(UNIQ1)],
              ["Unique after banding age and dropping state", I(UNIQ2)]]),
          watch="No programme count needs suppressing, but exact fields would let a classmate identify thousands of people."),

        T("4.13.6", "Consent, Purpose and Legal Basis",
          defn="Each use of personal data needs a lawful basis (consent, contract, legal duty, vital or public interest, "
               "legitimate interest) and must fit the purpose it was collected for.",
          simple="'We already have the data' is not a reason to use it for something new.",
          grid=dict(label="Purpose and basis", widths=[1.4, 2.0, 1.6], size=11, rows=[
              ["Dataset", "Typical basis", "Check before new use"],
              ["Mercyland", "Provision of care", "Ethics approval"],
              ["SuperMart Online", "Contract", "Marketing needs consent"],
              ["Wellbeing survey", "Consent", "Only the stated research"]]),
          watch="Using survey data for a purpose the respondents were not told about breaks their consent."),

        T("4.13.7", "Bias and Fairness in Analysis",
          defn="Bias is a systematic difference in how groups are treated or represented; a fairness check compares outcomes "
               "across groups at comparable levels.",
          simple="Compare like with like, and report small groups honestly.",
          grid=dict(label="Gender pay gap as a fairness test (tSal)", widths=[2.4, 1, 1.4], rows=[
              ["Measure", "Value", "Meaning"],
              ["Raw gap, all levels", P(GAP_ALL), "Headline figure"],
              [f"Largest within-level gap ({GAPS.idxmax()})", P(GAPS.max()), "Investigate"],
              ["'Other' gender respondents", str(N_OTHER), "Too few to report"]]),
          watch=f"The headline {P(GAP_ALL)} hides a {P(GAPS.max())} gap at {GAPS.idxmax()} level; do not silently drop small groups."),

        T("4.13.8", "Retention and Secure Handling",
          defn="Retention rules say how long data is kept and when it is deleted; secure handling covers storage, sharing and "
               "transfer.",
          simple="Fewer copies, fewer leaks.",
          grid=dict(label="Handling rules for the bootcamp data", widths=[1.4, 3.4], rows=[
              ["Control", "Practice"],
              ["Retention", "Training extracts deleted at the end of the bootcamp"],
              ["Storage", "Encrypted laptop; no personal data on USB sticks"],
              ["Sharing", "Restricted links, never email attachments"],
              ["AI tools", "No personal data in public AI tools"]]),
          watch="Most leaks come from forgotten copies: old extracts, email attachments, downloads folders."),

        T("4.13.9", "Governance and Accountability",
          defn="Data governance assigns roles and approvals: who owns the data, who may access it, who approves new uses.",
          simple="Every extract can be traced to an approval.",
          grid=dict(label="Roles for a Mercyland readmission study", widths=[1.6, 3.2], rows=[
              ["Role", "Action"],
              ["Data owner", f"Approves the {len(NEED)}-column extract"],
              ["Analyst", "Documents the minimisation and checks"],
              ["Data Protection Officer", "Reviews the new use"],
              ["Ethics committee", "Approves research use"]]),
          watch="If nobody approved the extract, nobody is accountable when it leaks."),
    ],
    takeaways=["Every dataset describes people; ethics is part of the job.",
               "Remove direct identifiers; generalise quasi-identifiers.",
               "Minimise to the question; suppress small cells.",
               "Check purpose, legal basis and fairness.",
               "Retain briefly, handle securely, govern every extract."],
)


M414 = dict(
    id="4.14", title="Internal Benchmarks, PDF Checks and Reconciliation",
    subtitle="Benchmark each SuperMart store against the chain median, verify the PDF bulletin against the ledger, align "
             "periods and definitions, show uncertainty and write a decision memo.",
    objectives=["Benchmark stores fairly against the chain median.",
                "Verify a PDF report against the ledger.",
                "Reconcile time periods and definitions.",
                "Show uncertainty and missing data honestly.",
                "Write a one-page benchmark decision memo."],
    topics=[
        T("4.14.1", "Beginner: understand benchmark fitness",
          defn="A fit benchmark uses the same definition, the same period and comparable units; for a chain, the median store is "
               "a robust internal benchmark.",
          simple="Sales per m² removes the size effect: big is not the same as efficient.",
          grid=dict(label=f"FY2025 sales per m² (chain median {N(MED_SPM)})", widths=[2.2, 1.3, 1.1], rows=[
              ["Store", "Sales per m²", "vs median"]] + [
              [r.Store_Name, N(r.spm), f"{r.vs:+.1%}"] for _, r in pd.concat([STS.head(3), STS.tail(2)]).iterrows()]),
          watch=f"{STS.iloc[0].Store_Name} is the most efficient store, not Lagos - Ikeja: Ikeja is big, "
                f"{STS.iloc[0].Store_Name.split(' - ')[0]} sells more per m²."),

        T("4.14.2", "Beginner: extract and verify the PDF",
          defn="PDF verification extracts the printed table (Get Data › From PDF) and compares each figure with the ledger within "
               "a rounding tolerance.",
          simple="Printed ₦m to 1 decimal: a difference up to ±₦0.05m is rounding.",
          grid=dict(label="SuperMart_Monthly_Sales_Bulletin_2025.pdf vs ledger (₦m)", widths=[2.0, 0.9, 1.0, 0.9, 0.9], size=11, rows=[
              ["Store", "Bulletin FY", "Months added", "Ledger", "Status"]] + [
              [n, f"{BUL.fy[n]:.1f}", f"{BUL.months[n]:.1f}", f"{BUL.ledger[n]:.2f}",
               "OK" if abs(BUL['diff'][n]) <= 0.05 + 1e-9 else "CHECK"] for n in ["Lagos - Ikeja", BIGM]]),
          watch=f"{int((BUL['diff'].abs() <= 0.05 + 1e-9).sum())} of {len(BUL)} stores agree within ±₦0.05m. Adding printed months can be ₦{abs(BUL.mdiff[BIGM]):.1f}m out from rounding "
                "alone: reconcile to the ledger, not to a sum of rounded numbers."),

        T("4.14.3", "Intermediate: reconcile time periods and definitions",
          defn="Two figures are comparable only if they cover the same period and use the same definition (gross or net, with or "
               "without VAT, cancelled orders in or out).",
          simple="Label every number with its period and definition.",
          grid=dict(label="Same company group, different '2025 sales'", widths=[2.4, 1.3, 1.6], size=11, rows=[
              ["Measure", "Value", "Definition"],
              ["SuperMart Stores FY2025", Nm(TOT25), "Jan-Dec 2025"],
              ["SuperMart Stores Apr 2024-Mar 2025", Nm(APRFY), "April year-end"],
              ["SuperMart Online 2025 gross", Nm(GROSS25), "All orders"],
              ["SuperMart Online 2025 net", Nm(NET25), "Delivered + Completed"]]),
          watch=f"Comparing one firm's gross with another's net overstates by {GROSS25 / NET25 - 1:.0%}: align definitions first."),

        T("4.14.4", "Intermediate: uncertainty and missing data",
          defn="Uncertainty is how far a figure could be from the truth; a 95% interval for a rate is rate ± 1.96 x "
               "SQRT(rate x (1 - rate) / n).",
          simple="Overlapping intervals mean the difference may not be real: do not rank them.",
          grid=dict(label="2023 readmission rate by department (tLOS)", widths=[2.2, 1, 0.9, 1.4], size=11, rows=[
              ["Department", "Admissions", "Rate", "95% interval"]] + [
              [d, I(r.n), P(r.p), f"{P(r.p - r.moe)} - {P(r.p + r.moe)}"] for d, r in RE.iterrows()]),
          watch=f"{TOPD}'s interval sits above {len(SEPD)} departments' intervals; the others overlap, so do not rank them. "
                f"Missing data: {QMISS} of {I(NDIRTY)} lines in the Week 2 extract have no quantity, so report units as a range."),

        T("4.14.5", "Advanced practical: benchmark decision memo",
          defn="A decision memo states the decision needed, the benchmark used, the finding, its uncertainty and a recommendation, "
               "on one page.",
          simple="Decision first; evidence next; limits and next step last.",
          grid=dict(label="Memo: which stores to review for the FY2026 refurbishment budget?", widths=[1.3, 3.5], size=11, rows=[
              ["Section", "Content"],
              ["Benchmark", "Chain median sales per m², FY2025 ledger reconciled to the PDF within ±₦0.05m"],
              ["Finding", f"{STS.iloc[-1].Store_Name} lowest: {N(STS.iloc[-1].spm)} per m², {abs(STS.iloc[-1].vs):.0%} below median"],
              ["Uncertainty", "One year of data; store age not adjusted; figures unaudited"],
              ["Recommendation", "Site review for the bottom three stores before allocating budget"]]),
          watch="An example decision for practice: the memo recommends a review, not a final allocation."),
    ],
    takeaways=["Benchmark on the same definition, period and units; use the median.",
               "Verify a PDF against the ledger within a rounding tolerance.",
               "Label every figure with its period and definition.",
               "Show intervals; do not rank overlapping ones; state missing data.",
               "One-page memo: decision, evidence, limits, next step."],
)

for _m in (M411, M412, M413, M414):
    for _t in _m["topics"]:
        if _t.get("fx"):
            _t.setdefault("stack", True)

MODULES = [M411, M412, M413, M414]
