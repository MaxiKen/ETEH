"""Capstone edition-3 slide copy: Modules P.8 (portfolio, anonymisation and CV) and P.9 (technical interviews and the roadmap
beyond Excel). Figures from w5_fx."""
from slidekit import NAVY, ORANGE
from v3kit import T
from w5_fx import (BAD_N, BAD_TOT, BD_LONG, CEN_DAYS, DLOG2, DM, EASY, EP_ALL, G25, LINES, LONG, M24, M25, MED, MH_TABLES,
                   N_LOSSSKU, NCV, NRAW2, NRISK, NSTORES, OCC, OK_N, OK_TOT, PROG, Q4M25, QD, RESCUED, S25, SKUS, SMO_TABLES,
                   SQL_EASY, SQL_MED, TX_N, XL_D40, XL_SS, XL_TOPM, YEAR, Nb, Nm, I, N, P)
from w5_fx import C3

FREQ = C3.R["Q5_frequent_attenders"]

M_P8 = dict(
    id="P.8", title="Portfolio, Anonymisation and CV",
    subtitle="Turning three capstones into evidence that gets interviews: write-ups, packaging, anonymisation rules and CV "
             "bullets with numbers you can defend.",
    objectives=["Explain why a portfolio beats a certificate.", "Write a six-heading project write-up.",
                "Anonymise and package files for reviewers.", "Write honest, numeric CV bullets and tool descriptions.",
                "Present yourself online, answer portfolio questions and keep it current."],
    topics=[
        T("P.8.1", "Why a Portfolio Beats a Certificate",
          defn="A certificate says you finished a course. A portfolio shows what you built, how you thought and what you found, "
               "and a recruiter can check it.",
          grid=dict(label="Three capstones, three headline numbers", widths=[1.6, 2.2, 2.2], size=12, rows=[
              ["Project", "Shows", "Headline number"],
              ["SuperMart (retail)", "Cleaning, modelling, diagnosis", f"Margin {P(M24)} › {P(M25)} on {P(G25)} growth"],
              ["Kingsway (education)", "Definitions, at-risk list, estimate", f"{P(PROG)} progression vs 90% target"],
              ["Mercyland (health)", "SQL, capacity, safety statistics", f"Occupancy {P(OCC)} vs 85% target"]])),

        T("P.8.2", "Structuring a Project Write-Up",
          defn="Each write-up has six headings: question, data, method, finding, recommendation, outcome. Add one dashboard "
               "screenshot and the reconciliation table.",
          grid=dict(label="Capstone 1 write-up", widths=[1.4, 4.5], size=11.5, rows=[
              ["Heading", "Content"],
              ["Question", f"Why did margin fall from {P(M24)} to {P(M25)} while sales grew {P(G25, 0)}?"],
              ["Data", f"{I(LINES)} POS lines (2022-2025), {NSTORES} stores, {I(SKUS)} SKUs; returns cleaned and reconciled"],
              ["Method", "Power Query cleaning, star schema, DAX, rate/mix decomposition"],
              ["Finding", f"Q4 2025 margin fell to {P(Q4M25)} as discount depth rose to {P(QD[(2025, 'Q4')])}"],
              ["Recommendation", "Cap Furniture discounts at 20%; review loss-making SKUs"],
              ["Outcome", "Board adopted the cap for Q1 2026 (training scenario)"]])),

        T("P.8.3", "Anonymising Portfolio Artefacts",
          defn="Never publish data that identifies a person. Replace names with labels, publish only totals, suppress small "
               "groups and blur names in screenshots.",
          grid=dict(label="Rules with Mercyland examples", widths=[1.6, 1.6, 2.7], size=12, rows=[
              ["Artefact", "Rule", "Example"],
              ["Frequent attenders", "Counts only", f"{I(FREQ.patients.sum())} patients admitted 3+ times in {YEAR}"],
              ["Escalation list", "Episode_ID only", "No names, no patient numbers"],
              ["Small groups", "Suppress cells under 5", f"{I(FREQ.patients.iloc[-1])} patients with {FREQ.admissions_2023.iloc[-1]} admissions: suppress"],
              ["Screenshots", "Crop or blur", "Name columns never visible"]]),
          watch="'The names are fictional' is no excuse: practise the habit you will need with real data."),

        T("P.8.4", "Packaging Files for Reviewers",
          defn="Send reviewers one folder in reading order, every file named with its version and date.",
          grid=dict(label="The portfolio folder", widths=[2.6, 3.3], size=12.5, rows=[
              ["File", "Content"], ["00_README.txt", "What is inside and how to open it"],
              ["01_Writeup_v1.0_2026-01-20.pdf", "The six-heading write-up"],
              ["02_Dashboard_v1.0_2026-01-20.pdf", "Dashboard export, names blurred"],
              ["03_Workbook_sample_v1.0.xlsx", "Workbook with a small anonymised sample"]])),

        T("P.8.5", "Writing CV Bullets for Analytics Roles",
          defn="A strong CV bullet is action verb + method + scale + result, and every number is defensible from your workbook.",
          grid=dict(label="Weak › strong", widths=[1.6, 4.3], size=11.5, rows=[
              ["Weak", "Strong"],
              ["Used Excel for analysis", f"Built a star-schema model over {I(LINES)} transactions, reconciling {Nm(S25)} with zero variance"],
              ["Found some problems", f"Diagnosed a {-DM:.1f}-point margin decline with rate/mix analysis; flagged {I(N_LOSSSKU)} loss-making SKUs"],
              ["Made a student report", f"Defined a progression measure for {I(NCV)} students and an at-risk list of {I(NRISK)} by trigger"],
              ["Did some SQL", f"Caught a join error that inflated revenue to {Nb(BAD_TOT)} (true: {Nb(OK_TOT)})"]])),

        T("P.8.6", "Describing Tools Honestly",
          defn="List only tools you can demonstrate, describe the level honestly, and say how you used AI and how you checked "
               "its output.",
          grid=dict(label="Describe honestly as", widths=[1.2, 4.7], size=12.5, rows=[
              ["Tool", "Level you can demonstrate"], ["Excel", "Power Query, Data Model, DAX, PivotTables, what-if, Solver"],
              ["SQL", "SELECT to multi-table JOIN, subqueries, window functions (SQLite)"],
              ["Statistics", "Descriptive statistics, correlation, regression, FORECAST.ETS"],
              ["AI", "Co-pilot drafting, every output verified against control totals"]]),
          watch="An overstated skill is found out in the first technical interview."),

        T("P.8.7", "LinkedIn and Professional Presence",
          defn="LinkedIn is a route to your portfolio: a clear headline, a short About section naming your three projects, and "
               "an occasional post sharing one insight.",
          grid=dict(label="Profile elements", widths=[1.2, 4.7], size=12.5, rows=[
              ["Element", "Example"], ["Headline", "Data Analyst | Excel, Power Query, DAX, SQL | Retail, Health, Education"],
              ["About", "Three projects: retail margin, student progression, hospital capacity"],
              ["Post", f"'Why {LONG}+ day stays fill {P(BD_LONG, 0)} of hospital bed-days' (anonymised training data)"]])),

        T("P.8.8", "Preparing for Portfolio Questions",
          defn="Expect 'walk me through it', 'why this measure?', 'how did you verify it?' and 'what was hardest?'. Answer "
               "with specifics from your build.",
          grid=dict(label="Answer shapes", widths=[2.0, 3.9], size=12.5, rows=[
              ["Question", "Answer shape"], ["Walk me through it", "Question › KPI cards › one chart › recommendation, in 60 s"],
              ["Why this measure?", "Margin = total profit ÷ total sales, because it matches the P&L"],
              ["How did you verify it?", "Seven reconciliation tests, all PASS, ₦0 difference"],
              ["What was hardest?", f"Cleaning the Kingsway export: {I(NRAW2)} rows, {I(DLOG2['Duplicate students'])} double registrations"]])),

        T("P.8.9", "Keeping the Portfolio Current",
          defn="Keep the portfolio current: add about one project a quarter, retire work you would no longer defend, and update "
               "write-ups when your methods improve.",
          grid=dict(label="A maintenance rhythm", widths=[1.8, 1.4, 2.6], size=12.5, rows=[
              ["Action", "Frequency", "Signal to a recruiter"], ["Add a project", "Quarterly", "You keep learning"],
              ["Retire weak work", "Each review", "You judge your own work"], ["Refresh numbers and links", "Twice a year", "You are careful"]])),
    ],
    takeaways=["A portfolio is checkable evidence; a certificate is not.", "Six headings: question, data, method, finding, "
               "recommendation, outcome.", "Never publish anything that identifies a person.",
               "CV bullets: verb + method + scale + result, all defensible.", "Keep the portfolio current."],
)

M_P9 = dict(
    id="P.9", title="Technical Interviews and the Roadmap Beyond Excel",
    subtitle="Interview formats, live Excel and SQL tests, behavioural questions, and a concrete route into Power BI, SQL and "
             "Python.",
    objectives=["Recognise the four interview formats.", "Work a live Excel test aloud.", "Answer SQL screens of rising difficulty.",
                "Explain your process and answer with STAR.", "Plan the next steps in Power BI, SQL and Python."],
    topics=[
        T("P.9.1", "Common Analytics Interview Formats",
          defn="Most analytics interviews use four formats: a take-home case, a live Excel test, a SQL screen and a behavioural "
               "interview.",
          grid=dict(label="Formats and how to practise", widths=[1.6, 2.0, 2.3], size=12.5, rows=[
              ["Format", "Tests", "Practise with"], ["Take-home case", "End-to-end analysis", "The three capstones"],
              ["Live Excel test", "Cleaning, PivotTables, speed", "Weekly assignments; the P.9 sheet"],
              ["SQL screen", "Filters, joins, windows", "supermart_online.db"], ["Behavioural", "Judgement, teamwork", "STAR stories"]])),

        T("P.9.2", "The Live Excel Test",
          defn="A live Excel test usually means: clean a file, build a PivotTable, answer a few questions and explain one number.",
          simple="Say what you are doing as you do it. Interviewers score the process as much as the answer.",
          fx=[('=SUMIFS(tSM[Sales],tSM[Year],2025,tSM[Zone],"South-South")', N(XL_SS), "FY2025 South-South sales"),
              ("PivotTable: Store × Margin, sorted", f"{XL_TOPM[0]} ({P(XL_TOPM[1])})", "Highest-margin store"),
              ('=COUNTIFS(tSM[Year],2025,tSM[Discount],">=0.4") / COUNTIFS(tSM[Year],2025)', P(XL_D40), "Lines at 40%+ discount")],
          fx_widths=[3.9, 1.6, 1.5], stack=True,
          watch="Explain one number: 'these deep-discount lines are mostly loss-making, which is why margin fell'."),

        T("P.9.3", "SQL Interview Questions",
          defn="SQL screens rise in difficulty: filter and aggregate, then a JOIN (state the cardinality first), then a window "
               "function.",
          code=SQL_MED.split("\n"), code_label="MEDIUM: A MANY-TO-ONE JOIN", code_size=11,
          grid=dict(label="Results on supermart_online.db", widths=[1.5, 0.8, 1.2], size=11.5, rows=[
              ["Brand", "Orders", "Net revenue"]] + [[r.brand_name, I(r.orders), Nm(r.net_revenue)] for r in MED.itertuples()]),
          order=["code", "grid"], split=0.6,
          more=[dict(title="Easy: Filter and Aggregate", pill="EASY",
                     chart=dict(kind="bar", full=True, h=3.6, cats=EASY.payment_method.tolist()[::-1],
                                series=[("Revenue", EASY.revenue.tolist()[::-1])], fmt='₦#,##0,,"m"', labels=True,
                                label_fmt='₦#,##0.0,,"m"', legend=False, label="DELIVERED REVENUE BY PAYMENT METHOD",
                                title=f"{EASY.payment_method.iloc[0]} leads: {I(EASY.orders.iloc[0])} delivered orders",
                                note="SELECT payment_method, COUNT(*), SUM(total_price) FROM transactions WHERE order_status = "
                                     "'Delivered' GROUP BY payment_method ORDER BY 3 DESC"),
                     notes="The easy question: one table, one filter, one GROUP BY. Ask students to say the grain aloud before "
                           "typing.")]),

        T("P.9.4", "Explaining Your Analytical Process",
          defn="'How do you approach a new dataset?' The answer that lands is a sequence: understand the question, profile the "
               "data, check quality, build, verify, communicate.",
          grid=dict(label="Profiling the Kingsway export, then deciding", widths=[2.0, 0.8, 3.0], size=12, rows=[
              ["Profile finding", "Count", "Decision"], ["Rows received", I(NRAW2), f"Reconcile to the {I(NCV)} registry cohort"],
              ["Padded IDs", I(DLOG2["Padded keys"]), "TRIM before de-duplicating"],
              ["Double registrations", I(DLOG2["Duplicate students"]), "Remove on the trimmed ID"],
              ["Scores above 100", I(DLOG2["Out-of-range score"]), "Query, then cap at 100"]])),

        T("P.9.5", "Behavioural Questions with STAR",
          defn="Answer behavioural questions with STAR: Situation, Task, Action, Result. Your capstone work is the material.",
          grid=dict(label="'Tell me about a time you caught an error'", widths=[1.1, 4.8], size=11.5, rows=[
              ["STAR", "Answer"],
              ["Situation", f"An AI-written zone query returned {Nb(BAD_TOT)} revenue ({I(BAD_N)} rows) from {I(TX_N)} transactions"],
              ["Task", "Verify it before it went into the deck"],
              ["Action", f"Checked rows with a delivery ({I(OK_N)}) and their total ({Nb(OK_TOT)}); found a join on customer_id; "
                         "rewrote it on delivery_id"],
              ["Result", "Correct figure published; a control-total check added to every SQL deliverable"]]),
          watch="Keep the Result about the outcome and what changed, not about you being clever."),

        T("P.9.6", "The Roadmap to Power BI",
          defn="Power BI is the natural next step: the Data Model, DAX and star-schema skills transfer directly; the new parts "
               "are the visual layer and publishing.",
          grid=dict(label="What transfers and what is new", widths=[1.6, 4.3], size=12.5, rows=[
              ["Item", "Detail"], ["Transfers", "Star schema, DAX measures, chart choices, slicers"],
              ["New", "Report pages, visual interactions, publishing, row-level security"],
              ["First project", "Rebuild Capstone 1 in Power BI Desktop (2-4 weeks)"]])),

        T("P.9.7", "The Roadmap to SQL and Databases",
          defn="Grow SQL from SELECT and JOIN to CTEs, window functions and indexes, then move to a server database such as "
               "PostgreSQL.",
          grid=dict(label="Practice databases in the pack", widths=[2.0, 1.4, 2.4], size=12.5, rows=[
              ["Database", "Tables", "Rows to practise on"],
              ["supermart_online.db", I(len(SMO_TABLES)), f"{I(TX_N)} transactions"],
              ["mercyland_hospital.db", I(len(MH_TABLES)), f"{I(EP_ALL)} admissions; {I(CEN_DAYS)} census days"]]),
          code=SQL_EASY.split("\n"), code_label="START HERE, THEN ADD A JOIN AND A WINDOW", code_size=11.5,
          order=["grid", "code"], stack=True),

        T("P.9.8", "The Roadmap to Python",
          defn="Python adds scale, automation and deeper statistics: pandas for tables, matplotlib for charts.",
          simple="Your most valuable asset is not syntax; it is knowing what to check. That transfers to every tool.",
          grid=dict(label="A first Python project", widths=[1.4, 4.5], size=12.5, rows=[
              ["Library", "Use it to"], ["pandas", "Re-create the SuperMart cleaning log from the dirty CSV"],
              ["matplotlib", "Draw the quarterly margin chart"], ["Time", "6-8 weeks alongside work"]]),
          chart=dict(kind="column", cats=["Power BI", "SQL", "Python"], series=[("Weeks (upper estimate)", [4, 6, 8])], fmt="0",
                     labels=True, legend=False, colors=[NAVY], point_colors=[ORANGE, NAVY, NAVY],
                     title="Suggested order and time (weeks)", h=2.4),
          order=["grid", "chart"], split=0.55),
    ],
    takeaways=["Prepare for four formats: case, Excel, SQL, behavioural.", "Work aloud: process is scored.",
               "State the join cardinality before writing the JOIN.", "STAR stories come from your capstones.",
               "Next: Power BI, then SQL, then Python."],
)

MODULES = [M_P8, M_P9]
