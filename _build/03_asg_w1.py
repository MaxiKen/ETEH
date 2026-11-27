"""Week 1 assignment - SuperMart Q1 2025 store review (modules 1.1-1.9). Builds student file + answer key."""
from __future__ import annotations

import pandas as pd

from asg_common import Asg
from wk_common import DATASETS_TEXT, L, sm_sales

sm = sm_sales()
q1 = sm[(sm.Order_Date >= "2025-01-01") & (sm.Order_Date < "2025-04-01")].copy()
q1 = q1.sort_values(["Order_Date", "Row_ID"]).reset_index(drop=True)
Q1COLS = ["Row_ID", "Order_ID", "Order_Date", "Ship_Date", "Ship_Mode", "Customer_ID", "Customer_Name", "Segment", "Store_ID",
          "Category", "Sub_Category", "Product_Name", "Sales", "Quantity", "Discount", "Profit"]
stores = L("sm_stores").copy()
stores["Opened"] = pd.to_datetime(stores.Opened)
STCOLS = ["Store_ID", "Store_Name", "City", "State", "Zone", "Store_Manager", "Annual_Target_2025"]
PHASE = 0.15
VAT = 0.075

# ---- truth used for cached answers
q1["Unit_Price"] = q1.Sales / q1.Quantity
q1["Margin_Pct"] = (q1.Profit / q1.Sales).where(q1.Sales != 0, 0)
q1["Days_to_Ship"] = (q1.Ship_Date - q1.Order_Date).dt.days
q1["Month"] = q1.Order_Date.dt.strftime("%b")
smap = stores.set_index("Store_ID")
q1["Store_Name"] = q1.Store_ID.map(smap.Store_Name)

sc = stores[["Store_ID"]].copy()
g = q1.groupby("Store_ID")
sc_sales = sc.Store_ID.map(g.Sales.sum()).fillna(0)
sc_profit = sc.Store_ID.map(g.Profit.sum()).fillna(0)
sc_target = sc.Store_ID.map(smap.Annual_Target_2025) * PHASE
sc_att = sc_sales / sc_target
sc_status = ["On target" if a >= 1 else ("Watch" if a >= 0.8 else "Behind") for a in sc_att]
sc_rank = sc_sales.rank(ascending=False, method="min").astype(int)


def build(key: bool):
    A = Asg(1, "ETEH Week 1 Assignment - SuperMart Q1 2025 Store Review", key,
            "Week 1 · Foundations of analytics and Excel · SuperMart Retail, Q1 2025")
    b = A.b

    # ---------------------------------------------------------------- T1 analytics thinking
    s = A.task("T1 Analytics Thinking", "Task 1 - Analytics thinking: question types, process and industry use", "1.1-1.4",
               "SuperMart's Regional Manager has sent you a list of questions from across ETEH's client organisations. Before touching any data, "
               "classify each question, put an investigation into the right order and connect analytics to real decisions.")
    s.band("1.2", "Classify each question (dropdown in the yellow cell)")
    types = "Descriptive,Diagnostic,Predictive,Prescriptive"
    qs = [("What were SuperMart's total Q1 2025 sales for each of its 12 stores?", "Descriptive", "Reports what happened."),
          ("Abuja - Wuse had the 4th-highest Q1 2025 sales of the 12 stores but made a loss. Why?", "Diagnostic", "Explains WHY something happened."),
          ("How many patients will Lifecare General Hospital admit next December?", "Predictive", "Estimates a future value."),
          ("At what discount level should SuperMart cap promotions to protect a 12% margin?", "Prescriptive", "Recommends an action / optimum."),
          ("Which Crestview University faculty had the highest at-risk rate last semester?", "Descriptive", "A past fact, sliced by a category."),
          ("Why are Online Travel Agent bookings at Eko Suites cancelled more often than corporate bookings?", "Diagnostic", "Cause behind an observed gap."),
          ("What salary should a data analyst with 5 years' experience expect, based on the ETEH Talent Salary Survey?", "Predictive", "Model-based estimate (regression, Week 4)."),
          ("How many nurses should each Lifecare ward roster next week, given the expected bed census?", "Prescriptive", "Turns a forecast into a decision.")]
    for i, (t, a, gd) in enumerate(qs, 1):
        A.q(s, f"1.{i}", t, a, kind="gen", marks=1, guide=gd, choices=types.split(","), header=(i == 1))
    s.gap()
    s.band("1.3", "Put the Abuja - Wuse loss investigation into the correct order of the analytics process")
    s.para("Enter the step number (1 = first, 6 = last) beside each activity. Activities are shown out of order.", "note")
    acts = [("Build a bar chart of Q1 profit by store with Abuja - Wuse highlighted", 5),
            ("Agree the question with the Regional Manager: why did Abuja - Wuse lose money in Q1 2025?", 1),
            ("Fix date formats, remove duplicate lines and check the totals against Finance", 3),
            ("Present a one-page recommendation to cap discounts at Abuja - Wuse", 6),
            ("Export the Q1 2025 sales lines for all 12 stores from the point-of-sale system", 2),
            ("Compare discount depth and margin by category for Wuse against the other stores", 4)]
    for i, (t, v) in enumerate(acts):
        A.q(s, chr(97 + i), t, v, kind="int", marks=1, choices=["1", "2", "3", "4", "5", "6"], header=(i == 0),
            guide="Ask → Collect → Clean → Analyse → Visualise → Act/Share" if i == 0 else "")
    s.band("1.4", "Analytics in industry")
    A.long(s, "1.9", "For EACH of Lifecare General Hospital, Crestview University and Eko Suites Hotel: name ONE key metric (KPI) and ONE decision that metric supports.",
           "Lifecare: bed occupancy % (or average length of stay / 30-day readmission rate) → how many nurses to roster per ward; when to open overflow beds.\n"
           "Crestview: progression rate / % of students at risk → which students get tutoring; which courses need extra support.\n"
           "Eko Suites: cancellation rate by channel (or occupancy, ADR, RevPAR) → overbooking level, deposit policy for OTA bookings, room pricing.",
           marks=6, guide="1 mark per sensible KPI, 1 per linked decision (3 × 2). Accept any defensible metric with the decision it drives.")
    A.end_task(s)

    # ---------------------------------------------------------------- T2 working with data
    s = A.task("T2 Working with Data", "Task 2 - Navigating and interrogating a real extract", "1.5-1.6",
               "Sheet D_Q1_Sales holds every SuperMart sales line for January-March 2025 (Excel Table tSalesQ1). Answer each question. "
               "You may use Sort, Filter, Find, the status bar or a formula - but the number in the yellow cell must be exact.")
    n_disc = int((q1.Discount >= 0.4).sum())
    imax = q1.Sales.idxmax()
    rows = [
        ("How many sales lines are in the Q1 2025 extract?", "=ROWS(tSalesQ1)", len(q1), "int", "Status bar Count on one column also works (500)."),
        ("What were total Q1 2025 sales (₦)?", "=SUM(tSalesQ1[Sales])", q1.Sales.sum(), "naira", "Exact to the naira."),
        ("How many lines belong to store ST02 (Lagos - Lekki)?", '=COUNTIF(tSalesQ1[Store_ID],"ST02")', int((q1.Store_ID == "ST02").sum()), "int", "Filter Store_ID = ST02."),
        ("What is the largest single sales line (₦)?", "=MAX(tSalesQ1[Sales])", q1.Sales.max(), "naira", "Sort Sales Z→A."),
        ("Which product was sold on that line?", "=INDEX(tSalesQ1[Product_Name],MATCH(MAX(tSalesQ1[Sales]),tSalesQ1[Sales],0))", q1.loc[imax, "Product_Name"], "gen", "XLOOKUP(MAX(..),..) equally valid."),
        ("How many lines carried a discount of 40% or more?", '=COUNTIF(tSalesQ1[Discount],">=0.4")', n_disc, "int", "Number filter ≥ 0.4."),
        ("What was the first order date at Abuja - Wuse (ST07) in Q1?", '=MINIFS(tSalesQ1[Order_Date],tSalesQ1[Store_ID],"ST07")', q1[q1.Store_ID == "ST07"].Order_Date.min(), "date", "Filter ST07, sort oldest → newest."),
        ("How many DISTINCT orders are in the extract?", "=SUMPRODUCT(1/COUNTIF(tSalesQ1[Order_ID],tSalesQ1[Order_ID]))", q1.Order_ID.nunique(), "int", "Copy column → Remove Duplicates also gives 240. 500 lines ≠ 500 orders."),
        ("How many DISTINCT customers bought in Q1?", "=SUMPRODUCT(1/COUNTIF(tSalesQ1[Customer_ID],tSalesQ1[Customer_ID]))", q1.Customer_ID.nunique(), "int", "COUNTA(UNIQUE(tSalesQ1[Customer_ID])) in Excel 365."),
    ]
    for i, (t, f, v, k, gd) in enumerate(rows, 1):
        A.q(s, f"2.{i}", t, f, v, k, marks=2, guide=gd, header=(i == 1))
    A.q(s, "2.10", "Which keyboard shortcut jumps from the top of the Sales column straight to its last filled cell?", "Ctrl + Down Arrow", kind="gen", marks=2,
        guide="Ctrl+Shift+Down selects the column; Ctrl+End goes to the last used cell of the sheet (half mark).",
        choices=["Ctrl + Down Arrow", "Ctrl + End", "Ctrl + Shift + L", "Alt + ="])
    A.end_task(s)

    # ---------------------------------------------------------------- T3 formulas
    s = A.task("T3 Formulas & Functions", "Task 3 - Store scorecard with lookups, conditional aggregation and absolute references", "1.7",
               "Build the Q1 2025 scorecard for the 12 stores. Every yellow cell must be a LIVE formula that refers to tSalesQ1, tStores (sheet D_Stores) "
               "or the input cells below. Q1 target = Annual_Target_2025 × the Q1 phasing input (absolute reference!).")
    inp = s.inputs([("Q1 phasing of the annual target", PHASE, "pct", "Q1 delivered 15.4%, 14.6% and 15.3% of annual sales in 2022-2024"),
                    ("VAT rate (Nigeria)", VAT, "pct", "Standard VAT rate used for the VAT-inclusive questions")])
    PH, VR = inp["Q1 phasing of the annual target"], inp["VAT rate (Nigeria)"]
    s.band("3.1", "Scorecard (fill every yellow cell with a formula that can be copied down)")
    top = s.r + 1
    n = len(sc)
    r1, r2 = top + 1, top + n
    ans = {
        "Store_Name": ("=XLOOKUP(B{r},tStores[Store_ID],tStores[Store_Name])", stores.Store_Name.tolist(), "gen"),
        "Zone": ("=XLOOKUP(B{r},tStores[Store_ID],tStores[Zone])", stores.Zone.tolist(), "gen"),
        "Q1_Sales": ("=SUMIFS(tSalesQ1[Sales],tSalesQ1[Store_ID],B{r})", sc_sales.tolist(), "naira"),
        "Q1_Profit": ("=SUMIFS(tSalesQ1[Profit],tSalesQ1[Store_ID],B{r})", sc_profit.tolist(), "naira"),
        "Margin": ("=IFERROR(F{r}/E{r},0)", (sc_profit / sc_sales).tolist(), "pct"),
        "Q1_Target": ("=XLOOKUP(B{r},tStores[Store_ID],tStores[Annual_Target_2025])*" + PH, sc_target.tolist(), "naira"),
        "Attainment": ("=E{r}/H{r}", sc_att.tolist(), "pct"),
        "Status": ('=IF(I{r}>=1,"On target",IF(I{r}>=0.8,"Watch","Behind"))', sc_status, "gen"),
        "Sales_Rank": (f"=RANK.EQ(E{{r}},$E${r1}:$E${r2})", sc_rank.tolist(), "int"),
    }
    A.fill(s, sc, ans, marks=15,
           guide="9 columns. Store_Name/Zone 2, SUMIFS 4, Margin 1, Target with $-locked phasing 3, Attainment 1, nested IF 2, RANK.EQ with locked range 2.")
    s.ws.set_column(2, 2, 26)
    s.ws.set_column(9, 9, 11)
    s.para("Status rule: On target if Attainment ≥ 100%; Watch if ≥ 80%; otherwise Behind.", "note")
    s.band("3.2", "Questions (formulas required)")
    A.q(s, "3.2", "VAT-inclusive Q1 sales for Lagos - Ikeja (use the VAT input with an absolute reference).", f"=E{r1}*(1+{VR})",
        sc_sales.iloc[0] * (1 + VAT), "naira", 2, "Must reference the VAT cell, not 7.5% typed in.", header=True)
    A.q(s, "3.3", "Total VAT due on all Q1 2025 sales.", f"=SUM(tSalesQ1[Sales])*{VR}", q1.Sales.sum() * VAT, "naira", 2, "")
    A.q(s, "3.4", "How many stores are 'Behind'?", f'=COUNTIF(J{r1}:J{r2},"Behind")', sc_status.count("Behind"), "int", 2, "COUNTIF on the Status column.")
    A.q(s, "3.5", "How many lines had a NEGATIVE profit AND a discount of at least 40%?", '=COUNTIFS(tSalesQ1[Profit],"<0",tSalesQ1[Discount],">=0.4")',
        int(((q1.Profit < 0) & (q1.Discount >= 0.4)).sum()), "int", 2, "Compare with 2.6: nearly every deep-discount line loses money.")
    A.q(s, "3.6", "Average sales per line for the Corporate segment, rounded to the nearest naira.", '=ROUND(AVERAGEIFS(tSalesQ1[Sales],tSalesQ1[Segment],"Corporate"),0)',
        round(q1[q1.Segment == "Corporate"].Sales.mean()), "naira", 2, "ROUND(...,0). Unrounded answer: half mark.")
    A.end_task(s)

    # ---------------------------------------------------------------- T4 tables
    s = A.task("T4 Excel Tables", "Task 4 - Calculated columns and structured references in tSalesQ1", "1.8",
               "On sheet D_Q1_Sales, five new columns have been added to the right of table tSalesQ1 (yellow). Fill each with ONE calculated-column formula "
               "(type it in the first row - the Table copies it down). Then answer the check questions with structured references.")
    s.grid([["Unit_Price", "Sales ÷ Quantity", " =[@Sales]/[@Quantity]"],
            ["Margin_Pct", "Profit ÷ Sales (0 if Sales is 0)", " =IF([@Sales]=0,0,[@Profit]/[@Sales])"],
            ["Days_to_Ship", "Ship_Date minus Order_Date", " =[@Ship_Date]-[@Order_Date]"],
            ["Month", "Three-letter month of Order_Date (Jan, Feb, Mar)", ' =TEXT([@Order_Date],"mmm")'],
            ["Store_Name", "Store name looked up from tStores", " =XLOOKUP([@Store_ID],tStores[Store_ID],tStores[Store_Name])"]]
           if key else
           [["Unit_Price", "Sales ÷ Quantity", ""], ["Margin_Pct", "Profit ÷ Sales (0 if Sales is 0)", ""], ["Days_to_Ship", "Ship_Date minus Order_Date", ""],
            ["Month", "Three-letter month of Order_Date (Jan, Feb, Mar)", ""], ["Store_Name", "Store name looked up from tStores", ""]],
           header=["New column", "Definition", "Formula used (key)" if key else "Write the formula you used"], merge_last=True)
    A.q(s, "4.1", "All five columns created as calculated columns on D_Q1_Sales (instructor inspects).", "5 columns, one formula each" if key else "", marks=5,
        guide="1 mark per column. Zero for a column of typed values or a formula that changes row to row.", header=True)
    A.q(s, "4.2", "Average days to ship, to 2 decimals.", "=ROUND(AVERAGE(tSalesQ1[Days_to_Ship]),2)", round(q1.Days_to_Ship.mean(), 2), "dec2", 2)
    A.q(s, "4.3", "How many lines have a negative Margin_Pct?", '=COUNTIF(tSalesQ1[Margin_Pct],"<0")', int((q1.Margin_Pct < 0).sum()), "int", 2)
    A.q(s, "4.4", "Total March sales using the Month column.", '=SUMIFS(tSalesQ1[Sales],tSalesQ1[Month],"Mar")', q1[q1.Month == "Mar"].Sales.sum(), "naira", 2)
    A.q(s, "4.5", "Total Q1 profit of 'Abuja - Wuse' using the Store_Name column.", '=SUMIFS(tSalesQ1[Profit],tSalesQ1[Store_Name],"Abuja - Wuse")',
        q1[q1.Store_Name == "Abuja - Wuse"].Profit.sum(), "naira", 2, "Negative: the loss that triggered Task 1.")
    A.q(s, "4.6", "Highest unit price in the extract.", "=MAX(tSalesQ1[Unit_Price])", q1.Unit_Price.max(), "naira", 2)
    A.end_task(s)

    # ---------------------------------------------------------------- T5 tidy data
    s = A.task("T5 Tidy Data", "Task 5 - Fix an untidy manager report", "1.9",
               "A zone manager e-mailed the report below. It looks tidy on screen but breaks several tidy-data laws, so it cannot be filtered, "
               "pivoted or joined. Diagnose it, then rebuild it as a tidy table driven by formulas on tSalesQ1.")
    s.band("5.0", "The report as received")
    q1["M"] = q1.Order_Date.dt.month
    rep = []
    for z in ["South-West", "South-South", "North-Central", "South-East"]:
        ids = stores[stores.Zone == z].Store_ID
        for sid in ids:
            v = [round(q1[(q1.Store_ID == sid) & (q1.M == m)].Sales.sum() / 1e6, 1) for m in (1, 2, 3)]
            rep.append([f"{smap.Store_Name[sid]} ({z})", *v, round(sum(v), 1), ""])
        tot = [round(q1[(q1.Store_ID.isin(ids)) & (q1.M == m)].Sales.sum() / 1e6, 1) for m in (1, 2, 3)]
        rep.append([f"{z.upper()} SUBTOTAL", *tot, round(sum(tot), 1), "see e-mail" if z == "North-Central" else ""])
    s.grid(rep, header=["Store (Zone)", "Jan Sales (₦m)", "Feb Sales (₦m)", "Mar Sales (₦m)", "Q1 Total", "Notes"], fmts=[None, "dec1", "dec1", "dec1", "dec1", None], c0=2)
    s.ws.set_column(2, 2, 40)
    A.long(s, "5.1", "List FOUR tidy-data problems in this report and the fix for each.",
           "1) Two variables in one column - 'Store (Zone)' → split into Store and Zone.\n"
           "2) A variable (Month) stored in column headers - Jan/Feb/Mar columns → unpivot to one Month column and one Sales column.\n"
           "3) Subtotal rows mixed with data rows → delete them; totals belong in a PivotTable / SUM, never in the data.\n"
           "4) Units and rounding baked into headers and values (₦m, 1 dp) → store exact naira; format for display only.\n"
           "Also accept: derived 'Q1 Total' column; free-text Notes column with no meaning; no store key (Store_ID).",
           marks=4, guide="1 mark per valid problem + fix (max 4).")
    s.band("5.2", "Rebuild it tidy: one row per store per month (formulas on tSalesQ1)")
    tid = pd.DataFrame([(sid, pd.Timestamp(2025, m, 1)) for sid in stores.Store_ID for m in (1, 2, 3)], columns=["Store_ID", "Month_Start"])
    zones = tid.Store_ID.map(smap.Zone).tolist()
    sales = [q1[(q1.Store_ID == a) & (q1.M == d.month)].Sales.sum() for a, d in zip(tid.Store_ID, tid.Month_Start)]
    top = s.r + 1
    A.fill(s, tid, {"Zone": ("=XLOOKUP(B{r},tStores[Store_ID],tStores[Zone])", zones, "gen"),
                    "Sales": ('=SUMIFS(tSalesQ1[Sales],tSalesQ1[Store_ID],B{r},tSalesQ1[Order_Date],">="&C{r},tSalesQ1[Order_Date],"<"&EDATE(C{r},1))', sales, "naira")},
           fmts={"Month_Start": "date"}, marks=12,
           guide="Zone lookup 4; SUMIFS with a date window (>= start, < EDATE(start,1)) 8. Exact naira, not ₦m.")
    r1, r2 = top + 1, top + len(tid)
    A.q(s, "5.3", "Control check: tidy-table total minus SUM(tSalesQ1[Sales]). What should it be?", f"=SUM(E{r1}:E{r2})-SUM(tSalesQ1[Sales])", 0, "naira", 2,
        "Must be 0. The e-mailed report's ₦m rounding makes its total drift.", header=True)
    A.q(s, "5.4", "In the tidy table, what does ONE row represent?", "One store in one month", kind="gen", marks=2,
        choices=["One store in one month", "One sales line", "One zone in one quarter", "One month for all stores"])
    A.end_task(s)

    # ---------------------------------------------------------------- data sheets
    calc = {
        "Unit_Price": ("=[@Sales]/[@Quantity]", "naira", q1.Unit_Price.tolist()),
        "Margin_Pct": ("=IF([@Sales]=0,0,[@Profit]/[@Sales])", "pct", q1.Margin_Pct.tolist()),
        "Days_to_Ship": ("=[@Ship_Date]-[@Order_Date]", "int", q1.Days_to_Ship.tolist()),
        "Month": ('=TEXT([@Order_Date],"mmm")', "gen", q1.Month.tolist()),
        "Store_Name": ("=XLOOKUP([@Store_ID],tStores[Store_ID],tStores[Store_Name])", "gen", q1.Store_Name.tolist()),
    }
    ds = b.sheet("D_Q1_Sales", "SuperMart sales lines, Q1 2025 (table tSalesQ1)", "500 lines · one row per product line of an order · ₦ values", "Data sheet - 500 rows")
    ds.r = 4
    A.build_table(ds, q1[Q1COLS], "tSalesQ1", calc, fmts={"Order_Date": "date", "Ship_Date": "date", "Sales": "naira", "Profit": "naira", "Discount": "pct0", "Quantity": "int"},
                  widths={"Product_Name": 40, "Customer_Name": 18}, style="Table Style Light 9")
    ds.ws.freeze_panes(5, 0)
    b.data_sheet("D_Stores", stores[STCOLS], "tStores", fmts={"Annual_Target_2025": "naira"}, title="SuperMart store directory (table tStores)",
                 widths={"Store_Name": 28, "Store_Manager": 18})

    return A.finish("Week 1 Assignment - SuperMart Q1 2025 Store Review",
                    ["Scenario: you have just joined SuperMart Retail's head-office analytics team in Ikeja. The Regional Manager wants a Q1 2025 store review "
                     "before Monday's management meeting - and is worried that Abuja - Wuse lost money.",
                     "Work through Tasks 1-5 in order. Each task names the slide modules it assesses. Answers go in the yellow cells; marks are shown beside every question.",
                     "Time guide: 4-5 hours. Everything you need is inside this workbook (tSalesQ1 sales lines and tStores directory)."],
                    [("SuperMart Stores", DATASETS_TEXT["SuperMart Stores"] + " - this file holds the Q1 2025 subset (500 lines).")],
                    rubric=[["Live formulas", "Where a formula is asked for, typed numbers score zero even if correct."],
                            ["Tolerance", "Money exact to ₦1 unless rounding is asked for; percentages ±0.1 pt."],
                            ["Follow-through", "If an earlier error is carried forward correctly, later marks are still awarded."],
                            ["Presentation", "Do not move cells. Number formats are already set."]])


if __name__ == "__main__":
    for k in (False, True):
        print("Week 1", "KEY" if k else "student", "->", build(k))
