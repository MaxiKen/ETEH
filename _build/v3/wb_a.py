"""Week 1 workbook, modules 1.1 - 1.3 (edition 3)."""
from __future__ import annotations

import numpy as np
import pandas as pd

from w1x import ModuleSheet

N = "naira"


def _im(series, value):
    """Index of the row holding `value` (for INDEX/MATCH cached answers)."""
    return int(np.flatnonzero(series.values == value)[0])


# =====================================================================================================
def m11(b, D):
    ST, IK, CAT = D.ST, D.IK, D.CAT
    m = ModuleSheet(b, "1.1 Intro to Analytics", "1.1  Introduction to Data Analytics",
                    "Tutor walk-through after Module 1.1: data and information, analytics, roles, sharp questions, "
                    "domain knowledge and deciding with evidence", "1.1.1 - 1.1.14")
    m.dataset([("tIkejaLines", "On this sheet, below", "The first 12 FY2025 sales lines of SuperMart Lagos - Ikeja. One row = one product line on an order."),
               ("tCat25", "On this sheet, below", "FY2025 sales and profit for each product category, with the 2024 margin to compare against."),
               ("tStores25", "Sheet Data_Stores25 (starts in A1)", "FY2025 results for the 12 SuperMart stores. One row = one store. "
                                                                    "Margin = Profit ÷ Sales and Attainment = Sales ÷ Target are calculated columns.")])
    m.h("tIkejaLines")
    m.s.table(IK, "tIkejaLines", fmts={"Order_Date": "date", "Sales": N, "Profit": N, "Quantity": "int"},
              widths={"Order_ID": 16, "Customer_Name": 17, "Sub_Category": 13, "Category": 15})
    m.h("tCat25")
    m.s.table(CAT[["Category", "Sales", "Profit", "Margin_2024"]], "tCat25",
              fmts={"Sales": N, "Profit": N, "Margin_2024": "pct"},
              calc={"Margin": ("=[@Profit]/[@Sales]", "pct", CAT.Margin.tolist())})

    # ---------------------------------------------------------------- 1.1.1-3
    m.topic("1.1.1-3", "Data, information and the difference between them",
            "Data are raw facts with no context, such as one till line. Information is data that has been processed, "
            "summarised or compared so that it means something to the reader. The facts do not change; the context does.")
    m.steps(["Look at one row of tIkejaLines: an order number, a date, an amount. On its own it answers no question. That is data.",
             "Click any empty cell and type =SUM(tIkejaLines[Sales]), then press Enter. Twelve facts become one number with meaning: that is information.",
             "Add context: =SUM(tIkejaLines[Profit])/SUM(tIkejaLines[Sales]) gives the margin on those lines. Press Ctrl+Shift+5 to show it as a %.",
             "Compare with a yardstick: on Data_Stores25, cell J2 is Lagos - Ikeja's Attainment (sales ÷ target). A number becomes a finding when you compare it with something."])
    m.worked([("Rows of data in tIkejaLines", "=ROWS(tIkejaLines)", len(IK), "int", "Twelve separate facts."),
              ("Total sales on those lines", "=SUM(tIkejaLines[Sales])", IK.Sales.sum(), N, "Information: one number that answers 'how much?'"),
              ("Total profit", "=SUM(tIkejaLines[Profit])", IK.Profit.sum(), N, ""),
              ("Margin on those lines", "=SUM(tIkejaLines[Profit])/SUM(tIkejaLines[Sales])", IK.Profit.sum() / IK.Sales.sum(), "pct",
               "Profit kept from every ₦1 of sales."),
              ("Lagos - Ikeja FY2025 attainment", "=Data_Stores25!J2", ST.Attainment[0], "pct", "Sales compared with the target: 103.1% means target beaten.")])
    m.turn()
    m.ex("How many items (Quantity) were sold on the 12 lines of tIkejaLines?", "=SUM(tIkejaLines[Quantity])", int(IK.Quantity.sum()), "int",
         short="Total quantity on tIkejaLines")
    m.ex("How many of the 12 lines are Technology? (Hint: COUNTIF.)", '=COUNTIF(tIkejaLines[Category],"Technology")',
         int((IK.Category == "Technology").sum()), "int", short="Technology lines")
    for item, ans in [("'SM-2025-144463, 2025-01-01, ₦474,430'", "Data"),
                      ("'Lagos - Ikeja reached 103.1% of its FY2025 target'", "Information"),
                      ("'Five of the 12 stores lost money in FY2025'", "Information"),
                      ("'11' in the Quantity column", "Data")]:
        m.ex_text(f"Data or Information? {item}", ans)

    # ---------------------------------------------------------------- 1.1.4-5
    loss = ST.Profit[ST.Profit < 0].sum()
    m.topic("1.1.4-5", "What data analytics is, and why we analyse data",
            "Data analytics is examining, cleaning, transforming and summarising data to answer questions and support decisions. "
            "We analyse to see what is really happening, spot problems early and act on evidence rather than opinion.")
    m.steps(["Ask a question: which stores lost money in FY2025?",
             "Find the data that answers it: the Profit column of tStores25 (sheet Data_Stores25).",
             "Analyse: type =COUNTIF(tStores25[Profit],\"<0\") to count the loss-making stores.",
             "Size the problem: =SUMIF(tStores25[Profit],\"<0\") adds up only the losses.",
             "Act: on Data_Stores25 click the Profit filter arrow › Number Filters › Less Than › 0 to see the five stores by name."])
    m.worked([("Stores that made a loss", '=COUNTIF(tStores25[Profit],"<0")', int((ST.Profit < 0).sum()), "int", "The question, answered."),
              ("Total of those losses", '=SUMIF(tStores25[Profit],"<0")', loss, N, "How big the problem is."),
              ("Chain profit, all 12 stores", "=SUM(tStores25[Profit])", ST.Profit.sum(), N, "The losses are hidden inside a healthy total.")])
    m.turn()
    m.ex("How many stores met or beat their FY2025 target (Attainment of 1, i.e. 100%, or more)?", '=COUNTIF(tStores25[Attainment],">=1")',
         int((ST.Attainment >= 1).sum()), "int", short="Stores at or above target")
    m.ex("What was the total profit of the stores that made a profit?", '=SUMIF(tStores25[Profit],">0")', ST.Profit[ST.Profit > 0].sum(), N,
         short="Profit of profitable stores")

    # ---------------------------------------------------------------- 1.1.6-7
    m.topic("1.1.6-7", "Data analytics in everyday life and in organisations",
            "The same cycle (question › data › answer › decision) runs when you choose a route to work and when a hospital plans its beds.")
    m.steps(["Name the decision (for example: which route to take to work).",
             "Name the data that informs it (live traffic speeds).",
             "Name the measure you compare (minutes per route).",
             "State the decision the measure supports (take the faster route)."], label="How to describe any analytics use")
    m.s.grid([["Bank app spending summary", "Your debit transactions by category", "Share of spending per category", "Cut the biggest category"],
              ["SuperMart Retail", "tStores25", "Attainment = Sales ÷ Target", "Support the stores below 100%"],
              ["Kingsway University", "tStudents", "At-risk rate by faculty", "Target tutoring at the highest rate"],
              ["Atlantic Pearl Hotel", "Bookings by market segment", "Cancellation rate", "Ask for deposits on high-cancellation bookings"]],
             header=["Situation / organisation", "Data", "Measure", "Decision it supports"])
    m.s.r += 1
    m.turn()
    m.ex_open("Describe one decision you made this week using data: the decision, the data and the measure.",
              "I chose between two data plans by comparing price per GB (data: the price list; measure: ₦ per GB; decision: the cheaper plan per GB).")
    m.ex_open("Pick one organisation in the table and write one question its managers could answer with its data.",
              "Atlantic Pearl Hotel: what share of online travel agent bookings were cancelled in 2023-2024, compared with corporate bookings?")

    # ---------------------------------------------------------------- 1.1.8-9
    m.topic("1.1.8-9", "The data analyst's role, and how it differs from a data scientist and a business analyst",
            "A data analyst turns data into answers about what happened and why. A data scientist builds models that predict or automate. "
            "A business analyst designs how a process or system should work.")
    m.steps(["Read the task and ask: is the output a report of what happened (analyst), a predictive model (scientist), or a process design (business analyst)?",
             "Check the tools: Excel, Power Query, SQL and dashboards point to the analyst; Python and machine learning to the scientist; process maps and requirements to the business analyst."],
            label="How to tell the roles apart")
    m.s.grid([["Weekly sales dashboard for store managers", "Data Analyst"],
              ["Model that scores every discharge for 30-day readmission risk", "Data Scientist"],
              ["Map how students register and write the requirements for a new system", "Business Analyst"]],
             header=["Task", "Role"])
    m.s.r += 1
    m.turn()
    for task, role in [("Find out why Furniture margin fell in FY2025", "Data Analyst"),
                       ("Build a model that predicts which students will fail", "Data Scientist"),
                       ("Design the steps of a new hotel check-in process", "Business Analyst"),
                       ("Clean the raw sales export and report sales by store", "Data Analyst")]:
        m.ex_text(f"Which role? {task}. (Type Data Analyst, Data Scientist or Business Analyst.)", role)

    # ---------------------------------------------------------------- 1.1.10
    fm = CAT.set_index("Category").Margin
    worst = ST.Store_Name[ST.Profit.idxmin()]
    m.topic("1.1.10", "Analytical thinking and problem-solving",
            "Analytical thinking means breaking a big question into smaller ones that data can answer, answering each with evidence, "
            "and joining the answers into an explanation.")
    m.steps(["Write the big question: why is the FY2025 chain margin only 12.7%?",
             "Split it into parts: which stores? which categories? (which discounts is answered in Module 1.2).",
             "Answer each part with one formula (below).",
             "Join the parts: a few loss-making stores and a very thin Furniture margin pull the chain margin down."])
    m.worked([("Chain margin FY2025", "=SUM(tStores25[Profit])/SUM(tStores25[Sales])", ST.Profit.sum() / ST.Sales.sum(), "pct", "The big question."),
              ("Part 1: stores that lost money", '=COUNTIF(tStores25[Profit],"<0")', int((ST.Profit < 0).sum()), "int", ""),
              ("Part 1: the biggest single loss", "=MIN(tStores25[Profit])", ST.Profit.min(), N, worst),
              ("Part 2: Furniture margin", '=SUMIFS(tCat25[Profit],tCat25[Category],"Furniture")/SUMIFS(tCat25[Sales],tCat25[Category],"Furniture")',
               fm["Furniture"], "pct", "Almost nothing kept from Furniture sales."),
              ("Part 2: Technology margin", '=SUMIFS(tCat25[Profit],tCat25[Category],"Technology")/SUMIFS(tCat25[Sales],tCat25[Category],"Technology")',
               fm["Technology"], "pct", "For comparison.")])
    m.turn()
    m.ex("What was the Office Supplies margin in FY2025?", '=SUMIFS(tCat25[Profit],tCat25[Category],"Office Supplies")/SUMIFS(tCat25[Sales],tCat25[Category],"Office Supplies")',
         fm["Office Supplies"], "pct", short="Office Supplies margin")
    low = CAT.Category[CAT.Margin.idxmin()]
    m.ex_text("Which category had the lowest FY2025 margin?", low,
              formula="=INDEX(tCat25[Category],MATCH(MIN(tCat25[Margin]),tCat25[Margin],0))", short="Lowest-margin category")

    # ---------------------------------------------------------------- 1.1.11-12
    m.topic("1.1.11-12", "Asking the right questions, and turning business problems into analytical questions",
            "A sharp question is specific, measurable and time-bound. Managers describe problems ('Furniture isn't making money'); "
            "the analyst rewrites each one as a question a formula can answer.")
    m.steps(["Write the problem exactly as it was said.",
             "Ask three things: which measure? which group? which period?",
             "Rewrite it as one question a formula can answer.",
             "List the columns you need, then write the formula."])
    m.s.grid([["How are sales?", "Which stores missed their FY2025 sales target, and by how much?", "Store_Name, Sales, Target"],
              ["Furniture isn't making money.", "By how many percentage points did the Furniture margin change from 2024 to 2025?", "Category, Margin_2024, Margin"],
              ["Are students doing well?", "What share of students in each faculty is at risk this semester?", "Faculty, At_Risk"]],
             header=["Vague / business problem", "Sharp analytical question", "Columns needed"])
    m.s.r += 1
    fch = CAT.Margin - CAT.Margin_2024
    m.worked([("Furniture margin change, 2024 › 2025",
               '=SUMIFS(tCat25[Margin],tCat25[Category],"Furniture")-SUMIFS(tCat25[Margin_2024],tCat25[Category],"Furniture")',
               float(fch[CAT.Category == "Furniture"].iloc[0]), "pct", "Negative: about 2 percentage points lower.")])
    m.turn()
    m.ex("By how many percentage points did the Office Supplies margin change from 2024 to 2025? (Type it as a %, e.g. -1.5%.)",
         '=SUMIFS(tCat25[Margin],tCat25[Category],"Office Supplies")-SUMIFS(tCat25[Margin_2024],tCat25[Category],"Office Supplies")',
         float(fch[CAT.Category == "Office Supplies"].iloc[0]), "pct", short="Office Supplies margin change")
    m.ex_open("Rewrite the vague question 'Is the hotel doing well?' as a sharp analytical question.",
              "What was Atlantic Pearl' cancellation rate by market segment in 2023-2024, and which segment was highest?")

    # ---------------------------------------------------------------- 1.1.13
    bn, en = 4, 9   # row positions of Benin City and Enugu in tStores25
    m.topic("1.1.13", "General knowledge and domain knowledge",
            "General knowledge is what most people know. Domain knowledge is how a particular industry works: which measures matter, "
            "what 'normal' looks like and why a number may be misleading.")
    m.steps(["Read the number.",
             "Say what general knowledge suggests.",
             "Add what someone who works in that industry would know.",
             "Only then draw a conclusion."])
    m.worked([("Benin City attainment (row 6 of Data_Stores25)", "=Data_Stores25!J6", ST.Attainment[bn], "pct", "General knowledge: beat target, good store."),
              ("Benin City profit", "=Data_Stores25!F6", ST.Profit[bn], N,
               "Domain knowledge: a store can hit a sales target by discounting heavily, and still lose money.")])
    m.turn()
    m.ex("What was the margin of Enugu - Independence Layout (row 11 of Data_Stores25)?", "=Data_Stores25!I11", ST.Margin[en], "pct",
         short="Enugu margin")
    m.ex_open("Mercyland's bed occupancy was 95% or higher on 63 days in 2023. What would a hospital manager's domain knowledge add?",
              "Above roughly 85-90% occupancy a hospital struggles to admit emergencies quickly, so the manager would plan extra staffed beds and staff for the peak months.")

    # ---------------------------------------------------------------- 1.1.14
    big = ST.Store_Name[ST.Sales.idxmax()]
    best = ST.Store_Name[ST.Attainment.idxmax()]
    m.topic("1.1.14", "Making decisions based on data rather than assumptions",
            "An assumption is a belief nobody has checked. Test it with one formula on the data, then keep it, change it or drop it.")
    m.steps(["Write the assumption: 'our biggest store is our best performer'.",
             "Find the column that tests it: Sales gives size; Attainment gives performance against target.",
             "Write one formula for each.",
             "Decide: here the evidence changes the assumption."])
    m.worked([("Store with the highest sales", "=INDEX(tStores25[Store_Name],MATCH(MAX(tStores25[Sales]),tStores25[Sales],0))", big, None,
               "INDEX/MATCH is covered in Week 2; for now, sort Data_Stores25 by Sales."),
              ("Store with the highest attainment", "=INDEX(tStores25[Store_Name],MATCH(MAX(tStores25[Attainment]),tStores25[Attainment],0))", best, None,
               "The best performer against its own target."),
              ("Stores that beat target yet lost money", '=COUNTIFS(tStores25[Attainment],">=1",tStores25[Profit],"<0")',
               int(((ST.Attainment >= 1) & (ST.Profit < 0)).sum()), "int", "Tests 'beating target means profitable'.")])
    m.turn()
    m.ex("How many stores missed their target but still made a profit?", '=COUNTIFS(tStores25[Attainment],"<1",tStores25[Profit],">0")',
         int(((ST.Attainment < 1) & (ST.Profit > 0)).sum()), "int", short="Missed target, still profitable")
    hm = ST.Store_Name[ST.Margin.idxmax()]
    m.ex_text("Which store had the highest FY2025 margin? (Sort Data_Stores25 by Margin, largest first.)", hm,
              formula="=INDEX(tStores25[Store_Name],MATCH(MAX(tStores25[Margin]),tStores25[Margin],0))", short="Highest-margin store",
              alts=(hm.split(" - ")[0],))
    m.finish()


# =====================================================================================================
def m12(b, D):
    Y, Q, DB, FU = D.Y, D.Q25, D.DB, D.FURN
    m = ModuleSheet(b, "1.2 Types of Analytics", "1.2  Types of Data Analytics",
                    "Tutor walk-through after Module 1.2: descriptive, diagnostic, predictive and prescriptive analytics on SuperMart data",
                    "1.2.1 - 1.2.6")
    m.dataset([("tYears", "On this sheet, below", "SuperMart sales, profit and orders for each year 2022-2025. Margin is a calculated column."),
               ("tDiscBand", "On this sheet, below", "FY2025 sales lines grouped by discount band: none, 1-20%, above 20%."),
               ("tQuarter25", "On this sheet, below", "FY2025 sales by quarter; Share is a calculated column."),
               ("tFurn25", "On this sheet, below", "FY2025 Furniture lines grouped by discount level, used to cost a discount cap."),
               ("tStudents", "Sheet Data_Students (starts in A1)", "Kingsway students; used in 1.2.6.")])
    m.h("tYears")
    ty = m.s.table(Y[["Year", "Sales", "Profit", "Orders"]], "tYears", fmts={"Year": "gen", "Sales": N, "Profit": N, "Orders": "int"},
              calc={"Margin": ("=[@Profit]/[@Sales]", "pct", Y.Margin.tolist())})
    m.h("tDiscBand")
    m.s.table(DB[["Discount_Band", "Lines", "Sales", "Profit"]].astype({"Discount_Band": str}), "tDiscBand",
              fmts={"Lines": "int", "Sales": N, "Profit": N}, calc={"Margin": ("=[@Profit]/[@Sales]", "pct", DB.Margin.tolist())},
              widths={"Discount_Band": 15})
    m.h("tQuarter25")
    m.s.table(Q, "tQuarter25", fmts={"Sales": N}, calc={"Share": ("=[@Sales]/SUM([Sales])", "pct", (Q.Sales / Q.Sales.sum()).tolist())})

    # ---------------------------------------------------------------- 1.2.1
    s24, s25 = Y.set_index("Year").Sales[2024], Y.set_index("Year").Sales[2025]
    m.topic("1.2.1", "Descriptive analytics: what happened?",
            "Descriptive analytics summarises past data with totals, averages, counts, rates and charts.")
    m.steps(["Total: =SUMIFS(tYears[Sales],tYears[Year],2025) gives FY2025 sales.",
             "Change: divide this year by last year and subtract 1 for growth.",
             "Rate: the Margin column already divides Profit by Sales for each year.",
             "Picture: select tYears[Year] and tYears[Sales], Insert › Line chart (below)."])
    m.worked([("FY2025 sales", "=SUMIFS(tYears[Sales],tYears[Year],2025)", s25, N, ""),
              ("Growth 2025 vs 2024", "=SUMIFS(tYears[Sales],tYears[Year],2025)/SUMIFS(tYears[Sales],tYears[Year],2024)-1", s25 / s24 - 1, "pct", ""),
              ("Sales, all four years", "=SUM(tYears[Sales])", Y.Sales.sum(), N, ""),
              ("Best margin of the four years", "=MAX(tYears[Margin])", Y.Margin.max(), "pct", "2024")])
    m.s.chart("line", [{"name": "Sales", "categories": ty.ref("Year"), "values": ty.ref("Sales")}],
              "SuperMart sales by year", legend=False, rows=15, size=(560, 270), y_fmt='₦#,##0,,"m"')
    m.turn()
    m.ex("What was the margin in 2023?", "=SUMIFS(tYears[Margin],tYears[Year],2023)", float(Y.set_index("Year").Margin[2023]), "pct", short="2023 margin")
    m.ex("How many orders were placed in the four years together?", "=SUM(tYears[Orders])", int(Y.Orders.sum()), "int", short="Orders 2022-2025")

    # ---------------------------------------------------------------- 1.2.2
    dbi = DB.set_index(DB.Discount_Band.astype(str))
    m.topic("1.2.2", "Diagnostic analytics: why did it happen?",
            "Diagnostic analytics drills into the data to find the cause of an outcome, by splitting it into segments and comparing them.")
    m.steps(["Start from the outcome: FY2025 margin was 12.7%.",
             "Split the lines into segments: tDiscBand groups them by discount.",
             "Compare the segments: Margin shows each band's profit per ₦ of sales.",
             "Name the cause carefully: deep discounts coincide with losses. Check other explanations before calling it proof."])
    m.worked([("Lines discounted above 20%", '=SUMIFS(tDiscBand[Lines],tDiscBand[Discount_Band],"Above 20%")', int(dbi.Lines["Above 20%"]), "int", ""),
              ("Their share of all FY2025 lines", '=SUMIFS(tDiscBand[Lines],tDiscBand[Discount_Band],"Above 20%")/SUM(tDiscBand[Lines])',
               dbi.Lines["Above 20%"] / DB.Lines.sum(), "pct", ""),
              ("Profit on those lines", '=SUMIFS(tDiscBand[Profit],tDiscBand[Discount_Band],"Above 20%")', dbi.Profit["Above 20%"], N, "A loss."),
              ("Profit on all other lines", '=SUM(tDiscBand[Profit])-SUMIFS(tDiscBand[Profit],tDiscBand[Discount_Band],"Above 20%")',
               DB.Profit.sum() - dbi.Profit["Above 20%"], N, "The deep-discount lines gave away about a third of this.")])
    m.turn()
    m.ex("What was the margin on lines discounted 1-20%?", '=SUMIFS(tDiscBand[Margin],tDiscBand[Discount_Band],"1-20%")', float(dbi.Margin["1-20%"]), "pct",
         short="Margin, 1-20% band")
    m.ex_text("Which discount band lost money?", "Above 20%", formula="=INDEX(tDiscBand[Discount_Band],MATCH(MIN(tDiscBand[Profit]),tDiscBand[Profit],0))",
              short="Loss-making band", alts=("above 20", ">20%"))

    # ---------------------------------------------------------------- 1.2.3
    g = round(float(s25 / s24 - 1), 3)
    q = Q.set_index("Quarter").Sales
    m.topic("1.2.3", "Predictive analytics: what might happen?",
            "Predictive analytics uses past patterns to estimate the future. Every prediction rests on an assumption, so state it and treat the result as an estimate.")
    m.steps(["Find a pattern: Share in tQuarter25 shows Q4 brought the biggest part of FY2025 sales.",
             "State the assumption in a yellow input cell: FY2026 grows at the same rate as FY2025.",
             "Estimate FY2026 sales: FY2025 sales × (1 + growth).",
             "Split the estimate by last year's quarter shares to estimate each quarter."])
    inp = m.s.inputs([("Assumed FY2026 growth (change me)", g, "pct", "Same as 2025 vs 2024, rounded. Try 10%.")])
    G = inp["Assumed FY2026 growth (change me)"]
    fy26 = q.sum() * (1 + g)
    m.worked([("Q4 share of FY2025 sales", '=SUMIFS(tQuarter25[Share],tQuarter25[Quarter],"Q4")', q["Q4"] / q.sum(), "pct", "The festive-season pattern."),
              ("Estimated FY2026 sales", f"=SUM(tQuarter25[Sales])*(1+{G})", fy26, N, "An estimate, not a promise."),
              ("Estimated Q4 2026 sales", f'=SUM(tQuarter25[Sales])*(1+{G})*SUMIFS(tQuarter25[Share],tQuarter25[Quarter],"Q4")',
               fy26 * q["Q4"] / q.sum(), N, "Same seasonal pattern assumed.")])
    m.turn()
    m.ex("What share of FY2025 sales came in Q3?", '=SUMIFS(tQuarter25[Share],tQuarter25[Quarter],"Q3")', q["Q3"] / q.sum(), "pct", short="Q3 share")
    m.ex("With growth left at its starting value, what is the estimated Q1 2026 sales figure?",
         f'=SUM(tQuarter25[Sales])*(1+{g})*SUMIFS(tQuarter25[Share],tQuarter25[Quarter],"Q1")', fy26 * q["Q1"] / q.sum(), N, tol=1,
         short="Q1 2026 estimate (growth as given)")

    # ---------------------------------------------------------------- 1.2.4
    def gain(cap, lost):
        hit = FU.Discount > cap
        newp = (FU.Profit + FU.Sales / (1 - FU.Discount) * (1 - cap) - FU.Sales) * (1 - lost)
        return float((newp - FU.Profit)[hit].sum())
    m.topic("1.2.4", "Prescriptive analytics: what should we do?",
            "Prescriptive analytics compares options and recommends one, showing the expected result of each and the assumptions behind it.")
    m.steps(["Write the option: cap Furniture discounts at 20%. A line sold above the cap is re-priced at the cap (same list price and cost).",
             "Put every assumption in a yellow cell: the cap, and the share of re-priced volume customers no longer buy.",
             "The Gain column works out, line group by line group, the extra profit if the cap applied: IF the discount is above the cap, "
             "new profit minus old profit, otherwise 0. Note the $ signs that lock the two input cells.",
             "Add up the Gain column and compare options by changing the inputs."])
    inp = m.s.inputs([("Discount cap (change me)", 0.2, "pct0", "Try 10% and 30%."),
                      ("Volume lost on re-priced lines (change me)", 0.0, "pct0", "Try 25% and 50%.")])
    C, LS = inp["Discount cap (change me)"], inp["Volume lost on re-priced lines (change me)"]
    hit = FU.Discount > 0.2
    m.h("tFurn25  (FY2025 Furniture lines by discount level)")
    gcol = ((FU.Profit + FU.Sales / (1 - FU.Discount) * 0.8 - FU.Sales) - FU.Profit).where(hit, 0)
    m.s.table(FU, "tFurn25", fmts={"Discount": "pct0", "Lines": "int", "Sales": N, "Profit": N},
              calc={"Gain": (f"=IF([@Discount]>{C},([@Profit]+[@Sales]/(1-[@Discount])*(1-{C})-[@Sales])*(1-{LS})-[@Profit],0)", N, gcol.tolist())})
    m.worked([("Furniture lines above the cap", f'=SUMIFS(tFurn25[Lines],tFurn25[Discount],">"&{C})', int(FU.Lines[hit].sum()), "int", ""),
              ("Their profit today", f'=SUMIFS(tFurn25[Profit],tFurn25[Discount],">"&{C})', FU.Profit[hit].sum(), N, "A loss today."),
              ("Profit gain from the cap", "=SUM(tFurn25[Gain])", gain(0.2, 0.0), N, "Recommend the cap, and state the volume assumption.")])
    m.turn()
    m.ex("Set the cap to 10% (volume lost 0%). What is the profit gain? Type the number you read in 'Profit gain from the cap'.",
         "=SUMPRODUCT((tFurn25[Discount]>0.1)*((tFurn25[Profit]+tFurn25[Sales]/(1-tFurn25[Discount])*(1-0.1)-tFurn25[Sales])-tFurn25[Profit]))",
         gain(0.1, 0.0), N, tol=1, short="Gain, 10% cap, 0% volume lost")
    m.ex("Set the cap back to 20% and volume lost to 50%. What is the profit gain now?",
         "=SUMPRODUCT((tFurn25[Discount]>0.2)*((tFurn25[Profit]+tFurn25[Sales]/(1-tFurn25[Discount])*(1-0.2)-tFurn25[Sales])*(1-0.5)-tFurn25[Profit]))",
         gain(0.2, 0.5), N, tol=1, short="Gain, 20% cap, 50% volume lost")
    m.note("The answer key uses SUMPRODUCT so its answers do not move when you change the inputs. SUMPRODUCT is covered in Week 2.")

    # ---------------------------------------------------------------- 1.2.5-6
    stu = D.STU
    m.topic("1.2.5-6", "Practical examples of each type, applied to retail, university and hospital data",
            "The four types are four questions: what happened, why, what might happen, what should we do. Each builds on the one before.")
    m.steps(["Read the question.",
             "Ask: is it about the past (descriptive), a cause (diagnostic), the future (predictive) or an action (prescriptive)?",
             "Answer it with the matching kind of work: a summary, a comparison, an estimate with assumptions, or options with results."],
            label="How to classify a question")
    m.worked([("Retail, descriptive: FY2025 margin", "=SUMIFS(tYears[Margin],tYears[Year],2025)", float(Y.set_index("Year").Margin[2025]), "pct", ""),
              ("University, descriptive: at-risk rate", '=COUNTIF(tStudents[At_Risk],"Yes")/COUNTA(tStudents[Matric_No])',
               (stu.At_Risk == "Yes").mean(), "pct", "Data_Students"),
              ("University, diagnostic: at-risk rate, attendance 70% or below",
               '=COUNTIFS(tStudents[At_Risk],"Yes",tStudents[Lecture_Attendance_Pct],"<=70")/COUNTIF(tStudents[Lecture_Attendance_Pct],"<=70")',
               ((stu.At_Risk == "Yes") & (stu.Lecture_Attendance_Pct <= 70)).sum() / (stu.Lecture_Attendance_Pct <= 70).sum(), "pct",
               "Far higher than average: low attendance goes with risk.")])
    m.turn()
    for qq, ans in [("How many students are at risk this semester?", "Descriptive"),
                    ("Why is Surgery's readmission rate the highest?", "Diagnostic"),
                    ("How many beds will Mercyland need next June?", "Predictive"),
                    ("Should SuperMart cap Furniture discounts at 20%?", "Prescriptive")]:
        m.ex_text(f"Which type? '{qq}'", ans)
    m.ex("How many Kingsway students are at risk?", '=COUNTIF(tStudents[At_Risk],"Yes")', int((stu.At_Risk == "Yes").sum()), "int",
         short="Students at risk")
    m.finish()


# =====================================================================================================
def m13(b, D):
    ST, DI = D.ST, D.DIRTY
    s25 = D.s25
    m = ModuleSheet(b, "1.3 Analytics Process", "1.3  The Data Analytics Process",
                    "Tutor walk-through after Module 1.3: from objective to recommendation, with data quality checks throughout",
                    "1.3.1 - 1.3.12")
    samp = s25[s25.Store_ID == "ST02"].sort_values(["Order_Date", "Row_ID"])
    samp = pd.concat([samp[samp.Discount == 0].head(3), samp[samp.Discount.between(0.01, 0.2)].head(3), samp[samp.Discount > 0.2].head(2)])
    samp = samp.sort_values("Order_Date")[["Order_ID", "Order_Date", "Sales", "Discount", "Profit"]].reset_index(drop=True)
    m.dataset([("tDirtySales", "Sheet Data_DirtySales (starts in A1)", "The raw FY2025 SuperMart export exactly as received: 3,387 rows with "
                                                                          "duplicates, blank rows, text numbers, mixed dates and inconsistent names."),
               ("tStores25", "Sheet Data_Stores25 (starts in A1)", "FY2025 store results, used for analysis and interpretation."),
               ("tSample25", "On this sheet, below", "Eight FY2025 Lagos - Lekki lines, used to practise transformation."),
               ("Stores directory CSV", "00_Core_Datasets/SuperMart_Stores_Directory.csv", "Used in 1.3.3 to practise importing."),
               ("Dirty sales CSV", "Week2_Data_Files/SuperMart_Sales_2025_DIRTY.csv", "Used in 1.3.4 to practise connecting.")])

    # ---------------------------------------------------------------- 1.3.1
    miss = ST[ST.Attainment < 1]
    m.topic("1.3.1", "Understanding the problem and defining objectives",
            "An objective says what the analysis must find, for whom and by when. A good one is SMART: Specific, Measurable, Achievable, Relevant, Time-bound.")
    m.steps(["Fuzzy request: 'look at the store data'.",
             "Make it SMART: 'By Friday, identify which of the 12 stores missed their FY2025 sales target and by how much, for the Head of Sales.'",
             "The objective tells you the measure (Sales vs Target), the group (stores), the period (FY2025) and the audience.",
             "Answer exactly that, no more: count the stores below target, then the total shortfall."])
    m.worked([("Stores below target", '=COUNTIF(tStores25[Attainment],"<1")', len(miss), "int", ""),
              ("Total shortfall of those stores", '=SUMIF(tStores25[Attainment],"<1",tStores25[Target])-SUMIF(tStores25[Attainment],"<1",tStores25[Sales])',
               (miss.Target - miss.Sales).sum(), N, "Target minus sales, stores below target only.")])
    m.turn()
    m.ex("By how much did Owerri - New Owerri (row 13 of Data_Stores25) miss its target? (Target minus Sales.)", "=Data_Stores25!H13-Data_Stores25!E13",
         int(ST.Target[11] - ST.Sales[11]), N, short="Owerri shortfall")
    m.ex_open("Make this SMART: 'check how the hospital is doing'.",
              "By the end of the month, report Mercyland's 2023 30-day readmission rate by department, highlighting departments above the hospital average, for the Medical Director.")

    # ---------------------------------------------------------------- 1.3.2
    m.topic("1.3.2", "Data collection and sources",
            "Data collection is gathering the data the objective needs. Primary data is collected by you for this purpose; secondary data was collected by someone else. "
            "Internal data comes from your own systems; external data from outside.")
    m.steps(["List the data the objective needs (for 1.3.1: sales and targets by store).",
             "For each, ask: does it already exist? Most analysis uses internal data that systems record automatically.",
             "Note who owns it and how often it is updated."])
    m.s.grid([["Till / point-of-sale system", "tSales, tDirtySales", "Internal, secondary"],
              ["Head office targets", "Target column of tStores25", "Internal, secondary"],
              ["Survey you run yourself", "ETEH Talent Salary Survey", "Primary"],
              ["National statistics office", "Inflation, population figures", "External, secondary"]],
             header=["Source", "Example in this bootcamp", "Type"])
    m.s.r += 1
    m.turn()
    for src, ans in [("Hotel booking engine records", "Internal"), ("Population figures from the statistics office", "External"),
                     ("A questionnaire you send to customers", "Primary")]:
        m.ex_text(f"Internal, External or Primary? {src}", ans)

    # ---------------------------------------------------------------- 1.3.3
    m.topic("1.3.3", "Manual data entry and data import",
            "Type small amounts of data by hand; import anything larger or anything that already exists in a system. Importing is faster and avoids typing errors.")
    m.steps(["Manual entry: type the three targets below. Each yellow cell only accepts a whole number above 0 (data validation).",
             "Import: Data › Get Data › From File › From Text/CSV.",
             "Choose 00_Core_Datasets › SuperMart_Stores_Directory.csv; check the preview and that Excel guessed each column type correctly.",
             "Click Load: the 12 stores arrive as a new table on a new sheet."])
    r0 = m.s.r
    m.s.grid([["ST01", "Lagos - Ikeja", None], ["ST02", "Lagos - Lekki", None], ["ST03", "Ibadan - Ring Road", None]],
             header=["Store_ID", "Store_Name", "Target (type it)"], fmts=[None, None, N])
    rng = f"D{r0 + 2}:D{r0 + 4}"
    for rr in range(r0 + 1, r0 + 4):
        m.ws.write_blank(rr, 3, None, b.nf(N, style="input"))
    m.ws.data_validation(r0 + 1, 3, r0 + 3, 3, {"validate": "integer", "criteria": ">", "value": 0,
                                                 "error_title": "Whole number only", "error_message": "Type the target as a whole number of naira."})
    m.s.r += 1
    m.turn()
    m.ex_formula_live("Type the FY2025 targets of the three stores (from Data_Stores25, column H) into the yellow cells above.",
                      f'=IF(COUNT({rng})<3,"Type all three targets",IF(SUM({rng})={int(ST.Target[:3].sum())},"✓ Correct: matches Data_Stores25","✗ Check against Data_Stores25"))',
                      "Type all three targets", key="₦142,000,000; ₦63,000,000; ₦37,000,000 (total ₦242,000,000).")
    m.do([("Import SuperMart_Stores_Directory.csv with Data › Get Data › From File › From Text/CSV, then Load.",
           "A new sheet with a 12-row table (ST01 to ST12) and 9 columns."),
          ("Check the Opened column: is it a date, aligned right?", "Yes: Excel recognised the yyyy-mm-dd dates.")])

    # ---------------------------------------------------------------- 1.3.4
    m.topic("1.3.4", "Connecting to files, folders and databases",
            "A connection remembers where the data lives. When the source file changes, one click on Refresh brings in the latest version. Power Query, which builds connections, is covered fully in Week 2.")
    m.steps(["Data › Get Data › From File › From Text/CSV.",
             "Choose Week2_Data_Files › SuperMart_Sales_2025_DIRTY.csv.",
             "In the preview, click the arrow next to Load › Load To… › Only Create Connection › OK.",
             "The connection appears in the Queries & Connections pane with its row count. Data › Refresh All updates it."])
    m.turn()
    m.ex("How many rows does the Queries & Connections pane report for the dirty CSV?", "=ROWS(tDirtySales)", len(DI), "int",
         short="Rows in the dirty CSV (same as tDirtySales)")
    m.do([("Right-click the connection › Properties: find the 'Refresh data when opening the file' option.", "A tick box on the Usage tab."),
          ("Name two other places Get Data can connect to.", "From Folder, From Database (or ODBC), From Web, From Workbook.")])

    # ---------------------------------------------------------------- 1.3.5
    txt_sales = int(DI.Sales.map(lambda v: isinstance(v, str)).sum())
    txt_disc = int(DI.Discount.map(lambda v: isinstance(v, str)).sum())
    blanks = int(DI.Order_ID.isna().sum())
    spaces = int(DI.Zone.dropna().map(lambda v: v != v.strip()).sum())
    exact_ik = int((DI.Store_Name == "Lagos - Ikeja").sum())
    any_ik = int(DI.Store_Name.dropna().str.lower().eq("lagos - ikeja").sum())
    m.topic("1.3.5", "Data cleaning and preparation",
            "Data cleaning is finding and correcting (or removing) errors and inconsistencies so the data is accurate, complete and consistent. "
            "First profile the problems with formulas, then fix them (fixing is practised in 1.6, 1.8 and 1.9).")
    m.steps(["Count the rows: =ROWS(tDirtySales).",
             "Find empty rows: =COUNTBLANK(tDirtySales[Order_ID]) counts rows with no order number.",
             "Find numbers stored as text: COUNTA counts every filled cell, COUNT only numbers, so the difference is text.",
             "Find extra spaces: compare each value's length with its TRIM-med length.",
             "Find spelling variants: COUNTIF ignores case, EXACT does not."])
    m.worked([("Rows in the raw export", "=ROWS(tDirtySales)", len(DI), "int", ""),
              ("Empty rows (no Order_ID)", "=COUNTBLANK(tDirtySales[Order_ID])", blanks, "int", "Delete them."),
              ("Sales stored as text", "=COUNTA(tDirtySales[Sales])-COUNT(tDirtySales[Sales])", txt_sales, "int", "e.g. '₦1,399,930': SUM skips these."),
              ("'Lagos - Ikeja' in any capitals", '=COUNTIF(tDirtySales[Store_Name],"Lagos - Ikeja")', any_ik, "int", "COUNTIF ignores case."),
              ("'Lagos - Ikeja' spelled exactly", '=SUMPRODUCT(--EXACT(tDirtySales[Store_Name],"Lagos - Ikeja"))', exact_ik, "int",
               "The rest are 'lagos - ikeja' and 'LAGOS - IKEJA'.")])
    m.turn()
    m.ex("How many Discount values are stored as text (for example '20%')?", "=COUNTA(tDirtySales[Discount])-COUNT(tDirtySales[Discount])", txt_disc, "int",
         short="Discounts stored as text")
    m.ex("How many Zone values have extra spaces before or after the name? (Hint: =SUMPRODUCT(--(LEN(range)<>LEN(TRIM(range)))).)",
         "=SUMPRODUCT(--(LEN(tDirtySales[Zone])<>LEN(TRIM(tDirtySales[Zone]))))", spaces, "int", short="Zones with extra spaces")

    # ---------------------------------------------------------------- 1.3.6
    def band(d):
        return "None" if d == 0 else ("1-20%" if d <= 0.2 else "Above 20%")
    m.topic("1.3.6", "Data transformation and organisation",
            "Transformation changes the structure or format of clean data so it is ready for analysis: new columns, groups, reshaping and combining tables. "
            "Cleaning fixes what is wrong; transformation adds what the question needs.")
    m.steps(["Start with tSample25 (below): five columns from the sales system.",
             "Year and Month: =YEAR([@Order_Date]) and =MONTH([@Order_Date]).",
             "Quarter: =\"Q\"&ROUNDUP(MONTH([@Order_Date])/3,0) joins the letter Q to the quarter number.",
             "Margin: =[@Profit]/[@Sales]. Discount_Band: a nested IF puts each discount in a band.",
             "Typed once in the first row, each formula fills the whole column because tSample25 is an Excel table."])
    m.h("tSample25")
    m.s.table(samp, "tSample25", fmts={"Order_Date": "date", "Sales": N, "Discount": "pct0", "Profit": N},
              calc={"Year": ("=YEAR([@Order_Date])", "gen", samp.Order_Date.dt.year.tolist()),
                    "Quarter": ('="Q"&ROUNDUP(MONTH([@Order_Date])/3,0)', None, [f"Q{(d.month - 1) // 3 + 1}" for d in samp.Order_Date]),
                    "Margin": ("=[@Profit]/[@Sales]", "pct", (samp.Profit / samp.Sales).tolist()),
                    "Discount_Band": ('=IF([@Discount]=0,"None",IF([@Discount]<=0.2,"1-20%","Above 20%"))', None, [band(x) for x in samp.Discount])},
              widths={"Order_ID": 16, "Discount_Band": 14})
    m.turn()
    m.ex("How many tSample25 lines are in the 'Above 20%' band?", '=COUNTIF(tSample25[Discount_Band],"Above 20%")',
         int((samp.Discount > 0.2).sum()), "int", short="Above-20% lines in tSample25")
    m.ex_text("What does =\"Q\"&ROUNDUP(MONTH(DATE(2025,8,14))/3,0) return?", "Q3")

    # ---------------------------------------------------------------- 1.3.7
    m.topic("1.3.7", "Data analysis",
            "Analysis applies calculations to answer the question: aggregate, compare, rank, segment and look at change over time.")
    m.steps(["Aggregate: total and average sales.",
             "Compare: chain attainment = total sales ÷ total target.",
             "Rank: =RANK.EQ(value, column) gives a store's position (1 = highest).",
             "Segment: SUMIF by Zone."])
    sw = ST[ST.Zone == "South-West"].Sales.sum()
    m.worked([("Total sales", "=SUM(tStores25[Sales])", ST.Sales.sum(), N, ""),
              ("Average sales per store", "=AVERAGE(tStores25[Sales])", ST.Sales.mean(), N, ""),
              ("Chain attainment", "=SUM(tStores25[Sales])/SUM(tStores25[Target])", ST.Sales.sum() / ST.Target.sum(), "pct", "Not the average of 12 percentages."),
              ("Lagos - Lekki's rank on attainment", "=RANK.EQ(Data_Stores25!J3,tStores25[Attainment])", int(ST.Attainment.rank(ascending=False)[1]), "int", "1 = best."),
              ("South-West sales", '=SUMIF(tStores25[Zone],"South-West",tStores25[Sales])', sw, N, "")])
    m.turn()
    m.ex("What were total South-East sales?", '=SUMIF(tStores25[Zone],"South-East",tStores25[Sales])', ST[ST.Zone == "South-East"].Sales.sum(), N,
         short="South-East sales")
    m.ex("What is Lagos - Ikeja's rank on Sales (row 2)?", "=RANK.EQ(Data_Stores25!E2,tStores25[Sales])", 1, "int", short="Ikeja sales rank")

    # ---------------------------------------------------------------- 1.3.8
    lk = ST.iloc[1]
    m.topic("1.3.8", "Interpretation of results",
            "Interpretation explains what a result means in the real world: is it good or bad, compared with what, and why does it matter?")
    m.steps(["Compare the number with something: a target, last year, another group.",
             "Say what it means for the business, in a sentence.",
             "Note any limits (small groups, missing data)."])
    m.worked([("Lagos - Lekki attainment", "=Data_Stores25!J3", lk.Attainment, "pct", ""),
              ("Naira above target", "=Data_Stores25!E3-Data_Stores25!H3", int(lk.Sales - lk.Target), N,
               "Interpretation: Lekki beat its target by 12.4%, about ₦7.8m, the best attainment in the chain.")])
    m.turn()
    m.ex("By how many naira did Jos - Rayfield (row 10) beat its target?", "=Data_Stores25!E10-Data_Stores25!H10", int(ST.Sales[8] - ST.Target[8]), N,
         short="Jos above target")
    m.ex_open("Write a one-sentence interpretation of Owerri - New Owerri's FY2025 attainment (row 13).",
              f"Owerri reached only {ST.Attainment[11]:.1%} of its target, the lowest in the chain, about ₦{(ST.Target[11] - ST.Sales[11]) / 1e6:.1f}m short, "
              f"although its margin ({ST.Margin[11]:.1%}) was the highest.")

    # ---------------------------------------------------------------- 1.3.9
    m.topic("1.3.9", "Data visualisation and presentation",
            "Visualisation shows data as charts so patterns are seen quickly. Pick the chart that matches the question.")
    m.steps(["Go to Data_Stores25 and select Store_Name, then hold Ctrl and select Attainment.",
             "Insert › Recommended Charts › Clustered Bar.",
             "Give it a title that states the message, e.g. 'Five stores met their FY2025 target'.",
             "Sort the table by Attainment first so the bars are in order. The chart below was built the same way."])
    m.s.chart("bar", [{"name": "Attainment", "categories": ["Data_Stores25", 1, 1, 12, 1], "values": ["Data_Stores25", 1, 9, 12, 9],
                       "labels": True, "label_fmt": "0%"}], "Five stores met their FY2025 target", y_fmt="0%", legend=False, rows=19,
              size=(640, 360))
    m.turn()
    for qq, ans in [("Sales of each of the 12 stores", "Bar"), ("Monthly occupancy across 2023", "Line"),
                    ("Attendance against exam score for 5,840 students", "Scatter")]:
        m.ex_text(f"Which chart? {qq}. (Bar, Line, Pie, Scatter or Histogram.)", ans, alts=("column",) if ans == "Bar" else ())

    # ---------------------------------------------------------------- 1.3.10
    m.topic("1.3.10", "Communicating findings and recommendations",
            "Lead with the answer, back it with two or three numbers, then say what to do next. Round numbers for readers.")
    m.steps(["Headline: the answer in one sentence.",
             "Evidence: two or three numbers, rounded.",
             "Recommendation: the action.",
             "Next step: how to test or start it."], label="A finding in four parts")
    m.worked([("Headline number: loss-making stores", '=COUNTIF(tStores25[Profit],"<0")', int((ST.Profit < 0).sum()), "int", "Five of our 12 stores lost money in FY2025."),
              ("Evidence, rounded to ₦m", '=ROUND(SUMIF(tStores25[Profit],"<0")/1000000,1)', round(ST.Profit[ST.Profit < 0].sum() / 1e6, 1), "dec1",
               "Together they lost about this many million naira."),
              ("Evidence: chain margin", "=SUM(tStores25[Profit])/SUM(tStores25[Sales])", ST.Profit.sum() / ST.Sales.sum(), "pct", "")])
    m.turn()
    m.ex("Round Lagos - Ikeja's FY2025 sales (row 2) to the nearest ₦0.1m and type it in millions, e.g. 146.4.", "=ROUND(Data_Stores25!E2/1000000,1)",
         round(ST.Sales[0] / 1e6, 1), "dec1", tol=0.001, short="Ikeja sales in ₦m, 1 dp")
    m.ex_open("Write a four-part finding (headline, evidence, recommendation, next step) about the stores below target.",
              "Headline: 7 of 12 stores missed their FY2025 target. Evidence: together they were about ₦34m short; Owerri reached only 81%. "
              "Recommendation: review targets and local promotions in those stores. Next step: meet the seven store managers this month.")

    # ---------------------------------------------------------------- 1.3.11
    tot25 = s25.Sales.sum()
    m.topic("1.3.11", "Accuracy and data quality",
            "Data quality is how fit data is for its purpose: accuracy, completeness, consistency, validity, timeliness and uniqueness. Check before you trust any result.")
    m.steps(["Accuracy: reconcile a total with its source (store totals must equal the sales lines).",
             "Completeness: count blanks.",
             "Validity: check values are in an allowed range (discounts between 0 and 80%).",
             "Uniqueness: duplicates are removed in 1.8.9 (60 repeated records in tDirtySales)."])
    m.worked([("tStores25 total sales", "=SUM(tStores25[Sales])", ST.Sales.sum(), N, ""),
              ("FY2025 total from the sales lines", '=SUMIFS(tSales[Sales],tSales[Order_Date],">="&DATE(2025,1,1))', tot25, N, ""),
              ("Reconciled?", '=IF(SUM(tStores25[Sales])=SUMIFS(tSales[Sales],tSales[Order_Date],">="&DATE(2025,1,1)),"Match","Investigate")',
               "Match" if ST.Sales.sum() == tot25 else "Investigate", None, "Accuracy check passed."),
              ("Discounts above 80% in tSales", '=COUNTIF(tSales[Discount],">0.8")', int((D.SALES.Discount > 0.8).sum()), "int", "Validity: none.")])
    m.turn()
    m.ex("Completeness: how many blank Order_Date cells are there in tDirtySales?", "=COUNTBLANK(tDirtySales[Order_Date])", int(DI.Order_Date.isna().sum()),
         "int", short="Blank dates in tDirtySales")
    m.ex("Validity: how many tSales lines have a negative Quantity?", '=COUNTIF(tSales[Quantity],"<0")', int((D.SALES.Quantity < 0).sum()), "int",
         short="Negative quantities")

    # ---------------------------------------------------------------- 1.3.12
    m.topic("1.3.12", "Common mistakes in data analysis",
            "Most wrong answers come from a few avoidable mistakes. The most common in Excel is averaging percentages.")
    m.steps(["Never average a column of rates: small stores count as much as big ones.",
             "Recalculate the rate from totals: total profit ÷ total sales.",
             "Compare the two and use the one from totals."])
    m.worked([("Average of the 12 store margins (wrong)", "=AVERAGE(tStores25[Margin])", ST.Margin.mean(), "pct", "Treats Owerri like Ikeja."),
              ("Chain margin from totals (right)", "=SUM(tStores25[Profit])/SUM(tStores25[Sales])", ST.Profit.sum() / ST.Sales.sum(), "pct", "")])
    m.turn()
    m.ex("What is the (wrong) average of the 12 attainment percentages?", "=AVERAGE(tStores25[Attainment])", ST.Attainment.mean(), "pct",
         short="Average of attainments (wrong method)")
    m.ex("What is the (right) chain attainment from totals?", "=SUM(tStores25[Sales])/SUM(tStores25[Target])", ST.Sales.sum() / ST.Target.sum(), "pct",
         short="Chain attainment from totals")
    m.finish()


def build(b, D):
    m11(b, D)
    m12(b, D)
    m13(b, D)
