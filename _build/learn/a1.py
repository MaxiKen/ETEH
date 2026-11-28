"""Week 1 assignment (Edition 3.1): SuperMart Q1 2025 store review, Week 1 skills only.

    Task 1  analytics words, the four types, the process, industries (1.1 - 1.4)
    Task 2  the Excel window and working with data (1.5 - 1.6)
    Task 3  first formulas: BODMAS and SUM / AVERAGE / COUNT / COUNTA / MIN / MAX (1.7)
    Task 4  Excel tables: sort, filter, Total Row (1.8)
    Task 5  tidy data (1.9)
No SUMIFS, IF or lookups: those are Week 2. Filter-and-read questions are done with the filter button and the Total Row.
"""
from __future__ import annotations

import pandas as pd

from wk_common import L, sm_sales

TYPES4 = ["Descriptive", "Diagnostic", "Predictive", "Prescriptive"]


def build(b):
    sm = sm_sales()
    q1 = sm[(sm.Order_Date >= "2025-01-01") & (sm.Order_Date < "2025-04-01")].copy()
    q1 = q1.sort_values(["Order_Date", "Row_ID"]).reset_index(drop=True)
    q1 = q1[["Row_ID", "Order_ID", "Order_Date", "Ship_Date", "Ship_Mode", "Customer_Name", "Segment", "Store_ID",
             "Store_Name", "Zone", "Category", "Sub_Category", "Product_Name", "Sales", "Quantity", "Discount", "Profit"]]
    st = L("sm_stores").copy()
    st["Opened"] = pd.to_datetime(st.Opened)
    st = st[["Store_ID", "Store_Name", "City", "State", "Zone", "Opened", "Floor_Area_sqm", "Store_Manager",
             "Annual_Target_2025"]]

    t1 = b.module("T1", "T1 Analytics Thinking", "Analytics thinking", "Modules 1.1 - 1.4: data and information, the four "
                  "types of analytics, the analytics process and how industries use data.")
    t2 = b.module("T2", "T2 Excel & Data", "The Excel window and working with data",
                  "Modules 1.5 - 1.6: the Excel window, data types, formats, Find & Replace, Flash Fill and validation.")
    t3 = b.module("T3", "T3 Formulas", "Formulas and the six essential functions",
                  "Module 1.7: operators, BODMAS, references and F4, SUM, AVERAGE, COUNT, COUNTA, MAX, MIN.")
    t4 = b.module("T4", "T4 Tables Sort Filter", "Excel tables: sort, filter and the Total Row",
                  "Module 1.8: tables (Ctrl+T), structured references, sort, filter, Total Row, calculated columns.")
    t5 = b.module("T5", "T5 Tidy Data", "Tidy data", "Module 1.9: the 12 tidy-data laws.")

    r_q1 = b.data("SalesQ1", q1, "One SuperMart sales line, January - March 2025",
                  fmts={"Order_Date": "date", "Ship_Date": "date", "Sales": "naira", "Profit": "naira", "Discount": "pct0"},
                  widths={"Product_Name": 36, "Customer_Name": 20}, used_in="T2, T3, T4")
    r_st = b.data("Stores", st, "One SuperMart store with its 2025 annual target",
                  fmts={"Opened": "date", "Annual_Target_2025": "naira"}, used_in="T2, T3, T4")
    for m_ in (t2, t3, t4):
        m_.set_refs([r_q1, r_st])

    # ================================================================== T1
    m = t1
    m.info("Scenario", "You have just joined SuperMart as a junior analyst. The Regional Manager wants a first review of "
                       "Q1 2025 (January to March) for the 12 stores. Before touching the data, show that you understand "
                       "the basics of analytics.")
    m.part("Part A · Data or information?")
    m.pick("The single value 48,896 in a cell, with no heading.", ["Data", "Information"], "Data", "A raw value with no context.")
    m.pick("'Lagos - Ikeja made the highest Q1 sales of the 12 stores.'", ["Data", "Information"], "Information",
           "Processed, with context and meaning.")
    m.mc("What is data analytics?", ["Typing data into Excel", "Examining data to find useful information and support decisions",
                                     "Buying software", "Drawing charts only"], "B", "Analytics turns data into decisions.")
    m.part("Part B · Which type of analytics?")
    m.pick("What were the total Q1 2025 sales of each store?", TYPES4, "Descriptive", "What happened.")
    m.pick("Why did Abuja - Wuse make a loss in Q1?", TYPES4, "Diagnostic", "Why it happened.")
    m.pick("What will Lagos - Lekki sell in Q2?", TYPES4, "Predictive", "What might happen.")
    m.pick("Which stores should get extra stock for the Easter promotion?", TYPES4, "Prescriptive", "What should be done.")
    m.part("Part C · Process and industries")
    m.mc("What is the FIRST step of the analytics process?", ["Make a chart", "Define the problem or question",
                                                             "Clean the data", "Present the results"], "B",
         "Everything starts with a clear question.")
    m.mc("Which step comes straight after collecting the data?", ["Cleaning the data", "Presenting", "Predicting",
                                                                  "Archiving"], "A", "Collect → clean → analyse → share.")
    m.pick("Tracking hospital bed occupancy each day", ["Retail", "Healthcare", "Banking", "Education"], "Healthcare",
           "Hospitals use data for beds, stays and safety.")
    m.pick("Finding students at risk of failing", ["Retail", "Healthcare", "Banking", "Education"], "Education",
           "Universities track scores and attendance.")

    # ================================================================== T2
    m = t2
    m.part("Part A · The Excel window")
    m.mc("Which part of the Excel window shows the address of the selected cell?", ["Formula Bar", "Name Box", "Status Bar",
                                                                                   "Ribbon"], "B", "Top left, next to the Formula Bar.")
    m.do("On Data_SalesQ1, click in the Name Box, type M2 and press Enter. Which product name is in cell M2?",
         str(q1.Product_Name.iloc[0]), "Name Box › M2.")
    m.num("On Data_Stores, select I2:I13 (Annual_Target_2025). What Sum does the Status Bar show?", None,
          int(st.Annual_Target_2025.sum()), "Read Sum on the Status Bar (bottom right).", fmt="naira", do=True)
    m.part("Part B · Data types and tools")
    D = ["Number", "Text", "Date/Time", "Logical"]
    m.pick("The Order_Date column of tSalesQ1", D, "Date/Time", "Dates are stored as numbers formatted as dates.")
    m.pick("The Store_ID column (e.g. ST01)", D, "Text", "Letters and numbers together = text.")
    m.pick("The Quantity column", D, "Number", "Whole numbers you can add up.")
    m.mc("Which tool fills a column by following a pattern you type in the first cell (e.g. first names)?",
         ["Flash Fill (Ctrl+E)", "Freeze Panes", "Go To", "Sort"], "A", "Flash Fill learns from your example.")
    m.num("On Data_SalesQ1, press Ctrl+F, type Lagos - Ikeja and click Find All. How many cells are found?", None,
          int((q1.Store_Name == "Lagos - Ikeja").sum()), "The count is shown at the bottom of the Find dialog.", do=True)
    m.mc("Which tool stops someone typing a Quantity of −5 or 'abc'?", ["Data Validation", "Flash Fill", "Find & Replace",
                                                                       "Freeze Panes"], "A",
         "Data › Data Validation › Whole number, greater than 0.")
    m.yn("Does formatting a number as currency (₦) change the value stored in the cell?", "No",
         "Formatting changes only how it looks.")

    # ================================================================== T3
    m = t3
    m.part("Part A · BODMAS and references")
    m.num("What does =20+10*2 give?", "=20+10*2", 40, "Multiply first: 10 × 2 = 20, then 20 + 20 = 40.")
    m.num("A sale of ₦10,000 plus 7.5% VAT. What does =10000*(1+7.5%) give?", "=10000*(1+7.5%)", 10750,
          "Brackets first: 1 + 0.075 = 1.075; × 10,000.", fmt="naira")
    m.mc("You copy =B2*$E$1 down one row. What does it become?", ["=B3*$E$2", "=B3*$E$1", "=B2*$E$1", "=C2*$E$1"], "B",
         "$E$1 is absolute (F4) and does not move; B2 is relative and becomes B3.")
    m.part("Part B · The six essential functions on tSalesQ1")
    m.num("What were total Q1 sales? (SUM of Sales)", "=SUM(tSalesQ1[Sales])", int(q1.Sales.sum()), "SUM of the Sales column.",
          fmt="naira")
    m.num("What was total Q1 profit?", "=SUM(tSalesQ1[Profit])", int(q1.Profit.sum()), "SUM of the Profit column.", fmt="naira")
    m.num("How many sales lines are there? (COUNTA of Order_ID)", "=COUNTA(tSalesQ1[Order_ID])", len(q1), "COUNTA counts filled cells.")
    m.num("What is the average sale per line? (nearest naira)", "=ROUND(AVERAGE(tSalesQ1[Sales]),0)", round(float(q1.Sales.mean())),
          "AVERAGE of Sales (round to the nearest naira).", fmt="naira")
    m.num("What is the largest single sale? (MAX)", "=MAX(tSalesQ1[Sales])", int(q1.Sales.max()), "MAX of Sales.", fmt="naira")
    m.num("What is the biggest loss on one line? (MIN of Profit; the answer is negative)", "=MIN(tSalesQ1[Profit])",
          int(q1.Profit.min()), "MIN of Profit.", fmt="naira")
    m.num("Overall Q1 profit margin = total profit ÷ total sales. What is it? (as a %)",
          "=SUM(tSalesQ1[Profit])/SUM(tSalesQ1[Sales])", float(q1.Profit.sum() / q1.Sales.sum()),
          "Divide the two totals; format as %.", fmt="pct")
    m.num("What is the total 2025 annual target of all 12 stores? (tStores)", "=SUM(tStores[Annual_Target_2025])",
          int(st.Annual_Target_2025.sum()), "SUM of the target column.", fmt="naira")

    # ================================================================== T4
    m = t4
    m.part("Part A · Tables")
    m.mc("Which shortcut turns a range into an Excel table?", ["Ctrl+T", "Ctrl+B", "Ctrl+E", "Ctrl+1"], "A",
         "Insert › Table does the same.")
    m.mc("In a table, what does =[@Sales]-[@Profit] in a new column do?", ["Adds a total row",
                                                                          "Works out Sales minus Profit for each row",
                                                                          "Sorts the table", "Deletes the column"], "B",
         "[@Column] means 'this row' (a calculated column).")
    m.part("Part B · Sort and filter (use the arrows on the header row)")
    top = q1.sort_values("Sales", ascending=False).iloc[0]
    m.do("Sort tSalesQ1 by Sales, Largest to Smallest. Which Product_Name is now in the first row?", str(top.Product_Name),
         "Sort Largest to Smallest; read the first row.")
    m.num("Filter Category to show only Furniture. How many lines are shown? (read the status bar: 'x of 500 records found')",
          None, int((q1.Category == "Furniture").sum()), "Filter › Furniture only.", do=True)
    m.num("Clear that filter. Filter Store_Name to Lagos - Ikeja only, turn on the Total Row (Table Design › Total Row) "
          "and choose Sum under Sales. What total appears?", None, int(q1.loc[q1.Store_Name == "Lagos - Ikeja", "Sales"].sum()),
          "The Total Row adds only the visible (filtered) rows.", fmt="naira", do=True)
    m.num("Keep Lagos - Ikeja filtered. In the Total Row choose Sum under Profit. What total appears?", None,
          int(q1.loc[q1.Store_Name == "Lagos - Ikeja", "Profit"].sum()), "Total Row › Sum under Profit.", fmt="naira", do=True)
    m.num("Clear the filters. Filter Zone to North-Central only. How many lines are shown?", None,
          int((q1.Zone == "North-Central").sum()), "Filter › North-Central.", do=True)
    m.yn("Does the Total Row include rows hidden by a filter?", "No", "It uses SUBTOTAL, which skips hidden rows.")

    # ================================================================== T5
    m = t5
    m.info("A messy sheet", "A store manager sends this sheet: (1) the title 'Q1 SALES!!' sits in row 1 above the headers; "
                            "(2) one column called 'Jan Feb Mar' holds three numbers in one cell, e.g. '10 / 12 / 9'; "
                            "(3) two rows are completely empty; (4) Store names are typed 'IKEJA', 'Ikeja ' and 'ikeja'; "
                            "(5) a total row is typed in the middle of the data.")
    L5 = ["One header row only", "One value per cell", "No empty rows", "Consistent spelling", "No totals inside the data"]
    m.pick("Problem (1): the title above the headers", L5, "One header row only", "Headers must be in the first row.")
    m.pick("Problem (2): 'Jan Feb Mar' in one cell", L5, "One value per cell", "Split into separate values.")
    m.pick("Problem (3): the empty rows", L5, "No empty rows", "Blank rows break tables and filters.")
    m.pick("Problem (4): IKEJA / Ikeja / ikeja", L5, "Consistent spelling", "One spelling per value.")
    m.pick("Problem (5): the typed total row", L5, "No totals inside the data", "Totals belong in the Total Row or a report.")
    m.yn("In tidy data, is each row one record (e.g. one sales line)?", "Yes", "Rows = records, columns = variables.")
