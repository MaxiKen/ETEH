"""Week 1 Section Examples (Edition 3.1): questions in teaching order.

Progression rule for Week 1
    1.1 - 1.4  words and ideas only (data, information, analytics, process, industries): no numbers to work out
    1.5 - 1.6  find things in Excel and read the status bar: still no formulas
    1.7        first formulas: operators, BODMAS, references, and the six essential functions on tStores25
    1.8        Excel tables: structured references, sort, filter, Total Row, calculated columns
    1.9        tidy data laws, spotting problems in tDirtySales
"""
from __future__ import annotations

import re

import pandas as pd

import w1_data as W

DATA_TYPES = ["Number", "Text", "Date/Time", "Logical", "Error"]
TYPES4 = ["Descriptive", "Diagnostic", "Predictive", "Prescriptive"]


def n(x):
    return f"{x:,.0f}"


def naira(x):
    return f"₦{x:,.0f}" if x >= 0 else f"-₦{-x:,.0f}"


def build(b):
    st = W.table("tStores25")
    stu = W.table("tStudents")
    sales = W.table("tSales")
    dirty = W.table("tDirtySales").copy()
    iso = re.compile(r"\d{4}-\d\d-\d\d$")
    dirty["Order_Date"] = [pd.Timestamp(v) if isinstance(v, str) and iso.match(v) else v for v in dirty.Order_Date]

    # ------------------------------------------------------------------ module sheets first (tab order)
    m11 = b.module("1.1", "1.1 Intro to Analytics", "Introduction to Data Analytics",
                   "Data, information, data analytics, why we analyse data, the analyst's role, asking the right "
                   "question, general and domain knowledge, and deciding from evidence rather than assumptions.")
    m12 = b.module("1.2", "1.2 Types of Analytics", "Types of Data Analytics",
                   "The four types: descriptive (what happened), diagnostic (why), predictive (what might happen) "
                   "and prescriptive (what should be done).")
    m13 = b.module("1.3", "1.3 Analytics Process", "The Data Analytics Process",
                   "The steps from defining the problem, through collecting, cleaning, transforming and analysing data, "
                   "to interpreting, visualising and communicating results; accuracy and common mistakes.")
    m14 = b.module("1.4", "1.4 Industry Applications", "Applications of Data Analytics Across Industries",
                   "How retail, education, healthcare, banking, government, marketing, HR, logistics and mining use data.")
    m15 = b.module("1.5", "1.5 Excel Interface", "Introduction to Microsoft Excel",
                   "Editions and platforms, file formats, saving safely, the Excel window (Ribbon, Name Box, Formula Bar, "
                   "Status Bar), the grid, navigating, and sheet tabs.")
    m16 = b.module("1.6", "1.6 Working with Data", "Working with Data in Excel",
                   "Entering values, the six data types, editing, copy and Paste Special, AutoFill, Flash Fill, "
                   "formatting (Ctrl+1), Freeze Panes, Find & Replace, Go To Special and Data Validation.")
    m17 = b.module("1.7", "1.7 Formulas & Functions", "Excel Formulas and Core Functions",
                   "Formulas start with =, operators and BODMAS, relative and absolute references (F4), "
                   "AutoSum and the six essential functions SUM, AVERAGE, COUNT, COUNTA, MAX, MIN, and error values.")
    m18 = b.module("1.8", "1.8 Excel Tables", "Organising Data with Excel Tables",
                   "Rows are records and columns are variables, Ctrl+T, table names, structured references such as "
                   "tStores25[Sales], sorting, filtering, Remove Duplicates, calculated columns and the Total Row.")
    m19 = b.module("1.9", "1.9 Tidy Data Laws", "Twelve Tidy Data Laws",
                   "The twelve laws that keep data easy to analyse, applied to the raw export tDirtySales.")

    # ------------------------------------------------------------------ data sheets (after the module sheets)
    ref_st = b.data("Stores25", st.drop(columns=["Margin", "Attainment"]),
                    "One SuperMart store, financial year 2025 (12 stores)",
                    fmts={"Sales": "naira", "Profit": "naira", "Target": "naira", "Orders": "int"},
                    widths={"Store_Name": 27, "Store_Manager": 18, "Sales": 14, "Profit": 13, "Target": 14},
                    used_in="1.5, 1.6, 1.7, 1.8",
                    calc={"Margin": ("=[@Profit]/[@Sales]", "pct", st.Margin.tolist()),
                          "Attainment": ("=[@Sales]/[@Target]", "pct", st.Attainment.tolist())})
    ref_stu = b.data("Students", stu, "One Kingsway University student, 2024/2025 first semester",
                     fmts={"Level": "gen"}, used_in="1.7, 1.8")
    ref_sa = b.data("Sales", sales, "One SuperMart sales line (one product on one order), 2022-2025",
                    fmts={"Row_ID": "gen", "Order_Date": "date", "Sales": "naira", "Profit": "naira", "Discount": "pct0"},
                    widths={"Order_ID": 16, "Customer_Name": 20, "Store_Name": 26, "Category": 15}, used_in="1.8")
    ref_d = b.data("DirtySales", dirty, "One line of the raw FY2025 store export, exactly as received (deliberately untidy)",
                   fmts={"Order_Date": "date", "Row_ID": "gen"}, widths={"Order_ID": 16, "Store_Name": 26, "Zone": 16},
                   used_in="1.6, 1.9")
    for m, refs in ((m15, [ref_st]), (m16, [ref_st, ref_d]), (m17, [ref_st, ref_stu]), (m18, [ref_st, ref_stu, ref_sa]),
                    (m19, [ref_d])):
        m.set_refs(refs)

    # ================================================================== 1.1
    m = m11
    m.part("Part A · Data or information?")
    m.mc("What is data?",
         ["Raw facts and figures that have not yet been organised or given meaning",
          "A report that explains what the numbers mean",
          "A decision taken by a manager",
          "A chart shown at a board meeting"], "A",
         "Data is the raw material: facts, numbers and text as they are recorded.")
    m.mc("What is information?",
         ["Any number typed into Excel",
          "Data that has been organised and given context so that it is useful",
          "A list of store codes",
          "A file saved in .csv format"], "B",
         "Information = data + context + meaning. It answers a question someone has.")
    m.pick("A cell holds 344, with no heading and no label. Is this data or information?", ["Data", "Information"], "Data",
           "On its own, 344 has no meaning: we do not know what it counts. It is data.")
    m.pick("'Lagos - Ikeja sold ₦146.4m in 2025, 3% above its target.' Data or information?", ["Data", "Information"],
           "Information", "The figure has context (which store, which year, compared with what), so it is information.")
    m.pick("A list of 9,994 sales lines exported from the till system. Data or information?", ["Data", "Information"],
           "Data", "A raw export has not been summarised or explained yet: it is data.")
    m.pick("A one-line summary: 'Five of our 12 stores made a loss in 2025.' Data or information?",
           ["Data", "Information"], "Information", "It has been organised to answer a question, so it is information.")
    m.part("Part B · Data analytics and the analyst")
    m.mc("Which sentence best describes data analytics?",
         ["Typing data into a spreadsheet",
          "Examining data to find patterns and answer questions that support decisions",
          "Buying a bigger computer",
          "Designing a company logo"], "B",
         "Data analytics turns data into answers that help people decide.")
    m.mc("What is the main reason organisations analyse data?",
         ["To make spreadsheets look nicer",
          "To make decisions based on evidence rather than guesses",
          "To avoid talking to customers",
          "To store more files"], "B",
         "Analysis replaces assumptions with evidence.")
    m.mc("Whose main job is to clean, analyse and report on data so managers can decide?",
         ["Data Analyst", "Data Scientist", "Business Analyst", "Database Administrator"], "A",
         "The data analyst turns data into reports and insights. A data scientist focuses on predictive models; "
         "a business analyst focuses on processes and requirements.")
    m.yn("Before you start analysing, should you first agree the exact question with the person who asked?", "Yes",
         "Asking the right question first saves time and makes sure the answer is useful.")
    m.mc("Knowing that a hospital worries about patients who come back within 30 days is an example of ...",
         ["General knowledge", "Domain knowledge", "Excel knowledge", "No knowledge needed"], "B",
         "Domain knowledge is knowledge of a particular industry or field.")
    m.tf("Data analytics is only useful for very large companies.", False,
         "Analytics is used everywhere: shops, schools, hospitals, and in everyday life.")

    # ================================================================== 1.2
    m = m12
    m.part("Part A · Which type of analytics?")
    m.note("Pick the type of analytics each question needs: Descriptive (what happened?), Diagnostic (why did it happen?), "
           "Predictive (what might happen?), Prescriptive (what should we do?).")
    for q, a, why in [
        ("What were total sales for each SuperMart store in 2025?", "Descriptive", "It reports what happened."),
        ("Why did Abuja - Wuse make a loss in 2025?", "Diagnostic", "It asks why something happened."),
        ("How many patients will Mercyland Hospital admit next December?", "Predictive", "It asks about the future."),
        ("What discount limit should SuperMart set to protect its profit?", "Prescriptive", "It asks what action to take."),
        ("How many Kingsway students were marked at risk last semester?", "Descriptive", "It counts what happened."),
        ("Why are online travel agent bookings at Atlantic Pearl Hotel cancelled more often?", "Diagnostic",
         "It looks for the cause."),
        ("Which students are likely to fail next semester if nothing changes?", "Predictive", "It estimates a future outcome."),
        ("How many nurses should each ward roster next week?", "Prescriptive", "It recommends an action."),
    ]:
        m.pick(q, TYPES4, a, why)
    m.part("Part B · Check your understanding")
    m.mc("Which type of analytics answers the question 'What might happen?'", TYPES4, "C",
         "Predictive analytics looks ahead.")
    m.tf("Descriptive analytics is about what has already happened.", True, "It summarises the past.")
    m.yn("Does prescriptive analytics recommend an action?", "Yes", "Prescriptive = what should be done.")

    # ================================================================== 1.3
    m = m13
    m.part("Part A · The steps of the process")
    m.mc("What is the FIRST step of the data analytics process?",
         ["Draw a chart", "Understand the problem and define the objective", "Clean the data", "Write the report"], "B",
         "Everything else depends on knowing what question you are answering.")
    m.mc("Which step comes straight after collecting the data?",
         ["Communicating findings", "Cleaning and preparing the data", "Defining the objective", "Building a dashboard"], "B",
         "Raw data nearly always has problems that must be fixed before analysis.")
    m.pick("Fixing misspelt store names and removing duplicate rows belongs to which step?",
           ["Collection", "Cleaning", "Analysis", "Presentation"], "Cleaning", "Fixing errors is data cleaning.")
    m.pick("Typing figures from paper forms into Excel is an example of which step?",
           ["Collection", "Cleaning", "Analysis", "Presentation"], "Collection",
           "Manual data entry is one way of collecting data.")
    m.pick("Comparing the sales of the 12 stores to find the best and worst belongs to which step?",
           ["Collection", "Cleaning", "Analysis", "Presentation"], "Analysis", "Comparing and summarising is analysis.")
    m.pick("Showing a bar chart of sales by store to the managers belongs to which step?",
           ["Collection", "Cleaning", "Analysis", "Presentation"], "Presentation",
           "Charts and slides present the results.")
    m.part("Part B · Quality and mistakes")
    m.yn("Should you check your totals against a trusted figure (for example from Finance) before sharing results?", "Yes",
         "Checking against a control total is the simplest accuracy check.")
    m.mc("Which of these is a common mistake in data analysis?",
         ["Checking the data for errors first", "Jumping to a conclusion before checking the data",
          "Writing down where the data came from", "Asking what decision the analysis supports"], "B",
         "Conclusions drawn from unchecked data are often wrong.")
    m.tf("Interpretation means explaining what the results mean for the business.", True,
         "Interpretation turns numbers into meaning: so what?")

    # ================================================================== 1.4
    m = m14
    IND = ["Retail", "Education", "Healthcare", "Banking", "Logistics", "Human Resources", "Marketing", "Mining"]
    m.part("Part A · Which industry?")
    m.note("Pick the industry that each measure belongs to.")
    for q, a in [("Bed occupancy and patient readmission rate", "Healthcare"),
                 ("Unusual card payments that might be fraud", "Banking"),
                 ("Students at risk of failing, based on attendance", "Education"),
                 ("On-time delivery rate and the best delivery routes", "Logistics"),
                 ("Staff turnover and fairness of pay", "Human Resources"),
                 ("Best-selling products and stock levels in each store", "Retail"),
                 ("Response rate to an SMS promotion", "Marketing"),
                 ("Daily tonnes of ore and the grade of the ore", "Mining")]:
        m.pick(q, IND, a, f"This is a typical {a.lower()} measure.")
    m.part("Part B · Check your understanding")
    m.yn("Can the same analytics process (question, data, clean, analyse, present) be used in every industry?", "Yes",
         "The process is the same; only the data and the domain knowledge change.")

    # ================================================================== 1.5
    m = m15
    m.part("Part A · Files and the Excel window")
    m.mc("Which file format can store macros?", [".xlsx", ".xlsm", ".csv", ".txt"], "B",
         ".xlsm = macro-enabled workbook. A normal .xlsx cannot store macros.")
    m.mc("Which file format is plain text with values separated by commas, and keeps no formatting or formulas?",
         [".xlsx", ".xlsb", ".csv", ".xls"], "C", "CSV = comma-separated values.")
    m.mc("Where can you see the address of the selected cell (for example C5)?",
         ["Status Bar", "Name Box", "Ribbon", "Sheet tab"], "B", "The Name Box is to the left of the Formula Bar.")
    m.mc("Where can you read and edit the full contents of the selected cell?",
         ["Formula Bar", "Quick Access Toolbar", "Status Bar", "Backstage view"], "A",
         "The Formula Bar shows what is really in the cell, including formulas.")
    m.mc("Which part of the window shows Sum, Average and Count when you select some numbers?",
         ["Name Box", "Ribbon", "Status Bar", "Formula Bar"], "C", "The Status Bar is at the bottom of the window.")
    m.tf("Excel for the web has every feature of desktop Excel.", False,
         "The web version is lighter; some features (for example Power Pivot) need the desktop version.")
    m.part("Part B · The grid and moving around")
    m.num("How many rows does one worksheet have?", None, 1048576, "Every modern worksheet has 1,048,576 rows.")
    m.num("How many columns does one worksheet have?", None, 16384, "16,384 columns, from A to XFD.")
    m.short("What is the address of the cell in column D, row 7?", "D7", "Column letter first, then row number.")
    m.mc("Which shortcut jumps to the last filled cell going down a column?",
         ["Ctrl + Down Arrow", "Ctrl + S", "Alt + =", "Ctrl + 1"], "A", "Ctrl + an arrow key jumps to the edge of the data.")
    m.part("Part C · Try it in Excel")
    m.do("Open sheet Data_Stores25. Click in the Name Box, type B7 and press Enter. Which store name is in cell B7?",
         st.Store_Name.iloc[5], "Row 1 holds the headers, so row 7 is the 6th store (ST06).")
    m.num("On Data_Stores25, select cells G2:G13 (the Orders column). What Sum does the Status Bar show?", None,
          int(st.Orders.sum()), f"The Status Bar adds the selected numbers: {n(st.Orders.sum())} orders. No formula needed.",
          do=True)
    m.num("Keep G2:G13 selected. What Count does the Status Bar show?", None, 12,
          "Count = how many cells are not empty: 12 stores.", do=True)

    # ================================================================== 1.6
    m = m16
    m.part("Part A · Data types")
    m.note("Excel's data types: Number, Text, Date/Time, Logical, Error (and Array, which you meet in Week 2).")
    for q, a in [("Lagos - Ikeja", "Text"), ("146388345", "Number"), ("TRUE", "Logical"),
                 ("2025-03-18 (stored as a real date)", "Date/Time"), ("#DIV/0!", "Error")]:
        m.pick(f"What data type is this value?   {q}", DATA_TYPES, a, f"{q.split(' (')[0]} is a {a} value.")
    m.part("Part B · Everyday tools")
    m.mc("Which shortcut opens the Format Cells dialog?", ["Ctrl + 1", "Ctrl + E", "Ctrl + T", "F4"], "A",
         "Ctrl+1 opens Format Cells (number, alignment, font, border, fill).")
    m.mc("Which shortcut runs Flash Fill?", ["Ctrl + 1", "Ctrl + E", "Ctrl + H", "Ctrl + Z"], "B",
         "Flash Fill (Ctrl+E) copies a pattern you have typed.")
    m.mc("You want the header row to stay visible when you scroll down. What do you use?",
         ["Wrap Text", "Freeze Panes", "Merge & Center", "AutoFill"], "B", "View › Freeze Panes › Freeze Top Row.")
    m.mc("You want to paste only the values, not the formulas. What do you use?",
         ["Paste Special › Values", "Cut", "Flash Fill", "Clear Formats"], "A", "Paste Special › Values (Ctrl+Alt+V, then V).")
    m.yn("Should you merge cells inside a data table?", "No",
         "Merged cells break sorting and filtering. Use Center Across Selection for titles instead.")
    m.mc("Which tool limits a cell to a drop-down list of allowed values?",
         ["Conditional Formatting", "Data Validation", "Freeze Panes", "Find & Replace"], "B", "Data › Data Validation › List.")
    m.part("Part C · Try it in Excel")
    m.num("On Data_DirtySales press Ctrl+F, type LAGOS - IKEJA, click Options, tick 'Match case', then click Find All. "
          "How many cells are found?", None, int((dirty.Store_Name == "LAGOS - IKEJA").sum()),
          "Find All lists every match at the bottom of the dialog: the store name typed in capitals.", do=True)
    m.do("On Data_Stores25, look at cell C2. Is the value Text or Number? (Type Text or Number.)", "Text",
         "C2 holds the zone 'South-West', which is text.")

    # ================================================================== 1.7
    m = m17
    m.example("the six essential functions on tStores25",
              ["Click an empty cell on this sheet and type = to start a formula.",
               "Type the function name and an opening bracket, e.g. =SUM(",
               "Type the column, e.g. tStores25[Sales] (or go to Data_Stores25 and select E2:E13), close the bracket, press Enter."],
              [("Total sales of all 12 stores", "=SUM(tStores25[Sales])", "=SUM(tStores25[Sales])", int(st.Sales.sum()), "naira"),
               ("Average sales per store", "=AVERAGE(tStores25[Sales])", "=AVERAGE(tStores25[Sales])", float(st.Sales.mean()), "naira"),
               ("Profit margin of Lagos - Ikeja (Profit ÷ Sales)", "=Data_Stores25!F2/Data_Stores25!E2",
                "=Data_Stores25!F2/Data_Stores25!E2", float(st.Profit.iloc[0] / st.Sales.iloc[0]), "pct")])
    m.part("Part A · How formulas work")
    m.mc("Every formula in Excel starts with ...", ["+", "=", "#", "@"], "B", "The equal sign tells Excel a formula follows.")
    m.mc("Which of these is a FUNCTION?", ["=E2-F2", "=SUM(E2:E13)", "=E2*2", "=100/4"], "B",
         "A function is a built-in named formula, such as SUM.")
    m.num("What result does =10+2*3 give?", "=10+2*3", 16, "Multiply first (BODMAS): 2×3 = 6, then 10 + 6 = 16.")
    m.num("What result does =(10+2)*3 give?", "=(10+2)*3", 36, "Brackets first: 10 + 2 = 12, then 12 × 3 = 36.")
    m.num("What result does =100-20/4 give?", "=100-20/4", 95, "Divide first: 20 ÷ 4 = 5, then 100 − 5 = 95.")
    m.mc("In the formula =$H$2, the dollar signs make the reference ...", ["Relative", "Absolute", "Mixed", "Wrong"], "B",
         "$ fixes the column and the row, so the reference does not move when copied.")
    m.mc("Which key cycles a reference through $H$2, H$2, $H2 and H2?", ["F2", "F4", "F9", "Ctrl"], "B", "F4 toggles the $ signs.")
    m.mc("You copy =E2/H2 from row 2 down to row 3. What does it become?", ["=E2/H2", "=E3/H3", "=F2/I2", "=E3/H2"], "B",
         "Relative references move with the formula: one row down.")
    m.part("Part B · The six essential functions on tStores25")
    m.num("What is the total Target of all 12 stores? (SUM of the Target column)", "=SUM(tStores25[Target])",
          int(st.Target.sum()), "=SUM(tStores25[Target]) adds every store's target.", fmt="naira")
    m.num("What is the largest Profit of any store? (MAX)", "=MAX(tStores25[Profit])", int(st.Profit.max()),
          "=MAX(tStores25[Profit]). It belongs to Lagos - Ikeja.", fmt="naira")
    m.num("What is the smallest Profit? (MIN). Hint: it is a loss, so the answer is negative.", "=MIN(tStores25[Profit])",
          int(st.Profit.min()), "=MIN(tStores25[Profit]). Abuja - Wuse lost this amount.", fmt="naira")
    m.num("How many stores are in the table? (COUNTA of Store_Name)", "=COUNTA(tStores25[Store_Name])", 12,
          "COUNTA counts non-empty cells, including text.")
    m.num("What is the average number of Orders per store? Round to 1 decimal place.", "=ROUND(AVERAGE(tStores25[Orders]),1)",
          round(float(st.Orders.mean()), 1), "=AVERAGE(tStores25[Orders]) then round to 1 decimal place.", fmt="dec1")
    m.num("What is the average Final_Score of all Kingsway students (tStudents)? Round to 1 decimal place.",
          "=ROUND(AVERAGE(tStudents[Final_Score]),1)", round(float(stu.Final_Score.mean()), 1),
          "=AVERAGE(tStudents[Final_Score]) over 5,840 students.", fmt="dec1")
    m.num("What is the profit margin of Jos - Rayfield (row 10 on Data_Stores25)? Profit ÷ Sales, as a percentage.",
          "=Data_Stores25!F10/Data_Stores25!E10", float(st.Profit.iloc[8] / st.Sales.iloc[8]),
          "=F10/E10 on Data_Stores25. Format the cell as a percentage.", fmt="pct")
    m.part("Part C · Error values")
    m.mc("What does #DIV/0! mean?", ["The column is too narrow", "You divided by zero or by a blank cell",
                                     "A function name is misspelt", "A lookup found no match"], "B",
         "Check the divisor (the number you divide by).")
    m.mc("You typed =SUMM(E2:E13). Which error appears?", ["#VALUE!", "#REF!", "#NAME?", "#N/A"], "C",
         "#NAME? = Excel does not recognise the name. Fix the spelling to SUM.")

    # ================================================================== 1.8
    m = m18
    m.example("structured references",
              ["Click any cell in the data and press Ctrl+T to turn it into a table (the data sheets are already tables).",
               "Name the table on the Table Design tab, e.g. tStores25.",
               "Refer to a whole column by name: tStores25[Orders]. The formula keeps working when rows are added."],
              [("Total orders", "=SUM(tStores25[Orders])", "=SUM(tStores25[Orders])", int(st.Orders.sum()), "int"),
               ("Number of students", "=COUNTA(tStudents[Matric_No])", "=COUNTA(tStudents[Matric_No])", len(stu), "int")])
    m.part("Part A · Tables and records")
    m.mc("In a dataset, each row is a ... and each column is a ...",
         ["variable / record", "record / variable", "chart / table", "sheet / workbook"], "B",
         "One row = one record (e.g. one store). One column = one variable (e.g. Sales).")
    m.mc("Which shortcut turns a range into an Excel table?", ["Ctrl + T", "Ctrl + 1", "Ctrl + E", "Alt + ="], "A",
         "Ctrl+T (Insert › Table).")
    m.mc("What is the structured reference to the Profit column of tStores25?",
         ["tStores25!Profit", "tStores25[Profit]", "Profit(tStores25)", "F:F"], "B", "TableName[ColumnName].")
    m.yn("Can two tables in the same workbook have the same name?", "No", "Every table name must be unique in the workbook.")
    m.mc("Why is an Excel table better than a normal range?",
         ["It cannot be sorted", "It grows automatically and formulas fill down by themselves",
          "It hides the data", "It only works in Excel for the web"], "B",
         "New rows join the table and calculated columns fill automatically.")
    m.part("Part B · Try it in Excel")
    m.do("Sort tStores25 by Profit, Smallest to Largest (click the arrow on the Profit header). Which store is now in the first row?",
         st.sort_values("Profit").Store_Name.iloc[0], "The smallest profit (the biggest loss) comes first.")
    m.num("Filter tStores25 so that Zone shows only South-South. How many stores are shown?", None,
          int((st.Zone == "South-South").sum()), "Port Harcourt - GRA, Benin City - Ring Road and Uyo - Ewet. "
          "Clear the filter afterwards.", do=True)
    m.num("Filter tStudents so that At_Risk shows only Yes. How many students are found? "
          "(Read it at the bottom left: 'x of 5840 records found'.)", None, int((stu.At_Risk == "Yes").sum()),
          "The status bar shows the filtered count. Clear the filter afterwards.", do=True)
    m.num("Filter tSales so that Category shows only Furniture. How many rows are found?", None,
          int((sales.Category == "Furniture").sum()), "Read the count at the bottom left of the window.", do=True)
    m.do("Look at the Attainment column of tStores25 (a calculated column). Which store has the highest attainment? "
         "(Sort Largest to Smallest.)", st.sort_values("Attainment", ascending=False).Store_Name.iloc[0],
         "Attainment = Sales ÷ Target. Lagos - Lekki reached 112.4% of its target.")
    m.num("Turn on the Total Row (Table Design › Total Row) for tStores25 and choose Sum under Profit. What total appears?",
          "=SUM(tStores25[Profit])", int(st.Profit.sum()), "The Total Row uses SUBTOTAL(109,[Profit]) = the same as "
          "=SUM(tStores25[Profit]). Turn the Total Row off again afterwards.", fmt="naira", do=True)

    # ================================================================== 1.9
    m = m19
    LAWS = {
        "1": "Law 1: One record per row", "2": "Law 2: One variable per column", "3": "Law 3: Atomic values",
        "5": "Law 5: No spacer rows", "6": "Law 6: ISO dates", "7": "Law 7: Blank vs zero",
        "8": "Law 8: Raw separate from reports", "9": "Law 9: Control your vocabulary",
        "10": "Law 10: Calculate, don't type", "11": "Law 11: No merged cells"}
    m.part("Part A · Which law is broken?")
    m.mc("tDirtySales has the store written as 'Lagos - Ikeja', 'LAGOS - IKEJA' and 'lagos - ikeja'. Which law is broken?",
         [LAWS["1"], LAWS["6"], LAWS["9"], LAWS["11"]], "C", "One value should be written exactly one way.")
    m.mc("Row_ID 6748 appears twice in tDirtySales. Which law is broken?",
         [LAWS["1"], LAWS["3"], LAWS["7"], LAWS["10"]], "A", "Every record must appear exactly once.")
    m.mc("Some Order_Date values look like 14/05/2025 or 04-May-2025. Which law is broken?",
         [LAWS["2"], LAWS["6"], LAWS["8"], LAWS["9"]], "B", "Dates should be real dates shown as yyyy-mm-dd.")
    m.mc("A cell holds 'Abuja, FCT' (city and state together). Which law is broken?",
         [LAWS["3"], LAWS["5"], LAWS["7"], LAWS["11"]], "A", "Atomic values: one value per cell. Split city and state.")
    m.mc("Sales for Q1, Q2, Q3 and Q4 are in four separate columns. Which law is broken?",
         [LAWS["1"], LAWS["2"], LAWS["6"], LAWS["10"]], "B", "Quarter is one variable, so it needs one column.")
    m.mc("Someone typed 103% in the Attainment column instead of a formula. Which law is broken?",
         [LAWS["5"], LAWS["8"], LAWS["10"], LAWS["11"]], "C", "Derived values should be formulas, e.g. =[@Sales]/[@Target].")
    m.part("Part B · Check your understanding")
    m.yn("Is it a good idea to type 0 when a value is unknown?", "No",
         "Leave it blank. AVERAGE ignores blanks but counts zeros, so a typed 0 drags the average down.")
    m.yn("Should the raw data be kept unchanged, with cleaning done in a separate copy or step?", "Yes",
         "Law 8: raw, clean and report are separate layers.")
    m.num("How many completely empty rows are in tDirtySales? (The slides for Law 5 tell you.)", None,
          int(dirty.isna().all(axis=1).sum()), "15 spacer rows: Law 5 says data must be one continuous block.")
    m.mc("A Discount column holds both 0.2 and 20%. Which law does this break?",
         [LAWS["2"], LAWS["5"], LAWS["8"], LAWS["11"]], "A",
         "One variable in two forms breaks Law 2 (see the Watch Out on the Law 2 slide).")
