"""Week 1 workbook, modules 1.7 - 1.9 (edition 3)."""
from __future__ import annotations

import datetime as dt

import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell

from w1x import ModuleSheet
from xlhelp import LAST_COL

N = "naira"


def A(r, c, ab=False):
    return xl_rowcol_to_cell(r, c, ab, ab)


def serial(d):
    return (d - dt.date(1899, 12, 30)).days


# =====================================================================================================
def m17(b, D):
    ST, STU, IK = D.ST, D.STU, D.IK
    F = b.F
    m = ModuleSheet(b, "1.7 Formulas & Functions", "1.7  Excel Formulas and Core Functions",
                    "Tutor walk-through after Module 1.7: formulas, order of operations, references, the core functions and formula errors",
                    "1.7.1 - 1.7.14")
    m.dataset([("tStores25", "Sheet Data_Stores25 (starts in A1)", "Row 2 = Lagos - Ikeja. Columns: E Sales, F Profit, G Orders, H Target, I Margin, J Attainment. "
                                                                    "From this sheet we refer to its cells as Data_Stores25!E2 (a reference to another sheet)."),
               ("tStudents", "Sheet Data_Students (starts in A1)", "Crestview students, used in 1.7.14."),
               ("tPayMix, tDept23, tHotelSeg", "Sheet 1.4 Industry Applications", "Summary tables reused in 1.7.14."),
               ("Practice ranges", "On this sheet", "Small ranges for writing and copying your own formulas (white cells).")])
    ik, lk = ST.iloc[0], ST.iloc[1]

    # ---------------------------------------------------------------- 1.7.1
    m.topic("1.7.1", "What is a formula?",
            "A formula is an expression that calculates a value. In Excel every formula starts with = and can combine values, cell references, operators and functions.")
    m.steps(["Click an empty cell and type =146388345-29366453. It works, but never updates: the numbers are typed in.",
             "Better: refer to the cells. =Data_Stores25!E2-Data_Stores25!F2 (Sales minus Profit = cost).",
             "Change the source and the result follows. Prefer references to typed numbers every time."])
    m.worked([("Typed numbers (never updates)", "=146388345-29366453", int(ik.Sales - ik.Profit), N, ""),
              ("Cell references: Ikeja cost", "=Data_Stores25!E2-Data_Stores25!F2", int(ik.Sales - ik.Profit), N, "Same answer, but live."),
              ("Ikeja margin", "=Data_Stores25!F2/Data_Stores25!E2", ik.Margin, "pct", "")])
    m.turn()
    m.ex("Write a formula for Lagos - Lekki's cost (row 3: Sales minus Profit).", "=Data_Stores25!E3-Data_Stores25!F3", int(lk.Sales - lk.Profit), N,
         short="Lekki cost")

    # ---------------------------------------------------------------- 1.7.2
    m.topic("1.7.2", "What is a function?",
            "A function is a ready-made formula built into Excel: =NAME(argument1, argument2, …). The arguments are the values you give it.")
    m.steps(["Type =SUM( and select the values, or use the table name: =SUM(tStores25[Sales]).",
             "Close the bracket and press Enter.",
             "Some functions take two arguments: =ROUND(value, decimals)."])
    m.worked([("=SUM(tStores25[Sales])", "=SUM(tStores25[Sales])", ST.Sales.sum(), N, "One argument: a range."),
              ("=MAX(tStores25[Sales])", "=MAX(tStores25[Sales])", ST.Sales.max(), N, ""),
              ("=ROUND(Ikeja margin, 3)", "=ROUND(Data_Stores25!F2/Data_Stores25!E2,3)", round(ik.Margin, 3), "dec3", "Two arguments.")])
    m.turn()
    m.ex("What is the smallest store Sales figure? (MIN)", "=MIN(tStores25[Sales])", ST.Sales.min(), N, short="Smallest Sales")
    m.ex("Round Lagos - Lekki's margin (row 3) to 3 decimals.", "=ROUND(Data_Stores25!F3/Data_Stores25!E3,3)", round(lk.Margin, 3), "dec3", tol=0.0004,
         short="Lekki margin, 3 decimals")

    # ---------------------------------------------------------------- 1.7.3
    m.topic("1.7.3", "Difference between a formula and a function",
            "A formula is the whole calculation you write; a function is one named tool used inside it. Every function is used inside a formula, but not every formula uses a function.")
    m.steps(["=F2/E2 is a formula with no function.", "SUM(E2:E13) is a function.", "=SUM(F2:F13)/SUM(E2:E13) is one formula using two functions."])
    m.worked([("Chain margin: one formula, two functions", "=SUM(tStores25[Profit])/SUM(tStores25[Sales])", ST.Profit.sum() / ST.Sales.sum(), "pct", "")])
    m.turn()
    m.ex_text("How many functions are in =ROUND(SUM(F2:F13)/SUM(E2:E13),3)?", "3")
    m.ex("What does =SUM(tStores25[Orders])/COUNT(tStores25[Orders]) return (average orders per store)?", "=SUM(tStores25[Orders])/COUNT(tStores25[Orders])",
         ST.Orders.mean(), "dec1", short="Average orders per store")

    # ---------------------------------------------------------------- 1.7.4
    m.topic("1.7.4", "The equal sign and formula structure",
            "Start with =, then the calculation. Brackets come in pairs; commas separate arguments (some regions use semicolons). Excel colours each reference to help you check it.")
    m.steps(["Without =, Excel stores the text 'SUM(E2:E13)'.", "Every ( needs a ).", "Press Enter to finish, F2 to edit, Esc to cancel."])
    m.turn()
    m.ex_text("What is wrong with SUM(E2:E13) typed into a cell? A) missing =  B) missing bracket  C) wrong range", "A")
    m.ex_text("What is wrong with =SUM(E2:E13 ? A) missing =  B) missing bracket  C) wrong range", "B")
    m.ex("Type =SUM(tStores25[Orders]) in the white cell. What does it return?", "=SUM(tStores25[Orders])", int(ST.Orders.sum()), "int", short="Total orders")

    # ---------------------------------------------------------------- 1.7.5
    m.topic("1.7.5", "Arithmetic operators and order of operations (BODMAS)",
            "Excel works in a fixed order: Brackets, then powers (^), then Division and Multiplication left to right, then Addition and Subtraction. Use brackets to make the order certain.")
    m.steps(["Operators: + - * / ^ (power), % (percent), & (join text), = <> > < (compare).",
             "Work out the brackets first, then ^, then * and /, then + and -.",
             "When in doubt, add brackets."])
    m.worked([("=10+20*3", "=10+20*3", 70, "int", "Multiply first."), ("=(10+20)*3", "=(10+20)*3", 90, "int", "Brackets first."),
              ("=100-20/4", "=100-20/4", 95, "int", "Divide first."), ("=2^3", "=2^3", 8, "int", "Power."),
              ("=50000*(1+7.5%)", "=50000*(1+7.5%)", 53750, N, "Price plus 7.5% VAT.")])
    m.turn()
    m.ex("Without Excel, work out =20-4*3+2. Type your answer.", "=20-4*3+2", 10, "int", short="20-4*3+2")
    m.ex("Work out =(20-4)*(3+2).", "=(20-4)*(3+2)", 80, "int", short="(20-4)*(3+2)")
    m.ex("Write one formula for Lagos - Ikeja's sales plus 7.5% VAT (row 2).", "=Data_Stores25!E2*(1+7.5%)", float(ik.Sales * 1.075), N, tol=1,
         short="Ikeja sales plus VAT")

    # ---------------------------------------------------------------- 1.7.6
    m.topic("1.7.6", "Cell references and reference types",
            "A cell reference identifies a cell or range by column letter and row number (E2, E2:E13), so a formula uses the cell's current value. "
            "Other types: another sheet (Data_Stores25!H2) and table columns (tStores25[Sales]).")
    m.steps(["Single cell: Data_Stores25!H2.", "Range: Data_Stores25!E2:E13.", "Table column: tStores25[Target], which grows with the table."])
    m.worked([("A cell on another sheet", "=Data_Stores25!H2", int(ik.Target), N, "Ikeja's target."),
              ("A range", "=SUM(Data_Stores25!E2:E13)", ST.Sales.sum(), N, ""),
              ("A table column (same cells)", "=SUM(tStores25[Sales])", ST.Sales.sum(), N, "")])
    m.turn()
    m.ex("What number is in Data_Stores25!G4?", "=Data_Stores25!G4", int(ST.Orders[2]), "int", short="Data_Stores25!G4")
    m.ex("Total all 12 targets with a range reference.", "=SUM(Data_Stores25!H2:H13)", int(ST.Target.sum()), N, short="Total target")

    # ---------------------------------------------------------------- 1.7.7
    lines = IK.head(4)[["Sub_Category", "Sales"]].reset_index(drop=True)
    m.topic("1.7.7", "Relative, absolute and mixed references (F4)",
            "When copied, a relative reference (C5) moves with the formula; an absolute reference ($G$5) stays fixed. Press F4 while the cursor is on a reference to cycle C5 › $C$5 › C$5 › $C5.")
    m.steps(["The VAT rate sits in one yellow cell.",
             "In the first white cell type = then click the Sales cell beside it, type *(1+ then click the VAT cell and press F4 to make it absolute, type ) and press Enter.",
             "Copy the formula down the other three white cells (double-click the fill handle).",
             "The check confirms the result AND that the rate is locked with $."])
    r0 = m.r
    m.ws.write(r0, 6, "VAT rate", F["lab"])
    m.ws.write_number(r0, 7, 0.075, b.nf("pct", style="input"))
    vat = A(r0, 7, True)
    m.ws.write(r0, 1, "Item", F["hdr"])
    m.ws.write(r0, 2, "Sales", F["hdr"])
    m.ws.write(r0, 3, "Incl. VAT (you)", F["hdr"])
    m.ws.merge_range(r0, 4, r0, 5, "Check", F["hdr"])
    for i, (nm, sv) in enumerate(zip(lines.Sub_Category, lines.Sales)):
        r = r0 + 1 + i
        m.ws.write(r, 1, nm, F["cell"])
        m.ws.write_number(r, 2, int(sv), b.nf(N))
        m.ws.write_blank(r, 3, None, b.nf(N, style="plain", border=2, border_color="#E0A526"))
        d, c = A(r, 3), A(r, 2)
        m.ws.merge_range(r, 4, r, 5, "", m.X["chk"])
        m.ws.write_formula(r, 4, f'=IF({d}="","",IF(ABS({d}-{c}*(1+{vat}))>0.01,"✗ Wrong result",IF(ISERROR(SEARCH("$",FORMULATEXT({d}))),"Lock the rate with $ (F4)","✓ Correct, rate locked")))',
                           m.X["chk"], "")
    m.r = r0 + 6
    m.n += 1
    m.key_txt.append((m.n, "1.7.7", f"First white cell: ={A(r0 + 1, 2)}*(1+{vat}); copied down, the row number changes and {vat} stays. "
                                    f"Results: {', '.join(f'₦{v * 1.075:,.2f}' for v in lines.Sales)}."))
    m.ws.write(r0, 9, f"Ex {m.n}", m.X["exid"])
    m.note(f"Ex {m.n}: the four white cells above are the exercise; each has its own check.")

    # ---------------------------------------------------------------- 1.7.8
    y = D.Y.set_index("Year").Sales
    m.topic("1.7.8", "Basic math formulas",
            "Most business measures are simple arithmetic on two columns: cost, margin, attainment, average order value and growth. Write each once, then copy it down.")
    m.steps(["Cost = Sales - Profit.", "Margin = Profit ÷ Sales.", "Attainment = Sales ÷ Target.", "Average order value = Sales ÷ Orders.",
             "Growth = (this year - last year) ÷ last year."])
    m.worked([("Ikeja cost", "=Data_Stores25!E2-Data_Stores25!F2", int(ik.Sales - ik.Profit), N, ""),
              ("Ikeja margin", "=Data_Stores25!F2/Data_Stores25!E2", ik.Margin, "pct", ""),
              ("Ikeja attainment", "=Data_Stores25!E2/Data_Stores25!H2", ik.Attainment, "pct", ""),
              ("Ikeja average order value", "=Data_Stores25!E2/Data_Stores25!G2", ik.Sales / ik.Orders, N, ""),
              ("Chain growth 2025 vs 2024", "=(SUMIFS(tYears[Sales],tYears[Year],2025)-SUMIFS(tYears[Sales],tYears[Year],2024))/SUMIFS(tYears[Sales],tYears[Year],2024)",
               (y[2025] - y[2024]) / y[2024], "pct", "tYears is on sheet 1.2.")])
    m.turn()
    m.ex("What is Lagos - Lekki's average order value (row 3)?", "=Data_Stores25!E3/Data_Stores25!G3", lk.Sales / lk.Orders, N, tol=1, short="Lekki AOV")
    m.ex("What was chain growth from 2023 to 2024?", "=(SUMIFS(tYears[Sales],tYears[Year],2024)-SUMIFS(tYears[Sales],tYears[Year],2023))/SUMIFS(tYears[Sales],tYears[Year],2023)",
         (y[2024] - y[2023]) / y[2023], "pct", short="Growth 2024 vs 2023")

    # ---------------------------------------------------------------- 1.7.9
    m.topic("1.7.9", "Core built-in functions",
            "Excel has over 450 functions in categories on the Formulas tab. A small core covers most analyst work: maths, statistics, text, date, logical (lookups follow in Week 2).")
    m.steps(["Maths: SUMIF(range, criteria, sum_range).", "Statistical: COUNTIF(range, criteria).", "Text: PROPER, TRIM, LEFT, LEN.",
             "Date: YEAR, MONTH, TODAY.", "Logical: IF(test, value_if_true, value_if_false)."])
    m.worked([("SUMIF: South-West sales", '=SUMIF(tStores25[Zone],"South-West",tStores25[Sales])', ST[ST.Zone == "South-West"].Sales.sum(), N, "Ikeja + Lekki + Ibadan."),
              ("COUNTIF: stores with a loss", '=COUNTIF(tStores25[Profit],"<0")', int((ST.Profit < 0).sum()), "int", ""),
              ("PROPER", '=PROPER("LAGOS - IKEJA")', "Lagos - Ikeja", None, ""),
              ("TRIM", '=TRIM(" North-Central")', "North-Central", None, "Removes the leading space."),
              ("YEAR", "=YEAR(DATE(2025,1,1))", 2025, "gen", ""),
              ("IF: Abuja - Wuse (row 8)", '=IF(Data_Stores25!F8<0,"Loss","Profit")', "Loss", None, "")])
    m.turn()
    m.ex("What were total North-Central sales? (SUMIF)", '=SUMIF(tStores25[Zone],"North-Central",tStores25[Sales])', ST[ST.Zone == "North-Central"].Sales.sum(), N,
         short="North-Central sales")
    m.ex("How many stores are in the South-South zone? (COUNTIF)", '=COUNTIF(tStores25[Zone],"South-South")', int((ST.Zone == "South-South").sum()), "int",
         short="South-South stores")
    m.ex_text('What does =LEFT("SM-2025-144463",2) return?', "SM", formula='=LEFT("SM-2025-144463",2)', short="LEFT of an Order_ID")
    m.ex_text('What does =IF(Data_Stores25!F6<0,"Loss","Profit") return for Benin City?', "Loss", formula='=IF(Data_Stores25!F6<0,"Loss","Profit")',
              short="IF for Benin City")

    # ---------------------------------------------------------------- 1.7.10
    m.topic("1.7.10", "Using Insert Function (fx)",
            "If you don't know a function's name or its arguments, let Excel guide you: the fx button opens a searchable list and a form for the arguments.")
    m.steps(["Click the white cell.", "Click fx beside the formula bar (or press Shift+F3).",
             "Search 'count cells that meet a condition' › COUNTIF › OK.",
             "Range: tStores25[Attainment]; Criteria: >=1 › read the preview › OK."])
    m.turn()
    m.ex("Use fx to insert COUNTIF counting the stores with Attainment of 1 or more. What is the result?", '=COUNTIF(tStores25[Attainment],">=1")',
         int((ST.Attainment >= 1).sum()), "int", short="COUNTIF via fx")
    m.ex_text("Which keys open Insert Function? (Type it as Shift+F3)", "Shift+F3", alts=("shift f3",))

    # ---------------------------------------------------------------- 1.7.11
    m.topic("1.7.11", "AutoSum and the six essential functions",
            "AutoSum (Alt+=) writes a SUM for the numbers above or to the left; its drop-down offers the other everyday functions: SUM, AVERAGE, COUNT, COUNTA, MAX, MIN.")
    m.steps(["Click under a column of numbers and press Alt+=; check the range Excel proposes; Enter.",
             "Use the AutoSum drop-down for Average, Count Numbers, Max and Min.",
             "COUNT counts numbers; COUNTA counts every filled cell."])
    m.worked([("SUM", "=SUM(tStores25[Sales])", ST.Sales.sum(), N, ""), ("AVERAGE", "=AVERAGE(tStores25[Sales])", ST.Sales.mean(), N, ""),
              ("COUNT", "=COUNT(tStores25[Sales])", 12, "int", ""), ("COUNTA", "=COUNTA(tStores25[Store_Name])", 12, "int", "Counts text too."),
              ("MAX", "=MAX(tStores25[Sales])", ST.Sales.max(), N, ""), ("MIN", "=MIN(tStores25[Sales])", ST.Sales.min(), N, "")])
    m.turn()
    m.ex("What is the largest store Profit?", "=MAX(tStores25[Profit])", ST.Profit.max(), N, short="Largest Profit")
    m.ex("What is the smallest number of Orders at any store?", "=MIN(tStores25[Orders])", int(ST.Orders.min()), "int", short="Fewest Orders")

    # ---------------------------------------------------------------- 1.7.12
    m.topic("1.7.12", "Understanding formula errors",
            "An error value tells you what went wrong. Read it before you change anything; use IFERROR only for errors you expect.")
    m.steps(["#DIV/0!: dividing by zero or a blank.", "#N/A: a lookup found no match.", "#NAME?: Excel doesn't recognise a name (often a typo).",
             "#VALUE!: wrong type, e.g. text in maths.", "#REF!: the referenced cell was deleted.", "####: the column is too narrow (not an error)."])
    m.worked([("IFERROR catches an expected error", '=IFERROR(1/0,"no sales yet")', "no sales yet", None, "Use it only when you expect the error.")])
    m.turn()
    for fx_, err in [("=1/0", "#DIV/0!"), ("=SUMM(E2:E13)", "#NAME?"), ('="abc"*2', "#VALUE!")]:
        m.ex_text(f"Which error does {fx_} show?", err)

    # ---------------------------------------------------------------- 1.7.13
    six = ST.iloc[:6][["Store_Name", "Sales", "Target"]].reset_index(drop=True)
    m.topic("1.7.13", "Copying and extending formulas across data",
            "Write a formula once in the first row, then extend it. Relative references adjust for each row.")
    m.steps(["In the first white cell type =C?/D? using that row's Sales and Target (click the cells).",
             "Select it and the five white cells below, then press Ctrl+D (fill down). Or double-click the fill handle.",
             "Each check confirms the row's result and that it is a formula, not a typed number."])
    r0 = m.r
    for j, h in enumerate(["Store_Name", "Sales", "Target", "Attainment (you)"]):
        m.ws.write(r0, 1 + j, h, F["hdr"])
    m.ws.merge_range(r0, 5, r0, 6, "Check", F["hdr"])
    for i, row in six.iterrows():
        r = r0 + 1 + i
        m.ws.write(r, 1, row.Store_Name, F["cell"])
        m.ws.write_number(r, 2, int(row.Sales), b.nf(N))
        m.ws.write_number(r, 3, int(row.Target), b.nf(N))
        m.ws.write_blank(r, 4, None, b.nf("pct", style="plain", border=2, border_color="#E0A526"))
        e, c, d = A(r, 4), A(r, 2), A(r, 3)
        m.ws.merge_range(r, 5, r, 6, "", m.X["chk"])
        m.ws.write_formula(r, 5, f'=IF({e}="","",IF(ABS({e}-{c}/{d})>0.00001,"✗",IF(ISFORMULA({e}),"✓","✗ typed: use a formula")))', m.X["chk"], "")
    m.r = r0 + 8
    m.n += 1
    m.key_txt.append((m.n, "1.7.13", f"First white cell: ={A(r0 + 1, 2)}/{A(r0 + 1, 3)}, filled down. Results: "
                                     + ", ".join(f"{v:.1%}" for v in six.Sales / six.Target) + "."))
    m.ws.write(r0, 8, f"Ex {m.n}", m.X["exid"])

    # ---------------------------------------------------------------- 1.7.14
    dept = D.DEPT
    m.topic("1.7.14", "Practical multi-sector calculations",
            "The same few patterns (ratio of totals, average, rate) answer questions in every sector.")
    m.steps(["Retail margin: SUM(profit) ÷ SUM(sales).", "Rate: COUNTIF(column, \"Yes\") ÷ COUNTA(ID column).",
             "Average from a summary table: weight each group by its size (SUMPRODUCT, Week 2).", "Ratio of totals: total ÷ count."])
    m.worked([("Retail: FY2025 margin", "=SUM(tStores25[Profit])/SUM(tStores25[Sales])", ST.Profit.sum() / ST.Sales.sum(), "pct", ""),
              ("University: at-risk rate", '=COUNTIF(tStudents[At_Risk],"Yes")/COUNTA(tStudents[Student_ID])', (STU.At_Risk == "Yes").mean(), "pct", ""),
              ("Hospital: average length of stay 2023, days", "=SUMPRODUCT(tDept23[Admissions],tDept23[Avg_LOS])/SUM(tDept23[Admissions])",
               float((dept.Admissions * dept.Avg_LOS).sum() / dept.Admissions.sum()), "dec1", ""),
              ("Hotel: cancellation rate 2023-2024", "=SUM(tHotelSeg[Cancelled])/SUM(tHotelSeg[Bookings])", D.HSEG.Cancelled.sum() / D.HSEG.Bookings.sum(), "pct", ""),
              ("E-commerce: average order value", "=SUM(tPayMix[Revenue])/SUM(tPayMix[Orders])", D.PAY.Revenue.sum() / D.PAY.Orders.sum(), N, "")])
    m.turn()
    m.ex("What is the average exam score in tStudents?", "=AVERAGE(tStudents[Exam_Score])", float(STU.Exam_Score.mean()), "dec2", tol=0.006,
         short="Average exam score")
    m.ex("What share of students got grade A?", '=COUNTIF(tStudents[Grade],"A")/COUNTA(tStudents[Student_ID])', (STU.Grade == "A").mean(), "pct",
         short="Grade A share")
    m.finish()


# =====================================================================================================
def m18(b, D):
    ST, SA = D.ST, D.SALES
    s25 = D.s25
    F = b.F
    jan = s25[(s25.Store_ID == "ST01") & (s25.Order_Date.dt.month == 1)].sort_values(["Order_Date", "Row_ID"])
    jan = jan[["Order_ID", "Order_Date", "Customer_Name", "Category", "Sub_Category", "Sales", "Quantity", "Discount", "Profit"]].reset_index(drop=True)
    njan = len(jan)
    prac = pd.concat([jan.iloc[:5], jan.iloc[[4]], jan.iloc[5:12], jan.iloc[[11]], jan.iloc[12:]]).reset_index(drop=True)
    m = ModuleSheet(b, "1.8 Excel Tables", "1.8  Organising Data with Excel Tables",
                    "Tutor walk-through after Module 1.8: datasets, converting ranges to tables, structured references, sorting, filtering, duplicates and the Total Row",
                    "1.8.1 - 1.8.12")
    m.dataset([("tSales", "Sheet Data_Sales (starts in A1)", f"{len(SA):,} SuperMart sales lines, 2022-2025: one row per product line."),
               ("tStores25", "Sheet Data_Stores25 (starts in A1)", "12 store rows, used for structured references."),
               ("tDirtySales", "Sheet Data_DirtySales (starts in A1)", "The raw FY2025 export, used for Remove Duplicates (on a copy)."),
               ("Ikeja January practice range", "On this sheet, below", f"Lagos - Ikeja's {njan} lines from January 2025 plus 2 repeated rows. "
                                                                       "It is a plain range, NOT a table: you convert it in 1.8.4 and use it for 1.8.4 - 1.8.12.")])
    m.h("Ikeja January practice range (plain range: convert it to a table named tIkejaJan in 1.8.4)")
    m.r += 1                                   # keep a blank row so Ctrl+T does not pick up the heading
    top = m.s.frame(prac, fmts={"Order_Date": "date", "Sales": N, "Quantity": "int", "Discount": "pct0", "Profit": N})
    last = top + len(prac) - 1
    col = {c: chr(ord("B") + i) for i, c in enumerate(prac.columns)}
    R = {c: f"${col[c]}${top + 1}:${col[c]}${last + 1}" for c in prac.columns}
    m.ws.set_column(1, 1, 16)
    m.ws.set_column(3, 3, 17)
    m.ws.set_column(4, 5, 14)

    # ---------------------------------------------------------------- 1.8.1
    m.topic("1.8.1", "What is a dataset?",
            "A dataset is a structured collection of related data, usually a table in which each row is one record and each column one variable.")
    m.steps(["Open Data_Sales: one header row, then one row per sales line.", "ROWS counts records; COLUMNS counts variables."])
    m.worked([("Records in tSales", "=ROWS(tSales)", len(SA), "int", ""), ("Variables in tSales", "=COLUMNS(tSales)", SA.shape[1], "int", "")])
    m.turn()
    m.ex("How many records are in tStudents?", "=ROWS(tStudents)", len(D.STU), "int", short="Records in tStudents")
    m.ex("How many variables (columns) does tStores25 have?", "=COLUMNS(tStores25)", 10, "int", short="Columns in tStores25")

    # ---------------------------------------------------------------- 1.8.2
    m.topic("1.8.2", "Rows as records, columns as variables",
            "Read across a row to learn everything about one record; read down a column to see one variable for every record. Statistics calls them observations and variables; databases call them records and fields.")
    m.steps(["On Data_Sales, row 4 is the record with Row_ID 3.", "Column K (Profit) is one variable for all 9,994 records."])
    m.turn()
    m.ex("What is the Profit of the record with Row_ID 3 (row 4 of Data_Sales)?", "=Data_Sales!K4", int(SA.Profit[2]), N, short="Profit of Row_ID 3")
    m.ex_text("In statistics, a column is called a …", "Variable", alts=("variables",))

    # ---------------------------------------------------------------- 1.8.3
    m.topic("1.8.3", "Normal range vs Excel table",
            "A range is any block of cells. An Excel table is a range turned into a named, structured object that Excel manages: headers, styles, filters, and it grows with the data.")
    m.steps(["Range formula: =SUM(Data_Sales!H2:H9995): fixed addresses that miss new rows.", "Table formula: =SUM(tSales[Sales]): readable, and includes new rows."])
    m.worked([("Range reference", "=SUM(Data_Sales!H2:H9995)", SA.Sales.sum(), N, ""), ("Structured reference", "=SUM(tSales[Sales])", SA.Sales.sum(), N, "Same today; only this one grows.")])
    m.turn()
    m.ex("Total Profit in tSales with a structured reference.", "=SUM(tSales[Profit])", SA.Profit.sum(), N, short="Total Profit in tSales")

    # ---------------------------------------------------------------- 1.8.4
    m.topic("1.8.4", "Converting a range to a table with Ctrl+T",
            "Click any cell inside the data and press Ctrl+T. Excel selects the whole block, provided it has no blank rows or columns.")
    m.steps(["Click any cell in the Ikeja January practice range.", "Press Ctrl+T; check the range covers every row and column; keep 'My table has headers' ticked; OK.",
             "Table Design › Table Name: type tIkejaJan and press Enter."])
    m.turn()
    m.ex_formula_live("Convert the practice range and name it tIkejaJan. The check looks for the table.",
                      '=IFERROR(IF(ROWS(INDIRECT("tIkejaJan"))>0,"✓ Table tIkejaJan found: "&ROWS(INDIRECT("tIkejaJan"))&" rows"),"Not yet: convert the range and name it tIkejaJan")',
                      "Not yet: convert the range and name it tIkejaJan", key=f"tIkejaJan with {len(prac)} rows (before duplicates are removed).")

    # ---------------------------------------------------------------- 1.8.5
    m.topic("1.8.5", "Creating and formatting Excel tables",
            "A new table gets a style straight away. Table Design lets you change the style and turn parts on or off without touching the cells.")
    m.steps(["Click inside tIkejaJan › Table Design › Table Styles: pick a light style.", "Tick or untick Banded Rows, First Column, Filter Button.",
             "Format whole columns: select a column of the table (click its header edge) › Ctrl+1."])
    m.do([("Apply a light table style and turn on First Column.", "Order_ID values are bold."),
          ("Format the Discount column of tIkejaJan as 0%.", "0.2 shows as 20%.")])

    # ---------------------------------------------------------------- 1.8.6
    m.topic("1.8.6", "Table headers and structured references",
            "A structured reference refers to a table or column by name instead of by address, e.g. tStores25[Sales] for the whole Sales column, [@Profit] for this row's value.")
    m.steps(["Type =SUM(tStores25[ and pick the column from the list.", "Inside a table, [@Column] means this row's value.", "tStores25[[#Totals],[Sales]] refers to the Total Row (when it is on)."])
    m.worked([("=SUM(tStores25[Sales])", "=SUM(tStores25[Sales])", ST.Sales.sum(), N, ""),
              ('=COUNTIF(tStores25[Attainment],">=1")', '=COUNTIF(tStores25[Attainment],">=1")', int((ST.Attainment >= 1).sum()), "int", "")])
    m.turn()
    m.ex("Write =SUM(tStores25[Profit]). What is it?", "=SUM(tStores25[Profit])", ST.Profit.sum(), N, short="Sum of tStores25[Profit]")
    m.ex("How many stores are in the South-East zone? Use a structured reference.", '=COUNTIF(tStores25[Zone],"South-East")', int((ST.Zone == "South-East").sum()),
         "int", short="South-East stores")

    # ---------------------------------------------------------------- 1.8.7
    big = prac.Customer_Name[prac.Sales.idxmax()]
    furn = prac[prac.Category == "Furniture"]
    m.topic("1.8.7", "Sorting data",
            "Sorting reorders whole rows by one or more columns. In a table, every column moves together, so records stay intact.")
    m.steps(["Click the Sales header arrow in tIkejaJan › Sort Largest to Smallest.",
             "Several levels: Data › Sort (Alt, A, S, S) › Category A to Z › Add Level › Sales Largest to Smallest › OK."])
    m.turn()
    m.ex_text("After sorting by Sales (largest first), which customer is in the first row?", big,
              formula=f"=INDEX({R['Customer_Name']},MATCH(MAX({R['Sales']}),{R['Sales']},0))", short="Top customer by Sales")
    m.ex_text("After the two-level sort (Category A-Z, then Sales largest first), which customer is in the first row?",
              furn.Customer_Name[furn.Sales.idxmax()])

    # ---------------------------------------------------------------- 1.8.8
    m.topic("1.8.8", "Filtering data",
            "A filter hides rows you don't need, without deleting them. Each column arrow offers text, number or date filters.")
    m.steps(["Category arrow › untick Select All › tick Technology › OK.", "The status bar shows 'x of y records found'.", "Clear the filter: Data › Clear (or the arrow › Clear Filter)."])
    m.turn()
    m.ex("Filter Category = Technology. How many rows are shown?", f'=COUNTIF({R["Category"]},"Technology")', int((prac.Category == "Technology").sum()), "int",
         short="Technology rows")
    m.ex("Clear it, then Number Filters › Less Than › 0 on Profit. How many rows are shown?", f'=COUNTIF({R["Profit"]},"<0")', int((prac.Profit < 0).sum()), "int",
         short="Rows with a loss")

    # ---------------------------------------------------------------- 1.8.9
    m.topic("1.8.9", "Removing duplicates",
            "Remove Duplicates deletes rows whose chosen columns repeat an earlier row. Decide first what makes a row unique, and work on a copy.")
    m.steps(["Clear any filters. Click inside tIkejaJan › Data › Remove Duplicates.", "Keep every column ticked (a duplicate = the same in every column) › OK.",
             "Excel reports how many duplicates it removed and how many unique rows remain.",
             "For tDirtySales: right-click the Data_DirtySales tab › Move or Copy › Create a copy, then Remove Duplicates on the copy."])
    m.turn()
    m.ex_text("How many duplicate rows does Remove Duplicates remove from tIkejaJan?", "2")
    m.ex_text("On a copy of Data_DirtySales, Remove Duplicates with all columns ticked. How many duplicate values does Excel report removing?", "74")
    m.note("Why 74? 60 repeated records plus 14 of the 15 empty rows (empty rows are identical to each other, so Excel keeps one and removes 14). "
           "Delete the one remaining empty row: 3,387 - 74 - 1 = 3,312 rows, matching the clean FY2025 file.")

    # ---------------------------------------------------------------- 1.8.10
    m.topic("1.8.10", "Auto-expansion and calculated columns",
            "Type in the row under a table and the table grows to include it. Type a formula in an empty column next to it and the formula fills every row: a calculated column.")
    m.steps(["Type Margin in the first empty column to the right of tIkejaJan's headers.", "In its first row type =[@Profit]/[@Sales] and press Enter: every row fills.",
             "Format the column as %."])
    m.turn()
    m.ex_formula_live("Add the Margin calculated column to tIkejaJan. The check counts its filled rows.",
                      '=IFERROR(IF(COUNT(INDIRECT("tIkejaJan[Margin]"))=ROWS(INDIRECT("tIkejaJan")),"✓ Margin fills every row","Margin column incomplete"),"Not yet: add a Margin column to tIkejaJan")',
                      "Not yet: add a Margin column to tIkejaJan", key="Margin = [@Profit]/[@Sales] in every row of tIkejaJan.")

    # ---------------------------------------------------------------- 1.8.11
    tech = jan[jan.Category == "Technology"].Sales.sum()
    m.topic("1.8.11", "Table design, Total Row and naming tables",
            "Give every table a meaningful name (t + a clear noun). The Total Row adds quick totals that respect filters: it uses SUBTOTAL.")
    m.steps(["Table Design › tick Total Row (or Ctrl+Shift+T).", "Click the Total Row cell under Sales › choose Sum.", "Filter Category = Technology: the total changes to the visible rows only.",
             "Formulas › Name Manager (Ctrl+F3) lists every table and name."])
    m.turn()
    m.ex("With the Total Row on and Category filtered to Technology, what does the Sales total show?", f'=SUMIF({R["Category"]},"Technology",{R["Sales"]})', int(tech), N,
         short="Technology Sales total")
    m.ex_text("Which function does the Total Row write?", "SUBTOTAL")

    # ---------------------------------------------------------------- 1.8.12
    m.topic("1.8.12", "Why Excel tables are better than normal ranges",
            "Tables remove a whole class of mistakes: rows missed by formulas, formulas copied to only some rows, unreadable references and records scrambled by sorting one column.")
    m.steps(["Write =SUM(tIkejaJan[Sales]) somewhere on this sheet.", "Add a row under tIkejaJan with Sales 100000.", "Watch the SUM grow by 100,000. A fixed range like B5:B36 would not."])
    m.turn()
    m.ex_text("Does =SUM(C10:C20) include a value typed into C21 later? (yes or no)", "no")
    m.do([("Add a new row to tIkejaJan and check that Margin fills in by itself.", "The new row has a Margin value without copying anything.")])
    m.finish()


# =====================================================================================================
def m19(b, D):
    ST, DI = D.ST, D.DIRTY
    s25 = D.s25
    F = b.F
    m = ModuleSheet(b, "1.9 Tidy Data Laws", "1.9  Twelve Tidy Data Laws",
                    "Tutor walk-through after Module 1.9: recognise and fix each breach of the twelve laws", "1.9.1 - 1.9.12")
    m.dataset([("tDirtySales", "Sheet Data_DirtySales (starts in A1)", "The raw FY2025 export: it breaks several laws on purpose."),
               ("tStores25", "Sheet Data_Stores25 (starts in A1)", "A tidy table, used as the model."),
               ("Untidy practice ranges", "On this sheet", "Small ranges, each breaking one law, for you to fix in the white cells.")])

    # ---------------------------------------------------------------- 1.9.1
    m.topic("1.9.1", "The tidy data foundation",
            "Tidy data: each variable forms a column, each observation a row, and each type of observational unit a table (Hadley Wickham, 2014). Tables, PivotTables, Power Query, SQL and Power BI all expect this shape.")
    m.steps(["Ask of any sheet: is there one header row?", "Is each row one record, each column one variable, each cell one value?", "If not, the twelve laws below say how to fix it."])
    m.turn()
    for lay, ans in [("Months as columns: Jan, Feb, Mar …", "Untidy"), ("One row per sale with Order_Date, Store_Name and Sales columns", "Tidy"),
                     ("A 'Lagos total' row in the middle of the sales lines", "Untidy")]:
        m.ex_text(f"Tidy or Untidy? {lay}", ans)

    # ---------------------------------------------------------------- 1.9.2
    sub = pd.DataFrame({"Store_Name": ["Lagos - Ikeja", "Lagos - Lekki", "Lagos total", "Abuja - Wuse", "Abuja - Garki"],
                        "Sales": [int(ST.Sales[0]), int(ST.Sales[1]), int(ST.Sales[0] + ST.Sales[1]), int(ST.Sales[6]), int(ST.Sales[7])]})
    m.topic("1.9.2", "Law 1: one record per row",
            "Each row describes exactly one record, and every record appears once. Duplicates, subtotal rows and rows holding two records all break counts and totals.")
    m.steps(["Look at the range below: the 'Lagos total' row is a subtotal inside the data.", "SUM counts Ikeja and Lekki twice.",
             "Fix: delete subtotal rows from data; put totals in a report (or the Total Row)."])
    top = m.s.frame(sub, fmts={"Sales": N})
    rs, rn = f"C{top + 1}:C{top + 5}", f"B{top + 1}:B{top + 5}"
    m.worked([("SUM including the subtotal row (wrong)", f"=SUM({rs})", int(sub.Sales.sum()), N, "Double-counts Lagos.")])
    m.turn()
    m.ex("What is the correct total of the four stores? (Hint: SUMIF with \"<>Lagos total\".)", f'=SUMIF({rn},"<>Lagos total",{rs})', int(sub.Sales.sum() - sub.Sales[2]), N,
         short="Correct four-store total")

    # ---------------------------------------------------------------- 1.9.3
    stores3 = ["Lagos - Ikeja", "Lagos - Lekki", "Port Harcourt - GRA"]
    q = s25[s25.Store_Name.isin(stores3)].groupby(["Store_Name", s25.Order_Date.dt.quarter]).Sales.sum().unstack()
    q = q.reindex(stores3)
    m.topic("1.9.3", "Law 2: one variable per column",
            "Each column holds one variable in one form. Quarters spread across columns (Q1, Q2, …) is the most common breach: it looks like a report but cannot be filtered or summed by quarter.")
    m.steps(["The wide range below has one column per quarter (untidy).", "The long range beside it has one Quarter column (tidy): Store, Quarter and Sales.",
             "Fill the white Sales cells from the wide range (copy, or type =the matching wide cell).", "Power Query's Unpivot does this in one click (Week 2)."])
    r0 = m.r
    for j, h in enumerate(["Store (wide, untidy)", "Q1", "Q2", "Q3", "Q4"]):
        m.ws.write(r0, 1 + j, h, F["hdr"])
    for i, s in enumerate(stores3):
        m.ws.write(r0 + 1 + i, 1, s, F["cell"])
        for k in range(4):
            m.ws.write_number(r0 + 1 + i, 2 + k, int(q.iloc[i, k]), b.nf(N))
    m.ws.write(r0, 7, "Store", F["hdr"])
    m.ws.write(r0, 8, "Quarter", F["hdr"])
    m.ws.write(r0, 9, "Sales (you)", F["hdr"])
    m.ws.write(r0, 10, "Check", F["hdr"])
    wide = f"$C${r0 + 2}:$F${r0 + 4}"
    wst = f"$B${r0 + 2}:$B${r0 + 4}"
    whd = f"$C${r0 + 1}:$F${r0 + 1}"
    k = 0
    for s in stores3:
        for qq in range(1, 5):
            r = r0 + 1 + k
            m.ws.write(r, 7, s, F["cell"])
            m.ws.write(r, 8, f"Q{qq}", F["cell"])
            m.ws.write_blank(r, 9, None, b.nf(N, style="plain", border=2, border_color="#E0A526"))
            j_, h_, i_ = A(r, 9), A(r, 7), A(r, 8)
            m.ws.write_formula(r, 10, f'=IF({j_}="","",IF({j_}=INDEX({wide},MATCH({h_},{wst},0),MATCH({i_},{whd},0)),"✓","✗"))', m.X["chk"], "")
            k += 1
    m.ws.set_column(7, 7, 18)
    longr = f"J{r0 + 2}:J{r0 + 13}"
    m.r = r0 + 14
    m.turn()
    m.ex_formula_live("Fill all 12 white Sales cells of the long range. The check compares totals.",
                      f'=IF(COUNT({longr})<12,COUNT({longr})&" of 12 filled",IF(SUM({longr})=SUM({wide}),"✓ Long range complete and matches","✗ Totals differ: check the rows marked ✗"))',
                      "0 of 12 filled", key=f"Total of the 12 cells = ₦{int(q.values.sum()):,}.")

    # ---------------------------------------------------------------- 1.9.4
    m.topic("1.9.4", "Law 3: atomic values (one value per cell)",
            "Each cell holds a single value that cannot sensibly be split further. Combined values ('Abuja, FCT', '₦1,399,930' as text) are hard to filter, sort and count.")
    m.steps(["Split text: Data › Text to Columns › Delimited, Flash Fill, or a formula.",
             "Area after ' - ': =TRIM(MID(name,FIND(\"-\",name)+1,99)).",
             "Text money to a number: =VALUE(SUBSTITUTE(SUBSTITUTE(text,\"₦\",\"\"),\",\",\"\"))."])
    m.worked([("Area of 'Lagos - Ikeja'", '=TRIM(MID("Lagos - Ikeja",FIND("-","Lagos - Ikeja")+1,99))', "Ikeja", None, ""),
              ("'₦1,399,930' as a number", '=VALUE(SUBSTITUTE(SUBSTITUTE("₦1,399,930","₦",""),",",""))', 1399930, N, "Now SUM can add it.")])
    m.turn()
    m.ex_text('What does =TRIM(MID("Enugu - Independence Layout",FIND("-","Enugu - Independence Layout")+1,99)) return?', "Independence Layout",
              formula='=TRIM(MID("Enugu - Independence Layout",FIND("-","Enugu - Independence Layout")+1,99))', short="Area of Enugu store")
    m.ex('Convert the text "₦239,240" to a number with VALUE and SUBSTITUTE. What number do you get?', '=VALUE(SUBSTITUTE(SUBSTITUTE("₦239,240","₦",""),",",""))',
         239240, "int", short="₦239,240 as a number")

    # ---------------------------------------------------------------- 1.9.5
    m.topic("1.9.5", "Law 4: one header row",
            "Data starts with exactly one row of unique column names, directly above the first record. Titles above the data or two stacked header rows confuse tables, PivotTables and Power Query.")
    m.steps(["Below: two stacked header rows (Q1 above Sales and Profit).", "Combine them into one row of unique names in the white cells: Store, Q1_Sales, Q1_Profit.",
             "Short names, no spaces (use _), no blanks."])
    r0 = m.r
    for j, (a, b_) in enumerate([("", "Store"), ("Q1", "Sales"), ("", "Profit")]):
        m.ws.write(r0, 1 + j, a, F["hdr"])
        m.ws.write(r0 + 1, 1 + j, b_, F["hdr"])
        m.ws.write_blank(r0 + 3, 1 + j, None, m.X["in"])
    m.ws.write(r0 + 2, 1, "One header row (you):", F["note"])
    m.r = r0 + 5
    c1, c2, c3 = A(r0 + 3, 1), A(r0 + 3, 2), A(r0 + 3, 3)
    m.turn()
    m.ex_formula_live("Type the single header row into the three white cells.",
                      f'=IF(COUNTA({c1}:{c3})<3,"Type all three names",IF(AND(EXACT({c1},"Store"),EXACT({c2},"Q1_Sales"),EXACT({c3},"Q1_Profit")),"✓ One header row of unique names","✗ Use Store, Q1_Sales, Q1_Profit"))',
                      "Type all three names", key="Store | Q1_Sales | Q1_Profit")

    # ---------------------------------------------------------------- 1.9.6
    m.topic("1.9.6", "Law 5: no spacer rows or blank columns",
            "A dataset is one continuous block. Excel treats a blank row as the end of the data: Ctrl+Down, Ctrl+T, sorting and filtering stop at the gap.")
    m.steps(["The range below has one empty row.", "Click its first cell and press Ctrl+T: the proposed range stops at the gap.",
             "Fix: select the empty row › right-click › Delete. For many, F5 › Special › Blanks, then Home › Delete › Delete Sheet Rows."])
    gap = ST.iloc[:6][["Store_Name", "Sales"]].reset_index(drop=True)
    top = m.s.frame(pd.concat([gap.iloc[:3], pd.DataFrame({"Store_Name": [None], "Sales": [None]}), gap.iloc[3:]]).reset_index(drop=True), fmts={"Sales": N})
    rr = f"B{top + 1}:B{top + 7}"
    m.turn()
    m.ex("How many blank cells are there in the Store_Name column of that range?", f"=COUNTBLANK({rr})", 1, "int", short="Blank Store_Name cells")
    m.ex("How many completely empty rows does tDirtySales have?", "=COUNTBLANK(tDirtySales[Order_ID])", int(DI.Order_ID.isna().sum()), "int",
         short="Empty rows in tDirtySales")

    # ---------------------------------------------------------------- 1.9.7
    dates = [("14/05/2025", dt.date(2025, 5, 14), '=DATE(RIGHT({c},4),MID({c},4,2),LEFT({c},2))'),
             ("2025.06.10", dt.date(2025, 6, 10), '=DATE(LEFT({c},4),MID({c},6,2),RIGHT({c},2))'),
             ("23/02/2025", dt.date(2025, 2, 23), None), ("2025.11.18", dt.date(2025, 11, 18), None)]
    m.topic("1.9.7", "Law 6: ISO dates (yyyy-mm-dd)",
            "Dates are stored as real Excel dates, in one column, displayed as yyyy-mm-dd. 04/05/2025 is 4 May in Nigeria and 5 April in the USA; 2025-05-04 means the same everywhere.")
    m.steps(["Test: =ISNUMBER(cell) is TRUE for a real date, FALSE for text.",
             "dd/mm/yyyy text: =DATE(RIGHT(t,4),MID(t,4,2),LEFT(t,2)).", "yyyy.mm.dd text: =DATE(LEFT(t,4),MID(t,6,2),RIGHT(t,2)).",
             "Format the result: Ctrl+1 › Custom › yyyy-mm-dd."])
    r0 = m.r
    for j, h in enumerate(["Text date", "Real date (you)", "Check"]):
        m.ws.write(r0, 1 + j * 2, h, F["hdr"])
        m.ws.write(r0, 2 + j * 2, "", F["hdr"])
    for i, (t, d, fx) in enumerate(dates):
        r = r0 + 1 + i
        m.ws.merge_range(r, 1, r, 2, t, F["cell"])
        c = A(r, 1)
        if fx:
            m.ws.merge_range(r, 3, r, 4, "", b.nf("date", style="result"))
            m.ws.write_formula(r, 3, fx.format(c=c), b.nf("date", style="result"), serial(d))
        else:
            m.ws.merge_range(r, 3, r, 4, "", b.nf("date", style="plain", border=2, border_color="#E0A526"))
        m.ws.merge_range(r, 5, r, 6, "", m.X["chk"])
        m.ws.write_formula(r, 5, f'=IF({A(r, 3)}="","",IF({A(r, 3)}=DATE({d.year},{d.month},{d.day}),"✓ {d.isoformat()}","✗"))', m.X["chk"],
                           f"✓ {d.isoformat()}" if fx else "")
    m.r = r0 + 6
    m.n += 1
    m.key_txt.append((m.n, "1.9.7", "Row 3: =DATE(RIGHT(t,4),MID(t,4,2),LEFT(t,2)) gives 2025-02-23. Row 4: =DATE(LEFT(t,4),MID(t,6,2),RIGHT(t,2)) gives 2025-11-18."))
    m.ws.write(r0, 8, f"Ex {m.n}", m.X["exid"])
    m.note(f"Ex {m.n}: rows 1-2 are worked; convert rows 3-4 in the white cells.")
    m.turn()
    m.ex("How many Order_Date values in tDirtySales are stored as text (not real dates)?", "=COUNTA(tDirtySales[Order_Date])-COUNT(tDirtySales[Order_Date])",
         int(DI.Order_Date.map(lambda v: isinstance(v, str)).sum()), "int", short="Text dates in tDirtySales")

    # ---------------------------------------------------------------- 1.9.8
    m.topic("1.9.8", "Law 7: missing values, blank vs zero",
            "A blank means unknown or not applicable; a zero means measured and nothing. AVERAGE ignores blanks but includes zeros, so typing 0 for 'not yet known' drags averages down.")
    m.steps(["Four scores below; the last is unknown and left blank.", "Read the AVERAGE.", "Type 0 in the white cell and watch the average fall. Delete it again."])
    r0 = m.r
    for j, v in enumerate([70, 80, 90]):
        m.ws.write_number(r0, 1 + j, v, b.nf("int"))
    m.ws.write_blank(r0, 4, None, m.X["in"])
    rg = f"B{r0 + 1}:E{r0 + 1}"
    m.ws.write(r0, 5, "AVERAGE:", F["lab"])
    m.ws.write_formula(r0, 6, f"=AVERAGE({rg})", b.nf("dec1", style="result"), 80)
    m.ws.write(r0, 7, "COUNT:", F["lab"])
    m.ws.write_formula(r0, 8, f"=COUNT({rg})", b.nf("int", style="result"), 3)
    m.r = r0 + 2
    m.turn()
    m.ex("With 0 typed in the white cell, what does AVERAGE show?", "=AVERAGE(70,80,90,0)", 60, "dec1", short="AVERAGE with a zero")
    m.ex_text("A Discount cell is empty. Does that mean 'no discount' or 'not recorded'?", "not recorded", alts=("not known", "unknown"))

    # ---------------------------------------------------------------- 1.9.9
    m.topic("1.9.9", "Law 8: separate raw data from reports",
            "Keep the original data exactly as received, do the work in a separate layer, and build summaries on their own sheets. If something goes wrong, start again from the raw data.")
    m.steps(["Raw: the import, never edited (Data_DirtySales).", "Clean: a fixed copy, or a Power Query output (tSales is the clean version).", "Report: totals, PivotTables and charts that read from the clean layer."])
    m.turn()
    m.ex_text("Which sheet in this workbook holds the raw export exactly as received?", "Data_DirtySales")
    m.ex_text("Should a PivotTable read from the raw, clean or report layer?", "clean")

    # ---------------------------------------------------------------- 1.9.10
    zn = DI.Zone.dropna()
    m.topic("1.9.10", "Law 9: control your vocabulary",
            "Each category value is written exactly one way, from an agreed list. To Excel, 'South-West' and ' South-West' are different values, so counts and PivotTables split one group into several.")
    m.steps(["COUNTIF finds only the exact text (it ignores case, not spaces).", "TRIM inside SUMPRODUCT counts every variant.",
             "Fix: TRIM and PROPER (or Find & Replace), then prevent new variants with a Data Validation list."])
    m.worked([('COUNTIF(Zone,"South-West")', '=COUNTIF(tDirtySales[Zone],"South-West")', int((zn == "South-West").sum()), "int", "Misses the variants with spaces."),
              ("All South-West rows (TRIM)", '=SUMPRODUCT(--(TRIM(tDirtySales[Zone])="South-West"))', int((zn.str.strip() == "South-West").sum()), "int", "")])
    m.turn()
    m.ex("How many rows are really North-Central (use TRIM)?", '=SUMPRODUCT(--(TRIM(tDirtySales[Zone])="North-Central"))', int((zn.str.strip() == "North-Central").sum()),
         "int", short="All North-Central rows")
    m.ex_text('What does =PROPER(TRIM("  lagos - ikeja")) return?', "Lagos - Ikeja", formula='=PROPER(TRIM("  lagos - ikeja"))', short="PROPER(TRIM())")

    # ---------------------------------------------------------------- 1.9.11
    m.topic("1.9.11", "Law 10: store calculations, not typed results",
            "Derived values are calculated with formulas from the data, never typed as fixed numbers. A typed 103% never changes and nobody can see how it was worked out.")
    m.steps(["The white cell holds a typed attainment for Lagos - Ikeja (1.031).", "Replace it with a formula: =Data_Stores25!E2/Data_Stores25!H2.",
             "The check uses ISFORMULA to tell a formula from a typed number."])
    m.turn()
    r = m.r
    c = A(r, 8)
    m.ex_formula_live("Replace the typed 1.031 in the white cell with a formula.",
                      f'=IF(ISBLANK({c}),"",IF(ISFORMULA({c}),IF(ABS({c}-Data_Stores25!E2/Data_Stores25!H2)<0.00001,"✓ Formula, correct result","✗ Formula, wrong result"),"✗ Still a typed number"))',
                      "✗ Still a typed number", key="=Data_Stores25!E2/Data_Stores25!H2 (103.1%).")
    m.ws.write_number(r, 8, 1.031, m.X["in"])
    m.ex_text("Does a value pasted with Paste Special › Values update when the source changes? (yes or no)", "no")

    # ---------------------------------------------------------------- 1.9.12
    m.topic("1.9.12", "Laws 11 and 12: no merged cells; document everything",
            "Law 11: never merge cells inside data. Law 12: document what each file and column means, its unit, its source and every change made: a data dictionary.")
    m.steps(["Use Center Across Selection for titles instead of merging (1.6.10).",
             "Fill in the data dictionary for tStores25 below: meaning and unit/type of each column.",
             "Keep the dictionary on its own sheet in the workbook, or in a README."])
    dd = [("Store_ID", "Store code, ST01-ST12", "Text"), ("Store_Name", None, None), ("Zone", None, None), ("Store_Manager", None, None),
          ("Sales", "Total FY2025 sales for the store", "₦, number"), ("Profit", None, None), ("Orders", None, None), ("Target", "Annual sales target for 2025", "₦, number"),
          ("Margin", None, None), ("Attainment", "Sales ÷ Target", "%, formula")]
    r0 = m.r
    m.ws.write(r0, 1, "Column", F["hdr"])
    m.ws.merge_range(r0, 2, r0, 7, "Meaning", F["hdr"])
    m.ws.merge_range(r0, 8, r0, 10, "Unit / type", F["hdr"])
    for i, (cn, mean, unit) in enumerate(dd):
        r = r0 + 1 + i
        m.ws.write(r, 1, cn, F["lab"])
        m.ws.merge_range(r, 2, r, 7, mean or "", F["cell"] if mean else m.X["in"])
        m.ws.merge_range(r, 8, r, 10, unit or "", F["cell"] if unit else m.X["in"])
    m.r = r0 + len(dd) + 2
    m.turn()
    m.ex_open("Fill in the white cells of the data dictionary above.",
              "Store_Name: store's city and area (Text). Zone: geopolitical zone (Text, 4 values). Store_Manager: manager's name (Text). "
              "Profit: FY2025 profit, sales minus cost (₦, number, can be negative). Orders: distinct orders in FY2025 (count). Margin: Profit ÷ Sales (%, formula).")
    m.ex_text("Which alignment option should you use instead of merging cells for a title?", "Center Across Selection", alts=("centre across selection",))
    m.finish()


def build(b, D):
    m17(b, D)
    m18(b, D)
    m19(b, D)
