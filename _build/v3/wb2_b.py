"""Week 2 workbook, modules 2.5 - 2.7 (edition 3): cleaning in Excel, transforming in Power Query, choosing the tool."""
from __future__ import annotations

import numpy as np
import pandas as pd

from w1x import ModuleSheet
from wb2_common import M_DIRTY, N, code, lit_grid, serial
from w2_fx import CTRL, DF

DD = "Data_Dirty25!"


def _num(v):
    return isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool) and v == v


# =====================================================================================================
def m25(b, D):
    DR, S = D.DR, D.S
    nb = DR[DR.Order_ID.notna()]
    m = ModuleSheet(b, "2.5 Excel Cleaning", "2.5  Data Cleaning with Excel",
                    "Tutor walk-through after Module 2.5: measure the twelve problems in the raw file, fix each with Excel tools and "
                    "formulas, and prove the result against the POS control totals", "2.5.1 - 2.5.12")
    m.dataset([("tDirty25", "Sheet Data_Dirty25 (starts in A1)", "The raw FY2025 export, 3,387 rows. Cell addresses on the slides are cells of "
                                                                 "this sheet (R23 holds '₦239,240'). Practise on a copy of the sheet."),
               ("tCatMap", "On this sheet, below", "Mapping table: every Category spelling found in the file and its standard value."),
               ("tSales25", "Sheet Data_Sales25 (starts in A1)", "The cleaned result. Your cleaned copy must match its control totals.")])
    variants = ["Furniture", "FURNITURE", "furniture", "Furnitre", "Office Supplies", "Office Supply", "Office-Supplies", "office supplies",
                "Technology", "Technolgy", "Tech", "technology"]
    std = {"fur": "Furniture", "off": "Office Supplies", "tec": "Technology"}
    cmap = pd.DataFrame({"Variant": variants, "Standard": [std[v[:3].lower()] for v in variants]})
    cnt = [sum(1 for c in DR.Category if c == v) for v in variants]
    m.h("tCatMap")
    m.s.table(cmap, "tCatMap", widths={"Variant": 18, "Standard": 18},
              calc={"Cells_in_file": ("=SUMPRODUCT(--EXACT(tDirty25[Category],[@Variant]))", "int", cnt)})

    # ---------------------------------------------------------------- 2.5.1-2
    zsp = sum(1 for v in DR.Zone if isinstance(v, str) and v != " ".join(v.split()))
    neg = sum(1 for v in DR.Quantity if _num(v) and v < 0)
    m.topic("2.5.1-2", "Understanding dirty data, and the twelve common problems",
            "Dirty data is any value you cannot trust for the job. Before fixing anything, measure each problem with one detection "
            "formula over the whole column: the count tells you how big the job is and later proves the fix worked.")
    lit_grid(m, [["1  Duplicates", "7  Bad or mixed dates"], ["2  Missing values", "8  Out-of-range values"],
                 ["3  Inconsistent case and spacing", "9  Combined values in one column"], ["4  Inconsistent categories", "10  Hidden characters"],
                 ["5  Wrong data types", "11  Mixed units"], ["6  Numbers stored as text", "12  Structure: blank rows, extra headers"]],
             header=["Problem", "Problem"])
    m.worked([("Blank spacer rows", "=COUNTBLANK(tDirty25[Order_ID])", DF["blank_rows"], "int", "Problem 12."),
              ("Exact duplicate rows (incl. extra blank rows)", "=ROWS(tDirty25)-ROWS(UNIQUE(tDirty25))", DF["dup_all"], "int",
               "60 real copies + 14 extra blank rows."),
              ("Sales stored as text", "=SUMPRODUCT(--ISTEXT(tDirty25[Sales]))", DF["sales_text"], "int", "Problem 6."),
              ("Placeholder names 'N/A'", '=COUNTIF(tDirty25[Customer_Name],"N/A")', DF["na_cust"], "int", "Problem 2."),
              ("Negative quantities", '=COUNTIF(tDirty25[Quantity],"<0")', neg, "int", "Problem 8."),
              ("Zones with extra spaces", "=SUMPRODUCT(--(LEN(tDirty25[Zone])<>LEN(TRIM(tDirty25[Zone]))))", zsp, "int", "Problem 3."),
              ("Order dates stored as text", "=SUMPRODUCT(--ISTEXT(tDirty25[Order_Date]))", DF["text_dates"], "int", "Problem 7.")])
    m.turn()
    m.ex("How many Discount cells are stored as text?", "=SUMPRODUCT(--ISTEXT(tDirty25[Discount]))", DF["disc_text"], "int",
         short="Discount stored as text")
    m.ex("How many Customer_Name cells contain a line feed, CHAR(10)?",
         "=SUMPRODUCT(--ISNUMBER(SEARCH(CHAR(10),tDirty25[Customer_Name])))", DF["lf"], "int", short="Names with a line feed")
    m.ex_text("Which of the twelve problems is 'Category spelt Furnitre'? Type its number.", "4", alts=("inconsistent categories",))

    # ---------------------------------------------------------------- 2.5.3
    m.topic("2.5.3", "Removing duplicates",
            "Data › Remove Duplicates deletes every repeat of the columns you tick, keeping the first. It cannot be undone after saving "
            "and leaves no record, so copy the sheet first and write down the message Excel shows.")
    m.steps(["Right-click the Data_Dirty25 tab › Move or Copy › tick Create a copy. Work on the copy.",
             "Delete the 15 blank rows first (2.5.10), otherwise Excel counts 14 of them as duplicates.",
             "Click in the table › Table Design › Remove Duplicates (or Data › Remove Duplicates). Keep every column ticked › OK.",
             "Read the message: 60 duplicate values found and removed; 3,312 unique values remain."])
    m.worked([("Rows before (blank rows removed)", "=ROWS(tDirty25)-COUNTBLANK(tDirty25[Order_ID])", DF["after_blank"], "int", ""),
              ("Unique rows", "=ROWS(UNIQUE(tDirty25))-1", DF["after_blank"] - DF["dup_lines"], "int", "Minus 1: the blank row."),
              ("Duplicates removed", "=ROWS(tDirty25)-COUNTBLANK(tDirty25[Order_ID])-(ROWS(UNIQUE(tDirty25))-1)", DF["dup_lines"], "int", "")])
    m.turn()
    m.ex("How many rows should your copy have after removing blank rows and duplicates?", "=ROWS(UNIQUE(tDirty25))-1",
         DF["after_blank"] - DF["dup_lines"], "int", short="Rows after de-duplication")
    m.ex_text("If you remove duplicates BEFORE deleting the blank rows, how many rows does Excel report removed?", str(DF["dup_all"]))

    # ---------------------------------------------------------------- 2.5.4
    m.topic("2.5.4", "Finding and replacing values",
            "Find & Replace (Ctrl+H) changes every match at once. By default it matches part of a cell and ignores case, so replacing "
            "'Tech' would also hit 'Technology'. Select the column first and tick Match entire cell contents.")
    m.steps(["On your copy, select the Category column only.",
             "Ctrl+H › Find what: Tech › Replace with: Technology.",
             "Options › tick Match entire cell contents.",
             "Replace All. Excel reports 8 replacements."])
    m.worked([("Cells a default Replace would change", '=COUNTIF(tDirty25[Category],"*Tech*")', DF["tech_partial"], "int",
               "Partial match, any case: 'Technology' would become 'Technologynology'."),
              ("Cells that are exactly 'Tech'", '=COUNTIF(tDirty25[Category],"Tech")', 8, "int", "With Match entire cell contents.")])
    m.turn()
    m.ex("How many Category cells are exactly 'Technolgy' (case-sensitive)?", '=SUMPRODUCT(--EXACT(tDirty25[Category],"Technolgy"))',
         cnt[variants.index("Technolgy")], "int", short="'Technolgy' cells")
    m.ex_text("Which Find & Replace option stops 'Tech' matching inside 'Technology'?", "Match entire cell contents",
              alts=("match entire cell",))

    # ---------------------------------------------------------------- 2.5.5
    tsv = sum(float(v.replace("₦", "").replace(",", "")) for v in DR.Sales if isinstance(v, str))
    m.topic("2.5.5", "Fixing numbers stored as text",
            "A number typed with ₦ or commas arrives as text: it sits left-aligned, shows a green triangle and SUM ignores it. "
            "Strip the symbols with SUBSTITUTE and convert with VALUE.")
    m.steps(["Insert a helper column next to Sales on your copy.",
             "Type the formula from the worked solution, using the Sales cell of that row, and fill it down.",
             "Copy the helper column › Paste Special › Values over Sales; delete the helper.",
             "Recount: =SUMPRODUCT(--ISTEXT(tDirty25[Sales])) on your copy must be 0."])
    m.worked([("R23 as stored", f"={DD}R23", DR.Sales[21], None, "Text: note the ₦ and the comma."),
              ("R23 converted", f'=VALUE(SUBSTITUTE(SUBSTITUTE({DD}R23,"₦",""),",",""))', 239240, N, "Strip ₦ and commas, then convert."),
              ("T27 '20%' converted", f'=VALUE(SUBSTITUTE({DD}T27,"%",""))/100', 0.2, "pct0", ""),
              ("Sales hidden from SUM", '=SUMPRODUCT(ISTEXT(tDirty25[Sales])*IFERROR(--SUBSTITUTE(SUBSTITUTE(tDirty25[Sales],"₦",""),",",""),0))',
               tsv, N, "234 text values that a plain SUM leaves out.")])
    m.turn()
    t2 = next(i for i, v in enumerate(DR.Sales) if isinstance(v, str) and i > 21)
    m.ex(f"Convert the text in R{t2 + 2} of Data_Dirty25 to a number. What is it?",
         f'=VALUE(SUBSTITUTE(SUBSTITUTE({DD}R{t2 + 2},"₦",""),",",""))', float(DR.Sales[t2].replace("₦", "").replace(",", "")), N,
         short=f"R{t2 + 2} as a number")
    m.ex("How many Sales cells are stored as text?", "=SUMPRODUCT(--ISTEXT(tDirty25[Sales]))", DF["sales_text"], "int", short="Sales stored as text")

    # ---------------------------------------------------------------- 2.5.6
    zd = len({v for v in DR.Zone if v})
    m.topic("2.5.6", "Cleaning text with TRIM and CLEAN",
            "TRIM removes leading and trailing spaces and reduces inner runs to one. CLEAN removes non-printing characters 0-31, such as "
            "the line feed CHAR(10). Neither removes the non-breaking space CHAR(160): SUBSTITUTE it with a normal space first.")
    m.worked([("L47 as stored", f'="["&{DD}L47&"]"', "[" + DR.Zone[45] + "]", None, "Brackets show the hidden space."),
              ("=TRIM(L47)", f"=TRIM({DD}L47)", DR.Zone[45].strip(), None, ""),
              ("Characters in G31", f"=LEN({DD}G31)", len(DR.Customer_Name[29]), "int", "Omolara, CHAR(160) instead of a space, Yusuf, then CHAR(10)."),
              ("CLEAN then LEN", f"=LEN(CLEAN({DD}G31))", len(DR.Customer_Name[29].replace("\n", "")), "int", "Line feed gone; CHAR(160) stays."),
              ("Fully cleaned", f'=TRIM(CLEAN(SUBSTITUTE({DD}G31,CHAR(160)," ")))', "Omolara Yusuf", None, ""),
              ("Distinct Zone values before", '=ROWS(UNIQUE(FILTER(tDirty25[Zone],tDirty25[Zone]<>"")))', zd, "int", ""),
              ("Distinct Zone values after TRIM", '=ROWS(UNIQUE(TRIM(FILTER(tDirty25[Zone],tDirty25[Zone]<>""))))', 4, "int", "")])
    m.turn()
    m.ex("How many Customer_Name cells contain a non-breaking space, CHAR(160)?",
         "=SUMPRODUCT(--ISNUMBER(SEARCH(CHAR(160),tDirty25[Customer_Name])))", DF["nbsp"], "int", short="Names with CHAR(160)")
    m.ex("How many Zone cells have extra spaces?", "=SUMPRODUCT(--(LEN(tDirty25[Zone])<>LEN(TRIM(tDirty25[Zone]))))", zsp, "int",
         short="Zones with extra spaces")

    # ---------------------------------------------------------------- 2.5.7
    nbad = int(len(nb) - sum(1 for v in nb.Category if v in ("Furniture", "Office Supplies", "Technology")))
    m.topic("2.5.7", "Standardising categorical values",
            "One category spelt several ways splits every count. List every variant once in a mapping table (tCatMap) and look each cell "
            "up. PROPER fixes case but not spelling.")
    m.steps(["Get the variants: =UNIQUE(tDirty25[Category]) lists them (it ignores case, so check with EXACT as tCatMap does).",
             "Write the standard value beside each variant: that is tCatMap.",
             "In a helper column: =XLOOKUP(TRIM(O2),tCatMap[Variant],tCatMap[Standard],\"CHECK\") and fill down.",
             "Any 'CHECK' means a new variant: add it to tCatMap."])
    m.worked([("O23 as stored", f"={DD}O23", DR.Category[21], None, ""),
              ("Mapped with tCatMap", f'=XLOOKUP(TRIM({DD}O23),tCatMap[Variant],tCatMap[Standard],"CHECK")', "Furniture", None, ""),
              ("O118 with PROPER", f"=PROPER({DD}O118)", DR.Category[116].title(), None, "Case fixed, spelling not."),
              ("O118 mapped", f'=XLOOKUP(TRIM({DD}O118),tCatMap[Variant],tCatMap[Standard],"CHECK")', "Furniture", None, ""),
              ("Non-standard Category cells", '=ROWS(tDirty25)-COUNTBLANK(tDirty25[Category])-SUMPRODUCT(--EXACT(tDirty25[Category],"Furniture"))'
                                              '-SUMPRODUCT(--EXACT(tDirty25[Category],"Office Supplies"))-SUMPRODUCT(--EXACT(tDirty25[Category],"Technology"))',
               nbad, "int", "The 132 cells the mapping fixes.")])
    m.turn()
    m.ex("How many cells say exactly 'Office Supply'?", "=XLOOKUP(\"Office Supply\",tCatMap[Variant],tCatMap[Cells_in_file])",
         cnt[variants.index("Office Supply")], "int", short="'Office Supply' cells")
    m.ex("How many Category spellings does tCatMap map to 'Technology'?", '=COUNTIF(tCatMap[Standard],"Technology")', 4, "int",
         short="Spellings of Technology")

    # ---------------------------------------------------------------- 2.5.8
    od = [v for v in DR.Order_Date if isinstance(v, str)]
    m.topic("2.5.8", "Fixing incorrect dates",
            "A date stored as text cannot be sorted, grouped by month or subtracted. The file has three text patterns; repair each with "
            "its own formula. DATE with LEFT, MID and RIGHT works whatever the PC's regional settings.")
    m.worked([("dd/mm/yyyy values", '=SUMPRODUCT(--ISNUMBER(SEARCH("/",tDirty25[Order_Date])))', sum("/" in v for v in od), "int", ""),
              ("dd-Mon-yyyy values", '=SUMPRODUCT(ISTEXT(tDirty25[Order_Date])*ISNUMBER(SEARCH("-",tDirty25[Order_Date])))',
               sum("-" in v for v in od), "int", ""),
              ("yyyy.mm.dd values", '=SUMPRODUCT(--ISNUMBER(SEARCH(".",tDirty25[Order_Date])))', sum("." in v for v in od), "int", ""),
              ("C24 '14/05/2025' repaired", f"=DATE(RIGHT({DD}C24,4),MID({DD}C24,4,2),LEFT({DD}C24,2))", serial("2025-05-14"), "date", ""),
              ("C63 '04-May-2025' repaired", f"=DATEVALUE({DD}C63)", serial("2025-05-04"), "date", "English month names: fine on most PCs."),
              ("C72 '2025.06.10' repaired", f'=DATE(LEFT({DD}C72,4),MID({DD}C72,6,2),RIGHT({DD}C72,2))', serial("2025-06-10"), "date",
               "Safer than DATEVALUE(SUBSTITUTE(…)).")])
    m.turn()
    first_dot = next(i for i, v in enumerate(DR.Order_Date) if isinstance(v, str) and "." in v and i != 70)
    fv = DR.Order_Date[first_dot]
    m.ex(f"Repair C{first_dot + 2} of Data_Dirty25 ('{fv}'). Type the month number of the date.",
         f"=MONTH(DATE(LEFT({DD}C{first_dot + 2},4),MID({DD}C{first_dot + 2},6,2),RIGHT({DD}C{first_dot + 2},2)))", int(fv[5:7]), "int",
         short=f"Month of C{first_dot + 2}")
    m.ex("How many Order_Date values are text in total?", "=SUMPRODUCT(--ISTEXT(tDirty25[Order_Date]))", DF["text_dates"], "int",
         short="Text order dates")

    # ---------------------------------------------------------------- 2.5.9
    qreal = int(sum(1 for v in nb.Quantity if v is None))
    q = [v for v in DR.Quantity if _num(v)]
    sb = sum(1 for o, s_ in zip(DR.Order_Date, DR.Ship_Date) if isinstance(o, pd.Timestamp) and s_ < o)
    m.topic("2.5.9", "Handling missing values",
            "A blank is not zero. For each column decide: recover the value, flag the row, or exclude it, and write the decision down. "
            "Typing 0 into a blank Quantity changes the average and makes Sales ÷ Quantity a #DIV/0! error.")
    m.worked([("Blank Quantity on real lines", "=COUNTBLANK(tDirty25[Quantity])-COUNTBLANK(tDirty25[Order_ID])", qreal, "int",
               "Flag; ask the store."),
              ("S31 flagged", f'=IF({DD}S31="","Missing","")', "Missing", None, ""),
              ("Average Quantity, blanks ignored", "=AVERAGE(tDirty25[Quantity])", float(np.mean(q)), "dec2", "AVERAGE skips blanks."),
              ("Average if blanks were typed as 0", "=SUM(tDirty25[Quantity])/(ROWS(tDirty25)-COUNTBLANK(tDirty25[Order_ID]))",
               float(sum(q) / len(nb)), "dec2", "Lower, and wrong."),
              ("Ship date before order date (real dates)", "=SUMPRODUCT(ISNUMBER(tDirty25[Order_Date])*(tDirty25[Ship_Date]<tDirty25[Order_Date]))",
               sb, "int", f"{DF['ship_before']} once the text dates are repaired too.")])
    m.turn()
    m.ex("How many real lines (not blank rows) have no Quantity?", "=COUNTBLANK(tDirty25[Quantity])-COUNTBLANK(tDirty25[Order_ID])", qreal,
         "int", short="Real lines missing Quantity")
    m.ex("What is the average Quantity, ignoring blanks? (2 decimal places)", "=AVERAGE(tDirty25[Quantity])", float(np.mean(q)), "dec2",
         short="Average Quantity")

    # ---------------------------------------------------------------- 2.5.10
    m.topic("2.5.10", "Removing blank rows and empty columns",
            "Blank rows break a table: sorting, filtering and Ctrl+A stop at the first gap. Detect them with a key column that is never "
            "blank in a real row, then delete whole sheet rows.")
    m.steps(["On your copy, select the Order_ID column.", "Home › Find & Select › Go To Special › Blanks › OK.",
             "Home › Delete › Delete Sheet Rows.", "Column1 (M) is empty: =COUNTA(tDirty25[Column1]) is 0. Delete the column."])
    m.worked([("Blank rows", "=COUNTBLANK(tDirty25[Order_ID])", DF["blank_rows"], "int", ""),
              ("Values in Column1", "=COUNTA(tDirty25[Column1])", 0, "int", "Empty: delete it."),
              ("Rows after deleting blanks", "=ROWS(tDirty25)-COUNTBLANK(tDirty25[Order_ID])", DF["after_blank"], "int", ""),
              ("Columns after deleting Column1", "=COLUMNS(tDirty25)-1", DR.shape[1] - 1, "int", "")])
    m.turn()
    m.ex("How many rows remain after deleting the blank rows?", "=ROWS(tDirty25)-COUNTBLANK(tDirty25[Order_ID])", DF["after_blank"], "int",
         short="Rows after blank rows removed")
    m.ex_text("Which Go To Special option selects the empty cells?", "Blanks")

    # ---------------------------------------------------------------- 2.5.11
    p2 = DR.Product_ID[1]
    m.topic("2.5.11", "Using Text to Columns",
            "Data › Text to Columns splits one column into several at a delimiter or at fixed widths. It overwrites the cells to the "
            "right without asking, so insert empty columns first. The formula equivalents are TEXTBEFORE and TEXTAFTER.")
    m.steps(["Insert 3 empty columns to the right of Product_ID on your copy.", "Select Product_ID › Data › Text to Columns › Delimited › Next.",
             "Other: - › Next.", "Click the third column › Column data format: Text (keeps leading zeros) › Finish."])
    m.worked([("N2 as stored", f"={DD}N2", DR.Product_ID[0], None, ""),
              ("Cat_Code", f'=TEXTBEFORE({DD}N2,"-")', DR.Product_ID[0].split("-")[0], None, ""),
              ("Sub_Code", f'=TEXTBEFORE(TEXTAFTER({DD}N2,"-"),"-")', DR.Product_ID[0].split("-")[1], None, ""),
              ("Item_No", f'=TEXTAFTER({DD}N2,"-",2)', DR.Product_ID[0].split("-")[2], None, "Text, like Text to Columns with Text format.")])
    m.turn()
    m.ex_text("What Sub_Code does N3 split into?", p2.split("-")[1], formula=f'=TEXTBEFORE(TEXTAFTER({DD}N3,"-"),"-")', short="Sub_Code of N3")
    m.ex("How many FY2025 clean lines have a Product_ID starting with TEC?", '=COUNTIF(tSales25[Product_ID],"TEC-*")',
         int(S.Product_ID.str.startswith("TEC-").sum()), "int", short="TEC lines")

    # ---------------------------------------------------------------- 2.5.12
    big = sum(1 for v in DR.Sales if _num(v) and v > S.Sales.max())
    m.topic("2.5.12", "Verifying a clean dataset",
            "Control totals are agreed figures from the source system (row count, totals) that the cleaned data must match. Cleaning is "
            "finished only when every one matches. Here the POS control totals are checked against tSales25.")
    rows = [("Rows", "=ROWS(tSales25)", CTRL["rows"], "int"), ("Total Sales", "=SUM(tSales25[Sales])", CTRL["sales"], N),
            ("Total Profit", "=SUM(tSales25[Profit])", CTRL["profit"], N), ("Total Quantity", "=SUM(tSales25[Quantity])", CTRL["qty"], "int"),
            ("Distinct orders", "=COUNTA(UNIQUE(tSales25[Order_ID]))", CTRL["orders"], "int"),
            ("Distinct customers", "=COUNTA(UNIQUE(tSales25[Customer_ID]))", CTRL["customers"], "int")]
    m.h("POS control totals (typed from the POS report)")
    ctl = m.s.inputs([(f"Control: {lab}", val, kind, "") for lab, _, val, kind in rows])
    m.worked([(f"{lab}: clean data", fx, val, kind, "") for lab, fx, val, kind in rows])
    m.worked([(f"{lab}: check", f'=IF({fx[1:]}={ctl["Control: " + lab]},"PASS","FAIL")', "PASS", None, "") for lab, fx, val, kind in rows],
             label="Checks")
    m.turn()
    m.ex(f"How many numeric Sales values in tDirty25 are larger than the largest clean line? (These are the ×100 keying errors, "
         f"{DF['x100']} in all once the text values are converted.)", '=COUNTIF(tDirty25[Sales],">"&MAX(tSales25[Sales]))', big, "int",
         short="Sales above the clean maximum")
    m.ex("How many distinct customers are in the clean data?", "=COUNTA(UNIQUE(tSales25[Customer_ID]))", CTRL["customers"], "int",
         short="Distinct customers")
    m.finish()


# =====================================================================================================
def m26(b, D):
    DR, S = D.DR, D.S
    m = ModuleSheet(b, "2.6 PQ Transforms", "2.6  Cleaning and Transforming in Power Query",
                    "Tutor walk-through after Module 2.6: the same clean-up as 2.5, recorded as Power Query steps, plus columns, grouping, "
                    "pivoting and refresh", "2.6.1 - 2.6.20")
    m.dataset([("SuperMart_Sales_2025_DIRTY.csv", "Week2_Data_Files folder", "The file you clean in the Power Query Editor (query qSales_Raw "
                                                                            "from 2.2.7, or a new one)."),
               ("tDirty25", "Sheet Data_Dirty25 (starts in A1)", "The same file as Excel opens it: the answer key counts on it."),
               ("tSales25", "Sheet Data_Sales25 (starts in A1)", "What the finished query must produce: 3,312 rows matching the control totals.")])
    code(m, M_DIRTY, "The finished cleaning query (Home › Advanced Editor › paste; set the DataFolder parameter first, 2.3.7)")
    m.note("Build it step by step through the topics below; paste the whole script only to check your work. Five Sales values are 100 times "
           "too large (2.5.12): correct them from the receipts before the totals match.")

    # ---------------------------------------------------------------- 2.6.1-3
    m.topic("2.6.1-3", "Removing columns, blank rows and duplicate records",
            "Home › Remove Columns drops columns; Remove Rows › Remove Blank Rows drops rows where every value is empty; Remove Rows › "
            "Remove Duplicates keeps the first of each set of identical rows. Select ALL columns before removing duplicates: the key "
            "decides what counts as a duplicate.")
    m.steps(["Right-click Column1 › Remove.", "Home › Remove Rows › Remove Blank Rows.",
             "Ctrl+A to select every column › Home › Remove Rows › Remove Duplicates.",
             "Status bar: 20 COLUMNS, 3,312 ROWS (with profiling on the entire data set)."])
    m.worked([("Columns after removing Column1", "=COLUMNS(tDirty25)-1", DR.shape[1] - 1, "int", ""),
              ("Rows after Remove Blank Rows", "=ROWS(tDirty25)-COUNTBLANK(tDirty25[Order_ID])", DF["after_blank"], "int", ""),
              ("Rows after Remove Duplicates (all columns)", "=ROWS(UNIQUE(tDirty25))-1", CTRL["rows"], "int", "Matches the POS control."),
              ("Rows if the key were Order_ID only", "=ROWS(UNIQUE(tSales25[Order_ID]))", CTRL["orders"], "int",
               "Wrong: 1,625 real lines lost.")])
    m.turn()
    m.ex("With Order_ID as the only key, how many real lines would be lost?", "=ROWS(tSales25)-ROWS(UNIQUE(tSales25[Order_ID]))",
         CTRL["rows"] - CTRL["orders"], "int", short="Lines lost with Order_ID key")
    m.ex_text("Which columns must be selected for Remove Duplicates to remove only exact copies?", "All", alts=("all columns", "every column"))

    # ---------------------------------------------------------------- 2.6.4-5
    m.topic("2.6.4-5", "Filtering rows and replacing values",
            "A filter in Power Query removes rows from the result (it does not just hide them). Replace Values swaps one value for "
            "another and is case-sensitive: 'FURNITURE' and 'furniture' need separate steps, or a mapping (2.5.7).")
    code(m, '''    Losses    = Table.SelectRows(Typed, each [Profit] < 0),
    Furnitre  = Table.ReplaceValue(Losses, "Furnitre", "Furniture", Replacer.ReplaceValue, {"Category"}),
    NA        = Table.ReplaceValue(Furnitre, "N/A", null, Replacer.ReplaceValue, {"Customer_Name"})''',
         "Steps written by Number Filters › Less Than and Transform › Replace Values")
    loss = S[S.Profit < 0]
    m.worked([("Lines with Profit < 0", '=COUNTIF(tSales25[Profit],"<0")', len(loss), "int", ""),
              ("Sales on those lines", '=SUMIF(tSales25[Profit],"<0",tSales25[Sales])', int(loss.Sales.sum()), N, ""),
              ("Loss on those lines", '=SUMIF(tSales25[Profit],"<0")', int(loss.Profit.sum()), N, ""),
              ("Cells exactly 'Furnitre'", '=SUMPRODUCT(--EXACT(tDirty25[Category],"Furnitre"))', 8, "int", ""),
              ("Cells exactly 'FURNITURE'", '=SUMPRODUCT(--EXACT(tDirty25[Category],"FURNITURE"))', 16, "int", "A separate Replace step.")])
    m.turn()
    m.ex("How many FY2025 lines remain if you filter Discount to 'greater than or equal to 0.3'?", '=COUNTIF(tSales25[Discount],">=0.3")',
         int((S.Discount >= 0.3).sum()), "int", short="Lines with Discount >= 30%")
    m.ex("How many Customer_Name cells will Replace Values 'N/A' › null change?", '=COUNTIF(tDirty25[Customer_Name],"N/A")', DF["na_cust"],
         "int", short="N/A names replaced")

    # ---------------------------------------------------------------- 2.6.6
    m.topic("2.6.6", "Changing data types with locale",
            "Change Type › Using Locale converts text using another region's rules. The same text 03/04/2025 is 3 April in Nigeria and "
            "the UK and 4 March in the USA. With the wrong locale some values error and, worse, some convert silently to the wrong date.")
    lit_grid(m, [["English (United States)", DF["us_err"], DF["us_wrong"], DF["us_ok"]], ["English (United Kingdom)", 0, 0, DF["d_slash"]]],
             header=["Locale used on the 106 dd/mm/yyyy values", "Errors", "Silently wrong", "Correct"])
    m.steps(["Right-click Order_Date › Change Type › Using Locale…", "Data Type: Date; Locale: English (United Kingdom) › OK.",
             "Check View › Column quality: 0% Error.", "Spot-check: Row_ID of C24 (14/05/2025) shows 14 May 2025."])
    m.worked([("dd/mm/yyyy values to convert", '=SUMPRODUCT(--ISNUMBER(SEARCH("/",tDirty25[Order_Date])))', DF["d_slash"], "int", ""),
              ("Day 13 or more (US locale errors)", '=SUMPRODUCT(ISNUMBER(SEARCH("/",tDirty25[Order_Date]))*(IFERROR(--LEFT(tDirty25[Order_Date],2),0)>12))',
               DF["us_err"], "int", "No 13th month, so the US reading fails."),
              ("Day 12 or less and day <> month (silently wrong)",
               '=SUMPRODUCT(ISNUMBER(SEARCH("/",tDirty25[Order_Date]))*(IFERROR(--LEFT(tDirty25[Order_Date],2),99)<=12)'
               '*(IFERROR(--LEFT(tDirty25[Order_Date],2),0)<>IFERROR(--MID(tDirty25[Order_Date],4,2),0)))', DF["us_wrong"], "int",
               "Day and month swap without any error.")])
    m.turn()
    m.ex_text("Which locale reads 14/05/2025 correctly?", "English (United Kingdom)", alts=("en-gb", "uk", "english (uk)"))
    m.ex("How many of the dd/mm/yyyy dates does the US locale read correctly by luck (day = month)?", f"={DF['d_slash']}-{DF['us_err']}-{DF['us_wrong']}",
         DF["us_ok"], "int", short="Correct by luck")

    # ---------------------------------------------------------------- 2.6.7-8
    loc = [v for v in DR.Store_Location if v]
    m.topic("2.6.7-8", "Splitting columns by delimiter, and merging columns",
            "Split Column › By Delimiter cuts one column into several at a character such as ', '. Merge Columns joins several into one "
            "with a separator. Both keep the original only if you duplicate it first.")
    code(m, '''    Split  = Table.SplitColumn(Prev, "Store_Location", Splitter.SplitTextByDelimiter(", ", QuoteStyle.Csv), {"City", "State"}),
    Merged = Table.AddColumn(Split, "Store_Label", each [Store_ID] & " · " & [Store_Name], type text)''',
         "Split Store_Location; add a Store_Label (Add Column › Merge Columns keeps the originals)")
    m.worked([("Store_Location of row 2", f"={DD}K2", DR.Store_Location[0], None, ""),
              ("City", f'=TEXTBEFORE({DD}K2,", ")', DR.Store_Location[0].split(", ")[0], None, ""),
              ("State", f'=TEXTAFTER({DD}K2,", ")', DR.Store_Location[0].split(", ")[1], None, ""),
              ("Store label for ST05", '=XLOOKUP("ST05",tStores[Store_ID],tStores[Store_ID]&" · "&tStores[Store_Name])', "ST05 · Benin City - Ring Road",
               None, "")])
    m.turn()
    m.ex("How many lines in tDirty25 are in Lagos State? (Hint: COUNTIF with \"*, Lagos\".)", '=COUNTIF(tDirty25[Store_Location],"*, Lagos")',
         sum(1 for v in loc if v.endswith(", Lagos")), "int", short="Lagos State lines")
    m.ex_text("What label does the merge give store ST09?", "ST09 · Jos - Rayfield",
              formula='=XLOOKUP("ST09",tStores[Store_ID],tStores[Store_ID]&" · "&tStores[Store_Name])', short="Label for ST09")

    # ---------------------------------------------------------------- 2.6.9-10
    m.topic("2.6.9-10", "Trimming and cleaning text, and extracting text",
            "Transform › Format › Trim and Clean work like TRIM and CLEAN, except Text.Trim removes only leading and trailing spaces. "
            "Add Column › Extract takes the first or last characters, a range, or the text before or after a delimiter.")
    code(m, '''    Trimmed = Table.TransformColumns(Prev, {{"Zone", Text.Trim, type text}, {"Customer_Name", Text.Clean, type text}}),
    CatCode = Table.AddColumn(Trimmed, "Cat_Code", each Text.Start([Product_ID], 3), type text),
    ItemNo  = Table.AddColumn(CatCode, "Item_No", each Text.AfterDelimiter([Product_ID], "-", 1), type text)''',
         "Trim, Clean and two extractions")
    m.worked([("Zone values with extra spaces", "=SUMPRODUCT(--(LEN(tDirty25[Zone])<>LEN(TRIM(tDirty25[Zone]))))",
               sum(1 for v in DR.Zone if isinstance(v, str) and v != " ".join(v.split())), "int", ""),
              ("Distinct Zone values after Trim", '=ROWS(UNIQUE(TRIM(FILTER(tDirty25[Zone],tDirty25[Zone]<>""))))', 4, "int", ""),
              ("Text.Start([Product_ID], 3) on N2", f"=LEFT({DD}N2,3)", DR.Product_ID[0][:3], None, "Excel twin: LEFT."),
              ("Text.AfterDelimiter(…, \"-\", 1) on N2", f'=TEXTAFTER({DD}N2,"-",2)', DR.Product_ID[0].split("-")[2], None,
               "M counts delimiters from 0; Excel's instance_num from 1.")])
    m.turn()
    m.ex("How many FY2025 clean lines have Cat_Code FUR?", '=COUNTIF(tSales25[Product_ID],"FUR-*")', int(S.Product_ID.str.startswith("FUR").sum()),
         "int", short="FUR lines")
    m.ex_text("Text.Trim removes inner double spaces too: true or false?", "False", alts=("f",))

    # ---------------------------------------------------------------- 2.6.11-13
    bands = [("None", S.Discount == 0), ("1-20%", (S.Discount > 0) & (S.Discount <= 0.2)), ("Above 20%", S.Discount > 0.2)]
    r1 = S.iloc[1]
    m.topic("2.6.11-13", "Conditional columns, custom columns and index columns",
            "Add Column › Conditional Column builds if-then-else rules from menus. Custom Column takes any M formula, e.g. [Profit] / "
            "[Sales]. Index Column numbers the rows from 0 or 1 in their current order, so sort first.")
    code(m, '''    Band   = Table.AddColumn(Prev, "Disc_Band", each if [Discount] = 0 then "None" else if [Discount] <= 0.2 then "1-20%" else "Above 20%", type text),
    Margin = Table.AddColumn(Band, "Margin", each [Profit] / [Sales], Percentage.Type),
    Sorted = Table.Sort(Margin, {{"Order_Date", Order.Ascending}, {"Row_ID", Order.Ascending}}),
    LineNo = Table.AddIndexColumn(Sorted, "Line_No", 1, 1, Int64.Type)''', "Disc_Band, Margin and Line_No")
    m.worked([("None: lines", "=COUNTIFS(tSales25[Discount],0)", int(bands[0][1].sum()), "int", ""),
              ("1-20%: lines", '=COUNTIFS(tSales25[Discount],">0",tSales25[Discount],"<=0.2")', int(bands[1][1].sum()), "int", ""),
              ("Above 20%: lines", '=COUNTIFS(tSales25[Discount],">0.2")', int(bands[2][1].sum()), "int", ""),
              ("Above 20%: profit", '=SUMIFS(tSales25[Profit],tSales25[Discount],">0.2")', int(S[bands[2][1]].Profit.sum()), N, "Deep discounts lose money."),
              (f"Margin of Row_ID {r1.Row_ID}", f"=XLOOKUP({r1.Row_ID},tSales25[Row_ID],tSales25[Profit])/XLOOKUP({r1.Row_ID},tSales25[Row_ID],tSales25[Sales])",
               r1.Profit / r1.Sales, "pct", "The custom column for one row."),
              ("Row_ID with Line_No 1", "=INDEX(tSales25[Row_ID],1)", int(S.Row_ID.iloc[0]), "gen", "tSales25 is sorted by date, then Row_ID.")])
    r2 = S.iloc[2]
    m.turn()
    m.ex("What is the total Sales of the 'None' band?", "=SUMIFS(tSales25[Sales],tSales25[Discount],0)", int(S[bands[0][1]].Sales.sum()), N,
         short="Sales, None band")
    m.ex(f"What is the Margin of Row_ID {r2.Row_ID}?", f"=XLOOKUP({r2.Row_ID},tSales25[Row_ID],tSales25[Profit])/XLOOKUP({r2.Row_ID},tSales25[Row_ID],tSales25[Sales])",
         r2.Profit / r2.Sales, "pct", short=f"Margin of {r2.Row_ID}")
    m.ex("Which Row_ID gets Line_No 100?", "=INDEX(tSales25[Row_ID],100)", int(S.Row_ID.iloc[99]), "gen", short="Row_ID at Line_No 100")

    # ---------------------------------------------------------------- 2.6.14
    zs = S.groupby("Zone").agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"))
    m.topic("2.6.14", "Grouping data with Group By",
            "Home › Group By collapses rows that share values into one row each, with the aggregations you choose. Advanced lets you "
            "group by several columns and add several aggregations at once.")
    code(m, '''    Grouped = Table.Group(Prev, {"Zone"}, {{"Sales", each List.Sum([Sales]), type number}, {"Profit", each List.Sum([Profit]), type number},
                  {"Lines", each Table.RowCount(_), Int64.Type}}),
    Margin  = Table.AddColumn(Grouped, "Margin", each [Profit] / [Sales], Percentage.Type)''', "Group By Zone")
    m.worked([(f"{z}: Sales", f'=SUMIFS(tSales25[Sales],tSales25[Zone],"{z}")', int(zs.loc[z, "Sales"]), N, "") for z in
              ["South-West", "South-South", "North-Central", "South-East"]] +
             [("South-West margin", '=SUMIFS(tSales25[Profit],tSales25[Zone],"South-West")/SUMIFS(tSales25[Sales],tSales25[Zone],"South-West")',
               zs.loc["South-West", "Profit"] / zs.loc["South-West", "Sales"], "pct", "Sum ÷ sum, never an average of line margins.")])
    m.turn()
    m.ex("What is North-Central's margin?", '=SUMIFS(tSales25[Profit],tSales25[Zone],"North-Central")/SUMIFS(tSales25[Sales],tSales25[Zone],"North-Central")',
         zs.loc["North-Central", "Profit"] / zs.loc["North-Central", "Sales"], "pct", short="North-Central margin")
    m.ex("How many rows does Group By Zone return?", "=ROWS(UNIQUE(tSales25[Zone]))", 4, "int", short="Rows after Group By Zone")

    # ---------------------------------------------------------------- 2.6.15
    cq = S.assign(Q=S.Order_Date.dt.quarter).pivot_table(index="Category", columns="Q", values="Sales", aggfunc="sum")
    m.topic("2.6.15", "Pivoting and unpivoting columns",
            "Pivot Column turns the values of one column into column headers (long to wide). Unpivot Columns does the reverse (wide to "
            "long), which is the shape analysis and PivotTables need. 2.4.3 unpivoted the PDF bulletin.")
    code(m, '''    Quarter = Table.AddColumn(Prev, "Quarter", each "Q" & Text.From(Date.QuarterOfYear([Order_Date])), type text),
    Keep    = Table.SelectColumns(Quarter, {"Category", "Quarter", "Sales"}),
    Pivoted = Table.Pivot(Keep, List.Sort(List.Distinct(Keep[Quarter])), "Quarter", "Sales", List.Sum),
    Back    = Table.UnpivotOtherColumns(Pivoted, {"Category"}, "Quarter", "Sales")''', "Category × Quarter: pivot, then unpivot back")
    m.worked([("Technology, Q4", '=SUMIFS(tSales25[Sales],tSales25[Category],"Technology",tSales25[Order_Date],">="&DATE(2025,10,1))',
               int(cq.loc["Technology", 4]), N, "One cell of the pivoted table."),
              ("Cells in the pivoted table", "=ROWS(UNIQUE(tSales25[Category]))*4", 12, "int", "3 categories × 4 quarters."),
              ("Rows after unpivoting it back", "=ROWS(UNIQUE(tSales25[Category]))*4", 12, "int", "One row per category per quarter.")])
    m.turn()
    m.ex("What are Furniture's Q1 sales?", '=SUMIFS(tSales25[Sales],tSales25[Category],"Furniture",tSales25[Order_Date],"<"&DATE(2025,4,1))',
         int(cq.loc["Furniture", 1]), N, short="Furniture Q1")
    m.ex_text("Which command turns month columns into rows: Pivot Column or Unpivot Columns?", "Unpivot Columns", alts=("unpivot",))

    # ---------------------------------------------------------------- 2.6.16-17
    c88 = S[S.Customer_ID == "SMC-0388"]
    m.topic("2.6.16-17", "Handling null values, and sorting and filtering",
            "null is Power Query's empty value: not zero and not empty text. Replace null with 0 only when zero is true. Sorting before "
            "Remove Duplicates decides which row is kept (the first), e.g. each customer's latest order.")
    code(m, '''    Flag    = Table.AddColumn(Prev, "Qty_Flag", each if [Quantity] = null then "Missing" else "OK", type text),
    Newest  = Table.Buffer(Table.Sort(Flag, {{"Order_Date", Order.Descending}})),     // Buffer fixes the sort before de-duplicating
    Latest  = Table.Distinct(Newest, {"Customer_ID"})''', "Flag nulls; one row per customer with the latest order")
    m.worked([("Quantity null on real lines", "=COUNTBLANK(tDirty25[Quantity])-COUNTBLANK(tDirty25[Order_ID])", DF["qty_blank"], "int",
               "Flag them; Unit_Price = Sales / null stays null."),
              ("Customers (rows after Distinct)", "=COUNTA(UNIQUE(tSales25[Customer_ID]))", CTRL["customers"], "int", ""),
              ("SMC-0388 latest order", '=MAXIFS(tSales25[Order_Date],tSales25[Customer_ID],"SMC-0388")', serial(c88.Order_Date.max()), "date",
               "Kept when sorted newest first."),
              ("SMC-0388 first order", '=MINIFS(tSales25[Order_Date],tSales25[Customer_ID],"SMC-0388")', serial(c88.Order_Date.min()), "date",
               "Kept if sorted oldest first.")])
    c1 = S[S.Customer_ID == "SMC-0001"]
    m.turn()
    m.ex("What is the latest 2025 order date of customer SMC-0001? (Format the cell as a date.)",
         '=MAXIFS(tSales25[Order_Date],tSales25[Customer_ID],"SMC-0001")', serial(c1.Order_Date.max()), "date", short="SMC-0001 latest order")
    m.ex_text("Replace null Quantity with 0: right or wrong?", "Wrong", alts=("no", "false"))

    # ---------------------------------------------------------------- 2.6.18-20
    m.topic("2.6.18-20", "Reordering and renaming, the step history, and refreshing after source changes",
            "Rename columns early to one style (Transaction_ID, not transaction_id). Rename steps to say what they did. Steps that name "
            "columns break if the source renames them, so agree fixed names with whoever sends the file and check counts after refresh.")
    m.do([("Connect to the Transactions sheet (2.3.1). Double-click each header: transaction_id › Transaction_ID, customer_id › Customer_ID.",
           "Names match the Customers sheet for Week 3 merges."),
          ("Drag Order_Date next to Order_ID.", "A Reordered Columns step appears."),
          ("In qSales_Raw, right-click 'Removed Blank Rows' › Rename › 'Removed 15 spacer rows'.", "Steps read like an audit trail."),
          ("Right-click a step › Properties › Description: why it exists.", "Hover over the step to read it."),
          ("Test: in a copy of the CSV rename the header Zone to Region › Refresh.", "Expression.Error: The column 'Zone' of the table wasn't found.")])
    m.turn()
    m.ex_text("A refresh fails with \"The column 'Zone' of the table wasn't found\". What changed in the source?", "Column renamed",
              alts=("renamed", "the column was renamed", "zone was renamed"))
    m.ex_text("New rows arrive in the same layout. Does Refresh work without changes: yes or no?", "Yes")
    m.finish()


# =====================================================================================================
def m27(b, D):
    DR, S = D.DR, D.S
    m = ModuleSheet(b, "2.7 Excel vs PQ", "2.7  Excel vs Power Query Cleaning",
                    "Tutor walk-through after Module 2.7: what each tool is good for, the cost of repeating manual work, auditability "
                    "and choosing the tool", "2.7.1 - 2.7.6")
    m.dataset([("tDirty25 and tSales25", "Data_Dirty25 and Data_Sales25 (A1)", "The same file before and after cleaning: the gap shows what a missed "
                                                                               "step costs.")])

    # ---------------------------------------------------------------- 2.7.1-2
    dsum = int(sum(v for v in DR.Sales if _num(v)))
    m.topic("2.7.1-2", "What manual Excel cleaning is good for, and what Power Query does better",
            "Excel is quickest for one-off, small, visual fixes. Power Query wins when the job repeats, the data is large or several "
            "people must get the same answer: the steps are written down and replay identically.")
    m.worked([("Careless SUM of the raw Sales column", "=SUM(tDirty25[Sales])", dsum, N, "Skips text, counts duplicates and ×100 errors."),
              ("Clean total", "=SUM(tSales25[Sales])", CTRL["sales"], N, ""),
              ("Error in the careless total", "=SUM(tDirty25[Sales])-SUM(tSales25[Sales])", dsum - CTRL["sales"], N,
               "One forgotten step, one wrong board report.")])
    lit_grid(m, [["Analyst A (forgot duplicates)", "3,372", "₦13.2m too high"], ["Analyst B (forgot text Sales)", "3,312", "₦56.7m missing"],
                 ["Power Query (all steps recorded)", "3,312", "Matches the control"]], header=["Who cleaned it", "Rows left", "Effect on Sales"])
    m.turn()
    m.ex("By how much does the careless SUM of tDirty25[Sales] differ from the clean total?", "=SUM(tDirty25[Sales])-SUM(tSales25[Sales])",
         dsum - CTRL["sales"], N, short="Careless total error")
    m.ex_text("Tidy 40 names once for a letter: Excel or Power Query?", "Excel")

    # ---------------------------------------------------------------- 2.7.3
    m.topic("2.7.3", "Reproducibility and the cost of manual cleaning",
            "Manual cleaning costs the same every month and can differ each time. A query costs more once, then a refresh. Type your own "
            "figures in the yellow cells; the totals follow.")
    inp = m.s.inputs([("Manual cleaning, hours per month", 2, "dec1", "Example figure: replace with your own"),
                      ("Building the query, hours (once)", 3, "dec1", ""),
                      ("Checking a refresh, hours per month", 0.1, "dec1", "About 5 minutes")])
    h_m, h_b, h_r = inp["Manual cleaning, hours per month"], inp["Building the query, hours (once)"], inp["Checking a refresh, hours per month"]
    m.worked([("Manual: hours per year", f"={h_m}*12", 24, "dec1", ""),
              ("Power Query: hours in year 1", f"={h_b}+{h_r}*12", 3 + 1.2, "dec1", ""),
              ("Hours saved in year 1", f"={h_m}*12-({h_b}+{h_r}*12)", 24 - 4.2, "dec1", ""),
              ("Month in which the query has paid for itself", f"=ROUNDUP({h_b}/({h_m}-{h_r}),0)", 2, "int", "")])
    m.turn()
    m.ex("With the example figures, how many hours does Power Query save in year 1?", f"={h_m}*12-({h_b}+{h_r}*12)", 19.8, "dec1",
         short="Hours saved in year 1")
    m.ex_open("Name one risk of manual cleaning that has nothing to do with time.", "Two people (or the same person on two days) clean "
                                                                                     "differently, so the numbers cannot be reproduced.")

    # ---------------------------------------------------------------- 2.7.4
    m.topic("2.7.4", "Auditability: proving how a number was produced",
            "An auditor asks where the data came from, what was removed and why. Power Query answers from the Source step and the named "
            "Applied Steps; manual cleaning answers from memory.")
    lit_grid(m, [["Where did the data come from?", "'An export'", "Source step shows the file"],
                 ["Why fewer rows?", "'I removed some'", "60 duplicates, 15 blank rows"],
                 ["Can you do it again?", "'Roughly'", "Refresh"]], header=["Question", "Manual answer", "Power Query answer"])
    m.turn()
    m.ex("How many rows in total did the cleaning remove (blank rows + duplicates)?", "=ROWS(tDirty25)-ROWS(tSales25)",
         len(DR) - len(S), "int", short="Rows removed")
    m.ex_text("Where in Power Query do you see the file a query reads?", "Source step", alts=("source", "the source step"))

    # ---------------------------------------------------------------- 2.7.5-6
    m.topic("2.7.5-6", "Choosing the right tool, and a practical comparison on the same file",
            "Use Power Query for anything that repeats or feeds a report; use Excel for one-off fixes and final touches. On the dirty "
            "sales file the same clean-up takes far fewer recorded steps than manual actions, and only the steps can be replayed.")
    lit_grid(m, [["Blank rows, duplicates", 2, 2], ["Text Sales, text dates", 7, 3], ["Spaces, case, categories", 6, 3],
                 ["Every following month", "All again", "Refresh"]], header=["Task", "Excel actions", "Power Query steps"])
    m.turn()
    m.ex_text("Combine the 12 monthly files every month: Excel or Power Query?", "Power Query", alts=("pq",))
    m.ex_text("Fix one typo in a finished report: Excel or Power Query?", "Excel")
    m.ex_text("Clean 3,312 POS lines that arrive every month: Excel or Power Query?", "Power Query", alts=("pq",))
    m.finish()


def build(b, D):
    m25(b, D)
    m26(b, D)
    m27(b, D)
