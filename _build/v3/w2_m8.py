"""Week 2 edition-3 slide copy: Module 2.8A (functions 1-33) and 2.8B (functions 34-65).
One slide per function: definition with syntax, when to use it, worked formulas on our tables, one trap.
Every result is computed below from w2_data (tables tSales25, tStores, tDirty25, tAdm23)."""
import numpy as np
import pandas as pd

from v3kit import T
from w2_fx import DF, N, I, P
import w2_data as D

S = D.table("tSales25")
ST = D.table("tStores").set_index("Store_ID")
DR = D.dirty_raw()
A = D.table("tAdm23")

TOT = int(S.Sales.sum())
AVG = S.Sales.mean()
MARGIN = S.Profit.sum() / TOT
TECH = int(S[S.Category == "Technology"].Sales.sum())
SWTECH = int(S[(S.Category == "Technology") & (S.Zone == "South-West")].Sales.sum())
_c1 = S[(S.Segment == "Corporate") & (S.Store_ID == "ST01")]
LOSS = int((S.Profit < 0).sum())
_fl = S[(S.Category == "Furniture") & (S.Profit < 0)]
_se = S[S.Zone == "South-East"].sort_values("Sales").iloc[-1]
MODE = S.Quantity.value_counts()
_mx = S.loc[S.Sales.idxmax()]
_mn = S.loc[S.Profit.idxmin()]
LG = S.Sales.nlargest(3).tolist()
SM = S.Profit.nsmallest(3).tolist()
TU = int(S[S.Category == "Technology"].Quantity.sum())
TARGET = int(ST.Target_2025.sum())


def st(sid):
    x = S[S.Store_ID == sid]
    sales, tgt = x.Sales.sum(), ST.loc[sid, "Target_2025"]
    return dict(name=ST.loc[sid, "Store_Name"], sales=sales, target=tgt, att=sales / tgt, margin=x.Profit.sum() / sales,
                mgr=ST.loc[sid, "Store_Manager"], city=ST.loc[sid, "City"])


IK, OW, JO, LK = st("ST01"), st("ST12"), st("ST09"), st("ST02")
SE_STORES = ST[ST.Zone == "South-East"].Store_Name.tolist()
SS = ST[ST.Zone == "South-South"].sort_values("Sales_2025", ascending=False)
RANK = ST.sort_values("Sales_2025", ascending=False)
ATT = (ST.Sales_2025 / ST.Target_2025).sort_values(ascending=False)
CUST = S.Customer_ID.value_counts()
CHAIR_S = int(S.Product_Name.str.lower().str.contains("chair").sum())
CHAIR_F = int(S.Product_Name.str.contains("chair").sum())
CAT_UP = DR[DR.Category != ""].Category.str.strip().str.upper().nunique()
STORE_RAW = DR[DR.Store_Name != ""].Store_Name.nunique()
STORE_PROPER = DR[DR.Store_Name != ""].Store_Name.str.title().nunique()
P1, P2, P3 = S.Product_ID.head(3).tolist()
O1, O2, O3 = S.Order_ID.head(3).tolist()
A0, A1 = A.iloc[0], A.iloc[1]
LOS = (A.Discharge_Date - A.Admission_Date).dt.days.mean()
_sc = S[S.Ship_Mode == "Standard Class"]
NWD_STD = np.mean([np.busday_count(o.date(), (s + pd.Timedelta(days=1)).date()) for o, s in zip(_sc.Order_Date, _sc.Ship_Date)])
NWD_1 = int(np.busday_count("2025-01-01", "2025-01-07"))
NWD_1H = int(np.busday_count("2025-01-01", "2025-01-07", holidays=["2025-01-01"]))


def d(s):
    return pd.Timestamp(s).strftime("%d-%b-%Y")


def wd(s, n):
    r = np.busday_offset(np.datetime64(s), n, roll="forward") if np.is_busday(np.datetime64(s)) else \
        np.busday_offset(np.datetime64(s), n - 1, roll="forward")
    return pd.Timestamp(r).strftime("%a %d-%b-%Y")


def F(tag, name, defn, simple, fx, watch=None, like=None, head=None, widths=None):
    return T(tag, name, defn=defn, simple=simple, like=like, fx=fx, fx_label=f"WORKED EXAMPLES  ·  {name}",
             fx_head=head, fx_widths=widths or ([3.3, 1.5, 1.6] if len(fx[0]) == 3 else None), watch=watch)


TOPICS_A = [
    F("2.8.1", "SUM", "=SUM(number1, [number2], …) adds numbers, ranges or table columns. Text and blanks are ignored.",
      "Totals: the first check on any dataset.",
      [["=SUM(tSales25[Sales])", N(TOT), "FY2025 sales"], ["=COUNT(tSales25[Sales])", I(len(S)), "Numbers summed"],
       ["=COUNTA(tSales25[Sales])", I(len(S)), "Non-empty cells"]],
      "SUM skips numbers stored as text without warning. If COUNT is below COUNTA, some values are text."),
    F("2.8.2", "SUMIF", "=SUMIF(range, criteria, [sum_range]) adds the cells in sum_range whose matching cell in range meets "
      "one condition.", "One condition: sales of one category, one store, one month.",
      [["=SUMIF(tSales25[Category], \"Technology\", tSales25[Sales])", N(TECH), "Technology sales"],
       ["=SUMIF(tSales25[Profit], \"<0\")", N(S[S.Profit < 0].Profit.sum()), "Total of all losses"]],
      "Criteria with an operator go in quotes: \"<0\", \">=0.3\". Without sum_range, range itself is added."),
    F("2.8.3", "SUMIFS", "=SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], …) adds the cells "
      "that meet every condition.", "Two or more conditions. Note the sum range comes first, unlike SUMIF.",
      [["=SUMIFS(tSales25[Sales], tSales25[Category], \"Technology\", tSales25[Zone], \"South-West\")", N(SWTECH),
        "Technology in the South-West"]], "All ranges must be the same size. Table columns guarantee this."),
    F("2.8.4", "AVERAGE", "=AVERAGE(number1, [number2], …) returns the arithmetic mean: the sum divided by the count of numbers.",
      "The typical value, when the data has no extreme values.",
      [["=AVERAGE(tSales25[Sales])", N(AVG), "Average line"], ["=SUM(tSales25[Sales]) / COUNT(tSales25[Sales])", N(AVG),
                                                                 "Same, by hand"]],
      f"A few huge lines pull the average up: the median line is only {N(S.Sales.median())} (2.8.13)."),
    F("2.8.5", "AVERAGEIFS", "=AVERAGEIFS(average_range, criteria_range1, criteria1, …) averages the cells that meet every condition.",
      "The average for one group: a segment at a store, a category in a month.",
      [["=AVERAGEIFS(tSales25[Sales], tSales25[Segment], \"Corporate\", tSales25[Store_ID], \"ST01\")", N(_c1.Sales.mean()),
        "Corporate line at Ikeja"],
       ["=COUNTIFS(tSales25[Segment], \"Corporate\", tSales25[Store_ID], \"ST01\")", I(len(_c1)), "Lines averaged"]],
      "Returns #DIV/0! when nothing matches. Check the count with COUNTIFS."),
    F("2.8.6", "COUNT", "=COUNT(value1, [value2], …) counts cells that contain numbers (dates are numbers too).",
      "Checks that a numeric column really holds numbers.",
      [["=COUNT(tSales25[Quantity])", I(len(S)), "Numeric cells"], ["=COUNT(tDirty25[Sales])",
                                                                    I(pd.to_numeric(DR.Sales, errors="coerce").notna().sum()),
                                                                    "Dirty file: text Sales not counted"]],
      "COUNT ignores text, so a column of numbers-as-text counts as 0."),
    F("2.8.7", "COUNTA", "=COUNTA(value1, [value2], …) counts cells that are not empty, whatever they contain.",
      "Counts records: rows with an ID, responses, entries.",
      [["=COUNTA(tSales25[Order_ID])", I(len(S)), "Sales lines"], ["=COUNTA(UNIQUE(tSales25[Order_ID]))",
                                                                   I(S.Order_ID.nunique()), "Distinct orders"]],
      "A cell holding a formula that returns \"\" looks empty but COUNTA counts it."),
    F("2.8.8", "COUNTBLANK", "=COUNTBLANK(range) counts empty cells, including cells with formulas that return \"\".",
      "Measures missing values before cleaning.",
      [["=COUNTBLANK(tDirty25[Quantity])", I((DR.Quantity == "").sum()), "Empty Quantity cells"],
       ["=COUNTBLANK(tDirty25[Order_ID])", I(DF["blank_rows"]), "Blank spacer rows"],
       ["Difference", I(DF["qty_blank"]), "Real lines missing Quantity"]],
      "Subtract the blank rows first, or you overstate the missing values."),
    F("2.8.9", "COUNTIF", "=COUNTIF(range, criteria) counts the cells in range that meet one condition.",
      "How many: losses, a category, a store.",
      [["=COUNTIF(tSales25[Profit], \"<0\")", I(LOSS), "Loss-making lines"],
       ["=COUNTIF(tSales25[Profit], \"<0\") / ROWS(tSales25)", P(LOSS / len(S)), "Share of lines"]],
      "Text criteria are not case-sensitive, and wildcards work: \"*chair*\"."),
    F("2.8.10", "COUNTIFS", "=COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], …) counts rows meeting every condition.",
      "Two or more conditions at once.",
      [["=COUNTIFS(tSales25[Category], \"Furniture\", tSales25[Profit], \"<0\")", I(len(_fl)), "Loss-making Furniture lines"],
       ["…, tSales25[Discount], \">=0.3\")", I((_fl.Discount >= 0.3).sum()), "… at 30% discount or more"]],
      "Conditions are joined by AND. For OR, add two COUNTIFS (or use SUM with an array)."),
    F("2.8.11", "SUBTOTAL", "=SUBTOTAL(function_num, ref1, …) aggregates only the visible rows when function_num is 101-111 "
      "(109 = SUM, 103 = COUNTA). 1-11 ignore filtered rows but include manually hidden ones.",
      "Totals that follow a filter. The table Total Row uses it.",
      [["=SUBTOTAL(109, tSales25[Sales]), no filter", N(TOT), "Whole year"],
       ["Same, filter Zone = South-West", N(S[S.Zone == "South-West"].Sales.sum()), "Visible rows only"]],
      "SUM ignores the filter and always returns the full total."),
    F("2.8.12", "AGGREGATE", "=AGGREGATE(function_num, options, array, [k]) runs one of 19 functions (14 = LARGE) and can skip "
      "errors (option 6) or hidden rows.", "LARGE, SMALL or a percentile with conditions, without a helper column.",
      [["=AGGREGATE(14, 6, tSales25[Sales] / (tSales25[Zone]=\"South-East\"), 1)", N(_se.Sales), "Largest South-East line"]],
      f"Rows outside the zone divide by 0; option 6 skips those errors. Result: {_se.Product_Name}, {_se.Store_Name}."),
    F("2.8.13", "MEDIAN", "=MEDIAN(number1, …) returns the middle value when the numbers are sorted.",
      "The typical value when a few extremes distort the average.",
      [["=MEDIAN(tSales25[Sales])", N(S.Sales.median()), "Middle line"], ["=AVERAGE(tSales25[Sales])", N(AVG), "Mean line"]],
      "Median far below the average means a few very large values. Report both."),
    F("2.8.14", "MODE.SNGL", "=MODE.SNGL(number1, …) returns the most frequent value. MODE.MULT returns all ties.",
      "The most common quantity, score or size.",
      [["=MODE.SNGL(tSales25[Quantity])", str(MODE.index[0]), "Most common quantity"],
       [f"=COUNTIF(tSales25[Quantity], {MODE.index[0]})", I(MODE.iloc[0]), "Lines with that quantity"],
       [f"=COUNTIF(tSales25[Quantity], {MODE.index[1]})", I(MODE.iloc[1]), "The runner-up"]],
      "Returns #N/A when no value repeats. Works on numbers only."),
    F("2.8.15", "MAX", "=MAX(number1, …) returns the largest number.", "The largest sale, the latest date, the highest score.",
      [["=MAX(tSales25[Sales])", N(_mx.Sales), "Largest line"], ["=MAX(tSales25[Order_Date])", "2025-12-30", "Latest order date"]],
      f"Check the maximum is genuine: {_mx.Product_Name} at {_mx.Store_Name}. Keying errors show up here first."),
    F("2.8.16", "MIN", "=MIN(number1, …) returns the smallest number.", "The worst loss, the earliest date.",
      [["=MIN(tSales25[Profit])", N(_mn.Profit), "Worst single loss"], ["=MIN(tSales25[Order_Date])", "2025-01-01", "First order date"]],
      f"The worst loss is a {_mn.Product_Name} at {_mn.Store_Name}, sold at {_mn.Discount:.0%} discount."),
    F("2.8.17", "LARGE", "=LARGE(array, k) returns the k-th largest value.", "Top-N lists.",
      [[f"=LARGE(tSales25[Sales], {k})", N(v), f"Rank {k}"] for k, v in enumerate(LG, 1)],
      "Put k (1, 2, 3) in cells and refer to them, so the list extends by filling down."),
    F("2.8.18", "SMALL", "=SMALL(array, k) returns the k-th smallest value.", "Bottom-N lists: worst losses, fastest times.",
      [[f"=SMALL(tSales25[Profit], {k})", N(v), f"Rank {k}"] for k, v in enumerate(SM, 1)],
      "Blanks and text are ignored; zeros are not."),
    F("2.8.19", "ROUND", "=ROUND(number, num_digits) rounds to num_digits decimal places; a negative num_digits rounds to "
      "tens, hundreds, thousands.", "Rounds the value itself, not just its display.",
      [["=AVERAGE(tSales25[Sales])", f"{AVG:,.2f}", "Unrounded"], ["=ROUND(AVERAGE(tSales25[Sales]), 0)", f"{round(AVG):,}", "Whole naira"],
       ["=ROUND(AVERAGE(tSales25[Sales]), -3)", f"{round(AVG, -3):,.0f}", "Nearest ₦1,000"]],
      "Format buttons only change what you see. Calculations still use every decimal unless you ROUND."),
    F("2.8.20", "ROUNDUP", "=ROUNDUP(number, num_digits) rounds away from zero.", "Anything you cannot have part of: cartons, buses, staff.",
      [["=SUMIFS(tSales25[Quantity], tSales25[Category], \"Technology\")", I(TU), "Technology units"],
       ["=ROUNDUP(… / 12, 0)", I(np.ceil(TU / 12)), f"Cartons of 12 ({TU / 12:.2f} rounded up)"]]),
    F("2.8.21", "ROUNDDOWN", "=ROUNDDOWN(number, num_digits) rounds towards zero.", "Cautious figures: never overstate.",
      [["=SUM(tSales25[Profit]) / SUM(tSales25[Sales])", f"{MARGIN:.5f}", "Exact margin"],
       ["=ROUNDDOWN(…, 3)", f"{np.floor(MARGIN * 1000) / 1000:.3f}", "Never rounds up"],
       ["=ROUNDDOWN(19.99, 0)", "19", "Whole units only"]]),
    F("2.8.22", "ABS", "=ABS(number) returns the number without its sign.", "Size of a difference, whatever its direction.",
      [["=MIN(tSales25[Profit])", f"{_mn.Profit:,}", "Worst loss"], ["=ABS(MIN(tSales25[Profit]))", f"{abs(_mn.Profit):,}",
                                                                      "As a positive amount"]],
      "Use ABS for variance size, but keep the sign for direction: over or under target matters."),
    F("2.8.23", "IF", "=IF(logical_test, value_if_true, [value_if_false]) returns one value if a test is TRUE and another if FALSE.",
      "Every yes/no decision in a sheet.",
      [["=SUM(tSales25[Sales])", N(TOT), "Actual"], ["=SUM(tStores[Target_2025])", N(TARGET), "Target"],
       ["=IF(SUM(tSales25[Sales]) >= SUM(tStores[Target_2025]), \"Met\", \"Missed\")", "Missed" if TOT < TARGET else "Met",
        "The verdict"]], "Text results need quotes; numbers and cell references do not."),
    F("2.8.24", "IFS", "=IFS(test1, value1, [test2, value2], …) returns the value for the first TRUE test.",
      "Several bands without nested IFs.",
      [["=SUM(tSales25[Profit]) / SUM(tSales25[Sales]) in B2", f"{MARGIN:.1%}", "Chain margin"],
       ["=IFS(B2 >= 0.15, \"Strong\", B2 >= 0.08, \"Healthy\", TRUE, \"Weak\")", "Healthy", "First TRUE test wins"]],
      "Order matters: test the highest band first. End with TRUE as the catch-all, or no match gives #N/A."),
    F("2.8.25", "AND", "=AND(logical1, [logical2], …) returns TRUE only if every test is TRUE.", "Several conditions must all hold.",
      [[f"Ikeja sales >= target ({N(IK['sales'])} vs {N(IK['target'])})", "TRUE", f"{IK['att']:.1%} of target"],
       [f"Ikeja margin > 10%", "TRUE", f"Margin {IK['margin']:.1%}"], ["=AND(test1, test2)", "TRUE", "Both met"]],
      "Inside IF: =IF(AND(…), \"Bonus\", \"\"). AND alone only returns TRUE or FALSE."),
    F("2.8.26", "OR", "=OR(logical1, [logical2], …) returns TRUE if at least one test is TRUE.", "Any one condition is enough.",
      [[f"Owerri attainment < 85%", "TRUE", f"{OW['att']:.1%} of target"], ["Owerri margin < 5%", "FALSE", f"Margin {OW['margin']:.1%}"],
       ["=OR(test1, test2)", "TRUE", "Review the store"]]),
    F("2.8.27", "NOT", "=NOT(logical) reverses TRUE and FALSE.", "Tests that read more naturally the other way round.",
      [["=COUNTIF(tSales25[Profit], \"<0\")=0", "FALSE", f"'No losses' is false ({LOSS} losses)"],
       ["=NOT(COUNTIF(tSales25[Profit], \"<0\")=0)", "TRUE", "Losses exist"]]),
    F("2.8.28", "IFERROR", "=IFERROR(value, value_if_error) returns value_if_error if value is any error; otherwise value.",
      "Replace an expected error with a message.",
      [["=SUMIFS(tSales25[Profit], tSales25[Store_ID], \"ST13\") / SUMIFS(tSales25[Sales], …, \"ST13\")", "#DIV/0!",
        "No such store: 0 ÷ 0"], ["=IFERROR(…, \"No sales\")", "No sales", "Readable result"]],
      "IFERROR hides every error, including typing mistakes. Prefer IFNA for lookups."),
    F("2.8.29", "IFNA", "=IFNA(value, value_if_na) replaces only the #N/A error.", "Lookups: 'not found' without hiding other errors.",
      [["=IFNA(VLOOKUP(\"ST01\", tStores[[Store_ID]:[Store_Manager]], 8, FALSE), \"Unknown store\")", IK["mgr"], "Found"],
       ["Same with \"ST13\"", "Unknown store", "Not found"]]),
    F("2.8.30", "SWITCH", "=SWITCH(expression, value1, result1, [value2, result2], …, [default]) compares one value with a list.",
      "Code-to-label conversions.",
      [["=SWITCH(LEFT([@Product_ID],3), \"FUR\", \"Furniture\", \"OFF\", \"Office Supplies\", \"TEC\", \"Technology\", \"Unknown\")",
        "Furniture", P1]],
      "SWITCH tests exact matches only. For ranges such as >= 0.15 use IFS."),
    F("2.8.31", "TEXT", "=TEXT(value, format_text) turns a number or date into text with a number format.",
      "Numbers inside sentences and labels.",
      [["=TEXT(SUM(tSales25[Sales]), \"₦#,##0\")", N(TOT), "Formatted total"],
       ["=\"FY2025 sales: \" & TEXT(SUM(tSales25[Sales]), \"₦#,##0\")", f"FY2025 sales: {N(TOT)}", "Live headline"],
       ["=TEXT(DATE(2025,12,31), \"dd mmm yyyy\")", "31 Dec 2025", "Date as text"]],
      "The result is text: it cannot be summed. Keep the number in its own cell."),
    F("2.8.32", "CONCAT", "=CONCAT(text1, [text2], …) joins text items; the & operator does the same.", "Labels from several columns.",
      [["=CONCAT([@Store_Name], \" (\", [@Store_ID], \")\")", f"{IK['name']} (ST01)", "Store label"],
       ["=[@Store_Name] & \" (\" & [@Store_ID] & \")\"", f"{IK['name']} (ST01)", "Same with &"]],
      "CONCAT has no separator argument. Use TEXTJOIN for lists."),
    F("2.8.33", "TEXTJOIN", "=TEXTJOIN(delimiter, ignore_empty, text1, …) joins text with a separator, optionally skipping blanks.",
      "A list in one cell.",
      [["=TEXTJOIN(\", \", TRUE, FILTER(tStores[Store_Name], tStores[Zone]=\"South-East\"))", ", ".join(SE_STORES),
        "South-East stores"]], head=["Formula", "Result", "What it does"], widths=[2.7, 2.4, 1.0]),
]

TOPICS_B = [
    F("2.8.34", "LEFT", "=LEFT(text, [num_chars]) returns the first num_chars characters.", "Codes at the start of an ID.",
      [[f"=LEFT(\"{P1}\", 3)", P1[:3], "Category code"], [f"=LEFT(\"{P3}\", 3)", P3[:3], ""]],
      "The result is text, even when it looks like a number."),
    F("2.8.35", "RIGHT", "=RIGHT(text, [num_chars]) returns the last num_chars characters.", "Numbers at the end of an ID.",
      [[f"=RIGHT(\"{O1}\", 6)", O1[-6:], "Order number (text)"], [f"=VALUE(RIGHT(\"{O1}\", 6))", f"{int(O1[-6:])}", "As a number"]]),
    F("2.8.36", "MID", "=MID(text, start_num, num_chars) returns num_chars characters starting at position start_num.",
      "Text in the middle of an ID.",
      [[f"=MID(\"{O1}\", 4, 4)", O1[3:7], "Year inside the Order_ID"], [f"=MID(\"{P1}\", 5, 2)", P1[4:6], "Sub-category code"]],
      "Count positions from 1: in SM-2025-…, S is 1, M is 2, the dash is 3, so the year starts at 4."),
    F("2.8.37", "LEN", "=LEN(text) returns the number of characters, spaces included.", "Detects extra spaces and wrong-length codes.",
      [[f"=LEN(\"{P1}\")", str(len(P1)), "Normal Product_ID"],
       ["=SUMPRODUCT(--(LEN(tDirty25[Zone]) <> LEN(TRIM(tDirty25[Zone]))))", I(DF["zone_space"]), "Zones with extra spaces"]]),
    F("2.8.38", "FIND", "=FIND(find_text, within_text, [start_num]) returns the position of find_text. Case-sensitive; no wildcards.",
      "Locates a delimiter so LEFT, MID or RIGHT can cut at it.",
      [[f"=FIND(\"-\", \"{P1}\")", str(P1.find("-") + 1), "First dash"],
       [f"=FIND(\"-\", \"{P1}\", FIND(\"-\", \"{P1}\") + 1)", str(P1.find("-", P1.find("-") + 1) + 1), "Second dash"],
       ["=FIND(\"chair\", \"Task Swivel Chair\")", "#VALUE!", "Case differs: not found"]]),
    F("2.8.39", "SEARCH", "=SEARCH(find_text, within_text, [start_num]) is FIND without case-sensitivity, and it allows wildcards.",
      "Keyword searches in names and descriptions.",
      [["=SEARCH(\"chair\", \"Office Star - Contemporary Task Swivel Chair\")",
        str("Office Star - Contemporary Task Swivel Chair".lower().find("chair") + 1), "Position"],
       ["=SUMPRODUCT(--ISNUMBER(SEARCH(\"chair\", tSales25[Product_Name])))", I(CHAIR_S), "Lines mentioning chair"],
       ["Same count with FIND", I(CHAIR_F), "Only lower-case 'chair'"]]),
    F("2.8.40", "SUBSTITUTE", "=SUBSTITUTE(text, old_text, new_text, [instance_num]) replaces text by matching it.",
      "Removing or swapping characters: ₦, commas, CHAR(160).",
      [["=SUBSTITUTE(\"₦239,240\", \"₦\", \"\")", "239,240", "Symbol removed"],
       ["=SUBSTITUTE([@Customer_Name], CHAR(160), \" \")", "Omolara Yusuf", "Web space fixed (line feed stays)"],
       ["Names with CHAR(160) in tDirty25", I(DF["nbsp"]), "Size of the problem"]]),
    F("2.8.41", "REPLACE", "=REPLACE(old_text, start_num, num_chars, new_text) replaces characters by position.",
      "Masking part of a code or phone number.",
      [["=REPLACE(\"07083197857\", 5, 4, \"****\")", "0708****857", "Masked phone"],
       ["=REPLACE(\"SM-2025-107503\", 4, 4, \"2026\")", "SM-2026-107503", "Year swapped"]],
      "REPLACE works by position, SUBSTITUTE by matching text."),
    F("2.8.42", "TRIM", "=TRIM(text) removes leading and trailing spaces and reduces inner runs of spaces to one.",
      "The first fix for any imported text.",
      [["=TRIM(\"  South-West \")", "South-West", "Spaces removed"],
       ["Distinct Zone values before / after TRIM", f"{DF['zone_distinct']} / {DF['zone_distinct_trim']}", "In tDirty25"]],
      "TRIM does not remove CHAR(160). Substitute it first."),
    F("2.8.43", "CLEAN", "=CLEAN(text) removes non-printing characters with codes 0-31, such as the line feed CHAR(10).",
      "Text from systems and web pages.",
      [["=LEN([@Customer_Name]) for 'Omolara Yusuf' + line feed", "14", "13 characters + line feed"],
       ["=LEN(CLEAN([@Customer_Name]))", "13", "Line feed removed"],
       ["Names with a line feed in tDirty25", I(DF["lf"]), "Size of the problem"]]),
    F("2.8.44", "PROPER", "=PROPER(text) capitalises the first letter of each word and lower-cases the rest.",
      "Names and places typed in any case.",
      [["=PROPER(\"lagos - ikeja\")", "Lagos - Ikeja", "Fixed case"],
       ["Distinct Store_Name values before / after PROPER", f"{STORE_RAW} / {STORE_PROPER}", "In tDirty25"]],
      "PROPER gets acronyms wrong: 'PORT HARCOURT - GRA' becomes 'Port Harcourt - Gra'. Check the result against the store list."),
    F("2.8.45", "UPPER", "=UPPER(text) converts text to capitals.", "Codes, and comparing text whatever its case.",
      [["=UPPER(TRIM(\"office supplies \"))", "OFFICE SUPPLIES", "Case and spaces"],
       ["Distinct Category before / after UPPER(TRIM())", f"12 / {CAT_UP}", "Misspellings remain"]],
      "Case fixes cannot fix spelling: Furnitre and Technolgy need a mapping table."),
    F("2.8.46", "LOWER", "=LOWER(text) converts text to lower case.", "Emails, usernames, case-free keys.",
      [[f"=LOWER(SUBSTITUTE(\"{IK['mgr']}\", \" \", \".\"))", IK["mgr"].lower().replace(" ", "."), "Login ID from a name"],
       ["=LOWER(\"LAGOS - IKEJA\") = LOWER(\"lagos - ikeja\")", "TRUE", "Case-free comparison"]]),
    F("2.8.47", "VALUE", "=VALUE(text) converts text that looks like a number, date or time into a number.",
      "The last step of fixing numbers stored as text.",
      [["=VALUE(SUBSTITUTE(SUBSTITUTE(\"₦239,240\", \"₦\", \"\"), \",\", \"\"))", "239240", "A real number"],
       ["Text Sales values in tDirty25", I(DF["sales_text"]), "Cells to convert"],
       ["Their total once converted", N(DF["text_sales_value"]), "Hidden from SUM before"]]),
    F("2.8.48", "TEXTBEFORE", "=TEXTBEFORE(text, delimiter, [instance_num], …) returns the text before a delimiter (Excel 365/2024).",
      "Splitting with a formula.",
      [["=TEXTBEFORE(\"Ikeja, Lagos\", \", \")", "Ikeja", "City"], [f"=TEXTBEFORE(\"{P1}\", \"-\")", P1.split("-")[0],
                                                                  "Category code"]]),
    F("2.8.49", "TEXTAFTER", "=TEXTAFTER(text, delimiter, [instance_num], …) returns the text after a delimiter. instance -1 "
      "means the last one.", "The second part, or the last word.",
      [["=TEXTAFTER(\"Uyo, Akwa Ibom\", \", \")", "Akwa Ibom", "State"], [f"=TEXTAFTER(\"{P1}\", \"-\", -1)", P1.split("-")[-1],
                                                                          "Item number"],
       ["=TEXTAFTER(\"Omolara Yusuf\", \" \", -1)", "Yusuf", "Surname"]]),
    F("2.8.50", "TODAY", "=TODAY() returns the current date. It recalculates whenever the workbook does.",
      "Ages, days since, days until.",
      [["=TODAY()", "Today's date", "Changes daily"], ["=TODAY() - DATE(2025,12,31)", "Days since FY end", "Format as a number"]],
      "Results change every day. Type the date (Ctrl+;) when the value must stay fixed."),
    F("2.8.51", "NOW", "=NOW() returns the current date and time.", "Time stamps on dashboards.",
      [["=\"Last refreshed: \" & TEXT(NOW(), \"dd-mmm-yyyy hh:mm\")", "Last refreshed: …", "Updates on recalculation"],
       ["=NOW() - TODAY()", "e.g. 0.6", "Time only, as a fraction of a day"]]),
    F("2.8.52", "DATE", "=DATE(year, month, day) builds a date from three numbers. Out-of-range months and days roll over.",
      "Repairing text dates safely, whatever the PC's regional settings.",
      [["=DATE(RIGHT(\"14/05/2025\",4), MID(\"14/05/2025\",4,2), LEFT(\"14/05/2025\",2))", "14-May-2025", "From dd/mm/yyyy"],
       ["=DATE(2025, 13, 1)", "01-Jan-2026", "Month 13 rolls over"],
       ["dd/mm/yyyy text dates in tDirty25", I(DF["d_slash"]), "To repair"]]),
    F("2.8.53", "DATEDIF", "=DATEDIF(start_date, end_date, unit) returns the gap in \"d\" days, \"m\" whole months or \"y\" whole years.",
      "Ages, tenure, length of stay.",
      [[f"=DATEDIF([@Admission_Date], [@Discharge_Date], \"d\")", str((A0.Discharge_Date - A0.Admission_Date).days),
        f"{A0.Admission_ID}"],
       [f"Same for {A1.Admission_ID}", str((A1.Discharge_Date - A1.Admission_Date).days), ""],
       [f"Average over all {len(A):,} stays in tAdm23", f"{LOS:.2f}", "Days"]],
      "DATEDIF does not appear in formula autocomplete: type it in full. A start_date after end_date gives #NUM!."),
    F("2.8.54", "EOMONTH", "=EOMONTH(start_date, months) returns the last day of the month, months before or after start_date.",
      "Month-end billing and reporting dates.",
      [[f"=EOMONTH(\"{d(A0.Discharge_Date)}\", 0)", d(A0.Discharge_Date + pd.offsets.MonthEnd(0)), "End of this month"],
       [f"=EOMONTH(\"{d(A0.Discharge_Date)}\", 1)", d(A0.Discharge_Date + pd.offsets.MonthEnd(1)), "End of next month"],
       [f"=EOMONTH(\"{d(A0.Discharge_Date)}\", -1) + 1", d(A0.Discharge_Date - pd.offsets.MonthBegin(1)), "First of this month"]]),
    F("2.8.55", "NETWORKDAYS", "=NETWORKDAYS(start_date, end_date, [holidays]) counts weekdays from start to end, both included, "
      "minus holidays.", "Working-day service levels.",
      [["=NETWORKDAYS(\"01-Jan-2025\", \"06-Jan-2025\")", str(NWD_1), "Wed to Mon"], ["Same, with 1 Jan as a holiday", str(NWD_1H), ""],
       ["Average for Standard Class lines", f"{NWD_STD:.2f}", "Working days, order to ship"]],
      "Both ends count: an order shipped the same weekday gives 1, not 0."),
    F("2.8.56", "WORKDAY", "=WORKDAY(start_date, days, [holidays]) returns the date a number of working days later.",
      "Due dates that skip weekends.",
      [[f"=WORKDAY(\"{d('2023-01-18')}\", 7)", wd("2023-01-18", 7), "Follow-up call"],
       [f"=WORKDAY(\"{d('2023-01-21')}\", 7)", wd("2023-01-21", 7), "From a Saturday"]],
      "List public holidays in a table and pass it as the third argument."),
    F("2.8.57", "VLOOKUP", "=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup]) finds a value in the first column "
      "and returns the value from column col_index_num. FALSE = exact match.", "Classic lookup, found in older files.",
      [["=VLOOKUP(\"ST09\", tStores[[Store_ID]:[Store_Manager]], 8, FALSE)", JO["mgr"], "Manager of ST09"],
       ["Same with \"ST02\"", LK["mgr"], ""], ["Same with \"ST13\"", "#N/A", "No such store"]],
      "Always use FALSE. Omitting it means approximate match, which returns wrong answers on unsorted data."),
    F("2.8.58", "XLOOKUP", "=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], …) finds a value and returns the "
      "matching item. Exact match by default.", "The modern lookup: any direction, built-in 'not found'.",
      [["=XLOOKUP(\"ST09\", tStores[Store_ID], tStores[Target_2025], \"not found\")", N(JO["target"]), "Target"],
       ["=XLOOKUP(\"ST09\", tStores[Store_ID], tStores[Sales_2025]) / (target)", P(JO["att"]), "Attainment"],
       ["Same with \"ST13\"", "not found", ""]]),
    F("2.8.59", "INDEX", "=INDEX(array, row_num, [column_num]) returns the value at a position.", "Paired with MATCH for flexible lookups.",
      [["=INDEX(tStores[City], 9)", JO["city"], "9th city"], ["=INDEX(tStores[City], MATCH(\"ST09\", tStores[Store_ID], 0))",
                                                               JO["city"], "INDEX-MATCH"]]),
    F("2.8.60", "MATCH", "=MATCH(lookup_value, lookup_array, [match_type]) returns the position of a value. 0 = exact.",
      "Finding where something is.",
      [["=MATCH(\"ST09\", tStores[Store_ID], 0)", "9", "Position"], ["=MATCH(\"ST13\", tStores[Store_ID], 0)", "#N/A", "Not in list"]],
      "Always type the 0: the default 1 assumes sorted data."),
    F("2.8.61", "FILTER", "=FILTER(array, include, [if_empty]) returns the rows where include is TRUE, spilling the result.",
      "Live extracts that update with the data (Excel 365/2021+).",
      [["=FILTER(tStores[[Store_Name]:[Store_Name]], tStores[Zone]=\"South-South\")", " · ".join(SS.Store_Name), "Stores"],
       ["=FILTER(tStores[Sales_2025], tStores[Zone]=\"South-South\")", " · ".join(N(v) for v in SS.Sales_2025), "Their sales"]],
      "Leave room below: if anything blocks the spill range you get #SPILL!.", widths=[2.6, 2.3, 0.9]),
    F("2.8.62", "UNIQUE", "=UNIQUE(array, [by_col], [exactly_once]) returns the distinct values.", "Lists and distinct counts.",
      [["=UNIQUE(tStores[Zone])", f"{ST.Zone.nunique()} zones spill", "Distinct list"],
       ["=ROWS(UNIQUE(tSales25[Customer_ID]))", I(len(CUST)), "Distinct customers"],
       ["=ROWS(UNIQUE(tSales25[Customer_ID],,TRUE))", I((CUST == 1).sum()), "Customers seen exactly once"]]),
    F("2.8.63", "SORT", "=SORT(array, [sort_index], [sort_order], [by_col]) returns the array sorted. -1 = descending.",
      "Live rankings.",
      [["=SORT(tStores[[Store_Name]:[Sales_2025]], 9, -1): row 1", f"{RANK.Store_Name.iloc[0]}  {N(RANK.Sales_2025.iloc[0])}", "Rank 1"],
       ["Row 2", f"{RANK.Store_Name.iloc[1]}  {N(RANK.Sales_2025.iloc[1])}", ""],
       ["Row 12", f"{RANK.Store_Name.iloc[-1]}  {N(RANK.Sales_2025.iloc[-1])}", "Last"]],
      "sort_index counts columns inside the array: Sales_2025 is the 9th column from Store_Name. To sort by something not shown, use SORTBY.", widths=[2.6, 2.4, 0.8]),
    F("2.8.64", "SORTBY", "=SORTBY(array, by_array1, [sort_order1], …) sorts an array by other arrays or calculations.",
      "Rank by a measure you do not display.",
      [["=SORTBY(tStores[Store_Name], tStores[Sales_2025] / tStores[Target_2025], -1)", ST.loc[ATT.index[0], "Store_Name"],
        f"1st: {ATT.iloc[0]:.1%}"],
       ["… second", ST.loc[ATT.index[1], "Store_Name"], f"{ATT.iloc[1]:.1%}"],
       ["… last", ST.loc[ATT.index[-1], "Store_Name"], f"{ATT.iloc[-1]:.1%}"]],
      "Ranked by attainment, Ikeja (largest sales) is not first.", widths=[3.3, 1.7, 1.0]),
    F("2.8.65", "LET", "=LET(name1, value1, [name2, value2], …, calculation) names intermediate results inside one formula.",
      "Long formulas that are readable, and calculate each part once.",
      [["=LET(s, SUM(tSales25[Sales]), p, SUM(tSales25[Profit]), p / s)", f"{MARGIN:.4f}", "Chain margin"],
       ["=LET(gap, SUM(tSales25[Sales]) - SUM(tStores[Target_2025]), IF(gap >= 0, \"Ahead by \", \"Behind by \") & "
        "TEXT(ABS(gap), \"₦#,##0\"))", f"Behind by {N(abs(TOT - TARGET))}", "Target gap"]],
      "Names cannot look like cell references (s1, ab12). Use words: sales, target, gap.", widths=[3.4, 1.6, 0.9]),
]

M28A = dict(
    id="2.8A", title="The 65 Analytical Excel Functions (1-33)",
    subtitle="Aggregation, counting, statistics, rounding, logic and joining text: functions 1 to 33, each on our sales data.",
    objectives=["Total, average and count with conditions.", "Aggregate visible rows with SUBTOTAL and AGGREGATE.",
                "Describe data with MEDIAN, MODE, MAX, MIN, LARGE and SMALL.", "Round correctly and use ABS.",
                "Decide with IF, IFS, AND, OR, NOT, IFERROR, IFNA, SWITCH; join text."],
    topics=TOPICS_A,
    takeaways=["SUMIFS, AVERAGEIFS and COUNTIFS cover most conditional summaries.",
               "SUBTOTAL 109 follows filters; AGGREGATE can skip errors.",
               "Report the median beside the average when values are skewed.",
               "ROUND changes the value; formatting only changes the display.",
               "IFS for bands, SWITCH for codes, IFNA for lookups; TEXT and TEXTJOIN for labels."],
)

M28B = dict(
    id="2.8B", title="Text, Date, Lookup and Dynamic Array Functions",
    subtitle="Functions 34 to 65: cutting and cleaning text, working with dates, looking values up, and spilling live lists.",
    objectives=["Cut text at positions and delimiters.", "Clean text and convert it to numbers.",
                "Build and measure dates and working days.", "Look up values with XLOOKUP, VLOOKUP and INDEX-MATCH.",
                "Spill live lists with FILTER, UNIQUE, SORT, SORTBY; simplify with LET."],
    topics=TOPICS_B,
    takeaways=["FIND or SEARCH locate; LEFT, MID, RIGHT cut; TEXTBEFORE and TEXTAFTER do both.",
               "Clean in order: SUBSTITUTE CHAR(160), CLEAN, TRIM, then fix case and convert with VALUE.",
               "DATE rebuilds text dates safely; NETWORKDAYS counts both ends.",
               "XLOOKUP with an if_not_found value; VLOOKUP always with FALSE.",
               "FILTER, UNIQUE, SORT and SORTBY spill; leave room; LET makes long formulas readable."],
)

MODULES = [M28A, M28B]
