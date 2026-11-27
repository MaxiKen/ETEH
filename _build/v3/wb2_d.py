"""Week 2 workbook, modules 2.9 - 2.11 (edition 3): the M language, PivotTables and PivotCharts.

2.9  M runs inside Power Query, not in cells. Each topic shows the M, an Excel twin where one exists (live formula), and asks
     for the result you get in the Power Query Editor. Those results come from w2_fx (computed from the practice files).
2.10 Real PivotTables on tSM are on sheet '2.10 Live Pivots' (Excel builds them when the file opens). Every pivot number used
     in a topic also has a formula twin (SUMIFS / COUNTIFS on tSM), so the learner can check the pivot against it.
2.11 Charts drawn from the formula-twin tables, so they display in any viewer, plus 'do it' tasks to insert real PivotCharts.
"""
from __future__ import annotations

import pandas as pd

import w2_m9_11 as M9
from w1x import ModuleSheet
from wb2_common import M_CRESTVIEW, M_LIFECARE, M_STORE_MONTH, N, code, lit_grid, serial
from w2_fx import CVF, LQ, SMF
from common import SEC

TOP = {t["tag"]: t for mod in M9.MODULES for t in mod["topics"]}
ZONES = ["South-West", "South-South", "North-Central", "South-East"]
YEARS = [2022, 2023, 2024, 2025]
LIVE = {}


def head(m, tag):
    t = TOP[tag]
    m.topic(tag, t["title"], " ".join(x for x in (t.get("defn"), t.get("simple"), t.get("like")) if x))
    return t


def watch(m, t):
    if t.get("watch"):
        m.note("Watch out: " + t["watch"])


def grid(m, t, widths=None):
    g = t.get("grid")
    if g:
        lit_grid(m, g["rows"][1:], header=g["rows"][0], merge_last=True)


def _lit(m, task, ans, alts=()):
    a = str(ans)
    extra = (f"{int(ans):,}",) if isinstance(ans, int) else ()
    m.ex_text(task, a, alts=tuple(alts) + extra)


def zsum(z=None, y=None, extra=""):
    f = "SUMIFS(tSM[Sales]"
    if z:
        f += f',tSM[Zone],"{z}"'
    if y:
        f += f",tSM[Year],{y}"
    return f + extra + ")"


# =====================================================================================================
def m29(b, D):
    S, A = D.S, D.A
    m = ModuleSheet(b, "2.9 M Language", "2.9  The M Language", "Tutor walk-through after Module 2.9: reading and writing the M "
                    "behind every Power Query step, on Lifecare, Crestview and SuperMart", "2.9.1 - 2.9.11")
    m.dataset([("Lifecare_Admissions_2023_DIRTY.csv", "Week2_Data_Files", f"{LQ['dedup'] + 63:,} raw rows. The main M example (M_LIFECARE below)."),
               ("Crestview_Student_Performance.csv", "00_Core_Datasets", f"{CVF['n']:,} students. Used for if … then … else (2.9.3, 2.9.11)."),
               ("tSales25", "Sheet Data_Sales25 (starts in A1)", "Excel twins of M functions, and the store-by-month Group By (2.9.10)."),
               ("tDirty25", "Sheet Data_Dirty25 (starts in A1)", "Excel's view of nulls and text (2.9.4).")])
    m.note("M runs in the Power Query Editor, not in cells. Where a topic asks for a Power Query result, build the query on the file named "
           "above and type the number you see; where Excel has a twin function, the worked solution shows it live.")

    t = head(m, "2.9.1")
    grid(m, t)
    m.turn()
    _lit(m, "Which M function turns the first row into column headers?", "Table.PromoteHeaders", ("promoteheaders",))
    m.ex_text("Which Excel function gives the same result as Text.Proper(\"chuKwuemEKa f. odili\")? Type the result of "
              "=PROPER(\"chuKwuemEKa f. odili\").", "Chukwuemeka F. Odili", formula='=PROPER("chuKwuemEKa f. odili")', short="PROPER twin")

    t = head(m, "2.9.2")
    grid(m, t)
    m.worked([("Excel compares text without case", '="ruby ward"="Ruby Ward"', True, None, "Excel: TRUE"),
              ("EXACT behaves like M", '=EXACT("ruby ward","Ruby Ward")', False, None, "M's = is case-sensitive, like EXACT.")])
    watch(m, t)
    m.turn()
    m.ex_text('In M, what does "ruby ward" = "Ruby Ward" return? (true / false)', "false")
    _lit(m, "In Power Query on the Lifecare file, filter Ward to exactly 'Ruby Ward' before the ward mapping. How many rows?", LQ["ruby_raw"])

    t = head(m, "2.9.3")
    grid(m, t)
    lit_grid(m, [[k, v] for k, v in CVF["bands"].items()], header=["Attendance_Band2", "Students"])
    m.turn()
    _lit(m, "Add the Band column of M_CRESTVIEW (below, 2.9.11). How many students are in the 90%+ band?", CVF["bands"]["90%+"])
    m.ex_text("Which keyword ends the list of steps and names the step to return?", "in")
    m.ex_text("Write the Excel twin of the band rule for a cell holding 82 (use IFS). What does it return?", "75-89%",
              formula='=IFS(82>=90,"90%+",82>=75,"75-89%",TRUE,"Below 75%")', short="IFS twin for 82")

    t = head(m, "2.9.4")
    grid(m, t)
    q = [v for v in D.DR.Quantity if v is not None]
    m.worked([("Excel's SUM skips empty cells, like List.Sum skips null", "=SUM(tDirty25[Quantity])", int(sum(q)), "int", ""),
              ("Cells it skipped", "=COUNTBLANK(tDirty25[Quantity])", sum(1 for v in D.DR.Quantity if v is None), "int", "")])
    watch(m, t)
    m.turn()
    m.ex_text("What does null + 1 return in M?", "null")
    m.ex_text("Which M type holds whole numbers: type number or Int64.Type?", "Int64.Type", alts=("int64",))

    t = head(m, "2.9.5")
    grid(m, t)
    m.steps(t["points"], label="Rules")
    m.turn()
    _lit(m, "In M_LIFECARE (2.9.7) how many rows does the Dedup step return?", LQ["dedup"])
    m.ex_text("A later step refers to a step you deleted. What does Power Query show?", "Error", alts=("an error", "expression.error"))

    t = head(m, "2.9.6")
    m.steps(["Power Query Editor › View › tick Formula Bar.", "Click the Dedup step in Applied Steps: the bar shows = Table.Distinct(Promoted).",
             "Click fx to add a step: Power Query writes = Dedup (a reference to the previous step).",
             "Type = Table.RowCount(Dedup) in the bar and press Enter: the preview becomes a single number."])
    m.do([("Show the formula bar and select Dedup.", "= Table.Distinct(Promoted)"),
          ("Add a step with fx and type = Table.RowCount(Dedup).", f"{LQ['dedup']:,}. Delete the step afterwards.")])

    t = head(m, "2.9.7")
    code(m, M_LIFECARE, "M_LIFECARE: the whole Lifecare cleaning query (Home › Advanced Editor › paste; create the DataFolder parameter first, 2.3.7)")
    grid(m, t)
    m.do([("Blank Query › Advanced Editor › paste M_LIFECARE › Done.", f"{LQ['clean']:,} rows load."),
          ("Display Options › tick Display line numbers.", "Each step on its own numbered line.")])
    m.turn()
    _lit(m, "Change the last line to 'in Exception' and filter Exception to non-null values. How many rows?", sum(LQ["exceptions"].values()))

    t = head(m, "2.9.8")
    grid(m, t)
    m.worked([("Excel MID starts at 1", '=MID("SM-2025-107503",4,4)', "2025", None, 'M: Text.Middle("SM-2025-107503", 3, 4)'),
              ("Number.Round twin", "=ROUND(0.12749,3)", 0.127, "0.000", ""),
              ("Date.EndOfMonth twin", "=EOMONTH(DATE(2024,2,10),0)", serial("2024-02-29"), "date", "")])
    watch(m, t)
    m.turn()
    m.ex_text('What does Text.Middle("SM-2025-107503", 8, 6) return? (Use the Excel twin =MID(…, 9, 6) to check.)', "107503",
              formula='=MID("SM-2025-107503",9,6)', short="Text.Middle offset 8")
    m.ex_text("Which family prefix holds the function that counts rows of a table?", "Table.", alts=("table",))

    t = head(m, "2.9.9")
    lit_grid(m, t["fx"], header=t["fx_head"], merge_last=True)
    grid(m, t)
    m.worked([("Excel twin of the Gender rule, on 'MALE '", '=IF(OR(LOWER(TRIM("MALE "))="m",LOWER(TRIM("MALE "))="male"),"Male","Female")',
               "Male", None, "")])
    m.turn()
    _lit(m, "After the Gender step of M_LIFECARE, how many rows are Female?", LQ["gender"]["Female"])
    _lit(m, "After the AdmType step, how many rows are Emergency?", LQ["atype"]["Emergency"])

    t = head(m, "2.9.10")
    lit_grid(m, t["fx"], header=t["fx_head"], merge_last=True)
    code(m, M_STORE_MONTH, "M_STORE_MONTH: store-by-month Group By (reference the clean query from 2.2.9)")
    m.worked([("Store-month rows (Excel twin)", '=ROWS(UNIQUE(tSales25[Store_Name]&TEXT(tSales25[Order_Date],"yyyy-mm")))', 144, "int", "12 × 12"),
              ("Abuja - Garki, Feb 2025 profit", '=SUMIFS(tSales25[Profit],tSales25[Store_Name],"Abuja - Garki",tSales25[Order_Date],">="&DATE(2025,2,1),'
               'tSales25[Order_Date],"<"&DATE(2025,3,1))', int(S[(S.Store_Name == "Abuja - Garki") & (S.Order_Date.dt.month == 2)].Profit.sum()), N, "")])
    ik = S[(S.Store_Name == "Lagos - Ikeja") & (S.Order_Date.dt.month == 12)].Sales.sum()
    m.turn()
    m.ex("What did Lagos - Ikeja sell in 2025-12? (Your Group By result, or the SUMIFS twin.)",
         '=SUMIFS(tSales25[Sales],tSales25[Store_Name],"Lagos - Ikeja",tSales25[Order_Date],">="&DATE(2025,12,1))', int(ik), N,
         short="Ikeja, December 2025")

    t = head(m, "2.9.11")
    grid(m, t)
    code(m, M_CRESTVIEW, "M_CRESTVIEW: bands and the mentor flag (needs a CoreFolder parameter pointing at 00_Core_Datasets)")
    m.turn()
    _lit(m, "Filter Needs_Mentor to true. How many students?", CVF["mentor"])
    m.ex_open("Write the M condition for 'needs a mentor' in your own words and in M.",
              "Exam score below 60 AND attendance below 75%:  each [Exam_Score] < 60 and [Attendance_Pct] < 75")
    m.finish()


# =====================================================================================================
def _zy_table(m, name, years, total=True):
    SM = m._SM
    df = pd.DataFrame({"Zone": ZONES})
    calc = {}
    for y in years:
        calc[f"Sales_{y}"] = (f'=SUMIFS(tSM[Sales],tSM[Zone],[@Zone],tSM[Year],{y})', "naira",
                              [float(SM[(SM.Zone == z) & (SM.Year == y)].Sales.sum()) for z in ZONES])
    if total:
        calc["Total"] = ('=SUMIFS(tSM[Sales],tSM[Zone],[@Zone])', "naira", [float(SM[SM.Zone == z].Sales.sum()) for z in ZONES])
    return m.s.table(df, name, calc=calc, widths={"Zone": 16})


def m210(b, D):
    S, SM, ST = D.S, D.SM, D.ST
    m = ModuleSheet(b, "2.10 PivotTables", "2.10  PivotTables", "Tutor walk-through after Module 2.10: build, group, filter and "
                    "format PivotTables on tSM, and check every number against a formula twin", "2.10.1 - 2.10.19")
    m._SM = SM
    m.dataset([("tSM", "Sheet Data_SM (starts in A1)", f"{len(SM):,} clean SuperMart lines 2022-2025: the source of every pivot."),
               ("Live pivots", "Sheet 2.10 Live Pivots", "Five real PivotTables on tSM. If one looks empty, press Data › Refresh All."),
               ("tSales25, tStores", "Data_Sales25, Data_Stores", "FY2025 lines and the 12 stores, for formula twins.")])
    ys = lambda z, y: int(SM[(SM.Zone == z) & (SM.Year == y)].Sales.sum())  # noqa: E731

    t = head(m, "2.10.1")
    m.h("Formula twin of Live Pivot 1 (Zone in Rows, Year in Columns, Sum of Sales)")
    _zy_table(m, "tZoneYear", YEARS)
    m.turn()
    m.ex("What were South-East sales in 2023? (Read Live Pivot 1 or tZoneYear.)", "=" + zsum("South-East", 2023), ys("South-East", 2023), N,
         short="South-East 2023")

    t = head(m, "2.10.2")
    grid(m, t)
    m.worked([("Grand total, one formula", "=SUM(tSM[Sales])", int(SM.Sales.sum()), N, "The pivot's Grand Total must match."),
              ("Sum of the twin table", "=SUM(tZoneYear[Total])", int(SM.Sales.sum()), N, "")])
    m.steps(t["points"], label="Why pivots")
    m.turn()
    m.ex("What is the 2025 column total of tZoneYear?", "=SUM(tZoneYear[Sales_2025])", int(SM[SM.Year == 2025].Sales.sum()), N, short="2025 total")

    t = head(m, "2.10.3")
    grid(m, t)
    m.worked([("Headers in tSM", "=COLUMNS(tSM[#Headers])", SM.shape[1], "int", ""),
              ("Sales stored as text", "=SUMPRODUCT(--ISTEXT(tSM[Sales]))", 0, "int", "0: safe to pivot"),
              ("Blank Order_IDs", "=COUNTBLANK(tSM[Order_ID])", 0, "int", "")])
    watch(m, t)
    m.turn()
    m.ex("How many Sales cells in tDirty25 are text, the reason a pivot on it would default to Count?", "=SUMPRODUCT(--ISTEXT(tDirty25[Sales]))",
         sum(1 for v in D.DR.Sales if isinstance(v, str)), "int", short="Text Sales in tDirty25")

    t = head(m, "2.10.4")
    lit_grid(m, t["diagram"][1]["table"][1:], header=t["diagram"][1]["table"][0])
    m.worked([(f"{z}, {c}, 2025", f'=SUMIFS(tSM[Sales],tSM[Zone],"{z}",tSM[Category],"{c}",tSM[Year],2025)',
               int(SM[(SM.Zone == z) & (SM.Category == c) & (SM.Year == 2025)].Sales.sum()), N, "")
              for z, c in [("South-West", "Technology"), ("South-East", "Furniture")]], label="Worked solution (formula twins)")
    m.turn()
    m.ex("What did North-Central sell in Office Supplies in 2025?", '=SUMIFS(tSM[Sales],tSM[Zone],"North-Central",tSM[Category],"Office Supplies",tSM[Year],2025)',
         int(SM[(SM.Zone == "North-Central") & (SM.Category == "Office Supplies") & (SM.Year == 2025)].Sales.sum()), N, short="NC Office Supplies 2025")

    t = head(m, "2.10.5")
    m.steps(t["steps"])
    grid(m, t)
    watch(m, t)
    m.do([("Click in tSM › Insert › PivotTable › New Worksheet › drag Zone to Rows, Sales to Values.", f"Grand Total ₦{int(SM.Sales.sum()):,}.")])
    m.turn()
    m.ex("What is Sum of Sales for South-West in your pivot?", "=" + zsum("South-West"), int(SM[SM.Zone == "South-West"].Sales.sum()), N,
         short="South-West, all years")

    t = head(m, "2.10.6")
    grid(m, t)
    m.worked([("Rows a 2022-2024 range covers", "=COUNTIFS(tSM[Year],\"<2025\")", SMF["fixed"][0], "int", ""),
              ("Sales a fixed range would see", "=SUMIFS(tSM[Sales],tSM[Year],\"<2025\")", int(SMF["fixed"][1]), N, "The table sees all.")])
    m.turn()
    m.ex("How many 2025 rows would a fixed 2022-2024 range miss?", "=COUNTIFS(tSM[Year],2025)", int((SM.Year == 2025).sum()), "int", short="Rows missed")

    t = head(m, "2.10.7")
    grid(m, t)
    m.steps(t["points"], label="Tips")
    m.do([("Live Pivot 3 shows Zone › Category. Drag Category above Zone in the field list.", "Category becomes the outer level.")])
    m.turn()
    m.ex("South-South Technology sales in 2025?", '=SUMIFS(tSM[Sales],tSM[Zone],"South-South",tSM[Category],"Technology",tSM[Year],2025)',
         int(SM[(SM.Zone == "South-South") & (SM.Category == "Technology") & (SM.Year == 2025)].Sales.sum()), N, short="SS Technology 2025")

    t = head(m, "2.10.8")
    grid(m, t)
    watch(m, t)
    m.turn()
    m.ex_text("Zone (4 items) or Store_Name (12 items): which belongs in Columns?", "Zone")
    m.ex("How many value columns would Store_Name × Year produce in Columns?", "=ROWS(UNIQUE(tSM[Store_Name]))*ROWS(UNIQUE(tSM[Year]))",
         SM.Store_Name.nunique() * 4, "int", short="Store × Year columns")

    t = head(m, "2.10.9")
    grid(m, t)
    m.worked([("Sum of Sales (2025)", "=SUM(tSales25[Sales])", int(S.Sales.sum()), N, ""),
              ("Count of Sales", "=COUNT(tSales25[Sales])", len(S), "int", ""),
              ("Average of Sales", "=AVERAGE(tSales25[Sales])", float(S.Sales.mean()), N, ""),
              ("Distinct Count of Order_ID", "=ROWS(UNIQUE(tSales25[Order_ID]))", int(S.Order_ID.nunique()), "int", "Data Model only in a pivot.")])
    watch(m, t)
    m.turn()
    m.ex("What would Max of Sales show for 2025?", "=MAX(tSales25[Sales])", int(S.Sales.max()), N, short="Max of Sales 2025")

    t = head(m, "2.10.10")
    grid(m, t)
    m.steps(t["points"], label="More")
    s25 = SM[SM.Year == 2025]
    m.do([("Live Pivot 2 has Year in Filters. Choose 2025.", f"Lagos - Ikeja ₦{int(s25[s25.Store_Name == 'Lagos - Ikeja'].Sales.sum()):,}.")])
    m.turn()
    m.ex("What profit did Jos - Rayfield make in 2025?", '=SUMIFS(tSM[Profit],tSM[Store_Name],"Jos - Rayfield",tSM[Year],2025)',
         int(s25[s25.Store_Name == "Jos - Rayfield"].Profit.sum()), N, short="Jos profit 2025")

    t = head(m, "2.10.11")
    grid(m, t)

    def q(n, y=2025):
        a, b_ = 3 * n - 2, 3 * n + 1
        e = f"DATE({y + (b_ > 12)},{(b_ - 1) % 12 + 1},1)"
        return f'=SUMIFS(tSM[Sales],tSM[Order_Date],">="&DATE({y},{a},1),tSM[Order_Date],"<"&{e})'
    m.worked([("Q1 2025 (twin)", q(1), SMF["q25"][1], N, ""), ("Q4 2025 (twin)", q(4), SMF["q25"][4], N, "")])
    m.do([("New pivot on tSM: Order_Date in Rows. Right-click a date › Group › Quarters and Years.", "Years › Quarters, Q4 largest each year.")])
    watch(m, t)
    m.turn()
    m.ex("What were Q3 2025 sales?", q(3), SMF["q25"][3], N, short="Q3 2025")

    t = head(m, "2.10.12")
    grid(m, t)
    m.worked([("Under ₦50k", '=COUNTIF(tSales25[Sales],"<50000")', int((S.Sales < 5e4).sum()), "int", ""),
              ("₦50k to 200k", '=COUNTIFS(tSales25[Sales],">=50000",tSales25[Sales],"<200000")', int(((S.Sales >= 5e4) & (S.Sales < 2e5)).sum()), "int", ""),
              ("₦200k to 1m", '=COUNTIFS(tSales25[Sales],">=200000",tSales25[Sales],"<1000000")', int(((S.Sales >= 2e5) & (S.Sales < 1e6)).sum()), "int", "")])
    m.turn()
    m.ex("How many FY2025 lines are ₦1m and above?", '=COUNTIF(tSales25[Sales],">=1000000")', int((S.Sales >= 1e6).sum()), "int", short="Lines ≥ ₦1m")

    t = head(m, "2.10.13")
    grid(m, t)
    m.worked([("Count of Sales, South-West 2025", '=COUNTIFS(tSM[Zone],"South-West",tSM[Year],2025)', int(SMF["zone_cnt"].loc["South-West", "count"]), "int",
               "A line count, not naira."),
              ("Sum of Sales, South-West 2025", "=" + zsum("South-West", 2025), ys("South-West", 2025), N, "")])
    watch(m, t)
    m.turn()
    m.ex("Count of Sales for South-East in 2025?", '=COUNTIFS(tSM[Zone],"South-East",tSM[Year],2025)', int(SMF["zone_cnt"].loc["South-East", "count"]),
         "int", short="Count SE 2025")
    m.ex_text("A pivot built on tDirty25 shows which summary of Sales by default: Sum or Count?", "Count")

    t = head(m, "2.10.14")
    grid(m, t)
    m.worked([("South-West: 2025 as % of row total", f"={zsum('South-West', 2025)}/{zsum('South-West')}", float(SMF["share25"]["South-West"]), "pct", ""),
              ("South-West: difference from 2024", f"={zsum('South-West', 2025)}-{zsum('South-West', 2024)}", int(SMF["diff25"]["South-West"]), N, "")])
    m.steps(t["points"], label="Settings")
    m.turn()
    m.ex("South-South: 2025 minus 2024?", f"={zsum('South-South', 2025)}-{zsum('South-South', 2024)}", int(SMF["diff25"]["South-South"]), N,
         short="SS difference from 2024")

    t = head(m, "2.10.15")
    grid(m, t)
    ik = int(S[S.Store_Name == "Lagos - Ikeja"].Sales.sum())
    m.worked([("₦#,##0", '=TEXT(SUMIFS(tSales25[Sales],tSales25[Store_Name],"Lagos - Ikeja"),"₦#,##0")', f"₦{ik:,}", None, ""),
              ("₦#,##0.0,,\"m\"", '=TEXT(SUMIFS(tSales25[Sales],tSales25[Store_Name],"Lagos - Ikeja"),"₦#,##0.0,,""m""")', f"₦{ik / 1e6:,.1f}m", None, "")])
    watch(m, t)
    m.turn()
    m.ex_text('Type the format code that shows 146388345 as ₦146.4m.', '₦#,##0.0,,"m"', alts=('₦0.0,,"m"',))

    t = head(m, "2.10.16")
    grid(m, t)
    pr = "SUMIFS(tSales25[Profit],tSales25[Store_Name],tStores[Store_Name])"
    sp = S.groupby("Store_Name").Profit.sum().sort_values()
    m.worked([("Top store by sales", "=INDEX(SORTBY(tStores[Store_Name],tStores[Sales_2025],-1),1)", ST.sort_values("Sales_2025").Store_Name.iloc[-1], None, ""),
              ("Biggest loss", f"=SMALL({pr},1)", int(sp.iloc[0]), N, ""),
              ("Its store", f"=INDEX(SORTBY(tStores[Store_Name],{pr},1),1)", sp.index[0], None, "")])
    m.turn()
    m.ex_text("Which store has the 2nd biggest loss in 2025?", sp.index[1], formula=f"=INDEX(SORTBY(tStores[Store_Name],{pr},1),2)",
              short="2nd biggest loss")

    t = head(m, "2.10.17")
    grid(m, t)
    m.steps(t["points"], label="How")
    watch(m, t)
    m.worked([("Jan-Nov lines", '=COUNTIF(tSales25[Order_Date],"<"&DATE(2025,12,1))', SMF["jan_nov"][0], "int", ""),
              ("Jan-Nov sales", '=SUMIFS(tSales25[Sales],tSales25[Order_Date],"<"&DATE(2025,12,1))', int(SMF["jan_nov"][1]), N, "")])
    m.turn()
    m.ex("How many December lines does Refresh All add?", '=COUNTIF(tSales25[Order_Date],">="&DATE(2025,12,1))', SMF["dec"][0], "int",
         short="December lines")

    t = head(m, "2.10.18")
    m.steps(t["steps"])
    grid(m, t)
    m.worked([("Sum ÷ sum (calculated field)", "=SUM(tSales25[Profit])/SUM(tSales25[Sales])", float(S.Profit.sum() / S.Sales.sum()), "pct", ""),
              ("Average of line margins (wrong)", "=AVERAGE(tSales25[Profit]/tSales25[Sales])", float(SMF["avg_line_margin"]), "pct", "")])
    watch(m, t)
    m.turn()
    x = S[S.Store_Name == "Lagos - Ikeja"]
    m.ex("What margin would the Margin calculated field show for Lagos - Ikeja?",
         '=SUMIFS(tSales25[Profit],tSales25[Store_Name],"Lagos - Ikeja")/SUMIFS(tSales25[Sales],tSales25[Store_Name],"Lagos - Ikeja")',
         float(x.Profit.sum() / x.Sales.sum()), "pct", short="Ikeja margin")

    t = head(m, "2.10.19")
    m.steps(t["steps"])
    grid(m, t)
    q4 = SM[(SM.Order_Date >= "2025-10-01") & (SM.Order_Date < "2026-01-01")]
    m.do([("Click Live Pivot 2 › Insert Slicer › Zone. Click South-West.", "Only South-West stores remain."),
          ("Insert Timeline › Order_Date › Quarters › select 2025 Q4.", "The pivot shows Q4 2025 only.")])
    m.turn()
    m.ex("With Zone = South-South and the timeline on Q4 2025, what is the total?",
         '=SUMIFS(tSM[Sales],tSM[Zone],"South-South",tSM[Order_Date],">="&DATE(2025,10,1),tSM[Order_Date],"<"&DATE(2026,1,1))',
         int(q4[q4.Zone == "South-South"].Sales.sum()), N, short="SS, Q4 2025")
    m.finish()


def live_sheet(b):
    s = b.sheet("2.10 Live Pivots", "2.10  Live PivotTables on tSM", "Excel builds these when the file opens. If one looks empty, press "
                "Data › Refresh All. Each one has a formula twin on sheet 2.10 PivotTables.", "2.10, 2.11")
    LIVE["sheet"] = s


PIVOTS = [("Pivot 1", "Sum of Sales by Zone (Rows) and Year (Columns). Twin: tZoneYear on 2.10 PivotTables.",
           dict(rows=["Zone"], columns=["Year"], values=[{"field": "Sales", "function": "sum", "num_format": "₦#,##0"}]), 10),
          ("Pivot 2", "Sales and Profit by store, Year in Filters (choose 2025). Twin: 2.10.10.",
           dict(rows=["Store_Name"], values=[{"field": "Sales", "function": "sum", "num_format": "₦#,##0"},
                                             {"field": "Profit", "function": "sum", "num_format": "₦#,##0"}], filters=["Year"]), 20),
          ("Pivot 3", "Zone › Category, Sum of Sales (all years). Twin: 2.10.7.",
           dict(rows=["Zone", "Category"], values=[{"field": "Sales", "function": "sum", "num_format": "₦#,##0"}]), 24),
          ("Pivot 4", "Count of Sales and Sum of Sales by Zone (the 2.10.13 trap, side by side).",
           dict(rows=["Zone"], values=[{"field": "Sales", "function": "count", "num_format": "#,##0", "name": "Count of Sales"},
                                       {"field": "Sales", "function": "sum", "num_format": "₦#,##0", "name": "Sum of Sales"}]), 10),
          ("Pivot 5", "Sum of Sales by Year: the source for a PivotChart (2.11.4).",
           dict(rows=["Year"], values=[{"field": "Sales", "function": "sum", "num_format": "₦#,##0"}]), 10)]


def after_data(b, D, tsm):
    s = LIVE["sheet"]
    for tag, what, spec, h in PIVOTS:
        s.band(tag, what)
        s.pivot(tsm, height=h, **spec)


# =====================================================================================================
def m211(b, D):
    SM, A = D.SM, D.A
    m = ModuleSheet(b, "2.11 PivotCharts", "2.11  PivotCharts", "Tutor walk-through after Module 2.11: charts from pivots, chart "
                    "types and filtering, with charts drawn from formula twins so they show in any viewer", "2.11.1 - 2.11.11")
    m._SM = SM
    m.dataset([("tSM", "Sheet Data_SM (starts in A1)", "Source of Live Pivots 1 and 5, and of every twin table here."),
               ("tAdm23", "Sheet Data_Adm23 (starts in A1)", "Lifecare 2023 admissions, for the bar chart in 2.11.11."),
               ("Crestview_Student_Performance.csv", "00_Core_Datasets", "Grades for the pie chart in 2.11.11.")])
    sw = SMF["sw_year"]

    t = head(m, "2.11.1")
    tz = _zy_table(m, "tZone2225", [2022, 2025], total=False)
    m.s.chart("column", [{"name": "2022", "categories": tz.ref("Zone"), "values": tz.ref("Sales_2022")},
                         {"name": "2025", "categories": tz.ref("Zone"), "values": tz.ref("Sales_2025")}],
              "Sales by zone, 2022 and 2025", y_fmt='₦#,##0,,"m"', rows=17)
    m.turn()
    m.ex("By how much did South-West sales grow from 2022 to 2025?", f"={zsum('South-West', 2025)}-{zsum('South-West', 2022)}",
         int(sw[2025] - sw[2022]), N, short="SW growth 2022-2025")

    t = head(m, "2.11.2")
    q4 = SMF["q4"]
    df = pd.DataFrame({"Year": YEARS})
    tq = m.s.table(df, "tQ4Share", fmts={"Year": "gen"}, calc={
        "Q4_Sales": ('=SUMIFS(tSM[Sales],tSM[Year],[@Year],tSM[Order_Date],">="&DATE([@Year],10,1))', "naira", [float(q4.loc[y, "Q4"]) for y in YEARS]),
        "Q4_Share": ('=[@Q4_Sales]/SUMIFS(tSM[Sales],tSM[Year],[@Year])', "pct", [float(q4.loc[y, "share"]) for y in YEARS])})
    m.s.chart("column", [{"name": "Q4 sales", "categories": tq.ref("Year"), "values": tq.ref("Q4_Sales"), "labels": True, "label_fmt": '₦#,##0.0,,"m"'}],
              "Q4 sales by year", legend=False, y_fmt='₦#,##0,,"m"', rows=16)
    m.steps(t["points"], label="What the chart shows")
    m.turn()
    m.ex("What share of 2023 sales came in Q4?", "=XLOOKUP(2023,tQ4Share[Year],tQ4Share[Q4_Share])", float(q4.loc[2023, "share"]), "pct", short="Q4 share 2023")

    t = head(m, "2.11.3")
    grid(m, t)
    m.turn()
    m.ex_text("Can a PivotChart be a scatter chart? (Yes / No)", "No")

    t = head(m, "2.11.4")
    m.steps(t["steps"])
    grid(m, t)
    m.do([("Click in Live Pivot 1 › PivotTable Analyze › PivotChart › Clustered Column.", "16 columns: 4 zones × 4 years."),
          ("Click in Live Pivot 5 › PivotChart › Line.", "One line rising from 2023 to 2025.")])

    t = head(m, "2.11.5")
    grid(m, t)
    m.turn()
    m.ex("How many bars does the Zone × Year PivotChart draw?", "=ROWS(UNIQUE(tSM[Zone]))*ROWS(UNIQUE(tSM[Year]))", 16, "int", short="Bars")

    t = head(m, "2.11.6")
    grid(m, t)
    watch(m, t)
    m.turn()
    m.ex_text("A PivotChart's Legend (Series) is which pivot area?", "Columns")

    t = head(m, "2.11.7")
    m.worked([(f"South-West {y}", "=" + zsum("South-West", y), int(sw[y]), N, "") for y in YEARS], label="Worked solution (what the pivot shows)")
    m.steps(t["points"], label="What happens")
    m.do([("On the Zone × Year PivotChart, click the Zone field button › tick South-West only.", "Live Pivot 1 filters to South-West too.")])
    m.turn()
    m.ex("South-West sales in 2023?", "=" + zsum("South-West", 2023), int(sw[2023]), N, short="SW 2023")

    t = head(m, "2.11.8")
    grid(m, t)
    watch(m, t)
    m.turn()
    m.ex("What should the value axis minimum of a column chart be?", "=0", 0, "int", short="Axis minimum")

    t = head(m, "2.11.9")
    grid(m, t)
    m.turn()
    m.ex_text("Which chart type suits 12 months in sequence?", "Line")
    m.ex_text("Which chart suits 12 stores' share of sales: pie or sorted bar?", "Sorted bar", alts=("bar",))

    t = head(m, "2.11.10")
    m.steps(t["steps"])
    z25 = SM[SM.Year == 2025].groupby("Zone").Sales.sum().sort_values(ascending=False)
    sb = lambda i: f"=INDEX(SORTBY(UNIQUE(tSM[Zone]),SUMIFS(tSM[Sales],tSM[Zone],UNIQUE(tSM[Zone]),tSM[Year],2025),-1),{i})"  # noqa: E731
    m.worked([("1st zone, 2025", sb(1), z25.index[0], None, ""), ("Last zone, 2025", sb(4), z25.index[3], None, "")])
    m.turn()
    m.ex_text("After sorting Z to A by Sum of Sales (2025), which zone is 3rd?", z25.index[2], formula=sb(3), short="3rd zone 2025")

    t = head(m, "2.11.11")
    grid(m, t)
    dep = A.Department.value_counts()
    td = m.s.table(pd.DataFrame({"Department": list(dep.index)}), "tDeptChart", widths={"Department": 28},
                   calc={"Admissions": ("=COUNTIF(tAdm23[Department],[@Department])", "int", [int(v) for v in dep.values])})
    m.s.chart("bar", [{"name": "Admissions", "categories": td.ref("Department"), "values": td.ref("Admissions"), "labels": True, "label_fmt": "#,##0"}],
              "Lifecare admissions by department, 2023", legend=False, rows=15)
    cv = pd.read_csv(SEC / "00_Core_Datasets" / "Crestview_Student_Performance.csv", usecols=["Grade"])
    gr = cv.Grade.value_counts().reindex(["A", "B", "C"]).reset_index()
    gr.columns = ["Grade", "Students"]
    tg = m.s.table(gr, "tGradeChart", fmts={"Students": "int"})
    m.note("Grade counts come from Crestview_Student_Performance.csv (a pivot on that file in Power Query or Excel gives the same).")
    m.s.chart("pie", [{"name": "Students", "categories": tg.ref("Grade"), "values": tg.ref("Students"), "labels": True,
                       "points": [{"fill": {"color": c}} for c in ["#13807A", "#0B2545", "#E0A526"]]}],
              "Crestview grades: B dominates", rows=15, size=(480, 290))
    m.turn()
    m.ex("How many 2023 admissions did Endocrinology have?", '=COUNTIF(tAdm23[Department],"Endocrinology")', int(dep["Endocrinology"]), "int",
         short="Endocrinology admissions")
    m.ex("What share of Crestview students got a B?", '=XLOOKUP("B",tGradeChart[Grade],tGradeChart[Students])/SUM(tGradeChart[Students])',
         float(CVF["grades"]["B"]), "pct", short="Share with grade B")
    m.finish()


def build(b, D):
    m29(b, D)
    m210(b, D)
    live_sheet(b)
    m211(b, D)
