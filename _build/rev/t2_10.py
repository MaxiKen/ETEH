from revlib import T

W = "SuperMart Retail"
TRY = "ETEH_Week2_Section_Examples.xlsx  ›  sheets '2.10 PivotTables' and '2.10 Live Pivots'"
TSM = "tSM (all 9,994 SuperMart lines 2022–2025, sheet D_SuperMart_All)"


def P(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", W)
    return T(tag, **k)


TOPICS = [
    P("2.10.1",
      what="A **PivotTable** summarises a long list in seconds. You choose what to **group by** (zones, years) and "
           "what to **add up** (sales), and Excel builds the summary table. Drag a field to a new place and the "
           "summary redraws instantly. Your data is never changed.",
      like="A PivotTable is like sorting a big pile of receipts into labelled trays (one per zone and year), then "
           "writing the total on each tray. Want trays by store instead? Just re-sort.",
      how=["Click inside your data table.", "Insert > PivotTable.", "Drag fields into Rows, Columns and Values."],
      tip="Build PivotTables on clean Tables, never on messy ranges.",
      trap="Thinking a PivotTable changes your data; it is only a view of it.",
      case="Head Office wants total sales by zone for each year, from " + TSM + ".",
      steps=["Insert a PivotTable from tSM.", "Drag Zone to Rows and Year to Columns.", "Drag Sales to Values."],
      table=[["Zone", "2022", "2025", "Grand Total"], ["South-West", "₦147.9m", "₦250.1m", "₦725.5m"],
             ["South-South", "₦128.7m", "₦213.1m", "₦678.8m"], ["North-Central", "₦103.8m", "₦147.1m", "₦501.2m"],
             ["South-East", "₦103.8m", "₦122.9m", "₦391.7m"],
             ["All zones", "₦484.2m", "₦733.2m", "₦2,297.2m"]],
      so="9,994 lines became one small table in three drags. South-West grew from ₦147.9m to ₦250.1m.",
      split=4.20),

    P("2.10.2",
      what="PivotTables let you answer **new questions in seconds**: change a field, add a filter, switch the "
           "measure. Doing the same with formulas means writing and rewriting many SUMIFS. Each new 'cut' of the "
           "data is just a drag.",
      like="Formulas are like cooking each dish from scratch; a PivotTable is like a buffet, where you just pick "
           "what you want on the plate.",
      how=["Write down the business question.", "Map it: 'by what?' goes to Rows/Columns, 'how much?' to Values.",
           "Change fields to answer follow-up questions."],
      tip="Use formulas to check a PivotTable's numbers once, then trust the pivot.",
      trap="Copying pivot numbers into another sheet as fixed values; they won't update.",
      case="The sales zone × year summary on the 2.10 sheet was built twice: once with SUMIFS formulas and once as "
           "a live pivot.",
      steps=["Count the formulas in the formula version.", "Count the drags in the pivot version.",
             "Check both grand totals match."],
      table=[["Approach", "Work needed", "Grand total"], ["SUMIFS grid", "16 SUMIFS + 9 SUM formulas", "₦2,297,200,862"],
             ["PivotTable", "3 drags", "₦2,297,200,862"]],
      so="Same answer, far less work. And a new question (by store?) is one more drag, not 16 new formulas.",
      split=4.20),

    P("2.10.3",
      what="A PivotTable needs tidy data: **one header row** with unique names, **one record per row**, no blank "
           "rows or merged cells, **one data type per column**, and **no total rows** inside the data. Best of all, "
           "make the source an **Excel Table**.",
      like="A PivotTable is like a sorting machine at a post office: it works only if every letter has an address "
           "in the same place, and nothing is taped together.",
      how=["Check headers: one row, no blanks, no duplicates.", "Remove blank rows, merged cells and total rows.",
           "Convert the range to a Table (Ctrl+T)."],
      tip="Clean with Power Query first, then pivot the clean Table.",
      trap="Leaving a 'Total' row in the data, so the pivot counts everything twice.",
      case="Before pivoting, the analyst checks that " + TSM + " meets every requirement.",
      steps=["Count the headers and look for blanks.", "Check Sales is numeric and dates are real dates.",
             "Confirm the source is the Table tSM."],
      table=[["Requirement", "tSM status"], ["One header row, unique names", "24 headers, none blank"],
             ["No blank rows or columns", "Cleaned in 2.5 and 2.6"], ["Consistent types", "Sales numeric, dates real"],
             ["Source is an Excel Table", "Yes: tSM"]],
      so="tSM passes every check, so the pivot's totals can be trusted.",
      split=4.90),

    P("2.10.4",
      what="Every PivotTable has **four areas**. **Rows** groups down the side. **Columns** groups across the top. "
           "**Values** holds the numbers being summarised. **Filters** limits the whole table to chosen items.",
      like="The four areas are like a school timetable: days down the side (Rows), periods across the top "
           "(Columns), subjects in the boxes (Values), and the class name at the top (Filter).",
      how=["Ask 'by what?': put that field in Rows (and maybe Columns).", "Ask 'how much?': put that in Values.",
           "Ask 'only for which?': put that in Filters."],
      tip="Put the field with more items in Rows; fewer items in Columns.",
      trap="Putting a measure like Sales in Rows, which lists every amount.",
      case="Two live pivots on the 2.10 Live Pivots sheet use all four areas.",
      steps=["Pivot 1: Rows = Zone, Columns = Year, Values = Sum of Sales.",
             "Pivot 2: Rows = Store_Name, Values = Sales and Profit, Filter = Year.",
             "Set Pivot 2's filter to 2025."],
      table=[["Area", "Pivot 1", "Pivot 2"], ["Rows", "Zone", "Store_Name"], ["Columns", "Year", "(none)"],
             ["Values", "Sum of Sales", "Sales, Profit"], ["Filters", "(none)", "Year = 2025"]],
      so="Rows and Columns set the shape, Values the numbers, and Filters the scope.",
      split=4.20),

    P("2.10.5",
      what="To create a PivotTable: click any cell in a Table, choose **Insert > PivotTable**, pick **New "
           "Worksheet**, and click OK. Give the new sheet a clear name, like PT_Sales_By_Zone, so others can find it.",
      like="Creating a pivot is like setting up an empty notice board: first the board, then you pin the items "
           "where you want them.",
      how=["Click inside the Table.", "Insert > PivotTable > New Worksheet > OK.", "Rename the sheet to describe the pivot."],
      tip="Name pivot sheets clearly: PT_ followed by what they show.",
      trap="Selecting only part of the data before inserting, so the pivot misses rows.",
      case="The analyst builds a first pivot of sales by zone from " + TSM + ".",
      steps=["Click any cell in tSM; Insert > PivotTable.", "Choose New Worksheet; rename it PT_Sales_By_Zone.",
             "Drag Zone to Rows and Sales to Values."],
      table=[["Row Labels", "Sum of Sales"], ["North-Central", "₦501,239,892"], ["South-East", "₦391,721,905"],
             ["South-South", "₦678,781,240"], ["South-West", "₦725,457,825"]],
      so="A first pivot in under a minute: South-West is the largest zone across 2022–2025."),

    P("2.10.6",
      what="A pivot's source can be a **fixed range**, an **Excel Table** or the **Data Model**. A fixed range "
           "ignores new rows. A Table includes them on refresh. The Data Model also links several tables and "
           "offers Distinct Count.",
      like="A fixed range is like a printed guest list; a Table is like a sign-up sheet that grows as people "
           "arrive.",
      how=["Convert the data to a Table (Ctrl+T) first.", "Build the pivot from the Table name.",
           "Tick 'Add this data to the Data Model' if you need Distinct Count or relationships."],
      tip="Always build on a Table, so new rows are included.",
      trap="Building on A1:X6683 and missing every row added later.",
      case="A pivot was built on a fixed range covering 2022–2024. Then the 2025 lines were added below it.",
      steps=["Refresh the fixed-range pivot and check the total.", "Rebuild it on the Table tSM.",
             "Refresh and compare."],
      table=[["Source", "Rows seen", "Total sales"], ["Fixed range (2022–2024 only)", "6,682", "₦1,564.0m"],
             ["Table tSM", "9,994", "₦2,297.2m"]],
      so="The fixed range silently missed all 3,312 lines of 2025 (₦733.2m). A Table would not.",
      split=4.20),

    P("2.10.7",
      what="Fields in **Rows** list items down the side. Add a second row field and it **nests** under the first, "
           "e.g. Zone, then Category within each zone. You can expand and collapse each group.",
      like="Nested rows are like folders inside folders: open 'South-West' to see its Furniture, Office Supplies "
           "and Technology.",
      how=["Drag the main field to Rows.", "Drag a second field below it.", "Use the +/− buttons to expand or collapse."],
      tip="Put the bigger group first (Zone before Category).",
      trap="Nesting too many fields, so the table becomes too long to read.",
      case="Management asks for FY2025 sales by zone, broken down by category.",
      steps=["Filter Year = 2025.", "Drag Zone, then Category, to Rows.", "Drag Sales to Values; expand South-West."],
      table=[["South-West (FY2025)", "Sum of Sales"], ["Furniture", "₦71,705,712"], ["Office Supplies", "₦82,463,506"],
             ["Technology", "₦95,959,148"], ["South-West total", "₦250,128,366"]],
      so="Inside South-West, Technology leads. This is the same ₦95.96m found with SUMIFS in 2.8.3.",
      split=4.90),

    P("2.10.8",
      what="Fields in **Columns** spread the values across the top, making a grid. One or two column fields work "
           "well. Too many columns make a grid nobody can read.",
      like="Columns are like the dates across the top of a calendar: 7 days fit nicely; 365 across one line would "
           "not.",
      how=["Put a field with few items in Columns (Year, Quarter).", "Keep the field with many items in Rows.",
           "Check the pivot still fits on one screen."],
      tip="Keep the columns to about 12 or fewer.",
      trap="Putting Month (48 items) or Store × Month in Columns.",
      case="The analyst compares two layouts for " + TSM + ".",
      steps=["Put Year in Columns: count the columns.", "Swap to Month in Columns: count again.",
             "Decide which layout is readable."],
      table=[["Field in Columns", "Value columns", "Verdict"], ["Year", "4", "Easy to read"],
             ["Quarter (within Year)", "16", "Readable with care"], ["Month (all years)", "48", "Too wide"]],
      so="Year in Columns gives a tidy 4-column grid. Long lists belong in Rows.",
      split=4.20),

    P("2.10.9",
      what="The **Values** area summarises numbers. Right-click > **Value Field Settings** to choose **Sum, Count, "
           "Average, Max, Min** and more. **Distinct Count** is available only when the pivot uses the Data Model.",
      like="Value settings are like asking different questions about the same jar of coins: how much in total? "
           "how many coins? what's the average coin?",
      how=["Drag the measure to Values.", "Right-click > Value Field Settings.", "Choose the summary and rename the field."],
      tip="Rename value fields (e.g. 'Lines' instead of 'Count of Sales').",
      trap="Counting lines and calling them orders; one order can have many lines.",
      case="For FY2025, the analyst puts Sales in Values four times, with a different summary each time.",
      steps=["Filter Year = 2025.", "Add Sales to Values four times.", "Set Sum, Count, Average; use Distinct Count on Order_ID."],
      table=[["Summary", "FY2025 result"], ["Sum of Sales", "₦733,215,255"], ["Count of Sales (lines)", "3,312"],
             ["Average of Sales", "₦221,381"], ["Distinct Count of Order_ID", "1,687 orders"]],
      so="The same field answers different questions. Lines (3,312) and orders (1,687) are not the same.",
      split=4.90),

    P("2.10.10",
      what="A **Report Filter** (the Filters area) limits the **whole** PivotTable to chosen items, such as one "
           "year. It appears above the pivot. Tick 'Select Multiple Items' to choose more than one.",
      like="A report filter is like choosing which year's photo album to open: everything you see inside belongs "
           "to that year.",
      how=["Drag the field (e.g. Year) to Filters.", "Click the drop-down above the pivot.", "Pick one or more items."],
      tip="Show the filter clearly, so readers know what they're looking at.",
      trap="Forgetting a filter is on and reporting a partial total as the full one.",
      case="Pivot 2 on the Live Pivots sheet shows sales and profit by store, with a Year filter.",
      steps=["Drag Year to Filters.", "Choose 2025.", "Sort stores by sales."],
      table=[["Store (Year = 2025)", "Sales", "Profit"], ["Lagos - Ikeja", "₦146,388,345", "₦29,366,453"],
             ["Port Harcourt - GRA", "₦93,922,995", "₦24,357,072"], ["Lagos - Lekki", "₦70,819,092", "₦18,561,342"]],
      so="With the filter set to 2025, Ikeja leads both sales and profit.",
      split=4.20),

    P("2.10.11",
      what="Right-click a date field > **Group** to group dates into **Years, Quarters and Months**. Thousands of "
           "separate dates become a short list you can read and drill into.",
      like="Grouping dates is like putting daily receipts into monthly envelopes, and monthly envelopes into "
           "yearly boxes.",
      how=["Put Order_Date in Rows.", "Right-click a date > Group.", "Choose Years and Quarters (or Months)."],
      tip="Always group by Years too, or Q1 2022 and Q1 2025 get added together.",
      trap="Grouping by Months only, which mixes January of every year.",
      case="Head Office wants FY2025 sales by quarter from " + TSM + ".",
      steps=["Put Order_Date in Rows (1,237 distinct dates).", "Group by Years and Quarters.", "Read the 2025 rows."],
      table=[["2025 quarter", "Sales"], ["Q1", "₦123,144,860"], ["Q2", "₦133,764,372"], ["Q3", "₦196,251,956"],
             ["Q4", "₦280,054,067"]],
      so="Q4 is the peak (38% of 2025 sales), and it is the peak in every year from 2022 to 2025.",
      split=4.90),

    P("2.10.12",
      what="Numbers can be grouped into **bands**: right-click > **Group**, then set Starting at, Ending at and By. "
           "For uneven bands (₦0–50k, ₦50k–200k…), add a band column in the data or Power Query instead.",
      like="Number bands are like age groups at a sports day: under-10s, 10–14, 15 and over. Each child goes into "
           "one group.",
      how=["Equal bands: put the field in Rows; right-click > Group > By.", "Uneven bands: add a band column first.",
           "Count or sum within each band."],
      tip="Use a band column in the data for bands that should never change.",
      trap="Overlapping band edges, so a value of exactly ₦50,000 is counted twice.",
      case="The analyst groups FY2025 sales lines by size.",
      steps=["Add a band: under ₦50k, ₦50k–200k, ₦200k–1m, ₦1m+.", "Put the band in Rows.", "Count lines per band."],
      table=[["Line size (FY2025)", "Lines"], ["Under ₦50k", "1,616"], ["₦50k–200k", "858"], ["₦200k–1m", "691"],
             ["₦1m and above", "147"]],
      so="Almost half of all lines are under ₦50k, but just 147 lines are ₦1m or more.",
      split=4.90),

    P("2.10.13",
      what="If a Values column contains **any text or blanks**, Excel may summarise it as **Count** instead of "
           "**Sum**. The pivot then shows how many rows there are, not how much money. Always check the label: "
           "'Count of Sales' is a warning sign.",
      like="It's like asking for the total value of a pile of receipts and being told '518': someone counted the "
           "receipts instead of adding them.",
      how=["Read the value field label: Sum of… or Count of…?", "If it says Count, check the column for text.",
           "Fix the data, refresh, and switch to Sum."],
      tip="Check every new value field says 'Sum of'.",
      trap="Reporting a count of rows as if it were money.",
      case="The raw sales file has 234 Sales values stored as text, like '₦1,234,560'. The pivot defaults to Count.",
      steps=["Put Zone in Rows and Sales in Values.", "Notice 'Count of Sales'.", "Fix the text values; switch to Sum."],
      table=[["Zone (FY2025)", "Count of Sales", "Sum of Sales"], ["North-Central", "778", "₦147,098,128"],
             ["South-East", "518", "₦122,905,857"], ["South-South", "921", "₦213,082,904"],
             ["South-West", "1,095", "₦250,128,366"]],
      so="'518' is a count of lines, not money. Only Sum gives the real ₦122.9m.",
      split=4.20),

    P("2.10.14",
      what="**Show Values As** changes how a number is shown without changing the data: **% of Grand Total**, "
           "**% of Row Total**, **Difference From** (e.g. previous year), **Running Total** and more.",
      like="Show Values As is like describing a test score as '72', '72% of the top mark' or '8 better than last "
           "term': same result, different view.",
      how=["Right-click a value > Show Values As.", "Pick the view, e.g. % of Row Total.",
           "For Difference From, choose the base field and item (Year, previous)."],
      tip="Add the same field twice: one as an amount, one as a %.",
      trap="Forgetting which total the % refers to (row, column or grand).",
      case="Management wants each zone's sales shown as a share of its 4-year total, and as change from last year.",
      steps=["Copy the Zone × Year pivot.", "Show Values As > % of Row Total.",
             "Add Sales again: Difference From > Year > (previous)."],
      table=[["Zone", "2025 % of zone total", "2025 vs 2024"], ["South-West", "34.5%", "+₦62.6m"],
             ["South-East", "31.4%", "+₦29.3m"], ["South-South", "31.4%", "+₦32.4m"],
             ["North-Central", "29.3%", "−₦0.3m"]],
      so="Every zone except North-Central grew in 2025; North-Central was flat (−₦0.3m).",
      split=4.20),

    P("2.10.15",
      what="Format pivot numbers through **Value Field Settings > Number Format**, not by formatting cells. Field "
           "formatting stays when the pivot is refreshed or rearranged. Use #,##0 for money and 0.0% for rates.",
      like="Field formatting is like printing prices on the product packaging rather than on a sticker: the "
           "sticker falls off when things move.",
      how=["Right-click a value > Value Field Settings.", "Click Number Format.", "Choose the format, e.g. #,##0."],
      tip="Format in Value Field Settings, not with the ribbon.",
      trap="Formatting cells directly and losing it on the next refresh.",
      case="The analyst formats the FY2025 store pivot for a board report.",
      steps=["Sales: Number Format ₦#,##0.", "Margin: 0.0%.", "Refresh and rearrange; check the formats stay."],
      table=[["Field", "Format code", "Shows"], ["Sum of Sales", "₦#,##0", "₦146,388,345"], ["Margin", "0.0%", "20.1%"],
             ["Lines", "#,##0", "3,312"]],
      so="Formats set on the field survive refreshes, so the report always looks right.",
      split=4.20),

    P("2.10.16",
      what="Right-click a value > **Sort Largest to Smallest** to rank items by their total, not alphabetically. "
           "**Value Filters** like **Top 10** show only the leaders; **Label Filters** filter by name.",
      like="Sorting a pivot is like arranging a football league table by points, not by team name.",
      how=["Right-click a value > Sort > Largest to Smallest.", "Row label drop-down > Value Filters > Top 10.",
           "Change 10 to the number you want."],
      tip="Sort by value to show who matters most.",
      trap="Sorting the row labels alphabetically and calling it a ranking.",
      case="Management wants the top FY2025 stores by sales, and the stores with the biggest losses.",
      steps=["Filter Year = 2025; Rows = Store_Name; Values = Sales, Profit.",
             "Sort by Sales: Top 3 value filter.", "Then sort Profit smallest to largest."],
      table=[["Ranking (FY2025)", "Store", "Value"], ["Top sales", "Lagos - Ikeja", "₦146.4m"],
             ["2nd sales", "Port Harcourt - GRA", "₦93.9m"], ["Biggest loss", "Abuja - Wuse", "−₦8.8m"],
             ["2nd biggest loss", "Onitsha - Main Market", "−₦6.2m"]],
      so="Sorting reveals both stories: who sells most, and where profit leaks.",
      split=4.20),

    P("2.10.17",
      what="A PivotTable does **not** update when its data changes. Use **Data > Refresh All** (Ctrl+Alt+F5) or "
           "right-click > Refresh. Tick **PivotTable Options > Data > Refresh data when opening the file** for "
           "shared reports.",
      like="A pivot is like a photo of the data: if the room changes, you must take a new photo.",
      how=["Change or add data in the source Table.", "Data > Refresh All.", "Check the totals updated."],
      tip="Refresh All before every presentation.",
      trap="Presenting a pivot that still shows last month's numbers.",
      case="The 2025 monthly pivot was built on Jan–Nov data. Then the December file arrived.",
      steps=["Note the total before refresh.", "Add the December rows to the source (or Power Query folder).",
             "Refresh All and compare."],
      table=[["State", "Lines", "Sales"], ["Jan–Nov (before refresh)", "2,850", "₦649,385,936"],
             ["December added", "462", "₦83,829,319"], ["After Refresh All", "3,312", "₦733,215,255"]],
      so="Until refreshed, the pivot was ₦83.8m short. Refresh brings in every new row.",
      split=4.20),

    P("2.10.18",
      what="A **Calculated Field** adds a new measure from other fields, e.g. **Margin = Profit / Sales**. It works "
           "on the **totals** (sum of profit ÷ sum of sales), which is the correct way to calculate a ratio.",
      like="It's like working out a class's pass rate from total passes ÷ total pupils, not by averaging each "
           "class's pass rate.",
      how=["PivotTable Analyze > Fields, Items & Sets > Calculated Field.", "Name it Margin; formula = Profit / Sales.",
           "Format as 0.0%."],
      tip="Calculate ratios from totals, never by averaging line ratios.",
      trap="Averaging the margin of each line; small loss-making lines distort it.",
      case="Head Office wants the true FY2025 chain margin in the pivot.",
      steps=["Add the calculated field Margin = Profit / Sales.", "Compare with the average of line margins.",
             "Explain the difference."],
      table=[["Method", "FY2025 margin"], ["Calculated field: SUM(Profit) ÷ SUM(Sales)", "12.7%"],
             ["Average of each line's margin (wrong)", "11.6%"]],
      so="The calculated field gives the true 12.7%. Averaging line margins understates it by over a point.",
      split=4.90),

    P("2.10.19",
      what="A **Slicer** is a panel of buttons that filters one or more pivots with a click. A **Timeline** does "
           "the same for dates with a slider. Connect one slicer to several pivots with **Report Connections**.",
      like="A slicer is like a TV remote: press 'South-West' and every screen in the room switches channel.",
      how=["Click in a pivot > PivotTable Analyze > Insert Slicer.", "Insert Timeline for a date field.",
           "Slicer > Report Connections: tick every pivot it should control."],
      tip="Connect one slicer to all the pivots on a dashboard.",
      trap="Adding a separate slicer per pivot, so they show different filters.",
      case="The analyst adds a Zone slicer and an Order_Date timeline to the live pivots.",
      steps=["Insert Slicer > Zone; click South-West.", "Insert Timeline > Order_Date; select 2025.",
             "Then narrow the timeline to Q4 2025 and clear the slicer."],
      table=[["Selection", "Sales"], ["Zone = South-West, 2025", "₦250,128,366"], ["All zones, Q4 2025", "₦280,054,067"]],
      so="Two clicks answer two questions. With Report Connections, every pivot follows the same selection.",
      split=4.90),
]
