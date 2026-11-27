from revlib import T

DIRTY = "SuperMart_Sales_2025_DIRTY.csv"

TOPICS = [
    T("2.2.1",
      what="**Power Query** is Excel's built-in tool for bringing data in and cleaning it. Every click you make "
           "(remove a column, fix a spelling, change a type) is **recorded as a step**. Next time, you click "
           "**Refresh** and Power Query repeats every step on the new data automatically.",
      like="Power Query is like a written recipe. The first time you cook jollof you note every step. Next time anyone "
           "can follow the recipe and get the same pot, without you standing in the kitchen.",
      how=["Data › Get Data › choose the source › Transform Data.",
           "Clean the data with the ribbon buttons; each click becomes a step.",
           "Close & Load; next month click Data › Refresh All."],
      tip="Clean inside Power Query, so the work can be repeated with one click.",
      trap="Cleaning the loaded sheet by hand, which must then be redone every month.",
      who="SuperMart Retail",
      case="The 2025 POS export " + DIRTY + " has blank rows, duplicates and text numbers. The analyst cleans it "
           "once in Power Query.",
      steps=["Data › From Text/CSV › choose the file › Transform Data.",
             "Apply the clean-up steps from the ribbon (14 in total).",
             "Close & Load; check the final row count."],
      table=[["Stage", "Rows"],
             ["Rows in the raw export", "3,387"],
             ["Steps recorded", "14"],
             ["Rows loaded, clean", "3,312"],
             ["Effort next time", "One click: Refresh"]],
      split=5.20,
      so="Once recorded, the 14 steps rerun on any new export. The cleaning is done once and reused forever."),

    T("2.2.2",
      what="Power Query exists because manual cleaning is slow, error-prone and impossible to check. It gives four "
           "big advantages: **repeatable** (same steps every time), **auditable** (every step is listed and "
           "visible), **scalable** (no helper columns, even for large data) and **reversible** (delete a step instead "
           "of undoing many edits).",
      like="It is like a bank's standing order compared with walking to the bank every month. Set it up once, it "
           "happens the same way every month, and there's a record of every payment.",
      how=["Notice any cleaning task you do more than once.",
           "Rebuild it as Power Query steps.",
           "Next time, Refresh instead of repeating the work."],
      tip="If you'll do a cleaning task twice, build it in Power Query.",
      trap="Thinking 'it's just this once'; monthly reports always come back.",
      who="SuperMart Retail",
      case="SuperMart's 2025 sales arrive as 12 monthly files. The analyst compares cleaning them by hand with "
           "one Power Query that combines the folder.",
      steps=["Count the monthly files: 12, from 107 to 462 lines each.",
             "By hand: open, clean and paste each file in turn.",
             "Power Query: build the steps once on the folder, then Refresh."],
      table=[["Advantage", "By hand (12 files)", "Power Query"],
             ["Repeatable", "Redo 12 times", "Built once"],
             ["Auditable", "No record of edits", "Every step listed"],
             ["Reversible", "Undo is limited", "Delete one step"],
             ["New file in 2026", "Start again", "Drop it in, Refresh"]],
      split=4.70,
      so="The manual approach costs the same effort every month. Power Query's cost is paid once."),

    T("2.2.3",
      what="**Traditional Excel cleaning** changes the cells themselves: Find & Replace, TRIM helper columns, filter "
           "and delete. **Power Query** changes the **recipe** that produces the cells. Fix the traditional way and "
           "you fix this month's file; fix it in Power Query and every future month is fixed too.",
      like="Traditional cleaning is like mopping a wet floor. Power Query is like fixing the leaking pipe: the floor "
           "stays dry next month too.",
      how=["List each cleaning task you do.",
           "Find the Power Query button for it (Remove Rows, Replace Values, Trim).",
           "Record it once as a step instead of repeating it in cells."],
      tip="Use Power Query for recurring files and Excel for one-off quick checks.",
      trap="Mixing the two: cleaning in Power Query, then making hand edits on the loaded table.",
      who="SuperMart Retail",
      case=DIRTY + " has four common problems. The analyst compares how each is fixed in Excel and in Power Query.",
      steps=["Count each problem in the raw file.",
             "Note the Excel fix (manual, in the cells).",
             "Note the Power Query step (recorded, repeatable)."],
      table=[["Problem in the file", "Count", "Excel fix", "Power Query step"],
             ["Blank rows", "15", "Filter and delete", "Remove Blank Rows"],
             ["Duplicate rows", "60", "Remove Duplicates", "Remove Duplicates"],
             ["Sales stored as text", "234", "VALUE helper column", "Replace ₦, change type"],
             ["Discount as '20%' text", "173", "Find & Replace", "Replace Values"]],
      split=4.20,
      so="The fixes are similar, but only the Power Query ones are recorded and run again by themselves next time."),

    T("2.2.4",
      what="Every Power Query does three things, called **ETL**. **Extract**: connect to the source and read it. "
           "**Transform**: clean it, reshape it, set the types. **Load**: send the result to a worksheet table or the "
           "Data Model. Most of the value is created in Transform, so do all your cleaning there, before loading.",
      like="ETL is like processing cassava: harvest it (extract), peel, wash and grate it (transform), then pack the "
           "garri into bags for sale (load). Nobody peels cassava after it's packed.",
      how=["Extract: Data › Get Data › pick the source.",
           "Transform: clean in the Power Query Editor.",
           "Load: Close & Load To › Table or Data Model."],
      tip="Finish all cleaning before Load; the loaded table should only need formatting.",
      trap="Loading raw data first and fixing it in the sheet afterwards.",
      who="SuperMart Retail",
      case="The analyst runs the full ETL cycle on " + DIRTY + " and records what happens at each stage.",
      steps=["Extract: From Text/CSV › the DIRTY file.",
             "Transform: apply the clean-up steps.",
             "Load: Close & Load To › Table on a new sheet."],
      table=[["Stage", "What you click", "Result"],
             ["Extract", "Data › From Text/CSV", "3,387 rows × 21 columns"],
             ["Transform", "14 recorded steps", "Blanks, duplicates, types fixed"],
             ["Load", "Close & Load To › Table", "3,312 clean rows"]],
      split=4.70,
      so="3,312 rows matches the POS control total, so the cleaning removed exactly the bad rows and nothing else."),

    T("2.2.5",
      what="Open the **Power Query Editor** in two ways. For **new** data: Data › Get Data › choose the source › "
           "**Transform Data**. To **change an existing** query: Data › **Queries & Connections** › right-click the "
           "query › **Edit**. The Editor is a separate window. **Close & Load** saves your changes and sends the result "
           "to Excel.",
      like="The Editor is like a workshop behind the shop. Customers only see the finished goods on the shelf (the "
           "sheet), but all the work happens in the back.",
      how=["New data: Get Data › source › click Transform Data (not Load).",
           "Existing query: Queries & Connections › right-click › Edit.",
           "Finish with Close & Load (or Close & Load To)."],
      tip="Click Transform Data, not Load, so you check the data before it lands.",
      trap="Clicking Load straight away and loading messy data into the sheet.",
      who="SuperMart Retail",
      case="The analyst opens " + DIRTY + " for the first time, then returns later to edit the query.",
      steps=["Data › From Text/CSV › select the file.",
             "In the preview window, click Transform Data.",
             "Later: Queries & Connections › right-click › Edit."],
      table=[["Goal", "Where to click", "Result"],
             ["Start a new query", "Get Data › Transform Data", "Editor opens with a preview"],
             ["Change a query", "Queries & Connections › Edit", "Editor opens at the last step"],
             ["Save and return", "Home › Close & Load", "Table refreshed in Excel"]],
      split=4.70,
      so="Two doors lead into the same Editor: one for new data, one for changing existing queries."),

    T("2.2.6",
      what="The Power Query Editor has five areas. The **Ribbon** (Home, Transform, Add Column, View) holds the "
           "buttons. The **Queries pane** (left) lists your queries. The **preview grid** (middle) shows the data, with "
           "a type icon on each column header. **Applied Steps** (right) lists every step. The **Formula Bar** shows "
           "the code behind the selected step.",
      like="It is like a car dashboard: the speedometer, fuel gauge and warning lights each tell you something "
           "different. Once you know where each one is, you can drive with confidence.",
      how=["Turn on the Formula Bar: View › Formula Bar.",
           "Read the type icon on each header: ABC (text), 123 (whole number), calendar (date).",
           "Click a step in Applied Steps to see the data at that point."],
      tip="Read the column type icons before doing anything else.",
      trap="Working with the Formula Bar hidden; you lose sight of what each step does.",
      who="SuperMart Retail",
      case="The analyst has just opened " + DIRTY + " in the Editor and tours the screen before starting.",
      steps=["Find each of the five areas.",
             "Read the type icons on Order_Date and Sales.",
             "Click the Source step and read the Formula Bar."],
      table=[["Area", "Where", "What you see here"],
             ["Queries pane", "Left", "SuperMart_Sales_2025_DIRTY"],
             ["Preview grid", "Middle", "21 columns with type icons"],
             ["Applied Steps", "Right", "Source, Promoted Headers, Changed Type"],
             ["Formula Bar", "Above the grid", "= Csv.Document(File.Contents(…))"]],
      split=4.20,
      so="Each area answers one question: what queries exist, what the data looks like, what was done, and how."),

    T("2.2.7",
      what="A **query** is a named list of steps. In **Applied Steps** each step uses the result of the step before "
           "it, from top to bottom. You can click any step to see the data at that moment, rename steps, delete them "
           "or drag them into a new order. **Query Settings** holds the query's name and description.",
      like="Applied Steps are like a step-by-step recipe card. Click any line to see what the pot looked like at that "
           "stage, and cross out a step if it wasn't needed.",
      how=["Click each step from the top and watch the data change.",
           "Right-click › Rename to give steps clear names.",
           "Right-click › Delete to remove a wrong step."],
      tip="Rename steps so a colleague understands the query without asking.",
      trap="Deleting a step in the middle without checking the steps after it still work.",
      who="SuperMart Retail",
      case="The cleaning query for " + DIRTY + " has 14 steps. The analyst clicks through them and notes the row "
           "count after each.",
      steps=["Click Source and note the row count (bottom-left of the Editor).",
             "Click each step in turn and note the count.",
             "Find the steps that removed rows."],
      table=[["Step", "What it does", "Rows after"],
             ["1 Source", "Reads the CSV file", "3,387"],
             ["2 Removed Blank Rows", "Drops 15 spacer rows", "3,372"],
             ["3 Removed Columns", "Deletes empty Column1", "3,372"],
             ["4 Removed Duplicates", "Drops 60 exact copies", "3,312"],
             ["14 Loaded", "Close & Load", "3,312"]],
      split=4.70,
      so="Only two steps remove rows, and they explain the whole difference: 15 blanks plus 60 duplicates is 75."),

    T("2.2.8",
      what="On the **View** tab, tick three profiling tools. **Column Quality** shows the % of valid, error and empty "
           "values. **Column Distribution** shows distinct and unique counts. **Column Profile** shows min, max, "
           "average and top values. They turn 'I think this column is messy' into exact numbers.",
      like="It is like a doctor's lab test. Instead of saying 'you look tired', she shows your blood count and points to "
           "the exact number that is low.",
      how=["View › tick Column Quality, Column Distribution, Column Profile.",
           "Click 'based on top 1000 rows' at the bottom › choose entire data set.",
           "Note any column with errors, empties or too many distinct values."],
      tip="Switch profiling to the entire data set; by default it checks only 1,000 rows.",
      trap="Trusting a profile based on the first 1,000 rows only.",
      who="SuperMart Retail",
      case="After removing blank rows, the analyst profiles " + DIRTY + " (3,372 rows) before cleaning further.",
      steps=["Turn on the three profiling tools on the View tab.",
             "Set the profile to the entire data set.",
             "Read Column Quality for each column."],
      table=[["Column", "Problem found", "Valid"],
             ["Order_Date", "335 errors (mixed formats)", "90.1%"],
             ["Sales", "234 errors (text with ₦)", "93.1%"],
             ["Discount", "173 errors ('20%' text)", "94.9%"],
             ["Quantity", "31 empty", "99.1%"],
             ["Category", "12 distinct values for 3 categories", "100%"]],
      split=4.70,
      so="The profile gives an exact to-do list. Category is 100% 'valid' but has 12 spellings, so check distinct counts too."),

    T("2.2.9",
      what="The **Queries pane** helps you organise many queries. **Groups** are folders. **Duplicate** makes an "
           "independent copy of a query. **Reference** makes a new query that **starts from another query's result**. "
           "Reference is the professional pattern: clean once, then build several outputs on top of the clean query.",
      like="Reference is like one central kitchen supplying several restaurants. Fix the stew recipe in the central "
           "kitchen and every restaurant serves the improved stew.",
      how=["Right-click the clean query › Reference.",
           "Rename the new query and add only its extra steps.",
           "Right-click › Move To Group › e.g. 02 Staging or 03 Final."],
      tip="Use Reference when a new query depends on cleaned data.",
      trap="Using Duplicate instead; every fix must then be made in each copy.",
      who="SuperMart Retail",
      case="The analyst has one clean query, qSales_Clean (3,312 rows). Management wants sales by store and sales "
           "by month.",
      steps=["Right-click qSales_Clean › Reference › rename qSales_ByStore.",
             "Group By Store_Name, sum Sales.",
             "Repeat with Reference for qSales_ByMonth."],
      table=[["Query", "Starts from", "Rows"],
             ["qSales_Clean", "The DIRTY CSV", "3,312"],
             ["qSales_ByStore", "Reference to qSales_Clean", "12 stores"],
             ["qSales_ByMonth", "Reference to qSales_Clean", "12 months"]],
      split=4.70,
      so="A fix in qSales_Clean flows into both summaries automatically. The cleaning lives in one place."),

    T("2.2.10",
      what="**Close & Load To** offers several destinations. **Table**: rows on a worksheet (up to 1,048,576). "
           "**Only Create Connection**: no output; it feeds other queries. **Add to the Data Model**: rows are stored "
           "in Excel's internal engine, which handles millions of rows and links between tables, and feeds "
           "PivotTables and DAX (Week 3).",
      like="It is like choosing where to keep goods: on the shop shelf (Table) where people see them, or in the "
           "warehouse (Data Model) where far more fits and you bring out only what's needed.",
      how=["Home › Close & Load (arrow) › Close & Load To.",
           "Choose Table, PivotTable report or Only Create Connection.",
           "Tick 'Add this data to the Data Model' for big or linked tables."],
      tip="Load staging queries as Connection only; they don't need a sheet.",
      trap="Loading every query to a sheet, which makes the file big and slow.",
      who="Multi-Sector",
      case="The analyst has four queries of different sizes. She chooses where each one should load.",
      steps=["Check each query's row count.",
             "Ask: will people read the rows, or only summaries?",
             "Choose the load destination."],
      table=[["Query", "Rows", "Load to"],
             ["qSales_ByStore", "12", "Table (people read it)"],
             ["qSales_Clean (SuperMart 2025)", "3,312", "Table or Data Model"],
             ["Lifecare admissions 2019–2024", "54,860", "Data Model"],
             ["Diabetes screening", "100,000", "Data Model"]],
      split=4.70,
      so="Small results go to sheets; large tables go to the Data Model, which keeps the file light and fast."),
]
