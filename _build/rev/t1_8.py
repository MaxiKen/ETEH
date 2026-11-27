from revlib import T

SALES = "SuperMart_Stores_Sales_2022_2025.csv"

TOPICS = [
    T("1.8.1",
      what="A **dataset** is a collection of related records stored in one consistent shape: one header row at the top, "
           "one row per record, one column per detail, and no gaps, totals or notes mixed into the data. When a dataset "
           "has this shape, every Excel tool (sort, filter, formulas, PivotTables) works on it without trouble.",
      like="A dataset is like a well-kept class register: names down the side, dates across the top, one tick per box. "
           "Anyone can count attendance from it quickly.",
      how=["Check there is exactly one header row, in row 1.",
           "Check every row is one record and every column one detail.",
           "Check there are no blank rows, subtotals or notes inside."],
      tip="Read the data dictionary before you use any dataset.",
      trap="Treating a printed report, with titles and subtotals, as a dataset.",
      who="Bootcamp Files",
      case="The bootcamp catalogue (00_Data_Dictionary_and_Catalogue.xlsx) describes each dataset. The analyst checks "
           "that the SuperMart sales file has the right shape.",
      steps=["Open the Catalogue sheet and find the SuperMart sales row.",
             "Open " + SALES + " and confirm the rows and columns.",
             "Check that Row_ID is unique (no duplicates)."],
      table=[["Property", "SuperMart sales file"],
             ["Header rows", "1 (row 1)"],
             ["Records (rows)", "9,994 sales lines"],
             ["Columns", "21"],
             ["One row means", "One product line in an order"]],
      split=5.20,
      so="The file has the shape of a proper dataset, so every tool in this course will work on it directly."),

    T("1.8.2",
      what="In a dataset each **row is a record**: one thing that happened or one item, such as one sale, one admission "
           "or one student. Each **column is a variable**: one detail about that record, such as date, amount or ward. "
           "Knowing exactly what one row means (the **grain**) stops you counting the wrong thing.",
      like="Think of a church offering register: each line is one envelope (a record), and the columns are name, date "
           "and amount (variables). Count lines and you count envelopes, not people.",
      how=["Ask: what does one row represent?",
           "Find the ID column that makes each row unique.",
           "Count unique IDs of other things (orders, patients) separately."],
      tip="Write down the grain ('one row = one …') at the top of your notes.",
      trap="Counting rows and calling them customers or orders when a row is something smaller.",
      who="Multi-Sector",
      case="A manager asks 'how many orders?' and 'how many patients?' The analyst checks what one row means in each "
           "file before counting.",
      steps=["Read the grain in the catalogue for each file.",
             "Count all rows (records).",
             "Count unique Order_ID or Patient_ID values with Remove Duplicates or a PivotTable."],
      table=[["File", "One row =", "Rows", "Unique"],
             ["SuperMart sales", "order line", "9,994", "5,009 orders"],
             ["Lifecare admissions", "admission", "54,860", "48,803 patients"],
             ["Crestview students", "student", "6,607", "6,607 students"]],
      split=4.40,
      so="9,994 rows are only 5,009 orders. Know the grain before you count, or your answers will be inflated."),

    T("1.8.3",
      what="A **normal range** is just a block of cells; Excel does not know where it starts or ends. An **Excel Table** "
           "(made with **Ctrl+T**) is a named object with a fixed edge. It grows automatically when you add rows, keeps "
           "filter buttons on the headers, and lets formulas use column names instead of cell addresses.",
      like="A normal range is like goods spread loose on a market table. An Excel Table is the same goods in a labelled "
           "crate: easy to move, count and add to without losing anything.",
      how=["Click any cell in the data.",
           "Press Ctrl+T and confirm 'My table has headers'.",
           "Name it in Table Design › Table Name (e.g. tblSales)."],
      tip="Turn every dataset into a Table before analysing it.",
      trap="Writing formulas on a normal range; new rows added later are left out.",
      who="SuperMart Retail",
      case="In January 2026 SuperMart will add new sales rows to the bottom of the sales sheet. The analyst compares how "
           "a normal range and a Table handle this.",
      steps=["On a normal range, write =SUM(R2:R9995) and add a new row 9996.",
             "Convert the data to a Table and write =SUM(tblSales[Sales]).",
             "Add a new row and compare which total includes it."],
      table=[["Capability", "Normal range", "Excel Table"],
             ["New rows included", "No (fixed R2:R9995)", "Yes, automatically"],
             ["Filter buttons", "Must add", "Built in"],
             ["Formulas by name", "No", "tblSales[Sales]"],
             ["Banded rows", "Manual", "Automatic"]],
      split=4.70,
      so="The Table total includes new rows by itself; the range total silently misses them."),

    T("1.8.4",
      what="To convert data into a Table, click **any one cell** inside it and press **Ctrl+T**. Excel selects the whole "
           "block, asks if your data has headers, and creates a Table with banded rows and filter buttons. Then give it a "
           "clear name in **Table Design › Table Name**.",
      like="It is like registering a business with CAC. The shop is the same, but now it has an official name and "
           "boundaries that everyone recognises.",
      how=["Click one cell inside the data (no need to select it all).",
           "Press Ctrl+T, tick 'My table has headers', click OK.",
           "Type a name in Table Design › Table Name, e.g. tblSales."],
      tip="Name tables with a 'tbl' prefix: tblSales, tblStores.",
      trap="Leaving the default name Table1; formulas become hard to read.",
      who="SuperMart Retail",
      case="The analyst turns the SuperMart sales sheet and the store directory into Tables so they can be used in "
           "formulas and PivotTables.",
      steps=["Open " + SALES + ", click A1 and press Ctrl+T, then OK.",
             "Name it tblSales; check the Name Box list shows it.",
             "Repeat for SuperMart_Stores_Directory.csv and name it tblStores."],
      table=[["Step", "What you see"],
             ["Ctrl+T", "Dotted border around A1:U9995"],
             ["OK", "Banded rows, filter buttons"],
             ["Table Name", "tblSales (9,994 rows)"],
             ["Second table", "tblStores (12 rows)"]],
      split=5.20,
      so="Two keystrokes and a name turned plain ranges into Tables that formulas can refer to by name."),

    T("1.8.5",
      what="The **Table Design** tab controls how a Table looks: the style (colours), banded rows, a bold header, a "
           "total row, and the filter buttons. These settings change only the look, never the data. A plain, light "
           "style with banded rows is easiest to read.",
      like="Choosing a table style is like choosing the uniform for a school. It makes everything look neat and "
           "consistent, but the pupils inside are the same.",
      how=["Click inside the Table and open Table Design.",
           "Pick a light style from Table Styles.",
           "Tick Header Row and Banded Rows; untick Banded Columns."],
      tip="Use one light style for all data tables in a workbook.",
      trap="Dark, heavy styles that make numbers hard to read when printed.",
      who="SuperMart Retail",
      case="The tblStores table will be printed for the area managers' meeting. The analyst makes it easy to read.",
      steps=["Click inside tblStores › Table Design › Table Styles › a light blue style.",
             "Keep Header Row and Banded Rows ticked.",
             "Format Annual_Target_2025 with a ₦ number format."],
      table=[["Setting", "Choice", "Reason"],
             ["Table style", "Light", "Prints clearly"],
             ["Banded rows", "On", "Easy to follow across"],
             ["Banded columns", "Off", "Less clutter"],
             ["Filter buttons", "On", "Quick filtering"]],
      split=4.90,
      so="A light, banded table is easy on the eye on screen and on paper, and the data is untouched."),

    T("1.8.6",
      what="In a Table, formulas can use **column names** instead of addresses. These are called **structured "
           "references**. =SUM(tblSales[Sales]) adds the whole Sales column. Inside the table, [@Profit] means 'Profit "
           "on this row'. They are easier to read and always cover every row.",
      like="It is like calling a person by name ('Mama Bisi') instead of 'the woman in the third stall from the "
           "left'. The name still works if she moves stalls.",
      how=["Type =SUM( and click the column header area of the Table.",
           "Excel writes tblSales[Sales] for you.",
           "In a new column, type = and click a cell in the same row to get [@Column]."],
      tip="Use short, clear headers without spaces (Sales, Order_Date).",
      trap="Vague headers like 'Column1' or 'Amt' make structured references meaningless.",
      who="SuperMart Retail",
      case="The analyst calculates SuperMart's 4-year totals and adds a margin column using structured references in "
           "tblSales.",
      steps=["Type =SUM(tblSales[Sales]) and =SUM(tblSales[Profit]).",
             "Add a column called Margin and type =[@Profit]/[@Sales].",
             "Read the formulas aloud: they explain themselves."],
      table=[["Formula", "Result"],
             ["=SUM(tblSales[Sales])", "₦2,297,200,862"],
             ["=COUNTA(tblSales[Order_ID])", "9,994"],
             ["=[@Profit]/[@Sales] (Row_ID 6684)", "42.0%"]],
      split=5.20,
      so="Structured references read like sentences and always include every row of the Table."),

    T("1.8.7",
      what="**Sorting** puts rows in order by one or more columns: smallest to largest, A to Z, or newest first. Use "
           "**Data › Sort** to sort by several levels, for example by store, then by sales. In a Table, whole rows move "
           "together, so records never get mixed up.",
      like="Sorting is like arranging the queue at a bank: first by ticket colour (level 1), then by ticket number "
           "(level 2). Everyone keeps their own ticket.",
      how=["Click inside the Table.",
           "Data › Sort › Add Level for each column you want.",
           "Choose the order for each level and click OK."],
      tip="Sort the whole Table, never a single column on its own.",
      trap="Selecting one column and sorting it alone; values get separated from their rows.",
      who="SuperMart Retail",
      case="The finance team wants to see the worst loss-making sales lines of 2025 first, to understand what went "
           "wrong.",
      steps=["Filter Order_Date to 2025.",
             "Data › Sort › Profit › Smallest to Largest.",
             "Read the top rows: product, store, discount and loss."],
      table=[["Top row after sorting", "Value"],
             ["Product", "Cubify CubeX 3D Printer"],
             ["Store", "Onitsha - Main Market"],
             ["Sales / Discount", "₦7,999,980 at 50% off"],
             ["Profit", "−₦3,839,990"]],
      split=5.20,
      so="One sort shows the single worst sale: a 3D printer sold at half price lost ₦3.8m. Deep discounts again."),

    T("1.8.8",
      what="**Filtering** hides the rows you don't need, so you see only the ones that match your conditions. The rows "
           "are not deleted. Every Table has filter buttons on its headers, with tick boxes, a search box, and number, "
           "text and date filters. You can filter several columns at once.",
      like="Filtering is like using the search on a shopping app: 'phones, under ₦200,000, Lagos delivery'. The other "
           "products still exist; you just don't see them.",
      how=["Click the filter arrow on a header.",
           "Tick values, or use Number/Date Filters for conditions.",
           "Read the row count at the bottom left; clear filters when done."],
      tip="Check the status bar: 'X of Y records found' confirms the filter.",
      trap="Forgetting a filter is on, then reporting totals from hidden data.",
      who="SuperMart Retail",
      case="Abuja - Wuse lost ₦8.8m in 2025. The area manager wants to see only that store's loss-making lines.",
      steps=["Filter Order_Date to 2025 and Store_Name to Abuja - Wuse.",
             "Profit filter › Number Filters › Less Than › 0.",
             "Select the Profit column and read Count and Sum in the Status Bar."],
      table=[["Abuja - Wuse (FY2025)", "Value"],
             ["All lines", "317"],
             ["Loss-making lines", "162"],
             ["Total loss on those lines", "−₦11.8m"]],
      split=5.20,
      so="More than half of Wuse's lines lose money. The manager now has a list of 162 lines to investigate."),

    T("1.8.9",
      what="**Data › Remove Duplicates** deletes rows that repeat, keeping the first one. You choose which columns "
           "decide a duplicate. Ticking **all columns** removes only exact copies. Ticking only one column, such as "
           "Order_ID, removes every row that shares that value, even real, different records.",
      like="It is like a register where the same pupil was written twice by mistake. Remove the exact double entry, "
           "but don't delete twins just because they share a surname.",
      how=["Copy the data first (or work on a copy of the file).",
           "Data › Remove Duplicates; tick the columns that define a duplicate.",
           "Read the message: how many removed, how many remain."],
      tip="Tick all columns unless you are sure one column should be unique.",
      trap="Removing duplicates on Order_ID; orders with several products lose real lines.",
      who="SuperMart Retail",
      case="The dirty 2025 sales file has 60 exact duplicate rows. A junior analyst wants to remove duplicates using "
           "Order_ID only. What would happen?",
      steps=["On a copy, delete blank rows, then Remove Duplicates with all columns ticked.",
             "Then, on the clean 3,312 rows, try Remove Duplicates with only Order_ID ticked.",
             "Compare the rows left with the true 3,312 lines."],
      table=[["Columns ticked", "Removed", "Left", "Correct?"],
             ["All columns", "60", "3,312", "Yes"],
             ["Then Order_ID only", "1,625", "1,687", "No: real lines lost"]],
      split=4.70,
      so="Ticking only Order_ID deletes 1,625 genuine product lines. Choose duplicate keys carefully."),

    T("1.8.10",
      what="Tables grow by themselves. Type in the row just below a Table and it **auto-expands** to include the new "
           "row, copying formats and formulas. Type a formula in one cell of a new column and Excel fills the whole "
           "column: a **calculated column**. Every row uses the same formula.",
      like="It is like a family WhatsApp group: anyone new who joins automatically gets the group rules and all future "
           "messages, without being added to each one separately.",
      how=["Click the first empty column next to the Table and type a header.",
           "Type the formula in the first row and press Enter; it fills down.",
           "Type a new record below the Table; it joins the Table."],
      tip="Let calculated columns fill themselves; don't copy formulas manually.",
      trap="Typing a different formula in one row of a calculated column; it breaks consistency.",
      who="SuperMart Retail",
      case="The analyst adds a Margin column and a Loss_Flag column to tblSales, then adds a new January 2026 sale to "
           "test auto-expansion.",
      steps=["In V1 type Margin; in V2 type =[@Profit]/[@Sales] and press Enter.",
             "In W1 type Loss_Flag; in W2 type =IF([@Profit]<0,\"Loss\",\"OK\").",
             "Type a new row under the last row; check the formulas appear."],
      table=[["Behaviour", "Normal range", "Excel Table"],
             ["Formula in new column", "Copy down by hand", "Fills all 9,994 rows"],
             ["New row typed below", "Outside the data", "Joins the Table"],
             ["New row's formulas", "Missing", "Added automatically"]],
      split=4.70,
      so="One formula per column, filled automatically and kept for every new row. Much less risk of mistakes."),

    T("1.8.11",
      what="**Table Design › Total Row** adds a total line under the Table. Click any cell in it to choose Sum, Average, "
           "Count, Max, Min and more. The total row uses **SUBTOTAL**, so it calculates only the rows you can see. When "
           "you filter, the totals change to match.",
      like="It is like the running total on a supermarket receipt. Remove some items from the basket (filter them out) "
           "and the total updates to only what you are buying.",
      how=["Click inside the Table › Table Design › tick Total Row.",
           "In the total row under Sales, choose Sum from the drop-down.",
           "Apply a filter and watch the total change."],
      tip="Use the total row for quick checks of filtered groups.",
      trap="Typing =SUM() under a Table; it does not react to filters like SUBTOTAL does.",
      who="SuperMart Retail",
      case="The analyst wants quick totals for different years and stores in tblSales without writing new formulas "
           "each time.",
      steps=["Tick Total Row; choose Sum under Sales and Count under Order_ID.",
             "Filter Order_Date to 2025.",
             "Then also filter Store_Name to Lagos - Ikeja."],
      table=[["Filter applied", "Sum of Sales", "Count"],
             ["None (2022–2025)", "₦2,297,200,862", "9,994"],
             ["2025 only", "₦733,215,255", "3,312"],
             ["2025, Lagos - Ikeja", "₦146,388,345", "663"]],
      split=4.70,
      so="The total row follows every filter, so you can check any group's total in seconds."),

    T("1.8.12",
      what="Tables prevent the most common spreadsheet failures. Formulas that miss new rows, PivotTables that don't "
           "pick up new data, sorts that scramble records and charts that stop at the old last row all disappear when "
           "the data is in a Table. It is the cheapest reliability upgrade in Excel.",
      like="A Table is like putting your money in a bank account instead of under the mattress. Same money, but it is "
           "tracked, protected and easy to add to.",
      how=["Convert every dataset to a Table with Ctrl+T.",
           "Build formulas, PivotTables and charts from the Table name.",
           "Add new data at the bottom of the Table and click Refresh."],
      tip="Build every PivotTable and chart on a Table, not a range.",
      trap="Pointing reports at fixed ranges like A1:U9995 that will not grow.",
      who="SuperMart Retail",
      case="Each month SuperMart adds about 300 new sales lines. The analyst compares what breaks with a range and with a "
           "Table.",
      steps=["Build one report on A1:U9995 and one on tblSales.",
             "Paste a month of new rows under the data.",
             "Refresh both and compare totals."],
      table=[["Failure", "With a range", "With a Table"],
             ["SUM misses new rows", "Yes", "No"],
             ["PivotTable misses data", "Yes", "No (refresh)"],
             ["Chart stops early", "Yes", "No"],
             ["Sort mixes up rows", "Possible", "No"]],
      split=4.70,
      so="With a Table, monthly updates just work. With a range, every report has to be fixed by hand each month."),
]
