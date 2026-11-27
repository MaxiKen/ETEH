from revlib import T

RAWX = "SuperMart_Online_RAW_Export.xlsx"
PIPE = "Lifecare_Admissions_2023_pipe.txt"
FOLDER = "SuperMart_Monthly_Sales_2025"

TOPICS = [
    T("2.3.1",
      what="After choosing a workbook in **From Workbook**, always click **Transform Data**, not Load. The first "
           "jobs in the Editor are to check the headers were promoted, and to **set the right type** for each column. "
           "Data that comes from other people's workbooks often has numbers, dates and codes stored in the wrong type.",
      like="It is like receiving a parcel: you open it and check the contents at the door, before you put anything "
           "away in your cupboard.",
      how=["From Workbook › pick the sheet › Transform Data.",
           "Check row 1 became the headers (Use First Row as Headers).",
           "Check each column's type icon and correct any wrong ones."],
      tip="Check every column's type before you do anything else.",
      trap="Trusting a workbook's types; codes and phone numbers are often stored as numbers.",
      who="SuperMart Online",
      case="In the Customers sheet of " + RAWX + ", phone numbers are stored as numbers, so the leading 0 is "
           "gone (7083197857). Dates of birth arrive as date-and-time values.",
      steps=["From Workbook › Customers › Transform Data.",
             "phone_number: Add Column › Custom Column: \"0\" & Text.From([phone_number]).",
             "date_of_birth: change the type to Date."],
      table=[["Column", "As imported", "After fixing"],
             ["phone_number", "7083197857 (number)", "07083197857 (text)"],
             ["date_of_birth", "1998-12-04 00:00:00", "1998-12-04 (date)"],
             ["Rows", "2,500", "2,500"]],
      split=4.70,
      so="Two small type fixes make all 2,500 phone numbers usable and the birth dates ready for age calculations."),

    T("2.3.2",
      what="**From Text/CSV** is the right door for every separated text file, even if the name ends in .txt. The "
           "preview window changes as you adjust the settings, so you can see the result before committing. Clicking "
           "**Transform Data** opens the Editor with two steps already made: **Promoted Headers** and **Changed Type**.",
      like="It is like trying on clothes in a shop mirror: you adjust the fit until it looks right, then you buy.",
      how=["Data › From Text/CSV › choose the file.",
           "Check the preview: right number of columns, readable text.",
           "Click Transform Data and review the automatic Changed Type step."],
      tip="Check the automatic Changed Type step; it guesses from the first 200 rows only.",
      trap="Accepting every automatic type without looking, so an ID column becomes a number.",
      who="SuperMart Retail",
      case="The analyst imports the January file, 2025-01_SuperMart_Sales.csv, from " + FOLDER + ".",
      steps=["From Text/CSV › 2025-01_SuperMart_Sales.csv.",
             "Check the preview, then click Transform Data.",
             "Read the two automatic steps in Applied Steps."],
      table=[["What you check", "What you should see"],
             ["File Origin", "65001: Unicode (UTF-8)"],
             ["Delimiter", "Comma"],
             ["Columns and rows", "21 columns, 155 rows"],
             ["Automatic steps", "Promoted Headers, Changed Type"]],
      split=5.20,
      so="The preview confirmed the shape before loading. 155 rows matches January's line count."),

    T("2.3.3",
      what="For .txt files, use **From Text/CSV** and set the delimiter: Tab, or **Custom** for a pipe (|) or "
           "another character. Some system exports start with report information (title, date, user) above the "
           "headers; remove those lines with **Remove Rows › Remove Top Rows** before promoting headers.",
      like="It is like a letter with a long address block on top. You skip the address and start reading at "
           "'Dear Sir'.",
      how=["From Text/CSV › Delimiter › Custom › type |.",
           "If needed, Home › Remove Rows › Remove Top Rows.",
           "Check the row count against a control total."],
      tip="Compare the row count with an independent control total.",
      trap="Assuming the import is complete because no error appeared.",
      who="Lifecare Hospital",
      case="The analyst imports " + PIPE + ". The core dataset says Lifecare had 10,918 admissions in 2023, which "
           "gives an independent check.",
      steps=["From Text/CSV › the pipe file › Delimiter Custom |.",
             "Transform Data; check there are 22 columns.",
             "Compare the row count with 10,918."],
      table=[["Measure", "Result"],
             ["Delimiter", "Custom: |"],
             ["Columns", "22"],
             ["Rows imported", "10,918"],
             ["Control total (core dataset)", "10,918, so PASS"]],
      split=5.20,
      so="The two counts agree, so no admission was lost or added during import."),

    T("2.3.4",
      what="A workbook can hold several sheets. **Append** them into one table only if they have the **same "
           "columns** (for example one sheet per month). If the sheets hold **different kinds** of data (customers, "
           "products, sellers), load each as its own query; later you link them with Merge or relationships.",
      like="You can pour several bags of the same rice into one big drum. But you would never pour rice, beans and "
           "garri into the same drum.",
      how=["Open the workbook in the Navigator and compare the sheets' columns.",
           "Same columns: Home › Append Queries as New.",
           "Different columns: keep separate queries."],
      tip="Compare the column headers before appending anything.",
      trap="Appending different tables, which creates a mess of mostly-empty columns.",
      who="SuperMart Online",
      case=RAWX + " has 5 sheets. The analyst decides which can be appended and which must stay separate.",
      steps=["Click each sheet in the Navigator and note its columns.",
             "Ask: do these sheets describe the same kind of record?",
             "Load each different table as its own query."],
      table=[["Sheet", "Columns", "One row is…", "Decision"],
             ["Transactions", "15", "a payment", "Own query"],
             ["Deliveries", "13", "a delivery", "Own query"],
             ["Customers", "15", "a customer", "Own query"],
             ["Products / Sellers", "15 / 14", "a product / seller", "Own queries"]],
      split=4.20,
      so="All five sheets describe different records, so none should be appended. They are linked later with Merge Queries or Data Model relationships (Week 3)."),

    T("2.3.5",
      what="**Get Data › From File › From Folder** reads every file in a folder. **Combine Files** applies the same "
           "steps to each file and stacks them into one table, with a **Source.Name** column showing which file each "
           "row came from. Filter the file list first, so only the right files are combined.",
      like="It is like a teacher collecting 12 class registers and copying them into one school register, writing "
           "the class name beside every pupil so each one can be traced back.",
      how=["Get Data › From Folder › pick the folder › Transform Data.",
           "Filter Extension to .csv; remove stray files.",
           "Click the Combine Files button on the Content column."],
      tip="Keep only the source files in the folder; never save outputs there.",
      trap="Leaving an old copy or a combined output in the folder, which duplicates rows on refresh.",
      who="SuperMart Retail",
      case=FOLDER + " holds 12 monthly CSVs. Finance wants one 2025 table that matches the POS control totals.",
      steps=["From Folder › " + FOLDER + " › Transform Data.",
             "Filter Extension = .csv, then Combine Files.",
             "Check total rows and total sales against the POS control."],
      table=[["File", "Rows", "Sales"],
             ["2025-01_SuperMart_Sales.csv", "155", "₦44.0m"],
             ["2025-02_SuperMart_Sales.csv", "107", "₦20.3m"],
             ["2025-11_SuperMart_Sales.csv", "459", "₦118.4m"],
             ["All 12 files combined", "3,312", "₦733.2m"]],
      split=4.70,
      so="12 files became one 3,312-row table that matches the POS total of ₦733,215,255. Source.Name traces each row."),

    T("2.3.6",
      what="Every query remembers the **full path** of its source in the Source step, such as "
           "C:\\Users\\Ada\\Downloads\\file.csv. If the file moves, or a colleague opens the workbook on another "
           "computer, the query breaks. Keep source files in a stable, shared folder and keep the path in **one place**.",
      like="A file path is like a home address. If you move house without telling anyone, the post goes to the old "
           "address and never reaches you.",
      how=["Check the path: Edit the query › click Source › read the Formula Bar.",
           "Keep data in a stable shared folder, not Downloads or Desktop.",
           "If the path changes: Home › Data Source Settings › Change Source."],
      tip="Keep all source files in one documented, shared folder.",
      trap="Building queries on files in your own Downloads folder.",
      who="SuperMart Retail",
      case="Ada builds the monthly sales query from her Downloads folder, then emails the workbook to Musa. On "
           "Musa's computer, Refresh fails.",
      steps=["Musa clicks Refresh: the error 'could not find the folder' appears.",
             "He opens Source and sees C:\\Users\\Ada\\Downloads\\…",
             "He changes the source to the shared folder and refreshes."],
      table=[["Path problem", "What you see", "Fix"],
             ["File in someone's Downloads", "Folder not found", "Use a shared folder"],
             ["File renamed", "File not found", "Keep stable file names"],
             ["Path typed in many queries", "Must edit each one", "Use one parameter (2.3.7)"]],
      split=4.70,
      so="A path that only works on one computer is a hidden fault. Shared folders and one path setting fix it."),

    T("2.3.7",
      what="A **parameter** is a named value that queries can use, such as **pFolderPath** holding the data folder. "
           "Create it with **Home › Manage Parameters › New Parameter**. Change the parameter once and every query "
           "using it points to the new place. Passwords must **never** be stored in parameters or cells; Excel keeps "
           "credentials separately.",
      like="A parameter is like a single master switch for a building. Flip it once and every room changes, instead "
           "of walking to each room's switch.",
      how=["Home › Manage Parameters › New: pFolderPath (Text).",
           "In each Source step, replace the typed path with pFolderPath.",
           "To move the data, change only the parameter's value."],
      tip="Put every folder path in one parameter.",
      trap="Typing passwords into a query or a cell in a shared workbook.",
      who="ETEH Practice Files",
      case="The Week 2 workbook has queries for the monthly folder, the pipe text file and the EU targets. All of "
           "them read from Week2_Data_Files.",
      steps=["Manage Parameters › New: pFolderPath (Text), value = your Week2_Data_Files folder path.",
             "Edit each Source step to use pFolderPath & the file name.",
             "Change pFolderPath to your own folder; Refresh All."],
      table=[["Query", "Source step uses"],
             ["Monthly sales folder", "pFolderPath & \"SuperMart_Monthly_Sales_2025\""],
             ["Lifecare pipe file", "pFolderPath & \"" + PIPE + "\""],
             ["EU store targets", "pFolderPath & \"SuperMart_Store_Targets_EU_semicolon.csv\""]],
      split=4.20,
      so="Moving the files now means changing one value. All three queries follow automatically."),

    T("2.3.8",
      what="**Data › Refresh All** (Ctrl+Alt+F5) reruns every query from its source and updates PivotTables. To "
           "refresh one query, right-click it in Queries & Connections › Refresh. After every refresh, **check the row "
           "count and a known total**; no error message does not mean the result is right.",
      like="Refresh is like re-checking a stock count after a new delivery. Count again, and compare with the "
           "delivery note.",
      how=["Add or replace the source file.",
           "Data › Refresh All (Ctrl+Alt+F5).",
           "Check the row count and total against the control figures."],
      tip="After every refresh, check the row count and one known total.",
      trap="Assuming the refresh worked because no error appeared.",
      who="SuperMart Retail",
      case="The December file arrives late. The folder query was built with 11 files. The analyst drops "
           "2025-12_SuperMart_Sales.csv into " + FOLDER + " and refreshes.",
      steps=["Before: note rows and sales with 11 files.",
             "Copy the December file into the folder; Refresh All.",
             "Compare the new totals with the POS control."],
      table=[["Measure", "Before (11 files)", "After refresh (12 files)"],
             ["Rows", "2,850", "3,312"],
             ["Sales", "₦649.4m", "₦733.2m"],
             ["Matches POS control?", "No", "Yes, PASS"]],
      split=4.70,
      so="One refresh added December's 462 lines with no rebuilding. The control check proves it worked."),
]
