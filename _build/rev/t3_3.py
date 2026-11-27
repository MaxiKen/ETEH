from revlib import T

ON = "SuperMart Online"
TRY = "ETEH_Week3_Section_Examples.xlsx  ›  sheet '3.3 ODBC Connections'  ·  supermart_online.db"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", ON)
    return T(tag, **k)


TOPICS = [
    O("3.3.1",
      what="**ODBC** (Open Database Connectivity) is a standard way for programs like Excel to talk to databases. "
           "Excel speaks ODBC; a **driver** translates it for the specific database (SQLite, SQL Server, MySQL…). "
           "One method works for many databases.",
      like="ODBC is like a universal travel adaptor: Excel's plug stays the same, and the adaptor (driver) fits "
           "whichever socket (database) you meet.",
      how=["Install the driver for your database.", "In Excel: Data > Get Data > From Other Sources > From ODBC.",
           "Choose the connection and the data to load."],
      tip="Use ODBC when Excel has no built-in connector for your database.",
      trap="Thinking the data is copied permanently; the connection reads it again on every refresh.",
      case="The analyst wants SuperMart Online data from supermart_online.db inside Excel.",
      steps=["Excel (the application) asks for data via ODBC.", "The SQLite3 ODBC driver translates the request.",
             "supermart_online.db answers; the result lands in Excel."],
      table=[["Layer", "Role", "In this example"], ["Application", "Asks for data", "Excel / Power Query"],
             ["ODBC driver", "Translates", "SQLite3 ODBC Driver"], ["Database", "Stores and answers", "supermart_online.db"]],
      so="Excel never needs to understand SQLite itself; the driver does the translating.",
      split=4.20),

    O("3.3.2",
      what="A **driver** is the translator for one database. It must match your Office version: **64-bit Excel "
           "needs a 64-bit driver**. A mismatch is the most common ODBC error.",
      like="A driver is like an interpreter: you need one who speaks both your language and the other person's, "
           "and who turns up to the right meeting room (32 or 64-bit).",
      how=["Check Office bitness: File > Account > About Excel.", "Install the matching SQLite3 ODBC driver.",
           "Confirm it appears in 'ODBC Data Sources (64-bit)'."],
      tip="Check bitness before installing anything.",
      trap="Installing a 32-bit driver for 64-bit Excel.",
      case="A colleague's connection fails with 'Data source name not found and no default driver specified'.",
      steps=["Check Excel: 64-bit.", "Open ODBC Data Sources (64-bit) > Drivers: SQLite3 is missing.",
             "Install the 64-bit SQLite3 ODBC driver and retry."],
      table=[["Situation", "Result"], ["64-bit Excel + 32-bit driver", "Data source name not found"],
             ["64-bit Excel + 64-bit driver", "Connects"], ["Driver name typed differently", "Data source name not found"]],
      so="The error text was misleading; the real cause was a driver of the wrong bitness.",
      split=4.90),

    O("3.3.3",
      what="A **DSN** (Data Source Name) is a connection saved on **your computer** under a name. It's handy, but "
           "a workbook using it breaks on a colleague's PC. A **DSN-less** connection puts the full connection "
           "string in the workbook, so it travels with the file.",
      like="A DSN is like a speed-dial number saved on your phone; DSN-less is like writing the full phone number "
           "on the document.",
      how=["For shared files, choose DSN-less: Data source = (None).", "Paste the connection string.",
           "Store the folder path in a parameter, not typed into every query."],
      tip="Use DSN-less connections for files you will share.",
      trap="Sharing a workbook that depends on a DSN only your PC has.",
      case="The analyst's report must work on the whole team's computers.",
      steps=["Choose (None) as the DSN in From ODBC.",
             "Paste: Driver={SQLite3 ODBC Driver};Database=…\\supermart_online.db;",
             "Build the path from the pFolderPath parameter."],
      table=[["Connection", "Stored where", "Works on a colleague's PC?"], ["DSN", "Windows on your computer", "Only if they create the same DSN"],
             ["DSN-less", "Inside the workbook", "Yes, if the driver is installed"]],
      so="DSN-less connections travel with the workbook. Only the driver must be installed on each PC.",
      split=4.20),

    O("3.3.4",
      what="To connect: **Data > Get Data > From Other Sources > From ODBC**, choose a DSN or paste a connection "
           "string, then either pick a table in the Navigator or, under **Advanced options**, paste a **SQL "
           "statement**. Finally choose **Load To**.",
      like="It's like ordering by phone: dial the shop (connection), say exactly what you want (SQL), and choose "
           "delivery or pick-up (Load To).",
      how=["Data > Get Data > From Other Sources > From ODBC.", "Paste the connection string; open Advanced options.",
           "Paste the SQL; Load To a Table or the Data Model."],
      tip="Paste tested SQL from DB Browser into Advanced options.",
      trap="Loading every table 'just in case'.",
      case="The analyst loads the 2025 zone scorecard (query 3.2.17) into Excel.",
      steps=["Connect DSN-less to supermart_online.db.", "Advanced options: paste the 3.2.17 SQL.",
             "Load To > Table on a new sheet."],
      table=[["Step", "Choice", "Result"], ["Connector", "From ODBC (DSN-less)", "Connected"],
             ["Advanced options", "3.2.17 SQL", "Only the summary is fetched"],
             ["Load To", "Table", "6 rows: one per zone"]],
      so="Five steps, and Excel holds a 6-row scorecard that refreshes from the database.",
      split=4.20),

    O("3.3.5",
      what="Many databases need a login: **Windows authentication**, a **username and password**, or a "
           "**token**. Power Query keeps credentials in its own secure store, not in the workbook. **Never** type "
           "a password into a cell or a connection string.",
      like="Credentials are like your ATM PIN: you type it when asked, but you never write it on the card.",
      how=["Prefer Windows (single sign-on) authentication.", "Enter passwords only when Power Query prompts.",
           "Manage them in Data > Get Data > Data Source Settings."],
      tip="Use Windows authentication where available.",
      trap="Putting a password in the connection string, visible to anyone with the file.",
      case="SQLite files have no login, but the team's future SQL Server does. The analyst sets up safe access.",
      steps=["Choose Windows authentication in the connector.", "If a password is required, type it in the prompt only.",
             "Check Data Source Settings lists the source with saved permissions."],
      table=[["Method", "Security", "Advice"], ["Windows authentication", "Strong; no password stored", "Preferred"],
             ["Prompted username + password", "Stored in the credential store", "Acceptable"],
             ["Password in connection string", "Visible to anyone with the file", "Never"]],
      so="Credentials belong in the secure store. The workbook should never contain a password.",
      split=4.20),

    O("3.3.6",
      what="You can import a **whole table** and filter in Excel, or write **SQL** so the database filters, "
           "joins and summarises first, and sends only the result. SQL moves far less data and is usually faster.",
      like="A whole table is like buying the entire market stall to get one tomato; SQL is asking the trader for "
           "the one tomato.",
      how=["Ask: do I need every row, or a summary?", "For a summary, write SQL and paste it in Advanced options.",
           "Compare the rows that arrive in Excel."],
      tip="Let the database do the heavy lifting.",
      trap="Importing whole tables and rebuilding the summary in Excel every month.",
      case="The analyst needs the 2025 zone scorecard. Option A imports the whole transactions table; option B "
           "sends the 3.2.17 SQL.",
      steps=["A: import transactions (all rows, all columns).", "B: import the 3.2.17 query result.",
             "Compare the rows that reach Excel."],
      table=[["Approach", "Rows into Excel", "Work left in Excel"], ["A: whole table", "11,000", "Filter, join, group"],
             ["B: SQL query", "6", "None"], ["Rows saved", "10,994", "0"]],
      so="The SQL route brought in 6 rows instead of 11,000. The database did the work.",
      split=4.20),

    O("3.3.7",
      what="A database query refreshes like any Power Query query: **Data > Refresh All** reruns the SQL and "
           "reloads the result. The SQL is saved in the workbook, so anyone can refresh it and get today's numbers.",
      like="Refreshing is like re-reading a live scoreboard: the question stays the same, the answer is current.",
      how=["Data > Refresh All (Ctrl+Alt+F5).", "Set Connection Properties > Refresh on open if needed.",
           "Record the refresh time on the report."],
      tip="Show a 'last refreshed' time on reports that use live data.",
      trap="Refreshing during busy hours on a live system, slowing it down.",
      case="The 2025 zone scorecard is refreshed each month as more 2025 orders arrive in the database.",
      steps=["Open the workbook.", "Data > Refresh All.", "Check the total order count changed as expected."],
      table=[["Refresh setting", "Advice", "Why"], ["Refresh All before sending", "Always", "Numbers are current"],
             ["Refresh on open", "For shared reports", "Readers see today's data"],
             ["Background refresh", "Off for big queries", "Avoid half-loaded results"]],
      so="One click reruns the saved SQL, so the report never needs rebuilding.",
      split=4.20),

    O("3.3.8",
      what="A **parameter** lets one cell (e.g. the year) drive the query, so one query serves many reports. In "
           "Power Query, a parameter like **pYear** is spliced into the SQL. Never build SQL from free text typed "
           "by users: that allows **SQL injection**.",
      like="A parameter is like the blank on a form: same form every time, a different date filled in.",
      how=["Create a parameter (pYear) or a yellow input cell.", "Build the SQL with the parameter in the WHERE clause.",
           "Change the value and refresh."],
      tip="Allow only safe values (e.g. a list of years) in the parameter.",
      trap="Pasting user text straight into SQL, letting someone change the query.",
      case="The yellow cell on sheet 3.3 holds the year for a 'top sellers' query.",
      steps=["Set the year to 2024; refresh.", "Change it to 2025; refresh.",
             "Test what happens if someone types 2025' OR '1'='1."],
      table=[["Year typed", "Top seller", "Orders in scope"], ["2024", "Port Harcourt Depot 170 (₦12.1m)", "2,776"],
             ["2025", "Ibadan Market Online (₦19.6m)", "3,356"], ["2025' OR '1'='1", "Every order", "11,000 (injection)"]],
      so="One cell switched the report between years. The last row shows why free text must never reach SQL.",
      split=4.20),

    O("3.3.9",
      what="Most ODBC problems are one of four: **driver mismatch** (32/64-bit), **wrong path or file**, "
           "**credentials** not saved, or **SQL that doesn't fit the dialect or data types**, such as comparing "
           "text dates with numbers.",
      like="Troubleshooting is like a car that won't start: check fuel (driver), keys (credentials), the road "
           "(path) and the map (SQL) one at a time.",
      how=["Read the error message carefully.", "Match it to driver, path, credentials or SQL.", "Test the SQL alone in DB Browser."],
      tip="Test the SQL in DB Browser first; if it works there, the problem is the connection.",
      trap="Comparing text dates or years with numbers, which silently returns the wrong rows.",
      case="A 2025 query returns 0 rows, and another returns every row. Both run without error.",
      steps=["WHERE strftime('%Y', order_date) = 2025 (a number).", "WHERE order_date >= 20250101 (a number).",
             "Fix both by comparing text: '2025' and '2025-01-01'."],
      table=[["Condition", "Rows", "Verdict"], ["strftime('%Y', …) = 2025", "0", "Wrong: text vs number"],
             ["order_date >= 20250101", "11,000", "Wrong: text always ranks above numbers"],
             ["order_date >= '2025-01-01'", "3,356", "Correct"]],
      so="No error message appeared, yet two queries were wrong. Compare dates as text in SQLite.",
      split=4.20),
]
