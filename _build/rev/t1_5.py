from revlib import T

SALES = "SuperMart_Stores_Sales_2022_2025.csv"

TOPICS = [
    T("1.5.1",
      what="**Microsoft Excel** is a spreadsheet program: a big grid of boxes called **cells** where you store numbers and "
           "text, and write **formulas** that calculate from them. Its most powerful feature is automatic recalculation. "
           "Change one input and every result that depends on it updates instantly. Excel can also sort, filter, chart and "
           "summarise data.",
      like="Excel is like a calculator, a notebook and a drawing board joined together, and the calculator never forgets "
           "a sum. Change one figure in your ajo list and the totals redo themselves.",
      how=["Type values into cells (one value per cell).",
           "Write a formula that starts with = to calculate from those cells.",
           "Change an input and watch every dependent result update."],
      tip="Let formulas do the arithmetic, so results update when inputs change.",
      trap="Calculating on a phone calculator and typing the answer into Excel.",
      who="SuperMart Retail",
      case="The first sales line in the SuperMart file (Row_ID 1) is 2 message books sold at Abuja - Wuse for ₦16,448. "
           "The analyst wants the price per book and wants it to update if the quantity is corrected.",
      steps=["Open " + SALES + " and click cell V2 (the empty cell after Profit).",
             "Type =R2/S2 and press Enter: Sales ÷ Quantity.",
             "Change S2 from 2 to 4 and watch V2 update by itself."],
      table=[["Cell", "Holds", "Shows"],
             ["R2 (Sales)", "16448", "16448"],
             ["S2 (Quantity)", "2", "2"],
             ["V2", "=R2/S2", "8224"],
             ["V2 after S2 = 4", "=R2/S2", "4112"]],
      so="You typed the formula once and Excel redid the sum the moment an input changed. That is why analysts use "
         "Excel."),

    T("1.5.2",
      what="Excel comes in **editions**. Excel 2016, 2019 and 2021 are bought once and do not get new features. "
           "**Microsoft 365** is a subscription that gets new features every month. Newer functions such as "
           "**XLOOKUP**, **UNIQUE**, **FILTER** and **TEXTSPLIT** only work in newer editions, so a file built in 365 can "
           "show errors on an old laptop.",
      like="It is like phone software. An old phone that no longer updates can still call and text, but the newest "
           "apps simply won't install on it.",
      how=["Check your edition: File › Account › About Excel.",
           "Before you use a new function, check it exists in your edition.",
           "If you share files, use functions your colleagues' edition supports."],
      tip="Find out which Excel edition your organisation and your client use.",
      trap="Sending a 365-only formula to someone on Excel 2016; they will see #NAME?.",
      who="ETEH Learner",
      case="Chiamaka builds her Week 2 work in Microsoft 365 and emails it to her manager, who uses Excel 2016. Some "
           "cells show #NAME?. Which functions caused it?",
      steps=["List the functions used in the workbook.",
             "Check each one against the edition table below.",
             "Replace 365-only functions, or ask the manager to use Excel for the web."],
      table=[["Function", "2016 / 2019", "2021", "Microsoft 365"],
             ["SUM, AVERAGE, IF, VLOOKUP", "Yes", "Yes", "Yes"],
             ["XLOOKUP, UNIQUE, FILTER", "No", "Yes", "Yes"],
             ["TEXTSPLIT, VSTACK", "No", "No", "Yes"]],
      split=4.90,
      so="#NAME? usually means 'this Excel doesn't know that function'. Check editions before sharing modern "
         "formulas."),

    T("1.5.3",
      what="Excel runs in three main places. **Excel for Windows** (desktop) has every feature. **Excel for Mac** has "
           "most of them, but not Power Pivot. **Excel for the web** runs in a browser and is great for viewing and "
           "editing together, but it lacks some advanced tools. Know which one you are using before you follow a "
           "tutorial.",
      like="Think of a full kitchen, a small flat kitchen and a gas camping stove. You can cook in all three, but a big "
           "owambe meal needs the full kitchen.",
      how=["Check where you are working: browser, Windows or Mac.",
           "Use the web version to view and share.",
           "Use Windows desktop Excel for Power Query and Power Pivot work."],
      tip="Do the heavy analysis on Windows desktop Excel when you can.",
      trap="Following a Windows tutorial on Excel for the web and assuming a missing button is your mistake.",
      who="ETEH Learner",
      case="Tunde studies on a Mac at home and on a Windows laptop at work. Week 3 uses Power Pivot. Where should he do "
           "each task?",
      steps=["Read the feature table below.",
             "Plan Week 3 Power Pivot exercises for the Windows laptop.",
             "Use the Mac or the browser for formulas, tables and charts."],
      table=[["Feature", "Windows", "Mac", "Web"],
             ["Formulas, Tables, charts", "Yes", "Yes", "Yes"],
             ["Power Query", "Full", "Most", "Basic"],
             ["Power Pivot (Data Model)", "Yes", "No", "No"],
             ["Macros (VBA)", "Yes", "Yes", "No"]],
      split=4.90,
      so="Everyone can do Weeks 1–2 on any Excel. Power Pivot work in Week 3 needs Excel for Windows."),

    T("1.5.4",
      what="The **file format** (the letters after the dot) decides what a file can hold. **.xlsx** is the normal Excel "
           "workbook: formulas, formatting, many sheets. **.xlsm** also holds macros. **.xlsb** is a compact format for "
           "very big files. **.csv** is plain text: values only, one sheet, no formulas or colours. The old **.xls** "
           "holds only 65,536 rows.",
      like="Formats are like containers. A .csv is a nylon bag: it carries the goods but nothing fancy. A .xlsx is a "
           "proper box with compartments and labels.",
      how=["Look at the extension (turn on File name extensions in File Explorer).",
           "Open .csv files, then save your work as .xlsx (File › Save As).",
           "Never use .xls for large data."],
      tip="Receive data as .csv; save your analysis as .xlsx.",
      trap="Working in a .csv file; formulas, formatting and extra sheets are lost when it is saved.",
      who="Bootcamp Files",
      case="The pack contains many formats. Before Week 2 you should know what each file can and cannot hold, and "
           "whether it would fit in an old .xls file.",
      steps=["Open the 00_Core_Datasets folder and note each file's extension.",
             "Open Lifecare_Diabetes_Screening.csv and press Ctrl+Down to find the last row.",
             "Compare its rows with the .xls limit of 65,536."],
      table=[["File", "Format", "Rows", "Fits in .xls?"],
             ["ETEH_Week1_Section_Examples", ".xlsx", "many sheets", "n/a"],
             ["SuperMart_Stores_Sales", ".csv", "9,994", "Yes"],
             ["Lifecare_Admissions", ".csv", "54,860", "Yes"],
             ["Lifecare_Diabetes_Screening", ".csv", "100,000", "No (rows lost)"]],
      split=4.70,
      so="The diabetes file would lose 34,465 rows in an old .xls. Always work in .xlsx."),

    T("1.5.5",
      what="Good file habits protect you from the two worst mistakes: **losing work** and **overwriting the original "
           "data**. Keep raw files in a separate folder you never edit, save often with **Ctrl+S**, and use clear file "
           "names with the date and a version number. Then anyone can tell which file is newest.",
      like="It is like keeping your original certificates in a safe envelope and carrying only photocopies around. "
           "If a copy gets spoilt, the original is still there.",
      how=["Create folders: 00_RAW (never edit), 01_WORK, 02_REPORTS.",
           "Copy the raw file into 01_WORK before opening it.",
           "Save as YYYY-MM-DD_Topic_v1.xlsx; press Ctrl+S often."],
      tip="Never edit the file in your raw folder; always work on a copy.",
      trap="Names like 'final', 'final2', 'final_FINAL'; nobody knows which is newest.",
      who="SuperMart Retail",
      case="The analyst starts the SuperMart Q1 2025 review. She sets up her folders and names her files so that her "
           "manager can always find the latest version.",
      steps=["Put " + SALES + " in 00_RAW and copy it to 01_WORK.",
             "Open the copy and save it as an .xlsx with a date and version.",
             "After major changes, save as v2 instead of overwriting v1."],
      table=[["Item", "Good example"],
             ["Raw data", "00_RAW/SuperMart_Stores_Sales_2022_2025.csv"],
             ["Working file", "01_WORK/2025-04-07_SuperMart_Q1_v1.xlsx"],
             ["Next version", "01_WORK/2025-04-09_SuperMart_Q1_v2.xlsx"],
             ["Report sent", "02_REPORTS/2025-04-10_SuperMart_Q1_Report.pdf"]],
      split=4.30,
      so="Dated, versioned names and an untouched raw folder mean you can always go back to any step."),

    T("1.5.6",
      what="Excel has three safety nets. **AutoSave** saves your file continuously when it is stored in OneDrive or "
           "SharePoint. **AutoRecover** keeps a backup copy every few minutes in case Excel or your laptop crashes. "
           "**Version History** lets you open and restore earlier saved versions of a file.",
      like="AutoSave is like a phone that backs up every photo the moment you take it. AutoRecover is the spare "
           "copy. Version History is the album of older photos.",
      how=["Store work files in OneDrive and switch on AutoSave (top-left).",
           "Set AutoRecover to 5 minutes: File › Options › Save.",
           "Restore an older copy: File › Info › Version History."],
      tip="Set AutoRecover to every 5 minutes; power cuts happen.",
      trap="Relying on AutoRecover instead of saving; it is a backup, not a save.",
      who="ETEH Learner",
      case="During a NEPA outage Emeka's laptop battery dies while he is cleaning the SuperMart file. He had last pressed "
           "Ctrl+S an hour ago. What can he recover?",
      steps=["Restart Excel; the Document Recovery panel opens on the left.",
             "Open the most recent recovered version and check it.",
             "Save it at once with a new version number."],
      table=[["Safety net", "Where", "Protects against"],
             ["AutoSave", "OneDrive / SharePoint file", "Forgetting to save"],
             ["AutoRecover", "File › Options › Save", "Crash or power cut"],
             ["Version History", "File › Info", "A bad change you saved"]],
      split=4.90,
      so="With AutoRecover set to 5 minutes, Emeka loses at most 5 minutes of work, not an hour."),

    T("1.5.7",
      what="The Excel window has a few main areas. At the top are the **Title bar** and **Quick Access Toolbar**, then "
           "the **Ribbon** of commands. Below that are the **Name Box** (shows where you are) and the **Formula Bar** "
           "(shows what is really in the cell). Then comes the grid of cells, and at the bottom the **sheet tabs** and the "
           "**Status Bar**.",
      like="Like the dashboard of a car: speedometer, fuel gauge, mirrors. Once you know where each thing is, you stop "
           "searching and just drive.",
      how=["Click any cell and look at the Name Box (its address).",
           "Look at the Formula Bar (its true content).",
           "Select several numbers and read the Status Bar (their sum)."],
      tip="Glance at the Name Box and Formula Bar before editing anything.",
      trap="Trusting what a cell shows; the Formula Bar shows what it really holds.",
      who="SuperMart Retail",
      case="A new analyst opens the SuperMart sales file for the first time and wants to find their way around before "
           "doing any work.",
      steps=["Click cell R2 and read the Name Box and the Formula Bar.",
             "Click the Data tab on the Ribbon to see the Sort & Filter group.",
             "Select R2:R5 and read Sum at the bottom right (Status Bar)."],
      table=[["Area", "What you see after the steps"],
             ["Name Box", "R2"],
             ["Formula Bar", "16448"],
             ["Ribbon (Data tab)", "Sort, Filter, Get Data …"],
             ["Status Bar", "Sum of the 4 selected sales"]],
      split=5.20,
      so="Five areas tell you where you are, what a cell holds and what a selection adds up to, without any formula."),

    T("1.5.8",
      what="The **Ribbon** is the strip of commands at the top. It is organised into **tabs** (Home, Insert, Formulas, "
           "Data, View …). Each tab is divided into **groups** of related buttons. Extra tabs appear only when needed: "
           "click inside a table and a **Table Design** tab appears.",
      like="The Ribbon is like a big supermarket. Tabs are the aisles, groups are the shelves, and buttons are the items. "
           "Once you know the aisles, you find things fast.",
      how=["Click each tab once to see its groups.",
           "Learn where your most-used buttons live.",
           "Click inside a table or chart and look for the extra tab."],
      tip="Learn the Home, Insert and Data tabs first; they hold most daily tools.",
      trap="Hunting for a button that only appears when a table or chart is selected.",
      who="Bootcamp Map",
      case="This bootcamp uses some Ribbon tabs far more than others. Here is where each week's main tools live, so you "
           "know which aisle to walk to.",
      steps=["Open any bootcamp workbook.",
             "Click each tab in the table below and find the named button.",
             "Click inside a table (Ctrl+T) and find the Table Design tab."],
      table=[["Tab", "You will use it for", "Week"],
             ["Home", "Formatting, Sort & Filter", "1"],
             ["Formulas", "Functions, Name Manager", "1–2"],
             ["Data", "Get Data, Remove Duplicates", "2"],
             ["Insert", "PivotTables, charts", "2–3"]],
      split=4.90,
      so="Four tabs hold most of what you need. Learning them first saves hours of searching."),

    T("1.5.9",
      what="The **Quick Access Toolbar** (QAT) is a small row of icons above or below the Ribbon. You choose what goes on "
           "it. Press **Alt** and Excel shows a letter or number on every command (**KeyTips**). Press Alt then the "
           "number to use a QAT button without the mouse.",
      like="The QAT is like the speed-dial on your phone: your most-called contacts, one tap away instead of scrolling "
           "through the whole contact list.",
      how=["Right-click any Ribbon button › Add to Quick Access Toolbar.",
           "Press Alt and note the number shown on each QAT button.",
           "Use Alt + that number to run it."],
      tip="Pin the 5–6 commands you use most, in the order you use them.",
      trap="Pinning dozens of buttons; then the QAT is as slow as the Ribbon.",
      who="ETEH Learner",
      case="Ngozi cleans the SuperMart file every week. She repeats the same few commands, so she pins them to her QAT in "
           "this order.",
      steps=["Right-click Paste Values, Sort, Filter and Remove Duplicates; add each to the QAT.",
             "Press Alt and read the numbers.",
             "Practise: Alt+3 sorts the selected data."],
      table=[["QAT position", "Command", "Shortcut"],
             ["1", "Save", "Alt+1"],
             ["2", "Paste Values", "Alt+2"],
             ["3", "Sort", "Alt+3"],
             ["4", "Filter", "Alt+4"],
             ["5", "Remove Duplicates", "Alt+5"]],
      so="Five pinned commands, each one keystroke away, cut her weekly cleaning time noticeably."),

    T("1.5.10",
      what="Clicking **File** opens the **Backstage view**: a full page for working with the whole file rather than with "
           "cells. Here you Save As, Print, Export to PDF, see Version History, and set **properties** such as title and "
           "author. **Info › Check for Issues › Inspect Document** finds hidden sheets, comments and personal details "
           "before you share.",
      like="If the grid is the inside of a parcel, the Backstage view is the outside: the label, address, seal and "
           "receipt. You check them before sending.",
      how=["Click File › Info and fill in Title and Author.",
           "Run Check for Issues › Inspect Document before sharing.",
           "Use Export › Create PDF for read-only reports."],
      tip="Inspect every file before you send it outside the team.",
      trap="Sending a workbook that still contains hidden sheets or private comments.",
      who="SuperMart Retail",
      case="The Q1 review workbook is ready for the CEO. It still has a hidden scratch sheet and notes with customer "
           "phone numbers. The analyst prepares it for sending.",
      steps=["File › Info › Check for Issues › Inspect Document › Inspect.",
             "Remove hidden worksheets and comments that should not go out.",
             "File › Export › Create PDF for a version that cannot be edited."],
      table=[["Backstage area", "What it does"],
             ["Info", "Properties, protection, Version History"],
             ["Save As", "New name, folder or format"],
             ["Export", "PDF copy for readers"],
             ["Inspect Document", "Finds hidden or private content"]],
      split=4.90,
      so="Two minutes in the Backstage view stops private data or scratch work from reaching the CEO's inbox."),

    T("1.5.11",
      what="The **Name Box** (left of the formula bar) shows the address of the selected cell, such as R2. Type an "
           "address in it and press Enter to jump there instantly. The **Formula Bar** shows what the cell really "
           "contains: a typed value or a formula. What the cell *displays* can look different because of formatting.",
      like="The Name Box is like your house address: type it into Google Maps and you go straight there. The Formula "
           "Bar is the house's title document: the truth behind the paint.",
      how=["Click in the Name Box, type an address (e.g. A9995), press Enter.",
           "Click a cell and read the Formula Bar.",
           "Compare it with what the cell displays."],
      tip="Check the Formula Bar when a number looks odd; formatting can hide decimals.",
      trap="Scrolling for minutes through thousands of rows instead of typing the address.",
      who="SuperMart Retail",
      case="The SuperMart sales file has 9,994 lines plus a header. The analyst wants to jump to the last row at once and "
           "check what the final Sales cell holds.",
      steps=["Click the Name Box, type R9995 and press Enter.",
             "Read the Formula Bar to see the stored value.",
             "Type R2:R9995 in the Name Box to select the whole Sales column."],
      table=[["Type in Name Box", "Result"],
             ["R9995", "Jumps to the last Sales cell"],
             ["R2:R9995", "Selects all 9,994 sales"],
             ["A1", "Back to the top-left cell"]],
      so="Typing an address is faster than scrolling, and the Formula Bar always shows the true value."),

    T("1.5.12",
      what="**Column headers** are the letters across the top (A, B, C … XFD). **Row headers** are the numbers down the "
           "side (1, 2, 3 …). Click a header to select the whole column or row. The **active cell** is the one with "
           "the thick border: it receives what you type next, even when a bigger range is selected.",
      like="The headers are like seat numbers in a stadium: block letter and row number. 'Block R, row 2' takes you to "
           "exactly one seat.",
      how=["Click a column letter to select the whole column.",
           "Press Ctrl+Space (column) or Shift+Space (row) from any cell.",
           "Double-click a column's right border to auto-fit its width."],
      tip="Double-click column borders to fit the widest value.",
      trap="Typing while a range is selected without noticing which cell is active.",
      who="SuperMart Retail",
      case="In the SuperMart file the Product_Name column (Q) is too narrow to read, and the analyst wants to format the "
           "whole Sales column (R) as naira.",
      steps=["Double-click the border between Q and R headers to auto-fit Product_Name.",
             "Click header R to select the whole Sales column.",
             "Press Ctrl+1 › Number › Currency › ₦ and click OK."],
      table=[["Action", "Shortcut", "Selects"],
             ["Click header R", "Ctrl+Space", "All of column R"],
             ["Click row number 2", "Shift+Space", "All of row 2"],
             ["Click corner box", "Ctrl+A", "The whole sheet"]],
      so="Header clicks act on a whole column or row in one move. There is no need to drag across 9,994 cells."),

    T("1.5.13",
      what="The **Status Bar** at the bottom of the window shows quick statistics for any cells you select: **Sum**, "
           "**Average** and **Count**, and you can add Minimum and Maximum. It is the fastest way to check a total without "
           "writing a formula. Right-click the Status Bar to choose what it shows.",
      like="It is like the digital display on a fuel pump: as the fuel flows you see litres and naira instantly, without "
           "doing any sums yourself.",
      how=["Right-click the Status Bar and tick Minimum and Maximum.",
           "Select a column of numbers.",
           "Read Sum, Average, Count, Min and Max at the bottom right."],
      tip="Use the Status Bar to double-check any total before you report it.",
      trap="A Count higher than expected often means a header or text cells were selected.",
      who="SuperMart Retail",
      case="Before building any report, the analyst wants a quick check of the whole SuperMart Sales column for "
           "2022–2025, without writing formulas.",
      steps=["Click R2, then press Ctrl+Shift+Down to select R2:R9995.",
             "Read the Status Bar.",
             "Compare Count with the number of rows you expect (9,994)."],
      table=[["Status Bar shows", "Value (R2:R9995)"],
             ["Sum", "₦2,297,200,862"],
             ["Average", "₦229,858"],
             ["Count", "9,994"],
             ["Min / Max", "₦444 / ₦22,638,480"]],
      so="In two seconds you know the total (₦2.30bn over 4 years) and that no rows are missing."),

    T("1.5.14",
      what="Every worksheet is a grid of **1,048,576 rows** and **16,384 columns** (A to XFD). Each cell has an address: "
           "column letter plus row number, such as C7. The grid is huge, but it has a limit. If a file has more rows, Excel "
           "cuts them off when it opens it, and you may not notice.",
      like="A sheet is like a very big exercise book with a fixed number of pages. Write more than it holds and the extra "
           "simply falls off the end.",
      how=["Press Ctrl+Down in a full column to see the last row used.",
           "Compare it with the number of records you expect.",
           "Use Power Query / the Data Model (Week 2–3) for data above the limit."],
      tip="After opening big files, always check the last row number.",
      trap="Opening a file with more than 1,048,576 rows and not noticing that rows were lost.",
      who="Bootcamp Files",
      case="How much of the Excel grid do the bootcamp files use? The analyst checks the largest ones against the row "
           "limit.",
      steps=["Open each file and press Ctrl+Down from cell A1.",
             "Note the last row (data rows + 1 header).",
             "Work out what share of the 1,048,576-row limit it uses."],
      table=[["File", "Data rows", "Share of limit"],
             ["Lifecare_Diabetes_Screening", "100,000", "9.5%"],
             ["Lifecare_Admissions_2019_2024", "54,860", "5.2%"],
             ["EkoSuites_Hotel_Bookings", "36,275", "3.5%"],
             ["SuperMart_Stores_Sales", "9,994", "1.0%"]],
      split=4.70,
      so="All bootcamp files fit easily, but you now know how to check. Some real company files do not fit."),

    T("1.5.15",
      what="Fast **navigation** saves hours. **Ctrl+Arrow** jumps to the edge of the data, **Ctrl+Home** goes back to A1, "
           "**Ctrl+End** goes to the last used cell, **Ctrl+PageDown/PageUp** moves between sheets, and **Ctrl+G** opens "
           "Go To. Add **Shift** to any jump to select everything along the way.",
      like="It is like taking the expressway instead of the side streets. You reach the same place, but in seconds "
           "instead of half an hour.",
      how=["From A1, press Ctrl+Down to reach the last row.",
           "Press Ctrl+Right to reach the last column.",
           "Add Shift to select while you jump."],
      tip="Learn Ctrl+Arrow and Ctrl+Shift+Arrow first; you will use them every day.",
      trap="Blank cells stop Ctrl+Arrow early; you may think the data ends there.",
      who="SuperMart Retail",
      case="The analyst needs to know the size of the SuperMart file and select the whole Sales column, without "
           "scrolling.",
      steps=["Click A1 and press Ctrl+Down (last row), then Ctrl+Right (last column).",
             "Click R2 and press Ctrl+Shift+Down to select all sales.",
             "Press Ctrl+Home to go back to A1."],
      table=[["Keys", "From", "Lands on"],
             ["Ctrl+Down", "A1", "A9995 (last row)"],
             ["Ctrl+Right", "A1", "U1 (Profit, last column)"],
             ["Ctrl+Shift+Down", "R2", "Selects R2:R9995"],
             ["Ctrl+Home", "anywhere", "A1"]],
      so="Four shortcuts tell you the file is 9,994 rows × 21 columns and select a whole column, in seconds."),

    T("1.5.16",
      what="A **workbook** is the file. **Worksheets** are its pages, shown as tabs at the bottom. Right-click a tab to "
           "rename, colour, move, copy, hide or protect it. A professional workbook has a clear order: a cover or README "
           "first, then raw data, then working sheets, then the report.",
      like="A workbook is like a file folder in an office. The sheets are the documents inside, arranged in a sensible "
           "order with labels so anyone can find what they need.",
      how=["Double-click a tab to rename it with a clear name.",
           "Right-click › Tab Color to group related sheets.",
           "Drag tabs into order: cover, data, work, report."],
      tip="Start every workbook with a cover sheet that explains what is inside.",
      trap="Leaving tabs named Sheet1, Sheet2, Sheet3.",
      who="Bootcamp Files",
      case="Open ETEH_Week2_Section_Examples.xlsx. Its sheets follow a clear structure that you should copy in your own "
           "work.",
      steps=["Look at the tab order from left to right.",
             "Notice that the cover comes first and the data sheets (D_) come last.",
             "Press Ctrl+PageDown to move through the tabs one by one."],
      table=[["Tabs", "Role"],
             ["00_Cover", "Contents, links to every sheet"],
             ["2.1 External Formats … 2.11 PivotCharts", "One sheet per module"],
             ["D_SuperMart_All, D_Sales_2025 …", "Data sheets used by the modules"]],
      split=4.90,
      so="A cover first, lessons in order and data at the end: anyone can open the file and know where to go."),

    T("1.5.17",
      what="You can **customise** Excel to work faster: pin favourite commands to the Quick Access Toolbar, hide the "
           "Ribbon to get more screen space (**Ctrl+F1**), show extra tabs such as Developer, and change the default font "
           "and number of sheets for new workbooks. Small settings save time every day.",
      like="Like adjusting the seat and mirrors before driving someone else's car. It takes a minute, but the whole "
           "journey is safer and more comfortable.",
      how=["File › Options › General: set the default font and sheet count.",
           "File › Options › Customize Ribbon: tick Developer.",
           "Press Ctrl+F1 to hide or show the Ribbon."],
      tip="Customise once on every computer you work on.",
      trap="Changing so much that you can't follow tutorials or help a colleague.",
      who="ETEH Learner",
      case="Bola works on a small 13-inch laptop and keeps running out of screen space when viewing the 21 columns of "
           "the SuperMart file.",
      steps=["Press Ctrl+F1 to collapse the Ribbon.",
             "Set View › Zoom to 85%.",
             "Freeze the header row: View › Freeze Panes › Freeze Top Row."],
      table=[["Setting", "Where", "Benefit"],
             ["Hide Ribbon", "Ctrl+F1", "More rows on screen"],
             ["Zoom 85%", "View › Zoom", "More columns on screen"],
             ["Freeze top row", "View › Freeze Panes", "Headers stay visible"]],
      split=4.90,
      so="Three quick settings make a small laptop screen workable for a 21-column file."),
]
