"""Week 1 edition-3 slide copy: Modules 1.5 (Excel interface) and 1.6 (working with data)."""
from v3kit import T

M15 = dict(
    id="1.5", title="Introduction to Microsoft Excel",
    subtitle="Editions, file formats, saving safely, the parts of the Excel window and moving around fast.",
    objectives=["Explain what Excel is and choose the right edition, platform and file format.",
                "Create, open and save workbooks safely, using AutoSave, AutoRecover and Version History.",
                "Name every part of the Excel window: ribbon, Quick Access Toolbar, Name Box, formula bar, status bar.",
                "Navigate large sheets quickly with the keyboard and organise worksheets.",
                "Customise the interface to work faster."],
    topics=[
        T("1.5.1", "What Microsoft Excel Is",
          defn="Microsoft Excel is a spreadsheet application: a grid of cells, organised into worksheets and "
               "workbooks, used to store, calculate, analyse and visualise data.",
          simple="Excel is a giant, intelligent sheet of squared paper. Each square (cell) can hold a value or a "
                 "formula, and formulas recalculate instantly when the values change.",
          like="Like a market trader's notebook that adds itself up: change one sale and every total updates.",
          grid=dict(label="What analysts use Excel for", widths=[1.4, 3.6], rows=[
              ["Task", "Example in this bootcamp"],
              ["Store and clean data", "Tidy the SuperMart sales export into an Excel table"],
              ["Calculate", "Margin, target attainment, length of stay"],
              ["Summarise", "Sales by store with SUMIFS or a PivotTable (Week 2)"],
              ["Visualise and share", "Charts and a one-page dashboard (Weeks 2-3)"]])),

        T("1.5.2", "Excel Editions: 2016, 2019, 2021 and Microsoft 365",
          simple="Excel 2016, 2019, 2021 and 2024 are bought once and never gain new features. Microsoft 365 is a "
                 "subscription that keeps getting new features. Newer functions do not work in older editions.",
          like="Like phone software: an old phone that no longer updates still calls and texts, but cannot run the "
               "newest apps.",
          grid=dict(label="What each edition can do", widths=[1.4, 1.5, 3.2], size=12, rows=[
              ["Edition", "Licence", "Notable functions available"],
              ["Excel 2016", "One-off purchase", "Core functions, Power Query (Get & Transform), PivotTables"],
              ["Excel 2019", "One-off purchase", "Adds IFS, SWITCH, TEXTJOIN, CONCAT, MAXIFS, MINIFS"],
              ["Excel 2021", "One-off purchase", "Adds XLOOKUP, FILTER, SORT, UNIQUE, LET (dynamic arrays)"],
              ["Microsoft 365", "Subscription", "All of the above plus the newest (TEXTSPLIT, LAMBDA; Copilot as an add-on)"]]),
          watch="Support for Office 2016 and 2019 ended in October 2025. If a file uses XLOOKUP, check that everyone "
                "who opens it has Excel 2021 or 365."),

        T("1.5.3", "Excel Platforms: Web, Desktop and Mac",
          simple="The same workbook opens on Windows, Mac and in a browser, but not every feature exists everywhere. "
                 "This bootcamp assumes Excel for Windows (desktop).",
          grid=dict(label="Platforms compared", widths=[1.6, 2, 2, 2], size=12, rows=[
              ["", "Windows desktop", "Mac desktop", "Excel for the web"],
              ["Cost", "Microsoft 365 or one-off licence", "Microsoft 365 or one-off licence", "Free with a Microsoft account"],
              ["Power Query", "Full", "Yes, fewer data sources", "Limited"],
              ["Power Pivot / data model", "Yes", "No", "No"],
              ["VBA macros", "Run and edit", "Run and edit", "Do not run"],
              ["Shortcut key", "Ctrl", "Cmd (mostly)", "Ctrl (browser may intercept some)"]])),

        T("1.5.4", "Excel File Formats: .xlsx, .xlsm, .xls, .xlsb and .csv",
          simple="The file extension decides what the file can hold. Use .xlsx by default; .csv only for moving raw "
                 "data between systems.",
          grid=dict(label="Which format, when", widths=[0.9, 2.8, 2.4], size=12, rows=[
              ["Format", "What it holds", "Use it for"],
              [".xlsx", "Default workbook: sheets, formulas, charts; no macros", "Almost everything"],
              [".xlsm", "Workbook with VBA macros", "Files that need macros"],
              [".xls", "Old 97-2003 format: 65,536 rows × 256 columns", "Opening very old files only"],
              [".xlsb", "Binary workbook: smaller and faster to open", "Very large workbooks"],
              [".csv", "Plain text, one sheet, values only: no formulas or formatting", "Exporting and importing raw data"]]),
          watch="Saving as .csv keeps only the active sheet and throws away formulas and formatting. Keep a .xlsx copy."),

        T("1.5.5", "Creating, Opening and Saving Workbooks Safely",
          simple="Save early, save often, and use clear file names. Most lost work comes from never saving or from "
                 "overwriting the only copy.",
          grid=dict(label="Shortcuts", widths=[1.2, 2.2], rows=[
              ["Keys", "Action"], ["Ctrl+N", "New blank workbook"], ["Ctrl+O", "Open"], ["Ctrl+S", "Save"],
              ["F12", "Save As (new name or format)"], ["Ctrl+W", "Close the workbook"]]),
          points=["Name files clearly: SuperMart_Sales_FY2025_v2.xlsx", "Keep the raw export untouched; work on a copy",
                  "Downloaded files open in Protected View: enable editing only if you trust the source",
                  "Save to OneDrive or SharePoint to switch on AutoSave"],
          points_label="SAFE HABITS", order=["grid", "points"], split=0.4),

        T("1.5.6", "AutoSave, AutoRecover and Version History",
          simple="Three different safety nets. Know which one you have, because AutoRecover is not a substitute for "
                 "saving.",
          grid=dict(label="Three safety nets", widths=[1.4, 2.6, 2.4], size=12, rows=[
              ["Feature", "What it does", "Where / when"],
              ["AutoSave", "Saves every change automatically, every few seconds", "Files stored on OneDrive or SharePoint (Microsoft 365)"],
              ["AutoRecover", "Keeps a recovery copy in case Excel crashes", "Every 10 minutes by default: File › Options › Save"],
              ["Version History", "Lets you open or restore an earlier version", "File › Info › Version History (OneDrive or SharePoint)"]]),
          watch="With AutoSave on, every change is saved at once. Use Save a Copy before experimenting, or restore an "
                "earlier version from Version History."),

        T("1.5.7", "Excel Window Anatomy",
          simple="Learn the names of the parts once; every later lesson refers to them.",
          diagram=("excel", dict(labels={"Title bar": "file name and AutoSave", "Ribbon": "tabs, groups and commands",
                                         "Name Box": "address of the active cell", "Formula bar": "contents of the active cell",
                                         "Grid": "columns A, B, C… and rows 1, 2, 3…",
                                         "Sheet tabs": "the worksheets in the workbook",
                                         "Status bar": "mode, quick totals, views and zoom"},
                                 highlight=())), diagram_label="The Excel window"),

        T("1.5.8", "The Ribbon: Tabs, Groups and Galleries",
          defn="The ribbon is the command bar across the top of Excel. It is divided into tabs; each tab holds groups "
               "of related commands, and some commands open galleries of ready-made options.",
          diagram=("excel", dict(highlight=("ribbon",), labels={
              "Tab": "Home: formatting · Insert: tables, charts · Formulas: functions · Data: import, sort, filter", "Group": "Font, Alignment, Number…",
              "Gallery": "Cell Styles, Table Styles: click a preview to apply",
              "Dialog launcher": "small arrow at a group's corner opens the full options",
              "Contextual tab": "appears only when needed, e.g. Table Design",
              "Ctrl+F1": "hide or show the ribbon"})), diagram_label="Ribbon parts"),

        T("1.5.9", "Quick Access Toolbar and KeyTips",
          defn="The Quick Access Toolbar (QAT) is a small, customisable toolbar for your most-used commands. KeyTips "
               "are letters shown on the ribbon when you press Alt, so any command can be run from the keyboard.",
          simple="Put the commands you use most on the QAT. Press Alt and follow the letters to run any ribbon command "
                 "without the mouse.",
          grid=dict(label="KeyTips worth learning", widths=[1.3, 2.4], rows=[
              ["Keys", "Command"], ["Alt, H, O, I", "AutoFit column width"], ["Alt, W, F, F", "Freeze Panes"],
              ["Alt, A, S, S", "Sort dialog"], ["Alt+= (together)", "AutoSum"]]),
          steps=["Click the small arrow at the end of the QAT", "Tick a command, or choose More Commands",
                 "Pick any command and click Add"],
          steps_label="ADD A COMMAND TO THE QAT", order=["grid", "steps"]),

        T("1.5.10", "Backstage View and Workbook Properties",
          defn="Backstage view is the full-screen File menu where you manage the workbook as a file: Info, New, Open, "
               "Save As, Print, Share, Export and Options.",
          simple="Everything you do to the file (not to the cells) happens behind the File tab.",
          grid=dict(label="Backstage view: what is where", widths=[1.2, 3.4], size=12, rows=[
              ["Page", "Use it to"],
              ["Info", "See properties, protect the workbook, inspect for hidden data, open Version History"],
              ["Save As / Save a Copy", "Save under a new name, place or format (.xlsx, .csv, PDF)"],
              ["Print", "Preview and print (Ctrl+P)"],
              ["Options", "Change Excel settings: AutoRecover, default font, ribbon"]]),
          points=["Title and Tags: help people find the file", "Author and Last Modified By", "Created and modified dates"],
          points_label="WORKBOOK PROPERTIES (File › Info)", order=["grid", "points"], split=0.62),

        T("1.5.11", "Name Box and Formula Bar",
          defn="The Name Box shows the address or name of the active cell or range. The formula bar shows the actual "
               "contents of the active cell: the typed value or the formula behind a result.",
          diagram=("excel", dict(highlight=("namebox", "formulabar"), labels={
              "Rule": "the cell shows the answer, the formula bar shows how it was worked out: read it before trusting a result",
              "Name Box": "type K500 and press Enter to jump there", "Select a range": "type A1:D20 and press Enter",
              "Name a range": "select cells, type a name, press Enter",
              "Formula bar": "shows the formula; edit it here or press F2 in the cell",
              "Ctrl+Shift+U": "expand the formula bar for long formulas"})), diagram_label="Name Box and formula bar"),

        T("1.5.12", "Column Headers, Row Headers and the Active Cell",
          simple="Columns are lettered (A, B, C…), rows are numbered (1, 2, 3…). A cell's address is its column then "
                 "its row, like C4. The active cell is the one with the bold border: what you type goes there.",
          diagram=("excel", dict(highlight=("colhdr", "rowhdr", "active"), active="C4", labels={
              "Column header": "click to select the whole column (Ctrl+Space)",
              "Row header": "click to select the whole row (Shift+Space)",
              "Active cell": "C4: column C, row 4; shown in the Name Box",
              "Resize": "drag a header edge; double-click to AutoFit"})), diagram_label="Headers and the active cell"),

        T("1.5.13", "Status Bar and View Modes",
          simple="The status bar at the bottom shows Excel's mode and instant statistics for whatever you select: no "
                 "formula needed.",
          diagram=("excel", dict(highlight=("status",), labels={
              "Mode": "Ready, Enter or Edit",
              "Quick statistics": "select cells to see Average, Count and Sum; right-click to add Min and Max",
              "Normal view": "everyday working", "Page Layout view": "see pages, headers and footers",
              "Page Break Preview": "drag the blue page breaks", "Zoom": "slider, or Ctrl + mouse wheel"})),
          diagram_label="Status bar and views",
          watch="Quick statistics are a check, not a result: they vanish when you click elsewhere. Write a formula for "
                "anything you report."),

        T("1.5.14", "The Grid: 1,048,576 Rows × 16,384 Columns",
          simple="A worksheet has 1,048,576 rows and 16,384 columns (A to XFD): over 17 billion cells. That is far "
                 "more than most datasets need, but very large files become slow.",
          grid=dict(label="Grid limits", widths=[2.4, 2.2], rows=[
              ["Limit", "Value"], ["Rows per sheet", "1,048,576"], ["Columns per sheet", "16,384 (last is XFD)"],
              ["Characters in one cell", "32,767"], ["Old .xls format", "65,536 rows × 256 columns"]]),
          points=["Ctrl+Down Arrow jumps to the last filled row of a column", "Ctrl+End jumps to the last used cell",
                  "Our largest dataset (54,860 admissions) uses about 5% of the rows"],
          points_label="SEE IT YOURSELF", order=["grid", "points"], split=0.5),

        T("1.5.15", "Navigating Efficiently",
          simple="The mouse is slow on large data. A handful of shortcuts move you anywhere instantly.",
          grid=dict(label="Navigation shortcuts", widths=[1.8, 3.2], size=12.5, rows=[
              ["Keys", "Moves to / selects"],
              ["Ctrl+Arrow", "The edge of the current block of data"],
              ["Ctrl+Shift+Arrow", "Selects to the edge of the data"],
              ["Ctrl+Home / Ctrl+End", "Cell A1 / the last used cell"],
              ["Ctrl+Page Down / Page Up", "Next / previous worksheet"],
              ["F5 or Ctrl+G", "Go To a cell, range or name"],
              ["Ctrl+A", "Select the current region (press again: whole sheet)"]]),
          watch="Ctrl+Arrow stops at the first blank cell. A blank inside a column makes it stop early, which is "
                "another reason to avoid gaps in data."),

        T("1.5.16", "Worksheets, Sheet Tabs and Workbook Structure",
          defn="A workbook is an Excel file; a worksheet is one grid inside it, reached by its sheet tab.",
          simple="Organise a workbook like a report: raw data on its own sheet, workings next, summary last, with clear "
                 "tab names.",
          grid=dict(label="Sheet tab actions", widths=[1.6, 2.4], size=12, rows=[
              ["Action", "How"], ["Insert", "Shift+F11 or the + button"], ["Rename", "Double-click the tab"],
              ["Move or copy", "Right-click › Move or Copy"], ["Colour", "Right-click › Tab Color"],
              ["Hide / unhide", "Right-click › Hide / Unhide"]]),
          points=["Cover or README: what the file is", "Data: the raw import, never edited",
                  "Workings: calculations", "Report: the summary people read"],
          points_label="A GOOD STRUCTURE", order=["grid", "points"], split=0.55,
          watch="Clicking several tabs groups them: anything you type then goes into every grouped sheet. Click one "
                "tab to ungroup."),

        T("1.5.17", "Customising the Excel Interface",
          simple="A few settings make Excel faster and safer for analysis work. All of them live in File › Options.",
          grid=dict(label="Useful settings", widths=[2.2, 3.2], size=12.5, rows=[
              ["Setting", "Where"],
              ["Add your own tab or group", "Options › Customize Ribbon"],
              ["Add commands to the QAT", "Options › Quick Access Toolbar"],
              ["AutoRecover every 5 minutes", "Options › Save"],
              ["Default font and new-sheet count", "Options › General"],
              ["Show or hide gridlines and headings", "View tab › Show group"],
              ["Dark or light theme", "File › Account › Office Theme"]])),
    ],
    takeaways=["Use .xlsx by default, .csv only for raw data exchange; newer functions need Excel 2021 or 365.",
               "Save early and often; know the difference between AutoSave, AutoRecover and Version History.",
               "Ribbon → tabs → groups → commands; Alt shows KeyTips for every command.",
               "Name Box = where you are; formula bar = what the cell really contains.",
               "Move with Ctrl+Arrow, Ctrl+Home/End and F5; keep raw data, workings and report on separate sheets."],
)

M16 = dict(
    id="1.6", title="Working with Data in Excel",
    subtitle="Entering, editing and formatting data correctly, and the tools that keep large sheets readable and clean.",
    objectives=["Enter text, numbers and dates so Excel stores the correct data type.",
                "Edit, insert, delete, copy and paste data, including Paste Special.",
                "Fill series and split or join text with AutoFill and Flash Fill.",
                "Format numbers, dates, currency and alignment; apply styles and conditional formatting.",
                "Use Freeze Panes, page setup, Find & Replace, Go To Special and data validation."],
    topics=[
        T("1.6.1", "Entering Text, Numbers and Dates",
          simple="Type into the active cell and press Enter (moves down) or Tab (moves right). Excel decides the type "
                 "from what you type, and shows it by alignment: text left, numbers and dates right.",
          grid=dict(label="What you type → what Excel stores", widths=[1.4, 1.6, 2.6], size=12.5, rows=[
              ["You type", "Excel stores", "Note"],
              ["Lagos - Ikeja", "Text", "Left-aligned"],
              ["474430", "Number", "Right-aligned; format it later"],
              ["2025-01-01", "Date (serial 45658)", "ISO format is never misread"],
              ["01/02/2025", "Date: 1 Feb or 2 Jan?", "Depends on the computer's region"],
              ["007", "Number 7", "Type '007 to keep the zeros as text"]]),
          points=["Ctrl+; enters today's date", "Ctrl+Shift+; enters the time", "Alt+Enter starts a new line in a cell",
                  "Esc cancels an entry"],
          points_label="HANDY KEYS", order=["grid", "points"], split=0.64),

        T("1.6.2", "The Six Excel Data Types",
          defn="A data type is the kind of value a cell holds. Excel recognises six: Number, Text, Date/Time, Logical, "
               "Error and Array.",
          simple="The type decides what formulas can do. SUM adds numbers but silently skips text, so a number stored "
                 "as text drops out of the total.",
          like="Like a bus head count: if a child's ticket is recorded as 'luggage', the passenger total is wrong even "
               "though the child is on the bus.",
          grid=dict(label="The six types", widths=[1.2, 2.2, 2.4], size=12, rows=[
              ["Type", "Example", "Test"],
              ["Number", "474430, 0.2, ₦733,215,255", "=ISNUMBER(A2)"],
              ["Text", "Lagos - Ikeja, '007", "=ISTEXT(A2)"],
              ["Date/Time", "2025-01-01 (stored as 45658)", "Changes when formatted as Number"],
              ["Logical", "TRUE, FALSE", "=ISLOGICAL(A2)"],
              ["Error", "#DIV/0!, #N/A, #VALUE!", "=ISERROR(A2)"],
              ["Array", "A list that spills, e.g. =UNIQUE(...)", "Blue border around the spill"]])),

        T("1.6.3", "Editing, Replacing and Clearing Cell Contents",
          simple="Typing over a cell replaces it. To change only part of it, edit in place. 'Delete' removes contents "
                 "but keeps formats; Clear lets you choose.",
          grid=dict(label="Edit and clear", widths=[1.8, 3.2], size=12.5, rows=[
              ["Action", "How"],
              ["Replace the whole cell", "Select it and type"],
              ["Edit part of a cell", "F2 (or double-click), edit, Enter"],
              ["Cancel an edit", "Esc"],
              ["Delete contents only", "Delete key (formats stay)"],
              ["Clear formats, contents, comments or all", "Home › Editing › Clear"],
              ["Undo / redo", "Ctrl+Z / Ctrl+Y"]])),

        T("1.6.4", "Inserting, Deleting, Hiding and Resizing Rows and Columns",
          simple="Insert and delete shift the cells around them; formulas adjust automatically. Hiding keeps the data "
                 "but takes it out of view.",
          grid=dict(label="Row and column actions", widths=[1.5, 3], size=12.5, rows=[
              ["Action", "Shortcut / how"],
              ["Insert", "Select row(s) or column(s), Ctrl+Plus"],
              ["Delete", "Select row(s) or column(s), Ctrl+Minus"],
              ["Hide", "Ctrl+9 (rows), Ctrl+0 (columns), or right-click › Hide"],
              ["Unhide", "Select both sides, right-click › Unhide"],
              ["AutoFit width", "Double-click the right edge of the column header"]]),
          watch="Hidden rows are still included by SUM, and easy to miss. Delete data you do not need; use filters to "
                "hide temporarily."),

        T("1.6.5", "Copy, Cut, Paste and Paste Special",
          defn="Paste Special pastes only selected parts of what was copied, such as values, formats or formulas, or "
               "transforms it as it pastes (transpose, add, multiply).",
          simple="Ordinary paste copies everything. Paste Special lets you choose, for example 'values only' to freeze "
                 "formula results.",
          grid=dict(label="Paste Special options (Ctrl+Alt+V)", widths=[1.4, 3.2], size=12.5, rows=[
              ["Option", "Use it to"],
              ["Values", "Replace formulas with their results"],
              ["Formats", "Copy formatting only"],
              ["Formulas", "Copy formulas without formatting"],
              ["Transpose", "Turn rows into columns"],
              ["Multiply / Divide", "Convert in bulk, e.g. divide by 1,000,000 for millions"]]),
          points=["Ctrl+C copy, Ctrl+X cut, Ctrl+V paste", "Cut moves the cells; formulas pointing to them follow",
                  "Esc clears the moving dashed border"],
          points_label="BASICS", order=["grid", "points"], split=0.62),

        T("1.6.6", "AutoFill, Series and Custom Lists",
          defn="AutoFill extends a pattern from selected cells into neighbouring cells by dragging the fill handle "
               "(the small square at the bottom-right of the selection).",
          simple="Type the start of a pattern, then drag or double-click the fill handle and Excel continues it.",
          grid=dict(label="Start with → AutoFill gives", widths=[1.6, 2.4], size=12.5, rows=[
              ["Start with", "Fill gives"], ["Jan", "Feb, Mar, Apr…"], ["Q1", "Q2, Q3, Q4, Q1…"],
              ["ST01", "ST02, ST03… ST12"], ["2025-01-01", "2025-01-02, 2025-01-03…"], ["5, 10", "15, 20, 25…"],
              ["A formula", "The same formula, adjusted per row"]]),
          steps=["Double-click the fill handle to fill down to the end of the data", "Home › Fill › Series sets a step and "
                 "stop value", "File › Options › Advanced › Edit Custom Lists adds your own list, e.g. the four zones"],
          steps_label="MORE CONTROL", order=["grid", "steps"], split=0.45),

        T("1.6.7", "Flash Fill (Ctrl+E)",
          defn="Flash Fill detects a pattern from one or two examples you type and fills the rest of the column to "
               "match.",
          simple="Show Excel what you want once; Flash Fill does the other thousand rows.",
          grid=dict(label="Examples from our data", widths=[2, 1.6, 2], size=12.5, rows=[
              ["Source value", "You type", "Flash Fill continues"],
              ["Lagos - Ikeja (Store_Name)", "Lagos", "Port Harcourt, Abuja, Jos…"],
              ["Gbenga Anyanwu (Customer_Name)", "Gbenga", "Adaeze, Abimbola, Femi…"],
              ["SM-2025-144463 (Order_ID)", "2025", "The year of every order"]]),
          steps=["Type the first result next to the data", "Move to the cell below", "Press Ctrl+E (or Data › Flash Fill)",
                 "Check several rows: Flash Fill guesses"],
          steps_label="HOW", order=["grid", "steps"], split=0.6,
          watch="Flash Fill results are typed values, not formulas: they do not update if the source changes."),

        T("1.6.8", "Formatting Cells with Ctrl+1",
          defn="Cell formatting changes how a value looks (font, borders, fill, number format, alignment) without "
               "changing the value stored.",
          simple="Ctrl+1 opens Format Cells, which has every formatting option in one place.",
          grid=dict(label="The Format Cells tabs", widths=[1.2, 3.2], size=12.5, rows=[
              ["Tab", "Controls"], ["Number", "How numbers, money, percentages and dates display"],
              ["Alignment", "Position, wrap text, orientation"], ["Font", "Typeface, size, bold, colour"],
              ["Border", "Lines around cells"], ["Fill", "Background colour"], ["Protection", "Locked and hidden cells"]]),
          watch="Formatting never changes the value: 0.127438 shown as 12.7% is still 0.127438 in calculations."),

        T("1.6.9", "Formatting Dates, Currency, Percentages and Decimals",
          simple="The same stored number can look completely different. Choose the format that makes the meaning "
                 "obvious.",
          grid=dict(label="One value, many formats", widths=[1.3, 1.9, 2.4, 1.6], size=12, rows=[
              ["Stored value", "Format", "Custom code", "Displays"],
              ["733215255", "Number, comma", "#,##0", "733,215,255"],
              ["733215255", "Currency (naira)", "\"₦\"#,##0", "₦733,215,255"],
              ["733215255", "Millions", "\"₦\"#,##0.0,,\"m\"", "₦733.2m"],
              ["0.127438", "Percentage, 1 dp", "0.0%", "12.7%"],
              ["45658", "Date, ISO", "yyyy-mm-dd", "2025-01-01"],
              ["45658", "Date, long", "d mmmm yyyy", "1 January 2025"]]),
          watch="Typing 12.7 into a %-formatted cell gives 12.7%, but typing 12.7 then formatting gives 1270%. Enter "
                "rates as decimals (0.127)."),

        T("1.6.10", "Alignment, Wrap Text and the Merge-Cell Trap",
          simple="Alignment positions text in the cell; Wrap Text splits long text onto several lines. Merging cells "
                 "looks neat but breaks sorting, filtering and copying.",
          grid=dict(label="Alignment tools", widths=[1.7, 3], size=12.5, rows=[
              ["Tool", "Use"], ["Left / Centre / Right", "Text left, numbers right, headers to match their data"],
              ["Wrap Text", "Long headers on two lines"], ["Indent", "Show sub-items under a heading"],
              ["Center Across Selection", "Centre a title over columns without merging"]]),
          steps=["Select the title cells across the columns", "Ctrl+1 › Alignment", "Horizontal: Center Across Selection"],
          steps_label="THE SAFE ALTERNATIVE TO MERGE", order=["grid", "steps"], split=0.58,
          watch="Never merge cells inside a data table. You cannot sort or filter through merged cells (Tidy Law 11)."),

        T("1.6.11", "Cell Styles, Themes and Conditional Formatting Basics",
          defn="Conditional formatting automatically changes a cell's format according to its value, using rules, "
               "colour scales, data bars or icon sets.",
          simple="Cell Styles apply saved formatting in one click; Themes change fonts and colours for the whole "
                 "workbook; conditional formatting makes problems stand out by themselves.",
          like="Conditional formatting is a traffic light: you do not read the number, you see red and stop.",
          ex=dict(t="tStores25", cols=["Store_Name", "Profit", "Margin"], n=5, sort="Margin", asc=True,
                  note="Rule: Margin less than 0 → red fill. These five stores would turn red."),
          steps=["Select the Margin column", "Home › Conditional Formatting › Highlight Cells Rules › Less Than",
                 "Type 0, choose Light Red Fill, OK"],
          steps_label="HIGHLIGHT LOSSES", order=["ex", "steps"], split=0.48),

        T("1.6.12", "Freeze Panes and Split",
          simple="Freeze Panes keeps the header row (and first column) visible while you scroll. Split divides the "
                 "window into panes that scroll separately.",
          grid=dict(label="Options (View › Freeze Panes)", widths=[1.7, 2.8], size=12.5, rows=[
              ["Option", "Keeps visible"], ["Freeze Top Row", "Row 1"], ["Freeze First Column", "Column A"],
              ["Freeze Panes", "Everything above and left of the active cell"],
              ["Split", "Four scrollable panes: compare distant rows"]]),
          steps=["To freeze row 1 and column A: click B2", "View › Freeze Panes › Freeze Panes (Alt, W, F, F)",
                 "Same command again: Unfreeze Panes"],
          steps_label="FREEZE HEADERS AND IDs", order=["grid", "steps"], split=0.55),

        T("1.6.13", "Printing and Page Setup",
          simple="Set the page up before printing or saving as PDF, and always check Print Preview first.",
          grid=dict(label="Page Layout settings", widths=[1.6, 3], size=12.5, rows=[
              ["Setting", "Use"], ["Orientation", "Landscape for wide tables"], ["Margins", "Narrow to fit more"],
              ["Print Area", "Print only the selected cells"], ["Print Titles", "Repeat the header row on every page"],
              ["Scale to Fit", "Fit all columns on one page wide"], ["Header / Footer", "Page numbers, file name, date"]]),
          steps=["Ctrl+P: preview", "Page Setup: landscape, fit all columns on one page", "Print Titles: rows to repeat $1:$1",
                 "Print, or save as PDF"],
          steps_label="PRINT A LONG TABLE", order=["grid", "steps"], split=0.58),

        T("1.6.14", "Find & Replace, Go To Special and Data Validation",
          simple="Find & Replace fixes many cells at once. Go To Special selects cells of one kind (blanks, formulas). "
                 "Data validation stops bad values being typed in the first place.",
          grid=dict(label="Three tools", widths=[1.6, 1.6, 2.8], size=12, rows=[
              ["Tool", "Open with", "Example"],
              ["Find & Replace", "Ctrl+H", "Replace 'LAGOS - IKEJA' with 'Lagos - Ikeja' (Match case)"],
              ["Go To Special", "F5 › Special", "Select all blanks in a column to inspect them"],
              ["Data validation", "Data › Data Validation", "Zone allows only a list: North-Central, South-East, South-South, South-West"]]),
          steps=["Select the Zone column", "Data › Data Validation › Allow: List", "Source: the four zone names, comma-separated",
                 "Error Alert tab: Stop, with a clear message"],
          steps_label="A DROP-DOWN LIST FOR ZONE", order=["grid", "steps"], split=0.58,
          watch="Replace All with an empty 'Find what' or a short word can change far more cells than you meant. Use "
                "Find All first to see the matches."),
    ],
    takeaways=["Enter dates as yyyy-mm-dd; Excel shows type by alignment (text left, numbers right).",
               "Six types: Number, Text, Date/Time, Logical, Error, Array. Text-numbers drop out of SUM.",
               "Paste Special › Values freezes results; AutoFill and Flash Fill save typing.",
               "Formatting changes the look, never the value; avoid merged cells, use Center Across Selection.",
               "Freeze Panes, Print Titles, Find & Replace, Go To Special and data validation keep big sheets usable."],
)

MODULES = [M15, M16]
