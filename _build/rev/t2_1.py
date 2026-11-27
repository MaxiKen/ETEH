from revlib import T

W2 = "Week2_Data_Files"
DIRTY = "SuperMart_Sales_2025_DIRTY.csv"
PIPE = "Lifecare_Admissions_2023_pipe.txt"
EU = "SuperMart_Store_Targets_EU_semicolon.csv"
RAWX = "SuperMart_Online_RAW_Export.xlsx"

TOPICS = [
    T("2.1.1",
      what="**External data** is any data that comes from outside the workbook you are working in: a till system "
           "export, a hospital system file, a database, a web page or a file a colleague emails you. Almost all real "
           "analysis starts with external data, so the key skill is bringing it in **the same way every time**, "
           "without retyping or damaging it.",
      like="External data is like foodstuff bought from the market. You don't grow it yourself; you bring it home, "
           "check it, wash it and store it properly before you cook with it.",
      how=["Find where the data lives (file, folder, database, web page).",
           "Note its format: CSV, TXT, XLSX, JSON, HTML or PDF.",
           "Choose the matching import route under Data › Get Data."],
      tip="Connect to the source; never copy and paste or retype data.",
      trap="Pasting a copy of the data, so next month you must redo everything by hand.",
      who="ETEH Practice Files",
      case="The " + W2 + " folder holds the Week 2 practice files. Each one is a different kind of external data, "
           "and each needs its own import route.",
      steps=["Open the " + W2 + " folder and list the files.",
             "Note each file's extension (.csv, .txt, .xlsx, .json, .html, .pdf).",
             "Match each one to its route in Data › Get Data."],
      table=[["Practice file", "Format", "Import route"],
             [DIRTY, "CSV", "From Text/CSV"],
             [PIPE, "TXT (pipe |)", "From Text/CSV"],
             [RAWX, "Workbook", "From Workbook"],
             ["SuperMart_Online_Sellers.json", "JSON", "From JSON"],
             ["SuperMart_Monthly_Sales_2025 (12 files)", "Folder", "From Folder"]],
      split=4.20,
      so="Every file arrives through Data › Get Data. The route depends only on the format, so learn the format first."),

    T("2.1.2",
      what="Four file formats appear in almost every job. **.xlsx** keeps numbers, dates, formulas and several "
           "sheets. **.csv** is plain text with values separated by commas, so it keeps no types or formatting. "
           "**.txt** is plain text with any separator (pipe, tab). **.pdf** stores a page layout, not data, so tables "
           "must be extracted before you can trust them.",
      like="An .xlsx file is like a packed lunch box with compartments. A .csv is the same food poured into one bag; "
           "you must sort it out again. A .pdf is a photo of the lunch: nice to look at, but you can't eat it.",
      how=["Check the file extension (turn on File name extensions in Explorer).",
           "Ask: does this format keep types, or is everything text?",
           "Plan the checks: leading zeros, dates and separators for text files."],
      tip="Treat every CSV and TXT value as text until you have checked it.",
      trap="Double-clicking a CSV to open it; Excel guesses the types and can strip leading zeros.",
      who="SuperMart and Lifecare",
      case="SuperMart and Lifecare send the analyst data in four formats. Before importing, she notes what each "
           "format keeps and what can go wrong.",
      steps=["Open each practice file's folder and check the extension.",
             "For CSV and TXT, open in Notepad to see the raw text.",
             "Write down the risk for each format before importing."],
      table=[["Format", "Keeps types?", "Example file", "Watch out for"],
             ["XLSX", "Yes", RAWX, "Several sheets"],
             ["CSV", "No, all text", DIRTY, "Leading zeros, dates"],
             ["TXT", "No, all text", PIPE, "The pipe | separator"],
             ["PDF", "Layout only", "Monthly_Sales_Bulletin_2025.pdf", "Extract tables first"]],
      split=4.20,
      so="Only XLSX keeps types. With CSV and TXT you decide the types; with PDF you must extract the table first."),

    T("2.1.3",
      what="To import from another workbook, use **Data › Get Data › From File › From Workbook**. The **Navigator** "
           "window lists every sheet and named Table in the file. Tick only the ones you need. Choose a named Table "
           "or a specific sheet by name, so the import still works if someone reorders the sheets.",
      like="It is like a library catalogue: instead of carrying home the whole shelf, you look up the exact book you "
           "need and borrow only that one.",
      how=["Data › Get Data › From File › From Workbook; pick the file.",
           "In the Navigator, click each sheet to preview it.",
           "Tick only what you need and click Transform Data."],
      tip="Import named Tables or named sheets, never 'the first sheet'.",
      trap="Ticking every sheet 'just in case' and loading thousands of rows nobody uses.",
      who="SuperMart Online",
      case="The online team sends " + RAWX + ". The analyst needs only the transactions and the customers.",
      steps=["From Workbook › choose " + RAWX + ".",
             "Click each of the 5 sheets in the Navigator and read the row count.",
             "Tick Transactions and Customers only, then Transform Data."],
      table=[["Sheet in the Navigator", "Data rows", "Needed?"],
             ["Transactions", "11,000", "Yes"],
             ["Customers", "2,500", "Yes"],
             ["Deliveries", "10,430", "No"],
             ["Products", "800", "No"],
             ["Sellers", "200", "No"]],
      split=4.90,
      so="Two of five sheets are needed. Ticking only those keeps the workbook small and the purpose clear."),

    T("2.1.4",
      what="**Data › From Text/CSV** opens a preview with three important settings: **File Origin** (the encoding, "
           "usually 65001: Unicode UTF-8), **Delimiter** (the separator, usually Comma) and **Data Type Detection**. "
           "Get these right at the start and you avoid broken characters, everything in one column, and lost "
           "leading zeros.",
      like="It is like tuning a radio: the right station (encoding) and the right volume (delimiter). Get either wrong "
           "and you hear noise, even though the broadcast is fine.",
      how=["Data › From Text/CSV; select the file.",
           "Check File Origin and Delimiter in the preview.",
           "For ID or phone columns, click Transform Data and set them to Text."],
      tip="Set phone numbers and ID codes to Text before the data loads.",
      trap="Letting Excel turn phone numbers like 07083197857 into 7083197857.",
      who="SuperMart Online",
      case="SuperMart_Online_customers.csv holds 2,500 customers with phone numbers like 07083197857. The "
           "analyst imports it and checks the phone column.",
      steps=["From Text/CSV › choose SuperMart_Online_customers.csv.",
             "Check: File Origin 65001 UTF-8, Delimiter Comma.",
             "Transform Data › set phone_number to Text › Close & Load."],
      table=[["Setting", "If wrong", "Correct choice"],
             ["File Origin", "Strange symbols, e.g. â‚¦", "65001: Unicode (UTF-8)"],
             ["Delimiter", "Everything in one column", "Comma"],
             ["phone_number type", "07083197857 → 7083197857", "Text"]],
      split=4.70,
      so="With phone_number set to Text, all 2,500 numbers keep their leading 0 and can actually be called."),

    T("2.1.5",
      what="A **.txt** file contains no information about its own structure, so you must tell Excel how the columns "
           "are separated. Common separators are the **pipe** (|) and the **tab**. Some older bank and hospital "
           "systems use **fixed width** instead: each column takes an exact number of characters.",
      like="A .txt file is like a long message with no punctuation. You have to know where one word ends and the next "
           "begins before it makes sense.",
      how=["Open the file in Notepad and look at the first line.",
           "Identify the separator: | or tab, or fixed positions.",
           "From Text/CSV › Delimiter › Custom (type |) or Tab."],
      tip="Always look at the raw file in Notepad first.",
      trap="Leaving the delimiter as Comma for a pipe file; everything lands in one column.",
      who="Lifecare Hospital",
      case="Lifecare's system exports 2023 admissions as " + PIPE + ", with fields separated by |.",
      steps=["Open the file in Notepad: Admission_ID|Patient_ID|Patient_Name|…",
             "Data › From Text/CSV › choose the file.",
             "Set Delimiter to Custom › type | › Load."],
      table=[["Delimiter setting", "Columns", "Rows"],
             ["Comma (wrong)", "1 long column", "10,918"],
             ["Custom | (correct)", "22 columns", "10,918"]],
      split=4.90,
      so="One setting turns a single column of text into 22 proper fields for 10,918 admissions."),

    T("2.1.6",
      what="A **delimiter** is the character that separates fields (comma, semicolon, tab, pipe). A **text "
           "qualifier**, usually a double quote (\"), wraps a value that itself contains the delimiter, so \"Ikeja, "
           "Lagos\" stays one field. A quote inside a value is written twice (\"\").",
      like="Quotes are like an envelope. Whatever is inside the envelope, even commas, is delivered as one letter.",
      how=["Open the raw file in Notepad and find a value with a comma.",
           "Check it is wrapped in double quotes.",
           "In the import dialog, keep the text qualifier as \" (the default)."],
      tip="Let Power Query handle quotes; never 'fix' them by hand in Notepad.",
      trap="Splitting on every comma, so one product name becomes three columns and every column shifts.",
      who="SuperMart Retail",
      case="Row 6683 in 2025-01_SuperMart_Sales.csv sells a 'Linden 10\" Round Wall Clock, Black'. The name holds "
           "both a comma and an inch mark.",
      steps=["Open the file in Notepad and find Row_ID 6683.",
             "See the name stored as \"Linden 10\"\" Round Wall Clock, Black\".",
             "Import with From Text/CSV and check Product_Name."],
      table=[["Element", "Role"],
             ["Comma ,", "Separates the 21 fields"],
             ["Quote \" around the value", "Keeps the inner comma inside one field"],
             ["Double quote \"\"", "Stands for one real \" (10 inches)"],
             ["Result", "Product_Name = Linden 10\" Round Wall Clock, Black"]],
      split=4.20,
      so="The qualifier keeps the row at exactly 21 fields. Without it, every column after Product_Name would shift."),

    T("2.1.7",
      what="**Import safely**: never paste over existing data, load to a dedicated sheet (or the Data Model), keep "
           "the connection so it can be refreshed, and **check the row count against the source** straight away. Note "
           "the file name, date and row count in a README sheet so later problems can be traced.",
      like="It is like counting your change before leaving the shop counter. Once you walk away, a missing ₦500 is "
           "hard to prove.",
      how=["Load to a new sheet named 01_RAW (or only to the Data Model).",
           "Compare the row count with the source system's total.",
           "Write source, date and row count in 00_README."],
      tip="Always write down the expected row count before you import.",
      trap="Assuming an import is complete without checking the count.",
      who="SuperMart Retail",
      case="The 2025 sales arrive as 12 monthly CSVs in SuperMart_Monthly_Sales_2025. The POS control total for the "
           "year is 3,312 lines. The analyst imports and checks.",
      steps=["Import the folder (Get Data › From Folder › Combine).",
             "Read the row count at the bottom of Queries & Connections.",
             "Record it in 00_README next to the control total."],
      table=[["Check", "Value"],
             ["Files imported", "12 (2025-01 to 2025-12)"],
             ["Rows: January / November", "155 / 459"],
             ["Rows loaded in total", "3,312"],
             ["POS control total", "3,312, so PASS"]],
      split=5.20,
      so="The counts match, so the import is complete. If a month went missing, the gap would show immediately."),

    T("2.1.8",
      what="**Encoding** decides how bytes become characters. **UTF-8** handles every language and symbols like ₦; "
           "reading it as **Windows-1252** turns ₦ into â‚¦. **Locale** decides how numbers and dates are read: in "
           "Europe '142 000 000,00' means 142 million, with a comma for the decimal point.",
      like="It is like reading a price written by a French trader. '1.500,00' is one thousand five hundred to him, but "
           "it looks like one and a half to you. Know who wrote it before you read it.",
      how=["Strange symbols? Re-import with File Origin 65001: UTF-8.",
           "European numbers? Set the locale (Using Locale) or use NUMBERVALUE.",
           "Check: =ISNUMBER(cell) must be TRUE after conversion."],
      tip="Always check that converted values are real numbers with ISNUMBER.",
      trap="Reading European numbers with Nigerian settings; they stay text or change value.",
      who="SuperMart Retail",
      case="A franchise consultant sends " + EU + ". It is semicolon-separated, and the target for Ikeja is written "
           "142 000 000,00.",
      steps=["Import with Delimiter set to Semicolon.",
             "Test =ISNUMBER() on the target column: FALSE.",
             "Convert with =NUMBERVALUE(I2, \",\", \" \") and test again."],
      table=[["Store", "As received (text)", "NUMBERVALUE result"],
             ["ST01 Lagos - Ikeja", "142 000 000,00", "142,000,000"],
             ["ST02 Lagos - Lekki", "63 000 000,00", "63,000,000"],
             ["ST04 Port Harcourt - GRA", "97 000 000,00", "97,000,000"]],
      split=4.70,
      so="NUMBERVALUE told Excel the comma is the decimal and the space groups thousands. Now the targets add up."),
]
