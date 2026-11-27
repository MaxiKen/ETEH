from revlib import T

DIRTY = "SuperMart_Sales_2025_DIRTY.csv"
LDIRTY = "Lifecare_Admissions_2023_DIRTY.csv"

TOPICS = [
    T("1.9.1",
      what="**Tidy data** is data in one simple, fixed shape: each **column** is one variable, each **row** is one "
           "record, and each **cell** holds one value. Every tool in this bootcamp (Tables, PivotTables, Power Query, "
           "SQL, Power BI) expects this shape. The next slides give the 12 'laws' that keep data tidy.",
      like="Tidy data is like a well-arranged provision shop: every item on its own labelled shelf. You find anything in "
           "seconds. An untidy shop has the same goods, but you waste an hour searching.",
      how=["Check columns: one variable each, one header row.",
           "Check rows: one record each; no totals or blank rows.",
           "Check cells: one value each, in a consistent format."],
      tip="Tidy the data first; analysis becomes simple afterwards.",
      trap="Building reports on untidy data and patching problems with complicated formulas.",
      who="SuperMart Retail",
      case="The clean SuperMart sales file and the dirty 2025 export hold the same kind of data. Compare their shapes "
           "against the three tidy rules.",
      steps=["Open SuperMart_Stores_Sales_2022_2025.csv and " + DIRTY + ".",
             "Check each against the three rules.",
             "List what breaks the rules in the dirty file."],
      table=[["Tidy rule", "Clean file", "Dirty file"],
             ["One variable per column", "Yes", "Discount mixes 0.2 and 20%"],
             ["One record per row", "Yes", "60 duplicates, 15 blank rows"],
             ["One value per cell", "Yes", "Store_Location 'Ikeja, Lagos'"]],
      split=4.70,
      so="Same data, different shape. The clean file is ready to analyse; the dirty one needs fixing first."),

    T("1.9.2",
      what="**Law 1: one record per row.** Each row must describe exactly one thing (one sale, one admission) and "
           "nothing else. Duplicate rows, subtotal rows and rows that combine two records all break counts and totals, "
           "because every row is counted as one record.",
      like="At a wedding, each guest signs the register once. If someone signs twice, or one line says 'the Okafor "
           "family (6)', the head count for food is wrong.",
      how=["Find the ID that should be unique (e.g. Admission_ID).",
           "Check for exact duplicates with Remove Duplicates (on a copy).",
           "Delete subtotal or 'Total' rows inside the data."],
      tip="Check that the ID column has no repeats before counting.",
      trap="Leaving duplicated rows in; counts and naira totals are both inflated.",
      who="Lifecare Hospital",
      case="The dirty 2023 admissions export, " + LDIRTY + ", has 11,026 rows. Some admissions appear twice. How much do "
           "the duplicates distort the totals?",
      steps=["Make a copy and run Data › Remove Duplicates with all columns ticked.",
             "Note how many rows were removed.",
             "Sum Billing_Amount for the removed rows."],
      table=[["Measure", "Value"],
             ["Rows in the export", "11,026"],
             ["Exact duplicate rows", "63"],
             ["Billing counted twice", "about ₦83.7m"]],
      split=5.20,
      so="63 repeated rows add ₦83.7m of billing that was never charged. One record per row protects every total."),

    T("1.9.3",
      what="**Law 2: one variable per column.** A column must hold one kind of information in one form. If a Discount "
           "column mixes 0.2 and '20%', or a Contact column mixes phone numbers and emails, Excel cannot calculate with "
           "it reliably.",
      like="It is like a drawer meant for spoons. Once forks and knives get mixed in, you can't grab a spoon without "
           "looking, and the count of spoons is wrong.",
      how=["Filter each column and scan the distinct values.",
           "Look for mixed forms (0.2 and 20%, numbers and text).",
           "Convert everything to one form, or split into separate columns."],
      tip="Store percentages as decimals (0.2) and format them as %.",
      trap="Averaging a column where some values are text; the text ones are silently skipped.",
      who="SuperMart Retail",
      case="In " + DIRTY + " the Discount column holds values like 0.2 and also '20%' and '0%' typed as text. The "
           "analyst wants the average discount.",
      steps=["Filter the Discount column and look at the values listed.",
             "Count how many are text with a % sign.",
             "Convert them to decimals (e.g. 20% → 0.2) before averaging."],
      table=[["Discount value", "Stored as", "Used by AVERAGE?"],
             ["0.2", "Number", "Yes"],
             ["'20%' (text)", "Text", "No, skipped"],
             ["Rows with text %", "173 rows", "All skipped"]],
      split=4.90,
      so="173 discounts are invisible to formulas until converted. One form per column makes every row count."),

    T("1.9.4",
      what="**Law 3: keep values atomic**, meaning one value per cell. 'Ikeja, Lagos' is two values (city and state). "
           "'chuKwuemEKa f. odili' holds a first name, an initial and a surname. Combined values are hard to filter, "
           "sort and count, so split them into separate columns.",
      like="It is like buying rice and beans already mixed in one bag. If you only want rice for tonight, you must "
           "pick it out grain by grain. Keep them in separate bags.",
      how=["Find columns that hold two or more facts.",
           "Split them: Data › Text to Columns, or Flash Fill (Ctrl+E).",
           "Give each new column a clear header."],
      tip="Split combined values as soon as you receive the data.",
      trap="Filtering 'Lagos' in a column of 'City, State' values; partial matches cause mistakes.",
      who="SuperMart Retail",
      case="In " + DIRTY + " Store_Location holds 'Ikeja, Lagos', 'Abuja, FCT' and so on. The manager wants sales by "
           "state.",
      steps=["Insert two empty columns to the right of Store_Location.",
             "Data › Text to Columns › Delimited › Comma; trim the spaces.",
             "Name the new columns City and State."],
      table=[["Store_Location", "City", "State"],
             ["Ikeja, Lagos", "Ikeja", "Lagos"],
             ["Lekki, Lagos", "Lekki", "Lagos"],
             ["Abuja, FCT", "Abuja", "FCT"],
             ["Ibadan, Oyo", "Ibadan", "Oyo"]],
      split=4.90,
      so="With State in its own column, both Lagos stores group together, and sales by state takes one PivotTable."),

    T("1.9.5",
      what="**Law 4: one header row.** Data must start with exactly one row of column names, in row 1. Report titles, "
           "logos and dates belong on a separate sheet. Two stacked header rows (for example 'Q1' above 'Sales' and "
           "'Profit') confuse Tables, PivotTables and Power Query.",
      like="A school register has one line of headings: name, then the dates. If someone adds a second heading line "
           "halfway down, the counting machine reads it as a pupil.",
      how=["Delete title rows above the headers.",
           "Combine stacked headers into one (e.g. Q1_Sales).",
           "Put titles and notes on a README sheet."],
      tip="Keep row 1 for headers only; put titles on a separate sheet.",
      trap="A title in row 1 and headers in row 3; Ctrl+T then treats the title as a header.",
      who="SuperMart Retail",
      case="A store manager sends a summary with 'SuperMart Q1 Report' in row 1, a blank row 2, 'Q1 2024' and 'Q1 2025' "
           "in row 3, and 'Sales' and 'Profit' under each in row 4.",
      steps=["Move the title to a README sheet and delete the blank row.",
             "Combine rows 3 and 4 into single headers.",
             "Convert to a Table (Ctrl+T) and check the headers."],
      table=[["Before (two header rows)", "After (one header row)"],
             ["Q1 2024 / Sales", "Sales_Q1_2024 = ₦93.2m"],
             ["Q1 2024 / Profit", "Profit_Q1_2024 = ₦11.4m"],
             ["Q1 2025 / Sales", "Sales_Q1_2025 = ₦123.1m"],
             ["Q1 2025 / Profit", "Profit_Q1_2025 = ₦23.5m"]],
      split=4.70,
      so="One clear header row lets Excel understand every column, so Tables and PivotTables work at once."),

    T("1.9.6",
      what="**Law 5: no spacer rows or blank columns** inside data. People add blank rows to make sheets 'look nicer', "
           "but Excel treats a blank row as the end of the data. Ctrl+Down stops there, Ctrl+T selects only part of the "
           "data, and sorts can leave rows behind.",
      like="A blank row is like a broken bridge on the expressway. Drivers (Excel tools) think the road ends there and "
           "never reach the towns beyond it.",
      how=["Press Ctrl+Down from A1 and check you reach the last row.",
           "Ctrl+G › Special › Blanks to find empty cells.",
           "Delete blank rows; use formatting (not empty rows) for spacing."],
      tip="Use borders or row height for spacing, never empty rows.",
      trap="Ctrl+T or a PivotTable silently using only the rows above the first blank row.",
      who="SuperMart Retail",
      case="In " + DIRTY + " the first blank row is row 212. An analyst presses Ctrl+T in A1 and builds a report. Only a "
           "small part of the data is included.",
      steps=["Click A1 and press Ctrl+Down: it stops at row 211.",
             "Ctrl+G › Special › Blanks shows 15 blank rows.",
             "Delete them, then press Ctrl+Down again: it reaches the last row."],
      table=[["Before deleting blanks", "Value"],
             ["Ctrl+Down from A1 stops at", "Row 211"],
             ["Rows Ctrl+T would include", "210 of 3,372 records"],
             ["Blank rows in the file", "15"]],
      split=5.20,
      so="One blank row hid 94% of the data from the report. Remove spacer rows before you analyse."),

    T("1.9.7",
      what="**Law 6: use ISO dates (YYYY-MM-DD)**, such as 2025-03-18. Dates must be real Excel dates, not text, and "
           "shown in one format. ISO dates are never ambiguous: 04/05/2025 can mean 4 May or 5 April depending on the "
           "computer's settings, but 2025-05-04 always means 4 May.",
      like="It is like writing your date of birth on a form. '04/05' could confuse an American officer, but '2025-05-04' "
           "is clear to everyone, everywhere.",
      how=["Check alignment: real dates align right; text dates align left.",
           "Convert text dates (Data › Text to Columns › Date, or Power Query).",
           "Ctrl+1 › Custom › yyyy-mm-dd for display."],
      tip="Format every date column as yyyy-mm-dd.",
      trap="Dates stored as text; they sort as words and can't be grouped by month.",
      who="SuperMart Retail",
      case="In " + DIRTY + " 335 Order_Date values are in mixed formats. Sorting by date puts November before February. "
           "The analyst standardises them.",
      steps=["Filter Order_Date and look for left-aligned (text) values.",
             "Convert each pattern to a real date.",
             "Format the whole column as yyyy-mm-dd and sort again."],
      table=[["As found", "Problem", "ISO date"],
             ["14/05/2025", "Text, day first", "2025-05-14"],
             ["04-May-2025", "Text with month name", "2025-05-04"],
             ["2025.06.10", "Dots, not a date", "2025-06-10"],
             ["2025-03-18", "Already correct", "2025-03-18"]],
      split=4.90,
      so="After conversion all 3,312 dates sort, filter and group correctly. Mixed text dates do none of these."),

    T("1.9.8",
      what="**Law 7: blank and zero mean different things.** A **blank** means unknown or not applicable. A **zero** "
           "means the value was measured and is nothing. AVERAGE ignores blanks but includes zeros, so typing 0 for "
           "'no data yet' drags averages down.",
      like="If a student missed an exam because of illness, recording 0% says she sat it and scored nothing. Leaving it "
           "blank says 'no score yet'. Her class average depends on which you choose.",
      how=["Decide what 'missing' means for each column.",
           "Leave truly unknown values blank (or use a clear code in a note).",
           "Use 0 only when the measured amount really was zero."],
      tip="Write in the data dictionary what a blank means for each column.",
      trap="Typing 0 for 'no data yet'; averages and minimums become wrong.",
      who="SuperMart Retail",
      case="Suppose SuperMart adds a 13th store, Kano, which only opens in 2026. In the 2025 store table, should Kano's "
           "sales be 0 or blank?",
      steps=["List the 12 stores' 2025 sales, then add Kano.",
             "Try Kano = 0 and calculate =AVERAGE() of all 13 cells.",
             "Try Kano = blank and calculate the average again."],
      table=[["Kano's 2025 sales cell", "AVERAGE of stores", "Correct?"],
             ["0", "₦56.4m (÷13)", "No: Kano didn't trade"],
             ["Blank", "₦61.1m (÷12)", "Yes"]],
      split=4.90,
      so="A zero for a store that didn't trade lowers the average by ₦4.7m. Blank means 'not applicable'."),

    T("1.9.9",
      what="**Law 8: keep raw data separate from your work and reports.** Keep the original file exactly as received "
           "(the **raw** layer), do cleaning in a copy or in Power Query (the **clean** layer), and build summaries and "
           "charts on separate sheets (the **report** layer). If something goes wrong, you can always start again from "
           "raw.",
      like="It is like keeping a seed yam aside before you cook the rest. Whatever happens in the kitchen, you can "
           "still plant next season.",
      how=["Save the original in 01_RAW and never edit it.",
           "Clean in 02_CLEAN (a copy, or Power Query from Week 2).",
           "Build reports in 03_REPORT from the clean data only."],
      tip="Make the raw file read-only (right-click › Properties › Read-only).",
      trap="Cleaning the only copy; a wrong Replace All can never be undone after saving.",
      who="Lifecare Hospital",
      case="Lifecare's analyst receives " + LDIRTY + " and must produce a ward report. She sets up three layers so "
           "she can always check her work.",
      steps=["Save the export in 01_RAW and set it to read-only.",
             "Make a cleaned copy in 02_CLEAN: fix wards, genders, duplicates.",
             "Build the ward summary in 03_REPORT from the clean copy."],
      table=[["Layer", "Contains", "Edit?"],
             ["01_RAW", "Export exactly as received (11,026 rows)", "Never"],
             ["02_CLEAN", "Fixed spellings, no duplicates", "Yes, step by step"],
             ["03_REPORT", "Ward summary, charts", "Yes"]],
      split=4.70,
      so="If a cleaning step goes wrong, the analyst goes back to the untouched raw file. Nothing is lost."),

    T("1.9.10",
      what="**Law 9: control your vocabulary.** Each category must be written exactly one way. To Excel, 'Female', "
           "'female' and 'F' are three different values. Counts, filters and PivotTables split one group into many. Use "
           "**Data Validation** lists so only the agreed spellings can be entered.",
      like="If people at a bus stop write 'Oshodi', 'Osodi' and 'OSH' on their tickets, the conductor's count for each "
           "stop is wrong. One agreed spelling fixes it.",
      how=["Filter each category column and list the variants.",
           "Replace each variant with the agreed value (Ctrl+H, Match case).",
           "Add Data Validation › List so new entries stay consistent."],
      tip="Agree one spelling for every category and enforce it with a list.",
      trap="COUNTIF(range,\"Female\") ignores 'F', so counts look right but are short.",
      who="Lifecare Hospital",
      case="In " + LDIRTY + " Gender is written six ways and Ward 16 ways (for only 4 wards). The analyst counts "
           "female patients.",
      steps=["Filter Gender: Female, female, F, Male, male, M.",
             "Type =COUNTIF(C:C,\"Female\") and note the result.",
             "Replace F and female with Female, then count again."],
      table=[["Measure", "Value"],
             ["Gender spellings", "6 (for 2 values)"],
             ["Ward spellings", "16 (for 4 wards)"],
             ["COUNTIF \"Female\" before fixing", "5,420"],
             ["True female count", "5,541"]],
      split=5.20,
      so="121 women coded 'F' were missing from the count. One controlled spelling per category fixes this for good."),

    T("1.9.11",
      what="**Law 10: store calculations, not typed results.** If you type a result (like 103%), it never changes, and "
           "nobody can see how it was worked out. A **formula** (=Sales/Target) updates when the data changes and shows "
           "its logic in the Formula Bar.",
      like="A typed result is like a photo of a clock: it shows the right time once, then it is wrong for the rest of "
           "the day. A formula is the working clock.",
      how=["Never type a number you calculated elsewhere.",
           "Write the formula in the cell, pointing to the inputs.",
           "Check: click the cell, and the Formula Bar should show =…"],
      tip="If a number came from a calculation, the cell should contain that calculation.",
      trap="Pasting results from a calculator; when inputs change, the report is silently wrong.",
      who="SuperMart Retail",
      case="Ikeja's 2025 sales were ₦146.4m against a ₦142m target. One analyst types 103.1%; another writes "
           "=B2/C2. Then Head Office revises Ikeja's target to ₦150m (a what-if).",
      steps=["Type 103.1% in one cell and =B2/C2 in another.",
             "Change the target in C2 from 142,000,000 to 150,000,000.",
             "Compare the two cells."],
      table=[["Approach", "Before change", "After target = ₦150m"],
             ["Typed 103.1%", "103.1%", "103.1% (wrong)"],
             ["Formula =B2/C2", "103.1%", "97.6% (correct)"]],
      split=4.90,
      so="Only the formula noticed that Ikeja is now below target. Typed results go stale without warning."),

    T("1.9.12",
      what="**Law 11: no merged cells** in data, because merging breaks sorting, filtering and PivotTables (use Center "
           "Across Selection for titles). **Law 12: document everything**: what each file and column means, its unit, "
           "where it came from, and what you changed. A **data dictionary** does this.",
      like="Documentation is like the label on a medicine bottle: name, dose, expiry date. Without it, even the right "
           "medicine is dangerous to use.",
      how=["Unmerge all cells in data sheets; fill values into every row.",
           "Create a README or dictionary sheet: column, type, meaning, example.",
           "Log every cleaning step with the date."],
      tip="Keep a data dictionary with every dataset you share.",
      trap="Handing over a file that only you understand.",
      who="Bootcamp Files",
      case="The pack's 00_Data_Dictionary_and_Catalogue.xlsx documents every bootcamp dataset. Use it as the model for "
           "your own work.",
      steps=["Open the Catalogue sheet: one row per dataset.",
             "Open the Dictionary sheet: one row per column (77 columns described).",
             "Find the entry for Row_ID in the SuperMart sales file."],
      table=[["Dictionary column", "Example entry"],
             ["Dataset", "SuperMart_Stores_Sales_2022_2025"],
             ["Column", "Row_ID"],
             ["Type", "Integer"],
             ["Description", "Line number, unique per sales line"]],
      split=4.90,
      so="Anyone can now understand each column without asking you. Good documentation makes your data trustworthy."),
]
