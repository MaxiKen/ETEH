from revlib import T

DIRTY = "SuperMart_Sales_2025_DIRTY.csv"

TOPICS = [
    T("2.5.1",
      what="**Dirty data** is any value you can't trust for the job: duplicates, numbers stored as text, the same "
           "category spelt several ways, missing or impossible values. **Cleaning** is not about making the sheet look "
           "nice; it means making every value clear, the right type and consistent, so totals and counts come out true.",
      like="Dirty data is like rice with small stones in it. It looks fine in the bag, but if you cook it without "
           "picking, someone bites a stone at dinner.",
      how=["Scan each column with a filter to see its values.",
           "Count each problem before fixing anything.",
           "Fix, then re-count to prove the problem is gone."],
      tip="Count each problem first; the count is your proof of cleaning.",
      trap="Starting to fix things without measuring, so you can't prove what changed.",
      who="SuperMart Retail",
      case=DIRTY + " is the raw 2025 POS export, 3,372 data rows. Before cleaning, the analyst lists what is wrong "
           "and how much.",
      steps=["Turn on filters (Ctrl+Shift+L) and scan each column.",
             "Count each problem with a formula or filter.",
             "List the problems and counts on a 00_README sheet."],
      table=[["Problem found", "Count"],
             ["Exact duplicate rows", "60"],
             ["Sales stored as text (₦239,240)", "234"],
             ["Spellings of 3 categories", "12"],
             ["Missing Quantity", "31"],
             ["Negative Quantity", "9"]],
      split=5.20,
      so="Each problem would distort a total or a count. Now there is a measured list to fix and to check against."),

    T("2.5.2",
      what="Cleaning works best as a **checklist**, not guesswork. Twelve problems come up again and again: "
           "duplicates, missing values, inconsistent case and spacing, inconsistent categories, wrong types, numbers "
           "as text, bad dates, out-of-range values, combined values, hidden characters, mixed units, and structure "
           "faults (blank rows, extra headers).",
      like="It is like a mechanic's service checklist: oil, brakes, tyres, lights… every car, every time. Nothing is "
           "missed because the list never forgets.",
      how=["Write one detection formula per problem.",
           "Run each formula over the whole column.",
           "Record the count next to each problem."],
      tip="Keep your detection formulas; rerun them after cleaning to prove zero.",
      trap="Checking only the problems you happen to notice.",
      who="SuperMart Retail",
      case="The analyst converts " + DIRTY + " to a Table named tDirty and writes one detection formula per "
           "problem.",
      steps=["Convert the data to a Table (Ctrl+T) named tDirty.",
             "Type each detection formula below.",
             "Note each count on the checklist."],
      table=[["Problem", "Detection formula", "Count"],
             ["Blank rows", "=COUNTBLANK(tDirty[Order_ID])", "15"],
             ["Sales as text", "=SUMPRODUCT(--ISTEXT(tDirty[Sales]))", "234"],
             ["Placeholder names", "=COUNTIF(tDirty[Customer_Name],\"N/A\")", "18"],
             ["Negative Quantity", "=COUNTIF(tDirty[Quantity],\"<0\")", "9"]],
      split=4.20,
      so="Four formulas give exact counts. After cleaning, the same formulas should all return 0."),

    T("2.5.3",
      what="**Data › Remove Duplicates** deletes repeated rows. It deletes permanently, so first decide **what "
           "makes a row unique** (the key, e.g. Row_ID), **count** the duplicates, and work on a **copy**. Ticking all "
           "columns removes only exact copies; ticking one column keeps the first row for each value of that column.",
      like="It is like a school removing pupils listed twice on the register. First check they really are the same "
           "child, not two children with the same name.",
      how=["Flag repeats: =COUNTIF($A$2:A2,A2)>1 (TRUE = seen before).",
           "Filter TRUE and review a few rows.",
           "On a copy: Data › Remove Duplicates; compare the count."],
      tip="Check the duplicates are real copies before removing them.",
      trap="Removing on a non-unique column (e.g. Order_ID) and deleting genuine order lines.",
      who="SuperMart Retail",
      case="Some POS lines in " + DIRTY + " were exported twice. Row_ID should be unique.",
      steps=["Add a helper column: =COUNTIF($A$2:A2,A2)>1.",
             "Filter TRUE: 60 rows; check they match their originals.",
             "Data › Remove Duplicates (all columns) on a copy."],
      table=[["Check", "Result"],
             ["Rows with a repeated Row_ID", "60"],
             ["Sales counted twice", "₦13.2m"],
             ["Rows before (blanks removed)", "3,372"],
             ["Rows after Remove Duplicates", "3,312"]],
      split=5.20,
      so="60 duplicate lines would have added ₦13.2m of sales that never happened. 3,312 matches the POS count."),

    T("2.5.4",
      what="**Ctrl+H** (Find and Replace) changes text across a range. It is powerful and dangerous: by default it "
           "matches **part of a cell** and **ignores case**. Tick **Match entire cell contents** so only whole-cell "
           "matches change, select the column first, and click **Find All** to preview before replacing.",
      like="It is like announcing 'Will Ade please come forward' in a hall. Ade comes, but so do Adewale, Adebayo "
           "and Folade. Be exact about who you're calling.",
      how=["Select only the column to change.",
           "Ctrl+H › Options › tick Match entire cell contents.",
           "Click Find All to preview, then Replace All."],
      tip="Always tick Match entire cell contents and preview with Find All.",
      trap="Replacing part-matches, which silently changes cells you never meant to touch.",
      who="SuperMart Retail",
      case="8 Category cells in " + DIRTY + " say 'Tech'. The analyst wants them to read 'Technology'.",
      steps=["Ctrl+H: Find 'Tech', Replace with 'Technology'.",
             "Click Find All without Match entire cell: 637 cells found.",
             "Tick Match entire cell contents: only 8 cells found."],
      table=[["Setting", "Cells changed", "Result"],
             ["Part-match (default)", "637", "'Technology' becomes 'Technologynology'"],
             ["Match entire cell contents", "8", "Only 'Tech' cells fixed"]],
      split=4.70,
      so="Without the tick, 629 correct cells would have been damaged. Find All showed the danger before it happened."),

    T("2.5.5",
      what="**Numbers stored as text** look like numbers but Excel treats them as words. Signs: they align left, "
           "show a green triangle, and **SUM and AVERAGE skip them**. Fix them with VALUE(), Text to Columns › Finish, "
           "or Paste Special › Multiply by 1, after removing symbols like ₦ and commas.",
      like="It is like a fake banknote in a pile of real ones. It looks the same, but the bank's counting machine "
           "rejects it, so your total comes up short.",
      how=["Spot them: =ISTEXT(cell) or left alignment.",
           "Select the column; Ctrl+H to remove ₦ and commas.",
           "Convert: Data › Text to Columns › Finish."],
      tip="After converting, check with =SUMPRODUCT(--ISTEXT(range)): it must be 0.",
      trap="Trusting a SUM over a column that contains text numbers.",
      who="SuperMart Retail",
      case="234 Sales values in " + DIRTY + " arrived as text, such as '₦239,240'. SUM ignores them.",
      steps=["Count text values: =SUMPRODUCT(--ISTEXT(tDirty[Sales])).",
             "Select the Sales column; replace ₦ and , with nothing.",
             "Text to Columns › Finish; count text values again."],
      table=[["Measure", "Before", "After"],
             ["Sales cells stored as text", "234", "0"],
             ["Value hidden from SUM", "₦57.1m", "₦0"],
             ["Example cell", "'₦239,240 (left-aligned)", "239,240 (number)"]],
      split=4.70,
      so="₦57.1m of sales was invisible to every formula. After conversion, every sale counts."),

    T("2.5.6",
      what="**TRIM** removes extra spaces: before, after and doubled between words. **CLEAN** removes hidden "
           "control characters such as line breaks. Use them together: =TRIM(CLEAN(A2)). Neither removes the "
           "**non-breaking space** CHAR(160), common in web and system exports; replace it first with "
           "SUBSTITUTE(A2, CHAR(160), \" \").",
      like="Extra spaces are like invisible dust on a lens. You can't see them, but every photo (every match and "
           "count) comes out slightly wrong.",
      how=["Detect: =LEN(A2)<>LEN(TRIM(A2)) shows TRUE for extra spaces.",
           "Fix: =TRIM(CLEAN(SUBSTITUTE(A2, CHAR(160), \" \"))).",
           "Paste the results as values over the original."],
      tip="Use the full TRIM(CLEAN(SUBSTITUTE())) combination for system exports.",
      trap="Using TRIM alone and wondering why some names still don't match.",
      who="SuperMart Retail",
      case="In " + DIRTY + " the Zone column has values like ' South-South' and 'South-East   '. Some customer names "
           "end in hidden characters.",
      steps=["Count Zone cells with extra spaces: 268.",
             "Add a helper column: =TRIM(L2); count the distinct zones.",
             "Fix names with =TRIM(CLEAN(SUBSTITUTE(G2, CHAR(160), \" \")))."],
      table=[["Check", "Before", "After"],
             ["Zone cells with extra spaces", "268", "0"],
             ["Distinct Zone values", "12", "4"],
             ["Names with hidden characters", "41", "0"]],
      split=4.90,
      so="12 'zones' were really 4. Invisible spaces split each zone into several groups in every report."),

    T("2.5.7",
      what="Before you can group data, each category must be spelt **one way**. Fix the case with PROPER, UPPER or "
           "LOWER and the spaces with TRIM. For typos and abbreviations, use a small **mapping table** (variant → "
           "standard) with XLOOKUP. Then add a **Data Validation** list so new entries stay clean.",
      like="It is like a church choir with members in different coloured robes. Give everyone the same robe and you "
           "can instantly see who belongs to which choir.",
      how=["List the distinct values (filter, or =UNIQUE()).",
           "Build a mapping table: Variant | Standard.",
           "Helper column: =XLOOKUP(LOWER(TRIM(O2)), Map[Variant], Map[Standard])."],
      tip="Keep the original column and add the standard one beside it.",
      trap="Fixing spellings one by one with Find & Replace and missing some.",
      who="SuperMart Retail",
      case="Category in " + DIRTY + " has 12 spellings for 3 real categories. A PivotTable shows 12 rows instead of 3.",
      steps=["List the 12 distinct spellings.",
             "Build the mapping table (8 variants → 3 standards).",
             "Add Category_Fixed with XLOOKUP; rebuild the PivotTable."],
      table=[["Variant (count)", "Standard"],
             ["Office Supply (30), Office-Supplies (24), office supplies (19)", "Office Supplies"],
             ["FURNITURE (16), furniture (10), Furnitre (8)", "Furniture"],
             ["Technolgy (10), Tech (8), technology (7)", "Technology"],
             ["PivotTable rows", "12 before, 3 after"]],
      split=4.20,
      so="One mapping table fixed every variant at once. The PivotTable now shows exactly 3 categories."),

    T("2.5.8",
      what="Dates go wrong in three ways: stored as **text**, entered in the **wrong order** (03/04: 3 April or "
           "March 4?), or **impossible** (31/02/2025). Test with =ISNUMBER(cell): a real date gives TRUE. Repair text "
           "dates with DATEVALUE, Text to Columns (choosing DMY), or DATE(RIGHT(), MID(), LEFT()).",
      like="A text date is like a wall calendar drawing: it shows a date, but you can't count the days between two "
           "pictures. A real date is a number Excel can count with.",
      how=["Find text dates: =ISNUMBER(B2) returns FALSE.",
           "Group them by pattern (dd/mm/yyyy, dd-Mon-yyyy, yyyy.mm.dd).",
           "Use the matching fix for each pattern; recheck ISNUMBER."],
      tip="After fixing, check the earliest and latest dates make sense.",
      trap="Letting Excel guess the order; 04/05 can silently become April 5.",
      who="SuperMart Retail",
      case="335 Order_Date values in " + DIRTY + " are text in three patterns. The analyst repairs each pattern.",
      steps=["Helper column: =ISNUMBER(C2); filter FALSE (335 rows).",
             "Apply the fix for each pattern (see the table).",
             "Check: no FALSE left; dates run 2025-01-01 to 2025-12-30."],
      table=[["Pattern (count)", "Example", "Fix"],
             ["dd/mm/yyyy (106)", "14/05/2025", "=DATE(RIGHT(C2,4), MID(C2,4,2), LEFT(C2,2))"],
             ["dd-Mon-yyyy (113)", "04-May-2025", "=DATEVALUE(C2)"],
             ["yyyy.mm.dd (116)", "2025.06.10", "=DATEVALUE(SUBSTITUTE(C2, \".\", \"-\"))"]],
      split=4.20,
      so="All 335 dates become real dates, so sorting, monthly grouping and date maths now work."),

    T("2.5.9",
      what="Every missing value needs a **clear decision**, never silence. The four honest options: **leave it "
           "blank** (unknown), **fill it** with a documented value, **exclude** the row from that analysis, or **flag** "
           "it for the source to correct. Never type 0 for 'unknown'. Record each decision and its count.",
      like="If a pupil's exam script is lost, the teacher doesn't write 0. She marks it 'missing', tells the head "
           "teacher, and decides with them what to do.",
      how=["Count blanks per column with =COUNTBLANK().",
           "Decide: leave blank, fill, exclude or flag.",
           "Write the decision and the count in 00_README."],
      tip="Write down every missing-value decision and how many rows it affects.",
      trap="Typing 0 or a guess into blanks; averages and totals become wrong.",
      who="SuperMart Retail",
      case=DIRTY + " has missing quantities, placeholder names and impossible negative quantities. The analyst "
           "decides how to handle each.",
      steps=["Count each problem with COUNTBLANK and COUNTIF.",
             "Choose an option for each problem.",
             "Add an Exception column and record the decisions."],
      table=[["Problem", "Count", "Decision"],
             ["Quantity blank", "31", "Flag; ask the store (don't type 0)"],
             ["Customer_Name = 'N/A'", "18", "Set to blank; report to IT"],
             ["Negative Quantity", "9", "Flag as possible returns"]],
      split=4.70,
      so="58 problem rows are flagged and explained, not hidden. Anyone can see what was decided and why."),

    T("2.5.10",
      what="**Blank rows** break data into pieces: Ctrl+Down stops, Ctrl+T selects part of the data, and filters "
           "miss rows. **Empty columns** do damage too. Remove blank rows with **Go To Special › Blanks** or a filter, "
           "and delete columns that hold nothing. Then press **Ctrl+End** to check where Excel thinks the data ends.",
      like="Blank rows are like empty carriages in the middle of a train. The engine (Excel) thinks the train ends "
           "at the gap and leaves the rest behind.",
      how=["Select a key column › Ctrl+G › Special › Blanks.",
           "Home › Delete › Delete Sheet Rows.",
           "Delete empty columns; press Ctrl+End to check the last cell."],
      tip="Select one key column for Go To Special, not the whole sheet.",
      trap="Running Go To Special on all columns; it also deletes rows with just one empty cell.",
      who="SuperMart Retail",
      case=DIRTY + " has 15 blank rows spread through the data, and an empty export column, Column1, in column M.",
      steps=["Select column B (Order_ID) › Go To Special › Blanks: 15 cells.",
             "Delete Sheet Rows; the data is now one block.",
             "Check =COUNTA(M:M) is 1 (header only); delete column M."],
      table=[["Defect", "How to detect", "Fix"],
             ["15 blank rows", "Go To Special › Blanks on Order_ID", "Delete Sheet Rows"],
             ["Empty Column1 (M)", "=COUNTA(M:M) = 1", "Delete the column"],
             ["Rows after fixing", "3,387 → 3,372", "One continuous block"]],
      split=4.20,
      so="The data is now one block of 3,372 rows. Filters, Tables and sorting all see every row."),

    T("2.5.11",
      what="**Data › Text to Columns** splits one column into several, by a **delimiter** (comma, space, dash) or by "
           "**fixed width**. It overwrites the columns to the right, so **insert empty columns first**. In the last "
           "step you can set each new column's format, for example Text to keep leading zeros.",
      like="It is like cutting a long loaf into slices along marks you made: one cut at every mark, and each slice "
           "lands on its own plate.",
      how=["Insert enough empty columns to the right.",
           "Data › Text to Columns › Delimited › choose the delimiter.",
           "Step 3: set each column's format (Text for codes) › Finish."],
      tip="Insert empty columns first, or the split overwrites your data.",
      trap="Splitting in place and losing the column next door.",
      who="SuperMart Retail",
      case="Product_ID in " + DIRTY + " holds three codes in one, e.g. TEC-AC-10003614. The manager wants to filter "
           "by the category code.",
      steps=["Insert 3 empty columns to the right of Product_ID.",
             "Text to Columns › Delimited › Other: type -.",
             "Set the number part to Text › Finish; name the columns."],
      table=[["Product_ID", "Cat_Code", "Sub_Code", "Item_No"],
             ["TEC-AC-10003614", "TEC", "AC", "10003614"],
             ["OFF-ST-10002485", "OFF", "ST", "10002485"],
             ["OFF-PA-10003022", "OFF", "PA", "10003022"]],
      split=4.70,
      so="Every one of the 3,372 IDs splits into exactly three parts, so the codes can now be filtered and counted."),

    T("2.5.12",
      what="Cleaning is finished when **checks prove it**, not when the sheet looks right. Compare the clean data "
           "with **control totals** from the source system: row count, total sales, total quantity, distinct orders. "
           "Rerun your detection formulas; they should all return 0. Put the result in one PASS/FAIL cell.",
      like="It is like a cashier closing the till: the cash in the drawer must match the day's sales report exactly "
           "before she goes home.",
      how=["Get control totals from the source system.",
           "Calculate the same totals on the clean data.",
           "=IF(AND(rows=control, sum=control), \"PASS\", \"FAIL\")."],
      tip="Ask for the source system's control totals before you start cleaning.",
      trap="Declaring the data clean because it 'looks right'.",
      who="SuperMart Retail",
      case="The POS system reports control totals for 2025. The analyst checks the cleaned table against every one.",
      steps=["Count rows: =ROWS(tSales).",
             "Sum Sales, Profit and Quantity; count distinct orders.",
             "Compare each with the POS control figure."],
      table=[["Check", "POS control", "Clean data"],
             ["Rows", "3,312", "3,312"],
             ["Total Sales", "₦733,215,255", "₦733,215,255"],
             ["Total Profit", "₦93,439,242", "₦93,439,242"],
             ["Total Quantity / Distinct orders", "12,476 / 1,687", "12,476 / 1,687"]],
      split=4.70,
      so="Every figure matches the source system. The data is proven clean and ready for analysis."),
]
