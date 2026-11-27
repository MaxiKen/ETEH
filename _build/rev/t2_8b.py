from revlib import T

W, ON, LC = "SuperMart Retail", "SuperMart Online", "Lifecare Hospital"
TRY = "ETEH_Week2_Section_Examples.xlsx  ›  sheet '2.8B Functions 34-65'"


def F(tag, what, like, how, tip, trap, case, steps, table, so, who=W, split=4.70):
    return T(tag, what=what, like=like, how=how, tip=tip, trap=trap, who=who, case=case, steps=steps,
             table=table, so=so, split=split, tryit=TRY)


TOPICS = [
    # ---------- text parsing ----------
    F("2.8.34",
      "**LEFT** takes characters from the **start** of a text. Syntax: **=LEFT(text, num_chars)**. Use it when the "
      "piece you want always sits at the beginning, such as the 3-letter category code in a Product_ID.",
      "LEFT is like reading only the first word on a signboard: 'FUR…' tells you it's the furniture section before "
      "you read the rest.",
      ["Check the piece is always the same length at the start.", "Type =LEFT(cell, number of characters).", "Fill down and spot-check a few rows."],
      "Check the code really has a fixed length before using LEFT.",
      "Using LEFT on codes whose first part changes length (use TEXTBEFORE).",
      "The analyst needs the category code from each Product_ID in the order sample on the 2.8B sheet.",
      ["In the Product_ID table, add a column LEFT_CatCode.", "Type =LEFT([@Product_ID], 3).", "Fill down."],
      [["Product_ID", "=LEFT(…, 3)"], ["FUR-FU-10003878", "FUR"], ["OFF-BI-10002012", "OFF"], ["OFF-ST-10002743", "OFF"]],
      "Every Product_ID starts with a 3-letter code, so LEFT(…, 3) reliably gives the category."),

    F("2.8.35",
      "**RIGHT** takes characters from the **end** of a text. Syntax: **=RIGHT(text, num_chars)**. The result is "
      "always **text**, even when it looks like a number, so wrap it in VALUE if you need to calculate with it.",
      "RIGHT is like reading the last digits of a phone number to identify a caller: you only need the end.",
      ["Check how many characters the ending has.", "Type =RIGHT(cell, number).", "Wrap in VALUE(…) if you need a number."],
      "Remember RIGHT returns text; use VALUE to make it a number.",
      "Trying to SUM RIGHT results and getting 0.",
      "SuperMart Order_IDs look like SM-2025-107503. The analyst needs the 6-digit order number.",
      ["Add a column RIGHT_OrderNo.", "Type =RIGHT([@Order_ID], 6).", "Check: the result is left-aligned (text)."],
      [["Order_ID", "=RIGHT(…, 6)"], ["SM-2025-107503", "107503"], ["SM-2025-144463", "144463"], ["SM-2025-151750", "151750"]],
      "RIGHT pulled the order number. It is text, which is fine for an ID you won't add up."),

    F("2.8.36",
      "**MID** takes characters from the **middle** of a text. Syntax: **=MID(text, start_num, num_chars)**. You "
      "say where to start and how many characters to take. Use it for codes where the piece sits at a fixed position.",
      "MID is like a cashier reading the middle digits of a card number: start at the 4th digit and read 4.",
      ["Count the start position (the first character is 1).", "Count how many characters to take.", "Type =MID(cell, start, length)."],
      "Count positions carefully; spaces and dashes count too.",
      "Starting one position too early or late, so the result is shifted.",
      "Order_IDs (SM-2025-107503) hold the year in the middle. The analyst extracts it.",
      ["Count: 'S'=1, 'M'=2, '-'=3, so the year starts at 4.", "Type =MID([@Order_ID], 4, 4).", "Fill down; every 2025 order shows 2025."],
      [["Argument", "Value"], ["text", "SM-2025-107503"], ["start_num", "4"], ["num_chars", "4"], ["Result", "2025"]],
      "MID extracted the year. Like RIGHT, the result is text: \"2025\", not the number 2025."),

    F("2.8.37",
      "**LEN** counts the characters in a text, **including spaces**. Syntax: **=LEN(text)**. Its best use is "
      "checking: if LEN(A2) is bigger than LEN(TRIM(A2)), the cell has extra hidden spaces.",
      "LEN is like counting the letters on a name tag, including the gaps. Extra gaps show up in the count even though "
      "you can't see them.",
      ["Type =LEN(cell) to count characters.", "Compare with =LEN(TRIM(cell)).", "A difference means hidden spaces."],
      "Use LEN versus LEN(TRIM) to find invisible spaces.",
      "Thinking two values are identical because they look identical.",
      "Zone looks fine in the dirty 2025 file, but there are 12 different zone values. The analyst checks for spaces.",
      ["Check one ID: =LEN(\"FUR-FU-10003878\") = 15.", "In tDirty, compare LEN([@Zone]) with LEN(TRIM([@Zone])).", "Count the rows where they differ."],
      [["Check", "Result"], ["LEN of a Product_ID", "15"], ["Zone values with extra spaces", "268"], ["Distinct zones: before / after TRIM", "12 / 4"]],
      "268 Zone values had hidden spaces. That is why 4 real zones appeared as 12."),

    F("2.8.38",
      "**FIND** returns the **position** of one text inside another. Syntax: **=FIND(find_text, within_text, "
      "[start_num])**. It is **case-sensitive** and returns #VALUE! if nothing is found. Use start_num to find the "
      "second occurrence.",
      "FIND is like searching a register for 'Ade' with capital A exactly: 'ade' in lower case would not count.",
      ["Type =FIND(\"-\", cell) for the first dash.", "For the second one, start after the first: FIND(\"-\", cell, first+1).", "Use the position in LEFT, MID or RIGHT."],
      "Use FIND when the exact character or case matters.",
      "Using FIND for words typed in mixed case (use SEARCH).",
      "Product_IDs have two dashes: FUR-FU-10003878. The analyst needs the position of the second dash.",
      ["First dash: =FIND(\"-\", [@Product_ID]) = 4.", "Second: =FIND(\"-\", [@Product_ID], FIND(\"-\", [@Product_ID]) + 1).", "Read the result."],
      [["Formula", "Result"], ["=FIND(\"-\", \"FUR-FU-10003878\")", "4"], ["Second dash (start at 5)", "7"], ["=FIND(\"chair\", \"…Chair Mats\")", "#VALUE! (case)"]],
      "The second dash is at position 7. FIND missed 'Chair' because it looked for lower-case 'chair'."),

    F("2.8.39",
      "**SEARCH** also returns the position of text, but **ignores case** and allows wildcards (* and ?). Syntax: "
      "**=SEARCH(find_text, within_text)**. Wrap it as **ISNUMBER(SEARCH(…))** for a 'contains' test.",
      "SEARCH is like a friendly librarian: ask for 'chair' and you also get books titled 'CHAIR' or 'Chair'.",
      ["Type =SEARCH(\"word\", cell).", "Wrap it: =ISNUMBER(SEARCH(\"word\", cell)) for TRUE/FALSE.", "Add IFERROR if you want text instead of #VALUE!."],
      "Use ISNUMBER(SEARCH(…)) to test whether text contains a word.",
      "Expecting SEARCH to tell you how many times a word appears; it gives only the first position.",
      "The furniture buyer wants every 2025 line whose product name mentions 'chair'.",
      ["On one row: =SEARCH(\"chair\", [@Product_Name]).", "For the table, count ISNUMBER(SEARCH(\"chair\", tSales[Product_Name])).", "Compare with the same count using FIND."],
      [["Test", "Result"], ["SEARCH in '…Task Swivel Chair'", "40"], ["Lines containing 'chair' (SEARCH)", "220"], ["Same count with FIND", "7"]],
      "SEARCH found 220 chair lines; FIND found only 7 because most say 'Chair' with a capital C."),

    # ---------- cleaning ----------
    F("2.8.40",
      "**SUBSTITUTE** replaces **specific text** with other text, wherever it appears. Syntax: "
      "**=SUBSTITUTE(text, old_text, new_text, [instance])**. It is ideal for removing symbols (₦, commas) or "
      "invisible characters like CHAR(160).",
      "SUBSTITUTE is like 'Find and Replace' inside a formula: every '₦' becomes nothing, every odd space becomes a "
      "normal one.",
      ["Identify the unwanted text or character.", "Type =SUBSTITUTE(cell, old, new).", "Nest SUBSTITUTEs to remove several things."],
      "Replace CHAR(160) with a normal space before you TRIM.",
      "Expecting TRIM or CLEAN to remove CHAR(160); neither does.",
      "Customer names in the dirty 2025 file contain non-breaking spaces, CHAR(160), copied from a website.",
      ["Type =SUBSTITUTE([@Customer_Name], CHAR(160), \" \").", "Then apply TRIM.", "Count affected names to see the size of the problem."],
      [["Check", "Result"], ["Names containing CHAR(160)", "41"], ["…of which also have line feeds", "22"], ["After SUBSTITUTE + TRIM", "Normal spacing"]],
      "41 names had a hidden character that TRIM alone would leave behind. SUBSTITUTE removes it first."),

    F("2.8.41",
      "**REPLACE** swaps characters at a **known position**. Syntax: **=REPLACE(old_text, start_num, num_chars, "
      "new_text)**. Unlike SUBSTITUTE, it works by position, not content. It is ideal for masking parts of phone "
      "numbers or IDs.",
      "REPLACE is like covering the middle of a card number with tape: you know exactly which digits to hide.",
      ["Count the start position.", "Count how many characters to replace.", "Type =REPLACE(cell, start, count, \"****\")."],
      "Mask personal data before sharing a file.",
      "Using REPLACE when the position changes from row to row.",
      "The online team must share a customer list without showing full phone numbers.",
      ["Phone 07083197857 has 11 digits.", "Type =REPLACE([@phone_number], 5, 4, \"****\").", "Fill down."],
      [["phone_number", "=REPLACE(…, 5, 4, \"****\")"], ["07083197857", "0708****857"], ["09113561597", "0911****597"], ["09055667651", "0905****651"]],
      "Digits 5–8 are hidden. Enough remains to recognise a number, but not to call it.", who=ON),

    F("2.8.42",
      "**TRIM** removes spaces at the **start and end** and turns runs of spaces **inside** into a single space. "
      "Syntax: **=TRIM(text)**. It is the most useful cleaning function, because extra spaces break lookups and "
      "counts silently.",
      "TRIM is like straightening a queue: no one standing far in front or behind, and just one step between people.",
      ["Type =TRIM(cell).", "Fill down.", "Check with LEN that the length dropped."],
      "TRIM every text key before using it in a lookup.",
      "Thinking TRIM removes CHAR(160) or line feeds (use SUBSTITUTE and CLEAN).",
      "A customer name was keyed as '  FUNMILAYO   BUHARI' with extra spaces.",
      ["Type =TRIM([@Name_as_keyed]).", "Compare LEN before and after.", "Apply the same fix to Zone in the dirty file."],
      [["Before", "After TRIM"], ["'  FUNMILAYO   BUHARI'", "'FUNMILAYO BUHARI'"], ["'Tochukwu  Kanu'", "'Tochukwu Kanu'"], ["Zone: 12 distinct", "4 distinct"]],
      "TRIM fixed the spacing. On Zone it turned 12 'different' values back into 4 real zones.", who=ON),

    F("2.8.43",
      "**CLEAN** removes **non-printing characters** (codes 0–31), such as line feeds and tabs. Syntax: "
      "**=CLEAN(text)**. They come from exports and copied text; you can't see them, but they stop matches working.",
      "CLEAN is like shaking sand out of a shoe: you couldn't see the grains, but you could feel them.",
      ["Type =CLEAN(cell).", "Combine: =TRIM(CLEAN(SUBSTITUTE(cell, CHAR(160), \" \"))).", "Check with LEN."],
      "Use the full chain: SUBSTITUTE → CLEAN → TRIM.",
      "Expecting CLEAN to remove CHAR(160); it is above code 31.",
      "One customer name was keyed as 'Ade Eze' followed by a hidden line feed.",
      ["Type =CLEAN([@SUBSTITUTE_nbsp]).", "Check LEN before and after.", "Then apply TRIM."],
      [["Step", "Result"], ["Keyed value", "'Ade Eze' + line feed"], ["LEN before / after CLEAN", "8 / 7"], ["Dirty file names with line feeds", "22"]],
      "CLEAN removed the invisible line feed. 22 names in the dirty file need the same fix.", who=ON),

    F("2.8.44",
      "**PROPER** capitalises the **first letter of each word** and lowercases the rest. Syntax: **=PROPER(text)**. "
      "It makes names and labels look consistent. Check unusual names afterwards.",
      "PROPER is like a teacher rewriting the class list neatly, with each name starting with a capital letter.",
      ["Clean first (SUBSTITUTE, CLEAN, TRIM).", "Type =PROPER(cleaned cell).", "Spot-check names with unusual capitals."],
      "Apply PROPER last, after cleaning.",
      "Trusting PROPER on names like McDonald (becomes Mcdonald).",
      "Customer names were keyed in every style: capitals, lower case, mixed.",
      ["Type =PROPER([@TRIM]).", "Fill down.", "Compare the before and after columns."],
      [["Keyed", "=PROPER(…)"], ["FUNMILAYO BUHARI", "Funmilayo Buhari"], ["amara attah", "Amara Attah"], ["emmanuel LAWAL", "Emmanuel Lawal"]],
      "Every name now follows the same style, ready for letters, labels and reports.", who=ON),

    F("2.8.45",
      "**UPPER** turns all letters into **capitals**. Syntax: **=UPPER(text)**. It is the safest way to make codes "
      "and category keys match, because every casing variant becomes one. It does **not** fix spelling mistakes.",
      "UPPER is like asking everyone to write their name in block letters: 'furniture' and 'Furniture' now look the "
      "same.",
      ["Type =UPPER(TRIM(cell)).", "Count distinct values before and after.", "Map the remaining misspellings with a lookup table."],
      "Use UPPER(TRIM()) for matching keys.",
      "Expecting UPPER to fix 'Furnitre' or 'Tech'; it only fixes case.",
      "Category in the dirty 2025 file has 12 variants for 3 real categories.",
      ["Type =UPPER(TRIM([@Category])).", "Count the distinct results.", "List what remains."],
      [["Stage", "Distinct values"], ["Raw Category", "12"], ["After UPPER(TRIM())", "8"], ["Left over: misspellings", "FURNITRE, TECH, TECHNOLGY…"]],
      "UPPER removed the case variants (12 → 8). The misspellings still need a mapping table."),

    F("2.8.46",
      "**LOWER** turns all letters into **lower case**. Syntax: **=LOWER(text)**. It is most useful for email "
      "addresses and for building matching keys. Combine it with TRIM.",
      "LOWER is like writing every email address the same way on a guest list, so no guest is listed twice.",
      ["Type =LOWER(TRIM(cell)).", "Fill down.", "Use the result to match or remove duplicates."],
      "Store email addresses in lower case.",
      "Treating two emails as different people because of capital letters.",
      "Some online customers typed their emails in capitals.",
      ["Type =LOWER([@email]).", "Fill down.", "Use it to find duplicates."],
      [["email (as typed)", "=LOWER(…)"], ["FUNMILAYO.BUHARI31@GMAIL.COM", "funmilayo.buhari31@gmail.com"], ["ADE.EZE105@GMAIL.COM", "ade.eze105@gmail.com"]],
      "Every email now has one standard form, so duplicates can be found reliably.", who=ON, split=4.20),

    F("2.8.47",
      "**VALUE** turns text that looks like a number into a **real number**. Syntax: **=VALUE(text)**. Remove "
      "currency signs and commas first with SUBSTITUTE; otherwise VALUE may return #VALUE!.",
      "VALUE is like changing a cheque written in words into cash: until you convert it, you can't spend it.",
      ["Remove symbols: SUBSTITUTE(SUBSTITUTE(cell, \"₦\", \"\"), \",\", \"\").", "Wrap in VALUE(…).", "Check the result is right-aligned and adds up."],
      "Test with COUNT: converted values are counted as numbers.",
      "Leaving ₦ or commas in the text, so VALUE fails.",
      "The dirty 2025 file has Sales values stored as text, like '₦239,240'.",
      ["Type =VALUE(SUBSTITUTE(SUBSTITUTE([@Sales], \"₦\", \"\"), \",\", \"\")).", "Fill down.", "Total the converted values."],
      [["Check", "Result"], ["'₦239,240' → number", "239,240"], ["Text Sales values in the file", "234"], ["Their total once converted", "₦57,079,213"]],
      "234 text values worth ₦57.1m were invisible to SUM until VALUE converted them."),

    F("2.8.48",
      "**TEXTBEFORE** returns the text **before** a separator. Syntax: **=TEXTBEFORE(text, delimiter)**. It is "
      "available in Excel 365 and 2021, and replaces the harder LEFT + FIND combination.",
      "TEXTBEFORE is like reading an address up to the first comma: 'Ikeja, Lagos' gives you 'Ikeja'.",
      ["Identify the separator, e.g. \", \".", "Type =TEXTBEFORE(cell, \", \").", "Add [if_not_found] for rows with no separator."],
      "Include the space in the separator (\", \"), so no space is left over.",
      "Using it in older Excel versions; it returns #NAME?.",
      "The dirty file has Store_Location as 'City, State'. The analyst needs the city.",
      ["Type =TEXTBEFORE([@Store_Location], \", \").", "Fill down.", "Count distinct cities."],
      [["Store_Location", "=TEXTBEFORE(…, \", \")"], ["Ikeja, Lagos", "Ikeja"], ["Abuja, FCT", "Abuja"], ["Port Harcourt, Rivers", "Port Harcourt"]],
      "The city comes out cleanly, even when it has two words like Port Harcourt."),

    F("2.8.49",
      "**TEXTAFTER** returns the text **after** a separator. Syntax: **=TEXTAFTER(text, delimiter, "
      "[instance])**. Use instance −1 to take what comes after the **last** separator, e.g. a surname.",
      "TEXTAFTER is like reading the end of an address after the comma: 'Ikeja, Lagos' gives you 'Lagos'.",
      ["Type =TEXTAFTER(cell, \", \").", "Use instance −1 to count from the end.", "Combine with UPPER or TRIM as needed."],
      "Use −1 to get the last word, whatever comes before it.",
      "Assuming every row has the separator; add if_not_found.",
      "The analyst needs the state from Store_Location, and customers' surnames.",
      ["State: =TEXTAFTER([@Store_Location], \", \").", "Surname: =UPPER(TEXTAFTER([@TRIM], \" \", -1)).", "Fill down."],
      [["Input", "Result"], ["Ikeja, Lagos", "Lagos"], ["Uyo, Akwa Ibom", "Akwa Ibom"], ["Funmilayo Buhari (surname)", "BUHARI"]],
      "TEXTAFTER gives the state (even two-word states) and the surname without counting characters."),

    # ---------- dates ----------
    F("2.8.50",
      "**TODAY** returns **today's date**. Syntax: **=TODAY()**, with no arguments. It updates every time the file "
      "recalculates, so it suits 'as at' dates, ages and days overdue.",
      "TODAY is like a wall calendar that turns its own page every morning.",
      ["Type =TODAY().", "Subtract a date from it to get days elapsed.", "Format the result as a number."],
      "Use TODAY for ageing; it updates by itself.",
      "Using TODAY in a report that must show a fixed date; it will change.",
      "The analyst wants to know how many days have passed since the end of FY2025.",
      ["Type =TODAY().", "Type =TODAY() - DATE(2025, 12, 31).", "Reopen tomorrow: the answer grows by 1."],
      [["Formula", "Result"], ["=TODAY()", "Today's date"], ["=TODAY() - DATE(2025, 12, 31)", "Days since year end (changes daily)"]],
      "Dates are numbers, so you can subtract them. The result moves forward every day."),

    F("2.8.51",
      "**NOW** returns today's date **and the current time**. Syntax: **=NOW()**. Use it for 'last refreshed' "
      "stamps or for measuring hours. Because it includes a time, it is never equal to a plain date.",
      "NOW is like a receipt printer that stamps both the date and the exact time.",
      ["Type =NOW().", "Format as dd-mmm-yyyy hh:mm.", "Use NOW() - TODAY() to see just the time part."],
      "Use NOW for time stamps; use TODAY for date comparisons.",
      "Testing NOW() = a date, which is almost never TRUE.",
      "The analyst adds a 'last refreshed' stamp to the dashboard.",
      ["Type =\"Last refreshed: \" & TEXT(NOW(), \"dd-mmm-yyyy hh:mm\").", "Press F9 to recalculate.", "The stamp updates."],
      [["Formula", "Shows"], ["=NOW()", "Date and time"], ["=NOW() - TODAY()", "Just the time (a fraction of a day)"], ["Stamp sentence", "Last refreshed: dd-mmm-yyyy hh:mm"]],
      "Readers can see how fresh the figures are. NOW also changes whenever the file recalculates."),

    F("2.8.52",
      "**DATE** builds a real date from a **year, month and day**. Syntax: **=DATE(year, month, day)**. It is the "
      "safe way to turn text pieces into a date, and it handles overflow: DATE(2025, 13, 1) is 1 January 2026.",
      "DATE is like filling in a form with separate boxes for day, month and year. Excel then puts them together "
      "properly.",
      ["Get the pieces (with LEFT, MID, RIGHT if needed).", "Type =DATE(year, month, day).", "Format the cell as a date."],
      "Build dates with DATE instead of typing them as text.",
      "Joining pieces with & and getting text that only looks like a date.",
      "The dirty file has text dates in dd/mm/yyyy, which a US-set Excel reads wrongly.",
      ["For '03/02/2025', the day is LEFT(…, 2) and the month is MID(…, 4, 2).", "Type =DATE(RIGHT(c, 4), MID(c, 4, 2), LEFT(c, 2)).", "The result is 3 February 2025."],
      [["Formula", "Result"], ["=DATE(2025, 12, 31)", "31-Dec-2025"], ["=DATE(2025, 13, 1)", "01-Jan-2026"], ["dd/mm text dates in dirty file", "106"]],
      "DATE puts day, month and year in the right places, whatever the computer's date setting."),

    F("2.8.53",
      "**DATEDIF** returns the gap between two dates in complete **days (\"d\"), months (\"m\") or years (\"y\")**. "
      "Syntax: **=DATEDIF(start, end, unit)**. It is an older function that isn't in the formula list, but it still "
      "works.",
      "DATEDIF is like counting how many full nights a patient stayed in hospital, from admission to discharge.",
      ["Type =DATEDIF(start, end, \"d\").", "Use \"m\" or \"y\" for whole months or years.", "Make sure start is before end, or you get #NUM!."],
      "Use \"y\" for ages; it counts only complete years.",
      "Putting the later date first, which returns #NUM!.",
      "Lifecare wants the length of stay for each 2024 admission.",
      ["Type =DATEDIF([@Admission_Date], [@Discharge_Date], \"d\").", "Fill down.", "Average the results."],
      [["Admission_ID", "Admitted → discharged", "Days"], ["ADM2024-51042", "1 Jan → 20 Jan 2024", "19"], ["ADM2024-51043", "1 Jan → 8 Jan 2024", "7"], ["All 3,819 stays in 2024", "Average", "15.75"]],
      "Each stay is measured in whole days. Across 2024 the average stay was about 16 days.", who=LC, split=4.20),

    F("2.8.54",
      "**EOMONTH** returns the **last day of a month**, a set number of months before or after a date. Syntax: "
      "**=EOMONTH(start_date, months)**. 0 = this month, 1 = next month. It handles 28, 29, 30 and 31-day months.",
      "EOMONTH is like asking 'when is month-end?' without having to remember which months have 30 days.",
      ["Type =EOMONTH(date, 0) for this month's end.", "Use 1 for next month, −1 for last month.", "Format the result as a date."],
      "Use EOMONTH for month-end deadlines and billing dates.",
      "Forgetting to format the cell, so you see a number like 45322.",
      "Lifecare bills HMOs at the end of the month of discharge.",
      ["Type =EOMONTH([@Discharge_Date], 0).", "Fill down.", "Try months = 1 to see the next billing date."],
      [["Discharge date", "Formula", "Result"], ["20-Jan-2024", "=EOMONTH(…, 0)", "31-Jan-2024"], ["20-Jan-2024", "=EOMONTH(…, 1)", "29-Feb-2024"]],
      "EOMONTH got February 2024 right (29 days, a leap year) with no extra work.", who=LC, split=4.20),

    F("2.8.55",
      "**NETWORKDAYS** counts **working days** (Monday to Friday) between two dates, counting both ends. Syntax: "
      "**=NETWORKDAYS(start, end, [holidays])**. Add a holiday list so public holidays are not counted.",
      "NETWORKDAYS is like counting only school days between two dates, skipping weekends and public holidays.",
      ["Type =NETWORKDAYS(start, end).", "Make a list of public holidays.", "Add it as the third argument."],
      "Always supply a holiday list for Nigerian public holidays.",
      "Forgetting that both the start and end day are counted.",
      "Operations wants shipping time in working days, not calendar days.",
      ["Order 1 Jan 2025, shipped 6 Jan 2025: =NETWORKDAYS(…) = 4.", "Add 1 Jan (New Year) as a holiday: result 3.", "Average by Ship_Mode."],
      [["Measure", "Result"], ["1 Jan → 6 Jan 2025", "4 working days"], ["Same, with 1 Jan as holiday", "3"], ["Standard Class average (2025)", "4.44 working days"]],
      "Working days are what customers feel. The holiday list changes the answer."),

    F("2.8.56",
      "**WORKDAY** gives the date a set number of **working days** after a start date, skipping weekends and "
      "holidays. Syntax: **=WORKDAY(start, days, [holidays])**. Use it to set realistic due dates.",
      "WORKDAY is like saying 'your result will be ready in 7 working days' and working out the actual date.",
      ["Type =WORKDAY(start, number of days).", "Add the holiday list.", "Format the result as a date."],
      "Use WORKDAY for promised dates, so they never land on a weekend.",
      "Adding days with + and promising a Sunday.",
      "Lifecare calls patients 7 working days after discharge.",
      ["Type =WORKDAY([@Discharge_Date], 7).", "Fill down.", "Check the weekday of the result."],
      [["Discharge date", "+7 working days"], ["Sat 20-Jan-2024", "Tue 30-Jan-2024"], ["Mon 08-Jan-2024", "Wed 17-Jan-2024"], ["Wed 03-Jan-2024", "Fri 12-Jan-2024"]],
      "Every follow-up call lands on a weekday, even after a Saturday discharge.", who=LC),

    # ---------- lookups ----------
    F("2.8.57",
      "**VLOOKUP** finds a value in the **first column** of a table and returns a value from another column in the "
      "same row. Syntax: **=VLOOKUP(value, table, col_index, FALSE)**. Always use **FALSE** for an exact match.",
      "VLOOKUP is like finding a name in a phone directory, then reading across to the number in the column on the "
      "right.",
      ["Make sure the lookup column is the first column of the range.", "Count the column number to return.", "Type =VLOOKUP(value, range, number, FALSE)."],
      "Always type FALSE as the fourth argument.",
      "Leaving out FALSE, so Excel returns a nearby wrong match.",
      "Head Office types a store ID and wants the store manager's name.",
      ["Type ST09 in cell E63.", "=VLOOKUP($E$63, tStores[[Store_ID]:[Store_Manager]], 8, FALSE).", "Try ST02, then ST13."],
      [["Store_ID typed", "Result"], ["ST09", "Ladi Danjuma"], ["ST02", "Tunde Bakare"], ["ST13", "#N/A (no such store)"]],
      "VLOOKUP found the manager. Column 8 is counted from Store_ID, so inserting a column would break it."),

    F("2.8.58",
      "**XLOOKUP** finds a value in **any column** and returns from **any other column**, left or right. Syntax: "
      "**=XLOOKUP(value, lookup_array, return_array, [if_not_found])**. It is exact by default and has its own "
      "not-found message.",
      "XLOOKUP is like a smart phone contact search: type the name and get the number, wherever it is stored.",
      ["Pick the column to search.", "Pick the column to return.", "Add a not-found message."],
      "Prefer XLOOKUP in Excel 365 and 2021.",
      "Selecting search and return columns of different lengths.",
      "Head Office wants the 2025 target for a typed store ID.",
      ["Type ST09 in E63.", "=XLOOKUP($E$63, tStores[Store_ID], tStores[Target_2025], \"not found\").", "Try ST13."],
      [["Store_ID typed", "Result"], ["ST09", "₦61,000,000"], ["ST09 attainment (sales ÷ target)", "108.6%"], ["ST13", "not found"]],
      "XLOOKUP needs no column number, and shows a clear message for a missing store."),

    F("2.8.59",
      "**INDEX** returns the value at a given **position** in a range. Syntax: **=INDEX(range, row_num, "
      "[col_num])**. On its own it needs a number; with MATCH it becomes a flexible lookup.",
      "INDEX is like a cinema seat number: 'row 9' takes you straight to the seat, without searching.",
      ["Select the column you want a value from.", "Give the row position.", "Use MATCH to work out the position."],
      "Pair INDEX with MATCH for lookups that survive column changes.",
      "Giving a position bigger than the range, which returns #REF!.",
      "The analyst wants the city for store ST09.",
      ["=INDEX(tStores[City], 9) gives the 9th city.", "Replace 9 with MATCH($E$63, tStores[Store_ID], 0).", "Change E63 to test."],
      [["Formula", "Result"], ["=INDEX(tStores[City], 9)", "Jos"], ["=INDEX(tStores[City], MATCH(\"ST09\", …, 0))", "Jos"]],
      "INDEX + MATCH found the city without counting columns, so adding columns won't break it."),

    F("2.8.60",
      "**MATCH** returns the **position** of a value in a list. Syntax: **=MATCH(value, range, 0)**. Use **0** for an "
      "exact match. Its main job is to give INDEX the row number.",
      "MATCH is like asking 'what number am I in the queue?': it gives your position, not your name.",
      ["Select a one-column list.", "Type =MATCH(value, list, 0).", "Feed the result into INDEX."],
      "Always use 0 for an exact match.",
      "Leaving out 0, so MATCH returns a nearby wrong position.",
      "Where is ST09 in the store list?",
      ["Type =MATCH(\"ST09\", tStores[Store_ID], 0).", "Read the position.", "Try ST13: the result is #N/A."],
      [["Lookup", "=MATCH(…, 0)"], ["ST09", "9"], ["ST12", "12"], ["ST13", "#N/A"]],
      "MATCH says ST09 is 9th in the list. INDEX then uses that number to fetch any column."),

    # ---------- dynamic arrays ----------
    F("2.8.61",
      "**FILTER** returns all rows that meet a condition, and the results **spill** into the cells below. Syntax: "
      "**=FILTER(array, include, [if_empty])**. It is available in Excel 365 and 2021. Change the condition and the "
      "list updates.",
      "FILTER is like asking a waiter for 'only the vegetarian dishes': you get a new, shorter menu, and the full menu "
      "stays the same.",
      ["Select the columns to return.", "Write the condition, e.g. tStores[Zone]=E74.", "Add an if_empty message."],
      "Leave empty cells below; spilled results need room.",
      "Typing in the spill area, which gives #SPILL!.",
      "Head Office wants the stores in a chosen zone, with their 2025 sales.",
      ["Type South-South in E74.", "=FILTER(tStores[[Store_Name]:[Store_Name]], tStores[Zone]=$E$74).", "Do the same for Sales_2025 next to it."],
      [["Store_Name", "Sales_2025"], ["Port Harcourt - GRA", "₦93,922,995"], ["Benin City - Ring Road", "₦62,116,000"], ["Uyo - Ewet", "₦57,043,909"]],
      "Three South-South stores spill out. Type another zone in E74 and the list redraws itself."),

    F("2.8.62",
      "**UNIQUE** returns the **distinct** values in a list, spilling them. Syntax: **=UNIQUE(array, [by_col], "
      "[exactly_once])**. Set exactly_once to TRUE to list only values that appear **once**.",
      "UNIQUE is like writing each guest's name once on a list, even if they came to the party three times.",
      ["Type =UNIQUE(column).", "Wrap in ROWS(…) to count the distinct values.", "Try exactly_once = TRUE."],
      "Use ROWS(UNIQUE(…)) to count distinct values.",
      "Running UNIQUE on uncleaned text, so 'Lagos' and 'Lagos ' appear twice.",
      "The analyst wants the list of zones, and how many distinct customers bought in 2025.",
      ["=UNIQUE(tStores[Zone]).", "=ROWS(UNIQUE(tSales[Customer_ID])).", "=ROWS(UNIQUE(tSales[Customer_ID], , TRUE))."],
      [["Formula", "Result"], ["=UNIQUE(tStores[Zone])", "4 zones spill"], ["Distinct customers", "693"], ["Customers on just one line", "96"]],
      "693 customers bought in 2025; 96 of them appear on only one sales line."),

    F("2.8.63",
      "**SORT** returns a **sorted copy** of a range, spilling the result. Syntax: **=SORT(array, [sort_index], "
      "[sort_order])**. sort_order −1 means high to low. The original data is not touched.",
      "SORT is like photocopying a class list and arranging the copy by score: the original register stays in order.",
      ["Select the range to sort.", "Choose the column number to sort by.", "Use −1 for descending."],
      "Use SORT for live rankings that update themselves.",
      "Forgetting −1, so the lowest values come first.",
      "Head Office wants stores ranked by 2025 sales, highest first.",
      ["=SORT(CHOOSE({1,2}, tStores[Store_Name], tStores[Sales_2025]), 2, -1).", "Press Enter; 12 rows spill.", "Read the top three."],
      [["Rank", "Store", "Sales_2025"], ["1", "Lagos - Ikeja", "₦146,388,345"], ["2", "Port Harcourt - GRA", "₦93,922,995"], ["3", "Lagos - Lekki", "₦70,819,092"], ["12", "Owerri - New Owerri", "₦31,585,710"]],
      "Ikeja leads by a wide margin. The ranking updates if any store's sales change.", split=4.20),

    F("2.8.64",
      "**SORTBY** sorts one range by **another** range or calculation. Syntax: **=SORTBY(array, by_array, "
      "[order])**. The sort column doesn't have to appear in the result, so you can sort by a calculated value.",
      "SORTBY is like lining pupils up by their exam score while only reading out their names.",
      ["Select what to show (e.g. store names).", "Give what to sort by (e.g. sales ÷ target).", "Use −1 for highest first."],
      "Use SORTBY to rank by a calculated measure.",
      "Using by-arrays of a different size from the array.",
      "Head Office wants store names ranked by target attainment, not by sales.",
      ["=SORTBY(tStores[Store_Name], tStores[Sales_2025] / tStores[Target_2025], -1).", "Press Enter; 12 names spill.", "Compare with the SORT ranking."],
      [["Rank", "Store", "Attainment"], ["1", "Lagos - Lekki", "112.4%"], ["2", "Jos - Rayfield", "108.6%"], ["3", "Benin City - Ring Road", "105.3%"], ["12", "Owerri - New Owerri", "81.0%"]],
      "Lekki tops attainment, although Ikeja tops sales. The measure you sort by changes the story.", split=4.20),

    F("2.8.65",
      "**LET** gives **names** to parts of a formula, so each part is calculated once and the formula is easier to "
      "read. Syntax: **=LET(name1, value1, name2, value2, …, result)**. The last argument is the answer.",
      "LET is like writing 'Let x = total sales' at the top of a maths answer, then using x instead of rewriting the "
      "whole sum.",
      ["Name each piece: s, SUM(tSales[Sales]).", "Add more names as needed.", "End with the calculation that uses the names."],
      "Give names that explain themselves, like actual and target.",
      "Forgetting the final calculation, which gives an error.",
      "Head Office wants one readable formula that says how far the chain is from its 2025 target.",
      ["=LET(actual, SUM(tSales[Sales]), target, SUM(tStores[Target_2025]),", "gap, actual - target,", "IF(gap >= 0, \"Ahead by \", \"Behind by \") & TEXT(ABS(gap), \"₦#,##0\"))"],
      [["Formula", "Result"], ["=LET(s, SUM(Sales), p, SUM(Profit), p / s)", "0.1274 (12.7%)"], ["Target gap with LET", "Behind by ₦13,784,745"]],
      "LET turns a long formula into readable steps. The chain is ₦13.8m behind target."),
]
