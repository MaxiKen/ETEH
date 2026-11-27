from revlib import T

TS = "tSales (the 3,312 clean 2025 lines, sheet D_Sales_2025)"
W = "SuperMart Retail"


def F(tag, what, like, how, tip, trap, case, steps, table, so, who=W, split=4.70):
    return T(tag, what=what, like=like, how=how, tip=tip, trap=trap, who=who, case=case, steps=steps,
             table=table, so=so, split=split,
             tryit="ETEH_Week2_Section_Examples.xlsx  ›  sheet '2.8A Functions 1-33'")


TOPICS = [
    F("2.8.1",
      "**SUM** adds all the numbers in a range. It skips text and blanks. Syntax: **=SUM(range)**. Because it skips "
      "text silently, check that **COUNT** (numbers) equals **COUNTA** (all entries) on the same column first.",
      "SUM is like a cashier adding up a basket. If one item has no price tag (text), it is quietly left out, and "
      "the total is short.",
      ["Click the cell for the answer.", "Type =SUM( and select the column, or use the Table name.", "Press Enter; compare COUNT and COUNTA."],
      "Check COUNT = COUNTA before trusting a SUM.",
      "Trusting a SUM over a column that has numbers stored as text.",
      "Head Office wants total 2025 sales from " + TS + ".",
      ["Type =SUM(tSales[Sales]).", "Check =COUNT(tSales[Sales]) and =COUNTA(tSales[Sales]).", "Both counts should be 3,312."],
      [["Formula", "Result"], ["=SUM(tSales[Sales])", "₦733,215,255"], ["=COUNT(tSales[Sales])", "3,312"], ["=COUNTA(tSales[Sales])", "3,312"]],
      "COUNT equals COUNTA, so every sale was a number and the total of ₦733.2m is complete."),

    F("2.8.2",
      "**SUMIF** adds values that meet **one condition**. Syntax: **=SUMIF(range, criteria, sum_range)**. range is "
      "where the condition is checked; sum_range is what gets added. Criteria can be text (\"Technology\"), a cell, or "
      "a comparison like \">1000\".",
      "SUMIF is like adding up only the green receipts in a pile: look at each receipt's colour, and add it only if "
      "it is green.",
      ["Pick the column to check (range).", "Write the condition (criteria).", "Pick the column to add (sum_range)."],
      "Always fill in sum_range; without it, SUMIF adds the condition column itself.",
      "Leaving out sum_range and getting 0 or a wrong total.",
      "The buyer for electronics asks for 2025 Technology sales.",
      ["Type =SUMIF(tSales[Category], \"Technology\", tSales[Sales]).", "Press Enter.", "Compare with total sales (₦733.2m)."],
      [["Argument", "Value used"], ["range", "tSales[Category]"], ["criteria", "\"Technology\""], ["sum_range", "tSales[Sales]"], ["Result", "₦271,730,811"]],
      "Technology brought in ₦271.7m, about 37% of all 2025 sales."),

    F("2.8.3",
      "**SUMIFS** adds values that meet **several conditions** at once. Syntax: **=SUMIFS(sum_range, range1, "
      "criteria1, range2, criteria2, …)**. Note: the **sum_range comes first**, the opposite of SUMIF.",
      "SUMIFS is like counting money only from green receipts that are also from Lagos. Every condition must be "
      "true for a receipt to count.",
      ["Start with the column to add (sum_range).", "Add pairs: range, criteria.", "Add as many pairs as you need."],
      "Remember the order: SUMIFS starts with what you are adding.",
      "Copying the SUMIF order into SUMIFS, which gives a wrong result or an error.",
      "Management asks: how much Technology did the South-West zone sell in 2025?",
      ["Type =SUMIFS(tSales[Sales],", "add tSales[Category], \"Technology\",", "add tSales[Zone], \"South-West\") and press Enter."],
      [["Argument", "Value used"], ["sum_range", "tSales[Sales]"], ["Condition 1", "Category = Technology"], ["Condition 2", "Zone = South-West"], ["Result", "₦95,959,148"]],
      "South-West sold ₦96.0m of Technology, over a third of all Technology sales."),

    F("2.8.4",
      "**AVERAGE** gives the mean: the total divided by the count. It **ignores blanks** but **includes zeros**. "
      "Syntax: **=AVERAGE(range)**. If missing values were typed as 0, the average comes out too low.",
      "AVERAGE is like sharing a pot of food equally. Each plate (each number) gets the same share, including an "
      "empty plate someone put on the table.",
      ["Type =AVERAGE( and select the column.", "Press Enter.", "Check for zeros that really mean 'unknown'."],
      "Make sure zeros are real zeros before averaging.",
      "Typing 0 for 'unknown' and dragging the average down.",
      "The analyst wants the average value of one 2025 sales line.",
      ["Type =AVERAGE(tSales[Sales]).", "Check: total ÷ count gives the same answer.", "Compare with the median (2.8.13)."],
      [["Formula", "Result"], ["=AVERAGE(tSales[Sales])", "₦221,381"], ["Check: 733,215,255 ÷ 3,312", "₦221,381"]],
      "The average line is worth ₦221,381. A few very large sales pull this up; see MEDIAN."),

    F("2.8.5",
      "**AVERAGEIFS** averages values that meet **several conditions**. Syntax: **=AVERAGEIFS(average_range, "
      "range1, criteria1, …)**. Like SUMIFS, the **range to average comes first**.",
      "It is like finding the average height of girls in JSS2 only: first pick the pupils who meet both conditions, "
      "then average them.",
      ["Start with the column to average.", "Add range and criteria pairs.", "Press Enter; check the matching count with COUNTIFS."],
      "Check the count with COUNTIFS; an average of very few rows is unreliable.",
      "Averaging when no rows match, which gives #DIV/0!.",
      "Ikeja's manager asks: what is the average Corporate sales line at Lagos - Ikeja (ST01)?",
      ["Type =AVERAGEIFS(tSales[Sales], tSales[Segment], \"Corporate\", tSales[Store_ID], \"ST01\").", "Check the count with COUNTIFS (same conditions).", "Compare with the chain average."],
      [["Measure", "Result"], ["Matching lines (COUNTIFS)", "218"], ["Average Corporate line at Ikeja", "₦270,640"], ["Chain average line", "₦221,381"]],
      "Ikeja's Corporate lines average ₦270,640, about 22% above the chain average."),

    F("2.8.6",
      "**COUNT** counts cells that hold **numbers only**. It ignores text, blanks and errors. Syntax: "
      "**=COUNT(range)**. Compare it with COUNTA: if COUNT is lower, some 'numbers' are stored as text.",
      "COUNT is like a teacher counting only scripts with a score written on them, and skipping any marked 'absent'.",
      ["Type =COUNT( and select a numeric column.", "Type =COUNTA( for the same column.", "If COUNT < COUNTA, look for text numbers."],
      "Use COUNT versus COUNTA as a quick test for text numbers.",
      "Using COUNT to count records; it misses rows where that column is text or blank.",
      "The analyst checks that every Quantity in " + TS + " is a real number.",
      ["Type =COUNT(tSales[Quantity]).", "Type =COUNTA(tSales[Quantity]).", "Compare the two results."],
      [["Formula", "Result"], ["=COUNT(tSales[Quantity])", "3,312"], ["=COUNTA(tSales[Quantity])", "3,312"], ["Difference", "0, so all numeric"]],
      "Every quantity is a real number. In the dirty file, the same test reveals the text values."),

    F("2.8.7",
      "**COUNTA** counts every cell that is **not empty**, whatever it holds. Syntax: **=COUNTA(range)**. Use it to "
      "count records (on an ID column) and as the comparison for COUNT.",
      "COUNTA is like counting every chair with something on it: a person, a bag or a coat. Only empty chairs are "
      "skipped.",
      ["Pick a column that is always filled (an ID).", "Type =COUNTA( and select it.", "Compare with the expected record count."],
      "Count records on an ID column that is never blank.",
      "Counting on an optional column, so blanks lower the count.",
      "The analyst confirms how many sales lines the clean table holds.",
      ["Type =COUNTA(tSales[Order_ID]).", "Compare with the POS control total.", "Count distinct orders separately (2.8B)."],
      [["Formula", "Result"], ["=COUNTA(tSales[Order_ID])", "3,312 lines"], ["POS control", "3,312"], ["Distinct orders (for comparison)", "1,687"]],
      "3,312 lines match the POS control. Remember: lines are not orders; orders are counted differently."),

    F("2.8.8",
      "**COUNTBLANK** counts **empty** cells in a range. Syntax: **=COUNTBLANK(range)**. Use it to check that "
      "required fields are filled in. Note: it also counts formulas that return \"\" (empty text).",
      "COUNTBLANK is like checking an attendance register for empty boxes: each one is a pupil nobody marked.",
      ["Select a column that should always be filled.", "Type =COUNTBLANK( and select it.", "Investigate any result above 0."],
      "Run COUNTBLANK on every required column before analysis.",
      "Assuming no blanks because the first screen of data looks full.",
      "The analyst checks Quantity in the raw dirty file (table tDirty) for missing values.",
      ["Type =COUNTBLANK(tDirty[Quantity]).", "Subtract the 15 blank spacer rows.", "The rest are real lines with no quantity."],
      [["Part", "Count"], ["=COUNTBLANK(tDirty[Quantity])", "46"], ["Blank spacer rows", "15"], ["Real lines missing Quantity", "31"]],
      "31 real sales lines have no quantity. They must be flagged, not filled with 0."),

    F("2.8.9",
      "**COUNTIF** counts cells that meet **one condition**. Syntax: **=COUNTIF(range, criteria)**. Criteria can be "
      "text, a number, or a comparison in quotes such as \"<0\". Use it to count exceptions and duplicates.",
      "COUNTIF is like a gateman counting only cars with Lagos plates: one rule, one tally.",
      ["Select the column to check.", "Write the condition in quotes, e.g. \"<0\".", "Press Enter."],
      "Put comparisons in quotes: \"<0\", \">=1000\".",
      "Typing <0 without quotes, which Excel rejects.",
      "Management asks how many 2025 lines lost money.",
      ["Type =COUNTIF(tSales[Profit], \"<0\").", "Press Enter.", "Divide by 3,312 for the share."],
      [["Formula", "Result"], ["=COUNTIF(tSales[Profit], \"<0\")", "620 lines"], ["Share of all lines", "18.7%"]],
      "620 lines (almost one in five) lost money. COUNTIF turns a feeling into a number."),

    F("2.8.10",
      "**COUNTIFS** counts rows that meet **several conditions**. Syntax: **=COUNTIFS(range1, criteria1, range2, "
      "criteria2, …)**. Every condition must be true for a row to count.",
      "COUNTIFS is like counting pupils who are in JSS3 **and** in the science club **and** scored above 70: only "
      "those who pass all three checks.",
      ["Add the first range and condition.", "Add more range, condition pairs.", "Press Enter; try removing one condition to compare."],
      "Build it one condition at a time and check the count each time.",
      "Using ranges of different sizes, which gives #VALUE!.",
      "The analyst asks: how many Furniture lines lost money at a discount of 30% or more?",
      ["=COUNTIFS(tSales[Category], \"Furniture\", tSales[Profit], \"<0\")", "Add: tSales[Discount], \">=0.3\".", "Compare both counts."],
      [["Conditions", "Lines"], ["Furniture and Profit < 0", "235"], ["… and Discount 30%+", "167"]],
      "167 of 235 loss-making Furniture lines had discounts of 30% or more. Deep discounts drive the losses."),

    F("2.8.11",
      "**SUBTOTAL** adds (or counts, averages…) only the **visible** rows when a filter is on. Syntax: "
      "**=SUBTOTAL(function_num, range)**. Use **109** for SUM ignoring hidden rows, **101** for AVERAGE, **103** for "
      "COUNTA.",
      "SUBTOTAL is like counting only the people still in the room after some have stepped outside. SUM counts "
      "everyone on the register.",
      ["Type =SUBTOTAL(109, tSales[Sales]).", "Apply a filter to the data.", "Watch SUBTOTAL change while SUM stays the same."],
      "Use SUBTOTAL(109, …) for totals above filtered lists.",
      "Using SUM above a filtered list and reporting the unfiltered total.",
      "The analyst filters " + TS + " to the South-West zone and wants the visible total.",
      ["Type =SUBTOTAL(109, tSales[Sales]) above the table.", "Filter Zone = South-West.", "Compare with =SUM(tSales[Sales])."],
      [["Formula", "No filter", "Zone = South-West"], ["=SUM(tSales[Sales])", "₦733.2m", "₦733.2m"], ["=SUBTOTAL(109, tSales[Sales])", "₦733.2m", "₦250.1m"]],
      "Only SUBTOTAL follows the filter. Use it whenever readers will filter the list."),

    F("2.8.12",
      "**AGGREGATE** does 19 calculations (SUM, LARGE, MEDIAN…) and can **skip errors and hidden rows**. Syntax: "
      "**=AGGREGATE(function_num, options, array, [k])**. Option **6** ignores errors. It is useful when a column "
      "contains #N/A or #DIV/0!.",
      "AGGREGATE is like a sieve that also removes stones: it does the maths while ignoring the bits that would break "
      "it.",
      ["Choose the function number (14 = LARGE, 9 = SUM).", "Choose option 6 to ignore errors.", "Give the range and k if needed."],
      "Use AGGREGATE when a column contains errors you can't remove yet.",
      "Using it to hide errors you should actually fix.",
      "The analyst wants the largest single 2025 sale in the South-East zone, without a helper column.",
      ["=AGGREGATE(14, 6, tSales[Sales] / (tSales[Zone]=\"South-East\"), 1)", "Non-South-East rows divide by 0, creating errors that option 6 skips.", "Press Enter."],
      [["Part", "Meaning"], ["14", "LARGE"], ["6", "Ignore error values"], ["Result", "₦7,999,980 (3D printer, Onitsha)"]],
      "AGGREGATE found the zone's largest sale in one formula. The errors were created on purpose and skipped."),

    F("2.8.13",
      "**MEDIAN** returns the **middle** value when the numbers are sorted. Syntax: **=MEDIAN(range)**. Unlike the "
      "average, it isn't pulled up by a few very large values, so it is better for typical sales, incomes or waiting "
      "times.",
      "If ten people earn ₦100,000 and one earns ₦50 million, the average says everyone is rich. The median shows "
      "what a typical person earns.",
      ["Type =MEDIAN( and select the column.", "Press Enter.", "Compare it with AVERAGE."],
      "Report the median when a few very large values exist.",
      "Calling the average 'typical' when the data is skewed.",
      "Is ₦221,381 a typical 2025 sales line? The analyst checks the median.",
      ["Type =MEDIAN(tSales[Sales]).", "Type =AVERAGE(tSales[Sales]).", "Compare the two."],
      [["Measure", "Result"], ["Median line", "₦53,810"], ["Average line", "₦221,381"]],
      "Half of all lines are below ₦53,810. A few huge sales make the average about four times larger."),

    F("2.8.14",
      "**MODE.SNGL** returns the **most common** value. Syntax: **=MODE.SNGL(range)**. Use it for typical basket "
      "sizes or quantities. It returns #N/A if no value repeats.",
      "MODE is like the most popular dish at a canteen: the one ordered more often than any other.",
      ["Type =MODE.SNGL( and select the column.", "Press Enter.", "Check how often it occurs with COUNTIF."],
      "Check how often the mode occurs; a 'most common' value can still be rare.",
      "Using MODE on unique values like IDs; it returns #N/A.",
      "The store team asks: what quantity do customers most often buy on one line?",
      ["Type =MODE.SNGL(tSales[Quantity]).", "Count it: =COUNTIF(tSales[Quantity], 2).", "Compare with the next most common."],
      [["Quantity", "Lines"], ["2 (the mode)", "781"], ["3", "759"], ["5", "441"]],
      "Customers most often buy 2 units, but 3 is almost as common. Shelf packs of 2–3 suit most buyers."),

    F("2.8.15",
      "**MAX** returns the **largest** number in a range. Syntax: **=MAX(range)**. Use it to find peaks, and as a "
      "check: a maximum that is impossibly high signals a data error.",
      "MAX is like finding the tallest pupil in a class line-up. If the 'tallest' is 5 metres, someone wrote the "
      "height wrongly.",
      ["Type =MAX( and select the column.", "Press Enter.", "Ask: is this value believable?"],
      "Always ask whether the maximum is believable.",
      "Reporting a maximum without checking if it is a typo.",
      "Head Office asks for the single largest 2025 sales line.",
      ["Type =MAX(tSales[Sales]).", "Find the row with Ctrl+F or a filter.", "Check it is a genuine sale."],
      [["Formula", "Result"], ["=MAX(tSales[Sales])", "₦13,999,960"], ["Product", "Canon imageCLASS 2200 copier"], ["Store", "Lagos - Lekki"]],
      "The largest line is a ₦14.0m copier sale at Lekki: large, but believable for office equipment."),

    F("2.8.16",
      "**MIN** returns the **smallest** number in a range. Syntax: **=MIN(range)**. On a profit column it finds the "
      "worst loss. As a check, a minimum of 0 in a price column or a negative duration signals bad data.",
      "MIN is like finding the lowest score in a test. It tells you who needs help, or that a score was entered "
      "wrongly.",
      ["Type =MIN( and select the column.", "Press Enter.", "Look up the row behind the value."],
      "Use MIN on every numeric column as a quick sense check.",
      "Ignoring a negative minimum in a column that can't be negative.",
      "Management asks for the worst single loss in 2025.",
      ["Type =MIN(tSales[Profit]).", "Filter Profit to that value.", "Read the product, store and discount."],
      [["Field", "Value"], ["=MIN(tSales[Profit])", "−₦3,839,990"], ["Product", "Cubify CubeX 3D Printer"], ["Store / discount", "Onitsha - Main Market / 50%"]],
      "The worst loss came from a 50% discount on a 3D printer. One discount decision cost ₦3.8m."),

    F("2.8.17",
      "**LARGE** returns the **k-th largest** value. Syntax: **=LARGE(range, k)**. k = 1 is the same as MAX; k = 3 "
      "is the third largest. Use it for top-N lists without sorting the data.",
      "LARGE is like announcing the 1st, 2nd and 3rd positions at a sports day: you pick out each place in order.",
      ["Type =LARGE(range, 1) for the top value.", "Change k to 2, 3… for the next places.", "Use cell references for k to build a top-N table."],
      "Put 1, 2, 3 in cells and refer to them as k.",
      "Expecting LARGE to return the name; it returns only the value (use lookups for names).",
      "The analyst builds a top-3 list of 2025 sales lines.",
      ["Type 1, 2, 3 in cells A2:A4.", "In B2: =LARGE(tSales[Sales], A2); fill down.", "Look up the products with XLOOKUP (2.8B)."],
      [["k", "=LARGE(tSales[Sales], k)"], ["1", "₦13,999,960"], ["2", "₦11,199,968"], ["3", "₦10,499,970"]],
      "The top three lines are each over ₦10m. The list updates itself when the data changes."),

    F("2.8.18",
      "**SMALL** returns the **k-th smallest** value. Syntax: **=SMALL(range, k)**. It is the mirror of LARGE: use it "
      "for bottom-N lists, such as the worst losses.",
      "SMALL is like listing the three cheapest phones in a shop, without rearranging the whole display.",
      ["Type =SMALL(range, 1) for the smallest.", "Change k for the next ones.", "Pair with lookups to show names."],
      "Use SMALL on Profit to list the worst losses.",
      "Using SMALL on a column with zeros or blanks you didn't expect.",
      "Management wants the three worst losses of 2025.",
      ["Type 1, 2, 3 in cells A2:A4.", "In B2: =SMALL(tSales[Profit], A2); fill down.", "Look up the product for each."],
      [["k", "=SMALL(tSales[Profit], k)"], ["1", "−₦3,839,990"], ["2", "−₦3,399,980"], ["3", "−₦2,929,484"]],
      "The three worst lines lost ₦10.2m together. Each should be reviewed with the store."),

    F("2.8.19",
      "**ROUND** rounds a number to a set number of digits, and **changes the stored value**. Syntax: "
      "**=ROUND(number, digits)**. Digits 2 = two decimals; 0 = whole number; **−3 = nearest thousand**. Number "
      "formatting only changes how it looks.",
      "ROUND is like a trader saying '₦221,000' instead of '₦221,381.42'. The figure is easier to say, and he "
      "actually changes the price.",
      ["Type =ROUND(value, digits).", "Use negative digits to round to tens, hundreds, thousands.", "Use formatting instead if you only want it to look rounded."],
      "Round final results, not the inputs to further calculations.",
      "Rounding every step, so small errors add up in the total.",
      "The analyst reports the average line value to the nearest ₦1,000.",
      ["Type =ROUND(AVERAGE(tSales[Sales]), -3).", "Compare with the unrounded average.", "Try digits 0 and −2 too."],
      [["Formula", "Result"], ["=AVERAGE(tSales[Sales])", "221,381.42"], ["=ROUND(…, 0)", "221,381"], ["=ROUND(…, −3)", "221,000"]],
      "−3 rounds to the nearest thousand, which suits a headline figure."),

    F("2.8.20",
      "**ROUNDUP** always rounds **away from zero** (upwards for positive numbers). Syntax: **=ROUNDUP(number, "
      "digits)**. Use it when falling short is not allowed: cartons, buses, staff, beds.",
      "If 45 pupils go on an excursion and each bus holds 40, you need 2 buses, not 1.1. You always round up.",
      ["Divide the amount by the capacity.", "Wrap it: =ROUNDUP(amount / capacity, 0).", "Check the result covers everything."],
      "Use ROUNDUP for anything you can't supply in part.",
      "Using ROUND, which can round down and leave items with no carton.",
      "2025 Technology units must be packed in cartons of 12. How many cartons are needed?",
      ["Total units: =SUMIFS(tSales[Quantity], tSales[Category], \"Technology\").", "Divide by 12.", "Wrap in ROUNDUP(…, 0)."],
      [["Step", "Result"], ["Technology units", "2,363"], ["Units ÷ 12", "196.92"], ["=ROUNDUP(…, 0)", "197 cartons"]],
      "196 cartons would leave 11 units unpacked. ROUNDUP guarantees enough cartons."),

    F("2.8.21",
      "**ROUNDDOWN** always rounds **towards zero**. Syntax: **=ROUNDDOWN(number, digits)**. Use it when you must "
      "never go over a limit, such as entitlements, allowances or a cautious figure.",
      "If a ₦10,000 bonus pool is shared between 3 people, each gets ₦3,333, not ₦3,334, or the pool runs out.",
      ["Write the calculation.", "Wrap it: =ROUNDDOWN(calculation, digits).", "Check the result never exceeds the true value."],
      "Use ROUNDDOWN for amounts that must not exceed a limit.",
      "Using ROUNDDOWN for quantities you must fully supply (use ROUNDUP).",
      "The analyst reports the 2025 chain margin cautiously, to 3 decimal places.",
      ["Margin: =SUM(tSales[Profit]) / SUM(tSales[Sales]).", "Wrap in ROUNDDOWN(…, 3).", "Compare with ROUND(…, 3)."],
      [["Formula", "Result"], ["Exact margin", "0.12744…"], ["=ROUNDDOWN(…, 3)", "0.127 (12.7%)"], ["=ROUND(…, 3)", "0.127 (12.7%)"]],
      "Here both give 12.7%. ROUNDDOWN guarantees you never overstate it."),

    F("2.8.22",
      "**ABS** returns the size of a number **without its sign**: =ABS(−5) is 5. Syntax: **=ABS(number)**. Use it to "
      "measure gaps or errors in either direction, e.g. how far from target, whether above or below.",
      "ABS is like measuring the distance between two towns: it's 50 km whichever direction you travel.",
      ["Calculate the difference (actual − target).", "Wrap it: =ABS(difference).", "Use it in tests like ABS(diff) > limit."],
      "Use ABS to test 'too far from target in either direction'.",
      "Using ABS where the direction matters (over vs under target).",
      "The analyst wants the worst 2025 loss shown as a positive amount for a chart.",
      ["Type =MIN(tSales[Profit]).", "Wrap it: =ABS(MIN(tSales[Profit])).", "Use the result in the chart label."],
      [["Formula", "Result"], ["=MIN(tSales[Profit])", "−3,839,990"], ["=ABS(MIN(tSales[Profit]))", "3,839,990"]],
      "ABS removed the minus sign, so the loss can be shown as a size on a chart."),

    F("2.8.23",
      "**IF** returns one result when a condition is TRUE and another when it is FALSE. Syntax: **=IF(test, "
      "value_if_true, value_if_false)**. It is the basis of flags and labels.",
      "IF is like a gate: 'If you have a ticket, enter; otherwise, wait outside.'",
      ["Write the test, e.g. A2>=B2.", "Write what to show if TRUE.", "Write what to show if FALSE."],
      "Use IF for two outcomes; for more, use IFS.",
      "Nesting many IFs inside each other; hard to read and easy to break.",
      "Did SuperMart meet its total 2025 target of ₦747m across all 12 stores?",
      ["Total sales: =SUM(tSales[Sales]).", "Total target: =SUM(tStores[Target_2025]).", "=IF(total sales >= total target, \"Met\", \"Missed\")."],
      [["Item", "Value"], ["Total sales", "₦733.2m"], ["Total target", "₦747.0m"], ["IF result", "Missed"]],
      "The chain missed its target by ₦13.8m, even though some stores beat theirs."),

    F("2.8.24",
      "**IFS** checks several conditions **in order** and returns the result for the **first TRUE** one. Syntax: "
      "**=IFS(test1, value1, test2, value2, …, TRUE, else_value)**. End with TRUE as a catch-all.",
      "IFS is like a grading table: check 70+ first, then 60+, then 50+. The first band that fits gives the grade.",
      ["List the bands from highest to lowest.", "Write each test and its result.", "End with TRUE, \"Other\"."],
      "Always end IFS with TRUE and a default result.",
      "Putting the bands in the wrong order, so everything matches the first test.",
      "The analyst labels the 2025 chain margin: 15%+ Strong, 8%+ Healthy, otherwise Weak.",
      ["Margin: =SUM(tSales[Profit]) / SUM(tSales[Sales]).", "=IFS(m >= 0.15, \"Strong\", m >= 0.08, \"Healthy\", TRUE, \"Weak\").", "Read the result."],
      [["Test (in order)", "TRUE?"], ["Margin 12.7% >= 15%", "No"], ["Margin 12.7% >= 8%", "Yes: Healthy"], ["TRUE (default)", "Not reached"]],
      "The chain's margin is Healthy. IFS stopped at the first test that was true."),

    F("2.8.25",
      "**AND** returns TRUE only if **all** conditions are TRUE. Syntax: **=AND(test1, test2, …)**. It is usually "
      "used inside IF: =IF(AND(…), \"Yes\", \"No\").",
      "AND is like a job requirement: a degree AND 3 years' experience. Missing either one, and you don't qualify.",
      ["Write each test separately first.", "Combine: =AND(test1, test2).", "Wrap in IF for a readable label."],
      "Test each condition on its own before combining them.",
      "Using AND when any one condition should be enough (use OR).",
      "Did Lagos - Ikeja (ST01) meet its target AND keep a margin above 10%?",
      ["Test 1: Ikeja sales >= target (₦146.4m vs ₦142m).", "Test 2: Ikeja margin > 10% (20.1%).", "=AND(test1, test2)."],
      [["Test", "Result"], ["Sales >= target (103%)", "TRUE"], ["Margin > 10% (20.1%)", "TRUE"], ["AND", "TRUE"]],
      "Ikeja passed both tests, so AND returns TRUE."),

    F("2.8.26",
      "**OR** returns TRUE if **at least one** condition is TRUE. Syntax: **=OR(test1, test2, …)**. Use it for "
      "warning flags, where any single problem is enough to raise attention.",
      "OR is like a smoke alarm with two sensors: if either one detects smoke, the alarm rings.",
      ["Write each warning test.", "Combine: =OR(test1, test2).", "Wrap in IF for a label such as \"Review\"."],
      "Use OR for warning flags; better one extra review than one missed problem.",
      "Using OR where every condition must hold (use AND).",
      "Should Owerri (ST12) be reviewed? The rule: target attainment below 85% OR margin below 5%.",
      ["Test 1: attainment ₦31.6m ÷ ₦39m < 85%.", "Test 2: margin < 5%.", "=OR(test1, test2)."],
      [["Test", "Result"], ["Attainment < 85% (81%)", "TRUE"], ["Margin < 5% (31.0%)", "FALSE"], ["OR", "TRUE: review"]],
      "Owerri is very profitable but far below target. One failed test is enough to flag it."),

    F("2.8.27",
      "**NOT** reverses TRUE and FALSE. Syntax: **=NOT(test)**. It makes rules easier to read when they are naturally "
      "negative, such as NOT(ISBLANK(A2)): 'A2 is not blank'.",
      "NOT is like a 'No Entry' sign: it turns 'enter' into 'don't enter'.",
      ["Write the positive test, e.g. ISBLANK(A2).", "Wrap it: NOT(ISBLANK(A2)).", "Use it inside IF or validation."],
      "Use NOT to express 'is not' rules clearly.",
      "Stacking several NOTs, which makes the logic hard to follow.",
      "The analyst writes a check: are there any loss-making lines at all?",
      ["Count losses: =COUNTIF(tSales[Profit], \"<0\") → 620.", "Test 'no losses': 620 = 0 → FALSE.", "Reverse it: =NOT(COUNTIF(tSales[Profit], \"<0\")=0)."],
      [["Formula", "Result"], ["=COUNTIF(tSales[Profit], \"<0\")=0", "FALSE"], ["=NOT(…)", "TRUE: losses exist"]],
      "NOT turned 'there are no losses: FALSE' into the clearer 'there are losses: TRUE'."),

    F("2.8.28",
      "**IFERROR** shows a replacement when a formula returns **any** error (#DIV/0!, #N/A, #VALUE!…). Syntax: "
      "**=IFERROR(formula, value_if_error)**. It makes reports tidy, but it also hides real mistakes, so use it with "
      "care.",
      "IFERROR is like a curtain over a broken window. It looks tidy from outside, but the window is still broken.",
      ["Write and test the formula without IFERROR first.", "Wrap it: =IFERROR(formula, \"message\").", "Use a clear message, not a blank."],
      "Fix real errors first; use IFERROR only for expected cases.",
      "Wrapping everything in IFERROR, so genuine mistakes are never seen.",
      "A report asks for the margin of store ST13, which does not exist. Profit ÷ Sales is 0 ÷ 0.",
      ["=SUMIFS(Profit for ST13) / SUMIFS(Sales for ST13) gives #DIV/0!.", "Wrap: =IFERROR(…, \"No sales\").", "Read the result."],
      [["Formula", "Result"], ["Profit ÷ Sales for ST13", "#DIV/0!"], ["=IFERROR(…, \"No sales\")", "No sales"]],
      "The report shows a clear message instead of an error. Check that ST13 truly has no sales."),

    F("2.8.29",
      "**IFNA** replaces **only #N/A** errors, the kind a lookup returns when no match is found. Other errors still "
      "show. Syntax: **=IFNA(formula, value_if_na)**. It is the right wrapper for lookups.",
      "IFNA is like a receptionist who says 'not in the directory' only when a name isn't listed, but still reports "
      "a broken phone.",
      ["Write the lookup.", "Wrap it: =IFNA(lookup, \"Not found\").", "Leave other errors visible."],
      "Use IFNA, not IFERROR, around lookups.",
      "Hiding a #REF! or #VALUE! bug with IFERROR.",
      "Someone asks for the manager of store ST13. The store directory has only ST01–ST12.",
      ["=VLOOKUP(\"ST13\", tStores[[Store_ID]:[Store_Manager]], 8, FALSE) gives #N/A.", "Wrap: =IFNA(…, \"Unknown store\").", "Test with ST01: Folake Adeyemi."],
      [["Lookup", "Result"], ["ST01 with IFNA", "Folake Adeyemi"], ["ST13 without IFNA", "#N/A"], ["ST13 with IFNA", "Unknown store"]],
      "A missing store gets a clear message, while real formula errors would still appear."),

    F("2.8.30",
      "**SWITCH** compares one value with a list of options and returns the matching result. Syntax: "
      "**=SWITCH(value, option1, result1, option2, result2, …, default)**. It is neater than nested IFs when "
      "translating codes.",
      "SWITCH is like a hotel key board: room 101's key on hook 101, room 102's on hook 102, and a 'no such room' "
      "sign for anything else.",
      ["Get the code (e.g. LEFT(Product_ID, 3)).", "List each code and its meaning.", "Add a default result at the end."],
      "Always add a default result for unexpected codes.",
      "Leaving out the default, so a new code returns #N/A.",
      "Product IDs start with FUR, OFF or TEC. The analyst turns the code into a category name.",
      ["Get the code: LEFT(Product_ID, 3).", "=SWITCH(code, \"FUR\", \"Furniture\", \"OFF\", \"Office Supplies\", \"TEC\", \"Technology\", \"Unknown\").", "Test on the first line."],
      [["Product_ID", "Code", "SWITCH result"], ["FUR-FU-10003878", "FUR", "Furniture"], ["OFF-ST-10002485", "OFF", "Office Supplies"], ["TEC-AC-10003614", "TEC", "Technology"]],
      "One formula translates every code, and any new code shows 'Unknown' instead of an error."),

    F("2.8.31",
      "**TEXT** turns a number into text with a chosen format. Syntax: **=TEXT(value, format)**, e.g. "
      "TEXT(A1, \"#,##0\") or TEXT(date, \"yyyy-mm\"). Use it to build sentences and labels. The result is text, so "
      "it can't be added up.",
      "TEXT is like writing an amount on a cheque in words and figures: it's for reading, not for further sums.",
      ["Choose the format code (#,##0 or yyyy-mm).", "Join it into a sentence with &.", "Keep the original number for calculations."],
      "Use TEXT only for display; keep numbers as numbers for maths.",
      "Using TEXT results in calculations; they are text.",
      "The analyst wants a headline sentence that updates with the data.",
      ["Type =\"FY2025 sales: \" & TEXT(SUM(tSales[Sales]), \"₦#,##0\").", "Press Enter.", "Change the data and watch it update."],
      [["Formula part", "Result"], ["SUM(tSales[Sales])", "733215255"], ["TEXT(…, \"₦#,##0\")", "₦733,215,255"], ["Full sentence", "FY2025 sales: ₦733,215,255"]],
      "A live, formatted sentence: perfect for report titles and dashboard headlines."),

    F("2.8.32",
      "**CONCAT** joins text from cells and ranges into one string. Syntax: **=CONCAT(text1, text2, …)**. It adds no "
      "spaces or separators; you must include them yourself. The & operator does the same job.",
      "CONCAT is like sewing pieces of fabric together: you also have to sew in any gaps (spaces) you want.",
      ["List the pieces in order.", "Add separators as text, e.g. \" (\".", "Check the result reads correctly."],
      "Remember to add your own spaces and separators.",
      "Forgetting separators and getting 'LagosIkejaST01'.",
      "Reports should show stores as 'Lagos - Ikeja (ST01)'.",
      ["=CONCAT(Store_Name, \" (\", Store_ID, \")\").", "Press Enter.", "Compare with Store_Name & \" (\" & Store_ID & \")\"."],
      [["Piece", "Value"], ["Store_Name", "Lagos - Ikeja"], ["Separator", "\" (\""], ["Store_ID + \")\"", "ST01)"], ["Result", "Lagos - Ikeja (ST01)"]],
      "CONCAT built a readable label. Both CONCAT and & work the same way."),

    F("2.8.33",
      "**TEXTJOIN** joins text with a **separator** you choose and can **skip empty cells**. Syntax: "
      "**=TEXTJOIN(separator, ignore_empty, range)**. It is the best way to list many values in one cell.",
      "TEXTJOIN is like threading beads on a string with a knot between each one. Empty spaces are skipped, so there "
      "are no double knots.",
      ["Choose the separator, e.g. \", \".", "Set ignore_empty to TRUE.", "Give the range (or a FILTER result)."],
      "Set ignore_empty to TRUE to avoid ', ,' gaps.",
      "Joining a huge range into one cell; it becomes unreadable.",
      "The analyst lists the South-East stores in one cell for a summary.",
      ["Filter the stores: FILTER(tStores[Store_Name], tStores[Zone]=\"South-East\").", "Wrap: =TEXTJOIN(\", \", TRUE, …).", "Read the result."],
      [["Part", "Value"], ["Separator", "\", \""], ["ignore_empty", "TRUE"], ["Result", "Enugu - Independence Layout, Onitsha - Main Market, Owerri - New Owerri"]],
      "One cell names all three South-East stores, and it updates if the store list changes.", split=4.20),
]
