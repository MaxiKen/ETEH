from revlib import T

DIRTY = "SuperMart_Sales_2025_DIRTY.csv"

TOPICS = [
    T("2.7.1",
      what="**Manual cleaning** means typing, Find & Replace, TRIM helper columns, filtering and deleting in the "
           "sheet itself. It is quick, visible and **perfectly fine for one-off, small jobs** that will never be "
           "repeated. It stops being the right choice once the task repeats, the file is shared, or someone must check "
           "how the numbers were made.",
      like="Manual cleaning is like sweeping your room by hand: perfect for one room, once. To clean a whole office "
           "block every day, you need a system.",
      how=["Ask: will I ever do this again?",
           "Ask: does anyone need to check how it was done?",
           "If both answers are no, clean by hand."],
      tip="Clean by hand only when the job is small and truly one-off.",
      trap="Saying 'just this once' about a report that comes back every month.",
      who="Multi-Sector",
      case="The analyst receives four cleaning requests in one week. She decides which ones to do by hand.",
      steps=["For each request, ask: will it repeat?",
             "Ask: is it large, shared or checked by others?",
             "Choose manual or Power Query."],
      table=[["Request", "Repeats?", "Tool"],
             ["Tidy 40 names for a one-off letter", "No", "Manual"],
             ["Fix one typo in a finished report", "No", "Manual"],
             ["Clean the monthly POS export", "Every month", "Power Query"],
             ["Combine 12 monthly files", "Every year", "Power Query"]],
      split=4.70,
      so="Manual cleaning wins for small one-offs. Anything that repeats belongs in Power Query."),

    T("2.7.2",
      what="**Automated cleaning in Power Query** defines the steps once and replays them. Its strengths: **the "
           "same result every time** (no person-to-person differences), **nothing forgotten** (every row is "
           "processed), and **a visible record** of each step. The cost is the time to build it the first time.",
      like="It is like a bread factory's mixing machine versus mixing by hand. By hand, each baker's loaf is slightly "
           "different. The machine gives the same loaf every time.",
      how=["Build the cleaning steps once in Power Query.",
           "Save the workbook with the query.",
           "Each new file: drop it in and Refresh."],
      tip="Invest the setup time once for any task that repeats.",
      trap="Judging Power Query only by the first build time, not the time saved later.",
      who="SuperMart Retail",
      case="Three analysts clean " + DIRTY + " by hand and each gets a slightly different result. The Power Query "
           "version gives one answer.",
      steps=["Compare the three manual results: each differs.",
             "Run the Power Query cleaning.",
             "Compare rows and sales with the POS control."],
      table=[["Who cleaned it", "Rows left", "Effect on sales"],
             ["Analyst A (forgot duplicates)", "3,372", "₦13.2m too high"],
             ["Analyst B (forgot text Sales)", "3,312", "₦57.1m missing"],
             ["Analyst C (did everything)", "3,312", "Correct"],
             ["Power Query (every time)", "3,312", "Correct"]],
      split=4.70,
      so="Manual results depend on who did the work. Power Query gives the right answer every time, whoever refreshes."),

    T("2.7.3",
      what="**Reproducible** means anyone can get the same result from the same source. Manual cleaning isn't, "
           "because the clicks and keystrokes aren't recorded. Work out its real cost: **hours per cycle × cycles per "
           "year**, plus the risk that a mistake reaches a decision.",
      like="A recipe kept only in Mama's head isn't reproducible. When she travels, nobody can make her soup. Write it "
           "down, and anyone can.",
      how=["Time one manual cleaning cycle.",
           "Multiply by the number of cycles per year.",
           "Compare with the one-off time to build the query."],
      tip="Put a number on manual effort; it makes the case for automation.",
      trap="Counting only today's time, not the next 12 months.",
      who="SuperMart Retail",
      case="An example to try with your own figures: the monthly POS clean-up takes about 2 hours by hand. Building "
           "the Power Query takes about 3 hours once.",
      steps=["Manual: 2 hours × 12 months.",
             "Power Query: 3 hours once, then a few minutes a month to refresh and check.",
             "Compare the yearly totals."],
      table=[["Approach (example figures)", "First month", "Whole year"],
             ["Manual cleaning", "2 hours", "24 hours"],
             ["Power Query", "3 hours", "About 4 hours"]],
      split=4.70,
      so="Power Query costs more in month one and far less over the year. Do this sum with your real times."),

    T("2.7.4",
      what="An auditor asks: **'Show me how this total was produced.'** With manual cleaning you only have the final "
           "file, and the history is lost. With Power Query you can show the **source file**, every **named step** in "
           "order, the **M code** and the row count after each step.",
      like="It is like a receipt for every purchase. 'I spent ₦50,000' is a claim; a pile of receipts is proof.",
      how=["Name each Applied Step by what it does.",
           "Click through the steps, showing the row count after each.",
           "Show the source path in the Source step."],
      tip="Name steps so they can be read aloud to an auditor.",
      trap="Answering an audit question with 'I cleaned it by hand; trust me'.",
      who="SuperMart Retail",
      case="An auditor asks why reported 2025 sales are lower than the raw export suggests. The analyst opens the "
           "cleaning query and walks through it.",
      steps=["Open the query; show the Source step and file name.",
             "Click 'Removed 60 duplicate lines': show the row count fall.",
             "Show the final row count matches the POS control (3,312)."],
      table=[["Audit question", "Manual answer", "Power Query answer"],
             ["Where did the data come from?", "'An export'", "Source step shows the file"],
             ["Why fewer rows?", "'I removed some'", "60 duplicates, 15 blank rows"],
             ["What was the effect?", "Unknown", "Duplicates added ₦13.2m"]],
      split=4.20,
      so="Every number can be traced back to the raw file, step by step. That is what auditability means."),

    T("2.7.5",
      what="Use **Excel by hand** for one-off fixes, quick looks and small requests. Use **Power Query** for anything "
           "**repeated, large, shared, audited**, or feeding a PivotTable. A simple rule: **if you'd do the same "
           "cleaning a second time, build the query.** Professionals combine both: Power Query for cleaning, Excel for "
           "the final presentation.",
      like="It is like choosing transport: a keke for a short trip around the corner, a bus for the daily commute. "
           "Pick the one that fits the journey.",
      how=["Ask: repeat? size? shared? audited?",
           "Any yes: use Power Query for the cleaning.",
           "Use Excel for final formatting and comments."],
      tip="Clean in Power Query; present and comment in Excel.",
      trap="Typing corrections into a query's output table; the next refresh erases them.",
      who="Multi-Sector",
      case="The analyst plans the tools for a monthly management pack built from three datasets.",
      steps=["List each task in the pack.",
             "Apply the repeat / size / shared test.",
             "Assign Power Query or Excel."],
      table=[["Task", "Tool", "Reason"],
             ["Clean 3,312 POS lines", "Power Query", "Monthly and large"],
             ["Combine Lifecare's monthly files", "Power Query", "Repeats every month"],
             ["Write the 3-line summary", "Excel (typed)", "Needs human judgement"],
             ["Format the final table", "Excel", "Presentation only"]],
      split=4.70,
      so="Each tool does what it is best at: Power Query repeats the cleaning; people add the judgement."),

    T("2.7.6",
      what="The best way to see the difference is to **clean the same file both ways** and compare: the number of "
           "actions, the final row count and total, and what happens when **next month's file** arrives.",
      like="It is like two farmers ploughing identical plots, one with a hoe and one with a tractor. Comparing the "
           "two plots settles the argument better than any speech.",
      how=["Make two copies of the same raw file.",
           "Clean one by hand and one in Power Query.",
           "Compare actions, results and next-month effort."],
      tip="Compare final results against the control totals, not just effort.",
      trap="Comparing only the first month; the real difference shows from month two.",
      who="SuperMart Retail",
      case="The analyst cleans " + DIRTY + " twice: once by hand in Excel, once in Power Query, and counts the "
           "actions.",
      steps=["Do each cleaning task in Excel; count the actions.",
             "Do the same tasks in Power Query; count the steps.",
             "Repeat for next month's file."],
      table=[["Task", "Excel actions", "Power Query steps"],
             ["Blank rows, duplicates", "2", "2"],
             ["Text Sales, text dates", "7", "3"],
             ["Categories, split, trim", "6", "4"],
             ["Next month's file", "Repeat all 15", "Refresh: 0"]],
      split=4.70,
      so="The first month is similar. Every month after, Excel needs 15 actions again; Power Query needs one click."),
]
