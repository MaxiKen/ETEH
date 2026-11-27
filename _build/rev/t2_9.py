from revlib import T

W, LC, MS = "SuperMart Retail", "Lifecare Hospital", "Multi-Sector"
TRY = "ETEH_Week2_Section_Examples.xlsx  ›  sheet '2.9 M Language'"
LDIRTY = "Lifecare_Admissions_2023_DIRTY.csv (Week2_Data_Files)"


def M(tag, **k):
    k.setdefault("tryit", TRY)
    return T(tag, **k)


TOPICS = [
    M("2.9.1",
      what="**M** is the language Power Query writes behind the scenes. Every click you make (remove a column, change "
           "a type, filter rows) is saved as **one line of M**. The **Applied Steps** list is simply a friendly view "
           "of that M script.",
      like="M is like the recipe card behind a dish. You see the meal (your clean table); the card lists every step "
           "the cook took, in order, so anyone can make it again.",
      how=["Do your steps with the ribbon as usual.", "Turn on View > Formula Bar to see each step's M.",
           "Open Home > Advanced Editor to see the whole script."],
      tip="You don't need to write M from scratch; learn to read it first.",
      trap="Thinking the Applied Steps and the M are different things; they are the same query.",
      who=LC,
      case="The first query on the 2.9 sheet loads " + LDIRTY + " and fixes patient names typed like "
           "'chuKwuemEKa f. odili'.",
      steps=["Load the CSV and promote headers.", "Transform Patient_Name with Text.Proper.",
             "Open the Advanced Editor and match each step to its M line."],
      table=[["Applied Step", "M function behind it"], ["Source", "Csv.Document(File.Contents(…))"],
             ["Promoted Headers", "Table.PromoteHeaders(Source, …)"],
             ["Proper Names", "Table.TransformColumns(…, Text.Proper)"],
             ["Result", "'chuKwuemEKa f. odili' → 'Chukwuemeka F. Odili'"]],
      so="Three clicks made three lines of M. Reading them tells you exactly what the query does.",
      split=4.90),

    M("2.9.2",
      what="M runs inside the **Power Query engine**, not in worksheet cells, so it can handle millions of rows. "
           "M is **case-sensitive**: Text.Proper works but text.proper fails. Column names and text values are also "
           "case-sensitive: \"Ruby Ward\" is not \"RUBY WARD\".",
      like="M is like a strict gateman with a guest list: 'Ade' is on the list, so 'ADE' and 'ade' are turned away "
           "until someone writes them the same way.",
      how=["Copy function names exactly: capital letters matter.", "Refer to columns exactly as named: [Ward], not [ward].",
           "Standardise text values before you filter or group."],
      tip="Clean text (trim, fix case) before filtering in M.",
      trap="Filtering on \"Ruby Ward\" and silently missing 'RUBY WARD' and 'Ruby  Ward'.",
      who=LC,
      case="After removing duplicates, " + LDIRTY + " has 10,963 rows. The Ward column has 16 spellings of 4 wards.",
      steps=["Filter Ward = \"Ruby Ward\" on the raw column and note the count.",
             "Apply the WardMap step (trim, lower case, fix 'Saphire').", "Filter again and compare."],
      table=[["Test", "Result"], ["Ward spellings in the raw file", "16"],
             ["Rows exactly \"Ruby Ward\" (raw)", "2,635"], ["Ruby Ward rows after WardMap", "2,811"],
             ["text.proper instead of Text.Proper", "Error: name not recognised"]],
      so="The exact filter missed 176 Ruby Ward patients. Case and spacing matter in M.",
      split=4.90),

    M("2.9.3",
      what="M has a few building blocks. **let … in** wraps the steps. **each** means 'for every row'. **[Column]** "
           "is that row's value. **//** starts a comment. Text uses \"double quotes\", lists use **{ }**, records use "
           "**[ ]**, and step names with spaces are written **#\"Like This\"**.",
      like="M syntax is like the layout of a recipe: a title (let), numbered steps, notes in the margin (//), and the "
           "finished dish at the end (in).",
      how=["Read let … in as 'here are the steps … here is the answer'.",
           "Read each [Column] as 'for every row, look at this column'.",
           "Use // comments to explain business rules."],
      tip="Write a // comment above any step with a business rule.",
      trap="Leaving out the comma between steps, or adding one after the last step.",
      who="Crestview University",
      case="Crestview (6,607 students) wants an attendance band added in M.",
      steps=["Table.AddColumn(Typed, \"Attendance_Band2\", each …)",
             "each if [Attendance_Pct] >= 90 then \"90%+\" else if [Attendance_Pct] >= 75 then \"75-89%\" else "
             "\"Below 75%\"",
             "Close the step and check the new column."],
      table=[["Attendance_Band2", "Students"], ["90%+", "1,746"], ["75-89%", "2,465"], ["Below 75%", "2,396"],
             ["Total", "6,607"]],
      so="One line with each and [Attendance_Pct] labelled all 6,607 students, and the bands add up."),

    M("2.9.4",
      what="Every M column should have a **type**: **text**, **Int64.Type** (whole number), **number** (decimal), "
           "**date**, **logical** (true/false) and so on. A missing value is **null**, which is not 0 and not \"\". "
           "Types are set with **Table.TransformColumnTypes**.",
      like="Types are like labelled containers in a kitchen: rice in the rice bin, oil in the oil can. Put the wrong "
           "thing in and the recipe fails.",
      how=["Set types on purpose; don't rely on the automatic step.",
           "Give a culture (e.g. \"en-US\") when numbers contain commas.",
           "Check columns with Column quality for errors and nulls."],
      tip="Set types explicitly, with a culture for numbers and dates.",
      trap="Treating null as 0; it means 'unknown', not zero.",
      who=LC,
      case="In " + LDIRTY + ", 441 Billing_Amount values arrive with commas, like '2,507,101.35', and 70 rows have no "
           "discharge date.",
      steps=["Table.TransformColumnTypes(…, {{\"Billing_Amount\", type number}}, \"en-US\")",
             "Type Admission_Date and Discharge_Date as date; Age as Int64.Type.",
             "Check how many Discharge_Date values are null."],
      table=[["Column", "M type", "Result"], ["Billing_Amount", "type number (en-US)", "441 comma values parsed"],
             ["Admission_Date", "type date", "All valid"], ["Discharge_Date", "type date", "70 null"],
             ["Age", "Int64.Type", "All whole numbers"]],
      so="With the right types and culture, every bill becomes a number, and missing dates stay honestly null.",
      split=4.20),

    M("2.9.5",
      what="A query is one **let** expression. Each step is a **name = expression**, and usually uses the step "
           "before it. That is why deleting a step in the middle can break every step after it: they still refer "
           "to the deleted name.",
      like="The step chain is like a relay race: each runner takes the baton from the one before. Remove a runner "
           "and the next one has nobody to take it from.",
      how=["Read each step's formula to see which step it uses.",
           "Before deleting a step, check what refers to it.",
           "If a later step breaks, point it at the step before the deleted one."],
      tip="Rename steps clearly (Dedup, WardMap) so the chain reads like a story.",
      trap="Deleting a middle step and leaving the next step pointing at a name that no longer exists.",
      who=LC,
      case="The full Lifecare cleaning query on the 2.9 sheet (example A) runs from Source to Clean.",
      steps=["Click each step and watch the row count.", "Note that Names uses Dedup.",
             "Delete Dedup and read the error on Names."],
      table=[["Step", "Uses", "Rows after"], ["Promoted", "Source", "11,026"], ["Dedup", "Promoted", "10,963"],
             ["Clean (exceptions removed)", "Exception", "10,838"],
             ["Names, if Dedup is deleted", "Dedup (gone)", "Error"]],
      so="Each step builds on the last. Deleting Dedup breaks Names, because Names still refers to it.",
      split=4.20),

    M("2.9.6",
      what="The **Formula Bar** shows the M for whichever step is selected. Turn it on with **View > Formula Bar**. "
           "It is the best place to check what a step does, and to make small edits without redoing clicks.",
      like="The Formula Bar is like reading the label on each jar in a row: click a jar and its label tells you "
           "exactly what is inside.",
      how=["Tick View > Formula Bar.", "Click a step in Applied Steps.", "Read or edit the M, then press Enter."],
      tip="Click through steps with the Formula Bar open to understand someone else's query.",
      trap="Editing a step's formula and breaking the step names that later steps rely on.",
      who=LC,
      case="A new analyst opens the Lifecare cleaning query and wants to know what the Dedup step does.",
      steps=["Turn on the Formula Bar.", "Click Promoted, then Dedup.", "Compare the formula and row counts."],
      table=[["Step selected", "Formula Bar shows", "Rows"], ["Promoted", "= Table.PromoteHeaders(Source, …)", "11,026"],
             ["Dedup", "= Table.Distinct(Promoted)", "10,963"]],
      so="One click showed that Dedup removes exact duplicates: 63 rows.",
      split=4.20),

    M("2.9.7",
      what="**Home > Advanced Editor** shows the **whole query** as one M script. Use it to copy a query to "
           "another file, to make many edits at once, or to paste in a query someone has written for you.",
      like="The Advanced Editor is like seeing a whole recipe book page at once, instead of one step at a time.",
      how=["Create a blank query: Data > Get Data > From Other Sources > Blank Query.",
           "Open Home > Advanced Editor and paste the script.",
           "Click Done; fix any names (like pFolderPath) it reports."],
      tip="Save a copy of the M before big edits.",
      trap="Pasting a query that refers to a parameter or query your file doesn't have.",
      who=LC,
      case="The trainer shares the complete Lifecare cleaning query (example A on the 2.9 sheet).",
      steps=["Create the pFolderPath parameter (as in 2.3).", "Paste example A into a blank query's Advanced Editor.",
             "Click Done and check the final row count."],
      table=[["Check", "Result"], ["Steps in the pasted script", "12"], ["Duplicates removed", "63"],
             ["Exception rows removed", "125"], ["Clean rows", "10,838"]],
      so="One paste rebuilt a 12-step cleaning query. Without pFolderPath it would fail at the first step.",
      split=4.90),

    M("2.9.8",
      what="M functions come in **families**, named by what they work on: **Text.**, **Number.**, **Date.**, "
           "**List.** and **Table.**. Most ribbon buttons are one of these functions.",
      like="Function families are like sections in a hardware shop: plumbing, electrical, paint. Knowing the section "
           "tells you where to look.",
      how=["Guess the family from what you're changing (text, number, date…).",
           "Type the family name and a dot in the Formula Bar to see options.",
           "Check the result in a custom column."],
      tip="Start with the family name; the rest follows.",
      trap="Mixing up families, e.g. using a Text function on a number column.",
      who=MS,
      case="The 2.9 sheet lists one example from each family, using values from the bootcamp files.",
      steps=["Add a blank query to test each expression.", "Type the example in the Formula Bar.",
             "Compare your result with the table."],
      table=[["Family", "Example (our data)", "Result"], ["Text.", "Text.BeforeDelimiter(\"Benin City, Edo\", \",\")", "\"Benin City\""],
             ["Number.", "Number.Round(0.12749, 3)", "0.127"], ["Date.", "Date.EndOfMonth(#date(2024, 2, 10))", "29-Feb-2024"],
             ["List.", "List.Contains({\"m\", \"male\"}, \"male\")", "true"]],
      so="Each family handles one kind of data. Ribbon buttons write these functions for you.",
      split=4.20),

    M("2.9.9",
      what="M's text, number and date functions are like Excel's but **stricter**: **Text.Trim**, **Text.Lower**, "
           "**Text.Proper**, **Number.From**, **Date.Year**, **Duration.Days** and more. Combine them inside "
           "**each** to clean a whole column at once.",
      like="These functions are like tools on a tailor's table: scissors, tape, needle. Each does one small job; "
           "together they make a finished garment.",
      how=["Pick the column to fix.", "Combine functions inside each, e.g. Text.Proper(Text.Trim(_)).",
           "Check the distinct values afterwards."],
      tip="Lower-case the text before comparing it.",
      trap="Comparing 'M' with \"Male\" without standardising case first.",
      who=LC,
      case="Gender in the de-duplicated Lifecare file has 6 spellings; Admission_Type also has 6.",
      steps=["Gender: each if List.Contains({\"m\", \"male\"}, Text.Lower(_)) then \"Male\" else \"Female\".",
             "Admission_Type: lower-case, trim, map 'emerg…' to Emergency.", "Check the distinct values."],
      table=[["Column", "Before", "After"], ["Gender", "6 spellings", "Female 5,510 / Male 5,453"],
             ["Admission_Type", "6 spellings", "Urgent 3,676 / Emergency 3,670 / Elective 3,617"]],
      so="A few text functions turned 12 messy spellings into 5 clean values across 10,963 patients.",
      split=4.20),

    M("2.9.10",
      what="**Table.** functions work on the **whole table**: Table.SelectRows (filter), Table.AddColumn, "
           "Table.Group, Table.Sort, Table.Distinct, Table.NestedJoin (merge) and more. Most Applied Steps are one "
           "of these.",
      like="Table functions are like a factory machine that processes the whole tray at once, not one biscuit at "
           "a time.",
      how=["Add a helper column with Table.AddColumn.", "Summarise with Table.Group.",
           "Add a calculated measure to the grouped table."],
      tip="Group on clean, typed columns only.",
      trap="Grouping before the text is clean, so one store splits into several groups.",
      who=W,
      case="Example C on the 2.9 sheet summarises FY2025 sales by store and month, starting from the clean "
           "sales query.",
      steps=["Table.AddColumn: Month = Date.ToText([Order_Date], \"yyyy-MM\").",
             "Table.Group by Store_Name and Month: sum Sales and Profit, count distinct orders.",
             "Table.AddColumn: Margin = [Profit] / [Sales]."],
      table=[["Check", "Result"], ["Rows after Table.Group", "144 (12 stores × 12 months)"],
             ["Lagos - Ikeja, best month", "2025-12: ₦19,318,023"],
             ["Abuja - Garki, Feb 2025 profit", "−₦86,903"]],
      so="Three table functions turned 3,312 lines into 144 store-month rows ready for a chart.",
      split=4.90),

    M("2.9.11",
      what="Most real queries follow one pattern: **load** the source, **promote** headers, **set types**, add "
           "**business rules** as new columns, **filter** out bad rows, then **group** or load. Once you know it, you "
           "can reuse it on any dataset.",
      like="The pattern is like a standard checklist a pilot uses on every flight: the aircraft changes, the "
           "checklist doesn't.",
      how=["Load and promote headers.", "Set types; add rule columns with Table.AddColumn.",
           "Filter or group, then load."],
      tip="Keep the same step order in every query, so anyone can read it.",
      trap="Hard-coding file paths instead of using a parameter like pFolderPath.",
      who=MS,
      case="The 2.9 sheet applies the pattern to three bootcamp datasets: Lifecare, Crestview and SuperMart.",
      steps=["A: clean the Lifecare 2023 dirty extract.", "B: add the Crestview attendance band and mentor flag.",
             "C: group SuperMart FY2025 sales by store and month."],
      table=[["Dataset", "Rule in M", "Result"], ["Lifecare", "Remove duplicates and exceptions", "11,026 → 10,838 rows"],
             ["Crestview", "Exam < 60 and attendance < 75%", "67 students need a mentor"],
             ["SuperMart", "Group by store and month", "144 rows"]],
      so="The same pattern handled a hospital, a university and a retailer. Only the rules changed.",
      split=4.20),
]
