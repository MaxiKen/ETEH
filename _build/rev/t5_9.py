from revlib import T

TRY = "ETEH_Capstone_Section_Examples.xlsx  ›  sheet 'P.9 Interviews & Roadmap'"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", "Interview practice · pack datasets")
    return T(tag, **k)


TOPICS = [
    O("P.9.1",
      what="Most analytics interviews use four formats: a **take-home case**, a **live Excel test**, a **SQL screen** "
           "and a **behavioural interview**. Prepare for each separately.",
      like="Like WAEC subjects: you can't pass Maths by revising English.",
      how=["Learn what each format tests.", "Practise with timed exercises.", "Prepare stories for behavioural questions."],
      tip="Ask the recruiter which formats to expect. Most will tell you.",
      trap="Preparing only for the conversation and not the test.",
      case="How the bootcamp maps to each format.",
      steps=["Match each format to practice material.", "Time yourself.", "Review your mistakes."],
      table=[["Format", "Tests", "Practise with"], ["Take-home case", "End-to-end analysis", "The three capstones"],
             ["Live Excel test", "Cleaning, PivotTables, speed", "Weekly assignments; P.9 sheet"],
             ["SQL screen", "Filters, joins, windows", "supermart_online.db queries"],
             ["Behavioural", "Judgement, errors, conflict", "Your defence and STAR stories"]],
      so="Everything in this bootcamp doubles as interview practice.",
      split=4.20),

    O("P.9.2",
      what="A **live Excel test** usually means: clean a file, build a PivotTable, answer a few questions and explain "
           "one number. Success comes from **process**, and from **saying what you're doing** aloud.",
      like="Like a maths exam where showing your working earns most of the marks.",
      how=["Ctrl+T, then check data types.", "Answer each question with an explicit formula or Pivot.", "Explain one number in plain words."],
      tip="Say your filters aloud: 'filtering to 2025 and the zone'.",
      trap="Working in silence. Silence looks like uncertainty.",
      case="Three test questions on tSM (FY2025).",
      steps=["SUMIFS for the zone.", "PivotTable for margin by store.", "COUNTIFS ÷ COUNTIFS for the share."],
      table=[["Question", "Answer"], ["Total South-South sales", "₦213,082,904"],
             ["Store with the highest margin", "Owerri - New Owerri (31.0%)"],
             ["Share of lines with 40%+ discount", "11.4%"], ["Explain one number", "Those lines are mostly loss-making"]],
      so="Correct answers plus a clear explanation is what passes the test.",
      split=4.90),

    O("P.9.3",
      what="**SQL screens** usually rise in difficulty: filter and aggregate, then a **JOIN** (say the cardinality "
           "first), then a **window function**. Interviewers care about your reasoning on keys as much as the syntax.",
      like="Like a driving test that goes from car park to main road to motorway.",
      how=["Easy: WHERE + GROUP BY.", "Medium: JOIN, saying 'many transactions to one product'.", "Hard: RANK() OVER (PARTITION BY …)."],
      tip="Before writing a join, say which key links the tables and how many rows match.",
      trap="Joining on the wrong key and not noticing the total jumped.",
      case="Three questions on supermart_online.db.",
      steps=["Revenue by payment method (delivered).", "Top brand by net revenue.", "Top seller per zone."],
      table=[["Level", "Question", "Top answer"], ["Easy", "Delivered revenue by payment method", "Paystack: 2,325 orders, ₦301.0m"],
             ["Medium", "Top brand (many-to-one JOIN)", "HomeStyle: 423 orders, ₦214.6m"],
             ["Hard", "Best seller in each zone (RANK)", "South-West: Jos Premium, ₦43.9m"]],
      so="Explaining each clause matters as much as getting the answer.",
      split=4.20),

    O("P.9.4",
      what="'How do you approach a new dataset?' The answer that lands is a **sequence**: understand the question, "
           "**profile** the data, check quality, build, **verify** against controls, communicate. Use a real example.",
      like="Like a doctor: ask about symptoms and run tests before prescribing.",
      how=["Question first.", "Profile: rows, blanks, distinct values.", "Verify, then communicate."],
      tip="Give one specific thing profiling revealed and what you did about it.",
      trap="Jumping straight to 'I'd make some charts'.",
      case="Profiling the dirty Lifecare 2023 extract.",
      steps=["Count rows.", "Count blanks and odd values.", "Decide how to handle each."],
      table=[["Profile finding", "Count", "Decision"], ["Rows", "11,026", "Reconcile to the 10,918 control"],
             ["Duplicates / negative bills", "90 / 18", "Remove both: 11,026 → 10,918"], ["Missing discharge date", "70", "Exclude from LOS; disclose"],
             ["Ward name variants", "661", "Map to 4 wards"]],
      so="A specific, ordered answer shows real experience.",
      who="Lifecare Hospital · dirty 2023 extract",
      split=4.70),

    O("P.9.5",
      what="Answer behavioural questions with **STAR**: **S**ituation, **T**ask, **A**ction, **R**esult. Your capstone "
           "work is the material. 'Tell me about a time you found an error' needs a real error.",
      like="Like telling a short story with a beginning, middle and end, and a lesson.",
      how=["Situation and task in one sentence each.", "Action: what you did.", "Result: what changed, with a number."],
      tip="Keep it under two minutes, most of it on the Action.",
      trap="A story with no result, or one where 'we' did everything.",
      case="The ₦7.19bn wrong join (SuperMart Online SQL project).",
      steps=["Write each STAR line.", "Check the numbers.", "Practise aloud."],
      table=[["STAR", "Answer"], ["Situation", "An AI-written zone query returned ₦7.19bn (56,725 rows)"],
             ["Task", "Verify it before it went into the deck"],
             ["Action", "Checked rows (should be 10,430) and the ₦1.32bn control; found the join on customer_id"],
             ["Result", "Rewrote on delivery_id; control check added to every SQL job"]],
      so="A real error, caught by your own check, is one of the strongest interview stories.",
      who="SuperMart Online · SQL project",
      split=4.90),

    O("P.9.6",
      what="**Power BI** is the natural next step. Your Data Model, DAX and star-schema skills transfer directly; "
           "the main new thing is the visual layer and publishing.",
      like="Like moving from a manual to an automatic car: you already know the roads and the rules.",
      how=["Install Power BI Desktop.", "Rebuild Capstone 1 with the same schema and DAX.", "Learn visuals and publishing."],
      tip="Rebuild something you know well, so only the tool is new.",
      trap="Starting with a new dataset and a new tool at the same time.",
      case="Rebuilding Capstone 1 in Power BI.",
      steps=["Load the same tables.", "Copy the measures.", "Rebuild the dashboard."],
      table=[["Item", "Detail"], ["Transfers directly", "Star schema, DAX measures, chart choices"],
             ["New to learn", "Power BI visuals, report pages, publishing"],
             ["Check", "Revenue still = ₦733,215,255"], ["Time", "About 2–4 weeks"]],
      so="You're closer to Power BI than you think.",
      who="Roadmap · next tools",
      split=4.90),

    O("P.9.7",
      what="Grow your **SQL** from SELECT and JOIN to **CTEs**, **window functions** and **indexes**, then move to a "
           "server database such as PostgreSQL. Practise on real databases with messy keys.",
      like="Like learning to cook: from jollof for two to cooking for a party of 200.",
      how=["Window functions and CTEs on the pack databases.", "Learn indexes and query plans.", "Then try PostgreSQL."],
      tip="Re-run your queries and check the totals against controls every time.",
      trap="Only practising on tiny, tidy tutorial tables.",
      case="Practice databases already in the pack.",
      steps=["Rewrite P.9 queries with CTEs.", "Add an index; compare speed.", "Move one database to PostgreSQL."],
      table=[["Database", "Rows to practise on"], ["supermart_online.db", "11,000 transactions, 5 tables"],
             ["lifecare_hospital.db", "54,860 admissions, 5 tables"], ["Time", "About 4–6 weeks"]],
      so="The real practice material is already on your laptop.",
      who="Roadmap · next tools",
      split=4.90),

    O("P.9.8",
      what="**Python** adds scale, automation and deeper statistics: **pandas** for tables, **matplotlib** for charts. "
           "Your most valuable asset isn't syntax; it's the discipline of profiling, verifying and reporting honestly.",
      like="Like a good driver learning a new car: the road sense comes with you.",
      how=["Learn pandas basics.", "Re-create a cleaning log you already know.", "Check results match the Excel version."],
      tip="Your first Python project should reproduce a number you already trust.",
      trap="Learning syntax for months without solving a real problem.",
      case="A first Python project.",
      steps=["Load SuperMart_Sales_2025_DIRTY.csv.", "Re-create the cleaning log.", "Plot the margin chart."],
      table=[["Item", "Detail"], ["pandas", "Cleaning and summarising tables"], ["matplotlib", "Charts like the margin trend"],
             ["Check", "Same totals as the Excel cleaning log"], ["Time", "About 6–8 weeks"]],
      so="If you can define, reconcile and explain, Python is just a new tool.",
      who="Roadmap · next tools",
      split=4.90),
]
