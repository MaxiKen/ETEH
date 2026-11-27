from revlib import T

TRY = "ETEH_Capstone_Section_Examples.xlsx  ›  sheet 'P.1 Project Lifecycle'"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", "SuperMart Retail · Capstone 1")
    return T(tag, **k)


TOPICS = [
    O("P.1.1",
      what="A **project lifecycle** is a fixed order of stages that every analysis follows: **scope, acquire, clean, "
           "model, visualise, verify, hand over**. Each stage produces something you can show, so others can check "
           "and repeat the work.",
      like="Like building a house: plan, foundation, walls, roof, inspection, keys. Skip the foundation and the walls crack.",
      how=["Agree the question before touching data.", "Finish each stage with an output you can show.", "Don't publish until verification passes."],
      tip="The process is part of the deliverable, not just the final chart.",
      trap="Opening the data and making charts before anyone agrees the question.",
      case="Capstone 1: why did FY2025 sales grow while margin fell? Each stage has one output.",
      steps=["List the seven stages.", "Name the output of each.", "Tick each output off before moving on."],
      table=[["Stage", "Output in Capstone 1"], ["1 Scope", "Signed one-page brief"], ["2 Acquire", "Provenance table for 3 files"],
             ["3 Clean", "Cleaning log: 3,387 → 3,312 rows"], ["4–5 Model, visualise", "Measures + one-screen dashboard"],
             ["6–7 Verify, hand over", "Reconciliation PASS + 00_README"]],
      so="Every stage leaves evidence, so a reviewer can follow the work from question to answer.",
      split=4.90),

    O("P.1.2",
      what="**Scoping** turns a vague request into a clear question. Write down the **decision** it supports, the "
           "**measures**, the **dimensions** (ways to slice), the **period**, the **audience** and the **deadline**, "
           "then get it **signed off**.",
      like="Like agreeing the menu, guest count and budget with the client before cooking for an owambe.",
      how=["Write the decision and the question.", "List measures, dimensions and period.", "Get the requester to sign it."],
      tip="If the question has no decision behind it, ask what will change when you answer it.",
      trap="Starting with a favourite chart or tool instead of the question.",
      case="The one-page brief for Capstone 1.",
      steps=["Fill in each brief item.", "Check it with the requester.", "Get two sign-offs."],
      table=[["Brief item", "Capstone 1 entry"], ["Decision", "Pricing and discount actions before Q1 2026"],
             ["Question", "Why did sales grow 20.4% while margin fell from 13.4% to 12.7%?"],
             ["Period", "FY2024 vs FY2025 (calendar years)"], ["Audience / deadline", "Board; draft in 2 weeks, final in 3"],
             ["Sign-off", "Commercial Director + Head of Finance"]],
      so="One page fixes the question, so nobody argues later about what was asked.",
      split=4.20),

    O("P.1.3",
      what="**Acquisition** means getting the data and recording where it came from (**provenance**): source "
           "system, extraction date, row count, owner and storage place. Keep the first copy **read-only** in a "
           "01_RAW folder.",
      like="Like keeping the receipt when you buy a phone: it proves where it came from and when.",
      how=["Copy each file into 01_RAW.", "Record source, date, rows and owner.", "Never edit the RAW copy."],
      tip="Write the row count down at once. It becomes your first control total.",
      trap="Analysing the first file someone emails you without noting where it came from.",
      case="Three files for Capstone 1, recorded before anything else.",
      steps=["Copy to 01_RAW.", "Count the rows.", "Fill in the provenance table."],
      table=[["File", "Source", "Rows", "Owner"], ["SuperMart_Stores_Sales_2022_2025.csv", "POS ledger export", "9,994", "Finance"],
             ["SuperMart_Stores_Directory.csv", "Store master", "12", "Operations"],
             ["SuperMart_Sales_2025_DIRTY.csv", "Manual store returns", "3,387", "Store admin team"]],
      so="Every file can be traced to a source, a date and an owner.",
      split=4.20),

    O("P.1.4",
      what="**Cleaning** fixes problems in the data using named, repeatable steps (Power Query). Count the rows "
           "**before and after** each step, and check the final count against an independent **control total**.",
      like="Like counting the tomatoes before and after sorting out the bad ones: the numbers must add up.",
      how=["Write down the rows received.", "Record rows removed at each step.", "Compare the final count with the control total."],
      tip="Log why each row was removed, not just how many.",
      trap="Deleting odd-looking rows without recording them.",
      case="Cleaning the dirty FY2025 store returns.",
      steps=["Start: 3,387 rows.", "Remove spacer rows, then exact duplicates.", "Compare with the ledger's FY2025 lines."],
      table=[["Step", "Rows removed", "Rows left"], ["Rows received", "0", "3,387"], ["Remove blank spacer rows", "15", "3,372"],
             ["Remove exact duplicates", "60", "3,312"], ["Ledger control total", "n/a", "3,312: difference 0"]],
      so="The counts reconcile and sales match the ledger to ₦0. Gate 3 passes.",
      split=4.70),

    O("P.1.5",
      tryit="ETEH_Capstone_Section_Examples.xlsx  ›  sheet 'P.3 Capstone1 SuperMart'",
      what="**Modelling** builds the layer that does the maths: a **star schema** (one fact table of sales linked "
           "to dimension tables like Store and Date), a marked date table, and **explicit measures** written in DAX.",
      like="Like a well-arranged kitchen: ingredients in one place, recipes in another, so any cook gets the same dish.",
      how=["Load the fact table and dimensions.", "Link many-to-one on keys (Store_ID, Date).", "Write each measure once in DAX."],
      tip="State the grain of the fact table: here, one row per sales line.",
      trap="Adding relationships until the totals 'look right'.",
      case="Capstone 1 measures and their FY2025 values.",
      steps=["Build FactSales + DimStore + DimDate.", "Write the measures.", "Check each against the ledger."],
      table=[["Measure", "DAX", "FY2025"], ["Revenue", "SUM(FactSales[Sales])", "₦733,215,255"],
             ["Gross Margin %", "DIVIDE([Profit],[Revenue])", "12.74%"], ["Orders", "DISTINCTCOUNT(…[Order_ID])", "1,687"],
             ["Basket Value", "DIVIDE([Revenue],[Orders])", "₦434,627"]],
      so="Every measure is defined once and matches the ledger, so every visual agrees.",
      split=4.20),

    O("P.1.6",
      what="The **dashboard** shows the answers: KPI cards, charts chosen for each question, and slicers connected "
           "to every visual. Map **each visual to a question** in the brief. If a visual answers nothing, delete it.",
      like="Like a car dashboard: speed, fuel, warnings. No decorations the driver doesn't need.",
      how=["List the brief's questions.", "Choose one visual per question.", "Connect slicers to every visual."],
      tip="Keep it to one screen.",
      trap="Adding charts because they look impressive.",
      case="Visual-to-question map for the Capstone 1 dashboard.",
      steps=["Write each question.", "Pick the visual.", "Test the slicers."],
      table=[["Visual", "Question it answers"], ["KPI cards", "Revenue, margin, basket value, growth vs last year"],
             ["Category × month columns", "Where did growth come from?"], ["Store ranking with 100% line", "Which stores missed target?"],
             ["Margin by quarter", "When did margin fall?"], ["Zone, Category, Segment slicers", "Connected to every visual"]],
      so="Every visual earns its place by answering a question in the brief.",
      split=4.90),

    O("P.1.7",
      what="**Verification** proves the numbers before anyone sees them: reconcile totals to **control totals**, "
           "test every slicer, and open the file on a **second machine**. Nothing is published until this passes.",
      like="Like a pharmacist double-checking a prescription before handing it over.",
      how=["Get control totals from the source owner.", "Difference = workbook − control; must be 0.", "Test slicers and a second machine."],
      tip="Put the reconciliation cells in the workbook so anyone can see them.",
      trap="Assuming a dashboard is right because it looks consistent.",
      case="Reconciliation tests for Capstone 1, as live formulas.",
      steps=["Enter the control totals.", "Calculate the workbook values.", "Show PASS or FAIL."],
      table=[["Test", "Control total", "Workbook", "Result"], ["SuperMart lines (all years)", "9,994", "9,994", "PASS"],
             ["FY2025 sales", "₦733,215,255", "₦733,215,255", "PASS"], ["FY2025 profit", "₦93,439,242", "₦93,439,242", "PASS"],
             ["FY2025 quantity", "12,476", "12,476", "PASS"]],
      so="Every difference is 0, so the numbers can be published.",
      split=4.20),

    O("P.1.8",
      what="**Handover** means someone else can run and maintain the work without you. Give them the workbook, a "
           "**00_README** sheet, the cleaning log, measure definitions, refresh steps, known limits and a contact.",
      like="Like leaving clear notes for the next shift nurse: what's done, what to watch, who to call.",
      how=["Write the 00_README.", "Include refresh steps and limits.", "Ask a colleague to refresh it unaided."],
      tip="If a colleague can refresh it without calling you, the handover worked.",
      trap="A workbook that only its author can refresh.",
      case="The Capstone 1 00_README.",
      steps=["Fill in each section.", "Test the refresh steps.", "Record the version."],
      table=[["README section", "Content"], ["Purpose", "Explain the FY2025 margin decline"],
             ["Refresh", "New CSV in 01_RAW → Refresh All → reconciliation cells = 0"],
             ["Limitations", "Calendar-year view; targets not adjusted for store size or opening date"],
             ["Version / date", "v1.0, 2026-01-20"]],
      so="The README answers the questions a new owner would otherwise have to ask.",
      split=4.90),

    O("P.1.9",
      what="A **quality gate** is a checkpoint with a clear pass rule at the end of each stage. You don't move to "
           "the next stage until the gate passes. Seven stages, seven gates.",
      like="Like WAEC requirements: you can't sit the next level until you pass this one.",
      how=["Write a pass rule for each stage.", "Record the evidence.", "Stop and fix if a gate fails."],
      tip="Find broken joins at Gate 4, not in the final presentation.",
      trap="Checking everything only at the end.",
      case="Gates for Capstone 1.",
      steps=["Name each gate.", "State the pass rule.", "Point to the evidence."],
      table=[["Gate", "Pass rule", "Evidence"], ["1 Scope", "Brief signed", "Signed brief"],
             ["3 Cleaning", "Rows = control total", "3,312 = 3,312"], ["4 Model", "Many-to-one links; date table marked", "Model diagram"],
             ["6 Verification", "Dashboard total = ₦733,215,255", "Reconciliation cell"],
             ["7 Handover", "README, log, measures, refresh steps", "00_README sheet"]],
      so="Problems are caught at the stage that caused them, when they're cheapest to fix.",
      split=4.20),

    O("P.1.10",
      what="**Time-boxing** gives each stage a fixed share of the time. A common split: scoping 10%, acquisition "
           "10%, **cleaning 30%**, modelling 20%, visuals 15%, verification 10%, handover 5%.",
      like="Like a budget for the month: rent, food, transport. Without it, the money is gone by week two.",
      how=["Enter the total hours.", "Share × hours for each stage.", "Track actual against plan."],
      tip="Cleaning and verification are the stages people underestimate most.",
      trap="Spending most of the time on the dashboard's look.",
      case="A 40-hour capstone. Change the hours and the plan updates.",
      steps=["Type 40 in the hours cell.", "Each stage: share × 40.", "Check the total is 40."],
      table=[["Stage", "Share", "Hours"], ["Scoping + acquisition", "20%", "8"], ["Cleaning", "30%", "12"],
             ["Modelling", "20%", "8"], ["Visualisation", "15%", "6"], ["Verification + handover", "15%", "6"]],
      so="Cleaning gets the most time because that's where credibility is won or lost.",
      split=4.70),

    O("P.1.11",
      what="Keep stakeholders informed at three moments: a **scope** for sign-off, a **draft preview** mid-way, and "
           "a **final pack** with the QA note. Surprises at the end destroy trust.",
      like="Like a tailor's fitting: you try it on half-way, not on the day of the wedding.",
      how=["Week 1: send the scope.", "Mid-way: show a draft.", "End: final pack + QA note."],
      tip="Agree the control totals with Finance before you polish anything.",
      trap="Disappearing for three weeks and presenting a finished surprise.",
      case="The Capstone 1 stakeholder plan.",
      steps=["Plan each moment.", "Name the person.", "State the purpose."],
      table=[["When", "What", "Purpose"], ["Week 1", "One-page scope to Commercial Director", "Sign-off on question and period"],
             ["Week 2", "Draft dashboard to Finance", "Agree control totals before polish"],
             ["Week 3", "Final pack + QA note to board", "Decision on Q4 discount policy"]],
      so="Each person sees the work at the point where their input matters.",
      split=4.20),

    O("P.1.12",
      what="After delivery, hold a short **post-project review**: did it change a decision, what took longer, what "
           "to automate next time, and what data problem to raise with the owner. Write the answers down.",
      like="Like a football team watching the match video: what worked, what didn't, what to practise.",
      how=["Ask the four questions.", "Record honest answers.", "Turn each into an action."],
      tip="The review is where one project improves the next.",
      trap="Moving on without learning anything.",
      case="Capstone 1 review (training scenario).",
      steps=["Ask each question.", "Write the answer.", "Assign an action."],
      table=[["Question", "Answer recorded"], ["Did it change a decision?", "Board capped Furniture discounts at 20% for Q1 2026"],
             ["What took longer?", "Cleaning the 331 mixed-format dates"],
             ["Automate next time", "Monthly folder import instead of manual copy"],
             ["Raise with data owner", "Stores typing '₦' into the Sales column"]],
      so="Four short answers turn into concrete improvements for the next project.",
      split=4.90),
]
