from revlib import T

TRY = "ETEH_Capstone_Section_Examples.xlsx  ›  sheet 'P.7 Executive Defence'"
SOL1 = "03_Instructor_Projects  ›  Capstone1_SuperMart_Stores_SOLUTION.xlsx  ›  '7 Memo & Defence'"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", "SuperMart Retail · Capstone 1 defence")
    return T(tag, **k)


TOPICS = [
    O("P.7.1",
      what="The **defence** checks that you understand your own work: the **definitions**, the **decisions**, the "
           "**limitations** and the **evidence**. A plain build you can explain beats a beautiful one you can't.",
      like="Like a driving test: the examiner watches how you drive, not how shiny the car is.",
      how=["Know every definition by heart.", "Know why you excluded anything.", "Have the reconciliation ready to show."],
      tip="Panels mostly probe three things: definitions, exclusions and reconciliation.",
      trap="Spending all your preparation on slide design.",
      case="What the panel checks on Capstone 1.",
      steps=["List the panel's areas.", "Write the likely question.", "Name the evidence."],
      table=[["Area", "Likely question", "Evidence"], ["Definition", "How is margin defined?", "Profit ÷ net sales after discount"],
             ["Decisions", "Is it price or mix?", "Rate −1.19 pts vs mix +0.51 pts"],
             ["Reconciliation", "Is the fall real?", "₦733,215,255 = ledger, ₦0 difference"],
             ["Limitations", "Biggest weakness?", "One year of comparison; no competitor prices"]],
      so="If you can answer these four, you can defend the work.",
      tryit=SOL1,
      split=4.20),

    O("P.7.2",
      what="Plan the ten minutes: **2** for the question and headline, **4** for the evidence and a live demo, **2** "
           "for the recommendation, **2** for limitations and next steps. Then take questions.",
      like="Like a good sermon: the point first, the proof next, the call to action last.",
      how=["Open with the headline number.", "Show the evidence live.", "Close with the ask and the limits."],
      tip="Rehearse with a timer. Most people overrun the evidence section.",
      trap="Opening with how you cleaned the data.",
      case="The Capstone 1 defence plan.",
      steps=["Write one line per segment.", "Time each segment.", "Check the total = 10."],
      table=[["Segment", "Minutes", "Content"], ["Question and headline", "2", "Margin fell 0.7 pts while sales grew 20%"],
             ["Evidence + live demo", "4", "Slicer → all visuals; exception list; reconciliation"],
             ["Recommendation", "2", "20% discount cap; owner, date, effect"],
             ["Limitations, next steps", "2", "Calendar-year view; next: price test"]],
      so="Ten minutes, with the decision clearly in the middle.",
      split=4.20),

    O("P.7.3",
      what="Demonstrate only **three interactions**: change a slicer and watch every visual respond, open the "
           "exception list, and show the reconciliation cell. **Practise until it's automatic**, and keep screenshots as backup.",
      like="Like a cooking demo: three dishes you've made a hundred times, not a new recipe live.",
      how=["Click a slicer: everything updates.", "Open the exception list.", "Show the reconciliation = 0."],
      tip="Keep screenshots ready in case the file or projector fails.",
      trap="Clicking around to 'explore' in front of the panel.",
      case="The Capstone 1 live demo.",
      steps=["Select one store.", "Show its exceptions.", "Show the reconciliation cell."],
      table=[["Demo step", "What it proves", "Time"], ["Slicer on Abuja - Wuse", "All visuals are connected", "30 s"],
             ["Loss-making stores list", "The exceptions are real (5 stores)", "30 s"],
             ["Reconciliation cell", "₦733,215,255 matches the ledger", "30 s"]],
      so="Ninety seconds proves the whole build works.",
      split=4.20),

    O("P.7.4",
      what="Prepare answers to the **ten predictable questions**: source and date, definition, exclusions, "
           "reconciliation, sample size, assumptions, biggest limitation, what you'd change, refresh and owner.",
      like="Like revising past questions before an exam: the same questions keep coming back.",
      how=["Write the ten questions.", "Write a one-line answer with a number.", "Rehearse aloud."],
      tip="Each answer should point to a sheet or cell you can open.",
      trap="Defending how the chart looks when asked about the measure.",
      case="Prepared answers for Capstone 1.",
      steps=["Draft the answers.", "Check each number.", "Rehearse with a friend."],
      table=[["Question", "Prepared answer"], ["Source and date?", "POS ledger export, 05-Jan-2026, Finance"],
             ["What did you exclude?", "Nothing from FY totals; 15 blank + 60 duplicate rows from store returns"],
             ["Is the sample big enough?", "3,312 FY2025 lines: all transactions, not a sample"],
             ["Q4 seasonality?", "Q4 2024 margin was 16.2% vs 9.8% in 2025"],
             ["Who owns it?", "Commercial Analytics; README lists the contact"]],
      so="With answers prepared, questions become a chance to show your evidence.",
      split=4.90),

    O("P.7.5",
      what="When challenged: **restate** the question, answer with **evidence**, and **admit** what's genuinely "
           "uncertain. 'Let me show you the reconciliation' beats any argument.",
      like="Like a referee checking the replay instead of arguing with the player.",
      how=["Restate the challenge calmly.", "Show the evidence on screen.", "Admit the limit and offer a next step."],
      tip="Admitting a real limit builds trust; bluffing destroys it.",
      trap="Defending a number you can't trace.",
      case="A clinician says Lifecare's length of stay is understated (example).",
      steps=["Show the definition.", "Show what was excluded.", "Offer a re-run."],
      table=[["Step", "Response"], ["Restate", "'You think LOS is too low because some stays are missing?'"],
             ["Evidence", "70 of 11,026 rows had no discharge date and were excluded"],
             ["Admit", "Those stays can't be measured yet"], ["Next step", "Re-run once discharges are entered"]],
      so="The challenge ends with more trust, not less.",
      who="Lifecare Hospital · Capstone 3 defence",
      split=4.90),

    O("P.7.6",
      what="Translate every technical choice into a **business consequence**. Panels reward 'what it means for you', "
           "not technical vocabulary.",
      like="Like a mechanic saying 'your brakes will fail in a month', not 'your pads are 2 mm thick'.",
      how=["Write the technical statement.", "Ask 'so what for the business?'", "Say the consequence."],
      tip="Test your sentence on someone outside analytics.",
      trap="Saying 'I used a star schema' and stopping there.",
      case="Capstone translations.",
      steps=["List three technical choices.", "Translate each.", "Practise saying it."],
      table=[["Technical statement", "Business translation"], ["Star schema with DAX measures", "New months drop in and every number updates itself"],
             ["Rate/mix decomposition", "The fall came from deeper discounts, not from selling different things"],
             ["95% interval for readmissions", "Only Oncology is clearly above the hospital rate"]],
      so="The panel understands the consequence, so it can decide.",
      split=4.90),

    O("P.7.7",
      what="End with a **specific ask**: what to approve, **who** decides, **by when**, and the **expected effect**. "
           "Ending with 'Any questions?' wastes the moment the panel is ready to decide.",
      like="Like a sales pitch that ends with 'Can we sign today?', not 'So… yeah'.",
      how=["Name the decision.", "Name the owner and date.", "State the effect as a range."],
      tip="Use the same sentence as your written recommendation.",
      trap="A vague close: 'We should look at discounts.'",
      case="The Capstone 1 close.",
      steps=["Write the ask.", "Add the effect and the risk control.", "Practise it word for word."],
      table=[["Element", "Capstone 1"], ["Decision", "Cap discounts at 20% from Q1 FY2026"], ["Owner", "Commercial Director"],
             ["How", "Pilot in 3 stores for 8 weeks"], ["Expected effect", "₦59.6m–₦71.5m a year (estimate)"],
             ["Risk control", "Stop if volume falls more than 50%"]],
      so="One clear sentence the panel can say yes to.",
      tryit=SOL1,
      split=4.90),

    O("P.7.8",
      what="Within a day, write down what the panel **probed**, what you **couldn't answer**, and what you'll "
           "**change**. That note is your fastest way to improve, and it shows you can review your own work.",
      like="Like a football team watching the match video the next morning.",
      how=["List the questions asked.", "Mark the weak answers.", "Turn each into a fix."],
      tip="Put the fix into the README or template, so it's permanent.",
      trap="Relying on memory a week later.",
      case="A Capstone 1 reflection note.",
      steps=["Write the note within 24 hours.", "Pick one fix.", "Apply it."],
      table=[["Reflection item", "Example"], ["Panel probed", "The definition of discount depth"],
             ["Couldn't answer well", "How depth is weighted (by sales)"], ["Fix", "Add the definition to the README"],
             ["Benefit", "The next panel finds it written down"]],
      so="Each defence makes the next one stronger.",
      split=4.90),
]
