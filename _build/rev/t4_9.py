from revlib import T

TRY = "ETEH_Week4_Section_Examples.xlsx  ›  sheet '4.9 Slide Presentations'  ·  data in D_Hotel"
HT = "EkoSuites Hotel"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", HT)
    return T(tag, **k)


TOPICS = [
    O("4.9.1",
      what="A short analytical talk follows a fixed path: **headline**, **situation**, the **evidence** (where, "
           "when, who), the **options**, the **recommendation** and **next steps**. Plan it as a storyboard first.",
      like="Like a film trailer: the hook, the story, and the reason to buy a ticket, in a few minutes.",
      how=["Write one title per slide on paper first.", "Give each slide a role and a time.", "Only then build the slides."],
      tip="Storyboard in a worksheet: slide, role, title, proof number, minutes.",
      trap="Opening PowerPoint before you know the message.",
      case="A 10-slide talk on EkoSuites cancellations (36,275 bookings, 2023–24).",
      steps=["List the 10 slide roles.", "Write each title.", "Add minutes and total them."],
      table=[["Slides", "Role", "Minutes"], ["1", "Headline: cancellations cost ₦4.30bn", "0.5"],
             ["2–6", "Situation, where, when, who, signal", "6"], ["7–8", "Option A (deposits), Option B (loyalty)", "2.5"],
             ["9–10", "Recommendation and next steps", "2"]],
      so="Ten slides, 11 minutes, one story from problem to decision.",
      split=4.70),

    O("4.9.2",
      what="**One message per slide**, stated in the **title**, proved by **one number or chart**. If a slide "
           "has two messages, split it into two slides.",
      like="Like road signs: each one tells you one thing, quickly.",
      how=["Write the title as a full-sentence finding.", "Add the one number that proves it.", "Remove everything else."],
      tip="Calculate the proof number with a formula, so the title stays true.",
      trap="A slide titled 'Analysis' with four charts on it.",
      case="Titles and proof numbers for slides 2–6.",
      steps=["Write the finding.", "COUNTIFS for the proof.", "Check they agree."],
      table=[["Slide title", "Proof number"], ["One booking in three is cancelled", "32.8%"], ["OTA bookings cancel most", "36.5%"],
             ["Bookings made 6+ months ahead cancel three times in four", "73.9%"], ["Repeat guests almost never cancel", "1.7%"],
             ["Guests who make special requests cancel less", "20.2% vs 43.2% with none"]],
      so="Each slide makes one clear point, backed by one checkable number.",
      split=4.90),

    O("4.9.3",
      what="Slides are seen from a distance. Use **bigger text** (labels 18pt or more), **one highlighted series**, "
           "**direct labels** instead of legends, and no 3D. A report chart usually needs changes before it works on a slide.",
      like="Like a road sign compared with a book page: big, bold and simple.",
      how=["Increase all font sizes.", "Grey everything; colour the one bar that matters.", "Label values directly on the bars."],
      tip="Test it: stand 3 metres from your screen.",
      trap="Pasting a dense report chart straight onto a slide.",
      case="Prepare the cancel-rate-by-lead-time chart for slide 4 (tLeadBand from 4.3).",
      steps=["Column chart of cancel rate by band.", "Highlight the 180+ days bar.", "Labels 18pt or more; delete the legend."],
      table=[["Booked ahead", "Cancel rate", "Colour"], ["0–7 days", "8.9%", "Grey"], ["31–90 days", "25.3%", "Grey"],
             ["91–180 days", "44.9%", "Grey"], ["180+ days", "73.9%", "Highlight"]],
      so="The audience sees the rising cancellations at once, and the bar you want them to remember.",
      split=4.70),

    O("4.9.4",
      what="Slides hold **headlines and evidence, not paragraphs**. Keep titles under about **15 words**, use at most "
           "about **6 short bullets**, and put the detail in your speaker notes.",
      like="Like a billboard: if drivers can't read it in a few seconds, it's failed.",
      how=["Count title words.", "Cut bullets to 6 or fewer.", "Move explanations into the speaker notes."],
      tip="If you'd read the slide aloud word for word, it has too much text.",
      trap="Full sentences of method on the slide.",
      case="Check the slide 4 title length with a formula.",
      steps=["=LEN(TRIM(t))-LEN(SUBSTITUTE(TRIM(t),\" \",\"\"))+1.", "Compare with 15.", "Shorten if needed."],
      table=[["Check", "Result"], ["Slide 4 title", "Bookings made 6+ months ahead cancel three times in four"],
             ["Word count", "10"], ["Limit", "About 15"], ["Verdict", "Pass"]],
      so="A 10-word finding title can be read in a glance.",
      split=4.90),

    O("4.9.5",
      what="Use **one font family**, **one colour palette**, **one chart style** and **one footer** on every slide. "
           "Keep bright colours for emphasis only. Consistency makes the content stand out.",
      like="Like a school uniform: nobody notices the clothes, so they notice the person.",
      how=["Set fonts and colours in the slide master.", "Use the same chart style throughout.", "Keep accent colours for highlights."],
      tip="Save the palette as a theme so every chart uses it.",
      trap="A new colour scheme on every slide.",
      case="The visual rules for the EkoSuites deck.",
      steps=["Choose the palette.", "Set the slide master.", "Apply to all charts."],
      table=[["Element", "Rule"], ["Main colour", "ETEH navy #0B2545"], ["Second colour", "Teal #13807A"],
             ["Highlight only", "Gold #E0A526"], ["Footer", "Same on every slide"]],
      so="A consistent look makes the deck feel professional and lets the findings stand out.",
      split=4.90),

    O("4.9.6",
      what="A table on a slide holds about **5 rows × 4 columns** at most. Round numbers, highlight the key row, "
           "and move the full table to the appendix.",
      like="Like a menu board: a few choices, easy to read from the queue.",
      how=["Keep only the rows that support the message.", "Round to what matters (%, ₦m).", "Highlight one row; full table in the appendix."],
      tip="If the table needs more than 5 rows, use a bar chart instead.",
      trap="Pasting a 30-row table and saying 'as you can see'.",
      case="Cancel rate by market segment for slide 3.",
      steps=["COUNTIFS per segment.", "Round to one decimal.", "Highlight OTA."],
      table=[["Segment", "Bookings", "Cancel rate"], ["Online Travel Agent", "23,214", "36.5%"],
             ["Walk-in / Travel Agent", "10,528", "29.9%"], ["Corporate", "2,017", "10.9%"],
             ["Airline Crew", "125", "29.6%"], ["Complimentary", "391", "0%"]],
      so="Five rows show the point: OTA is the largest segment and cancels the most.",
      split=4.70),

    O("4.9.7",
      what="Prepare for three kinds of question: about the **data** ('where's this from?'), the **method** ('how "
           "did you calculate it?') and the **implications** ('what would it cost?'). Keep **back-up slides** ready.",
      like="Like a lawyer preparing for cross-examination: know the likely questions in advance.",
      how=["List likely questions of each type.", "Prepare a short answer and a back-up slide for each.", "Rehearse with a colleague."],
      tip="'I'll check and come back to you' is better than guessing.",
      trap="Being surprised by 'how many bookings is that?'",
      case="Prepare for the EkoSuites board's questions.",
      steps=["Data question.", "Method question.", "Implication question."],
      table=[["Question type", "Prepared answer"], ["Data: where from?", "36,275 bookings, 2023–24 (D_Hotel)"],
             ["Method: how calculated?", "Cancelled bookings ÷ all bookings"], ["Implication: who do deposits affect?", "7,510 OTA bookings over 90 days ahead"],
             ["Follow-up: how risky are they?", "60.1% of those were cancelled"]],
      so="With answers ready, questions build confidence instead of doubt.",
      split=4.90),

    O("4.9.8",
      what="Allow about **one minute per slide**, give more time to the evidence, and keep time for questions. "
           "**Rehearse aloud** with a timer: talks almost always run longer than planned.",
      like="Like cooking for guests: plan the timing so everything is ready together.",
      how=["Add minutes per slide in the storyboard.", "Total them and compare with your slot.", "Rehearse twice with a timer."],
      tip="Know which slide to skip if you run late.",
      trap="Rushing the recommendation because time ran out.",
      case="Does the storyboard fit a 12-minute talk slot plus 5 minutes of questions?",
      steps=["=SUM of the minutes column.", "Compare with 12.", "Mark a slide to skip if needed."],
      table=[["Item", "Minutes"], ["Total planned talk time", "11"], ["Slot", "12 + 5 for questions"],
             ["Spare", "1"], ["Slide to skip if late", "Slide 6 (signal)"]],
      so="The plan fits with one minute spare, and there's a slide to drop if needed.",
      split=4.90),

    O("4.9.9",
      what="Make decks **accessible**: strong colour **contrast**, never colour alone (add labels), **alt text** "
           "on every chart, and a checked **reading order**. Use Review > Check Accessibility.",
      like="Like ramps beside stairs: they help some people a lot and get in nobody's way.",
      how=["Review > Check Accessibility.", "Add alt text that states the finding.", "Add labels as well as colour."],
      tip="Good alt text gives the message: '180+ day bookings cancel 74% of the time'.",
      trap="Red and green as the only way to tell bars apart.",
      case="Accessibility check on the EkoSuites deck.",
      steps=["Run Check Accessibility.", "Fix each issue.", "Re-check."],
      table=[["Risk", "Fix"], ["Colour-only meaning", "Add value labels"], ["Charts without alt text", "Alt text stating the finding"],
             ["Wrong reading order", "Fix in the Selection Pane"], ["Low contrast", "Navy text on a light background"]],
      so="Everyone in the room, and anyone reading later, gets the same message.",
      split=4.90),

    O("4.9.10",
      what="When presenting, say the **headline first**, then the **evidence**, then the **ask**. Talk to the "
           "audience, not the screen, and end with the decision you need.",
      like="Like a good waiter: the specials first, the details if asked, then 'what would you like?'",
      how=["Open with the headline sentence.", "Walk through the evidence, one slide at a time.", "Close with a clear ask."],
      tip="Your last sentence should be the decision you want.",
      trap="Ending with 'Any questions?' and no ask.",
      case="Open and close the EkoSuites talk.",
      steps=["Opening line.", "Evidence.", "Closing ask."],
      table=[["Moment", "What to say"], ["Open", "Cancellations cost ₦4.30bn of bookings in 2023–24"],
             ["Evidence", "OTA, long lead times and first-time guests drive it"], ["Close (the ask)", "Approve a pilot of deposits on OTA bookings over 90 days from Q1"]],
      so="The board knows what you found and what you need from them.",
      split=4.90),
]
