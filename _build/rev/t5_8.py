from revlib import T

TRY = "ETEH_Capstone_Section_Examples.xlsx  ›  sheet 'P.8 Portfolio & CV'"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", "Your portfolio · the three capstones")
    return T(tag, **k)


TOPICS = [
    O("P.8.1",
      what="A **certificate** says you finished a course. A **portfolio** shows what you built, how you thought and "
           "what you found. A recruiter can check a portfolio in five minutes; a certificate must be taken on trust.",
      like="Like a tailor showing finished clothes, not only the certificate from the training centre.",
      how=["Pick your three best projects.", "Document each one the same way.", "Make each easy to open and check."],
      tip="Three well-documented projects beat a long list of courses.",
      trap="A portfolio of tutorials everyone else has done too.",
      case="The three capstones as portfolio pieces.",
      steps=["One write-up per capstone.", "Add a dashboard screenshot.", "Add the control totals."],
      table=[["Project", "Shows", "Headline number"], ["SuperMart (retail)", "Cleaning, modelling, diagnosis", "₦47.6m lost to deep discounts"],
             ["Crestview (education)", "Definitions, at-risk list, estimate", "86.4% vs 90% target"],
             ["Lifecare (health)", "SQL, capacity, safety, privacy", "89.5% occupancy vs 85%"]],
      so="Three domains, three real findings: plenty for an interviewer to ask about.",
      split=4.20),

    O("P.8.2",
      what="Each **write-up** has six headings: **question, data, method, finding, recommendation, outcome**. Add one "
           "dashboard screenshot and one table of control totals. Keep it to two pages.",
      like="Like a news story: what happened, how we know, and what happens next.",
      how=["Start with the question.", "Give scale and method in one line each.", "Finish with the finding and action."],
      tip="Label training outcomes honestly as a training scenario.",
      trap="Ten pages describing every cleaning step.",
      case="The Capstone 1 write-up.",
      steps=["Fill each heading.", "Check every number.", "Cut to two pages."],
      table=[["Heading", "Content"], ["Question", "Why did margin fall from 13.4% to 12.7% while sales grew 20%?"],
             ["Data", "9,994 POS lines, 12 stores, 1,862 SKUs; store returns cleaned"],
             ["Method", "Power Query, star schema, DAX, rate/mix split"],
             ["Finding", "Q4 margin 9.8% as discount depth rose to 14.8%"],
             ["Outcome", "Board adopted the Furniture cap for Q1 2026 (training scenario)"]],
      so="A reviewer understands the whole project in two minutes.",
      split=4.90),

    O("P.8.3",
      what="**Never publish data that identifies a person.** Replace names with labels, publish only totals, "
           "suppress **small groups**, blur names in screenshots, and say clearly that the data is anonymised.",
      like="Like a newspaper reporting 'a 45-year-old trader from Aba', not her name and address.",
      how=["Remove names and IDs.", "Publish group totals; hide tiny groups.", "Check screenshots and file metadata."],
      tip="Pick one small-group rule and state it. The P.8 sheet hides groups under 5; the Crestview paper uses 10.",
      trap="Removing names but leaving Student_IDs or exact dates.",
      case="Making the Lifecare work safe to share.",
      steps=["Drop Patient_ID and names.", "Report counts only.", "Add the statement."],
      table=[["Artefact", "Rule", "Lifecare example"], ["Frequent attenders", "Counts only", "12 patients admitted 3+ times in 2023"],
             ["Escalation list", "Admission_ID only", "814 Q4 rows, no names"],
             ["Statement", "Say it's anonymised", "'Anonymised ETEH training data'"]],
      so="You show real skill without exposing a single patient.",
      split=4.20),

    O("P.8.4",
      what="Send reviewers **one folder** in reading order: README, write-up, dashboard export and a workbook with "
           "sample data. Name every file with the version and date.",
      like="Like a well-packed school bag: the first book you need is on top.",
      how=["Number the files in reading order.", "Add version and date to each name.", "Include a sample, not full data."],
      tip="A reviewer should never have to ask which file to open first.",
      trap="A zip of 30 files called 'final', 'final2' and 'new'.",
      case="The Capstone 1 portfolio folder.",
      steps=["Export the PDFs.", "Name the files.", "Zip the folder."],
      table=[["File", "Content"], ["00_README.txt", "What's inside and how to open it"], ["01_Writeup.pdf", "The six-heading write-up"],
             ["02_Dashboard.pdf", "One-screen dashboard export"], ["03_Workbook_sample.xlsx", "Model with sample data"],
             ["Naming", "Each ends _v1.0_2026-01-20"]],
      so="The reviewer reads the story first, then opens the file.",
      split=4.90),

    O("P.8.5",
      what="A strong **CV bullet** follows: **action verb + method + scale + result**. Every bullet carries a number "
           "you can defend from your workbook.",
      like="Like a football CV: 'scored 14 goals in 20 games', not 'good at football'.",
      how=["Start with an action verb.", "Name the method and the scale.", "End with the result."],
      tip="If an interviewer asks 'how did you get that number?', you must be able to show it.",
      trap="Inventing time savings you never measured.",
      case="Bullets drawn from the capstones.",
      steps=["Pick one project.", "Write the bullet.", "Check every number."],
      table=[["Weak", "Strong"], ["Used Excel for analysis", "Built a star-schema model over 9,994 transactions, reconciling ₦733.2m with zero variance"],
             ["Did some SQL", "Caught a join error that inflated revenue to ₦7.19bn (true: ₦1.32bn)"],
             ["Worked on student data", "Defined progression for 6,607 students; built an at-risk list of 901"]],
      so="Specific, checkable bullets get interviews.",
      split=4.20),

    O("P.8.6",
      what="List only tools you can **demonstrate**. Describe the level honestly, and say how you used AI and how "
           "you **checked** its output. Overstated skills are found out in the first technical question.",
      like="Like saying you can swim: you may be asked to jump into the pool.",
      how=["List each tool.", "Describe the level you've actually used.", "Have a project that proves it."],
      tip="'SELECT to multi-table JOIN' is more convincing than 'expert in SQL'.",
      trap="Listing a tool you've only watched a video about.",
      case="Honest tool descriptions from the P.8 sheet.",
      steps=["Write the claim.", "Link it to a project.", "Cut anything you can't show."],
      table=[["Tool", "Describe honestly as"], ["Excel", "Power Query, Data Model, DAX, PivotTables, what-if"],
             ["SQL", "SELECT to multi-table JOIN, subqueries, window functions (SQLite)"],
             ["Statistics", "Descriptive stats, correlation, regression, forecasting"],
             ["AI", "Co-pilot use, checked against control totals"]],
      so="Every claim on your CV is backed by a project you can open.",
      split=4.90),

    O("P.8.7",
      what="**LinkedIn** is a route to your portfolio: a clear **headline**, a short About section naming your three "
           "projects, and an occasional post sharing one finding.",
      like="Like a shop signboard: it tells passers-by what you sell and where the door is.",
      how=["Headline: role + tools + domains.", "About: three projects, one line each.", "Link to the portfolio."],
      tip="Post one finding at a time, with a chart and the 'so what'.",
      trap="'Aspiring data scientist' with no link to any work.",
      case="A profile built from the capstones.",
      steps=["Write the headline.", "Write the About section.", "Add the portfolio link."],
      table=[["Element", "Example"], ["Headline", "Data Analyst | Excel, Power Query, DAX, SQL | Retail, Health, Education"],
             ["About", "Three projects: retail margin, student progression, hospital capacity"],
             ["Post idea", "'Why 20% sales growth still cut margin', with one chart"]],
      so="A recruiter goes from your profile to your work in one click.",
      split=4.90),

    O("P.8.8",
      what="Expect: 'Walk me through this dashboard', 'Why this measure?', 'How did you verify it?', 'What was "
           "hardest?' Answer with **specifics from your build**, not general talk about process.",
      like="Like a chef explaining one dish: the ingredients, the timing and the mistake they learnt from.",
      how=["Prepare a 60-second walkthrough.", "Have your verification evidence ready.", "Pick a real 'hardest moment'."],
      tip="Concrete beats general: name the number, the sheet, the decision.",
      trap="'I followed best practice' with no example.",
      case="Answers drawn from the capstones.",
      steps=["Write each answer.", "Time the walkthrough.", "Rehearse aloud."],
      table=[["Question", "Answer shape"], ["Walk me through it", "KPI cards → Q4 margin chart → recommendation, in 60 s"],
             ["How did you verify it?", "7 reconciliation tests, all PASS, ₦0 difference"],
             ["What was hardest?", "70 stays with no discharge date: excluded and disclosed"]],
      so="Specific answers show you really did the work.",
      split=4.90),

    O("P.8.9",
      what="Keep the portfolio **current**: add about one project a quarter, retire work you'd no longer defend, "
           "and update write-ups when your methods improve. Record the changes in the README.",
      like="Like a shop window: change the display or customers assume the shop has closed.",
      how=["Add one project a quarter.", "Retire weak or outdated work.", "Log changes in the README."],
      tip="Extend a capstone in a new tool, e.g. rebuild Capstone 1 in Power BI.",
      trap="The same three projects, untouched, for two years.",
      case="A maintenance plan.",
      steps=["Set a quarterly reminder.", "Review each project.", "Update the revision history."],
      table=[["Action", "Frequency", "Signal"], ["Add a project", "Quarterly", "You keep learning"],
             ["Retire weak work", "Each review", "You judge your own work"],
             ["Update the README history", "Each change", "Visible improvement over time"]],
      so="A living portfolio tells employers you're still growing.",
      split=4.20),
]
