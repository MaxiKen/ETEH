from revlib import T

TRY = "ETEH_Capstone_Section_Examples.xlsx  ›  sheet 'P.4 Capstone2 Crestview'"
SOL = "03_Instructor_Projects  ›  Capstone2_Crestview_SOLUTION.xlsx  ›  "


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", "Crestview University · Capstone 2")
    return T(tag, **k)


TOPICS = [
    O("P.4.1",
      what="Capstone 2: the Senate wants **90%** of students in good standing (**progressing**). Which faculties are "
           "behind, what drives performance, and who needs support in Week 5? Fix the **cohort** (who counts) before "
           "calculating anything.",
      like="Like a coach checking how many players are fit for the next match, and who needs extra training.",
      how=["Define the cohort: one row per registered student.", "Write the progression rule.", "Compare with the 90% target."],
      tip="Put the thresholds (62 and 65%) in input cells so anyone can see and change them.",
      trap="Measuring progression on whoever happens to be in the file.",
      case="The headline progression figure, 2024/25.",
      steps=["Cohort = ROWS(tCV).", "Progressed = COUNTIFS(score >= 62, attendance >= 65).", "Gap = progression − 90%."],
      table=[["Measure", "Value"], ["Cohort students", "6,607"], ["Progressed", "5,706"], ["Progression", "86.4%"],
             ["Gap to 90% target", "−3.6 points"]],
      so="Crestview is 3.6 points short of the Senate target.",
      split=4.90),

    O("P.4.2",
      what="The **inventory** for Capstone 2 is one registry export with faculty, department, level, attendance, "
           "study habits and exam score. The published cohort (6,607 students) is the **control total**.",
      like="Like checking the class register against the school's official headcount.",
      how=["Count the raw rows.", "List the defects you can see.", "Note the control total and its owner."],
      tip="Write the control total down before cleaning, so you know where you must end up.",
      trap="Assuming the raw row count is the number of students.",
      case="Crestview_Registry_Export_RAW.csv.",
      steps=["Count rows.", "Count department spellings.", "Record the control."],
      table=[["Item", "Value"], ["Raw rows", "6,652"], ["Department spellings", "28 (14 real departments)"],
             ["Faculties / levels", "4 / 100–400"], ["Defect types", "8"], ["Control total (registry)", "6,607 students"]],
      so="6,652 rows must become 6,607 students. The 45-row gap has to be explained.",
      tryit=SOL + "'2 Clean & Reconcile'",
      split=4.90),

    O("P.4.3",
      what="Clean the registry in a safe order: **TRIM the IDs first**, map M/F codes, turn '84%' into 84, map "
           "department spellings to 14 names, standardise Yes/No, cap the one score of 101, **then** remove duplicates.",
      like="Like tidying names on a guest list before crossing out repeats: 'Ade ' and 'Ade' are the same guest.",
      how=["TRIM Student_ID.", "Fix codes, text numbers and spellings.", "Remove duplicates last."],
      tip="VALUE(\"84%\") gives 0.84, not 84. Remove the % sign first.",
      trap="Removing duplicates before trimming: 25 padded IDs would survive as 'different' students.",
      case="The Capstone 2 defect log.",
      steps=["Log each defect.", "Apply the fix.", "Check the cohort = 6,607."],
      table=[["Defect", "Rows", "Fix"], ["Attendance typed '84%'", "330", "VALUE(SUBSTITUTE(…,\"%\",\"\"))"],
             ["Department variants", "198", "Lookup to department master"], ["Padded Student_IDs", "25", "TRIM before lookups"],
             ["Double registrations", "45", "Remove duplicates on trimmed ID"], ["Score of 101", "1", "Query, then cap at 100"]],
      so="6,652 rows become 6,607 unique students, matching the registry.",
      tryit=SOL + "'2 Clean & Reconcile'",
      split=4.70),

    O("P.4.4",
      what="**Progression %** = students meeting the rule ÷ cohort. The rule is **exam of 62 or more and attendance of 65% or more**. "
           "**At risk** is everyone else: exam < 62 **or** attendance < 65%.",
      like="Like WAEC credit rules: written down before the exam, the same for everyone.",
      how=["Write the rule in words.", "Put thresholds in input cells.", "COUNTIFS with \">=\"&cell."],
      tip="Use the same denominator (the cohort) every time you compare.",
      trap="Changing the rule or the denominator to get a nicer number.",
      case="The progression definition and its values.",
      steps=["Enter 62, 65 and 90%.", "COUNTIFS for progressed.", "Divide by the cohort."],
      table=[["Component", "Definition", "Value"], ["Cohort", "All students on the registry extract", "6,607"],
             ["Progressed", "Exam 62+ AND attendance 65+", "5,706"], ["At risk", "Exam < 62 OR attendance < 65", "901 (13.6%)"],
             ["Progression %", "Progressed ÷ cohort", "86.4%"]],
      so="One written rule gives one defensible number: 86.4%.",
      split=4.20),

    O("P.4.5",
      tryit=SOL + "'4 Dashboard'",
      what="Break the at-risk rate down by group to find **where risk concentrates**. If one group stands out, "
           "that's where support pays off most.",
      like="Like a doctor checking which symptom most patients share before choosing a treatment.",
      how=["At-risk rate by faculty.", "At-risk rate by attendance band.", "Compare the spread."],
      tip="Flat bars by faculty and steep bars by attendance tell you the cause is behaviour, not faculty.",
      trap="Targeting a faculty because it has the most students, not the highest rate.",
      case="At-risk rate by attendance band.",
      steps=["Band attendance.", "At-risk count ÷ students per band.", "Compare with faculty rates."],
      table=[["Group", "Students", "At-risk rate"], ["Attendance below 70%", "1,573", "53.5%"],
             ["Attendance 70–79%", "1,681", "3.1%"], ["Attendance 90%+", "1,746", "0.0%"],
             ["Faculties (range)", "1,014–2,250", "12.7% to 14.6%"]],
      so="Faculty rates are flat; attendance is the big separator (r = 0.58 with exam score).",
      split=4.70),

    O("P.4.6",
      what="An **at-risk list** must be usable: student ID, faculty, department, level, the **trigger** (why "
           "flagged) and the **support route**. It supports a human adviser; it is not a judgement.",
      like="Like a nurse's triage list: who to see first and why, not a diagnosis.",
      how=["Filter students meeting the at-risk rule.", "Add the trigger.", "Map each trigger to a support route."],
      tip="Restrict the list to advisers: it contains student IDs.",
      trap="Presenting a flag as a final verdict on a student.",
      case="901 students split by trigger.",
      steps=["COUNTIFS each trigger.", "Assign the route.", "Sort by score."],
      table=[["Trigger", "Students", "Support route"], ["Low attendance only", "585", "Pastoral follow-up"],
             ["Low score only", "161", "Academic tutoring"], ["Both", "155", "Adviser meeting in Week 5"],
             ["Total", "901", "Restricted list"]],
      so="Each adviser knows why a student is on the list and what to offer.",
      split=4.70),

    O("P.4.7",
      what="Estimate the effect of an intervention from the data, **state the assumption** and give a **range**. "
           "A slope shows an association, not proof that the programme will cause the change.",
      like="Like a farmer estimating harvest from rainfall: a useful guess with a range, not a promise.",
      how=["SLOPE(exam, attendance).", "Apply an assumed lift to low attenders.", "Re-apply the rule; report a range."],
      tip="Say how many students move only because attendance is part of the rule.",
      trap="Claiming the programme 'will' raise progression to an exact figure.",
      case="Assume a Week-5 programme lifts low attenders by 10 points.",
      steps=["Slope = 0.196 exam points per attendance point.", "Lift 740 low attenders by 10.", "Re-apply the rule."],
      table=[["Item", "Value"], ["Exam gain per +10 attendance", "+1.96 points"], ["Students moved to progression", "652"],
             ["…through attendance alone", "585"], ["Progression if achieved", "96.2% (estimate)"],
             ["Range for +5 to +15 points", "585–694 students"]],
      so="Only 67 of the 652 depend on the predicted exam gain. Report it as an estimate, and test with a comparison group.",
      split=4.90),

    O("P.4.8",
      what="The Senate dashboard fits one screen: KPI cards (Cohort, Progression %, At-risk count, Average "
           "attendance), at-risk rate by faculty and by level, and the restricted at-risk list.",
      like="Like a school notice board: the key numbers big, the details one step away.",
      how=["KPI cards at the top.", "Faculty and level bars, sorted.", "List on a restricted page."],
      tip="Put flat faculty bars beside the steep attendance chart: the message is instant.",
      trap="Leaving the named list on the public page.",
      case="At-risk rate by level.",
      steps=["COUNTIFS by level.", "Rate = 1 − progressed ÷ students.", "Sort and chart."],
      table=[["Level", "Students", "At-risk rate"], ["100", "1,994", "13.4%"], ["200", "1,767", "14.5%"],
             ["300", "1,508", "12.7%"], ["400", "1,338", "13.9%"]],
      so="Levels are as flat as faculties. Target students by behaviour, not by level or faculty.",
      split=4.70),

    O("P.4.9",
      what="The **Senate paper** follows a fixed order: context, finding, evidence, recommendation (with owner) and "
           "risks. The effect is stated as a **range** and called an association.",
      like="Like a well-argued letter to a commissioner: problem, proof, request, risks.",
      how=["Context and finding first.", "One strong piece of evidence.", "Recommendation, owner, range, risks."],
      tip="If the pilot result differs, a range stated honestly still stands.",
      trap="Overselling the estimate as certain.",
      case="The Capstone 2 Senate paper.",
      steps=["Write each section.", "Check every number.", "Name the owner."],
      table=[["Section", "Content"], ["Context", "Target: 90% in good standing"], ["Finding", "86.4%, 3.6 points short"],
             ["Evidence", "53.5% at risk below 70% attendance"],
             ["Recommendation", "Week-5 attendance + tutorials for 901 students; owners: Deans"],
             ["Effect", "585–694 students (estimate, not proven cause)"]],
      so="A clear, honest paper that survives challenge.",
      split=4.90),

    O("P.4.10",
      what="Verify the cohort against the registry, the at-risk count against the rule, and that every slicer "
           "drives every visual. Handover includes the **definitions**, because they're what people will question.",
      like="Like checking the exam register against the hall attendance before releasing results.",
      how=["Cohort = registry figure.", "Progressed = examinations office figure.", "Test every slicer."],
      tip="Mark the workbook restricted: it contains student IDs.",
      trap="Handing over without the definitions document.",
      case="Capstone 2 verification.",
      steps=["Compare with each control.", "Record the result.", "Name the owner."],
      table=[["Check", "Control", "Result"], ["Cohort", "Registry: 6,607", "Match"],
             ["Progressed", "Examinations office: 5,706", "Match"], ["At risk", "Rule applied: 901", "Match"],
             ["Duplicates removed", "Defect log: 45", "Documented"]],
      so="Every headline number traces back to an official control.",
      split=4.70),
]
