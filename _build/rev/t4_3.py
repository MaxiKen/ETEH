from revlib import T

TRY = "ETEH_Week4_Section_Examples.xlsx  ›  sheet '4.3 Correlation'"
CV = "Crestview University"
HT = "EkoSuites Hotel"
SAL = "ETEH Talent Survey"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    return T(tag, **k)


TOPICS = [
    O("4.3.1",
      what="**Correlation** measures how strongly two number columns **move together**, from **−1** (one goes "
           "up as the other goes down) through **0** (no straight-line link) to **+1** (they rise together).",
      like="Like two friends walking: some walk in step, some walk opposite ways, some ignore each other.",
      how=["Pick two number columns.", "=CORREL(column1, column2).", "Read the sign and size."],
      tip="Rough guide: under 0.3 weak, 0.3 to 0.7 moderate, over 0.7 strong.",
      trap="Reading correlation as proof that one causes the other.",
      who=CV,
      case="Which student habits move with exam scores? (6,607 students, tCV)",
      steps=["CORREL of each habit with Exam_Score.", "Sort by size.", "Describe each in words."],
      table=[["Pair", "r", "Reading"], ["Attendance % and exam score", "0.58", "Moderate positive"],
             ["Hours studied and exam score", "0.45", "Moderate positive"], ["Previous score and exam score", "0.18", "Weak"],
             ["Sleep hours and exam score", "-0.02", "No straight-line link"]],
      so="Attendance and study hours move most with exam scores; sleep shows no link in this data.",
      split=4.70),

    O("4.3.2",
      what="**Pearson's r** (=CORREL or =PEARSON) measures a **straight-line** relationship between two number "
           "columns. It's sensitive to extreme values and misses curved patterns.",
      like="Like checking how well a ruler fits along a scatter of dots.",
      how=["=CORREL(tCV[Attendance_Pct], tCV[Exam_Score]).", "Square it (r²) for the share of variation shared.", "Plot the pair to check (4.3.6)."],
      tip="r² is easier to explain: 0.58² ≈ 0.34, so about a third.",
      trap="Using Pearson on text codes like Faculty.",
      who=CV,
      case="How strongly is attendance linked to exam score?",
      steps=["CORREL on the two columns.", "Square for r².", "Write the sentence."],
      table=[["Measure", "Result"], ["Pearson r", "0.58"], ["r²", "About 0.34"], ["Number of students (n)", "6,607"],
             ["Sentence", "Higher attendance goes with higher scores"]],
      so="Attendance is moderately linked to scores; about a third of score variation moves with attendance.",
      split=4.90),

    O("4.3.3",
      what="**Spearman's rank correlation** correlates the **ranks** instead of the values. It handles skewed "
           "data and outliers, and catches any steady up-or-down pattern, even curved. In Excel: rank both columns "
           "with RANK.AVG, then CORREL the ranks.",
      like="Like comparing race positions (1st, 2nd, 3rd) rather than exact finishing times.",
      how=["=RANK.AVG(x, column) for each column.", "=CORREL(rank1, rank2).", "Compare with Pearson."],
      tip="If Spearman is much higher than Pearson, outliers or a curve are hiding the link.",
      trap="Using RANK.EQ: ties need average ranks.",
      who=CV,
      case="A sample of 300 students: hours studied vs exam score, Pearson vs Spearman (tSpear).",
      steps=["Rank_Hours and Rank_Score with RANK.AVG.", "CORREL the ranks.", "Compare with Pearson."],
      table=[["Data", "Pearson r", "Spearman"], ["Hours studied, 300 sampled", "0.42", "0.50"],
             ["Hours studied, all 6,607", "0.45", "0.48"], ["Attendance %, all 6,607", "0.58", "0.67"]],
      so="Spearman is higher in each case: a few unusual scores weaken Pearson, but the ranks show a steady link.",
      split=4.70),

    O("4.3.4",
      what="A **correlation matrix** shows every pair of columns in one grid. The diagonal is always 1. Read one "
           "row for 'what moves with the target', and check the other cells for inputs that move together.",
      like="Like a distance chart in a road atlas: every city against every other city.",
      how=["List the number columns across and down.", "CORREL in every cell.", "Highlight the target row."],
      tip="Inputs that are strongly linked to each other repeat the same information.",
      trap="Reading the diagonal 1s as findings.",
      who=CV,
      case="The analyst builds a matrix of four habits and Exam_Score.",
      steps=["CORREL for every pair.", "Read the Exam_Score row.", "Check the habit-to-habit cells."],
      table=[["Finding", "Value"], ["Exam score with attendance", "0.58"], ["Exam score with hours studied", "0.45"],
             ["Attendance with hours studied", "-0.01"], ["Largest habit-to-habit link", "About 0.02"]],
      so="The habits don't overlap: attendance and study hours each bring separate information about scores.",
      split=4.90),

    O("4.3.5",
      what="**Data > Data Analysis > Correlation** builds the whole matrix in one step. Select all the number "
           "columns **side by side** with headers, tick **Labels in first row**. The result is static: rerun it "
           "when the data changes.",
      like="Like a photocopier: fast and complete, but the copy doesn't update when the original changes.",
      how=["Place the number columns next to each other.", "Data > Data Analysis > Correlation.", "Tick Labels; choose an output cell."],
      tip="Use CORREL formulas when you need the matrix to stay live.",
      trap="Including ID columns, which gives meaningless coefficients.",
      who=CV,
      case="The analyst runs the ToolPak Correlation on five columns and checks it against the live matrix.",
      steps=["Select Hours, Attendance, Previous, Tutoring, Exam.", "Run Correlation with Labels ticked.", "Compare with the CORREL grid."],
      table=[["Setting", "Correct choice", "Why"], ["Input range", "Number columns only", "IDs and text break it"],
             ["Grouped by", "Columns", "Each column is a variable"], ["Labels in first row", "Ticked", "Names in the output"],
             ["Exam vs attendance", "0.58 in both", "ToolPak matches CORREL"]],
      so="The ToolPak and CORREL give the same numbers; only the CORREL version updates.",
      split=4.20),

    O("4.3.6",
      what="**Always plot the pair** before quoting r. Very different patterns (a curve, a cluster, a few "
           "extreme points) can give the same r. A scatter chart shows what the number hides.",
      like="Like reading a menu description and then seeing the actual plate: look before you trust.",
      how=["Insert > Scatter on the two columns.", "Add a linear trendline with R².", "Look for curves, groups and extreme points."],
      tip="Re-run CORREL without the unusual points to see how much they matter; don't delete them.",
      trap="Quoting r without ever looking at the chart.",
      who=CV,
      case="The scatter of hours vs score shows 43 students scoring above 80 at every level of study.",
      steps=["Plot hours against exam score.", "Spot the high scores spread across all hours.", "Recompute r without them, for comparison."],
      table=[["Pair", "All 6,607", "Without the 43 above 80"], ["Hours and exam score", "0.45", "0.52"],
             ["Attendance and exam score", "0.58", "0.67"]],
      so="A few high scorers, unrelated to study time, weaken r. The chart shows it; the number alone doesn't.",
      split=4.70),

    O("4.3.7",
      what="**Correlation is not causation.** Two columns can move together because A causes B, B causes A, a "
           "third factor drives both, or by chance. Only a test or controlled change proves cause.",
      like="Ice-cream sales and drowning both rise in hot weather; ice cream doesn't cause drowning.",
      how=["Find the correlation.", "List other explanations.", "Say 'associated with', not 'causes'."],
      tip="Ask: if we changed A on purpose, would B really change?",
      trap="Launching a policy based on a correlation alone.",
      who=HT,
      case="What goes with cancellations at EkoSuites? (36,275 bookings; tLeadBand)",
      steps=["CORREL of lead time with Cancelled (1/0).", "Cancel rate by lead-time band.", "CORREL of special requests with Cancelled."],
      table=[["Measure", "Result"], ["r(lead time, cancelled)", "0.44"], ["Cancel rate, booked 0–7 days ahead", "8.9%"],
             ["Cancel rate, booked 180+ days ahead", "73.9%"], ["r(special requests, cancelled)", "-0.25"]],
      so="Guests who make requests cancel less, but forcing requests won't stop cancellations: committed guests make requests.",
      split=4.90),

    O("4.3.8",
      what="A **confounder** is a third factor linked to both columns, which creates or exaggerates a link. Test "
           "it by comparing **within groups** of the suspected confounder. Some suspects turn out innocent.",
      like="Like blaming a new cook for slow service when the restaurant also got twice as many customers that week.",
      how=["Name a suspected third factor.", "AVERAGEIFS within each group of it.", "Compare the within-group gap with the raw gap."],
      tip="Report both the raw and the adjusted number.",
      trap="Claiming the adjusted gap means there's 'no real gap'.",
      who=SAL,
      case="Women in the survey earn 9.0% less on average. Do seniority or experience explain it? (tGapSeniority, tGapExperience)",
      steps=["Gap within each seniority level.", "Gap within each experience band.", "Weight the band gaps by headcount."],
      table=[["View", "Gender pay gap"], ["Raw (all women vs all men)", "9.0%"], ["Seniority held constant", "9.1%: not a confounder"],
             ["Experience held constant", "4.6%: about half explained"], ["Experience, women vs men", "8.7 vs 9.5 years on average"]],
      so="Experience explains about half the gap; seniority explains none. The remaining 4.6% is still real.",
      split=4.90),

    O("4.3.9",
      what="**Simpson's paradox**: a pattern seen in every group **reverses** when the groups are added together. "
           "It happens when the groups have very different sizes and rates. Always check key subgroups.",
      like="Like a student who beats a friend in every subject, yet has a lower overall average because they took more of the hardest subject.",
      how=["Compare the rates overall.", "Split by a grouping column (here year).", "Check whether the direction flips."],
      tip="If the groups are very unequal in size, split before you conclude.",
      trap="Reporting only the overall rate.",
      who=HT,
      case="Does Deluxe cancel more than Standard Twin? (tSimpson)",
      steps=["Cancel rate by room type overall.", "Same split for 2023 and 2024.", "Check where most bookings fall."],
      table=[["View", "Deluxe", "Standard Twin"], ["2023", "2.6%", "11.6%"], ["2024", "37.3%", "38.7%"],
             ["All years", "34.2%", "32.9%"], ["Share of bookings in 2024", "91%", "79%"]],
      so="Deluxe cancels less in each year, but more overall, because most Deluxe bookings are in high-cancellation 2024.",
      split=4.70),

    O("4.3.10",
      what="Report correlation **responsibly**: give **r, n and the method**, show the scatter, describe the "
           "strength in words, say 'associated with', name possible confounders, and check subgroups.",
      like="Like a doctor's note: the reading, how it was measured and what it doesn't tell you.",
      how=["r, n, Pearson or Spearman, and a chart.", "Strength in words.", "Confounders and subgroup checks."],
      tip="Check the same correlation in each subgroup before generalising.",
      trap="'Attendance causes higher scores.'",
      who=CV,
      case="The analyst writes the attendance finding for the faculty report.",
      steps=["Give r and n.", "Check r within each faculty.", "Write the careful sentence."],
      table=[["Rule", "Example from our data"], ["r, n, method", "Pearson r = 0.58, n = 6,607, with scatter"],
             ["Check subgroups", "r is 0.56 to 0.67 in each of the 4 faculties"], ["Careful wording", "Attendance is associated with higher scores"],
             ["Name confounders", "Motivation may drive both attendance and scores"]],
      so="The link holds in every faculty, and the wording claims only what the data shows.",
      split=4.90),
]
