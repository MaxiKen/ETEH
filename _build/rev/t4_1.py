from revlib import T

TRY = "ETEH_Week4_Section_Examples.xlsx  ›  sheet '4.1 Descriptive Stats'  ·  D_Salary, D_Orders"
SAL = "ETEH Talent Survey"
ON = "SuperMart Online"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", SAL)
    return T(tag, **k)


TOPICS = [
    O("4.1.1",
      what="**Descriptive statistics** sum up a whole column in a few numbers: where the **centre** is (typical "
           "value), how **spread out** the values are, and the **shape** (lopsided or balanced). They describe "
           "what happened; they don't explain why.",
      like="It's like describing a crowd in three sentences: how tall people typically are, how much heights "
           "vary, and whether a few giants stand out.",
      how=["Centre: AVERAGE, MEDIAN, MODE.", "Spread: MIN, MAX, quartiles, STDEV.S.", "Shape: SKEW and KURT."],
      tip="Always run these before any deeper analysis of a new dataset.",
      trap="Reporting one average and calling it the analysis.",
      case="HR has the ETEH Talent salary survey: 1,781 professionals (D_Salary). The analyst summarises salaries.",
      steps=["Count the rows: 1,781.", "Calculate centre and spread.", "Check the shape."],
      table=[["Family", "Measure", "Annual salary"], ["Centre", "Median", "₦11,000,000"],
             ["Spread", "Middle 50% (Q1 to Q3)", "₦7,000,000 to ₦16,000,000"], ["Shape", "Skewness", "0.09 (balanced)"],
             ["Size", "Count", "1,781 respondents"]],
      so="In four numbers HR knows the typical salary, the usual range and that the spread is balanced.",
      split=4.20),

    O("4.1.2",
      what="The **mean** (=AVERAGE) adds everything and divides by the count. It uses every value, so a few very "
           "large values pull it up. It's ideal for balanced data and for totals (mean × count = sum).",
      like="The mean is like sharing a pot of money equally: one rich person joining the group raises everyone's "
           "share on paper.",
      how=["=AVERAGE(tSal[Annual_Salary_NGN]).", "Compare it with the median.", "If they're far apart, the data is lopsided."],
      tip="Mean × count = total; that's why finance likes the mean.",
      trap="Quoting the mean of lopsided data as 'typical'.",
      case="The analyst compares the mean for balanced salaries and lopsided online order values.",
      steps=["Mean and median of salary.", "Mean and median of order value (D_Orders).", "Remove the 10 largest orders and recompute."],
      table=[["Data", "Mean", "Median"], ["Salary (1,781)", "₦11,344,180", "₦11,000,000"],
             ["Order value (11,000)", "₦126,792", "₦29,300"], ["Order value without top 10", "₦124,114", "₦29,250"]],
      so="For salaries the mean is fair. For orders, a few big ones push the mean to over 4 times the median.",
      split=4.20),

    O("4.1.3",
      what="The **median** (=MEDIAN) is the **middle value** after sorting: half are below, half above. Extreme "
           "values don't move it, so use it for lopsided data like prices, incomes and waiting times.",
      like="The median is the person standing in the middle of a line sorted by height; a giant at the end doesn't change who that is.",
      how=["=MEDIAN(range).", "Compare with AVERAGE.", "If mean ÷ median is above about 1.2, report the median as 'typical'."],
      tip="For money, start with the median.",
      trap="Saying 'the average order is ₦126,792' when most orders are far smaller.",
      who=ON,
      case="Marketing asks: what does a typical SuperMart Online order look like?",
      steps=["=MEDIAN(tOrders[total_price]).", "=AVERAGE(tOrders[total_price]).", "Divide mean by median."],
      table=[["Measure", "Order value"], ["Median (typical order)", "₦29,300"], ["Mean", "₦126,792"],
             ["Mean ÷ median", "4.33"], ["Orders at or below ₦29,300", "Half of 11,000"]],
      so="A typical order is ₦29,300. The mean is over 4 times higher because of a few large electronics orders.",
      split=4.90),

    O("4.1.4",
      what="The **mode** is the **most common value** (=MODE.SNGL for numbers, or a count for text). It's the "
           "only 'centre' that works for categories, like the most common job level or basket size.",
      like="The mode is the most popular dish on the menu: what most people ordered.",
      how=["Numbers: =MODE.SNGL(range).", "Text: COUNTIF per category, or a PivotTable count.", "Report what the mode tells you."],
      tip="Use the mode for categories and whole-number counts.",
      trap="Taking an 'average' of categories like seniority.",
      case="HR and operations want the most common value in four columns.",
      steps=["MODE.SNGL on salary.", "Count Seniority values.", "MODE.SNGL on quantity and on length of stay."],
      table=[["Column", "Mode", "How often"], ["Salary", "₦12,000,000", "Round-number answers"], ["Seniority", "Senior", "766 of 1,781"],
             ["Order quantity (Online)", "1 item", "5,519 of 11,000"], ["Length of stay 2023 (Lifecare)", "6 days", "Most common stay"]],
      so="Most orders contain 1 item, and 'Senior' is the most common level. The mode also exposes round-number answers.",
      split=4.20),

    O("4.1.5",
      what="The **range** is max minus min; one extreme value changes it. The **interquartile range (IQR)** = Q3 "
           "minus Q1 is the spread of the **middle 50%**, and isn't affected by extremes.",
      like="The range is the gap between the shortest and tallest person; the IQR is the gap across the middle half of the crowd.",
      how=["Range: =MAX(r)-MIN(r).", "Q1 and Q3: =QUARTILE.INC(r,1) and (r,3).", "IQR = Q3 - Q1."],
      tip="Report the IQR alongside the median.",
      trap="Using the range to describe 'typical' spread.",
      case="HR wants to describe how widely salaries vary.",
      steps=["MIN and MAX; range.", "Q1 and Q3.", "IQR."],
      table=[["Measure", "Salary"], ["Minimum / maximum", "₦2,500,000 / ₦25,000,000"], ["Range", "₦22,500,000"],
             ["Q1 / Q3", "₦7,000,000 / ₦16,000,000"], ["IQR (middle 50%)", "₦9,000,000"]],
      so="Half of all professionals earn between ₦7m and ₦16m. The ₦22.5m range is set by just the extremes.",
      split=4.90),

    O("4.1.6",
      what="**Standard deviation** (SD) is the typical distance of values from the mean, in the original units. "
           "**Variance** is SD squared. Use **STDEV.S** for a sample, **STDEV.P** only when you have every member "
           "of the population.",
      like="SD is like how far, on average, students sit from the middle of the classroom.",
      how=["=STDEV.S(range) for survey or sample data.", "=VAR.S(range) for variance (units squared).",
           "Compare SD with the mean."],
      tip="Most business data is a sample: use STDEV.S.",
      trap="Reporting variance in ₦² to managers.",
      case="The analyst measures the spread of salaries, exam scores and 2023 hospital stays.",
      steps=["STDEV.S for each column.", "VAR.S for salary.", "Read SD next to the mean."],
      table=[["Column", "Mean", "SD (STDEV.S)"], ["Salary", "₦11,344,180", "₦5,136,701"],
             ["Exam score (Crestview)", "67.2", "3.9"], ["Length of stay 2023 (Lifecare)", "15.5 days", "8.8 days"]],
      so="Exam scores sit tightly around 67; hospital stays vary a lot around 15.5 days.",
      split=4.20),

    O("4.1.7",
      what="The **coefficient of variation (CV)** = SD ÷ mean, as a %. It removes the units, so you can compare "
           "the spread of naira, days and exam points fairly.",
      like="CV is like comparing how wobbly a bicycle and a lorry are, relative to their size.",
      how=["=STDEV.S(r)/AVERAGE(r).", "Format as %.", "Compare across columns."],
      tip="Use CV whenever the columns have different units or sizes.",
      trap="Comparing SDs of columns measured in different units.",
      who="Multi-Sector",
      case="Which is most variable: salaries, online order values, hospital stays or exam scores? (tCVcompare)",
      steps=["Mean and SD for each.", "CV = SD ÷ mean.", "Rank by CV."],
      table=[["Measure", "Mean", "SD", "CV"], ["Online order value", "₦126,792", "₦278,268", "219.5%"],
             ["Length of stay 2023", "15.5 days", "8.8 days", "56.7%"], ["Annual salary", "₦11,344,180", "₦5,136,701", "45.3%"],
             ["Exam score", "67.2", "3.9", "5.8%"]],
      so="Order values are by far the most variable; exam scores are the most consistent.",
      split=4.20),

    O("4.1.8",
      what="**Skewness** (=SKEW) measures lopsidedness. About **0** = balanced; **positive** = a long tail of high "
           "values; **negative** = a long tail of low values. Above about 1, report the median.",
      like="Skewness is like a seesaw: balanced at zero, tipped when a few heavy values sit on one end.",
      how=["=SKEW(range).", "Read the sign and size.", "Choose median or mean accordingly."],
      tip="Positive skew is normal for money and waiting times.",
      trap="Running methods that assume balance on strongly skewed data.",
      who="Multi-Sector",
      case="The analyst checks the shape of three columns.",
      steps=["SKEW on order value.", "SKEW on salary.", "SKEW on exam score."],
      table=[["Column", "Skewness", "Shape and choice"], ["Online order value", "4.61", "Long right tail: use median"],
             ["Exam score", "1.64", "A few very high scores"], ["Salary", "0.09", "Balanced: mean is fine"]],
      so="Order values are strongly lopsided, which is why their median (₦29,300) is the honest 'typical' value.",
      split=4.20),

    O("4.1.9",
      what="**Kurtosis** (=KURT) measures how **heavy the tails** are compared with a normal bell curve. Above 0 "
           "means more extreme values than expected; below 0 means flatter, with fewer extremes.",
      like="Kurtosis is like the weather: some places have mild days with occasional storms (heavy tails), "
           "others are steady all year.",
      how=["=KURT(range).", "Above 0: expect surprises in the tails.", "Below 0: values are spread flat."],
      tip="High kurtosis is a warning that averages and SDs can mislead.",
      trap="Assuming a bell curve without checking.",
      who="Multi-Sector",
      case="The analyst checks the tails of four columns.",
      steps=["KURT on each column.", "Compare with 0.", "Note what it means."],
      table=[["Column", "Kurtosis", "Meaning"], ["Online order value", "28.9", "Very heavy tail: rare huge orders"],
             ["Exam score", "10.5", "Very peaked: 81% score 63–71"], ["Salary", "-1.14", "Flat: spread evenly"],
             ["Length of stay 2023", "-1.23", "Flat: stays spread evenly"]],
      so="Orders have rare extreme values; exam scores bunch tightly; salaries and stays spread out evenly.",
      split=4.20),

    O("4.1.10",
      what="**Percentiles** split sorted data into 100 parts: P90 is the value 90% fall below. **Quartiles** are "
           "P25, P50 (median) and P75. Use =PERCENTILE.INC. Pay bands and service targets use percentiles.",
      like="Percentiles are like exam rankings: 'top 10%' means you beat 90% of the class.",
      how=["=PERCENTILE.INC(range, 0.9) for P90.", "=QUARTILE.INC(range, 1) for Q1.", "Build a small table of P10 to P90."],
      tip="Use percentiles to set bands; they're robust to extremes.",
      trap="Mixing PERCENTILE.INC and .EXC in the same report.",
      case="HR builds salary bands from the survey (tSalPct).",
      steps=["Enter k = 0.1, 0.25, 0.5, 0.75, 0.9.", "=PERCENTILE.INC(tSal[Annual_Salary_NGN], k).", "Use them as bands."],
      table=[["Percentile", "Salary", "Use"], ["P10", "₦4,500,000", "Entry band floor"], ["P25 / P50", "₦7,000,000 / ₦11,000,000", "Lower band / midpoint"],
             ["P75", "₦16,000,000", "Upper band"], ["P90", "₦18,500,000", "Top 10% start here"]],
      so="Anyone earning over ₦18.5m is in the top 10%; the middle half earn ₦7m to ₦16m.",
      split=4.20),

    O("4.1.11",
      what="A common outlier rule: a value is an **outlier** if it's more than **1.5 × IQR** below Q1 or above Q3. "
           "Outliers aren't automatically errors. **Investigate before removing**; some are your most important records.",
      like="An outlier is like a very tall student in class: unusual, but real. You check, you don't erase.",
      how=["Q1, Q3 and IQR.", "Upper fence = Q3 + 1.5 × IQR; lower fence = Q1 - 1.5 × IQR.", "COUNTIFS above / below the fences."],
      tip="Check what share of the total the outliers hold before removing anything.",
      trap="Deleting outliers automatically.",
      who=ON,
      case="An analyst wants to 'clean out' high-value outliers from SuperMart Online orders.",
      steps=["Q1 ₦9,550, Q3 ₦96,462.50; IQR ₦86,912.50.", "Upper fence ₦226,831.25.", "Count orders above it and their revenue share."],
      table=[["Measure", "Result"], ["Upper fence", "₦226,831.25"], ["Lower fence", "Below zero: no low outliers"],
             ["Orders above the fence", "1,564"], ["Their share of revenue", "72.4%"]],
      so="Removing these 'outliers' would delete 72% of revenue. They're real high-value orders, not errors.",
      split=4.90),

    O("4.1.12",
      what="The **Analysis ToolPak** is a free Excel add-in for ready-made statistics. Turn it on once: "
           "**File > Options > Add-ins > Excel Add-ins > Go**, tick **Analysis ToolPak**. A **Data Analysis** "
           "button appears on the Data tab.",
      like="It's like a toolbox that came with your house but was locked in the store-room: unlock it once and it's always there.",
      how=["File > Options > Add-ins.", "Manage: Excel Add-ins > Go.", "Tick Analysis ToolPak > OK."],
      tip="On a Mac: Tools > Excel Add-ins.",
      trap="Looking for Data Analysis before enabling the add-in.",
      case="A learner can't find Data Analysis on the Data tab.",
      steps=["Enable the ToolPak.", "Data > Data Analysis.", "Check the list of tools."],
      table=[["Tool", "Gives", "Module"], ["Descriptive Statistics", "Full summary table", "4.1"],
             ["Histogram", "Frequency table and chart", "4.2"], ["Correlation", "Correlation matrix", "4.3"],
             ["Regression", "Full regression output", "4.4"]],
      so="One setting unlocks the tools used throughout Week 4.",
      split=4.20),

    O("4.1.13",
      what="**Data > Data Analysis > Descriptive Statistics** (tick Summary statistics) gives 16 lines: mean, "
           "standard error, median, mode, SD, variance, kurtosis, skewness, range, min, max, sum, count and a 95% "
           "confidence level. The output is **static**: rerun it when data changes.",
      like="It's like a medical check-up printout: all the main readings on one page.",
      how=["Data > Data Analysis > Descriptive Statistics.", "Select the column with its header; tick Labels.",
           "Tick Summary statistics and Confidence Level."],
      tip="Compare the output with live formulas once, so you trust both.",
      trap="Forgetting the output doesn't update when data changes.",
      case="The analyst runs Descriptive Statistics on salary and compares it with the live replica on sheet 4.1.",
      steps=["Run the ToolPak on Annual_Salary_NGN.", "Read the key lines.", "Compare with the replica table."],
      table=[["Output line", "Salary", "Check"], ["Mean / Median", "₦11,344,180 / ₦11,000,000", "Close: balanced"],
             ["Standard Error", "₦121,717", "SD ÷ √1,781"], ["Count", "1,781", "Matches the rows"],
             ["Confidence Level (95%)", "₦238,724", "Mean ± this"]],
      so="The average salary is ₦11.34m, give or take about ₦0.24m at 95% confidence.",
      split=4.20),

    O("4.1.14",
      what="Report statistics **honestly**: give the **sample size (n)**, the right **centre** for the shape, a "
           "**spread**, sensible rounding, and what was **excluded**. An average without these can't be checked.",
      like="It's like a weather forecast: '30°C' means little without 'Lagos, tomorrow afternoon, ±2°C'.",
      how=["State n.", "Pick median or mean based on the shape; add a spread.", "Round for readers; state exclusions."],
      tip="One sentence can carry n, centre, spread and scope.",
      trap="Averages of tiny groups, like 7 'Other' respondents.",
      case="The analyst writes summary sentences for the HR report.",
      steps=["Write the salary sentence.", "Write the order value sentence.", "Check each has n, centre, spread and scope."],
      table=[["Rule", "Example sentence"], ["n + centre + spread", "Median salary ₦11.0m (IQR ₦7m–₦16m), n = 1,781"],
             ["Skewed money", "Typical online order ₦29,300 (median); mean ₦126,792"],
             ["Small groups", "Only 7 'Other' respondents: no average reported"], ["Exclusions", "Cancelled orders included (gross bookings)"]],
      so="Each sentence can be checked and can't be misread.",
      split=4.90),
]
