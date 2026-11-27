from revlib import T

TRY = "ETEH_Week4_Section_Examples.xlsx  ›  sheet '4.2 Distributions & Z'"
CV = "Crestview University"
RT = "SuperMart Retail"
LC = "Lifecare Hospital"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    return T(tag, **k)


TOPICS = [
    O("4.2.1",
      what="A **distribution** shows **how often each value occurs**. Its shape (balanced, lopsided, flat or "
           "peaked) decides which summary is honest. Look at the shape before you choose mean or median.",
      like="It's like looking at a crowd from above: are people bunched in the middle, spread evenly, or "
           "mostly at one end?",
      how=["Build a histogram (4.2.3).", "Describe the shape in words.", "Pick the summary that suits it (4.2.9)."],
      tip="Check SKEW and a histogram together.",
      trap="Choosing the summary before seeing the shape.",
      who="Multi-Sector",
      case="The analyst compares the shape of three columns from different datasets.",
      steps=["Histogram of each column.", "Read the shape.", "Choose the summary."],
      table=[["Column", "Shape", "Honest summary"], ["Online order value", "Long right tail (skew 4.61)", "Median + IQR"],
             ["Exam score (Crestview)", "Tall peak at 67, thin tail to 100", "Median + percentiles"],
             ["Attendance % (Crestview)", "Flat from 60% to 100%", "Mean ± SD"]],
      so="Three columns, three shapes, three different honest summaries.",
      split=4.20),

    O("4.2.2",
      what="The **normal distribution** is the balanced bell curve. In a normal shape about **68%** of values sit "
           "within 1 SD of the mean, **95%** within 2 SD and **99.7%** within 3 SD. Real data is only roughly normal, so check.",
      like="Like heights of adults: most are near average, fewer are very short or very tall, and it's balanced on both sides.",
      how=["Mean and SD.", "COUNTIFS between mean − k×SD and mean + k×SD.", "Compare with 68 / 95 / 99.7%."],
      tip="If the shares don't match, don't use 'mean ± 2 SD' rules.",
      trap="Assuming every column is a bell curve.",
      who=CV,
      case="Is the exam score (mean 67.2, SD 3.9) close enough to normal to set grade bands by SD?",
      steps=["Count scores within ±1, ±2 and ±3 SD.", "Divide by 6,607.", "Compare with the normal rule."],
      table=[["Interval", "Normal expects", "Exam scores"], ["Mean ± 1 SD", "68%", "75.4%"],
             ["Mean ± 2 SD", "95%", "97.8%"], ["Mean ± 3 SD", "99.7%", "99.2%"]],
      so="Scores bunch more tightly than a bell curve, yet have more extreme highs. Use percentiles for grade bands.",
      split=4.70),

    O("4.2.3",
      what="A **histogram** groups a number column into **bins** and counts each bin, showing the shape. Use "
           "**Insert > Histogram**, the ToolPak, or COUNTIFS per bin. Too few bins hide detail; too many look noisy.",
      like="Like sorting exam scripts into piles by score band and comparing pile heights.",
      how=["Choose equal bins (here 2 points wide).", "Count each bin with COUNTIFS.", "Plot as columns with no gaps."],
      tip="Overlay the normal curve to see where the data differs.",
      trap="Unequal bin widths, which distort the shape.",
      who=CV,
      case="The analyst bins 6,607 exam scores and compares each bin with what a normal curve expects (tHist).",
      steps=["Bins 55–56 up to 79+.", "=COUNTIFS for each bin.", "Normal_expected from NORM.DIST."],
      table=[["Bin", "Students", "Normal expects"], ["63–64", "872", "854"], ["67–68", "1,476", "1,338"],
             ["75–76", "64", "147"], ["79+", "51", "12"]],
      so="The centre is taller than normal, and the top bin has 4 times more students than expected.",
      split=4.70),

    O("4.2.4",
      what="**Pareto analysis** ranks items by contribution and adds a **running % of the total**. It finds the "
           "'vital few'. The '80/20 rule' is a rule of thumb; measure the real split.",
      like="Like finding that a few friends make most of the calls on your phone bill.",
      how=["Sort customers by sales, largest first.", "Running total ÷ grand total = Cum_Share.",
           "Count rows until Cum_Share reaches 80%."],
      tip="Report 'top N = X% of sales', not just '80/20'.",
      trap="Assuming the split is 80/20 without checking.",
      who=RT,
      case="How many SuperMart customers make up 80% of FY2025 sales of ₦733,215,255? (D_Pareto_FY25)",
      steps=["Count customers: 693.", "=COUNTIFS(tPareto[Cum_Share],\"<0.8\")+1.", "Top 10 share with SUMIFS."],
      table=[["Measure", "Result"], ["Customers who bought in FY2025", "693"], ["Customers needed for 80% of sales", "260 (37.5%)"],
             ["Top 20% of customers (139)", "60.2% of sales"], ["Top 10 customers", "12.0% of sales"]],
      so="Sales are spread wider than '80/20': it takes 37.5% of customers to reach 80% of sales.",
      split=4.90),

    O("4.2.5",
      what="A **percentile rank** says what share of values sit **below** a value (=PERCENTRANK.INC, 0 to 100%). "
           "Group ranks into **bands** (top quartile, and so on) for fair league tables.",
      like="Like saying 'you finished ahead of 82% of runners' instead of 'you came 3rd'.",
      how=["=PERCENTRANK.INC([Sales],[@Sales]).", "Bands with IFS: >=75% top quartile, >=50%, >=25%.", "Sort by rank."],
      tip="Percentile ranks work the same for 12 stores or 12,000 customers.",
      trap="Reading percentile rank as 'percentage of target'.",
      who=RT,
      case="Rank the 12 stores by FY2025 sales and put each in a band (tStoreZ).",
      steps=["PERCENTRANK.INC on Sales.", "IFS to assign the band.", "Read the top and bottom."],
      table=[["Store", "Sales", "Percentile rank", "Band"], ["Lagos - Ikeja", "₦146,388,345", "100%", "Top quartile"],
             ["Port Harcourt - GRA", "₦93,922,995", "90.9%", "Top quartile"], ["Benin City - Ring Road", "₦62,116,000", "63.6%", "Upper middle"],
             ["Abuja - Wuse", "₦43,421,758", "27.2%", "Lower middle"], ["Owerri - New Owerri", "₦31,585,710", "0%", "Bottom quartile"]],
      so="Owerri has the lowest sales of the 12 stores; Ikeja has the highest.",
      split=4.20),

    O("4.2.6",
      what="A **z-score** says how many **SDs a value is from the mean**: =STANDARDIZE(x, mean, SD), or "
           "(x − mean) ÷ SD. 0 = average, +2 = well above, −2 = well below. Above 3 in size is rare.",
      like="Like saying how unusually tall someone is for their own group, not in centimetres.",
      how=["Mean and STDEV.S.", "=STANDARDIZE(value, mean, SD).", "Flag |z| > 2 for a closer look."],
      tip="z-scores let you compare a salary with an exam score.",
      trap="Trusting z-rules on very skewed data.",
      who="Multi-Sector",
      case="How unusual are these values within their own column?",
      steps=["Exam: mean 67.2, SD 3.9.", "Salary: mean ₦11,344,180, SD ₦5,136,701.", "STANDARDIZE each value."],
      table=[["Value", "z-score", "Reading"], ["Exam score 75", "2.00", "Well above average"],
             ["Exam score 100", "8.42", "Extremely rare: check it"], ["Salary ₦25,000,000", "2.66", "Top of the survey"],
             ["Salary ₦2,500,000", "-1.72", "Low, but not extreme"]],
      so="A score of 100 is 8 SDs above average; confirm it's real before using it.",
      split=4.20),

    O("4.2.7",
      what="To compare things on **several measures with different units** (₦, %, days), turn each into a "
           "**z-score** and average them into a **composite**. State the weights you chose.",
      like="Like a decathlon: seconds and metres are converted to points before adding.",
      how=["z-score each measure.", "Composite = AVERAGE of the z-scores (or a weighted mix).", "Rank with RANK.EQ."],
      tip="Show the separate z-scores too, so readers see why a store ranks where it does.",
      trap="Adding ₦ sales to % margin directly.",
      who=RT,
      case="Rank the 12 stores on sales, margin and target attainment with equal weights (tStoreZ).",
      steps=["z_Sales, z_Margin, z_Attain.", "Composite_z = average of the three.", "Overall_Rank."],
      table=[["Store", "Rank on sales", "Composite z", "Overall rank"], ["Lagos - Ikeja", "1", "1.30", "1"],
             ["Jos - Rayfield", "4", "0.81", "3"], ["Port Harcourt - GRA", "2", "0.65", "4"],
             ["Owerri - New Owerri", "12", "-0.44", "8"], ["Abuja - Wuse", "9", "-1.15", "12"]],
      so="Owerri is last on sales but has the best margin, so it rises to 8th; Abuja - Wuse is last overall.",
      split=4.20),

    O("4.2.8",
      what="Find anomalies **systematically**, not by eye. Combine three checks: **statistical** (|z| > 2), "
           "**business rules** (impossible values) and **comparison** (against last period). Flagged rows are investigated, not deleted.",
      like="Like a bank's fraud alerts: unusual payments are flagged automatically, then a person checks them.",
      how=["Add a z column with STANDARDIZE.", "Flag with =IF(ABS(z)>2,\"Anomaly\",\"\").", "Count flags and look for clusters."],
      tip="Clustered flags usually mean a real event; single flags can be typos.",
      trap="Deleting flagged rows without asking why.",
      who=LC,
      case="Which days in 2023 had unusual bed occupancy? (tCensus23, 365 days)",
      steps=["Mean 89.5%, SD 4.5 points.", "z and Flag columns.", "Count and date the flags."],
      table=[["Measure", "Result"], ["Days with |z| > 2", "10 of 365"], ["Direction", "All 10 are low days (80.4% or below)"],
             ["When", "9 of the 10 fall between 28 Oct and 12 Nov"], ["Lowest day", "78.85% on 11 Nov 2023"]],
      so="The flags cluster in early November: ask operations what happened, rather than treating them as errors.",
      split=4.90),

    O("4.2.9",
      what="Match the summary to the **shape** and the **question**: mean ± SD for balanced data, median + IQR "
           "for lopsided money, percentile bands for rankings, z-scores for mixed units, Pareto for concentration.",
      like="Like choosing the right tool from a toolbox: a hammer is great, but not for screws.",
      how=["Check the shape.", "Ask what the reader wants to know.", "Use the matching summary."],
      tip="When unsure, report the median and IQR alongside the mean.",
      trap="Using the mean for everything.",
      who="Multi-Sector",
      case="The analyst picks summaries for five report questions.",
      steps=["Read each question.", "Check the data shape.", "Pick the summary."],
      table=[["Question / shape", "Bootcamp example", "Use"], ["Balanced, no outliers", "Crestview attendance %", "Mean ± SD"],
             ["Lopsided money", "Online order value, hospital bills", "Median + IQR"], ["League table", "12 stores", "Percentile bands"],
             ["Different units", "Sales ₦, margin %, attainment %", "Composite z-score"], ["Unusual days", "Bed occupancy", "z-score flags"]],
      so="The right summary depends on the shape and the question, not on habit.",
      split=4.20),
]
