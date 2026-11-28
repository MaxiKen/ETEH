"""Week 4 Section Examples (Edition 3.1): questions in teaching order.

Progression rule for Week 4 (Weeks 1-3 are assumed)
    4.1 - 4.2   descriptive statistics, distributions, percentiles, z-scores (one function per question)
    4.3 - 4.4   correlation and simple regression (CORREL, SLOPE, INTERCEPT, RSQ, FORECAST.LINEAR)
    4.5         time series: moving average and a linear trend forecast (ETS ideas as multiple choice)
    4.6         what-if: a three-cell profit model for Goal Seek, data tables, scenarios and Solver
    4.7 - 4.9   storytelling, reports, slides (ideas only)
    4.10        auditing, protection, macros
    4.11 - 4.13 AI co-pilots, verification, ethics
    4.14        internal benchmark lab: sales per m², the PDF bulletin check, the memo
"""
from __future__ import annotations

import numpy as np
import pandas as pd

import w4_data as W
import w4_m11_14 as F14
from common import W2D


def build(b):
    sal = W.table("tSal").copy()
    sal["Age"] = sal.Age.astype(int)
    sal["Years_of_Experience"] = sal.Years_of_Experience.astype(int)
    cv = W.table("tCV")
    mon = W.table("tMonthly")
    dash = W.table("tDash25")
    st = W.table("tStores").copy()
    st["Sales_2025"] = st.Store_ID.map(dash.groupby("Store_ID").Sales.sum()).astype("int64")

    M = {}

    def mod(code, sheet, title, learned):
        M[code] = b.module(code, sheet, title, learned)
        return M[code]

    m41 = mod("4.1", "4.1 Descriptive Stats", "Descriptive Statistics and the Analysis ToolPak",
              "Mean, median, mode, range, IQR, variance, standard deviation, coefficient of variation, skewness, "
              "kurtosis, percentiles and quartiles, the 1.5 × IQR outlier rule and the ToolPak Descriptive Statistics.")
    m42 = mod("4.2", "4.2 Distributions", "Distributions, Pareto, Percentiles and Z-Scores",
              "Distribution shape, the normal curve, histograms, Pareto (80/20), percentile rank, z-scores and "
              "standardising measures, and spotting anomalies.")
    m43 = mod("4.3", "4.3 Correlation", "Correlation Analysis",
              "Pearson and Spearman correlation, correlation matrices, reading scatterplots first, correlation vs "
              "causation, confounders and Simpson's paradox.")
    m44 = mod("4.4", "4.4 Regression", "Linear Regression",
              "Simple and multiple regression, the equation and coefficients, R², adjusted R², p-values, residuals, "
              "prediction, multicollinearity and limits.")
    m45 = mod("4.5", "4.5 Forecasting", "Time Series and Forecasting",
              "Trend, seasonality, moving averages, exponential smoothing, FORECAST.ETS and its confidence interval, "
              "forecast accuracy (MAE, MAPE) and backtesting.")
    m46 = mod("4.6", "4.6 What-If & Solver", "Goal Seek, Data Tables, Scenarios and Solver",
              "Goal Seek, one- and two-variable data tables, Scenario Manager and summaries, and Solver "
              "(objective, variable cells, constraints).")
    m47 = mod("4.7", "4.7 Storytelling", "Data Storytelling",
              "The three-act narrative, the headline message, message hierarchy, evidence for the audience, titles that "
              "state the finding, annotation and avoiding misleading charts.")
    m48 = mod("4.8", "4.8 Reports & PDF", "Report Writing and PDF Production",
              "Report structure, the executive summary, layout, table design, page setup, PDF export, metadata and "
              "version control.")
    m49 = mod("4.9", "4.9 Presentations", "Slide Presentations",
              "One message per slide, charts for projection, text density, visual identity, timing and handling questions.")
    m410 = mod("4.10", "4.10 Auditing & Macros", "Formula Auditing, Protection and Light Automation",
               "Trace precedents and dependents, Show Formulas, Evaluate Formula, Error Checking, the Watch Window, "
               "protecting sheets, recording macros and Office Scripts.")
    m411 = mod("4.11", "4.11 AI Co-pilots", "AI Co-pilots for Analysts",
               "What AI can and cannot do, prompt structure, schema prompts, AI for formulas, SQL and narrative, and "
               "what must never be pasted into an AI tool.")
    m412 = mod("4.12", "4.12 Verifying AI", "Verifying AI Output and Handling Hallucinations",
               "Hallucinations, the verification protocol, control totals, plausible-but-wrong output, provenance and sign-off.")
    m413 = mod("4.13", "4.13 Data Ethics", "Data Ethics, PII and PHI",
               "Personal and health data, minimisation, anonymisation and aggregation, consent and purpose, bias, "
               "retention and governance.")
    m414 = mod("4.14", "4.14 Benchmark Lab", "Internal Benchmarks, PDF Checks and Reconciliation",
               "Sales per m² against the chain median, checking the PDF bulletin against the ledger, periods and "
               "definitions, uncertainty and the decision memo.")

    r_sal = b.data("Salary", sal, "One salary-survey respondent (Nigerian tech and business roles)",
                   fmts={"Annual_Salary_NGN": "naira"}, used_in="4.1, 4.2")
    r_st = b.data("Students", cv, "One Kingsway University student (study habits and final score)", used_in="4.2, 4.3, 4.4")
    r_mo = b.data("Monthly", mon, "One month of SuperMart chain sales, Jan 2022 - Dec 2025 (t = month number)",
                  fmts={"Month": "date", "Sales": "naira"}, used_in="4.5")
    r_sb = b.data("StoreBench", st, "One SuperMart store: floor area, 2025 target and 2025 sales",
                  fmts={"Annual_Target_2025": "naira", "Sales_2025": "naira"}, used_in="4.14")
    for m_, refs in ((m41, [r_sal]), (m42, [r_sal, r_st]), (m43, [r_st]), (m44, [r_st]), (m45, [r_mo]), (m414, [r_sb])):
        m_.set_refs(refs)
    f414 = m414.link_files([W2D / "SuperMart_Monthly_Sales_Bulletin_2025.pdf"])

    S = "tSalary[Annual_Salary_NGN]"
    x = sal.Annual_Salary_NGN
    # ================================================================== 4.1
    m = m41
    m.example("the three averages", ["Each formula reads one column of tSalary on the Data_Salary sheet.",
                                     "Compare them: when the mean and median are far apart, the data is skewed."],
              [("Mean salary", f"=AVERAGE({S})", f"=AVERAGE({S})", float(x.mean()), "naira"),
               ("Median salary", f"=MEDIAN({S})", f"=MEDIAN({S})", float(x.median()), "naira")])
    m.part("Part A · Ideas")
    m.mc("Which average is least affected by a few very large values?", ["Mean", "Median", "Maximum", "Sum"], "B",
         "The median is the middle value, so extremes barely move it.")
    m.mc("What does the standard deviation tell you?", ["The middle value", "How spread out values are around the mean",
                                                       "The most common value", "The number of rows"], "B",
         "Small SD = values close to the mean.")
    m.mc("The interquartile range (IQR) is ...", ["Q3 − Q1", "Max − Min", "Mean − Median", "Q2 × 2"], "A",
         "The spread of the middle 50% of values.")
    m.mc("Using the 1.5 × IQR rule, a value is an outlier if it is ...",
         ["above Q3 + 1.5 × IQR or below Q1 − 1.5 × IQR", "above the mean", "equal to the median", "negative"], "A",
         "The fences are Q1 − 1.5 IQR and Q3 + 1.5 IQR.")
    m.tf("The Analysis ToolPak must be switched on (File › Options › Add-ins) before Data › Data Analysis appears.", True,
         "Enable it once per computer.")
    m.part("Part B · Work it out on tSalary")
    m.num("What is the most common salary (mode)?", f"=MODE.SNGL({S})", float(x.mode()[0]), "MODE.SNGL returns the most frequent value.",
          fmt="naira")
    m.num("What is the range (highest minus lowest salary)?", f"=MAX({S})-MIN({S})", float(x.max() - x.min()),
          "MAX − MIN.", fmt="naira")
    q1, q3 = float(x.quantile(.25)), float(x.quantile(.75))
    m.num("What is the first quartile (Q1)?", f"=QUARTILE.INC({S},1)", q1, "QUARTILE.INC(range, 1).", fmt="naira")
    m.num("What is the interquartile range (Q3 − Q1)?", f"=QUARTILE.INC({S},3)-QUARTILE.INC({S},1)", q3 - q1,
          "Q3 − Q1.", fmt="naira")
    m.num("What is the sample standard deviation? (nearest naira)", f"=ROUND(STDEV.S({S}),0)", round(float(x.std())),
          "STDEV.S for a sample.", fmt="naira")
    cvv = float(x.std() / x.mean())
    m.num("Coefficient of variation = standard deviation ÷ mean. What is it? (as a %)", f"=STDEV.S({S})/AVERAGE({S})", cvv,
          "SD ÷ mean.", fmt="pct")

    # ================================================================== 4.2
    m = m42
    sc = cv.Final_Score
    m.part("Part A · Ideas")
    m.mc("In a normal distribution, about what share of values lie within 1 standard deviation of the mean?",
         ["50%", "68%", "95%", "99.7%"], "B", "68-95-99.7 rule.")
    m.mc("What is a z-score?", ["A rank", "How many standard deviations a value is above or below the mean",
                                "A percentage", "A count"], "B", "z = (value − mean) ÷ SD.")
    m.mc("Pareto analysis (the 80/20 rule) asks ...", ["which few items make up most of the total", "the average",
                                                       "the mode", "the trend"], "A",
         "Sort largest first and add a cumulative % column.")
    m.pick("Salaries with a long tail of high earners", ["Right (positive) skew", "Left (negative) skew", "Symmetric"],
           "Right (positive) skew", "Long tail to the right: mean above median.")
    m.part("Part B · Work it out on tStudents (and tSalary for the last question)")
    m.num("What is the mean Final_Score? (1 decimal place)", "=AVERAGE(tStudents[Final_Score])", float(sc.mean()),
          "AVERAGE.", fmt="dec1")
    sd = float(sc.std())
    m.num("A student scored 80. What is the z-score? (2 decimal places)",
          "=STANDARDIZE(80,AVERAGE(tStudents[Final_Score]),STDEV.S(tStudents[Final_Score]))",
          (80 - sc.mean()) / sd, "z = (80 − mean) ÷ SD.", fmt="dec2")
    m.num("What is the 90th percentile of Final_Score?", "=PERCENTILE.INC(tStudents[Final_Score],0.9)",
          float(sc.quantile(.9)), "90% of students scored at or below this.", fmt="dec1")
    m.num("How many students scored 2 or more standard deviations above the mean? (hint: first work out mean + 2 × SD)",
          '=COUNTIFS(tStudents[Final_Score],">="&(AVERAGE(tStudents[Final_Score])+2*STDEV.S(tStudents[Final_Score])))',
          int((sc >= sc.mean() + 2 * sd).sum()), "COUNTIFS with the cut-off mean + 2 SD.")
    top = sal.groupby("Job_Family").size().sort_values(ascending=False)
    m.num("Pareto: what share of respondents work in the largest Job_Family (Software Engineering)?",
          '=COUNTIFS(tSalary[Job_Family],"Software Engineering")/ROWS(tSalary)', float(top.iloc[0] / len(sal)),
          f"{top.iloc[0]} of {len(sal):,}.", fmt="pct")

    # ================================================================== 4.3
    m = m43
    m.part("Part A · Ideas")
    m.mc("A correlation of r = −0.9 means ...", ["no relationship", "a strong negative relationship",
                                                  "a weak positive relationship", "one causes the other"], "B",
         "Close to −1: as one goes up, the other goes down.")
    m.mc("What should you look at BEFORE trusting r?", ["The scatterplot", "The font", "The file size", "The sheet name"], "A",
         "A curve or a single outlier can fool r.")
    m.yn("If ice-cream sales and drownings rise together, does ice cream cause drowning?", "No",
         "A confounder (hot weather) drives both.")
    m.mc("Spearman correlation is used when ...", ["data are ranks or the relationship is monotonic but not straight",
                                                   "there are no numbers", "r is exactly 1", "you have one column"], "A",
         "Spearman works on ranks.")
    m.part("Part B · Work it out on tStudents")
    for col, lab in (("Study_Hours_Wk", "Study_Hours_Wk"), ("Lecture_Attendance_Pct", "Lecture_Attendance_Pct"),
                     ("Sleep_Hrs_Night", "Sleep_Hrs_Night")):
        m.num(f"What is the correlation between {lab} and Final_Score? (2 decimal places)",
              f"=CORREL(tStudents[{col}],tStudents[Final_Score])", float(np.corrcoef(cv[col], sc)[0, 1]),
              "CORREL(x, y).", fmt="dec2")
    m.pick("Which of the three has the strongest relationship with Final_Score?",
           ["Study_Hours_Wk", "Lecture_Attendance_Pct", "Sleep_Hrs_Night"], "Lecture_Attendance_Pct",
           "Largest r (furthest from 0).")

    # ================================================================== 4.4
    m = m44
    xa, ya = cv.Lecture_Attendance_Pct.astype(float), sc.astype(float)
    sl, ic = np.polyfit(xa, ya, 1)
    r2 = float(np.corrcoef(xa, ya)[0, 1] ** 2)
    X, Y = "tStudents[Lecture_Attendance_Pct]", "tStudents[Final_Score]"
    m.part("Part A · Ideas")
    m.mc("In Score = 25 + 0.5 × Attendance, what does 0.5 mean?", ["The score with zero attendance",
                                                                  "Each extra attendance point adds 0.5 to the predicted score",
                                                                  "R²", "The error"], "B", "The slope (coefficient).")
    m.mc("R² = 0.22 means ...", ["22% of the variation in y is explained by the model", "the model is 22% wrong",
                                 "y rises by 0.22", "22 rows"], "A", "R² is the share of variation explained.")
    m.mc("A p-value below 0.05 for a coefficient usually means ...", ["the coefficient is statistically significant",
                                                                       "the model is perfect", "remove it", "R² is 0.05"],
         "A", "Unlikely to be zero by chance.")
    m.yn("Is it safe to predict far outside the range of the data used to fit the model?", "No",
         "Extrapolation: the relationship may not hold there.")
    m.part("Part B · Fit Final_Score on Lecture_Attendance_Pct")
    m.num("What is the slope? (3 decimal places)", f"=SLOPE({Y},{X})", float(sl), "SLOPE(y, x).", fmt="dec3")
    m.num("What is the intercept? (2 decimal places)", f"=INTERCEPT({Y},{X})", float(ic), "INTERCEPT(y, x).", fmt="dec2")
    m.num("What is R²? (as a %)", f"=RSQ({Y},{X})", r2, "RSQ(y, x) = r².", fmt="pct")
    m.num("Predict the Final_Score of a student with 95% attendance. (1 decimal place)", f"=FORECAST.LINEAR(95,{Y},{X})",
          float(ic + sl * 95), "intercept + slope × 95.", fmt="dec1")

    # ================================================================== 4.5
    m = m45
    ms = mon.Sales.astype(float)
    m.part("Part A · Ideas")
    m.mc("Seasonality means ...", ["a pattern that repeats every year (or week)", "a one-off spike", "the overall trend",
                                   "random noise"], "A", "e.g. December peaks every year.")
    m.mc("A 3-month moving average ...", ["smooths the series by averaging each month with the two before it",
                                          "predicts the next 3 years", "removes the trend", "counts months"], "A",
         "It smooths short-term ups and downs.")
    m.mc("FORECAST.ETS.CONFINT gives ...", ["the width of the uncertainty band around a forecast", "the trend",
                                            "the season length", "the actual value"], "A", "Forecast ± CONFINT = the band.")
    m.mc("MAPE = 8% means ...", ["forecasts were off by 8% on average", "8 forecasts", "R² is 8%", "growth of 8%"], "A",
         "Mean absolute percentage error.")
    m.yn("Should you test a forecast on past months it did not see (backtesting) before trusting it?", "Yes",
         "Hold out the last months, forecast them, compare.")
    m.part("Part B · Work it out on tMonthly")
    m.num("What were total sales in 2025?", '=SUMIFS(tMonthly[Sales],tMonthly[Year],2025)',
          float(mon.loc[mon.Year == 2025, "Sales"].sum()), "SUMIFS by Year.", fmt="naira")
    m.num("What is the 3-month moving average for Dec 2025 (Oct, Nov, Dec 2025)? (nearest naira)",
          '=ROUND(AVERAGEIFS(tMonthly[Sales],tMonthly[t],">=46"),0)', round(float(ms.iloc[45:48].mean())),
          "Average of months t = 46, 47 and 48 (AVERAGEIFS with t >= 46).", fmt="naira")
    m.num("Using the straight-line trend, forecast sales for month t = 49 (Jan 2026). (nearest naira)",
          "=ROUND(FORECAST.LINEAR(49,tMonthly[Sales],tMonthly[t]),0)",
          round(float(np.polyval(np.polyfit(mon.t, ms, 1), 49))), "FORECAST.LINEAR(49, sales, t).", fmt="naira")
    g = float(mon.loc[mon.Year == 2025, "Sales"].sum() / mon.loc[mon.Year == 2024, "Sales"].sum() - 1)
    m.num("Growth from 2024 to 2025 (as a %)?",
          '=SUMIFS(tMonthly[Sales],tMonthly[Year],2025)/SUMIFS(tMonthly[Sales],tMonthly[Year],2024)-1', g,
          "2025 ÷ 2024 − 1.", fmt="pct")

    # ================================================================== 4.6
    m = m46
    m.info("Model", "Price ₦2,500 per unit, Units sold 4,000, Cost ₦1,800 per unit. "
           "Profit = (Price − Cost) × Units.")
    m.part("Part A · Which tool?")
    T = ["Goal Seek", "Data Table", "Scenario Manager", "Solver"]
    m.pick("What price gives a profit of exactly ₦4,000,000?", T, "Goal Seek", "One input, one target.")
    m.pick("Show profit for prices ₦2,000 to ₦3,000 in steps of ₦100.", T, "Data Table", "One input, many values.")
    m.pick("Compare a Best, Base and Worst case, each changing several inputs.", T, "Scenario Manager",
           "Named sets of inputs with a summary report.")
    m.pick("Choose units of three products to maximise profit within a budget.", T, "Solver",
           "An objective, variable cells and constraints.")
    m.part("Part B · Work it out")
    m.num("What is the profit in the model?", "=(2500-1800)*4000", 2_800_000, "(2,500 − 1,800) × 4,000.", fmt="naira")
    m.num("Goal Seek: what price gives a profit of ₦4,000,000?", "=1800+4000000/4000", 2800,
          "Price = Cost + Target ÷ Units = 1,800 + 1,000.", fmt="naira")
    m.num("Data table row: what is the profit at a price of ₦2,200?", "=(2200-1800)*4000", 1_600_000,
          "(2,200 − 1,800) × 4,000.", fmt="naira")
    m.num("Worst case: price ₦2,300, units 3,500, cost ₦1,900. What is the profit?", "=(2300-1900)*3500", 1_400_000,
          "(2,300 − 1,900) × 3,500.", fmt="naira")
    m.tf("Solver is an add-in that must be enabled before it appears on the Data tab.", True, "File › Options › Add-ins.")

    # ================================================================== 4.7
    m = m47
    m.part("Storytelling")
    m.mc("Which is the best headline for a report slide?", ["Sales chart", "Q1 data",
                                                            "Five stores made a loss in 2025; three are in the North",
                                                            "Analysis"], "C", "A title that states the finding.")
    m.mc("The three acts of a data story are ...", ["Setup, conflict, resolution", "Charts, tables, text",
                                                    "Start, middle, appendix", "Data, SQL, DAX"], "A",
         "Situation → problem → recommended action.")
    m.yn("Should a column chart's axis be cut (not start at zero) to make a small difference look large?", "No",
         "That misleads the reader.")
    m.mc("An annotation on a chart ...", ["points out the one thing the reader must notice", "is decoration",
                                          "replaces the title", "hides the data"], "A", "e.g. an arrow: 'Promotion week'.")
    m.tf("Different audiences (board vs analysts) need different levels of detail.", True, "Choose evidence for the audience.")

    # ================================================================== 4.8
    m = m48
    m.part("Reports and PDF")
    m.mc("Where does the executive summary go?", ["At the end", "At the start, on one page", "In the appendix",
                                                  "In the footer"], "B", "Busy readers may read nothing else.")
    m.mc("Before exporting to PDF you should set ...", ["the print area, orientation and page breaks", "a password",
                                                        "the macro security", "the zoom to 400%"], "A",
         "Page Layout › Print Area / Orientation / Breaks.")
    m.yn("Should you remove personal metadata (author, comments) before sending a report outside the company?", "Yes",
         "File › Info › Check for Issues › Inspect Document.")
    m.mc("Which file name follows good version control?", ["report final FINAL2.xlsx", "Mercyland_Report_2025-11_v03.pdf",
                                                           "new.pdf", "copy of copy.pdf"], "B",
         "Name_Topic_Date_version.")

    # ================================================================== 4.9
    m = m49
    m.part("Slides")
    m.mc("Each slide should carry ...", ["one message", "every chart you made", "as much text as fits", "no title"], "A",
         "One message per slide, stated in the title.")
    m.mc("Charts for projection need ...", ["larger fonts and fewer series", "3-D effects", "small labels",
                                            "many colours"], "A", "Readable from the back of the room.")
    m.yn("For a 10-minute slot, is 40 slides a good plan?", "No", "About one slide per minute is a common guide.")
    m.tf("You should prepare backup slides for likely questions.", True, "Put them after the closing slide.")

    # ================================================================== 4.10
    m = m410
    m.part("Auditing and protection")
    m.mc("Which button shows the cells a formula uses?", ["Trace Precedents", "Trace Dependents", "Watch Window",
                                                          "Protect Sheet"], "A", "Formulas › Trace Precedents.")
    m.mc("Which shortcut shows every formula instead of its result?", ["Ctrl + `", "Ctrl + S", "F4", "Ctrl + T"], "A",
         "Formulas › Show Formulas (Ctrl + grave accent).")
    m.mc("Evaluate Formula lets you ...", ["step through a formula one part at a time", "delete formulas",
                                           "protect the sheet", "record a macro"], "A", "Formulas › Evaluate Formula.")
    m.yn("Does Protect Sheet stop people editing locked cells?", "Yes", "Unlock input cells first (Format Cells › Protection).")
    m.mc("A recorded macro is saved in which file type?", [".xlsx", ".xlsm", ".csv", ".pdf"], "B",
         ".xlsx cannot hold macros.")
    m.tf("Office Scripts run in Excel for the web and can be scheduled with Power Automate.", True, "VBA macros run in desktop Excel.")

    # ================================================================== 4.11
    m = m411
    m.part("AI co-pilots")
    m.mc("Which is the strongest prompt for a formula?", ["Write a formula",
                                                          "Table tSales has columns Region and Sales. Write an Excel formula "
                                                          "for total Sales where Region is 'North'.", "Help", "Excel?"], "B",
         "Give the table, columns and the exact goal.")
    m.mc("A schema prompt gives the AI ...", ["table and column names (no real rows)", "the full customer list",
                                              "passwords", "nothing"], "A", "Structure, not personal data.")
    m.yn("May you paste patient names and diagnoses into a public AI chatbot?", "No",
         "Never paste personal or health data into an external AI tool.")
    m.tf("AI can write SQL quickly, but you must still run and check it.", True, "Check row counts and totals.")

    # ================================================================== 4.12
    m = m412
    m.part("Verifying AI output")
    m.mc("A hallucination is ...", ["confident output that is wrong or made up", "a slow answer", "a chart",
                                    "a macro"], "A", "It often looks plausible.")
    m.mc("An AI says total 2025 sales were ₦733.2m. How do you check it?", ["Trust it",
                                                                           "Recalculate the control total from the source "
                                                                           "(e.g. =SUM(Sales))", "Ask the AI again",
                                                                           "Round it"], "B", "The control-total discipline.")
    m.yn("If AI output fails a check and you cannot see why, should you discard it?", "Yes",
         "Do the work yourself instead.")
    m.tf("The analyst who signs off is responsible for the numbers, not the AI tool.", True, "Responsibility and sign-off.")

    # ================================================================== 4.13
    m = m413
    m.part("Ethics, PII and PHI")
    D = ["PII", "PHI", "Neither"]
    m.pick("A customer's phone number", D, "PII", "Identifies a person.")
    m.pick("A patient's diagnosis linked to their name", D, "PHI", "Health information about an identifiable person.")
    m.pick("Total sales per state for 2025", D, "Neither", "Aggregated business data.")
    m.mc("Data minimisation means ...", ["collect and keep only the data you need for the purpose", "delete everything",
                                         "make data smaller files", "remove all numbers"], "A", "Less data, less risk.")
    m.yn("Is showing counts of fewer than 5 patients in a public report a privacy risk?", "Yes",
         "Small counts can identify people; aggregate or suppress them.")
    m.mc("Nigeria's data-protection law is the ...", ["Nigeria Data Protection Act (2023)", "Companies Act",
                                                       "Tax Act", "Electoral Act"], "A", "NDPA 2023.")

    # ================================================================== 4.14
    m = m414
    st["per"] = st.Sales_2025 / st.Floor_Area_sqm
    med = float(st.per.median())
    m.part("Part A · Benchmark on tStoreBench")
    m.num("Lagos - Ikeja: what are its 2025 sales per m²? (nearest naira)",
          '=ROUND(XLOOKUP("Lagos - Ikeja",tStoreBench[Store_Name],tStoreBench[Sales_2025])'
          '/XLOOKUP("Lagos - Ikeja",tStoreBench[Store_Name],tStoreBench[Floor_Area_sqm]),0)',
          round(float(st.set_index("Store_Name").per["Lagos - Ikeja"])), "Sales ÷ floor area.", fmt="naira")
    m.num("What is the chain median of sales per m²? (add a column Sales_per_sqm = [@Sales_2025]/[@Floor_Area_sqm] first; "
          "nearest naira)", "=ROUND(MEDIAN(tStoreBench[Sales_2025]/tStoreBench[Floor_Area_sqm]),0)", round(med),
          "MEDIAN of the new column.", fmt="naira")
    m.num("How many stores are below the chain median?",
          "=SUMPRODUCT(--(tStoreBench[Sales_2025]/tStoreBench[Floor_Area_sqm]<MEDIAN(tStoreBench[Sales_2025]/tStoreBench[Floor_Area_sqm])))",
          int((st.per < med).sum()), "With 12 stores, 6 are below the median.")
    low = st.sort_values("per").Store_Name.iloc[0]
    m.short("Which store has the lowest sales per m²? (type the Store_Name)", low, "Sort the new column smallest to largest.")
    m.part("Part B · Check the PDF bulletin")
    bul = F14.BUL
    m.num(f"Open {f414} › SuperMart_Monthly_Sales_Bulletin_2025.pdf. What FY figure (₦m) does it show for Lagos - Ikeja?",
          None, float(bul.fy["Lagos - Ikeja"]), "Read from the store table in the PDF.", fmt="dec1", do=True)
    m.num("Ledger check: Lagos - Ikeja 2025 sales in ₦m (2 decimal places)?",
          '=XLOOKUP("Lagos - Ikeja",tStoreBench[Store_Name],tStoreBench[Sales_2025])/1000000',
          float(bul.ledger["Lagos - Ikeja"]), "146,388,345 ÷ 1,000,000.", fmt="dec2")
    m.yn("The PDF rounds to ₦0.1m. Is a difference of ₦0.01m between the PDF and the ledger acceptable?", "Yes",
         "Within rounding (±₦0.05m).")
    m.mc("Which belongs in the decision memo?", ["Only a chart", "The finding, the evidence, the uncertainty and the action",
                                                 "Every formula", "The raw data"], "B", "Short, decision-focused.")
