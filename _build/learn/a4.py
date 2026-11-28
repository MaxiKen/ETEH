"""Week 4 assignment (Edition 3.1): statistics, forecasting, what-if and responsible reporting. Week 1-4 skills only.

    Task 1  descriptive statistics and outliers on the salary survey (4.1 - 4.2)
    Task 2  correlation and simple regression on Kingsway students (4.3 - 4.4)
    Task 3  forecasting Atlantic Pearl Hotel room nights (4.5)
    Task 4  what-if: a room-price model for Goal Seek and a data table (4.6)
    Task 5  storytelling, reports, auditing, AI and ethics (4.7 - 4.13)
"""
from __future__ import annotations

import numpy as np

from wk_common import L, hotel


def build(b):
    sal = L("sal_clean").copy()
    sal["Age"] = sal.Age.astype(int)
    sal["Years_of_Experience"] = sal.Years_of_Experience.astype(int)
    cv = L("cv_perf")[["Matric_No", "Faculty", "Level", "Last_Sem_Score", "Tutorials_Attended", "Study_Hours_Wk",
                       "Lecture_Attendance_Pct", "Final_Score"]].copy()
    h = hotel()
    st = h[h.Reservation_Status == "Stayed"]
    hm = st.groupby(st.Check_In.dt.to_period("M")).agg(Bookings=("Reservation_ID", "size"),
                                                       Room_Nights=("Nights", "sum")).reset_index()
    hm["Month"] = hm.Check_In.dt.to_timestamp()
    hm["Month_No"] = range(1, len(hm) + 1)
    hm = hm[["Month", "Month_No", "Bookings", "Room_Nights"]]

    t1 = b.module("T1", "T1 Descriptive Stats", "Descriptive statistics and outliers", "Modules 4.1 - 4.2: mean, median, "
                  "quartiles, IQR, standard deviation, the 1.5 × IQR rule, percentiles and z-scores.")
    t2 = b.module("T2", "T2 Correlation", "Correlation and regression", "Modules 4.3 - 4.4: CORREL, scatterplots, "
                  "causation, SLOPE, INTERCEPT, RSQ and FORECAST.LINEAR.")
    t3 = b.module("T3", "T3 Forecasting", "Forecasting", "Module 4.5: trend, seasonality, moving averages, linear trend "
                  "forecasts and forecast error.")
    t4 = b.module("T4", "T4 What-If", "What-if analysis", "Module 4.6: Goal Seek, data tables, scenarios and Solver.")
    t5 = b.module("T5", "T5 Report AI Ethics", "Reporting, AI and ethics", "Modules 4.7 - 4.13: storytelling, reports, "
                  "slides, auditing, AI co-pilots, verifying AI and data ethics.")
    r_s = b.data("Salary", sal, "One salary-survey respondent", fmts={"Annual_Salary_NGN": "naira"}, used_in="T1")
    r_c = b.data("Students", cv, "One Kingsway University student", used_in="T2")
    r_h = b.data("HotelMonthly", hm, "One month of stayed bookings at Atlantic Pearl Hotel (Month_No 1 = Jul 2023)",
                 fmts={"Month": "date"}, used_in="T3")
    t1.set_refs([r_s])
    t2.set_refs([r_c])
    t3.set_refs([r_h])

    # ================================================================== T1
    m = t1
    S = "tSalary[Annual_Salary_NGN]"
    x = sal.Annual_Salary_NGN
    q1, q3 = float(x.quantile(.25)), float(x.quantile(.75))
    fence = q3 + 1.5 * (q3 - q1)
    m.info("Scenario", "A recruitment firm wants a salary benchmark for tech and business roles in Nigeria, a study-habits "
                       "brief for Kingsway University and a demand forecast for Atlantic Pearl Hotel.")
    m.num("Median salary?", f"=MEDIAN({S})", float(x.median()), "MEDIAN.", fmt="naira")
    m.num("Third quartile (Q3)?", f"=QUARTILE.INC({S},3)", q3, "QUARTILE.INC(range, 3).", fmt="naira")
    m.num("Upper outlier fence = Q3 + 1.5 × IQR?", f"=QUARTILE.INC({S},3)+1.5*(QUARTILE.INC({S},3)-QUARTILE.INC({S},1))",
          fence, "Q3 + 1.5 × (Q3 − Q1).", fmt="naira")
    m.yn(f"Is a salary of ₦25,000,000 above the upper fence (so an outlier)?", "Yes" if 25_000_000 > fence else "No",
         f"The fence is ₦{fence:,.0f}.")
    da = sal[sal.Job_Family == "Data & Analytics"].Annual_Salary_NGN
    m.num("Median salary for the Data & Analytics job family? (filter the table, or use MEDIAN on the filtered column)",
          None, float(da.median()), "Filter Job_Family, then MEDIAN of the visible values (or a PivotTable).", fmt="naira",
          do=True)
    m.num("Sample standard deviation of salary? (nearest naira)", f"=ROUND(STDEV.S({S}),0)", round(float(x.std())),
          "STDEV.S.", fmt="naira")
    z = (20_000_000 - x.mean()) / x.std()
    m.num("z-score of a ₦20,000,000 salary? (2 decimal places)", f"=STANDARDIZE(20000000,AVERAGE({S}),STDEV.S({S}))",
          float(z), "(20,000,000 − mean) ÷ SD.", fmt="dec2")
    m.mc("The mean salary is above the median. The distribution is ...", ["right (positive) skewed",
                                                                         "left (negative) skewed", "perfectly normal",
                                                                         "uniform"], "A", "A tail of high salaries.")

    # ================================================================== T2
    m = t2
    X, Y = "tStudents[Last_Sem_Score]", "tStudents[Final_Score]"
    xa, ya = cv.Last_Sem_Score.astype(float), cv.Final_Score.astype(float)
    sl, ic = np.polyfit(xa, ya, 1)
    r = float(np.corrcoef(xa, ya)[0, 1])
    m.num("Correlation between Last_Sem_Score and Final_Score? (2 decimal places)", f"=CORREL({X},{Y})", r, "CORREL.",
          fmt="dec2")
    m.num("Correlation between Tutorials_Attended and Final_Score? (2 decimal places)",
          f"=CORREL(tStudents[Tutorials_Attended],{Y})",
          float(np.corrcoef(cv.Tutorials_Attended.astype(float), ya)[0, 1]), "CORREL.", fmt="dec2")
    m.num("Regression of Final_Score on Last_Sem_Score: slope? (3 decimal places)", f"=SLOPE({Y},{X})", float(sl),
          "SLOPE(y, x).", fmt="dec3")
    m.num("Intercept? (2 decimal places)", f"=INTERCEPT({Y},{X})", float(ic), "INTERCEPT(y, x).", fmt="dec2")
    m.num("R² (as a %)?", f"=RSQ({Y},{X})", r * r, "RSQ.", fmt="pct")
    m.num("Predicted Final_Score for a Last_Sem_Score of 70? (1 decimal place)", f"=FORECAST.LINEAR(70,{Y},{X})",
          float(ic + sl * 70), "intercept + slope × 70.", fmt="dec1")
    m.yn("With an R² this small, does Last_Sem_Score alone predict Final_Score well?", "No",
         "R² of about 1% means the line explains almost none of the variation; look for better predictors.")
    m.yn("Does this prove that a higher last-semester score CAUSES a higher final score?", "No",
         "Correlation is not causation: ability and habits drive both.")

    # ================================================================== T3
    m = t3
    rn = hm.Room_Nights.astype(float)
    n = len(hm)
    m.num("Total room nights over all months?", "=SUM(tHotelMonthly[Room_Nights])", int(rn.sum()), "SUM.")
    m.short("Which month had the fewest room nights? (type it as yyyy-mm, e.g. 2023-07)",
            hm.loc[rn.idxmin(), "Month"].strftime("%Y-%m"), "Sort Room_Nights smallest to largest.")
    m.num(f"3-month moving average of the last three months (Month_No {n - 2} to {n})? (nearest whole number)",
          f'=ROUND(AVERAGEIFS(tHotelMonthly[Room_Nights],tHotelMonthly[Month_No],">={n - 2}"),0)',
          round(float(rn.iloc[-3:].mean())), "Average of the last three months.")
    m.num(f"Linear trend forecast for Month_No {n + 1}? (nearest whole number)",
          f"=ROUND(FORECAST.LINEAR({n + 1},tHotelMonthly[Room_Nights],tHotelMonthly[Month_No]),0)",
          round(float(np.polyval(np.polyfit(hm.Month_No, rn, 1), n + 1))), "FORECAST.LINEAR(next month, y, x).")
    m.mc("A hotel's December is busy every year. That pattern is called ...", ["seasonality", "trend", "an outlier",
                                                                              "a residual"], "A",
         "FORECAST.ETS can model it.")
    m.mc("Forecast 500, actual 400. What is the absolute percentage error?", ["20%", "25%", "80%", "100"], "B",
         "|400 − 500| ÷ 400 = 25%.")

    # ================================================================== T4
    m = t4
    m.info("Model", "Room rate ₦80,000 per night; 5,000 room nights a month; cost ₦50,000 per room night. "
                    "Monthly profit = (Rate − Cost) × Room nights.")
    m.num("Monthly profit in the model?", "=(80000-50000)*5000", 150_000_000, "(80,000 − 50,000) × 5,000.", fmt="naira")
    m.num("Goal Seek: what room rate gives a monthly profit of ₦200,000,000?", "=50000+200000000/5000", 90_000,
          "Rate = Cost + Target ÷ Room nights.", fmt="naira")
    m.num("Data table: profit at 4,000 room nights (rate ₦80,000)?", "=(80000-50000)*4000", 120_000_000,
          "(80,000 − 50,000) × 4,000.", fmt="naira")
    m.pick("Compare Low, Base and High seasons, each with its own rate and room nights.",
           ["Goal Seek", "Data Table", "Scenario Manager", "Solver"], "Scenario Manager", "Named sets of inputs.")

    # ================================================================== T5
    m = t5
    m.mc("Which slide title is best?", ["Room nights", "Hotel data 2024", "December is the busiest month in both years: "
                                        "plan festive staffing early", "Chart"], "C", "State the finding and the action.")
    m.mc("Which goes first in a report?", ["Executive summary", "Appendix", "Raw data", "Methodology"], "A",
         "One page, decision first.")
    m.mc("A formula gives #REF!. Which tool shows which cells it uses?", ["Trace Precedents", "Flash Fill", "Slicer",
                                                                         "Solver"], "A", "Formulas › Trace Precedents.")
    m.yn("An AI tool says the median salary is ₦9.5m. Should you check it with =MEDIAN before using it?", "Yes",
         "Verify every AI number with a control calculation.")
    m.yn("May you paste the salary survey with names and phone numbers into a public AI chatbot?", "No",
         "Personal data must never go into external AI tools.")
    m.pick("Respondent_ID with Job_Title and salary, but no name", ["Personal data (PII)", "Anonymised", "Health data (PHI)"],
           "Anonymised", "No direct identifier; still check small groups.")
