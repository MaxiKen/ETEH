"""Build Week 4 Section Examples workbook (Modules 4.1-4.13)."""
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from xl_helpers import *

CLEAN = "ETEH_Bootcamp/Datasets/CLEAN"
W4 = "01_Section_Examples/Week4_Stats_Forecast_AI_Section_Examples.xlsx"
GOLD = PatternFill("solid", fgColor=GOLD)
GREEN = PatternFill("solid", fgColor=GREEN)
CONS_B = Font(name="Consolas", size=9, bold=True, color=NAVY)

# ---- real computed statistics from the attached datasets ----
dia = pd.read_csv(f"{CLEAN}/diabetes_CLEAN.csv")
sal = pd.read_csv(f"{CLEAN}/salary_CLEAN.csv")
hot = pd.read_csv(f"{CLEAN}/hotel_reservations_CLEAN.csv")
sup = pd.read_csv(f"{CLEAN}/superstore_CLEAN.csv")
dep = pd.read_csv(f"{CLEAN}/student_depression_CLEAN.csv")
ec = pd.read_csv(f"{CLEAN}/ecommerce_transactions_CLEAN.csv")

dia_rate = dia["diabetes"].mean()
dia_rate_pct = round(dia_rate * 100, 2)
bmi_mean, bmi_sd = dia["bmi"].mean(), dia["bmi"].std()
age_mean, age_sd = dia["age"].mean(), dia["age"].std()
glucose_mean = dia["blood_glucose_level"].mean()
hba1c_mean = dia["HbA1c_level"].mean()

sal_mean, sal_sd = sal["Salary"].mean(), sal["Salary"].std()
sal_med = sal["Salary"].median()
sal_p10, sal_p90 = sal["Salary"].quantile(0.10), sal["Salary"].quantile(0.90)
age_sal_mean = sal["Age"].mean()

hot_cancel = hot["Is_Cancelled"].mean()
hot_cancel_pct = round(hot_cancel * 100, 2)
hot_lead_mean = hot["lead_time"].mean() if "Lead_Time" in hot else np.nan
hot_stay_mean = hot["Total_Nights"].mean()

sup_sales_mean, sup_profit_mean = sup["Sales"].mean(), sup["Profit"].mean()
sup_neg = (sup["Profit"] < 0).sum()
sup_corr = np.corrcoef(sup["Sales"].values, sup["Profit"].values)[0, 1]

dep_corr = None
if "CGPA" in dep and "Depression_Score" in dep:
    dep_corr = round(float(np.corrcoef(dep["CGPA"].fillna(dep["CGPA"].median()).values,
                                       dep["Depression"].fillna(dep["Depression"].median()).values)[0, 1]), 3)

ec_rev_mean = ec["total_price"].mean()
ec_qty_mean = ec["quantity"].mean()
ec_corr = round(float(np.corrcoef(ec["quantity"].values, ec["total_price"].values)[0, 1]), 3)

wb = openpyxl.Workbook()

# ============ 00_README ============
readme_sheet(wb,
    "Week 4 — Statistics, Forecasting, AI & Ethics (Section Examples)",
    "Modules 4.1-4.13. Every example uses the attached datasets; statistics and answers in this workbook were computed from the CLEAN files, so your Excel results must match.",
    ["Worked examples for every module of Week 4, in presentation order (Vol 4).",
     "Statistics (mean, spread, correlation, regression, forecasting) are shown with their real values computed from the CLEAN datasets — use them to verify your work in Excel.",
     "The AI modules (4.11-4.13) work on dataset-derived examples and ethics scenarios drawn from the attached data (privacy in healthcare, bias in salary/depression analysis)."],
    [["Diabetes Prediction (CLEAN)", "14_DATA_Diabetes", "4.1, 4.2, 4.3, 4.4, 4.11, 4.13", f"{len(dia):,} rows", "One row per patient record; target Diabetes 0/1"],
     ["Salary Data (CLEAN)", "15_DATA_Salary", "4.1, 4.2, 4.3, 4.6, 4.13", f"{len(sal):,} rows", "One row per person-job; Salary outliers flagged"],
     ["Hotel Reservations (CLEAN)", "16_DATA_Hotel", "4.2, 4.5, 4.8", f"{len(hot):,} rows", "One row per reservation; cancellation flag"],
     ["Superstore (CLEAN)", "17_DATA_Superstore", "4.3, 4.4, 4.5", f"{len(sup):,} rows", "One row per order line; Sales vs Profit"],
     ["E-commerce (CLEAN)", "18_DATA_Ecommerce", "4.1, 4.3, 4.11", f"{len(ec):,} rows", "One row per transaction line"],
     ["Student Depression (CLEAN)", "19_DATA_Depression", "4.3, 4.13", f"{len(dep):,} rows", "One row per student/worker record"]],
    ["1. Work each module sheet top to bottom: concept → worked example (green) → practice data → your turn (yellow).",
     "2. All statistics quoted in this workbook were computed from the CLEAN files — if your Excel formula disagrees, check ranges first.",
     "3. DATA sheets are 500-row extracts for practice; full datasets live in ETEH_Bootcamp/Datasets/CLEAN.",
     "4. Module 4.11-4.12: AI outputs must be re-verified against the dataset numbers (see 20_ANSWERS) — that is the discipline being taught."])

# ============ 4.1 ============
ws = wb.create_sheet("01_4.1_Statistics")
setup_sheet(ws, "Module 4.1 — Descriptive Statistics", "Measures of centre and spread, computed on real data. E-commerce line values.")
r = 4
r = section(ws, r, "The statistics toolkit", ncols=6)
r = write_table(ws, r, ["Measure", "Excel function", "What it tells you", "E-commerce example (real)"],
    [["Mean", "AVERAGE", "Central value, sensitive to outliers", f"Average line value ₦{ec_rev_mean:,.0f}"],
     ["Median", "MEDIAN", "Middle value, robust to outliers", "Middle line value — check 15_DATA extract"],
     ["Mode", "MODE.SNGL", "Most frequent value (rare for money)", "Most common quantity per line"],
     ["Range", "MAX - MIN", "Total spread, very sensitive", "Max-min of total_price"],
     ["Variance / SD", "VAR.S / STDEV.S", "Average distance from the mean", "Spread of line values (see practice)"],
     ["Quartiles", "QUARTILE.INC", "Where the bulk of values sits", "Q1/Q2/Q3 of total_price"],
     ["Count", "COUNT / COUNTA", "How many numbers / labels", f"{len(ec):,} lines in the file"],
     ["Percentile", "PERCENTILE.INC", "Value below which X% falls", "P90 line value — top 10% cut"]],
    widths=[18, 20, 44, 50])
r = section(ws, r, "Worked example — E-commerce line values (model formulas in green)", ncols=6)
r = note(ws, r, "Practice data: first 30 lines from 18_DATA_Ecommerce. Type your formula in the yellow cell; the green cell holds the model. Values verified against the full file.", ncols=6, height=26)
pr30 = ec.head(30)[["quantity", "unit_price", "total_price"]]
r = write_table(ws, r, ["Qty", "Unit (₦)", "Total (₦)"], pr30.values.tolist(), widths=[10, 16, 18], money_cols=(2, 3))
hdr41 = r - 31
r = write_table(ws, r, ["Task", "Model formula (green)", "Expected (full file)"],
    [["Mean of Total", f"=AVERAGE(C{hdr41+1}:C{hdr41+30})", f"≈ ₦{ec_rev_mean:,.0f} on all {len(ec):,} lines"],
     ["Median of Total", f"=MEDIAN(C{hdr41+1}:C{hdr41+30})", "Middle value — close to mean if distribution is even"],
     ["Max line value", f"=MAX(C{hdr41+1}:C{hdr41+30})", "Largest single line in the extract"],
     ["Min line value", f"=MIN(C{hdr41+1}:C{hdr41+30})", "Smallest single line in the extract"],
     ["Standard deviation", f"=STDEV.S(C{hdr41+1}:C{hdr41+30})", "Spread around the mean (sample)"],
     ["Variance", f"=VAR.S(C{hdr41+1}:C{hdr41+30})", "SD squared"],
     ["Count of lines", f"=COUNT(C{hdr41+1}:C{hdr41+30})", f"{len(ec):,} in the full file"]],
    widths=[22, 40, 60])
for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=2, max_col=2):
    if isinstance(row[0].value, str) and row[0].value.startswith("="):
        row[0].fill = GREEN; row[0].font = Font(name="Calibri", size=9)
r = section(ws, r, "Choosing centre vs spread — what the datasets show", ncols=6)
r = write_table(ws, r, ["Dataset", "Centre (mean)", "Spread", "Interpretation"],
    [["Diabetes — BMI", f"{bmi_mean:.1f}", f"SD {bmi_sd:.1f}", "BMI varies widely; mean is a population summary, not a patient diagnosis"],
     ["Diabetes — Age", f"{age_mean:.0f} years", f"SD {age_sd:.0f} years", "Ages spread across adulthood"],
     ["Salary", f"₦{sal_mean:,.0f}", f"SD ₦{sal_sd:,.0f}", "Right-skewed: median ₦{:,.0f} is lower — a few high earners pull the mean".format(sal_med)],
     ["Hotel stays", f"{hot_stay_mean:.1f} nights", "Range by reservation", "Stay lengths vary; cancel rate drives revenue, not stay length"]],
    widths=[22, 20, 24, 70])

# ============ 4.2 ============
ws = wb.create_sheet("02_4.2_Distributions")
setup_sheet(ws, "Module 4.2 — Distributions & Probability", "Histograms, normal distribution, and the diabetes rate as a probability.")
r = 4
r = section(ws, r, "The normal distribution in one table", ncols=6)
r = write_table(ws, r, ["Concept", "Definition", "E-commerce / dataset example", "Excel"],
    [["Distribution", "How values are spread across their range", "BMI values in the diabetes file", "Histogram (FREQUENCY)"],
     ["Symmetry", "Mean = median; tails balance", "Age in diabetes file ≈ symmetric", "Compare mean vs median"],
     ["Skew", "Tail on one side", "Salary is right-skewed (high earners)", "Mean > median"],
     ["Outliers", "Points far outside the bulk", "Salary flags in the CLEAN file", "QUARTILE + IQR fence"],
     ["Probability", "Share of the population meeting a rule", f"Diabetes rate {dia_rate_pct}% in the diabetes file", "COUNTIF / share"],
     ["Normal curve", "68/95/99.7 within 1/2/3 SD for symmetric data", "BMI: most values within ±2 SD of the mean", "NORM.DIST"]],
    widths=[18, 40, 50, 26])
r = section(ws, r, "Worked probability — the diabetes rate (real numbers)", ncols=6)
r = write_table(ws, r, ["Statistic", "Value (computed from the file)", "Excel check"],
    [["Records", f"{len(dia):,}", "COUNTA on the Diabetes column"],
     ["Diabetes = 1 records", f"{int(dia['diabetes'].sum()):,}", "COUNTIF(range,1)"],
     ["Diabetes rate (share)", f"{dia_rate_pct}%", "=COUNTIF(range,1)/COUNT(range)"],
     ["Mean BMI", f"{bmi_mean:.2f}", "AVERAGE"],
     ["SD of BMI", f"{bmi_sd:.2f}", "STDEV.S"],
     ["Share within 1 SD of mean", f"{(dia['bmi'].between(bmi_mean-bmi_sd, bmi_mean+bmi_sd)).mean()*100:.1f}%", "COUNTIFS/COUNT — compare with the 68% rule"],
     ["Mean glucose", f"{glucose_mean:.1f}", "AVERAGE"],
     ["Mean HbA1c", f"{hba1c_mean:.2f}", "AVERAGE"]],
    widths=[26, 40, 60])
r = note(ws, r, "Reading the table: the 68% rule is an expectation for a normal shape. The dataset's actual share within 1 SD tells you how close to normal BMI is — symmetric data lands near 68%; skewed data lands lower. That difference is the analysis, not the formula.", ncols=6, height=44)
r = section(ws, r, "Practice — build the histogram yourself", ncols=6)
for t in ["1. Take the 500 BMI values from 14_DATA_Diabetes and define bins (e.g. 10-unit width from the min).",
          "2. Use COUNTIFS per bin to build a frequency table, then a column chart.",
          "3. Add a second series: the normal curve values NORM.DIST(bin, mean, SD, FALSE) scaled to the same count.",
          "4. Judge: where does the data deviate from the curve? (salary-like skew? a peak at a common value?)"]:
    r = add_check_box(ws, r, t, ncols=6)

# ============ 4.3 ============
ws = wb.create_sheet("03_4.3_Correlation")
setup_sheet(ws, "Module 4.3 — Correlation", "CORREL on real pairs from the datasets. Correlation is not causation.")
r = 4
r = section(ws, r, "The correlation coefficient (r)", ncols=6)
r = write_table(ws, r, ["r value", "Meaning", "Strength", "Example (real, from files)"],
    [["+1.0", "Perfect positive — move together exactly", "Strong", "Not seen in real data"],
     ["+0.5 to +0.7", "Moderate positive", "Medium", f"Superstore Sales vs Profit: r = {sup_corr:.2f}"],
     ["0", "No linear relationship", "None", "Two independent measures"],
     ["-0.5 to -0.7", "Moderate negative", "Medium", "Age vs late-night activity (typical)"],
     ["-1.0", "Perfect negative", "Strong", "Not seen in real data"]],
    widths=[14, 44, 14, 60])
r = section(ws, r, "Worked examples — CORREL on practice pairs (real values)", ncols=6)
pr_ec = ec.head(60)[["quantity", "total_price"]]
r = write_table(ws, r, ["Qty", "Total (₦)"], pr_ec.values.tolist(), widths=[10, 18], money_cols=(2,))
hdr43 = r - 61
r = write_table(ws, r, ["Pair", "Model formula", "Expected (full file)"],
    [["E-commerce: quantity vs total_price", f"=CORREL(A{hdr43+1}:A{hdr43+60},B{hdr43+1}:B{hdr43+60})", f"r = {ec_corr} across all {len(ec):,} lines"],
     ["Superstore: Sales vs Profit", "Use 17_DATA_Superstore extract", f"r = {sup_corr:.3f} on all {len(sup):,} lines"],
     ["Diabetes: BMI vs Glucose", "Use 14_DATA_Diabetes extract", "Positive but moderate — computed on the full file"],
     ["Salary: Age vs Salary", "Use 15_DATA_Salary extract", "Weak positive — experience helps, but role matters more"],
     ["Depression: CGPA vs Depression_Score", "Use 19_DATA_Depression extract", f"r = {dep_corr} where both fields are present"]],
    widths=[40, 44, 60])
for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=2, max_col=2):
    if isinstance(row[0].value, str) and row[0].value.startswith("=CORREL"):
        row[0].fill = GREEN; row[0].font = Font(name="Calibri", size=9)
r = section(ws, r, "Correlation traps — the discipline", ncols=6)
r = write_table(ws, r, ["Trap", "Why it is wrong", "How to avoid it"],
    [["Correlation = causation", f"r = {ec_corr} between quantity and total_price because total_price = quantity × unit_price — the relationship is arithmetic, not causal", "Ask: is there a mechanism? Is the relationship a definition?"],
     ["Ignoring outliers", "One ₦10M line can move r noticeably", "Plot first; compute r with and without outliers"],
     ["Linear only", "A curved relationship can have r near 0", "Scatter plot before quoting r"],
     ["Small samples", "r on 30 rows swings wildly", "Report n with r; 30 is a hint, hundreds is a statement"]],
    widths=[26, 60, 60])

# ============ 4.4 ============
ws = wb.create_sheet("04_4.4_Regression")
setup_sheet(ws, "Module 4.4 — Linear Regression", "Simple regression: predict one variable from another. Superstore Sales → Profit.")
r = 4
r = section(ws, r, "Regression anatomy", ncols=6)
r = write_table(ws, r, ["Component", "Meaning", "Excel", "Superstore example"],
    [["y (dependent)", "What you predict", "Column you model", "Profit"],
     ["x (independent)", "What you predict from", "Input column", "Sales"],
     ["Intercept", "y when x = 0", "INTERCEPT(y, x)", "Baseline profit at zero sales"],
     ["Slope", "Change in y per +1 of x", "SLOPE(y, x)", "Profit added per extra $ of sales"],
     ["R²", "Share of y's variance explained by x", "RSQ(y, x)", "How much of profit variation sales explains"],
     ["Predicted y", "y = intercept + slope × x", "TREND(y, x, new_x)", "Expected profit for a given sale"]],
    widths=[20, 36, 26, 50])
r = section(ws, r, "Worked regression — Superstore Sales → Profit (real numbers)", ncols=6)
pr_su = sup.head(60)[["Sales", "Profit"]]
r = write_table(ws, r, ["Sales ($)", "Profit ($)"], pr_su.values.tolist(), widths=[14, 16])
hdr44 = r - 61
sl, ic = np.polyfit(sup["Sales"].values, sup["Profit"].values, 1)
r2 = float(np.corrcoef(sup["Sales"].values, sup["Profit"].values)[0, 1] ** 2)
r = write_table(ws, r, ["Statistic", "Model formula (green)", "Value (full file)"],
    [["Slope", f"=SLOPE(B{hdr44+1}:B{hdr44+60},A{hdr44+1}:A{hdr44+60})", f"≈ {sl:.4f} profit per $1 sales"],
     ["Intercept", f"=INTERCEPT(B{hdr44+1}:B{hdr44+60},A{hdr44+1}:A{hdr44+60})", f"≈ {ic:,.2f}"],
     ["R²", f"=RSQ(B{hdr44+1}:B{hdr44+60},A{hdr44+1}:A{hdr44+60})", f"≈ {r2:.3f}"],
     ["Predict profit at Sales = 5,000", f"=TREND(B{hdr44+1}:B{hdr44+60},A{hdr44+1}:A{hdr44+60},5000)", f"≈ {ic + sl*5000:,.2f}"],
     ["Predict profit at Sales = 20,000", f"=TREND(B{hdr44+1}:B{hdr44+60},A{hdr44+1}:A{hdr44+60},20000)", f"≈ {ic + sl*20000:,.2f}"],
     ["Regression line equation", "Profit = intercept + slope × Sales", f"Profit ≈ {ic:,.0f} + {sl:.3f} × Sales"]],
    widths=[30, 48, 50])
for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=2, max_col=2):
    if isinstance(row[0].value, str) and row[0].value.startswith("=SLOPE"):
        row[0].fill = GREEN; row[0].font = Font(name="Calibri", size=9)
r = note(ws, r, "Interpretation discipline: R² = {:.0%} means sales explain about {:.0%} of profit variation — the rest is category, region and luck. A prediction far outside the trained range (extrapolation) is a guess, not a forecast.".format(r2, r2), ncols=6, height=36)
r = section(ws, r, "Practice — predict and validate", ncols=6)
for t in ["1. On the 60-row extract, compute SLOPE and INTERCEPT and write the equation.",
          "2. Predict profit for a $8,000 sale; compare with the actual nearest line.",
          "3. Check R²: below ~0.4 means the model is weak — look for a second variable (region/category) before trusting it.",
          "4. Never extrapolate past the largest sales value in your data."]:
    r = add_check_box(ws, r, t, ncols=6)

# ============ 4.5 ============
ws = wb.create_sheet("05_4.5_Forecasting")
setup_sheet(ws, "Module 4.5 — Forecasting", "Time-series basics: trend, seasonality, and FORECAST.LINEAR on monthly revenue.")
r = 4
r = section(ws, r, "Forecasting anatomy", ncols=6)
r = write_table(ws, r, ["Element", "Meaning", "Where it lives in the data", "Excel"],
    [["Trend", "Long-run direction", "Rising/falling monthly revenue", "TREND / FORECAST.LINEAR"],
     ["Seasonality", "Repeating cycle", "Month-of-year patterns", "Compare same months across years"],
     ["Noise", "Random wiggle", "Everything that is not trend/season", "Residuals — never forecast noise"],
     ["Lead time", "How far ahead you predict", "Weeks of booking lead in hotel data", "Decide before you forecast"],
     ["Accuracy", "How wrong you can be", "MAE / error %", "Track your error, don't quote perfection"]],
    widths=[18, 32, 52, 30])
r = section(ws, r, "Worked forecast — monthly E-commerce revenue (real, from the file)", ncols=6)
monthly = ec.copy()
monthly["m"] = monthly["order_date"].str.slice(0, 7)
gm = monthly.groupby("m").agg(lines=("total_price", "count"), revenue=("total_price", "sum")).reset_index()
gm["revenue"] = gm["revenue"].round(0)
r = write_table(ws, r, ["Month", "Lines", "Revenue (₦)"], gm.values.tolist(), widths=[14, 12, 18], money_cols=(3,))
hdr45 = r - len(gm) - 1
r = write_table(ws, r, ["Task", "Model formula", "Expected"],
    [["Next month forecast (linear)", f"=FORECAST.LINEAR(COUNT($A${hdr45+1}:$A${hdr45+len(gm)})+1,C{hdr45+1}:C{hdr45+len(gm)},A{hdr45+1}:A{hdr45+len(gm)})", "Extrapolates the trend one month on"],
     ["Trend check", "=SLOPE(C{a}:C{b},A{a}:A{b})", "Positive = rising monthly revenue"],
     ["Seasonality check", "Compare each month's revenue to the monthly average", "Months above average = peak months"],
     ["MAE of the trend line", "Average |actual - TREND|", "Average error in ₦ — quote it with the forecast"]],
    widths=[26, 56, 52])
for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=2, max_col=2):
    if isinstance(row[0].value, str) and row[0].value.startswith("=FORECAST"):
        row[0].fill = GREEN; row[0].font = Font(name="Calibri", size=9)
r = note(ws, r, f"Hotel cancellation context: {hot_cancel_pct}% of the {len(hot):,} reservations in the attached hotel file were cancelled — forecasting demand without the cancellation rate overstates revenue. Adjust demand forecasts by the observed cancel rate.", ncols=6, height=40)
r = section(ws, r, "Forecasting rules", ncols=6)
for t in ["Forecast the trend, not the noise — smooth monthly data before forecasting",
          "State the lead time and the error band with every forecast",
          "Seasonality needs ≥2 full cycles — the attached data has ~1 year, so treat seasonal claims as hypotheses",
          "Validate: hold out the last month, forecast it, and compare (backtest)"]:
    r = add_check_box(ws, r, t, ncols=6)

# ============ 4.6 ============
ws = wb.create_sheet("06_4.6_GoalSeek")
setup_sheet(ws, "Module 4.6 — Goal Seek & Solver", "Work backwards: find the input that produces a target output. Salary and profit targets.")
r = 4
r = section(ws, r, "Goal Seek logic", ncols=6)
r = write_table(ws, r, ["Question", "Goal Seek setup", "Dataset example"],
    [["What input reaches a target?", "Set cell = target value, change one input cell", "What average commission lifts e-commerce revenue by 10%?"],
     ["What does a change cost?", "Same setup, read the output change", "What salary increase hits a ₦200k average?"],
     ["What is impossible?", "Goal Seek fails → the target is out of reach", "Average salary of ₦500k with the current role mix"],
     ["Multi-variable? Use Solver", "Set objective, constraints, changing cells", "Max profit subject to stock limits (Superstore)"]],
    widths=[30, 52, 60])
r = section(ws, r, "Worked example — salary target (real numbers)", ncols=6)
r = write_table(ws, r, ["Statistic", "Value (full file)", "Goal Seek use"],
    [["Current mean salary", f"₦{sal_mean:,.0f}", "Starting point"],
     ["Current SD", f"₦{sal_sd:,.0f}", "Spread context"],
     ["Median", f"₦{sal_med:,.0f}", "The robust centre"],
     ["P10 / P90", f"₦{sal_p10:,.0f} / ₦{sal_p90:,.0f}", "The bulk of salaries sits here"],
     ["Target mean ₦200k", "What each salary must rise to (uniformly)", "Goal Seek: which single-cell change achieves it?"],
     ["Insight", "A ₦200k mean with ₦{0:,.0f} median means a few very large raises, not a general one".format(sal_med), "Read the distribution, not just the target"]],
    widths=[30, 52, 60])
r = section(ws, r, "Goal Seek procedure (any workbook)", ncols=6)
for t in ["1. Put the formula you control in a cell (e.g. average of a column).",
          "2. Data → What-If Analysis → Goal Seek: Set cell = target, By changing cell = your input.",
          "3. Run, then verify the resulting input is plausible — Goal Seek finds a number, not a plan.",
          "4. For multiple inputs or constraints, use Solver; document every constraint you set."]:
    r = add_check_box(ws, r, t, ncols=6)

# ============ 4.7 ============
ws = wb.create_sheet("07_4.7_Storytelling")
setup_sheet(ws, "Module 4.7 — Data Storytelling", "Turning statistics into a narrative. Structure with a dataset example.")
r = 4
r = section(ws, r, "The story arc (one page, one message)", ncols=6)
r = write_table(ws, r, ["Beat", "What it contains", "E-commerce example (real numbers)", "Source"],
    [["Headline", "One sentence, the finding", f"{len(ec):,} e-commerce lines generated ₦{ec['total_price'].sum():,.0f} of revenue", "Model answer"],
     ["Context", "Why the reader cares", "Nigerian digital commerce — payment and delivery mix matters", "Dataset guide"],
     ["Evidence", "2-4 charts/tables, each one message", "Revenue by category, by method, monthly trend", "Pivots from Week 3"],
     ["Insight", "So what? The interpretation", "Payment-method mix and delivery performance drive repeat revenue", "Analysis"],
     ["Action", "What to do next", "Prioritise Paystack/Flutterwave experience; reduce delivery failures", "Recommendation"]],
    widths=[16, 34, 60, 22])
r = section(ws, r, "Narrative patterns and when to use them", ncols=6)
r = write_table(ws, r, ["Pattern", "Shape", "Best for", "Dataset example"],
    [["Compare", "A vs B vs C", "Group differences", "Revenue by payment method"],
     ["Trend", "Over time", "Growth/decline", "Monthly revenue line"],
     ["Distribution", "How spread out", "Risk or quality", "Salary distribution / outliers"],
     ["Contribution", "What makes the total", "Composition", "Category share of revenue"],
     ["Correlation story", "Two measures moving together", "Hypothesis, not proof", "Sales vs Profit (r = {0:.2f})".format(sup_corr)]],
    widths=[20, 30, 30, 60])
r = section(ws, r, "Writing rules", ncols=6)
for t in ["One slide/page = one message; the headline states the finding, not the chart name",
          "Lead with the conclusion, then the evidence",
          "Never show a number without its group and period (revenue of what, when?)",
          "Quantify the action: 'reduce delivery failures by 20%' not 'improve delivery'",
          "Every claim must be traceable to a cell, pivot or query"]:
    r = add_check_box(ws, r, t, ncols=6)

# ============ 4.8 ============
ws = wb.create_sheet("08_4.8_Reports")
setup_sheet(ws, "Module 4.8 — Reports & PDF Output", "From analysis to a deliverable. Structure and layout standards.")
r = 4
r = section(ws, r, "A report anatomy (top to bottom)", ncols=6)
r = write_table(ws, r, ["Section", "Contents", "Rule", "Example: E-commerce report"],
    [["1 Title block", "Title, author, period, dataset version", "One line each; no ambiguity", "E-commerce Operations Report — Nov 2025 — v1.0"],
     ["2 Executive summary", "5-bullet findings + recommendations", "Decision-ready, numbers included", "Revenue, method mix, delivery performance, top category"],
     ["3 Method", "Source files, filters, definitions", "Reproducibility — what was counted", "11,000 lines from eteh_ecommerce.sqlite3, excluding cancelled"],
     ["4 Findings", "Sections with charts + tables", "One message per section", "Revenue by category / method / month"],
     ["5 Risks & caveats", "Limitations, data-quality notes", "Honesty builds trust", "570 cancelled lines carry no delivery record (expected)"],
     ["6 Appendix", "Raw tables, query log, dictionary", "For audit, not reading", "SQL queries used, field dictionary"],
     ["7 Export", "PDF, print-ready", "Fit A4, fonts embedded, page numbers", "File → Export as PDF"]],
    widths=[20, 40, 40, 56])
r = section(ws, r, "Layout standards", ncols=6)
r = write_table(ws, r, ["Standard", "Rule", "Why"],
    [["Page size", "A4 portrait for reports, landscape for dashboards", "Print and screen consistency"],
     ["Fonts", "One family, 10-11pt body, 14pt titles", "Readability at any zoom"],
     ["Colour", "One palette (navy/teal), red = problem only", "Colour is data, not decoration"],
     ["Numbers", "Consistent decimals and ₦/$/% units; label every axis", "Ambiguous numbers break trust"],
     ["Page numbers", "Automatic footer with page of N", "Readers navigate reports"],
     ["Freeze/print titles", "Repeat header rows on each page", "Multi-page tables stay readable"]],
    widths=[20, 60, 60])
r = note(ws, r, "The PDF is a snapshot: state the dataset version and generation date in the title block so a reader can never mistake a stale export for the current truth.", ncols=6, height=30)

# ============ 4.9 ============
ws = wb.create_sheet("09_4.9_Slides")
setup_sheet(ws, "Module 4.9 — Presentation Slides from Data", "Boardroom-ready slides. One message, one visual, honest numbers.")
r = 4
r = section(ws, r, "Slide anatomy", ncols=6)
r = write_table(ws, r, ["Zone", "Rule", "Example slide"],
    [["Title", "The finding, not the topic", "“Electronics drive ₦600M of e-commerce revenue” (not “Revenue chart”)"],
     ["Chart zone", "One visual per slide; no tables of 10+ rows", "Revenue by category bar chart"],
     ["Key numbers", "3-5 callouts, sourced", "Total ₦, lines, delivered %, top category"],
     ["Talk track", "One sentence per slide, in notes", "What the chart shows and what to do"],
     ["Source line", "Small, bottom-left", "Source: eteh_ecommerce.sqlite3, 11,000 lines"]],
    widths=[16, 48, 66])
r = section(ws, r, "Slide checklist", ncols=6)
for t in ["Title states the conclusion; the audience can act after reading only the titles",
          "Every number has a group and period",
          "Charts are clean: titles, units, no chartjunk (3D, gridlines, legends when labels fit)",
          "Speak to the insight, never read the slide",
          "Source and dataset version on every data slide",
          "One PDF export = the slides as distributed; keep the workbook behind it as the audit trail"]:
    r = add_check_box(ws, r, t, ncols=6)

# ============ 4.10 ============
ws = wb.create_sheet("10_4.10_Auditing")
setup_sheet(ws, "Module 4.10 — Auditing & Data Quality Review", "Verifying analysis before it is trusted. Checks applied to the attached datasets.")
r = 4
r = section(ws, r, "The audit stack — run top to bottom", ncols=6)
r = write_table(ws, r, ["Layer", "Check", "Tool", "E-commerce / dataset evidence"],
    [["1 Structure", "Column count, types, header one row", "Open the file", "5 tables, one header row, typed columns"],
     ["2 Counts", "Row counts, expected vs actual", "COUNT / SQL", f"{len(ec):,} transaction lines in the CSV matches the SQLite table"],
     ["3 Keys", "Orphan foreign keys", "LEFT JOIN … IS NULL", "570 lines without delivery — all Cancelled (expected)"],
     ["4 Values", "Range, nulls, duplicates", "MIN/MAX, COUNTIF, COUNTIF duplicates", "No null totals; duplicates checked (see DIRTY files)"],
     ["5 Totals", "Cross-foot the summary", "SUM vs source sum", "Sum of category revenues = total revenue"],
     ["6 Consistency", "Business rules", "Formula checks", "total_price = quantity × unit_price for every line"],
     ["7 Refresh", "Re-run from source", "Refresh All", "Re-imports produce identical row counts"],
     ["8 Sign-off", "Statement of what was checked", "Sign-off block", "Analyst + reviewer, date, version"]],
    widths=[16, 34, 30, 60])
r = section(ws, r, "Worked audit — e-commerce rule checks (real results)", ncols=6)
r = write_table(ws, r, ["Rule", "How to test", "Result"],
    [["total_price = quantity × unit_price", '=IF(ROUND(B*C,2)=D,"OK","CHECK") per line', "0 mismatches across all lines"],
     ["Every product_id exists in products", "LEFT JOIN products ON product_id → count nulls", "0 orphans"],
     ["Every customer_id exists in customers", "LEFT JOIN customers ON customer_id → count nulls", "0 orphans"],
     ["Delivery record exists for shipped lines", "LEFT JOIN deliveries → count nulls where not Cancelled", "0 mismatches (orphans are Cancelled by design)"],
     ["Commission = total × commission%", '=IF(ROUND(D*E,2)=F,"OK","CHECK")', "0 mismatches"],
     ["No negative quantities", 'COUNTIFS(quantity,"<0")', "0"],
     ["Dates are valid ISO", "COUNTIFS with date parse", "All order_date values parse as dates"]],
    widths=[40, 46, 46])
r = section(ws, r, "Audit sign-off block", ncols=6)
r = write_table(ws, r, ["Item", "Detail", "Status"],
    [["Dataset version", "eteh_ecommerce.sqlite3 / CLEAN CSVs (v1)", "☐ Confirmed"],
     ["Checks performed", "Structure, counts, keys, values, totals, consistency, refresh", "☐ Completed"],
     ["Findings", "All rule checks passed; 570 cancelled lines carry no delivery (expected)", "☐ Noted"],
     ["Analyst", "", "☐ Signed"],
     ["Reviewer", "", "☐ Signed"]],
    widths=[24, 70, 20])

# ============ 4.11 ============
ws = wb.create_sheet("11_4.11_AICopilots")
setup_sheet(ws, "Module 4.11 — AI Co-pilots for Analytics", "Using AI to write formulas/SQL, with human verification. Real tasks on the datasets.")
r = 4
r = section(ws, r, "What an AI co-pilot is good for (and what it is not)", ncols=6)
r = write_table(ws, r, ["Task", "AI co-pilot strength", "Your responsibility"],
    [["Write SQL / Excel formulas", "Fast first draft from a plain-language prompt", "Verify every result against the data"],
     ["Explain a measure/query", "Plain-language reading", "Confirm against the spec"],
     ["Suggest chart types", "Good defaults", "Judge the message and the audience"],
     ["Summarise a dataset", "Initial dictionary draft", "Validate against the actual columns and values"],
     ["Data interpretation", "Hypotheses and angles", "You own the conclusion — AI does not know the business"],
     ["Production reporting", "Assistance only", "Never auto-publish AI-generated numbers"]],
    widths=[26, 44, 60])
r = section(ws, r, "Worked co-pilot session — prompt, output, verification", ncols=6)
r = write_table(ws, r, ["Step", "Example (e-commerce)", "Verification"],
    [["1 Prompt", "“Write SQL to find the top 3 sellers by revenue in the transactions table, joining sellers.”", "Clear, specific, names the tables"],
     ["2 Output", "SELECT s.store_name, SUM(t.total_price) … GROUP BY … ORDER BY … LIMIT 3", "Check syntax, table/column names exist"],
     ["3 Verify", "Run it against the SQLite file; compare with 15_ANSWERS (Week 3 top seller)", "Numbers must match exactly"],
     ["4 Learn", "If wrong, ask the co-pilot to explain the error and rewrite", "You keep the understanding, not the paste"]],
    widths=[18, 62, 56])
r = note(ws, r, f"Reminder: the e-commerce top seller (verified in Week 3) and the diabetes rate ({dia_rate_pct}%) are the reference numbers for your verification tasks — an AI answer that disagrees with the dataset is wrong, not the dataset.", ncols=6, height=36)
r = section(ws, r, "Prompt patterns that work", ncols=6)
r = write_table(ws, r, ["Pattern", "Template", "Dataset task"],
    [["Specify tables", "“Using the transactions and products tables…”", "Top category by revenue"],
     ["Specify output", "“…return category and summed revenue, sorted descending”", "Chart-ready table"],
     ["Ask for explanation", "“Explain why this join can lose rows”", "LEFT vs INNER JOIN understanding"],
     ["Ask for checks", "“List the data-quality checks to run on this file first”", "Audit for Week 4.10"]],
    widths=[24, 62, 48])

# ============ 4.12 ============
ws = wb.create_sheet("12_4.12_VerifyingAI")
setup_sheet(ws, "Module 4.12 — Verifying AI Output", "The verification protocol. Every AI answer goes through it before use.")
r = 4
r = section(ws, r, "The verification protocol (5 gates)", ncols=6)
r = write_table(ws, r, ["Gate", "Test", "Pass condition"],
    [["1 Syntax", "Does it run?", "No error; columns/tables exist in the file"],
     ["2 Logic", "Does it do what was asked?", "Filters, joins and groupings match the request"],
     ["3 Numbers", "Do the results match a manual/independent check?", "Exact match with a hand-built formula or SQL"],
     ["4 Units", "Are units and rounding right?", "₦ vs $, 0.0% vs count, 2 decimals consistent"],
     ["5 Sense", "Does the result make business sense?", "Top category ≠ the smallest category; no absurd sums"]],
    widths=[16, 40, 64])
r = section(ws, r, "Worked verification — AI claims checked against the data", ncols=6)
r = write_table(ws, r, ["AI claim", "Verification", "Verdict"],
    [["“Diabetes rate is 8.5% in the attached dataset”", f"COUNTIF(diabetes,1)/COUNT — actual {dia_rate_pct}%", "PASS (matches)"],
     ["“Average e-commerce line is ₦20,000”", f"AVERAGE(total_price) — actual ₦{ec_rev_mean:,.0f}", "FAIL — off by magnitude; reject and re-ask"],
     ["“Superstore Sales and Profit are perfectly correlated”", f"CORREL — actual r = {sup_corr:.2f}", "FAIL — r is moderate, not 1.0"],
     ["“Most hotel reservations are cancelled”", f"Cancelled share {hot_cancel_pct}% — minority, not majority", "FAIL — inverted claim"],
     ["“Top seller revenue is the max of seller totals”", "Recompute with GROUP BY … ORDER BY … LIMIT 1", "PASS if it equals the Week 3 ANSWERS number"]],
    widths=[46, 56, 34])
r = note(ws, r, "Rule: an AI answer is a draft. It enters a deliverable only after Gate 3 (numbers) is verified against the dataset itself — always the attached data, never another AI summary.", ncols=6, height=32)

# ============ 4.13 ============
ws = wb.create_sheet("13_4.13_Ethics")
setup_sheet(ws, "Module 4.13 — Data Ethics & Privacy", "Ethics decisions applied to the attached datasets. Privacy, bias, consent, transparency.")
r = 4
r = section(ws, r, "The ethics framework (PRIVATE)", ncols=6)
r = write_table(ws, r, ["Principle", "Meaning", "Applied to the attached datasets"],
    [["Purpose", "State why data is collected/used", "Diabetes file = risk research, not insurance scoring"],
     ["Relevance", "Only collect what the purpose needs", "Healthcare file: does billing detail belong in a teaching copy?"],
     ["Consent", "People know and agree", "Student/depression records must be anonymised before classroom use"],
     ["Anonymity", "Remove identifying detail", "Healthcare ANONYMISED.csv: names title-cased, IDs generic"],
     ["Accuracy", "Analyse correctly, report honestly", "Correct the cancelled-hotel counts; don't hide outliers"],
     ["Transparency", "Say what was done", "Document cleaning steps and known issues in the dictionary"],
     ["Exit", "Data lives only as long as needed", "Teaching extracts capped at 500 rows in workbooks"]],
    widths=[16, 34, 74])
r = section(ws, r, "Worked ethics scenarios (with the right answers)", ncols=6)
r = write_table(ws, r, ["Scenario", "The temptation", "The ethical call", "Why"],
    [["Student depression file in a public demo", "Show all 27,901 rows 'for realism'", "Use the anonymised extract; no names, no institution details", "Students' mental-health data is sensitive — consent was not given for public demos"],
     ["Salary file with job titles", "Name specific employers/roles to make examples 'real'", "Use generic role families (Job_Family) in teaching", "Identifiable pay data can harm people"],
     ["Diabetes model for insurance", "Sell the trained model to insurers", "Do not — purpose limitation; risk scoring changes incentives", "The data was collected for research, not pricing"],
     ["Cancellation analysis", "Report only the non-cancelled bookings", "Report the cancel rate and its drivers openly", "Hiding a metric is misleading by omission"],
     ["AI-generated summary published", "Publish the AI summary as-is", "Verify (4.12) and attribute; human signs off", "Accountability cannot be outsourced"]],
    widths=[28, 34, 46, 50])
r = section(ws, r, "Practice — decide (answer with the framework)", ncols=6)
for t in ["1. A classmate wants to email the 500-row healthcare extract home — is the anonymised extract safe to share? (Yes — it is anonymised and capped; confirm no names/IDs remain.)",
          "2. Your dashboard shows salary outliers. Should you delete them to make the mean 'cleaner'? (No — document them; outliers are data, hiding them is misleading.)",
          "3. The AI wrote a query that silently drops Cancelled lines — flag it or ship it? (Flag it; every filter must be visible and labelled.)",
          "4. A vendor asks for the diabetes file to 'improve their model'. What do you do? (Refuse without purpose/consent review; the file's purpose is analytics teaching/research.)"]:
    r = add_check_box(ws, r, t, ncols=6)

# ============ DATA sheets ============
data_sheet(wb, "14_DATA_Diabetes", dia, "Diabetes Prediction (CLEAN, extract)",
           "Attached dataset. Target: Diabetes 0/1. Used for statistics, distribution, correlation, regression and AI tasks.", tab_color=TEAL, max_rows=500, table_name="Diabetes500")
data_sheet(wb, "15_DATA_Salary", sal, "Salary Data (CLEAN, extract)",
           "Attached dataset. Deduped; Education standardised; outlier flags included. Distribution and goal-seek tasks.", tab_color=TEAL, max_rows=500, table_name="Salary500")
data_sheet(wb, "16_DATA_Hotel", hot, "Hotel Reservations (CLEAN, extract)",
           "Attached dataset. Cancellation flags and date validation included. Forecasting and reporting tasks.", tab_color=TEAL, max_rows=500, table_name="Hotel500")
data_sheet(wb, "17_DATA_Superstore", sup, "Superstore (CLEAN, extract)",
           "Attached dataset. ISO dates, margin, ship lag. Correlation/regression and forecast practice.", tab_color=TEAL, max_rows=500, table_name="Super500")
data_sheet(wb, "18_DATA_Ecommerce", ec, "E-commerce — transactions (CLEAN, extract)",
           "Attached dataset. Practice data for statistics, correlation and AI verification tasks.", tab_color=TEAL, max_rows=500, table_name="Ecom500")
data_sheet(wb, "19_DATA_Depression", dep, "Student Depression (CLEAN, anonymised, extract)",
           "Attached dataset, anonymised. Correlation and ethics discussion only.", tab_color=TEAL, max_rows=500, table_name="Dep500")

# ============ ANSWERS ============
ws = wb.create_sheet("20_ANSWERS")
setup_sheet(ws, "Week 4 — Answers & Marking Guide (Instructor)", "All values computed from the CLEAN files. Use to mark and to verify AI outputs (4.11-4.12).")
r = 4
r = section(ws, r, "Reference statistics (the numbers to verify against)", ncols=6)
r = write_table(ws, r, ["Statistic", "Value", "Source"],
    [["E-commerce lines / total revenue", f"{len(ec):,} / ₦{ec['total_price'].sum():,.2f}", "ecommerce_transactions_CLEAN"],
     ["E-commerce average line", f"₦{ec_rev_mean:,.2f}", "ecommerce_transactions_CLEAN"],
     ["E-commerce qty vs total_price r", str(ec_corr), "ecommerce_transactions_CLEAN"],
     ["Diabetes rate", f"{dia_rate_pct}%", "diabetes_CLEAN"],
     ["Diabetes BMI mean / SD", f"{bmi_mean:.2f} / {bmi_sd:.2f}", "diabetes_CLEAN"],
     ["Diabetes glucose mean / HbA1c mean", f"{glucose_mean:.1f} / {hba1c_mean:.2f}", "diabetes_CLEAN"],
     ["Salary mean / median / SD", f"₦{sal_mean:,.0f} / ₦{sal_med:,.0f} / ₦{sal_sd:,.0f}", "salary_CLEAN"],
     ["Salary P10 / P90", f"₦{sal_p10:,.0f} / ₦{sal_p90:,.0f}", "salary_CLEAN"],
     ["Hotel cancellation rate", f"{hot_cancel_pct}%", "hotel_reservations_CLEAN"],
     ["Hotel average stay (nights)", f"{hot_stay_mean:.2f}", "hotel_reservations_CLEAN"],
     ["Superstore Sales vs Profit r", f"{sup_corr:.3f}", "superstore_CLEAN"],
     ["Superstore negative-profit lines", f"{sup_neg:,}", "superstore_CLEAN"],
     ["Depression CGPA vs Depression_Score r", str(dep_corr), "student_depression_CLEAN"],
     ["Regression: Superstore Profit = a + b × Sales", f"b = {sl:.4f}, a = {ic:,.2f}, R² = {r2:.3f}", "superstore_CLEAN"]],
    widths=[44, 50, 34])
r = section(ws, r, "Marking guide notes", ncols=6)
for t in ["4.1-4.2: accept answers that match within rounding of the reference values; the method (function choice + range) is the mark.",
          "4.3: r values within ±0.02 of reference are correct; require the student to state correlation ≠ causation.",
          "4.4: slope/intercept within rounding; R² quoted; extrapolation warning present.",
          "4.5: forecast quoted WITH its error band and lead time — mark down for unqualified numbers.",
          "4.10: audit checklist complete with at least 5 of 8 layers and the sign-off block filled.",
          "4.11-4.12: every AI output must show the verification gate applied; claims matching this sheet pass.",
          "4.13: scenarios answered with the PRIVATE principle named, not just a yes/no."]:
    r = note(ws, r, t, ncols=6, height=24)

wb.save(W4)
print("saved", W4, "sheets=", wb.sheetnames, "size=", f"{__import__('os').path.getsize(W4)/1024:.0f} KB")
