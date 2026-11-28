"""Week 4 edition-3 slide copy: Modules 4.3 (correlation) and 4.4 (linear regression).
Data: tCV (Kingsway), tSal (salary survey), tHotel (Atlantic Pearl bookings). Every number is computed here from w4_data."""
import numpy as np
import pandas as pd
from scipy import stats

from v3kit import T
from w4_fx import CV, EX, HOT, I, N, Nm, P, SAL, Y, F, F2, ols

# ---------------------------------------------------------------- 4.3 figures
DRIVERS = ["Lecture_Attendance_Pct", "Study_Hours_Wk", "Last_Sem_Score", "Tutorials_Attended", "Sleep_Hrs_Night", "Exercise_Hrs_Wk"]
SHORT = {"Lecture_Attendance_Pct": "Attendance %", "Study_Hours_Wk": "Hours studied", "Last_Sem_Score": "Previous score",
         "Tutorials_Attended": "Tutoring", "Sleep_Hrs_Night": "Sleep hours", "Exercise_Hrs_Wk": "Physical activity"}
R = {c: EX.corr(CV[c]) for c in DRIVERS}
RS = {c: EX.corr(CV[c], method="spearman") for c in DRIVERS}
MX = CV[["Final_Score", "Lecture_Attendance_Pct", "Study_Hours_Wk", "Last_Sem_Score"]].corr()
HI = EX >= 80
R_NOHI = CV[~HI].Final_Score.corr(CV[~HI].Study_Hours_Wk)

H = HOT.assign(C=(HOT.Reservation_Status == "Cancelled").astype(int))
R_LEAD = H.Days_Ahead.corr(H.C)
LB = pd.cut(H.Days_Ahead, [-1, 7, 30, 90, 180, 10_000], labels=["0-7", "8-30", "31-90", "91-180", "181+"])
LEAD = H.groupby(LB, observed=True).C.agg(["mean", "size"])
SEG = H.groupby("Sales_Channel").agg(rate=("C", "mean"), lead=("Days_Ahead", "mean"), n=("C", "size"))

S2 = SAL[SAL.Gender.isin(["Male", "Female"])].assign(F=lambda d: (d.Gender == "Female").astype(int))
GM = S2.groupby("Gender").agg(sal=("Annual_Salary_NGN", "mean"), exp=("Years_of_Experience", "mean"))
RAW = ols(S2.Annual_Salary_NGN, S2, ["F"])
ADJ = ols(S2.Annual_Salary_NGN, S2, ["F", "Years_of_Experience"])
XB = pd.cut(S2.Years_of_Experience, [-0.1, 2, 5, 10, 15, 60], labels=["0-2", "3-5", "6-10", "11-15", "16+"])
GAPX = S2.groupby([XB, "Gender"], observed=True).Annual_Salary_NGN.mean().unstack()
GAPX["gap"] = GAPX.Female / GAPX.Male - 1

SP = H[H.Room_Category.isin(["Deluxe", "Standard Twin"])]
SPA = SP.groupby("Room_Category").C.mean()
SPY = SP.groupby(["Check_In_Year", "Room_Category"]).C.agg(["mean", "size"])

# ---------------------------------------------------------------- 4.4 figures
S1 = ols(Y, SAL, ["Years_of_Experience"])
M4 = ols(Y, SAL, ["Years_of_Experience", "Age", "Is_Masters", "Is_PhD"])
M3 = ols(Y, SAL, ["Years_of_Experience", "Is_Masters", "Is_PhD"])
A1, B1 = S1["coef"]["Intercept"], S1["coef"]["Years_of_Experience"]
C4 = M4["coef"]
R_AGE_EXP = SAL.Age.corr(SAL.Years_of_Experience)
VIF = 1 / (1 - R_AGE_EXP ** 2)
RES = pd.Series(S1["resid"])
RB = pd.cut(SAL.Years_of_Experience, [-0.1, 2, 5, 10, 15, 20, 60], labels=["0-2", "3-5", "6-10", "11-15", "16-20", "21+"])
RESB = RES.groupby(RB, observed=True).mean()
PROF = dict(Years_of_Experience=10, Age=35, Is_Masters=1, Is_PhD=0)
PRED = C4["Intercept"] + sum(C4[k] * v for k, v in PROF.items())
PI = 1.96 * M4["sey"]
XMAX = SAL.Years_of_Experience.max()


def pv(p):
    return "below 0.001" if p < 0.001 else f"{p:.3f}"


def rlab(r):
    a = abs(r)
    return "strong" if a >= 0.5 else "moderate" if a >= 0.3 else "weak" if a >= 0.1 else "none"


M43 = dict(
    id="4.3", title="Correlation Analysis",
    subtitle="Measure association honestly with Pearson and Spearman, build a matrix, read the scatter first, and avoid the causation traps.",
    objectives=["Explain what a correlation coefficient measures and its limits.",
                "Calculate Pearson (CORREL) and Spearman (rank) correlation.",
                "Build and read a correlation matrix.",
                "Check scatterplots before trusting r.",
                "Recognise causation traps, confounders and Simpson's paradox, and report correlation responsibly."],
    topics=[
        T("4.3.1", "What Correlation Measures",
          defn="Correlation (r) measures how closely two numeric columns move together in a straight line, from -1 (perfectly "
               "opposite) through 0 (no linear link) to +1 (perfectly together).",
          simple="It measures association only: not cause, not slope, not size of effect.",
          grid=dict(label=f"Exam score against six factors (tCV, {I(len(CV))} students)", widths=[1.8, 0.8, 1.3], rows=[
              ["Factor", "r", "Strength"]] + [[SHORT[c], f"{R[c]:+.2f}", rlab(R[c])] for c in sorted(DRIVERS, key=lambda c: -abs(R[c]))]),
          watch="r near 0 means no straight-line link; a curved relationship can still be strong."),

        T("4.3.2", "Pearson Correlation",
          defn="Pearson's r is the standard correlation coefficient for two numeric columns: CORREL(array1, array2) or PEARSON.",
          simple="Square it to get the share of variation the two columns share: r = 0.58 means about a third.",
          fx=[("=CORREL(tCV[Lecture_Attendance_Pct],tCV[Final_Score])", f"{R['Lecture_Attendance_Pct']:.3f}", "Attendance and score"),
              ("=CORREL(tCV[Study_Hours_Wk],tCV[Final_Score])", f"{R['Study_Hours_Wk']:.3f}", "Study hours and score"),
              ("=CORREL(tCV[Sleep_Hrs_Night],tCV[Final_Score])", f"{R['Sleep_Hrs_Night']:.3f}", "Sleep and score")],
          fx_widths=[3.6, 0.8, 1.6],
          watch="Pearson is pulled by outliers and only sees straight lines; check the scatter (4.3.6)."),

        T("4.3.3", "Spearman Rank Correlation",
          defn="Spearman's rho is Pearson's r calculated on the ranks of the values instead of the values themselves.",
          simple="It measures whether one column rises as the other rises, in a straight line or not, and resists outliers. In Excel: "
                 "RANK.AVG each column, then CORREL the ranks.",
          grid=dict(label="Pearson and Spearman, exam score (tCV)", widths=[1.8, 1, 1], rows=[
              ["Factor", "Pearson r", "Spearman"]] + [[SHORT[c], f"{R[c]:.3f}", f"{RS[c]:.3f}"] for c in DRIVERS[:4]]),
          watch="When Spearman is clearly above Pearson, the link is steady but not a straight line; describe it that way."),

        T("4.3.4", "Correlation Matrix",
          defn="A correlation matrix shows r for every pair of columns in a square grid, with 1s on the diagonal.",
          simple="Scan one row for what links to the outcome, and the other cells for factors that overlap each other.",
          grid=dict(label="Exam score and three drivers (tCV)", widths=[1.5, 0.9, 0.9, 0.9, 0.9], size=11.5, rows=[
              ["", "Score", "Attend.", "Hours", "Previous"]] + [
              [n] + [f"{MX.iloc[i, j]:.2f}" for j in range(4)] for i, n in enumerate(["Score", "Attend.", "Hours", "Previous"])]),
          watch="The drivers barely correlate with each other here (all near 0): each adds its own information."),

        T("4.3.5", "Building a Correlation Matrix in Excel",
          defn="Data › Data Analysis › Correlation builds the matrix in one step; a CORREL formula in every cell builds a live one.",
          simple="The ToolPak output is static; formulas update when the data changes.",
          steps=["Data › Data Analysis › Correlation", "Input range: the numeric columns, with Labels in first row",
                 "Output range: a blank cell › OK", "Or type =CORREL(col_i, col_j) in each cell for a live matrix"],
          points=["The ToolPak fills only the lower triangle.", "Conditional formatting (colour scale) shows the pattern.",
                  "Keep only numeric columns; text columns break the tool."],
          watch="Non-numeric cells inside the range stop the ToolPak with an error; clean the columns first."),

        T("4.3.6", "Reading Scatterplots Before Trusting r",
          defn="A scatterplot shows every pair of values as a point; it reveals curves, clusters and outliers that r cannot show.",
          simple="Plot first, then calculate: the same r can come from very different pictures.",
          grid=dict(label="Hours studied against exam score (tCV)", widths=[2.6, 1.6], rows=[
              ["Check", "Result"],
              ["r, all students", f"{R['Study_Hours_Wk']:.3f}"],
              [f"r without the {I(HI.sum())} scores of 80+", f"{R_NOHI:.3f}"],
              ["Hours studied range", f"{CV.Study_Hours_Wk.min()} to {CV.Study_Hours_Wk.max()} a week"]]),
          steps=["Select the two columns, x first", "Insert › Scatter", "Add a trendline", "Look for curves, clusters, outliers"],
          split=0.55,
          watch=f"Removing {I(HI.sum())} high scorers changes r; know which points drive it before you quote it."),

        T("4.3.7", "Correlation Is Not Causation",
          defn="Two columns can move together because one causes the other, because both depend on a third factor, or by chance.",
          simple="Correlation tells you where to look; only an experiment or a strong design can show cause.",
          grid=dict(label=f"Atlantic Pearl: lead time and cancellation (tHotel, r = {R_LEAD:.2f})", widths=[1.4, 1.2, 1.2], rows=[
              ["Days booked ahead", "Cancelled", "Bookings"]] + [[k, P(r["mean"]), I(r["size"])] for k, r in LEAD.iterrows()]),
          watch="Lead time does not 'cause' cancellation: early bookers simply have more time to change plans. Act on the risk, "
                "not on a supposed cause."),

        T("4.3.8", "Confounding Variables",
          defn="A confounder is a third variable linked to both columns that creates or inflates the association between them.",
          simple="Test a suspected confounder by comparing like with like: hold it fixed and see if the gap remains.",
          grid=dict(label="Salary gap, female vs male (tSal)", widths=[2.8, 1.6], rows=[
              ["Comparison", "Female minus male"],
              ["Raw average salary", f"{Nm(RAW['coef']['F'], 2)} ({P(GM.sal['Female'] / GM.sal['Male'] - 1)})"],
              ["Average experience", f"{GM.exp['Female'] - GM.exp['Male']:+.1f} years"],
              ["Same experience (regression)", f"{Nm(ADJ['coef']['F'], 2)} (p {pv(ADJ['p']['F'])})"]]),
          watch="Experience explains about half the raw gap; the rest is not explained by experience. Report both numbers."),

        T("4.3.9", "Simpson's Paradox",
          defn="Simpson's paradox is when a trend that appears in every group reverses when the groups are combined.",
          simple="It happens when the groups are of very different sizes; always check the split before the total.",
          grid=dict(label="Atlantic Pearl cancellation rate: Deluxe vs Standard Twin (tHotel)", widths=[1.5, 1.2, 1.4], rows=[
              ["Arrivals", "Deluxe", "Standard Twin"],
              ["2023", P(SPY.loc[(2023, "Deluxe"), "mean"]), P(SPY.loc[(2023, "Standard Twin"), "mean"])],
              ["2024", P(SPY.loc[(2024, "Deluxe"), "mean"]), P(SPY.loc[(2024, "Standard Twin"), "mean"])],
              ["Both years", P(SPA["Deluxe"]), P(SPA["Standard Twin"])]]),
          watch="Deluxe cancels less in each year but more overall, because most Deluxe bookings fall in 2024, the high-cancellation year."),

        T("4.3.10", "Reporting Correlation Responsibly",
          defn="A responsible correlation statement gives r, the number of rows, the shape seen in the scatter, and states that it "
               "shows association, not cause.",
          simple="Name the population, the period and any confounders you checked.",
          grid=dict(label="Two versions of the same finding", widths=[1.1, 3.7], size=11.5, rows=[
              ["Version", "Wording"],
              ["Misleading", "Attendance drives exam results: raising attendance will raise scores."],
              ["Responsible", f"Across {I(len(CV))} Kingsway students in 2024/25, attendance and exam score are strongly associated "
                              f"(r = {R['Lecture_Attendance_Pct']:.2f}); this shows association, not cause."]]),
          watch="Never write 'drives', 'causes' or 'leads to' about a correlation alone."),
    ],
    takeaways=["r measures straight-line association from -1 to +1, nothing more.",
               "Spearman ranks the values: robust to outliers and curves.",
               "A matrix shows every pair; plot the scatter before trusting r.",
               "Confounders and Simpson's paradox can create or reverse a link.",
               "Report r, n, shape and 'association, not cause'."],
)


M44 = dict(
    id="4.4", title="Linear Regression",
    subtitle="Fit a line, read R-squared and the coefficients, check the residuals, and predict with honest uncertainty.",
    objectives=["Explain what a regression estimates and when to use it.",
                "Fit simple and multiple regressions with LINEST or the ToolPak.",
                "Read coefficients, R-squared, adjusted R-squared and p-values.",
                "Check residuals and multicollinearity.",
                "Predict with a range, and report the model's limits."],
    topics=[
        T("4.4.1", "What Linear Regression Does",
          defn="Linear regression fits the straight line (or plane) that best predicts an outcome from one or more inputs, by "
               "minimising the squared prediction errors.",
          simple="It answers 'how much does the outcome change, on average, when an input rises by one unit?'",
          grid=dict(label="Salary survey (tSal): the question and the answer", widths=[1.5, 3.3], rows=[
              ["Part", "In this example"],
              ["Outcome (y)", "Annual_Salary_NGN"],
              ["Input (x)", "Years_of_Experience"],
              ["Answer", f"Each extra year goes with about {N(B1)} more salary, on average"]]),
          watch="A regression describes the data it was fitted on; it is not proof of cause."),

        T("4.4.2", "Simple Linear Regression",
          defn="Simple linear regression uses one input: y = a + b x, where a is the intercept and b the slope.",
          simple="SLOPE, INTERCEPT and RSQ give the three key numbers; a scatter trendline shows the same line.",
          fx=[("=SLOPE(tSal[Annual_Salary_NGN],tSal[Years_of_Experience])", N(B1), "b: per year of experience"),
              ("=INTERCEPT(tSal[Annual_Salary_NGN],tSal[Years_of_Experience])", N(A1), "a: predicted at 0 years"),
              ("=RSQ(tSal[Annual_Salary_NGN],tSal[Years_of_Experience])", f"{S1['r2']:.3f}", "R-squared")],
          fx_widths=[3.8, 1.1, 1.6],
          watch="The intercept is a prediction at x = 0; it is only meaningful if x = 0 is in the data (here it is: new entrants)."),

        T("4.4.3", "Multiple Linear Regression",
          defn="Multiple regression uses several inputs at once; each coefficient is the effect of its input with the others held "
               "fixed.",
          simple="Text categories enter as 0/1 dummy columns (Is_Masters, Is_PhD); the left-out level is the baseline.",
          grid=dict(label="Salary on four inputs (tSal, LINEST)", widths=[1.9, 1.4, 1.3], rows=[
              ["Input", "Coefficient", "p-value"]] + [
              [k.replace("_", " "), N(C4[k]), pv(M4["p"][k])] for k in ["Intercept", "Years_of_Experience", "Age", "Is_Masters", "Is_PhD"]]),
          watch="Baseline = Bachelor's or High School. 'Is_PhD = ₦2.86m' means compared with that baseline, not with a Master's."),

        T("4.4.4", "The Regression Equation and Coefficients",
          defn="The regression equation writes the prediction as the intercept plus each coefficient times its input.",
          simple="Read each coefficient in the outcome's units per one unit of the input.",
          grid=dict(label="Salary equation (tSal)", widths=[1, 3.8], size=11.5, rows=[
              ["Model", "Equation"],
              ["Simple", f"Salary = {N(A1)} + {N(B1)} x Years"],
              ["Multiple", f"Salary = {N(C4['Intercept'])} + {N(C4['Years_of_Experience'])} x Years {N(C4['Age']).replace('-', '- ')} x Age "
                           f"+ {N(C4['Is_Masters'])} x Masters + {N(C4['Is_PhD'])} x PhD"]]),
          watch="The years coefficient drops from simple to multiple because education and age now share the credit."),

        T("4.4.5", "R-Squared",
          defn="R-squared is the share of the variation in the outcome that the model explains, from 0 to 1.",
          simple="RSQ for one input; LINEST(...,TRUE,TRUE) or the ToolPak for several.",
          grid=dict(label="How much of salary variation is explained?", widths=[2.6, 1.2], rows=[
              ["Model", "R-squared"],
              ["Experience only", P(S1["r2"])],
              ["Experience, age, Master's, PhD", P(M4["r2"])],
              ["Unexplained by the fuller model", P(1 - M4["r2"])]]),
          watch="A high R-squared does not make the model right, and a low one is normal for human behaviour."),

        T("4.4.6", "Adjusted R-Squared",
          defn="Adjusted R-squared is R-squared penalised for the number of inputs, so it only rises if a new input adds real "
               "explanatory power.",
          simple="Compare models with adjusted R-squared, not plain R-squared.",
          grid=dict(label="Does Age earn its place?", widths=[2.6, 1, 1], rows=[
              ["Model", "R-squared", "Adjusted"],
              ["Years, Master's, PhD", f"{M3['r2']:.4f}", f"{M3['adj']:.4f}"],
              ["Years, Age, Master's, PhD", f"{M4['r2']:.4f}", f"{M4['adj']:.4f}"]],
              note="Adjusted = 1 - (1 - R2) x (n - 1) / (n - k - 1)"),
          watch="Adding Age raises adjusted R-squared by only 0.0002: it adds almost nothing once experience is in."),

        T("4.4.7", "P-Values and Statistical Significance",
          defn="A coefficient's p-value is the probability of an estimate at least this far from zero if the input had no effect; "
               "below 0.05 is the usual threshold for 'significant'.",
          simple="p = T.DIST.2T(ABS(coefficient / standard error), n - k - 1).",
          grid=dict(label="Multiple regression p-values (tSal)", widths=[1.9, 1.3, 1.6], rows=[
              ["Input", "p-value", "Verdict"]] + [
              [k.replace("_", " "), pv(M4["p"][k]), "Significant" if M4["p"][k] < 0.05 else "Not significant"]
              for k in ["Years_of_Experience", "Age", "Is_Masters", "Is_PhD"]]),
          watch="Significant means 'unlikely to be zero', not 'large' or 'important'; read the coefficient's size too."),

        T("4.4.8", "Residual Analysis",
          defn="A residual is actual minus predicted. Residuals should scatter randomly around zero; a pattern means the model "
               "misses something.",
          simple="Plot residuals against the input, or average them by band.",
          grid=dict(label="Average residual by experience (simple model)", widths=[1.4, 1.6], rows=[
              ["Years", "Avg residual"]] + [[k, Nm(v, 2)] for k, v in RESB.items()]),
          watch="Negative at both ends, positive in the middle: salary rises then flattens, so a straight line overpredicts the extremes."),

        T("4.4.9", "Making Predictions with a Regression",
          defn="A prediction plugs input values into the equation; a rough 95% range is the prediction plus or minus 2 standard errors "
               "of the estimate.",
          simple="Always give the range, not just the point.",
          grid=dict(label="Predict: 10 years' experience, age 35, Master's", widths=[2.4, 2], rows=[
              ["Item", "Value"],
              ["Point prediction", N(PRED)],
              ["Standard error of estimate", N(M4["sey"])],
              ["Rough 95% range", f"{Nm(PRED - PI)} to {Nm(PRED + PI)}"]]),
          watch=f"Do not predict outside the data: experience runs 0 to {XMAX:.0f} years, so a 40-year profile is extrapolation."),

        T("4.4.10", "Multicollinearity",
          defn="Multicollinearity is when inputs are strongly correlated with each other, so the model cannot separate their effects.",
          simple="Check the correlation between inputs; a variance inflation factor (VIF) above about 5 is a warning.",
          grid=dict(label="Age and experience in tSal", widths=[2.4, 1.6], rows=[
              ["Check", "Value"],
              ["r (Age, Years_of_Experience)", f"{R_AGE_EXP:.3f}"],
              ["VIF = 1 / (1 - r squared)", F(VIF)],
              ["Age coefficient (multiple)", f"{N(C4['Age'])} (p {pv(M4['p']['Age'])})"]]),
          watch="The negative Age coefficient is not 'older people earn less': age and experience overlap, so drop one."),

        T("4.4.11", "Regression Limitations",
          defn="A regression assumes a straight-line relationship, independent errors with constant spread, and inputs measured "
               "without bias; it describes association in the sample.",
          simple="Know what the model cannot tell you before you use it.",
          grid=dict(label="Limits of the salary model", widths=[1.8, 3], rows=[
              ["Limit", "In this model"],
              ["Linearity", "Residuals curve at both ends (4.4.8)"],
              ["Range", f"0 to {XMAX:.0f} years only"],
              ["Omitted inputs", "Job family, city and employer are not in the model"],
              ["Cause", "Survey data: association, not cause"]]),
          watch="A model with omitted inputs can give biased coefficients; say which inputs were left out."),

        T("4.4.12", "Reporting a Regression Honestly",
          defn="An honest regression report states the outcome, inputs, sample, coefficients with their uncertainty, adjusted "
               "R-squared, the residual checks and the limits.",
          simple="Lead with the plain-language finding, then the numbers.",
          grid=dict(label="Summary for HR", widths=[1.2, 3.6], size=11.5, rows=[
              ["Part", "Wording"],
              ["Finding", f"Each year of experience goes with about {Nm(C4['Years_of_Experience'], 2)} more salary; a PhD adds about "
                          f"{Nm(C4['Is_PhD'], 2)} over a Bachelor's."],
              ["Fit", f"The model explains {P(M4['adj'], 0)} of the variation (adjusted R-squared)."],
              ["Limits", f"{I(len(SAL))} survey respondents; association only; curve at the extremes."]]),
          watch="Never present a coefficient as a promise: 'each year adds ₦0.57m' is an average, not a pay rule."),
    ],
    takeaways=["Regression estimates the average change in y per unit of x.",
               "Multiple regression holds the other inputs fixed; dummies need a baseline.",
               "Compare models with adjusted R-squared; read p-values and sizes together.",
               "Patterned residuals and correlated inputs are warnings.",
               "Predict with a range, inside the data's range only."],
)

for _m in (M43, M44):
    for _t in _m["topics"]:
        if _t.get("fx"):
            _t.setdefault("stack", True)

MODULES = [M43, M44]
