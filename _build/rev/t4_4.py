from revlib import T

TRY = "ETEH_Week4_Section_Examples.xlsx  ›  sheet '4.4 Regression'  ·  data in D_Salary"
SAL = "ETEH Talent Survey"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", SAL)
    return T(tag, **k)


TOPICS = [
    O("4.4.1",
      what="**Linear regression** fits a straight line, **y = a + b × x**, that predicts a number (like salary) "
           "from one or more inputs (like years of experience). It tells you **how much y changes per unit of x** "
           "and how well the line fits.",
      like="Like a taxi fare rule: a fixed starting charge plus a set amount per kilometre.",
      how=["Choose the outcome (y) and the input (x).", "Fit the line: SLOPE and INTERCEPT.", "Check the fit: RSQ and STEYX."],
      tip="Always plot x against y first.",
      trap="Treating the line as a law of nature.",
      case="HR asks: how does salary rise with years of experience? (1,781 respondents)",
      steps=["SLOPE(salary, experience).", "INTERCEPT(salary, experience).", "RSQ and STEYX for the fit."],
      table=[["Output", "Value", "Meaning"], ["Slope b", "₦614,499", "Extra pay per year of experience"],
             ["Intercept a", "₦5,706,904", "Line's salary at 0 years"], ["R²", "0.67", "Experience explains 67% of the variation"],
             ["Standard error", "₦2,948,324", "Typical miss of a prediction"]],
      so="Predicted salary = ₦5.71m + ₦614,499 × years. The fit is good, but predictions miss by about ₦2.9m.",
      split=4.20),

    O("4.4.2",
      what="**Simple regression** uses **one** input. In Excel use SLOPE / INTERCEPT / RSQ for quick answers, "
           "**=LINEST(y, x, TRUE, TRUE)** for the full statistics, or **Data Analysis > Regression** for a report.",
      like="Like three ways to weigh yourself: a quick bathroom scale, a detailed clinic scale, or a full medical.",
      how=["Quick: =SLOPE, =INTERCEPT, =RSQ.", "Full: =LINEST(y, x, TRUE, TRUE) spills 5 rows × 2 columns.", "Report: ToolPak Regression."],
      tip="LINEST and the functions stay live; the ToolPak output is static.",
      trap="Swapping the y and x arguments.",
      case="The analyst fits salary on experience using LINEST.",
      steps=["Select a 5 × 2 area or let it spill.", "=LINEST(tSal[Annual_Salary_NGN], tSal[Years_of_Experience], TRUE, TRUE).", "Read each row."],
      table=[["LINEST row", "Slope column", "Intercept column"], ["Coefficient", "₦614,499", "₦5,706,904"],
             ["Standard error", "₦10,208", "₦116,832"], ["R² / SE of estimate", "0.67", "₦2,948,324"],
             ["F / degrees of freedom", "3,624", "1,779"]],
      so="LINEST gives the same slope and intercept as SLOPE and INTERCEPT, plus the error figures.",
      split=4.20),

    O("4.4.3",
      what="**Multiple regression** uses **several inputs at once**. Each coefficient shows that input's effect "
           "**with the others held constant**. Text categories (education) need **0/1 dummy columns**.",
      like="Like working out a house price from size, location and age together, rather than one at a time.",
      how=["Put the inputs side by side (Age, Experience, Is_Masters, Is_PhD).",
           "=LINEST(y, tSal[[Age]:[Is_PhD]], TRUE, TRUE).", "Coefficients come out in REVERSE column order."],
      tip="Leave one category out (Bachelor's/High School) as the baseline.",
      trap="Reading LINEST's first coefficient as the first input.",
      case="Add age and education to experience to explain salary.",
      steps=["Dummy columns: Is_Masters, Is_PhD.", "LINEST on the four inputs.", "Label the reversed output."],
      table=[["Input", "Coefficient"], ["Each year of experience", "+₦569,385"], ["Master's (vs baseline)", "+₦1,872,832"],
             ["PhD (vs baseline)", "+₦2,858,830"], ["Each year of age", "-₦34,495"]],
      so="With experience held constant, a PhD goes with about ₦2.86m more pay than the baseline.",
      split=4.90),

    O("4.4.4",
      what="Each **coefficient** = the expected change in the outcome for **one more unit** of that input, with "
           "the others unchanged. A **dummy** coefficient is the gap compared with the baseline group. The "
           "**intercept** is the prediction when all inputs are 0, which may be unrealistic.",
      like="Like a price list: each extra item adds its own amount to the bill.",
      how=["Name the unit: '₦ per year'.", "Dummies: '₦ more than the baseline'.", "Check the intercept makes sense."],
      tip="Write each coefficient as a sentence in business units.",
      trap="Reading the intercept literally when 0 is impossible (age 0).",
      case="The analyst writes a sentence for each coefficient of the multiple model.",
      steps=["Read each coefficient.", "Add the unit and the 'holding others constant' phrase.", "Judge the intercept."],
      table=[["Coefficient", "Value", "Sentence"], ["Experience", "₦569,385", "Per extra year, same age and education"],
             ["Master's", "₦1,872,832", "More than Bachelor's / High School"], ["PhD", "₦2,858,830", "More than Bachelor's / High School"],
             ["Intercept", "₦6,192,397", "Age 0, no experience: not meaningful"]],
      so="Coefficients become plain sentences; the intercept is only a starting point for the line.",
      split=4.20),

    O("4.4.5",
      what="**R²** is the share of variation in the outcome that the model explains, from 0 to 1. Higher is "
           "better, but a high R² doesn't prove cause, and a modest R² can still be useful.",
      like="Like a weather forecast that gets 7 out of 10 days right: useful, not perfect.",
      how=["Simple: =RSQ(y, x).", "Multiple: third row of LINEST.", "Say 'explains X% of the variation'."],
      tip="Report R² with the typical prediction error (STEYX).",
      trap="Chasing R² by adding every column.",
      case="How much salary variation do different models explain?",
      steps=["RSQ with age only.", "RSQ with experience only.", "LINEST R² for the multiple model."],
      table=[["Model", "R²", "Explains"], ["Age only", "0.59", "59% of salary variation"],
             ["Experience only", "0.67", "67%"], ["Experience + age + education", "0.71", "71%"]],
      so="Adding age and education lifts R² only from 0.67 to 0.71: experience does most of the work.",
      split=4.70),

    O("4.4.6",
      what="**Adjusted R²** gives a small penalty for each extra input, so it only rises if the new input adds "
           "real information. Use it to **compare models with different numbers of inputs**.",
      like="Like a recipe judged on taste per ingredient: extra ingredients must earn their place.",
      how=["Adjusted R² = 1 − (1 − R²) × (n − 1) ÷ df.", "df comes from LINEST row 4.", "Compare models on adjusted R²."],
      tip="If adjusted R² barely moves, the extra inputs aren't worth the complexity.",
      trap="Comparing plain R² across models of different sizes.",
      case="Is the four-input model worth it compared with experience alone?",
      steps=["R² and df for each model.", "Compute adjusted R².", "Compare."],
      table=[["Model", "Inputs", "R²", "Adjusted R²"], ["Experience only", "1", "0.671", "0.671"],
             ["Experience, age, Master's, PhD", "4", "0.710", "0.709"]],
      so="Adjusted R² still rises (0.671 to 0.709), so education adds real information.",
      split=4.20),

    O("4.4.7",
      what="A **p-value** answers: if this input truly had no effect, how likely is a coefficient this large? "
           "Below **0.05** we call it **statistically significant**. Significant doesn't mean large or important.",
      like="Like a smoke alarm: it tells you there's probably something there, not how big the fire is.",
      how=["t = coefficient ÷ standard error.", "=T.DIST.2T(ABS(t), df).", "Compare with 0.05."],
      tip="Always report the size of the effect next to the p-value.",
      trap="Dropping an input just because p is slightly above 0.05.",
      case="Which inputs in the multiple model are significant?",
      steps=["t for each coefficient.", "T.DIST.2T with df = 1,776.", "Mark p < 0.05."],
      table=[["Input", "t statistic", "p-value", "Verdict"], ["Experience", "20.6", "Below 0.0001", "Significant"],
             ["PhD", "13.9", "Below 0.0001", "Significant"], ["Master's", "11.9", "Below 0.0001", "Significant"],
             ["Age", "-1.48", "0.14", "Not significant"]],
      so="Experience and education matter; age adds nothing once experience is in the model.",
      split=4.20),

    O("4.4.8",
      what="A **residual** = actual minus predicted. Plot residuals against the input. A **random cloud** is good. "
           "A **curve** means a straight line is the wrong shape; a **funnel** means errors grow with x.",
      like="Like checking a tailor's work: if every suit is tight at the shoulders, the pattern is wrong, not just one suit.",
      how=["Predicted = a + b × x for each row.", "Residual = actual − predicted.", "Scatter residuals vs x."],
      tip="Average the residuals by band of x to see a curve clearly.",
      trap="Trusting a good R² without looking at residuals.",
      case="Residuals from the simple model for 200 sampled respondents (tResid).",
      steps=["Predicted and Residual columns.", "Scatter against experience.", "Average residual by experience band."],
      table=[["Experience band", "People", "Average residual"], ["0–2 years", "32", "-₦1,758,146"],
             ["6–15 years", "86", "+₦832,645"], ["Over 20 years", "13", "-₦2,353,984"]],
      so="The line over-predicts at both ends and under-predicts in the middle: pay rises fast, then levels off.",
      split=4.70),

    O("4.4.9",
      what="To **predict**, put values into the equation: =INTERCEPT + SLOPE × x, or =FORECAST.LINEAR(x, known_y, known_x). "
           "Give a **range**, not just one number: about ± 2 × the standard error. Stay inside the data's range.",
      like="Like a delivery estimate: 'between 2 and 4 pm' is more honest than '3:07 pm'.",
      how=["Type the profile into input cells.", "Predicted value from the equation.", "Range = prediction ± 2 × SE."],
      tip="Keep inputs inside what the data covers (0 to 34 years).",
      trap="Quoting a single exact figure as certain.",
      case="Predict pay for a profile: 6 years' experience, age 31, Master's degree.",
      steps=["Simple model: intercept + slope × 6.", "Multiple model with LINEST coefficients.", "Add ± 2 × SE."],
      table=[["Prediction", "Result"], ["Simple model (experience only)", "₦9,393,897"], ["Multiple model", "₦10,412,180"],
             ["About 95% range (multiple)", "₦4,871,176 to ₦15,953,185"]],
      so="A fair offer is around ₦10.4m, but people like this can earn anywhere from ₦4.9m to ₦16.0m.",
      split=4.90),

    O("4.4.10",
      what="**Multicollinearity** means two inputs carry nearly the same information. Their coefficients become "
           "unstable and can flip sign. Check CORREL between inputs and the **VIF** = 1 ÷ (1 − r²). Above 5 is a "
           "warning, above 10 serious.",
      like="Like two witnesses who saw exactly the same thing: the second one adds little.",
      how=["CORREL between inputs.", "VIF = 1/(1 − r²).", "Drop or combine one of the pair."],
      tip="If a coefficient's sign surprises you, check multicollinearity first.",
      trap="Saying 'older people earn less' from the negative age coefficient.",
      case="Age and experience are both in the model. Is that a problem?",
      steps=["CORREL(Age, Years_of_Experience).", "VIF.", "Compare the age coefficient alone vs in the model."],
      table=[["Measure", "Result"], ["r(Age, Experience)", "0.94"], ["VIF", "8.2: warning"],
             ["Age coefficient on its own", "+₦480,650 per year"], ["Age coefficient with experience", "-₦34,495 (not significant)"]],
      so="Age and experience overlap so much that age flips sign. Keep experience and drop age.",
      split=4.90),

    O("4.4.11",
      what="Regression only describes **straight-line links in the range of data you have**. It doesn't prove "
           "cause, can't predict far outside the data, and is limited by what's missing from the data.",
      like="Like a map of Lagos: very useful in Lagos, useless for finding your way in Kano.",
      how=["Note the data range.", "Name what isn't in the data.", "Check the shape (residuals)."],
      tip="Put the limitations on the same page as the results.",
      trap="Predicting for 40 years' experience when the data stops at 34.",
      case="The analyst lists the limits of the salary model.",
      steps=["Check the range of experience.", "Describe the sample.", "List missing drivers."],
      table=[["Limitation", "In this example"], ["Range", "Experience runs 0–34 years only"],
             ["Sample", "1,781 self-reported salaries, not every professional"], ["Missing drivers", "Company size and city aren't in the data"],
             ["Shape", "Pay levels off; a straight line misses this"]],
      so="The model is useful for typical profiles, not for extreme cases or proof of cause.",
      split=4.90),

    O("4.4.12",
      what="A regression report gives the **equation in business units**, **R²** and the typical error, **n**, "
           "which inputs are **significant**, and the **limitations**. Use 'associated with', not 'causes'.",
      like="Like a product label: what's inside, how much, and the warnings.",
      how=["One sentence for the main effect, with R² and n.", "List significant inputs.", "State the limits."],
      tip="Write for a manager: naira, not 'beta coefficients'.",
      trap="Pasting the raw ToolPak output into a report.",
      case="The analyst writes the summary for HR.",
      steps=["Main-effect sentence.", "Education effect sentence.", "Limitations sentence."],
      table=[["Element", "Example"], ["Main effect", "Each extra year is associated with about ₦614,499 more pay"],
             ["Fit and size", "R² = 0.67, n = 1,781, typical miss ₦2.9m"], ["Other inputs", "PhD +₦2.86m, Master's +₦1.87m (vs Bachelor's or less)"],
             ["Limits", "0–34 years only; survey data; not proof of cause"]],
      so="HR gets numbers they can use, with the warnings attached.",
      split=4.90),
]
