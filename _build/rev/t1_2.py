from revlib import T

TOPICS = [
    T("1.2.1",
      what="**Descriptive analytics** answers the simplest question: **what happened?** It turns raw rows into totals, "
           "counts, averages, percentages and simple charts for a period. Monthly sales reports, attendance summaries and "
           "hospital bed reports are all descriptive. Every other type of analytics starts here, because you cannot "
           "explain or predict something you have not measured.",
      like="It is like your bank statement at month end. It doesn't tell you why you overspent. It just shows clearly "
           "what came in, what went out and what is left.",
      how=["Pick the measure (e.g. Sales) and the period (e.g. 2025).",
           "Group the rows by something useful (zone, store, month).",
           "Add totals and a share of total, then label units and period."],
      tip="Always show the total alongside the parts so readers can see the scale.",
      trap="Listing raw rows and calling it a report; nobody can read 3,000 lines.",
      who="SuperMart Retail",
      case="Head Office asks: how did each geopolitical zone perform in 2025? The analyst summarises the 3,312 sales lines "
           "of 2025 by the Zone column.",
      steps=["Filter Order_Date to 2025.",
             "For each Zone, count distinct orders and add up Sales.",
             "Divide each zone's sales by the ₦733.2m total to get its share."],
      table=[["Zone (FY2025)", "Orders", "Sales", "Share"],
             ["South-West", "538", "₦250.1m", "34.1%"],
             ["South-South", "470", "₦213.1m", "29.1%"],
             ["North-Central", "406", "₦147.1m", "20.1%"],
             ["South-East", "273", "₦122.9m", "16.8%"]],
      so="The South-West brings in about one naira in every three. This is a clear picture of what happened, but it does "
         "not yet say why."),

    T("1.2.2",
      what="**Diagnostic analytics** asks **why did it happen?** After descriptive analytics spots something (a fall, a "
           "spike, a gap), you dig deeper. You split the data into groups, compare them, and look for the group where the "
           "problem is concentrated. The aim is to find the cause, not only the symptom.",
      like="Your car's fuel is finishing faster than usual. Diagnosis means checking: is it the traffic, the tyres, or a "
           "leak? You compare each possible cause until one stands out.",
      how=["Start from the result you want to explain (e.g. 901 students at risk).",
           "Split it by a likely cause (attendance, study hours, faculty).",
           "Compare the rate in each group and look for the big gap."],
      tip="Compare rates (percentages), not raw counts, because groups have different sizes.",
      trap="Stopping at the first explanation without comparing it with the other groups.",
      who="Crestview University",
      case="Crestview has 901 students at risk out of 6,607. Is the problem spread evenly, or is there a clear cause? The "
           "analyst tests attendance first.",
      steps=["Use the Attendance_Band column to group all 6,607 students.",
             "Count the students and the At_Risk = Yes students in each band.",
             "Divide at-risk students by students to get the at-risk rate."],
      table=[["Attendance", "Students", "At risk", "Rate"],
             ["Below 70%", "1,573", "841", "53.5%"],
             ["70–79%", "1,681", "52", "3.1%"],
             ["80–89%", "1,607", "8", "0.5%"],
             ["90–100%", "1,746", "0", "0.0%"]],
      so="841 of the 901 at-risk students attend fewer than 70% of classes. Poor attendance is the main cause, so it is "
         "where help should go first."),

    T("1.2.3",
      what="**Predictive analytics** asks **what is likely to happen next?** It uses patterns from the past, such as "
           "trends, seasons and relationships, to estimate the future. A forecast is never certain. It is a reasoned "
           "best guess that helps people plan stock, staff and budgets before events happen.",
      like="It is like a farmer planning for the rainy season. He has watched the rains for years, so he plants before the "
           "rain comes, without waiting to be sure.",
      how=["Line up past values in time order (years or months).",
           "Use a trend function, e.g. =FORECAST.LINEAR(next_year, values, years).",
           "Present the result as an estimate, not a promise."],
      tip="Say what the forecast is based on and give it as a round estimate.",
      trap="Treating a forecast as a guaranteed number, or forecasting from only one or two points.",
      who="SuperMart Retail",
      case="The Finance Director needs a first estimate of 2026 sales for the budget. The analyst has four years of yearly "
           "sales from the SuperMart file.",
      steps=["Sum Sales for each year 2022–2025 and list the years in one column.",
             "In a new cell, type =FORECAST.LINEAR(2026, sales_cells, year_cells).",
             "Round the answer and label it 'trend forecast'."],
      table=[["Year", "Sales", "Type"],
             ["2022", "₦484.2m", "Actual"],
             ["2023", "₦470.5m", "Actual"],
             ["2024", "₦609.2m", "Actual"],
             ["2025", "₦733.2m", "Actual"],
             ["2026", "about ₦795.7m", "Trend forecast"]],
      so="If the 4-year trend continues, 2026 sales will be roughly ₦796m. That is a planning figure the budget can start "
         "from."),

    T("1.2.4",
      what="**Prescriptive analytics** asks **what should we do?** It compares possible actions, works out the likely result "
           "of each, and recommends the best one within the real limits of the business (money, staff, rules). It is the "
           "most valuable stage because it ends with a decision.",
      like="Google Maps doesn't only show the traffic (descriptive) or predict your arrival time (predictive). It "
           "recommends the best route. That is prescriptive.",
      how=["List the options you could take (including 'do nothing').",
           "Calculate the result of each option with the same data.",
           "Recommend the option with the best result, and state your assumptions."],
      tip="Always include 'do nothing' as an option so the benefit is clear.",
      trap="Recommending an action without showing its expected result in naira or numbers.",
      who="SuperMart Retail",
      case="In 2025, 455 lines with discounts above 20% lost ₦47.6m. Furniture, whose margin fell to 1.4%, had 167 of "
           "them. Management asks: what if no Furniture discount could go above 20%?",
      steps=["Filter the 2025 Furniture lines where Discount is above 0.2.",
             "Rebuild each line's profit as if its discount had been 0.2 (same list price and cost).",
             "Compare total profit for both options."],
      table=[["Option for the 167 Furniture lines", "Profit FY2025"],
             ["Keep current discounts", "−₦17.5m"],
             ["Cap Furniture discounts at 20%", "+₦0.4m"],
             ["Gain from the cap", "+₦18.0m"]],
      so="If customers still buy at 20% off, the cap turns a ₦17.5m loss into a small profit: about ₦18.0m gained. "
         "The recommendation is to cap Furniture discounts at 20%."),

    T("1.2.5",
      what="The four types work best together, as steps on one problem. **Descriptive** spots the issue, "
           "**diagnostic** finds the cause, **predictive** says who or what is likely to be affected next, and "
           "**prescriptive** recommends the action. Each step uses the answer from the step before it.",
      like="A good doctor goes through all four steps. 'Your blood pressure is high' (what). 'It is the salt' (why). 'It "
           "may cause a stroke' (next). 'Cut salt and take this drug' (action).",
      how=["Write the four questions: what, why, what next, what to do.",
           "Answer them in order on the same dataset.",
           "Make sure each answer leads to the next question."],
      tip="Present the four answers in order; it tells a complete story.",
      trap="Jumping straight to an action without first measuring and explaining the problem.",
      who="EkoSuites Hotel",
      case="EkoSuites Lagos had 36,275 bookings in 2023–2024. The manager feels too many guests cancel. The analyst works "
           "through all four types using EkoSuites_Hotel_Bookings_2023_2024.csv.",
      steps=["Count cancelled bookings and divide by all bookings.",
             "Compare the cancellation rate by Lead_Time_Days (days booked ahead).",
             "Use the pattern to flag risky new bookings and propose an action."],
      table=[["Type", "Question", "Answer"],
             ["Descriptive", "What share cancel?", "32.8% of bookings"],
             ["Diagnostic", "Why?", "181+ days ahead: 73.9% cancel"],
             ["Predictive", "Who next?", "Long lead, no special requests"],
             ["Prescriptive", "What to do?", "Deposit for bookings 90+ days out"]],
      split=4.90,
      so="Together the four types turn a vague feeling ('too many cancel') into a specific policy the hotel can try and "
         "measure."),

    T("1.2.6",
      what="In this bootcamp you will use the same four types on three organisations that appear again and again: "
           "**SuperMart Retail** (12 shops across Nigeria), **Crestview University** (6,607 students) and **Lifecare "
           "General Hospital** (four 130-bed wards). Meeting the same data many times lets you build skill step by step.",
      like="It is like learning to drive on the same few roads near your house. Once you know them well, you can learn "
           "harder moves without getting lost.",
      how=["Open each organisation's main file in 00_Core_Datasets.",
           "For each one, write one 'what', one 'why' and one 'what to do'.",
           "Keep your notes; later weeks build on them."],
      tip="Get to know these three datasets well; every week reuses them.",
      trap="Treating each example as new; the same files come back each week.",
      who="Multi-Sector",
      case="Here is one finding for each of the three organisations, taken through what happened, why it happened, and "
           "what to do. All the numbers come from the bootcamp files.",
      steps=["SuperMart: margin by discount size (sales file, 2025).",
             "Crestview: at-risk rate by attendance band (student file).",
             "Lifecare: 30-day readmission by department (admissions, 2023)."],
      table=[["Organisation", "What happened", "Why", "What to do"],
             ["SuperMart", "Margin fell to 12.7%", "Discounts above 20%", "Cap Furniture at 20%"],
             ["Crestview", "901 students at risk", "Attendance below 70%", "Early attendance alerts"],
             ["Lifecare", "10.3% readmitted", "Oncology highest: 13.2%", "Follow-up calls"]],
      split=4.30,
      so="The same simple steps work in a shop, a university and a hospital. You will repeat them with more skill every "
         "week."),
]
