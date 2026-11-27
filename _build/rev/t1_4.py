from revlib import T

TOPICS = [
    T("1.4.1",
      what="Shops and online stores use analytics to decide what to stock, how to price, where to open, and which "
           "customers to reward. Their most common measures, called **KPIs** (key performance indicators), are sales, "
           "number of orders, average order value, profit margin and number of customers.",
      like="A mama-put owner who notices that rice finishes by 1 p.m. but beans are left over is doing retail analytics. "
           "Tomorrow she cooks more rice and less beans.",
      how=["Sum Sales and Profit; count distinct Order_ID and Customer_ID.",
           "Average order value = Sales ÷ Orders.",
           "Margin = Profit ÷ Sales; compare with last year."],
      tip="Report a small, fixed set of KPIs every month so trends are easy to see.",
      trap="Counting sales lines as orders; one order can have many lines.",
      who="SuperMart Retail",
      case="The board wants a one-glance KPI card for SuperMart's 2025 performance, using the 3,312 sales lines of 2025.",
      steps=["Filter Order_Date to 2025.",
             "Sum Sales and Profit, and count unique Order_ID and Customer_ID values.",
             "Calculate average order value and margin."],
      table=[["KPI (FY2025)", "Value"],
             ["Sales", "₦733.2m"],
             ["Orders", "1,687"],
             ["Average order value", "₦434,627"],
             ["Profit margin", "12.7% (13.4% in 2024)"],
             ["Customers", "693"]],
      so="Five numbers summarise a year of trading. The falling margin stands out as the KPI to investigate."),

    T("1.4.2",
      what="Universities and schools use analytics to track attendance, results and progression, and to find students who "
           "need help early, before they fail. Common measures are pass rate, grade distribution, attendance rate and the "
           "number of students at risk.",
      like="A good class teacher notices who is always absent or sleeping in class and calls the parents early, before "
           "exam results come out.",
      how=["Count students by grade or status.",
           "Work out rates (e.g. at-risk ÷ students) for each faculty.",
           "Rank faculties and focus support where rates are highest."],
      tip="Compare rates between faculties, not raw counts.",
      trap="Ranking faculties by number of at-risk students; big faculties always look worse.",
      who="Crestview University",
      case="The Deputy Vice-Chancellor wants to know whether any faculty has a higher share of at-risk students. The analyst "
           "uses Crestview_Student_Performance.csv.",
      steps=["Group the 6,607 students by Faculty.",
             "Count the students and the At_Risk = Yes students in each faculty.",
             "Divide to get each faculty's at-risk rate."],
      table=[["Faculty", "Students", "At-risk rate"],
             ["Arts & Education", "1,014", "14.6%"],
             ["Science", "1,943", "14.1%"],
             ["Engineering", "1,400", "13.9%"],
             ["Social Sciences", "2,250", "12.7%"]],
      so="The rates are close (12.7–14.6%), so the problem is university-wide. Support should target low attendance, "
         "not one faculty."),

    T("1.4.3",
      what="Hospitals use analytics to manage beds, staff and money, and to improve patient care. Key measures are bed "
           "**occupancy** (the share of beds in use), length of stay, readmission rate and billing. High occupancy means "
           "patients may wait for a bed; low occupancy means wasted capacity.",
      like="A hospital is like a hotel for sick people. The manager must know every day how many 'rooms' are full, so "
           "there is always space for the next emergency.",
      how=["Occupancy % = patients in beds ÷ staffed beds.",
           "Average it by month to see busy and quiet periods.",
           "Count the days above a safety level (e.g. 95%)."],
      tip="Watch the busy days, not only the yearly average.",
      trap="Reporting only the annual average; it hides the dangerous peak days.",
      who="Lifecare Hospital",
      case="Lifecare has 4 wards of 130 beds (520 in total). The medical director asks how full the hospital was in 2023, "
           "using Lifecare_Daily_Bed_Census.csv.",
      steps=["Filter Census_Date to 2023.",
             "Average Occupancy_Pct by month and for the whole year.",
             "Count the days when occupancy was 95% or more."],
      table=[["Measure (2023)", "Value"],
             ["Average occupancy", "89.5%"],
             ["Busiest month (June)", "95.2%"],
             ["Quietest month (November)", "82.5%"],
             ["Days at 95% or more", "44 days"]],
      so="On average there was spare room, but on 44 days the hospital was almost full. June needs extra staff and beds "
         "planned."),

    T("1.4.4",
      what="Banks and fintechs use analytics to track payments, spot fraud, manage loans and reduce failed transactions. "
           "A common measure is the **failure rate**: the share of payments that did not go through. Even a small rate "
           "matters when there are thousands of transactions.",
      like="When a POS machine says 'transaction failed' but the money leaves your account, you lose trust. Banks study "
           "failure rates to prevent this.",
      how=["Count transactions by payment method.",
           "Count how many of each failed.",
           "Failure rate = failed ÷ total, for each method."],
      tip="Always show the count next to the rate so small groups are not over-read.",
      trap="Comparing raw failed counts; the busiest method will always have the most failures.",
      who="SuperMart Online",
      case="SuperMart Online's finance team asks which payment method fails most often. The analyst checks payment_status "
           "in all 11,000 transactions.",
      steps=["Count transactions by payment_method.",
             "Count those with payment_status = Failed.",
             "Divide failed by total for each method."],
      table=[["Payment method", "Transactions", "Failed", "Rate"],
             ["Card", "347", "12", "3.5%"],
             ["USSD", "764", "23", "3.0%"],
             ["Bank Transfer", "2,182", "64", "2.9%"],
             ["Paystack", "3,133", "88", "2.8%"],
             ["Cash on Delivery", "1,928", "48", "2.5%"]],
      so="Paystack has the most failures (88) only because it is busiest; card has the highest rate. Both are below 4%."),

    T("1.4.5",
      what="Government and public bodies use analytics to plan services, target help to people who need it most, and "
           "measure whether programmes work. Health screening, school enrolment and immunisation coverage are common "
           "examples. Because budgets are limited, data shows where each naira does the most good.",
      like="When an LGA has only enough vaccines for some villages, data on where children live and who is unvaccinated "
           "tells them which villages to visit first.",
      how=["Group people by a factor such as age or location.",
           "Work out the rate of the condition in each group.",
           "Focus the programme on the groups with the highest rates."],
      tip="Target limited resources at the groups with the highest rates.",
      trap="Spreading a small budget evenly when the need is very uneven.",
      who="Lifecare Diabetes Screening",
      case="A state health programme can fund free diabetes tests for only part of the population. Lifecare_Diabetes_"
           "Screening.csv holds 100,000 screening results to guide the choice.",
      steps=["Group the 100,000 people into age bands.",
             "Count people with Diabetes = Yes in each band.",
             "Divide to get the diabetes rate per band."],
      table=[["Age band", "Screened", "Diabetes rate"],
             ["Under 30", "32,435", "0.8%"],
             ["30–44", "19,973", "4.2%"],
             ["45–59", "22,537", "10.9%"],
             ["60 and over", "25,055", "19.7%"]],
      so="Diabetes is about 25 times more common at 60+ than under 30. Free tests should go first to people aged 45 and "
         "over."),

    T("1.4.6",
      what="Marketing teams use analytics to understand customer groups (called **segments**), measure which campaigns "
           "work, and decide where to spend. A key idea is value per customer: some groups are smaller but each customer "
           "spends much more.",
      like="A suya seller knows the office workers who buy every Friday spend more each time than passers-by. He makes "
           "sure they are served first.",
      how=["Group customers by Segment.",
           "Count unique customers and sum sales for each segment.",
           "Sales per customer = segment sales ÷ customers."],
      tip="Look at both total sales and sales per customer.",
      trap="Judging a segment by total sales only; a small group can be the most valuable.",
      who="SuperMart Retail",
      case="The marketing manager has a budget for a loyalty scheme. Which customer segment is most valuable? The "
           "analyst uses the Segment column for 2025.",
      steps=["Filter to 2025 and group by Segment.",
             "Count unique Customer_ID and sum Sales for each segment.",
             "Divide sales by customers."],
      table=[["Segment (FY2025)", "Customers", "Sales", "Per customer"],
             ["Consumer", "361", "₦331.9m", "₦919,404"],
             ["Corporate", "204", "₦241.8m", "₦1,185,529"],
             ["Home Office", "128", "₦159.5m", "₦1,245,803"]],
      so="Home Office is the smallest group but each customer spends the most, so it is a strong target for the loyalty "
         "scheme."),

    T("1.4.7",
      what="**People analytics** (HR analytics) uses data about staff to make fairer, better decisions on hiring, pay, "
           "promotion and keeping good people. Pay is often compared using the **median**: the middle value when "
           "salaries are sorted, which is not pulled up by a few very high earners.",
      like="If a class of 10 includes one child of a billionaire, the average pocket money looks huge. The median, the "
           "middle child, shows what is normal.",
      how=["Group staff by a factor such as experience or role.",
           "Use =MEDIAN() of salary for each group.",
           "Compare groups and look for unfair gaps."],
      tip="Use the median for pay; it resists extreme salaries.",
      trap="Using the average salary, which a few top earners can distort.",
      who="ETEH Talent Salary Survey",
      case="An HR manager sets a pay scale for new analysts. The analyst uses the 1,781 replies in ETEH_Talent_Salary_"
           "Survey.csv to see how pay grows with experience.",
      steps=["Create experience bands from Years_of_Experience.",
             "Use MEDIAN on Annual_Salary_NGN for each band.",
             "Check the median for Job_Title = Data Analyst."],
      table=[["Experience", "Respondents", "Median salary"],
             ["0–2 years", "306", "₦4.4m"],
             ["3–5 years", "368", "₦7.0m"],
             ["6–10 years", "452", "₦11.0m"],
             ["11+ years", "655", "₦16.7m"]],
      so="Median pay almost quadruples from early career to 11+ years. That gives HR a fair, evidence-based scale."),

    T("1.4.8",
      what="Logistics companies use analytics to deliver on time at the lowest cost. They track delivery time, the "
           "share of late deliveries, and performance by courier or route. Comparing couriers on the same measure shows "
           "who is reliable.",
      like="If you send goods from Lagos to Onitsha every week, you soon learn which motor park and which bus company "
           "gets them there fastest and safest.",
      how=["Delivery days = delivered date − shipped date.",
           "Average the days and work out the % over the promise (e.g. 5 days).",
           "Compare couriers and routes."],
      tip="Measure the late share as well as the average.",
      trap="Judging a courier on a few complaints instead of all deliveries.",
      who="SuperMart Online",
      case="SuperMart Online uses 8 couriers. The operations head asks which are fastest and which are most often late "
           "(more than 5 days), using the 8,286 delivered orders.",
      steps=["Open SuperMart_Online_deliveries.csv and keep delivery_status = Delivered.",
             "Add a column: delivered_date − shipped_date.",
             "Average it, and work out the % above 5 days, for each courier."],
      table=[["Courier", "Avg days", "Over 5 days"],
             ["Gokada", "3.9", "26.0%"],
             ["FedEx Lagos", "3.9", "27.2%"],
             ["DHL Nigeria", "4.0", "28.6%"],
             ["UPS Abuja", "4.1", "30.2%"],
             ["ABC Transport", "4.1", "31.7%"]],
      so="More than 1 delivery in 4 takes over 5 days with every courier. ABC Transport and UPS Abuja are late most often, so review them first."),

    T("1.4.9",
      what="Mining and oil companies use analytics to decide where to dig or drill. They record what is found at each "
           "depth in each hole (the **grade**: how much valuable mineral is in the rock) and keep only the zones worth "
           "the cost of extraction. Getting it wrong costs millions.",
      like="Before buying land, a farmer tests the soil in several spots. He farms where the soil is rich and leaves the "
           "rocky parts.",
      how=["Record each sample: hole, depth from–to, grade.",
           "Set a cut-off grade below which mining does not pay.",
           "Flag each interval as 'ore' or 'waste' with IF."],
      tip="Agree the cut-off grade with engineers before classifying.",
      trap="Averaging grades without weighting by interval length.",
      who="Iron-ore site, Kogi (illustrative figures)",
      case="A geologist has four core samples from one drill hole near Itakpe. Mining pays only above 30% iron (the "
           "cut-off). These figures are for practice only.",
      steps=["List each interval with its iron grade (Fe %).",
             "In a new column: =IF(grade>=30,\"Ore\",\"Waste\").",
             "Add up the metres classified as ore."],
      table=[["Depth (m)", "Fe grade", "Class"],
             ["0–12", "18%", "Waste"],
             ["12–30", "36%", "Ore"],
             ["30–41", "41%", "Ore"],
             ["41–55", "22%", "Waste"]],
      so="29 of the 55 metres (12–41 m) are worth mining. One IF formula turns raw samples into a dig-or-skip "
         "decision."),

    T("1.4.10",
      what="Data-driven decisions change results when the organisation acts on the finding and then measures again. "
           "The pattern is always the same: a clear problem, a finding from the data, a specific action, and a number to "
           "track. The four bootcamp organisations each have one such finding waiting in their files.",
      like="A footballer who watches replays sees he always misses shots with his left foot. He trains his left foot, "
           "and the next season's goals show if it worked.",
      how=["State the problem with a number.",
           "Show the finding that explains it.",
           "Propose one action and the number you will track."],
      tip="Always name the number you will track after the action.",
      trap="Making a change without a 'before' number to compare with.",
      who="Multi-Sector",
      case="Each bootcamp organisation has a clear problem, a finding hidden in its data, and an action it could take. "
           "You will reproduce all four findings during the course.",
      steps=["SuperMart: profit by discount size (2025 sales).",
             "EkoSuites: cancellation rate by days booked ahead.",
             "Crestview and Lifecare: at-risk by attendance; readmission by department."],
      table=[["Organisation", "Finding", "Action"],
             ["SuperMart", "Lines over 20% off lost ₦47.6m", "Cap Furniture discounts at 20%"],
             ["EkoSuites", "181+ days ahead: 73.9% cancel", "Deposit on long bookings"],
             ["Crestview", "53.5% of low attenders at risk", "Attendance alerts"],
             ["Lifecare", "Oncology readmits 13.2%", "Follow-up calls"]],
      split=4.30,
      so="Each finding leads to one action and one number to watch. That is what 'data-driven' means in practice."),
]
