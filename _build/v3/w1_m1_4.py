"""Week 1 edition-3 slide copy: Modules 1.1 to 1.4 (analytics foundations).

Each topic only carries the parts it needs (see v3kit.T). All figures are taken from the named tables in
_build/w1_data.py; table names are the Excel table names used in the Week 1 examples workbook.
"""
from v3kit import T

neg = lambda d: d[d.Profit < 0]  # noqa: E731

M11 = dict(
    id="1.1", title="Introduction to Data Analytics",
    subtitle="Data versus information, what analytics is, who does it, and how to think and question like an analyst.",
    objectives=["Define data, information and data analytics, and tell data from information.",
                "Explain why organisations analyse data, with everyday and business examples.",
                "Describe the data analyst's role and how it differs from data scientist and business analyst.",
                "Turn a business problem into a clear analytical question.",
                "Use evidence, general knowledge and domain knowledge instead of assumptions."],
    topics=[
        T("1.1.1", "What is Data?",
          defn="Data are raw facts (numbers, text, dates or codes) recorded about events, people or things, before "
               "anyone has processed or interpreted them.",
          simple="Data is whatever gets written down when something happens: a sale at a till, a patient's age, a "
                 "student's score. Storing it neatly in rows and columns does not change that: each row is still one "
                 "raw event.",
          like="Data is like the loose ingredients in a market basket: rice, tomatoes, pepper and onions. Useful, but "
               "nobody calls them a meal yet.",
          ex=dict(t="tIkejaLines", cols=["Order_ID", "Order_Date", "Category", "Sub_Category", "Sales", "Quantity", "Profit"], n=3,
                  note="This table is structured data, not information. The headers say what each value is, but nothing "
                       "has been summed, compared or interpreted: each row is one sale as the till recorded it."),
          watch="A neat table is not automatically information. Ask: has anything been summarised, calculated or "
                "compared to answer a question? If not, it is still data."),

        T("1.1.2", "What is Information?",
          defn="Information is data that has been processed (summarised, calculated or compared) and put in context, "
               "so that it answers a question and is useful for a decision.",
          simple="We turn data into information by processing it: totals, averages, percentages, comparisons with a "
                 "target. 'Lagos - Ikeja sold ₦146.4m in FY2025' is information.",
          like="If data is the raw ingredients, information is the pot of jollof rice: the same ingredients, cleaned, "
               "cooked and served so anyone can enjoy them.",
          ex=dict(t="tStores25", cols=["Store_Name", "Zone", "Sales", "Target", "Attainment"], n=4, sort="Sales",
                  note="Also a table, but now information: 3,312 raw sales lines were summed into 12 store totals and "
                       "compared with targets (top 4 by sales shown).")),

        T("1.1.3", "Difference Between Data and Information",
          simple="Data is the input; information is the output. Both can sit in a table: the difference is processing, "
                 "not layout.",
          diagram=("pyramid", dict(levels=["Data", "Information", "Knowledge", "Wisdom"],
                                   notes=["474430 · 11 · Furniture · 2025-01-01 (one sales line)",
                                          "Lagos - Ikeja sold ₦146.4m in FY2025: 103.1% of target",
                                          "Lines discounted above 20% lost ₦47.6m in FY2025",
                                          "Decision: cap Furniture discounts at 20%"])),
          grid=dict(label="Data vs information", widths=[1.2, 2.2, 2.2], size=12, rows=[
              ["", "Data", "Information"],
              ["Form", "Records as captured: 3,312 FY2025 lines", "Summaries: 12 store totals vs target"],
              ["Use", "Input to analysis; proof", "Output for decisions"],
              ["Test", "Nothing summed or compared", "Answers a question"]]),
          order=["diagram", "grid"], diagram_label="The DIKW pyramid: from data to wisdom"),

        T("1.1.4", "What is Data Analytics?",
          defn="Data analytics is the process of examining, cleaning, transforming and modelling data to discover useful "
               "information, draw conclusions and support decision-making.",
          simple="Analytics is asking a question, finding the data that answers it, working it out carefully and "
                 "explaining what it means, so someone can make a better decision.",
          like="A doctor does not guess. She orders tests (data), reads the results (analysis) and prescribes "
               "(decision). Analytics does the same for organisations.",
          ex=dict(t="tYears", cols=["Year", "Sales", "Profit", "Margin", "Growth"], n=4,
                  note="Analytics turns 9,994 lines into one story: sales grew 20.4% in 2025 but margin slipped from 13.4% to 12.7%.")),

        T("1.1.5", "Why Do We Analyse Data?",
          simple="Organisations analyse data to stop guessing. Data shows what is working, what is failing and where to "
                 "put limited money, staff and time.",
          points=["Measure performance against targets", "Spot problems early, before they grow",
                  "Explain the causes of good and bad results", "Forecast demand and plan ahead",
                  "Choose between options and check the result afterwards"],
          points_label="FIVE REASONS",
          ex=dict(t="tStores25", cols=["Store_Name", "Sales", "Profit", "Margin"], rows=neg, n=5, sort="Profit", asc=True,
                  note="The chain made ₦93.4m profit in FY2025, yet 5 of 12 stores lost money. Only analysis reveals this."),
          order=["points", "ex"], split=0.42),

        T("1.1.6", "Data Analytics in Everyday Life",
          simple="You already use analytics daily. Every time a decision is based on past numbers rather than a hunch, "
                 "that is analytics.",
          grid=dict(label="Everyday examples", widths=[1.6, 2.4, 2.4], rows=[
              ["Situation", "Data used", "Decision"],
              ["Bank app spending summary", "Your debit transactions by category", "Cut spending on the biggest category"],
              ["Map app traffic", "Live speed of phones on each road", "Take the faster route to work"],
              ["Online shop suggestions", "What similar shoppers bought", "Show products you are likely to buy"],
              ["Weather forecast", "Years of temperature and rainfall records", "Carry an umbrella today"],
              ["Fuel budgeting", "Litres bought and kilometres driven", "Plan the monthly fuel budget"]])),

        T("1.1.7", "Data Analytics in Business and Organisations",
          simple="Every department asks questions that data can answer. These are the five organisations whose data you "
                 "will use throughout this bootcamp.",
          grid=dict(label="Our case organisations", widths=[1.7, 2.6, 2.0], size=12, rows=[
              ["Organisation", "Typical question", "Data"],
              ["SuperMart Retail (12 stores)", "Which stores missed their sales target?", "9,994 sales lines, 2022-2025"],
              ["SuperMart Online marketplace", "Which payment methods do buyers prefer?", "11,000 orders, 5 linked tables"],
              ["Kingsway University", "Which students are at risk of failing?", "5,840 student records"],
              ["Mercyland Specialist Hospital", "Which departments have most readmissions?", "58,856 admissions, 2019-2024"],
              ["Atlantic Pearl Hotel, Lekki", "Which booking channels cancel most?", "38,400 reservations, 2023-2024"]])),

        T("1.1.8", "The Role of a Data Analyst",
          defn="A data analyst collects, cleans, analyses and interprets data to answer business questions, and "
               "communicates the findings to support decisions.",
          like="An analyst is like a translator between the numbers and the managers. The data speaks one language; "
               "decision-makers speak another.",
          points=["Agree the question with the stakeholder", "Find, import and clean the data",
                  "Calculate, compare and spot patterns", "Build reports, charts and dashboards",
                  "Explain findings and recommend actions"],
          points_label="WHAT AN ANALYST DOES",
          grid=dict(label="Skills", widths=[1, 2.2], rows=[
              ["Type", "Examples"],
              ["Technical", "Excel, Power Query, pivot tables, SQL, Power BI, basic statistics"],
              ["Business", "Domain knowledge, asking good questions, understanding targets"],
              ["People", "Clear writing, presenting, listening, honesty about limits"]]),
          order=["points", "grid"], split=0.45),

        T("1.1.9", "Data Analyst vs Data Scientist vs Business Analyst",
          simple="All three work with data, but they answer different questions and produce different outputs.",
          like="In a football club, the analyst reports last match's statistics, the data scientist builds a model that "
               "predicts injuries, and the business analyst redesigns how tickets are sold.",
          grid=dict(label="Three roles compared", widths=[1.2, 2, 2, 2], size=12, rows=[
              ["", "Data Analyst", "Data Scientist", "Business Analyst"],
              ["Main question", "What happened and why?", "What will happen? Can we automate it?", "How should the process or system work?"],
              ["Typical work", "Clean data, summarise, report, dashboards", "Statistics, machine learning, prediction models", "Gather requirements, map processes, specify changes"],
              ["Main tools", "Excel, SQL, Power BI", "Python or R, SQL, cloud platforms", "Process maps, documents, workshops, Excel"],
              ["Output", "Reports and recommendations", "Models and predictions", "Requirements and process designs"]])),

        T("1.1.10", "Analytical Thinking and Problem-Solving",
          defn="Analytical thinking is the ability to break a complex problem into smaller parts, examine the evidence "
               "for each part and reason logically to a conclusion.",
          like="A mechanic does not replace the whole engine because of a noise. He checks one part at a time until "
               "he finds the cause.",
          diagram=("flow", dict(steps=["Define the problem", "Break it into parts", "Gather evidence", "Test each idea",
                                       "Conclude and act"])), diagram_label="The problem-solving cycle",
          grid=dict(label="Breaking down 'Why is profit low?'", widths=[1.6, 3], size=12, rows=[
              ["Part to check", "Evidence (FY2025)"],
              ["Which stores?", "5 of 12 stores lost money; Abuja - Wuse lost ₦8.8m"],
              ["Which category?", "Furniture margin only 1.4% (Technology 18.7%)"],
              ["Which discounts?", "Lines discounted above 20% lost ₦47.6m"]]),
          order=["diagram", "grid"]),

        T("1.1.11", "Asking the Right Questions Before Analysing Data",
          simple="A clear question saves hours. Before touching the data, agree exactly what is being asked: which "
                 "measure, for whom, over what period and compared with what.",
          grid=dict(label="Vague question → sharp question", widths=[1.5, 3.2], rows=[
              ["Vague", "Sharp (specific, measurable, time-bound)"],
              ["How are sales?", "Which stores missed their FY2025 sales target, and by how much?"],
              ["Are students doing well?", "What share of students in each faculty is at risk this semester?"],
              ["Is the hospital busy?", "On how many days in 2023 was bed occupancy 95% or higher?"],
              ["Do customers cancel?", "What is the cancellation rate for each booking channel?"]]),
          points=["What decision will this answer support?", "Which measure, and in what unit?",
                  "Which period, and compared with what (target, last year)?"],
          points_label="QUESTIONS TO ASK FIRST", order=["grid", "points"], split=0.62,
          watch="Do not start building until the stakeholder agrees the question. A perfect answer to the wrong "
                "question is wasted work."),

        T("1.1.12", "Understanding Business Problems and Translating Them into Analytical Questions",
          simple="Managers describe problems in business words. The analyst rewrites each one as a question that data "
                 "can answer, names the data needed and the measure to calculate.",
          diagram=("flow", dict(steps=["Business problem", "Analytical question", "Data needed", "Measure"],
                                sub=["'Profit is falling though sales grow'", "'Which category's margin fell from 2024 to 2025?'",
                                     "tSales: Category, Year, Sales, Profit", "Margin = Profit / Sales"])),
          ex=dict(t="tCat25", cols=["Category", "Sales", "Profit", "Margin_2024", "Margin"], n=3,
                  note="The answer: Furniture margin fell from 3.5% to 1.4%; Office Supplies from 19.1% to 16.1%."),
          order=["diagram", "ex"]),

        T("1.1.13", "General Knowledge and Domain Knowledge",
          defn="General knowledge is broad understanding of how the world works. Domain knowledge is specialised "
               "understanding of a particular industry: its processes, terms, rules and normal values.",
          simple="Data shows a pattern; knowledge explains whether it is normal, important or a mistake.",
          like="Anyone can see a patient has a temperature of 39°C (general). A nurse knows at once it is a fever that "
               "needs action (domain).",
          grid=dict(label="Same numbers, two kinds of knowledge", widths=[2.1, 2, 2.4], size=12, rows=[
              ["What the data shows", "General knowledge says", "Domain knowledge adds"],
              ["Q4 is 38.2% of SuperMart's FY2025 sales", "Festive season: people buy more", "Order stock and staff for Q4 early"],
              ["Bed occupancy 95% or higher on 63 days in 2023", "The hospital was very busy", "Above about 85% is widely seen as a safety risk"],
              ["Online travel agent bookings: 37.8% cancelled", "Many people change plans", "Overbook carefully; ask for deposits"]])),

        T("1.1.14", "Making Decisions Based on Data Rather Than Assumptions",
          simple="An assumption is a belief nobody has checked. Data-driven decisions test the belief against the "
                 "evidence first. Often the data agrees; sometimes it surprises you.",
          like="'This road is always faster' is an assumption. Checking the map app before you leave is data-driven.",
          grid=dict(label="Assumption vs what our data shows", widths=[2, 3], size=12, rows=[
              ["Assumption", "What the data shows (FY2025)"],
              ["Discounts always help the business", "Lines discounted above 20% lost ₦47.6m"],
              ["A store that beats its target is profitable", "Benin City: 105.3% of target, yet a ₦0.7m loss"],
              ["Every store is profitable", "5 of 12 stores lost money"]]),
          ex=dict(t="tDiscBand", cols=["Discount_Band", "Lines", "Sales", "Profit", "Margin"], n=3, note=" "),
          order=["grid", "ex"], split=0.5),
    ],
    takeaways=["Data is raw recorded facts, even in a neat table; information is data processed to answer a question.",
               "Analytics is the process from question to decision: collect, clean, analyse, interpret, communicate.",
               "Analysts answer 'what happened and why'; scientists predict; business analysts redesign processes.",
               "Start with a sharp question; break problems into parts; test assumptions with evidence.",
               "Domain knowledge tells you whether a number is normal, important or an error."],
)

M12 = dict(
    id="1.2", title="Types of Data Analytics",
    subtitle="The four questions analytics answers: what happened, why, what might happen and what should we do.",
    objectives=["Define descriptive, diagnostic, predictive and prescriptive analytics.",
                "Match a business question to the right type of analytics.",
                "Apply all four types to the retail, university and hospital datasets."],
    topics=[
        T("1.2.1", "Descriptive Analytics: What Happened",
          defn="Descriptive analytics summarises historical data to describe what has happened, using totals, "
               "averages, counts, percentages and trends.",
          simple="It is the report card: sales last year, number of students, admissions per month. Most day-to-day "
                 "analyst work is descriptive.",
          like="Like checking your bank statement at month end: it tells you exactly where the money went.",
          ex=dict(t="tYears", cols=["Year", "Sales", "Orders", "Margin"], n=4,
                  note="Descriptive: sales rose from ₦484.2m (2022) to ₦733.2m (2025)."),
          fx=[["=SUM(tStores25[Sales])", "₦733,215,255", "Total FY2025 sales"],
              ["=AVERAGE(tStores25[Sales])", "₦61,101,271", "Average sales per store"]],
          fx_label="DESCRIPTIVE TOOLS IN EXCEL", fx_widths=[2.7, 1.5, 1.3], order=["ex", "fx"], split=0.44),

        T("1.2.2", "Diagnostic Analytics: Why It Happened",
          defn="Diagnostic analytics examines data to find the causes of an outcome, by drilling down, comparing "
               "segments and looking for relationships.",
          simple="Descriptive says 'margin fell'. Diagnostic asks 'why?' and slices the data by store, category or "
                 "discount until the cause appears.",
          like="A doctor who sees a fever (what happened) runs tests to find out whether it is malaria or typhoid "
               "(why it happened).",
          ex=dict(t="tDiscBand", cols=["Discount_Band", "Lines", "Sales", "Profit", "Margin"], n=3,
                  note="Drill-down: the 455 lines discounted above 20% lost ₦47.6m, a third of the ₦141.0m earned by all other lines."),
          watch="Two things moving together is a clue, not proof. Check other explanations before naming a cause."),

        T("1.2.3", "Predictive Analytics: What Might Happen",
          defn="Predictive analytics uses historical data, patterns and statistical models to estimate what is likely "
               "to happen in the future.",
          simple="If sales rise every Q4, we expect next Q4 to be strong too. Predictions are estimates with a margin of "
                 "error, never certainties.",
          like="A weather forecast: '70% chance of rain' is based on past patterns, and it is sometimes wrong.",
          grid=dict(label="FY2025 sales by quarter: a repeating pattern to predict from", widths=[1, 1.4, 1.2], rows=[
              ["Quarter", "Sales", "Share of year"],
              ["Q1", "₦123.1m", "16.8%"], ["Q2", "₦133.8m", "18.2%"], ["Q3", "₦196.3m", "26.8%"], ["Q4", "₦280.1m", "38.2%"]]),
          points=["Trend: the general direction over time", "Seasonality: a pattern that repeats each year",
                  "Forecast error: in our Week 4 backtest the best method was still 16.4% off on average"],
          points_label="KEY IDEAS (WEEK 4 GOES DEEPER)", order=["grid", "points"], split=0.48),

        T("1.2.4", "Prescriptive Analytics: What Should Be Done",
          defn="Prescriptive analytics recommends the best course of action by comparing options and their likely "
               "outcomes, often with scenarios or optimisation.",
          simple="It goes beyond 'what will happen' to 'what should we do about it', and shows the expected result of "
                 "each choice.",
          like="A map app does not just predict traffic; it recommends the fastest route and tells you the arrival time.",
          grid=dict(label="Options for a Furniture discount cap (FY2025 model)", widths=[1.4, 1.6, 1.6, 1.6], rows=[
              ["Furniture cap", "No volume lost", "25% volume lost", "50% volume lost"],
              ["10%", "+₦35.57m", "+₦30.60m", "+₦25.63m"],
              ["20%", "+₦17.96m", "+₦17.85m", "+₦17.73m"],
              ["30%", "+₦8.15m", "+₦9.52m", "+₦10.89m"]],
              note="Profit gain if Furniture lines above the cap are re-priced to the cap; lost volume is an assumption."),
          watch="A recommendation is only as good as its assumptions. Always state them, as the note above does."),

        T("1.2.5", "Practical Examples of Each Type",
          simple="The four types build on each other. Each step answers a harder question and adds more value.",
          diagram=("ladder", dict(steps=[("Descriptive", "What happened?"), ("Diagnostic", "Why did it happen?"),
                                         ("Predictive", "What might happen?"), ("Prescriptive", "What should we do?")])),
          diagram_label="Rising value and difficulty",
          grid=dict(label="One question of each type", widths=[1.3, 2.4, 2.6], size=12, rows=[
              ["Type", "Everyday example", "Business example"],
              ["Descriptive", "How much did I spend on transport last month?", "What were sales per store in FY2025?"],
              ["Diagnostic", "Why was my phone bill high?", "Why did Furniture margin fall?"],
              ["Predictive", "Will my data plan last till month end?", "How much will we sell in Q4 2026?"],
              ["Prescriptive", "Which data plan should I buy?", "What discount cap should we set?"]]),
          order=["diagram", "grid"]),

        T("1.2.6", "Applying the Four Types to Retail, University and Hospital Datasets",
          simple="The same four questions work in any sector. These are real results from our three main datasets.",
          grid=dict(label="Four types × three datasets", widths=[1.2, 2.2, 2.2, 2.2], size=11.5, rows=[
              ["", "SuperMart Retail", "Kingsway University", "Mercyland Hospital (2023)"],
              ["Descriptive", "FY2025 sales ₦733.2m; margin 12.7%", "5,840 students; 1,591 at risk", "12,626 admissions; readmission 8.7%"],
              ["Diagnostic", "Discounts above 20% lost ₦47.6m", "At-risk share highest in Engineering (36.8%)", "Cardiology has the highest readmission rate (14.1%)"],
              ["Predictive", "Q4 brings 38% of the year's sales", "Low attendance flags likely failure early", "Occupancy peaks mid-year (June 98.1%)"],
              ["Prescriptive", "Cap Furniture discounts at 20%: about +₦17.9m", "Offer tutoring to flagged students first", "Plan extra staffed beds before peak months"]])),
    ],
    takeaways=["Descriptive = what happened; diagnostic = why; predictive = what might happen; prescriptive = what to do.",
               "Each type builds on the one before it and adds more value.",
               "Predictions and recommendations rest on assumptions: always state them.",
               "The same four questions apply to retail, education, healthcare and every other sector."],
)

M13 = dict(
    id="1.3", title="The Data Analytics Process",
    subtitle="Eight stages from a business question to a recommendation, and how to keep the work accurate.",
    objectives=["List the stages of the analytics process in order.",
                "Write SMART objectives and identify data sources and collection methods.",
                "Explain cleaning, transformation, analysis, interpretation and visualisation.",
                "Communicate findings with a clear recommendation.",
                "Check data quality and avoid the most common analysis mistakes."],
    topics=[
        T("1.3.1", "Understanding the Problem and Defining Objectives",
          defn="An analytical objective is a clear statement of what the analysis must find out, for whom and by "
               "when. A good objective is SMART: Specific, Measurable, Achievable, Relevant and Time-bound.",
          simple="Every project starts here. If the objective is fuzzy, every later stage goes in the wrong direction.",
          diagram=("flow", dict(steps=["Define", "Collect", "Clean", "Transform", "Analyse", "Interpret", "Visualise",
                                       "Communicate"], highlight=0)),
          grid=dict(label="Fuzzy vs SMART objective", widths=[1.3, 3.4], rows=[
              ["", "Objective"],
              ["Fuzzy", "Look at the store data."],
              ["SMART", "By Friday, identify which of the 12 stores missed their FY2025 sales target and by how much, "
                        "for the Head of Sales."]]),
          order=["diagram", "grid"]),

        T("1.3.2", "Data Collection and Sources",
          defn="Data collection is gathering the data needed to meet the objective. Primary data is collected by you "
               "for this purpose; secondary data was collected by someone else for another purpose.",
          simple="First ask: does the data already exist? Most analysis uses internal data that systems record "
                 "automatically.",
          grid=dict(label="Where our datasets come from", widths=[1.8, 2.4, 1.3], size=12, rows=[
              ["Source", "Example in this bootcamp", "Type"],
              ["Till / point-of-sale system", "SuperMart sales lines (tSales)", "Internal"],
              ["Student records system", "Kingsway student records (tStudents)", "Internal"],
              ["Hospital admissions system", "Mercyland admissions (tAdmissions23)", "Internal"],
              ["Online booking engine", "Atlantic Pearl bookings (tBookings)", "Internal"],
              ["Survey", "ETEH Talent Salary Survey (tSalary)", "Primary"],
              ["Published statistics", "Inflation or population figures", "External, secondary"]])),

        T("1.3.3", "Manual Data Entry and Data Import",
          simple="Small data can be typed in. Anything larger, or anything that already exists in a system, should be "
                 "imported: it is faster and avoids typing errors.",
          grid=dict(label="Manual entry vs import", widths=[1.2, 2.2, 2.2], rows=[
              ["", "Manual entry", "Import"],
              ["Use for", "A few values: targets, a lookup list", "Files and system exports"],
              ["Speed", "Slow", "Fast, and repeatable"],
              ["Main risk", "Typing errors", "Wrong types: dates, leading zeros"]]),
          steps=["Data tab › Get Data › From File › From Text/CSV", "Choose the file, check the preview and the column types",
                 "Click Load (or Transform Data to clean first in Power Query)"],
          steps_label="IMPORT A CSV", order=["grid", "steps"], split=0.5,
          watch="Double-clicking a CSV lets Excel guess the types: codes like 007 lose their zeros and dates can flip "
                "day and month."),

        T("1.3.4", "Connecting to Files, Folders and Databases",
          simple="Instead of copying data in, Excel can connect to it. A connection remembers where the data lives, "
                 "so one click on Refresh brings in the latest version.",
          grid=dict(label="Data › Get Data: where Excel can connect", widths=[1.2, 2.4, 2.6], size=12, rows=[
              ["Source", "Route", "Example"],
              ["Workbook", "From File › From Workbook", "A store's monthly sales workbook"],
              ["Text / CSV", "From File › From Text/CSV", "SuperMart sales export (.csv)"],
              ["Folder", "From File › From Folder", "12 monthly files combined into one table"],
              ["Database", "From Database (or ODBC)", "supermart_online.db, used in Week 3"],
              ["Web", "From Other Sources › From Web", "A table on a statistics web page"]]),
          watch="Power Query, which builds these connections, is covered fully in Week 2. Here you only need to know the routes."),

        T("1.3.5", "Data Cleaning and Preparation",
          defn="Data cleaning is detecting and correcting (or removing) errors and inconsistencies in data so that it "
               "is accurate, complete and consistent.",
          simple="Real exports are messy. Analysts often spend more time cleaning than analysing, because wrong inputs "
                 "give wrong answers.",
          like="Like picking stones out of beans before cooking: skip it and someone breaks a tooth.",
          grid=dict(label="Problems found in the dirty SuperMart file (tDirtySales, 3,387 rows)", widths=[1.6, 1.7, 2.3],
                    size=12, rows=[
                        ["Problem", "Example", "Fix"],
                        ["Inconsistent case", "LAGOS - IKEJA (Row 7486)", "PROPER or a lookup to the store list"],
                        ["Number stored as text", "₦1,399,930 in Sales (234 rows)", "Remove ₦ and commas; VALUE"],
                        ["Mixed date formats", "14/05/2025, 04-May-2025 (335 rows)", "Convert to real dates"],
                        ["Extra spaces", "' North-Central' in Zone", "TRIM"],
                        ["Duplicates and blank rows", "60 repeated records (e.g. Row 6748); 15 empty rows", "Remove Duplicates; delete blanks"]])),

        T("1.3.6", "Data Transformation and Organisation",
          defn="Data transformation is changing the structure or format of clean data so it is ready for analysis: "
               "deriving new columns, grouping, reshaping and combining tables.",
          simple="Cleaning fixes what is wrong; transformation adds what is missing. We create the columns the "
                 "question needs.",
          grid=dict(label="Transformations used in this course", widths=[1.8, 1.8, 2.4], size=12, rows=[
              ["Start with", "Create", "How"],
              ["Order_Date", "Year, Month, Quarter", "=YEAR([@Order_Date]), =MONTH(...)"],
              ["Sales, Profit", "Margin", "=[@Profit]/[@Sales]"],
              ["Discount", "Discount_Band", "IF: none, 1-20%, above 20%"],
              ["Lines per store", "One row per store", "Group and summarise (SUMIFS or a PivotTable)"],
              ["Sales + store list", "Sales with Zone and Target", "Lookup on Store_ID (XLOOKUP)"]])),

        T("1.3.7", "Data Analysis",
          defn="Data analysis is applying calculations and techniques (aggregating, comparing, segmenting, finding "
               "trends and relationships) to answer the question.",
          simple="Now the numbers are worked out: totals, rates, rankings and comparisons with a target or last year.",
          points=["Aggregate: total, average, count", "Compare: against target or last year",
                  "Rank: best and worst performers", "Segment: by store, category or channel", "Trend: change over time"],
          points_label="CORE TECHNIQUES",
          ex=dict(t="tStores25", cols=["Store_Name", "Sales", "Target", "Attainment"], n=5, sort="Attainment", asc=True,
                  note="Comparing with target: the 5 stores furthest below their FY2025 target."),
          order=["points", "ex"], split=0.38),

        T("1.3.8", "Interpretation of Results",
          defn="Interpretation is explaining what the results mean in the real-world context: whether they are good or "
               "bad, why they matter and what they suggest.",
          simple="A number is not a finding. '37.8%' becomes a finding when you say 'more than one in three online "
                 "travel agent bookings is cancelled, three times the corporate rate'.",
          ex=dict(t="tHotelSeg", cols=["Sales_Channel", "Bookings", "Cancelled", "Cancel_Rate"], n=5),
          points=["Compare with something: a target, last year, another group", "Say what it means for the business",
                  "Note limits: small groups (Airline Crew has only 125 bookings) are less reliable"],
          points_label="HOW TO INTERPRET", order=["ex", "points"], split=0.55),

        T("1.3.9", "Data Visualisation and Presentation",
          defn="Data visualisation is the graphical representation of data, using charts, tables and dashboards, so "
               "that patterns and comparisons can be seen quickly.",
          simple="Pick the chart that matches the question. The wrong chart hides the message.",
          grid=dict(label="Choose the chart from the question", widths=[2.2, 1.4, 2.4], size=12, rows=[
              ["If the question is…", "Use", "Example"],
              ["Compare categories", "Bar / column", "Sales by store"],
              ["Change over time", "Line", "Monthly occupancy in 2023"],
              ["Part of a whole (few parts)", "Pie or 100% bar", "Payment method share"],
              ["Relationship between two measures", "Scatter", "Attendance vs exam score"],
              ["Spread of values", "Histogram", "Length of stay"]]),
          watch="Avoid 3D charts, too many colours and pies with more than five slices. Charts are covered in Week 2."),

        T("1.3.10", "Communicating Findings and Recommendations",
          simple="Managers are busy. Lead with the answer, back it with two or three numbers, then say what to do next.",
          points=["**Headline:** Five of our 12 stores lost money in FY2025.",
                  "**Evidence:** Lines discounted above 20% lost ₦47.6m; Furniture margin was only 1.4%.",
                  "**Recommendation:** Cap Furniture discounts at 20% (about +₦17.9m a year in our model).",
                  "**Next step:** Pilot the cap in two stores for one quarter and track the margin."],
          points_label="A FINDING IN FOUR PARTS",
          grid=dict(label="Do and avoid", widths=[1, 1], rows=[
              ["Do", "Avoid"], ["Answer first", "Story of every step"], ["Round: ₦17.9m", "₦17,957,612"],
              ["One message per chart", "Jargon"], ["State assumptions", "Hiding bad news"]]),
          order=["points", "grid"], split=0.62),

        T("1.3.11", "Accuracy and Data Quality",
          defn="Data quality is how fit data is for its purpose, usually judged on accuracy, completeness, consistency, "
               "validity, timeliness and uniqueness.",
          simple="Before trusting any result, check the data. A beautiful chart built on bad data is still wrong.",
          grid=dict(label="Six quality dimensions, with checks", widths=[1.3, 2.4, 2.4], size=12, rows=[
              ["Dimension", "Question", "Example check"],
              ["Accuracy", "Are values correct?", "Workbook total = source total (₦733,215,255)"],
              ["Completeness", "Is anything missing?", "Count blanks: 15 empty rows in tDirtySales"],
              ["Consistency", "Same thing, same way?", "One spelling per store name"],
              ["Validity", "Allowed type and range?", "Discount between 0% and 80%"],
              ["Timeliness", "Up to date?", "Data runs to the end of FY2025"],
              ["Uniqueness", "Any duplicates?", "60 duplicate records in tDirtySales"]])),

        T("1.3.12", "Common Mistakes in Data Analysis",
          simple="Most wrong answers come from a handful of avoidable mistakes. Learn them now and check for them every time.",
          grid=dict(label="Mistake → how to avoid it", widths=[2.1, 2.9], size=12, rows=[
              ["Mistake", "Avoid it by"],
              ["Analysing before agreeing the question", "Write a SMART objective first"],
              ["Skipping cleaning", "Profile every column: blanks, types, duplicates"],
              ["Wrong join that double-counts", "Check totals: a wrong join in Week 3 shows ₦7.19bn instead of ₦1.32bn"],
              ["Averaging percentages", "Recalculate from totals: total profit / total sales"],
              ["Calling a correlation a cause", "Look for other explanations"],
              ["Overwriting the raw data", "Keep raw data on its own sheet; work on a copy"]])),
    ],
    takeaways=["Define → Collect → Clean → Transform → Analyse → Interpret → Visualise → Communicate.",
               "A SMART objective keeps every later stage on track.",
               "Import rather than type; connect when the data will be refreshed.",
               "Clean and check quality before analysing; reconcile totals to the source.",
               "Lead with the answer, support it with evidence and end with a recommendation."],
)

M14 = dict(
    id="1.4", title="Applications of Data Analytics Across Industries",
    subtitle="How retail, education, healthcare, finance, government, marketing, HR, logistics and mining use data.",
    objectives=["Describe typical analytics questions in nine industries.",
                "Recognise the key measures each industry tracks.",
                "Read a real result from our datasets for retail, education, healthcare, marketing, HR and logistics."],
    topics=[
        T("1.4.1", "Retail and E-commerce",
          simple="Retailers track what sells, where, when and at what profit. Online shops add clicks, baskets, "
                 "payments and deliveries.",
          points=["Sales and margin by store, category and product", "Target attainment", "Discount effect on profit",
                  "Average order value and basket size", "Payment method and delivery success"],
          points_label="KEY MEASURES",
          ex=dict(t="tPayMix", cols=["Payment_Method", "Orders", "Revenue", "Share"], n=6,
                  note="SuperMart Online: Paystack and Flutterwave carry over half of all 11,000 orders."),
          order=["points", "ex"], split=0.38),

        T("1.4.2", "Education and Universities",
          simple="Universities use data to track admission, attendance, results and progression, and to support "
                 "students before they fail.",
          points=["Pass and progression rates", "Attendance and study hours", "At-risk students by faculty",
                  "Course and lecturer evaluations"],
          points_label="KEY MEASURES",
          ex=dict(t="tFaculty", cols=["Faculty", "Students", "Avg_Score", "At_Risk", "At_Risk_Rate"], n=4,
                  note="Kingsway University, 2024/2025 first semester: 1,591 of 5,840 students are flagged at risk."),
          order=["points", "ex"], split=0.36),

        T("1.4.3", "Healthcare and Hospitals",
          simple="Hospitals analyse admissions, beds, length of stay, readmissions and billing to improve care and "
                 "use scarce beds well.",
          points=["Bed occupancy", "Average length of stay (ALOS)", "30-day readmission rate",
                  "Billing and HMO claims"],
          points_label="KEY MEASURES",
          ex=dict(t="tDept23", cols=["Department", "Admissions", "Avg_LOS", "Readmission_Rate"], n=5,
                  note="Mercyland Specialist Hospital, 2023: 12,626 admissions; overall 30-day readmission 8.7%."),
          order=["points", "ex"], split=0.34),

        T("1.4.4", "Banking and Finance",
          simple="Banks hold huge volumes of transactions. Analytics protects money and decides who gets credit.",
          like="A bank's fraud system is like a security guard who knows your habits: a ₦900,000 card payment at 3 am "
               "in another country looks wrong at once.",
          grid=dict(label="Typical uses", widths=[1.5, 3], rows=[
              ["Use", "What the data does"],
              ["Fraud detection", "Flags transactions that do not match a customer's normal pattern"],
              ["Credit scoring", "Estimates how likely a borrower is to repay"],
              ["Customer churn", "Finds customers likely to move to another bank"],
              ["Branch and ATM planning", "Uses transaction volumes to place cash and staff"]])),

        T("1.4.5", "Government and Public Administration",
          simple="Governments use data to plan services, set budgets and measure results for millions of people.",
          grid=dict(label="Typical uses", widths=[1.6, 3], rows=[
              ["Area", "Example question"],
              ["Population and planning", "Where are new schools and clinics needed?"],
              ["Economy", "How are prices (inflation) and jobs changing?"],
              ["Revenue", "Which taxpayers are under-declaring?"],
              ["Public health", "Where is a disease outbreak spreading?"],
              ["Transport", "Which roads need repair first?"]]),
          watch="Public data is about real people: protect privacy and publish only what cannot identify individuals."),

        T("1.4.6", "Marketing and Customer Management",
          simple="Marketing analytics finds out which customers and channels bring the most value, and which ones cost "
                 "more than they return.",
          points=["Customer segments", "Conversion and cancellation rates by channel", "Campaign return",
                  "Repeat customers and loyalty"],
          points_label="KEY MEASURES",
          ex=dict(t="tHotelSeg", cols=["Sales_Channel", "Bookings", "Cancel_Rate"], n=5,
                  note="Atlantic Pearl Hotel, 2023-2024: corporate bookings cancel far less often than online agency bookings."),
          order=["points", "ex"], split=0.40),

        T("1.4.7", "Human Resources (People Analytics)",
          defn="People analytics is the use of workforce data to improve decisions about hiring, pay, performance, "
               "retention and wellbeing.",
          simple="HR uses data to pay fairly, hire well and understand why people leave.",
          ex=dict(t="tSeniority", cols=["Seniority", "People", "Median_Salary", "Avg_Experience"], n=4,
                  note="ETEH Talent Salary Survey, 1,781 professionals: pay rises with seniority."),
          points=["Headcount and turnover", "Pay by role, level and gender", "Time to hire", "Absence and engagement"],
          points_label="KEY MEASURES", order=["ex", "points"], split=0.6,
          watch="Salary and personal data is sensitive: report groups, never named individuals."),

        T("1.4.8", "Logistics and Transportation",
          simple="Logistics analytics makes sure goods arrive on time at the lowest cost: routes, couriers, stock and "
                 "delivery success.",
          points=["On-time and delivered rate", "Delivery time (days)", "Cost per delivery", "Courier comparison"],
          points_label="KEY MEASURES",
          ex=dict(t="tCouriers", cols=["Courier", "Deliveries", "Delivered", "Delivered_Rate"], n=8,
                  note="SuperMart Online: delivered rates sit within 78-81% for every courier, so no courier stands out."),
          order=["points", "ex"], split=0.36),

        T("1.4.9", "Geology, Mining and Natural Resources",
          simple="Mining and oil companies collect huge amounts of sensor, drilling and production data. Analytics "
                 "finds resources, controls cost and keeps people safe.",
          grid=dict(label="Typical uses", widths=[1.6, 3], rows=[
              ["Use", "What the data does"],
              ["Exploration", "Combines drill-hole samples and maps to estimate where ore or oil lies"],
              ["Grade control", "Tracks the quality of ore mined each day against the plan"],
              ["Equipment maintenance", "Uses sensor readings to service machines before they break down"],
              ["Safety and environment", "Monitors incidents, gas levels and water quality"],
              ["Production reporting", "Tonnes or barrels produced against target"]]),
          like="Exploration is like looking for a lost ring in a field: every test hole tells you where to dig next."),

        T("1.4.10", "Real-World Data-Driven Decisions",
          simple="In every case below, data replaced a guess with evidence and pointed to a specific action.",
          grid=dict(label="From our datasets: evidence → decision", widths=[1.6, 2.6, 2.4], size=12, rows=[
              ["Organisation", "Evidence", "Decision"],
              ["SuperMart Retail", "Lines discounted above 20% lost ₦47.6m in FY2025", "Cap Furniture discounts at 20%"],
              ["Kingsway University", "1,591 students flagged at risk", "Target tutoring at flagged students"],
              ["Mercyland Hospital", "Occupancy 95% or more on 63 days in 2023", "Plan surge staffing for peak months"],
              ["Atlantic Pearl Hotel", "37.8% of online agency bookings cancelled", "Require deposits on agency bookings"],
              ["SuperMart Online", "Paystack and Flutterwave: 52.5% of orders", "Keep both gateways reliable at checkout"]])),
    ],
    takeaways=["Every industry asks the same kinds of questions about performance, causes, forecasts and actions.",
               "Each industry has its own key measures: margin, pass rate, occupancy, cancellation, delivered rate.",
               "Domain knowledge tells you which measures matter and what 'normal' looks like.",
               "Personal data (patients, staff, customers) must be protected and reported in groups."],
)

MODULES = [M11, M12, M13, M14]
