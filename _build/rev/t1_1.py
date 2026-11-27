from revlib import T

TOPICS = [
    T("1.1.1",
      what="**Data** is a collection of raw facts: numbers, names, dates and codes written down when something happens. "
           "A sale at a till, a patient's temperature, or a student's exam score are all data. On its own, one piece of "
           "data tells you very little. The number 474,430 means nothing until you know it is a sale, in naira, at a "
           "SuperMart store, on 1 January 2025.",
      like="Data is like the loose ingredients in your market basket: rice, tomatoes, pepper and onions. They are useful, "
           "but nobody calls them a meal until they are cleaned, cooked and served as jollof.",
      how=["Open SuperMart_Stores_Sales_2022_2025.csv in Excel.",
           "Look at any single row: each cell is one raw fact (date, store, product, amount).",
           "Ask of every cell: what is it, in what unit, and when was it recorded?"],
      tip="Always find out what each column means and its unit before using it.",
      trap="Treating a lone number as meaningful without knowing what it measures.",
      who="SuperMart Retail",
      case="Row 6,684 of the SuperMart sales file holds: SM-2025-144463, 2025-01-01, Gbenga Anyanwu, Lagos - Ikeja, "
           "Howard Miller wall clock, 474430, 11, 0, 199261. Each value is raw data. Read alone, none of them answers a question.",
      steps=["Open the file and press Ctrl+G, type A6685 and press Enter to jump to that row.",
             "Read across the row and write down what each value is.",
             "Use the header row (row 1) to name each value: Sales, Quantity, Discount, Profit."],
      table=[["Raw value", "Column", "What it really means"],
             ["474430", "Sales", "₦474,430 paid for the clocks"],
             ["11", "Quantity", "11 clocks in this line"],
             ["0", "Discount", "No discount was given"],
             ["199261", "Profit", "₦199,261 profit on this line"]],
      so="A number only becomes useful when you know its label, its unit and its context. The header row is your first "
         "dictionary."),

    T("1.1.2",
      what="**Information** is data that has been organised so that a person can understand it quickly. We get information "
           "by adding labels and units, and by doing simple calculations such as totals, averages and percentages. "
           "'SuperMart sold ₦733.2 million in 2025' is information. It came from adding up 3,312 raw sales lines.",
      like="If data is the raw ingredients, information is the pot of jollof rice: the same ingredients, washed, measured, "
           "cooked and served on a plate so anyone can enjoy it.",
      how=["Start from the raw rows.",
           "Group them by something people care about (store, month, category).",
           "Add a total, average or percentage, and label it clearly with the unit."],
      tip="Every summary should say what was counted, the unit and the period.",
      trap="Giving totals without the period or unit, e.g. '146 million' with no ₦ and no year.",
      who="SuperMart Retail",
      case="Head Office wants to know how big the Lagos - Ikeja store is compared with the whole chain in 2025. The "
           "3,312 raw 2025 sales lines cannot answer that until they are summed and compared.",
      steps=["Filter Order_Date to 2025 and add up the Sales column for all stores.",
             "Filter Store_Name to Lagos - Ikeja and add up Sales again.",
             "Divide the store total by the chain total and format it as a percentage."],
      table=[["Measure (FY2025)", "Value"],
             ["Chain sales, 12 stores", "₦733,215,255"],
             ["Lagos - Ikeja sales", "₦146,388,345"],
             ["Ikeja share of the chain", "20.0%"]],
      so="One clear sentence ('Ikeja brings in one naira of every five') is information. It was hidden inside thousands of "
         "rows of data."),

    T("1.1.3",
      what="Data is the **input**; information is the **output**. Data sits in rows, one row for each sale or patient. "
           "Information is the short summary built from those rows: a table, a total or a chart. Analysts keep both. "
           "They keep the raw rows so they can always check their work, and the summary so managers can act. Sending a "
           "manager 3,000 raw rows is not helpful. Sending a 12-line summary is.",
      like="A market woman keeps every receipt in her bag (data). At month end she tells her husband: 'I sold 40 bags of rice "
           "and made ₦60,000 profit' (information). She needs both: the receipts prove the summary.",
      how=["Keep the raw rows untouched on one sheet.",
           "Build the summary on a separate sheet.",
           "Make sure the summary total equals the raw total."],
      who="SuperMart Retail",
      case="The 2025 sales file has 3,312 rows. The Head of Sales only needs one line per store: sales against the 2025 "
           "target from SuperMart_Stores_Directory.csv.",
      steps=["Sum 2025 Sales for each of the 12 stores.",
             "Bring in each store's Annual_Target_2025 from the directory.",
             "Divide sales by target to get '% of target'."],
      table=[["Store (FY2025)", "Sales", "Target", "% of target"],
             ["Lagos - Lekki", "₦70.8m", "₦63.0m", "112%"],
             ["Lagos - Ikeja", "₦146.4m", "₦142.0m", "103%"],
             ["Abuja - Wuse", "₦43.4m", "₦52.0m", "84%"],
             ["Owerri - New Owerri", "₦31.6m", "₦39.0m", "81%"]],
      so="Twelve lines of information replace 3,312 lines of data. The manager can see at once which stores beat their "
         "target and which stores need help."),

    T("1.1.4",
      what="**Data analytics** means using data to answer a question and decide what to do. It follows a simple path: ask "
           "a clear question, collect the data, clean it, analyse it (count, compare, look for patterns) and explain the "
           "answer. Bookkeeping records what happened. Analytics goes further and asks why it happened and what we "
           "should do next.",
      like="Analytics works like a good doctor. Your symptoms are the data. Tests and questions are the analysis. The "
           "diagnosis is the insight, and the prescription is the decision.",
      how=["Write the question in one sentence.",
           "Find the data that can answer it.",
           "Clean, calculate and compare.",
           "Explain the answer and recommend an action."],
      who="SuperMart Retail",
      case="SuperMart's sales grew 20.4% in 2025, yet its profit margin fell from 13.4% to 12.7%. The board asks: why are "
           "we selling more but keeping less of each naira?",
      steps=["Question: which kind of sale loses money?",
             "Group 2025 sales lines by discount size (0%, up to 20%, above 20%).",
             "Add up profit for each group and compare."],
      table=[["Discount on the line", "Lines", "Profit FY2025"],
             ["No discount", "1,590", "₦104.2m"],
             ["1–20%", "1,267", "₦36.9m"],
             ["Above 20%", "455", "−₦47.6m"]],
      so="Deep discounts above 20% lost ₦47.6 million. That is the 'why'. The action that follows from it is to cap "
         "discounts. This is analytics: from question to decision."),

    T("1.1.5",
      what="Organisations analyse data for five main reasons: to **see** how they are performing, to **cut waste**, to "
           "**grow** sales or results, to **spot risks** early, and to **prove** results to bosses, owners or regulators. "
           "Without data, people argue from opinion. With data, they can agree on the facts first.",
      like="It is like checking your bank app before spending. You see where your money went, you notice the small daily "
           "spends that add up, and you can prove to yourself that the new budget is working.",
      how=["Pick one of the five reasons (see, cut, grow, spot risk, prove).",
           "Choose the number that shows it (profit, cost, readmissions…).",
           "Compare it across groups or over time."],
      who="SuperMart Retail",
      case="The CEO wants to 'see' which stores actually make money. Total sales look healthy everywhere, so she asks for "
           "profit by store for 2025.",
      steps=["Sum Sales and Profit for each store for 2025.",
             "Work out the margin: Profit ÷ Sales.",
             "Sort the stores by profit, lowest first."],
      table=[["Store (FY2025)", "Sales", "Profit", "Margin"],
             ["Abuja - Wuse", "₦43.4m", "−₦8.8m", "−20.4%"],
             ["Onitsha - Main Market", "₦47.8m", "−₦6.2m", "−13.0%"],
             ["Ibadan - Ring Road", "₦32.9m", "−₦4.1m", "−12.5%"],
             ["Lagos - Ikeja", "₦146.4m", "₦29.4m", "20.1%"]],
      so="Five of the 12 stores lost money in 2025, even though they all had good sales. Seeing this is the first step "
         "towards fixing it."),

    T("1.1.6",
      what="You already do analytics in daily life. You compare data-bundle prices, check which route has less traffic, or "
           "notice that you spend more on transport at the start of the month. Each time, you collect facts, compare them, "
           "and change what you do. In this course you will do the same thing, only with bigger files and Excel.",
      like="Choosing when to leave home to beat Third Mainland Bridge traffic is analytics. You remember past journey "
           "times, compare them, and pick the best time.",
      how=["Notice a question in your daily life.",
           "Write down a few facts over several days.",
           "Compare them and act on what you see."],
      who="Crestview University",
      case="Crestview's Digital Wellbeing survey asked 705 students how many hours a day they spend on social media and how "
           "long they sleep. Is there a pattern?",
      steps=["Open Crestview_Digital_Wellbeing_Survey.csv.",
             "Group students by daily social media hours.",
             "Find the average sleep hours for each group."],
      table=[["Social media per day", "Students", "Average sleep"],
             ["Under 3 hours", "49", "8.2 hours"],
             ["3–5 hours", "362", "7.4 hours"],
             ["5–7 hours", "259", "6.1 hours"],
             ["7 hours or more", "35", "5.0 hours"]],
      so="The more time students spend on social media, the less they sleep: about three hours less a night at the top "
         "end. The same facts could help you plan your own day."),

    T("1.1.7",
      what="In a business, every department asks its own questions. Finance asks about profit, Operations about delivery "
           "speed, and Sales about targets. When everyone works from the **same trusted data**, called a **single source of "
           "truth**, meetings are spent deciding what to do instead of arguing over whose figure is right.",
      like="In a family meeting about the NEPA bill, everyone looks at the same bill. Nobody argues about the amount, so the "
           "talk moves straight to who pays what.",
      how=["Agree one master file or database for the whole organisation.",
           "Let each team build its own view from that file.",
           "Never retype figures into separate spreadsheets."],
      who="SuperMart Retail",
      case="Three SuperMart departments use the same 2025 sales file, each answering its own question. Because they share "
           "one file, their numbers always agree.",
      steps=["Finance sums Profit and divides by Sales.",
             "Operations finds the average days between Order_Date and Ship_Date.",
             "Sales compares each store's sales with its 2025 target."],
      table=[["Department", "Question", "Answer from the file"],
             ["Finance", "What margin did we make?", "12.7% (₦93.4m profit)"],
             ["Operations", "How fast do we ship?", "3.9 days on average"],
             ["Sales", "How many stores hit target?", "5 of 12 stores"]],
      so="One shared file gives three teams three answers that agree with each other. That is the value of a single source "
         "of truth."),

    T("1.1.8",
      what="A **data analyst** is the link between data and the people who make decisions. Each day an analyst works out "
           "what the question really is, gets the data from files or databases, cleans it, does the calculations, builds "
           "tables and charts, and explains the answer in plain words. They do not just produce numbers. They help people "
           "take the right action.",
      like="An analyst is like an interpreter at a meeting between a Yoruba trader and a Hausa buyer. Each side speaks its "
           "own language (data and business), and the interpreter makes sure both sides understand each other.",
      how=["Clarify the question with the person asking.",
           "Get and clean the data.",
           "Calculate, check and summarise.",
           "Explain the answer and recommend an action."],
      who="SuperMart Retail",
      case="The Head of Sales sends a WhatsApp message: 'Which stores missed their 2025 target?' The analyst uses the sales "
           "file and the store directory to answer within the hour.",
      steps=["Sum 2025 sales for each store and look up each store's target.",
             "Calculate sales ÷ target and flag the stores below 100%.",
             "Reply with the short list and one sentence of advice."],
      table=[["Missed target (FY2025)", "% of target"],
             ["Owerri - New Owerri", "81%"],
             ["Abuja - Wuse", "84%"],
             ["Ibadan - Ring Road", "89%"],
             ["Onitsha - Main Market (+3 more)", "90%"]],
      so="The analyst turned a vague request into a clear list: 7 stores missed target, and Owerri and Abuja - Wuse missed "
         "by the most. That gives the manager a place to start."),

    T("1.1.9",
      what="Three jobs sound alike but do different work. A **data analyst** explains what happened and why, using Excel, "
           "SQL and dashboards. A **business analyst** studies how work is done and designs better processes. A **data "
           "scientist** builds statistical or machine-learning models that predict what will happen next. They often work "
           "on the same problem from different angles.",
      like="Think of building a house. The analyst is the quantity surveyor measuring what is already there. The business "
           "analyst is the architect redesigning the rooms. The data scientist is the engineer predicting how much load "
           "the roof can carry.",
      who="Lifecare Hospital",
      case="In 2023, 10.3% of Lifecare's 10,918 admissions came back within 30 days. Lifecare's management wants fewer "
           "readmissions, and each of the three roles helps in a different way.",
      steps=["The analyst measures the readmission rate by department.",
             "The business analyst maps the discharge process to find weak steps.",
             "The data scientist builds a model that flags high-risk patients."],
      table=[["Role", "Main tool", "Output at Lifecare"],
             ["Data analyst", "Excel, SQL", "Oncology highest: 13.2%"],
             ["Business analyst", "Process maps", "New discharge checklist"],
             ["Data scientist", "Python, models", "Risk score per patient"]],
      so="All three roles help, but this bootcamp trains the first one: the analyst who finds and explains the facts that "
         "everyone else builds on."),

    T("1.1.10",
      what="**Analytical thinking** means breaking a big, vague problem into small parts you can check one at a time. "
           "Instead of saying 'profit is bad', you ask: which category, which store, which month, and which kind of sale? "
           "Each answer narrows the search until you find the real cause.",
      like="It is like a mechanic whose customer says 'the car won't start'. He doesn't change the whole engine. He checks "
           "the battery, then the fuel, then the plugs, one at a time, until he finds the fault.",
      how=["Write down the big problem.",
           "Split it into parts (by category, store, month…).",
           "Find the part that explains most of the problem, then split that part again."],
      who="SuperMart Retail",
      case="'SuperMart's profit is weak' is too vague to act on. The analyst breaks the 2025 profit down by category, then "
           "looks inside the weakest category.",
      steps=["Sum Sales and Profit by Category and work out each margin.",
             "Furniture has by far the lowest margin, so split it by Sub_Category.",
             "Find the sub-category that loses money."],
      table=[["Level", "Item", "Margin FY2025"],
             ["Category", "Technology", "18.7%"],
             ["Category", "Furniture", "1.4%"],
             ["Inside Furniture", "Tables", "−13.4% (−₦8.1m)"],
             ["Inside Furniture", "Chairs", "8.0%"]],
      so="Two simple splits took us from 'profit is weak' to 'Tables are losing ₦8.1 million'. Now the manager has "
         "something specific to fix."),

    T("1.1.11",
      what="Before you open Excel, turn a vague request into a **sharp question**. A good question names five things: the "
           "**measure** (sales, cancellations), the **period**, the **level of detail** (by store, by month), the "
           "**comparison** (against last year or a target) and the **decision** it will support. A sharp question saves "
           "hours of aimless clicking.",
      like="Telling a danfo driver 'take me somewhere in Lagos' gets you nowhere. Saying 'Oshodi, under the bridge, before 8 "
           "a.m.' gets you there.",
      how=["Write the request exactly as you received it.",
           "Add the measure, period, level of detail and comparison.",
           "Confirm the sharper question with the person who asked."],
      who="EkoSuites Hotel",
      case="The EkoSuites manager says: 'How are our bookings?' The analyst sharpens this to: 'What share of 2023–2024 "
           "bookings were cancelled, and which booking channel cancels most?'",
      steps=["Open EkoSuites_Hotel_Bookings_2023_2024.csv (36,275 bookings).",
             "Count bookings by Booking_Status and work out the cancelled share.",
             "Repeat the count for each Market_Segment."],
      table=[["Booking channel", "Cancelled"],
             ["All bookings", "32.8%"],
             ["Online Travel Agent", "36.5%"],
             ["Walk-in / Travel Agent", "29.9%"],
             ["Corporate", "10.9%"]],
      so="The sharp question produced a useful answer: online travel agent bookings cancel most. A vague 'how are bookings?' "
         "would never have found this."),

    T("1.1.12",
      what="Managers describe **problems** in everyday words: 'too many students are failing'. A spreadsheet only "
           "understands **columns, filters and formulas**. Your job is to translate each word of the problem into "
           "something you can calculate: which column, which rule, which group, and which number.",
      like="A tailor does this every day. The customer says 'I want a fine agbada for my brother's wedding'. The tailor "
           "turns that into exact measurements, fabric and a delivery date.",
      how=["Underline each important word in the manager's sentence.",
           "Match each word to a column and a rule.",
           "Write the calculation that answers it."],
      who="Crestview University",
      case="The Senate says: 'Too many students are failing; who needs help?' The analyst turns 'failing' into a clear "
           "rule using Crestview_Student_Performance.csv (6,607 students).",
      steps=["Agree the rule: at risk = Exam_Score below 62 OR Attendance_Pct below 65.",
             "Add an At_Risk column with that rule and count the Yes rows.",
             "Break the count down by attendance band to see where the risk sits."],
      table=[["Business words", "In the spreadsheet", "Result"],
             ["'Failing'", "Exam < 62 or attendance < 65%", "901 students"],
             ["'Too many'", "901 ÷ 6,607", "13.6% at risk"],
             ["'Who?'", "Attendance below 70%", "53.5% of them at risk"]],
      so="After translation, the vague worry becomes a number (901 students) and a clear cause (poor attendance). Now "
         "the Senate can plan support."),

    T("1.1.13",
      what="Excel skill alone is not enough. You also need **domain knowledge**: understanding how the business really works, "
           "including its rules, seasons and limits. Domain knowledge helps you spot numbers that are impossible or "
           "suspicious, and suggest actions that make sense in the real world.",
      like="A pharmacist who sees a prescription for '100 tablets a day' knows at once it is a typing error. That is domain "
           "knowledge. A machine just reads the number.",
      how=["Learn the basic rules of the sector you analyse.",
           "Check each column against those rules (ranges, dates, units).",
           "Ask a domain expert when something looks strange."],
      who="Multi-Sector",
      case="Three files in this bootcamp each contain values that only domain knowledge can judge. The formulas run without "
           "errors, but are the numbers sensible?",
      steps=["Check Lifecare Length_of_Stay: all 2023 stays are between 1 and 30 days.",
             "Check SuperMart Discount: 307 lines in 2025 had more than 40% off.",
             "Check Diabetes screening: HbA1c of 6.5% or more is the clinical diabetes mark."],
      table=[["File", "What the formula sees", "What domain knowledge asks"],
             ["Lifecare", "LOS 1–30 days", "Why does no stay go past 30 days?"],
             ["SuperMart", "Discount 0.8", "Who approved 80% off?"],
             ["Diabetes", "HbA1c 6.6", "Above the 6.5% diagnosis line"]],
      so="Domain knowledge turns 'the formula worked' into 'the answer makes sense'. Always ask whether a number could "
         "really be true."),

    T("1.1.14",
      what="Many decisions are made on **assumptions**: what the boss believes, what happened yesterday, or what 'everyone "
           "knows'. Being **data-driven** means testing the belief against the full data before spending money or "
           "changing policy. Sometimes the data agrees with the belief. Often it does not.",
      like="At the market you don't guess the weight of a bag of rice by looking at it. You put it on the scale. Data is the "
           "scale for business beliefs.",
      how=["Write the belief as a statement that can be checked.",
           "Find the data that can test it.",
           "Count, compare, and accept or reject the belief."],
      tip="Test beliefs with the complete data, not a few memorable cases.",
      trap="Acting on the loudest opinion or yesterday's single complaint.",
      who="SuperMart Online",
      case="A manager says: 'Nigerians prefer cash, so most of our online customers pay on delivery. Let's hire more cash "
           "riders.' The analyst checks payment_method in all 11,000 transactions.",
      steps=["Open SuperMart_Online_transactions.csv.",
             "Count the transactions for each payment_method.",
             "Divide each count by 11,000 to get its share."],
      table=[["Payment method", "Transactions", "Share"],
             ["Paystack", "3,133", "28.5%"],
             ["Flutterwave", "2,646", "24.1%"],
             ["Bank Transfer", "2,182", "19.8%"],
             ["Cash on Delivery", "1,928", "17.5%"]],
      so="Only about 1 order in 6 is paid in cash. The belief was wrong, so the money is better spent on smooth online "
         "payments than on more cash riders."),
]
