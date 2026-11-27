from revlib import T

TOPICS = [
    T("1.3.1",
      what="Every analysis starts by **understanding the problem** and agreeing a clear **objective**. You agree what will "
           "be measured, for which period, at what level of detail, against what comparison, and which decision it will "
           "support. Write this down and agree it with the person who asked. It keeps the work focused and stops arguments "
           "later.",
      like="Before a tailor cuts any fabric, he agrees the style, the measurements and the date of the event. Cutting "
           "first and asking later wastes cloth and money.",
      how=["Meet the person asking and write their request in their words.",
           "Agree the measure, period, level of detail and comparison.",
           "Send back a one-paragraph summary and get a 'yes' before starting."],
      tip="Put the agreed objective at the top of your workbook or report.",
      trap="Starting analysis on 'just look at the numbers' with no agreed objective.",
      who="SuperMart Retail",
      case="The CEO asks for a 'Q1 review'. The analyst agrees the objective: compare January–March 2025 with "
           "January–March 2024 for sales, orders and profit, to decide whether the new pricing is working.",
      steps=["Filter Order_Date to January–March for 2024 and for 2025.",
             "Sum Sales and Profit, and count distinct Order_ID, for each year.",
             "Work out the change: (2025 − 2024) ÷ 2024."],
      table=[["Measure", "Q1 2024", "Q1 2025", "Change"],
             ["Sales", "₦93.2m", "₦123.1m", "+32.1%"],
             ["Orders", "179", "240", "+34.1%"],
             ["Profit", "₦11.4m", "₦23.5m", "+105.4%"]],
      so="Because the objective was clear, the answer is clear: Q1 2025 profit more than doubled on the same quarter "
         "last year."),

    T("1.3.2",
      what="**Data collection** means getting the data you need from the right **source**. Common sources are files "
           "exported from systems (CSV or Excel), databases, surveys and forms, and published reports (PDF or web pages). "
           "Each source has strengths and typical problems, so always note where your data came from and when.",
      like="Cooking egusi soup means buying ingredients from different places: the market, the supermarket, the farm. Each "
           "has different quality, so you check what you buy.",
      how=["List the questions you must answer.",
           "Find the source that holds each piece of data.",
           "Record the source, date and owner of every file you use."],
      tip="Keep a short 'data sources' note in every workbook.",
      trap="Using a file without knowing where it came from or how old it is.",
      who="Bootcamp Files",
      case="The bootcamp pack itself shows all the main source types. Look at them in 01_Section_Examples and note what "
           "to watch out for in each.",
      steps=["Open the 00_Core_Datasets folder and the Week2_Data_Files folder.",
             "Match each file to a source type (export, database, survey, report).",
             "Write one watch-out for each."],
      table=[["Source type", "Pack example", "Watch out for"],
             ["System export", "SuperMart sales CSV", "Text numbers, codes"],
             ["Database", "supermart_online.db", "Must join tables"],
             ["Survey", "Digital Wellbeing (705)", "Self-reported answers"],
             ["Report / PDF", "Monthly Sales Bulletin", "Hard to copy cleanly"]],
      split=4.90,
      so="Knowing the source tells you what problems to expect before you start, so you can check for them."),

    T("1.3.3",
      what="There are two ways to get data into Excel. **Manual entry** means typing it, which is fine for small, stable "
           "lists like ward names. **Importing** means letting Excel read a file or database, which is the only safe way "
           "for large data. Typing thousands of rows is slow and introduces errors.",
      like="Writing down your 5 family phone numbers by hand is fine. Copying 2,000 church members' numbers by hand is "
           "madness. You would ask for the list as a file.",
      how=["Count the rows you need.",
           "Type only small lookup lists (roughly under 50 rows).",
           "Import everything else with Data › Get Data or File › Open."],
      tip="Type small reference lists once; import everything that changes or is large.",
      trap="Retyping data that already exists in a file; every keystroke is a chance for error.",
      who="Lifecare Hospital",
      case="Lifecare's analyst needs two files: the ward list (Lifecare_Wards.csv) and every admission from 2019 to 2024 "
           "(Lifecare_Admissions_2019_2024.csv). Which method suits each?",
      steps=["Open each file and press Ctrl+Down to see how many rows it has.",
             "Decide: small and stable → type; large or changing → import.",
             "Import the large file with Data › Get Data › From Text/CSV."],
      table=[["File", "Rows", "Best method"],
             ["Lifecare_Wards", "4", "Type it (or import)"],
             ["SuperMart_Stores_Directory", "12", "Type it (or import)"],
             ["SuperMart_Stores_Sales", "9,994", "Import"],
             ["Lifecare_Admissions", "54,860", "Import"]],
      so="Four ward names can be typed in a minute. 54,860 admissions must be imported; typing them would take weeks and "
         "be full of mistakes."),

    T("1.3.4",
      what="Excel can **connect** directly to a file, a whole folder of files, or a database using **Data › Get Data**. "
           "A connection remembers where the data lives. When the source file is updated, you click **Refresh** and the "
           "new rows flow in without copying and pasting again.",
      like="A connection is like a water pipe from the tank to your kitchen. Once it is laid, you just open the tap. You "
           "don't carry buckets from the tank every day.",
      how=["Go to Data › Get Data and choose From File, From Folder or From Database.",
           "Pick the file, check the preview, then click Load.",
           "Next time, click Data › Refresh All to bring in new rows."],
      tip="Keep source files in a fixed folder so connections don't break.",
      trap="Moving or renaming the source file; the connection can no longer find it.",
      who="Lifecare Hospital",
      case="Each month Lifecare's records team sends an updated admissions extract. The analyst wants to stop copying and "
           "pasting it into the report.",
      steps=["Data › Get Data › From File › From Text/CSV.",
             "Choose Lifecare_Admissions_2019_2024.csv and click Load.",
             "When a new extract replaces the file, click Data › Refresh All."],
      table=[["Connect to", "Excel menu", "Pack example"],
             ["One file", "From Text/CSV", "Lifecare admissions (54,860 rows)"],
             ["A folder", "From Folder", "Monthly sales files"],
             ["A database", "From Database", "lifecare_hospital.db"]],
      so="A connection turns a monthly hour of copy-and-paste into one click on Refresh, and removes paste mistakes."),

    T("1.3.5",
      what="**Data cleaning** means fixing problems in the raw data before you analyse it. Typical problems are blank rows, "
           "duplicate rows, extra spaces, different spellings of the same thing, numbers stored as text, and dates in "
           "mixed formats. Dirty data gives wrong totals, so analysts often spend more time cleaning than analysing.",
      like="It is like picking stones out of beans before cooking. Skip it, and someone breaks a tooth. In analytics, "
           "someone takes a wrong decision.",
      how=["Scan each column for blanks, odd spellings and odd formats.",
           "Remove blank and duplicate rows; trim spaces; fix spellings.",
           "Convert text numbers and dates; check the row count again."],
      tip="Clean a copy and keep the original raw file untouched.",
      trap="Summing a column before cleaning it; text numbers and duplicates distort the total.",
      who="SuperMart Retail",
      case="SuperMart_Sales_2025_DIRTY.csv should hold the 3,312 sales lines of 2025, but it has 3,387 rows. The analyst "
           "lists every problem before fixing anything.",
      steps=["Count blank rows and exact duplicate rows.",
             "Filter Store_Name and Order_Date to spot mixed spellings and formats.",
             "Look for Sales values with ₦ signs or commas (stored as text)."],
      table=[["Problem found", "Rows affected"],
             ["Completely blank rows", "15"],
             ["Exact duplicate rows", "60"],
             ["Store names in wrong case", "169"],
             ["Dates in mixed formats", "335"],
             ["Sales typed as text (₦239,240)", "234"]],
      so="Removing 15 blank and 60 duplicate rows brings 3,387 back to 3,312 real lines. The other fixes make every value "
         "usable."),

    T("1.3.6",
      what="**Transformation** means reshaping clean data so it is ready for analysis. Examples are splitting one column "
           "into several, joining columns, creating new columns (such as month or age group), and putting data into one "
           "tidy table. You do not change the facts, only the shape.",
      like="Transforming is like cutting a whole yam into slices so it fits in the pot. It is still the same yam, just "
           "ready to cook.",
      how=["Decide which new columns the analysis needs.",
           "Split or combine columns (Text to Columns, Flash Fill, formulas).",
           "Add helper columns such as Year or Month from a date."],
      tip="Add new columns next to the original; never overwrite the raw column.",
      trap="Typing the split values by hand instead of using a tool or a formula.",
      who="SuperMart Retail",
      case="Each SuperMart Order_ID looks like SM-2025-144463: a prefix, the year and a serial number. The analyst needs "
           "the year as its own column.",
      steps=["Select the Order_ID column and choose Data › Text to Columns.",
             "Choose Delimited and tick Other, typing a hyphen (-).",
             "Send the three parts into three new columns: Prefix, Year, Serial."],
      table=[["Order_ID", "Prefix", "Year", "Serial"],
             ["SM-2025-144463", "SM", "2025", "144463"],
             ["SM-2025-119438", "SM", "2025", "119438"],
             ["SM-2025-121160", "SM", "2025", "121160"]],
      so="One messy code becomes three clean columns. Now you can filter or count by year without reading every ID."),

    T("1.3.7",
      what="**Data analysis** is the step where you calculate and compare to answer the question. You use totals, "
           "averages, percentages, rankings and comparisons between groups. Good analysis always compares: against "
           "another group, another period or a target. A number on its own rarely tells a story.",
      like="Weighing yourself once tells you very little. Comparing it with last month's weight tells you whether "
           "your diet is working.",
      how=["Choose the groups to compare (category, store, month).",
           "Calculate the same measures for each group.",
           "Rank them and look at the highest, the lowest and the gaps."],
      tip="Calculate both size (sales) and quality (margin) for every group.",
      trap="Ranking by sales alone; the biggest group may be the least profitable.",
      who="SuperMart Retail",
      case="The merchandising manager wants to know which product category deserves more shelf space in 2026. The analyst "
           "compares the three categories for 2025.",
      steps=["Group the 2025 lines by Category.",
             "Count lines, and sum Sales and Profit for each category.",
             "Work out margin = Profit ÷ Sales and compare."],
      table=[["Category (FY2025)", "Lines", "Sales", "Margin"],
             ["Technology", "624", "₦271.7m", "18.7%"],
             ["Office Supplies", "2,002", "₦246.1m", "16.1%"],
             ["Furniture", "686", "₦215.4m", "1.4%"]],
      so="Furniture sells almost as much as the others but keeps only ₦1.40 of every ₦100. More space should go to "
         "Technology."),

    T("1.3.8",
      what="**Interpretation** means explaining what a result really means in the real world. The same number can be read "
           "wrongly if you ignore the context: the size of the group, the target, the cost or the season. A good "
           "interpretation always asks 'compared with what?' and 'what else could explain this?'",
      like="If a friend says 'I made ₦2 million this month', you still don't know if that is good. It depends on what "
           "they spent and what they expected.",
      how=["Write the raw result in one sentence.",
           "Add the context: target, cost, last year, group size.",
           "Rewrite the sentence so a manager cannot misread it."],
      tip="Put every key number next to its comparison (target, last year, average).",
      trap="Calling the biggest number 'the best' without checking costs and targets.",
      who="SuperMart Retail",
      case="A manager reads the 2025 store table and announces: 'Ikeja is our best store and Abuja - Wuse is fine.' The "
           "analyst checks each claim against targets and profit.",
      steps=["Compare each store's sales with its Annual_Target_2025.",
             "Work out each store's margin (Profit ÷ Sales).",
             "Rewrite each claim with the full context."],
      table=[["Result", "Naive reading", "Sound reading"],
             ["Ikeja ₦146.4m sales", "Best store", "Biggest; Lekki beat target most (112%)"],
             ["Owerri 31.0% margin", "Top performer", "Best margin, but only 81% of target"],
             ["Abuja - Wuse ₦43.4m", "Doing fine", "Lost ₦8.8m (−20.4% margin)"]],
      split=4.70,
      so="The same numbers tell a different story with context. 'Biggest' is not the same as 'best', and sales are not "
         "the same as profit."),

    T("1.3.9",
      what="**Visualisation** means showing results as charts so people see the pattern in seconds. A line chart shows "
           "change over time, a bar chart compares groups, and a table gives exact values. Choose the chart that answers "
           "the question, and give it a title that states the message, not just the topic.",
      like="A good chart works like a road sign. Drivers don't read a paragraph at 80 km/h. They glance at the sign and "
           "understand at once.",
      how=["Decide the message first (e.g. 'sales peak in November').",
           "Choose the chart: line for time, bar for comparing groups.",
           "Write the message as the chart title and label the axes."],
      tip="Make the chart title a sentence that states the finding.",
      trap="3-D effects, too many colours, or a title like 'Chart 1'.",
      who="SuperMart Retail",
      case="The operations manager must plan staff and stock for 2026. The analyst charts 2025 sales by month to show when "
           "the busy season is.",
      steps=["Sum 2025 Sales by month (January to December).",
             "Insert › Line Chart using the 12 monthly totals.",
             "Title it: 'Sales peak in November; Q4 brings 38% of the year'."],
      table=[["Month (2025)", "Sales", "Note"],
             ["February", "₦20.3m", "Lowest month"],
             ["September", "₦87.9m", "Busy season starts"],
             ["November", "₦118.4m", "Peak month"],
             ["Oct–Dec total", "₦280.1m", "38.2% of the year"]],
      so="The chart shows at a glance that the last quarter is the busy season, so extra staff and stock are needed from "
         "September."),

    T("1.3.10",
      what="Analysis only helps if decision-makers understand it. **Communicating findings** means giving a short "
           "headline, the key evidence, a clear recommendation and the expected benefit, in plain words. Put the answer "
           "first, then the proof. Busy managers often read only the first two lines.",
      like="When the news says 'Fuel price rises to ₦X from Monday', the headline comes first and the details follow. "
           "Your report should work the same way.",
      how=["Write the headline: the finding in one sentence with a number.",
           "Add two or three pieces of evidence.",
           "State the recommendation and the expected benefit."],
      tip="Lead with the answer and the naira impact; keep details for the appendix.",
      trap="Telling the story of every step you took before reaching the point.",
      who="SuperMart Retail",
      case="The analyst must send the CEO a WhatsApp-length message about the discount finding from the 2025 sales "
           "file.",
      steps=["Headline: state the loss and its cause.",
             "Evidence: number of lines and the comparison.",
             "Action and benefit: the cap and the naira gain."],
      table=[["Part", "What the analyst writes"],
             ["Headline", "Deep discounts lost ₦47.6m in 2025"],
             ["Evidence", "455 lines had over 20% off; all others made profit"],
             ["Action", "Cap discounts at 20% without approval"],
             ["Benefit", "Up to ₦71.5m more profit a year"]],
      split=4.90,
      so="Four short lines give the CEO everything needed to decide. That is better than a 20-slide walk through the "
         "method."),

    T("1.3.11",
      what="**Accuracy** and **data quality** decide whether people can trust your work. Good data is complete (nothing "
           "missing), unique (no duplicates), valid (values make sense), consistent (same spelling and format) and "
           "timely. One wrong total can destroy trust in all your future reports.",
      like="A trader whose scale is faulty loses customers fast. Once they discover it, they won't trust anything he "
           "sells, even the honest items.",
      how=["Check row counts against the source.",
           "Check for duplicates, blanks and impossible values.",
           "Reconcile your total to a known control figure."],
      tip="Reconcile every report total to an independent control figure.",
      trap="Assuming the data is correct because it came from 'the system'.",
      who="SuperMart Retail",
      case="An analyst sums Sales in the dirty 2025 file without cleaning it first. How far is the result from the "
           "truth?",
      steps=["Count exact duplicate rows and add up their Sales.",
             "Find the Sales values stored as text (with ₦ and commas) and add them up.",
             "Compare with the clean file: 3,312 lines, ₦733.2m."],
      table=[["Quality problem", "Effect on the Sales total"],
             ["60 duplicate rows", "Adds ₦13.2m that never happened"],
             ["234 text values (₦…)", "SUM skips ₦57.1m of real sales"],
             ["Clean control total", "₦733.2m from 3,312 lines"]],
      so="Two simple quality problems push the total in opposite directions, so the error is hidden. Always reconcile "
         "to a trusted control total."),

    T("1.3.12",
      what="Beginners make the same few mistakes again and again. They start without a clear question, skip cleaning, "
           "mix up counts and sums, average percentages that should be weighted, ignore missing values, and present "
           "numbers without context. Knowing these traps helps you avoid most errors.",
      like="Like learner drivers who forget the handbrake or the mirror, analysts slip on the same few basics. A short "
           "checklist before you move prevents most crashes.",
      how=["Keep a short checklist of common mistakes.",
           "Run through it before you send any report.",
           "Ask a colleague to check one total independently."],
      tip="Use a 5-point checklist: question, clean, count vs sum, context, reconcile.",
      trap="Averaging store margins instead of dividing total profit by total sales.",
      who="SuperMart Retail",
      case="An analyst reports SuperMart's 2025 margin by taking the simple average of the 12 store margins. The finance "
           "team gets a different answer. Who is right?",
      steps=["Work out each store's margin, then take the simple AVERAGE of the 12.",
             "Now divide total profit (₦93.4m) by total sales (₦733.2m).",
             "Compare the two answers and explain the gap."],
      table=[["Method", "Margin FY2025"],
             ["Simple average of 12 store margins", "8.9%"],
             ["Total profit ÷ total sales (correct)", "12.7%"]],
      so="A simple average treats tiny Owerri the same as giant Ikeja. Weight by size: divide total profit by total "
         "sales."),
]
