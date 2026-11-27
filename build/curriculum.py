# -*- coding: utf-8 -*-
"""
ETEH Bootcamp - Curriculum specification.

Each module (excluding the removed last-two "lab" modules) is described with a
learning objective and one or more WORKED EXAMPLE blocks.  Every block is drawn
from one of the curated, standardised datasets so the examples are simple,
relatable and grounded in real data ("drawn from the ideas of the datasets").

A block spec:
  title   : short title
  intro   : one relatable sentence framing the example
  src     : ("<dataset_key>", dict)  -> slice a real dataset
            or ("custom", <list-of-dicts>) -> hand-built small table
  show    : columns to display (subset of the source)
  derive  : {new_col: "=B{r}*C{r}"}  row-wise Excel formulas (live)
  totals  : True/False  -> add a SUM totals row (live formulas)
  worked  : explanation of the calculation performed
  task    : "Your turn" practice prompt

Datasets are loaded by make_workbooks.py from Datasets/Clean/*.csv
"""

# dataset keys -> file
DATASET_FILES = {
    "ecom_tx":  "ecommerce_transactions_reconciled.csv",
    "ecom_cust":"ecommerce_customers.csv",
    "ecom_sell":"ecommerce_sellers.csv",
    "ecom_prod":"ecommerce_products.csv",
    "ecom_del": "ecommerce_deliveries.csv",
    "health":   "healthcare_records.csv",
    "salary":   "salary_data.csv",
    "supers":   "superstore_sales.csv",
    "stperf":   "student_performance.csv",
    "stdep":    "student_depression.csv",
    "social":   "social_media_addiction.csv",
    "diab":     "diabetes_prediction.csv",
    "hotel":    "hotel_reservations.csv",
}

WEEKS = [
 {
  "week": 1, "file": "Week1_Section_Examples.xlsx",
  "title": "Week 1 - Data Analytics Foundations & Core Excel Mastery",
  "modules": [
   {"id":"1.1","name":"Introduction to Data Analytics",
    "objective":"Distinguish raw data from actionable information; understand the DIKW chain and the analyst's role.",
    "blocks":[
      {"title":"From raw rows to a decision","intro":"A hospital log is just numbers until we give it context.",
       "src":("health",{"cols":["name","age","medical_condition","billing_amount"],"n":8}),
       "show":["name","age","medical_condition","billing_amount"],
       "derive":{"cost_band":"=IF(D{r}>30000,\"High\",\"Standard\")"},
       "totals":True,
       "worked":"Each patient is a raw record. Adding a Cost Band column turns raw billing into a category a manager can act on. Total billing = SUM(D2:D9).",
       "task":"Add a column that labels patients over 60 as 'Senior'. Then total billing for Seniors only."},
      {"title":"Data vs Information","intro":"A single shop sale is data; the daily total is information.",
       "src":("supers",{"cols":["order_id","category","sales","profit"],"n":8}),
       "show":["order_id","category","sales","profit"],
       "derive":{"margin_pct":"=D{r}/C{r}"},
       "totals":True,
       "worked":"Margin % = Profit / Sales. Summing Sales and Profit across rows gives the day's information total.",
       "task":"Which category has the highest total profit in your slice?"}]},

   {"id":"1.2","name":"Types of Data Analytics",
    "objective":"Climb the maturity ladder: descriptive, diagnostic, predictive, prescriptive.",
    "blocks":[
      {"title":"Descriptive: what happened?","intro":"Count how many bookings were cancelled last quarter.",
       "src":("hotel",{"cols":["booking_id","booking_status","avg_price_per_room"],"n":10}),
       "show":["booking_id","booking_status","avg_price_per_room"],
       "derive":{"status_code":"=IF(B{r}=\"Canceled\",1,0)"},
       "totals":True,
       "worked":"Descriptive analytics summarises the past: count cancellations with a status flag (1=cancelled) and SUM it.",
       "task":"Compute the average room price for cancelled vs not-cancelled bookings."},
      {"title":"Predictive hint","intro":"Experience predicts salary - a glimpse of predictive analytics.",
       "src":("salary",{"cols":["age","years_of_experience","salary"],"n":8,"sort_by":"years_of_experience","ascending":True}),
       "show":["age","years_of_experience","salary"],
       "derive":{"salary_k":"=C{r}/1000"},
       "totals":False,
       "worked":"Predictive analytics uses patterns (here, salary tends to rise with experience). We pre-scale salary to thousands for readability.",
       "task":"Do you see salary rising as years of experience increase? Describe the pattern in one sentence."}]},

   {"id":"1.3","name":"The Data Analytics Process",
    "objective":"Follow a disciplined lifecycle: define, acquire, clean, model, communicate.",
    "blocks":[
      {"title":"Acquire & clean a student extract","intro":"Raw survey export before and after tidying.",
       "src":("stperf",{"cols":["hours_studied","attendance","exam_score"],"n":8}),
       "show":["hours_studied","attendance","exam_score"],
       "derive":{"score_band":"=IF(C{r}>=70,\"Pass\",\"Retake\")"},
       "totals":True,
       "worked":"Step 1-3 of the lifecycle: acquire the extract, confirm types, then engineer a Pass/Retake band from Exam Score.",
       "task":"List the lifecycle steps and write one line on what you would do at the 'communicate' stage."}]},

   {"id":"1.4","name":"Applications Across Industries",
    "objective":"See how each sector uses the same Excel toolkit on its own data.",
    "blocks":[
      {"title":"Five industries, five datasets","intro":"One technique, many domains.",
       "src":("custom",[
         {"Industry":"Healthcare","Dataset":"Healthcare Records","Question":"Avg billing by condition?"},
         {"Industry":"Retail","Dataset":"Superstore Sales","Question":"Profit by region?"},
         {"Industry":"Marketplace","Dataset":"E-Commerce Transactions","Question":"Revenue by seller?"},
         {"Industry":"Education","Dataset":"Student Performance","Question":"Factors driving exam score?"},
         {"Industry":"Hospitality","Dataset":"Hotel Reservations","Question":"What drives cancellations?"},
       ]),
       "show":["Industry","Dataset","Question"],"derive":{},"totals":False,
       "worked":"Every industry runs the same process on domain data. Match each dataset to its analytical question.",
       "task":"Pick one industry and write the three KPIs you would track monthly."}]},

   {"id":"1.5","name":"Introduction to Microsoft Excel",
    "objective":"Learn the Excel anatomy: ribbon, grid, data types, file formats.",
    "blocks":[
      {"title":"Know your data types","intro":"A column's type decides what you can calculate.",
       "src":("supers",{"cols":["order_date","ship_mode","sales","quantity"],"n":8}),
       "show":["order_date","ship_mode","sales","quantity"],
       "derive":{},"totals":True,
       "worked":"Order date is a Date, Ship Mode is Text, Sales/Quantity are Numbers. Mixing types breaks formulas - keep each column one type.",
       "task":"Which column above could NEVER be summed, and why?"}]},

   {"id":"1.6","name":"Working with Data in Excel",
    "objective":"Enter, correct, fill and shape data efficiently.",
    "blocks":[
      {"title":"AutoFill a sequence","intro":"Build a clean ID column with AutoFill.",
       "src":("custom",[
         {"transaction_id":"TXN-001","product":"Wristwatch","qty":3},
         {"transaction_id":"TXN-002","product":"Wardrobe","qty":1},
         {"transaction_id":"TXN-003","product":"Sneakers","qty":2},
         {"transaction_id":"TXN-004","product":"Headphones","qty":5},
       ]),
       "show":["transaction_id","product","qty"],
       "derive":{"line_total":"=C{r}*5000"},
       "totals":True,
       "worked":"Type the first two IDs, drag the fill handle, Excel continues the pattern. Line Total = qty x unit price (5,000).",
       "task":"Add a 5th row and use Flash Fill to extend the transaction_id pattern."}]},

   {"id":"1.7","name":"Excel Formulas and Core Functions",
    "objective":"Master the equal sign, operators, references and core functions (SUM/AVERAGE/COUNTIF).",
    "blocks":[
      {"title":"SUM, AVERAGE, COUNTIF","intro":"Summarise a small salary list by gender.",
       "src":("salary",{"cols":["gender","years_of_experience","salary"],"n":10}),
       "show":["gender","years_of_experience","salary"],
       "derive":{"salary_k":"=C{r}/1000"},
       "totals":True,
       "worked":"Use =AVERAGE(C2:C11) for mean salary, =COUNTIF(A2:A11,\"Female\") to count women. Absolute vs relative references matter when you copy down.",
       "task":"Add a column computing years_of_experience + 1 (a tenure projection) for each row."},
      {"title":"IF for bands","intro":"Convert a score into a grade with IF.",
       "src":("stperf",{"cols":["hours_studied","exam_score"],"n":8}),
       "show":["hours_studied","exam_score"],
       "derive":{"grade":"=IF(B{r}>=70,\"A\",IF(B{r}>=60,\"B\",\"C\"))"},
       "totals":False,
       "worked":"Nested IF maps score to grade: >=70 A, >=60 B, else C. IF is the workhorse of business rules.",
       "task":"Change the rule so 80+ is 'A' and 50-79 is 'B'."}]},

   {"id":"1.8","name":"Organising Data with Excel Tables",
    "objective":"Convert ranges into official Tables for structured references and auto-fill.",
    "blocks":[
      {"title":"Table + structured reference","intro":"Turn a student extract into a Table (Ctrl+T).",
       "src":("stperf",{"cols":["hours_studied","attendance","previous_scores","exam_score"],"n":8}),
       "show":["hours_studied","attendance","previous_scores","exam_score"],
       "derive":{"score_x2":"=D{r}*2"},
       "totals":True,
       "worked":"Inside a Table, totals use =SUBTOTAL(109,[exam_score]). Structured refs ([column]) survive column moves.",
       "task":"Insert a Total Row and show the average of Attendance."}]},

   {"id":"1.9","name":"Data Entry Best Practices: The Twelve Tidy Data Laws",
    "objective":"Apply tidy-data rules: atomic cells, one table per observation, consistent labels.",
    "blocks":[
      {"title":"Messy vs Tidy","intro":"Compare a wide, messy table with its tidy form.",
       "src":("custom",[
         {"customer":"CUST00001","Jan_sales":120,"Feb_sales":95,"Mar_sales":140},
         {"customer":"CUST00002","Jan_sales":80,"Feb_sales":60,"Mar_sales":75},
         {"customer":"CUST00003","Jan_sales":200,"Feb_sales":210,"Mar_sales":190},
       ]),
       "show":["customer","Jan_sales","Feb_sales","Mar_sales"],
       "derive":{"quarter_total":"=B{r}+C{r}+D{r}"},
       "totals":True,
       "worked":"The wide table mixes months into columns (a tidy-data sin). Quarter Total is still computable, but a tidy form (customer, month, sales) is better for pivoting.",
       "task":"Rewrite the table in tidy form: columns customer, month, sales. How many rows result?"}]},
  ]},

 {
  "week": 2, "file": "Week2_Section_Examples.xlsx",
  "title": "Week 2 - External Data, Power Query & PivotTables",
  "modules": [
   {"id":"2.1","name":"Understanding External Data Sources",
    "objective":"Know where data lives (CSV, XLSX, DB) and which format carries what.",
    "blocks":[
      {"title":"Format trade-offs","intro":"Match a format to a job.",
       "src":("custom",[
         {"Format":"CSV","Best_for":"Flat exports, SQL feeds","Weakness":"No types/sheets"},
         {"Format":"XLSX","Best_for":"Modelled workbooks","Weakness":"Not version-control friendly"},
         {"Format":"SQLite","Best_for":"Relational, joined data","Weakness":"Needs a engine"},
         {"Format":"PDF","Best_for":"Reports","Weakness":"Hard to re-extract"},
       ]),
       "show":["Format","Best_for","Weakness"],"derive":{},"totals":False,
       "worked":"CSV is king for raw data; SQLite for joined e-commerce; PDF only for final reports.",
       "task":"Which format would you use to share a cleaned, analysed dashboard? Why?"}]},

   {"id":"2.2","name":"Introduction to Power Query",
    "objective":"Understand the ETL engine and the query editor layout.",
    "blocks":[
      {"title":"A first query","intro":"Load the e-commerce products and trim names.",
       "src":("ecom_prod",{"cols":["product_id","product_name","category","price"],"n":8}),
       "show":["product_id","product_name","category","price"],
       "derive":{"price_k":"=D{r}/1000"},
       "totals":True,
       "worked":"Power Query's Applied Steps record every transform (trim, change type). Price shown in thousands for readability.",
       "task":"List the three stages of ETL and what Power Query does at each."}]},

   {"id":"2.3","name":"Connecting and Importing with Power Query",
    "objective":"Use the connectors you'll actually need: CSV, workbook, folder.",
    "blocks":[
      {"title":"Import & type","intro":"Bring in the salary survey and set Salary to Decimal.",
       "src":("salary",{"cols":["job_title","education_level","years_of_experience","salary"],"n":8}),
       "show":["job_title","education_level","years_of_experience","salary"],
       "derive":{"salary_k":"=D{r}/1000"},
       "totals":True,
       "worked":"Connecting is one step; fixing types (Salary as Decimal, not Text) is where quality is won.",
       "task":"Why must Salary be a number before you can average it?"}]},

   {"id":"2.4","name":"External Web and PDF Data Sources",
    "objective":"Source, import, clean, validate and interpret external feeds.",
    "blocks":[
      {"title":"Validate an extract","intro":"Check a hospital extract for blanks before trusting it.",
       "src":("health",{"cols":["name","age","medical_condition","billing_amount"],"n":8}),
       "show":["name","age","medical_condition","billing_amount"],
       "derive":{"billing_k":"=D{r}/1000"},
       "totals":True,
       "worked":"Validation rule: billing_amount must be > 0 and age between 0 and 120. Flag any row that breaks the rule.",
       "task":"Write two validation rules you would apply to this extract."}]},

   {"id":"2.5","name":"Cleaning and Transforming in Power Query",
    "objective":"Twenty transformations: split, trim, pivot, unpivot, fill, replace.",
    "blocks":[
      {"title":"Trim & standardise","intro":"Fix mixed-case categories in one step.",
       "src":("stdep",{"cols":["gender","dietary_habits","cgpa","depression"],"n":8}),
       "show":["gender","dietary_habits","cgpa","depression"],
       "derive":{"cgpa_x10":"=C{r}*10"},
       "totals":True,
       "worked":"Power Query's 'Capitalise Each Word' fixes inconsistent casing in dietary_habits. CGPA scaled x10 for display.",
       "task":"Which transformation would you use to turn full names into first/last columns?"}]},

   {"id":"2.6","name":"Excel Cleaning versus Power Query Cleaning",
    "objective":"Decide when manual in-cell cleaning is right vs. a repeatable query.",
    "blocks":[
      {"title":"Repeatable vs one-off","intro":"A one-off fix vs a monthly refresh.",
       "src":("custom",[
         {"Task":"Fix 3 typos","Use":"In-cell edit","Reason":"One-off"},
         {"Task":"Clean 50k rows monthly","Use":"Power Query","Reason":"Repeatable, auditable"},
         {"Task":"Standardise labels","Use":"Power Query","Reason":"Consistent steps"},
       ]),
       "show":["Task","Use","Reason"],"derive":{},"totals":False,
       "worked":"If you will redo it, use Power Query. One-off edits are fine in-cell.",
       "task":"Classify 'convert USD text to number' as one-off or repeatable and justify."}]},

   {"id":"2.7","name":"The 65 Analytical Excel Functions",
    "objective":"Purpose, syntax and a real example for the core function families.",
    "blocks":[
      {"title":"SUMIFS in action","intro":"Total superstore sales for a region and category.",
       "src":("supers",{"cols":["region","category","sales"],"n":10}),
       "show":["region","category","sales"],
       "derive":{"sales_k":"=C{r}/1000"},
       "totals":True,
       "worked":"=SUMIFS(sales, region, \"West\", category, \"Technology\") sums only matching rows. Build the criteria from the table above.",
       "task":"Write a SUMIFS that totals Sales where Region = 'East'."},
      {"title":"XLOOKUP a price","intro":"Pull a product price from the catalogue.",
       "src":("ecom_prod",{"cols":["product_id","product_name","price"],"n":8}),
       "show":["product_id","product_name","price"],
       "derive":{},"totals":True,
       "worked":"=XLOOKUP(\"PROD00018\", product_id, price) returns that product's price. Lookups replace manual searching.",
       "task":"Use XLOOKUP to find the price of PROD00001 in the table above."}]},

   {"id":"2.8","name":"Introduction to the Power Query M Language",
    "objective":"Read let-expressions, types and step folding in M.",
    "blocks":[
      {"title":"Read an M step","intro":"A simple M expression and what it returns.",
       "src":("custom",[
         {"M_code":"let Source = Csv.Document(...) in Source","Meaning":"Load the CSV"},
         {"M_code":"Table.TransformColumns(.., trim)","Meaning":"Trim text columns"},
         {"M_code":"Table.SelectRows(.., each [age]>0)","Meaning":"Keep valid ages"},
       ]),
       "show":["M_code","Meaning"],"derive":{},"totals":False,
       "worked":"Every Power Query step is an M function. SelectRows keeps only rows meeting a condition (here age>0).",
       "task":"Translate 'keep rows where billing_amount > 0' into M-style pseudo-code."}]},

   {"id":"2.9","name":"Introduction to PivotTables",
    "objective":"Design the four-area PivotTable and avoid the Sum-vs-Count trap.",
    "blocks":[
      {"title":"Region x Category revenue","intro":"Pivot superstore sales by Region (rows) and Category (columns).",
       "src":("supers",{"cols":["region","category","sales","profit"],"n":14}),
       "show":["region","category","sales","profit"],
       "derive":{"margin":"=D{r}/C{r}"},
       "totals":True,
       "worked":"Drag Region to Rows, Category to Columns, Sales to Values (Sum). Watch the Values aggregation - ensure it is Sum, not Count.",
       "task":"Which Region-Category cell has the highest Sales in your slice?"}]},

   {"id":"2.10","name":"Introduction to PivotCharts",
    "objective":"Create the visual twin of a PivotTable and use the five-zone interface.",
    "blocks":[
      {"title":"Chart the margins","intro":"Turn the Region revenue pivot into a column chart.",
       "src":("supers",{"cols":["region","sales","profit"],"n":4,"group_by":"region"}),
       "show":["region","sales","profit"],
       "derive":{"profit_ratio":"=C{r}/B{r}"},
       "totals":True,
       "worked":"A PivotChart reads the same cache as its PivotTable. Profit Ratio = Profit/Sales shows efficiency per region.",
       "task":"Which region has the best profit ratio in the summary above?"}]},
  ]},

 {
  "week": 3, "file": "Week3_Section_Examples.xlsx",
  "title": "Week 3 - SQL, Data Model, DAX & Dashboards",
  "modules": [
   {"id":"3.1","name":"Database Fundamentals",
    "objective":"Tables, keys, relational design and normalisation.",
    "blocks":[
      {"title":"Keys & relationships","intro":"The e-commerce schema links five tables by keys.",
       "src":("custom",[
         {"Table":"customers","Primary_Key":"customer_id","Foreign_Keys":"-"},
         {"Table":"sellers","Primary_Key":"seller_id","Foreign_Keys":"-"},
         {"Table":"products","Primary_Key":"product_id","Foreign_Keys":"seller_id"},
         {"Table":"deliveries","Primary_Key":"delivery_id","Foreign_Keys":"customer_id"},
         {"Table":"transactions","Primary_Key":"transaction_id","Foreign_Keys":"customer_id, product_id, seller_id, delivery_id"},
       ]),
       "show":["Table","Primary_Key","Foreign_Keys"],"derive":{},"totals":False,
       "worked":"A primary key uniquely identifies a row; a foreign key points to a parent. transactions is the fact table joining the other four.",
       "task":"Draw the join path from transactions to sellers."}]},

   {"id":"3.2","name":"Introduction to SQL and SQLite",
    "objective":"SELECT, WHERE, ORDER BY, aggregate and GROUP BY.",
    "blocks":[
      {"title":"Top categories by revenue","intro":"A real query against ecommerce.db.",
       "src":("ecom_tx",{"cols":["transaction_id","product_id","total_price","order_date"],"n":8}),
       "show":["transaction_id","product_id","total_price","order_date"],
       "derive":{"price_k":"=C{r}/1000"},
       "totals":True,
       "worked":"SQL: SELECT category, SUM(total_price) FROM transactions JOIN products ... GROUP BY category. In Excel we pre-sum with a helper column.",
       "task":"Write the SQL that counts transactions per payment_method."}]},

   {"id":"3.3","name":"Connecting Excel to Databases with ODBC",
    "objective":"Reach a database from Excel via Get Data > From Database.",
    "blocks":[
      {"title":"One query, many tools","intro":"The same SQL feeds Excel and a dashboard.",
       "src":("ecom_sell",{"cols":["seller_id","store_name","store_state","total_revenue"],"n":8}),
       "show":["seller_id","store_name","store_state","total_revenue"],
       "derive":{"rev_k":"=D{r}/1000"},
       "totals":True,
       "worked":"ODBC lets Excel pull a live query result. total_revenue shown in thousands.",
       "task":"Name two benefits of querying the DB instead of pasting a CSV."}]},

   {"id":"3.4","name":"Merging and Appending in Power Query",
    "objective":"The six join kinds, fuzzy match and expansion.",
    "blocks":[
      {"title":"Left join transactions to products","intro":"Enrich sales with category via a key.",
       "src":("ecom_tx",{"cols":["transaction_id","product_id","total_price"],"n":8}),
       "show":["transaction_id","product_id","total_price"],
       "derive":{"price_k":"=C{r}/1000"},
       "totals":True,
       "worked":"A Left Join keeps every transaction and adds product columns where the key matches. total_price in thousands.",
       "task":"What rows would an Inner Join drop compared to a Left Join here?"}]},

   {"id":"3.5","name":"The Excel Data Model and Star Schema",
    "objective":"Relationships, fact/dimension tables and the date table.",
    "blocks":[
      {"title":"Star schema sketch","intro":"One fact, four dimensions.",
       "src":("custom",[
         {"Table":"transactions (fact)","Grain":"one line item","Links_to":"customers, products, sellers, deliveries"},
         {"Table":"products (dim)","Grain":"one product","Links_to":"transactions"},
         {"Table":"customers (dim)","Grain":"one customer","Links_to":"transactions"},
       ]),
       "show":["Table","Grain","Links_to"],"derive":{},"totals":False,
       "worked":"The fact table holds measures (total_price); dimensions hold descriptive attributes. Relationships replace VLOOKUP.",
       "task":"Why is a star schema faster than a single flat 50-column table?"}]},

   {"id":"3.6","name":"Power Pivot and DAX: Foundations and Context",
    "objective":"DAX vs Excel formulas and the two evaluation contexts.",
    "blocks":[
      {"title":"A measure vs a column","intro":"Total revenue as a DAX measure.",
       "src":("ecom_tx",{"cols":["transaction_id","seller_id","total_price"],"n":8}),
       "show":["transaction_id","seller_id","total_price"],
       "derive":{"price_k":"=C{r}/1000"},
       "totals":True,
       "worked":"DAX: Total Revenue := SUM(transactions[total_price]). A measure respects filter context (slicers); a calculated column does not.",
       "task":"Write a DAX measure for total commission_amount."}]},

   {"id":"3.7","name":"Chart Anatomy and the 17 Excel Chart Types",
    "objective":"Make any chart readable; choose the right type.",
    "blocks":[
      {"title":"Pick the right chart","intro":"Match a question to a chart type.",
       "src":("custom",[
         {"Question":"Trend over time?","Chart":"Line"},
         {"Question":"Parts of a whole?","Chart":"Pie/Donut"},
         {"Question":"Compare categories?","Chart":"Column/Bar"},
         {"Question":"Distribution?","Chart":"Histogram"},
       ]),
       "show":["Question","Chart"],"derive":{},"totals":False,
       "worked":"Charts encode questions. Time -> line; composition -> pie; comparison -> column.",
       "task":"Which chart shows the relationship between study hours and exam score?"}]},

   {"id":"3.8","name":"Dashboard Development",
    "objective":"Plan, layout, KPI cards, slicers and publication.",
    "blocks":[
      {"title":"KPI cards from e-commerce","intro":"Three headline numbers for a seller dashboard.",
       "src":("ecom_tx",{"cols":["transaction_id","seller_id","total_price","order_status"],"n":8}),
       "show":["transaction_id","seller_id","total_price","order_status"],
       "derive":{"price_k":"=C{r}/1000"},
       "totals":True,
       "worked":"A dashboard leads with KPI cards (total revenue, orders, avg order value). Add a Slicer on seller_id to make it interactive.",
       "task":"Name the three KPI cards you would put on a Hospital Operations dashboard."}]},

   {"id":"3.9","name":"Three Dashboard Blueprints",
    "objective":"Full specs for the three capstone dashboards.",
    "blocks":[
      {"title":"Blueprint: Retail","intro":"Superstore - sales, profit, ship mode.",
       "src":("supers",{"cols":["region","category","sales","profit"],"n":10}),
       "show":["region","category","sales","profit"],
       "derive":{"margin":"=D{r}/C{r}"},
       "totals":True,
       "worked":"Blueprint = KPI row + Region map + Category column + Segment slicer. Margin column flags efficiency.",
       "task":"Sketch the three visuals you would place on this dashboard."}]},

   {"id":"3.10","name":"External Data in SQL, DAX and Dashboards",
    "objective":"A repeatable source-to-dashboard pipeline.",
    "blocks":[
      {"title":"Pipeline check","intro":"Validate each hand-off.",
       "src":("custom",[
         {"Stage":"Source (CSV/DB)","Check":"Row count matches"},
         {"Stage":"Clean (PQ)","Check":"No null keys"},
         {"Stage":"Model (DAX)","Check":"Measures aggregate correctly"},
         {"Stage":"Visual","Check":"Numbers tie to source"},
       ]),
       "show":["Stage","Check"],"derive":{},"totals":False,
       "worked":"Each stage needs a reconciliation check so the dashboard's numbers tie back to the raw source.",
       "task":"What single check proves your dashboard total equals the source total?"}]},
  ]},

 {
  "week": 4, "file": "Week4_Section_Examples.xlsx",
  "title": "Week 4 - Statistics, Forecasting, AI & Communication",
  "modules": [
   {"id":"4.1","name":"Descriptive Statistics and the Analysis ToolPak",
    "objective":"Central tendency, spread, shape and ToolPak summaries.",
    "blocks":[
      {"title":"Summarise salaries","intro":"Mean, median, stdev of a salary slice.",
       "src":("salary",{"cols":["age","years_of_experience","salary"],"n":12}),
       "show":["age","years_of_experience","salary"],
       "derive":{"salary_k":"=C{r}/1000"},
       "totals":True,
       "worked":"Descriptive stats answer 'what is typical?'. Use AVERAGE, MEDIAN, STDEV.S. Salary shown in thousands.",
       "task":"Which is higher here, mean or median salary, and what does that imply?"}]},

   {"id":"4.2","name":"Distributions, Pareto, Percentiles and Z-Scores",
    "objective":"Read shape, apply 80/20, standardise and detect outliers.",
    "blocks":[
      {"title":"Pareto on categories","intro":"A few categories drive most sales.",
       "src":("supers",{"cols":["category","sales"],"n":6,"group_by":"category"}),
       "show":["category","sales"],
       "derive":{"sales_k":"=B{r}/1000"},
       "totals":True,
       "worked":"Rank categories by sales (desc), add a running % - usually ~20% of items drive ~80% of value (Pareto).",
       "task":"After sorting, what share of total sales do the top two categories represent?"}]},

   {"id":"4.3","name":"Correlation Analysis",
    "objective":"Measure association honestly with Pearson/Spearman.",
    "blocks":[
      {"title":"Study hours vs score","intro":"Do more hours mean higher scores?",
       "src":("stperf",{"cols":["hours_studied","previous_scores","exam_score"],"n":12}),
       "show":["hours_studied","previous_scores","exam_score"],
       "derive":{"score_x2":"=C{r}*2"},
       "totals":True,
       "worked":"Correlation (r) near +1 means strong positive link. Here exam_score tends to rise with hours_studied.",
       "task":"Would you call the relationship between hours_studied and exam_score strong, weak or none?"}]},

   {"id":"4.4","name":"Linear Regression",
    "objective":"Fit a line, read R-squared and coefficients, validate.",
    "blocks":[
      {"title":"Experience predicts salary","intro":"Fit salary on experience.",
       "src":("salary",{"cols":["years_of_experience","salary"],"n":12,"sort_by":"years_of_experience","ascending":True}),
       "show":["years_of_experience","salary"],
       "derive":{"salary_k":"=B{r}/1000"},
       "totals":True,
       "worked":"Regression: salary = a + b*experience. The slope b is the average raise per year; R-squared is fit quality.",
       "task":"If slope = 3000, what extra salary does 5 more years of experience imply?"}]},

   {"id":"4.5","name":"Time Series and Forecasting",
    "objective":"Trend, seasonality, moving averages, FORECAST.ETS.",
    "blocks":[
      {"title":"Monthly e-commerce revenue","intro":"Aggregate transactions by month, then forecast.",
       "src":("ecom_tx",{"cols":["order_date","total_price"],"n":10,"group_by":"month"}),
       "show":["order_date","total_price"],
       "derive":{"rev_k":"=B{r}/1000"},
       "totals":True,
       "worked":"Group transactions by month (substr(order_date,1,7)), plot the trend, then FORECAST.ETS projects the next period.",
       "task":"Is the monthly revenue trend rising or falling in your slice?"}]},

   {"id":"4.6","name":"Goal Seek, Data Tables, Scenarios and Solver",
    "objective":"What-if analysis answering management questions.",
    "blocks":[
      {"title":"Hit a revenue target","intro":"How many units to reach a goal?",
       "src":("custom",[
         {"Scenario":"Base","Units":100,"Price":5000,"Revenue":500000},
         {"Scenario":"Grow","Units":120,"Price":5000,"Revenue":600000},
         {"Scenario":"Premium","Units":100,"Price":5500,"Revenue":550000},
       ]),
       "show":["Scenario","Units","Price","Revenue"],
       "derive":{},"totals":True,
       "worked":"Goal Seek finds the Units that make Revenue hit a target. Data Tables show Revenue across price x units.",
       "task":"Use Goal Seek logic: how many units at 5,000 price yield 750,000 revenue?"}]},

   {"id":"4.7","name":"Data Storytelling",
    "objective":"Turn a finding into a decision with message hierarchy.",
    "blocks":[
      {"title":"Lead with the message","intro":"One sentence, then the evidence.",
       "src":("health",{"cols":["medical_condition","billing_amount"],"n":6,"group_by":"medical_condition"}),
       "show":["medical_condition","billing_amount"],
       "derive":{"billing_k":"=B{r}/1000"},
       "totals":True,
       "worked":"Story = headline ('Cancer drives the highest billing') + chart + recommendation. Billing shown in thousands.",
       "task":"Write the one-sentence headline for the table above."}]},

   {"id":"4.8","name":"Report Writing and PDF Production",
    "objective":"Structure, layout, page setup and clean PDF export.",
    "blocks":[
      {"title":"Report skeleton","intro":"The sections every analytical report needs.",
       "src":("custom",[
         {"Section":"Executive Summary","Purpose":"The decision in 5 lines"},
         {"Section":"Method","Purpose":"Data + approach"},
         {"Section":"Findings","Purpose":"Charts + tables"},
         {"Section":"Recommendation","Purpose":"What to do"},
       ]),
       "show":["Section","Purpose"],"derive":{},"totals":False,
       "worked":"A report argues: summary first, evidence after, recommendation last. Export to PDF with set print areas.",
       "task":"Which section would a busy CEO read first, and why?"}]},

   {"id":"4.9","name":"Slide Presentations",
    "objective":"Decks that argue: one message per slide, charts not decoration.",
    "blocks":[
      {"title":"One message per slide","intro":"Map findings to slides.",
       "src":("supers",{"cols":["region","sales","profit"],"n":4,"group_by":"region"}),
       "show":["region","sales","profit"],
       "derive":{"profit_ratio":"=C{r}/B{r}"},
       "totals":True,
       "worked":"Each slide = one claim + one chart. Region profit summary supports the 'where we earn' slide.",
       "task":"Draft the title of a slide whose message is 'The West region is most profitable'."}]},

   {"id":"4.10","name":"Formula Auditing, Protection and Light Automation",
    "objective":"Trace errors, protect workbooks, basic macros.",
    "blocks":[
      {"title":"Trace a formula","intro":"Find why a total is wrong.",
       "src":("custom",[
         {"Cell":"B7","Formula":"=SUM(B2:B6)","Risk":"missing row 7"},
         {"Cell":"C7","Formula":"=B7*1.1","Risk":"hard-coded tax"},
         {"Cell":"D7","Formula":"=VLOOKUP(A7,..)","Risk":"missing key"},
       ]),
       "show":["Cell","Formula","Risk"],"derive":{},"totals":False,
       "worked":"Trace Precedents shows which cells feed a formula. Protect sheets but leave input cells unlocked.",
       "task":"Which of the three formulas above would break if a new row is inserted?"}]},

   {"id":"4.11","name":"AI Co-pilots for Analysts",
    "objective":"Prompt design, schema prompts, safe LLM use.",
    "blocks":[
      {"title":"A good analysis prompt","intro":"Give the model context and constraints.",
       "src":("custom",[
         {"Bad_prompt":"analyse sales","Why_weak":"no data, no goal"},
         {"Good_prompt":"Given the ecommerce transactions table (cols: total_price, seller_id), write a SQL query for top 5 sellers by revenue.","Why_strong":"schema + goal + output"},
       ]),
       "show":["Bad_prompt","Why_weak"],"derive":{},"totals":False,
       "worked":"Strong prompts name the table, columns and desired output. Never paste PII into a public model.",
       "task":"Rewrite 'clean the data' into a prompt that specifies the dataset and the exact issues to fix."}]},

   {"id":"4.12","name":"Verifying AI Output and Handling Hallucinations",
    "objective":"A verification protocol that makes AI output safe.",
    "blocks":[
      {"title":"Trust but verify","intro":"Reconcile any AI number to source.",
       "src":("ecom_tx",{"cols":["transaction_id","total_price"],"n":6}),
       "show":["transaction_id","total_price"],
       "derive":{"price_k":"=B{r}/1000"},
       "totals":True,
       "worked":"Protocol: (1) ask for the query, (2) run it yourself, (3) tie the total to the source. Never trust a number you can't reproduce.",
       "task":"An AI says total revenue is 999,999,999. What one check disproves or confirms it?"}]},

   {"id":"4.13","name":"Data Ethics, PII and PHI",
    "objective":"Consent, minimisation, anonymisation, bias, retention.",
    "blocks":[
      {"title":"Spot the PII","intro":"Which fields must be protected?",
       "src":("health",{"cols":["name","age","medical_condition","billing_amount"],"n":6}),
       "show":["name","age","medical_condition","billing_amount"],
       "derive":{"billing_k":"=D{r}/1000"},
       "totals":True,
       "worked":"Name + medical condition = PHI. Anonymise (drop name, generalise age band) before sharing externally.",
       "task":"Which two columns above should be removed before a public release?"}]},

   {"id":"4.14","name":"External Benchmarks, PDF Extraction and Reconciliation",
    "objective":"Source, import, clean, validate, interpret external benchmarks.",
    "blocks":[
      {"title":"Reconcile to source","intro":"Prove your total equals the extract's total.",
       "src":("hotel",{"cols":["booking_id","avg_price_per_room","booking_status"],"n":10}),
       "show":["booking_id","avg_price_per_room","booking_status"],
       "derive":{"price_k":"=B{r}/1000"},
       "totals":True,
       "worked":"Reconciliation: SUM(avg_price_per_room) here must equal the same sum in the raw CSV. Differences mean a load error.",
       "task":"If your Excel sum is 12,340 but the CSV sum is 12,355, name two possible causes."}]},
  ]},

 {
  "week": 5, "file": "Week5_Section_Examples.xlsx",
  "title": "Week 5 - Capstone, Executive Defence & Career Launch",
  "modules": [
   {"id":"P.1","name":"The Seven-Stage Project Lifecycle",
    "objective":"Scope, acquire, clean, model, analyse, communicate, hand over.",
    "blocks":[
      {"title":"Lifecycle checklist","intro":"A repeatable structure for all three capstones.",
       "src":("custom",[
         {"Stage":"1 Scope","Output":"Brief & KPIs"},
         {"Stage":"2 Acquire","Output":"Raw datasets"},
         {"Stage":"3 Clean","Output":"Tidy tables"},
         {"Stage":"4 Model","Output":"Joins / measures"},
         {"Stage":"5 Analyse","Output":"Findings"},
         {"Stage":"6 Communicate","Output":"Dashboard + report"},
         {"Stage":"7 Hand over","Output":"Files + README"},
       ]),
       "show":["Stage","Output"],"derive":{},"totals":False,
       "worked":"Every capstone follows these seven stages. Skip none - stage 3 (clean) is where most projects fail.",
       "task":"For the Healthcare capstone, write the KPI for stage 1."}]},

   {"id":"P.2","name":"Source Governance, Citation and Capstone Delivery",
    "objective":"Document provenance, cite datasets, deliver reproducibly.",
    "blocks":[
      {"title":"Provenance log","intro":"Record where each file came from.",
       "src":("custom",[
         {"Dataset":"Healthcare Records","Source":"Provided bootcamp set","Licence":"Bootcamp use"},
         {"Dataset":"Superstore Sales","Source":"Provided bootcamp set","Licence":"Bootcamp use"},
         {"Dataset":"E-Commerce","Source":"Provided bootcamp set (SQL)","Licence":"Bootcamp use"},
       ]),
       "show":["Dataset","Source","Licence"],"derive":{},"totals":False,
       "worked":"A provenance log (file, source, licence, date) makes your work reproducible and citable.",
       "task":"Add a row to the log for the Student Performance dataset."}]},

   {"id":"P.3","name":"Capstone 1: SuperMart Retail Analytics",
    "objective":"A complete retail build: basket, margin, region, ship mode.",
    "blocks":[
      {"title":"Retail KPIs","intro":"The numbers a retail manager wants.",
       "src":("supers",{"cols":["region","category","sales","profit"],"n":12}),
       "show":["region","category","sales","profit"],
       "derive":{"margin":"=D{r}/C{r}"},
       "totals":True,
       "worked":"Retail capstone = Sales, Profit, Margin%, by Region & Category, plus Ship Mode mix. Margin flags efficiency.",
       "task":"Which Region-Category pair has the best margin in your slice?"}]},

   {"id":"P.4","name":"Capstone 2: Crestview University Progression",
    "objective":"Education build: factors driving exam performance.",
    "blocks":[
      {"title":"Drivers of performance","intro":"Which factors move the exam score?",
       "src":("stperf",{"cols":["hours_studied","attendance","sleep_hours","exam_score"],"n":12}),
       "show":["hours_studied","attendance","sleep_hours","exam_score"],
       "derive":{"score_x2":"=D{r}*2"},
       "totals":True,
       "worked":"Education capstone correlates Exam Score with study hours, attendance, sleep. Flag the strongest driver.",
       "task":"Name the single factor you would tell students to improve first."}]},

   {"id":"P.5","name":"Capstone 3: Lifecare Hospital Operations",
    "objective":"Healthcare build: occupancy, length of stay, billing.",
    "blocks":[
      {"title":"Operational KPIs","intro":"What a hospital COO tracks.",
       "src":("health",{"cols":["medical_condition","length_of_stay","billing_amount"],"n":10,"group_by":"medical_condition"}),
       "show":["medical_condition","length_of_stay","billing_amount"],
       "derive":{"billing_k":"=C{r}/1000"},
       "totals":True,
       "worked":"Hospital capstone = avg Length of Stay and Billing by condition. Billing in thousands; LOS in days.",
       "task":"Which condition has the longest average stay in your slice?"}]},

   {"id":"P.6","name":"Workbook QA and Stress Testing",
    "objective":"Prove a workbook is correct, fast, portable, maintainable.",
    "blocks":[
      {"title":"QA checklist","intro":"Five checks before hand-over.",
       "src":("custom",[
         {"Check":"Totals tie to source","Pass_criteria":"diff = 0"},
         {"Check":"No #REF! errors","Pass_criteria":"0 errors"},
         {"Check":"Opens on another PC","Pass_criteria":"links resolved"},
         {"Check":"Slicers work","Pass_criteria":"filters apply"},
         {"Check":"README present","Pass_criteria":"yes"},
       ]),
       "show":["Check","Pass_criteria"],"derive":{},"totals":False,
       "worked":"QA = reconciliation + error scan + portability + documentation. Tick each before you submit.",
       "task":"Which check would catch a broken VLOOKUP after a column insert?"}]},

   {"id":"P.7","name":"The Ten-Minute Executive Defence",
    "objective":"Present and defend a build to a panel: structure, evidence, Q&A.",
    "blocks":[
      {"title":"Defence script","intro":"Open with the decision, close with the ask.",
       "src":("ecom_tx",{"cols":["seller_id","total_price"],"n":6,"group_by":"seller_id"}),
       "show":["seller_id","total_price"],
       "derive":{"rev_k":"=B{r}/1000"},
       "totals":True,
       "worked":"Defence = 2 min context, 5 min findings (top sellers by revenue), 3 min recommendation + Q&A.",
       "task":"Write the one-sentence recommendation you would open with."}]},

   {"id":"P.8","name":"Portfolio, Anonymisation and CV",
    "objective":"Package three capstones into interview-winning evidence.",
    "blocks":[
      {"title":"Portfolio entries","intro":"One card per project.",
       "src":("custom",[
         {"Project":"Retail Analytics","Stack":"Excel, PQ, Pivot","Result":"+18% margin insight"},
         {"Project":"University Progression","Stack":"Stats, Regression","Result":"Top driver: attendance"},
         {"Project":"Hospital Ops","Stack":"SQL, Dashboard","Result":"LOS reduced 2 days"},
       ]),
       "show":["Project","Stack","Result"],"derive":{},"totals":False,
       "worked":"Each portfolio entry states the stack, the question and the quantified result. Quantify every claim.",
       "task":"Write a one-line CV bullet for the Retail Analytics project."}]},

   {"id":"P.9","name":"Technical Interviews and the Roadmap Beyond Excel",
    "objective":"Interview formats, live tests, and a concrete next-step roadmap.",
    "blocks":[
      {"title":"Interview prep map","intro":"What to practise before the call.",
       "src":("custom",[
         {"Topic":"Excel","Drill":"Pivot + XLOOKUP"},
         {"Topic":"SQL","Drill":"JOIN + GROUP BY"},
         {"Topic":"Stats","Drill":"Regression interpretation"},
         {"Topic":"Story","Drill":"Explain a past analysis"},
       ]),
       "show":["Topic","Drill"],"derive":{},"totals":False,
       "worked":"Interviews test the same skills as the bootcamp. Practise the live Excel/SQL drill daily.",
       "task":"Which one SQL drill would you practise tonight, and why?"}]},

   {"id":"P.10","name":"The 100-Mark Assessment Rubric",
    "objective":"The published marking scheme distinguishing pass from distinction.",
    "blocks":[
      {"title":"Rubric bands","intro":"Know what earns the marks.",
       "src":("custom",[
         {"Criterion":"Data cleaning","Weight":"25","Distinction":"No errors, documented"},
         {"Criterion":"Analysis","Weight":"30","Distinction":"Right method, insight"},
         {"Criterion":"Dashboard","Weight":"25","Distinction":"Clean, interactive"},
         {"Criterion":"Communication","Weight":"20","Distinction":"Clear story"},
       ]),
       "show":["Criterion","Weight","Distinction"],"derive":{},"totals":True,
       "worked":"Total = 100. Spend effort where the weight is: Analysis (30) and Cleaning (25) dominate.",
       "task":"If you score full on Cleaning and Analysis but low on Communication, what is your max total?"}]},
  ]},
]
