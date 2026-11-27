from revlib import T

ON = "SuperMart Online"
TRY = "ETEH_Week3_Section_Examples.xlsx  ›  sheet '3.2 SQL with SQLite'  ·  ETEH_Week3_SQL_Examples.sql"


def S(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", ON)
    return T(tag, **k)


TOPICS = [
    S("3.2.1",
      what="**SQL** (Structured Query Language) is how you ask a database questions. You describe **what** you "
           "want ('total revenue by category') and the database works out **how** to get it. Analysts mostly use "
           "**SELECT** to read data.",
      like="SQL is like ordering at a restaurant: you say what you want on the plate; the kitchen decides how to "
           "cook it.",
      how=["Open the database in a SQL tool.", "Write a SELECT statement.", "Run it and read the result grid."],
      tip="Start every query by writing the business question as a comment (-- …).",
      trap="Thinking SQL changes the data; SELECT only reads it.",
      case="The analyst asks supermart_online.db a first question: how many order lines are there, and what are "
           "they worth?",
      steps=["Type: SELECT COUNT(*), SUM(total_price) FROM transactions;", "Run the query.",
             "Read the one-row answer."],
      table=[["Result column", "Value"], ["COUNT(*)", "11,000 order lines"], ["SUM(total_price)", "₦1,394,713,700"]],
      so="One line of SQL summarised all 11,000 orders (2020–2025) without opening a single row.",
      split=4.90),

    S("3.2.2",
      what="SQL is a standard, but each database adds its own **dialect**: SQLite, PostgreSQL, SQL Server (T-SQL), "
           "MySQL and others differ in date functions, row limits and text handling. The core (SELECT, WHERE, "
           "GROUP BY, JOIN) is the same everywhere.",
      like="Dialects are like English in Lagos, London and New York: same language, a few different words.",
      how=["Learn the core clauses once.", "Check the dialect's date and limit syntax.", "Test text matching (case sensitivity)."],
      tip="Note the dialect at the top of every saved query.",
      trap="Copying a query between databases and assuming it behaves the same.",
      case="The team's queries run on SQLite now, but may move to SQL Server later.",
      steps=["Compare the row-limit syntax.", "Compare how the year is extracted from a date.",
             "Test LIKE '%laptop%' against product names containing 'Laptop'."],
      table=[["Feature", "SQLite", "SQL Server"], ["First 10 rows", "LIMIT 10", "TOP 10"],
             ["Year of a date", "strftime('%Y', order_date)", "YEAR(order_date)"],
             ["Join text", "a || b", "a + b"], ["LIKE '%laptop%' finds 'Laptop'?", "Yes (10 rows)", "Depends on collation"]],
      so="The ideas carry over; a few keywords change. SQLite's LIKE ignored case and found all 10 laptops.",
      split=4.20),

    S("3.2.3",
      what="**SQLite** is a complete SQL database stored in **one file**. There's no server to install or log in "
           "to, so it is ideal for learning and for sharing data. It supports almost all the SQL an analyst needs.",
      like="SQLite is like a USB stick with a whole filing cabinet inside: copy one file and you've copied the "
           "entire database.",
      how=["Find the .db file.", "Open it in a SQLite tool (no installation of a server).", "Query it like any other database."],
      tip="Keep a backup copy of the .db file before you experiment.",
      trap="Expecting server features like user logins; SQLite has none.",
      case="The bootcamp databases are SQLite files in Week3_Data_Files.",
      steps=["Look at the files and their sizes.", "Open supermart_online.db.", "Count its tables."],
      table=[["File", "Size", "Main table"], ["supermart_online.db", "5.4 MB", "transactions: 11,000 rows"],
             ["lifecare_hospital.db", "17.1 MB", "admissions: 54,860 rows"]],
      so="Two complete databases fit in about 22 MB, and each opens with no installation or password.",
      split=4.20),

    S("3.2.4",
      what="Common tools for SQLite: **DB Browser for SQLite** (free, with a table browser and result grid), the "
           "**sqlite3** command line (fast, scriptable), and **Excel / Power Query** through an ODBC driver "
           "(Module 3.3).",
      like="The tools are like ways into the same building: the front door (DB Browser), the staff entrance "
           "(command line) and the loading bay (Excel).",
      how=["Install DB Browser for SQLite.", "File > Open Database > choose the .db file.",
           "Open the Execute SQL tab, paste a query and press Ctrl+Enter."],
      tip="Beginners: start with DB Browser for SQLite.",
      trap="Editing data in the Browse Data tab by accident; changes save when you click Write Changes.",
      case="A new analyst sets up DB Browser for SQLite to run the bootcamp queries.",
      steps=["Open supermart_online.db.", "Open ETEH_Week3_SQL_Examples.sql in the Execute SQL tab.",
             "Run the first query (3.1.3)."],
      table=[["Tool", "Strength", "Use it for"], ["DB Browser for SQLite", "Visual, beginner-friendly", "Learning, exploring"],
             ["sqlite3 command line", "Fast, scriptable", "Repeat jobs"], ["Excel via ODBC", "Results land in Excel", "Reports (3.3)"]],
      so="The same .sql file runs in any of these tools and gives the same answer.",
      split=4.20),

    S("3.2.5",
      what="Every query starts with **SELECT columns FROM table**. **SELECT *** returns every column; it's fine "
           "for a quick look, but saved queries should list the columns they need. Add **LIMIT** to preview a few "
           "rows.",
      like="SELECT is like telling a shop assistant which items to bring from the store room; FROM is which store "
           "room.",
      how=["Write SELECT followed by column names, separated by commas.", "Write FROM and the table name.",
           "Add LIMIT 5 while exploring."],
      tip="List the columns by name in saved queries.",
      trap="Running SELECT * on a huge table without LIMIT.",
      case="The analyst previews the products table.",
      steps=["SELECT product_id, product_name, category, price", "FROM products", "LIMIT 5;"],
      table=[["product_id", "product_name", "category", "price"], ["PROD00001", "Wristwatch", "Fashion", "70,000"],
             ["PROD00002", "Wardrobe", "Home", "330,000"], ["PROD00003", "Palm Oil 5L", "Food", "13,000"],
             ["PROD00004", "Indomie Noodles Carton", "Food", "12,650"]],
      so="Four chosen columns, first five rows: a quick, safe look at a table of 800 products.",
      split=4.20),

    S("3.2.6",
      what="**WHERE** keeps only the rows that meet a condition. Operators: **= <> > < >= <=**, **BETWEEN**, "
           "**IN (…)**, **LIKE '%text%'** and **IS NULL**. Text goes in **single quotes**. Missing values need "
           "**IS NULL**, never = NULL.",
      like="WHERE is like a bouncer with a guest list: only rows that match the rule get in.",
      how=["Add WHERE after FROM.", "Write the condition, e.g. price > 500000.", "Use IS NULL for missing values."],
      tip="Use IS NULL and IS NOT NULL for missing values.",
      trap="Writing = NULL, which matches nothing.",
      case="The buyer wants Electronics priced above ₦500,000. Operations asks how many orders have no delivery.",
      steps=["WHERE category = 'Electronics' AND price > 500000", "WHERE delivery_id IS NULL",
             "Compare with WHERE delivery_id = NULL."],
      table=[["Query", "Rows"], ["Electronics over ₦500,000", "22 (top: Dell Inspiron ₦950,000)"],
             ["delivery_id IS NULL", "570"], ["delivery_id = NULL", "0 (wrong)"]],
      so="IS NULL finds the 570 orders with no delivery. = NULL silently finds none.",
      split=4.90),

    S("3.2.7",
      what="**AND** needs both conditions, **OR** needs either, **NOT** reverses one. SQL does AND **before** OR, "
           "so when you mix them, use **brackets** to say what you mean.",
      like="It's like 'tea or coffee with sugar': does sugar go with both, or only coffee? Brackets settle it.",
      how=["Write each condition separately.", "Join with AND / OR.", "Put brackets around the OR part when mixing."],
      tip="Always use brackets when AND and OR appear together.",
      trap="Leaving out brackets, so the conditions group the wrong way.",
      case="Marketing wants Lagos or FCT customers who are not Bronze members.",
      steps=["WHERE (state = 'Lagos' OR state = 'FCT') AND NOT membership_tier = 'Bronze'",
             "Run it again without the brackets.", "Compare the counts."],
      table=[["Version", "Customers"], ["With brackets (correct)", "210"], ["Without brackets", "247"],
             ["Correct split", "Gold 74, Platinum 51, Silver 85"]],
      so="Without brackets, SQL included Bronze members from Lagos: 37 customers too many.",
      split=4.90),

    S("3.2.8",
      what="**ORDER BY column** sorts the result; add **DESC** for largest first. You can sort by several columns "
           "in turn. Without ORDER BY, the row order is **not guaranteed**.",
      like="ORDER BY is like arranging exam scripts by score before handing them back, highest first.",
      how=["Add ORDER BY at the end of the query.", "Add DESC for descending.", "Add a second column to break ties."],
      tip="Sort on purpose; never rely on the order rows happen to come back in.",
      trap="Assuming the result is sorted because it looked sorted last time.",
      case="Finance wants the biggest orders first.",
      steps=["SELECT order_id, order_date, total_price FROM transactions", "ORDER BY total_price DESC",
             "Read the top rows."],
      table=[["order_id", "order_date", "total_price"], ["ORD2025101804782", "2025-10-18", "₦3,600,000"],
             ["ORD2023040900384", "2023-04-09", "₦3,386,500"], ["ORD2023031100917", "2023-03-11", "₦3,230,000"]],
      so="Sorted by value, the largest order (₦3.6m, October 2025) comes first.",
      split=4.20),

    S("3.2.9",
      what="**LIMIT n** returns only the first n rows. Combined with **ORDER BY**, it gives a top-n list. Without "
           "ORDER BY, LIMIT gives you n rows in no particular order, not the top n.",
      like="LIMIT is like saying 'just show me the first 10 in the queue'. If the queue isn't sorted, they aren't "
           "the best 10.",
      how=["Sort with ORDER BY first.", "Add LIMIT n at the very end.", "Use LIMIT 5 to preview big queries."],
      tip="LIMIT always goes last, after ORDER BY.",
      trap="Using LIMIT 10 without ORDER BY and calling it the top 10.",
      case="Finance wants the 10 biggest orders. A junior analyst forgets ORDER BY.",
      steps=["SELECT … FROM transactions LIMIT 3 (no ORDER BY).", "Add ORDER BY total_price DESC.", "Compare the first row."],
      table=[["Query", "First row returned"], ["LIMIT 3, no ORDER BY", "₦21,600 (just the first stored row)"],
             ["ORDER BY total_price DESC LIMIT 10", "₦3,600,000 (the true largest)"]],
      so="Without sorting, LIMIT returned a ₦21,600 order. Only ORDER BY + LIMIT gives a real top 10.",
      split=4.70),

    S("3.2.10",
      what="**GROUP BY** collapses rows into **one row per group**, and aggregate functions summarise each group: "
           "**COUNT, SUM, AVG, MIN, MAX**. Every column in SELECT must be either grouped or aggregated.",
      like="GROUP BY is like sorting receipts into piles by category, then writing the count and total on each "
           "pile.",
      how=["Choose the grouping column.", "Add aggregates: COUNT(*), SUM(total_price).", "GROUP BY the grouping column."],
      tip="Filter with WHERE first, then group.",
      trap="Selecting a column that is neither grouped nor aggregated.",
      case="Management wants delivered and completed revenue by product category.",
      steps=["JOIN products to get the category.", "WHERE order_status IN ('Delivered', 'Completed')",
             "GROUP BY category, ORDER BY revenue DESC."],
      table=[["category", "orders", "revenue"], ["Electronics", "1,446", "₦648,437,900"], ["Home", "1,279", "₦327,026,350"],
             ["Fashion", "1,564", "₦110,075,900"], ["Beauty (lowest)", "1,123", "₦10,542,400"]],
      so="Electronics earns the most revenue, though Health has more orders (1,648) at small values.",
      split=4.20),

    S("3.2.11",
      what="**WHERE** filters **rows** before grouping; **HAVING** filters **groups** after grouping. Use HAVING "
           "for conditions on totals, like 'sellers with more than ₦15m revenue'.",
      like="WHERE is checking tickets at the gate; HAVING is choosing the winning teams after the match is scored.",
      how=["Filter rows with WHERE.", "Group with GROUP BY.", "Filter the totals with HAVING SUM(…) > n."],
      tip="Ask: is this condition about one row (WHERE) or a total (HAVING)?",
      trap="Putting SUM(…) in WHERE, which gives an error.",
      case="The marketplace team wants sellers with more than ₦15m of delivered revenue.",
      steps=["WHERE order_status = 'Delivered'", "GROUP BY store_name",
             "HAVING SUM(total_price) > 15000000, ORDER BY revenue DESC."],
      table=[["store_name", "orders", "revenue"], ["Jos Premium", "354", "₦82,585,050"],
             ["Alaba International Digital", "361", "₦65,316,100"], ["Ibadan Market Online", "354", "₦63,857,000"],
             ["Stores passing HAVING", "26 of 85", "Each above ₦15m"]],
      so="26 of the 85 store names with delivered sales cleared ₦15m. HAVING tested totals that WHERE never sees.",
      split=4.20),

    S("3.2.12",
      what="**AS** gives a column or table a short name (an **alias**). You can also build **calculated "
           "columns**: total_price * commission_percentage / 100.0 AS recalculated. Aliases make results readable "
           "and joins shorter.",
      like="An alias is like a nickname: 'Oluwaseun' becomes 'Seun', easier to say, same person.",
      how=["Name calculations: … AS recalculated.", "Shorten tables: FROM transactions t.", "Use 100.0, not 100, to keep decimals."],
      tip="Give every calculated column a clear alias.",
      trap="Dividing by 100 (integer) instead of 100.0, which can cut off decimals.",
      case="Finance checks that commission_amount was calculated correctly on every order.",
      steps=["ROUND(total_price * commission_percentage / 100.0, 2) AS recalculated",
             "commission_amount − recalculated AS diff", "Count rows where diff is not 0."],
      table=[["order_id", "total_price", "comm %", "diff"], ["ORD2024050500001", "21,600", "7.5", "0"],
             ["ORD2025052100002", "240,000", "12.5", "0"], ["All 11,000 orders", "(all)", "(all)", "0 mismatches"]],
      so="The recalculation matched on all 11,000 orders: commissions were charged correctly.",
      split=4.20),

    S("3.2.13",
      what="**SELECT DISTINCT** returns each different value **once**. **COUNT(DISTINCT column)** counts them. "
           "It is the quickest way to profile a column: how many couriers, customers or states are there?",
      like="DISTINCT is like listing every team that played in a league, once each, however many matches they "
           "played.",
      how=["SELECT DISTINCT column FROM table.", "Add ORDER BY to sort the list.", "Use COUNT(DISTINCT …) for the number."],
      tip="Profile every new column with COUNT(DISTINCT …).",
      trap="Using DISTINCT to hide duplicate rows instead of fixing why they appear.",
      case="Operations asks which couriers deliver for SuperMart Online.",
      steps=["SELECT DISTINCT courier_name FROM deliveries ORDER BY courier_name;",
             "COUNT(DISTINCT courier_name) and COUNT(*).", "Read the list."],
      table=[["Check", "Result"], ["Deliveries", "10,430"], ["Distinct couriers", "8"],
             ["First three", "ABC Transport, DHL Nigeria, FedEx Lagos"], ["Customers who ordered", "2,473 of 2,500"]],
      so="10,430 deliveries were handled by just 8 couriers.",
      split=4.90),

    S("3.2.14",
      what="**CASE WHEN … THEN … ELSE … END** adds if-then logic inside SQL. Use it to create bands, flags and "
           "labels, then group by them. Conditions are checked in order; the first true one wins.",
      like="CASE is like a sorting desk at the post office: small parcels here, medium there, large over there.",
      how=["Write WHEN conditions from smallest to largest.", "End with ELSE and END AS band_name.",
           "GROUP BY the band."],
      tip="Number the band labels (1 Under…, 2…) so they sort correctly.",
      trap="Overlapping conditions in the wrong order, so bands swallow each other.",
      case="Finance wants orders grouped into value bands.",
      steps=["CASE WHEN total_price < 50000 THEN '1 Under ₦50k' …", "… ELSE '4 ₦1m+' END AS value_band",
             "GROUP BY value_band with COUNT and SUM."],
      table=[["value_band", "orders", "revenue"], ["1 Under ₦50k", "6,872", "₦115.8m"], ["2 ₦50k–250k", "2,712", "₦304.2m"],
             ["3 ₦250k–1m", "1,167", "₦596.7m"], ["4 ₦1m+", "249", "₦378.0m"]],
      so="Just 249 orders of ₦1m+ bring in more revenue (₦378m) than 6,872 small orders (₦116m).",
      split=4.20),

    S("3.2.15",
      what="A **subquery** is a query **inside** another query. Use it for two-step questions, like 'customers who "
           "spent more than the average customer': first work out the average, then compare each customer to it.",
      like="A subquery is like a teacher who first works out the class average, then lists pupils above it.",
      how=["Write the inner query and test it alone.", "Wrap it in brackets.", "Use it in WHERE, FROM or SELECT."],
      tip="Test the inner query on its own first.",
      trap="A subquery that returns several rows where one value is expected.",
      case="Marketing wants to know how many customers spent more than the average customer.",
      steps=["Inner: SUM(total_price) per customer → average spend.", "Outer: count customers whose spend > that average.",
             "Run and read."],
      table=[["Step", "Result"], ["Customers who ordered", "2,473"], ["Average spend per customer", "₦563,976"],
             ["Customers above average", "854"]],
      so="Only 854 of 2,473 customers (35%) beat the average: a few big spenders pull the average up.",
      split=4.90),

    S("3.2.16",
      what="Useful text functions: **UPPER, LOWER, TRIM, LENGTH, SUBSTR, REPLACE**. Useful date functions in "
           "SQLite: **date()** and **strftime('%Y', d)** for the year, **'%m'** for the month. Other databases use "
           "YEAR() or EXTRACT().",
      like="These functions are like kitchen tools: one peels (TRIM), one cuts (SUBSTR), one measures (LENGTH).",
      how=["Apply a function to a column in SELECT.", "Group by the result, e.g. the year.", "Give it an alias."],
      tip="Group by strftime('%Y', …) for yearly reports.",
      trap="Comparing dates stored in different formats.",
      case="Management wants orders and revenue by year (excluding cancelled orders).",
      steps=["strftime('%Y', order_date) AS order_year", "WHERE order_status <> 'Cancelled'",
             "GROUP BY order_year ORDER BY order_year."],
      table=[["order_year", "orders", "revenue"], ["2020", "286", "₦31,539,700"], ["2023", "1,985", "₦233,311,200"],
             ["2024", "2,638", "₦330,200,900"], ["2025", "3,160", "₦408,150,100"]],
      so="Revenue grew every year, reaching ₦408.2m in 2025 from ₦31.5m in 2020.",
      split=4.20),

    S("3.2.17",
      what="A full query follows a fixed order: **SELECT → FROM → JOIN → WHERE → GROUP BY → HAVING → ORDER BY "
           "→ LIMIT**. Read it as a sentence: show these columns, from these tables, for these rows, per group, "
           "sorted this way.",
      like="It's like a recipe: ingredients (FROM, JOIN), preparation (WHERE), cooking (GROUP BY), serving (ORDER BY).",
      how=["Write the business question as a -- comment.", "Build clause by clause, running as you go.",
           "Check the totals against a known figure."],
      tip="Run the query after each clause you add.",
      trap="Writing the whole query at once and not knowing which part is wrong.",
      case="Management asks: which customer zones drove 2025 net revenue, and how fast did we deliver?",
      steps=["FROM transactions JOIN customers, LEFT JOIN deliveries.",
             "WHERE 2025 and status Delivered/Completed; GROUP BY zone.",
             "SUM revenue, AVG days to deliver; ORDER BY revenue DESC."],
      table=[["customer_zone", "orders", "net_revenue", "avg days"], ["South-West", "781", "₦91,272,200", "6.1"],
             ["South-South", "691", "₦78,505,850", "6.1"], ["North-Central", "478", "₦71,038,500", "6.1"],
             ["North-East (lowest)", "201", "₦31,590,450", "6.0"]],
      so="South-West leads 2025 revenue; delivery speed is about 6 days in every zone.",
      split=4.20),
]
