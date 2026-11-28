"""Week 3 Section Examples (Edition 3.1): questions in teaching order.

Progression rule for Week 3 (Weeks 1-2 are assumed)
    3.1        database words: tables, keys, why relational, views, star schema (no SQL yet)
    3.2        a little SQL in DB Browser: simple single queries, each answer checked by running it on supermart_online.db
    3.3        bringing the data into Excel (ODBC or CSV)
    3.4        Power Query merge (join types) and append on the SuperMart Online CSV files
    3.5        the Data Model and relationships
    3.6A/B     DAX measures; every number can also be checked with a Week 2 formula on the data sheets
    3.7 - 3.9  choosing charts, dashboard design, the three blueprints
    3.10       profiling a raw file extract before loading it
"""
from __future__ import annotations

import sqlite3

import pandas as pd

import w3_data as W
from common import CORE, STAGE, W2D, W3D

NET = '{"Delivered","Completed"}'


def build(b):
    db = W3D / "supermart_online.db"
    con = sqlite3.connect(db)

    def v(sql):
        return con.execute(sql).fetchone()[0]

    tx = W.smo("transactions").copy()
    tx["order_date"] = pd.to_datetime(tx.order_date)
    prod = W.smo("products")
    cust = W.smo("customers")

    M = {}

    def mod(code, sheet, title, learned):
        M[code] = b.module(code, sheet, title, learned)
        return M[code]

    m31 = mod("3.1", "3.1 Database Basics", "Database Fundamentals",
              "Databases, tables, rows and fields, primary and foreign keys, relational vs flat files, why relational "
              "databases cope as data grows, the SuperMart table map, data types, indexes, views, star schema, fact and "
              "dimension tables, OLTP vs OLAP.")
    m32 = mod("3.2", "3.2 SQL with SQLite", "Introduction to SQL and SQLite",
              "DB Browser for SQLite (Database Structure, Browse Data, Execute SQL) and simple queries: SELECT, FROM, "
              "WHERE, AND / OR / NOT, ORDER BY, LIMIT, GROUP BY, a two-table JOIN and DISTINCT.")
    m33 = mod("3.3", "3.3 ODBC Connections", "Connecting Excel to Databases with ODBC",
              "ODBC and drivers, DSN and DSN-less connections, credentials, importing a table vs writing SQL, refreshing, "
              "troubleshooting, the CSV export route and a first PivotTable on the imported data.")
    m34 = mod("3.4", "3.4 Merge & Append", "Merging and Appending in Power Query",
              "Merge vs append, the four join kinds (inner, left outer, full outer, left anti), expand vs aggregate, the "
              "Cartesian trap, appending tables, cleaning join keys and checking row counts.")
    m35 = mod("3.5", "3.5 Data Model & Star", "The Excel Data Model and Star Schema",
              "The Data Model, relationships and cardinality, fact and dimension tables, the date table, and why "
              "relationships beat VLOOKUP columns.")
    m36a = mod("3.6A", "3.6A DAX Foundations", "Power Pivot and DAX: Foundations and Context",
               "Power Pivot, calculated columns vs measures, DAX syntax, row and filter context, CALCULATE, ALL, VALUES, "
               "RELATED, SUMX, COUNTROWS, DISTINCTCOUNT, DIVIDE and IF / SWITCH.")
    m36b = mod("3.6B", "3.6B DAX Time & KPIs", "Power Pivot and DAX: Time Intelligence and KPIs",
               "TOTALYTD, SAMEPERIODLASTYEAR, % of total, BLANK vs zero, measure tables, format strings, implicit vs "
               "explicit measures, retail / hospital / university KPI measures and debugging.")
    m37 = mod("3.7", "3.7 Chart Types", "Chart Anatomy and the 17 Excel Chart Types",
              "What makes a chart effective, chart parts (axes, legend, labels, titles) and when to use each chart type.")
    m38 = mod("3.8", "3.8 Dashboard Build", "Dashboard Development",
              "Dashboard purpose and design principles, planning, layout, colour, KPI cards, slicers and timelines, "
              "interaction, printing, accessibility, performance and the QA checklist.")
    m39 = mod("3.9", "3.9 Blueprints", "Three Dashboard Blueprints",
              "The SuperMart retail, Kingsway progression and Mercyland operations dashboards.")
    m310 = mod("3.10", "3.10 Files to SQL & DAX", "File Extracts into SQL, DAX and Dashboards",
               "Profiling an extract, importing it into DB Browser as a staging table, simple SQL checks, cleaning in "
               "Power Query, DAX measures and reconciling a governed dashboard.")

    # ------------------------------------------------------------------ data sheets
    r_tx = b.data("Transactions", tx, "One SuperMart Online order line (transactions table of supermart_online.db)",
                  fmts={"order_date": "date", "unit_price": "naira", "total_price": "naira", "commission_amount": "naira",
                        "commission_percentage": "dec1"}, used_in="3.5, 3.6A, 3.6B")
    r_pr = b.data("Products", prod, "One product listed on SuperMart Online", fmts={"price": "naira", "discount_price": "naira"},
                  widths={"description": 40}, used_in="3.5, 3.6A")
    r_cu = b.data("Customers", cust, "One registered SuperMart Online customer", used_in="3.5")
    r_dd = b.data("DimDate", W.dimdate(), "One calendar day, 2020-2025 (the date table)", fmts={"Date": "date"},
                  used_in="3.5, 3.6B")

    # SQL examples file with the folder note made true for its new home
    sql_src = (W3D / "ETEH_Week3_SQL_Examples.sql").read_text(encoding="utf-8")
    sql_fixed = STAGE / "learn_files" / "ETEH_Week3_SQL_Examples.sql"
    sql_fixed.parent.mkdir(parents=True, exist_ok=True)
    sql_fixed.write_text(sql_src.replace("(folder Week3_Data_Files)", "(both databases are in this Ex3.2_Files folder)"), encoding="utf-8")
    f31 = m31.link_files([W3D / "supermart_online_01_schema.sql"])
    f32 = m32.link_files([db, W3D / "mercyland_hospital.db", sql_fixed])
    f34 = m34.link_files([CORE / f"SuperMart_Online_{t}.csv" for t in ("transactions", "products", "customers", "deliveries")])
    f310 = m310.link_files([W2D / "Mercyland_Episodes_2023_DIRTY.csv"])
    m35.set_refs([r_tx, r_pr, r_cu, r_dd])
    m36a.set_refs([r_tx, r_pr])
    m36b.set_refs([r_tx, r_dd])

    # ================================================================== 3.1
    m = m31
    m.part("Part A · Database words")
    m.mc("What is a primary key?", ["Any column with numbers", "A column whose value identifies each row uniquely",
                                    "The first column of every table", "A password for the database"], "B",
         "Each value appears once and is never blank, e.g. transaction_id.")
    m.mc("In the transactions table, customer_id points to a row in the customers table. What is customer_id in "
         "transactions?", ["A primary key", "A foreign key", "An index", "A view"], "B",
         "A foreign key links a row to the primary key of another table.")
    m.mc("Why does a relational database keep a customer's state on ONE row of customers, not on every order line?",
         ["It looks nicer", "Each fact is stored once, so a change is made in one place and cannot disagree",
          "SQL cannot read repeated text", "Order lines cannot hold text"], "B",
         "Store each fact once; order lines only keep customer_id and look the details up.")
    m.mc("What is a view?", ["A saved query that looks like a table", "A copy of the database file",
                             "A type of chart", "A backup"], "A", "supermart_online.db has a view called v_order_lines.")
    m.mc("What does an index do?", ["Deletes duplicates", "Makes lookups and filters on a column faster",
                                    "Changes data types", "Encrypts the table"], "B",
         "Like the index of a book: it finds rows without reading every page.")
    m.pick("transactions (one row per order line, with amounts)", ["Fact table", "Dimension table"], "Fact table",
           "Fact tables hold the events and the numbers you add up.")
    m.pick("customers (one row per customer, with name, state and tier)", ["Fact table", "Dimension table"],
           "Dimension table", "Dimension tables describe the who / what / where / when.")
    m.pick("A till system recording every sale as it happens", ["OLTP", "OLAP"], "OLTP",
           "OLTP = online transaction processing: many small writes.")
    m.pick("A model built for analysing sales by month, region and category", ["OLTP", "OLAP"], "OLAP",
           "OLAP = online analytical processing: reading and summarising.")
    m.yn("Can one Excel worksheet hold 5,000,000 order lines?", "No",
         "A worksheet stops at 1,048,576 rows; a database has no such limit, and Excel reaches it through a connection.")
    m.part("Part B · Read the schema file")
    m.do(f"Open {f31} › supermart_online_01_schema.sql in Notepad. Which column is the PRIMARY KEY of the sellers table?",
         "seller_id", "seller_id TEXT PRIMARY KEY.")
    m.num("In the same file, how many tables are created (count the CREATE TABLE statements)?", None, 6,
          "nigeria_states, customers, sellers, products, deliveries and transactions.", do=True)

    # ================================================================== 3.2
    m = m32

    def sql(text, q, fmt="int", how=None):
        val = v(q)
        if isinstance(val, str):
            m.short(text + "\n" + q, val, how or q)
        else:
            m.num(text + "\n" + q, None, val, (how + "  " if how else "") + "Query: " + q, fmt=fmt, do=True)

    m.example("running a query",
              [f"Open DB Browser for SQLite › Open Database › {f32} › supermart_online.db.",
               "Click the Execute SQL tab, type the query, then press Ctrl+Enter (or the ▶ button).",
               "Read the result in the grid below. Each question below gives the query to run (or to finish)."])
    m.part("Part A · SQL ideas")
    m.mc("What does SQL stand for?", ["Simple Question Language", "Structured Query Language", "Sorted Quick List",
                                      "System Quality Level"], "B", "SQL is the standard language for databases.")
    m.tf("SQLite keeps a whole database in one ordinary file (.db).", True, "No server is needed.")
    m.mc("In which order must these clauses be written?", ["FROM, SELECT, WHERE", "SELECT, FROM, WHERE, GROUP BY, ORDER BY",
                                                          "WHERE, FROM, SELECT", "ORDER BY, SELECT, FROM"], "B",
         "SELECT ... FROM ... WHERE ... GROUP BY ... ORDER BY ... LIMIT.")
    m.mc("Which DB Browser tab lists every table with its columns and data types?",
         ["Browse Data", "Database Structure", "Execute SQL", "Edit Pragmas"], "B",
         "Look at the structure first, then browse the rows, then query.")
    m.part("Part B · Run the queries")
    sql("How many order lines are in the transactions table?", "SELECT COUNT(*) FROM transactions;")
    sql("How many customers live in Lagos state?", "SELECT COUNT(*) FROM customers WHERE state = 'Lagos';")
    sql("How many order lines were Delivered AND paid with Paystack?",
        "SELECT COUNT(*) FROM transactions WHERE order_status = 'Delivered' AND payment_method = 'Paystack';")
    sql("What is the name of the most expensive product?",
        "SELECT product_name FROM products ORDER BY price DESC LIMIT 1;")
    sql("How many order lines were Cancelled? (GROUP BY order_status and read the Cancelled row)",
        "SELECT COUNT(*) FROM transactions WHERE order_status = 'Cancelled';")
    sql("How many different customers placed at least one order?", "SELECT COUNT(DISTINCT customer_id) FROM transactions;")
    sql("Which payment method has the most order lines? (GROUP BY, then the top row)",
        "SELECT payment_method, COUNT(*) AS lines FROM transactions GROUP BY payment_method ORDER BY lines DESC LIMIT 1;")
    sql("How many order lines have a total_price of ₦100,000 or more?",
        "SELECT COUNT(*) FROM transactions WHERE total_price >= 100000;")
    sql("How many order lines are for products in the Electronics category? (JOIN on product_id)",
        "SELECT COUNT(*) FROM transactions t JOIN products p ON p.product_id = t.product_id WHERE p.category = 'Electronics';")
    sql("How many order lines are dated on or after 1 January 2025? (dates are text: keep the quotes)",
        "SELECT COUNT(*) FROM transactions WHERE order_date >= '2025-01-01';")
    sql("What is the net revenue (Delivered + Completed lines) in naira?",
        "SELECT SUM(total_price) AS net_revenue FROM transactions WHERE order_status IN ('Delivered','Completed');",
        fmt="naira")

    # ================================================================== 3.3
    m = m33
    m.part("ODBC")
    m.mc("What is ODBC?", ["A chart type", "A standard way for programs such as Excel to connect to databases",
                           "A file format", "A DAX function"], "B", "Open Database Connectivity.")
    m.mc("What must be installed before Excel can connect to SQLite with ODBC?", ["A printer", "The SQLite ODBC driver",
                                                                                 "Power BI", "Python"], "B",
         "Each database type needs its own driver.")
    m.mc("A DSN-less connection stores the connection details ...", ["in the Windows ODBC settings",
                                                                     "inside the connection string in the workbook",
                                                                     "in a text file on the desktop", "nowhere"], "B",
         "A DSN stores them in Windows; DSN-less puts them in the connection string.")
    m.yn("Is it safe to type a database password into a workbook you will email to others?", "No",
         "Use Windows authentication or ask for credentials at refresh; never save passwords in shared files.")
    m.mc("Why write a SQL query in the connection instead of importing the whole table?",
         ["It is required", "Only the rows and columns you need come into Excel", "It adds more rows", "It removes keys"],
         "B", "Less data, faster refresh.")
    m.tf("A database query in Excel can be refreshed to bring in new rows.", True, "Data › Refresh All.")

    # ================================================================== 3.4
    m = m34
    J = ["Inner", "Left Outer", "Full Outer", "Left Anti"]
    m.part("Part A · Which join kind?")
    m.pick("Keep every transaction and add the product details where they match.", J, "Left Outer",
           "Left Outer keeps all rows of the first (left) table.")
    m.pick("Keep only rows that match in both tables.", J, "Inner", "Inner keeps matches only.")
    m.pick("Find customers who have NO transactions.", J, "Left Anti",
           "Left Anti keeps rows of the left table with no match on the right.")
    m.pick("Keep everything from both tables, matched where possible.", J, "Full Outer", "Full Outer keeps all rows of both.")
    m.mc("What is the difference between Merge and Append?",
         ["None", "Merge adds columns from another table (a join); Append adds rows (stacks tables)",
          "Append adds columns; Merge adds rows", "Merge deletes rows"], "B", "Merge = side by side; Append = on top.")
    m.mc("Before merging on Store name, you should ...", ["sort both tables", "TRIM and fix the case of the key in both tables",
                                                          "delete the key column", "add a chart"], "B",
         "'Lagos - Ikeja ' and 'Lagos - Ikeja' do not match until cleaned.")
    m.yn("If the key is repeated on BOTH sides of a merge, can the row count grow far beyond the original?", "Yes",
         "That is the Cartesian trap: always check the row count after a merge.")
    m.num("You append a table of 155 rows to a table of 107 rows (same columns). How many rows does the result have?",
          "=155+107", 262, "Append stacks the rows: 155 + 107 = 262.")
    m.part("Part B · Try it in Power Query")
    m.num(f"Load {f34} › SuperMart_Online_transactions.csv and SuperMart_Online_products.csv. Merge transactions with "
          "products on product_id (Left Outer). How many rows does the result have?", None, len(tx),
          "Each product_id appears once in products, so the row count stays 11,000.", do=True)
    m.num("Merge transactions with SuperMart_Online_deliveries.csv on delivery_id using Left Anti. How many transactions "
          "have no delivery?", None, int(v("SELECT COUNT(*) FROM transactions WHERE delivery_id IS NULL")),
          "The cancelled lines have no delivery_id.", do=True)
    m.num("Merge SuperMart_Online_customers.csv with transactions on customer_id using Left Anti. How many customers have "
          "never ordered?", None, int(v("SELECT COUNT(*) FROM customers WHERE customer_id NOT IN "
                                        "(SELECT customer_id FROM transactions)")),
          "Left Anti on customers keeps the customers with no matching transaction.", do=True)

    # ================================================================== 3.5
    m = m35
    m.part("The Data Model")
    m.mc("What is the Excel Data Model?", ["A chart template", "A store inside the workbook that holds several tables and the "
                                           "relationships between them", "A printed report", "A macro"], "B",
         "Power Pivot works on the Data Model.")
    m.pick("tProducts and tTransactions are related on product_id. Which table is on the 'one' side?",
           ["tProducts", "tTransactions"], "tProducts", "Each product appears once in tProducts and many times in tTransactions.")
    m.pick("What is the cardinality of the relationship tCustomers[customer_id] → tTransactions[customer_id]?",
           ["One-to-many", "One-to-one", "Many-to-many"], "One-to-many", "One customer, many order lines.")
    m.mc("Why does a model need a date table such as tDimDate?", ["To store passwords",
                                                                  "So time intelligence (YTD, last year) works on a complete calendar",
                                                                  "To hold product prices", "It is not needed"], "B",
         "Mark it as a date table: one row per day, no gaps.")
    m.yn("Is a relationship better than copying columns across with VLOOKUP?", "Yes",
         "Relationships keep one copy of each fact, stay small and update on refresh.")
    m.num("How many days does tDimDate hold? (=ROWS(tDimDate))", "=ROWS(tDimDate)", len(W.dimdate()),
          "2020-01-01 to 2025-12-31, one row per day.")

    # ================================================================== 3.6A
    m = m36a
    rev = float(tx.total_price.sum())
    net = float(tx.loc[tx.order_status.isin(["Delivered", "Completed"]), "total_price"].sum())
    m.example("measures and their Excel check",
              ["Power Pivot › Measures › New Measure. Type the DAX on the left.",
               "With no filters, the measure gives the same number as the Week 2 formula on the right.",
               "Put the measure in a PivotTable: it recalculates for every row and filter (filter context)."],
              [("Total Revenue := SUM(tTransactions[total_price])", "=SUM(tTransactions[total_price])",
                "=SUM(tTransactions[total_price])", rev, "naira"),
               ("Order Lines := COUNTROWS(tTransactions)", "=ROWS(tTransactions)", "=ROWS(tTransactions)", len(tx), "int")])
    m.part("Part A · DAX ideas")
    m.mc("What is the main difference between a calculated column and a measure?",
         ["None", "A column is worked out row by row and stored; a measure is worked out on demand for the current filters",
          "A measure is stored in every row", "Columns only work in charts"], "B",
         "Use measures for totals and ratios; columns for row-level labels.")
    m.mc("Which function changes the filter context of a calculation?", ["SUM", "CALCULATE", "RELATED", "DIVIDE"], "B",
         "CALCULATE(expression, filters...).")
    m.mc("Which function returns a value from the 'one' side of a relationship into a row of the 'many' side?",
         ["RELATED", "ALL", "COUNTROWS", "VALUES"], "A", "e.g. RELATED(tProducts[category]) in tTransactions.")
    m.mc("Why use DIVIDE(a, b) instead of a / b?", ["It is faster to type", "It returns BLANK instead of an error when b is 0",
                                                    "It rounds the answer", "It only works on dates"], "B",
         "DIVIDE handles division by zero safely.")
    m.mc("Which function removes filters from a table or column?", ["ALL", "RELATED", "SUMX", "DIVIDE"], "A",
         "Used for % of total: DIVIDE([Total Revenue], CALCULATE([Total Revenue], ALL(...))).")
    m.part("Part B · Work out each measure (check it with a formula)")
    m.num("Customers := DISTINCTCOUNT(tTransactions[customer_id]). What does it return with no filters?",
          "=COUNTA(UNIQUE(tTransactions[customer_id]))", int(tx.customer_id.nunique()),
          "Excel check: =COUNTA(UNIQUE(tTransactions[customer_id])).", dynamic=True)
    m.num(f"Net Revenue := CALCULATE([Total Revenue], tTransactions[order_status] IN {NET}). What does it return?",
          '=SUMIFS(tTransactions[total_price],tTransactions[order_status],"Delivered")'
          '+SUMIFS(tTransactions[total_price],tTransactions[order_status],"Completed")', net,
          "Excel check: two SUMIFS added together.", fmt="naira")
    m.num('Cancelled Lines := CALCULATE(COUNTROWS(tTransactions), tTransactions[order_status] = "Cancelled"). Value?',
          '=COUNTIFS(tTransactions[order_status],"Cancelled")', int((tx.order_status == "Cancelled").sum()),
          "Excel check: COUNTIFS.")
    m.num("Avg Line Value := DIVIDE([Total Revenue], [Order Lines]). What does it return? (nearest naira)",
          "=ROUND(SUM(tTransactions[total_price])/ROWS(tTransactions),0)", round(rev / len(tx)),
          "Total revenue ÷ 11,000 lines.", fmt="naira")

    # ================================================================== 3.6B
    m = m36b

    def net_between(a, b_):
        mask = tx.order_status.isin(["Delivered", "Completed"]) & (tx.order_date >= a) & (tx.order_date <= b_)
        return float(tx.loc[mask, "total_price"].sum())

    def sumifs_period(a, z):
        part = (f'tTransactions[order_date],">="&DATE({a}),tTransactions[order_date],"<="&DATE({z})')
        return (f'=SUMIFS(tTransactions[total_price],tTransactions[order_status],"Delivered",{part})'
                f'+SUMIFS(tTransactions[total_price],tTransactions[order_status],"Completed",{part})')

    n25 = net_between("2025-01-01", "2025-12-31")
    n24 = net_between("2024-01-01", "2024-12-31")
    q1 = net_between("2025-01-01", "2025-03-31")
    m.part("Part A · Time intelligence ideas")
    m.mc("What does TOTALYTD([Net Revenue], tDimDate[Date]) show for March 2025?",
         ["March only", "January to March 2025 added together", "The whole of 2024", "The average month"], "B",
         "Year-to-date = from 1 January up to the current date in the filter.")
    m.mc("What does SAMEPERIODLASTYEAR do?", ["Deletes last year", "Shifts the dates in the filter back one year",
                                               "Adds one year", "Counts the years"], "B",
         "Used for 'last year' comparisons.")
    m.mc("A PivotTable hides a seller whose [Net Revenue] is BLANK. What makes it show 0 instead?",
         ["[Net Revenue] + 0", "DELETE the seller", "SUMX", "A slicer"], "A",
         "Adding + 0 turns BLANK into 0; use it only where the zero matters.")
    m.mc("An explicit measure is ...", ["one you write yourself in DAX", "one Excel makes when you drag a field to Values",
                                         "a column", "a chart"], "A",
         "Explicit measures are named, reusable and documented; implicit ones are not.")
    m.part("Part B · Work out each value (check it with a formula)")
    m.num("Net Revenue for calendar year 2025?", sumifs_period("2025,1,1", "2025,12,31"), n25,
          "Delivered + Completed lines dated in 2025 (Excel check: SUMIFS with a date window).", fmt="naira")
    m.num("Net Revenue LY := CALCULATE([Net Revenue], SAMEPERIODLASTYEAR(tDimDate[Date])). Value for 2025 (i.e. 2024)?",
          sumifs_period("2024,1,1", "2024,12,31"), n24, "The same measure, one year earlier.", fmt="naira")
    m.num("Growth % = DIVIDE([Net Revenue] - [Net Revenue LY], [Net Revenue LY]). Value for 2025? (1 decimal place)",
          f"=({n25}-{n24})/{n24}", (n25 - n24) / n24, "(2025 − 2024) ÷ 2024.", fmt="pct")
    m.num("Net Revenue YTD := TOTALYTD([Net Revenue], tDimDate[Date]). Value at 31 March 2025?",
          sumifs_period("2025,1,1", "2025,3,31"), q1, "January to March 2025.", fmt="naira")

    # ================================================================== 3.7
    m = m37
    C = ["Line", "Clustered column", "Pie", "Scatter"]
    m.part("Which chart?")
    m.pick("Monthly sales over two years.", C, "Line", "Lines show trends over time.")
    m.pick("Sales of 12 stores compared side by side.", C, "Clustered column", "Columns compare categories.")
    m.pick("Share of 3 categories in one total (adds to 100%).", C, "Pie", "Only for a few parts of one whole.")
    m.pick("Study hours against final score for 5,840 students.", C, "Scatter", "Scatter shows the relationship between two numbers.")
    m.mc("Which chart shows how a starting value moves to an ending value through ups and downs?",
         ["Waterfall", "Radar", "Doughnut", "Treemap"], "A", "e.g. FY2024 margin → FY2025 margin, step by step.")
    m.mc("Which chart shows the spread of values in bins?", ["Histogram", "Pie", "Line", "Bubble"], "A",
         "A histogram counts values in ranges (bins).")
    m.mc("A good chart title ...", ["just says 'Chart 1'", "states the finding, e.g. 'Five stores made a loss in 2025'",
                                    "is always empty", "repeats the axis names"], "B", "Titles that state the finding.")
    m.yn("Should a column chart's value axis normally start at zero?", "Yes", "Otherwise small differences look huge.")

    # ================================================================== 3.8
    m = m38
    m.part("Dashboard design")
    m.mc("What should you decide FIRST when planning a dashboard?", ["The colours", "The audience, their questions and decisions",
                                                                   "The font", "How many charts fit"], "B",
         "Every visual must answer a question someone has.")
    m.mc("What is a KPI card?", ["A business card", "A big single number (with comparison) for one key measure",
                                 "A PivotTable", "A slicer"], "B", "e.g. Net Revenue ₦386.7m, +23.8% vs last year.")
    m.mc("How do you make one slicer filter several PivotTables?", ["Copy it", "Slicer › Report Connections",
                                                                  "Freeze Panes", "Group the charts"], "B",
         "Report Connections links the slicer to every pivot built on the same source.")
    m.mc("Which tool filters a dashboard by date with a sliding bar?", ["Timeline", "Sparkline", "Data bar", "Legend"], "A",
         "Insert › Timeline (needs a real date column).")
    m.yn("Should every colour on a dashboard mean something?", "Yes", "Use colour for meaning (e.g. red = below target), not decoration.")
    m.yn("Is it acceptable to rely on red and green alone to show good and bad?", "No",
         "Colour-blind users may not see the difference; add labels or icons.")
    m.tf("A dashboard QA checklist includes checking totals against the source.", True, "Reconcile before you publish.")

    # ================================================================== 3.9
    m = m39
    m.part("The three blueprints")
    m.pick("Net sales, margin and target attainment by store and month.", ["SuperMart", "Kingsway", "Mercyland"], "SuperMart",
           "The retail sales dashboard.")
    m.pick("Pass rates and at-risk students by faculty and level.", ["SuperMart", "Kingsway", "Mercyland"], "Kingsway",
           "The university progression dashboard.")
    m.pick("Bed occupancy, length of stay and readmissions by ward.", ["SuperMart", "Kingsway", "Mercyland"], "Mercyland",
           "The hospital operations dashboard.")
    m.yn("Do all three blueprints follow the same steps: question, model, measures, visuals, checks?", "Yes",
         "Only the data and the measures change.")

    # ================================================================== 3.10
    m = m310
    ep = pd.read_csv(W2D / "Mercyland_Episodes_2023_DIRTY.csv", dtype=str, keep_default_na=False)
    m.part("Profile the extract before loading")
    m.mc("Why profile a file extract before loading it?", ["To make it bigger",
                                                           "Problems found now are cheap; problems found on a dashboard are not",
                                                           "It is not needed", "To add charts"], "B",
         "Count rows, check columns, types, blanks and odd values first.")
    m.num(f"Open {f310} › Mercyland_Episodes_2023_DIRTY.csv with Data › From Text/CSV. How many rows (not counting the "
          "header)?", None, len(ep), f"{len(ep):,} rows.", do=True)
    m.num("How many columns does the file have?", None, len(ep.columns), ", ".join(ep.columns), do=True)
    m.mc("Patient names in the file look like 'yetuNde p. garBa'. Which fix is right?",
         ["Delete the names", "Use PROPER (Excel) or Text.Proper (Power Query)", "Sort them", "Leave them"], "B",
         "PROPER gives 'Yetunde P. Garba'.")
    m.mc("A staging table is ...", ["the final dashboard", "a raw copy of the extract loaded as-is before cleaning",
                                    "a chart", "a password file"], "B", "Load raw → check → clean → model.")
    m.yn("Should a governed dashboard show a reconciliation check (e.g. rows and totals vs the source)?", "Yes",
         "It proves the numbers are complete.")
    con.close()
