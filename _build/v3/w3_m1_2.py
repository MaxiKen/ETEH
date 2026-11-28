"""Week 3 edition-3 slide copy: Modules 3.1 (database fundamentals) and 3.2 (SQL and SQLite).
Every figure comes from running the SQL shown on supermart_online.db / mercyland_hospital.db (w3_fx.SQL, w3_fx.sandbox)."""
import w3_data as D
from v3kit import T
from w3_fx import CU, DV, GROSS, I, N, N_LINES, NO_DELIV, PR, SE, SQL, TX, q, rows, sandbox, v

# ---------------------------------------------------------------- facts computed once
OBJ = q("SELECT type, COUNT(*) AS n FROM sqlite_master WHERE name NOT LIKE 'sqlite_%' GROUP BY type").set_index("type").n
LAST_NAMES = CU.last_name.nunique()
PK_ERR, FK_ERR, CHK_ERR = sandbox([
    "INSERT INTO transactions SELECT 'TRX000001', 'ORD-TEST', customer_id, product_id, seller_id, delivery_id, quantity, "
    "unit_price, total_price, commission_percentage, commission_amount, order_status, payment_method, payment_status, "
    "order_date FROM transactions LIMIT 1",
    "INSERT INTO transactions (transaction_id, order_id, customer_id, product_id, seller_id, quantity, order_date) "
    "VALUES ('TRX-TEST', 'ORD-TEST', 'CUST00001', 'NOPE', 'SELL0001', 1, '2025-01-01')",
    "INSERT INTO transactions (transaction_id, order_id, customer_id, product_id, seller_id, quantity, order_date) "
    "VALUES ('TRX-TEST', 'ORD-TEST', 'CUST00001', 'PROD00001', 'SELL0001', 0, '2025-01-01')"])
FIRST_TRX = TX.transaction_id.min()
FK_LINE = q("SELECT t.order_id, t.product_id, p.product_name FROM transactions t JOIN products p ON p.product_id = t.product_id "
            "ORDER BY t.transaction_id LIMIT 1").iloc[0]
C985 = v("SELECT COUNT(*) FROM transactions WHERE customer_id = 'CUST00985'")
VIEW_COLS = len(q("SELECT * FROM v_order_lines LIMIT 1").columns)
LC_ADM = v("SELECT COUNT(*) FROM admissions", "lc")


def _plan(sql):
    p = q("EXPLAIN QUERY PLAN " + sql)
    return p.detail.iloc[-1].replace("USING COVERING INDEX", "USING INDEX")


PLAN_DATE = _plan("SELECT SUM(total_price) FROM transactions WHERE order_date BETWEEN '2024-01-01' AND '2024-01-31'")
PLAN_PAY = _plan("SELECT SUM(total_price) FROM transactions WHERE payment_method = 'USSD'")
JAN24 = q("SELECT COUNT(*) AS n, SUM(total_price) AS v FROM transactions WHERE order_date BETWEEN '2024-01-01' AND '2024-01-31'").iloc[0]


SIZES = {k: D.DBS[k].stat().st_size / 1e6 for k in D.DBS}

# ---------------------------------------------------------------- edition 3.1 additions (why relational, structure, browsing)
N_STATES = v("SELECT COUNT(*) FROM nigeria_states")
N_CU = len(CU)
SCALE_ROWS = [[I(n), I(n * 4), I(N_CU * 4 + n), "Yes" if n <= 1_048_575 else "No"]
              for n in (len(TX), 100_000, 1_000_000, 5_000_000)]
LAGOS_N = v("SELECT COUNT(*) FROM customers WHERE state = 'Lagos'")
USSD_N = v("SELECT COUNT(*) FROM transactions WHERE payment_method = 'USSD'")
CANC_N = v("SELECT COUNT(*) FROM transactions WHERE order_status = 'Cancelled'")
TOP_PAY = q("SELECT payment_method, COUNT(*) AS n FROM transactions GROUP BY payment_method ORDER BY n DESC LIMIT 1").payment_method[0]

M31 = dict(
    id="3.1", title="Database Fundamentals",
    subtitle="How a relational database is built and why: tables, keys and links, and why it copes as data grows.",
    objectives=["Explain what a database is and how it differs from a spreadsheet.",
                "Identify primary keys, foreign keys and what one row of a table means.",
                "Explain why relational (linked) tables beat one big sheet, especially as data grows.",
                "Read the table map of supermart_online.db: tables, keys and links.",
                "Describe data types, indexes and views, and tell a recording design from a star."],
    topics=[
        T("3.1.1", "What Is a Database?",
          defn="A database is an organised collection of related data, stored and managed by a database management system "
               "(DBMS) that enforces rules on what may be stored.",
          simple="A spreadsheet accepts anything you type. A database checks every entry against its rules: types, unique "
                 "IDs and valid links between tables, and refuses bad data.",
          like="A spreadsheet is a notebook you can write anywhere in; a database is a bank ledger where every entry must follow the rules.",
          grid=dict(label="supermart_online.db at a glance", widths=[1.6, 1.4, 2.6], rows=[
              ["Object", "Count", "Examples"],
              ["Tables", I(OBJ["table"]), "transactions, customers, products, sellers, deliveries, nigeria_states"],
              ["Views", I(OBJ["view"]), "v_order_lines"],
              ["Indexes", I(OBJ["index"]), "ix_trx_date, ix_trx_customer, ix_prod_cat"],
              ["Largest table", I(len(TX)), "transactions: one row per order line"]]),
          watch="A database is not a spreadsheet you edit freely: changes go through rules, and some are refused on purpose."),

        T("3.1.2", "Tables, Rows and Fields",
          defn="A table stores one kind of thing. Each row (record) is one instance and each column (field) holds one "
               "attribute, with a declared data type.",
          simple="It looks like a tidy Excel table, but every field has a type, so text cannot slip into a number column. "
                 "The grain is what one row means: here, one customer.",
          like="Like a class register: one line per pupil, the same boxes for every pupil.",
          ex=dict(t="tCustomers", cols=["customer_id", "first_name", "last_name", "state", "membership_tier"], n=4, size=11.5,
                  note=f"Showing 4 of {I(len(CU))} rows · {CU.shape[1]} fields · grain: one row per customer"),
          watch="Never mix grains in one table, for example customer rows and order rows together."),

        T("3.1.3", "Primary Keys",
          defn="A primary key is a column (or set of columns) whose value is unique and never empty in every row, so it "
               "identifies exactly one record.",
          simple="transaction_id is the key of transactions. The database checks the key on every insert and refuses a duplicate.",
          like="Like a National Identification Number: no two people share one.",
          **SQL("""
SELECT COUNT(*) AS rows_,
       COUNT(DISTINCT transaction_id) AS distinct_ids
FROM transactions;""", extra=[["Insert a repeat of " + FIRST_TRX, PK_ERR.replace("Error: ", "")]],
                widths=[1.6, 2.2], split=0.52),
          watch=f"Never use names as keys: the {I(len(CU))} customers share just {LAST_NAMES} last names."),

        T("3.1.4", "Foreign Keys",
          defn="A foreign key is a column whose values must match the primary key of another table, which protects "
               "referential integrity.",
          simple="transactions.product_id must exist in products.product_id. An order for a product that does not exist is refused.",
          like="Like a receipt that quotes a product code: the code must be in the shop's price list.",
          grid=dict(label="One order line and the product it points to", widths=[1.7, 2.6], size=11.5, rows=[
              ["Table · column", "Value"],
              ["transactions · order_id", FK_LINE.order_id],
              ["transactions · product_id (foreign key)", FK_LINE.product_id],
              ["products · product_id (primary key)", FK_LINE.product_id],
              ["products · product_name", FK_LINE.product_name],
              ["Insert an order for product 'NOPE'", FK_ERR.replace("Error: ", "Refused: ")]]),
          watch="SQLite checks foreign keys only after PRAGMA foreign_keys = ON; other databases check them always."),

        T("3.1.5", "Relational versus Flat Files",
          defn="A flat file stores all data in one table and repeats details on every row; a relational design stores each "
               "entity once in its own table and links tables with keys.",
          simple="The view v_order_lines is the flat version: customer name, tier, state and zone are copied onto every order line.",
          like="Flat is writing a customer's address on every receipt; relational is writing the customer number and keeping one address card.",
          grid=dict(label="Customer details: flat versus relational", widths=[2.6, 1.2, 1.2], rows=[
              ["Measure", "Flat view", "Relational"],
              ["Customer cells (4 details)", I(len(TX) * 4), I(len(CU) * 4)],
              ["Repeated copies", I(len(TX) * 4 - len(CU) * 4), "0"],
              ["Rows to edit if CUST00985 moves", I(C985), "1"]]),
          watch="In a flat file it is easy to change a repeated fact in some rows and miss the others."),

        T("3.1.6", "Why Relational Databases?",
          defn="A relational database stores each kind of thing (customers, products, orders) in its own table and links the "
               "tables with keys, instead of copying the same details onto every row.",
          simple="Each fact is written once, in one place. An order line keeps only customer_id; the customer's name, tier and "
                 "state live on one row of customers.",
          like="Like a school: one register of pupils, one list of subjects, and a results sheet that only writes the pupil number and subject code.",
          grid=dict(label="One big sheet versus linked tables", widths=[1.9, 2.0, 1.7], size=11.5, rows=[
              ["Question", "One big sheet", "Relational tables"],
              ["Where is a customer's state?", f"On every order line ({I(C985)} rows for CUST00985)", "One row in customers"],
              ["The customer moves state", f"Edit {I(C985)} rows; miss one and totals disagree", "Edit 1 row"],
              ["A wrong value is typed", "Accepted", "Refused by types, keys and links"],
              ["Several people at once", "One person edits the file", "Many users read and write"]]),
          watch="Relational does not make analysis harder: in 3.3 to 3.5 the linked tables come into Excel and work together."),

        T("3.1.7", "Relational Databases as Data Grows",
          defn="Scalability is how well a way of storing data copes as the rows, the users and the years grow.",
          simple="In a flat sheet every new order copies the customer's details again. In a relational design a new order adds "
                 "one short row that points to the customer.",
          like="Writing the customer's full address on every receipt is fine for ten receipts, not for a million.",
          grid=dict(label="Customer details stored as the business grows", widths=[1.3, 1.6, 1.7, 1.4], size=11.5, rows=[
              ["Order lines", "Flat: customer cells", "Relational: customer cells", "Fits one Excel sheet?"]] + SCALE_ROWS,
              note=f"Same {I(len(CU))} customers throughout; 4 details per customer (name, tier, state, zone); relational keeps one customer_id per line."),
          watch="An Excel sheet stops at 1,048,576 rows. A database has no such limit, and Excel can still analyse it through a connection (3.3) and the Data Model (3.5)."),

        T("3.1.8", "How the SuperMart Online Tables Link",
          defn="A database diagram (the schema) shows each table, what one row means, its primary key and the foreign keys "
               "that point to other tables.",
          simple="transactions sits in the middle: every order line points to one customer, one product, one seller and one delivery.",
          grid=dict(label="supermart_online.db: tables and links", widths=[1.3, 1.4, 1.3, 2.3, 0.8], size=11, rows=[
              ["Table", "One row is", "Primary key", "Points to (foreign keys)", "Rows"],
              ["transactions", "one order line", "transaction_id", "customers, products, sellers, deliveries", I(len(TX))],
              ["customers", "one customer", "customer_id", "nigeria_states (state)", I(len(CU))],
              ["products", "one product", "product_id", "sellers (seller_id)", I(len(PR))],
              ["sellers", "one store", "seller_id", "nigeria_states (store_state)", I(len(SE))],
              ["deliveries", "one shipment", "delivery_id", "customers, nigeria_states", I(len(DV))],
              ["nigeria_states", "one state", "state", "none", I(N_STATES)]]),
          watch="Open the Database Structure tab in DB Browser (3.2) and you will see exactly these tables and columns."),

        T("3.1.9", "Database Data Types",
          defn="A data type defines what values a column may hold (INTEGER, REAL, TEXT, DATE). Constraints such as "
               "NOT NULL, UNIQUE and CHECK add business rules.",
          simple="transactions declares quantity INTEGER CHECK (quantity > 0); sellers declares store_name UNIQUE, so no two stores share a name. SQLite stores dates as text.",
          **SQL("""
SELECT typeof(quantity)    AS quantity,
       typeof(total_price) AS total_price,
       typeof(order_date)  AS order_date
FROM transactions LIMIT 1;""", extra=[["Insert quantity = 0", CHK_ERR.replace("Error: CHECK constraint failed: ", "CHECK failed: "), ""]],
                widths=[1.6, 1.3, 1.1], split=0.5),
          watch="Money stored as text cannot be summed; give every column the type of the data it holds."),

        T("3.1.10", "Indexes and Query Performance",
          defn="An index is a sorted lookup structure on one or more columns that lets the DBMS find matching rows without "
               "reading the whole table.",
          simple="Index the columns you filter and join on most. EXPLAIN QUERY PLAN shows whether a query uses an index (SEARCH) "
                 "or reads every row (SCAN).",
          like="Like the index at the back of a textbook: go straight to the page.",
          grid=dict(label="EXPLAIN QUERY PLAN on transactions", widths=[1.6, 3.2], size=11.5, rows=[
              ["Filter", "Plan"],
              ["order_date (indexed)", PLAN_DATE],
              ["payment_method (no index)", PLAN_PAY],
              ["January 2024 result", f"{I(JAN24.n)} orders, {N(JAN24.v)}"]]),
          watch="Do not index every column: each index must be updated on every insert, which slows writing."),

        T("3.1.11", "Views",
          defn="A view is a named, stored query that can be queried like a table; it holds no data of its own and runs its "
               "query each time it is used.",
          simple=f"v_order_lines joins transactions, customers, products, sellers and deliveries into one {VIEW_COLS}-column "
                 "'table'. Analysts query it without writing the joins.",
          like="Like a shop window: a chosen arrangement of stock, without moving the stock.",
          **SQL("""
SELECT customer_zone, COUNT(*) AS lines
FROM v_order_lines
GROUP BY customer_zone
ORDER BY lines DESC;""", widths=[1.4, 0.9], split=0.55),
          watch="A view is not a copy: it always shows the current data in the tables underneath."),

        T("3.1.12", "Star Schema versus the Recording Design",
          defn="A recording design splits data into many small linked tables so each fact is stored once; a star schema has "
               "one central fact table joined directly to dimension tables.",
          simple="The recording design suits taking orders; a star suits analysis and PivotTables, because every "
                 "description is one step from the facts.",
          like="The recording design is a filing cabinet of many folders; a star is hub-and-spoke buses, one ride from the centre.",
          grid=dict(label="supermart_online.db two ways", widths=[1.7, 2.3, 1.6], rows=[
              ["Shape", "Tables", "Best for"],
              ["Recording design (as stored)", f"{I(OBJ['table'])} tables; zone is 2 steps from an order", "Recording orders"],
              ["Star (for analysis)", "1 fact + customers, products, sellers, deliveries, DimDate", "PivotTables"]]),
          watch="Chains such as transactions → customers → nigeria_states make filters slow and confusing in a model."),

        T("3.1.13", "Fact Tables and Dimension Tables",
          defn="A fact table records measurable events at a stated grain plus keys to dimensions; a dimension table "
               "describes the who, what, where and when used to filter and group facts.",
          simple="Facts are long and narrow and mostly numbers; dimensions are short and wide and mostly text.",
          like="Facts are till receipts; dimensions are the customer and product cards the receipts point to.",
          grid=dict(label="Tables in supermart_online.db", widths=[1.4, 1.1, 0.9, 2.4], rows=[
              ["Table", "Type", "Rows", "Holds"],
              ["transactions", "Fact", I(len(TX)), "quantity, total_price, commission_amount"],
              ["customers", "Dimension", I(len(CU)), "name, state, zone, membership_tier"],
              ["products", "Dimension", I(len(PR)), "category, subcategory, brand_name"],
              ["sellers", "Dimension", I(len(SE)), "store_name, store_state, rating"]]),
          watch="Do not sum numbers held in a dimension (such as customers.total_spent) alongside the fact rows."),

        T("3.1.14", "OLTP versus OLAP",
          defn="Online transaction processing (OLTP) systems record many small, fast writes; online analytical processing "
               "(OLAP) systems answer large read-only questions over history.",
          simple="The marketplace app is OLTP: one order at a time, many linked tables. The analyst's Data Model is OLAP: years of "
                 "orders summarised, in a star.",
          like="OLTP is the till recording each sale; OLAP is the manager's report reading all the till rolls.",
          grid=dict(label="The same orders, two jobs", widths=[1.4, 2.2, 2.2], rows=[
              ["Aspect", "OLTP (marketplace app)", "OLAP (Data Model)"],
              ["Typical action", "Insert one order", "Sum six years of orders"],
              ["Data here", f"{I(len(TX))} order lines, {TX.order_date.min()[:4]}–{TX.order_date.max()[:4]}", "Totals by month, zone, category"],
              ["Shape", "Many linked tables", "Star"]]),
          watch="Heavy reports run on the live OLTP system slow down the people taking orders."),
    ],
    takeaways=["A database enforces types, keys and links; a spreadsheet does not.",
               "Primary keys identify one row; foreign keys must point at a real row.",
               "Relational tables store each fact once, so they stay correct and fast as data grows.",
               "Indexes speed up filters; views save queries for re-use.",
               "Record in linked tables (OLTP); analyse in a star (OLAP)."],
)

# ---------------------------------------------------------------- 3.2
CNT = q("SELECT COUNT(*) AS n, SUM(total_price) AS v FROM transactions").iloc[0]
LIKE25 = v("SELECT COUNT(*) FROM transactions WHERE order_date LIKE '2025%'")
LAPTOPS = v("SELECT COUNT(*) FROM products WHERE product_name LIKE '%laptop%'")
BR_OK = v("""SELECT COUNT(*) FROM customers WHERE (membership_tier = 'Gold' OR membership_tier = 'Platinum') AND zone = 'South-West'""")
BR_BAD = v("""SELECT COUNT(*) FROM customers WHERE membership_tier = 'Gold' OR membership_tier = 'Platinum' AND zone = 'South-West'""")
FIRST_STORED = v("SELECT total_price FROM transactions LIMIT 1")
TOP1 = v("SELECT MAX(total_price) FROM transactions")
PAY_N = q("SELECT payment_method, COUNT(*) AS lines FROM transactions GROUP BY payment_method ORDER BY lines DESC")

M32 = dict(
    id="3.2", title="Introduction to SQL and SQLite",
    subtitle="A little time inside the database: explore it in DB Browser for SQLite and run simple one-line and multi-line queries.",
    objectives=["Explain what SQL and SQLite are, and why analysts use them.",
                "Find your way round DB Browser: Database Structure, Browse Data and Execute SQL.",
                "Write simple queries: SELECT, WHERE, AND/OR, ORDER BY and LIMIT.",
                "Count and total with COUNT, SUM and GROUP BY, and join two tables on a key.",
                "Check a SQL answer against Excel, where the analysis happens."],
    topics=[
        T("3.2.1", "What Is SQL?",
          defn="SQL (Structured Query Language) is the standard declarative language for defining, querying and changing "
               "data in relational databases.",
          simple="You say what you want, such as 'total value of all order lines', and the database works out how to get it. "
                 "Analysts mostly use SELECT, which only reads.",
          like="Like ordering at a restaurant: you name the dish; the kitchen decides how to cook it.",
          **SQL("""
SELECT COUNT(*)         AS order_lines,
       SUM(total_price) AS gross_value
FROM transactions;""", fmt={"gross_value": "n"}, widths=[1, 1.4], split=0.55),
          watch="SELECT never changes data. INSERT, UPDATE and DELETE do, so practise on a copy of the file."),

        T("3.2.2", "What Is SQLite?",
          defn="SQLite is a self-contained, serverless SQL database engine that stores a whole database in a single file.",
          simple="Nothing to install or log in to: copy the file and you have copied the database. It supports nearly all the "
                 "SQL an analyst needs.",
          like="Like a USB stick with a whole filing cabinet inside.",
          grid=dict(label="The Week 3 databases (01_Section_Examples › Week_3 › Ex3.2_Files)", widths=[1.8, 0.8, 2], rows=[
              ["File", "Size", "Largest table"],
              ["supermart_online.db", f"{SIZES['smo']:.1f} MB", f"transactions: {I(len(TX))} rows"],
              ["mercyland_hospital.db", f"{SIZES['lc']:.1f} MB", f"admissions: {I(LC_ADM)} rows"]]),
          watch="SQLite has no user logins or permissions: anyone with the file can read all of it."),

        T("3.2.3", "Tools for Running SQLite Queries",
          defn="A SQL client is a program that connects to a database, sends SQL and shows the results.",
          simple="DB Browser for SQLite is free and visual; the sqlite3 shell is fast and scriptable; Excel reads SQLite "
                 "through an ODBC driver (Module 3.3).",
          grid=dict(label="Three ways into the same file", widths=[1.8, 1.8, 1.6], rows=[
              ["Tool", "Strength", "Use it for"],
              ["DB Browser for SQLite", "Visual, beginner-friendly", "Learning, exploring"],
              ["sqlite3 command line", "Fast, scriptable", "Repeat jobs"],
              ["Excel via ODBC", "Results land in Excel", "Reports (3.3)"]]),
          steps=["DB Browser › Open Database (read only) › supermart_online.db", "Execute SQL tab › type the query",
                 "Press Ctrl+Enter (or F5) to run", "Read the result grid and the row count below it"],
          order=["steps", "grid"], split=0.48,
          watch="In the Browse Data tab, edits are saved when you click Write Changes; open the file read-only to explore."),

        T("3.2.4", "Look Around: the Database Structure Tab",
          defn="The Database Structure tab of DB Browser lists every table, view and index in the file, with each table's "
               "columns and data types.",
          simple="Expand a table to read its columns. The CREATE statement beside it shows the primary key, the links and the "
                 "rules, so you understand the design before you write any SQL.",
          steps=["Open Database Read Only › supermart_online.db", "Database Structure tab › expand Tables",
                 "Expand transactions: read each column and type", "Browse Data tab › Table: customers to see rows"],
          grid=dict(label="What you see for transactions", widths=[1.5, 0.9, 1.9], size=11.5, rows=[
              ["Column", "Type", "Meaning"],
              ["transaction_id", "TEXT", "Primary key: one per line"],
              ["customer_id", "TEXT", "Foreign key to customers"],
              ["quantity", "INTEGER", "Items; must be above 0"],
              ["total_price", "REAL", "Line value in naira"],
              ["order_date", "DATE", "Stored as text 'yyyy-mm-dd'"]]),
          order=["steps", "grid"], split=0.47,
          watch="Open the file read only while you explore; then nothing can be changed by accident."),

        T("3.2.5", "Browse Data: Look Before You Query",
          defn="The Browse Data tab shows a table's rows like a spreadsheet, with a filter box above every column and the row "
               "count at the bottom.",
          simple="Pick the customers table and type Lagos in the filter box above state: the count at the bottom changes. "
                 "The same question in SQL gives the same number.",
          grid=dict(label="Browse Data filter versus the same SQL", widths=[1.6, 2.5, 0.8], size=11.5, rows=[
              ["Filter typed in Browse Data", "Same question in Execute SQL", "Rows"],
              ["customers › state: Lagos", "SELECT COUNT(*) FROM customers WHERE state = 'Lagos';", I(LAGOS_N)],
              ["transactions › payment_method: USSD", "… FROM transactions WHERE payment_method = 'USSD';", I(USSD_N)],
              ["transactions › order_status: Cancelled", "… FROM transactions WHERE order_status = 'Cancelled';", I(CANC_N)]]),
          watch="Browse Data is for looking. For an answer you can repeat and share, write the query in Execute SQL."),

        T("3.2.6", "SELECT and FROM",
          defn="SELECT lists the columns to return; FROM names the table to read them from.",
          simple="SELECT * returns every column: fine for a first look, but saved queries should name their columns. Add "
                 "LIMIT to preview a few rows.",
          like="SELECT is the list of items you want; FROM is which store room to fetch them from.",
          **SQL("""
SELECT product_id, product_name,
       category, price
FROM products
LIMIT 4;""", fmt={"price": "n"}, size=11.5, split=0.42),
          watch=f"Running SELECT * on a big table without LIMIT pulls every row; products has {I(len(PR))}, admissions {I(LC_ADM)}."),

        T("3.2.7", "Filtering Rows with WHERE",
          defn="WHERE keeps only the rows for which a condition is true, using comparison operators, BETWEEN, IN, LIKE and IS NULL.",
          simple="Text goes in single quotes. A missing value is NULL, which you test with IS NULL; = NULL never matches anything.",
          like="Like a bouncer with a guest list: only rows that meet the rule get in.",
          **SQL("""
SELECT COUNT(*) AS ussd_lines
FROM transactions
WHERE payment_method = 'USSD';""", extra=[[f"Without a delivery (WHERE delivery_id IS NULL): {I(NO_DELIV)}"],
                                          [f"Dated in 2025 (WHERE order_date LIKE '2025%'): {I(LIKE25)}"]], widths=[1], split=0.55),
          watch="= NULL returns no rows and no error, so the mistake is silent. Always write IS NULL or IS NOT NULL."),

        T("3.2.8", "Combining Conditions: AND, OR, NOT",
          defn="AND requires both conditions, OR requires at least one, and NOT reverses a condition. AND is evaluated before OR.",
          simple="Question: Gold or Platinum customers in the South-West. Without brackets, SQL reads it as Gold anywhere, or "
                 "Platinum in the South-West.",
          like="Like 'tea or coffee with sugar': does sugar go with both? Brackets settle it.",
          **SQL("""
SELECT COUNT(*) AS customers
FROM customers
WHERE (membership_tier = 'Gold'
       OR membership_tier = 'Platinum')
  AND zone = 'South-West';""", extra=[["Without brackets: " + I(BR_BAD)]], widths=[1], split=0.56),
          watch=f"The bracket-free version runs without error and returns {I(BR_BAD)} instead of {I(BR_OK)}. Bracket every mix of AND and OR."),

        T("3.2.9", "Sorting with ORDER BY",
          defn="ORDER BY sorts the result by one or more columns, ascending by default or descending with DESC.",
          simple="Sort by several columns in turn: ORDER BY category, price DESC. Without ORDER BY the row order is not guaranteed.",
          like="Like handing back exam scripts in score order, highest first.",
          **SQL("""
SELECT order_id, order_date, total_price
FROM transactions
ORDER BY total_price DESC
LIMIT 3;""", fmt={"total_price": "n"}, size=11.5, split=0.48),
          watch="A result that looked sorted last time is not sorted unless the query says ORDER BY."),

        T("3.2.10", "Limiting Rows with LIMIT",
          defn="LIMIT n returns at most n rows of the result; with ORDER BY it produces a top-n list.",
          simple="LIMIT without ORDER BY gives you some n rows, in storage order, not the top n.",
          like="'Show me the first 10 in the queue': if the queue isn't sorted, they aren't the best 10.",
          grid=dict(label="Same table, two queries", widths=[2.6, 1.8], rows=[
              ["Query", "First row returned"],
              ["SELECT total_price FROM transactions LIMIT 1", N(FIRST_STORED) + " (first stored row)"],
              ["… ORDER BY total_price DESC LIMIT 1", N(TOP1) + " (the true largest)"]]),
          watch="Never label a LIMIT result 'top 10' unless the query sorts first."),

        T("3.2.11", "Aggregating with GROUP BY",
          defn="GROUP BY collapses rows that share values into one row per group, and aggregate functions (COUNT, SUM, AVG, "
               "MIN, MAX) summarise each group.",
          simple="One row per payment method, with its count and total; AS names a result column. It is the SQL "
                 "version of a PivotTable.",
          like="Like sorting receipts into piles by payment method, then writing the count and total on each pile.",
          **SQL("""
SELECT payment_method,
       COUNT(*)         AS lines,
       SUM(total_price) AS value
FROM transactions
GROUP BY payment_method
ORDER BY lines DESC
LIMIT 4;""", fmt={"value": "n"}, size=11.5, split=0.5),
          watch="Every column in SELECT must be in GROUP BY or inside COUNT, SUM, AVG, MIN or MAX; otherwise the result is not one row per group."),

        T("3.2.12", "Joining Two Tables on a Key",
          defn="A JOIN puts two tables side by side by matching a foreign key in one to the primary key in the other.",
          simple="transactions holds only product_id; products holds the name and category. JOIN … ON matches them, so each "
                 "order line can show its product. This is the relational design at work.",
          like="Like reading a receipt's product code and looking it up on the shop's price list.",
          **SQL("""
        SELECT t.order_id, p.product_name, p.category
        FROM transactions t
        JOIN products p ON p.product_id = t.product_id
        LIMIT 3;""", size=11, split=0.5, code_size=11.5),
          watch="Join on the key (product_id), never on a name. In Power Query the same idea is called Merge (3.4)."),

        T("3.2.13", "Removing Duplicates with DISTINCT",
          defn="SELECT DISTINCT returns each different combination of values once; COUNT(DISTINCT column) counts them.",
          simple="The fastest profile of a column: how many couriers, states or customers are there?",
          like="Like listing every team in a league once, however many matches each played.",
          **SQL("""
SELECT COUNT(*)                     AS deliveries,
       COUNT(DISTINCT courier_name) AS couriers,
       COUNT(DISTINCT state)        AS states
FROM deliveries;""", widths=[1, 1, 1], split=0.6),
          watch="Using DISTINCT to hide duplicate rows treats the symptom; find out why the duplicates appear."),

        T("3.2.14", "Writing a Multi-Line Query",
          defn="A multi-line query is still one question, written in the order SELECT, FROM, WHERE, GROUP BY, ORDER BY, LIMIT.",
          simple="Read it as a sentence: show these columns, from this table, for these rows, per group, sorted this way. "
                 "v_order_lines is the view from 3.1, so no join is needed.",
          **SQL("""
SELECT customer_zone,
       COUNT(*)         AS orders,
       SUM(total_price) AS net_revenue
FROM v_order_lines
WHERE order_status IN ('Delivered','Completed')
GROUP BY customer_zone
ORDER BY net_revenue DESC
LIMIT 4;""", fmt={"net_revenue": "n"}, size=11, split=0.56, code_size=11.5,
                head=["zone", "orders", "net_revenue"]),
          watch="Write the whole query at once and you won't know which clause is wrong. Run it after each added clause."),

        T("3.2.15", "From a Quick Query to Analysis in Excel",
          defn="In this course SQL is how you look inside a database and check what is there; the analysis itself is done in Excel.",
          simple="Answer a quick question in DB Browser, then bring the table into Excel (3.3) and use PivotTables, functions and "
                 "the Data Model (3.5 and 3.6). Both must give the same totals.",
          grid=dict(label="One question, two places, one answer", widths=[1.3, 2.1, 2.5, 0.8], size=10.5, rows=[
              ["Question", "In DB Browser (SQL)", "In Excel", "Answer"],
              ["Order lines in total", "SELECT COUNT(*) FROM transactions;", "=ROWS(tTransactions)", I(N_LINES)],
              ["Cancelled lines", "… WHERE order_status = 'Cancelled';", "=COUNTIFS(tTransactions[order_status],\"Cancelled\")", I(CANC_N)],
              ["Top payment method", "GROUP BY payment_method ORDER BY … LIMIT 1", "PivotTable, sort largest first", TOP_PAY]]),
          watch="Check the SQL count against the Excel row count after every import; if they differ, the import is incomplete."),

    ],
    takeaways=["Explore first: Database Structure, then Browse Data, then Execute SQL.",
               "SELECT only reads; WHERE filters rows (IS NULL for blanks); bracket AND/OR mixes.",
               "ORDER BY before LIMIT for a true top-n; GROUP BY for counts and totals per group.",
               "JOIN matches a foreign key to a primary key.",
               "SQL checks the data; the analysis happens in Excel, and both must agree."],
)

MODULES = [M31, M32]
