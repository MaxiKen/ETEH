"""Week 3 edition-3 slide copy: Modules 3.1 (database fundamentals) and 3.2 (SQL and SQLite).
Every figure comes from running the SQL shown on supermart_online.db / mercyland_hospital.db (w3_fx.SQL, w3_fx.sandbox)."""
import sqlite3

import w3_data as D
from v3kit import T
from w3_fx import CU, DV, GROSS, I, N, NO_DELIV, PR, SE, SQL, TX, q, rows, sandbox, v

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
C985 = v("SELECT COUNT(*) FROM transactions WHERE customer_id = 'CUST00985'")
VIEW_COLS = len(q("SELECT * FROM v_order_lines LIMIT 1").columns)
LC_ADM = v("SELECT COUNT(*) FROM admissions", "lc")


def _plan(sql):
    p = q("EXPLAIN QUERY PLAN " + sql)
    return p.detail.iloc[-1].replace("USING COVERING INDEX", "USING INDEX")


PLAN_DATE = _plan("SELECT SUM(total_price) FROM transactions WHERE order_date BETWEEN '2024-01-01' AND '2024-01-31'")
PLAN_PAY = _plan("SELECT SUM(total_price) FROM transactions WHERE payment_method = 'USSD'")
JAN24 = q("SELECT COUNT(*) AS n, SUM(total_price) AS v FROM transactions WHERE order_date BETWEEN '2024-01-01' AND '2024-01-31'").iloc[0]


def _acid():
    src, mem = sqlite3.connect(D.DBS["smo"]), sqlite3.connect(":memory:", isolation_level=None)
    src.backup(mem)
    cnt = lambda: mem.execute("SELECT COUNT(*) FROM transactions").fetchone()[0]  # noqa: E731
    stock = lambda: mem.execute("SELECT stock_quantity FROM products WHERE product_id = 'PROD00001'").fetchone()[0]  # noqa: E731
    before = (cnt(), stock())
    mem.execute("BEGIN")
    mem.execute("INSERT INTO transactions (transaction_id, order_id, customer_id, product_id, seller_id, quantity, order_date) "
                "VALUES ('TRX-NEW', 'ORD-NEW', 'CUST00001', 'PROD00001', 'SELL0001', 2, '2025-12-31')")
    mem.execute("UPDATE products SET stock_quantity = stock_quantity - 2 WHERE product_id = 'PROD00001'")
    inside = (cnt(), stock())
    mem.execute("ROLLBACK")
    after = (cnt(), stock())
    return before, inside, after


ACID = _acid()


def _attach():
    c = sqlite3.connect(D.DBS["smo"])
    c.execute(f"ATTACH DATABASE '{D.DBS['lc']}' AS lh")
    a = c.execute("SELECT COUNT(*) FROM main.transactions").fetchone()[0]
    b = c.execute("SELECT COUNT(*) FROM lh.admissions").fetchone()[0]
    c.close()
    return a, b


ATT = _attach()
SIZES = {k: D.DBS[k].stat().st_size / 1e6 for k in D.DBS}

M31 = dict(
    id="3.1", title="Database Fundamentals",
    subtitle="Tables, keys, normalisation, types, indexes, views and transactions, and the two shapes a database can take.",
    objectives=["Explain what a database is and how it differs from a spreadsheet.",
                "Identify primary keys, foreign keys and the grain of a table.",
                "Recognise flat and normalised designs, and apply 1NF to 3NF.",
                "Describe data types, constraints, indexes, views, schemas and transactions.",
                "Tell a normalised (OLTP) schema from a star (OLAP) schema."],
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
          **SQL("""
SELECT COUNT(*) AS orphan_lines
FROM transactions t
LEFT JOIN products p
  ON p.product_id = t.product_id
WHERE p.product_id IS NULL;""", extra=[["Insert an order for product 'NOPE'"], [FK_ERR.replace("Error: ", "")]],
                widths=[1], split=0.52),
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

        T("3.1.6", "Normalisation: First Normal Form",
          defn="A table is in first normal form (1NF) when every cell holds a single value, there are no repeating groups of "
               "columns, and each row has a key.",
          simple="One box per answer: no 'Phone1, Phone2, Phone3' columns and no 'city, state' squeezed into one cell.",
          ex=dict(t="tDeliveries", cols=["delivery_id", "delivery_address", "city", "state"], n=3, size=11.5,
                  note=f"Showing 3 of {I(len(DV))} rows · city and state have their own columns, so you can filter by state"),
          watch="If 'Calabar, Cross River' sits in one cell, every filter by state needs text splitting first."),

        T("3.1.7", "Normalisation: Second and Third Normal Form",
          defn="Second normal form (2NF) requires every non-key column to depend on the whole key; third normal form (3NF) "
               "forbids a non-key column depending on another non-key column.",
          simple="Zone depends on state, not on the customer, so the strict design keeps state → zone once, in nigeria_states. "
                 "customers also keeps a copy of zone, so we check the copy against the lookup.",
          **SQL("""
SELECT COUNT(*) AS customers_zone_differs
FROM customers c
JOIN nigeria_states s ON s.state = c.state
WHERE s.zone <> c.zone;""", extra=[["States in the lookup: " + I(v("SELECT COUNT(*) FROM nigeria_states"))]],
                widths=[1], split=0.55),
          watch="A derived fact copied into several tables can drift apart; the lookup table is the single source of truth."),

        T("3.1.8", "Database Data Types",
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

        T("3.1.9", "Indexes and Query Performance",
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

        T("3.1.10", "Views",
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

        T("3.1.11", "Schemas and Namespaces",
          defn="A schema is a named container (namespace) for database objects, written schema.table, which keeps names "
               "apart and controls access.",
          simple="In SQLite the open file is the schema main. ATTACH adds another database under its own name, so one query "
                 "can read both.",
          like="Like surnames: two people called Ade are told apart as Ade Okafor and Ade Bello.",
          grid=dict(label="ATTACH DATABASE 'mercyland_hospital.db' AS lh", widths=[1.6, 1.8, 1], rows=[
              ["Reference", "Schema", "Rows"],
              ["main.transactions", "main (SuperMart Online)", I(ATT[0])],
              ["lh.admissions", "lh (Mercyland)", I(ATT[1])]]),
          watch="With several schemas attached, qualify table names: an unqualified name may not mean the table you think."),

        T("3.1.12", "Transactions and ACID",
          defn="A transaction is a group of changes that succeed or fail together. ACID: Atomic, Consistent, Isolated and Durable.",
          simple="A sale must add an order line and reduce stock together. BEGIN starts the group; COMMIT keeps it; ROLLBACK undoes all of it.",
          like="Like a bank transfer: the money leaves one account and arrives in the other, or neither happens.",
          grid=dict(label="New order on a copy of the database", widths=[1.8, 1.2, 1.4], rows=[
              ["Moment", "transactions", "PROD00001 stock"],
              ["Before BEGIN", I(ACID[0][0]), I(ACID[0][1])],
              ["Inside the transaction", I(ACID[1][0]), I(ACID[1][1])],
              ["After ROLLBACK", I(ACID[2][0]), I(ACID[2][1])]]),
          watch="Without a transaction, a crash between the two changes leaves a sale recorded with stock never reduced."),

        T("3.1.13", "Star Schema versus Normalised Schema",
          defn="A normalised schema splits data into many small related tables to avoid repetition; a star schema has one "
               "central fact table joined directly to dimension tables.",
          simple="Normalised suits recording orders; a star suits analysis, PivotTables and DAX, because every description is "
                 "one join from the facts.",
          like="Normalised is a filing cabinet of many folders; a star is hub-and-spoke buses, one ride from the centre.",
          grid=dict(label="supermart_online.db two ways", widths=[1.7, 2.3, 1.6], rows=[
              ["Shape", "Tables", "Best for"],
              ["Normalised (as stored)", f"{I(OBJ['table'])} tables; zone is 2 joins from an order", "Recording orders"],
              ["Star (for analysis)", "1 fact + customers, products, sellers, deliveries, DimDate", "Pivots and DAX"]]),
          watch="Chains such as transactions → customers → nigeria_states make filters slow and confusing in a model."),

        T("3.1.14", "Fact Tables and Dimension Tables",
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

        T("3.1.15", "OLTP versus OLAP",
          defn="Online transaction processing (OLTP) systems record many small, fast writes; online analytical processing "
               "(OLAP) systems answer large read-only questions over history.",
          simple="The marketplace app is OLTP: one order at a time, normalised. The analyst's Data Model is OLAP: years of "
                 "orders summarised, in a star.",
          like="OLTP is the till recording each sale; OLAP is the manager's report reading all the till rolls.",
          grid=dict(label="The same orders, two jobs", widths=[1.4, 2.2, 2.2], rows=[
              ["Aspect", "OLTP (marketplace app)", "OLAP (Data Model)"],
              ["Typical action", "Insert one order", "Sum six years of orders"],
              ["Data here", f"{I(len(TX))} order lines, {TX.order_date.min()[:4]}–{TX.order_date.max()[:4]}", "Totals by month, zone, category"],
              ["Shape", "Normalised", "Star"]]),
          watch="Heavy reports run on the live OLTP system slow down the people taking orders."),
    ],
    takeaways=["A database enforces types, keys and links; a spreadsheet does not.",
               "Primary keys identify one row; foreign keys must point at a real row.",
               "Normalise to store each fact once; check copies against the lookup.",
               "Indexes speed up filters and joins; views save queries; transactions keep changes together.",
               "Record in a normalised (OLTP) design; analyse in a star (OLAP)."],
)

# ---------------------------------------------------------------- 3.2
CNT = q("SELECT COUNT(*) AS n, SUM(total_price) AS v FROM transactions").iloc[0]
LAPTOPS = v("SELECT COUNT(*) FROM products WHERE product_name LIKE '%laptop%'")
BR_OK = v("""SELECT COUNT(*) FROM customers WHERE (membership_tier = 'Gold' OR membership_tier = 'Platinum') AND zone = 'South-West'""")
BR_BAD = v("""SELECT COUNT(*) FROM customers WHERE membership_tier = 'Gold' OR membership_tier = 'Platinum' AND zone = 'South-West'""")
FIRST_STORED = v("SELECT total_price FROM transactions LIMIT 1")
TOP1 = v("SELECT MAX(total_price) FROM transactions")
HAV_ALL = v("SELECT COUNT(DISTINCT seller_id) FROM transactions WHERE order_status IN ('Delivered','Completed')")
HAV_OK = v("""SELECT COUNT(*) FROM (SELECT seller_id FROM transactions WHERE order_status IN ('Delivered','Completed')
              GROUP BY seller_id HAVING SUM(total_price) > 15000000)""")
COMM_BAD = v("SELECT COUNT(*) FROM transactions WHERE ABS(total_price * commission_percentage / 100.0 - commission_amount) > 0.01")
SUB = q("""SELECT COUNT(*) AS n, AVG(spend) AS avg FROM (SELECT customer_id, SUM(total_price) AS spend FROM transactions
           GROUP BY customer_id)""").iloc[0]
SUB_ABOVE = v("""SELECT COUNT(*) FROM (SELECT customer_id, SUM(total_price) AS spend FROM transactions GROUP BY customer_id)
                 WHERE spend > (SELECT AVG(spend) FROM (SELECT SUM(total_price) AS spend FROM transactions GROUP BY customer_id))""")

M32 = dict(
    id="3.2", title="Introduction to SQL and SQLite",
    subtitle="Ask a database questions: SELECT, WHERE, ORDER BY, GROUP BY, HAVING, CASE, subqueries and functions.",
    objectives=["Explain what SQL is, what a dialect is and what SQLite is.",
                "Run queries in DB Browser for SQLite or the sqlite3 shell.",
                "Select, filter, sort and limit rows; combine conditions safely.",
                "Group and aggregate, filter groups with HAVING, and add calculated columns.",
                "Use DISTINCT, CASE, subqueries, text and date functions in a complete query."],
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

        T("3.2.2", "SQL Dialects",
          defn="A SQL dialect is one database product's version of SQL: the standard core plus its own functions and syntax.",
          simple="SELECT, WHERE, GROUP BY and JOIN are the same everywhere. Row limits, date functions and text joining differ.",
          like="Like English in Lagos, London and New York: the same language with a few different words.",
          grid=dict(label="SQLite versus SQL Server", widths=[1.6, 2.2, 1.8], size=11.5, rows=[
              ["Feature", "SQLite", "SQL Server (T-SQL)"],
              ["First 10 rows", "… LIMIT 10", "SELECT TOP 10 …"],
              ["Year of a date", "strftime('%Y', order_date)", "YEAR(order_date)"],
              ["Join two texts", "first_name || ' ' || last_name", "first_name + ' ' + last_name"],
              ["LIKE '%laptop%' matches 'Laptop'", f"Yes ({LAPTOPS} products)", "Depends on collation"]]),
          watch="A query copied between databases may run but give a different answer; test it on the target system."),

        T("3.2.3", "What Is SQLite?",
          defn="SQLite is a self-contained, serverless SQL database engine that stores a whole database in a single file.",
          simple="Nothing to install or log in to: copy the file and you have copied the database. It supports nearly all the "
                 "SQL an analyst needs.",
          like="Like a USB stick with a whole filing cabinet inside.",
          grid=dict(label="The Week 3 databases (folder Week3_Data_Files)", widths=[1.8, 0.8, 2], rows=[
              ["File", "Size", "Largest table"],
              ["supermart_online.db", f"{SIZES['smo']:.1f} MB", f"transactions: {I(len(TX))} rows"],
              ["mercyland_hospital.db", f"{SIZES['lc']:.1f} MB", f"admissions: {I(LC_ADM)} rows"]]),
          watch="SQLite has no user logins or permissions: anyone with the file can read all of it."),

        T("3.2.4", "Tools for Running SQLite Queries",
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

        T("3.2.5", "SELECT and FROM",
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

        T("3.2.6", "Filtering Rows with WHERE",
          defn="WHERE keeps only the rows for which a condition is true, using comparison operators, BETWEEN, IN, LIKE and IS NULL.",
          simple="Text goes in single quotes. A missing value is NULL, which you test with IS NULL; = NULL never matches anything.",
          like="Like a bouncer with a guest list: only rows that meet the rule get in.",
          **SQL("""
SELECT 'delivery_id IS NULL' AS test, COUNT(*) AS lines
FROM transactions WHERE delivery_id IS NULL
UNION ALL
SELECT 'delivery_id = NULL', COUNT(*)
FROM transactions WHERE delivery_id = NULL;""", widths=[1.6, 0.8], split=0.6),
          watch="= NULL returns no rows and no error, so the mistake is silent. Always write IS NULL or IS NOT NULL."),

        T("3.2.7", "Combining Conditions: AND, OR, NOT",
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

        T("3.2.8", "Sorting with ORDER BY",
          defn="ORDER BY sorts the result by one or more columns, ascending by default or descending with DESC.",
          simple="Sort by several columns in turn: ORDER BY category, price DESC. Without ORDER BY the row order is not guaranteed.",
          like="Like handing back exam scripts in score order, highest first.",
          **SQL("""
SELECT order_id, order_date, total_price
FROM transactions
ORDER BY total_price DESC
LIMIT 3;""", fmt={"total_price": "n"}, size=11.5, split=0.48),
          watch="A result that looked sorted last time is not sorted unless the query says ORDER BY."),

        T("3.2.9", "Limiting Rows with LIMIT",
          defn="LIMIT n returns at most n rows of the result; with ORDER BY it produces a top-n list.",
          simple="LIMIT without ORDER BY gives you some n rows, in storage order, not the top n.",
          like="'Show me the first 10 in the queue': if the queue isn't sorted, they aren't the best 10.",
          grid=dict(label="Same table, two queries", widths=[2.6, 1.8], rows=[
              ["Query", "First row returned"],
              ["SELECT total_price FROM transactions LIMIT 1", N(FIRST_STORED) + " (first stored row)"],
              ["… ORDER BY total_price DESC LIMIT 1", N(TOP1) + " (the true largest)"]]),
          watch="Never label a LIMIT result 'top 10' unless the query sorts first."),

        T("3.2.10", "Aggregating with GROUP BY",
          defn="GROUP BY collapses rows that share values into one row per group, and aggregate functions (COUNT, SUM, AVG, "
               "MIN, MAX) summarise each group.",
          simple="Every column in SELECT must be either in GROUP BY or inside an aggregate.",
          like="Like sorting receipts into piles by category, then writing the count and total on each pile.",
          **SQL("""
SELECT p.category,
       COUNT(*)           AS lines,
       SUM(t.total_price) AS revenue
FROM transactions t
JOIN products p ON p.product_id = t.product_id
WHERE t.order_status IN ('Delivered','Completed')
GROUP BY p.category
ORDER BY revenue DESC
LIMIT 4;""", fmt={"revenue": "n"}, size=11.5, split=0.56),
          watch="SQLite lets you select a column that is neither grouped nor aggregated and returns an arbitrary value; other databases refuse."),

        T("3.2.11", "Filtering Groups with HAVING",
          defn="HAVING filters groups after GROUP BY, using conditions on aggregates; WHERE filters rows before grouping.",
          simple="Sellers whose net revenue is above ₦15 million: the total exists only after grouping, so the test goes in HAVING.",
          like="WHERE checks tickets at the gate; HAVING picks the winners after the match.",
          **SQL("""
SELECT seller_id, SUM(total_price) AS net_rev
FROM transactions
WHERE order_status IN ('Delivered','Completed')
GROUP BY seller_id
HAVING SUM(total_price) > 15000000
ORDER BY net_rev DESC
LIMIT 3;""", fmt={"net_rev": "n"}, extra=[[f"{HAV_OK} of {HAV_ALL} sellers pass", ""]], size=11.5, split=0.58),
          watch="SUM(...) inside WHERE is an error: aggregates do not exist yet when WHERE runs."),

        T("3.2.12", "Aliases and Calculated Columns",
          defn="An alias (AS) renames a column or table in a query; a calculated column is an expression computed for every row.",
          simple="Recalculate commission from its parts and compare with the stored amount: a quick integrity check on every row.",
          like="Like a nickname: Oluwaseun becomes Seun, the same person.",
          **SQL("""
SELECT COUNT(*) AS lines_checked,
       SUM(ABS(total_price * commission_percentage
               / 100.0 - commission_amount) > 0.01)
         AS mismatches
FROM transactions;""", widths=[1, 1], split=0.6),
          watch="Divide by 100.0, not 100: in many databases integer ÷ integer drops the decimals."),

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

        T("3.2.14", "Conditional Logic with CASE",
          defn="CASE WHEN condition THEN value … ELSE value END returns a value from the first true condition, row by row.",
          simple="Use CASE to build bands and flags, then GROUP BY the band. Conditions are tested in order.",
          like="Like a post-office sorting desk: small parcels here, medium there, large over there.",
          **SQL("""
SELECT CASE WHEN total_price < 50000  THEN '1 <50k'
            WHEN total_price < 250000 THEN '2 50k-250k'
            WHEN total_price < 1000000 THEN '3 250k-1m'
            ELSE '4 1m+' END AS band,
       COUNT(*) AS lines, SUM(total_price) AS value
FROM transactions
GROUP BY band;""", fmt={"value": "m"}, size=11.5, split=0.62),
          watch="Order matters: put WHEN total_price < 1000000 first and every smaller order lands in that band."),

        T("3.2.15", "Subqueries",
          defn="A subquery is a SELECT nested inside another statement, whose result is used as a value, a list or a table.",
          simple="'Customers who spent more than the average customer' takes two steps: work out the average, then compare.",
          like="A teacher first works out the class average, then lists pupils above it.",
          **SQL("""
SELECT COUNT(*) AS above_average
FROM (SELECT customer_id, SUM(total_price) AS spend
      FROM transactions GROUP BY customer_id)
WHERE spend > (SELECT AVG(spend) FROM
      (SELECT SUM(total_price) AS spend
       FROM transactions GROUP BY customer_id));""",
                extra=[[f"of {I(SUB.n)} buyers; average {N(SUB.avg)}"]], widths=[1], split=0.62, code_size=11.5),
          watch="A subquery used as one value must return one row; if it returns several, the query fails."),

        T("3.2.16", "Useful SQL String and Date Functions",
          defn="Scalar functions transform one value per row: text (UPPER, TRIM, LENGTH, SUBSTR, REPLACE) and dates (date, "
               "strftime in SQLite).",
          simple="strftime('%Y', order_date) gives the year as text; '%m' gives the month. Other databases use YEAR() or EXTRACT().",
          **SQL("""
SELECT strftime('%Y', order_date) AS yr,
       COUNT(*)         AS lines,
       SUM(total_price) AS value
FROM transactions
GROUP BY yr
ORDER BY yr;""", fmt={"value": "n", "yr": "raw"}, size=11.5, split=0.5),
          watch="strftime returns text: compare it with '2025', not 2025, or every row fails the test (3.3.9)."),

        T("3.2.17", "Writing a Complete Analytical Query",
          defn="A complete query is written in the order SELECT, FROM, JOIN, WHERE, GROUP BY, HAVING, ORDER BY, LIMIT.",
          simple="Read it as a sentence: show these columns, from these tables, for these rows, per group, sorted this way. "
                 "Build it one clause at a time.",
          **SQL("""
SELECT customer_zone,
       COUNT(*)                   AS orders,
       SUM(total_price)           AS net_revenue,
       ROUND(AVG(days_to_deliver), 1) AS avg_days
FROM v_order_lines
WHERE order_status IN ('Delivered','Completed')
GROUP BY customer_zone
ORDER BY net_revenue DESC
LIMIT 4;""", fmt={"net_revenue": "n", "avg_days": "d1"}, size=11, split=0.56, code_size=11.5,
                head=["zone", "orders", "net_revenue", "avg_days"]),
          watch="Write the whole query at once and you won't know which clause is wrong. Run it after each added clause."),
    ],
    takeaways=["SQL describes what you want; SELECT only reads.",
               "Filter with WHERE (IS NULL for blanks); bracket every AND/OR mix.",
               "ORDER BY before LIMIT for a true top-n.",
               "GROUP BY + aggregates; HAVING filters the groups.",
               "Build a query clause by clause and check the count each time."],
)

MODULES = [M31, M32]
