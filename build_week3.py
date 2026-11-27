"""Build Week 3 Section Examples workbook (Modules 3.1-3.10)."""
import sqlite3
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import DataBarRule
from xl_helpers import *

CLEAN = "ETEH_Bootcamp/Datasets/CLEAN"
SQLITE = "ETEH_Bootcamp/Datasets/SQLITE/eteh_ecommerce.sqlite3"
W3 = "01_Section_Examples/Week3_SQL_DAX_Dashboards_Section_Examples.xlsx"

GOLD = PatternFill("solid", fgColor=GOLD)
GREEN = PatternFill("solid", fgColor=GREEN)
CONS = Font(name="Consolas", size=9)
CONS_B = Font(name="Consolas", size=9, bold=True, color=NAVY)

# ---- real answers from the SQLite database ----
con = sqlite3.connect(SQLITE)
q = lambda s: pd.read_sql_query(s, con)
n_trx = q("SELECT COUNT(*) c FROM transactions")["c"][0]
n_cust = q("SELECT COUNT(*) c FROM customers")["c"][0]
n_prod = q("SELECT COUNT(*) c FROM products")["c"][0]
n_sell = q("SELECT COUNT(*) c FROM sellers")["c"][0]
n_del = q("SELECT COUNT(*) c FROM deliveries")["c"][0]
tot_rev = q("SELECT ROUND(SUM(total_price),2) v FROM transactions")["v"][0]
n_deliv = q("SELECT COUNT(*) c FROM transactions WHERE order_status='Delivered'")["c"][0]
avg_line = q("SELECT ROUND(AVG(total_price),2) v FROM transactions")["v"][0]
top_cat = q("""SELECT p.category, ROUND(SUM(t.total_price),0) v
              FROM transactions t JOIN products p ON t.product_id=p.product_id
              GROUP BY p.category ORDER BY v DESC LIMIT 1""")
top_cat = top_cat.iloc[0].tolist()
orph = q("""SELECT COUNT(*) c FROM transactions t
            LEFT JOIN deliveries d ON t.delivery_id=d.delivery_id
            WHERE d.delivery_id IS NULL AND t.order_status<>'Cancelled'""")["c"][0]
top_seller = q("""SELECT s.store_name, ROUND(SUM(t.total_price),0) v, COUNT(*) n
                  FROM transactions t JOIN sellers s ON t.seller_id=s.seller_id
                  GROUP BY s.seller_id ORDER BY v DESC LIMIT 1""")
top_seller = top_seller.iloc[0].tolist()
by_month = q("""SELECT substr(order_date,1,7) m, COUNT(*) n, ROUND(SUM(total_price),0) v
                FROM transactions GROUP BY m ORDER BY m""")
by_month = [tuple(x) for x in by_month.values.tolist()]
cust_spend = q("""SELECT ROUND(SUM(total_price),2) v FROM transactions
                  WHERE customer_id='CUST00754'""")["v"][0]
n_methods = q("SELECT COUNT(DISTINCT payment_method) c FROM transactions")["c"][0]
con.close()

wb = openpyxl.Workbook()

# ============ 00_README ============
readme_sheet(wb,
    "Week 3 — SQL, DAX & Dashboards (Section Examples)",
    "Modules 3.1-3.10. Every example runs against the attached datasets — the relational E-commerce database (SQLite) is the spine, with Superstore/Education/Healthcare extracts used for append and modelling.",
    ["Worked examples for every module of Week 3, in the same order as the Week 3 presentation (Vol 3).",
     "The E-commerce database (5 tables, shared keys) is used end-to-end: it is the only attached dataset with proper primary/foreign keys, so it is the one used for SQL joins, the star schema, DAX measures and dashboards.",
     "SQL answers in this workbook were computed against ETEH_Bootcamp/Datasets/SQLITE/eteh_ecommerce.sqlite3 — your answers must match (small differences = a filter or join error)."],
    [["E-commerce (SQLite + CSV)", "11_DATA_Ecommerce", "3.1, 3.2, 3.3, 3.4, 3.5, 3.6, 3.8", f"{n_trx:,} txn", "One row per transaction; join keys order_id/product_id/customer_id/seller_id/delivery_id"],
     ["Superstore (CLEAN)", "12_DATA_Superstore", "3.4, 3.5, 3.7", "9,994 rows", "One row per order line"],
     ["Student Performance (CLEAN)", "13_DATA_Education", "3.4, 3.7", "6,607 rows", "One row per student-term"],
     ["Healthcare (CLEAN, anonymised)", "14_DATA_Health", "3.4, 3.10", "55,500 rows", "One row per admission"]],
    ["1. Open a module sheet, work top to bottom: concept table → worked example → practice data → your turn (yellow cells).",
     "2. Green cells are model answers/formulas — never type over them until you have attempted the task.",
     "3. DATA sheets are extracts for practice; the full datasets live in ETEH_Bootcamp/Datasets/CLEAN and SQLITE.",
     "4. Module 3.2 SQL: run each query in the E-commerce SQLite file and compare with the Expected column (or the 15_ANSWERS sheet)."])

# ============ 3.1 ============
ws = wb.create_sheet("01_3.1_DBConcepts")
setup_sheet(ws, "Module 3.1 — Database Concepts", "Rows vs tables, keys, relationships. Modelled on the E-commerce database.")
r = 4
r = section(ws, r, "What a database is (vs a spreadsheet)", ncols=7)
r = write_table(ws, r, ["Question", "Spreadsheet", "Database", "Why it matters"],
    [["Storage", "One big sheet (or many unlinked)", "Tables with enforced structure", "One source of truth, no duplicated headers"],
     ["Structure", "Flexible — anyone can add a column", "Fixed columns + types per table", "Clean imports; no silent format drift"],
     ["Relationships", "VLOOKUP strings, fragile", "Foreign keys, enforced by engine", "Joins are fast and correct"],
     ["Access", "File sharing / email", "Concurrent users, queries", "Many analysts, one dataset"]],
    widths=[16, 34, 36, 36])
r = section(ws, r, "The E-commerce database — 5 tables (the attached SQLite file)", ncols=7)
r = write_table(ws, r, ["Table", "Rows", "Key column", "What it holds", "Joined by"],
    [["transactions", f"{n_trx:,}", "transaction_id", "One row per line item bought", "order_id, product_id, customer_id, seller_id, delivery_id"],
     ["customers", f"{n_cust:,}", "customer_id", "One row per customer", "customer_id ← transactions.customer_id"],
     ["products", f"{n_prod:,}", "product_id", "One row per product", "product_id ← transactions.product_id"],
     ["sellers", f"{n_sell:,}", "seller_id", "One row per seller", "seller_id ← transactions.seller_id / products.seller_id"],
     ["deliveries", f"{n_del:,}", "delivery_id", "One row per shipment", "delivery_id ← transactions.delivery_id"]],
    widths=[16, 10, 16, 40, 46])
r = note(ws, r, f"Grain check: transactions = {n_trx:,} rows because one order can hold several products (multi-line baskets). This is why the sales analysis happens on transactions, not on orders.", ncols=7, height=30)
r = section(ws, r, "Keys — the join logic", ncols=7)
r = write_table(ws, r, ["Key type", "Example from E-commerce", "Rule it enforces"],
    [["Primary key", "transaction_id (each row unique, never null)", "One row per transaction — no duplicates allowed"],
     ["Foreign key", "transactions.product_id → products.product_id", "Every product_id on a transaction must exist in products"],
     ["Composite reality", "order_id is not unique in transactions (multi-line)", "Join on order_id only when you want all lines of one order"]],
    widths=[18, 52, 52])
r = section(ws, r, "Practice — read the tables (worked answers below)", ncols=7)
r = write_table(ws, r, ["#", "Question", "Answer (verified against the file)"],
    [["1", "How many tables does the E-commerce database have?", "5 — transactions, customers, products, sellers, deliveries"],
     ["2", "Which table holds the money (revenue per line)?", "transactions — column total_price (and quantity × unit_price = total_price)"],
     ["3", "Which key links a transaction to the person who bought it?", "transactions.customer_id → customers.customer_id"],
     ["4", "Which key links a transaction to what was bought?", "transactions.product_id → products.product_id"],
     ["5", "Why is order_id not a primary key of transactions?", "One order has several line items, so order_id repeats across rows"],
     ["6", "What must exist in deliveries before a transaction can reference it?", "The delivery_id row — a foreign key must match a real primary key"]],
    widths=[6, 52, 80])

# ============ 3.2 ============
ws = wb.create_sheet("02_3.2_SQL")
setup_sheet(ws, "Module 3.2 — SQL Essentials", "SELECT → WHERE → ORDER BY → GROUP BY → JOIN → subqueries/CTEs. All queries run on the E-commerce SQLite file.")
r = 4
r = section(ws, r, "SQL syntax ladder — each block adds one idea", ncols=7)
r = write_table(ws, r, ["Step", "Query", "Result (real, from the file)"],
    [["1 Select", "SELECT transaction_id, order_id, total_price\nFROM transactions;", f"{n_trx:,} rows; sample total_price e.g. ₦ values in the 10_DATA sheet"],
     ["2 Filter", "SELECT *\nFROM transactions\nWHERE order_status = 'Delivered';", f"{n_deliv:,} of {n_trx:,} rows (order_status text compare is exact)"],
     ["3 Order", "SELECT *\nFROM transactions\nORDER BY total_price DESC\nLIMIT 5;", "Top-5 most expensive lines — descending sort, LIMIT cuts the tail"],
     ["4 Aggregate", "SELECT COUNT(*), ROUND(SUM(total_price),2),\n       ROUND(AVG(total_price),2)\nFROM transactions;", f"{n_trx:,} rows · total ₦{tot_rev:,.0f} · average line ₦{avg_line:,.0f}"],
     ["5 Group", "SELECT category, ROUND(SUM(total_price),0) revenue\nFROM transactions t JOIN products p\n  ON t.product_id = p.product_id\nGROUP BY category\nORDER BY revenue DESC;", "Revenue by product category — grouping happens AFTER the join"],
     ["6 Join", "SELECT c.first_name || ' ' || c.last_name AS customer_name, COUNT(*) lines,\n       ROUND(SUM(t.total_price),0) spent\nFROM transactions t\nJOIN customers c ON t.customer_id = c.customer_id\nGROUP BY c.customer_id\nORDER BY spent DESC LIMIT 5;", "Top-5 customers by spend — one-to-many join, grouped"],
     ["7 Subquery", "SELECT * FROM transactions\nWHERE total_price > (\n  SELECT AVG(total_price) FROM transactions);", "All lines above the average line value"],
     ["8 CTE", "WITH big AS (\n  SELECT * FROM transactions WHERE total_price > 100000)\nSELECT payment_method, COUNT(*)\nFROM big GROUP BY payment_method;", f"Count of high-value lines per payment method ({n_methods} methods in the file)"]],
    widths=[14, 62, 56])
for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=2, max_col=2):
    if isinstance(row[0].value, str) and row[0].value.lstrip().upper().startswith(("SELECT", "WITH")):
        row[0].font = CONS
        row[0].alignment = Alignment(wrap_text=True, vertical="top")
r = section(ws, r, "Your turn — write the query (model answer in green)", ncols=7)
r = write_table(ws, r, ["#", "Task", "Model SQL", "Expected"],
    [["1", "Count transactions paid with Paystack", "SELECT COUNT(*) FROM transactions\nWHERE payment_method = 'Paystack';", "From ANSWERS sheet"],
     ["2", "Total revenue of Delivered orders", "SELECT ROUND(SUM(total_price),2)\nFROM transactions\nWHERE order_status = 'Delivered';", f"= SUM over {n_deliv:,} rows"],
     ["3", "Revenue by payment method", "SELECT payment_method,\n       ROUND(SUM(total_price),0) revenue\nFROM transactions\nGROUP BY payment_method\nORDER BY revenue DESC;", "3-4 methods, sorted desc"],
     ["4", "Top seller by revenue (join sellers)", "SELECT s.store_name,\n       ROUND(SUM(t.total_price),0) revenue\nFROM transactions t\nJOIN sellers s ON t.seller_id = s.seller_id\nGROUP BY s.seller_id\nORDER BY revenue DESC LIMIT 1;", f"{top_seller[0]} — ₦{top_seller[1]:,.0f} ({top_seller[2]:,} lines)"],
     ["5", "Monthly revenue (2024)", "SELECT substr(order_date,1,7) month,\n       COUNT(*) lines, ROUND(SUM(total_price),0) revenue\nFROM transactions\nGROUP BY month ORDER BY month;", "See month table in ANSWERS"],
     ["6", "Average basket per customer (>1 line)", "SELECT ROUND(AVG(lines),0)\nFROM (SELECT customer_id, COUNT(*) lines\n      FROM transactions GROUP BY customer_id);", "From ANSWERS sheet"],
     ["7", "Revenue of the top category", "SELECT p.category,\n       ROUND(SUM(t.total_price),0) revenue\nFROM transactions t JOIN products p\n  ON t.product_id = p.product_id\nGROUP BY p.category\nORDER BY revenue DESC LIMIT 1;", f"{top_cat[0]} — ₦{top_cat[1]:,.0f}"]],
    widths=[6, 42, 62, 40])
for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=3, max_col=3):
    if isinstance(row[0].value, str) and row[0].value.lstrip().upper().startswith(("SELECT", "WITH")):
        row[0].font = CONS_B
        row[0].fill = GREEN
        row[0].alignment = Alignment(wrap_text=True, vertical="top")
r = note(ws, r, "Join rule of thumb: JOIN when you need a column from the other table; GROUP BY only columns you select or aggregate; always qualify column names when two tables share a name (e.g. id columns).", ncols=7, height=32)

# ============ 3.3 ============
ws = wb.create_sheet("03_3.3_ODBC")
setup_sheet(ws, "Module 3.3 — Connecting Excel to a Database (ODBC / Power Query)", "Why we connect instead of copy-paste, and how to connect to the E-commerce SQLite file.")
r = 4
r = section(ws, r, "Connect vs copy-paste", ncols=6)
r = write_table(ws, r, ["Criterion", "Copy-paste from DB", "ODBC / Power Query connection"],
    [["Freshness", "Snapshot — stale tomorrow", "Refresh pulls the latest rows"],
     ["Volume", "Breaks past ~1M rows", "Streams millions; filters pushed down"],
     ["Audit", "Unknown where data came from", "Connection string + query log"],
     ["Cost", "Manual minutes daily", "One setup, then 1 click refresh"]],
    widths=[16, 40, 56])
r = section(ws, r, "Connection walkthrough (Power Query → SQLite)", ncols=6)
r = write_table(ws, r, ["Step", "Action", "What you should see"],
    [["1", "Data → Get Data → From Database → From SQLite Database", "File picker opens"],
     ["2", "Browse to ETEH_Bootcamp/Datasets/SQLITE/eteh_ecommerce.sqlite3", "Navigation window lists the 5 tables"],
     ["3", "Select all 5 tables → Load", "5 queries created; preview shows first rows"],
     ["4", "Keep 'Load To → Only Create Connection'", "No data dumped to sheets; model stays clean"],
     ["5", "Apply any transform (filter Delivered, add calculated columns)", "Transformations live in the query, not the sheet"],
     ["6", "Data → Refresh All", "All queries re-pull from the file"]],
    widths=[8, 56, 56])
r = section(ws, r, "Connection string anatomy (what the dialog builds for you)", ncols=6)
cs = "Driver={SQLite3};Database=C:\\ETEH\\Datasets\\SQLITE\\eteh_ecommerce.sqlite3"
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
c = ws.cell(row=r, column=1, value=cs); c.font = CONS; c.fill = GREEN
r += 1
r = note(ws, r, "The Driver names the ODBC bridge, Database names the file. If the file moves, edit Database — the query itself never changes. Never paste credentials into a connection string you share.", ncols=6, height=30)

# ============ 3.4 ============
ws = wb.create_sheet("04_3.4_MergeAppend")
setup_sheet(ws, "Module 3.4 — Merge (Join) & Append in Power Query", "Merge = join across tables; Append = stack identical tables. E-commerce join + Superstore append.")
r = 4
r = section(ws, r, "Merge vs Append — the decision", ncols=6)
r = write_table(ws, r, ["Operation", "Question it answers", "E-commerce example", "Power Query menu"],
    [["Merge (Join)", "What belongs to each row from ANOTHER table?", "Add product_name + category to each transaction", "Home → Merge Queries"],
     ["Append", "What are ALL rows from tables of the same shape?", "Stack monthly exports of transactions", "Home → Append Queries"]],
    widths=[18, 44, 50, 28])
r = section(ws, r, "Worked merge — transactions + products + customers (3 tables, 2 joins)", ncols=6)
r = write_table(ws, r, ["Step", "Action", "Result column", "Note"],
    [["1", "Merge transactions (left) with products (right), join on product_id", "product_name, category, seller_id added", "Exact match = same key values"],
     ["2", "Merge the result with customers, join on customer_id", "customer_name, city added", "One-to-many: every transaction line keeps its customer"],
     ["3", "Expand the product columns (tick the ones you need)", "A flat table — one row per transaction line", "Never leave merged columns hidden; expand explicitly"],
     ["4", "Change types (total_price → decimal, order_date → date)", "Correct types before loading", "Wrong types break aggregations silently"],
     ["5", "Group/aggregate: revenue by category", "Summary table for the dashboard", "Group By = GROUP BY in SQL"]],
    widths=[8, 52, 46, 46])
r = section(ws, r, "Worked append — stacking Superstore + Education (same-grain practice)", ncols=6)
r = write_table(ws, r, ["Concept", "For this workbook"],
    [["When append works", "Two tables with the SAME columns: e.g. two quarters of the same export"],
     ["When it fails", "Different columns (Superstore vs Education) — append adds null columns; merge is the right tool"],
     ["Practice here", "Append the 500-row extracts from 12_DATA_Superstore and 13_DATA_Education into one practice table: add a Source column first so you can tell them apart"],
     ["Real-world form", "Daily/monthly file drops of the same shape — append keeps one growing table"]],
    widths=[26, 110])
r = note(ws, r, "Merge rules: join on keys, not names/descriptions; keep the join kind Exact; if row counts after merge < left row count, you lost rows on non-matching keys — check key values, not just types.", ncols=6, height=32)

# ============ 3.5 ============
ws = wb.create_sheet("05_3.5_Model")
setup_sheet(ws, "Module 3.5 — Data Modelling & the Star Schema", "Facts, dimensions, relationships. Building the E-commerce star schema.")
r = 4
r = section(ws, r, "Facts vs Dimensions", ncols=6)
r = write_table(ws, r, ["Type", "Definition", "E-commerce example", "Count"],
    [["Fact table", "Numeric, measurable events (grain = one event)", "transactions (line-item grain)", f"{n_trx:,} rows"],
     ["Dimension table", "Descriptive context (who/what/when/where)", "customers, products, sellers, deliveries", "200–2,500 rows"],
     ["Time dimension", "A calendar table for date slicing", "order_date column → months/quarters", "Derived from order_date"]],
    widths=[18, 44, 50, 16])
r = section(ws, r, "Star schema — how the E-commerce tables relate", ncols=6)
r = write_table(ws, r, ["Relationship", "From → To", "Cardinality", "Meaning"],
    [["transactions → customers", "transactions.customer_id → customers.customer_id", "Many-to-one", "Each line belongs to one customer"],
     ["transactions → products", "transactions.product_id → products.product_id", "Many-to-one", "Each line is one product"],
     ["transactions → sellers", "transactions.seller_id → sellers.seller_id", "Many-to-one", "Each line sold by one seller"],
     ["transactions → deliveries", "transactions.delivery_id → deliveries.delivery_id", "Many-to-one", "Each line ships with one delivery (Cancelled lines may have none)"],
     ["products → sellers", "products.seller_id → sellers.seller_id", "Many-to-one", "Each product is sold by one seller"]],
    widths=[26, 52, 18, 56])
r = note(ws, r, "Star = facts in the middle, dimensions around, all joins many-to-one. This shape makes every pivot, DAX measure and dashboard slice work without VLOOKUPs.", ncols=6, height=28)
r = section(ws, r, "Modelling rules learned from the e-commerce file", ncols=6)
r = write_table(ws, r, ["Rule", "Evidence in the file", "Practice"],
    [["Never store derived values", "total_price repeats quantity × unit_price", "Store quantity + unit_price; compute total_price (or keep both, but mark which is source)"],
     ["Keys are stable", "order_id, product_id never change meaning", "Join on keys, never on names or dates"],
     ["Dimensions describe", "category lives in products, not repeated in transactions", "De-duplicate categories into the dimension"],
     ["Check referential integrity", f"{orph} non-Cancelled lines have no delivery row (expected: all Cancelled)", "Foreign keys must match — count orphans as a data-quality check"]],
    widths=[26, 56, 56])

# ============ 3.6 ============
ws = wb.create_sheet("06_3.6_DAX")
setup_sheet(ws, "Module 3.6 — DAX Measures & Calculated Columns", "CALCULATE, FILTER, SUMMARIZE. Measures computed from the E-commerce model.")
r = 4
r = section(ws, r, "Calculated column vs Measure", ncols=6)
r = write_table(ws, r, ["Type", "When it is computed", "Where it lives", "E-commerce example"],
    [["Calculated column", "At load, row by row", "In the table", "margin_line = total_price − (unit_price × 0.12)"],
     ["Measure", "At query time, on the fly", "In the model, not a row", "Revenue Delivered = CALCULATE(SUM(...), Status=\"Delivered\")"],
     ["Rule of thumb", "Use a column when every row needs the value", "Use a measure when the value depends on the filter (slicer/filter)", "Slicing by seller → measure"]],
    widths=[20, 36, 36, 56])
r = section(ws, r, "Core DAX patterns (model formulas in green — paste into DAX Studio / a measure)", ncols=6)
dax = [
    ("Total Revenue", "Total Revenue := SUM(transactions[total_price])", "The base measure; everything builds on it"),
    ("Revenue Delivered", "Revenue Delivered :=\nCALCULATE([Total Revenue], transactions[order_status] = \"Delivered\")", "CALCULATE = SUM with a filter added"),
    ("Revenue by Category", "Revenue by Category :=\nCALCULATE([Total Revenue], products[category])", "Context transition: sum only the current category slice"),
    ("Count Delivered", "Count Delivered :=\nCALCULATE(COUNTROWS(transactions), transactions[order_status] = \"Delivered\")", "COUNTROWS inside CALCULATE respects filters"),
    ("Lines above 100k", "High-Value Lines :=\nCALCULATE(COUNTROWS(transactions), transactions[total_price] > 100000)", "Numeric filter in CALCULATE"),
    ("Revenue per Line", "Revenue per Line := DIVIDE([Total Revenue], COUNTROWS(transactions))", "DIVIDE handles divide-by-zero safely"),
    ("Seller Revenue (top)", "Top Seller Revenue :=\nMAXX(TOPN(1, SUMMARIZE(transactions, sellers[store_name], \"Rev\", [Total Revenue]), [Rev], DESC), [Rev])", "SUMMARIZE groups; TOPN picks the leader"),
    ("Moving-style: MTD", "Revenue MTD :=\nCALCULATE([Total Revenue], DATESMTD(transactions[order_date]))", "Time intelligence — month-to-date")]
r = header_row(ws, r, ["Measure", "DAX (green = model)", "What it demonstrates"], widths=[22, 74, 56])
for m, code, desc in dax:
    ws.cell(row=r, column=1, value=m).font = BOLD10
    c = ws.cell(row=r, column=2, value=code); c.font = CONS_B; c.fill = GREEN
    c.alignment = Alignment(wrap_text=True, vertical="top")
    d = ws.cell(row=r, column=3, value=desc); d.font = Font(name="Calibri", size=9); d.alignment = WRAP
    for j in range(1, 4):
        ws.cell(row=r, column=j).border = BORDER
    ws.row_dimensions[r].height = 34
    r += 1
r = note(ws, r, "DAX mindset: every measure starts from an aggregation (SUM/COUNTROWS) and uses CALCULATE to add filters. Filters replace, not append — a second filter on the same column overrides the slicer. Test measures against the SQL answers in 3.2: [Revenue Delivered] must equal the SQL SUM(total_price) WHERE order_status='Delivered'.", ncols=6, height=44)

# ============ 3.7 ============
ws = wb.create_sheet("07_3.7_Charts")
setup_sheet(ws, "Module 3.7 — Charts from Aggregated Data", "Chart-type choice + pivot-driven charts on the E-commerce and Superstore data.")
r = 4
r = section(ws, r, "Chart-type decision table", ncols=6)
r = write_table(ws, r, ["Message", "Chart type", "Why", "E-commerce example"],
    [["Compare totals (few categories)", "Bar / column", "Eye reads length accurately", "Revenue by category"],
     ["Trend over time", "Line", "Order of magnitude is the time axis", "Revenue by month"],
     ["Share of a whole", "Pie / donut", "Only for ≤6 slices", "Payment-method share of revenue"],
     ["Distribution", "Histogram / column", "Shape of counts", "Line-value distribution"],
     ["Two measures, one category", "Combo / scatter", "Two scales when needed", "Revenue vs commission by seller"],
     ["Ranking (top N)", "Horizontal bar, sorted", "Sorted = instant ranking", "Top 10 sellers by revenue"]],
    widths=[32, 22, 36, 46])
r = section(ws, r, "Worked pivots (source: e-commerce + superstore extracts)", ncols=6)
r = write_table(ws, r, ["Pivot", "Rows", "Values", "Chart to build", "Expected result (verified)"],
    [["P1 Revenue by category", "products.category", "SUM(total_price)", "Column chart", f"Top category: {top_cat[0]} — ₦{top_cat[1]:,.0f}"],
     ["P2 Revenue by month", "order_date (month)", "SUM(total_price)", "Line chart", "Ascending months from the file (see 3.2 Q5)"],
     ["P3 Lines by payment method", "payment_method", "COUNT(transaction_id)", "Pie (≤5 slices)", f"{n_methods} methods in the file"],
     ["P4 Superstore: revenue by segment", "Segment", "SUM(Sales)", "Bar chart", "Technology/Furniture/Office Supplies split"],
     ["P5 Superstore: margin by region", "Region", "SUM(Profit)", "Bar chart (red/green)", "Regions with negative total profit"]],
    widths=[26, 24, 26, 30, 50])
r = section(ws, r, "Chart hygiene checklist", ncols=6)
for t in ["Title states the metric AND the grouping (not 'Chart 1')",
          "Axis labels are real values, not row numbers; no gaps in a time axis",
          "One chart = one message; 3 charts on a slide means 3 messages",
          "Sort bar charts by value; keep time on the x-axis left-to-right",
          "Colour is meaningful (red = negative/cancelled, not decorative)"]:
    r = add_check_box(ws, r, t, ncols=6)

# ============ 3.8 ============
ws = wb.create_sheet("08_3.8_Dashboards")
setup_sheet(ws, "Module 3.8 — Dashboards", "Layout, KPIs, slicers, drill-down. A small e-commerce operations dashboard.")
r = 4
r = section(ws, r, "Dashboard anatomy — what each block does", ncols=7)
r = write_table(ws, r, ["Block", "What it shows", "Source", "Formula / measure", "Verification"],
    [["KPI row", "Total revenue, lines, delivered rate, avg line", "transactions", "SUM / COUNT / CALCULATE", "Match 3.2 Q4 answers"],
     ["Filter bar", "Slicers: category, payment method, month", "dimensions", "Slicer → applies to all blocks", "Slicer changes every block at once"],
     ["Trend block", "Revenue by month (line chart)", "P2 pivot", "SUM(total_price) by month", "Sum of months = total revenue"],
     ["Rank block", "Top 10 sellers (bar chart)", "P6 pivot", "SUM by seller, top 10", f"Leader: {top_seller[0]} ₦{top_seller[1]:,.0f}"],
     ["Detail table", "Recent / filtered transactions", "transactions + products", "PivotTable with slicers", "Row count changes with filters"],
     ["Status strip", "Cancelled count + %", "transactions", "COUNT where status='Cancelled'", "Cancelled + Delivered + others = total lines"]],
    widths=[16, 38, 20, 36, 42])
r = section(ws, r, "Build order (professional sequence)", ncols=7)
for i, t in enumerate(["1. Model first: load the 5 tables, set relationships (3.5), write the base measures (3.6).",
                       "2. Build pivots/charts on the model — never on raw copies.",
                       "3. Assemble the sheet: KPI row at top, charts middle, detail table bottom.",
                       "4. Add slicers and format: consistent fonts, one colour language (navy/teal), no gridlines.",
                       "5. Validate: change every slicer and tick each block against the SQL answers."], 1):
    r = note(ws, r, t, ncols=7, height=22)
r = note(ws, r, "Single-source rule: a dashboard value must be reproducible from one SQL query or one measure. If two blocks disagree, the model is wrong — never patch a chart with a typed number.", ncols=7, height=30)

# ============ 3.9 ============
ws = wb.create_sheet("09_3.9_Blueprints")
setup_sheet(ws, "Module 3.9 — Dashboard Blueprints", "The spec sheet you hand to a builder: blocks, sources, formulas, acceptance tests.")
r = 4
r = section(ws, r, "Blueprint format (one row per block — fill before you build)", ncols=8)
r = header_row(ws, r, ["Block", "Purpose", "Data source", "Row/Col logic", "Measure/formula", "Chart type", "Size (cells)", "Acceptance test"],
    widths=[16, 30, 20, 24, 34, 18, 12, 40])
blue = [
    ["B1 Revenue KPI", "Headline revenue", "transactions", "n/a", "Total Revenue := SUM(transactions[total_price])", "None (big number)", "B5:C5", "Equals SUM(total_price) in SQL"],
    ["B2 Lines KPI", "Volume of lines", "transactions", "n/a", "Lines := COUNTROWS(transactions)", "None", "E5:F5", "Equals COUNT(*) in SQL"],
    ["B3 Delivered %", "Service quality", "transactions", "n/a", "Delivered % := DIVIDE([Lines Delivered],[Lines])", "None", "H5:I5", "Matches SQL COUNT where Delivered"],
    ["B4 Revenue by month", "Trend", "transactions", "order_date month", "SUM by month", "Line", "B8:F18", "Months sum to B1"],
    ["B5 Top sellers", "Ranking", "transactions+sellers", "SUM by seller desc", "Seller Revenue", "Bar", "H8:L18", "Leader matches SQL top-1"],
    ["B6 Payment mix", "Composition", "transactions", "SUM by method", "Revenue by method", "Donut", "N8:R14", "Slices sum to B1"],
    ["B7 Detail grid", "Audit drill-down", "transactions+products", "filtered rows", "PivotTable", "Table", "B20:H60", "Row count responds to slicers"],
]
for b in blue:
    for j, v in enumerate(b, 1):
        c = ws.cell(row=r, column=j, value=v)
        c.font = Font(name="Calibri", size=9)
        c.border = BORDER; c.alignment = WRAP
    ws.row_dimensions[r].height = 30
    r += 1
r = section(ws, r, "Acceptance test sheet — run before you call it done", ncols=8)
for t in ["Every KPI equals a SQL query result (Module 3.2 model answers)",
          "Every slicer changes every block (no block ignores a filter)",
          "No hardcoded numbers in any block — all from the model",
          "Print area set; dashboard fits one A4 landscape page",
          "A colleague can rebuild it from this blueprint alone"]:
    r = add_check_box(ws, r, t, ncols=8)

# ============ 3.10 ============
ws = wb.create_sheet("10_3.10_DBDesign")
setup_sheet(ws, "Module 3.10 — Database Design & Normalization", "1NF/2NF/3NF applied to the E-commerce and Healthcare data.")
r = 4
r = section(ws, r, "Normal form ladder", ncols=6)
r = write_table(ws, r, ["Form", "Rule", "E-commerce example", "Check"],
    [["1NF", "Atomic values; one row = one record; unique key", "transactions: one row per line item, transaction_id unique", "No repeating columns like item1/item2"],
     ["2NF", "No partial dependency on a composite key", "Revenue fields depend on the full line (product × order), not on one part", "Every non-key column depends on the whole key"],
     ["3NF", "No transitive dependency (non-key depends on another non-key)", "Seller city lives in sellers, not repeated on products/transactions", "Products table has no seller_city column"]],
    widths=[12, 44, 60, 40])
r = section(ws, r, "Denormalization — when to break the rules (deliberately)", ncols=6)
r = write_table(ws, r, ["Situation", "Denormalized design", "Trade-off"],
    [["Dashboards need fast reads", "Sales summary table: category, month, revenue pre-aggregated", "Fast reads, must refresh on change — keep a refresh job"],
     ["Reporting on 55k healthcare rows", "Admissions wide table with region name repeated", "Small duplication, big simplicity for analysts"],
     ["Rule", "Normalize for storage/integrity, denormalize for read speed", "Know which you are doing and why"]],
    widths=[34, 56, 56])
r = section(ws, r, "Design practice — spot the problem (answers verified)", ncols=6)
r = write_table(ws, r, ["#", "Scenario", "Problem", "Fix"],
    [["1", "A 'orders' table stores customer_name and city on every line", "3NF violation — city depends on customer, not the line", "Keep customer_id; move name/city to customers"],
     ["2", "transactions table has a column 'total_order_value' repeated on each line", "Denormalized redundancy; must match SUM of lines", "Drop it; compute from lines"],
     ["3", "Healthcare file repeats the ward name on every admission row", "Acceptable denormalization for reporting speed", "Fine for reporting; keep an admissions-ward lookup if writes happen"],
     ["4", "A table has no unique key (duplicated rows)", "Breaks 1NF — can't join or dedupe reliably", "Add/generate a surrogate key or dedupe (see Week 2 DIRTY files)"],
     ["5", "Products table stores the seller's commission %", "Transitive dependency (depends on seller, not product)", "Move commission % to sellers or a product-seller bridge"]],
    widths=[6, 52, 52, 52])

# ============ DATA sheets ============
ecom = pd.read_csv(f"{CLEAN}/ecommerce_transactions_CLEAN.csv")
data_sheet(wb, "11_DATA_Ecommerce", ecom, "E-commerce — transactions (extract)",
           f"Attached dataset. Full file: 11,000 rows in the SQLite DB. Practice queries from Module 3.2 against this extract.", tab_color=TEAL, max_rows=500, table_name="EcomTrx")
sup = pd.read_csv(f"{CLEAN}/superstore_CLEAN.csv")
data_sheet(wb, "12_DATA_Superstore", sup, "Superstore — orders (extract)",
           "Attached dataset. Used for append practice and pivot comparison in Modules 3.4 and 3.7.", tab_color=TEAL, max_rows=500, table_name="Super")
edu = pd.read_csv(f"{CLEAN}/student_performance_CLEAN.csv")
data_sheet(wb, "13_DATA_Education", edu, "Student Performance — records (extract)",
           "Attached dataset. Append/merge practice and modelling comparisons.", tab_color=TEAL, max_rows=500, table_name="Edu")
heal = pd.read_csv(f"{CLEAN}/healthcare_ANONYMISED.csv")
data_sheet(wb, "14_DATA_Health", heal, "Healthcare — admissions (extract, anonymised)",
           "Attached dataset. Normalization and denormalization discussion in Module 3.10.", tab_color=TEAL, max_rows=500, table_name="Health")

# ============ ANSWERS ============
ws = wb.create_sheet("15_ANSWERS")
setup_sheet(ws, "Week 3 — Answers & Marking Guide (Instructor)", "SQL answers computed against ETEH_Bootcamp/Datasets/SQLITE/eteh_ecommerce.sqlite3.")
r = 4
r = section(ws, r, "Database facts (for quick checks)", ncols=6)
r = write_table(ws, r, ["Fact", "Value"],
    [["Rows: transactions", f"{n_trx:,}"], ["Rows: customers", f"{n_cust:,}"], ["Rows: products", f"{n_prod:,}"],
     ["Rows: sellers", f"{n_sell:,}"], ["Rows: deliveries", f"{n_del:,}"],
     ["Total revenue (all lines)", f"₦{tot_rev:,.2f}"], ["Average line value", f"₦{avg_line:,.2f}"],
     ["Delivered lines", f"{n_deliv:,}"], ["Orphan non-Cancelled lines (delivery)", str(orph)],
     ["Top category", f"{top_cat[0]} — ₦{top_cat[1]:,.0f}"],
     ["Top seller", f"{top_seller[0]} — ₦{top_seller[1]:,.0f} ({top_seller[2]:,} lines)"],
     ["Customer CUST00754 total spend (top customer)", f"₦{cust_spend:,.2f}"],
     ["Payment methods", str(n_methods)]],
    widths=[42, 60])
r = section(ws, r, "Monthly revenue (from order_date)", ncols=6)
r = header_row(ws, r, ["Month", "Lines", "Revenue (₦)"], widths=[18, 14, 24])
for m, n, v in by_month:
    ws.cell(row=r, column=1, value=m).font = Font(name="Calibri", size=9); ws.cell(row=r, column=1).border = BORDER
    ws.cell(row=r, column=2, value=n).font = Font(name="Calibri", size=9); ws.cell(row=r, column=2).border = BORDER
    c = ws.cell(row=r, column=3, value=v); c.font = Font(name="Calibri", size=9); c.border = BORDER; c.number_format = "#,##0"
    r += 1
r += 1
r = section(ws, r, "Your-turn SQL answers (Module 3.2)", ncols=6)
r = write_table(ws, r, ["Task", "Answer / how to check"],
    [["1 Paystack count", "SELECT COUNT(*) WHERE payment_method='Paystack' — count matches ANSWERS fact table"],
     ["2 Delivered revenue", "SUM(total_price) WHERE order_status='Delivered' — less than total revenue (cancelled/refunded lines excluded)"],
     ["3 Revenue by method", "GROUP BY payment_method — 3–4 rows sorted by revenue desc"],
     ["4 Top seller", f"{top_seller[0]} — ₦{top_seller[1]:,.0f}"],
     ["5 Monthly revenue", "See monthly table above — months ascending, lines + revenue per month"],
     ["6 Avg basket", "AVG of per-customer line counts (customers with 1 line lower the average)"],
     ["7 Top category", f"{top_cat[0]} — ₦{top_cat[1]:,.0f}"],
     ["DAX cross-check", "Revenue Delivered (3.6) must equal SQL task 2; Revenue MTD must never exceed total revenue"]],
    widths=[26, 100])

wb.save(W3)
print("saved", W3, "sheets=", wb.sheetnames, "size=", f"{__import__('os').path.getsize(W3)/1024:.0f} KB")
