"""Week 3 Section Examples workbook (modules 3.1 - 3.10): databases, SQL, ODBC, joins, Data Model, DAX, charts, dashboards."""
from __future__ import annotations

import os
import sqlite3
import sys

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(__file__))
from common import SEC, W3D  # noqa: E402
from wk_common import DATASETS_TEXT, L, lc_adm, sm_sales, smo  # noqa: E402
from xlhelp import Book  # noqa: E402

OUT = SEC / "ETEH_Week3_Section_Examples.xlsx"
SQL_OUT = W3D / "ETEH_Week3_SQL_Examples.sql"
DB_SMO = W3D / "supermart_online.db"
DB_LC = W3D / "lifecare_hospital.db"

sm = sm_sales()
sm["Year"] = sm.Order_Date.dt.year
s25 = sm[sm.Year == 2025]
stores = L("sm_stores")
lc = lc_adm()
cv = L("cv_perf")
o = smo()
tx, dv, cu, pr, se = o["transactions"], o["deliveries"], o["customers"], o["products"], o["sellers"]
census = L("lc_census")
census["Census_Date"] = pd.to_datetime(census.Census_Date)

b = Book(str(OUT), "ETEH Week 3 - Section Examples", "Week 3 · Databases, SQL, Data Model, DAX and Dashboards")
F = b.F
SQL_LOG: list[str] = []


def q(sql, db=DB_SMO):
    with sqlite3.connect(db) as c:
        return pd.read_sql_query(sql, c)


def sqlblock(s, tag, title, sql, db=DB_SMO, n=12, name=None, fmts=None, note=None, widths=None):
    """Show a query and its real result (first n rows) as an Excel Table; log it to the .sql file."""
    res = q(sql, db)
    s.h2(f"{tag}  {title}")
    s.code(sql.strip())
    SQL_LOG.append(f"-- {tag} {title}\n-- database: {db.name}\n{sql.strip()};\n")
    shown = res.head(n).copy()
    shown.columns = [str(c_) for c_ in shown.columns]
    if name:
        s.table(shown, name, fmts=fmts or {}, widths=widths)
    else:
        s.frame(shown, fmts=fmts or {})
    s.para((f"Result: {len(res):,} row(s)" + (f", first {n} shown." if len(res) > n else ".")) + (" " + note if note else ""), "note")
    return res


# ======================================================================================
# 3.1 Database fundamentals
# ======================================================================================
s = b.sheet("3.1 Database Basics", "3.1  Database Fundamentals",
            "Tables, keys, normalisation, views, star vs normalised - on supermart_online.db", "3.1.1 - 3.1.15")
s.band("3.1.1-4", "The SuperMart Online database: tables, primary keys and foreign keys")
meta = [["customers", "customer_id", "state → nigeria_states", len(cu), "One customer"],
        ["sellers", "seller_id", "store_state → nigeria_states", len(se), "One marketplace seller"],
        ["products", "product_id", "seller_id → sellers", len(pr), "One product listing"],
        ["deliveries", "delivery_id (order_id UNIQUE)", "customer_id → customers", len(dv), "One shipment"],
        ["transactions", "transaction_id (order_id UNIQUE)", "customer_id, product_id, seller_id, delivery_id", len(tx), "One order line (fact)"],
        ["nigeria_states", "state", "-", 32, "Lookup: state → zone"]]
s.table(pd.DataFrame(meta, columns=["Table", "Primary key", "Foreign keys", "Rows", "Grain (one row =)"]), "tSchema",
        fmts={"Rows": "int"}, widths={"Table": 16, "Primary key": 26, "Foreign keys": 40, "Grain (one row =)": 22})
sqlblock(s, "3.1.3", "Primary key check - is every transaction_id unique?",
         "SELECT COUNT(*) AS rows_, COUNT(DISTINCT transaction_id) AS distinct_ids,\n       COUNT(*) - COUNT(DISTINCT transaction_id) AS duplicates\nFROM transactions", fmts={"rows_": "int", "distinct_ids": "int", "duplicates": "int"})
sqlblock(s, "3.1.4", "Foreign key check - orders pointing at a product that does not exist (orphans)",
         "SELECT COUNT(*) AS orphan_lines\nFROM transactions t\nLEFT JOIN products p ON p.product_id = t.product_id\nWHERE p.product_id IS NULL", note="0 = referential integrity holds.")
s.band("3.1.5-7", "Flat file versus relational: why the marketplace is split into tables")
flat_cells = 11000 * 30
s.calc([("Cells if everything sat in one flat sheet (v_order_lines: 11,000 x 30)", flat_cells, flat_cells, "int", "Customer name, tier, state repeated on every order"),
        ("Cells in the 5 normalised tables", "=" + "+".join(str(n * c_) for n, c_ in [(2500, 16), (200, 14), (800, 15), (10430, 13), (11000, 15)]),
         2500 * 16 + 200 * 14 + 800 * 15 + 10430 * 13 + 11000 * 15, "int", "Each fact stored once"),
        ("Saving", f"=1-({2500 * 16 + 200 * 14 + 800 * 15 + 10430 * 13 + 11000 * 15})/{flat_cells}",
         1 - (2500 * 16 + 200 * 14 + 800 * 15 + 10430 * 13 + 11000 * 15) / flat_cells, "pct", "And one update fixes a customer's address everywhere")])
s.grid([["1NF", "Atomic values, no repeating groups", "delivery_address '431 Road, Calabar' + separate city/state columns; never 'Phone1, Phone2' in one cell"],
        ["2NF", "Every non-key column depends on the whole key", "product_name lives in products, not in transactions"],
        ["3NF", "No column depends on another non-key column", "zone depends on state → kept in nigeria_states, looked up by state"]],
       header=["Form", "Rule", "In supermart_online.db"])
s.gap()
s.band("3.1.10-15", "Views, star schema, facts, dimensions, OLTP vs OLAP")
s.grid([["View", "v_order_lines joins transactions + customers + products + sellers + deliveries", "Analysts query one 'table'; logic lives in the database"],
        ["Fact table", "transactions (11,000 rows): quantity, total_price, commission_amount", "Numbers you SUM; foreign keys to dimensions"],
        ["Dimension tables", "customers, products, sellers, deliveries, a Date table", "Words you filter and group by"],
        ["Star schema", "Fact in the middle, each dimension one hop away", "Best shape for PivotTables, Power Pivot and DAX"],
        ["OLTP", "The live marketplace app writing orders one at a time", "Normalised, many small writes"],
        ["OLAP", "Your Data Model / dashboard reading millions of rows", "Star schema, big aggregate reads"]], header=["Concept", "In our data", "Why it matters"])

# ======================================================================================
# 3.2 SQL and SQLite
# ======================================================================================
s = b.sheet("3.2 SQL with SQLite", "3.2  Introduction to SQL and SQLite",
            "17 runnable queries on supermart_online.db - each shown with its real result", "3.2.1 - 3.2.17")
s.scenario("Open Week3_Data_Files/supermart_online.db in DB Browser for SQLite (free) or any SQLite tool, paste a query, press Run. "
           "Every query here is also saved in ETEH_Week3_SQL_Examples.sql. Results below were produced by running the exact SQL shown.")
sqlblock(s, "3.2.5", "SELECT and FROM - first rows of products", "SELECT product_id, product_name, category, price\nFROM products\nLIMIT 5", n=5, fmts={"price": "naira"}, widths=None)
sqlblock(s, "3.2.6", "WHERE - Electronics priced above ₦500,000",
         "SELECT product_id, product_name, brand_name, price\nFROM products\nWHERE category = 'Electronics' AND price > 500000\nORDER BY price DESC", n=8, fmts={"price": "naira"})
sqlblock(s, "3.2.7", "AND / OR / NOT - Lagos or Abuja customers who are not Bronze",
         "SELECT membership_tier, COUNT(*) AS customers\nFROM customers\nWHERE (state = 'Lagos' OR state = 'FCT') AND NOT membership_tier = 'Bronze'\nGROUP BY membership_tier", fmts={"customers": "int"})
sqlblock(s, "3.2.8-9", "ORDER BY + LIMIT - the 10 biggest orders",
         "SELECT order_id, order_date, product_id, quantity, total_price\nFROM transactions\nORDER BY total_price DESC\nLIMIT 10", n=10, fmts={"total_price": "naira", "quantity": "int"})
sqlblock(s, "3.2.10", "GROUP BY - revenue by category (completed and delivered orders only)",
         "SELECT p.category,\n       COUNT(*)            AS orders,\n       SUM(t.total_price)  AS revenue\nFROM transactions t\nJOIN products p ON p.product_id = t.product_id\nWHERE t.order_status IN ('Delivered', 'Completed')\nGROUP BY p.category\nORDER BY revenue DESC",
         name="tSqlCat", fmts={"orders": "int", "revenue": "naira"})
sqlblock(s, "3.2.11", "HAVING - sellers with more than ₦15m of delivered revenue",
         "SELECT s.store_name, COUNT(*) AS orders, SUM(t.total_price) AS revenue\nFROM transactions t\nJOIN sellers s ON s.seller_id = t.seller_id\nWHERE t.order_status = 'Delivered'\nGROUP BY s.store_name\nHAVING SUM(t.total_price) > 15000000\nORDER BY revenue DESC", n=10, fmts={"orders": "int", "revenue": "naira"},
         note="WHERE filters rows before grouping; HAVING filters the groups.")
sqlblock(s, "3.2.12", "Aliases and calculated columns - seller commission rate check",
         "SELECT order_id, total_price, commission_percentage AS comm_pct,\n       commission_amount,\n       ROUND(total_price * commission_percentage / 100.0, 2) AS recalculated,\n       commission_amount - ROUND(total_price * commission_percentage / 100.0, 2) AS diff\nFROM transactions\nLIMIT 6", n=6,
         fmts={"total_price": "naira", "commission_amount": "naira2", "recalculated": "naira2", "diff": "naira2"})
sqlblock(s, "3.2.13", "DISTINCT - which couriers deliver for us?", "SELECT DISTINCT courier_name\nFROM deliveries\nORDER BY courier_name", n=12)
sqlblock(s, "3.2.14", "CASE - order-value bands",
         "SELECT CASE WHEN total_price < 50000 THEN '1 Under ₦50k'\n            WHEN total_price < 250000 THEN '2 ₦50k-250k'\n            WHEN total_price < 1000000 THEN '3 ₦250k-1m'\n            ELSE '4 ₦1m+' END AS value_band,\n       COUNT(*) AS orders, SUM(total_price) AS revenue\nFROM transactions\nGROUP BY value_band\nORDER BY value_band",
         fmts={"orders": "int", "revenue": "naira"})
sqlblock(s, "3.2.15", "Subquery - customers who spent more than the average customer",
         "SELECT COUNT(*) AS above_average_customers\nFROM (SELECT customer_id, SUM(total_price) AS spend\n      FROM transactions GROUP BY customer_id) AS c\nWHERE spend > (SELECT AVG(spend) FROM\n                 (SELECT SUM(total_price) AS spend FROM transactions GROUP BY customer_id))",
         fmts={"above_average_customers": "int"})
sqlblock(s, "3.2.16", "String and date functions - revenue by year",
         "SELECT strftime('%Y', order_date) AS order_year,\n       COUNT(*) AS orders,\n       SUM(total_price) AS revenue,\n       ROUND(AVG(total_price), 0) AS avg_order\nFROM transactions\nWHERE order_status <> 'Cancelled'\nGROUP BY order_year\nORDER BY order_year",
         name="tSqlYear", fmts={"orders": "int", "revenue": "naira", "avg_order": "naira"})
sqlblock(s, "3.2.17", "A complete analytical query - 2025 zone scorecard",
         "-- Business question: which customer zones drove 2025 net revenue, and how fast did we deliver?\nSELECT c.zone                                         AS customer_zone,\n       COUNT(DISTINCT t.customer_id)                  AS customers,\n       COUNT(*)                                       AS orders,\n       SUM(t.total_price)                             AS net_revenue,\n       ROUND(SUM(t.total_price) * 1.0 / COUNT(*), 0)  AS avg_order_value,\n       ROUND(AVG(julianday(d.delivered_date) - julianday(t.order_date)), 1) AS avg_days_to_deliver\nFROM transactions t\nJOIN customers  c ON c.customer_id = t.customer_id\nLEFT JOIN deliveries d ON d.delivery_id = t.delivery_id\nWHERE t.order_date BETWEEN '2025-01-01' AND '2025-12-31'\n  AND t.order_status IN ('Delivered', 'Completed')\nGROUP BY c.zone\nORDER BY net_revenue DESC",
         name="tSqlZone25", fmts={"customers": "int", "orders": "int", "net_revenue": "naira", "avg_order_value": "naira", "avg_days_to_deliver": "dec1"})
sqlblock(s, "3.2 (hospital)", "Same skills on lifecare_hospital.db - ALOS and readmission by department",
         "SELECT department,\n       COUNT(*) AS admissions,\n       ROUND(AVG(length_of_stay), 1) AS alos_days,\n       ROUND(100.0 * SUM(readmit_30d = 'Yes') / COUNT(*), 1) AS readmit_pct\nFROM admissions\nGROUP BY department\nORDER BY readmit_pct DESC", db=DB_LC,
         fmts={"admissions": "int", "alos_days": "dec1", "readmit_pct": "dec1"})

# ======================================================================================
# 3.3 ODBC
# ======================================================================================
s = b.sheet("3.3 ODBC Connections", "3.3  Connecting Excel to Databases with ODBC",
            "Driver, DSN vs DSN-less, a parameterised query into Excel, refresh and troubleshooting", "3.3.1 - 3.3.9")
s.band("3.3.2-4", "Connect Excel to supermart_online.db")
s.grid([["1", "Install the SQLite3 ODBC driver (64-bit if your Office is 64-bit - File > Account > About Excel)"],
        ["2", "Data > Get Data > From Other Sources > From ODBC"],
        ["3", "DSN-less: choose (None) and paste the connection string below; or pick a DSN you created in 'ODBC Data Sources (64-bit)'"],
        ["4", "Advanced options > SQL statement: paste a query from sheet 3.2 - only the result comes into Excel"],
        ["5", "Load To... > Table, or Only Create Connection + Add to the Data Model (3.5)"]], header=["Step", "Action"], merge_last=True)
s.code("Driver={SQLite3 ODBC Driver};Database=C:\\ETEH\\01_Section_Examples\\Week3_Data_Files\\supermart_online.db;")
s.band("3.3.6", "Whole table versus SQL - how many rows travel to Excel?")
s.calc([("Import whole transactions table", len(tx), len(tx), "int", "Every row, then filter in Excel"),
        ("Import the 2025 zone scorecard query (3.2.17)", "=ROWS(tSqlZone25)", len(q("SELECT DISTINCT zone FROM customers")), "int", "The database does the work"),
        ("Rows saved", f"={len(tx)}-ROWS(tSqlZone25)", len(tx) - len(q("SELECT DISTINCT zone FROM customers")), "int", "")])
s.band("3.3.8", "Parameters in SQL - one yellow cell drives the query")
s.scenario("In Power Query, a parameter is spliced into the SQL. In Microsoft Query/ODBC the placeholder is '?'. The yellow cell holds the year; "
           "the formula builds the exact SQL text that will be sent - copy it into DB Browser to test.")
yr = s.inputs([("Order year parameter", 2024, "gen", "Try 2023 or 2025")])["Order year parameter"]
s.calc([("SQL that will be sent", f'="SELECT seller_id, SUM(total_price) AS revenue FROM transactions WHERE strftime(\'%Y\', order_date) = \'"&{yr}&"\' GROUP BY seller_id ORDER BY revenue DESC LIMIT 5"',
         "SELECT seller_id, SUM(total_price) AS revenue FROM transactions WHERE strftime('%Y', order_date) = '2024' GROUP BY seller_id ORDER BY revenue DESC LIMIT 5", None,
         "Never concatenate free text typed by users into SQL in production (SQL injection) - use proper parameters.")])
s.code('''// Power Query version with a parameter query named pYear
let
    Source = Odbc.Query("driver={SQLite3 ODBC Driver};database=" & pFolderPath & "supermart_online.db",
              "SELECT seller_id, SUM(total_price) AS revenue FROM transactions " &
              "WHERE strftime('%Y', order_date) = '" & Text.From(pYear) & "' GROUP BY seller_id")
in
    Source''')
s.band("3.3.5 / 3.3.9", "Credentials and troubleshooting")
s.grid([["Data source name not found and no default driver specified", "Driver name typed differently, or 32/64-bit mismatch", "Copy the exact name from ODBC Data Sources (64-bit)"],
        ["unable to open database file", "Path wrong or file on a disconnected drive", "Use the FolderPath parameter; check the .db exists"],
        ["Query returns 0 rows", "Date stored as text 'YYYY-MM-DD' but compared as a number", "Compare to strings: order_date >= '2025-01-01'"],
        ["Credentials prompt every refresh", "Privacy level / credentials not saved", "Data > Get Data > Data Source Settings > Edit Permissions"],
        ["Passwords in the connection string", "Visible to anyone with the file", "Use Windows authentication or the credential store, never typed passwords"]],
       header=["Error message", "Usual cause", "Fix"])

# ======================================================================================
# 3.4 Joins / Merge / Append
# ======================================================================================
s = b.sheet("3.4 Joins Merge Append", "3.4  Merging and Appending (Power Query joins = SQL joins)",
            "All six join kinds with real row counts, fuzzy matching, the Cartesian trap, append and cardinality", "3.4.1 - 3.4.16")
s.band("3.4.2-7", "Six join kinds: transactions (11,000 orders) × deliveries (10,430 shipments)")
s.scenario("570 orders were cancelled before shipping, so they have no delivery. How many rows each join returns depends entirely on the join kind - "
           "that is the whole lesson. Counts below come from running the SQL on supermart_online.db.")
joins = [
    ("Inner", "Only orders that have a delivery", "SELECT COUNT(*) FROM transactions t JOIN deliveries d ON d.delivery_id = t.delivery_id", "JoinKind.Inner"),
    ("Left outer", "All orders; delivery columns blank for the 570", "SELECT COUNT(*) FROM transactions t LEFT JOIN deliveries d ON d.delivery_id = t.delivery_id", "JoinKind.LeftOuter"),
    ("Right outer", "All deliveries; order columns blank if no order", "SELECT COUNT(*) FROM deliveries d LEFT JOIN transactions t ON t.delivery_id = d.delivery_id", "JoinKind.RightOuter"),
    ("Full outer", "Everything from both sides", "SELECT COUNT(*) FROM (SELECT t.order_id FROM transactions t LEFT JOIN deliveries d ON d.delivery_id = t.delivery_id UNION ALL SELECT d.order_id FROM deliveries d LEFT JOIN transactions t ON t.delivery_id = d.delivery_id WHERE t.transaction_id IS NULL)", "JoinKind.FullOuter"),
    ("Left anti", "Orders with NO delivery - the exception list", "SELECT COUNT(*) FROM transactions t LEFT JOIN deliveries d ON d.delivery_id = t.delivery_id WHERE d.delivery_id IS NULL", "JoinKind.LeftAnti"),
    ("Right anti", "Deliveries with NO order (should be 0)", "SELECT COUNT(*) FROM deliveries d LEFT JOIN transactions t ON t.delivery_id = d.delivery_id WHERE t.transaction_id IS NULL", "JoinKind.RightAnti")]
jr = []
for k, meaning, sql, pq in joins:
    n = int(q(sql).iloc[0, 0])
    jr.append([k, meaning, n, pq, sql])
    SQL_LOG.append(f"-- 3.4 {k} join row count\n{sql};\n")
jd = pd.DataFrame(jr, columns=["Join kind", "What you get", "Rows", "Power Query JoinKind", "SQL (row count)"])
tj = s.table(jd, "tJoins", fmts={"Rows": "int"}, widths={"Join kind": 12, "What you get": 36, "Power Query JoinKind": 20, "SQL (row count)": 60})
s.chart("bar", [{"name": "Rows", "categories": tj.ref("Join kind"), "values": tj.ref("Rows"), "labels": True, "label_fmt": "#,##0"}],
        "Rows returned by each join kind", legend=False, rows=15, size=(640, 280))
s.code('''// Power Query: Home > Merge Queries > choose key columns > Join Kind
Merged   = Table.NestedJoin(transactions, {"delivery_id"}, deliveries, {"delivery_id"}, "Delivery", JoinKind.LeftAnti),
NoShip   = Table.SelectColumns(Merged, {"order_id", "order_date", "order_status", "total_price"})''')
sqlblock(s, "3.4.6", "Left anti in practice - what happened to the orders that never shipped?",
         "SELECT t.order_status, t.payment_status, COUNT(*) AS orders, SUM(t.total_price) AS value\nFROM transactions t\nLEFT JOIN deliveries d ON d.delivery_id = t.delivery_id\nWHERE d.delivery_id IS NULL\nGROUP BY t.order_status, t.payment_status",
         fmts={"orders": "int", "value": "naira"}, note="All 570 are cancellations - the anti join proves there is no 'lost' shipment.")
sqlblock(s, "3.4.6b", "Left anti on customers - registered but never ordered",
         "SELECT c.membership_tier, COUNT(*) AS customers_without_orders\nFROM customers c\nLEFT JOIN transactions t ON t.customer_id = c.customer_id\nWHERE t.transaction_id IS NULL\nGROUP BY c.membership_tier", fmts={"customers_without_orders": "int"},
         note="27 customers - a re-activation campaign list for marketing.")
s.band("3.4.10", "The Cartesian product trap - joining on the wrong key")
s.calc([("Correct: join on delivery_id (one-to-one)", int(q("SELECT COUNT(*) FROM transactions t JOIN deliveries d ON d.delivery_id=t.delivery_id").iloc[0, 0]),
         int(q("SELECT COUNT(*) FROM transactions t JOIN deliveries d ON d.delivery_id=t.delivery_id").iloc[0, 0]), "int", "One row per shipped order"),
        ("Wrong: join on customer_id (many-to-many)", int(q("SELECT COUNT(*) FROM transactions t JOIN deliveries d ON d.customer_id=t.customer_id").iloc[0, 0]),
         int(q("SELECT COUNT(*) FROM transactions t JOIN deliveries d ON d.customer_id=t.customer_id").iloc[0, 0]), "int", "Each order matched to EVERY delivery of that customer"),
        ("Inflation factor", "=E" + str(s.r + 3) + "/E" + str(s.r + 2), 56725 / 10430, "dec2", "Revenue summed after this join would be overstated ~5x")])
SQL_LOG.append("-- 3.4.10 Cartesian trap: wrong key\nSELECT COUNT(*) FROM transactions t JOIN deliveries d ON d.customer_id = t.customer_id;\n")
s.band("3.4.8", "Fuzzy matching - seller names typed by hand in a supplier list")
sl = se.head(8)[["seller_id", "store_name"]].copy()
typed = [v.replace("Suppliers", "Supplier") if i % 3 == 0 else (v.upper() if i % 3 == 1 else v.replace(" ", "  ", 1)) for i, v in enumerate(sl.store_name)]
fz = pd.DataFrame({"Name_in_supplier_list": typed, "Official_store_name": sl.store_name.values, "seller_id": sl.seller_id.values})
s.table(fz, "tFuzzy", calc={"Exact_match": ("=EXACT([@Name_in_supplier_list],[@Official_store_name])", None, [a == b_ for a, b_ in zip(fz.Name_in_supplier_list, fz.Official_store_name)])},
        widths={"Name_in_supplier_list": 30, "Official_store_name": 30})
s.code('''// Power Query fuzzy merge: case-insensitive, ignore extra spaces, similarity 0.8
Fuzzy = Table.FuzzyNestedJoin(SupplierList, {"Name_in_supplier_list"}, sellers, {"store_name"}, "Match", JoinKind.LeftOuter,
                              [IgnoreCase = true, IgnoreSpace = true, Threshold = 0.8])''')
s.warn("Fuzzy matches are suggestions. Always review a sample - a 0.8 threshold can match two different sellers with similar names.")
s.band("3.4.11-13", "Appending: 12 monthly SuperMart files (Week 2) and a schema mismatch")
s.grid([["Append 2 tables", "Table.Combine({Jan, Feb})", "Rows add up: Jan + Feb"],
        ["Append 3+ tables", "Table.Combine({Jan, Feb, ... , Dec}) or From Folder", "3,312 rows = POS control total"],
        ["Schema mismatch", "December file exported with 'Sales_NGN' instead of 'Sales'", "Append creates TWO half-empty columns - rename before appending"],
        ["Different column order", "Same names, different order", "Harmless - Append matches by NAME, not position"]], header=["Situation", "M / action", "Result / fix"])
s.band("3.4.14-16", "Clean keys first, then verify cardinality")
sqlblock(s, "3.4.15", "Cardinality check - is order_id unique on BOTH sides before a one-to-one merge?",
         "SELECT 'transactions' AS table_name, COUNT(*) AS rows_, COUNT(DISTINCT order_id) AS distinct_keys FROM transactions\nUNION ALL\nSELECT 'deliveries', COUNT(*), COUNT(DISTINCT order_id) FROM deliveries", fmts={"rows_": "int", "distinct_keys": "int"},
         note="rows = distinct keys on both sides → one-to-one. If not, you have a one-to-many and must aggregate or expect more rows.")
s.para("Key cleaning recipe (Crestview padded IDs, Week 2): TRIM + UPPER both keys in Power Query before merging; count matched vs unmatched rows; record the counts in a Merge Log.", "plain")

# ======================================================================================
# 3.5 Data Model and star schema
# ======================================================================================
s = b.sheet("3.5 Data Model & Star", "3.5  The Excel Data Model and Star Schema",
            "Relationships and cardinality for SuperMart Online, plus a ready-made Date table", "3.5.1 - 3.5.7")
s.band("3.5.2", "Relationships to create (Data > Relationships, or Power Pivot Diagram View)")
rel = pd.DataFrame([["transactions", "customer_id", "customers", "customer_id", "Many-to-one (*:1)", "Single"],
                    ["transactions", "product_id", "products", "product_id", "Many-to-one (*:1)", "Single"],
                    ["transactions", "seller_id", "sellers", "seller_id", "Many-to-one (*:1)", "Single"],
                    ["transactions", "delivery_id", "deliveries", "delivery_id", "One-to-one (1:1) - 570 blanks on fact side", "Single"],
                    ["transactions", "order_date", "DimDate", "Date", "Many-to-one (*:1)", "Single"]],
                   columns=["Fact table", "Fact column", "Dimension table", "Dimension key", "Cardinality", "Cross-filter"])
s.table(rel, "tRelationships", widths={"Cardinality": 36, "Fact table": 14, "Dimension table": 15})
s.para("All five tables are on the D_ sheets of this workbook as Excel Tables. Select each > Power Pivot > Add to Data Model, then draw the relationships above.", "insight")
s.band("3.5.5", "The Date table - one row per day, marked as date table")
dates = pd.date_range("2020-01-01", "2025-12-31", freq="D")
dim = pd.DataFrame({"Date": dates, "Year": dates.year, "Quarter": "Q" + dates.quarter.astype(str), "Month_No": dates.month,
                    "Month": dates.strftime("%b"), "Year_Month": dates.strftime("%Y-%m"), "Weekday": dates.strftime("%a"),
                    "Is_Weekend": np.where(dates.day_of_week >= 5, "Yes", "No")})
s.calc([("Days in DimDate", "=ROWS(tDimDate)", len(dim), "int", "2020-01-01 to 2025-12-31 - covers every order_date"),
        ("Earliest order date", "=MIN(tTransactions[order_date])", (tx.order_date.min() - pd.Timestamp("1899-12-30")).days, "date", ""),
        ("Latest order date", "=MAX(tTransactions[order_date])", (tx.order_date.max() - pd.Timestamp("1899-12-30")).days, "date", ""),
        ("Gaps in the date table", "=MAX(tDimDate[Date])-MIN(tDimDate[Date])+1-ROWS(tDimDate)", 0, "int", "Must be 0 for time intelligence (3.6.22+)")])
s.para("Power Pivot > Design > Mark as Date Table > Date. Sort 'Month' by 'Month_No' (Sort by Column) so Jan..Dec appear in calendar order (3.5.6).", "note")
s.band("3.5.7", "Why relationships beat VLOOKUP columns")
s.grid([["VLOOKUP category into transactions", "11,000 extra cells; breaks if products sheet is re-sorted with approximate match", "Relationship: 0 extra cells, filters flow automatically"],
        ["VLOOKUP tier + state + zone too", "33,000 more cells, workbook slows", "One relationship to customers gives all 16 columns"],
        ["New products added", "Extend every VLOOKUP range", "Refresh - relationship just works"]], header=["Approach", "Cost", "Data Model alternative"])

# ======================================================================================
# 3.6 DAX (A: foundations, B: time intelligence & KPIs)
# ======================================================================================
net = tx[tx.order_status.isin(["Delivered", "Completed"])]
gross = tx.total_price.sum()
m = tx.merge(pr[["product_id", "category"]], on="product_id").merge(cu[["customer_id", "zone", "membership_tier"]], on="customer_id")
m["Year"] = m.order_date.dt.year
netm = m[m.order_status.isin(["Delivered", "Completed"])]
cat_rev = netm.groupby("category").total_price.sum()
s = b.sheet("3.6A DAX Foundations", "3.6  Power Pivot and DAX - Foundations and Context (3.6.1 - 3.6.21)",
            "Each measure: the DAX to type, what it means, and the answer you should get on this model", "3.6.1 - 3.6.21")
s.scenario("Build the model from sheet 3.5, then create these measures in a Measures table (3.6.31). The 'Expected result' is the grand total "
           "with no filters (or the filter stated) - if yours differs, check relationships and the order_status filter.")
dax_a = [
    ("3.6.4", "Gross Bookings", "Gross Bookings := SUM ( tTransactions[total_price] )", gross, "naira", "All 11,000 orders incl. cancelled/returned"),
    ("3.6.8", "Net Revenue", "Net Revenue := CALCULATE ( [Gross Bookings], tTransactions[order_status] IN { \"Delivered\", \"Completed\" } )", net.total_price.sum(), "naira", "CALCULATE changes the filter context"),
    ("3.6.17", "Orders", "Orders := COUNTROWS ( tTransactions )", len(tx), "int", ""),
    ("3.6.18", "Customers (distinct)", "Customers := DISTINCTCOUNT ( tTransactions[customer_id] )", tx.customer_id.nunique(), "int", "27 registered customers never ordered"),
    ("3.6.19", "Avg Order Value", "AOV := DIVIDE ( [Net Revenue], CALCULATE ( [Orders], tTransactions[order_status] IN { \"Delivered\", \"Completed\" } ) )", net.total_price.sum() / len(net), "naira", "DIVIDE returns BLANK instead of #DIV/0!"),
    ("3.6.15", "Revenue via SUMX", "Revenue X := SUMX ( tTransactions, tTransactions[quantity] * tTransactions[unit_price] )", float((tx.quantity * tx.unit_price).sum()), "naira", "Row-by-row iterator; equals Gross Bookings here"),
    ("3.6.16", "Avg items per order", "Avg Qty := AVERAGEX ( tTransactions, tTransactions[quantity] )", tx.quantity.mean(), "dec2", ""),
    ("3.6.9", "Net Revenue - Electronics", "Elec Rev := CALCULATE ( [Net Revenue], FILTER ( tProducts, tProducts[category] = \"Electronics\" ) )", cat_rev.get("Electronics", 0), "naira", "FILTER returns a table"),
    ("3.6.10", "% of all categories", "Cat Share := DIVIDE ( [Net Revenue], CALCULATE ( [Net Revenue], ALL ( tProducts[category] ) ) )", cat_rev.get("Electronics", 0) / cat_rev.sum(), "pct", "Shown for Electronics; ALL removes the category filter"),
    ("3.6.11", "Zone total ignoring tier", "Zone Rev := CALCULATE ( [Net Revenue], ALLEXCEPT ( tCustomers, tCustomers[zone] ) )", None, "naira", "In a Zone x Tier pivot, every tier cell shows its zone total"),
    ("3.6.12", "Categories sold", "Cats := COUNTROWS ( VALUES ( tProducts[category] ) )", pr.category.nunique(), "int", "VALUES respects filters; DISTINCT ignores blank row"),
    ("3.6.13", "Category (calc column on fact)", "Category = RELATED ( tProducts[category] )", None, None, "Calculated column in tTransactions - follows *:1 relationship"),
    ("3.6.14", "Orders per customer (calc column on customers)", "Order Count = COUNTROWS ( RELATEDTABLE ( tTransactions ) )", tx.groupby("customer_id").size().max(), "int", "Max value shown"),
    ("3.6.20", "Order size label", "Size := SWITCH ( TRUE (), [AOV] >= 250000, \"Large\", [AOV] >= 50000, \"Medium\", \"Small\" )",
     "Large" if net.total_price.mean() >= 250000 else ("Medium" if net.total_price.mean() >= 50000 else "Small"), None, "SWITCH(TRUE()) = readable IFS"),
    ("3.6.21", "Commission margin", "Comm % := VAR rev = [Gross Bookings] VAR comm = SUM ( tTransactions[commission_amount] ) RETURN DIVIDE ( comm, rev )",
     tx.commission_amount.sum() / gross, "pct", "VAR evaluates once - faster and readable"),
]
dd_ = pd.DataFrame([[a, b_, c, "" if d is None else d, f] for a, b_, c, d, e, f in dax_a], columns=["Topic", "Measure", "DAX", "Expected result", "Note"])
td = s.table(dd_, "tDaxA", widths={"Topic": 7, "Measure": 22, "DAX": 70, "Expected result": 16, "Note": 40})
for i, (_, _, _, val, kind, _) in enumerate(dax_a):
    if val is not None and kind:
        s.ws.write(td.r0 + 1 + i, td.ci("Expected result"), float(val) if not isinstance(val, str) else val, b.nf(kind, style="plain"))
s.band("3.6.6-7", "Row context vs filter context - same numbers, two views")
s.scenario("Filter context: a pivot with category on rows filters tTransactions to that category before [Net Revenue] runs. "
           "Formula twin: SUMIFS over the D_ tables reproduces each cell - use it to check your pivot.")
cr = netm.groupby("category").agg(Orders=("order_id", "count"), Net_Revenue=("total_price", "sum")).reset_index().sort_values("Net_Revenue", ascending=False)
s.table(cr, "tDaxCatCheck", fmts={"Orders": "int", "Net_Revenue": "naira"},
        calc={"Share": ("=[@Net_Revenue]/SUM([Net_Revenue])", "pct", (cr.Net_Revenue / cr.Net_Revenue.sum()).tolist())},
        total={"category": "Total", "Orders": "sum", "Net_Revenue": "sum"})

s = b.sheet("3.6B DAX Time & KPIs", "3.6  Power Pivot and DAX - Time Intelligence, KPIs and Performance (3.6.22 - 3.6.41)",
            "Time intelligence on DimDate, ranking, blanks - and KPI measures for all three organisations", "3.6.22 - 3.6.41")
ny = netm.groupby("Year").total_price.sum()
nm25 = netm[netm.Year == 2025].groupby(netm.order_date.dt.month).total_price.sum()
nm24 = netm[netm.Year == 2024].groupby(netm.order_date.dt.month).total_price.sum()
dax_b = [
    ("3.6.22", "Revenue YTD", "Rev YTD := TOTALYTD ( [Net Revenue], DimDate[Date] )", nm25.loc[:6].sum(), "naira", "Shown for June 2025 (Jan-Jun)"),
    ("3.6.23", "Revenue same period LY", "Rev LY := CALCULATE ( [Net Revenue], SAMEPERIODLASTYEAR ( DimDate[Date] ) )", ny[2024], "naira", "Shown for year 2025 → returns 2024"),
    ("3.6.23", "YoY growth", "YoY % := DIVIDE ( [Net Revenue] - [Rev LY], [Rev LY] )", ny[2025] / ny[2024] - 1, "pct", "2025 vs 2024"),
    ("3.6.24", "Revenue previous month", "Rev PM := CALCULATE ( [Net Revenue], DATEADD ( DimDate[Date], -1, MONTH ) )", nm25.loc[11], "naira", "Shown for Dec 2025 → Nov 2025"),
    ("3.6.25", "Rolling 3 months", "Rev R3M := CALCULATE ( [Net Revenue], DATESINPERIOD ( DimDate[Date], MAX ( DimDate[Date] ), -3, MONTH ) )", nm25.loc[10:12].sum(), "naira", "Shown for Dec 2025 (Oct-Dec)"),
    ("3.6.26", "Running total (all time)", "Rev Running := CALCULATE ( [Net Revenue], FILTER ( ALL ( DimDate ), DimDate[Date] <= MAX ( DimDate[Date] ) ) )", ny.sum(), "naira", "At 31-Dec-2025 = all-time net revenue"),
    ("3.6.27", "% of total revenue", "Rev % Total := DIVIDE ( [Net Revenue], CALCULATE ( [Net Revenue], ALL ( tProducts ) ) )", cat_rev.max() / cat_rev.sum(), "pct", "Electronics row"),
    ("3.6.28", "Seller rank", "Seller Rank := RANKX ( ALL ( tSellers[store_name] ), [Net Revenue] )", 1, "int", f"Rank 1 = {netm.merge(se[['seller_id','store_name']], on='seller_id').groupby('store_name').total_price.sum().idxmax()}"),
    ("3.6.29", "Revenue (zero not blank)", "Rev 0 := [Net Revenue] + 0", None, None, "Shows 0 for sellers with no sales (3 sellers) instead of hiding them"),
    ("3.6.30", "Safe ratio", "Return Rate := DIVIDE ( CALCULATE ( [Orders], tTransactions[order_status] = \"Returned\" ), [Orders], 0 )", (tx.order_status == "Returned").mean(), "pct", "511 returned of 11,000"),
]
ddb = pd.DataFrame([[a, b_, c, "" if d is None else d, f] for a, b_, c, d, e, f in dax_b], columns=["Topic", "Measure", "DAX", "Expected result", "Context / note"])
tdb = s.table(ddb, "tDaxB", widths={"Topic": 7, "Measure": 22, "DAX": 70, "Expected result": 16, "Context / note": 40})
for i, (_, _, _, val, kind, _) in enumerate(dax_b):
    if val is not None and kind:
        s.ws.write(tdb.r0 + 1 + i, tdb.ci("Expected result"), float(val), b.nf(kind, style="plain"))
s.band("3.6.23", "Formula twin: net revenue by year (check your SAMEPERIODLASTYEAR)")
yy = pd.DataFrame({"Year": ny.index.astype(int), "Net_Revenue": ny.values})
s.table(yy, "tNetYear", fmts={"Year": "gen", "Net_Revenue": "naira"},
        calc={"Prior_Year": ('=IFERROR(INDEX([Net_Revenue],MATCH([@Year]-1,[Year],0)),"")', "naira", [""] + yy.Net_Revenue.tolist()[:-1]),
              "YoY_Growth": ('=IFERROR([@Net_Revenue]/[@Prior_Year]-1,"")', "pct", [""] + (yy.Net_Revenue.values[1:] / yy.Net_Revenue.values[:-1] - 1).tolist())})
s.band("3.6.35-37", "KPI measures for the three organisations - definitions and expected values")
lc_rd = (lc.Readmit_30d == "Yes").mean()
occ = census.Occupancy_Pct.mean()
kpi = [["Retail (SuperMart Stores)", "Sales", "SUM ( Sales[Sales] )", s25.Sales.sum(), "naira", "FY2025"],
       ["Retail (SuperMart Stores)", "Margin %", "DIVIDE ( SUM ( Sales[Profit] ), SUM ( Sales[Sales] ) )", s25.Profit.sum() / s25.Sales.sum(), "pct", "FY2025"],
       ["Retail (SuperMart Stores)", "Target attainment", "DIVIDE ( [Sales], SUM ( Stores[Annual_Target_2025] ) )", s25.Sales.sum() / stores.Annual_Target_2025.sum(), "pct", "FY2025"],
       ["Hospital (Lifecare)", "ALOS (days)", "AVERAGE ( Admissions[length_of_stay] )", lc.Length_of_Stay.mean(), "dec1", "2019-2024"],
       ["Hospital (Lifecare)", "30-day readmission rate", "DIVIDE ( CALCULATE ( COUNTROWS ( Admissions ), Admissions[readmit_30d] = \"Yes\" ), COUNTROWS ( Admissions ) )", lc_rd, "pct", "2019-2024"],
       ["Hospital (Lifecare)", "Average bed occupancy", "AVERAGE ( BedCensus[occupancy_pct] )", occ, "pct", "Daily census, 520 staffed beds"],
       ["Hospital (Lifecare)", "Days above 95% occupancy", "COUNTROWS ( FILTER ( BedCensus, BedCensus[occupancy_pct] > 0.95 ) )", int((census.Occupancy_Pct > 0.95).sum()), "int", "Escalation days"],
       ["University (Crestview)", "Average exam score", "AVERAGE ( Students[Exam_Score] )", cv.Exam_Score.mean(), "dec1", "2024/25 first semester"],
       ["University (Crestview)", "Grade A rate (≥70)", "DIVIDE ( CALCULATE ( COUNTROWS ( Students ), Students[Exam_Score] >= 70 ), COUNTROWS ( Students ) )", (cv.Exam_Score >= 70).mean(), "pct", "2024/25 first semester"],
       ["University (Crestview)", "At-risk students", "CALCULATE ( COUNTROWS ( Students ), Students[At_Risk] = \"Yes\" )", int((cv.At_Risk == "Yes").sum()), "int", "Score < 62 or attendance < 65%"],
       ["University (Crestview)", "Average attendance", "AVERAGE ( Students[Attendance_Pct] ) / 100", cv.Attendance_Pct.mean() / 100, "pct", "2024/25 first semester"]]
kd = pd.DataFrame([[a, b_, c, d, f] for a, b_, c, d, e, f in kpi], columns=["Organisation", "KPI", "DAX", "Expected", "Scope"])
tk = s.table(kd, "tKpiDax", widths={"Organisation": 22, "KPI": 22, "DAX": 62, "Expected": 15, "Scope": 22})
for i, row in enumerate(kpi):
    s.ws.write(tk.r0 + 1 + i, tk.ci("Expected"), float(row[3]), b.nf(row[4], style="plain"))
s.band("3.6.29-41", "Blank vs zero, debugging, performance and documentation")
s.grid([["BLANK vs 0", "3 sellers have no sales: [Net Revenue] is BLANK, so they vanish from a pivot", "Add +0 only where showing them matters"],
        ["Debugging", "Measure gives the grand total in every row", "Missing/inactive relationship or filter on the wrong table"],
        ["Performance", "Iterators (SUMX/FILTER) over 11,000 rows are instant; over 10m rows prefer CALCULATE with simple filters", "Performance Analyzer (3.6.39)"],
        ["Implicit vs explicit", "Dragging total_price to Values makes an implicit 'Sum of total_price'", "Always write explicit measures - reusable and documented"],
        ["Format strings", "₦#,##0;-₦#,##0;- for currency measures", "Set once on the measure, applies everywhere"],
        ["Documentation", "Description field + a Measures sheet listing name, DAX, owner, date", "Hand-over ready (3.6.40)"]], header=["Topic", "Example in this model", "Practice"])

# ======================================================================================
# 3.7 Charts
# ======================================================================================
s = b.sheet("3.7 Chart Types", "3.7  Chart Anatomy and the 17 Excel Chart Types",
            "Each chart type drawn from a real bootcamp table - with when to use it", "3.7.1 - 3.7.26")
zone = s25.groupby("Zone").agg(Sales=("Sales", "sum"), Profit=("Profit", "sum")).reset_index().sort_values("Sales", ascending=False)
s.band("3.7.10 / 3.7.19", "Column and combo: FY2025 sales by zone with margin on a secondary axis")
tz = s.table(zone, "tChZone", fmts={"Sales": "naira", "Profit": "naira"}, calc={"Margin": ("=[@Profit]/[@Sales]", "pct", (zone.Profit / zone.Sales).tolist())})
s.chart("column", [{"name": "Sales", "categories": tz.ref("Zone"), "values": tz.ref("Sales")}], "Sales by zone (₦m) and margin (%)", y_fmt='₦#,##0,,"m"',
        secondary={"type": "line", "name": "Margin", "categories": tz.ref("Zone"), "values": tz.ref("Margin"), "y2_fmt": "0%"}, rows=16)
s.band("3.7.11", "Bar: stores ranked by FY2025 sales (long names read better horizontally)")
st_ = s25.groupby("Store_Name").Sales.sum().sort_values().reset_index()
tst = s.table(st_, "tChStores", fmts={"Sales": "naira"}, widths={"Store_Name": 28})
s.chart("bar", [{"name": "Sales", "categories": tst.ref("Store_Name"), "values": tst.ref("Sales"), "labels": True, "label_fmt": '#,##0.0,,"m"'}],
        "FY2025 sales by store (₦m)", legend=False, y_fmt='#,##0,,"m"', rows=18, size=(720, 360))
s.band("3.7.12-13", "Line and area: monthly sales 2022-2025")
mon = sm.groupby(sm.Order_Date.dt.to_period("M")).Sales.sum().reset_index()
mon["Month"] = mon.Order_Date.dt.to_timestamp()
mon = mon[["Month", "Sales"]]
tmn = s.table(mon, "tChMonthly", fmts={"Month": "mon", "Sales": "naira"})
s.chart("line", [{"name": "Sales", "categories": tmn.ref("Month"), "values": tmn.ref("Sales")}], "Monthly sales - a clear Q4 peak every year",
        y_fmt='₦#,##0,,"m"', legend=False, rows=16, extra=lambda ch: ch.set_x_axis({"date_axis": True, "num_format": "mmm yy", "major_unit": 6, "major_unit_type": "months"}))
s.chart("area", [{"name": "Sales", "categories": tmn.ref("Month"), "values": tmn.ref("Sales")}], "Same data as an area chart - emphasises volume",
        y_fmt='₦#,##0,,"m"', legend=False, rows=16)
s.band("3.7.14-15", "Pie and doughnut: SuperMart Online payment mix and membership tiers")
pm = tx.groupby("payment_method").total_price.sum().sort_values(ascending=False).reset_index()
tpm = s.table(pm, "tChPay", fmts={"total_price": "naira"}, widths={"payment_method": 18})
s.chart("pie", [{"name": "Revenue", "categories": tpm.ref("payment_method"), "values": tpm.ref("total_price"), "labels": True,
                 "points": [{"fill": {"color": c}} for c in ["#0B2545", "#13807A", "#E0A526", "#B23A48", "#6C8EAD", "#7FB069", "#8E6C8A"]]}],
        "Revenue by payment method", rows=16, size=(520, 300))
tier = cu.membership_tier.value_counts().reindex(["Bronze", "Silver", "Gold", "Platinum"]).reset_index()
tier.columns = ["Tier", "Customers"]
ttr = s.table(tier, "tChTier", fmts={"Customers": "int"})
s.chart("doughnut", [{"name": "Customers", "categories": ttr.ref("Tier"), "values": ttr.ref("Customers"), "labels": True,
                      "points": [{"fill": {"color": c}} for c in ["#B08D57", "#A7A9AC", "#E0A526", "#0B2545"]]}],
        "Customers by membership tier", rows=16, size=(520, 300))
s.band("3.7.16-17", "Scatter (and bubble): Crestview hours studied vs exam score")
samp = cv.sample(300, random_state=1)[["Student_ID", "Hours_Studied_Per_Week", "Exam_Score", "Attendance_Pct"]].reset_index(drop=True)
tsc = s.table(samp, "tChScatter", fmts={"Hours_Studied_Per_Week": "int", "Exam_Score": "int"})
s.chart("scatter", [{"name": "Students", "categories": tsc.ref("Hours_Studied_Per_Week"), "values": tsc.ref("Exam_Score"),
                     "trendline": {"type": "linear", "display_equation": True, "display_r_squared": True}}],
        "Hours studied vs exam score (300 sampled students)", x_title="Hours studied per week", y_title="Exam score", legend=False, rows=17)
s.para("Bubble chart (3.7.17): Insert > Scatter > Bubble with X = hours, Y = score, size = attendance. XlsxWriter cannot draw bubbles, so insert it yourself from tChScatter.", "note")
s.band("3.7.18", "Radar: four zones on five normalised measures")
zz = s25.groupby("Zone").agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"), Orders=("Order_ID", "nunique"), Customers=("Customer_ID", "nunique")).reset_index()
zz["Margin"] = zz.Profit / zz.Sales
rad = pd.DataFrame({"Measure": ["Sales", "Profit", "Orders", "Customers", "Margin"]})
for _, r_ in zz.iterrows():
    rad[r_.Zone] = [r_[k] / zz[k].max() for k in ["Sales", "Profit", "Orders", "Customers", "Margin"]]
trd = s.table(rad, "tChRadar", fmts={z: "pct0" for z in zz.Zone})
s.chart("radar", [{"name": z, "categories": trd.ref("Measure"), "values": trd.ref(z), "type": "line"} for z in zz.Zone],
        "Zone profile (each measure as % of the best zone)", rows=17, size=(560, 320))
s.band("3.7.20", "Waterfall: profit bridge FY2024 → FY2025 by category")
p24 = sm[sm.Year == 2024].groupby("Category").Profit.sum()
p25 = s25.groupby("Category").Profit.sum()
wf = [["FY2024 profit", p24.sum()]] + [[f"{c_} change", p25[c_] - p24[c_]] for c_ in p24.index] + [["FY2025 profit", p25.sum()]]
wfd = pd.DataFrame(wf, columns=["Step", "Amount"])
twf = s.table(wfd, "tChWaterfall", fmts={"Amount": "naira"}, widths={"Step": 26})
s.para("Select tChWaterfall > Insert > Waterfall. Right-click 'FY2025 profit' > Set as total. The bars show which category added the most profit.", "note")
s.band("3.7.21", "Funnel: SuperMart Online order funnel")
fn = pd.DataFrame({"Stage": ["Orders placed", "Not cancelled", "Shipped", "Delivered (not returned)"],
                   "Orders": [len(tx), int((tx.order_status != "Cancelled").sum()), len(dv), int((tx.order_status.isin(["Delivered", "Completed"])).sum())]})
tfn = s.table(fn, "tChFunnel", fmts={"Orders": "int"}, calc={"Pct_of_placed": ("=[@Orders]/INDEX([Orders],1)", "pct", (fn.Orders / fn.Orders.iloc[0]).tolist())}, widths={"Stage": 24})
s.chart("bar", [{"name": "Orders", "categories": tfn.ref("Stage"), "values": tfn.ref("Orders"), "labels": True, "label_fmt": "#,##0"}],
        "Order funnel (Insert > Funnel gives the centred version)", legend=False, rows=14, size=(620, 260),
        extra=lambda ch: ch.set_y_axis({"reverse": True}))
s.band("3.7.22-23", "Treemap and sunburst: revenue by category > subcategory")
tm = netm.merge(pr[["product_id", "subcategory"]], on="product_id").groupby(["category", "subcategory"]).total_price.sum().reset_index().sort_values(["category", "total_price"], ascending=[True, False])
s.table(tm, "tChTreemap", fmts={"total_price": "naira"}, widths={"category": 14, "subcategory": 20})
s.para("Select tChTreemap > Insert > Treemap (or Sunburst for rings). Hierarchy = category, subcategory; size = total_price.", "note")
s.band("3.7.24-25", "Histogram and box & whisker: Lifecare length of stay")
bins = list(range(0, 35, 5))
h_ = pd.cut(lc.Length_of_Stay, bins=bins + [100], right=False, labels=[f"{a}-{a + 4}" for a in bins[:-1]] + ["30+"]).value_counts().sort_index().reset_index()
h_.columns = ["LOS_days", "Admissions"]
h_["LOS_days"] = h_.LOS_days.astype(str)
thh = s.table(h_, "tChHist", fmts={"Admissions": "int"})
s.chart("column", [{"name": "Admissions", "categories": thh.ref("LOS_days"), "values": thh.ref("Admissions"), "gap": 5}],
        "Length of stay distribution (all admissions)", legend=False, rows=15, size=(600, 280))
bx = lc.groupby("Department").Length_of_Stay.describe()[["min", "25%", "50%", "75%", "max"]].reset_index()
bx.columns = ["Department", "Min", "Q1", "Median", "Q3", "Max"]
s.table(bx, "tChBox", fmts={k: "dec1" for k in ["Min", "Q1", "Median", "Q3", "Max"]}, widths={"Department": 26})
s.para("Insert > Box and Whisker on the raw Length_of_Stay by Department (Lifecare_Admissions_2019_2024.csv). The five-number summary above is what the box draws.", "note")
s.band("3.7.26", "Sparklines: monthly 2025 sales trend per store")
sp = s25.assign(M=s25.Order_Date.dt.month).pivot_table(index="Store_Name", columns="M", values="Sales", aggfunc="sum", fill_value=0)
sp.columns = [pd.Timestamp(2025, c_, 1).strftime("%b") for c_ in sp.columns]
sp = sp.reset_index()
top_sp = s.r + 1
s.frame(sp, fmts={c_: '#,##0,,"m"' for c_ in sp.columns[1:]})
s.ws.write(top_sp - 1, 14, "Trend", F["hdr"])
for i in range(len(sp)):
    s.ws.add_sparkline(top_sp + i, 14, {"range": f"'3.7 Chart Types'!C{top_sp + i + 1}:N{top_sp + i + 1}", "markers": True, "high_point": True, "series_color": "#0B2545"})
s.ws.set_column(14, 14, 18)

# ======================================================================================
# 3.8 Dashboard (working, formula-driven)
# ======================================================================================
s = b.sheet("3.8 Dashboard Build", "3.8  Dashboard Development - a working FY2025 SuperMart dashboard",
            "Pick a zone in the yellow cell: KPI cards, the store chart and the monthly trend all respond", "3.8.1 - 3.8.19")
zsel = s.inputs([("Zone (or All)", "All", "text", "Dropdown: All, South-West, South-South, North-Central, South-East")])["Zone (or All)"]
s.ws.data_validation(zsel.replace("$", ""), {"validate": "list", "source": ["All", "South-West", "South-South", "North-Central", "South-East"]})
crit = f'IF({zsel}="All","*",{zsel})'
s.band("3.8.7", "KPI cards")
k_sales, k_profit, k_orders = s25.Sales.sum(), s25.Profit.sum(), s25.Order_ID.nunique()
tgt_all = stores.Annual_Target_2025.sum()
cards = [("Net sales", f"=SUMIFS(tDash[Sales],tDash[Zone],{crit})", k_sales, "nairam"),
         ("Profit", f"=SUMIFS(tDash[Profit],tDash[Zone],{crit})", k_profit, "nairam"),
         ("Margin", f"=IFERROR(SUMIFS(tDash[Profit],tDash[Zone],{crit})/SUMIFS(tDash[Sales],tDash[Zone],{crit}),0)", k_profit / k_sales, "pct"),
         ("Target attainment", f"=SUMIFS(tDash[Sales],tDash[Zone],{crit})/SUMIFS(tDashStores[Target],tDashStores[Zone],{crit})", k_sales / tgt_all, "pct")]
card_fmt_l = b.wb.add_format({"bold": True, "font_color": "white", "bg_color": "#0B2545", "align": "center", "font_size": 10})
for j, (lab, f_, v, kind) in enumerate(cards):
    c0 = 1 + j * 3
    s.ws.merge_range(s.r, c0, s.r, c0 + 1, lab, card_fmt_l)
    s.ws.merge_range(s.r + 1, c0, s.r + 2, c0 + 1, "", b.nf(kind, style="result", font_size=18, align="center", valign="vcenter"))
    s.ws.write_formula(s.r + 1, c0, f_, b.nf(kind, style="result", font_size=18, align="center", valign="vcenter"), float(v))
s.r += 4
s.band("3.8.8 / 3.8.14", "Store performance for the selected zone (bar) and monthly trend (line)")
store_tab = SS = s25.groupby(["Store_ID", "Store_Name", "Zone"]).agg(Sales=("Sales", "sum"), Profit=("Profit", "sum")).reset_index()
store_tab = store_tab.merge(stores[["Store_ID", "Annual_Target_2025"]].rename(columns={"Annual_Target_2025": "Target"}), on="Store_ID")
tds = s.table(store_tab, "tDashStores", fmts={"Sales": "naira", "Profit": "naira", "Target": "naira"},
              calc={"Shown_Sales": (f'=IF(OR({zsel}="All",[@Zone]={zsel}),[@Sales],NA())', "naira", store_tab.Sales.tolist()),
                    "Attainment": ("=[@Sales]/[@Target]", "pct", (store_tab.Sales / store_tab.Target).tolist())},
              widths={"Store_Name": 26})
s.chart("bar", [{"name": "Sales", "categories": tds.ref("Store_Name"), "values": tds.ref("Shown_Sales"), "labels": True, "label_fmt": '#,##0.0,,"m"'}],
        "FY2025 sales - stores in the selected zone", legend=False, y_fmt='#,##0,,"m"', rows=18, size=(720, 340))
mrow = []
for mth in range(1, 13):
    mrow.append([pd.Timestamp(2025, mth, 1), s25[s25.Order_Date.dt.month == mth].Sales.sum()])
mm = pd.DataFrame(mrow, columns=["Month", "Sales_all"])
tmm = s.table(mm, "tDashMonth", fmts={"Month": "mon", "Sales_all": "naira"},
              calc={"Sales_selected": (f'=SUMIFS(tDash[Sales],tDash[Zone],{crit},tDash[Order_Date],">="&[@Month],tDash[Order_Date],"<="&EOMONTH([@Month],0))', "naira", mm.Sales_all.tolist())})
s.chart("line", [{"name": "Selected zone", "categories": tmm.ref("Month"), "values": tmm.ref("Sales_selected")}], "Monthly sales, FY2025 - selected zone",
        legend=False, y_fmt='₦#,##0,,"m"', rows=16)
s.band("3.8.2-6 / 3.8.17-19", "Design and QA checklist applied to this sheet")
s.grid([["Audience & question", "Regional managers: 'How is my zone doing against target?'", "✔"],
        ["Layout grid", "KPI row on top, comparison (bar) then trend (line) - Z-pattern reading", "✔"],
        ["Colour system", "Navy for data, gold only for the input, red reserved for problems", "✔"],
        ["One input drives everything", "Zone cell → every formula and both charts", "✔"],
        ["Accessibility", "Titles state the content; numbers labelled; no red/green-only meaning", "✔"],
        ["Reconciliation", "Net sales card with 'All' = ₦733,215,255 = POS control total", "✔"],
        ["Performance", "SUMIFS over 3,312 rows - instant; no volatile OFFSET/INDIRECT", "✔"]], header=["QA item", "How this dashboard meets it", "Status"])
s.para("With PivotTables instead of formulas: build pivots on tDash, insert Slicers (3.8.9) and connect one Zone slicer to all pivots via Report Connections (3.8.10).", "note")

# ======================================================================================
# 3.9 Blueprints
# ======================================================================================
s = b.sheet("3.9 Dashboard Blueprints", "3.9  Three Dashboard Blueprints",
            "KPI specification for each organisation with current values from the data", "3.9.1 - 3.9.3")
bp1 = [["Net sales FY2025", "SUM(Sales)", s25.Sales.sum(), "naira", "KPI card", "POS ledger"],
       ["Target attainment", "Sales ÷ Annual_Target_2025", s25.Sales.sum() / stores.Annual_Target_2025.sum(), "pct", "KPI card + bullet bar by store", "Store directory"],
       ["Margin %", "Profit ÷ Sales", s25.Profit.sum() / s25.Sales.sum(), "pct", "KPI card + combo by zone", "POS ledger"],
       ["Loss-making lines", "COUNT lines with Profit < 0", int((s25.Profit < 0).sum()), "int", "Table drill-down", "POS ledger"],
       ["Sales growth vs FY2024", "FY25 ÷ FY24 - 1", s25.Sales.sum() / sm[sm.Year == 2024].Sales.sum() - 1, "pct", "KPI card with arrow", "POS ledger"]]
bp2 = [["Students enrolled", "COUNT Student_ID", len(cv), "int", "KPI card", "Registry"],
       ["Average exam score", "AVERAGE Exam_Score", cv.Exam_Score.mean(), "dec1", "KPI card + bar by faculty", "Registry"],
       ["At-risk students", "At_Risk = Yes", int((cv.At_Risk == "Yes").sum()), "int", "KPI card + named list (restricted)", "Registry"],
       ["Average attendance", "AVERAGE Attendance_Pct", cv.Attendance_Pct.mean() / 100, "pct", "Gauge-style card", "Registry"],
       ["Grade A share (≥70)", "Grade = A ÷ students", (cv.Grade == "A").mean(), "pct", "Stacked bar by faculty", "Registry"]]
lc23 = lc[lc.Admission_Date.dt.year == 2023]
c23 = census[census.Census_Date.dt.year == 2023]
bp3 = [["Admissions 2023", "COUNT admissions", len(lc23), "int", "KPI card + monthly line", "EMR"],
       ["Average occupancy 2023", "Patients in house ÷ 520 beds", c23.Occupancy_Pct.mean(), "pct", "Line with 85% / 95% bands", "Bed census"],
       ["ALOS 2023 (days)", "AVERAGE Length_of_Stay", lc23.Length_of_Stay.mean(), "dec1", "Bar by department", "EMR"],
       ["30-day readmission rate 2023", "Readmit_30d = Yes ÷ admissions", (lc23.Readmit_30d == "Yes").mean(), "pct", "KPI card with RAG", "EMR + follow-up registry"],
       ["Average bill 2023", "AVERAGE Billing_Amount", lc23.Billing_Amount.mean(), "naira", "Bar by HMO", "Billing"]]
for tag, title, rows_, tname in [("3.9.1", "Blueprint 1: SuperMart Retail Sales Dashboard", bp1, "tBlueprint1"), ("3.9.2", "Blueprint 2: Crestview University Progression Dashboard", bp2, "tBlueprint2"),
                                 ("3.9.3", "Blueprint 3: Lifecare Hospital Operations Dashboard", bp3, "tBlueprint3")]:
    s.band(tag, title)
    d_ = pd.DataFrame([[a, b_, c, e, f] for a, b_, c, d, e, f in rows_], columns=["KPI", "Definition", "Current value", "Visual", "Source"])
    t_ = s.table(d_, tname, widths={"KPI": 26, "Definition": 30, "Current value": 16, "Visual": 30, "Source": 22})
    for i, r_ in enumerate(rows_):
        s.ws.write(t_.r0 + 1 + i, t_.ci("Current value"), float(r_[2]), b.nf(r_[3], style="plain"))

# ======================================================================================
# 3.10 Files into SQL, DAX and dashboards (rewritten module - internal files only)
# ======================================================================================
s = b.sheet("3.10 Files to SQL & DAX", "3.10  Bringing File Extracts into SQL, DAX and Dashboards",
            "Profile a CSV extract, load it into SQLite, exception-report with joins, then reconcile the dashboard", "3.10.1 - 3.10.5")
s.band("3.10.1", "Beginner: profile the extract before loading - Lifecare_Admissions_2023_DIRTY.csv")
dl = L("dirty_lc")
prof = pd.DataFrame({"Column": dl.columns, "Rows": len(dl), "Blank": dl.isna().sum().values, "Distinct": dl.nunique().values})
s.table(prof, "tLcProfile", fmts={"Rows": "int", "Blank": "int", "Distinct": "int"}, calc={"Blank_pct": ("=[@Blank]/[@Rows]", "pct", (prof.Blank / prof.Rows).tolist())},
        widths={"Column": 20})
s.band("3.10.2", "Beginner: ingest the extract into SQLite (staging table)")
s.code('''-- DB Browser for SQLite: File > Import > Table from CSV file... (header row, comma, UTF-8) → table stg_admissions_2023
-- or in the sqlite3 shell:
.mode csv
.import --skip 1 Lifecare_Admissions_2023_DIRTY.csv stg_admissions_2023
SELECT COUNT(*) FROM stg_admissions_2023;   -- must equal the CSV row count before any cleaning''')
s.calc([("CSV data rows", len(dl), len(dl), "int", "Count from the file itself"),
        ("Expected staging rows", len(dl), len(dl), "int", "Import must not drop or add rows")])
s.band("3.10.3", "Intermediate: SQL joins and an exception report")
sqlblock(s, "3.10.3a", "Admissions whose doctor is not in the doctors table (orphans)",
         "SELECT COUNT(*) AS orphan_admissions\nFROM admissions a\nLEFT JOIN doctors d ON d.doctor_id = a.doctor_id\nWHERE d.doctor_id IS NULL", db=DB_LC)
sqlblock(s, "3.10.3b", "Exception report - one row per rule, with counts",
         "SELECT 'Discharge before admission' AS rule, COUNT(*) AS rows_ FROM admissions WHERE discharge_date < admission_date\nUNION ALL\nSELECT 'Bill below zero', COUNT(*) FROM admissions WHERE billing_amount < 0\nUNION ALL\nSELECT 'Stay over 30 days', COUNT(*) FROM admissions WHERE length_of_stay > 30\nUNION ALL\nSELECT 'Room outside ward range', COUNT(*) FROM admissions a JOIN wards w ON w.ward_id = a.ward_id\n       WHERE a.room_number NOT BETWEEN w.first_room AND w.last_room",
         db=DB_LC, fmts={"rows_": "int"}, note="The production database passes every rule; run the same SQL on stg_admissions_2023 to see the dirty extract fail.")
s.band("3.10.4", "Intermediate: DAX measures and filter context on the hospital model")
s.grid([["Admissions", "COUNTROWS ( Admissions )", f"{len(lc):,}"],
        ["ALOS", "AVERAGE ( Admissions[length_of_stay] )", f"{lc.Length_of_Stay.mean():.1f} days"],
        ["Readmission %", "DIVIDE ( CALCULATE ( [Admissions], Admissions[readmit_30d] = \"Yes\" ), [Admissions] )", f"{lc_rd:.1%}"],
        ["Emergency readmission %", "CALCULATE ( [Readmission %], Admissions[admission_type] = \"Emergency\" )",
         f"{(lc[lc.Admission_Type == 'Emergency'].Readmit_30d == 'Yes').mean():.1%}"]], header=["Measure", "DAX", "Expected (all years)"])
s.band("3.10.5", "Advanced practical: reconcile a governed dashboard - SQL vs dashboard vs source")
lc_sql = q("SELECT COUNT(*) n, ROUND(SUM(billing_amount),2) bill FROM admissions", DB_LC)
s.calc([("Admissions - source CSV", len(lc), len(lc), "int", "Lifecare_Admissions_2019_2024.csv"),
        ("Admissions - SQL COUNT(*)", int(lc_sql.n[0]), int(lc_sql.n[0]), "int", "lifecare_hospital.db"),
        ("Billing - source CSV (₦)", round(float(lc.Billing_Amount.sum()), 2), round(float(lc.Billing_Amount.sum()), 2), "naira2", ""),
        ("Billing - SQL SUM (₦)", float(lc_sql.bill[0]), float(lc_sql.bill[0]), "naira2", ""),
        ("Reconciled?", f'=IF(AND({len(lc)}={int(lc_sql.n[0])},ABS({round(float(lc.Billing_Amount.sum()), 2)}-{float(lc_sql.bill[0])})<1),"PASS","FAIL")', "PASS", None,
         "Publish only when source, SQL and dashboard agree")])

# ---------------------------------------------------------------------- data sheets
b.data_sheet("D_Transactions", tx, "tTransactions", fmts={"order_date": "date", "unit_price": "naira", "total_price": "naira", "commission_amount": "naira2"},
             title="SuperMart Online - transactions (fact, 11,000 orders)", note="Source: supermart_online.db / SuperMart_Online_transactions.csv")
b.data_sheet("D_Deliveries", dv, "tDeliveries", fmts={"shipped_date": "date", "delivered_date": "date"}, title="SuperMart Online - deliveries (10,430)")
b.data_sheet("D_Customers", cu, "tCustomers", fmts={"date_of_birth": "date", "created_at": "date", "total_spent": "naira"}, title="SuperMart Online - customers (2,500)")
b.data_sheet("D_Products", pr, "tProducts", fmts={"price": "naira", "discount_price": "naira", "created_at": "date"}, title="SuperMart Online - products (800)")
b.data_sheet("D_Sellers", se, "tSellers", fmts={"total_revenue": "naira", "created_at": "date"}, title="SuperMart Online - sellers (200)")
b.data_sheet("D_DimDate", dim, "tDimDate", fmts={"Date": "date", "Year": "gen"}, title="Date dimension 2020-2025 (mark as date table)")
dash = s25[["Order_ID", "Order_Date", "Store_ID", "Store_Name", "Zone", "Category", "Sales", "Profit"]]
b.data_sheet("D_Dash_2025", dash, "tDash", fmts={"Order_Date": "date", "Sales": "naira", "Profit": "naira"}, title="SuperMart FY2025 lines feeding the 3.8 dashboard")

b.cover("Week 3 Section Examples - Databases, SQL, Data Model, DAX and Dashboards",
        ["One worksheet per module (3.1 - 3.10); DAX is split into foundations (3.6A) and time intelligence/KPIs (3.6B) like the slides.",
         "SQL results were produced by running the exact SQL shown against Week3_Data_Files/supermart_online.db and lifecare_hospital.db. "
         "All queries are collected in ETEH_Week3_SQL_Examples.sql.",
         "D_ sheets hold the SuperMart Online tables as Excel Tables - add them to the Data Model to build the star schema and test the DAX answers.",
         "3.8 is a working dashboard: change the yellow Zone cell and every KPI and chart updates."],
        [(k, DATASETS_TEXT[k]) for k in ["SuperMart Online", "SuperMart Stores", "Lifecare", "Crestview"]])
b.close()
SQL_OUT.write_text("-- ETEH Week 3 - SQL examples (SQLite). Run against the database named above each query.\n"
                   "-- Files: Week3_Data_Files/supermart_online.db, Week3_Data_Files/lifecare_hospital.db\n\n" + "\n".join(SQL_LOG), encoding="utf-8")
print("Week 3 ->", OUT, "|", SQL_OUT.name, len(SQL_LOG), "queries")
