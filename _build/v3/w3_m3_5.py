"""Week 3 edition-3 slide copy: Modules 3.3 (ODBC), 3.4 (merge and append) and 3.5 (Data Model and star schema).
Join counts come from running the SQL shown; append counts from the Week 2 monthly files; model facts from w3_data."""
import glob

import pandas as pd

import w3_data as D
from common import W2D
from v3kit import T
from w3_fx import (AOV, CODE, CU, DV, GROSS, I, LN, N, N_LINES, N_NET, NETREV, NO_DELIV, NT, P, PR, SE, SQL, TX, q, v)

ZONE_SQL = """SELECT customer_zone, COUNT(*) AS orders, SUM(total_price) AS net_revenue
FROM v_order_lines WHERE order_status IN ('Delivered','Completed') GROUP BY customer_zone"""
ZONE_ROWS = len(q(ZONE_SQL))


def _top(year):
    r = q(f"""SELECT t.seller_id, s.store_name, SUM(t.total_price) AS rev, COUNT(*) AS n FROM transactions t
              JOIN sellers s ON s.seller_id = t.seller_id
              WHERE strftime('%Y', t.order_date) = '{year}' GROUP BY t.seller_id ORDER BY rev DESC LIMIT 1""").iloc[0]
    return f"{r.seller_id} {r.store_name} ({N(r.rev)})"


SCOPE = {y: v(f"SELECT COUNT(*) FROM transactions WHERE strftime('%Y', order_date) = '{y}'") for y in (2024, 2025)}
INJ = v("SELECT COUNT(*) FROM transactions WHERE strftime('%Y', order_date) = '2025' OR '1'='1'")
BAD1 = v("SELECT COUNT(*) FROM transactions WHERE strftime('%Y', order_date) = 2025")
BAD2 = v("SELECT COUNT(*) FROM transactions WHERE order_date >= 20250101")
GOOD = v("SELECT COUNT(*) FROM transactions WHERE order_date >= '2025-01-01'")

M33 = dict(
    id="3.3", title="Connecting Excel to Databases with ODBC",
    subtitle="Drivers, DSNs, connection strings, credentials, SQL sent from Excel, parameters, refresh and the usual errors.",
    objectives=["Explain what ODBC and a driver do, and match the driver to Office's bitness.",
                "Choose between a DSN and a DSN-less connection string.",
                "Connect from Excel, send a SQL statement and load the result.",
                "Keep credentials out of the workbook and refresh safely.",
                "Use parameters without SQL injection, and fix the common ODBC errors."],
    topics=[
        T("3.3.1", "What Is ODBC?",
          defn="ODBC (Open Database Connectivity) is a standard programming interface that lets an application talk to any "
               "database for which an ODBC driver is installed.",
          simple="Excel speaks ODBC; the driver translates for the particular database. One method reaches SQLite, SQL Server, "
                 "MySQL and many more.",
          like="Like a universal travel adaptor: Excel's plug stays the same, the adaptor fits the socket.",
          grid=dict(label="Three layers in every connection", widths=[1.3, 1.5, 2.1], rows=[
              ["Layer", "Role", "In our example"],
              ["Application", "Asks for data", "Excel / Power Query"],
              ["ODBC driver", "Translates", "SQLite3 ODBC Driver"],
              ["Database", "Stores and answers", "supermart_online.db"]]),
          watch="A connection does not copy data for good: every refresh reads the database again."),

        T("3.3.2", "ODBC Drivers",
          defn="An ODBC driver is the software component that implements ODBC for one database product.",
          simple="The driver must match Office's bitness: 64-bit Excel needs a 64-bit driver. Check in File › Account › About Excel.",
          like="Like an interpreter who speaks both languages and turns up to the right meeting room.",
          grid=dict(label="Driver checks", widths=[2.4, 2.4], rows=[
              ["Situation", "Result"],
              ["64-bit Excel + 32-bit driver", "Data source name not found"],
              ["64-bit Excel + 64-bit driver", "Connects"],
              ["Driver name typed differently", "Data source name not found"]]),
          steps=["Windows search: ODBC Data Sources (64-bit)", "Drivers tab: find SQLite3 ODBC Driver",
                 "Copy its exact name into the connection string"],
          order=["steps", "grid"], split=0.45,
          watch="A 32/64-bit mismatch is the most common ODBC failure, and the message does not say so."),

        T("3.3.3", "DSN versus DSN-less Connections",
          defn="A DSN (data source name) is a connection saved in Windows under a name; a DSN-less connection keeps the full "
               "connection string inside the workbook.",
          simple="A DSN is handy but lives on your PC only. DSN-less travels with the file; colleagues need just the driver.",
          like="A DSN is a speed-dial entry on your phone; DSN-less is the full number written on the document.",
          **CODE("""
Driver={SQLite3 ODBC Driver};
Database=C:\\ETEH\\01_Section_Examples\\
         Week3_Data_Files\\supermart_online.db;""", label="DSN-LESS CONNECTION STRING"),
          grid=dict(label="Which to use", widths=[1.1, 1.6, 1.8], rows=[
              ["Type", "Stored", "On a colleague's PC"],
              ["DSN", "Windows, your PC", "Only if they create the same DSN"],
              ["DSN-less", "In the workbook", "Works if the driver is installed"]]),
          split=0.5,
          watch="A workbook built on a DSN only your PC has fails for everyone else with 'Data source name not found'."),

        T("3.3.4", "Configuring a Connection in Excel",
          defn="Excel's ODBC connector (Get Data › From Other Sources › From ODBC) opens a connection and lets you pick a "
               "table or send your own SQL statement.",
          simple=f"Paste the zone query from 3.2.17 under Advanced options: only its {ZONE_ROWS} result rows travel to Excel.",
          steps=["Data › Get Data › From Other Sources › From ODBC", "Data source name: (None); paste the connection string",
                 "Advanced options › SQL statement: paste the query", "OK › Transform Data to check, or Load To › Table"],
          grid=dict(label="What arrives", widths=[1.6, 1.8], rows=[
              ["Choice", "Result"],
              ["Navigator, table transactions", f"{I(N_LINES)} rows"],
              ["SQL statement (zone query)", f"{ZONE_ROWS} rows, one per zone"]]),
          order=["steps", "grid"], split=0.55,
          watch="Don't load every table 'just in case': each one is read again on every refresh."),

        T("3.3.5", "Authentication and Credential Security",
          defn="Authentication proves who is connecting: Windows (integrated) authentication, a database username and "
               "password, or a token.",
          simple="Power Query keeps credentials in its own store on your PC, not in the workbook. Never type a password into "
                 "a cell or a connection string.",
          like="Like an ATM PIN: you type it when asked, but never write it on the card.",
          grid=dict(label="Methods", widths=[2, 2.2, 1], rows=[
              ["Method", "Security", "Advice"],
              ["Windows authentication", "Strong; nothing stored in the file", "Preferred"],
              ["Prompted username + password", "Kept in the credential store", "Acceptable"],
              ["Password in connection string", "Visible to anyone with the file", "Never"]]),
          watch="SQLite has no logins at all: protect the .db file with folder permissions instead."),

        T("3.3.6", "Importing a Whole Table versus Writing SQL",
          defn="Query folding, or SQL pushdown, means sending the filtering and summarising to the database so only the result "
               "is transferred.",
          simple="Import the whole table and you filter in Excel; write SQL and the database does the work and sends the answer.",
          grid=dict(label="Rows that travel to Excel", widths=[2, 1.2, 1.6], rows=[
              ["Approach", "Rows", "Work left in Excel"],
              ["A: whole transactions table", I(N_LINES), "Filter, join, group"],
              ["B: the zone SQL query", I(ZONE_ROWS), "None"],
              ["Rows not transferred", I(N_LINES - ZONE_ROWS), ""]]),
          watch="On a server with millions of rows, approach A can take minutes or fail; approach B takes seconds."),

        T("3.3.7", "Refreshing Database Queries",
          defn="Refreshing re-runs a query's connection and SQL and replaces the loaded result with the current data.",
          simple="Data › Refresh All re-runs every saved query. The SQL is stored in the workbook, so anyone can refresh it.",
          grid=dict(label="Refresh settings (Query Properties)", widths=[2, 1.4, 2], rows=[
              ["Setting", "Advice", "Why"],
              ["Refresh All before sending", "Always", "Numbers are current"],
              ["Refresh data when opening the file", "Shared reports", "Readers see today's data"],
              ["Enable background refresh", "Off for big queries", "No half-loaded results"]]),
          watch="A refresh that fails leaves the old numbers on screen; read the status message before you send the file."),

        T("3.3.8", "Parameters in SQL Queries",
          defn="A query parameter is a named value supplied at run time, so one query can serve many reports.",
          simple="A parameter such as pYear is placed into the SQL. Never paste free text typed by users into SQL: that "
                 "invites SQL injection.",
          grid=dict(label="One query, three values typed", widths=[1.6, 2.6, 1.2], size=11.5, rows=[
              ["Year typed", "Top seller by revenue", "Rows in scope"],
              ["2024", _top(2024), I(SCOPE[2024])],
              ["2025", _top(2025), I(SCOPE[2025])],
              ["2025' OR '1'='1", "Every row (injection)", I(INJ)]]),
          watch="The third value turned 'one year' into 'every row' without any error. Validate parameters: a year is four digits."),

        T("3.3.9", "Troubleshooting ODBC Errors",
          defn="Most ODBC failures fall into four groups: driver mismatch, wrong path, missing credentials, and SQL that does "
               "not fit the dialect or the data types.",
          simple="SQLite stores dates as text. Compare text with text, or the filter silently returns the wrong rows.",
          grid=dict(label="Three ways to ask for 2025 orders", widths=[2.6, 0.8, 2.2], size=11.5, rows=[
              ["Condition", "Rows", "Verdict"],
              ["strftime('%Y', order_date) = 2025", I(BAD1), "Wrong: text compared with a number"],
              ["order_date >= 20250101", I(BAD2), "Wrong: text ranks above every number"],
              ["order_date >= '2025-01-01'", I(GOOD), "Correct"]]),
          watch="Zero rows or all rows, with no error, is the signature of a type mismatch in WHERE."),
    ],
    takeaways=["ODBC + a matching driver connects Excel to almost any database.",
               "Prefer DSN-less strings so the workbook travels.",
               "Send SQL so the database does the work; only results travel.",
               "Credentials stay in the credential store, never in cells.",
               "Validate parameters; compare text dates with text."],
)

# ---------------------------------------------------------------- 3.4
INNER = q("SELECT COUNT(*) AS n, SUM(t.total_price) AS v FROM transactions t JOIN deliveries d ON d.delivery_id = t.delivery_id").iloc[0]
LEFT = v("SELECT COUNT(*) FROM transactions t LEFT JOIN deliveries d ON d.delivery_id = t.delivery_id")
RIGHT = v("SELECT COUNT(*) FROM deliveries d LEFT JOIN transactions t ON t.delivery_id = d.delivery_id")
RANTI = v("SELECT COUNT(*) FROM deliveries d LEFT JOIN transactions t ON t.delivery_id = d.delivery_id WHERE t.transaction_id IS NULL")
NEVER_SOLD = q("""SELECT s.seller_id FROM sellers s LEFT JOIN transactions t ON t.seller_id = s.seller_id
                  WHERE t.transaction_id IS NULL ORDER BY s.seller_id""").seller_id.tolist()
PROD_NEVER = v("SELECT COUNT(*) FROM products p LEFT JOIN transactions t ON t.product_id = p.product_id WHERE t.transaction_id IS NULL")
C985 = q("SELECT COUNT(*) AS n, SUM(total_price) AS v FROM transactions WHERE customer_id = 'CUST00985'").iloc[0]
EXPAND = v("SELECT COUNT(*) FROM customers c LEFT JOIN transactions t ON t.customer_id = c.customer_id")
CART = q("SELECT COUNT(*) AS n, SUM(t.total_price) AS v FROM transactions t JOIN deliveries d ON d.customer_id = t.customer_id").iloc[0]
MFILES = sorted(glob.glob(str(W2D / "SuperMart_Monthly_Sales_2025" / "*.csv")))
MROWS = [len(pd.read_csv(f)) for f in MFILES]
DIRTY = pd.read_csv(W2D / "SuperMart_Sales_2025_DIRTY.csv", dtype=str, keep_default_na=False, encoding="utf-8-sig")
Z4 = {"South-West", "South-South", "North-Central", "South-East"}
ZRAW = DIRTY.Zone[DIRTY.Zone != ""]
Z_FAIL = int((~ZRAW.isin(Z4)).sum())
Z_FAIL2 = int((~ZRAW.str.strip().isin(Z4)).sum())

# fuzzy example: eight seller names as a supplier typed them (built for the exercise; same table tFuzzy in the workbook)
FZ = SE.head(8)[["seller_id", "store_name"]].copy()
# every third name is typed with a different word form (only fuzzy matching finds it), the others in capitals or with a double space
_FORMS = {"Trading Company": "Trading Co.", "Shopping Centre": "Shopping Center", "Global Concepts": "Global Concept",
          "Superstore": "Super Store", "Online Shop": "Online Shopping", "Retail Hub": "Retail-Hub", "Enterprises": "Enterprise",
          "Ventures": "Venture", "Stores": "Store", "Supplies": "Supply", "Merchants": "Merchant", "Emporium": "Emporium Ltd",
          "Mart": "Mart Ltd", "Depot": "Depot Ltd", "Plaza": "Plaza Ltd"}


def _variant(s):
    k = next(k for k in _FORMS if s.endswith(k))
    return s[: -len(k)] + _FORMS[k]


FZ["typed"] = [_variant(s) if i % 3 == 0 else (s.upper() if i % 3 == 1 else s.replace(" ", "  ", 1))
               for i, s in enumerate(FZ.store_name)]
FZ_EXACT = int((FZ.typed == FZ.store_name).sum())
FZ_NORM = int((FZ.typed.str.lower().str.split().str.join(" ") == FZ.store_name.str.lower()).sum())
FZ_MISS = FZ[FZ.typed.str.lower().str.split().str.join(" ") != FZ.store_name.str.lower()].typed.tolist()

M34 = dict(
    id="3.4", title="Merging and Appending in Power Query",
    subtitle="Six join kinds with real row counts, fuzzy matching, expand or aggregate, the Cartesian trap and appending.",
    objectives=["Tell merge (join) from append (union).",
                "Choose among inner, left, right, full, left anti and right anti joins.",
                "Expand or aggregate a merged column, and avoid the Cartesian trap.",
                "Append two or many tables and handle mismatched columns.",
                "Clean keys, verify cardinality and document every merge."],
    topics=[
        T("3.4.1", "Merge versus Append",
          defn="A merge combines two tables side by side by matching key values (a SQL JOIN); an append stacks tables with "
               "the same columns on top of each other (SQL UNION ALL).",
          simple="Merge adds columns; append adds rows.",
          like="Merge is stapling a delivery note to each order; append is putting February's forms under January's.",
          grid=dict(label="Our two jobs", widths=[2, 1.6, 2], rows=[
              ["Job", "Operation", "Result"],
              ["transactions + deliveries", "Merge on delivery_id", f"{I(LEFT)} rows, more columns"],
              [f"Jan ({MROWS[0]}) + Feb ({MROWS[1]}) files", "Append", f"{MROWS[0] + MROWS[1]} rows, same columns"]]),
          watch="Appending tables with different columns, or merging on a column that is not a key, gives a wrong table with no error."),

        T("3.4.2", "Inner Join",
          defn="An inner join returns only the rows whose key appears in both tables.",
          simple=f"Orders without a delivery ({I(NO_DELIV)} cancelled orders) vanish from the result, without any warning.",
          like="Like a register that lists only pupils who were both enrolled and present.",
          **SQL("""
SELECT COUNT(*)           AS lines,
       SUM(t.total_price) AS value
FROM transactions t
JOIN deliveries d
  ON d.delivery_id = t.delivery_id;""", fmt={"value": "n"}, extra=[["All orders: " + I(N_LINES), N(GROSS)]],
                widths=[1.3, 1.5], split=0.52),
          watch=f"The inner join dropped {N(GROSS - INNER.v)} of orders. Check what an inner join drops before trusting its totals."),

        T("3.4.3", "Left Outer Join",
          defn="A left outer join returns every row of the left table plus matching columns from the right; unmatched right "
               "columns are null.",
          simple="Use it when no row of your main table may be lost. Power Query: Merge Queries › Join Kind › Left Outer.",
          **SQL("""
SELECT COUNT(*)                AS all_orders,
       COUNT(d.courier_name)   AS with_courier,
       SUM(d.delivery_id IS NULL) AS courier_null
FROM transactions t
LEFT JOIN deliveries d
  ON d.delivery_id = t.delivery_id;""", widths=[1, 1, 1], split=0.6),
          watch="Put the main table first (on the left). Swap them and the main table becomes the side that can lose rows."),

        T("3.4.4", "Right Outer Join",
          defn="A right outer join returns every row of the right table plus matching columns from the left.",
          simple="It equals a left join with the tables swapped, so most analysts swap the order and keep to left joins.",
          grid=dict(label="Two ways to keep every delivery", widths=[2.6, 1.4, 0.8], rows=[
              ["Written as", "Keeps every row of", "Rows"],
              ["transactions RIGHT JOIN deliveries", "deliveries", I(RIGHT)],
              ["deliveries LEFT JOIN transactions", "deliveries", I(RIGHT)]]),
          watch="Mixing left and right joins in one query confuses every reader; pick one direction."),

        T("3.4.5", "Full Outer Join",
          defn="A full outer join returns all rows from both tables, matched where possible and null-filled where not.",
          simple="It is the reconciliation join: list everything from both sides and see what is missing from each.",
          like="Like comparing a bank statement with your own record book and ticking off the matches.",
          grid=dict(label="transactions full outer join deliveries", widths=[1.6, 1, 1.8], rows=[
              ["Group", "Rows", "Meaning"],
              ["In both", I(INNER.n), "Order shipped"],
              ["Orders only", I(LEFT - INNER.n), "Never shipped"],
              ["Deliveries only", I(RANTI), "No unknown shipments"],
              ["Total", I(LEFT + RANTI), "Full outer result"]]),
          watch="An inner join used for a reconciliation hides exactly the differences you were looking for."),

        T("3.4.6", "Left Anti Join",
          defn="A left anti join returns the rows of the left table that have no match in the right table.",
          simple="It answers 'which of mine are missing over there?' and gives you an exception list to investigate.",
          like="Checking the guest list against arrivals: the anti join lists who didn't come.",
          **SQL("""
SELECT t.order_status, t.payment_status,
       COUNT(*) AS orders, SUM(t.total_price) AS value
FROM transactions t
LEFT JOIN deliveries d
  ON d.delivery_id = t.delivery_id
WHERE d.delivery_id IS NULL
GROUP BY 1, 2;""", fmt={"value": "n"}, size=11.5, split=0.55),
          watch="Find out why rows are missing before calling them errors: here every one is a cancellation."),

        T("3.4.7", "Right Anti Join",
          defn="A right anti join returns the rows of the right table that have no match in the left table.",
          simple="The mirror check: which products never sold, which sellers never made a sale.",
          grid=dict(label="Things with no order lines", widths=[1.6, 0.6, 2.4], rows=[
              ["Check", "Rows", "Detail"],
              ["Products never sold", I(PROD_NEVER), f"All {I(len(PR))} products have sales"],
              ["Sellers never sold", I(len(NEVER_SOLD)), ", ".join(NEVER_SOLD)]]),
          watch="Checking one direction only misses the gap on the other side; run both anti joins."),

        T("3.4.8", "Fuzzy Matching",
          defn="A fuzzy merge matches keys that are similar rather than identical, scoring similarity from 0 to 1 against a threshold.",
          simple="Use it for hand-typed names when no ID exists. Every match is a suggestion to review.",
          grid=dict(label="A supplier's hand-typed list of 8 seller names", widths=[2, 1.2, 2.2], size=11.5, rows=[
              ["Method", "Matched", "Note"],
              ["Exact match", f"{FZ_EXACT} of 8", "Capitals, spaces and word forms fail"],
              ["Ignore case and spaces", f"{FZ_NORM} of 8", "Missed: " + "; ".join(FZ_MISS)],
              ["Fuzzy, threshold 0.8", "Candidates for all", "Review each one"]]),
          **CODE("""
Table.FuzzyNestedJoin(Supplier, {"typed"},
  sellers, {"store_name"}, "Match",
  JoinKind.LeftOuter,
  [IgnoreCase = true, IgnoreSpace = true,
   Threshold = 0.8])""", label="M", size=11.5),
          order=["code", "grid"], split=0.45,
          watch="A low threshold can match two different sellers with similar names; never accept fuzzy matches unread."),

        T("3.4.9", "Expanding versus Aggregating After a Merge",
          defn="After a Power Query merge, each row holds a nested table: Expand brings its columns out as rows; Aggregate "
               "summarises it to one value per row.",
          simple="Expanding a one-to-many merge adds rows; aggregating keeps one row per key.",
          like="Expand unpacks every item from each parcel; aggregate writes 'contains 14 items' on the parcel.",
          grid=dict(label="customers merged with transactions", widths=[1.2, 1, 2.6], rows=[
              ["Action", "Rows", "CUST00985"],
              ["Expand", I(EXPAND), f"{I(C985.n)} rows, one per order"],
              ["Aggregate", I(len(CU)), f"1 row: {I(C985.n)} orders, {N(C985.v)}"]]),
          watch="Expand a one-to-many merge and then sum a customer column: each customer's value repeats per order."),

        T("3.4.10", "The Cartesian Product Trap",
          defn="When the join key repeats on both sides, every matching row pairs with every other, multiplying rows (a "
               "partial Cartesian product).",
          simple="Join orders to deliveries on customer_id instead of delivery_id and each order meets every delivery of that customer.",
          like="Introducing every guest to every other guest instead of to their partner.",
          grid=dict(label="Same tables, two keys", widths=[2, 1, 1.8], rows=[
              ["Join key", "Rows", "Value summed"],
              ["delivery_id (correct)", I(INNER.n), N(INNER.v)],
              ["customer_id (wrong)", I(CART.n), N(CART.v)]]),
          watch=f"Names make poor keys: {SE.owner_name.nunique()} owner names cover {len(SE)} sellers, so joining on owner_name multiplies rows too. "
                "Join and group on seller_id."),

        T("3.4.11", "Appending Two Tables",
          defn="Append Queries stacks two tables, matching columns by name, not by position.",
          simple="Home › Append Queries. A column named differently in one table becomes a separate, half-empty column.",
          grid=dict(label="Week2_Data_Files › SuperMart_Monthly_Sales_2025", widths=[2.4, 0.8, 1.4], rows=[
              ["Table", "Rows", "Columns"],
              ["2025-01_SuperMart_Sales.csv", I(MROWS[0]), "21, same names"],
              ["2025-02_SuperMart_Sales.csv", I(MROWS[1]), "21, same names"],
              ["Appended", I(MROWS[0] + MROWS[1]), "21, no new columns"]]),
          watch="Column order does not matter in an append; column names do."),

        T("3.4.12", "Appending Three or More Tables",
          defn="Append Queries as New accepts three or more tables; the From Folder connector appends every file in a folder, "
               "including files added later.",
          simple="For files that arrive monthly, connect to the folder once and refresh.",
          like="Like a letterbox: each new letter joins the pile with no extra effort.",
          grid=dict(label="Twelve monthly files", widths=[1.5, 1.4, 2], rows=[
              ["Method", "Queries", "Result"],
              ["Manual append", f"{len(MFILES)} queries", "Edit it for every new month"],
              ["From Folder", "1 query", f"{I(sum(MROWS))} rows = FY2025 POS lines"]]),
          watch="Appending twelve queries by hand means somebody forgets month thirteen."),

        T("3.4.13", "Handling Schema Mismatches on Append",
          defn="A schema mismatch is a difference in column names or types between tables being appended.",
          simple="Power Query keeps every column and fills gaps with null, with no error. Check null counts after each append.",
          like="Two class lists, one headed 'Surname' and one 'Last name': you get two half-filled columns.",
          grid=dict(label=f"What if December ({MROWS[-1]} rows) arrived as Sales_NGN?", widths=[1.2, 2, 1.4], rows=[
              ["Column", "Null after append", "After renaming"],
              ["Sales", f"{I(MROWS[-1])} of {I(sum(MROWS))} ({P(MROWS[-1] / sum(MROWS))})", "0"],
              ["Sales_NGN", f"{I(sum(MROWS) - MROWS[-1])} ({P(1 - MROWS[-1] / sum(MROWS))})", "Column gone"]]),
          watch="The appended total looks plausible while one month's sales sit unused in another column."),

        T("3.4.14", "Cleaning Join Keys Before Merging",
          defn="Key cleaning makes both sides of a join comparable: same type, trimmed, consistent case and no hidden characters.",
          simple="Merges fail on differences you cannot see. Clean both keys the same way: Trim, Clean, a consistent case, and one type.",
          like="Two keys cut almost the same: a tiny difference stops the lock turning.",
          grid=dict(label="Zone in SuperMart_Sales_2025_DIRTY.csv (Week 2)", widths=[2, 1.3, 1.3], rows=[
              ["Check", "Before", "After Trim"],
              ["Distinct non-blank Zone values", I(ZRAW.nunique()), I(ZRAW.str.strip().nunique())],
              ["Rows that fail an exact match", I(Z_FAIL), I(Z_FAIL2)]]),
          watch="Cleaning the key in one table only still fails: both sides must be cleaned the same way."),

        T("3.4.15", "Verifying Join Cardinality",
          defn="Join cardinality states how many rows on each side share a key: one-to-one, one-to-many or many-to-many.",
          simple="Before merging, check the key is unique on the 'one' side. After, compare row counts in and out.",
          like="Count passengers on and off a bus: more off than on means something is wrong.",
          **SQL("""
SELECT 'transactions' AS tbl, COUNT(*) AS rows_,
       COUNT(DISTINCT order_id) AS keys
FROM transactions
UNION ALL
SELECT 'deliveries', COUNT(*), COUNT(DISTINCT order_id)
FROM deliveries;""", extra=[["left join", I(LEFT), "= left: safe"]], widths=[1.2, 0.9, 0.9], split=0.58),
          watch="More rows out than in means a duplicated key; fewer means unmatched rows were dropped."),

        T("3.4.16", "Documenting Merge and Append Logic",
          defn="Merge documentation records each step's name, join kind, key, expected cardinality and row counts.",
          simple="Rename steps (not 'Merged Queries1') and log the counts in a 00_README sheet, so the next analyst knows what is normal.",
          like="Like labelling the wires in a fuse box.",
          grid=dict(label="Merge log entry", widths=[1.5, 3], rows=[
              ["Item", "Entry"],
              ["Step name", "Merged Orders to Deliveries"],
              ["Join kind / key", "Left outer on delivery_id"],
              ["Cardinality", "One-to-one"],
              ["Rows in / out", f"{I(N_LINES)} / {I(LEFT)} ({I(NO_DELIV)} with no courier)"]]),
          watch="An undocumented merge cannot be checked: nobody knows what row count to expect."),
    ],
    takeaways=["Merge adds columns (join); append adds rows (union).",
               "Choose the join kind deliberately; inner joins drop rows silently.",
               "Anti joins produce exception lists in both directions.",
               "Join on true keys; check cardinality and row counts.",
               "Append by column name; clean keys; log every merge."],
)

# ---------------------------------------------------------------- 3.5
DD = D.dimdate()
ORDER_DAYS = LN.order_date.dt.normalize().nunique()

M35 = dict(
    id="3.5", title="The Excel Data Model and Star Schema",
    subtitle="Relationships instead of lookups: facts, dimensions, a date table and a tidy field list.",
    objectives=["Explain what the Data Model is and why it beats copied lookup columns.",
                "Create one-to-many relationships and state their cardinality.",
                "Identify fact and dimension tables and their grain.",
                "Build and mark a gap-free date table.",
                "Hide keys, sort months and format measures."],
    topics=[
        T("3.5.1", "What Is the Excel Data Model?",
          defn="The Data Model is a relational, in-memory database inside the workbook (the Power Pivot engine) that stores "
               "tables and the relationships between them.",
          simple="It compresses data, handles millions of rows and links tables, so one PivotTable can use fields from several tables.",
          like="Like a filing cabinet with drawers tied by string: pull a customer and their orders come too.",
          grid=dict(label="A PivotTable from two related tables", widths=[1.8, 1, 1.6], rows=[
              ["products[category]", "Lines", "Gross value"]] + [
              [c, I(n), N(val)] for c, n, val in
              LN.groupby("category").agg(n=("total_price", "size"), val=("total_price", "sum")).sort_values("val", ascending=False)
              .head(3).itertuples()] + [["All 7 categories", I(N_LINES), N(GROSS)]]),
          watch="Copying lookup columns into the fact table, instead of relating tables, is the habit the Data Model replaces."),

        T("3.5.2", "Relationships and Cardinality",
          defn="A relationship links two tables on a shared key column; most are one-to-many, with unique keys on the 'one' side.",
          simple="One product, many order lines. Draw relationships in Data › Relationships or Power Pivot's Diagram View.",
          like="One mother, several children: each child points to exactly one mother.",
          grid=dict(label="Relationships from transactions", widths=[1.3, 1.8, 2], rows=[
              ["Column", "To table (rows)", "Cardinality"],
              ["customer_id", f"customers ({I(len(CU))})", "Many-to-one"],
              ["product_id", f"products ({I(len(PR))})", "Many-to-one"],
              ["seller_id", f"sellers ({I(len(SE))})", "Many-to-one"],
              ["delivery_id", f"deliveries ({I(len(DV))})", f"One-to-one; {I(NO_DELIV)} blanks"],
              ["order_date", f"DimDate ({I(len(DD))})", "Many-to-one"]]),
          watch="If the key repeats on the 'one' side, Excel refuses the relationship; fix the duplicates first."),

        T("3.5.3", "Fact Tables in the Data Model",
          defn="In the Data Model, the fact table holds the events at one grain, the numeric columns measures add up, and "
               "the keys that relate it to each dimension.",
          simple="tTransactions: one row per order line, with quantity, total_price and commission_amount.",
          ex=dict(t="tTransactions", cols=["transaction_id", "customer_id", "product_id", "quantity", "total_price", "order_date"],
                  n=4, size=11, note=f"Showing 4 of {I(N_LINES)} rows · grain: one order line"),
          watch="Keep one grain per fact table: never mix order lines and monthly totals."),

        T("3.5.4", "Dimension Tables in the Data Model",
          defn="A dimension table has one row per item, a unique key and the descriptive columns that slice the facts.",
          simple="Dimensions are short and wide. Their columns go on Rows, Columns and slicers.",
          grid=dict(label="SuperMart Online dimensions", widths=[1.8, 1.3, 2.4], rows=[
              ["Dimension", "Unique key", "Useful columns"],
              [f"customers ({I(len(CU))} × {CU.shape[1]})", "customer_id", "membership_tier, state, zone"],
              [f"products ({I(len(PR))} × {PR.shape[1]})", "product_id", "category, subcategory, brand_name"],
              [f"sellers ({I(len(SE))} × {SE.shape[1]})", "seller_id", "store_name, store_zone, rating"]]),
          watch="A dimension with duplicate keys cannot sit on the 'one' side of a relationship."),

        T("3.5.5", "The Date Table",
          defn="A date table is a dimension with exactly one row per calendar day over the whole period, without gaps, "
               "marked as the model's date table.",
          simple="It makes time intelligence (3.6.22) reliable and shows days and months with no sales.",
          ex=dict(t="tDimDate", cols=["Date", "Year", "Quarter", "Month_No", "Month", "Weekday"], n=3, size=11.5,
                  note=f"Showing 3 of {I(len(DD))} days · {DD.Date.min():%Y-%m-%d} to {DD.Date.max():%Y-%m-%d}, no gaps"),
          grid=dict(label="Checks", widths=[2.4, 1.2], rows=[
              ["Check", "Result"], ["Days in DimDate", I(len(DD))], ["Days with an order", I(ORDER_DAYS)],
              ["Days with no order (still shown)", I(len(DD) - ORDER_DAYS)]]),
          stack=False, split=0.62,
          watch="Using the fact table's order_date as the calendar leaves out days with no orders, so time intelligence miscounts."),

        T("3.5.6", "Hiding, Sorting and Arranging Model Fields",
          defn="Model housekeeping hides keys and technical columns, sets Sort by Column for month names, and gives every "
               "measure a format.",
          simple="Users then see only clean, sensible fields in a sensible order.",
          grid=dict(label="Before and after", widths=[1.8, 1.8, 1.8], rows=[
              ["Setting", "Before", "After"],
              ["Month order", "Apr, Aug, Dec, Feb …", "Jan, Feb, Mar …"],
              ["Key columns in field list", "Visible", "Hidden"],
              ["Net Revenue format", f"{NETREV:.0f}", N(NETREV)]]),
          steps=["Power Pivot › Manage › select DimDate[Month]", "Home › Sort by Column › by Month_No",
                 "Right-click key columns › Hide from Client Tools"],
          order=["steps", "grid"], split=0.45,
          watch="Without Sort by Column, month names sort alphabetically, April first."),

        T("3.5.7", "Why Relationships Beat VLOOKUP Columns",
          defn="A lookup column copies dimension data into every fact row; a relationship stores each description once and "
               "links it by key.",
          simple="Relationships keep working as rows are added; lookup columns bloat the file and must be extended.",
          grid=dict(label="Bringing product and customer details to the facts", widths=[2.2, 1.2, 1.8], rows=[
              ["Approach", "Extra cells", "When rows are added"],
              ["VLOOKUP category", I(N_LINES), "Extend formulas"],
              ["VLOOKUP tier, state, zone", I(3 * N_LINES) + " more", "Extend formulas"],
              ["Relationships", "0", "Refresh; nothing to extend"]]),
          watch="A VLOOKUP column that was not filled down for new rows returns blanks that nobody notices."),
    ],
    takeaways=["The Data Model links tables; one pivot can use them all.",
               "Relationships are one-to-many with unique keys on the 'one' side.",
               "Facts hold events and numbers; dimensions hold descriptions.",
               "Always use a gap-free, marked date table.",
               "Hide keys, sort months by number, format measures."],
)

MODULES = [M33, M34, M35]
