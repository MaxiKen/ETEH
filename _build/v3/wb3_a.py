"""Week 3 workbook, modules 3.1 - 3.5 (edition 3): databases, SQL, ODBC, merging and appending, the Data Model.

Definitions and SQL come from the slide copy (wb3_common.DEF). The learner runs the SQL in DB Browser for SQLite; the worked
solution works the same number out with worksheet formulas on the Data_ sheets (tTransactions, tLines, ...), so every SQL
result on the slides has a live twin here. Cached values are computed in pandas.
"""
from __future__ import annotations

import pandas as pd

import w3_m3_5 as M35
from w1x import ModuleSheet
from wb3_common import NETS, N, f, net, netn, tp

TX_DS = ("tTransactions", "Sheet Data_Transactions (A1)", "The transactions table: 11,000 order lines, 15 columns.")
LI_DS = ("tLines", "Sheet Data_Lines (A1)", "The v_order_lines view: the same 11,000 lines with category, customer zone and courier "
                                             "joined in SQL. Use it for 'by category' and 'by zone' questions.")


def build(b, D):
    m31(b, D)
    m32(b, D)
    m33(b, D)
    m34(b, D)
    m35(b, D)


# =====================================================================================================
def m31(b, D):
    TX, P, C, SE, DV, ST, L = D.TX, D.P, D.C, D.SE, D.DV, D.ST, D.L
    m = ModuleSheet(b, "3.1 Database Fundamentals", "3.1  Database Fundamentals",
                    "Tutor walk-through after Module 3.1: tables, keys, normalisation, types, indexes, views, ACID, star schema, "
                    "OLTP and OLAP, on supermart_online.db", "3.1.1 - 3.1.15")
    m.dataset([TX_DS, ("tProducts / tCustomers / tSellers", "Sheets Data_Products, Data_Customers, Data_Sellers (A1)",
                       "The dimension tables: 800 products, 2,500 customers, 200 sellers."),
               ("tDeliveries / tStates", "Sheets Data_Deliveries, Data_States (A1)", "10,430 deliveries; 32 states with their zone."),
               ("supermart_online.db", "01_Section_Examples › Week3_Data_Files", "The database itself: open it in DB Browser for SQLite.")])
    sizes = {"tTransactions": 11000, "tCustomers": len(C), "tProducts": len(P), "tSellers": len(SE), "tDeliveries": len(DV), "tStates": len(ST)}
    tp(m, "3.1.1",
       [("Order lines", "=ROWS(tTransactions)", len(TX), "int", "One row per order line."),
        ("Customers", "=ROWS(tCustomers)", len(C), "int", ""),
        ("Products", "=ROWS(tProducts)", len(P), "int", ""),
        ("Rows in all six tables", "=ROWS(tTransactions)+ROWS(tCustomers)+ROWS(tProducts)+ROWS(tSellers)+ROWS(tDeliveries)+ROWS(tStates)",
         sum(sizes.values()), "int", "Six related tables in one file.")],
       [("n", "How many deliveries does the database hold?", "=ROWS(tDeliveries)", len(DV), "int", "Deliveries"),
        ("n", "How many sellers?", "=ROWS(tSellers)", len(SE), "int", "Sellers")],
       steps=["Install DB Browser for SQLite (sqlitebrowser.org) and open it.",
              "Open Database › Week3_Data_Files › supermart_online.db (Read Only is safest).",
              "Database Structure tab: read the 6 tables, the view and the indexes.",
              "Browse Data tab: pick a table and scroll its rows."])
    tp(m, "3.1.2",
       [("Fields in tTransactions", "=COLUMNS(tTransactions)", 15, "int", "Each column is a field with one type."),
        ("Records (rows)", "=ROWS(tTransactions)", len(TX), "int", ""),
        ("Cells", "=ROWS(tTransactions)*COLUMNS(tTransactions)", len(TX) * 15, "int", "Rows × fields.")],
       [("n", "How many fields does tStates have?", "=COLUMNS(tStates)", 2, "int", "Fields in tStates"),
        ("t", "Which field of tTransactions holds the date of the order?", "order_date", None, None, ("order date",))],
       steps=["On Data_Transactions, row 1 holds the field names; every row below is one record.",
              "Click any cell and read the field name above it.", "In DB Browser, Database Structure › transactions lists the same fields with their types."])
    tp(m, "3.1.3",
       [("Rows", "=ROWS(tTransactions)", len(TX), "int", ""),
        ("Distinct transaction_id", "=COUNTA(UNIQUE(tTransactions[transaction_id]))", int(TX.transaction_id.nunique()), "int",
         "Equal to rows: the key is unique."),
        ("Duplicate keys", "=ROWS(tTransactions)-COUNTA(UNIQUE(tTransactions[transaction_id]))", 0, "int", "")],
       [("n", "How many distinct seller_id values are in tSellers?", "=SUMPRODUCT(1/COUNTIF(tSellers[seller_id],tSellers[seller_id]))",
         int(SE.seller_id.nunique()), "int", "Distinct seller_id"),
        ("t", "Is customer_id a valid primary key of tCustomers? Compare ROWS with the distinct count. (Yes / No)",
         "Yes" if C.customer_id.is_unique else "No", None, None, ("y",))],
       steps=["Run the SQL in DB Browser (Execute SQL tab, F5).", "Compare rows_ with distinct_ids: equal means no duplicates.",
              "Here, the worksheet twin compares ROWS with COUNTA(UNIQUE(...))."], code=True)
    no_sale = int((~P.product_id.isin(TX.product_id)).sum())
    tp(m, "3.1.4",
       [("Lines whose product_id is not in tProducts", "=SUMPRODUCT(--(COUNTIF(tProducts[product_id],tTransactions[product_id])=0))", 0, "int",
         "The SQL orphan check, done with COUNTIF."),
        ("Lines whose seller_id is not in tSellers", "=SUMPRODUCT(--(COUNTIF(tSellers[seller_id],tTransactions[seller_id])=0))", 0, "int", "")],
       [("n", "How many order lines point at a customer_id that is not in tCustomers?",
         "=SUMPRODUCT(--(COUNTIF(tCustomers[customer_id],tTransactions[customer_id])=0))", 0, "int", "Orphan customer lines"),
        ("n", "The other direction: how many products have never been ordered?",
         "=SUMPRODUCT(--(COUNTIF(tTransactions[product_id],tProducts[product_id])=0))", no_sale, "int", "Products never ordered")],
       steps=["Run the LEFT JOIN … IS NULL query: 0 means every line has a real product.",
              "The twin: COUNTIF counts each line's product_id in tProducts; a count of 0 is an orphan.",
              "Swap the tables to find products that were never ordered (not an error: just unsold)."], code=True)
    cust_lines = TX.groupby("customer_id").size()
    tp(m, "3.1.5",
       [("Customers", "=ROWS(tCustomers)", len(C), "int", "Relational: each customer's details stored once."),
        ("Order lines", "=ROWS(tTransactions)", len(TX), "int", "A flat file would repeat the customer's name, state and tier on each."),
        ("Most lines for one customer", "=MAX(COUNTIF(tTransactions[customer_id],tCustomers[customer_id]))", int(cust_lines.max()), "int",
         "A flat file repeats that customer's details this many times.")],
       [("n", "How many order lines does CUST00985 have? (In a flat file, their address appears that many times.)",
         '=COUNTIF(tTransactions[customer_id],"CUST00985")', int(cust_lines.get("CUST00985", 0)), "int", "Lines for CUST00985"),
        ("o", "A customer moves from Lagos to Abuja. How many cells change in the relational design, and why is a flat file worse?",
         "One cell (state in tCustomers). In a flat file every one of their order lines must change, and a missed row leaves two addresses.")],
       steps=["Open Data_Customers: one row per customer.", "Open Data_Transactions: customer_id only, no name or address.",
              "The relationship (customer_id) brings the details back when needed."])
    tp(m, "3.1.6",
       [("Fields per transactions row", "=COLUMNS(tTransactions)", 15, "int", "One value per cell: 1NF."),
        ("Distinct payment methods", "=COUNTA(UNIQUE(tTransactions[payment_method]))", int(TX.payment_method.nunique()), "int",
         "Atomic values: one method per row, never 'Card, Transfer'."),
        ("Cells with a comma in payment_method", '=COUNTIF(tTransactions[payment_method],"*,*")', int(TX.payment_method.str.contains(",").sum()),
         "int", "0: no lists inside cells.")],
       [("n", "How many product_name cells contain a comma (a possible list in one cell)?", '=COUNTIF(tProducts[product_name],"*,*")',
         int(P.product_name.str.contains(",").sum()), "int", "Commas in product_name"),
        ("t", "A cell holds 'Phone; Charger'. Which normal form does that break? (1NF / 2NF / 3NF)", "1NF", None, None, ("first normal form",))],
       steps=["Check every cell holds one value (no lists, no 'A, B').", "Check every row is unique (a primary key).",
              "Check there are no repeating column groups (Item1, Item2, Item3).", "COUNTIF with '*,*' finds commas hiding lists."])
    sw_bad = int(((C.state == "Lagos") & (C.zone != "South-West")).sum())
    tp(m, "3.1.7",
       [("States in the lookup", "=ROWS(tStates)", len(ST), "int", "Zone is stored once per state in tStates."),
        ("Lagos customers not in South-West", '=COUNTIFS(tCustomers[state],"Lagos",tCustomers[zone],"<>South-West")', sw_bad, "int",
         "The copy of zone in tCustomers agrees here."),
        ("Lagos customers", '=COUNTIF(tCustomers[state],"Lagos")', int((C.state == "Lagos").sum()), "int", "")],
       [("n", "How many Kano customers are not recorded in the North-West zone?", '=COUNTIFS(tCustomers[state],"Kano",tCustomers[zone],"<>North-West")',
         int(((C.state == "Kano") & (C.zone != "North-West")).sum()), "int", "Kano customers not North-West"),
        ("t", "zone depends on state, not on customer_id. Which normal form does storing zone in customers bend? (2NF / 3NF)", "3NF", None, None,
         ("third normal form",))],
       steps=["Ask of each non-key field: does it depend on the whole key, and only on the key?",
              "zone depends on state (a transitive dependency): 3NF says keep it in tStates.",
              "Run the SQL: 0 customers whose zone differs from the lookup, so the copy is consistent."], code=True)
    tp(m, "3.1.8",
       [("Numeric quantity cells", "=COUNT(tTransactions[quantity])", len(TX), "int", "INTEGER in SQLite."),
        ("Numeric total_price cells", "=COUNT(tTransactions[total_price])", len(TX), "int", "REAL."),
        ("Smallest quantity", "=MIN(tTransactions[quantity])", int(TX.quantity.min()), "int", "The CHECK (quantity > 0) rule holds.")],
       [("n", "What is the largest quantity on one line?", "=MAX(tTransactions[quantity])", int(TX.quantity.max()), "int", "Largest quantity"),
        ("t", "SQLite stores order_date as which type? (integer / real / text)", "text", None, None, ("TEXT",))],
       steps=["Run the typeof() query: integer, real, text.", "SQLite keeps dates as ISO text (YYYY-MM-DD), which sorts correctly.",
              "Power Query converted order_date to a date for this sheet (3.3.4)."], code=True)
    tp(m, "3.1.9",
       [("Lines for one customer", '=COUNTIF(tTransactions[customer_id],"CUST00985")', int(cust_lines.get("CUST00985", 0)), "int",
         "With idx_tx_customer, SQLite finds these without reading all 11,000 rows."),
        ("Rows a full scan reads", "=ROWS(tTransactions)", len(TX), "int", "Without an index.")],
       [("n", "How many lines does SELL0200 have? (idx_tx_seller would find them directly.)", '=COUNTIF(tTransactions[seller_id],"SELL0200")',
         int((TX.seller_id == "SELL0200").sum()), "int", "Lines for SELL0200")],
       steps=["In DB Browser run: EXPLAIN QUERY PLAN SELECT * FROM transactions WHERE customer_id = 'CUST00985';",
              "Read 'SEARCH transactions USING INDEX idx_tx_customer': the index is used.",
              "Try a column without an index (payment_method): the plan says SCAN."],
       do=[("Run EXPLAIN QUERY PLAN for customer_id = 'CUST00985'.", "SEARCH transactions USING INDEX idx_tx_customer (customer_id=?)"),
           ("Run it for payment_method = 'Paystack'.", "SCAN transactions: no index on payment_method.")])
    zl = L.groupby("customer_zone").size().sort_values(ascending=False)
    tp(m, "3.1.10",
       [(f"{zl.index[0]} lines", f'=COUNTIF(tLines[customer_zone],"{zl.index[0]}")', int(zl.iloc[0]), "int", "tLines is the view, imported."),
        (f"{zl.index[1]} lines", f'=COUNTIF(tLines[customer_zone],"{zl.index[1]}")', int(zl.iloc[1]), "int", ""),
        ("All lines in the view", "=ROWS(tLines)", len(L), "int", "Same as transactions: the view adds columns, not rows.")],
       [("n", "How many North-East lines does the view return?", '=COUNTIF(tLines[customer_zone],"North-East")', int(zl["North-East"]), "int",
         "North-East lines"),
        ("t", "Does a view store its own copy of the data? (Yes / No)", "No", None, None, ("n",))],
       steps=["Database Structure › Views › v_order_lines: right-click › Copy Create statement to read its SQL.",
              "Run the GROUP BY on the view as if it were a table.", "Data_Lines here is that view, imported through ODBC."], code=True)
    tp(m, "3.1.11",
       [("SuperMart order lines (main)", "=ROWS(tTransactions)", len(TX), "int", "main.transactions"),
        ("Mercyland admissions 2023 (lh)", "=ROWS(tAdm23)", len(D.A), "int", "lh.admissions after ATTACH, filtered to 2023.")],
       [("t", "After ATTACH DATABASE 'mercyland_hospital.db' AS lh, how do you name its admissions table in a query?", "lh.admissions", None, None, ())],
       steps=["Open supermart_online.db, then run: ATTACH DATABASE 'mercyland_hospital.db' AS lh;",
              "Query main.transactions and lh.admissions in the same session.",
              "The prefix (main., lh.) is the namespace: it removes any doubt about which file a table is in."])
    tp(m, "3.1.12",
       [("Lines before the test", "=ROWS(tTransactions)", len(TX), "int", "Run the test on a copy of the database."),
        ("Expected after ROLLBACK", "=ROWS(tTransactions)", len(TX), "int", "Atomic: nothing kept.")],
       [("n", "On a copy, BEGIN; insert two lines; ROLLBACK. How many rows does transactions hold afterwards?", "=ROWS(tTransactions)",
         len(TX), "int", "Rows after ROLLBACK"),
        ("t", "Which ACID property means 'all or nothing'?", "Atomicity", None, None, ("atomic",))],
       steps=["Copy supermart_online.db to a scratch folder and open the copy (not read-only).",
              "Run BEGIN; INSERT … ; SELECT COUNT(*) FROM transactions; (11,001 inside the transaction).",
              "Run ROLLBACK; and count again: 11,000."])
    tp(m, "3.1.13",
       [("Fact rows", "=ROWS(tTransactions)", len(TX), "int", "Centre of the star."),
        ("Dimension rows (customers + products + sellers)", "=ROWS(tCustomers)+ROWS(tProducts)+ROWS(tSellers)", len(C) + len(P) + len(SE), "int",
         "The points of the star."),
        ("Date dimension rows", "=ROWS(tDimDate)", len(D.DD), "int", "Added for analysis (3.5.5).")],
       [("t", "In the star version, where does a customer's zone live? (fact / dimension)", "dimension", None, None, ("tCustomers", "customers")),
        ("n", "How many dimension rows would the star have, counting tDimDate as well?",
         "=ROWS(tCustomers)+ROWS(tProducts)+ROWS(tSellers)+ROWS(tDimDate)", len(C) + len(P) + len(SE) + len(D.DD), "int", "Dimension rows")],
       steps=["Normalised: zone in tStates, joined through state.", "Star: zone copied into the customer dimension, one join from the fact.",
              "Fewer joins makes the star faster and simpler for PivotTables and DAX."])
    tp(m, "3.1.14",
       [("Fact: numeric total_price cells", "=COUNT(tTransactions[total_price])", len(TX), "int", "Facts hold measures and keys."),
        ("Dimension: product categories", "=COUNTA(UNIQUE(tProducts[category]))", int(P.category.nunique()), "int", "Dimensions hold descriptions.")],
       [("n", "How many membership tiers does tCustomers describe?", "=COUNTA(UNIQUE(tCustomers[membership_tier]))", int(C.membership_tier.nunique()),
         "int", "Membership tiers"),
        ("t", "Is tDeliveries a fact or a dimension table here? (it records events with dates)", "fact", None, None, ("fact table",))],
       steps=["Facts: many rows, keys + numbers (transactions).", "Dimensions: fewer rows, descriptive text (products, customers, sellers).",
              "Relationships go from a dimension's key (one) to the fact's foreign key (many)."])
    tp(m, "3.1.15",
       [("OLTP: lines recorded one at a time", "=ROWS(tTransactions)", len(TX), "int", "Each checkout writes one line."),
        ("OLAP: net revenue, all lines summarised", "=" + net(t="tTransactions"), f(D.NT.total_price.sum()), N, "One query reads every line.")],
       [("n", "OLAP question: net revenue for 2025 (Delivered + Completed)?", "=" + net(order_year=2025), f(D.NT[D.NT.order_year == 2025].total_price.sum()),
         N, "Net revenue 2025"),
        ("t", "Is a checkout system that saves one order at a time OLTP or OLAP?", "OLTP", None, None, ())],
       steps=["OLTP: many small writes, normalised tables, current data.", "OLAP: few big reads, star schemas, history.",
              "Analysts query a copy (or the Data Model) so reports never slow the checkout."])
    m.finish()


# =====================================================================================================
def m32(b, D):
    TX, P, C, DV, L, NT = D.TX, D.P, D.C, D.DV, D.L, D.NT
    m = ModuleSheet(b, "3.2 SQL and SQLite", "3.2  Introduction to SQL and SQLite",
                    "Tutor walk-through after Module 3.2: run each query in DB Browser for SQLite, then check it against the worksheet twin",
                    "3.2.1 - 3.2.17")
    m.dataset([TX_DS, LI_DS, ("tProducts / tCustomers / tDeliveries", "Data_ sheets (A1)", "Used in the WHERE, DISTINCT and subquery examples."),
               ("ETEH_Week3_SQL_Examples.sql", "Week3_Data_Files", "Every query on the slides; open it in DB Browser's Execute SQL tab.")])
    gross = f(TX.total_price.sum())
    tp(m, "3.2.1",
       [("order_lines (COUNT(*))", "=ROWS(tTransactions)", len(TX), "int", ""),
        ("gross_value (SUM(total_price))", "=SUM(tTransactions[total_price])", gross, N, "Same as the SQL result.")],
       [("n", "Write SELECT SUM(commission_amount) FROM transactions; What does it return?", "=SUM(tTransactions[commission_amount])",
         f(TX.commission_amount.sum()), N, "Total commission"),
        ("n", "SELECT AVG(total_price) FROM transactions; (round to the naira)", "=AVERAGE(tTransactions[total_price])", f(TX.total_price.mean()), N,
         "Average line value")],
       steps=["DB Browser › Execute SQL tab.", "Type the query; press F5 (or Ctrl+Enter) to run it.", "Read the result grid; compare with the twin below."],
       code=True)
    tp(m, "3.2.2",
       [("First 4 products (SQLite: LIMIT 4)", "=ROWS(tProducts)", len(P), "int", "SQL Server writes SELECT TOP 4 … instead."),
        ("Lines in 2025 (strftime in SQLite, YEAR() in SQL Server)", "=COUNTIF(tLines[order_year],2025)", int((L.order_year == 2025).sum()), "int", "")],
       [("t", "SQLite uses LIMIT 5. What keyword does SQL Server use instead?", "TOP", None, None, ("top 5", "select top 5")),
        ("n", "How many lines are dated 2024? (strftime('%Y', order_date) = '2024')", "=COUNTIF(tLines[order_year],2024)", int((L.order_year == 2024).sum()),
         "int", "Lines in 2024")],
       steps=["Core SQL (SELECT, WHERE, GROUP BY, JOIN) is the same in every dialect.", "Row limits and date functions differ: LIMIT/TOP, strftime/YEAR.",
              "Write the core first; look up the dialect's function when you move databases."])
    tp(m, "3.2.3",
       [("Tables in supermart_online.db (rows)", "=ROWS(tTransactions)+ROWS(tCustomers)+ROWS(tProducts)+ROWS(tSellers)+ROWS(tDeliveries)+ROWS(tStates)",
         len(TX) + len(C) + len(P) + len(D.SE) + len(DV) + len(D.ST), "int", "One file, no server."),
        ("Mercyland 2023 admissions", "=ROWS(tAdm23)", len(D.A), "int", "mercyland_hospital.db, filtered to 2023.")],
       [("t", "Does SQLite need a database server running? (Yes / No)", "No", None, None, ("n",))],
       steps=["Find Week3_Data_Files › supermart_online.db and mercyland_hospital.db.", "Each is one ordinary file: copy, email or back it up like a workbook.",
              "Open either in DB Browser; nothing to install on a server."])
    tp(m, "3.2.4",
       [("Rows any tool returns for SELECT COUNT(*) FROM transactions", "=ROWS(tTransactions)", len(TX), "int",
         "DB Browser, sqlite3 and Excel (ODBC) read the same file.")],
       [("n", "Run SELECT COUNT(*) FROM customers; in your tool. What do you get?", "=ROWS(tCustomers)", len(C), "int", "Customers")],
       steps=["DB Browser: point and click, best for learning.", "sqlite3 command line: sqlite3 supermart_online.db, then type SQL.",
              "Excel: Data › Get Data › From Other Sources › From ODBC (3.3)."],
       do=[("Open supermart_online.db in DB Browser and run SELECT COUNT(*) FROM customers;", "2,500")])
    p4 = P.head(4)
    tp(m, "3.2.5",
       [("First product_id", "=INDEX(tProducts[product_id],1)", p4.product_id.iloc[0], "gen", "LIMIT 4 returns the first four rows."),
        ("Its name", "=INDEX(tProducts[product_name],1)", p4.product_name.iloc[0], "gen", ""),
        ("4th product's price", "=INDEX(tProducts[price],4)", f(p4.price.iloc[3]), N, "")],
       [("t", "Run the query. What is the product_name of the 3rd row?", p4.product_name.iloc[2], "=INDEX(tProducts[product_name],3)", "3rd product"),
        ("n", "SELECT price FROM products WHERE product_id = 'PROD00010'; What is it?",
         '=XLOOKUP("PROD00010",tProducts[product_id],tProducts[price])', f(P.set_index("product_id").price["PROD00010"]), N, "Price of PROD00010")],
       steps=["SELECT lists the columns, FROM names the table.", "Name columns rather than SELECT * so the result stays the same if the table changes.",
              "LIMIT keeps the preview small."], code=True)
    tp(m, "3.2.6",
       [("delivery_id IS NULL", "=COUNTBLANK(tTransactions[delivery_id])", int(TX.delivery_id.isna().sum()), "int", "NULL is a blank cell here."),
        ("delivery_id = NULL", "=0", 0, "int", "Never true in SQL: use IS NULL."),
        ("Cancelled lines", '=COUNTIF(tTransactions[order_status],"Cancelled")', int((TX.order_status == "Cancelled").sum()), "int",
         "The same 570: no delivery for cancelled orders.")],
       [("n", "SELECT COUNT(*) FROM transactions WHERE order_status = 'Returned';", '=COUNTIF(tTransactions[order_status],"Returned")',
         int((TX.order_status == "Returned").sum()), "int", "Returned lines"),
        ("n", "SELECT COUNT(*) FROM products WHERE category = 'Electronics' AND price > 500000;",
         '=COUNTIFS(tProducts[category],"Electronics",tProducts[price],">500000")', int(((P.category == "Electronics") & (P.price > 500000)).sum()),
         "int", "Electronics above ₦500k")],
       steps=["WHERE keeps rows where the condition is true.", "Text in single quotes: order_status = 'Returned'.",
              "Test for missing values with IS NULL / IS NOT NULL."], code=True)
    gp = C.membership_tier.isin(["Gold", "Platinum"])
    tp(m, "3.2.7",
       [("Gold or Platinum, South-West", '=COUNTIFS(tCustomers[membership_tier],"Gold",tCustomers[zone],"South-West")+'
                                         'COUNTIFS(tCustomers[membership_tier],"Platinum",tCustomers[zone],"South-West")',
         int((gp & (C.zone == "South-West")).sum()), "int", "Brackets: (Gold OR Platinum) AND South-West."),
        ("Without brackets", '=COUNTIF(tCustomers[membership_tier],"Gold")+COUNTIFS(tCustomers[membership_tier],"Platinum",tCustomers[zone],"South-West")',
         int(((C.membership_tier == "Gold") | ((C.membership_tier == "Platinum") & (C.zone == "South-West"))).sum()), "int",
         "AND binds first: every Gold customer slips in.")],
       [("n", "How many customers are NOT Bronze and live in Lagos?", '=COUNTIFS(tCustomers[membership_tier],"<>Bronze",tCustomers[state],"Lagos")',
         int(((C.membership_tier != "Bronze") & (C.state == "Lagos")).sum()), "int", "Lagos, not Bronze"),
        ("n", "Female customers who are Gold or Platinum?", '=COUNTIFS(tCustomers[gender],"Female",tCustomers[membership_tier],"Gold")+'
                                                         'COUNTIFS(tCustomers[gender],"Female",tCustomers[membership_tier],"Platinum")',
         int(((C.gender == "Female") & gp).sum()), "int", "Female Gold or Platinum")],
       steps=["AND: both true. OR: either true. NOT: reverses.", "AND is worked out before OR: always bracket OR groups.",
              "COUNTIFS is AND only, so an OR becomes two COUNTIFS added."], code=True)
    top = TX.sort_values("total_price", ascending=False)
    tp(m, "3.2.8",
       [("Largest line", "=LARGE(tTransactions[total_price],1)", f(top.total_price.iloc[0]), N, "ORDER BY total_price DESC, row 1."),
        ("Its order_id", "=XLOOKUP(LARGE(tTransactions[total_price],1),tTransactions[total_price],tTransactions[order_id])", top.order_id.iloc[0],
         "gen", ""),
        ("3rd largest", "=LARGE(tTransactions[total_price],3)", f(top.total_price.iloc[2]), N, "")],
       [("n", "ORDER BY total_price ASC LIMIT 1: what is the smallest line value?", "=SMALL(tTransactions[total_price],1)", f(TX.total_price.min()), N,
         "Smallest line"),
        ("n", "What is the 2nd largest line value?", "=LARGE(tTransactions[total_price],2)", f(top.total_price.iloc[1]), N, "2nd largest line")],
       steps=["ORDER BY column DESC = largest first; ASC (default) = smallest first.", "Sort on several columns: ORDER BY zone, total_price DESC.",
              "Without ORDER BY, row order is not guaranteed."], code=True)
    tp(m, "3.2.9",
       [("Rows without LIMIT", "=ROWS(tTransactions)", len(TX), "int", "SELECT * FROM transactions;"),
        ("Rows with LIMIT 10", "=MIN(10,ROWS(tTransactions))", 10, "int", "Preview only.")],
       [("n", "SELECT * FROM deliveries LIMIT 25; how many rows come back?", "=MIN(25,ROWS(tDeliveries))", 25, "int", "LIMIT 25"),
        ("t", "LIMIT 5 OFFSET 5 returns which rows? (e.g. 6-10)", "6-10", None, None, ("6 to 10", "rows 6-10"))],
       steps=["Add LIMIT n to preview a big table.", "OFFSET skips rows: LIMIT 5 OFFSET 5 = rows 6-10.", "Use ORDER BY with LIMIT for a 'top n'."])
    cat = NT.groupby("category").total_price.agg(["size", "sum"]).sort_values("sum", ascending=False)
    tp(m, "3.2.10",
       [(f"{cat.index[0]}: lines", "=" + netn(category=cat.index[0]), int(cat["size"].iloc[0]), "int", "COUNT(*) for the group."),
        (f"{cat.index[0]}: revenue", "=" + net(category=cat.index[0]), f(cat["sum"].iloc[0]), N, "SUM(total_price)."),
        (f"{cat.index[1]}: revenue", "=" + net(category=cat.index[1]), f(cat["sum"].iloc[1]), N, "")],
       [("n", "Add GROUP BY on Food: what is Food's net revenue?", "=" + net(category="Food"), f(cat.loc["Food", "sum"]), N, "Food net revenue"),
        ("n", "How many net lines does Beauty have?", "=" + netn(category="Beauty"), int(cat.loc["Beauty", "size"]), "int", "Beauty net lines")],
       steps=["GROUP BY returns one row per group.", "Every selected column is either grouped or aggregated (COUNT, SUM, AVG).",
              "The twin: SUMIFS by category on tLines, Delivered + Completed."], code=True)
    sell = NT.groupby("seller_id").total_price.sum().sort_values(ascending=False)
    tp(m, "3.2.11",
       [("SELL0200 net revenue", "=" + net(t="tTransactions", seller_id="SELL0200"), f(sell["SELL0200"]), N, "Top of the HAVING list."),
        ("Sellers above ₦15m", f"=SUMPRODUCT(--(SUMIFS(tLines[total_price],tLines[seller_id],tSellers[seller_id],tLines[order_status],\"Delivered\")"
                               f"+SUMIFS(tLines[total_price],tLines[seller_id],tSellers[seller_id],tLines[order_status],\"Completed\")>15000000))",
         int((sell > 15e6).sum()), "int", "Rows the HAVING clause keeps.")],
       [("n", "What is SELL0098's net revenue?", "=" + net(t="tTransactions", seller_id="SELL0098"), f(sell["SELL0098"]), N, "SELL0098 net revenue"),
        ("t", "Which clause filters rows before grouping: WHERE or HAVING?", "WHERE", None, None, ())],
       steps=["WHERE filters rows before grouping; HAVING filters the groups after.", "HAVING uses aggregates: HAVING SUM(total_price) > 15000000.",
              "ORDER BY and LIMIT then pick the top three."], code=True)
    tp(m, "3.2.12",
       [("lines_checked", "=ROWS(tTransactions)", len(TX), "int", ""),
        ("mismatches", "=SUMPRODUCT(--(ABS(tTransactions[total_price]*tTransactions[commission_percentage]/100-tTransactions[commission_amount])>0.01))",
         int(((TX.total_price * TX.commission_percentage / 100 - TX.commission_amount).abs() > 0.01).sum()), "int", "Calculated column, checked.")],
       [("n", "SELECT SUM(quantity * unit_price) AS check_total FROM transactions; does it equal SUM(total_price)? Type the check_total.",
         "=SUMPRODUCT(tTransactions[quantity],tTransactions[unit_price])", f((TX.quantity * TX.unit_price).sum()), N, "SUM(quantity × unit_price)"),
        ("t", "In SELECT total_price * 0.075 AS vat, what is 'vat' called?", "alias", None, None, ("an alias", "column alias"))],
       steps=["AS gives a column (or table) a short name: an alias.", "Calculated columns do arithmetic per row: quantity * unit_price.",
              "Use 100.0, not 100, so SQLite divides with decimals."], code=True)
    tp(m, "3.2.13",
       [("deliveries", "=ROWS(tDeliveries)", len(DV), "int", ""),
        ("couriers (COUNT DISTINCT)", "=COUNTA(UNIQUE(tDeliveries[courier_name]))", int(DV.courier_name.nunique()), "int", ""),
        ("states (COUNT DISTINCT)", "=COUNTA(UNIQUE(tDeliveries[state]))", int(DV.state.nunique()), "int", "")],
       [("n", "SELECT COUNT(DISTINCT zone) FROM deliveries;", "=COUNTA(UNIQUE(tDeliveries[zone]))", int(DV.zone.nunique()), "int", "Distinct zones"),
        ("n", "How many distinct customers placed at least one order? (COUNT(DISTINCT customer_id) FROM transactions)",
         "=COUNTA(UNIQUE(tTransactions[customer_id]))", int(TX.customer_id.nunique()), "int", "Distinct customers")],
       steps=["SELECT DISTINCT col returns each value once.", "COUNT(DISTINCT col) counts them.", "The twin: COUNTA(UNIQUE(...))."], code=True)
    bands = pd.cut(TX.total_price, [0, 50000, 250000, 1000000, 1e12], right=False).value_counts().sort_index()
    tp(m, "3.2.14",
       [("1 <50k: lines", '=COUNTIF(tTransactions[total_price],"<50000")', int(bands.iloc[0]), "int", "WHEN total_price < 50000"),
        ("2 50k-250k: lines", '=COUNTIFS(tTransactions[total_price],">=50000",tTransactions[total_price],"<250000")', int(bands.iloc[1]), "int",
         "CASE stops at the first true WHEN."),
        ("4 1m+: lines", '=COUNTIF(tTransactions[total_price],">=1000000")', int(bands.iloc[3]), "int", "ELSE")],
       [("n", "How many lines fall in '3 250k-1m'?", '=COUNTIFS(tTransactions[total_price],">=250000",tTransactions[total_price],"<1000000")',
         int(bands.iloc[2]), "int", "Band 3 lines"),
        ("n", "What is the value of the '4 1m+' band?", '=SUMIF(tTransactions[total_price],">=1000000")', f(TX[TX.total_price >= 1e6].total_price.sum()),
         N, "Band 4 value")],
       steps=["CASE WHEN … THEN … ELSE … END builds a new column.", "Order the WHENs: the first true one wins.",
              "GROUP BY the CASE column to count each band."], code=True)
    spend = TX.groupby("customer_id").total_price.sum()
    tp(m, "3.2.15",
       [("Buyers", "=COUNTA(UNIQUE(tTransactions[customer_id]))", len(spend), "int", "Inner query: one row per customer."),
        ("Average spend per buyer", "=SUM(tTransactions[total_price])/COUNTA(UNIQUE(tTransactions[customer_id]))", f(spend.mean()), N,
         "The AVG(spend) subquery."),
        ("Buyers above average", "=SUMPRODUCT(--(SUMIFS(tTransactions[total_price],tTransactions[customer_id],tCustomers[customer_id])>"
                                 "SUM(tTransactions[total_price])/COUNTA(UNIQUE(tTransactions[customer_id]))))", int((spend > spend.mean()).sum()), "int", "")],
       [("n", "How many products are priced above the average product price? (WHERE price > (SELECT AVG(price) FROM products))",
         "=COUNTIF(tProducts[price],\">\"&AVERAGE(tProducts[price]))", int((P.price > P.price.mean()).sum()), "int", "Products above average price")],
       steps=["Write and run the inner query on its own first.", "Wrap it in brackets inside the outer query.",
              "A subquery in WHERE must return one value; in FROM it acts as a table."], code=True)
    yr = TX.assign(y=TX.order_date.dt.year).groupby("y").total_price.agg(["size", "sum"])
    tp(m, "3.2.16",
       [("2020 lines", "=COUNTIF(tLines[order_year],2020)", int(yr.loc[2020, "size"]), "int", "strftime('%Y', order_date) = '2020'"),
        ("2025 value", "=SUMIFS(tLines[total_price],tLines[order_year],2025)", f(yr.loc[2025, "sum"]), N, "All statuses, as in the query."),
        ("Lines in May 2025 (order_month)", '=COUNTIF(tLines[order_month],"2025-05")', int((L.order_month == "2025-05").sum()), "int",
         "strftime('%Y-%m', order_date)")],
       [("n", "What is the 2023 value (all statuses)?", "=SUMIFS(tLines[total_price],tLines[order_year],2023)", f(yr.loc[2023, "sum"]), N, "2023 value"),
        ("n", "How many products have a name starting 'Samsung'? (WHERE product_name LIKE 'Samsung%')", '=COUNTIF(tProducts[product_name],"Samsung*")',
         int(P.product_name.str.lower().str.startswith("samsung").sum()), "int", "Samsung products")],
       steps=["strftime('%Y', date) gives the year as text; '%Y-%m' the month.", "UPPER, LOWER, TRIM, SUBSTR and LIKE work on text.",
              "LIKE 'A%' = starts with A; % is any run of characters (Excel's *)."], code=True)
    zz = NT.groupby("customer_zone").agg(n=("total_price", "size"), s=("total_price", "sum"), d=("days_to_deliver", "mean")).sort_values("s", ascending=False)
    tp(m, "3.2.17",
       [("South-West orders", "=" + netn(customer_zone="South-West"), int(zz.loc["South-West", "n"]), "int", ""),
        ("South-West net revenue", "=" + net(customer_zone="South-West"), f(zz.loc["South-West", "s"]), N, ""),
        ("South-West avg days to deliver", '=ROUND((SUMIFS(tLines[days_to_deliver],tLines[customer_zone],"South-West",tLines[order_status],"Delivered")+'
                                           'SUMIFS(tLines[days_to_deliver],tLines[customer_zone],"South-West",tLines[order_status],"Completed"))/'
                                           '(COUNTIFS(tLines[customer_zone],"South-West",tLines[order_status],"Delivered",tLines[days_to_deliver],"<>")+'
                                           'COUNTIFS(tLines[customer_zone],"South-West",tLines[order_status],"Completed",tLines[days_to_deliver],"<>")),1)',
         round(float(zz.loc["South-West", "d"]), 1), "dec1",
         "AVG skips NULL days, so the count must skip blank days too.")],
       [("n", "What is the North-East net revenue?", "=" + net(customer_zone="North-East"), f(zz.loc["North-East", "s"]), N, "North-East net revenue"),
        ("n", "How many net orders came from the South-East?", "=" + netn(customer_zone="South-East"), int(zz.loc["South-East", "n"]), "int",
         "South-East net orders")],
       steps=["Clause order: SELECT, FROM, WHERE, GROUP BY, HAVING, ORDER BY, LIMIT.", "Build it in steps, running after each clause.",
              "Check the totals against a twin before sharing."], code=True)
    m.finish()


# =====================================================================================================
def m33(b, D):
    TX, L, NT = D.TX, D.L, D.NT
    m = ModuleSheet(b, "3.3 ODBC", "3.3  Connecting Excel to Databases with ODBC",
                    "Tutor walk-through after Module 3.3: install the driver, connect, import tables or SQL, refresh and parameterise",
                    "3.3.1 - 3.3.9")
    m.dataset([TX_DS, LI_DS, ("SQLite3 ODBC Driver", "ch-werner.de/sqliteodbc (64-bit for 64-bit Excel)", "The driver Excel needs to read .db files.")])
    tp(m, "3.3.1",
       [("Rows delivered through the connection", "=ROWS(tTransactions)", len(TX), "int", "Excel › ODBC driver › database file.")],
       [("t", "Which layer translates Excel's request into the database's own language? (Excel / driver / database)", "driver", None, None,
         ("odbc driver", "the driver"))],
       steps=["Excel (the app) asks for data.", "The ODBC driver translates the request for that database.", "The database returns rows."],
       do=[("Check Excel's bitness: File › Account › About Excel.", "'64-bit' or '32-bit' in the first line.")])
    tp(m, "3.3.2",
       [("Columns that arrived", "=COLUMNS(tTransactions)", 15, "int", "Proof the driver works.")],
       [("t", "64-bit Excel needs which driver bitness? (32-bit / 64-bit)", "64-bit", None, None, ("64", "64 bit"))],
       steps=["Download the SQLite3 ODBC Driver installer that matches Excel's bitness.", "Install it (admin rights may be needed).",
              "Open ODBC Data Sources (64-bit) › Drivers tab: 'SQLite3 ODBC Driver' is listed."],
       do=[("Open 'ODBC Data Sources (64-bit)' from the Start menu › Drivers tab.", "SQLite3 ODBC Driver appears with a version number.")])
    tp(m, "3.3.3",
       [("Rows via a DSN-less string", "=ROWS(tTransactions)", len(TX), "int", "Same data either way; only where the settings live differs.")],
       [("t", "Which travels with the workbook: a DSN or a DSN-less connection string?", "DSN-less", None, None, ("dsn less", "connection string"))],
       steps=["DSN: a named connection saved in Windows on each PC.", "DSN-less: the full connection string saved in the workbook.",
              "Edit the Database= path to your own folder."], code=True)
    tp(m, "3.3.4",
       [("Rows of transactions loaded", "=ROWS(tTransactions)", len(TX), "int", "Navigator › transactions › Transform Data."),
        ("Dates after Power Query typing", "=COUNT(tTransactions[order_date])", len(TX), "int", "Text dates set to Date before loading.")],
       [("n", "Load the products table through ODBC. How many rows arrive?", "=ROWS(tProducts)", len(D.P), "int", "Products rows")],
       steps=["Data › Get Data › From Other Sources › From ODBC.", "Choose (None) for DSN, open Advanced options, paste the connection string.",
              "Navigator: tick transactions › Transform Data; set order_date to Date › Close & Load."],
       do=[("Connect to supermart_online.db and load transactions as a table.", "A green table of 11,000 rows; Queries & Connections lists the query.")])
    tp(m, "3.3.5",
       [("Credentials stored in this workbook", "=0", 0, "int", "SQLite files have no password; server databases use Windows or database sign-in.")],
       [("t", "Where does Power Query keep the credentials you type? (in the workbook / on your PC)", "on your PC", None, None,
         ("pc", "my pc", "data source settings"))],
       steps=["Prefer Windows authentication on company servers.", "Never put a password in the connection string or SQL.",
              "Data › Get Data › Data Source Settings shows and clears saved credentials."])
    n25 = int(((NT.order_year == 2025)).sum())
    tp(m, "3.3.6",
       [("Whole table: rows that travel", "=ROWS(tTransactions)", len(TX), "int", ""),
        ("SQL query: 2025 net lines only", "=" + netn(order_year=2025), n25, "int", "WHERE in the SQL: the database filters before sending."),
        ("Share of the table", "=(" + netn(order_year=2025) + ")/ROWS(tTransactions)", n25 / len(TX), "pct", "")],
       [("n", "A query returns only Electronics net lines. How many rows travel?", "=" + netn(category="Electronics"), int((NT.category == "Electronics").sum()),
         "int", "Electronics net lines"),
        ("t", "Data_Lines holds a view imported through ODBC. Which view?", "v_order_lines", None, None, ())],
       steps=["From ODBC › Advanced options › SQL statement: paste a SELECT.", "The database runs the WHERE and GROUP BY; Excel receives the result.",
              "Import a whole table only when you need every row."])
    tp(m, "3.3.7",
       [("Rows now", "=ROWS(tTransactions)", len(TX), "int", "After a refresh, compare with the database's COUNT(*)."),
        ("Latest order date", "=MAX(tTransactions[order_date])", f((TX.order_date.max() - pd.Timestamp("1899-12-30")).days), "date",
         "A refresh brings later dates when the database grows.")],
       [("t", "Which button re-runs every query in the workbook? (Data tab)", "Refresh All", None, None, ("refresh all",))],
       steps=["Data › Refresh All, or right-click the table › Refresh.", "Query Properties: Refresh every n minutes / Refresh when opening the file.",
              "After refreshing, check the row count against the database."])
    year_in = 2024
    tp(m, "3.3.8",
       lambda mm: [("Lines in the ParamYear year", f"=COUNTIF(tLines[order_year],{mm._param_addr})", int((L.order_year == year_in).sum()), "int",
                     "Change the year in the yellow cell: the count follows, like a query parameter.")],
       [("n", "Type 2022 as the parameter value. How many lines does the query return?", "=COUNTIF(tLines[order_year],2022)",
         int((L.order_year == 2022).sum()), "int", "Lines in 2022")],
       steps=["Put the parameter in a named cell (for example ParamYear).", "In Power Query, read it with Excel.CurrentWorkbook(){[Name=\"ParamYear\"]}[Content]{0}[Column1].",
              "Build the WHERE from it, or filter after the source step so query folding sends it to the database."],
       extra=lambda mm: _param(mm, year_in))
    tp(m, "3.3.9",
       [("Lines in 2025 (right: compare dates as ISO text)", "=COUNTIF(tLines[order_year],2025)", int((L.order_year == 2025).sum()), "int",
         "WHERE order_date >= '2025-01-01' works on ISO text."),
        ("Lines matched by a dd/mm/yyyy text compare", "=0", 0, "int", "WHERE order_date >= '01/01/2025' compares text: wrong rows.")],
       [("t", "Error: 'Data source name not found and no default driver specified'. What is the usual cause? (driver bitness / wrong SQL)",
         "driver bitness", None, None, ("bitness", "32/64-bit driver", "wrong driver"))],
       steps=["'Data source name not found': driver missing or the wrong bitness.", "'unable to open database file': check the Database= path.",
              "Wrong numbers, no error: check date formats and filters in the SQL."])
    m.finish()


def _param(m, year):
    m.h("Parameter cell")
    out = m.s.inputs([("ParamYear (type a year 2020-2025)", year, "gen", "Yellow input cell")])
    m.__dict__["_param_addr"] = list(out.values())[0]


# =====================================================================================================
def m34(b, D):
    TX, DV, C, SE, P, L = D.TX, D.DV, D.C, D.SE, D.P, D.L
    m = ModuleSheet(b, "3.4 Merge & Append", "3.4  Merging and Appending in Power Query",
                    "Tutor walk-through after Module 3.4: six join kinds with real row counts, fuzzy matching, expand or aggregate, the "
                    "Cartesian trap and appending files", "3.4.1 - 3.4.16")
    m.dataset([TX_DS, ("tDeliveries", "Sheet Data_Deliveries (A1)", "10,430 deliveries: the right-hand table in the join examples."),
               ("tFuzzy / tMonthFiles / tZoneRaw", "On this sheet (3.4.8, 3.4.11, 3.4.14)", "Small tables made for the exercises."),
               ("Week2_Data_Files", "01_Section_Examples › Week2_Data_Files", "The 12 monthly CSVs and the dirty sales file for appending and key cleaning.")])
    has_del = TX.delivery_id.notna()
    tp(m, "3.4.1",
       [("Merge: transactions + deliveries (rows kept, left join)", "=ROWS(tTransactions)", len(TX), "int", "Merge adds columns."),
        ("Append: 12 monthly files (rows)", "=ROWS(tDash)", len(D.DS), "int", "Append adds rows: 3,312 FY2025 lines.")],
       [("t", "Stacking January's file under December's is a merge or an append?", "append", None, None, ("an append",)),
        ("t", "Adding the courier name to each order line is a merge or an append?", "merge", None, None, ("a merge",))],
       steps=["Merge: Home › Merge Queries; choose the second table and the key columns.", "Append: Home › Append Queries; choose the tables to stack.",
              "Ask: am I adding columns (merge) or rows (append)?"])
    inner_v = f(TX[has_del].total_price.sum())
    tp(m, "3.4.2",
       [("Inner join rows", "=COUNTA(tTransactions[delivery_id])", int(has_del.sum()), "int", "Only lines with a matching delivery."),
        ("Inner join value", '=SUMIF(tTransactions[delivery_id],"<>",tTransactions[total_price])', inner_v, N, ""),
        ("Lines dropped", "=ROWS(tTransactions)-COUNTA(tTransactions[delivery_id])", int((~has_del).sum()), "int", "The 570 cancellations.")],
       [("n", "What value do the dropped lines carry?", '=SUMIF(tTransactions[delivery_id],"",tTransactions[total_price])',
         f(TX[~has_del].total_price.sum()), N, "Value dropped by the inner join")],
       steps=["Merge Queries: transactions (top), deliveries (bottom), key delivery_id on both.", "Join Kind: Inner (only matching rows).",
              "Read the footer: 'The selection matches 10,430 of 11,000 rows'."], code=True)
    tp(m, "3.4.3",
       [("all_orders", "=ROWS(tTransactions)", len(TX), "int", "Left outer keeps every left row."),
        ("with_courier", "=COUNTA(tLines[courier_name])", int(L.courier_name.notna().sum()), "int", ""),
        ("courier_null", "=COUNTBLANK(tLines[courier_name])", int(L.courier_name.isna().sum()), "int", "null where no delivery matched.")],
       [("n", "How many order lines have a delivery_status of 'Delivered' after the left join?", '=COUNTIF(tLines[delivery_status],"Delivered")',
         int((L.delivery_status == "Delivered").sum()), "int", "delivery_status Delivered")],
       steps=["Same merge, Join Kind: Left Outer.", "Expand the new column: tick courier_name, delivery_status.",
              "Rows without a match show null: keep them, they are real orders."], code=True)
    tp(m, "3.4.4",
       [("Deliveries kept by a right join", "=ROWS(tDeliveries)", len(DV), "int", "Every delivery, matched or not."),
        ("Deliveries with no order line", "=SUMPRODUCT(--(COUNTIF(tTransactions[delivery_id],tDeliveries[delivery_id])=0))",
         int((~DV.delivery_id.isin(TX.delivery_id)).sum()), "int", "")],
       [("t", "A right join of transactions to deliveries gives the same rows as which join with the tables swapped?", "left join", None, None,
         ("left outer", "left outer join", "left"))],
       steps=["Join Kind: Right Outer keeps every row of the second table.", "Most people swap the tables and use Left Outer instead: same result.",
              "Check the row count against the second table."])
    tp(m, "3.4.5",
       [("Full outer rows", "=ROWS(tTransactions)+SUMPRODUCT(--(COUNTIF(tTransactions[delivery_id],tDeliveries[delivery_id])=0))",
         len(TX) + int((~DV.delivery_id.isin(TX.delivery_id)).sum()), "int", "Every row of both, matched where possible.")],
       [("n", "How many rows of a full outer join have no delivery? (left side only)", "=COUNTBLANK(tTransactions[delivery_id])",
         int((~has_del).sum()), "int", "Unmatched left rows")],
       steps=["Join Kind: Full Outer.", "Unmatched rows from either side carry null on the other side.",
              "Use it for reconciliation: what is in one system and not the other?"])
    canc = TX[~has_del].groupby("payment_status").total_price.agg(["size", "sum"])
    tp(m, "3.4.6",
       [("Left anti: lines with no delivery", "=COUNTBLANK(tTransactions[delivery_id])", int((~has_del).sum()), "int", ""),
        ("… Failed payments", '=COUNTIFS(tTransactions[delivery_id],"",tTransactions[payment_status],"Failed")', int(canc.loc["Failed", "size"]), "int", ""),
        ("… Refunded", '=COUNTIFS(tTransactions[delivery_id],"",tTransactions[payment_status],"Refunded")', int(canc.loc["Refunded", "size"]), "int", "")],
       [("n", "What value do the Refunded anti-join lines carry?", '=SUMIFS(tTransactions[total_price],tTransactions[delivery_id],"",tTransactions[payment_status],"Refunded")',
         f(canc.loc["Refunded", "sum"]), N, "Refunded, no delivery")],
       steps=["Join Kind: Left Anti returns left rows with no match.", "It is the fastest way to list exceptions.", "Group the result to explain it."], code=True)
    no_ord_c = int((~C.customer_id.isin(TX.customer_id)).sum())
    tp(m, "3.4.7",
       [("Customers with no order line", "=SUMPRODUCT(--(COUNTIF(tTransactions[customer_id],tCustomers[customer_id])=0))", no_ord_c, "int",
         "Right anti: customers (right) with no transactions (left)."),
        ("Sellers with no order line", "=SUMPRODUCT(--(COUNTIF(tTransactions[seller_id],tSellers[seller_id])=0))", int((~SE.seller_id.isin(TX.seller_id)).sum()),
         "int", "")],
       [("n", "How many products have no order line?", "=SUMPRODUCT(--(COUNTIF(tTransactions[product_id],tProducts[product_id])=0))",
         int((~P.product_id.isin(TX.product_id)).sum()), "int", "Products with no line")],
       steps=["Join Kind: Right Anti returns right rows with no match.", "Here: customers who signed up but never ordered.",
              "A marketing list, not an error."])
    FZ = M35.FZ.reset_index(drop=True)
    ci = int((FZ.typed.str.lower() == FZ.store_name.str.lower()).sum())
    tr = int((FZ.typed.str.split().str.join(" ").str.lower() == FZ.store_name.str.lower()).sum())

    def fuzzy_table(mm):
        mm.h("tFuzzy: a supplier's hand-typed seller names")
        mm.s.table(FZ[["seller_id", "store_name", "typed"]], "tFuzzy", widths={"seller_id": 10, "store_name": 26, "typed": 26})
        mm.s.r += 1

    tp(m, "3.4.8",
       [("Equal in Excel (= ignores case)", "=SUMPRODUCT(--(tFuzzy[typed]=tFuzzy[store_name]))", ci, "int",
         "Excel's = is not case-sensitive; Power Query's exact match is."),
        ("Equal after TRIM (case and spaces ignored)", "=SUMPRODUCT(--(TRIM(tFuzzy[typed])=tFuzzy[store_name]))", tr, "int",
         "The slide's 'Ignore case and spaces': 7 of 8.")],
       [("n", "How many typed names fail even after TRIM? (These need fuzzy matching or a person.)",
         "=ROWS(tFuzzy)-SUMPRODUCT(--(TRIM(tFuzzy[typed])=tFuzzy[store_name]))", len(FZ) - tr, "int", "Still unmatched"),
        ("t", "Which typed name is still unmatched?", FZ[FZ.typed.str.split().str.join(" ").str.lower() != FZ.store_name.str.lower()].typed.iloc[0],
         None, None, ())],
       steps=["Merge Queries › tick 'Use fuzzy matching to perform the merge'.", "Fuzzy options: Similarity threshold 0.8, Ignore case, Match by combining text parts.",
              "Review every fuzzy match before trusting it."], code=True, extra=fuzzy_table)
    lines = TX.groupby("customer_id").size()
    tp(m, "3.4.9",
       [("Expand: rows after merging customers with transactions", "=ROWS(tTransactions)", len(TX), "int", "One row per order line."),
        ("Aggregate: rows (one per customer)", "=ROWS(tCustomers)", len(C), "int", "Count or Sum of the nested table."),
        ("Most lines for one customer", "=MAX(COUNTIF(tTransactions[customer_id],tCustomers[customer_id]))", int(lines.max()), "int", "")],
       [("n", "Aggregate (Sum of total_price) for CUST00985: what is it?", '=SUMIF(tTransactions[customer_id],"CUST00985",tTransactions[total_price])',
         f(TX[TX.customer_id == "CUST00985"].total_price.sum()), N, "CUST00985 total"),
        ("n", "Average order lines per customer who ordered?", "=ROWS(tTransactions)/COUNTA(UNIQUE(tTransactions[customer_id]))", f(len(TX) / TX.customer_id.nunique()),
         "dec1", "Lines per buyer")],
       steps=["After a merge, click the expand icon on the new column.", "Expand: one row per match (rows can multiply).",
              "Aggregate: one row kept; choose Sum, Count, Min, Max of the nested rows."])
    cart = int((TX.groupby("customer_id").size().reindex(DV.customer_id.unique()).fillna(0) *
                DV.groupby("customer_id").size()).sum())
    tp(m, "3.4.10",
       [("Join on order_id (correct key)", "=ROWS(tTransactions)", len(TX), "int", "One delivery per order line."),
        ("Join on customer_id (wrong key)", "=SUMPRODUCT(COUNTIF(tTransactions[customer_id],tDeliveries[customer_id]))", cart, "int",
         "Every line meets every delivery of the same customer.")],
       [("n", "How many times bigger is the wrong join than the transactions table? (1 decimal)",
         "=SUMPRODUCT(COUNTIF(tTransactions[customer_id],tDeliveries[customer_id]))/ROWS(tTransactions)", cart / len(TX), "dec1", "Blow-up factor")],
       steps=["Merge on a key that is unique on at least one side.", "Compare the row count after the merge with the left table.",
              "More rows than expected = a many-to-many (Cartesian) join."])
    mf = pd.DataFrame({"File": [M35.MFILES[i].name if hasattr(M35.MFILES[i], "name") else str(M35.MFILES[i]) for i in range(12)], "Rows": M35.MROWS})
    mf["File"] = mf.File.map(lambda s: str(s).replace("\\", "/").split("/")[-1])

    def month_table(mm):
        mm.h("tMonthFiles: Week2_Data_Files › SuperMart_Monthly_Sales_2025")
        mm.s.table(mf, "tMonthFiles", fmts={"Rows": "int"}, widths={"File": 34, "Rows": 9})
        mm.s.r += 1

    tp(m, "3.4.11",
       [("January + February rows", "=INDEX(tMonthFiles[Rows],1)+INDEX(tMonthFiles[Rows],2)", int(mf.Rows.iloc[0] + mf.Rows.iloc[1]), "int",
         "Append two: rows add up.")],
       [("n", "Append March and April. How many rows?", "=INDEX(tMonthFiles[Rows],3)+INDEX(tMonthFiles[Rows],4)", int(mf.Rows.iloc[2] + mf.Rows.iloc[3]),
         "int", "March + April")],
       steps=["Load January and February as queries (From Text/CSV › Transform).", "Home › Append Queries as New › Two tables.",
              "Check: rows = January + February."], extra=month_table)
    tp(m, "3.4.12",
       [("All 12 files", "=SUM(tMonthFiles[Rows])", int(mf.Rows.sum()), "int", "Append Queries › Three or more tables, or From Folder."),
        ("FY2025 lines in tDash", "=ROWS(tDash)", len(D.DS), "int", "Equal: nothing lost or doubled.")],
       [("n", "Which month has the most rows? Type its row count.", "=MAX(tMonthFiles[Rows])", int(mf.Rows.max()), "int", "Largest month"),
        ("n", "Rows in the second quarter (Apr-Jun)?", "=INDEX(tMonthFiles[Rows],4)+INDEX(tMonthFiles[Rows],5)+INDEX(tMonthFiles[Rows],6)",
         int(mf.Rows.iloc[3:6].sum()), "int", "Apr-Jun rows")],
       steps=["Data › Get Data › From File › From Folder › the monthly folder.", "Combine & Transform: one query appends every file.",
              "Next month, drop the new file in the folder and Refresh."], extra=None)
    tp(m, "3.4.13",
       [("December rows (column renamed Sales_NGN)", "=INDEX(tMonthFiles[Rows],12)", int(mf.Rows.iloc[11]), "int",
         "Append by name: Sales would be null and a new Sales_NGN column appear for these rows.")],
       [("t", "What must you do to December's query before appending? (rename the column / delete rows)", "rename the column", None, None,
         ("rename", "rename column", "rename sales_ngn to sales"))],
       steps=["Append matches columns by name, not position.", "Rename mismatched columns before appending (Sales_NGN › Sales).",
              "After appending, filter each new column for null to spot leftovers."])
    zr = M35.ZRAW.value_counts().rename_axis("Zone_as_typed").reset_index(name="Rows")
    zr["Shown_as"] = zr.Zone_as_typed.map(lambda s: s.replace(" ", "·"))

    def zone_table(mm):
        mm.h("tZoneRaw: distinct Zone values in SuperMart_Sales_2025_DIRTY.csv (· marks a space)")
        mm.s.table(zr[["Zone_as_typed", "Shown_as", "Rows"]], "tZoneRaw", fmts={"Rows": "int"}, widths={"Zone_as_typed": 18, "Shown_as": 18})
        mm.s.r += 1

    tp(m, "3.4.14",
       [("Distinct values as typed", "=ROWS(tZoneRaw)", len(zr), "int", "Four real zones hidden behind stray spaces."),
        ("Rows that fail an exact match", "=SUMPRODUCT(tZoneRaw[Rows]*(TRIM(tZoneRaw[Zone_as_typed])<>tZoneRaw[Zone_as_typed]))", M35.Z_FAIL, "int",
         "Each would get null from the merge."),
        ("Distinct after TRIM", "=COUNTA(UNIQUE(TRIM(tZoneRaw[Zone_as_typed])))", 4, "int", "")],
       [("n", "How many rows have a leading space before 'South-West'?", '=SUMIF(tZoneRaw[Zone_as_typed]," South-West",tZoneRaw[Rows])',
         int(M35.ZRAW.eq(" South-West").sum()), "int", "Leading-space South-West"),
        ("n", "How many rows are clean already?", "=SUMPRODUCT(tZoneRaw[Rows]*(TRIM(tZoneRaw[Zone_as_typed])=tZoneRaw[Zone_as_typed]))",
         int(len(M35.ZRAW) - M35.Z_FAIL), "int", "Clean rows")],
       steps=["Before merging, select the key column › Transform › Format › Trim (and Clean).", "Match case too: Format › Capitalize Each Word, or Lowercase on both sides.",
              "Check the distinct count after cleaning."], extra=zone_table)
    tp(m, "3.4.15",
       [("transactions rows", "=ROWS(tTransactions)", len(TX), "int", ""),
        ("transactions distinct order_id", "=COUNTA(UNIQUE(tTransactions[order_id]))", int(TX.order_id.nunique()), "int", "Unique: one side is 'one'."),
        ("deliveries distinct order_id", "=COUNTA(UNIQUE(tDeliveries[order_id]))", int(DV.order_id.nunique()), "int", "")],
       [("n", "Is delivery_id unique in tDeliveries? Type its distinct count.", "=COUNTA(UNIQUE(tDeliveries[delivery_id]))", int(DV.delivery_id.nunique()),
         "int", "Distinct delivery_id"),
        ("t", "Keys unique on both sides: is the join one-to-one, one-to-many or many-to-many?", "one-to-one", None, None, ("1:1", "one to one"))],
       steps=["Count rows and distinct keys on each side before merging.", "Unique on one side = one-to-many: rows stay as the many side.",
              "Not unique on either side: stop, the join will multiply rows."], code=True)
    tp(m, "3.4.16",
       [("Rows in (left)", "=ROWS(tTransactions)", len(TX), "int", "Log the counts in and out of every merge."),
        ("Rows out (left join)", "=ROWS(tLines)", len(L), "int", "Equal: safe."),
        ("Unmatched (courier null)", "=COUNTBLANK(tLines[courier_name])", int(L.courier_name.isna().sum()), "int", "")],
       [("o", "Write the merge-log entry for transactions LEFT JOIN deliveries.",
         "Left: transactions (11,000). Right: deliveries (10,430). Key: delivery_id. Kind: left outer. Out: 11,000 rows; 570 without a "
         "delivery, all cancellations. Checked by: name, date.")],
       steps=["Record left and right tables, the key and the join kind.", "Record rows in and out and the unmatched count, with the reason.",
              "Rename each step in Applied Steps so the query reads like the log."])
    m.finish()


# =====================================================================================================
def m35(b, D):
    TX, DD, L = D.TX, D.DD, D.L
    m = ModuleSheet(b, "3.5 Data Model", "3.5  The Excel Data Model and Star Schema",
                    "Tutor walk-through after Module 3.5: load tables to the Data Model, relate them, add a date table, and see why "
                    "relationships beat lookup columns", "3.5.1 - 3.5.7")
    m.dataset([TX_DS, ("tProducts / tCustomers / tSellers", "Data_ sheets (A1)", "The dimensions."),
               ("tDimDate", "Sheet Data_DimDate (A1)", "One row per day 2020-2025: the date table.")])
    cat = D.NT.groupby("category").total_price.sum()
    tp(m, "3.5.1",
       [("Electronics net revenue (PivotTable from two related tables)", "=" + net(category="Electronics"), f(cat["Electronics"]), N,
         "category from tProducts, total_price from tTransactions.")],
       [("n", "Build the PivotTable (category on rows). What does Home show?", "=" + net(category="Home"), f(cat["Home"]), N, "Home net revenue")],
       steps=["Select tTransactions › Insert › PivotTable › tick 'Add this data to the Data Model'.", "Repeat for tProducts (Power Pivot › Add to Data Model).",
              "Relate them (3.5.2); the PivotTable field list shows both tables."],
       do=[("Insert a PivotTable from tTransactions with 'Add this data to the Data Model' ticked.", "The field list has an 'All' tab listing tables.")])
    tp(m, "3.5.2",
       [("product_id unique in tProducts (one side)", "=ROWS(tProducts)-COUNTA(UNIQUE(tProducts[product_id]))", 0, "int", "0 duplicates."),
        ("Lines whose product is missing (many side)", "=SUMPRODUCT(--(COUNTIF(tProducts[product_id],tTransactions[product_id])=0))", 0, "int", "")],
       [("t", "tProducts[product_id] to tTransactions[product_id] is which cardinality?", "one-to-many", None, None, ("1:*", "1 to many", "one to many")),
        ("n", "How many relationships connect tTransactions to tProducts, tCustomers, tSellers and tDimDate?", "=4", 4, "int", "Relationships")],
       steps=["Power Pivot › Diagram View.", "Drag tTransactions[product_id] onto tProducts[product_id].",
              "Repeat for customer_id, seller_id, and order_date › tDimDate[Date]."])
    tp(m, "3.5.3",
       [("Fact rows", "=ROWS(tTransactions)", len(TX), "int", "Many rows, keys and numbers."),
        ("Numeric fields in the fact", "=4", 4, "int", "quantity, unit_price, total_price, commission_amount.")],
       [("n", "Sum the fact's quantity column.", "=SUM(tTransactions[quantity])", int(TX.quantity.sum()), "int", "Total quantity")],
       steps=["Keep only keys, dates and numbers in the fact.", "Leave descriptions (names, categories) to the dimensions.",
              "One row per event at a stated grain: here, one order line."])
    tp(m, "3.5.4",
       [("Product categories", "=COUNTA(UNIQUE(tProducts[category]))", int(D.P.category.nunique()), "int", ""),
        ("Customer zones", "=COUNTA(UNIQUE(tCustomers[zone]))", int(D.C.zone.nunique()), "int", "")],
       [("n", "How many distinct subcategories does tProducts describe?", "=COUNTA(UNIQUE(tProducts[subcategory]))", int(D.P.subcategory.nunique()),
         "int", "Subcategories")],
       steps=["One row per member (product, customer, seller), with a unique key.", "Descriptive columns used to slice: category, zone, tier.",
              "Hide the key from report view once related (3.5.6)."])
    no_order = int((~DD.Date.isin(TX.order_date)).sum())
    tp(m, "3.5.5",
       [("Days in tDimDate", "=ROWS(tDimDate)", len(DD), "int", "Every day, none missing."),
        ("First date", "=MIN(tDimDate[Date])", f((DD.Date.min() - pd.Timestamp("1899-12-30")).days), "date", ""),
        ("Days with no order", "=SUMPRODUCT(--(COUNTIF(tTransactions[order_date],tDimDate[Date])=0))", no_order, "int",
         "Why the date table, not the order dates.")],
       [("n", "How many days does 2024 have in tDimDate?", "=COUNTIF(tDimDate[Year],2024)", int((DD.Year == 2024).sum()), "int", "Days in 2024"),
        ("n", "Are the dates unique? Type the duplicate count.", "=ROWS(tDimDate)-COUNTA(UNIQUE(tDimDate[Date]))", 0, "int", "Duplicate dates")],
       steps=["Power Pivot › Design › Date Table › New, or load tDimDate.", "Mark as Date Table, choosing the Date column.",
              "Relate tDimDate[Date] to tTransactions[order_date]; sort Month by Month_No."])
    tp(m, "3.5.6",
       [("Columns in tTransactions", "=COLUMNS(tTransactions)", 15, "int", "Hide keys and raw numbers; show measures instead."),
        ("Key columns to hide", "=5", 5, "int", "customer_id, product_id, seller_id, delivery_id, order_id.")],
       [("t", "Month names sort Apr, Aug, Dec… How do you fix it? (Sort by Column …)", "Month_No", None, None, ("sort by month_no", "by month_no"))],
       steps=["In Power Pivot, right-click a key column › Hide from Client Tools.", "Select Month › Sort by Column › Month_No.",
              "Group measures in one table (3.6.31)."])
    tp(m, "3.5.7",
       [("Lookup columns needed without relationships", "=ROWS(tTransactions)*2", len(TX) * 2, "int",
         "Category and zone on 11,000 rows = 22,000 XLOOKUP cells."),
        ("With relationships", "=0", 0, "int", "The model looks them up when needed.")],
       [("t", "What does XLOOKUP('PROD00575', tProducts[product_id], tProducts[category]) return?",
         D.P.set_index("product_id").category["PROD00575"], '=XLOOKUP("PROD00575",tProducts[product_id],tProducts[category])', "Category of PROD00575")],
       steps=["A lookup column copies data into the fact: more size, more to refresh, more to break.",
              "A relationship stores nothing extra and works in both directions for filtering.", "Use the model for anything beyond a one-off check."])
    m.finish()
