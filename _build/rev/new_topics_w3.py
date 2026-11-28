"""Week 3 simplification (edition 3.1): topics dropped and topics added. Used by rev/topic_surgery.py.

Aim of 3.1-3.2: understand how a relational database is built and why, then spend a little time in DB Browser for SQLite
running simple single queries. 3.3 onward: bring the data into Excel, where the analysis happens.
"""
DROP = [
    # 3.1: normalisation forms and the advanced database internals
    "3.1.6", "3.1.7", "3.1.11", "3.1.12",
    # 3.2: dialects, HAVING, calculated columns, CASE, subqueries, string/date functions
    "3.2.2", "3.2.11", "3.2.12", "3.2.14", "3.2.15", "3.2.16",
    # 3.3: SQL parameters and injection
    "3.3.8",
    # 3.4: rarely used join kinds, fuzzy matching, schema drift, cardinality SQL, documentation
    "3.4.4", "3.4.7", "3.4.8", "3.4.13", "3.4.15", "3.4.16",
    # 3.6: advanced DAX
    "3.6.9", "3.6.11", "3.6.14", "3.6.16", "3.6.21",
    "3.6.24", "3.6.25", "3.6.26", "3.6.28", "3.6.30", "3.6.34", "3.6.39", "3.6.40",
]

WHY = '''
T("3.1.5n1", "Why Relational Databases?",
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
'''

SCALE = '''
T("3.1.5n2", "Relational Databases as Data Grows",
  defn="Scalability is how well a way of storing data copes as the rows, the users and the years grow.",
  simple="In a flat sheet every new order copies the customer's details again. In a relational design a new order adds "
         "one short row that points to the customer.",
  like="Writing the customer's full address on every receipt is fine for ten receipts, not for a million.",
  grid=dict(label="Customer details stored as the business grows", widths=[1.3, 1.6, 1.7, 1.4], size=11.5, rows=[
      ["Order lines", "Flat: customer cells", "Relational: customer cells", "Fits one Excel sheet?"]] + SCALE_ROWS,
      note=f"Same {I(len(CU))} customers throughout; 4 details per customer (name, tier, state, zone); relational keeps one customer_id per line."),
  watch="An Excel sheet stops at 1,048,576 rows. A database has no such limit, and Excel can still analyse it through a connection (3.3) and the Data Model (3.5)."),
'''

MAP = '''
T("3.1.5n3", "How the SuperMart Online Tables Link",
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
'''

STRUCT = '''
T("3.2.4n1", "Look Around: the Database Structure Tab",
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
'''

BROWSE = '''
T("3.2.4n2", "Browse Data: Look Before You Query",
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
'''

JOIN = '''
T("3.2.10n1", "Joining Two Tables on a Key",
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
'''

BRIDGE = '''
T("3.2.17n1", "From a Quick Query to Analysis in Excel",
  defn="In this course SQL is how you look inside a database and check what is there; the analysis itself is done in Excel.",
  simple="Answer a quick question in DB Browser, then bring the table into Excel (3.3) and use PivotTables, functions and "
         "the Data Model (3.5 and 3.6). Both must give the same totals.",
  grid=dict(label="One question, two places, one answer", widths=[1.5, 2.1, 1.9, 1.0], size=11, rows=[
      ["Question", "In DB Browser (SQL)", "In Excel", "Answer"],
      ["Order lines in total", "SELECT COUNT(*) FROM transactions;", "=ROWS(tTransactions)", I(N_LINES)],
      ["Cancelled lines", "… WHERE order_status = 'Cancelled';", "=COUNTIFS(… ,\\"Cancelled\\")", I(CANC_N)],
      ["Top payment method", "GROUP BY payment_method ORDER BY … LIMIT 1", "PivotTable, sort largest first", TOP_PAY]]),
  watch="Check the SQL count against the Excel row count after every import; if they differ, the import is incomplete."),
'''

CSV = '''
T("3.3.9n1", "Another Route: Export a Table to CSV",
  defn="DB Browser can export any table or view to a CSV file, which Excel imports with Data › From Text/CSV.",
  simple="Use this when the ODBC driver cannot be installed, for example on a Mac or a locked work laptop. The CSV is a "
         "snapshot: it does not refresh.",
  steps=["DB Browser › File › Export › Table(s) as CSV file", "Pick transactions; keep column names in the first line",
         "Save SuperMart_Online_transactions.csv", "Excel › Data › From Text/CSV › Transform Data › Close & Load"],
  grid=dict(label="ODBC connection versus CSV export", widths=[1.1, 1.7, 1.5], size=11.5, rows=[
      ["", "ODBC connection", "CSV export"],
      ["Setup", "Driver + connection", "None"],
      ["New data", "Data › Refresh All", "Export again"],
      ["Best for", "Reports you repeat", "One-off analysis"]]),
  order=["steps", "grid"], split=0.52,
  watch=f"Check the row count after loading: transactions must arrive with {I(N_LINES)} data rows."),
'''

PIVOT = '''
T("3.3.9n2", "Analyse the Imported Table in Excel",
  defn="Once the data is an Excel table, everything from Weeks 1 and 2 works on it: functions, PivotTables and charts.",
  simple="Load transactions as a table, insert a PivotTable, put order_status in Rows and total_price in Values. The "
         "totals match the SQL you ran in 3.2.",
  grid=dict(label="PivotTable on the imported transactions", widths=[1.5, 1.0, 1.7], size=11.5, rows=[
      ["order_status", "Lines", "Sum of total_price"]] + STATUS_ROWS + [["Grand Total", I(N_LINES), N(GROSS)]]),
  watch="If the Grand Total differs from SELECT SUM(total_price) FROM transactions, the import is incomplete: refresh and check."),
'''

INSERT = {
    "3.1.5": [WHY, SCALE, MAP],
    "3.2.4": [STRUCT, BROWSE],
    "3.2.10": [JOIN],
    "3.2.17": [BRIDGE],
    "3.3.9": [CSV, PIVOT],
}
