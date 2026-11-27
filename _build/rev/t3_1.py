from revlib import T

ON, LC = "SuperMart Online", "Lifecare Hospital"
TRY = "ETEH_Week3_Section_Examples.xlsx  ›  sheet '3.1 Database Basics'  ·  supermart_online.db"
DB = "supermart_online.db (Week3_Data_Files)"


def D(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", ON)
    return T(tag, **k)


TOPICS = [
    D("3.1.1",
      what="A **database** is an organised store of data, managed by software (a DBMS) that **enforces rules**: "
           "each column has a type, IDs must be unique, and links between tables must be valid. A spreadsheet "
           "accepts anything you type; a database refuses bad data.",
      like="A spreadsheet is like a notebook: you can write anything anywhere. A database is like a bank's ledger "
           "system: every entry must follow the rules, or it is rejected.",
      how=["Open the database file in a tool such as DB Browser for SQLite.", "Look at the list of tables and views.",
           "Check the row count of each table."],
      tip="Ask what one row of each table means before you analyse it.",
      trap="Treating a database like a spreadsheet you can edit freely.",
      case="SuperMart Online's marketplace data is stored in " + DB + ".",
      steps=["Open the file and list its objects.", "Count the rows in each table.", "Note the view and the indexes."],
      table=[["Object", "Count"], ["Tables", "6"], ["Views", "1 (v_order_lines)"], ["Indexes", "5"],
             ["Largest table", "transactions: 11,000 rows"]],
      so="One file holds six linked tables, a ready-made view and indexes for speed, all governed by rules.",
      split=4.90),

    D("3.1.2",
      what="A database **table** is like a tidy Excel sheet: each **row** is one record and each **column (field)** "
           "holds one kind of fact. The difference: every field has a **declared type**, so a text value can't sneak "
           "into a number column.",
      like="A table is like a class register: one line per pupil, the same boxes for every pupil, and the school "
           "won't accept a phone number in the 'age' box.",
      how=["Identify what one row represents (the grain).", "List the fields and their types.",
           "Look at a few rows to understand the values."],
      tip="Say the grain out loud: 'one row = one customer'.",
      trap="Mixing grains, e.g. customer rows and order rows in the same table.",
      case="The analyst explores the customers table in " + DB + ".",
      steps=["SELECT * FROM customers LIMIT 2;", "Count the rows and fields.", "State the grain."],
      table=[["customer_id", "Name", "City / State", "Tier"], ["CUST00001", "Funmilayo Buhari", "Kano / Kano", "Bronze"],
             ["CUST00002", "Amara Attah", "Ikorodu / Lagos", "Gold"]],
      so="customers has 2,500 rows and 16 fields. The grain is one row per customer.",
      split=4.20),

    D("3.1.3",
      what="A **primary key** is a column whose value is **unique** for every row and never empty, like "
           "transaction_id. It lets you point to exactly one row. The database rejects any duplicate.",
      like="A primary key is like a National Identification Number: no two people share one, and everyone has one.",
      how=["Find the ID column for the table.", "Check COUNT(*) equals COUNT(DISTINCT id).", "Try to insert a duplicate: it fails."],
      tip="Run the uniqueness check on every new table you receive.",
      trap="Using a name as a key; names repeat (2,500 customers share just 46 last names).",
      case="The analyst checks that transaction_id is a proper key in the transactions table.",
      steps=["SELECT COUNT(*), COUNT(DISTINCT transaction_id) FROM transactions;",
             "Subtract to get the duplicates.", "Try inserting a row with an existing transaction_id."],
      table=[["Check", "Result"], ["Rows", "11,000"], ["Distinct transaction_id", "11,000"], ["Duplicates", "0"],
             ["Insert a duplicate", "UNIQUE constraint failed"]],
      so="Every transaction_id is unique, and the database blocks any attempt to repeat one.",
      split=4.90),

    D("3.1.4",
      what="A **foreign key** is a column that points to the primary key of **another** table, like "
           "transactions.product_id → products.product_id. It protects **referential integrity**: you can't record "
           "an order for a product that doesn't exist.",
      like="A foreign key is like a receipt that quotes a product code: the code must exist in the shop's "
           "price list, or the receipt makes no sense.",
      how=["Find columns ending in _id that match another table's key.", "Check for orphans with a LEFT JOIN … IS NULL.",
           "Expect 0 orphans."],
      tip="Run an orphan check after every data load.",
      trap="Assuming links are valid without checking.",
      case="The analyst checks that every order line points to a real product, and tests the rule.",
      steps=["LEFT JOIN products ON product_id; count rows where p.product_id IS NULL.",
             "Turn on PRAGMA foreign_keys = ON.", "Try inserting an order for product 'NOPE'."],
      table=[["Check", "Result"], ["Order lines with no matching product", "0"],
             ["Admissions with no matching doctor (Lifecare)", "0"], ["Insert order for product 'NOPE'", "FOREIGN KEY constraint failed"]],
      so="All links are valid, and the database refuses an order for a product that doesn't exist.",
      split=4.90),

    D("3.1.5",
      what="A **flat file** puts everything in one big table, repeating details (customer name, tier, state) on "
           "every row. A **relational** design stores each thing once, in its own table, and links them with keys. "
           "Change a detail once and every order sees it.",
      like="A flat file is like writing a customer's full address on every receipt. Relational is like writing the "
           "customer number, with the address kept once on a customer card.",
      how=["Spot columns that repeat the same value for the same ID.", "Move them to their own table with a key.",
           "Link back with the key."],
      tip="If a fact repeats for the same ID, it belongs in its own table.",
      trap="Editing a repeated fact in some rows but not others.",
      case="The view v_order_lines shows the flat version: customer details copied onto all 11,000 order lines.",
      steps=["Count customer cells in the flat view (4 columns × 11,000).", "Count the same facts stored once (4 × 2,500).",
             "Count the rows to edit if CUST00985 moves state."],
      table=[["Measure", "Flat view", "Relational"], ["Customer cells", "44,000", "10,000"],
             ["Rows to edit if CUST00985 moves", "14", "1"]],
      so="77% of the customer cells in the flat view are repeats. Relational storage means one edit, not 14.",
      split=4.20),

    D("3.1.6",
      what="**First Normal Form (1NF)**: each cell holds **one value**, there are **no repeating columns** "
           "(Phone1, Phone2, Phone3), and each row has a **key**. It is the database version of tidy data.",
      like="1NF is like a form with one box per answer. You never squeeze two phone numbers into one box.",
      how=["Look for cells holding lists or several values.", "Look for numbered columns (Item1, Item2…).",
           "Split them into separate rows or columns."],
      tip="One cell, one value.",
      trap="Storing 'Calabar, Cross River' in one cell and then trying to filter by state.",
      case="The deliveries table stores a full address, but also keeps city and state in separate columns.",
      steps=["SELECT delivery_address, city, state FROM deliveries LIMIT 2;", "Check each column holds one value.",
             "Filter by state without parsing text."],
      table=[["delivery_address", "city", "state"], ["431 Road, Calabar", "Calabar", "Cross River"],
             ["266 Avenue, Maiduguri", "Maiduguri", "Borno"]],
      so="Because state has its own column, you can filter by state directly, without splitting text.",
      split=4.20),

    D("3.1.7",
      what="**2NF**: every non-key column depends on the **whole** key. **3NF**: no column depends on another "
           "non-key column. Example: zone depends on state, not on the customer, so the strict design keeps zone in "
           "a state lookup table.",
      like="3NF is like keeping the bus fare chart at the bus stop, not writing the fare on every passenger's "
           "ticket. If the fare changes, you fix one chart.",
      how=["For each column ask: does it describe this row's key?", "If it describes another column, move it to a lookup.",
           "Link the lookup by that column."],
      tip="If you keep a convenience copy, check it against the lookup after every load.",
      trap="Copying derived facts (like zone) into many tables and letting them drift apart.",
      case="In " + DB + ", zone depends on state. The lookup is nigeria_states, but customers also keeps a copy of zone.",
      steps=["product_name sits in products, not transactions (2NF).", "zone is looked up from nigeria_states (3NF).",
             "Compare customers.zone with the lookup."],
      table=[["Check", "Result"], ["States in nigeria_states", "32 (6 zones)"],
             ["Customers whose zone disagrees with the lookup", "0"], ["Customers with a state not in the lookup", "0"]],
      so="The copy matches today. The lookup table is the single source of truth if they ever disagree.",
      split=4.90),

    D("3.1.8",
      what="Databases declare a **type** for every column: **INTEGER**, **REAL** (decimal), **TEXT**, **DATE** and "
           "so on. Rules like **CHECK (quantity > 0)** add business logic. Note: SQLite stores dates as text in "
           "yyyy-mm-dd form.",
      like="Types are like labelled compartments in a toolbox: screws in one, nails in another. The wrong item "
           "won't fit.",
      how=["Read the CREATE TABLE statement for the types.", "Look for CHECK and NOT NULL rules.",
           "Test a value that breaks a rule."],
      tip="Store dates as yyyy-mm-dd so they sort correctly as text.",
      trap="Storing money as text, which can't be summed.",
      case="The transactions table declares quantity INTEGER CHECK (quantity > 0). The analyst tests it.",
      steps=["Read the schema for transactions.", "Check typeof() on a few columns.", "Try inserting quantity = 0."],
      table=[["Column", "Declared type", "Stored as"], ["quantity", "INTEGER, CHECK > 0", "integer"],
             ["total_price", "REAL", "real"], ["order_date", "DATE", "text (yyyy-mm-dd)"],
             ["Insert quantity = 0", "Rejected", "CHECK constraint failed"]],
      so="The database rejects a zero quantity before it can spoil any total.",
      split=4.20),

    D("3.1.9",
      what="An **index** is a sorted lookup that lets the database **jump** to matching rows instead of reading "
           "every row. Columns you filter or join on often (dates, IDs) should be indexed. **EXPLAIN QUERY PLAN** "
           "shows whether an index is used.",
      like="An index is like the index at the back of a textbook: you go straight to page 214 instead of reading "
           "the whole book.",
      how=["List indexes: they appear in the schema.", "Prefix a query with EXPLAIN QUERY PLAN.",
           "Look for SEARCH … USING INDEX (good) or SCAN (reads everything)."],
      tip="Filter on indexed columns, such as order_date, where possible.",
      trap="Indexing every column; each index slows down writes.",
      case="The analyst sums January 2024 sales, then filters by payment method, and compares the query plans.",
      steps=["EXPLAIN QUERY PLAN … WHERE order_date BETWEEN '2024-01-01' AND '2024-01-31'.",
             "EXPLAIN QUERY PLAN … WHERE payment_method = 'Card'.", "Compare the plans."],
      table=[["Filter", "Query plan"], ["order_date (indexed)", "SEARCH USING INDEX ix_trx_date"],
             ["payment_method (no index)", "SCAN transactions"], ["January 2024 result", "191 orders, ₦26,978,450"]],
      so="The date filter jumps straight to January's 191 rows; the payment filter must read all 11,000.",
      split=4.70),

    D("3.1.10",
      what="A **view** is a **saved query** that behaves like a table. It stores no data of its own; it runs its "
           "query whenever you use it. Views hide complex joins, so analysts can query one simple 'table'.",
      like="A view is like a shop window: it shows a chosen arrangement of goods from the store room, without "
           "moving the stock.",
      how=["Find views in the schema list.", "Query a view exactly like a table.", "Read the view's SQL to see where its data comes from."],
      tip="Ask your database team for views of common joins.",
      trap="Thinking a view is a copy; it always shows the current data.",
      case="v_order_lines joins transactions, customers, products, sellers and deliveries.",
      steps=["SELECT COUNT(*) FROM v_order_lines;", "Count its columns.", "Read its SQL: which tables does it join?"],
      table=[["Check", "Result"], ["Rows in v_order_lines", "11,000"], ["Columns", "30"],
             ["Tables it joins", "5"], ["Data it stores itself", "None"]],
      so="One simple query on the view replaces a five-table join every time.",
      split=4.90),

    D("3.1.11",
      what="A **schema** is a named container for tables, written as **schema.table**, e.g. sales.orders. It keeps "
           "names apart and controls access. In SQLite, the main file is **main**, and you can **ATTACH** another "
           "database under its own schema name.",
      like="Schemas are like surnames: two people called Ade are told apart as Ade Okafor and Ade Bello.",
      how=["Write the schema before the table name: main.transactions.", "ATTACH a second database with a short name.",
           "Query across both."],
      tip="Use schema names whenever two databases are open.",
      trap="Assuming a table name is unique when several schemas are attached.",
      case="The analyst opens " + DB + " and attaches lifecare_hospital.db as lh.",
      steps=["ATTACH 'lifecare_hospital.db' AS lh;", "SELECT COUNT(*) FROM main.transactions;",
             "SELECT COUNT(*) FROM lh.admissions;"],
      table=[["Reference", "Schema", "Rows"], ["main.transactions", "main (SuperMart Online)", "11,000"],
             ["lh.admissions", "lh (Lifecare)", "54,860"]],
      so="Two databases, one session: the schema name says exactly which table you mean.",
      split=4.20),

    D("3.1.12",
      what="A **transaction** groups several changes so they succeed or fail **together**. **ACID** means: "
           "**Atomic** (all or nothing), **Consistent** (rules kept), **Isolated** (users don't interfere), "
           "**Durable** (saved changes survive a crash).",
      like="A transaction is like a bank transfer: the money leaves one account and arrives in the other, or "
           "neither happens. It never just disappears.",
      how=["BEGIN the transaction.", "Make the related changes.", "COMMIT to save them, or ROLLBACK to undo all of them."],
      tip="Group changes that belong together in one transaction.",
      trap="Recording a sale without reducing stock, so the two disagree.",
      case="On a copy of " + DB + ", a new order must add a transaction and reduce product stock together.",
      steps=["BEGIN; reduce PROD00001 stock by 2; insert the order line.", "Check both changes inside the transaction.",
             "ROLLBACK and check again."],
      table=[["Moment", "Transactions", "PROD00001 stock"], ["Before", "11,000", "142"], ["Inside the transaction", "11,001", "140"],
             ["After ROLLBACK", "11,000", "142"]],
      so="Rollback undid both changes together. The data was never left half-updated.",
      split=4.20),

    D("3.1.13",
      what="A **normalised** schema has many small tables, which is best for recording transactions. A **star "
           "schema** has one central **fact** table surrounded by **dimension** tables, each one join away. It is "
           "best for analysis, PivotTables and DAX.",
      like="Normalised is like a filing cabinet with many folders; a star is like a hub-and-spoke bus network: "
           "every stop is one ride from the centre.",
      how=["Find the fact table (the numbers).", "List the dimensions it links to.", "Check each dimension is one join from the fact."],
      tip="Model for analysis as a star.",
      trap="Chaining lookups (state → zone) several joins away from the fact.",
      case="The analyst sketches " + DB + " as a star for the Data Model.",
      steps=["Fact: transactions.", "Dimensions: customers, products, sellers, deliveries (plus a Date table).",
             "Note that nigeria_states hangs off customers: a snowflake step."],
      table=[["Shape", "Tables", "Best for"], ["Normalised (as stored)", "6 tables, state lookup 2 joins away", "Recording orders"],
             ["Star (for analysis)", "1 fact + 5 dimensions incl. Date", "Pivots, Power Pivot, DAX"]],
      so="The same data can be shaped two ways. For analysis, keep every dimension one hop from the fact.",
      split=4.20),

    D("3.1.14",
      what="A **fact table** stores **events with numbers** (one row per order line) plus keys to dimensions. A "
           "**dimension table** stores the **descriptions** you filter and group by: customers, products, dates. "
           "Facts are long and narrow; dimensions are shorter and wider.",
      like="Facts are like till receipts (many, mostly numbers); dimensions are like the customer and product cards "
           "the receipts refer to.",
      how=["Look for tables of events with numbers: those are facts.", "Look for tables of descriptions: dimensions.",
           "Check the grain of the fact table."],
      tip="SUM facts; filter and group by dimensions.",
      trap="Summing a number stored in a dimension, e.g. total_spent from customers alongside order lines.",
      case="The analyst classifies the tables in " + DB + ".",
      steps=["Count rows and columns per table.", "Mark which have numbers to add up.", "Mark which describe things."],
      table=[["Table", "Type", "Rows", "Holds"], ["transactions", "Fact", "11,000", "quantity, total_price, commission"],
             ["customers", "Dimension", "2,500", "name, state, tier"], ["products", "Dimension", "800", "category, brand"],
             ["sellers", "Dimension", "200", "store, state, rating"]],
      so="One fact table with 11,000 order lines, described by the customer, product and seller dimensions.",
      split=4.20),

    D("3.1.15",
      what="**OLTP** systems run the business: many small, fast writes, one order at a time, normalised tables. "
           "**OLAP** systems analyse the business: big reads that summarise many rows, usually in a star schema.",
      like="OLTP is like the till at the checkout, recording each sale; OLAP is like the manager's monthly report "
           "that reads all the till rolls.",
      how=["Identify where data is captured (OLTP).", "Identify where it is analysed (OLAP).",
           "Copy from OLTP to OLAP on a schedule, not live."],
      tip="Query a copy for analysis, not the live till system.",
      trap="Running heavy reports on the live system and slowing down sales.",
      case="SuperMart Online's app records orders; the analyst's Data Model reads them for the dashboard.",
      steps=["OLTP: the marketplace app writes each order.", "Extract: " + DB + " holds the history.",
             "OLAP: the Data Model summarises it."],
      table=[["Aspect", "OLTP (marketplace app)", "OLAP (Data Model)"], ["Typical action", "Insert one order", "Sum years of orders"],
             ["Data in this pack", "11,000 orders, 2020–2025", "Summaries by month, state, category"],
             ["Shape", "Normalised", "Star"]],
      so="The same 11,000 orders serve two jobs: recording each sale, and analysing all of them.",
      split=4.20),
]
