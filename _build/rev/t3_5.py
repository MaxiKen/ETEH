from revlib import T

ON = "SuperMart Online"
TRY = "ETEH_Week3_Section_Examples.xlsx  ›  sheet '3.5 Data Model & Star'  ·  D_ data sheets"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", ON)
    return T(tag, **k)


TOPICS = [
    O("3.5.1",
      what="The **Data Model** is a database built into Excel (Power Pivot). It stores tables in compressed "
           "form, handles millions of rows, and links tables with **relationships** instead of VLOOKUP formulas. "
           "One PivotTable can then use fields from several tables at once.",
      like="The Data Model is like a filing cabinet with labelled drawers joined by string: pull a customer's "
           "file and their orders come with it.",
      how=["Select a table > Power Pivot > Add to Data Model (or tick 'Add this data to the Data Model').",
           "Repeat for each table.", "Insert > PivotTable > From Data Model."],
      tip="Load each source table once, then link them with relationships.",
      trap="Copying lookup columns into the big table instead of linking tables.",
      case="The analyst wants revenue by product category, but category lives in products and revenue in "
           "transactions.",
      steps=["Add transactions and products to the Data Model.", "Relate them on product_id.",
             "PivotTable: category in Rows, Sum of total_price in Values."],
      table=[["Category", "Order lines", "Revenue"], ["Electronics (highest)", "1,593", "₦717,407,700"],
             ["Home", "1,433", "₦368,892,800"], ["Beauty (lowest)", "1,239", "₦11,727,050"],
             ["All 7 categories", "11,000", "₦1,394,713,700"]],
      so="One relationship let a single PivotTable combine two tables. No VLOOKUP column was needed.",
      split=4.20),

    O("3.5.2",
      what="A **relationship** links two tables on one shared key. Most are **one-to-many**: one product, many "
           "sales. The 'one' side must have **unique** keys. **Cardinality** describes how many rows match on "
           "each side.",
      like="One-to-many is like one mother with several children: each child points to exactly one mother.",
      how=["Data > Relationships > New (or drag keys in Diagram View).",
           "Many side = fact table; one side = lookup table.", "Check the key is unique on the one side."],
      tip="The key on the 'one' side must have no duplicates or blanks.",
      trap="Relating two tables on a column that repeats on both sides.",
      case="The analyst sets up the SuperMart Online model, with transactions at the centre.",
      steps=["Relate transactions to customers, products and sellers.", "Relate transactions to deliveries.",
             "Relate order_date to DimDate[Date]."],
      table=[["From transactions", "To table", "Cardinality"], ["customer_id", "customers (2,500)", "Many-to-one"],
             ["product_id", "products (800)", "Many-to-one"], ["seller_id", "sellers (200)", "Many-to-one"],
             ["delivery_id", "deliveries (10,430)", "One-to-one; 570 blanks"]],
      so="Five relationships connect every table. Each lookup key is unique, so filters flow correctly.",
      split=4.20),

    O("3.5.3",
      what="A **fact table** records **events**: one row per sale, admission or payment. It holds the numbers "
           "you add up (quantity, revenue, commission) plus **keys** to the lookup tables. Its **grain** is what "
           "one row means.",
      like="A fact table is like a till roll: a long list of what happened, with codes pointing to who and what.",
      how=["State the grain: one row per order line.", "Keep numeric measures and keys.",
           "Leave descriptions (names, categories) in lookup tables."],
      tip="Write the grain down before building anything.",
      trap="Mixing grains, e.g. order lines and monthly totals in one table.",
      case="The analyst checks that transactions is a proper fact table.",
      steps=["Grain: one row per order line (11,000 rows).", "Find the measures.", "Find the keys."],
      table=[["Column", "Type", "Example"], ["quantity, total_price", "Measure (add up)", "Sum of total_price"],
             ["commission_amount", "Measure (add up)", "Sum of commission"],
             ["customer_id, product_id, seller_id", "Key", "Links to lookups"], ["order_date", "Key", "Links to DimDate"]],
      so="transactions holds 11,000 events with measures and keys. Descriptions live elsewhere.",
      split=4.20),

    O("3.5.4",
      what="A **dimension** (lookup) table describes things: customers, products, sellers, dates. It is short "
           "and wide, with **one row per item** and a **unique key**. You slice facts by its columns.",
      like="A dimension is like a phone contact card: one card per person, many details on it.",
      how=["Give it a unique key column.", "Keep descriptive columns: names, categories, places.",
           "Relate its key to the fact table."],
      tip="Check: rows = distinct keys.",
      trap="Leaving duplicate keys, which breaks the relationship.",
      case="The analyst reviews the three main lookup tables in SuperMart Online.",
      steps=["Count rows and distinct keys in each.", "List the useful descriptive columns.", "Relate each to transactions."],
      table=[["Dimension", "Key (unique)", "Useful columns"], ["customers (2,500 × 16)", "customer_id", "membership_tier, state, zone"],
             ["products (800 × 15)", "product_id", "category, brand_name, price"],
             ["sellers (200 × 14)", "seller_id", "store_name, store_zone, rating"]],
      so="Each lookup has one row per key, so it can slice the 11,000 transactions safely.",
      split=4.20),

    O("3.5.5",
      what="A **date table** has **one row for every day**, with no gaps, plus columns like Year, Quarter and "
           "Month. Mark it as the date table. It makes time calculations reliable and shows days and months with "
           "no sales.",
      like="A date table is like a wall calendar: every day is printed, even days when nothing happened.",
      how=["Create a table with every date from first to last year.", "Add Year, Quarter, Month_No, Month.",
           "Power Pivot > Design > Mark as Date Table; relate to order_date."],
      tip="Cover whole years, from 1 January to 31 December.",
      trap="Using the order_date column as the calendar; days with no orders are missing.",
      case="DimDate on sheet D_DimDate covers 2020–2025. Orders run from 20 Jan 2020 to 30 Dec 2025.",
      steps=["Count the rows in DimDate.", "Check for gaps: days spanned minus rows = 0.",
             "Count the distinct order dates."],
      table=[["Check", "Result"], ["Days in DimDate", "2,192 (2020-01-01 to 2025-12-31)"], ["Gaps", "0"],
             ["Days with at least one order", "1,996"], ["Days with no orders", "196, still shown"]],
      so="The date table fills the 196 empty days, so time calculations always have a complete calendar.",
      split=4.90),

    O("3.5.6",
      what="Make the model easy to use: **hide** keys and unused columns, set **Sort by Column** so months sort "
           "Jan to Dec, and give measures number formats. Users then see only clean, sensible fields.",
      like="It's like tidying a shop: stock rooms are hidden from customers, and shelves are arranged in a "
           "sensible order.",
      how=["Right-click technical columns > Hide from Client Tools.", "Select Month > Sort by Column > Month_No.",
           "Format measures (₦, %, thousands)."],
      tip="Hide every key on the fact side.",
      trap="Leaving months to sort alphabetically.",
      case="A PivotTable shows months in the wrong order and lists confusing ID columns.",
      steps=["Check the month order in Rows.", "Sort Month by Month_No.", "Hide customer_id, product_id etc. in transactions."],
      table=[["Setting", "Before", "After"], ["Month order", "Apr, Aug, Dec, Feb…", "Jan, Feb, Mar…"],
             ["Key columns in field list", "Visible", "Hidden"], ["Revenue format", "1394713700", "₦1,394,713,700"]],
      so="Three settings turned a confusing field list into one beginners can use.",
      split=4.20),

    O("3.5.7",
      what="A VLOOKUP column **copies** lookup data into every fact row: it bloats the file and must be "
           "extended when data grows. A **relationship** stores each description once and **links** it. It "
           "keeps working as rows are added.",
      like="VLOOKUP is like photocopying the full product catalogue onto every receipt; a relationship is like "
           "printing a product code that points to one catalogue.",
      how=["Remove VLOOKUP helper columns from the fact table.", "Add the lookup table to the Data Model.",
           "Create one relationship on the key."],
      tip="One relationship replaces many VLOOKUP columns.",
      trap="Adding a new VLOOKUP column every time someone needs another attribute.",
      case="The analyst compares adding category, tier, state and zone to transactions by VLOOKUP versus "
           "relationships.",
      steps=["Count the extra cells VLOOKUP would create.", "Count the extra cells a relationship creates.",
             "Compare what happens when new rows arrive."],
      table=[["Approach", "Extra cells", "When rows are added"], ["VLOOKUP category", "11,000", "Extend formulas"],
             ["VLOOKUP tier + state + zone", "33,000 more", "Extend formulas"],
             ["Relationships", "0", "Refresh; it just works"]],
      so="Relationships give access to all 16 customer columns and all product columns with 0 extra cells.",
      split=4.20),
]
