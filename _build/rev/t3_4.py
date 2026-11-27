from revlib import T

ON = "SuperMart Online"
TRY = "ETEH_Week3_Section_Examples.xlsx  ›  sheet '3.4 Joins Merge Append'  ·  supermart_online.db"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", ON)
    return T(tag, **k)


TOPICS = [
    O("3.4.1",
      what="**Merge** puts two tables **side by side**, matching rows on a key, like SQL **JOIN**. **Append** "
           "**stacks** tables with the same columns **on top of each other**, like SQL **UNION ALL**. Merge adds "
           "columns; Append adds rows.",
      like="Merge is like stapling each customer's delivery slip to their receipt. Append is like putting "
           "February's receipts under January's in the same box.",
      how=["Ask: do I need more columns (Merge) or more rows (Append)?",
           "Merge: Home > Merge Queries, pick the key in both tables.", "Append: Home > Append Queries."],
      tip="Same columns, different rows = Append. Different columns, shared key = Merge.",
      trap="Appending tables with different columns, which creates half-empty columns.",
      case="The analyst has two jobs: add courier details to each order, and combine the January and February "
           "2025 sales files.",
      steps=["Orders + courier details: they share delivery_id, so Merge.",
             "Jan + Feb files: same columns, so Append.", "Check the row and column counts after each."],
      table=[["Job", "Operation", "Result"], ["transactions + deliveries", "Merge on delivery_id", "11,000 rows, more columns"],
             ["Jan (155) + Feb (107) files", "Append", "262 rows, same columns"]],
      so="Merge widened the orders table; Append lengthened the sales table. The question decides the tool.",
      split=4.20),

    O("3.4.2",
      what="An **inner join** keeps only rows whose key appears in **both** tables. Anything without a match "
           "disappears, with no warning. Check what was dropped before you trust the totals.",
      like="An inner join is like a dance where only people who found a partner stay on the floor; everyone "
           "else quietly leaves.",
      how=["Merge with Join Kind = Inner (SQL: JOIN … ON).", "Compare rows before and after.",
           "Find out what the dropped rows were."],
      tip="Always compare row counts before and after an inner join.",
      trap="Reporting revenue after an inner join without noticing rows were dropped.",
      case="The analyst joins transactions (11,000 orders) to deliveries (10,430 shipments) with an inner join "
           "and sums total_price.",
      steps=["JOIN deliveries ON delivery_id.", "Count rows and sum total_price.", "Compare with the full transactions table."],
      table=[["Measure", "All orders", "After inner join"], ["Rows", "11,000", "10,430"],
             ["Revenue", "₦1,394,713,700", "₦1,317,681,100"], ["Dropped", "None", "570 orders (₦77,032,600)"]],
      so="The inner join silently removed 570 orders worth ₦77.0m: the ones that never shipped.",
      split=4.20),

    O("3.4.3",
      what="A **left outer join** keeps **every row of the left table** and adds matching columns from the "
           "right. Where there is no match, the right-side columns are **blank (null)**. Use it when you must not "
           "lose any row of your main table.",
      like="It's like a class register: every student stays on the list, and the 'parent phone' column is blank "
           "for anyone who didn't give one.",
      how=["Put the main table first (left).", "Merge with Join Kind = Left Outer (SQL: LEFT JOIN).",
           "Count the blanks in the new columns."],
      tip="Left outer is the safe default for 'add details to my main table'.",
      trap="Putting the tables the wrong way round, so the main table becomes the right side.",
      case="Operations wants every order listed with its courier, including orders that never shipped.",
      steps=["FROM transactions LEFT JOIN deliveries ON delivery_id.", "Count all rows.", "Count rows where courier is blank."],
      table=[["Rows", "Count", "Courier column"], ["All orders kept", "11,000", "Filled or blank"],
             ["Orders with a delivery", "10,430", "Filled"], ["Orders without a delivery", "570", "Blank (null)"]],
      so="Every one of the 11,000 orders survived. The 570 blanks show exactly which never shipped.",
      split=4.20),

    O("3.4.4",
      what="A **right outer join** keeps **every row of the right table**. It is the same as a left join with "
           "the tables swapped, so most analysts just swap the order and use a left join.",
      like="It's the same register as before, just read from the other side of the desk.",
      how=["Decide which table must keep every row.", "Make that table the left one and use a left join.",
           "Or keep the order and choose Right Outer."],
      tip="Prefer left joins; they are easier to read.",
      trap="Mixing left and right joins in one query, which confuses readers.",
      case="The courier team wants every shipment with its order details.",
      steps=["Right join: transactions RIGHT JOIN deliveries.", "Same thing: deliveries LEFT JOIN transactions.",
             "Count the rows both ways."],
      table=[["Written as", "Keeps every row of", "Rows"], ["transactions RIGHT JOIN deliveries", "deliveries", "10,430"],
             ["deliveries LEFT JOIN transactions", "deliveries", "10,430"]],
      so="Both versions return the same 10,430 shipments. Swapping the tables turns a right join into a left join.",
      split=4.20),

    O("3.4.5",
      what="A **full outer join** keeps **all rows from both tables**, matching where it can. Unmatched rows "
           "from either side appear with blanks. It is the tool for **reconciliation**: comparing two lists to "
           "see what is missing from each.",
      like="It's like comparing your bank statement with your own record book: you list everything from both and "
           "tick off the matches.",
      how=["Merge with Join Kind = Full Outer.", "Flag rows missing on the left, on the right, or matched.",
           "Count each group."],
      tip="Count the three groups: both sides, left only, right only.",
      trap="Using an inner join for a reconciliation, which hides the differences.",
      case="Finance reconciles orders against shipments: is anything missing from either list?",
      steps=["Full outer join transactions and deliveries on delivery_id.", "Label each row: both / orders only / deliveries only.",
             "Count each label."],
      table=[["Group", "Rows", "Meaning"], ["In both", "10,430", "Order shipped"], ["Orders only", "570", "Never shipped"],
             ["Deliveries only", "0", "No unknown shipments"], ["Total", "11,000", "Full outer result"]],
      so="Every shipment has an order; only orders are missing shipments. The reconciliation balances.",
      split=4.20),

    O("3.4.6",
      what="A **left anti join** returns the rows of the left table that have **no match** on the right. It "
           "answers 'which of mine are missing over there?' It is an exception list.",
      like="It's like checking the guest list against the people who arrived: the anti join lists who didn't "
           "show up.",
      how=["Merge with Join Kind = Left Anti.", "SQL: LEFT JOIN … WHERE right.key IS NULL.", "Group the result to explain it."],
      tip="Use anti joins to build exception lists for follow-up.",
      trap="Assuming missing rows are errors before checking why they are missing.",
      case="Were any paid orders lost before shipping? The analyst lists orders with no delivery.",
      steps=["transactions LEFT ANTI deliveries on delivery_id.", "GROUP BY order_status, payment_status.",
             "Count and sum each group."],
      table=[["order_status", "payment_status", "Orders", "Value"], ["Cancelled", "Failed", "299", "₦40,317,700"],
             ["Cancelled", "Refunded", "271", "₦36,714,900"], ["Total", "Both", "570", "₦77,032,600"]],
      so="All 570 are cancellations, so no shipment was lost. The same method finds 27 customers who never ordered.",
      split=4.20),

    O("3.4.7",
      what="A **right anti join** returns rows of the **right** table with **no match** on the left. It is the "
           "mirror check: which products never sold, which sellers never made a sale.",
      like="It's the guest list check from the host's side: which invitations went out but nobody used.",
      how=["Merge with Join Kind = Right Anti.", "Or swap the tables and use Left Anti.", "Investigate each row returned."],
      tip="Run anti joins in both directions when checking two tables.",
      trap="Checking only one direction and missing the other gap.",
      case="The marketplace team asks: are any products or sellers sitting idle?",
      steps=["transactions RIGHT ANTI products on product_id.", "transactions RIGHT ANTI sellers on seller_id.",
             "List the sellers returned."],
      table=[["Check", "Rows", "Detail"], ["Products never sold", "0", "All 800 have sales"],
             ["Sellers never sold", "3", "SELL0046, SELL0084, SELL0136"]],
      so="Every product sells, but 3 of 200 sellers have never made a sale: a list for the seller team.",
      split=4.20),

    O("3.4.8",
      what="A **fuzzy merge** matches keys that are **similar** rather than identical, using a similarity "
           "**threshold** (0 to 1). It helps with hand-typed names, but matches are **suggestions** to review. "
           "Use it only when no proper ID exists.",
      like="Fuzzy matching is like recognising a friend's handwriting even when the letters are messy; usually "
           "right, sometimes wrong.",
      how=["Home > Merge Queries > tick Use fuzzy matching.", "Set Threshold (e.g. 0.8), Ignore case, Ignore spaces.",
           "Review every match before using it."],
      tip="Join on an ID whenever one exists; names are not unique.",
      trap="Accepting fuzzy matches without review.",
      case="A supplier list has 8 hand-typed store names to match to the sellers table.",
      steps=["Exact match the names.", "Then match ignoring case and extra spaces.", "Then fuzzy match at 0.8 and review."],
      table=[["Method", "Names matched", "Issue"], ["Exact match", "2 of 8", "Capitals and double spaces fail"],
             ["Ignore case + spaces", "7 of 8", "'Sokoto Supplier' still missed"],
             ["Fuzzy, threshold 0.8", "Candidate for all 8", "Review each one"]],
      so="Even a correct name match is ambiguous: 'Sokoto Suppliers' is 8 different sellers. Use seller_id.",
      split=4.20),

    O("3.4.9",
      what="After a merge, the new column holds a **small table** for each row. **Expand** it to bring in "
           "columns (this can add rows when there are several matches). **Aggregate** it to get one summary per "
           "row, such as a count or a sum.",
      like="Expanding is like unpacking every item from each parcel onto the table; aggregating is like just "
           "writing 'contains 14 items' on the parcel.",
      how=["Click the expand icon on the merged column.", "Choose Expand (columns) or Aggregate (Count, Sum…).",
           "Check the row count afterwards."],
      tip="Aggregate when you want one row per customer.",
      trap="Expanding a one-to-many merge and then summing, which repeats values.",
      case="The analyst merges customers (2,500) with their transactions (left outer).",
      steps=["Expand the transactions column.", "Undo; aggregate it instead: Count and Sum of total_price.",
             "Compare row counts."],
      table=[["Action", "Rows", "Example: CUST00985"], ["Expand", "11,027", "14 rows, one per order"],
             ["Aggregate", "2,500", "1 row: 14 orders, ₦841,950"]],
      so="Expanding gave 11,027 rows (27 customers have no orders). Aggregating kept one row per customer.",
      split=4.20),

    O("3.4.10",
      what="If the key is **not unique** on either side, a merge **multiplies rows**: each row matches every "
           "matching row on the other side. Totals then grow far beyond the truth. This is the **Cartesian "
           "product trap**.",
      like="It's like introducing every guest to every other guest instead of to their partner. The number of "
           "handshakes explodes.",
      how=["Join on a key that is unique on at least one side.", "Compare rows in and out.",
           "If rows grow, stop and check the key."],
      tip="If a merge increases your row count, find out why before going on.",
      trap="Joining on a shared but non-unique column such as customer_id or store name.",
      case="An analyst joins transactions to deliveries on customer_id instead of delivery_id.",
      steps=["Join on delivery_id; count rows and sum revenue.", "Join on customer_id; count rows and sum revenue.",
             "Compare."],
      table=[["Join key", "Rows", "Revenue summed"], ["delivery_id (correct)", "10,430", "₦1,317,681,100"],
             ["customer_id (wrong)", "56,725", "₦7,191,095,600"]],
      so="The wrong key created 56,725 rows and over five times the real revenue, with no error message.",
      split=4.20),

    O("3.4.11",
      what="**Home > Append Queries** stacks two tables. Columns are matched by **name**, not position. A "
           "column called 'Sales' in one file and 'Sales_NGN' in the other becomes **two** half-empty columns. "
           "Rename first.",
      like="Appending is like stacking two piles of forms: it only works if each box on the form has the same "
           "label.",
      how=["Check both tables have the same column names.", "Home > Append Queries (as New).",
           "Check: rows out = rows in table 1 + table 2."],
      tip="Check the total: appended rows must equal the sum of the inputs.",
      trap="Assuming column order matters; only names do.",
      who="SuperMart Retail",
      case="The analyst appends the January and February 2025 sales files from SuperMart_Monthly_Sales_2025.",
      steps=["Load 2025-01 (155 rows) and 2025-02 (107 rows).", "Append Queries as New.", "Count the rows."],
      table=[["Table", "Rows", "Columns"], ["2025-01_SuperMart_Sales.csv", "155", "Same names"],
             ["2025-02_SuperMart_Sales.csv", "107", "Same names"], ["Appended", "262", "No new columns"]],
      so="155 + 107 = 262 and no extra columns appeared, so the append is correct.",
      split=4.20),

    O("3.4.12",
      what="To combine many tables, **Append Queries as New** accepts several at once. For files that arrive "
           "regularly, the **From Folder** connector appends every file in a folder, including future ones, "
           "automatically.",
      like="From Folder is like a letterbox: every new letter dropped in joins the pile with no extra effort.",
      how=["Few tables: Append Queries > Three or more tables.", "Regular files: Get Data > From Folder > Combine.",
           "Check the total against a control figure."],
      tip="Use From Folder for monthly files; new months join on refresh.",
      trap="Appending 12 queries by hand and forgetting to add month 13.",
      who="SuperMart Retail",
      case="All 12 monthly 2025 sales files must become one table that matches the POS control total.",
      steps=["Get Data > From Folder > SuperMart_Monthly_Sales_2025.", "Combine Files.", "Compare rows with the POS control."],
      table=[["Method", "Tables", "Result"], ["Manual append", "12 queries", "Must edit for each new month"],
             ["From Folder", "1 query", "3,312 rows = POS control total"]],
      so="One folder query gave all 3,312 rows and will pick up new files on refresh.",
      split=4.20),

    O("3.4.13",
      what="If appended files have **different columns**, Power Query keeps **all** columns and fills the gaps "
           "with **null**. There is no error, so check the **null percentage** of each column after appending.",
      like="It's like merging two class lists where one teacher wrote 'Surname' and the other 'Last name': you "
           "get two half-filled columns.",
      how=["After appending, turn on View > Column quality.", "Look for columns with unexpected empty %.",
           "Rename the odd column in its source query; refresh."],
      tip="Turn on Column quality after every append.",
      trap="Trusting the appended total when one month's values sit in a different column.",
      who="SuperMart Retail",
      case="Example: suppose the December file arrived with 'Sales_NGN' instead of 'Sales'. December has 462 rows.",
      steps=["Append all 12 files.", "Column quality shows Sales and Sales_NGN both partly empty.",
             "Rename Sales_NGN to Sales in the December step; refresh."],
      table=[["Column", "Empty before fix", "Empty after fix"], ["Sales", "462 rows (14%)", "0"],
             ["Sales_NGN", "2,850 rows (86%)", "Column gone"]],
      so="A renamed column hid all 462 December rows from Sales totals. Column quality caught it.",
      split=4.20),

    O("3.4.14",
      what="Merges fail on **invisible differences**: extra spaces, different capitals, or a number stored as "
           "text on one side. Clean **both** keys the same way (Trim, Clean, Upper, same data type) before "
           "merging.",
      like="It's like two keys cut almost the same: a tiny difference you can't see stops the lock turning.",
      how=["Trim and Clean the key in both tables.", "Make the case the same (Upper) if needed.",
           "Set the same data type on both keys."],
      tip="Count unmatched rows before and after cleaning the keys.",
      trap="Cleaning the key in one table but not the other.",
      who="SuperMart Retail",
      case="The analyst merges SuperMart_Sales_2025_DIRTY.csv to a 4-row zone lookup on Zone.",
      steps=["Merge on Zone as it is; count unmatched rows.", "Trim and Clean Zone in both tables.",
             "Merge again and count unmatched rows."],
      table=[["Check", "Before cleaning", "After cleaning"], ["Distinct Zone values", "12", "4"],
             ["Rows that fail to match", "268", "0"]],
      so="268 rows had spaces around Zone and failed to match. After Trim, all rows match one of the 4 zones.",
      split=4.20),

    O("3.4.15",
      what="Before trusting a merge, check **cardinality**: is the key unique on each side? Then compare rows in "
           "and out. More rows out means a **duplicated key**; fewer means **unmatched rows** were dropped.",
      like="It's like counting passengers on and off a bus: if more get off than got on, something is wrong.",
      how=["Compare COUNT(*) and COUNT(DISTINCT key) on both tables.", "Note rows before the merge.",
           "Compare rows after the merge."],
      tip="Rows = distinct keys on both sides means one-to-one.",
      trap="Skipping the check because the merge 'looked fine'.",
      case="Before merging transactions with deliveries on order_id, the analyst checks the keys.",
      steps=["COUNT(*) and COUNT(DISTINCT order_id) for each table.", "Left join on order_id.", "Compare rows in and out."],
      table=[["Table", "Rows", "Distinct order_id"], ["transactions", "11,000", "11,000"],
             ["deliveries", "10,430", "10,430"], ["Left join result", "11,000", "Same as left: safe"]],
      so="Keys are unique on both sides, so this is one-to-one and the left join keeps exactly 11,000 rows.",
      split=4.20),

    O("3.4.16",
      what="Document every merge and append: **step name**, **join kind**, **key**, **expected cardinality** "
           "and **row counts**. Put it in the 00_README sheet so the next analyst knows what each step does and "
           "what result is normal.",
      like="It's like labelling the wires in a fuse box: whoever opens it next knows what each one does.",
      how=["Rename steps clearly: Merged Orders to Deliveries.", "Record kind, key and expected counts.",
           "Add the log to 00_README."],
      tip="Record the expected row count so changes stand out.",
      trap="Leaving steps called 'Merged Queries1'.",
      case="The analyst writes a merge log for this module's main merge.",
      steps=["Rename the step.", "Write kind, key and cardinality.", "Write rows in and out."],
      table=[["Item", "Entry"], ["Step name", "Merged Orders to Deliveries"], ["Join kind / key", "Left outer on delivery_id"],
             ["Cardinality", "One-to-one"], ["Rows in / out", "11,000 / 11,000 (570 blank couriers)"]],
      so="If next month's rows out ever exceed rows in, the log shows something has changed.",
      split=4.90),
]
