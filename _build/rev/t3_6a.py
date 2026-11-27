from revlib import T

ON = "SuperMart Online"
TRY = "ETEH_Week3_Section_Examples.xlsx  ›  sheet '3.6A DAX Foundations'  ·  model from sheet 3.5"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", ON)
    return T(tag, **k)


TOPICS = [
    O("3.6.1",
      what="**Power Pivot** is the engine behind Excel's Data Model. It holds many related tables in memory, "
           "goes far beyond the sheet's 1,048,576-row limit, and adds **DAX**, a formula language for "
           "**measures** that work across tables.",
      like="If a normal PivotTable is a calculator, Power Pivot is a small accounting office: many ledgers, linked "
           "together, with staff (DAX) who can answer new questions.",
      how=["Load the tables into the Data Model.", "Link them with relationships.",
           "Write DAX measures and use them in PivotTables."],
      tip="Think 'tables + relationships + measures'.",
      trap="Treating Power Pivot as just a bigger worksheet.",
      case="The analyst builds the SuperMart Online model from the D_ sheets of the Week 3 workbook.",
      steps=["Add the six tables to the Data Model.", "Create the relationships from sheet 3.5.", "Check the row counts loaded."],
      table=[["Table", "Rows", "Role"], ["transactions", "11,000", "Fact"], ["deliveries", "10,430", "Fact details"],
             ["customers / products / sellers", "2,500 / 800 / 200", "Dimensions"], ["DimDate", "2,192", "Calendar"]],
      so="Six linked tables now sit in one model, ready for DAX measures.",
      split=4.20),

    O("3.6.2",
      what="The Data Model is always there, but the **Power Pivot tab** is hidden by default. Turn it on with "
           "**File > Options > Add-ins > COM Add-ins > Go** and tick **Microsoft Power Pivot for Excel**. You need "
           "it for Diagram View, calculated columns and Mark as Date Table.",
      like="It's like a room in your house with the door locked: the room exists, you just need to unlock it to go in.",
      how=["File > Options > Add-ins.", "Manage: COM Add-ins > Go.", "Tick Microsoft Power Pivot for Excel > OK."],
      tip="If the tab vanishes after a crash, re-tick the add-in.",
      trap="Thinking your Excel 'doesn't have' the Data Model because the tab is missing.",
      case="A learner can't find the Power Pivot tab needed for this module.",
      steps=["Open File > Options > Add-ins.", "Choose COM Add-ins > Go.", "Tick Power Pivot; the tab appears."],
      table=[["Task", "Needs the Power Pivot tab?"], ["Add tables via Power Query or PivotTable", "No"],
             ["Diagram View and relationships drawing", "Yes"], ["Calculated columns, Mark as Date Table", "Yes"]],
      so="Enabling the add-in once gives full access to the model for the rest of the module.",
      split=4.90),

    O("3.6.3",
      what="Tables enter the model three ways: **Power Query** (Close & Load To > Add this data to the Data "
           "Model), **Power Pivot > Add to Data Model** from a sheet table, or **Insert > PivotTable > Add this "
           "data to the Data Model**. Power Query is best, because cleaning steps are kept.",
      like="It's like stocking a pantry: buying from a trusted shop with a receipt (Power Query) is better than "
           "dropping in loose items with no record.",
      how=["Clean the data in Power Query.", "Close & Load To > Only Create Connection.",
           "Tick Add this data to the Data Model."],
      tip="Use 'Only Create Connection' so big tables don't fill worksheets.",
      trap="Loading the same table twice: once to a sheet and once to the model.",
      case="The analyst loads transactions (11,000 rows) into the model.",
      steps=["Get Data from the D_Transactions table.", "Close & Load To > Only Create Connection + Data Model.",
             "Check the row count in Power Pivot."],
      table=[["Route", "Cleaning kept?", "Recommended"], ["Power Query > Data Model", "Yes", "Yes"],
             ["Power Pivot > Add to Data Model", "No steps", "For quick tests"],
             ["PivotTable > Add to Data Model", "No steps", "For quick tests"]],
      so="All 11,000 rows sit in the model without taking a single worksheet row.",
      split=4.20),

    O("3.6.4",
      what="A **calculated column** works row by row and stores a value in every row, like a helper column. A "
           "**measure** calculates a single answer on demand for whatever the PivotTable is showing. Use columns "
           "to slice or filter; use measures for totals and ratios.",
      like="A calculated column is like writing a label on every item in the warehouse. A measure is like asking "
           "the stock-keeper 'how much is on this shelf?' whenever you need to know.",
      how=["Need a value to put on Rows, Columns or a slicer? Calculated column.",
           "Need a number in Values? Measure.", "Prefer measures; they use less memory."],
      tip="If you're going to add it up, make it a measure.",
      trap="Making a calculated column for every total; the file grows and nothing gets faster.",
      case="The analyst needs a 'net order?' flag for slicing and a Net Revenue figure for reporting.",
      steps=["Column: Is Net = Yes if order_status is Delivered or Completed.", "Measure: Net Revenue := CALCULATE(…).",
             "Compare what each produces."],
      table=[["Aspect", "Calculated column", "Measure"], ["Example", "Is Net (Yes/No)", "Net Revenue"],
             ["Stored", "11,000 values", "Nothing; calculated live"],
             ["Use in PivotTable", "Rows / slicer", "Values"], ["Result", "9,919 Yes, 1,081 No", "₦1,256,619,400"]],
      so="The column labels each row once; the measure answers for whatever cell you look at.",
      split=4.20),

    O("3.6.5",
      what="DAX looks like Excel but refers to **tables and columns**, not cells: **tTransactions[total_price]**. "
           "Measures are named with **:=** and referenced in brackets, e.g. **[Net Revenue]**. Functions like "
           "SUM take a column; others take a whole table.",
      like="Excel formulas point at seats in a cinema (B2, C5); DAX points at whole rows of seats by name.",
      how=["Columns: 'Table'[Column] (quotes only if the name has spaces).", "Measures: Name := expression.",
           "Reuse measures by name: [Gross Bookings]."],
      tip="Always write the table name before a column.",
      trap="Typing cell references like B2 in DAX.",
      case="The analyst writes the first measures on sheet 3.6A.",
      steps=["Gross Bookings := SUM ( tTransactions[total_price] ).", "Orders := COUNTROWS ( tTransactions ).",
             "Reuse them: Avg := DIVIDE ( [Gross Bookings], [Orders] )."],
      table=[["Element", "Example", "Result"], ["Column reference", "tTransactions[total_price]", "One column"],
             ["Measure", "Gross Bookings := SUM ( … )", "₦1,394,713,700"],
             ["Measure reuse", "[Gross Bookings]", "Same answer, reused"], ["Table function", "COUNTROWS ( tTransactions )", "11,000"]],
      so="Columns are named, measures are bracketed and reused, so formulas read almost like sentences.",
      split=4.20),

    O("3.6.6",
      what="**Row context** means 'the current row'. It exists in **calculated columns** and inside **iterator** "
           "functions like SUMX. A plain measure has **no** current row, so referring to a bare column there "
           "gives an error.",
      like="Row context is like a teacher marking one exam paper at a time: 'this paper's score' makes sense. "
           "Asked for 'the score' of the whole class, the teacher needs to know: total? average?",
      how=["In a calculated column, refer to columns directly.", "In a measure, wrap columns in SUM, MAX, SUMX…",
           "Use SUMX when a measure needs a row-by-row calculation."],
      tip="In a measure, every column needs an aggregation or an iterator.",
      trap="Writing a measure like := tTransactions[quantity] with no aggregation.",
      case="The analyst writes quantity × unit_price as a calculated column, then tries the same as a measure.",
      steps=["Calculated column: Line Rev = [quantity] * [unit_price].", "Look at TRX004782: 5 × ₦720,000.",
             "Try a measure := tTransactions[quantity]."],
      table=[["Where", "Formula", "Result"], ["Calculated column", "[quantity] * [unit_price]", "₦3,600,000 for TRX004782"],
             ["Measure (no aggregation)", "tTransactions[quantity]", "Error: a single value can't be determined"],
             ["Measure with iterator", "SUMX ( tTransactions, … )", "₦1,394,713,700"]],
      so="The column had a current row; the measure didn't until SUMX supplied one.",
      split=4.20),

    O("3.6.7",
      what="**Filter context** is the set of filters active when a measure is calculated: the PivotTable's row "
           "and column labels, slicers and report filters. The same measure gives a **different answer in each "
           "cell** because each cell has different filters.",
      like="Filter context is like asking 'how many students?' in different rooms: the question is the same, the "
           "answer depends on which room you're standing in.",
      how=["Put the measure in Values.", "Add a field to Rows: each row filters the data first.",
           "Add slicers to filter further."],
      tip="Read each cell as 'the measure, for these filters'.",
      trap="Thinking a measure stores numbers per category; it recalculates for each cell.",
      case="[Net Revenue] is placed in a PivotTable with product category on Rows.",
      steps=["Values: [Net Revenue].", "Rows: tProducts[category].", "Check the cells against the SUMIFS twin on sheet 3.6A."],
      table=[["Category (filter)", "Net orders", "[Net Revenue]"], ["Electronics", "1,446", "₦648,437,900"],
             ["Home", "1,279", "₦327,026,350"], ["Beauty (lowest)", "1,123", "₦10,542,400"],
             ["Grand total (no filter)", "9,919", "₦1,256,619,400"]],
      so="One measure gave a different answer in each row, because each row filtered the data to one category.",
      split=4.20),

    O("3.6.8",
      what="**CALCULATE** runs a measure with **changed filters**. It is the most important DAX function. Each "
           "filter you give it **replaces** any existing filter on that column, which is powerful but can surprise you.",
      like="CALCULATE is like a manager saying 'give me that figure, but only for delivered orders': same "
           "question, with a rule added.",
      how=["CALCULATE ( [measure], filter1, filter2… ).",
           "Net Revenue := CALCULATE ( [Gross Bookings], tTransactions[order_status] IN { \"Delivered\", \"Completed\" } ).",
           "Compare with the measure alone."],
      tip="Build simple measures first, then wrap them in CALCULATE.",
      trap="Putting order_status on Rows with this measure: every row shows the same net total.",
      case="Gross Bookings includes cancelled and returned orders. Finance wants net revenue.",
      steps=["Gross Bookings := SUM ( tTransactions[total_price] ).", "Net Revenue := CALCULATE ( [Gross Bookings], status IN {…} ).",
             "Find the difference."],
      table=[["Measure", "Orders", "Value"], ["Gross Bookings", "11,000", "₦1,394,713,700"],
             ["Net Revenue (Delivered + Completed)", "9,919", "₦1,256,619,400"],
             ["Removed: Cancelled + Returned", "1,081", "₦138,094,300"]],
      so="CALCULATE removed ₦138.1m of cancelled and returned orders by changing the filter.",
      split=4.20),

    O("3.6.9",
      what="**FILTER ( table, condition )** returns the rows that meet a condition, checking them one by one. Use "
           "it inside CALCULATE when a simple filter isn't enough, such as conditions on a measure or on several "
           "columns.",
      like="FILTER is like a sieve: you pour the table through and keep only the rows that pass.",
      how=["CALCULATE ( [measure], FILTER ( table, condition ) ).", "Filter the smallest table that works.",
           "Prefer simple column filters when they do the job."],
      tip="Filter a dimension table (e.g. tProducts) rather than the big fact table where possible.",
      trap="Using FILTER on the whole fact table for a simple equality test.",
      case="Management asks for Electronics net revenue, and for net revenue from big orders over ₦500,000.",
      steps=["Elec Rev := CALCULATE ( [Net Revenue], FILTER ( tProducts, [category] = \"Electronics\" ) ).",
             "Big Rev := CALCULATE ( [Net Revenue], FILTER ( tTransactions, [total_price] > 500000 ) ).",
             "Compare with total net revenue."],
      table=[["Measure", "Net orders", "Net revenue"], ["Electronics only", "1,446", "₦648,437,900"],
             ["Orders over ₦500,000", "672", "₦647,224,100"], ["All net orders", "9,919", "₦1,256,619,400"]],
      so="Just 672 big orders bring in about half of all net revenue.",
      split=4.20),

    O("3.6.10",
      what="**ALL ( column or table )** **removes filters**. Divide a measure by the same measure with ALL to get "
           "a **share of the total**: each category's part of the whole.",
      like="ALL is like stepping out of your room to count everyone in the building, so you can say what share "
           "your room holds.",
      how=["Total := CALCULATE ( [Net Revenue], ALL ( tProducts[category] ) ).", "Cat Share := DIVIDE ( [Net Revenue], [Total] ).",
           "Format as %."],
      tip="Name the share measure clearly: '% of all categories'.",
      trap="Using ALL on the whole table when you only meant one column; it removes more filters than intended.",
      case="Management wants each category's share of net revenue.",
      steps=["Rows: category; Values: [Net Revenue].", "Add [Cat Share].", "Check the shares add up to 100%."],
      table=[["Category", "Net revenue", "Cat Share"], ["Electronics", "₦648,437,900", "51.6%"],
             ["Home", "₦327,026,350", "26.0%"], ["Fashion", "₦110,075,900", "8.8%"], ["Beauty", "₦10,542,400", "0.8%"]],
      so="Electronics alone is over half of net revenue; Electronics and Home together are 77.6%.",
      split=4.20),

    O("3.6.11",
      what="**ALLEXCEPT ( table, column )** removes **every filter on that table except the column(s) you name**. "
           "Use it for 'share of the parent': each tier's share **of its own zone**, not of the whole company.",
      like="It's like comparing a class's score to its own school's total, not to the whole country's.",
      how=["Zone Rev := CALCULATE ( [Net Revenue], ALLEXCEPT ( tCustomers, tCustomers[zone] ) ).",
           "Share of Zone := DIVIDE ( [Net Revenue], [Zone Rev] ).", "Use a Zone × Tier PivotTable."],
      tip="Keep the parent column; drop the child.",
      trap="Using ALL here, which compares to the national total instead.",
      case="In a Zone × Tier PivotTable, the analyst wants each tier's share of its zone. South-West is shown.",
      steps=["Rows: zone, then membership_tier.", "Add [Zone Rev]: every South-West row shows ₦339,893,950.",
             "Add [Share of Zone]."],
      table=[["South-West tier", "Net revenue", "Share of zone"], ["Gold", "₦102,914,800", "30.3%"],
             ["Silver", "₦100,049,950", "29.4%"], ["Platinum", "₦68,660,950", "20.2%"], ["Bronze", "₦68,268,250", "20.1%"]],
      so="Gold leads within South-West. ALLEXCEPT kept the zone filter and dropped the tier filter.",
      split=4.20),

    O("3.6.12",
      what="**VALUES ( column )** returns the **distinct values visible** in the current filters. Wrap it in "
           "COUNTROWS to count them, or use it for dynamic titles. **DISTINCT** is almost the same but ignores the "
           "extra 'blank' row for unmatched keys.",
      like="VALUES is like asking 'which subjects are on today's timetable?' The answer changes with the day.",
      how=["Cats := COUNTROWS ( VALUES ( tProducts[category] ) ).", "Put it in a PivotTable by seller.",
           "Group the sellers by the answer."],
      tip="Use VALUES whenever you need 'which ones are showing right now'.",
      trap="Expecting VALUES to ignore slicers; it respects all current filters.",
      case="The marketplace team asks: how many categories does each seller sell?",
      steps=["Grand total: [Cats] = 7.", "Rows: seller_id; Values: [Cats].", "Count sellers by result."],
      table=[["Categories sold", "Sellers"], ["1", "19"], ["2 or 3", "114"], ["4 or 5", "59"], ["6 (none sells all 7)", "5"]],
      so="The same measure showed 7 overall but 1 to 6 per seller. Most sellers specialise in 2 or 3 categories.",
      split=4.90),

    O("3.6.13",
      what="**RELATED ( column )** fetches a value from the **'one' side** of a relationship while you're on the "
           "**'many' side**, in a calculated column or iterator. It's the model's version of VLOOKUP, with no "
           "column numbers.",
      like="RELATED is like reading a student's class teacher from the register: you follow the link, you don't "
           "search for it.",
      how=["Make sure the relationship exists.", "In tTransactions, add a column: Category = RELATED ( tProducts[category] ).",
           "Use it in row-by-row formulas."],
      tip="Prefer using the dimension column directly in PivotTables; use RELATED inside formulas.",
      trap="Using RELATED without a relationship, which gives an error.",
      case="The analyst wants each transaction's category inside a row-by-row formula.",
      steps=["Check transactions → products on product_id.", "Add Category = RELATED ( tProducts[category] ).",
             "Count rows by the new column."],
      table=[["Category (from RELATED)", "Order lines", "Share of lines"], ["Health (most lines)", "1,852", "16.8%"],
             ["Automotive", "1,760", "16.0%"], ["Beauty (fewest)", "1,239", "11.3%"],
             ["All 7 categories", "11,000", "100% (no blanks)"]],
      so="Every one of the 11,000 rows found its category through the relationship, with no lookup formula.",
      split=4.20),

    O("3.6.14",
      what="**RELATEDTABLE ( table )** goes the other way: from a row on the **'one' side**, it returns all the "
           "**related rows** on the many side. Use it to count or sum each customer's or product's transactions.",
      like="It's like looking up a teacher and getting the list of all their students.",
      how=["In tCustomers, add: Order Count = COUNTROWS ( RELATEDTABLE ( tTransactions ) ).",
           "Sort to find the highest.", "Count customers with blank (no orders)."],
      tip="Use it for dimension-level labels like 'repeat customer'.",
      trap="Using it in a measure when a simple measure by customer would do.",
      case="Marketing wants each customer's order count stored on the customer table.",
      steps=["Add the Order Count column to tCustomers.", "Find the maximum.", "Count customers with no orders."],
      table=[["Check", "Result"], ["Highest order count", "14 (CUST00985)"], ["Average per customer", "4.4 orders"],
             ["Customers with no orders", "27 (blank)"]],
      so="Each of the 2,500 customers now carries an order count, ready for slicing.",
      split=4.90),

    O("3.6.15",
      what="**SUMX ( table, expression )** works **row by row**: it calculates the expression for each row, then "
           "adds the results. Use it when you must multiply or combine columns **before** adding up.",
      like="SUMX is like a cashier ringing up each item (price × quantity) and then totalling the receipt, instead "
           "of adding up all the prices and all the quantities separately.",
      how=["Revenue X := SUMX ( tTransactions, [quantity] * [unit_price] ).", "Compare with SUM ( [total_price] ).",
           "Never add up prices on their own."],
      tip="Use SUMX whenever the calculation is 'per row, then total'.",
      trap="Using SUM ( unit_price ), which adds prices and means nothing.",
      case="The analyst rebuilds revenue from quantity and unit price.",
      steps=["SUMX ( tTransactions, quantity × unit_price ).", "Compare with SUM ( total_price ).", "Compare with SUM ( unit_price )."],
      table=[["Measure", "Result", "Meaning"], ["SUMX ( qty × unit_price )", "₦1,394,713,700", "Correct revenue"],
             ["SUM ( total_price )", "₦1,394,713,700", "Same; good check"], ["SUM ( unit_price )", "₦780,554,500", "Meaningless"]],
      so="SUMX matched the stored totals exactly; adding prices alone gave a meaningless number.",
      split=4.20),

    O("3.6.16",
      what="**AVERAGEX ( table, expression )** calculates something for each row of a table, then averages the "
           "results. The table you choose sets the **grain**: average **per order** and average **per customer** "
           "are very different.",
      like="It's like asking for the average spend per shopping trip versus per shopper: a shopper makes several trips.",
      how=["Per order: AVERAGEX over net orders.", "Per customer: AVERAGEX ( VALUES ( customer_id ), [Net Revenue] ).",
           "Label which grain each measure uses."],
      tip="Say 'average per ___' in every average's name.",
      trap="Calling both of these 'average spend' without saying the grain.",
      case="Marketing asks for 'average spend'. The analyst calculates it per order and per customer.",
      steps=["AOV = average of total_price over net orders.", "Avg per customer = average of each customer's net revenue.",
             "Compare."],
      table=[["Measure", "Grain", "Result"], ["Average items per order", "Order line", "1.79"],
             ["Average net order value", "Net order (9,919)", "₦126,688"],
             ["Average net spend per customer", "Customer (2,456)", "₦511,653"]],
      so="Per customer is about 4 times per order. Always say which average you mean.",
      split=4.20),

    O("3.6.17",
      what="**COUNTROWS ( table )** counts the rows in the current filters. It's the clearest way to count events "
           "like orders or admissions. **COUNT ( column )** counts only **non-blank** values, so it can quietly "
           "give a smaller number.",
      like="COUNTROWS counts people in the room; COUNT counts only those wearing name badges.",
      how=["Orders := COUNTROWS ( tTransactions ).", "Compare with COUNT ( tTransactions[delivery_id] ).",
           "Use COUNTROWS for event counts."],
      tip="Count events with COUNTROWS.",
      trap="Counting orders with COUNT on a column that has blanks.",
      case="Two analysts report different order totals for the same data.",
      steps=["Analyst A: COUNTROWS ( tTransactions ).", "Analyst B: COUNT ( tTransactions[delivery_id] ).", "Explain the gap."],
      table=[["Measure", "Result", "Why"], ["COUNTROWS ( tTransactions )", "11,000", "Every order"],
             ["COUNT ( [delivery_id] )", "10,430", "Skips 570 blank delivery_ids"]],
      so="The 570 missing orders were cancellations with no delivery. COUNTROWS counted them; COUNT didn't.",
      split=4.20),

    O("3.6.18",
      what="**DISTINCTCOUNT ( column )** counts **unique** values: how many different customers, patients or "
           "students. Distinct counts **don't add up** across groups, because one customer can appear in several "
           "groups.",
      like="It's like counting people at a market: one shopper visiting five stalls is still one shopper.",
      how=["Customers := DISTINCTCOUNT ( tTransactions[customer_id] ).", "Compare with COUNTROWS.",
           "Never add distinct counts from different groups."],
      tip="Use the grand total from the measure, never a sum of the rows.",
      trap="Adding category headcounts to get 'total customers'.",
      case="Management asks how many customers SuperMart Online has, and how many buy each category.",
      steps=["[Orders] = 11,000; [Customers] = 2,473.", "Rows: category; Values: [Customers].", "Add up the category rows."],
      table=[["Measure", "Value", "Meaning"], ["Orders (COUNTROWS)", "11,000", "Order lines"],
             ["Customers (DISTINCTCOUNT)", "2,473", "Unique buyers"],
             ["Sum of 7 category headcounts", "8,203", "Wrong: people counted repeatedly"]],
      so="2,473 people placed 11,000 orders. Adding category headcounts would claim 8,203 customers.",
      split=4.20),

    O("3.6.19",
      what="**DIVIDE ( numerator, denominator, [alternate] )** is safe division: if the denominator is 0 or blank "
           "it returns **blank** (or your alternate) instead of **#DIV/0!**.",
      like="DIVIDE is like a polite waiter who says 'none today' instead of dropping the tray when a dish runs out.",
      how=["AOV := DIVIDE ( [Net Revenue], [Net Orders] ).", "Add a third argument for a default, e.g. 0.",
           "Use it for every ratio."],
      tip="Use DIVIDE for every ratio in a model.",
      trap="Using / and getting errors when a slicer selects something with no rows.",
      case="A PivotTable shows average order value by seller. Some sellers have no orders.",
      steps=["Rows: seller_id; Values: [AOV].", "Find SELL0046, a seller with no orders.", "Compare / with DIVIDE."],
      table=[["Seller", "Net orders", "[AOV] with DIVIDE"], ["All sellers", "9,919", "₦126,688"],
             ["SELL0046 (no sales)", "0", "Blank, not #DIV/0!"]],
      so="DIVIDE kept the report clean for the 3 sellers with no orders.",
      split=4.20),

    O("3.6.20",
      what="**IF** and **SWITCH** add decisions to measures. **SWITCH ( TRUE (), test1, result1, test2, result2, "
           "default )** is the tidy way to write three or more bands, like IFS in Excel.",
      like="SWITCH is like a sorting table at the post office: check each rule in order, first match decides the tray.",
      how=["Size := SWITCH ( TRUE (), [AOV] >= 250000, \"Large\", [AOV] >= 50000, \"Medium\", \"Small\" ).",
           "Order the tests from highest to lowest.", "Always give a default."],
      tip="Put the strictest test first.",
      trap="Ordering the tests from lowest to highest, so everything matches the first test.",
      case="The analyst labels each category by its average net order value.",
      steps=["Rows: category; Values: [AOV], [Size].", "Read the labels.", "Check the grand total label."],
      table=[["Category", "AOV", "Size"], ["Electronics", "₦448,436", "Large"], ["Home", "₦255,689", "Large"],
             ["Fashion / Automotive", "₦70,381 / ₦54,061", "Medium"], ["Food / Health / Beauty", "₦39,466 or less", "Small"],
             ["All categories", "₦126,688", "Medium"]],
      so="Two categories are 'Large' ticket items; three are 'Small'. The overall label is Medium.",
      split=4.20),

    O("3.6.21",
      what="**VAR name = expression … RETURN result** stores an intermediate result once and reuses it. Measures "
           "become easier to read and faster, because the expression is not calculated twice.",
      like="VAR is like writing a sub-total on scrap paper so you don't redo the addition every time you need it.",
      how=["VAR rev = [Gross Bookings].", "VAR comm = SUM ( tTransactions[commission_amount] ).",
           "RETURN DIVIDE ( comm, rev )."],
      tip="Give variables meaningful names.",
      trap="Expecting a VAR to recalculate after CALCULATE changes the filters; it's fixed where defined.",
      case="Finance wants commission as a percentage of gross bookings.",
      steps=["Store gross bookings in rev.", "Store total commission in comm.", "Return comm ÷ rev."],
      table=[["Part", "Value"], ["rev (Gross Bookings)", "₦1,394,713,700"], ["comm (commission)", "₦143,638,048.75"],
             ["Comm % (RETURN)", "10.3%"]],
      so="Commission is about 10.3% of bookings, and the measure reads like a short calculation.",
      split=4.90),
]
