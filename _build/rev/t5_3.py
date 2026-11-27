from revlib import T

TRY = "ETEH_Capstone_Section_Examples.xlsx  ›  sheet 'P.3 Capstone1 SuperMart'"
SOL = "03_Instructor_Projects  ›  Capstone1_SuperMart_Stores_SOLUTION.xlsx  ›  "


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", "SuperMart Retail · Capstone 1")
    return T(tag, **k)


TOPICS = [
    O("P.3.1",
      what="Capstone 1 starts with the board's question: FY2025 sales grew, so **why did the margin fall**, which "
           "stores and products caused it, and what should change in FY2026? **Success** means reconciled numbers, a "
           "measured cause and a costed recommendation.",
      like="Like a shop owner asking: 'We're selling more than ever, so why is there less money in the drawer?'",
      how=["Compute the headline numbers first.", "Write the question and success criteria.", "Get the brief signed."],
      tip="Margin = profit ÷ sales. Always compare totals, never averages of line margins.",
      trap="Celebrating sales growth without looking at profit.",
      case="The headline numbers, as live formulas over the ledger.",
      steps=["SUMIFS sales by year.", "Margin = SUMIFS profit ÷ SUMIFS sales.", "Change in points × 100."],
      table=[["Measure", "FY2024", "FY2025"], ["Sales", "₦609,205,599", "₦733,215,255"], ["Growth", "n/a", "+20.4%"],
             ["Margin", "13.43%", "12.74%"], ["Margin change", "n/a", "−0.68 points"]],
      so="Sales rose 20.4% but margin fell about 0.7 points. That gap is the project.",
      split=4.70),

    O("P.3.2",
      what="A **data inventory** lists each source with its **grain** (what one row means), period and size. "
           "Knowing the grain stops you mixing up lines, orders and customers.",
      like="Like counting a market stall's stock: crates, not tomatoes, unless you say which.",
      how=["List every source.", "Write what one row means.", "Record rows, period and owner."],
      tip="One order can have several lines. Count orders with DISTINCTCOUNT(Order_ID).",
      trap="Calling 3,312 lines '3,312 orders'. They are only 1,687 orders.",
      case="Capstone 1 inventory.",
      steps=["Name the sources.", "State the grain.", "Count rows and orders."],
      table=[["Source", "Grain", "Scale"], ["POS ledger", "One sales line", "9,994 lines; 5,009 orders; 2022–2025"],
             ["Products (from ledger)", "One SKU", "1,862 SKUs; 3 categories; 17 sub-categories"],
             ["Store directory", "One store", "12 stores; 4 zones; FY2025 targets"],
             ["Dirty store returns", "One sales line", "3,387 rows → 3,312 after cleaning"]],
      so="With the grain written down, nobody confuses lines with orders.",
      split=4.20),

    O("P.3.3",
      what="The dirty store returns have **16 kinds of defect**: split columns, extra spaces, wrong case, ₦ and % "
           "typed as text, mixed dates, blanks, duplicates and more. Fix each with a named step and log the count.",
      like="Like sorting a sack of beans: remove stones, dust and broken ones, and count what you removed.",
      how=["List each defect and its count.", "Fix it with a named Power Query step.", "Reconcile to the ledger at the end."],
      tip="Remove exact duplicates on all columns, not on Order_ID alone: one order can have several lines.",
      trap="Treating a blank Quantity as 0.",
      case="Part of the defect log for SuperMart_Sales_2025_DIRTY.csv.",
      steps=["Log the defect.", "Apply the fix.", "Check the controls."],
      table=[["Defect", "Rows", "Fix"], ["Mixed date formats", "331", "Change Type with Locale (en-GB)"],
             ["Sales typed as '₦1,234,560'", "231", "Remove ₦ and commas, then number"],
             ["Discount typed as '20%'", "165", "Convert to 0.2"], ["Sales × 100 keying errors", "6", "Correct from receipt"],
             ["Duplicates + spacer rows", "60 + 15", "Remove: 3,387 → 3,312"]],
      so="After cleaning, rows, sales, profit, quantity, orders (1,687) and customers (693) all match the controls.",
      tryit=SOL + "'2 Clean & Reconcile'",
      split=4.70),

    O("P.3.4",
      what="The model links the sales lines to the store table on **Store_ID** and uses explicit measures: Sales, "
           "Profit, **Margin %** (total profit ÷ total sales), Orders, Growth % and Target Attainment.",
      like="Like a school report: one formula for the average, used the same way for every student.",
      how=["Relate FactSales to DimStore and DimDate.", "Write each measure once.", "Reconcile to the ledger."],
      tip="Use DIVIDE() in DAX so a zero never shows an error.",
      trap="Averaging line margins instead of dividing total profit by total sales.",
      case="Capstone 1 measures, FY2025.",
      steps=["Write the DAX.", "Evaluate for FY2025.", "Compare with the ledger."],
      table=[["Measure", "DAX", "FY2025"], ["Profit", "SUM(FactSales[Profit])", "₦93,439,242"],
             ["Growth %", "DIVIDE([Revenue]-[Revenue LY],[Revenue LY])", "20.4%"],
             ["Discount depth", "Sales-weighted discount", "13.8%"], ["Target attainment", "Revenue ÷ sum of targets", "98.2%"]],
      so="The chain as a whole reached 98.2% of target, even with 20.4% growth.",
      split=4.20),

    O("P.3.5",
      what="**Basket analysis** looks at whole orders: how big a typical order is and which categories are bought "
           "together. Add lines up to **order level** first, then measure.",
      like="Like checking what's in each shopper's bag at the till, not counting items on the shelf.",
      how=["Group lines by Order_ID.", "Count lines and sum value per order.", "Compare mean and median."],
      tip="When the mean is far above the median, a few big orders are pulling it up.",
      trap="Calling the average line value the average basket.",
      case="FY2025 orders.",
      steps=["Build one row per order.", "Mean and median basket.", "Count orders with 2+ categories."],
      table=[["Measure", "Value"], ["Orders", "1,687"], ["Lines per order (mean / most common)", "1.96 / 1"],
             ["Basket value (mean / median)", "₦434,627 / ₦148,257"],
             ["Orders with Office Supplies", "75.4%"], ["Orders with 2+ categories", "33.2%"]],
      so="Most shoppers buy one line. The mean is about 3× the median because of a few large orders.",
      split=4.90),

    O("P.3.6",
      what="To explain a margin fall, split it into a **rate effect** (margins fell inside categories) and a **mix "
           "effect** (sales moved between categories), and look at margin by **discount band**.",
      like="Like a trader asking: did each item earn less, or did I just sell more of the cheap items?",
      how=["Margin by discount band.", "Rate = share × change in margin.", "Mix = change in share × old margin."],
      tip="Rate + mix should add up to the total change. Check it.",
      trap="Blaming the product mix without measuring it.",
      case="FY2025 margin by discount band.",
      steps=["Band the discount.", "SUMIFS profit ÷ sales per band.", "Total the loss above 20%."],
      table=[["Discount band", "Share of sales", "Margin"], ["0%", "47.6%", "29.9%"], ["1–20%", "37.8%", "13.3%"],
             ["21–40%", "9.4%", "−17.6%"], ["Above 40%", "5.2%", "−92.5%"]],
      so="Lines above 20% discount lost ₦47.6m on 14.6% of sales. Rate effect −1.19 pts, mix +0.51.",
      tryit=SOL + "'4 Margin Diagnosis'",
      split=4.70),

    O("P.3.7",
      what="**Slow movers** are products that sell little or nothing. Compare the full product list with FY2025 "
           "sales using an **anti join** (items in one list but not the other).",
      like="Like a shop checking which goods have sat on the shelf all year.",
      how=["List all SKUs ever sold.", "Anti join with FY2025 sales.", "Flag loss-making SKUs too."],
      tip="State the window: 'no sale in FY2025', not just 'slow'.",
      trap="Calling a product slow from a few weeks of data.",
      case="SKU actions for FY2025.",
      steps=["Count SKUs.", "COUNTIFS for no FY2025 sale.", "COUNTIFS for negative profit."],
      table=[["Measure", "Value"], ["SKUs in the range", "1,862"], ["No sale in FY2025", "337"],
             ["…sold in FY2024 but not FY2025", "230"], ["Loss-making SKUs", "271 (₦38.3m lost)"],
             ["Revenue from bottom 30% of sellers", "1.3%"]],
      so="230 recently dead SKUs are clearance candidates; 271 loss-makers need repricing.",
      split=4.90),

    O("P.3.8",
      what="The Capstone 1 dashboard fits one screen: **KPI cards**, a **store league table**, margin by discount "
           "band and by quarter. Slicers for zone, category and segment drive every visual.",
      like="Like a scoreboard: the score, the table and the key moments, all at a glance.",
      how=["Four KPI cards at the top.", "Store table sorted by sales.", "Margin charts beside it."],
      tip="Show profit next to sales. Sales targets alone reward discounting.",
      trap="A store table showing sales only.",
      case="The store league table: five stores lost money.",
      steps=["Sort stores by FY2025 sales.", "Add profit and target attainment.", "Highlight negative profit."],
      table=[["Store", "FY2025 profit", "Target attainment"], ["Benin City - Ring Road", "−₦713,637", "105.3%"],
             ["Abuja - Garki", "−₦2,512,563", "96.0%"], ["Ibadan - Ring Road", "−₦4,118,850", "89.0%"],
             ["Onitsha - Main Market", "−₦6,220,517", "90.2%"], ["Abuja - Wuse", "−₦8,838,505", "83.5%"]],
      so="Benin City beat its sales target but still lost money. Profit must sit beside sales.",
      tryit=SOL + "'3 Dashboard'",
      split=4.70),

    O("P.3.9",
      what="The **insight memo** fits one page: a headline, three supporting numbers, the recommendation with an "
           "**owner** and **expected effect** shown as a range, and a note on the data.",
      like="Like a short letter to the landlord: the problem, the evidence, what you want, by when.",
      how=["Headline in one sentence.", "Three numbers as evidence.", "Recommendation with owner and range."],
      tip="Give a range, not a single number, when the effect depends on assumptions.",
      trap="A recommendation with no owner or cost.",
      case="Costing a 20% discount cap: lines above 20% are re-priced at 20%.",
      steps=["Count lines affected: 455.", "Recalculate profit for each volume-loss case.", "Report the range."],
      table=[["Volume lost on re-priced lines", "Profit gain"], ["0%", "₦71.5m"], ["25%", "₦65.5m"], ["50%", "₦59.6m"],
             ["Owner", "Commercial Director; pilot in 3 stores"]],
      so="A 20% cap adds ₦59.6m–₦71.5m even if half the re-priced volume is lost. It's a static model, so pilot first.",
      tryit=SOL + "'5 What-If'",
      split=4.90),

    O("P.3.10",
      what="Before handover, prove the headline numbers: sales and profit reconcile to the ledger, row counts are "
           "documented, slicers are tested and the file opens on another machine.",
      like="Like a final walk-through before handing over a new flat: lights, taps, doors, keys.",
      how=["Reconcile sales and profit.", "Document the row counts.", "Test slicers and a second machine."],
      tip="Record the evidence beside each check, not just a tick.",
      trap="Handing over a polished dashboard with no control totals.",
      case="Capstone 1 verification.",
      steps=["Run each check.", "Record the evidence.", "Mark the status."],
      table=[["Check", "Evidence", "Status"], ["Revenue = POS ledger", "₦733,215,255", "Pass"],
             ["Profit = ledger", "₦93,439,242", "Pass"], ["Row counts documented", "3,312 FY2025; 9,994 total", "Pass"],
             ["Slicers tested", "3 slicers × 6 visuals", "Pass"], ["Second machine", "Charts and ₦ symbol intact", "Pass"]],
      so="A reviewer can reproduce every headline number from the evidence.",
      split=4.70),
]
