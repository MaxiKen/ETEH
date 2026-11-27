"""Week 3 workbook, module 3.6 in two parts (edition 3): Power Pivot and DAX.

Each topic shows the slide's DAX measure, then a worksheet 'twin' that reproduces the measure's result with SUMIFS / COUNTIFS on
tLines or tTransactions (3.6.39: a measure is tested by comparing it with an independent calculation). The learner builds the
measure in Power Pivot on a copy of this workbook and types the value the PivotTable shows; the check compares it with the twin.
Small helper tables tCat (3.6A) and tYear (3.6B) hold one row per category / year with calculated twin columns.
"""
from __future__ import annotations

import pandas as pd

from w1x import ModuleSheet
from wb3_common import N, f, net, netn, tp

XL0 = pd.Timestamp("1899-12-30")
CATS = ["Electronics", "Home", "Fashion", "Automotive", "Food", "Health", "Beauty"]
MODEL = ("Data Model (your copy)", "Power Pivot › Manage", "tTransactions, tProducts, tCustomers, tSellers and tDimDate, related as in 3.5.2.")
TWIN = ("tLines / tTransactions", "Sheets Data_Lines, Data_Transactions (A1)", "Source of every worksheet twin on this sheet.")


def netd(d0, d1, t="tLines", extra=""):
    """Net revenue between two dates (inclusive), as SUMIFS on order_date."""
    a, b_ = pd.Timestamp(d0), pd.Timestamp(d1)
    rng = (f',{t}[order_date],">="&DATE({a.year},{a.month},{a.day}),{t}[order_date],"<="&DATE({b_.year},{b_.month},{b_.day}){extra}')
    return f'SUMIFS({t}[total_price],{t}[order_status],"Delivered"{rng})+SUMIFS({t}[total_price],{t}[order_status],"Completed"{rng})'


def per_seller(t="tLines"):
    return (f'SUMIFS({t}[total_price],{t}[seller_id],tSellers[seller_id],{t}[order_status],"Delivered")+'
            f'SUMIFS({t}[total_price],{t}[seller_id],tSellers[seller_id],{t}[order_status],"Completed")')


def build(b, D):
    m36a(b, D)
    m36b(b, D)


# =====================================================================================================
def m36a(b, D):
    TX, L, NT, C, P = D.TX, D.L, D.NT, D.C, D.P
    m = ModuleSheet(b, "3.6A DAX Foundations", "3.6A  Power Pivot and DAX: Foundations and Context",
                    "Tutor walk-through after Module 3.6A: build each measure in Power Pivot, then check it against the worksheet twin",
                    "3.6.1 - 3.6.21")
    m.dataset([MODEL, TWIN, ("tCat", "On this sheet (3.6.7)", "One row per category with twin columns: a PivotTable in formulas.")])
    NETV, NETN, GROSS = f(NT.total_price.sum()), len(NT), f(TX.total_price.sum())
    cat = NT.groupby("category").total_price.agg(["size", "sum"])
    tp(m, "3.6.1",
       [("Fact rows (tTransactions)", "=ROWS(tTransactions)", len(TX), "int", ""),
        ("Dimension rows (customers + products + sellers)", "=ROWS(tCustomers)+ROWS(tProducts)+ROWS(tSellers)", len(C) + len(P) + len(D.SE), "int", ""),
        ("Calendar rows (tDimDate)", "=ROWS(tDimDate)", len(D.DD), "int", "")],
       [("n", "How many rows will the Data Model hold in all, for these five tables?",
         "=ROWS(tTransactions)+ROWS(tCustomers)+ROWS(tProducts)+ROWS(tSellers)+ROWS(tDimDate)",
         len(TX) + len(C) + len(P) + len(D.SE) + len(D.DD), "int", "Rows in the model")],
       steps=["Save a copy of this workbook as Week3_Model.xlsx: build the model there.", "Power Pivot › Manage opens the model window.",
              "Each table is a tab; Diagram View shows the relationships."])
    tp(m, "3.6.2",
       [("Tables you will add", "=5", 5, "int", "tTransactions, tProducts, tCustomers, tSellers, tDimDate.")],
       [("t", "Which dialog turns the Power Pivot tab on? (File › Options › Add-ins › Manage: …)", "COM Add-ins", None, None, ("com add-ins", "com"))],
       steps=["File › Options › Add-ins.", "Manage: COM Add-ins › Go.", "Tick Microsoft Power Pivot for Excel › OK: the Power Pivot tab appears."],
       do=[("Enable the Power Pivot add-in.", "A Power Pivot tab on the ribbon with Manage, Measures, KPIs.")])
    tp(m, "3.6.3",
       [("tTransactions rows to add", "=ROWS(tTransactions)", len(TX), "int", "Power Pivot › Add to Data Model."),
        ("tDimDate rows to add", "=ROWS(tDimDate)", len(D.DD), "int", "")],
       [("n", "After adding tCustomers, how many rows does its tab in Power Pivot show?", "=ROWS(tCustomers)", len(C), "int", "tCustomers rows")],
       steps=["Click inside tTransactions › Power Pivot › Add to Data Model.", "Repeat for tProducts, tCustomers, tSellers, tDimDate.",
              "In Diagram View drag product_id, customer_id, seller_id and order_date › Date to relate them."],
       do=[("Add the five tables and draw the four relationships.", "Diagram View: four 1-to-* lines into tTransactions.")])
    tp(m, "3.6.4",
       [("Is Net = Yes (calculated column, 11,000 values stored)", "=" + netn(t="tTransactions"), NETN, "int", ""),
        ("Is Net = No", "=ROWS(tTransactions)-(" + netn(t="tTransactions") + ")", len(TX) - NETN, "int", ""),
        ("[Net Revenue] (measure, nothing stored)", "=" + net(t="tTransactions"), NETV, N, "")],
       [("n", "Add the calculated column Is Net = IF(tTransactions[order_status] IN {\"Delivered\",\"Completed\"}, \"Yes\", \"No\"). How many 'No'?",
         "=ROWS(tTransactions)-(" + netn(t="tTransactions") + ")", len(TX) - NETN, "int", "Is Net = No"),
        ("t", "A total that must respond to slicers: calculated column or measure?", "measure", None, None, ("a measure",))],
       steps=["Calculated column: Power Pivot › click the first empty column › type =… (row by row, stored).",
              "Measure: Power Pivot › Measures › New Measure (calculated live in the PivotTable).", "Rows and slicers use columns; Values use measures."])
    tp(m, "3.6.5",
       [("[Gross Bookings]", "=SUM(tTransactions[total_price])", GROSS, N, "SUM ( tTransactions[total_price] )"),
        ("[Orders]", "=ROWS(tTransactions)", len(TX), "int", "COUNTROWS ( tTransactions )")],
       [("n", "Create Commission := SUM ( tTransactions[commission_amount] ). What does it show with no filters?", "=SUM(tTransactions[commission_amount])",
         f(TX.commission_amount.sum()), N, "[Commission]")],
       steps=["Measure name, then :=, then the expression.", "Columns as Table[Column]; measures as [Measure].",
              "Spaces and line breaks are allowed: lay long measures out on several lines."], code=True)
    big = TX.loc[TX.total_price.idxmax()]
    tp(m, "3.6.6",
       [("TRX004782: quantity × unit_price (row context)",
         '=XLOOKUP("TRX004782",tTransactions[transaction_id],tTransactions[quantity])*XLOOKUP("TRX004782",tTransactions[transaction_id],tTransactions[unit_price])',
         f(TX.set_index("transaction_id").loc["TRX004782"].pipe(lambda r: r.quantity * r.unit_price)), N, "One row: a calculated column."),
        ("SUMX over every row", "=SUMPRODUCT(tTransactions[quantity],tTransactions[unit_price])", f((TX.quantity * TX.unit_price).sum()), N,
         "An iterator creates row context in a measure.")],
       [("n", "What does the calculated column = [quantity] * [unit_price] give on TRX000001?",
         '=XLOOKUP("TRX000001",tTransactions[transaction_id],tTransactions[quantity])*XLOOKUP("TRX000001",tTransactions[transaction_id],tTransactions[unit_price])',
         f(TX.iloc[0].quantity * TX.iloc[0].unit_price), N, "Row context on TRX000001")],
       steps=["In a calculated column, [quantity] means 'this row's quantity'.", "A measure has no current row: wrap the expression in SUMX.",
              f"The largest line ({big.transaction_id}) shows the row value."])

    def cat_table(mm):
        df = pd.DataFrame({"category": CATS})
        nl = [int(cat.loc[c, "size"]) for c in CATS]
        nv = [f(cat.loc[c, "sum"]) for c in CATS]
        calc = {"Net_Lines": ('COUNTIFS(tLines[category],[@category],tLines[order_status],"Delivered")+'
                              'COUNTIFS(tLines[category],[@category],tLines[order_status],"Completed")', "int", nl),
                "Net_Revenue": ('SUMIFS(tLines[total_price],tLines[category],[@category],tLines[order_status],"Delivered")+'
                                'SUMIFS(tLines[total_price],tLines[category],[@category],tLines[order_status],"Completed")', N, nv),
                "AOV": ("[@Net_Revenue]/[@Net_Lines]", N, [v / n for v, n in zip(nv, nl)]),
                "Cat_Share": ("[@Net_Revenue]/SUM([Net_Revenue])", "pct", [v / NETV for v in nv])}
        mm.h("tCat: each row is one filter context (category), worked out with SUMIFS")
        mm.s.table(df, "tCat", calc=calc, widths={"category": 14, "Net_Lines": 11, "Net_Revenue": 15, "AOV": 12, "Cat_Share": 10})
        mm.s.r += 1

    tp(m, "3.6.7",
       [("Electronics: [Net Revenue]", '=XLOOKUP("Electronics",tCat[category],tCat[Net_Revenue])', f(cat.loc["Electronics", "sum"]), N,
         "The category on rows is the filter."),
        ("Grand total (no filter)", "=SUM(tCat[Net_Revenue])", NETV, N, "Same as the measure with no filter.")],
       [("n", "Category on rows, [Net Revenue] in values: what does Beauty show?", '=XLOOKUP("Beauty",tCat[category],tCat[Net_Revenue])',
         f(cat.loc["Beauty", "sum"]), N, "Beauty [Net Revenue]"),
        ("n", "How many net lines are behind Fashion's cell?", '=XLOOKUP("Fashion",tCat[category],tCat[Net_Lines])', int(cat.loc["Fashion", "size"]), "int",
         "Fashion net lines")],
       steps=["Put tProducts[category] on rows and [Net Revenue] in values.", "Each cell is the measure under one filter: that category.",
              "The filter flows from tProducts to tTransactions along the relationship."], extra=cat_table)
    cr = TX[~TX.order_status.isin(["Delivered", "Completed"])]
    tp(m, "3.6.8",
       [("[Gross Bookings]", "=SUM(tTransactions[total_price])", GROSS, N, ""),
        ("[Net Revenue] (CALCULATE with a status filter)", "=" + net(t="tTransactions"), NETV, N, ""),
        ("Cancelled + Returned", "=SUM(tTransactions[total_price])-(" + net(t="tTransactions") + ")", f(cr.total_price.sum()), N, "")],
       [("n", "Write Returned Value := CALCULATE ( [Gross Bookings], tTransactions[order_status] = \"Returned\" ). Its value?",
         '=SUMIF(tTransactions[order_status],"Returned",tTransactions[total_price])', f(TX[TX.order_status == "Returned"].total_price.sum()), N,
         "[Returned Value]"),
        ("n", "CALCULATE ( [Net Revenue], tTransactions[payment_method] = \"Paystack\" )?", "=" + net(t="tTransactions", payment_method="Paystack"),
         f(NT[NT.payment_method == "Paystack"].total_price.sum()), N, "Net revenue, Paystack")],
       steps=["CALCULATE ( expression, filter1, filter2 … ) evaluates the expression under changed filters.",
              "A column filter replaces any existing filter on that column.", "Every other filter (slicers, rows) still applies."], code=True)
    bg = NT[NT.total_price > 500000]
    tp(m, "3.6.9",
       [("[Net Rev Big Orders]: lines", '=COUNTIFS(tLines[order_status],"Delivered",tLines[total_price],">500000")+'
                                         'COUNTIFS(tLines[order_status],"Completed",tLines[total_price],">500000")', len(bg), "int", ""),
        ("[Net Rev Big Orders]: value", "=" + net(total_price=">500000"), f(bg.total_price.sum()), N, "")],
       [("n", "Change the threshold to 1,000,000. What is the value?", "=" + net(total_price=">1000000"), f(NT[NT.total_price > 1e6].total_price.sum()), N,
         "Net revenue, lines over ₦1m"),
        ("n", "What share of net revenue do lines over ₦500,000 make? (as %)", "=(" + net(total_price=">500000") + ")/(" + net() + ")",
         f(bg.total_price.sum() / NT.total_price.sum()), "pct", "Share of big lines")],
       steps=["FILTER ( table, condition ) returns the rows where the condition is true.", "Use it inside CALCULATE for conditions on a measure or several columns.",
              "For a simple column test, a plain CALCULATE filter is faster (3.6.34)."], code=True)
    tp(m, "3.6.10",
       [("Electronics [Cat Share]", '=XLOOKUP("Electronics",tCat[category],tCat[Cat_Share])', f(cat.loc["Electronics", "sum"] / NETV), "pct",
         "Net revenue ÷ net revenue with ALL categories."),
        ("Denominator (ALL removes the category filter)", "=SUM(tCat[Net_Revenue])", NETV, N, "")],
       [("n", "What is Home's [Cat Share]?", '=XLOOKUP("Home",tCat[category],tCat[Cat_Share])', f(cat.loc["Home", "sum"] / NETV), "pct", "Home share"),
        ("n", "What is Food's [Cat Share]?", '=XLOOKUP("Food",tCat[category],tCat[Cat_Share])', f(cat.loc["Food", "sum"] / NETV), "pct", "Food share")],
       steps=["ALL ( column ) removes filters from that column.", "CALCULATE ( [M], ALL ( … ) ) gives the denominator for a share.",
              "DIVIDE the measure by it; format as %."], code=True)
    sw = NT[NT.customer_zone == "South-West"]
    tier = sw.groupby("membership_tier").total_price.sum()
    tp(m, "3.6.11",
       [("South-West net revenue", "=" + net(customer_zone="South-West"), f(sw.total_price.sum()), N, "Denominator: zone kept, tier removed."),
        ("Gold share of South-West", "=(" + net(customer_zone="South-West", membership_tier="Gold") + ")/(" + net(customer_zone="South-West") + ")",
         f(tier["Gold"] / sw.total_price.sum()), "pct", "")],
       [("n", "What is Platinum's [Share of Zone] in the South-West?",
         "=(" + net(customer_zone="South-West", membership_tier="Platinum") + ")/(" + net(customer_zone="South-West") + ")",
         f(tier["Platinum"] / sw.total_price.sum()), "pct", "Platinum share of SW"),
        ("n", "And Gold's share in the North-East?", "=(" + net(customer_zone="North-East", membership_tier="Gold") + ")/(" + net(customer_zone="North-East") + ")",
         f(NT[(NT.customer_zone == "North-East") & (NT.membership_tier == "Gold")].total_price.sum() / NT[NT.customer_zone == "North-East"].total_price.sum()),
         "pct", "Gold share of NE")],
       steps=["ALLEXCEPT ( table, column ) removes every filter on the table except that column.", "Here: keep zone, remove tier.",
              "Zone and tier on rows: each tier's share of its own zone."], code=True)
    sc = NT.groupby("seller_id").category.nunique()
    tp(m, "3.6.12",
       [("Categories sold by SELL0200", '=SUMPRODUCT(--((COUNTIFS(tLines[seller_id],"SELL0200",tLines[category],tCat[category],tLines[order_status],"Delivered")+'
                                         'COUNTIFS(tLines[seller_id],"SELL0200",tLines[category],tCat[category],tLines[order_status],"Completed"))>0))',
         int(sc["SELL0200"]), "int", "COUNTROWS ( VALUES ( category ) ) for that seller."),
        ("Sellers with net sales", "=SUMPRODUCT(--((" + per_seller() + ")>0))", len(sc), "int", "")],
       [("n", "How many categories has SELL0098 sold (net lines)?",
         '=SUMPRODUCT(--((COUNTIFS(tLines[seller_id],"SELL0098",tLines[category],tCat[category],tLines[order_status],"Delivered")+'
         'COUNTIFS(tLines[seller_id],"SELL0098",tLines[category],tCat[category],tLines[order_status],"Completed"))>0))', int(sc["SELL0098"]), "int",
         "Categories of SELL0098"),
        ("n", "How many distinct categories exist in tProducts?", "=COUNTA(UNIQUE(tProducts[category]))", int(P.category.nunique()), "int", "Categories")],
       steps=["VALUES ( column ) returns the distinct values visible in the current filter.", "COUNTROWS ( VALUES ( … ) ) counts them.",
              "DISTINCT is the same but ignores the blank row added for unmatched keys."], code=True)
    lc = L.groupby("category").size()
    tp(m, "3.6.13",
       [("Health lines (all statuses)", '=COUNTIF(tLines[category],"Health")', int(lc["Health"]), "int", "RELATED brings category onto each line."),
        ("Health share of lines", '=COUNTIF(tLines[category],"Health")/ROWS(tLines)', f(lc["Health"] / len(L)), "pct", "")],
       [("t", "What category does RELATED ( tProducts[category] ) return on TRX000002?", L.set_index("transaction_id").category["TRX000002"],
         '=XLOOKUP("TRX000002",tLines[transaction_id],tLines[category])', "Category on TRX000002"),
        ("n", "How many Automotive lines are there?", '=COUNTIF(tLines[category],"Automotive")', int(lc["Automotive"]), "int", "Automotive lines")],
       steps=["RELATED works on the many side, looking up the one side.", "Add it as a calculated column in tTransactions: = RELATED ( tProducts[category] ).",
              "It needs an active relationship."], code=True)
    per_c = TX.groupby("customer_id").size()
    tp(m, "3.6.14",
       [("Orders Placed for CUST00985", '=COUNTIF(tTransactions[customer_id],"CUST00985")', int(per_c["CUST00985"]), "int", ""),
        ("Average per customer (all 2,500)", "=ROWS(tTransactions)/ROWS(tCustomers)", len(TX) / len(C), "dec1", ""),
        ("Customers with no orders (blank)", "=SUMPRODUCT(--(COUNTIF(tTransactions[customer_id],tCustomers[customer_id])=0))",
         int((~C.customer_id.isin(TX.customer_id)).sum()), "int", "")],
       [("n", "Orders Placed for CUST00001?", '=COUNTIF(tTransactions[customer_id],"CUST00001")', int(per_c.get("CUST00001", 0)), "int", "CUST00001 orders"),
        ("n", "How many customers placed 10 or more order lines?", "=SUMPRODUCT(--(COUNTIF(tTransactions[customer_id],tCustomers[customer_id])>=10))",
         int((per_c >= 10).sum()), "int", "Customers with 10+ lines")],
       steps=["RELATEDTABLE works on the one side, returning the related many rows.", "In tCustomers add: = COUNTROWS ( RELATEDTABLE ( tTransactions ) ).",
              "Customers with no rows show blank."], code=True)
    tp(m, "3.6.15",
       [("SUMX ( qty × unit_price )", "=SUMPRODUCT(tTransactions[quantity],tTransactions[unit_price])", f((TX.quantity * TX.unit_price).sum()), N, ""),
        ("SUM ( total_price )", "=SUM(tTransactions[total_price])", GROSS, N, "Same: a check."),
        ("SUM ( unit_price )", "=SUM(tTransactions[unit_price])", f(TX.unit_price.sum()), N, "Meaningless: prices are not additive.")],
       [("n", "Net revenue with SUMX on Delivered + Completed rows only (quantity × unit_price)?",
         '=SUMPRODUCT(tTransactions[quantity],tTransactions[unit_price],--(tTransactions[order_status]="Delivered"))+'
         'SUMPRODUCT(tTransactions[quantity],tTransactions[unit_price],--(tTransactions[order_status]="Completed"))',
         f((NT.quantity * NT.unit_price).sum()), N, "SUMX on net rows")],
       steps=["SUMX ( table, expression ) works the expression out row by row, then adds.", "Use it when the row maths must happen first (quantity × price).",
              "Compare with a stored total to check."], code=True)
    nb = NT.customer_id.nunique()
    tp(m, "3.6.16",
       [("Average quantity per line", "=AVERAGE(tTransactions[quantity])", f(TX.quantity.mean()), "dec2", "Grain: line."),
        ("Average net order value", "=(" + net() + ")/(" + netn() + ")", NETV / NETN, N, "Grain: net line."),
        ("Net buyers", '=SUMPRODUCT(--((COUNTIFS(tTransactions[customer_id],tCustomers[customer_id],tTransactions[order_status],"Delivered")+'
                       'COUNTIFS(tTransactions[customer_id],tCustomers[customer_id],tTransactions[order_status],"Completed"))>0))', nb, "int", ""),
        ("Avg net spend per customer", "=(" + net() + ")/" + str(nb), NETV / nb, N, "Grain: customer (AVERAGEX over VALUES).")],
       [("n", "What is the average unit_price per line?", "=AVERAGE(tTransactions[unit_price])", f(TX.unit_price.mean()), N, "Average unit price"),
        ("o", "Why is the average spend per customer much bigger than the average order value?",
         "Different grain: a customer has several lines on average, so their total is several order values added together.")],
       steps=["AVERAGEX ( table, expression ) averages the expression over the table's rows.", "The table sets the grain: VALUES ( customer_id ) = per customer.",
              "Always say the grain next to an average."], code=True)
    tp(m, "3.6.17",
       [("COUNTROWS ( tTransactions )", "=ROWS(tTransactions)", len(TX), "int", "Every line."),
        ("COUNT ( delivery_id )", "=COUNTA(tTransactions[delivery_id])", int(TX.delivery_id.notna().sum()), "int", "Skips blanks.")],
       [("n", "COUNTROWS ( tDeliveries )?", "=ROWS(tDeliveries)", len(D.DV), "int", "Rows in tDeliveries"),
        ("n", "CALCULATE ( COUNTROWS ( tTransactions ), tTransactions[order_status] = \"Cancelled\" )?", '=COUNTIF(tTransactions[order_status],"Cancelled")',
         int((TX.order_status == "Cancelled").sum()), "int", "Cancelled rows")],
       steps=["COUNTROWS counts rows of a table: the safest count.", "COUNT / COUNTA count non-blank values in one column.",
              "Pick the one that matches the question."])
    tp(m, "3.6.18",
       [("Order lines (COUNTROWS)", "=ROWS(tTransactions)", len(TX), "int", ""),
        ("Customers (DISTINCTCOUNT)", "=COUNTA(UNIQUE(tTransactions[customer_id]))", int(TX.customer_id.nunique()), "int", "Unique buyers.")],
       [("n", "DISTINCTCOUNT ( tTransactions[seller_id] )?", "=COUNTA(UNIQUE(tTransactions[seller_id]))", int(TX.seller_id.nunique()), "int", "Distinct sellers"),
        ("n", "DISTINCTCOUNT ( tTransactions[product_id] )?", "=COUNTA(UNIQUE(tTransactions[product_id]))", int(TX.product_id.nunique()), "int",
         "Distinct products")],
       steps=["DISTINCTCOUNT ( column ) counts unique values in the current filter.", "Unique counts are not additive: categories' headcounts do not add up to the total.",
              "Show the total from the measure, not a sum of rows."])
    ss = NT.groupby("seller_id").total_price.agg(["size", "sum"])
    tp(m, "3.6.19",
       [("[AOV], all sellers", "=(" + net() + ")/(" + netn() + ")", NETV / NETN, N, ""),
        ("SELL0046 (no orders): DIVIDE gives blank", '=IFERROR((' + net(seller_id="SELL0046") + ')/(' + netn(seller_id="SELL0046") + '),"")', "", "gen",
         "In Excel, / gives #DIV/0!; DIVIDE returns BLANK.")],
       [("n", "What is SELL0200's [AOV]?", "=(" + net(seller_id="SELL0200") + ")/(" + netn(seller_id="SELL0200") + ")",
         f(ss.loc["SELL0200", "sum"] / ss.loc["SELL0200", "size"]), N, "AOV SELL0200"),
        ("n", "What is the Electronics [AOV]?", '=XLOOKUP("Electronics",tCat[category],tCat[AOV])', f(cat.loc["Electronics", "sum"] / cat.loc["Electronics", "size"]),
         N, "AOV Electronics")],
       steps=["DIVIDE ( numerator, denominator [, alternate] ).", "It returns BLANK (or the alternate) when the denominator is 0 or blank.",
              "Use it for every ratio measure."], code=True)

    def size(aov):
        return "Large" if aov >= 200000 else "Medium" if aov >= 50000 else "Small"

    aov = {c: cat.loc[c, "sum"] / cat.loc[c, "size"] for c in CATS}
    tp(m, "3.6.20",
       [("Electronics [Order Size]", '=IFS(XLOOKUP("Electronics",tCat[category],tCat[AOV])>=200000,"Large",XLOOKUP("Electronics",tCat[category],tCat[AOV])>=50000,"Medium",TRUE,"Small")',
         size(aov["Electronics"]), "gen", ""),
        ("Fashion [Order Size]", '=IFS(XLOOKUP("Fashion",tCat[category],tCat[AOV])>=200000,"Large",XLOOKUP("Fashion",tCat[category],tCat[AOV])>=50000,"Medium",TRUE,"Small")',
         size(aov["Fashion"]), "gen", "")],
       [("t", "What [Order Size] does Beauty get?", size(aov["Beauty"]),
         '=IFS(XLOOKUP("Beauty",tCat[category],tCat[AOV])>=200000,"Large",XLOOKUP("Beauty",tCat[category],tCat[AOV])>=50000,"Medium",TRUE,"Small")', "Beauty size"),
        ("t", "And Health?", size(aov["Health"]),
         '=IFS(XLOOKUP("Health",tCat[category],tCat[AOV])>=200000,"Large",XLOOKUP("Health",tCat[category],tCat[AOV])>=50000,"Medium",TRUE,"Small")', "Health size")],
       steps=["IF ( test, then, else ) for one test.", "SWITCH ( TRUE (), test1, r1, test2, r2, else ) for several, checked in order.",
              "Excel's twin is IFS."], code=True)
    tp(m, "3.6.21",
       [("rev", "=SUM(tTransactions[total_price])", GROSS, N, "VAR rev = [Gross Bookings]"),
        ("comm", "=SUM(tTransactions[commission_amount])", f(TX.commission_amount.sum()), "naira2", ""),
        ("[Comm %]", "=SUM(tTransactions[commission_amount])/SUM(tTransactions[total_price])", f(TX.commission_amount.sum() / TX.total_price.sum()), "pct", "")],
       [("n", "[Comm %] for Electronics lines only (all statuses)?",
         '=SUMIF(tLines[category],"Electronics",tLines[commission_amount])/SUMIF(tLines[category],"Electronics",tLines[total_price])',
         f(L[L.category == "Electronics"].commission_amount.sum() / L[L.category == "Electronics"].total_price.sum()), "pct", "Comm % Electronics")],
       steps=["VAR name = expression stores a value once.", "RETURN gives the result using the variables.",
              "Easier to read and to debug: RETURN one variable at a time."], code=True)
    m.finish()


# =====================================================================================================
def m36b(b, D):
    TX, L, NT, DS, DSS, A, CEN, CV = D.TX, D.L, D.NT, D.DS, D.DSS, D.A, D.CEN, D.CV
    m = ModuleSheet(b, "3.6B DAX Time & KPIs", "3.6B  Power Pivot and DAX: Time Intelligence, KPIs and Performance",
                    "Tutor walk-through after Module 3.6B: time-intelligence measures on tDimDate, ranking, blanks and errors, then KPI sets "
                    "for retail, hospital and university", "3.6.22 - 3.6.41")
    m.dataset([MODEL, TWIN, ("tYear", "On this sheet (3.6.23)", "Net revenue per year with last year and growth."),
               ("tDash / tDashStores / tAdm23 / tCensus / tCV", "Data_ sheets (A1)", "The KPI examples in 3.6.35 - 3.6.37.")])
    NETV = f(NT.total_price.sum())
    nt25 = NT[NT.order_year == 2025]
    mon = nt25.groupby(nt25.order_date.dt.month).total_price.sum()
    ytd = mon.cumsum()
    tp(m, "3.6.22",
       [("[Rev YTD], March 2025", "=" + netd("2025-01-01", "2025-03-31"), f(ytd[3]), N, "TOTALYTD: 1 January to the end of the month shown."),
        ("[Rev YTD], December 2025", "=" + netd("2025-01-01", "2025-12-31"), f(ytd[12]), N, "= the year's net revenue.")],
       [("n", "What does [Rev YTD] show for June 2025?", "=" + netd("2025-01-01", "2025-06-30"), f(ytd[6]), N, "YTD June 2025"),
        ("n", "And for September 2025?", "=" + netd("2025-01-01", "2025-09-30"), f(ytd[9]), N, "YTD September 2025")],
       steps=["Mark tDimDate as the date table; relate Date to order_date.", "Rev YTD := TOTALYTD ( [Net Revenue], tDimDate[Date] ).",
              "Year and Month from tDimDate on rows."], code=True)
    yr = NT.groupby("order_year").total_price.sum()

    def year_table(mm):
        df = pd.DataFrame({"Year": list(range(2020, 2026))})
        nv = [f(yr[y]) for y in df.Year]
        ly = [None] + nv[:-1]
        calc = {"Net_Revenue": ('SUMIFS(tLines[total_price],tLines[order_year],[@Year],tLines[order_status],"Delivered")+'
                                'SUMIFS(tLines[total_price],tLines[order_year],[@Year],tLines[order_status],"Completed")', N, nv),
                "Rev_LY": ('IF([@Year]=2020,"",SUMIFS(tLines[total_price],tLines[order_year],[@Year]-1,tLines[order_status],"Delivered")+'
                           'SUMIFS(tLines[total_price],tLines[order_year],[@Year]-1,tLines[order_status],"Completed"))', N, [x if x else "" for x in ly]),
                "YoY": ('IF([@Year]=2020,"",[@Net_Revenue]/[@Rev_LY]-1)', "pct", [""] + [nv[i] / nv[i - 1] - 1 for i in range(1, 6)])}
        mm.h("tYear: one row per year (the PivotTable twin for 3.6.23 and 3.6.26)")
        mm.s.table(df, "tYear", fmts={"Year": "gen"}, calc=calc, widths={"Year": 8, "Net_Revenue": 15, "Rev_LY": 15, "YoY": 9})
        mm.s.r += 1

    tp(m, "3.6.23",
       [("2025 [Rev LY]", "=XLOOKUP(2025,tYear[Year],tYear[Rev_LY])", f(yr[2024]), N, "SAMEPERIODLASTYEAR shifts the dates back a year."),
        ("2025 [YoY %]", "=XLOOKUP(2025,tYear[Year],tYear[YoY])", f(yr[2025] / yr[2024] - 1), "pct", "")],
       [("n", "What is the [YoY %] for 2022?", "=XLOOKUP(2022,tYear[Year],tYear[YoY])", f(yr[2022] / yr[2021] - 1), "pct", "YoY 2022"),
        ("n", "What does [Rev LY] show on the 2021 row?", "=XLOOKUP(2021,tYear[Year],tYear[Rev_LY])", f(yr[2020]), N, "Rev LY 2021")],
       steps=["Rev LY := CALCULATE ( [Net Revenue], SAMEPERIODLASTYEAR ( tDimDate[Date] ) ).", "YoY % := DIVIDE ( [Net Revenue] - [Rev LY], [Rev LY] ).",
              "2020 has no prior year: blank is correct."], code=True, extra=year_table)
    dec24 = f(NT[(NT.order_month == "2024-12")].total_price.sum())
    tp(m, "3.6.24",
       [("January 2025 [Net Revenue]", "=" + net(order_month="2025-01"), f(mon[1]), N, ""),
        ("January 2025 [Rev PM] (December 2024)", "=" + net(order_month="2024-12"), dec24, N, "DATEADD ( …, -1, MONTH )"),
        ("[MoM %]", "=(" + net(order_month="2025-01") + ")/(" + net(order_month="2024-12") + ")-1", f(mon[1] / dec24 - 1), "pct", "")],
       [("n", "What is the [MoM %] for February 2025?", "=(" + net(order_month="2025-02") + ")/(" + net(order_month="2025-01") + ")-1",
         f(mon[2] / mon[1] - 1), "pct", "MoM Feb 2025"),
        ("n", "What is [Rev PM] on the December 2025 row?", "=" + net(order_month="2025-11"), f(mon[11]), N, "Rev PM Dec 2025")],
       steps=["DATEADD ( dates, -1, MONTH ) shifts the visible dates back one month.", "Use it inside CALCULATE.",
              "Works across year boundaries (January compares with December)."], code=True)
    tp(m, "3.6.25",
       [("[Rev R3M], June 2025 (Apr–Jun)", "=" + netd("2025-04-01", "2025-06-30"), f(mon[4] + mon[5] + mon[6]), N, ""),
        ("[Rev R3M], December 2025 (Oct–Dec)", "=" + netd("2025-10-01", "2025-12-31"), f(mon[10] + mon[11] + mon[12]), N, "")],
       [("n", "What does [Rev R3M] show for March 2025 (Jan–Mar)?", "=" + netd("2025-01-01", "2025-03-31"), f(mon[1] + mon[2] + mon[3]), N, "R3M March 2025"),
        ("n", "And for September 2025?", "=" + netd("2025-07-01", "2025-09-30"), f(mon[7] + mon[8] + mon[9]), N, "R3M September 2025")],
       steps=["DATESINPERIOD ( dates, last date, -3, MONTH ) = the 3 months ending on the last visible date.",
              "DATESYTD ( dates ) = 1 January to the last visible date.", "Use either as a CALCULATE filter."], code=True)
    run = yr.cumsum()
    tp(m, "3.6.26",
       [("[Rev Running], 2022", "=SUMIFS(tYear[Net_Revenue],tYear[Year],\"<=2022\")", f(run[2022]), N, "All dates up to the end of 2022."),
        ("[Rev Running], 2025", "=SUM(tYear[Net_Revenue])", f(run[2025]), N, "= all-time net revenue.")],
       [("n", "What is [Rev Running] at the end of 2024?", "=SUMIFS(tYear[Net_Revenue],tYear[Year],\"<=2024\")", f(run[2024]), N, "Running 2024"),
        ("n", "[Rev Running] on 30 June 2023?", "=" + netd("2020-01-01", "2023-06-30"), f(NT[NT.order_date <= "2023-06-30"].total_price.sum()), N,
         "Running to 30 Jun 2023")],
       steps=["FILTER ( ALL ( tDimDate ), tDimDate[Date] <= MAX ( tDimDate[Date] ) ).", "ALL removes the year/month filter; the condition keeps dates up to the visible end.",
              "Wrap in CALCULATE ( [Net Revenue], … )."], code=True)
    zn = NT.groupby("customer_zone").total_price.sum()
    tp(m, "3.6.27",
       [("South-West [Zone Share]", "=(" + net(customer_zone="South-West") + ")/(" + net() + ")", f(zn["South-West"] / NETV), "pct", ""),
        ("All 6 zones", "=(" + net() + ")/(" + net() + ")", 1.0, "pct", "")],
       [("n", "What is the North-West [Zone Share]?", "=(" + net(customer_zone="North-West") + ")/(" + net() + ")", f(zn["North-West"] / NETV), "pct",
         "North-West share"),
        ("n", "And the North-East?", "=(" + net(customer_zone="North-East") + ")/(" + net() + ")", f(zn["North-East"] / NETV), "pct", "North-East share")],
       steps=["Share := DIVIDE ( [M], CALCULATE ( [M], ALL ( zone ) ) ).", "Format as 0.0%.",
              "Or: PivotTable › Show Values As › % of Grand Total (a quick check)."], code=True)
    ss = NT.groupby("seller_id").total_price.sum().sort_values(ascending=False)
    tp(m, "3.6.28",
       [("Rank 1 seller's net revenue", "=MAX(" + per_seller() + ")", f(ss.iloc[0]), N, ""),
        ("Rank 2", "=LARGE(" + per_seller() + ",2)", f(ss.iloc[1]), N, ""),
        ("Rank of SELL0057", "=SUMPRODUCT(--((" + per_seller() + ")>(" + net(seller_id="SELL0057") + ")))+1", int((ss > ss["SELL0057"]).sum() + 1), "int", "")],
       [("n", "What is SELL0170's [Seller Rank]?", "=SUMPRODUCT(--((" + per_seller() + ")>(" + net(seller_id="SELL0170") + ")))+1",
         int((ss > ss["SELL0170"]).sum() + 1), "int", "Rank SELL0170"),
        ("n", "What net revenue does the 10th seller have?", "=LARGE(" + per_seller() + ",10)", f(ss.iloc[9]), N, "10th seller")],
       steps=["RANKX ( ALL ( tSellers[seller_id] ), [Net Revenue], , DESC, DENSE ).", "ALL makes every seller compete, whatever is on rows.",
              "Rank by seller_id: 86 store names repeat."], code=True)
    tp(m, "3.6.29",
       [("Sellers with net sales (rows [Net Revenue] shows)", "=SUMPRODUCT(--((" + per_seller() + ")>0))", len(ss), "int", ""),
        ("Rows [Rev 0] shows", "=ROWS(tSellers)", len(D.SE), "int", "+ 0 turns BLANK into 0: every seller appears.")],
       [("n", "How many sellers does [Rev 0] add that [Net Revenue] hides?", "=ROWS(tSellers)-SUMPRODUCT(--((" + per_seller() + ")>0))", len(D.SE) - len(ss),
         "int", "Sellers shown only by Rev 0")],
       steps=["BLANK rows are hidden in PivotTables: usually what you want.", "Add + 0 only when the zero rows must show.",
              "Never use + 0 on a measure used in a ratio's denominator."], code=True)
    rr = L.groupby("category").order_status.apply(lambda s: (s == "Returned").mean())
    tp(m, "3.6.30",
       [("[Return Rate], all lines", '=COUNTIF(tTransactions[order_status],"Returned")/ROWS(tTransactions)', f((TX.order_status == "Returned").mean()), "pct", ""),
        ("Health (highest)", '=COUNTIFS(tLines[category],"Health",tLines[order_status],"Returned")/COUNTIF(tLines[category],"Health")', f(rr["Health"]), "pct", "")],
       [("n", "What is the Food [Return Rate]?", '=COUNTIFS(tLines[category],"Food",tLines[order_status],"Returned")/COUNTIF(tLines[category],"Food")',
         f(rr["Food"]), "pct", "Return rate Food"),
        ("n", "What is the Electronics [Return Rate]?", '=COUNTIFS(tLines[category],"Electronics",tLines[order_status],"Returned")/COUNTIF(tLines[category],"Electronics")',
         f(rr["Electronics"]), "pct", "Return rate Electronics")],
       steps=["DIVIDE ( a, b, 0 ) returns 0 instead of BLANK when b is 0.", "IFERROR exists in DAX but is slow: prefer DIVIDE.",
              "A seller with no orders shows 0.0%, not an error."], code=True)
    tp(m, "3.6.31",
       [("[Net Revenue] (kept in _Measures)", "=" + net(), NETV, N, ""),
        ("[Orders]", "=ROWS(tTransactions)", len(TX), "int", ""),
        ("[AOV]", "=(" + net() + ")/(" + netn() + ")", NETV / len(NT), N, "")],
       [("t", "In Power Pivot, after moving every measure to _Measures, what do you hide so only measures show?", "the dummy column", None, None,
         ("dummy column", "column1", "the column"))],
       steps=["Power Query › Enter Data › one column › load to the Data Model as _Measures.", "Move each measure: Measure Settings › Table name: _Measures.",
              "Hide the dummy column: _Measures moves to the top of the field list."])
    tp(m, "3.6.32",
       [("[Net Revenue] as ₦#,##0", "=" + net(), NETV, N, ""),
        ("[YoY %] 2025 as 0.0%", "=XLOOKUP(2025,tYear[Year],tYear[YoY])", f(yr[2025] / yr[2024] - 1), "pct", ""),
        ("[Orders] as #,##0", "=ROWS(tTransactions)", len(TX), "int", "")],
       [("t", "Which format string shows 0.238 as 23.8%?", "0.0%", None, None, ())],
       steps=["Power Pivot › Measures › Manage Measures › Edit.", "Category: Currency (₦) / Number / Percentage, with decimals.",
              "Set it once on the measure: every PivotTable inherits it."])
    tp(m, "3.6.33",
       [("Report A: Sum of total_price (implicit)", "=SUM(tTransactions[total_price])", f(TX.total_price.sum()), N, "Dragging the column in."),
        ("Report B: [Net Revenue] (explicit)", "=" + net(), NETV, N, ""),
        ("Difference", "=SUM(tTransactions[total_price])-(" + net() + ")", f(TX.total_price.sum()) - NETV, N, "Cancelled + Returned counted in A.")],
       [("n", "How much bigger is report A than report B, as a % of B?", "=SUM(tTransactions[total_price])/(" + net() + ")-1",
         f(TX.total_price.sum() / NETV - 1), "pct", "A over B")],
       steps=["Implicit: drag a numeric column to Values (Excel writes Sum of …).", "Explicit: a named measure with its definition.",
              "Use explicit measures only, so every report agrees."])
    tp(m, "3.6.34",
       [("Rows FILTER ( tTransactions, … ) scans", "=ROWS(tTransactions)", len(TX), "int", ""),
        ("Values a column filter scans (order_status)", "=COUNTA(UNIQUE(tTransactions[order_status]))", int(TX.order_status.nunique()), "int",
         "Why CALCULATE with a column filter is faster.")],
       [("n", "How many distinct payment_method values would a column filter on payment_method scan?", "=COUNTA(UNIQUE(tTransactions[payment_method]))",
         int(TX.payment_method.nunique()), "int", "Distinct payment methods")],
       steps=["Filter columns, not whole tables.", "Use measures, not calculated columns, for totals.", "Remove unused columns before loading (Power Query)."])
    s25 = f(DS.Sales.sum())
    o25 = int(DS.Order_ID.nunique())
    tp(m, "3.6.35",
       [("Sales", "=SUM(tDash[Sales])", s25, N, ""),
        ("Margin %", "=SUM(tDash[Profit])/SUM(tDash[Sales])", f(DS.Profit.sum() / DS.Sales.sum()), "pct2", ""),
        ("Orders", "=COUNTA(UNIQUE(tDash[Order_ID]))", o25, "int", "DISTINCTCOUNT ( Order_ID )"),
        ("Basket value", "=SUM(tDash[Sales])/COUNTA(UNIQUE(tDash[Order_ID]))", s25 / o25, N, ""),
        ("Target attainment", "=SUM(tDash[Sales])/SUM(tDashStores[Target])", s25 / f(DSS.Target.sum()), "pct", "")],
       [("n", "Units per basket (Quantity ÷ Orders, 1 decimal)?", "=SUM(tDash[Quantity])/COUNTA(UNIQUE(tDash[Order_ID]))", f(DS.Quantity.sum() / o25), "dec1",
         "Units per basket"),
        ("n", "South-West margin %?", '=SUMIF(tDash[Zone],"South-West",tDash[Profit])/SUMIF(tDash[Zone],"South-West",tDash[Sales])',
         f(DS[DS.Zone == "South-West"].Profit.sum() / DS[DS.Zone == "South-West"].Sales.sum()), "pct", "SW margin")],
       steps=["Load tDash and tDashStores to the model; relate Store_ID.", "Create Sales, Profit, Margin %, Orders, Basket, Units per Basket, Attainment.",
              "Check each against the twin here."])
    tp(m, "3.6.36",
       [("Admissions 2023", "=ROWS(tAdm23)", len(A), "int", "The slide counts all years 2019-2024; this sheet uses 2023."),
        ("ALOS 2023 (days)", "=AVERAGE(tAdm23[Length_of_Stay])", f(A.Length_of_Stay.mean()), "dec1", ""),
        ("30-day readmission rate 2023", '=COUNTIF(tAdm23[Readmit_30d],"Yes")/ROWS(tAdm23)', f((A.Readmit_30d == "Yes").mean()), "pct", ""),
        ("Average occupancy (all census days)", "=AVERAGE(tCensus[Occupancy_Pct])", f(CEN.Occupancy_Pct.mean()), "pct", "")],
       [("n", "Median length of stay in 2023?", "=MEDIAN(tAdm23[Length_of_Stay])", f(A.Length_of_Stay.median()), "dec1", "Median LOS 2023"),
        ("n", "How many census days were above 95% full?", '=COUNTIF(tCensus[Occupancy_Pct],">0.95")', int((CEN.Occupancy_Pct > 0.95).sum()), "int",
         "Days above 95%")],
       steps=["Admissions := COUNTROWS ( tAdm23 ); ALOS := AVERAGE ( tAdm23[Length_of_Stay] ).",
              "Readmit % := DIVIDE ( CALCULATE ( [Admissions], tAdm23[Readmit_30d] = \"Yes\" ), [Admissions] ).", "Occupancy from tCensus, not from admissions."])
    tp(m, "3.6.37",
       [("Enrolment", "=ROWS(tCV)", len(CV), "int", ""),
        ("Average exam score", "=AVERAGE(tCV[Exam_Score])", f(CV.Exam_Score.mean()), "dec1", ""),
        ("Grade A rate (70+)", '=COUNTIF(tCV[Exam_Score],">=70")/ROWS(tCV)', f((CV.Exam_Score >= 70).mean()), "pct", ""),
        ("At-risk students", '=COUNTIF(tCV[At_Risk],"Yes")', int((CV.At_Risk == "Yes").sum()), "int", "")],
       [("n", "Average attendance (as %)?", "=AVERAGE(tCV[Attendance_Pct])/100", f(CV.Attendance_Pct.mean() / 100), "pct", "Average attendance"),
        ("n", "At-risk share of students?", '=COUNTIF(tCV[At_Risk],"Yes")/ROWS(tCV)', f((CV.At_Risk == "Yes").mean()), "pct", "At-risk share")],
       steps=["Enrolment := COUNTROWS ( tCV ).", "Grade A % := DIVIDE ( CALCULATE ( [Enrolment], tCV[Exam_Score] >= 70 ), [Enrolment] ).",
              "Faculty and Level on rows; check each KPI against the twin."])
    tp(m, "3.6.38",
       [("No relationship: every year shows", "=" + net(), NETV, N, "The same total on every row = a missing relationship."),
        ("After relating: 2020", "=XLOOKUP(2020,tYear[Year],tYear[Net_Revenue])", f(yr[2020]), N, "")],
       [("n", "After the fix, what does 2023 show?", "=XLOOKUP(2023,tYear[Year],tYear[Net_Revenue])", f(yr[2023]), N, "2023 after the fix")],
       steps=["Same number on every row: the row field's table is not related to the fact.", "Check Diagram View: is the relationship there and active?",
              "Fix, then compare with a SUMIFS twin."])
    tp(m, "3.6.39",
       [("[Net Revenue] twin", "=" + net(t="tTransactions"), NETV, N, ""),
        ("[Net Revenue] 2025 twin", "=" + net(order_year=2025), f(yr[2025]), N, ""),
        ("Customers twin", "=COUNTA(UNIQUE(tTransactions[customer_id]))", int(TX.customer_id.nunique()), "int", "")],
       [("n", "Write the twin for [Net Revenue] in 2021. What is it?", "=" + net(order_year=2021), f(yr[2021]), N, "Net revenue 2021")],
       steps=["Put the measure in a PivotTable with no filters.", "Work the same number out with SUMIFS on the data sheet.",
              "Equal = tested. Different = find out why before anyone uses the measure."])
    tp(m, "3.6.40",
       [("Expected total for [Net Revenue]", "=" + net(), NETV, N, "Write it in the measure's Description.")],
       [("o", "Write the documentation entry for [AOV].",
         "Definition: net revenue ÷ net order lines (Delivered + Completed). Grain: order line. Source: tTransactions. "
         "Expected total ₦126,688. Owner and date of last twin check.")],
       steps=["For every measure: name, definition, grain, source, expected total, owner, date checked.",
              "Put it in the measure's Description box.", "Keep a Measures sheet in 00_README of the model workbook."])
    tp(m, "3.6.41",
       [("Total: SUM ( column )", "=SUM(tTransactions[total_price])", f(TX.total_price.sum()), N, ""),
        ("Filtered total: CALCULATE", "=" + net(), NETV, N, ""),
        ("Year to date (Dec 2025)", "=" + netd("2025-01-01", "2025-12-31"), f(yr[2025]), N, ""),
        ("Growth (2025)", "=XLOOKUP(2025,tYear[Year],tYear[YoY])", f(yr[2025] / yr[2024] - 1), "pct", "")],
       [("n", "Share pattern: Electronics share of net revenue?", "=(" + net(category="Electronics") + ")/(" + net() + ")",
         f(NT[NT.category == "Electronics"].total_price.sum() / NETV), "pct", "Electronics share")],
       steps=["Context first: what filters reach this cell?", "CALCULATE to change filters; iterators (SUMX) for row maths.",
              "Time intelligence on a marked date table; DIVIDE for every ratio."])
    m.finish()
