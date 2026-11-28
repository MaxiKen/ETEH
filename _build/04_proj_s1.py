"""Student Project 1 - SuperMart Online marketplace: SQL + Excel performance review."""
from __future__ import annotations

import sqlite3

import pandas as pd

from asg_common import Asg
from common import SPR, W3D
from proj_common import SKEY, brief_pdf
from wk_common import DATASETS_TEXT

NAME = "Project1_SuperMart_Online_SQL"
FOLDER = SPR / "Week_5_Capstone" / NAME
(FOLDER / "Data").mkdir(parents=True, exist_ok=True)
import shutil  # noqa: E402
from common import CORE  # noqa: E402
shutil.copy2(W3D / "supermart_online.db", FOLDER / "Data" / "supermart_online.db")
for _t in ("transactions", "customers", "products", "sellers", "deliveries"):
    shutil.copy2(CORE / f"SuperMart_Online_{_t}.csv", FOLDER / "Data" / f"SuperMart_Online_{_t}.csv")
con = sqlite3.connect(W3D / "supermart_online.db")


def one(sql):
    v = pd.read_sql_query(sql, con).iloc[0, 0]
    return v.item() if hasattr(v, "item") else v


NET = "order_status IN ('Delivered','Completed')"
Y25 = "t.order_date LIKE '2025%'"
AUDIT = [(f"Rows in {t}", f"SELECT COUNT(*) FROM {t};") for t in ["customers", "sellers", "products", "deliveries", "transactions"]]
AUDIT += [("Transactions with a NULL delivery_id", "SELECT COUNT(*) FROM transactions WHERE delivery_id IS NULL;"),
          ("...of which are NOT Cancelled", "SELECT COUNT(*) FROM transactions WHERE delivery_id IS NULL AND order_status <> 'Cancelled';"),
          ("Products with a blank (NULL) seller_id", "SELECT COUNT(*) FROM products WHERE seller_id IS NULL;"),
          ("First order date", "SELECT MIN(order_date) FROM transactions;"),
          ("Rows returned by the WRONG join transactions-deliveries ON customer_id", "SELECT COUNT(*) FROM transactions t JOIN deliveries d ON d.customer_id = t.customer_id;")]
QS = [
    ("Net revenue in 2025 (Delivered + Completed), ₦.", f"SELECT SUM(total_price) AS net_revenue_2025\nFROM transactions t\nWHERE {Y25}\n  AND t.{NET};", "naira"),
    ("Net revenue in 2024 (Delivered + Completed), ₦. (Same query, one change.)",
     f"SELECT SUM(total_price) AS net_revenue_2024\nFROM transactions t\nWHERE t.order_date LIKE '2024%'\n  AND t.{NET};", "naira"),
    ("Category with the highest 2025 net revenue.",
     f"SELECT p.category, SUM(t.total_price) AS rev\nFROM transactions t\nJOIN products p ON p.product_id = t.product_id\nWHERE {Y25} AND t.{NET}\nGROUP BY p.category\nORDER BY rev DESC\nLIMIT 1;", "gen"),
    ("Seller (store_name) with the highest 2025 net revenue.",
     f"SELECT s.store_name, SUM(t.total_price) AS rev\nFROM transactions t\nJOIN sellers s ON s.seller_id = t.seller_id\nWHERE {Y25} AND t.{NET}\nGROUP BY s.store_name\nORDER BY rev DESC\nLIMIT 1;", "gen"),
    ("Platform commission earned on 2025 net orders, ₦ (nearest naira).",
     f"SELECT ROUND(SUM(commission_amount), 0) AS commission\nFROM transactions t\nWHERE {Y25} AND t.{NET};", "naira"),
    ("How many 2025 order lines were Returned?",
     f"SELECT COUNT(*) AS returned_2025\nFROM transactions t\nWHERE {Y25}\n  AND t.order_status = 'Returned';", "int"),
    ("Payment method with the most Cancelled order lines (all years).",
     "SELECT payment_method, COUNT(*) AS cancelled\nFROM transactions\nWHERE order_status = 'Cancelled'\nGROUP BY payment_method\nORDER BY cancelled DESC\nLIMIT 1;", "gen"),
    ("Customer zone with the highest 2025 net revenue.",
     f"SELECT c.zone, SUM(t.total_price) AS rev\nFROM transactions t\nJOIN customers c ON c.customer_id = t.customer_id\nWHERE {Y25} AND t.{NET}\nGROUP BY c.zone\nORDER BY rev DESC\nLIMIT 1;", "gen"),
    ("How many different customers placed a net order in 2025?",
     f"SELECT COUNT(DISTINCT customer_id) AS active_2025\nFROM transactions t\nWHERE {Y25} AND t.{NET};", "int"),
    ("Product category with the most products listed (products table).",
     "SELECT category, COUNT(*) AS products\nFROM products\nGROUP BY category\nORDER BY products DESC\nLIMIT 1;", "gen"),
]
AUD_A = [one(s) for _, s in AUDIT]
QS_A = [one(s) for _, s, _ in QS]

# supporting details for guides
def frame(sql):
    return pd.read_sql_query(sql, con)


cat_top = frame(QS[2][1].replace("LIMIT 1", "LIMIT 3"))
pay = frame("SELECT payment_method, COUNT(*) AS n FROM transactions WHERE order_status = 'Cancelled' GROUP BY payment_method ORDER BY n DESC")
zone = frame(QS[7][1].replace("LIMIT 1", ""))
net25 = frame(f"SELECT t.*, c.zone FROM transactions t JOIN customers c USING(customer_id) WHERE {Y25} AND t.{NET}")
aov = net25.total_price.mean()
act = net25.customer_id.nunique()

(FOLDER / "Project1_SQL_Template.sql").write_text("\n".join(
    ["-- ETEH Student Project 1 - SuperMart Online. One query per question; keep the question comment above it.", ""] +
    [f"-- A{i}. {t}\n\n" for i, (t, _) in enumerate(AUDIT, 1)] + [f"-- Q{i}. {t}\n\n" for i, (t, _, _) in enumerate(QS, 1)]), encoding="utf-8")
(SKEY / "Project1_SuperMart_Online_SQL_KEY.sql").write_text("\n".join(
    ["-- ETEH Student Project 1 - ANSWER KEY", ""] + [f"-- A{i}. {t}  -> {a}\n{s}\n" for i, ((t, s), a) in enumerate(zip(AUDIT, AUD_A), 1)] +
    [f"-- Q{i}. {t}  -> {a}\n{s}\n" for i, ((t, s, _), a) in enumerate(zip(QS, QS_A), 1)]), encoding="utf-8")


def build(key):
    A = Asg("P1", "Student Project 1 - SuperMart Online Marketplace Performance Review", key,
            "Student Project 1 · SQL, data model and dashboard · supermart_online.db", path=(SKEY / f"{NAME}_KEY.xlsx") if key else (FOLDER / f"{NAME}_Workbook.xlsx"),
            label="Student Project 1", submission="Submit this workbook (answers), your .sql file, your dashboard workbook, a 1-page memo (PDF) and a 5-slide deck. "
                                                  "File names: Surname_Firstname_P1_*.")
    s = A.task("M1 Data Audit", "Milestone 1 - Audit the database before trusting it", "Weeks 1-3 (3.1-3.4)",
               "Open Data › supermart_online.db (the Data folder inside this project folder) in DB Browser for SQLite. Answer with SQL (write the queries in Project1_SQL_Template.sql).")
    for i, ((t, _), a) in enumerate(zip(AUDIT, AUD_A), 1):
        k = "gen" if isinstance(a, str) else "int"
        gd = {6: "All 570 are Cancelled orders - a NULL delivery is expected, not an error.", 7: "0 confirms the rule."}.get(i, "")
        if i == 10:
            bad, tot = con.execute("SELECT SUM(t.total_price) FROM transactions t JOIN deliveries d ON d.customer_id = t.customer_id").fetchone()[0], con.execute("SELECT SUM(total_price) FROM transactions").fetchone()[0]
            gd = f"{a:,} rows vs 11,000 transactions - a many-to-many join inflates revenue {bad / tot:.1f}x. Join on delivery_id."
        A.q(s, f"A{i}", t, a, None, k, 2 if i >= 6 else 1, gd, header=(i == 1))
    A.end_task(s)

    s = A.task("M2 SQL Questions", "Milestone 2 - Answer the Head of Marketplace's questions", "3.2-3.4",
               "Net revenue = total_price of orders with order_status Delivered or Completed. Result 2 marks, SQL 2 marks.")
    for i, ((t, sq, k), a) in enumerate(zip(QS, QS_A), 1):
        gd = {3: "Top 3: " + "; ".join(f"{r.category} ₦{r.rev:,.0f}" for r in cat_top.itertuples()),
              2: f"Growth 2025 vs 2024 = Q1 ÷ Q2 − 1 = {QS_A[0] / QS_A[1] - 1:.2%} (work it out in Excel).",
              7: "Cancelled lines: " + ", ".join(f"{r.payment_method} {r.n}" for r in pay.itertuples()) + " - differences are small; say so.",
              8: "Zones: " + ", ".join(f"{r.zone} ₦{r.rev:,.0f}" for r in zone.itertuples())}.get(i, "")
        A.q(s, f"Q{i}", t, a, None, k, 2, gd, header=True)
        A.long(s, "SQL", "Your SQL:", sq, marks=2, rows=max(3, sq.count("\n") + 1), guide="Any correct query earns the marks; check filters (year AND status).")
    A.end_task(s)

    s = A.task("M3 Excel Model", "Milestone 3 - Load to Excel, model and summarise", "3.3, 3.5-3.6",
               "Load transactions, customers, products, sellers and deliveries into the Data Model (Power Query from the SuperMart_Online_*.csv files in the Data folder, or ODBC). "
               "Add a DimDate table, build a star schema, add a sale_type Conditional Column (Net = Delivered or Completed) in Power Query, "
               "and build one Data Model PivotTable. Report the 2025 values it returns.")
    A.q(s, "3.1", "Net revenue 2025 (₦): Sum of total_price, sale_type = Net, Year = 2025 - must equal Q1.", QS_A[0], None, "naira", 2, header=True)
    A.q(s, "3.2", "Net Orders 2025 (count of Delivered + Completed transactions).", len(net25), None, "int", 2)
    A.q(s, "3.3", "Average Order Value 2025 = net revenue ÷ net orders, in a cell beside the PivotTable (₦, nearest naira).", round(aov), None, "naira", 2)
    A.q(s, "3.4", "Active customers 2025 (Distinct Count of customer_id, sale_type = Net) - must equal Q9.", act, None, "int", 2)
    A.q(s, "3.5", "Relationship deliveries → transactions: which key?", "delivery_id (1:1)", marks=2,
        choices=["delivery_id (1:1)", "customer_id (1:*)", "order_id (*:*)", "No relationship needed"])
    A.long(s, "3.6", "List how you built each value: the Power Query step, the PivotTable field and setting, or the cell ratio "
                     "(Net revenue, Net orders, AOV and growth 2025 vs 2024).",
           "sale_type: Add Column › Conditional Column, Net if order_status = Delivered or Completed, else Not net\n"
           "Net revenue: Values › Sum of total_price; Filters: sale_type = Net; Rows: DimDate Year\n"
           "Net orders: Values › Count of transaction_id, same filters\n"
           "AOV: cell beside the PivotTable = net revenue ÷ net orders\n"
           "Growth: second copy of total_price › Show Values As › % Difference From, Base field Year, Base item (previous)",
           marks=6, rows=6, guide="1.5 marks each for the sale_type column, net revenue / orders, AOV and growth; the Net filter is required.")
    A.long(s, "3.7", "Screenshot the Diagram View into the box below (or describe the relationships).",
           "transactions (fact) *→1 customers, *→1 products, *→1 sellers, 1↔1 deliveries (delivery_id), *→1 DimDate[Date] on order_date. Single direction filters.",
           marks=4, rows=4)
    A.end_task(s)

    s = A.task("M4 Dashboard & Story", "Milestone 4 - Dashboard, memo and 5-slide deck", "3.7-3.9, 4.7-4.9",
               "Build a one-page dashboard (KPI cards, revenue trend, category and zone views, slicers for year and zone). Write a one-page memo to the Head of Marketplace "
               "and a 5-slide deck. Marked against the rubric below.")
    s.grid([["Dashboard (10)", "4 KPI cards with YoY; trend line by month; category bar sorted; zone view; 2 slicers driving every visual; titles state findings; no 3-D"],
            ["Memo (10)", "BLUF first line; 3 findings each with a number from M2/M3; 2 recommendations with owner and next step; data caveats (cancelled/NULL deliveries)"],
            ["Deck (5)", "5 slides: headline, 3 findings, recommendation; one message per slide; numbers reconcile to the workbook"]],
           header=["Component", "What earns full marks"], merge_last=True)
    A.q(s, "4.1", "Dashboard (instructor marks the dashboard workbook).", "Rubric" if key else "", marks=10, header=True)
    A.q(s, "4.2", "Memo.", "Rubric" if key else "", marks=10)
    A.q(s, "4.3", "Deck.", "Rubric" if key else "", marks=5)
    A.long(s, "4.4", "Write your memo's opening (BLUF) sentence here.",
           f"SuperMart Online's 2025 net revenue grew {QS_A[0] / QS_A[1] - 1:.1%} to ₦{QS_A[0] / 1e6:,.1f}m, led by {QS_A[2]}; reducing the "
           f"{QS_A[5]:,} returned lines and winning repeat orders from the {QS_A[8]:,} active customers are the two levers for 2026.", marks=0, rows=2, guide="Model answer - not separately marked; part of 4.2.")
    A.end_task(s)
    return A.finish("Student Project 1 - SuperMart Online Marketplace Performance Review",
                    ["Client: Head of Marketplace, SuperMart Online. Question: how did the marketplace perform in 2025, where is revenue coming from, and what should change in 2026?",
                     "You will audit the database, answer business questions in SQL, build an Excel Data Model with PivotTables, and present a dashboard, memo and deck.",
                     "Duration: assigned at the end of Week 3, due at the end of Week 4 (about 12-15 hours)."],
                    [("SuperMart Online", DATASETS_TEXT["SuperMart Online"])],
                    files="Data folder (inside this project folder) › supermart_online.db and SuperMart_Online_*.csv (5 files). "
                          "SQL template: Project1_SQL_Template.sql (next to this workbook).",
                    rubric=[["SQL", "Correct result 2, correct SQL 2 (any equivalent query)."],
                            ["Reconciliation", "M3 PivotTable values must reconcile to M2 SQL results; unexplained differences lose the marks."],
                            ["Communication", "Findings must quote numbers from your own work."]])


def main():
    for k in (False, True):
        print("P1", "KEY" if k else "student", "->", build(k))
    brief_pdf(FOLDER / "Project1_Brief.pdf", "Student Project 1 - SuperMart Online Marketplace Performance Review",
              "SQL · Power Query merges · Data Model & PivotTables · dashboard · executive memo", "Student Project 1 brief", [
                  ("h", "1. The client and the question"),
                  ("p", "SuperMart Online is SuperMart's marketplace: 200 independent sellers list 800 products and ship to 2,500 registered customers across Nigeria. "
                        "The Head of Marketplace has asked ETEH for an independent review of 2025: <b>How did the marketplace perform, where does net revenue come from, "
                        "and what should we change in 2026?</b>"),
                  ("h", "2. The data"),
                  ("p", "Everything is in the <b>Data</b> folder inside this project folder: supermart_online.db (open it in DB Browser for SQLite) "
                        "and the same five tables as CSV files, SuperMart_Online_&lt;table&gt;.csv, for Power Query."),
                  ("t", [["Table", "Rows", "Grain / key", "Joins to"],
                         ["transactions", f"{AUD_A[4]:,}", "one order line · transaction_id", "customers, products, sellers (customer_id, product_id, seller_id); deliveries (delivery_id)"],
                         ["deliveries", f"{AUD_A[3]:,}", "one shipment · delivery_id", "transactions (1:1); NULL for cancelled orders"],
                         ["customers", f"{AUD_A[0]:,}", "one customer · customer_id", "transactions"],
                         ["products", f"{AUD_A[2]:,}", "one listed product · product_id", "transactions, sellers"],
                         ["sellers", f"{AUD_A[1]:,}", "one seller store · seller_id", "transactions, products"]], [28, 18, 50, 84]),
                  ("note", "Net revenue = total_price of orders whose order_status is Delivered or Completed. Cancelled and Returned orders are excluded. "
                           "Always check row counts after every join - one wrong ON clause in this database inflates revenue more than five-fold."),
                  ("h", "3. Milestones"),
                  ("t", [["Milestone", "What you produce", "Marks"],
                         ["M1 Data audit", "Row counts, NULL/orphan checks, join-safety check (SQL)", "15"],
                         ["M2 SQL questions", "10 business questions answered in SQL with results", "40"],
                         ["M3 Excel model", "Power Query load, star schema, DimDate, 4 PivotTable values reconciled to SQL", "20"],
                         ["M4 Dashboard & story", "One-page dashboard, one-page memo, 5-slide deck", "25"]], [40, 115, 25]),
                  ("h", "4. Deliverables"),
                  ("b", ["Project1_SuperMart_Online_SQL_Workbook.xlsx - answers in the yellow cells", "Your .sql file (one query per question)",
                         "Dashboard workbook (Data Model + dashboard sheet)", "Memo (1 page, PDF) and deck (5 slides, PPTX/PDF)"]),
                  ("h", "5. How you will be assessed"),
                  ("p", "Accuracy (results reconcile across SQL, Data Model and dashboard), method (correct joins and filters, documented queries), and communication "
                        "(findings quote your numbers; recommendations are specific). Pass 50%, distinction 70%."),
              ])


if __name__ == "__main__":
    main()
