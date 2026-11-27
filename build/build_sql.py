"""
ETEH Bootcamp - Phase A3: SQL artifacts.

1) ecommerce.db  -- the relational star of the bootcamp.
   Five normalised tables (transactions, sellers, products, deliveries, customers)
   with PRIMARY KEYs, FOREIGN KEYs and indexes so learners can practise JOINs,
   aggregation and window functions. Transactions were reconciled by
   build/reconcile_ecommerce.py so every foreign key resolves.

2) ecommerce_schema.sql        -- DDL (tables + FKs + indexes), re-runnable.
   ecommerce_sample_queries.sql -- annotated JOIN / aggregation / window examples.

3) all_datasets.db -- every cleaned dataset loaded as a flat table so the SQL
   module has one portable file to query.
"""
import os, sqlite3
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CLEAN = os.path.join(ROOT, "Datasets", "Clean")
SQL   = os.path.join(ROOT, "Datasets", "SQL")
os.makedirs(SQL, exist_ok=True)

ECOM_SCHEMA = """
PRAGMA foreign_keys = ON;

DROP TABLE IF EXISTS transactions;
DROP TABLE IF EXISTS deliveries;
DROP TABLE IF EXISTS products;
DROP TABLE IF EXISTS sellers;
DROP TABLE IF EXISTS customers;

CREATE TABLE customers (
    customer_id      TEXT PRIMARY KEY,
    first_name       TEXT,
    last_name        TEXT,
    email            TEXT,
    phone_number     TEXT,
    gender           TEXT,
    date_of_birth    TEXT,
    city             TEXT,
    state            TEXT,
    country          TEXT,
    membership_tier  TEXT,
    loyalty_points   INTEGER,
    total_orders     INTEGER,
    total_spent      REAL,
    created_at       TEXT
);

CREATE TABLE sellers (
    seller_id       TEXT PRIMARY KEY,
    store_name      TEXT,
    owner_name      TEXT,
    email           TEXT,
    phone_number    TEXT,
    store_city      TEXT,
    store_state     TEXT,
    store_country   TEXT,
    rating          REAL,
    total_sales     INTEGER,
    total_revenue   REAL,
    account_status  TEXT,
    created_at      TEXT
);

CREATE TABLE products (
    product_id            TEXT PRIMARY KEY,
    product_name          TEXT,
    description           TEXT,
    category              TEXT,
    subcategory           TEXT,
    brand_name            TEXT,
    seller_id             TEXT,
    price                 REAL,
    discount_price        REAL,
    commission_percentage REAL,
    stock_quantity        INTEGER,
    weight                REAL,
    total_sold            INTEGER,
    average_rating        REAL,
    created_at            TEXT,
    FOREIGN KEY (seller_id) REFERENCES sellers(seller_id)
);

CREATE TABLE deliveries (
    delivery_id      TEXT PRIMARY KEY,
    order_id         TEXT,
    customer_id      TEXT,
    delivery_address TEXT,
    city             TEXT,
    state            TEXT,
    country          TEXT,
    courier_name     TEXT,
    tracking_number  TEXT,
    delivery_status  TEXT,
    shipped_date     TEXT,
    delivered_date   TEXT,
    FOREIGN KEY (customer_id) REFERENCES customers(customer_id)
);

CREATE TABLE transactions (
    transaction_id          TEXT PRIMARY KEY,
    order_id                TEXT,
    customer_id             TEXT,
    product_id              TEXT,
    seller_id               TEXT,
    delivery_id             TEXT,
    quantity                INTEGER,
    unit_price              REAL,
    total_price             REAL,
    commission_percentage   REAL,
    commission_amount       REAL,
    order_status            TEXT,
    payment_method          TEXT,
    payment_status          TEXT,
    order_date              TEXT,
    FOREIGN KEY (customer_id)  REFERENCES customers(customer_id),
    FOREIGN KEY (product_id)   REFERENCES products(product_id),
    FOREIGN KEY (seller_id)    REFERENCES sellers(seller_id),
    FOREIGN KEY (delivery_id)  REFERENCES deliveries(delivery_id)
);

CREATE INDEX idx_txn_customer ON transactions(customer_id);
CREATE INDEX idx_txn_product  ON transactions(product_id);
CREATE INDEX idx_txn_seller   ON transactions(seller_id);
CREATE INDEX idx_txn_order    ON transactions(order_id);
CREATE INDEX idx_prod_seller  ON products(seller_id);
CREATE INDEX idx_del_customer ON deliveries(customer_id);
CREATE INDEX idx_del_order    ON deliveries(order_id);
"""

ECOM_FILES = {
    "customers":    "ecommerce_customers.csv",
    "sellers":      "ecommerce_sellers.csv",
    "products":     "ecommerce_products.csv",
    "deliveries":   "ecommerce_deliveries.csv",
    "transactions": "ecommerce_transactions_reconciled.csv",
}

SAMPLE_QUERIES = """-- =====================================================================
-- ETEH Bootcamp - ecommerce.db  Sample Analytical Queries
-- Run with:  sqlite3 ecommerce.db < ecommerce_sample_queries.sql
-- =====================================================================

-- 1) Revenue by product category (JOIN + aggregation)
SELECT p.category,
       ROUND(SUM(t.total_price), 2) AS revenue,
       COUNT(*)                     AS orders
FROM transactions t
JOIN products p ON p.product_id = t.product_id
GROUP BY p.category
ORDER BY revenue DESC;

-- 2) Top 10 sellers by revenue (JOIN + aggregation + ordering)
SELECT s.seller_id, s.store_name, s.store_state,
       ROUND(SUM(t.total_price), 2) AS revenue,
       COUNT(*)                     AS transactions
FROM transactions t
JOIN sellers s ON s.seller_id = t.seller_id
GROUP BY s.seller_id
ORDER BY revenue DESC
LIMIT 10;

-- 3) Customer lifetime value (JOIN customers -> transactions)
SELECT c.customer_id, c.first_name || ' ' || c.last_name AS customer,
       c.membership_tier,
       ROUND(SUM(t.total_price), 2) AS lifetime_spend,
       COUNT(*)                     AS orders
FROM customers c
JOIN transactions t ON t.customer_id = c.customer_id
GROUP BY c.customer_id
ORDER BY lifetime_spend DESC
LIMIT 15;

-- 4) Delivery performance: avg days in transit per courier
SELECT d.courier_name,
       ROUND(AVG(julianday(d.delivered_date) - julianday(d.shipped_date)), 2) AS avg_transit_days,
       COUNT(*) AS parcels
FROM deliveries d
WHERE d.delivered_date IS NOT NULL
GROUP BY d.courier_name
ORDER BY avg_transit_days;

-- 5) Monthly revenue trend
SELECT substr(t.order_date, 1, 7) AS month,
       ROUND(SUM(t.total_price), 2) AS revenue,
       ROUND(SUM(t.commission_amount), 2) AS commission
FROM transactions t
GROUP BY month
ORDER BY month;

-- 6) Product sell-through (sold vs stock)
SELECT p.product_id, p.product_name, p.category,
       p.total_sold, p.stock_quantity,
       ROUND(100.0 * p.total_sold / (p.total_sold + p.stock_quantity), 1) AS sell_through_pct
FROM products p
ORDER BY sell_through_pct DESC
LIMIT 15;

-- 7) State-level revenue with RANK() window function
SELECT c.state,
       ROUND(SUM(t.total_price), 2) AS revenue,
       RANK() OVER (ORDER BY SUM(t.total_price) DESC) AS rk
FROM transactions t
JOIN customers c ON c.customer_id = t.customer_id
GROUP BY c.state
ORDER BY revenue DESC
LIMIT 12;

-- 8) Payment method mix (aggregate + percentage)
SELECT payment_method,
       COUNT(*) AS txns,
       ROUND(100.0 * COUNT(*) / SUM(COUNT(*)) OVER (), 1) AS pct
FROM transactions
GROUP BY payment_method
ORDER BY txns DESC;
"""

def load_table(conn, table, csv_path):
    df = pd.read_csv(csv_path, dtype=str, keep_default_na=False)
    cols = list(df.columns)
    placeholders = ", ".join("?" * len(cols))
    col_list = ", ".join(cols)
    cur = conn.cursor()
    cur.execute(f"DELETE FROM {table}")
    for _, row in df.iterrows():
        vals = [None if v == "" else v for v in row.tolist()]
        cur.execute(f"INSERT INTO {table} ({col_list}) VALUES ({placeholders})", vals)
    conn.commit()
    print(f"   loaded {len(df):>6} rows -> {table}")

def build_ecommerce():
    db = os.path.join(SQL, "ecommerce.db")
    if os.path.exists(db):
        os.remove(db)
    conn = sqlite3.connect(db)
    conn.executescript(ECOM_SCHEMA)
    print("Building ecommerce.db:")
    for table, fname in ECOM_FILES.items():
        load_table(conn, table, os.path.join(CLEAN, fname))
    conn.execute("PRAGMA foreign_keys = ON;")
    fk = conn.execute("PRAGMA foreign_key_check").fetchall()
    print("   FK check:", "OK (no violations)" if not fk else f"VIOLATIONS: {fk}")
    conn.close()
    with open(os.path.join(SQL, "ecommerce_schema.sql"), "w") as f:
        f.write(ECOM_SCHEMA)
    with open(os.path.join(SQL, "ecommerce_sample_queries.sql"), "w") as f:
        f.write(SAMPLE_QUERIES)
    print("   wrote ecommerce_schema.sql + ecommerce_sample_queries.sql")

def build_all():
    db = os.path.join(SQL, "all_datasets.db")
    if os.path.exists(db):
        os.remove(db)
    conn = sqlite3.connect(db)
    flat = {
        "clean_healthcare_records": "healthcare_records.csv",
        "clean_salary_data": "salary_data.csv",
        "clean_superstore_sales": "superstore_sales.csv",
        "clean_student_depression": "student_depression.csv",
        "clean_student_performance": "student_performance.csv",
        "clean_social_media_addiction": "social_media_addiction.csv",
        "clean_diabetes_prediction": "diabetes_prediction.csv",
        "clean_hotel_reservations": "hotel_reservations.csv",
    }
    print("Building all_datasets.db:")
    for tname, fname in flat.items():
        df = pd.read_csv(os.path.join(CLEAN, fname))
        df.to_sql(tname, conn, if_exists="replace", index=False)
        print(f"   loaded {len(df):>7} rows -> {tname}")
    for table, fname in ECOM_FILES.items():
        df = pd.read_csv(os.path.join(CLEAN, fname), dtype=str, keep_default_na=False)
        df.to_sql("ecom_" + table, conn, if_exists="replace", index=False)
        print(f"   loaded {len(df):>6} rows -> ecom_{table}")
    conn.close()

if __name__ == "__main__":
    build_ecommerce()
    build_all()
    print("\nSQL artifacts complete.")
