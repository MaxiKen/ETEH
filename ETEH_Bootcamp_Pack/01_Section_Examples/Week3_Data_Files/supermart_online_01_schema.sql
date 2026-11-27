-- =====================================================================
-- SuperMart Online marketplace - relational schema (SQLite / ANSI SQL)
-- ETEH Four-Week Excel Data Analytics Bootcamp - Week 3
-- Load order: nigeria_states -> customers -> sellers -> products -> deliveries -> transactions
-- =====================================================================
PRAGMA foreign_keys = ON;

DROP TABLE IF EXISTS transactions;
DROP TABLE IF EXISTS deliveries;
DROP TABLE IF EXISTS products;
DROP TABLE IF EXISTS sellers;
DROP TABLE IF EXISTS customers;
DROP TABLE IF EXISTS nigeria_states;

CREATE TABLE nigeria_states (
    state            TEXT PRIMARY KEY,
    zone             TEXT NOT NULL
);

CREATE TABLE customers (
    customer_id      TEXT PRIMARY KEY,
    first_name       TEXT NOT NULL,
    last_name        TEXT NOT NULL,
    email            TEXT,
    phone_number     TEXT,
    gender           TEXT CHECK (gender IN ('Female','Male')),
    date_of_birth    DATE,
    city             TEXT,
    state            TEXT REFERENCES nigeria_states(state),
    zone             TEXT,
    country          TEXT DEFAULT 'Nigeria',
    membership_tier  TEXT CHECK (membership_tier IN ('Bronze','Silver','Gold','Platinum')),
    loyalty_points   INTEGER,
    total_orders     INTEGER,
    total_spent      REAL,
    created_at       DATE
);

CREATE TABLE sellers (
    seller_id        TEXT PRIMARY KEY,
    store_name       TEXT NOT NULL,
    owner_name       TEXT,
    email            TEXT,
    phone_number     TEXT,
    store_city       TEXT,
    store_state      TEXT REFERENCES nigeria_states(state),
    store_zone       TEXT,
    store_country    TEXT DEFAULT 'Nigeria',
    rating           REAL CHECK (rating BETWEEN 0 AND 5),
    total_sales      INTEGER,
    total_revenue    REAL,
    account_status   TEXT,
    created_at       DATE
);

CREATE TABLE products (
    product_id            TEXT PRIMARY KEY,
    product_name          TEXT NOT NULL,
    description           TEXT,
    category              TEXT NOT NULL,
    subcategory           TEXT,
    brand_name            TEXT,
    seller_id             TEXT REFERENCES sellers(seller_id),
    price                 REAL CHECK (price > 0),
    discount_price        REAL,
    commission_percentage REAL,
    stock_quantity        INTEGER,
    weight                REAL,
    total_sold            INTEGER,
    average_rating        REAL,
    created_at            DATE
);

CREATE TABLE deliveries (
    delivery_id      TEXT PRIMARY KEY,
    order_id         TEXT NOT NULL UNIQUE,
    customer_id      TEXT REFERENCES customers(customer_id),
    delivery_address TEXT,
    city             TEXT,
    state            TEXT REFERENCES nigeria_states(state),
    zone             TEXT,
    country          TEXT DEFAULT 'Nigeria',
    courier_name     TEXT,
    tracking_number  TEXT,
    delivery_status  TEXT CHECK (delivery_status IN ('Delivered','Pending')),
    shipped_date     DATE,
    delivered_date   DATE
);

CREATE TABLE transactions (
    transaction_id        TEXT PRIMARY KEY,
    order_id              TEXT NOT NULL UNIQUE,
    customer_id           TEXT NOT NULL REFERENCES customers(customer_id),
    product_id            TEXT NOT NULL REFERENCES products(product_id),
    seller_id             TEXT NOT NULL REFERENCES sellers(seller_id),
    delivery_id           TEXT REFERENCES deliveries(delivery_id),
    quantity              INTEGER CHECK (quantity > 0),
    unit_price            REAL,
    total_price           REAL,
    commission_percentage REAL,
    commission_amount     REAL,
    order_status          TEXT,
    payment_method        TEXT,
    payment_status        TEXT,
    order_date            DATE NOT NULL
);

CREATE INDEX ix_trx_customer ON transactions(customer_id);
CREATE INDEX ix_trx_product  ON transactions(product_id);
CREATE INDEX ix_trx_seller   ON transactions(seller_id);
CREATE INDEX ix_trx_date     ON transactions(order_date);
CREATE INDEX ix_prod_cat     ON products(category);

-- A reusable analytical view: one wide row per order line
DROP VIEW IF EXISTS v_order_lines;
CREATE VIEW v_order_lines AS
SELECT t.transaction_id, t.order_id, t.order_date,
       strftime('%Y', t.order_date) AS order_year,
       strftime('%Y-%m', t.order_date) AS order_month,
       t.customer_id, c.first_name || ' ' || c.last_name AS customer_name, c.membership_tier,
       c.state AS customer_state, c.zone AS customer_zone,
       t.product_id, p.product_name, p.category, p.subcategory, p.brand_name,
       t.seller_id, s.store_name, s.store_state,
       t.quantity, t.unit_price, t.total_price, t.commission_amount,
       t.order_status, t.payment_method, t.payment_status,
       d.courier_name, d.delivery_status, d.shipped_date, d.delivered_date,
       julianday(d.delivered_date) - julianday(t.order_date) AS days_to_deliver
FROM transactions t
JOIN customers c ON c.customer_id = t.customer_id
JOIN products  p ON p.product_id  = t.product_id
JOIN sellers   s ON s.seller_id   = t.seller_id
LEFT JOIN deliveries d ON d.delivery_id = t.delivery_id;
