
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
