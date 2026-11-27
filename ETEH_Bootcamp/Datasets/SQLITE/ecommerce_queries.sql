-- ============================================================
-- ETEH Bootcamp — E-commerce SQLite schema & teaching queries
-- Database: eteh_ecommerce.sqlite3  (Week 3, Module 3.2 / 3.4)
-- Run in DB Browser for SQLite or sqlite3 CLI.
-- ============================================================

-- Tables (grain declared):
--   customers    1 row per customer_id (PK)
--   sellers      1 row per seller_id (PK)
--   products     1 row per product_id (PK, FK seller_id -> sellers)
--   deliveries   1 row per delivery_id (PK, FK order_id, customer_id)
--   transactions 1 row per transaction_id (PK, FKs to all dims)
-- NOTE: 570 Cancelled transactions have NO delivery row (cancelled before
-- shipment). Use LEFT JOIN from transactions to deliveries to keep them.

-- ---------- Q1: Revenue by category (INNER JOIN) ----------
SELECT p.category, COUNT(*) AS lines, SUM(t.total_price) AS revenue
FROM transactions t
JOIN products p ON t.product_id = p.product_id
GROUP BY p.category ORDER BY revenue DESC;

-- ---------- Q2: Revenue by state (JOIN via deliveries) ----------
SELECT d.state, COUNT(*) AS deliveries, SUM(t.total_price) AS revenue
FROM transactions t
JOIN deliveries d ON t.order_id = d.order_id
GROUP BY d.state ORDER BY revenue DESC;

-- ---------- Q3: Keep cancelled orders (LEFT JOIN) ----------
SELECT t.order_status, COUNT(*) AS n, SUM(t.total_price) AS value
FROM transactions t
LEFT JOIN deliveries d ON t.order_id = d.order_id
GROUP BY t.order_status;

-- ---------- Q4: Top sellers by commission ----------
SELECT s.store_name, s.store_city, COUNT(*) AS lines,
       SUM(t.commission_amount) AS commission
FROM transactions t
JOIN sellers s ON t.seller_id = s.seller_id
GROUP BY s.seller_id ORDER BY commission DESC LIMIT 10;

-- ---------- Q5: Monthly revenue trend (for forecasting) ----------
SELECT order_month, SUM(total_price) AS revenue, COUNT(*) AS lines
FROM transactions GROUP BY order_month ORDER BY order_month;

-- ---------- Q6: Orphan check (should return 0 for customers/products) ----------
SELECT COUNT(*) AS orphan_customer_lines
FROM transactions t LEFT JOIN customers c ON t.customer_id = c.customer_id
WHERE c.customer_id IS NULL;

-- ---------- Q7: Payment mix ----------
SELECT payment_method, payment_status, COUNT(*) AS n, SUM(total_price) AS value
FROM transactions GROUP BY payment_method, payment_status ORDER BY value DESC;
