-- =====================================================================
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
