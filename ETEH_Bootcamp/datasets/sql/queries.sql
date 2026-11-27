
-- ===========================================================================
-- ETEH COMMERCE  -  analytical query pack
-- Database: eteh_commerce.db  (SQLite, foreign keys enforced)
-- Grain: one row per transaction line in `transactions`.
-- ===========================================================================

-- Q1. Monthly revenue, orders and units sold (time series for forecasting)
SELECT substr(order_date,1,7)            AS month,
       COUNT(DISTINCT order_id)          AS orders,
       SUM(quantity)                     AS units,
       ROUND(SUM(total_price),2)         AS revenue
FROM transactions
GROUP BY substr(order_date,1,7)
ORDER BY month;

-- Q2. Revenue by product category and sub-category (dashboard matrix)
SELECT p.category, p.subcategory,
       COUNT(*)                       AS lines,
       SUM(t.quantity)               AS units,
       ROUND(SUM(t.total_price),2)   AS revenue
FROM transactions t JOIN products p ON t.product_id = p.product_id
GROUP BY p.category, p.subcategory
ORDER BY revenue DESC;

-- Q3. Top 10 sellers by revenue, with account status (join + ranking)
SELECT s.store_name, s.store_state, s.account_status,
       ROUND(SUM(t.total_price),2) AS revenue, COUNT(*) AS lines
FROM transactions t JOIN sellers s ON t.seller_id = s.seller_id
GROUP BY s.seller_id ORDER BY revenue DESC LIMIT 10;

-- Q4. Customer Lifetime Value + loyalty tier
SELECT c.customer_id, c.first_name || ' ' || c.last_name AS customer,
       c.membership_tier, c.loyalty_points,
       COUNT(DISTINCT t.order_id)  AS orders,
       ROUND(SUM(t.total_price),2) AS total_spent,
       ROUND(AVG(t.total_price),2) AS avg_order_value
FROM transactions t JOIN customers c ON t.customer_id = c.customer_id
GROUP BY c.customer_id ORDER BY total_spent DESC LIMIT 15;

-- Q5. Delivery performance per courier
SELECT courier_name, COUNT(*) AS parcels,
       ROUND(AVG(julianday(delivered_date)-julianday(shipped_date)),1) AS avg_days,
       ROUND(100.0*SUM(CASE WHEN delivery_status='Delivered' THEN 1 ELSE 0 END)/COUNT(*),1) AS delivered_pct
FROM deliveries GROUP BY courier_name ORDER BY parcels DESC;

-- Q6. Payment method mix and cancellation rate (funnel)
SELECT payment_method, payment_status, COUNT(*) AS txns,
       ROUND(100.0*COUNT(*)/(SUM(COUNT(*)) OVER ()),2) AS pct_of_all
FROM transactions GROUP BY payment_method, payment_status ORDER BY txns DESC;

-- Q7. Monthly revenue per category (pivot source for Excel)
SELECT substr(t.order_date,1,7) AS month, p.category,
       ROUND(SUM(t.total_price),2) AS revenue
FROM transactions t JOIN products p ON t.product_id = p.product_id
GROUP BY month, p.category ORDER BY month, p.category;

-- Q8. Rolling 3-month revenue per seller (window function)
SELECT store_name, month, revenue,
       SUM(revenue) OVER (PARTITION BY store_name ORDER BY month
                          ROWS BETWEEN 2 PRECEDING AND CURRENT ROW) AS rolling_3m
FROM (SELECT s.store_name, substr(t.order_date,1,7) AS month,
             ROUND(SUM(t.total_price),2) AS revenue
      FROM transactions t JOIN sellers s ON t.seller_id=s.seller_id
      GROUP BY s.store_name, month);
