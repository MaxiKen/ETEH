-- ETEH Instructor Capstone - E-Commerce & Retail Analytics
-- Generated 2026-11-27

-- Revenue by Category
SELECT p.category, ROUND(SUM(t.total_price),2) revenue, COUNT(*) orders
            FROM transactions t JOIN products p ON p.product_id=t.product_id
            GROUP BY p.category ORDER BY revenue DESC;

-- Top 10 Sellers
SELECT s.store_name, s.store_state, ROUND(SUM(t.total_price),2) revenue, COUNT(*) txns
            FROM transactions t JOIN sellers s ON s.seller_id=t.seller_id
            GROUP BY s.seller_id ORDER BY revenue DESC LIMIT 10;

-- Customer LTV (Top 15)
SELECT c.customer_id, c.first_name||' '||c.last_name customer, c.membership_tier,
                   ROUND(SUM(t.total_price),2) lifetime_spend, COUNT(*) orders
            FROM customers c JOIN transactions t ON t.customer_id=c.customer_id
            GROUP BY c.customer_id ORDER BY lifetime_spend DESC LIMIT 15;

-- Monthly Revenue Trend
SELECT substr(order_date,1,7) month, ROUND(SUM(total_price),2) revenue,
                   ROUND(SUM(commission_amount),2) commission
            FROM transactions GROUP BY month ORDER BY month;

-- Delivery Performance
SELECT courier_name, ROUND(AVG(julianday(delivered_date)-julianday(shipped_date)),2) avg_days, COUNT(*) parcels
            FROM deliveries WHERE delivered_date IS NOT NULL
            GROUP BY courier_name ORDER BY avg_days;

-- Payment Mix
SELECT payment_method, COUNT(*) txns,
                   ROUND(100.0*COUNT(*)/SUM(COUNT(*)) OVER (),1) pct
            FROM transactions GROUP BY payment_method ORDER BY txns DESC;

