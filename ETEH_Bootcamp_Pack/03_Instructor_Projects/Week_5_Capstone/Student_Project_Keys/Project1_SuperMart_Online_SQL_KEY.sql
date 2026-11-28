-- ETEH Student Project 1 - ANSWER KEY

-- A1. Rows in customers  -> 2500
SELECT COUNT(*) FROM customers;

-- A2. Rows in sellers  -> 200
SELECT COUNT(*) FROM sellers;

-- A3. Rows in products  -> 800
SELECT COUNT(*) FROM products;

-- A4. Rows in deliveries  -> 10430
SELECT COUNT(*) FROM deliveries;

-- A5. Rows in transactions  -> 11000
SELECT COUNT(*) FROM transactions;

-- A6. Transactions with a NULL delivery_id  -> 570
SELECT COUNT(*) FROM transactions WHERE delivery_id IS NULL;

-- A7. ...of which are NOT Cancelled  -> 0
SELECT COUNT(*) FROM transactions WHERE delivery_id IS NULL AND order_status <> 'Cancelled';

-- A8. Products whose seller_id has no match in sellers  -> 0
SELECT COUNT(*) FROM products p LEFT JOIN sellers s ON s.seller_id = p.seller_id WHERE s.seller_id IS NULL;

-- A9. First order date  -> 2020-01-20
SELECT MIN(order_date) FROM transactions;

-- A10. Rows returned by the WRONG join transactions-deliveries ON customer_id  -> 56725
SELECT COUNT(*) FROM transactions t JOIN deliveries d ON d.customer_id = t.customer_id;

-- Q1. Net revenue in 2025 (Delivered + Completed), ₦.  -> 386719300.0
SELECT SUM(total_price) AS net_revenue_2025
FROM transactions t
WHERE strftime('%Y', t.order_date) = '2025'
  AND t.order_status IN ('Delivered','Completed');

-- Q2. Net revenue growth, 2025 vs 2024 (as a decimal, 4 dp).  -> 0.2383
SELECT ROUND(SUM(CASE WHEN strftime('%Y', order_date) = '2025' THEN total_price END) * 1.0 /
       SUM(CASE WHEN strftime('%Y', order_date) = '2024' THEN total_price END) - 1, 4) AS yoy
FROM transactions
WHERE order_status IN ('Delivered','Completed');

-- Q3. Category with the highest 2025 net revenue.  -> Electronics
SELECT p.category, SUM(t.total_price) AS rev
FROM transactions t
JOIN products p ON p.product_id = t.product_id
WHERE strftime('%Y', t.order_date) = '2025' AND t.order_status IN ('Delivered','Completed')
GROUP BY p.category
ORDER BY rev DESC
LIMIT 1;

-- Q4. Seller (store_name) with the highest 2025 net revenue.  -> Kaduna Merchants
SELECT s.store_name, SUM(t.total_price) AS rev
FROM transactions t
JOIN sellers s ON s.seller_id = t.seller_id
WHERE strftime('%Y', t.order_date) = '2025' AND t.order_status IN ('Delivered','Completed')
GROUP BY s.seller_id, s.store_name
ORDER BY rev DESC
LIMIT 1;

-- Q5. Platform commission earned on 2025 net orders, ₦ (nearest naira).  -> 39108808.0
SELECT ROUND(SUM(commission_amount), 0) AS commission
FROM transactions t
WHERE strftime('%Y', t.order_date) = '2025' AND t.order_status IN ('Delivered','Completed');

-- Q6. 2025 return rate = Returned ÷ (Delivered + Completed + Returned), 4 dp.  -> 0.056
SELECT ROUND(SUM(order_status = 'Returned') * 1.0 /
       SUM(order_status IN ('Delivered','Completed','Returned')), 4) AS return_rate
FROM transactions t
WHERE strftime('%Y', t.order_date) = '2025';

-- Q7. Payment method with the highest cancellation rate (all years). Rate = cancelled orders ÷ all orders for that method.  -> USSD
SELECT payment_method, ROUND(AVG(order_status = 'Cancelled'), 4) AS cancel_rate, COUNT(*) AS n
FROM transactions
GROUP BY payment_method
ORDER BY cancel_rate DESC
LIMIT 1;

-- Q8. Customer zone with the SLOWEST average order-to-delivery time for 2025 delivered orders. Days = julianday(d.delivered_date) - julianday(t.order_date); julianday turns a date into a day number.  -> South-West
SELECT c.zone, ROUND(AVG(julianday(d.delivered_date) - julianday(t.order_date)), 2) AS days
FROM transactions t
JOIN deliveries d ON d.delivery_id = t.delivery_id
JOIN customers c ON c.customer_id = t.customer_id
WHERE strftime('%Y', t.order_date) = '2025' AND t.order_status = 'Delivered'
GROUP BY c.zone
ORDER BY days DESC
LIMIT 1;

-- Q9. Share of customers (who ever placed a net order) with TWO or more net orders, 4 dp.  -> 0.9222
SELECT ROUND(AVG(n >= 2), 4) AS repeat_share FROM (
  SELECT customer_id, COUNT(*) AS n
  FROM transactions
  WHERE order_status IN ('Delivered','Completed')
  GROUP BY customer_id
);

-- Q10. How many listed products had NO net sale in 2025? (LEFT JOIN - put the year/status filter in the ON clause)  -> 18
SELECT COUNT(*) AS unsold_2025
FROM products p
LEFT JOIN transactions t
       ON t.product_id = p.product_id
      AND strftime('%Y', t.order_date) = '2025'
      AND t.order_status IN ('Delivered','Completed')
WHERE t.transaction_id IS NULL;
