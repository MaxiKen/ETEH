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

-- A8. Products with a blank (NULL) seller_id  -> 0
SELECT COUNT(*) FROM products WHERE seller_id IS NULL;

-- A9. First order date  -> 2020-01-20
SELECT MIN(order_date) FROM transactions;

-- A10. Rows returned by the WRONG join transactions-deliveries ON customer_id  -> 56725
SELECT COUNT(*) FROM transactions t JOIN deliveries d ON d.customer_id = t.customer_id;

-- Q1. Net revenue in 2025 (Delivered + Completed), ₦.  -> 386719300.0
SELECT SUM(total_price) AS net_revenue_2025
FROM transactions t
WHERE t.order_date LIKE '2025%'
  AND t.order_status IN ('Delivered','Completed');

-- Q2. Net revenue in 2024 (Delivered + Completed), ₦. (Same query, one change.)  -> 312300450.0
SELECT SUM(total_price) AS net_revenue_2024
FROM transactions t
WHERE t.order_date LIKE '2024%'
  AND t.order_status IN ('Delivered','Completed');

-- Q3. Category with the highest 2025 net revenue.  -> Electronics
SELECT p.category, SUM(t.total_price) AS rev
FROM transactions t
JOIN products p ON p.product_id = t.product_id
WHERE t.order_date LIKE '2025%' AND t.order_status IN ('Delivered','Completed')
GROUP BY p.category
ORDER BY rev DESC
LIMIT 1;

-- Q4. Seller (store_name) with the highest 2025 net revenue.  -> Kaduna Merchants
SELECT s.store_name, SUM(t.total_price) AS rev
FROM transactions t
JOIN sellers s ON s.seller_id = t.seller_id
WHERE t.order_date LIKE '2025%' AND t.order_status IN ('Delivered','Completed')
GROUP BY s.store_name
ORDER BY rev DESC
LIMIT 1;

-- Q5. Platform commission earned on 2025 net orders, ₦ (nearest naira).  -> 39108808.0
SELECT ROUND(SUM(commission_amount), 0) AS commission
FROM transactions t
WHERE t.order_date LIKE '2025%' AND t.order_status IN ('Delivered','Completed');

-- Q6. How many 2025 order lines were Returned?  -> 177
SELECT COUNT(*) AS returned_2025
FROM transactions t
WHERE t.order_date LIKE '2025%'
  AND t.order_status = 'Returned';

-- Q7. Payment method with the most Cancelled order lines (all years).  -> Paystack
SELECT payment_method, COUNT(*) AS cancelled
FROM transactions
WHERE order_status = 'Cancelled'
GROUP BY payment_method
ORDER BY cancelled DESC
LIMIT 1;

-- Q8. Customer zone with the highest 2025 net revenue.  -> South-West
SELECT c.zone, SUM(t.total_price) AS rev
FROM transactions t
JOIN customers c ON c.customer_id = t.customer_id
WHERE t.order_date LIKE '2025%' AND t.order_status IN ('Delivered','Completed')
GROUP BY c.zone
ORDER BY rev DESC
LIMIT 1;

-- Q9. How many different customers placed a net order in 2025?  -> 1749
SELECT COUNT(DISTINCT customer_id) AS active_2025
FROM transactions t
WHERE t.order_date LIKE '2025%' AND t.order_status IN ('Delivered','Completed');

-- Q10. Product category with the most products listed (products table).  -> Health
SELECT category, COUNT(*) AS products
FROM products
GROUP BY category
ORDER BY products DESC
LIMIT 1;
