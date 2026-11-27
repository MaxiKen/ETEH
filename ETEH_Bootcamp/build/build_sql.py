#!/usr/bin/env python3
"""
ETEH Bootcamp - Dataset build pipeline (Part 2 of 3)
The supplied e-commerce workbook has five related tables with foreign keys.
We build a CLEAN SQLite database (eteh_commerce.db) with proper PRIMARY/FOREIGN
keys (referential integrity enforced) plus a deliberately messy DIRTY copy for
students to clean. We also ship analytical SQL (joins/aggregations/windows) and
ready-made extracts for the Pivot / Dashboard modules.
"""
import os, random, sqlite3
import pandas as pd

ROOT = "/home/user/ETEH"
SRC  = os.path.join(ROOT, "e-commerce.xlsx")
OUT  = os.path.join(ROOT, "ETEH_Bootcamp", "datasets")
SQL  = os.path.join(OUT, "sql")
CLEAN= os.path.join(OUT, "clean")
DIRTY= os.path.join(OUT, "dirty")
os.makedirs(SQL, exist_ok=True); os.makedirs(CLEAN, exist_ok=True); os.makedirs(DIRTY, exist_ok=True)
random.seed(99)
DB = os.path.join(SQL, "eteh_commerce.db")

xl = pd.ExcelFile(SRC)
raw = {s: xl.parse(s) for s in xl.sheet_names}

def snake(df):
    df = df.copy()
    df.columns = [c.strip().lower().replace(' ', '_') for c in df.columns]
    return df

customers  = snake(raw['Customers'])
sellers    = snake(raw['Sellers'])
products   = snake(raw['Products'])
deliveries = snake(raw['Deliveries'])
trans_raw  = snake(raw['Transactions'])   # keep raw for the DIRTY copy

dfs = {'customers':customers, 'sellers':sellers, 'products':products,
       'deliveries':deliveries, 'transactions':trans_raw}
date_cols = {'customers':['date_of_birth','created_at'], 'sellers':['created_at'],
             'products':['created_at'], 'deliveries':['shipped_date','delivered_date'],
             'transactions':['order_date']}
for name, cols in date_cols.items():
    d = dfs[name]
    for c in cols:
        d[c] = pd.to_datetime(d[c], errors='coerce').dt.strftime('%Y-%m-%d')

# ---------------------------------------------------------------------------
# Referential cleanup -> CLEAN star schema (enforceable FKs)
# ---------------------------------------------------------------------------
valid_cust = set(customers['customer_id']); valid_sell = set(sellers['seller_id'])
valid_prod = set(products['product_id']);  valid_del  = set(deliveries['delivery_id'])

# dimensions: drop duplicate PK rows
for n in ['customers','sellers','products','deliveries']:
    pk = dfs[n].columns[0]
    dfs[n] = dfs[n].drop_duplicates(subset=[pk]).reset_index(drop=True)

# products must point to a real seller
dfs['products'] = dfs['products'][dfs['products']['seller_id'].isin(valid_sell)].reset_index(drop=True)
valid_prod = set(dfs['products']['product_id'])

# deliveries: rebuild the transaction<->delivery link via order_id
# (raw delivery_id key spaces differ, but order_id matches 1:1)
order_to_del = (dfs['deliveries'].drop_duplicates('order_id')
                .set_index('order_id')['delivery_id'].to_dict())
valid_del = set(dfs['deliveries']['delivery_id'])

# facts: relink delivery, keep only fully referenced lines, assign PK
trans = trans_raw.copy()
trans['delivery_id'] = trans['order_id'].map(order_to_del)
trans = trans[trans['customer_id'].isin(valid_cust) &
              trans['product_id'].isin(valid_prod) &
              trans['seller_id'].isin(valid_sell) &
              trans['delivery_id'].isin(valid_del)].copy()
trans = trans.drop_duplicates(subset=['transaction_id']).reset_index(drop=True)
trans['source_transaction_id'] = trans['transaction_id']
trans['transaction_id'] = ['TRX' + str(i+1).zfill(7) for i in range(len(trans))]
dfs['transactions'] = trans

customers, sellers, products, deliveries, trans = (dfs['customers'], dfs['sellers'],
        dfs['products'], dfs['deliveries'], dfs['transactions'])

# ---------------------------------------------------------------------------
# Analytics-friendly flat fact table (for non-SQL modules)
# ---------------------------------------------------------------------------
flat = (trans
        .merge(customers[['customer_id','first_name','last_name','city','state','membership_tier']],
               on='customer_id', how='left')
        .merge(products[['product_id','product_name','category','subcategory','brand_name']],
               on='product_id', how='left')
        .merge(sellers[['seller_id','store_name','store_state']], on='seller_id', how='left')
        .merge(deliveries[['delivery_id','delivery_status','courier_name']], on='delivery_id', how='left'))
flat = flat[['transaction_id','order_id','order_date','customer_id','first_name','last_name','city','state',
             'membership_tier','product_id','product_name','category','subcategory','brand_name',
             'seller_id','store_name','store_state','quantity','unit_price','total_price',
             'commission_amount','order_status','payment_method','payment_status','delivery_status','courier_name']]

# ---------------------------------------------------------------------------
# Write SQLite database with keys
# ---------------------------------------------------------------------------
if os.path.exists(DB): os.remove(DB)
con = sqlite3.connect(DB)
con.execute("PRAGMA foreign_keys = ON;")

def ddl(name, df, pk, fks=None):
    cols = []
    for c in df.columns:
        t = 'INTEGER' if df[c].dtype == 'int64' else ('REAL' if df[c].dtype == 'float64' else 'TEXT')
        cols.append(f'"{c}" {t}')
    cols.append(f'PRIMARY KEY("{pk}")')
    for fk_col, ref in (fks or {}).items():
        cols.append(f'FOREIGN KEY("{fk_col}") REFERENCES {ref}')
    return f'CREATE TABLE {name} (\n  ' + ',\n  '.join(cols) + '\n);'

con.execute(ddl('customers', customers, 'customer_id'))
con.execute(ddl('sellers', sellers, 'seller_id'))
con.execute(ddl('products', products, 'product_id', {'seller_id':'sellers(seller_id)'}))
con.execute(ddl('deliveries', deliveries, 'delivery_id', {'customer_id':'customers(customer_id)'}))
con.execute(ddl('transactions', trans, 'transaction_id',
                {'customer_id':'customers(customer_id)','product_id':'products(product_id)',
                 'seller_id':'sellers(seller_id)','delivery_id':'deliveries(delivery_id)'}))
for name, df in dfs.items():
    df.to_sql(name, con, if_exists='append', index=False)
con.commit()
orphans = con.execute("""SELECT COUNT(*) FROM transactions t LEFT JOIN customers c
                         ON t.customer_id=c.customer_id WHERE c.customer_id IS NULL""").fetchone()[0]
print("orphan transactions (should be 0):", orphans)
print("row counts:", {n: con.execute(f'SELECT COUNT(*) FROM {n}').fetchone()[0] for n in dfs})
con.close()

# ---------------------------------------------------------------------------
# Analytical SQL script + run sample outputs
# ---------------------------------------------------------------------------
queries = r'''
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
'''
with open(os.path.join(SQL, "queries.sql"), "w") as f:
    f.write(queries)

con = sqlite3.connect(DB)
def save(sql, fname):
    df = pd.read_sql_query(sql, con); df.to_csv(os.path.join(SQL, fname), index=False)
    print("wrote", fname, df.shape)
save("SELECT substr(order_date,1,7) AS month, COUNT(DISTINCT order_id) AS orders, SUM(quantity) AS units, ROUND(SUM(total_price),2) AS revenue FROM transactions GROUP BY month ORDER BY month", "out_monthly_revenue.csv")
save("""SELECT p.category, p.subcategory, COUNT(*) lines, SUM(t.quantity) units, ROUND(SUM(t.total_price),2) revenue
        FROM transactions t JOIN products p ON t.product_id=p.product_id
        GROUP BY p.category, p.subcategory ORDER BY revenue DESC""", "out_category_sales.csv")
save("""SELECT s.store_name, s.store_state, ROUND(SUM(t.total_price),2) revenue, COUNT(*) lines
        FROM transactions t JOIN sellers s ON t.seller_id=s.seller_id GROUP BY s.seller_id
        ORDER BY revenue DESC LIMIT 10""", "out_top_sellers.csv")
con.close()

# ---------------------------------------------------------------------------
# Exports: clean table CSVs + combined workbook + flat extract + pivot source
# ---------------------------------------------------------------------------
for name, df in dfs.items():
    df.to_csv(os.path.join(CLEAN, f"eteh_commerce_{name}.csv"), index=False)
with pd.ExcelWriter(os.path.join(CLEAN, "eteh_commerce.xlsx"), engine='xlsxwriter') as xw:
    for name, df in dfs.items():
        df.to_excel(xw, sheet_name=name, index=False)
flat.to_csv(os.path.join(CLEAN, "eteh_commerce_flat.csv"), index=False)
flat.to_excel(os.path.join(CLEAN, "eteh_commerce_flat.xlsx"), engine='xlsxwriter', index=False)
pivot_src = pd.read_sql_query("""
    SELECT substr(t.order_date,1,7) AS month, p.category,
           ROUND(SUM(t.total_price),2) AS revenue, SUM(t.quantity) AS units
    FROM transactions t JOIN products p ON t.product_id=p.product_id
    GROUP BY month, p.category""", sqlite3.connect(DB))
pivot_src.to_csv(os.path.join(CLEAN, "eteh_commerce_pivot_source.csv"), index=False)

# ---------------------------------------------------------------------------
# DIRTY commerce copy = the raw, messy original for students to clean
# (duplicate transaction_ids, orphan delivery keys, mixed cases, text money)
# ---------------------------------------------------------------------------
dtrans = trans_raw.copy()
dtrans['customer_id'] = dtrans['customer_id'].apply(lambda v: v if (random.random()<0.995 or not isinstance(v,str)) else 'CUST99999')
dtrans['order_status'] = dtrans['order_status'].apply(lambda v: v.lower() if (isinstance(v,str) and random.random()<0.10) else v)
dtrans['commission_amount'] = dtrans['commission_amount'].apply(lambda v: f"NGN {v:,.2f}" if (pd.notna(v) and random.random()<0.20) else v)
dtrans['order_date'] = dtrans['order_date'].apply(lambda v: pd.to_datetime(v).strftime('%d/%m/%Y') if (pd.notna(v) and random.random()<0.5) else v)
dtrans = pd.concat([dtrans, dtrans.iloc[:40]], ignore_index=True)
dtrans.to_csv(os.path.join(DIRTY, "eteh_commerce_transactions.csv"), index=False)

dprod = products.copy()
dprod['category'] = dprod['category'].apply(lambda v: v.lower() if random.random()<0.12 else v)
dprod['subcategory'] = dprod['subcategory'].apply(lambda v: f" {v}" if random.random()<0.10 else v)
dprod.to_csv(os.path.join(DIRTY, "eteh_commerce_products.csv"), index=False)
print("DONE part 2 - commerce SQL + extracts")
