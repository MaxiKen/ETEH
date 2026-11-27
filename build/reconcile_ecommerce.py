"""
ETEH Bootcamp - Reconcile the e-commerce transactions table.

The source transactions had ~70% of foreign keys left blank and a delivery_id
format mismatch (DEL + 7 digits vs the deliveries table's DEL + 6 digits).
This makes the data un-joinable out of the box.

We deterministically rebuild a FULLY REFERENTIALLY-INTACT transactions table:
  * product_id  -> kept if valid, else sampled from products
  * seller_id   -> always derived from the chosen product (guaranteed in sellers)
  * customer_id -> kept if valid, else sampled from customers
  * delivery_id -> chosen from deliveries belonging to that customer when possible,
                   otherwise a random valid delivery_id (format normalised to 6 digits)
All original transaction-level facts (id, order, qty, price, commission, status,
payment, date) are preserved untouched.
"""
import os, random
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CLEAN = os.path.join(ROOT, "Datasets", "Clean")
random.seed(2024)

def load(name):
    return pd.read_csv(os.path.join(CLEAN, name), dtype=str, keep_default_na=False)

tx  = load("ecommerce_transactions.csv")
cu  = load("ecommerce_customers.csv")
pr  = load("ecommerce_products.csv")
se  = load("ecommerce_sellers.csv")
de  = load("ecommerce_deliveries.csv")

cust_ids  = list(cu["customer_id"])
prod_ids  = list(pr["product_id"])
prod2sell = dict(zip(pr["product_id"], pr["seller_id"]))
valid_prod = set(prod_ids)
valid_cust = set(cust_ids)
valid_del  = set(de["delivery_id"])
# customer -> delivery_ids (for sensible delivery linkage)
cust2del = {}
for d, c in zip(de["delivery_id"], de["customer_id"]):
    cust2del.setdefault(c, []).append(d)

def norm_del(s):
    if not s or not s.startswith("DEL"):
        return None
    digits = s[3:].lstrip("0")
    return "DEL" + digits.zfill(6)

rows = []
n_fixed_cust = n_fixed_prod = n_fixed_del = 0
for _, r in tx.iterrows():
    # product + seller
    pid = r["product_id"] if r["product_id"] in valid_prod else random.choice(prod_ids)
    if r["product_id"] not in valid_prod:
        n_fixed_prod += 1
    sid = prod2sell[pid]
    # customer
    cid = r["customer_id"] if r["customer_id"] in valid_cust else random.choice(cust_ids)
    if r["customer_id"] not in valid_cust:
        n_fixed_cust += 1
    # delivery
    cand = cust2del.get(cid, [])
    if cand:
        did = random.choice(cand)
    else:
        did = random.choice(list(valid_del))
    rows.append({
        "transaction_id": r["transaction_id"],
        "order_id": r["order_id"],
        "customer_id": cid,
        "product_id": pid,
        "seller_id": sid,
        "delivery_id": did,
        "quantity": r["quantity"],
        "unit_price": r["unit_price"],
        "total_price": r["total_price"],
        "commission_percentage": r["commission_percentage"],
        "commission_amount": r["commission_amount"],
        "order_status": r["order_status"],
        "payment_method": r["payment_method"],
        "payment_status": r["payment_status"],
        "order_date": r["order_date"],
    })

out = pd.DataFrame(rows)
out.to_csv(os.path.join(CLEAN, "ecommerce_transactions_reconciled.csv"), index=False)
print(f"Reconciled transactions: {len(out)} rows")
print(f"  filled customer_id: {n_fixed_cust}, product_id: {n_fixed_prod}")
# verify integrity
assert out["customer_id"].isin(valid_cust).all(), "customer FK broken"
assert out["product_id"].isin(valid_prod).all(), "product FK broken"
assert out["seller_id"].isin(set(se["seller_id"])).all(), "seller FK broken"
assert out["delivery_id"].isin(valid_del).all(), "delivery FK broken"
# seller must match product's seller
m = out.merge(pr[["product_id","seller_id"]], on="product_id")
assert (m["seller_id_x"] == m["seller_id_y"]).all(), "seller/product mismatch"
print("Integrity: ALL foreign keys valid; seller_id matches product's seller. OK")
