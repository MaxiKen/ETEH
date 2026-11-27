"""Stage 1 - standardise, localise and (deliberately) dirty the supplied datasets.

Outputs
  ETEH_Bootcamp_Pack/01_Section_Examples/00_Core_Datasets/   clean master CSVs + data dictionary
  ETEH_Bootcamp_Pack/01_Section_Examples/Week2_Data_Files/    import-practice files (csv/txt/pdf/html/json/xlsx/folder)
  ETEH_Bootcamp_Pack/01_Section_Examples/Week3_Data_Files/    SQLite databases + SQL scripts
  _build/_stage/*.pkl                                          cached frames for later build stages
"""
from __future__ import annotations

import json
import shutil
import sqlite3
import sys

import numpy as np
import pandas as pd

sys.path.insert(0, str(__import__("pathlib").Path(__file__).parent))
from common import (CITY_STATE, CORE, FEMALE_FIRST, MALE_FIRST, MIDDLE_INITIALS, RAW, SEED, STAGE,
                    STATE_ZONE, SURNAMES, W2D, W3D, h)

rng = np.random.default_rng(SEED)
LOG: list[dict] = []          # dataset catalogue rows
DICT: list[dict] = []         # data dictionary rows


def note(dataset, column, dtype, description, example=""):
    DICT.append(dict(Dataset=dataset, Column=column, Type=dtype, Description=description, Example=str(example)))


def stage(name, df):
    df.to_pickle(STAGE / f"{name}.pkl")


class NamePool:
    """Injective, deterministic mapping from an original key to a Nigerian name."""

    def __init__(self, salt):
        self.salt = salt
        self.used: set[str] = set()
        self.map: dict = {}

    def get(self, key, gender="Female", initial=False):
        if key in self.map:
            return self.map[key]
        firsts = FEMALE_FIRST if str(gender).lower().startswith("f") else MALE_FIRST
        k = 0
        while True:
            seed = f"{self.salt}|{key}|{k}"
            first = firsts[h(seed + "f", len(firsts))]
            last = SURNAMES[h(seed + "l", len(SURNAMES))]
            name = f"{first} {MIDDLE_INITIALS[h(seed + 'm', len(MIDDLE_INITIALS))]}. {last}" if initial else f"{first} {last}"
            if name not in self.used:
                break
            k += 1
        self.used.add(name)
        self.map[key] = name
        return name


# ======================================================================================
# 1. SUPERMART STORES  (source: Sample - Superstore.csv)  -> Nigerian 12-store chain, ₦
# ======================================================================================
def build_supermart_stores():
    s = pd.read_csv(RAW / "Sample - Superstore.csv", encoding="latin-1")
    zone_map = {"West": "South-West", "East": "South-South", "Central": "North-Central", "South": "South-East"}
    zone_stores = {
        "South-West": [("ST01", "Lagos - Ikeja", "Ikeja", "Lagos"), ("ST02", "Lagos - Lekki", "Lekki", "Lagos"),
                       ("ST03", "Ibadan - Ring Road", "Ibadan", "Oyo")],
        "South-South": [("ST04", "Port Harcourt - GRA", "Port Harcourt", "Rivers"),
                        ("ST05", "Benin City - Ring Road", "Benin City", "Edo"), ("ST06", "Uyo - Ewet", "Uyo", "Akwa Ibom")],
        "North-Central": [("ST07", "Abuja - Wuse", "Abuja", "FCT"), ("ST08", "Abuja - Garki", "Abuja", "FCT"),
                          ("ST09", "Jos - Rayfield", "Jos", "Plateau")],
        "South-East": [("ST10", "Enugu - Independence Layout", "Enugu", "Enugu"),
                       ("ST11", "Onitsha - Main Market", "Onitsha", "Anambra"), ("ST12", "Owerri - New Owerri", "Owerri", "Imo")],
    }
    s["Zone"] = s["Region"].map(zone_map)
    # greedy-balance US states onto the three stores of each zone
    state_store = {}
    for reg, zone in zone_map.items():
        counts = s[s.Region == reg].groupby("State").size().sort_values(ascending=False)
        load = {st[0]: 0 for st in zone_stores[zone]}
        for state, n in counts.items():
            target = min(load, key=load.get)
            state_store[(reg, state)] = target
            load[target] += n
    store_info = {st[0]: st for z in zone_stores.values() for st in z}
    s["Store_ID"] = [state_store[(r, st)] for r, st in zip(s.Region, s.State)]
    s["Store_Name"] = s.Store_ID.map(lambda x: store_info[x][1])
    s["City"] = s.Store_ID.map(lambda x: store_info[x][2])
    s["State"] = s.Store_ID.map(lambda x: store_info[x][3])

    # customers -> Nigerian names
    pool = NamePool("smc")
    cust = s[["Customer ID", "Customer Name"]].drop_duplicates("Customer ID").sort_values("Customer ID").reset_index(drop=True)
    cust["Customer_ID"] = [f"SMC-{i:04d}" for i in range(1, len(cust) + 1)]
    cust["Customer_Name"] = [pool.get(cid, "Female" if h(cid, 2) else "Male") for cid in cust["Customer ID"]]
    s = s.merge(cust[["Customer ID", "Customer_ID", "Customer_Name"]], on="Customer ID")

    od = pd.to_datetime(s["Order Date"], format="%m/%d/%Y") + pd.DateOffset(years=8)
    sd = pd.to_datetime(s["Ship Date"], format="%m/%d/%Y") + pd.DateOffset(years=8)
    s["Order_ID"] = [f"SM-{d.year}-{oid.split('-')[-1]}" for oid, d in zip(s["Order ID"], od)]
    out = pd.DataFrame({
        "Row_ID": s["Row ID"],
        "Order_ID": s["Order_ID"],
        "Order_Date": od.dt.date,
        "Ship_Date": sd.dt.date,
        "Ship_Mode": s["Ship Mode"],
        "Customer_ID": s["Customer_ID"],
        "Customer_Name": s["Customer_Name"],
        "Segment": s["Segment"],
        "Store_ID": s["Store_ID"],
        "Store_Name": s["Store_Name"],
        "City": s["City"],
        "State": s["State"],
        "Zone": s["Zone"],
        "Product_ID": s["Product ID"],
        "Category": s["Category"],
        "Sub_Category": s["Sub-Category"],
        "Product_Name": s["Product Name"].str.replace("\u00a0", " ").str.strip(),
        "Sales": (s["Sales"] * 1000).round(0).astype(int),
        "Quantity": s["Quantity"],
        "Discount": s["Discount"],
        "Profit": (s["Profit"] * 1000).round(0).astype(int),
    }).sort_values(["Order_Date", "Row_ID"]).reset_index(drop=True)
    out["Row_ID"] = np.arange(1, len(out) + 1)

    stores = pd.DataFrame([dict(Store_ID=k, Store_Name=v[1], City=v[2], State=v[3], Zone=STATE_ZONE[v[3]]) for k, v in sorted(store_info.items())])
    stores["Zone"] = [z for z in stores.State.map(lambda st: [zz for zz, lst in zone_stores.items() if any(x[3] == st for x in lst)][0])]
    stores["Opened"] = pd.to_datetime(["2015-03-02", "2017-08-14", "2016-05-09", "2015-11-23", "2018-02-05", "2019-06-17",
                                       "2016-01-11", "2018-09-03", "2020-04-20", "2017-03-13", "2019-10-07", "2021-02-01"]).date
    stores["Floor_Area_sqm"] = [2400, 1850, 1600, 2100, 1400, 1150, 2200, 1500, 1050, 1700, 1300, 1100]
    stores["Store_Manager"] = ["Folake Adeyemi", "Tunde Bakare", "Kemi Oladipo", "Ebere Amadi", "Osagie Igbinedion",
                               "Uduak Etim", "Hauwa Bello", "Ibrahim Garba", "Ladi Danjuma", "Chinedu Okafor",
                               "Nkechi Eze", "Ikenna Nwosu"]
    stores["Annual_Target_2025"] = 0
    # Targets were "set by head office in Dec-2024"; designed so FY2025 attainment spreads 81%-112%
    attain = {"ST01": 1.03, "ST02": 1.12, "ST03": 0.88, "ST04": 0.97, "ST05": 1.05, "ST06": 0.92,
              "ST07": 0.84, "ST08": 0.95, "ST09": 1.08, "ST10": 1.01, "ST11": 0.90, "ST12": 0.81}
    act25 = out[pd.to_datetime(out.Order_Date).dt.year == 2025].groupby("Store_ID").Sales.sum()
    tgt = act25 / pd.Series(attain)
    stores["Annual_Target_2025"] = stores.Store_ID.map((tgt / 1e6).round(0) * 1e6).astype(int)

    out.to_csv(CORE / "SuperMart_Stores_Sales_2022_2025.csv", index=False)
    stores.to_csv(CORE / "SuperMart_Stores_Directory.csv", index=False)
    stage("sm_sales", out)
    stage("sm_stores", stores)
    for c, t, d, e in [
        ("Row_ID", "Integer", "Line number, unique per sales line", 1),
        ("Order_ID", "Text", "Receipt/order number SM-<year>-<serial>; one order can have many lines", "SM-2022-152156"),
        ("Order_Date", "Date (ISO)", "Date the customer ordered", "2025-11-08"),
        ("Ship_Date", "Date (ISO)", "Date goods left the store/warehouse", "2025-11-11"),
        ("Ship_Mode", "Text", "Same Day, First Class, Second Class or Standard Class", "Second Class"),
        ("Customer_ID", "Text", "Loyalty customer key", "SMC-0144"),
        ("Customer_Name", "Text", "Customer full name (PII - teaching data)", "Chioma Okafor"),
        ("Segment", "Text", "Consumer, Corporate or Home Office", "Consumer"),
        ("Store_ID", "Text", "Store key - joins to SuperMart_Stores_Directory", "ST01"),
        ("Store_Name", "Text", "Branch name", "Lagos - Ikeja"),
        ("City / State / Zone", "Text", "Store location and geopolitical zone", "Ikeja / Lagos / South-West"),
        ("Product_ID", "Text", "Product SKU code <CAT>-<SUB>-<serial>", "FUR-BO-10001798"),
        ("Category / Sub_Category", "Text", "Furniture, Office Supplies, Technology and 17 sub-categories", "Furniture / Bookcases"),
        ("Product_Name", "Text", "Product description", "Bush Somerset Collection Bookcase"),
        ("Sales", "Currency ₦", "Net line revenue after discount", 261960),
        ("Quantity", "Integer", "Units sold on the line", 2),
        ("Discount", "Decimal 0-0.8", "Discount rate applied to list price", 0.2),
        ("Profit", "Currency ₦", "Line profit after cost of goods and discount (can be negative)", 41914),
    ]:
        note("SuperMart_Stores_Sales_2022_2025", c, t, d, e)
    LOG.append(dict(Dataset="SuperMart_Stores_Sales_2022_2025.csv", Organisation="SuperMart Retail (12 stores)",
                    Rows=len(out), Columns=out.shape[1], Grain="One sales line (order x product)",
                    Period=f"{out.Order_Date.min()} to {out.Order_Date.max()}",
                    Main_Use="Formulas, PivotTables, PivotCharts, forecasting, dashboards, capstone 1",
                    Built_From="Sample - Superstore.csv (re-mapped to 12 Nigerian stores, dates +8 years, values x1,000 to ₦, Nigerian customer names)"))
    return out, stores


# ======================================================================================
# 2. SUPERMART ONLINE marketplace (source: e-commerce.xlsx) -> clean relational DB
# ======================================================================================
CATALOG = {  # product -> (category, subcategory, brand, low ₦, high ₦)
    "Air Freshener": ("Automotive", "Car Care", "Ambi Pur", 2500, 6000),
    "Brake Pads": ("Automotive", "Spare Parts", "Bosch", 18000, 45000),
    "Car Battery": ("Automotive", "Spare Parts", "Exide", 65000, 140000),
    "Car Cover": ("Automotive", "Accessories", "AutoGuard", 15000, 35000),
    "Car Vacuum": ("Automotive", "Tools", "Black+Decker", 25000, 55000),
    "Dashboard Polish": ("Automotive", "Car Care", "Turtle Wax", 3500, 8000),
    "Engine Oil 4L": ("Automotive", "Spare Parts", "TotalEnergies", 22000, 38000),
    "Jump Cables": ("Automotive", "Tools", "Bosch", 9000, 20000),
    "Seat Covers": ("Automotive", "Accessories", "AutoGuard", 20000, 60000),
    "Tyre Inflator": ("Automotive", "Tools", "Michelin", 25000, 60000),
    "Black Soap": ("Beauty", "Personal Care", "Dudu-Osun", 1200, 3500),
    "Body Lotion": ("Beauty", "Skincare", "Nivea", 4500, 12000),
    "Facial Cleanser": ("Beauty", "Skincare", "Garnier", 6000, 15000),
    "Hair Attachment": ("Beauty", "Hair", "X-pression", 3000, 9000),
    "Hair Cream": ("Beauty", "Hair", "Dark and Lovely", 3000, 7500),
    "Hair Gel": ("Beauty", "Hair", "Eco Styler", 2500, 6000),
    "Lip Balm": ("Beauty", "Makeup", "Vaseline", 1500, 4000),
    "Nail Polish": ("Beauty", "Makeup", "Sleek", 1500, 4500),
    "Perfume Oil": ("Beauty", "Fragrances", "Arabian Oud", 5000, 18000),
    "Shea Butter Cream": ("Beauty", "Skincare", "Nigerian Pride", 3000, 8000),
    "Bluetooth Speaker": ("Electronics", "Audio", "JBL", 15000, 75000),
    "Dell Inspiron": ("Electronics", "Laptops", "Dell", 450000, 950000),
    "HP Laptop 15": ("Electronics", "Laptops", "HP", 420000, 900000),
    "Infinix Note 30": ("Electronics", "Phones", "Infinix", 210000, 290000),
    "Nokia C30": ("Electronics", "Phones", "Nokia", 95000, 140000),
    "Power Bank 20000mAh": ("Electronics", "Accessories", "Oraimo", 15000, 32000),
    "Samsung A14": ("Electronics", "Phones", "Samsung", 180000, 250000),
    "Tecno Spark 10": ("Electronics", "Phones", "Tecno", 130000, 180000),
    "Wireless Earbuds": ("Electronics", "Audio", "Oraimo", 12000, 45000),
    "iPhone 12": ("Electronics", "Phones", "Apple", 380000, 560000),
    "Ankara Dress": ("Fashion", "Women Clothing", "Ankara House", 15000, 45000),
    "Designer Asoebi": ("Fashion", "Women Clothing", "Lagos Couture", 35000, 120000),
    "Handbag": ("Fashion", "Bags", "Fashion House", 12000, 60000),
    "Jeans Trousers": ("Fashion", "Men Clothing", "Levi's", 12000, 35000),
    "Lace Material": ("Fashion", "Fabrics", "Swiss Lace", 25000, 90000),
    "Leather Shoes": ("Fashion", "Shoes", "Aba Made", 18000, 55000),
    "Polo Shirt": ("Fashion", "Men Clothing", "Polo Club", 7000, 20000),
    "Senator Wear": ("Fashion", "Men Clothing", "Senator Couture", 25000, 70000),
    "Traditional Cap": ("Fashion", "Accessories", "Fila Weavers", 5000, 15000),
    "Wristwatch": ("Fashion", "Jewelry", "Casio", 15000, 85000),
    "Basmati Rice 50kg": ("Food", "Rice", "Royal Stallion", 140000, 200000),
    "Golden Morn 1kg": ("Food", "Beverages", "Nestle", 4500, 6500),
    "Groundnut Oil 3L": ("Food", "Vegetable Oil", "Kings", 9000, 14000),
    "Honey Beans 10kg": ("Food", "Beans", "Olam", 22000, 32000),
    "Indomie Noodles Carton": ("Food", "Snacks", "Dufil", 9500, 14000),
    "Knorr Cubes Pack": ("Food", "Spices", "Unilever", 2000, 4000),
    "Milo 1kg": ("Food", "Beverages", "Nestle", 8000, 12500),
    "Palm Oil 5L": ("Food", "Vegetable Oil", "Ola-Ola", 9000, 15000),
    "Peak Milk 400g": ("Food", "Beverages", "FrieslandCampina", 4000, 6500),
    "Yellow Garri 5kg": ("Food", "Garri", "Nigerian Pride", 4500, 8000),
    "Blood Pressure Monitor": ("Health", "Medical Supplies", "Omron", 25000, 55000),
    "Blood Tonic": ("Health", "Supplements", "Emzor", 2000, 5500),
    "Digital Thermometer": ("Health", "Medical Supplies", "Omron", 3500, 9000),
    "Face Masks Pack": ("Health", "Medical Supplies", "Medisure", 2500, 6000),
    "First Aid Kit": ("Health", "Medical Supplies", "Medisure", 8000, 25000),
    "Hand Sanitizer": ("Health", "Wellness", "Dettol", 1500, 4500),
    "Multivitamins": ("Health", "Supplements", "Centrum", 6000, 18000),
    "Pain Relief Balm": ("Health", "Wellness", "Robb", 800, 2500),
    "Paracetamol Pack": ("Health", "Medical Supplies", "Emzor", 600, 2000),
    "Vitamin C": ("Health", "Supplements", "Emzor", 1500, 6000),
    "Bed Sheet Set": ("Home", "Bedding", "Vono", 12000, 35000),
    "Center Rug": ("Home", "Decor", "HomeStyle", 25000, 85000),
    "Dining Table": ("Home", "Furniture", "HomeStyle", 180000, 650000),
    "Gas Cooker": ("Home", "Appliances", "Scanfrost", 120000, 380000),
    "Kitchen Utensils Set": ("Home", "Kitchen", "Royalty Line", 15000, 45000),
    "LED Bulb Pack": ("Home", "Decor", "Philips", 4000, 12000),
    "Microwave Oven": ("Home", "Appliances", "LG", 85000, 180000),
    "Plastic Chairs Set": ("Home", "Furniture", "Ultimate Plastics", 25000, 60000),
    "Standing Fan": ("Home", "Appliances", "Binatone", 35000, 75000),
    "Wardrobe": ("Home", "Furniture", "HomeStyle", 150000, 480000),
}
FEMALE_SET = {"Adaeze", "Aisha", "Amara", "Amina", "Blessing", "Chiamaka", "Chidinma", "Chioma", "Fatima", "Folake",
              "Funmilayo", "Hadiza", "Halima", "Ngozi", "Nkiru", "Oluchi", "Rashida", "Safiya", "Titilayo", "Zainab"}
MALE_SET = {"Abdul", "Abubakar", "Babatunde", "Chinedu", "Chukwuemeka", "Emmanuel", "Femi", "Gbenga", "Ibrahim",
            "Ifeanyi", "Kelechi", "Musa", "Mustapha", "Obinna", "Okey", "Olumide", "Sani", "Segun", "Tochukwu",
            "Tunde", "Usman", "Yusuf"}


def r50(x):
    return (np.round(np.asarray(x, dtype=float) / 50.0) * 50.0)


def build_supermart_online():
    x = pd.read_excel(RAW / "e-commerce.xlsx", sheet_name=None)
    t = x["Transactions"].dropna(subset=["transaction_id"]).copy()
    d = x["Deliveries"].copy()
    c = x["Customers"].copy()
    p = x["Products"].copy()
    s = x["Sellers"].copy()

    # ---- products: fix name, category, subcategory, brand and realistic ₦ prices
    p["product_name"] = p["product_name"].str.strip().replace({"Techno Spark 10": "Tecno Spark 10"})
    p["_pct"] = p.groupby("product_name")["price"].rank(pct=True, method="first")
    p["_disc_ratio"] = (p["discount_price"] / p["price"]).clip(0.5, 1)
    cat = p["product_name"].map(CATALOG)
    assert cat.notna().all(), p.loc[cat.isna(), "product_name"].unique()
    p["category"] = cat.str[0]
    p["subcategory"] = cat.str[1]
    p["brand_name"] = cat.str[2]
    lo, hi = cat.str[3].astype(float), cat.str[4].astype(float)
    p["price"] = r50(lo + (hi - lo) * (p["_pct"] - p.groupby("product_name")["_pct"].transform("min")) /
                     (1 - p.groupby("product_name")["_pct"].transform("min")).replace(0, 1))
    p["discount_price"] = r50(p["price"] * p["_disc_ratio"])
    p["description"] = "High quality " + p["product_name"].str.lower() + " available for immediate delivery nationwide"
    p = p.drop(columns=["_pct", "_disc_ratio"])

    # ---- transactions: reprice, recompute derived money fields, repair delivery key
    t = t.merge(p[["product_id", "discount_price"]], on="product_id", how="left")
    t["quantity"] = t["quantity"].astype(int)
    t["unit_price"] = t["discount_price"]
    t["total_price"] = (t["quantity"] * t["unit_price"]).round(2)
    t["commission_amount"] = (t["total_price"] * t["commission_percentage"] / 100).round(2)
    t = t.drop(columns=["discount_price"])
    t = t.drop(columns=["delivery_id"]).merge(d[["order_id", "delivery_id"]], on="order_id", how="left")
    t["order_date"] = pd.to_datetime(t["order_date"]).dt.date
    cols = ["transaction_id", "order_id", "customer_id", "product_id", "seller_id", "delivery_id", "quantity",
            "unit_price", "total_price", "commission_percentage", "commission_amount", "order_status",
            "payment_method", "payment_status", "order_date"]
    t = t[cols].sort_values("transaction_id").reset_index(drop=True)

    # ---- geography: repair impossible city/state pairs
    for df, ccol, scol in [(c, "city", "state"), (d, "city", "state"), (s, "store_city", "store_state")]:
        df[scol] = df[ccol].map(CITY_STATE)
    c["zone"] = c["state"].map(STATE_ZONE)
    d["zone"] = d["state"].map(STATE_ZONE)
    s["store_zone"] = s["store_state"].map(STATE_ZONE)

    # ---- customers: gender consistent with first name, phone as text, recomputed totals
    c["gender"] = np.where(c.first_name.isin(FEMALE_SET), "Female", np.where(c.first_name.isin(MALE_SET), "Male", c.gender))
    c["phone_number"] = "0" + c["phone_number"].astype(str)
    s["phone_number"] = "0" + s["phone_number"].astype(str)
    agg = t.groupby("customer_id").agg(total_orders=("order_id", "count"), total_spent=("total_price", "sum"))
    c = c.drop(columns=["total_orders", "total_spent"]).merge(agg, on="customer_id", how="left")
    c["total_orders"] = c["total_orders"].fillna(0).astype(int)
    c["total_spent"] = c["total_spent"].fillna(0).round(2)
    sagg = t.groupby("seller_id").agg(total_sales=("order_id", "count"), total_revenue=("total_price", "sum"))
    s = s.drop(columns=["total_sales", "total_revenue"]).merge(sagg, on="seller_id", how="left")
    s["total_sales"] = s["total_sales"].fillna(0).astype(int)
    s["total_revenue"] = s["total_revenue"].fillna(0).round(2)
    psold = t.groupby("product_id").quantity.sum()
    p["total_sold"] = p.product_id.map(psold).fillna(0).astype(int)
    for df in (c, s, p):
        df["created_at"] = pd.to_datetime(df["created_at"]).dt.date
    c["date_of_birth"] = pd.to_datetime(c["date_of_birth"]).dt.date
    for col in ("shipped_date", "delivered_date"):
        d[col] = pd.to_datetime(d[col]).dt.date

    c = c[["customer_id", "first_name", "last_name", "email", "phone_number", "gender", "date_of_birth", "city",
           "state", "zone", "country", "membership_tier", "loyalty_points", "total_orders", "total_spent", "created_at"]]
    s = s[["seller_id", "store_name", "owner_name", "email", "phone_number", "store_city", "store_state", "store_zone",
           "store_country", "rating", "total_sales", "total_revenue", "account_status", "created_at"]]
    d = d[["delivery_id", "order_id", "customer_id", "delivery_address", "city", "state", "zone", "country",
           "courier_name", "tracking_number", "delivery_status", "shipped_date", "delivered_date"]]
    p = p[["product_id", "product_name", "description", "category", "subcategory", "brand_name", "seller_id", "price",
           "discount_price", "commission_percentage", "stock_quantity", "weight", "total_sold", "average_rating", "created_at"]]

    tables = {"customers": c, "sellers": s, "products": p, "deliveries": d, "transactions": t}
    for k, df in tables.items():
        df.to_csv(CORE / f"SuperMart_Online_{k}.csv", index=False)
        stage(f"smo_{k}", df)
    LOG.append(dict(Dataset="SuperMart_Online_<table>.csv (5 tables) + supermart_online.db", Organisation="SuperMart Online marketplace",
                    Rows=sum(len(v) for v in tables.values()), Columns=sum(v.shape[1] for v in tables.values()),
                    Grain="customers, sellers, products, deliveries, transactions (1 line per order)",
                    Period=f"{t.order_date.min()} to {t.order_date.max()}",
                    Main_Use="SQL, joins, Power Query merges, Data Model / DAX, student project 1",
                    Built_From="e-commerce.xlsx (26,000 blank padding rows removed, delivery key repaired, city->state repaired, catalogue & ₦ prices standardised)"))
    for tname, cols_ in {
        "customers": "customer_id PK; names; email; phone (text, leading 0 kept); gender; date_of_birth; city/state/zone; membership_tier; loyalty_points; total_orders & total_spent (recomputed from transactions)",
        "sellers": "seller_id PK; store_name; owner; contact; store_city/state/zone; rating; total_sales & total_revenue (recomputed); account_status",
        "products": "product_id PK; product_name; category; subcategory; brand_name; seller_id FK; price & discount_price (₦); commission_percentage; stock; weight (kg); total_sold; average_rating",
        "deliveries": "delivery_id PK; order_id (1:1 with transactions); customer_id FK; address; courier; tracking; delivery_status; shipped/delivered dates (blank while Pending)",
        "transactions": "transaction_id PK; order_id; customer_id FK; product_id FK; seller_id FK; delivery_id FK (NULL for cancelled orders); quantity; unit_price; total_price = qty x unit; commission; order/payment status; order_date",
    }.items():
        note(f"SuperMart_Online_{tname}", "(table)", "SQL table", cols_)

    build_sqlite_supermart(tables)
    build_supermart_online_raw_export(tables)
    return tables


def _sql_literal(v):
    if v is None or (isinstance(v, float) and np.isnan(v)) or (v is pd.NaT):
        return "NULL"
    if isinstance(v, (int, np.integer)):
        return str(int(v))
    if isinstance(v, (float, np.floating)):
        return repr(round(float(v), 2))
    return "'" + str(v).replace("'", "''") + "'"


def build_sqlite_supermart(tables):
    schema = """-- =====================================================================
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
"""
    (W3D / "supermart_online_01_schema.sql").write_text(schema, encoding="utf-8")
    states = pd.DataFrame(sorted(STATE_ZONE.items()), columns=["state", "zone"])
    order = [("nigeria_states", states)] + [(k, tables[k]) for k in ["customers", "sellers", "products", "deliveries", "transactions"]]
    lines = ["-- SuperMart Online marketplace - data load (run AFTER supermart_online_01_schema.sql)",
             "PRAGMA foreign_keys = ON;", "BEGIN TRANSACTION;"]
    for name, df in order:
        cols = ", ".join(df.columns)
        lines.append(f"\n-- {name}: {len(df):,} rows")
        recs = df.astype(object).where(df.notna(), None).values.tolist()
        for i in range(0, len(recs), 500):
            chunk = recs[i:i + 500]
            vals = ",\n  ".join("(" + ", ".join(_sql_literal(v) for v in r) + ")" for r in chunk)
            lines.append(f"INSERT INTO {name} ({cols}) VALUES\n  {vals};")
    lines.append("COMMIT;")
    (W3D / "supermart_online_02_data.sql").write_text("\n".join(lines), encoding="utf-8")

    dbp = W3D / "supermart_online.db"
    if dbp.exists():
        dbp.unlink()
    con = sqlite3.connect(dbp)
    con.executescript(schema)
    con.executescript((W3D / "supermart_online_02_data.sql").read_text(encoding="utf-8"))
    con.commit()
    # integrity verification
    chk = con.execute("PRAGMA foreign_key_check").fetchall()
    assert not chk, chk[:5]
    n = con.execute("SELECT COUNT(*) FROM v_order_lines").fetchone()[0]
    assert n == len(tables["transactions"]), n
    con.execute("VACUUM")
    con.close()


def build_supermart_online_raw_export(tables):
    """Re-create the messy 'as exported' workbook from the clean tables (for Power Query practice)."""
    t = tables["transactions"].copy()
    # broken delivery key format + 26k-style padding rows (scaled down) + text ids
    t["delivery_id"] = t["delivery_id"].fillna("").map(lambda v: ("DEL0" + v[3:]) if v else "")
    pad = pd.DataFrame({c: [np.nan] * 400 for c in t.columns})
    t_raw = pd.concat([t, pad], ignore_index=True)
    c = tables["customers"].copy()
    wrong = rng.choice(len(c), 180, replace=False)
    states = sorted(set(CITY_STATE.values()))
    c.loc[wrong, "state"] = [states[h(i, len(states))] for i in wrong]
    c["phone_number"] = c["phone_number"].str.lstrip("0").astype("int64")
    p = tables["products"].copy()
    p.loc[p.product_name == "Polo Shirt", "product_name"] = " Polo Shirt"
    p.loc[p.product_name == "Tecno Spark 10", "product_name"] = "Techno Spark 10"
    with pd.ExcelWriter(W2D / "SuperMart_Online_RAW_Export.xlsx", engine="xlsxwriter") as xw:
        t_raw.to_excel(xw, sheet_name="Transactions", index=False)
        tables["sellers"].to_excel(xw, sheet_name="Sellers", index=False)
        p.to_excel(xw, sheet_name="Products", index=False)
        tables["deliveries"].to_excel(xw, sheet_name="Deliveries", index=False)
        c.drop(columns=["zone"]).to_excel(xw, sheet_name="Customers", index=False)


# ======================================================================================
# 3. LIFECARE GENERAL HOSPITAL (source: healthcare_dataset.csv)
# ======================================================================================
DEPT = {"Cancer": "Oncology", "Diabetes": "Endocrinology", "Obesity": "Endocrinology", "Asthma": "Respiratory",
        "Hypertension": "Cardiology", "Arthritis": "Orthopaedics & Rheumatology"}
HMO = {"Cigna": "AXA Mansard Health", "Medicare": "NHIA (National Health Insurance)", "UnitedHealthcare": "Hygeia HMO",
       "Blue Cross": "Reliance HMO", "Aetna": "Leadway Health"}
WARDS = [("W1", "Emerald Ward", 101, 200, 130), ("W2", "Sapphire Ward", 201, 300, 130),
         ("W3", "Ruby Ward", 301, 400, 130), ("W4", "Topaz Ward", 401, 500, 130)]
TOTAL_BEDS = sum(w[4] for w in WARDS)  # 520 staffed beds
BILL_RATE = 50  # source amounts are re-scaled x50 to realistic ₦ admission bills


def ward_of(room):
    for wid, name, lo, hi, _ in WARDS:
        if lo <= room <= hi:
            return wid, name
    return WARDS[-1][0], WARDS[-1][1]


def build_lifecare():
    hraw = pd.read_csv(RAW / "healthcare_dataset.csv")
    hraw["_key"] = hraw["Name"].str.lower().str.strip() + "|" + hraw["Gender"]
    pool = NamePool("lgh")
    names = {k: pool.get(k, g, initial=True) for k, g in hraw[["_key", "Gender"]].drop_duplicates().values}
    hraw["Patient_Name"] = hraw["_key"].map(names)
    # keep the source's random-casing defect for the dirty copy
    def scramble(n, k):
        return "".join(ch.upper() if h(k + str(i), 3) == 0 else ch.lower() for i, ch in enumerate(n))
    hraw["Patient_Name_Raw"] = [scramble(n, k) for n, k in zip(hraw.Patient_Name, hraw._key)]
    # doctors: a roster of consultants per department
    roster = {}
    dpool = NamePool("doc")
    doc_rows = []
    did = 1
    for dep in sorted(set(DEPT.values())):
        n = 8 if dep == "Endocrinology" else 5
        roster[dep] = []
        for j in range(n):
            nm = "Dr " + dpool.get(f"{dep}{j}", "Female" if h(dep + str(j), 2) else "Male")
            roster[dep].append((f"DOC{did:03d}", nm))
            doc_rows.append(dict(Doctor_ID=f"DOC{did:03d}", Doctor_Name=nm, Department=dep,
                                 Grade="Consultant" if j < 3 else "Senior Registrar"))
            did += 1
    hraw["Department"] = hraw["Medical Condition"].map(DEPT)
    pick = [roster[dep][h(doc, len(roster[dep]))] for dep, doc in zip(hraw.Department, hraw.Doctor)]
    hraw["Doctor_ID"] = [p_[0] for p_ in pick]
    hraw["Doctor_Name"] = [p_[1] for p_ in pick]
    w = [ward_of(r) for r in hraw["Room Number"]]
    hraw["Ward_ID"] = [a for a, _ in w]
    hraw["Ward"] = [b for _, b in w]
    hraw["HMO"] = hraw["Insurance Provider"].map(HMO)
    hraw["Billing_Amount"] = (hraw["Billing Amount"] * BILL_RATE).round(2)
    hraw["Admission_Date"] = pd.to_datetime(hraw["Date of Admission"])
    hraw["Discharge_Date"] = pd.to_datetime(hraw["Discharge Date"])

    # patient identity = name + gender + blood type
    hraw["_pkey"] = hraw["_key"] + "|" + hraw["Blood Type"]
    pk = {k: f"LGH-P{i:06d}" for i, k in enumerate(sorted(hraw._pkey.unique()), 1)}
    hraw["Patient_ID"] = hraw._pkey.map(pk)

    raw_keep = hraw.copy()
    # ---- clean: drop exact duplicates and negative bills
    clean = hraw.drop_duplicates(subset=["Name", "Age", "Gender", "Blood Type", "Medical Condition", "Date of Admission",
                                         "Doctor", "Hospital", "Insurance Provider", "Billing Amount", "Room Number",
                                         "Admission Type", "Discharge Date", "Medication", "Test Results"]).copy()
    n_dups = len(hraw) - len(clean)
    n_neg = int((clean.Billing_Amount < 0).sum())
    clean = clean[clean.Billing_Amount >= 0].copy()
    clean = clean.sort_values(["Admission_Date", "Patient_ID"]).reset_index(drop=True)
    clean["Admission_ID"] = [f"ADM{d.year}-{i:05d}" for i, d in enumerate(clean.Admission_Date, 1)]
    clean["Length_of_Stay"] = (clean.Discharge_Date - clean.Admission_Date).dt.days
    # readmission within 30 days of previous discharge for the same patient
    clean = clean.sort_values(["Patient_ID", "Admission_Date"])
    prev_dis = clean.groupby("Patient_ID")["Discharge_Date"].shift(1)
    gap = (clean.Admission_Date - prev_dis).dt.days
    observed = gap.between(0, 30)
    # Follow-up registry flag (includes readmissions to partner facilities): risk rises with age,
    # emergency admission, long stays and chronic conditions -> about 10% overall.
    rr = np.random.default_rng(SEED + 30)
    z = (-4.05 + 0.018 * clean.Age.clip(0, 90) + 0.55 * (clean["Admission Type"] == "Emergency")
         + 0.025 * clean.Length_of_Stay
         + clean["Medical Condition"].map({"Diabetes": 0.45, "Hypertension": 0.35, "Cancer": 0.55,
                                           "Asthma": 0.15, "Obesity": 0.10, "Arthritis": 0.0}).fillna(0))
    p_ = 1 / (1 + np.exp(-z))
    clean["Readmit_30d"] = np.where(observed | (rr.random(len(clean)) < p_), "Yes", "No")
    clean = clean.sort_values(["Admission_Date", "Admission_ID"]).reset_index(drop=True)
    adm = pd.DataFrame({
        "Admission_ID": clean.Admission_ID, "Patient_ID": clean.Patient_ID, "Patient_Name": clean.Patient_Name,
        "Age": clean.Age, "Gender": clean.Gender, "Blood_Type": clean["Blood Type"],
        "Medical_Condition": clean["Medical Condition"], "Department": clean.Department,
        "Admission_Type": clean["Admission Type"], "Admission_Date": clean.Admission_Date.dt.date,
        "Discharge_Date": clean.Discharge_Date.dt.date, "Length_of_Stay": clean.Length_of_Stay,
        "Ward_ID": clean.Ward_ID, "Ward": clean.Ward, "Room_Number": clean["Room Number"],
        "Doctor_ID": clean.Doctor_ID, "Doctor_Name": clean.Doctor_Name, "HMO": clean.HMO,
        "Billing_Amount": clean.Billing_Amount, "Medication": clean.Medication,
        "Test_Result": clean["Test Results"], "Readmit_30d": clean.Readmit_30d,
    })
    doctors = pd.DataFrame(doc_rows)
    wards = pd.DataFrame([dict(Ward_ID=a, Ward=b, First_Room=c, Last_Room=d_, Staffed_Beds=e) for a, b, c, d_, e in WARDS])
    patients = adm.groupby("Patient_ID").agg(Patient_Name=("Patient_Name", "first"), Gender=("Gender", "first"),
                                             Blood_Type=("Blood_Type", "first"),
                                             First_Admission=("Admission_Date", "min"),
                                             Admissions=("Admission_ID", "count")).reset_index()
    # daily census -> occupancy
    days = pd.date_range(clean.Admission_Date.min(), clean.Discharge_Date.max(), freq="D")
    delta = pd.Series(0, index=days)
    a_ct = clean.groupby("Admission_Date").size()
    d_ct = clean.groupby("Discharge_Date").size()
    delta = delta.add(a_ct, fill_value=0).sub(d_ct, fill_value=0)
    census = delta.cumsum().astype(int)
    census_df = pd.DataFrame({"Census_Date": census.index.date, "Patients_In_House": census.values,
                              "Staffed_Beds": TOTAL_BEDS})
    census_df["Occupancy_Pct"] = (census_df.Patients_In_House / census_df.Staffed_Beds).round(4)
    # trim warm-up / run-off periods (first/last 30 days are incomplete)
    census_df = census_df[(census.index >= clean.Admission_Date.min() + pd.Timedelta(days=30)) &
                          (census.index <= clean.Admission_Date.max())].reset_index(drop=True)

    adm.to_csv(CORE / "Lifecare_Admissions_2019_2024.csv", index=False)
    patients.to_csv(CORE / "Lifecare_Patients.csv", index=False)
    doctors.to_csv(CORE / "Lifecare_Doctors.csv", index=False)
    wards.to_csv(CORE / "Lifecare_Wards.csv", index=False)
    census_df.to_csv(CORE / "Lifecare_Daily_Bed_Census.csv", index=False)
    for nm, df in [("lc_adm", adm), ("lc_patients", patients), ("lc_doctors", doctors), ("lc_wards", wards),
                   ("lc_census", census_df), ("lc_raw", raw_keep)]:
        stage(nm, df)
    LOG.append(dict(Dataset="Lifecare_Admissions_2019_2024.csv (+ Patients, Doctors, Wards, Daily_Bed_Census) + lifecare_hospital.db",
                    Organisation="Lifecare General Hospital", Rows=len(adm), Columns=adm.shape[1],
                    Grain="One inpatient admission", Period=f"{adm.Admission_Date.min()} to {adm.Admission_Date.max()}",
                    Main_Use="Cleaning, LOS/occupancy/readmission KPIs, SQL, DAX, statistics, capstone 3",
                    Built_From=f"healthcare_dataset.csv ({n_dups} exact duplicates and {n_neg} negative bills removed; Nigerian names, departments, wards, HMOs; bills x{BILL_RATE} to ₦)"))
    for c_, t_, d_, e_ in [
        ("Admission_ID", "Text", "Unique admission key ADM<year>-<serial>", "ADM2019-00001"),
        ("Patient_ID", "Text", "Patient key (repeat patients share it)", "LGH-P000123"),
        ("Patient_Name", "Text", "Patient name - PII/PHI, teaching data only", "Ngozi Okafor"),
        ("Age / Gender / Blood_Type", "Integer / Text", "Demographics at admission", "62 / Male / A+"),
        ("Medical_Condition", "Text", "Primary diagnosis group (6)", "Diabetes"),
        ("Department", "Text", "Clinical department derived from condition", "Endocrinology"),
        ("Admission_Type", "Text", "Emergency, Urgent or Elective", "Urgent"),
        ("Admission_Date / Discharge_Date", "Date (ISO)", "Inpatient spell dates", "2023-08-20"),
        ("Length_of_Stay", "Integer (days)", "Discharge_Date - Admission_Date", 15),
        ("Ward_ID / Ward / Room_Number", "Text / Integer", "4 wards x 130 staffed beds = 520 beds", "W3 / Ruby Ward / 328"),
        ("Doctor_ID / Doctor_Name", "Text", "Admitting consultant (joins to Lifecare_Doctors)", "DOC007"),
        ("HMO", "Text", "Payer / health maintenance organisation", "Hygeia HMO"),
        ("Billing_Amount", "Currency ₦", "Total bill for the spell", 1276966.03),
        ("Medication / Test_Result", "Text", "Main medication; test outcome Normal/Abnormal/Inconclusive", "Paracetamol / Normal"),
        ("Readmit_30d", "Yes/No", "Patient was readmitted within 30 days of this discharge (Lifecare follow-up registry, incl. partner facilities)", "No"),
    ]:
        note("Lifecare_Admissions_2019_2024", c_, t_, d_, e_)
    build_sqlite_lifecare(adm, patients, doctors, wards, census_df)
    return adm


def build_sqlite_lifecare(adm, patients, doctors, wards, census):
    schema = """-- =====================================================================
-- Lifecare General Hospital - inpatient database (SQLite)
-- ETEH Bootcamp - Week 3 (SQL / ODBC) and Capstone 3
-- =====================================================================
PRAGMA foreign_keys = ON;
DROP VIEW IF EXISTS v_admissions;
DROP TABLE IF EXISTS bed_census;
DROP TABLE IF EXISTS admissions;
DROP TABLE IF EXISTS patients;
DROP TABLE IF EXISTS doctors;
DROP TABLE IF EXISTS wards;

CREATE TABLE wards (
    ward_id      TEXT PRIMARY KEY,
    ward         TEXT NOT NULL,
    first_room   INTEGER,
    last_room    INTEGER,
    staffed_beds INTEGER NOT NULL
);
CREATE TABLE doctors (
    doctor_id    TEXT PRIMARY KEY,
    doctor_name  TEXT NOT NULL,
    department   TEXT NOT NULL,
    grade        TEXT
);
CREATE TABLE patients (
    patient_id      TEXT PRIMARY KEY,
    patient_name    TEXT NOT NULL,
    gender          TEXT,
    blood_type      TEXT,
    first_admission DATE,
    admissions      INTEGER
);
CREATE TABLE admissions (
    admission_id      TEXT PRIMARY KEY,
    patient_id        TEXT NOT NULL REFERENCES patients(patient_id),
    age               INTEGER,
    medical_condition TEXT,
    department        TEXT,
    admission_type    TEXT CHECK (admission_type IN ('Emergency','Urgent','Elective')),
    admission_date    DATE NOT NULL,
    discharge_date    DATE,
    length_of_stay    INTEGER,
    ward_id           TEXT REFERENCES wards(ward_id),
    room_number       INTEGER,
    doctor_id         TEXT REFERENCES doctors(doctor_id),
    hmo               TEXT,
    billing_amount    REAL,
    medication        TEXT,
    test_result       TEXT,
    readmit_30d       TEXT
);
CREATE TABLE bed_census (
    census_date       DATE PRIMARY KEY,
    patients_in_house INTEGER,
    staffed_beds      INTEGER,
    occupancy_pct     REAL
);
CREATE INDEX ix_adm_date ON admissions(admission_date);
CREATE INDEX ix_adm_patient ON admissions(patient_id);
CREATE INDEX ix_adm_dept ON admissions(department);

CREATE VIEW v_admissions AS
SELECT a.*, p.patient_name, p.gender, p.blood_type, d.doctor_name, w.ward,
       strftime('%Y', a.admission_date) AS adm_year, strftime('%Y-%m', a.admission_date) AS adm_month
FROM admissions a
JOIN patients p ON p.patient_id = a.patient_id
JOIN doctors  d ON d.doctor_id  = a.doctor_id
JOIN wards    w ON w.ward_id    = a.ward_id;
"""
    (W3D / "lifecare_hospital_schema.sql").write_text(schema, encoding="utf-8")
    dbp = W3D / "lifecare_hospital.db"
    if dbp.exists():
        dbp.unlink()
    con = sqlite3.connect(dbp)
    con.executescript(schema)
    def ins(name, df):
        df = df.copy()
        df.columns = [c.lower() for c in df.columns]
        df = df.astype(object).where(df.notna(), None)
        q = f"INSERT INTO {name} ({', '.join(df.columns)}) VALUES ({', '.join('?' * df.shape[1])})"
        con.executemany(q, [tuple(str(v) if hasattr(v, 'isoformat') else (v.item() if hasattr(v, 'item') else v) for v in r) for r in df.values.tolist()])
    ins("wards", wards)
    ins("doctors", doctors)
    ins("patients", patients)
    a = adm.drop(columns=["Patient_Name", "Gender", "Blood_Type", "Ward", "Doctor_Name"])
    ins("admissions", a)
    ins("bed_census", census)
    con.commit()
    assert not con.execute("PRAGMA foreign_key_check").fetchall()
    con.execute("VACUUM")
    con.close()


# ======================================================================================
# 4. LIFECARE DIABETES SCREENING (source: diabetes_prediction_dataset.csv)
# ======================================================================================
def build_diabetes():
    d = pd.read_csv(RAW / "diabetes_prediction_dataset.csv")
    d.insert(0, "Screening_ID", [f"LDS-{i:06d}" for i in range(1, len(d) + 1)])
    out = pd.DataFrame({
        "Screening_ID": d.Screening_ID, "Gender": d.gender, "Age": d.age,
        "Hypertension": d.hypertension.map({0: "No", 1: "Yes"}), "Heart_Disease": d.heart_disease.map({0: "No", 1: "Yes"}),
        "Smoking_History": d.smoking_history.replace({"No Info": "Not Recorded", "never": "Never", "former": "Former",
                                                      "current": "Current", "not current": "Not Current", "ever": "Ever"}),
        "BMI": d.bmi, "HbA1c_Level": d.HbA1c_level, "Blood_Glucose_mg_dL": d.blood_glucose_level,
        "Diabetes": d.diabetes.map({0: "No", 1: "Yes"}),
    })
    out.to_csv(CORE / "Lifecare_Diabetes_Screening.csv", index=False)
    stage("lc_diab", out)
    LOG.append(dict(Dataset="Lifecare_Diabetes_Screening.csv", Organisation="Lifecare General Hospital - screening clinic",
                    Rows=len(out), Columns=out.shape[1], Grain="One screening visit", Period="Screening programme (undated)",
                    Main_Use="Distributions, z-scores, correlation, thresholds & confusion matrix, student project 3",
                    Built_From="diabetes_prediction_dataset.csv (ID added, codes labelled, smoking vocabulary standardised)"))
    for c_, t_, d_, e_ in [("Screening_ID", "Text", "Unique screening key", "LDS-000001"),
                           ("Gender / Age", "Text / Decimal", "Sex and age in years (infants < 1 recorded as decimals)", "Female / 54"),
                           ("Hypertension / Heart_Disease", "Yes/No", "Known comorbidities", "No"),
                           ("Smoking_History", "Text", "Never, Former, Current, Not Current, Ever, Not Recorded", "Never"),
                           ("BMI", "Decimal", "Body-mass index kg/m²", 27.32),
                           ("HbA1c_Level", "Decimal %", "Glycated haemoglobin; >= 6.5 is the diagnostic threshold", 6.6),
                           ("Blood_Glucose_mg_dL", "Integer", "Random blood glucose", 140),
                           ("Diabetes", "Yes/No", "Confirmed diabetes diagnosis", "No")]:
        note("Lifecare_Diabetes_Screening", c_, t_, d_, e_)
    return out


# ======================================================================================
# 5. CRESTVIEW UNIVERSITY (StudentPerformanceFactors, Student Depression, Social Media)
# ======================================================================================
FACULTIES = {
    "Science": ["Computer Science", "Microbiology", "Physics", "Biochemistry"],
    "Social Sciences": ["Economics", "Accounting", "Mass Communication", "Political Science"],
    "Engineering": ["Civil Engineering", "Electrical Engineering", "Mechanical Engineering"],
    "Arts & Education": ["English", "History", "Education"],
}
DEPT_CODE = {"Computer Science": "CSC", "Microbiology": "MCB", "Physics": "PHY", "Biochemistry": "BCH",
             "Economics": "ECO", "Accounting": "ACC", "Mass Communication": "MAC", "Political Science": "POL",
             "Civil Engineering": "CVE", "Electrical Engineering": "EEE", "Mechanical Engineering": "MEE",
             "English": "ENG", "History": "HIS", "Education": "EDU"}


def build_crestview():
    sp = pd.read_csv(RAW / "StudentPerformanceFactors.csv")
    depts = [(f, d) for f, lst in FACULTIES.items() for d in lst]
    weights = np.array([9, 6, 5, 6, 9, 9, 7, 6, 6, 7, 6, 5, 4, 5], dtype=float)
    weights /= weights.sum()
    choice = rng.choice(len(depts), len(sp), p=weights)
    level = rng.choice([100, 200, 300, 400], len(sp), p=[0.30, 0.27, 0.23, 0.20])
    entry = 2024 - (level // 100 - 1)
    serial = np.arange(1, len(sp) + 1)
    out = pd.DataFrame({
        "Student_ID": [f"CU/{e}/{DEPT_CODE[depts[ci][1]]}/{s:04d}" for e, ci, s in zip(entry, choice, serial)],
        "Faculty": [depts[i][0] for i in choice],
        "Department": [depts[i][1] for i in choice],
        "Level": level,
        "Gender": sp.Gender,
        "Hours_Studied_Per_Week": sp.Hours_Studied,
        "Attendance_Pct": sp.Attendance,
        "Sleep_Hours": sp.Sleep_Hours,
        "Previous_Score": sp.Previous_Scores,
        "Tutoring_Sessions": sp.Tutoring_Sessions,
        "Physical_Activity_Hrs": sp.Physical_Activity,
        "Parental_Involvement": sp.Parental_Involvement,
        "Access_to_Resources": sp.Access_to_Resources,
        "Extracurricular": sp.Extracurricular_Activities,
        "Motivation_Level": sp.Motivation_Level,
        "Internet_Access": sp.Internet_Access,
        "Family_Income": sp.Family_Income,
        "Teacher_Quality": sp.Teacher_Quality.fillna("Not Recorded"),
        "School_Type": sp.School_Type.replace({"Public": "Public Secondary", "Private": "Private Secondary"}),
        "Peer_Influence": sp.Peer_Influence,
        "Learning_Disability": sp.Learning_Disabilities,
        "Parental_Education": sp.Parental_Education_Level.fillna("Not Recorded"),
        "Distance_from_Home": sp.Distance_from_Home.fillna("Not Recorded"),
        "Exam_Score": sp.Exam_Score.clip(upper=100),
    })
    out["Grade"] = pd.cut(out.Exam_Score, [-1, 44, 49, 59, 69, 100], labels=["F", "D", "C", "B", "A"]).astype(str)
    out["Attendance_Band"] = pd.cut(out.Attendance_Pct, [0, 69, 79, 89, 100], labels=["<70%", "70-79%", "80-89%", "90-100%"]).astype(str)
    out["At_Risk"] = np.where((out.Exam_Score < 62) | (out.Attendance_Pct < 65), "Yes", "No")
    out.to_csv(CORE / "Crestview_Student_Performance.csv", index=False)
    stage("cv_perf", out)
    LOG.append(dict(Dataset="Crestview_Student_Performance.csv", Organisation="Crestview University", Rows=len(out),
                    Columns=out.shape[1], Grain="One student, 2024/2025 first semester", Period="2024/2025 session",
                    Main_Use="Formulas, lookups, pivots, correlation & regression, capstone 2",
                    Built_From="StudentPerformanceFactors.csv (IDs, faculty/department/level added; 1 score of 101 capped at 100; blanks labelled 'Not Recorded'; grade, attendance band and at-risk flag derived)"))
    for c_, t_, d_, e_ in [("Student_ID", "Text", "Matric number CU/<entry year>/<dept code>/<serial>", "CU/2024/CSC/0001"),
                           ("Faculty / Department / Level", "Text / Integer", "Academic home and level (100-400)", "Science / Computer Science / 100"),
                           ("Hours_Studied_Per_Week", "Integer", "Self-reported weekly private study hours", 23),
                           ("Attendance_Pct", "Integer %", "Lecture attendance this semester", 84),
                           ("Sleep_Hours", "Integer", "Average nightly sleep", 7),
                           ("Previous_Score", "Integer", "Previous semester average score", 73),
                           ("Tutoring_Sessions", "Integer", "Peer-tutoring sessions attended per month", 0),
                           ("Parental_Involvement ... Distance_from_Home", "Text", "Low/Medium/High style factors; blanks shown as 'Not Recorded'", "Medium"),
                           ("Exam_Score", "Integer 0-100", "Semester examination average", 67),
                           ("Grade", "Text", "A 70-100, B 60-69, C 50-59, D 45-49, F <45 (Nigerian 5-point scale)", "B"),
                           ("Attendance_Band", "Text", "<70%, 70-79%, 80-89%, 90-100%", "80-89%"),
                           ("At_Risk", "Yes/No", "Exam_Score < 62 OR Attendance_Pct < 65", "No")]:
        note("Crestview_Student_Performance", c_, t_, d_, e_)

    # wellbeing survey (sensitive) -------------------------------------------------
    dep = pd.read_csv(RAW / "Student Depression Dataset.csv")
    states = sorted(set(CITY_STATE.values()))
    dep_out = pd.DataFrame({
        "Respondent_ID": [f"CWS-{i:05d}" for i in range(1, len(dep) + 1)],
        "Gender": dep.Gender, "Age": dep.Age.astype(int),
        "State_of_Origin": [states[h(c_, len(states))] for c_ in dep.City],
        "Programme": dep.Degree.replace({"Class 12": "Pre-degree", "'Class 12'": "Pre-degree"}),
        "Academic_Pressure_1to5": dep["Academic Pressure"], "CGPA_5pt": (dep.CGPA / 2).round(2),
        "Study_Satisfaction_1to5": dep["Study Satisfaction"], "Sleep_Duration": dep["Sleep Duration"].str.strip("'"),
        "Dietary_Habits": dep["Dietary Habits"], "Study_Hours_Per_Day": dep["Work/Study Hours"],
        "Financial_Stress_1to5": dep["Financial Stress"],
        "Family_History_Mental_Illness": dep["Family History of Mental Illness"],
        "Suicidal_Thoughts_Ever": dep["Have you ever had suicidal thoughts ?"],
        "Depression_Screen_Positive": dep.Depression.map({0: "No", 1: "Yes"}),
    })
    dep_out = dep_out[dep.Profession.eq("Student").values].reset_index(drop=True)
    dep_out.to_csv(CORE / "Crestview_Wellbeing_Survey_SENSITIVE.csv", index=False)
    stage("cv_wellbeing", dep_out)
    LOG.append(dict(Dataset="Crestview_Wellbeing_Survey_SENSITIVE.csv", Organisation="Crestview University - Student Affairs",
                    Rows=len(dep_out), Columns=dep_out.shape[1], Grain="One survey respondent",
                    Period="2024/2025 wellbeing survey", Main_Use="Data ethics (special-category data), aggregation & small-cell suppression, correlation",
                    Built_From="Student Depression Dataset.csv (students only; CGPA converted to 5-point; city mapped to state of origin)"))
    note("Crestview_Wellbeing_Survey_SENSITIVE", "(all)", "Survey", "Contains mental-health answers - special-category data. Only aggregate results may leave the analytics team; suppress cells with < 10 respondents.")

    sm = pd.read_csv(RAW / "Students Social Media Addiction.csv")
    sm.columns = ["Student_ID", "Age", "Gender", "Academic_Level", "Country", "Avg_Daily_Usage_Hours",
                  "Most_Used_Platform", "Affects_Academic_Performance", "Sleep_Hours_Per_Night",
                  "Mental_Health_Score", "Relationship_Status", "Conflicts_Over_Social_Media", "Addicted_Score"]
    sm["Student_ID"] = [f"DWS-{i:04d}" for i in sm.Student_ID]
    sm.to_csv(CORE / "Crestview_Digital_Wellbeing_Survey.csv", index=False)
    stage("cv_social", sm)
    LOG.append(dict(Dataset="Crestview_Digital_Wellbeing_Survey.csv", Organisation="Crestview University - international partner survey",
                    Rows=len(sm), Columns=sm.shape[1], Grain="One student respondent", Period="2025 survey",
                    Main_Use="Week 1 data types & first formulas; Week 4 correlation (usage vs sleep vs addiction score)",
                    Built_From="Students Social Media Addiction.csv (IDs prefixed, columns standardised)"))
    return out


# ======================================================================================
# 6. ETEH TALENT SALARY SURVEY (Salary_Data.csv)
# ======================================================================================
EDU = {"Bachelor's Degree": "Bachelor's", "Bachelor's": "Bachelor's", "Master's Degree": "Master's",
       "Master's": "Master's", "PhD": "PhD", "phD": "PhD", "High School": "High School"}


def build_salary():
    s = pd.read_csv(RAW / "Salary_Data.csv", encoding="utf-8-sig")
    s["Annual_Salary_NGN"] = s.Salary * 100
    dirty = s.rename(columns={"Education Level": "Education_Level", "Job Title": "Job_Title",
                              "Years of Experience": "Years_of_Experience"})[
        ["Age", "Gender", "Education_Level", "Job_Title", "Years_of_Experience", "Annual_Salary_NGN"]]
    clean = dirty.dropna().copy()
    clean["Education_Level"] = clean.Education_Level.map(EDU)
    clean = clean.drop_duplicates()
    clean = clean[clean.Annual_Salary_NGN >= 1_000_000]  # removes 350/550/579-type keying errors
    clean = clean[clean.Years_of_Experience <= clean.Age - 16]
    clean.insert(0, "Respondent_ID", [f"TSS-{i:04d}" for i in range(1, len(clean) + 1)])
    def fam(t):
        t = t.lower()
        for k, v in [("data", "Data & Analytics"), ("software", "Software Engineering"), ("developer", "Software Engineering"),
                     ("engineer", "Engineering"), ("marketing", "Marketing"), ("sales", "Sales"), ("hr", "Human Resources"),
                     ("human", "Human Resources"), ("financ", "Finance"), ("account", "Finance"), ("product", "Product"),
                     ("project", "Project Management"), ("operations", "Operations"), ("design", "Design"),
                     ("research", "Research"), ("scientist", "Data & Analytics")]:
            if k in t:
                return v
        return "Other Business"
    clean["Job_Family"] = clean.Job_Title.map(fam)
    clean["Seniority"] = np.select([clean.Job_Title.str.contains("Director|VP|Chief|CEO|CTO", case=False),
                                    clean.Job_Title.str.contains("Senior|Manager|Lead|Principal", case=False),
                                    clean.Job_Title.str.contains("Junior", case=False)],
                                   ["Executive", "Senior", "Junior"], "Mid")
    clean.to_csv(CORE / "ETEH_Talent_Salary_Survey.csv", index=False)
    stage("sal_clean", clean)
    stage("sal_dirty", dirty)
    LOG.append(dict(Dataset="ETEH_Talent_Salary_Survey.csv", Organisation="ETEH Talent Partners (HR analytics)", Rows=len(clean),
                    Columns=clean.shape[1], Grain="One survey respondent", Period="2025 salary survey",
                    Main_Use="People analytics, regression (salary vs experience), student project 4",
                    Built_From=f"Salary_Data.csv ({len(dirty)} raw rows -> {len(clean)} after removing blanks, {int(dirty.duplicated().sum())} duplicates and keying errors; education labels unified; salary x100 to ₦)"))
    for c_, t_, d_, e_ in [("Respondent_ID", "Text", "Survey key", "TSS-0001"), ("Age", "Integer", "Age in years", 32),
                           ("Gender", "Text", "Male / Female / Other", "Male"),
                           ("Education_Level", "Text", "High School, Bachelor's, Master's, PhD", "Bachelor's"),
                           ("Job_Title / Job_Family / Seniority", "Text", "Role and derived groupings", "Data Analyst / Data & Analytics / Mid"),
                           ("Years_of_Experience", "Decimal", "Years of professional experience", 5),
                           ("Annual_Salary_NGN", "Currency ₦", "Gross annual salary", 9000000)]:
        note("ETEH_Talent_Salary_Survey", c_, t_, d_, e_)
    return clean


# ======================================================================================
# 7. EKO SUITES HOTEL (Hotel Reservations.csv)
# ======================================================================================
def build_hotel():
    ho = pd.read_csv(RAW / "Hotel Reservations.csv")
    yr = ho.arrival_year + 6
    dt = pd.to_datetime(dict(year=yr, month=ho.arrival_month, day=ho.arrival_date), errors="coerce")
    rooms = {"Room_Type 1": "Standard", "Room_Type 2": "Standard Twin", "Room_Type 3": "Business",
             "Room_Type 4": "Deluxe", "Room_Type 5": "Executive", "Room_Type 6": "Family Suite", "Room_Type 7": "Presidential Suite"}
    meals = {"Not Selected": "Room Only", "Meal Plan 1": "Bed & Breakfast", "Meal Plan 2": "Half Board", "Meal Plan 3": "Full Board"}
    seg = {"Online": "Online Travel Agent", "Offline": "Walk-in / Travel Agent", "Corporate": "Corporate",
           "Complementary": "Complimentary", "Aviation": "Airline Crew"}
    out = pd.DataFrame({
        "Booking_ID": ho.Booking_ID.str.replace("INN", "EKO-"),
        "Booking_Date": (dt - pd.to_timedelta(ho.lead_time, unit="D")).dt.date,
        "Arrival_Date": dt.dt.date,
        "Adults": ho.no_of_adults, "Children": ho.no_of_children,
        "Weekend_Nights": ho.no_of_weekend_nights, "Week_Nights": ho.no_of_week_nights,
        "Total_Nights": ho.no_of_weekend_nights + ho.no_of_week_nights,
        "Meal_Plan": ho.type_of_meal_plan.map(meals), "Parking_Required": ho.required_car_parking_space.map({0: "No", 1: "Yes"}),
        "Room_Type": ho.room_type_reserved.map(rooms), "Lead_Time_Days": ho.lead_time,
        "Market_Segment": ho.market_segment_type.map(seg), "Repeat_Guest": ho.repeated_guest.map({0: "No", 1: "Yes"}),
        "Previous_Cancellations": ho.no_of_previous_cancellations,
        "Previous_Stays": ho.no_of_previous_bookings_not_canceled,
        "Avg_Room_Rate_NGN": (ho.avg_price_per_room * 1000).round(0),
        "Special_Requests": ho.no_of_special_requests,
        "Booking_Status": ho.booking_status.map({"Canceled": "Cancelled", "Not_Canceled": "Honoured"}),
    })
    out["Booking_Value_NGN"] = out.Avg_Room_Rate_NGN * out.Total_Nights
    out = out.sort_values(["Arrival_Date", "Booking_ID"]).reset_index(drop=True)
    out.to_csv(CORE / "EkoSuites_Hotel_Bookings_2023_2024.csv", index=False)
    stage("hotel", out)
    LOG.append(dict(Dataset="EkoSuites_Hotel_Bookings_2023_2024.csv", Organisation="Eko Suites Hotel, Victoria Island",
                    Rows=len(out), Columns=out.shape[1], Grain="One room booking",
                    Period=f"{out.Arrival_Date.min()} to {out.Arrival_Date.max()}",
                    Main_Use="Pivots, what-if (Goal Seek, Data Tables, Solver), cancellation analysis, student project 2",
                    Built_From="Hotel Reservations.csv (dates +6 years so Feb-29 is valid, room/meal/segment codes labelled, rate x1,000 to ₦)"))
    for c_, t_, d_, e_ in [("Booking_ID", "Text", "Booking key", "EKO-00001"), ("Booking_Date / Arrival_Date", "Date", "When booked / when due to arrive", "2024-10-02"),
                           ("Adults / Children", "Integer", "Guests", 2), ("Weekend_Nights / Week_Nights / Total_Nights", "Integer", "Stay length", 3),
                           ("Meal_Plan", "Text", "Room Only, Bed & Breakfast, Half Board, Full Board", "Bed & Breakfast"),
                           ("Room_Type", "Text", "Standard ... Presidential Suite", "Deluxe"),
                           ("Lead_Time_Days", "Integer", "Days between booking and arrival", 224),
                           ("Market_Segment", "Text", "Booking channel", "Online Travel Agent"),
                           ("Avg_Room_Rate_NGN", "Currency ₦", "Average nightly rate", 65000),
                           ("Booking_Status", "Text", "Honoured or Cancelled", "Honoured"),
                           ("Booking_Value_NGN", "Currency ₦", "Avg_Room_Rate_NGN x Total_Nights", 195000)]:
        note("EkoSuites_Hotel_Bookings_2023_2024", c_, t_, d_, e_)
    return out


# ======================================================================================
# 8. DIRTY EXTRACTS (each with a defect log and control totals)
# ======================================================================================
def dirty_supermart_2025(sales):
    """SuperMart Stores FY2025 extract with 14 deliberate defect types."""
    r = np.random.default_rng(SEED + 1)
    base = sales[pd.to_datetime(sales.Order_Date).dt.year == 2025].copy().reset_index(drop=True)
    truth = base.copy()
    d = base.copy().astype(object)
    log = []
    n = len(d)

    def pick(k, exclude=()):
        pool = np.setdiff1d(np.arange(n), np.array(list(exclude), dtype=int))
        return r.choice(pool, k, replace=False)

    # 1 combined location column (Text-to-Columns lesson)
    d.insert(d.columns.get_loc("City"), "Store_Location", d.City + ", " + d.State)
    d = d.drop(columns=["City", "State"])
    log.append(("Combined column", "City and State stored together in Store_Location 'Ikeja, Lagos'", n, "Text to Columns / Split Column by delimiter ', '"))
    # 2 whitespace / casing on Zone & Store_Name
    idx = pick(int(n * 0.08)); d.loc[idx, "Zone"] = [z + "  " if i % 2 else " " + z for i, z in zip(idx, d.loc[idx, "Zone"])]
    log.append(("Leading/trailing spaces", "Zone has extra spaces", len(idx), "TRIM / Transform > Format > Trim"))
    idx2 = pick(int(n * 0.05)); d.loc[idx2, "Store_Name"] = [v.upper() if i % 2 else v.lower() for i, v in zip(idx2, d.loc[idx2, "Store_Name"])]
    log.append(("Inconsistent case", "Store_Name in UPPER or lower case", len(idx2), "PROPER / Format > Capitalize Each Word"))
    # 3 category vocabulary
    variants = {"Furniture": ["Furnitre", "furniture", "FURNITURE"], "Office Supplies": ["Office Supply", "office supplies", "Office-Supplies"],
                "Technology": ["Tech", "Technolgy", "technology"]}
    idx = pick(int(n * 0.04)); d.loc[idx, "Category"] = [variants[c_][h(i, 3)] for i, c_ in zip(idx, d.loc[idx, "Category"])]
    log.append(("Uncontrolled vocabulary", "Category misspelt/cased (Furnitre, Tech, Office Supply ...)", len(idx), "Find & Replace or mapping table / Replace Values"))
    # 4 mixed date formats
    od = pd.to_datetime(truth.Order_Date)
    idx = pick(int(n * 0.10))
    fm = ["%d/%m/%Y", "%d-%b-%Y", "%Y.%m.%d"]
    d["Order_Date"] = od.dt.strftime("%Y-%m-%d")
    d.loc[idx, "Order_Date"] = [od[i].strftime(fm[h(i, 3)]) for i in idx]
    log.append(("Mixed date formats", "Order_Date as text in dd/mm/yyyy, dd-Mon-yyyy and yyyy.mm.dd", len(idx), "DATE/DATEVALUE or Change Type with Locale (en-GB)"))
    # 5 numbers as text with currency symbol
    idx = pick(int(n * 0.07)); d.loc[idx, "Sales"] = ["₦" + f"{int(v):,}" for v in d.loc[idx, "Sales"]]
    log.append(("Numbers stored as text", "Sales typed as '₦1,234,560'", len(idx), "SUBSTITUTE + VALUE / Replace '₦' & ',' then Change Type"))
    # 6 discount as percentage text
    idx = pick(int(n * 0.05)); d.loc[idx, "Discount"] = [f"{float(v)*100:.0f}%" for v in d.loc[idx, "Discount"]]
    log.append(("Percent as text", "Discount typed as '20%' instead of 0.2", len(idx), "VALUE(SUBSTITUTE(x,\"%\",\"\"))/100"))
    # 7 hidden characters in names
    idx = pick(40); d.loc[idx, "Customer_Name"] = [v.replace(" ", "\u00a0", 1) + ("\n" if i % 2 else "") for i, v in zip(idx, d.loc[idx, "Customer_Name"])]
    log.append(("Non-printing characters", "Customer_Name contains CHAR(160) or line-feed", len(idx), "CLEAN + SUBSTITUTE(x,CHAR(160),\" \") / Clean"))
    # 8 missing values
    idx = pick(30); d.loc[idx, "Quantity"] = None
    log.append(("Missing values", "Quantity blank", len(idx), "Recover from source or flag; never assume 0"))
    idx = pick(18); d.loc[idx, "Customer_Name"] = "N/A"
    log.append(("Placeholder text", "Customer_Name = 'N/A'", len(idx), "Replace with blank (null) and flag"))
    # 9 ship date before order date
    idx = pick(12); d.loc[idx, "Ship_Date"] = [(od[i] - pd.Timedelta(days=int(1 + h(i, 4)))).strftime("%Y-%m-%d") for i in idx]
    log.append(("Impossible dates", "Ship_Date earlier than Order_Date", len(idx), "Flag with =Ship_Date<Order_Date; query the store"))
    # 10 keying outliers (x100)
    big = pick(6); d.loc[big, "Sales"] = [int(v) * 100 if not isinstance(v, str) else v for v in d.loc[big, "Sales"]]
    log.append(("Keying outliers", "Sales multiplied by 100 (extra zeros)", len(big), "Sort / MAX / IQR rule; correct from receipt"))
    # 11 negative quantity
    idx = pick(8); d.loc[idx, "Quantity"] = [-abs(int(q)) if q is not None and not pd.isna(q) else q for q in d.loc[idx, "Quantity"]]
    log.append(("Sign errors", "Quantity negative on sale lines", len(idx), "ABS() after confirming they are not returns"))
    d["Ship_Date"] = [v if isinstance(v, str) else pd.Timestamp(v).strftime("%Y-%m-%d") for v in d.Ship_Date]
    # 12 duplicates
    dup_idx = pick(60)
    d = pd.concat([d, d.loc[dup_idx]], ignore_index=True)
    log.append(("Duplicate records", "Exact duplicate sales lines", 60, "Remove Duplicates on all columns (check Row_ID)"))
    d = d.sample(frac=1, random_state=SEED).reset_index(drop=True)
    # 13 spacer rows + 14 empty column
    blanks = pd.DataFrame([[None] * d.shape[1]] * 15, columns=d.columns)
    cuts = np.linspace(0, len(d), 17).astype(int)
    parts = [d.iloc[cuts[i]:cuts[i + 1]] for i in range(16)]
    d = pd.concat([pd.concat([p_, blanks.iloc[[i]]]) if i < 15 else p_ for i, p_ in enumerate(parts)], ignore_index=True)
    d.insert(d.columns.get_loc("Product_ID"), "Column1", None)
    log.append(("Spacer rows", "Completely blank rows inside the data", 15, "Go To Special > Blanks / Remove Blank Rows"))
    log.append(("Empty column", "Column1 has no values", 1, "Delete column / Remove Columns"))
    logdf = pd.DataFrame(log, columns=["Defect", "What was injected", "Rows affected", "Expected fix"])
    ctrl = pd.DataFrame([
        ("Clean row count", len(truth)), ("Total Sales (₦)", int(truth.Sales.sum())), ("Total Profit (₦)", int(truth.Profit.sum())),
        ("Total Quantity", int(truth.Quantity.sum())), ("Distinct orders", truth.Order_ID.nunique()),
        ("Distinct customers", truth.Customer_ID.nunique()),
    ], columns=["Control total (after cleaning)", "Value"])
    return d, truth, logdf, ctrl


def dirty_lifecare_2023(raw, adm):
    r = np.random.default_rng(SEED + 2)
    yr = raw[raw.Admission_Date.dt.year == 2023].copy().reset_index(drop=True)
    d = pd.DataFrame({
        "Patient_Name": yr.Patient_Name_Raw, "Age": yr.Age.astype(object), "Gender": yr.Gender.astype(object),
        "Blood_Type": yr["Blood Type"], "Medical_Condition": yr["Medical Condition"],
        "Admission_Date": yr.Admission_Date.dt.strftime("%Y-%m-%d"), "Discharge_Date": yr.Discharge_Date.dt.strftime("%Y-%m-%d"),
        "Doctor_Name": yr.Doctor_Name, "Ward": yr.Ward.astype(object), "Room_Number": yr["Room Number"],
        "HMO": yr.HMO.astype(object), "Billing_Amount": yr.Billing_Amount.round(2).astype(object),
        "Admission_Type": yr["Admission Type"].astype(object), "Medication": yr.Medication, "Test_Result": yr["Test Results"],
    })
    n = len(d)
    log = [("Random casing", "Patient_Name in mixed case e.g. 'cHIoma okAfor' (as exported by the old EMR)", n, "PROPER / Capitalize Each Word"),
           ("Source duplicates", "Exact duplicate admissions already present in the EMR export", int(d.duplicated().sum()), "Remove Duplicates"),
           ("Negative bills", "Billing_Amount below zero (reversals keyed as admissions)", int((yr.Billing_Amount < 0).sum()), "Filter out & report to Finance")]
    idx = r.choice(n, int(n * 0.06), replace=False)
    wv = {"Emerald Ward": ["Emerald", "EMERALD WARD", "Emerald Ward "], "Sapphire Ward": ["Sapphire", "Saphire Ward", "SAPPHIRE WARD"],
          "Ruby Ward": ["Ruby", "Ruby  Ward", "RUBY WARD"], "Topaz Ward": ["Topaz", "Topaz ward", "TOPAZ WARD "]}
    d.loc[idx, "Ward"] = [wv[w][h(i, 3)] for i, w in zip(idx, d.loc[idx, "Ward"])]
    log.append(("Ward label variants", "Emerald / EMERALD WARD / Saphire Ward / trailing spaces", len(idx), "TRIM + mapping table"))
    idx = r.choice(n, int(n * 0.05), replace=False)
    d.loc[idx, "Gender"] = [{"Male": "M", "Female": "F"}[g] if i % 2 else g.lower() for i, g in zip(idx, d.loc[idx, "Gender"])]
    log.append(("Coded vs full labels", "Gender recorded as M/F/male/female", len(idx), "SWITCH / Replace Values"))
    idx = r.choice(n, int(n * 0.03), replace=False)
    d.loc[idx, "Admission_Type"] = [{"Emergency": "EMERG", "Urgent": "urgent", "Elective": "Elective "}[a] for a in d.loc[idx, "Admission_Type"]]
    log.append(("Abbreviations", "Admission_Type EMERG / urgent / 'Elective '", len(idx), "Standardise to Emergency/Urgent/Elective"))
    idx = r.choice(n, 70, replace=False); d.loc[idx, "Discharge_Date"] = None
    log.append(("Missing discharge date", "Patient still flagged in-house or date not keyed", 70, "Exclude from LOS; report count"))
    idx = r.choice(n, 25, replace=False)
    d.loc[idx, "Discharge_Date"] = [(pd.Timestamp(a) - pd.Timedelta(days=int(2 + h(i, 5)))).strftime("%Y-%m-%d") for i, a in zip(idx, d.loc[idx, "Admission_Date"])]
    log.append(("Discharge before admission", "Negative length of stay", 25, "Flag; correct from ward register"))
    idx = r.choice(n, 12, replace=False); d.loc[idx, "Age"] = [0 if i % 2 else 150 for i in idx]
    log.append(("Impossible ages", "Age 0 or 150 for adult admissions", 12, "Data Validation 0-120; query records"))
    idx = r.choice(n, 40, replace=False); d.loc[idx, "HMO"] = None
    log.append(("Missing payer", "HMO blank", 40, "Label 'Self-pay / Unknown' only if policy allows"))
    idx = r.choice(n, int(n * 0.04), replace=False); d.loc[idx, "Billing_Amount"] = [f"{float(v):,.2f}" for v in d.loc[idx, "Billing_Amount"]]
    log.append(("Thousands separator in text", "Billing_Amount stored as text '1,234,567.89'", len(idx), "Change Type with Locale / VALUE"))
    return d, pd.DataFrame(log, columns=["Defect", "What was injected", "Rows affected", "Expected fix"])


def dirty_crestview(perf):
    r = np.random.default_rng(SEED + 3)
    d = perf.drop(columns=["Grade", "Attendance_Band", "At_Risk"]).copy().astype(object)
    src = pd.read_csv(RAW / "StudentPerformanceFactors.csv")
    d["Exam_Score"] = src.Exam_Score.values.astype(object)  # includes the genuine 101
    for col in ["Teacher_Quality", "Parental_Education", "Distance_from_Home"]:
        d[col] = d[col].replace({"Not Recorded": None})
    n = len(d)
    log = [("Out-of-range score", "Exam_Score above 100 (source record)", int((src.Exam_Score > 100).sum()), "Validate 0-100; cap/query"),
           ("Missing categories", "Teacher_Quality / Parental_Education / Distance_from_Home blank", int(d[["Teacher_Quality", "Parental_Education", "Distance_from_Home"]].isna().sum().sum()), "Label 'Not Recorded'")]
    idx = r.choice(n, int(n * 0.06), replace=False)
    d.loc[idx, "Internet_Access"] = [{"Yes": ["Y", "yes", "YES"], "No": ["N", "no", "NO"]}[v][h(i, 3)] for i, v in zip(idx, d.loc[idx, "Internet_Access"])]
    log.append(("Yes/No variants", "Internet_Access as Y / yes / YES / N / no", len(idx), "UPPER+LEFT or mapping"))
    idx = r.choice(n, int(n * 0.05), replace=False); d.loc[idx, "Attendance_Pct"] = [f"{v}%" for v in d.loc[idx, "Attendance_Pct"]]
    log.append(("Percent as text", "Attendance_Pct typed '84%'", len(idx), "VALUE(SUBSTITUTE())"))
    idx = r.choice(n, int(n * 0.04), replace=False); d.loc[idx, "Gender"] = [{"Male": "M", "Female": "F"}[g] for g in d.loc[idx, "Gender"]]
    log.append(("Coded gender", "Gender as M/F", len(idx), "Standardise"))
    idx = r.choice(n, int(n * 0.03), replace=False)
    typo = {"Computer Science": "Computer Sci.", "Economics": "Econs", "Accounting": "Acct", "Microbiology": "Micro-biology",
            "Mass Communication": "Mass Comm", "Civil Engineering": "Civil Engr", "Electrical Engineering": "Elect. Engineering"}
    d.loc[idx, "Department"] = [typo.get(v, v.upper()) for v in d.loc[idx, "Department"]]
    log.append(("Department variants", "Abbreviated / upper-case department names", len(idx), "Lookup to department master"))
    dup = r.choice(n, 45, replace=False)
    d = pd.concat([d, d.loc[dup]], ignore_index=True)
    log.append(("Duplicate students", "Same Student_ID appears twice (double registration)", 45, "COUNTIF(ID)>1 / Remove Duplicates"))
    idx = r.choice(len(d), 25, replace=False); d.loc[idx, "Student_ID"] = [" " + v + " " for v in d.loc[idx, "Student_ID"]]
    log.append(("Padded keys", "Student_ID with leading/trailing spaces (breaks lookups)", 25, "TRIM before VLOOKUP/XLOOKUP"))
    d = d.sample(frac=1, random_state=SEED).reset_index(drop=True)
    return d, pd.DataFrame(log, columns=["Defect", "What was injected", "Rows affected", "Expected fix"])


def dirty_hotel(hotel):
    r = np.random.default_rng(SEED + 4)
    d = hotel.copy().astype(object)
    n = len(d)
    d["Arrival_Date"] = [x.strftime("%d/%m/%Y") for x in hotel.Arrival_Date]
    d["Booking_Date"] = [x.strftime("%d/%m/%Y") for x in hotel.Booking_Date]
    log = [("Day-first text dates", "Arrival_Date & Booking_Date as dd/mm/yyyy text", n, "Change Type with Locale en-GB / DATE()")]
    idx = r.choice(n, 4, replace=False); d.loc[idx, "Arrival_Date"] = "29/02/2023"
    log.append(("Invalid calendar date", "29/02/2023 (2023 is not a leap year)", 4, "Flag & query front desk"))
    idx = r.choice(n, int(n * 0.03), replace=False)
    d.loc[idx, "Market_Segment"] = [{"Online Travel Agent": "OTA", "Walk-in / Travel Agent": "Offline", "Corporate": "corporate",
                                     "Complimentary": "Comp", "Airline Crew": "Aviation"}[v] for v in d.loc[idx, "Market_Segment"]]
    log.append(("Segment codes", "OTA / Offline / corporate / Comp / Aviation", len(idx), "Mapping table"))
    idx = r.choice(n, int(n * 0.02), replace=False); d.loc[idx, "Booking_Status"] = [v.upper() for v in d.loc[idx, "Booking_Status"]]
    log.append(("Status casing", "CANCELLED / HONOURED", len(idx), "PROPER"))
    zero = int((hotel.Avg_Room_Rate_NGN == 0).sum())
    log.append(("Zero room rate", "Avg_Room_Rate_NGN = 0 (complimentary & data gaps) - genuine source values", zero, "Separate complimentary stays before averaging"))
    log.append(("Extreme child count", "Children up to 10 in one room (source records)", int((hotel.Children >= 9).sum()), "Validate occupancy rules"))
    dup = r.choice(n, 120, replace=False)
    d = pd.concat([d, d.loc[dup]], ignore_index=True)
    log.append(("Duplicate bookings", "Same Booking_ID twice (channel manager re-sent)", 120, "Remove duplicates on Booking_ID"))
    d = d.drop(columns=["Booking_Value_NGN", "Total_Nights"])
    log.append(("Derived columns removed", "Total_Nights and Booking_Value_NGN must be rebuilt", 0, "Calculated columns"))
    return d, pd.DataFrame(log, columns=["Defect", "What was injected", "Rows affected", "Expected fix"])


def dirty_diabetes(diab):
    r = np.random.default_rng(SEED + 5)
    base = diab.sample(12000, random_state=SEED).reset_index(drop=True)
    d = base.copy().astype(object)
    n = len(d)
    log = []
    idx = r.choice(n, 240, replace=False); d.loc[idx, "BMI"] = None
    log.append(("Missing BMI", "BMI blank (scale not available)", 240, "Keep blank; exclude from BMI statistics"))
    idx = r.choice(n, 60, replace=False); d.loc[idx, "BMI"] = [round(float(base.BMI[i]) * 10, 1) for i in idx]
    log.append(("Decimal-point slip", "BMI keyed x10 (e.g. 273.2)", 60, "Range check 10-70"))
    idx = r.choice(n, int(n * 0.04), replace=False); d.loc[idx, "Diabetes"] = [{"Yes": "1", "No": "0"}[v] for v in d.loc[idx, "Diabetes"]]
    log.append(("Mixed outcome coding", "Diabetes as 1/0 text mixed with Yes/No", len(idx), "Standardise"))
    idx = r.choice(n, int(n * 0.03), replace=False); d.loc[idx, "HbA1c_Level"] = [f"{v}%" for v in d.loc[idx, "HbA1c_Level"]]
    log.append(("Unit in value", "HbA1c_Level typed '6.6%'", len(idx), "Strip % and convert"))
    log.append(("Not Recorded smoking", "Smoking_History = 'Not Recorded' (genuine gap)", int((base.Smoking_History == "Not Recorded").sum()), "Keep as its own category"))
    dup = r.choice(n, 80, replace=False)
    d = pd.concat([d, d.loc[dup]], ignore_index=True)
    log.append(("Duplicate screening IDs", "Screening_ID repeated", 80, "Remove duplicates"))
    return d.sample(frac=1, random_state=SEED).reset_index(drop=True), pd.DataFrame(log, columns=["Defect", "What was injected", "Rows affected", "Expected fix"])


# ======================================================================================
# 9. WEEK 2 IMPORT-PRACTICE FILES (csv, txt, semicolon/EU locale, folder, html, json, pdf)
# ======================================================================================
def week2_files(sales, stores, adm, dirty_sm):
    # (a) monthly CSV folder for "Combine Files from Folder"
    folder = W2D / "SuperMart_Monthly_Sales_2025"
    if folder.exists():
        shutil.rmtree(folder)
    folder.mkdir()
    s25 = sales[pd.to_datetime(sales.Order_Date).dt.year == 2025]
    for m, g in s25.groupby(pd.to_datetime(s25.Order_Date).dt.month):
        g.to_csv(folder / f"2025-{m:02d}_SuperMart_Sales.csv", index=False)
    # (b) pipe-delimited TXT
    lc23 = adm[pd.to_datetime(adm.Admission_Date).dt.year == 2023]
    lc23.to_csv(W2D / "Lifecare_Admissions_2023_pipe.txt", sep="|", index=False)
    # (c) EU-locale CSV: semicolon delimiter, comma decimals, Windows-1252
    eu = stores.copy()
    eu["Floor_Area_sqm"] = eu.Floor_Area_sqm.map(lambda v: f"{v:.1f}".replace(".", ","))
    eu["Annual_Target_2025"] = eu.Annual_Target_2025.map(lambda v: f"{v:,.2f}".replace(",", " ").replace(".", ","))
    eu.to_csv(W2D / "SuperMart_Store_Targets_EU_semicolon.csv", sep=";", index=False, encoding="cp1252")
    # (d) dirty SuperMart extract as CSV (UTF-8 with BOM so ₦ survives)
    dirty_sm.to_csv(W2D / "SuperMart_Sales_2025_DIRTY.csv", index=False, encoding="utf-8-sig")
    # (e) JSON - seller directory
    smo_s = pd.read_pickle(STAGE / "smo_sellers.pkl")
    (W2D / "SuperMart_Online_Sellers.json").write_text(
        json.dumps({"source": "SuperMart Online seller API (training extract)", "generated": "2026-01-05",
                    "sellers": json.loads(smo_s.to_json(orient="records", date_format="iso"))}, indent=1), encoding="utf-8")
    # (f) HTML - store directory web page (for From Web / Web.Page)
    s24 = sales[pd.to_datetime(sales.Order_Date).dt.year == 2024].groupby("Store_ID").Sales.sum()
    rows = "\n".join(
        f"<tr><td>{r.Store_ID}</td><td>{r.Store_Name}</td><td>{r.City}</td><td>{r.State}</td><td>{r.Zone}</td>"
        f"<td>{pd.Timestamp(r.Opened).strftime('%d %b %Y')}</td><td>{r.Floor_Area_sqm:,}</td><td>{r.Store_Manager}</td></tr>"
        for r in stores.itertuples())
    html = f"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><title>SuperMart Retail - Store Directory</title>
<style>body{{font-family:Segoe UI,Arial;margin:32px;color:#0B2545}}table{{border-collapse:collapse}}td,th{{border:1px solid #ccd;padding:6px 10px}}th{{background:#0B2545;color:#fff}}</style></head>
<body><h1>SuperMart Retail Nigeria</h1><p>Store directory - published by Head Office Communications. Last updated 05 January 2026.</p>
<h2>Table 1: Store directory</h2>
<table id="stores"><thead><tr><th>Store ID</th><th>Store</th><th>City</th><th>State</th><th>Zone</th><th>Opened</th><th>Floor area (sq m)</th><th>Manager</th></tr></thead>
<tbody>
{rows}
</tbody></table>
<h2>Table 2: Opening hours</h2>
<table id="hours"><tr><th>Days</th><th>Hours</th></tr><tr><td>Monday - Saturday</td><td>08:00 - 21:00</td></tr><tr><td>Sunday</td><td>12:00 - 20:00</td></tr></table>
<p>Footnote: Floor area excludes car parks and warehouse space.</p></body></html>"""
    (W2D / "SuperMart_Store_Directory.html").write_text(html, encoding="utf-8")
    # (g) PDF bulletin - store x month cross-tab for 2025 (₦ millions)
    make_pdf_bulletin(s25, stores)


def make_pdf_bulletin(s25, stores):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    from common import FONT_DIR
    pdfmetrics.registerFont(TTFont("DejaVu", str(FONT_DIR / "DejaVuSans.ttf")))
    pdfmetrics.registerFont(TTFont("DejaVu-Bold", str(FONT_DIR / "DejaVuSans-Bold.ttf")))
    s = s25.copy()
    s["M"] = pd.to_datetime(s.Order_Date).dt.month
    ct = s.pivot_table(index="Store_Name", columns="M", values="Sales", aggfunc="sum", fill_value=0) / 1e6
    ct = ct.reindex(stores.Store_Name)
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    data = [["Store"] + months + ["FY 2025"]]
    for st, row in ct.iterrows():
        data.append([st] + [f"{v:,.1f}" for v in row.values] + [f"{row.sum():,.1f}"])
    tot = ct.sum()
    data.append(["Total"] + [f"{v:,.1f}" for v in tot.values] + [f"{tot.sum():,.1f}"])
    pdfp = W2D / "SuperMart_Monthly_Sales_Bulletin_2025.pdf"
    doc = SimpleDocTemplate(str(pdfp), pagesize=landscape(A4), leftMargin=12 * mm, rightMargin=12 * mm, topMargin=12 * mm,
                            bottomMargin=12 * mm, title="SuperMart Monthly Sales Bulletin 2025", author="SuperMart Finance")
    H = ParagraphStyle("h", fontName="DejaVu-Bold", fontSize=16, textColor=colors.HexColor("#0B2545"))
    P = ParagraphStyle("p", fontName="DejaVu", fontSize=9, leading=12)
    tbl = Table(data, repeatRows=1)
    tbl.setStyle(TableStyle([("FONT", (0, 0), (-1, -1), "DejaVu", 8), ("FONT", (0, 0), (-1, 0), "DejaVu-Bold", 8),
                             ("FONT", (0, -1), (-1, -1), "DejaVu-Bold", 8),
                             ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B2545")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                             ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#EEF4F8")), ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
                             ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#99A")),
                             ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F7F9FB")])]))
    z = s.groupby("Zone").agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"), Orders=("Order_ID", "nunique"))
    zdata = [["Zone", "Sales (₦m)", "Profit (₦m)", "Margin %", "Orders"]] + [
        [k, f"{r.Sales/1e6:,.1f}", f"{r.Profit/1e6:,.1f}", f"{r.Profit/r.Sales:.1%}", f"{r.Orders:,}"] for k, r in z.iterrows()]
    zt = Table(zdata)
    zt.setStyle(TableStyle([("FONT", (0, 0), (-1, -1), "DejaVu", 8), ("FONT", (0, 0), (-1, 0), "DejaVu-Bold", 8),
                            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#13807A")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                            ("ALIGN", (1, 0), (-1, -1), "RIGHT"), ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#99A"))]))
    doc.build([Paragraph("SuperMart Retail Nigeria — Monthly Sales Bulletin, FY 2025", H), Spacer(1, 4 * mm),
               Paragraph("Table 1. Net sales by store and month, ₦ millions (after discounts, excluding VAT). Source: SuperMart POS ledger, "
                         "closed 05 January 2026. Figures are rounded to one decimal place, so rows may not add exactly.", P),
               Spacer(1, 3 * mm), tbl, Spacer(1, 6 * mm),
               Paragraph("Table 2. Sales, profit and margin by zone, FY 2025.", P), Spacer(1, 2 * mm), zt, Spacer(1, 4 * mm),
               Paragraph("Note: the bulletin is a presentation document. Analysts must reconcile it to the ledger "
                         "(SuperMart_Stores_Sales_2022_2025.csv) before using it.", P)])


# ======================================================================================
def write_dictionary():
    cat = pd.DataFrame(LOG)
    dd = pd.DataFrame(DICT)
    with pd.ExcelWriter(CORE / "00_Data_Dictionary_and_Catalogue.xlsx", engine="xlsxwriter") as xw:
        wb = xw.book
        hdr = wb.add_format({"bold": True, "bg_color": "#0B2545", "font_color": "white", "text_wrap": True, "valign": "top", "border": 1})
        wrap = wb.add_format({"text_wrap": True, "valign": "top", "border": 1})
        for name, df, widths in [("Catalogue", cat, [44, 30, 9, 9, 34, 28, 52, 70]), ("Dictionary", dd, [34, 34, 16, 70, 26])]:
            df.to_excel(xw, sheet_name=name, index=False, startrow=2)
            ws = xw.sheets[name]
            ws.write(0, 0, "ETEH Bootcamp - " + ("Dataset catalogue (all files are training data; personal names are fictitious)" if name == "Catalogue" else "Column dictionary"),
                     wb.add_format({"bold": True, "font_size": 14, "font_color": "#0B2545"}))
            for j, col in enumerate(df.columns):
                ws.write(2, j, col, hdr)
                ws.set_column(j, j, widths[j] if j < len(widths) else 20, wrap)
            ws.freeze_panes(3, 0)
    cat.to_csv(STAGE / "catalogue.csv", index=False)


def main():
    sales, stores = build_supermart_stores()
    build_supermart_online()
    adm = build_lifecare()
    build_diabetes()
    build_crestview()
    build_salary()
    hotel = build_hotel()

    dsm, tsm, lsm, csm = dirty_supermart_2025(sales)
    stage("dirty_sm", dsm); stage("dirty_sm_truth", tsm); stage("dirty_sm_log", lsm); stage("dirty_sm_ctrl", csm)
    dlc, llc = dirty_lifecare_2023(pd.read_pickle(STAGE / "lc_raw.pkl"), adm)
    stage("dirty_lc", dlc); stage("dirty_lc_log", llc)
    dcv, lcv = dirty_crestview(pd.read_pickle(STAGE / "cv_perf.pkl"))
    stage("dirty_cv", dcv); stage("dirty_cv_log", lcv)
    dho, lho = dirty_hotel(hotel)
    stage("dirty_hotel", dho); stage("dirty_hotel_log", lho)
    ddi, ldi = dirty_diabetes(pd.read_pickle(STAGE / "lc_diab.pkl"))
    stage("dirty_diab", ddi); stage("dirty_diab_log", ldi)
    week2_files(sales, stores, adm, dsm)
    dlc.to_csv(W2D / "Lifecare_Admissions_2023_DIRTY.csv", index=False)
    write_dictionary()
    print("Datasets built.")
    for r in LOG:
        print(f"  {r['Dataset'][:70]:70s} rows={r['Rows']:>7,}")


if __name__ == "__main__":
    main()
