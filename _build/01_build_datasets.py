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
import neworgs as NG  # noqa: E402
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


STORE_TYPES = ["Ventures", "Enterprises", "Global Concepts", "Stores", "Mart", "Superstore", "Trading Company", "Emporium",
               "Supplies", "Online Shop", "Retail Hub", "Depot", "Plaza", "Merchants", "Shopping Centre"]


def seller_trading_names(s):
    """Data standard for sellers.store_name: the seller's registered trading name.
    Unique across the marketplace (UNIQUE in the database), no numeric suffixes, and any place in the
    name is the store's own city. Most follow the common '<Surname> <business type>' pattern; every
    fourth seller trades under its city. Deterministic: the same input always gives the same names."""
    used, out = set(), []
    for i, r in enumerate(s.itertuples(index=False)):
        surname = str(r.owner_name).split()[-1]
        stems = [r.store_city, surname] if i % 4 == 3 else [surname, r.store_city]
        name = None
        for k in range(len(STORE_TYPES)):
            for stem in stems:
                cand = f"{stem} {STORE_TYPES[(i * 7 + k) % len(STORE_TYPES)]}"
                if cand not in used:
                    name = cand
                    break
            if name:
                break
        assert name, r.seller_id
        used.add(name)
        out.append(name)
    return out


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
    s["store_name"] = seller_trading_names(s)

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
                    Main_Use="SQL, Power Query merges, Data Model and PivotTables, student project 1",
                    Built_From="e-commerce.xlsx (26,000 blank padding rows removed, delivery key repaired, city->state repaired, catalogue & ₦ prices standardised)"))
    for tname, cols_ in {
        "customers": "customer_id PK; names; email; phone (text, leading 0 kept); gender; date_of_birth; city/state/zone; membership_tier; loyalty_points; total_orders & total_spent (recomputed from transactions)",
        "sellers": "seller_id PK; store_name (registered trading name, unique; any place in it is the store city); owner; contact; store_city/state/zone; rating; total_sales & total_revenue (recomputed); account_status",
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
    store_name       TEXT NOT NULL UNIQUE,
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
# 3. MERCYLAND SPECIALIST HOSPITAL (original data generated by neworgs.py)
# ======================================================================================
WARDS = NG.WARDS
TOTAL_BEDS = NG.TOTAL_BEDS  # 200 staffed beds


def build_lifecare():
    pool = NamePool("msh")
    ep, patients, doctors, wards, census_df = NG.gen_hospital(lambda k, g: pool.get(k, g, initial=True))
    adm = ep
    adm.to_csv(CORE / "Mercyland_Inpatient_Episodes_2019_2024.csv", index=False)
    patients.to_csv(CORE / "Mercyland_Patients.csv", index=False)
    doctors.to_csv(CORE / "Mercyland_Consultants.csv", index=False)
    wards.to_csv(CORE / "Mercyland_Wards.csv", index=False)
    census_df.to_csv(CORE / "Mercyland_Daily_Bed_Census.csv", index=False)
    for nm, df in [("lc_adm", adm), ("lc_patients", patients), ("lc_doctors", doctors), ("lc_wards", wards),
                   ("lc_census", census_df)]:
        stage(nm, df)
    LOG.append(dict(Dataset="Mercyland_Inpatient_Episodes_2019_2024.csv (+ Patients, Consultants, Wards, Daily_Bed_Census) + mercyland_hospital.db",
                    Organisation="Mercyland Specialist Hospital, Abuja", Rows=len(adm), Columns=adm.shape[1],
                    Grain="One inpatient episode (admission to discharge)", Period=f"{adm.Admit_Date.min()} to {adm.Admit_Date.max()}",
                    Main_Use="Cleaning, LOS/occupancy/readmission KPIs, SQL, PivotTables, statistics, capstone 3",
                    Built_From="Original ETEH teaching data (neworgs.py): 6 diagnosis groups in 5 departments, 4 wards / 200 beds, "
                               "seasonal admissions with a second-quarter surge, medical inflation in bills"))
    for c_, t_, d_, e_ in [
        ("Episode_ID", "Text", "Unique episode key ADM<year>-<serial>", "ADM2019-00001"),
        ("Patient_No", "Text", "Patient key (repeat patients share it)", "MSH-P000123"),
        ("Patient_Name", "Text", "Patient name - PII/PHI, teaching data only", "Ngozi O. Okafor"),
        ("Age / Gender / Blood_Group", "Integer / Text", "Demographics at admission", "62 / Male / O+"),
        ("Diagnosis_Group", "Text", "Malaria, Diabetes, Hypertension, Pneumonia (children), Childbirth, Trauma", "Malaria"),
        ("Department", "Text", "Internal Medicine, Cardiology, Paediatrics, Obstetrics & Gynaecology, Surgery", "Internal Medicine"),
        ("Admission_Route", "Text", "Emergency, Urgent or Elective", "Urgent"),
        ("Admit_Date / Discharge_Date", "Date (ISO)", "Inpatient episode dates", "2023-08-20"),
        ("LOS_Days", "Integer (days)", "Discharge_Date - Admit_Date", 4),
        ("Ward_Code / Ward / Bed_No", "Text / Integer", "4 wards, 200 staffed beds", "W1 / Iroko Ward / 128"),
        ("Consultant_ID / Consultant_Name", "Text", "Admitting consultant (joins to Mercyland_Consultants)", "CON007"),
        ("HMO", "Text", "Payer: NHIA or one of four HMOs", "Avon HMO"),
        ("Bill_Amount_NGN", "Currency ₦", "Total bill for the episode", 516000.00),
        ("Medication / Lab_Result", "Text", "Main medicine; lab outcome Normal/Abnormal/Inconclusive", "Metformin / Abnormal"),
        ("Readmitted_30d", "Yes/No", "Readmitted within 30 days of this discharge (Mercyland follow-up registry, incl. partner facilities)", "No"),
    ]:
        note("Mercyland_Inpatient_Episodes_2019_2024", c_, t_, d_, e_)
    note("Mercyland_Daily_Bed_Census", "Census_Date / Inpatients / Open_Beds / Occupancy_Pct", "Date / Integer / Integer / Decimal",
         "Midnight count of inpatients against 200 staffed beds; above 100% means patients on trolleys or in overflow", "0.8650")
    build_sqlite_lifecare(adm, patients, doctors, wards, census_df)
    return adm


def build_sqlite_lifecare(adm, patients, doctors, wards, census):
    schema = """-- =====================================================================
-- Mercyland Specialist Hospital - inpatient database (SQLite)
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
    ward_code    TEXT PRIMARY KEY,
    ward         TEXT NOT NULL,
    first_bed    INTEGER,
    last_bed     INTEGER,
    open_beds    INTEGER NOT NULL
);
CREATE TABLE doctors (
    consultant_id   TEXT PRIMARY KEY,
    consultant_name TEXT NOT NULL,
    department      TEXT NOT NULL,
    grade           TEXT
);
CREATE TABLE patients (
    patient_no      TEXT PRIMARY KEY,
    patient_name    TEXT NOT NULL,
    gender          TEXT,
    blood_group     TEXT,
    first_episode   DATE,
    admissions      INTEGER
);
CREATE TABLE admissions (
    episode_id        TEXT PRIMARY KEY,
    patient_no        TEXT NOT NULL REFERENCES patients(patient_no),
    age               INTEGER,
    diagnosis_group   TEXT,
    department        TEXT,
    admission_route   TEXT CHECK (admission_route IN ('Emergency','Urgent','Elective')),
    admit_date        DATE NOT NULL,
    discharge_date    DATE,
    los_days          INTEGER,
    ward_code         TEXT REFERENCES wards(ward_code),
    bed_no            INTEGER,
    consultant_id     TEXT REFERENCES doctors(consultant_id),
    hmo               TEXT,
    bill_amount_ngn   REAL,
    medication        TEXT,
    lab_result        TEXT,
    readmitted_30d    TEXT
);
CREATE TABLE bed_census (
    census_date       DATE PRIMARY KEY,
    inpatients        INTEGER,
    open_beds         INTEGER,
    occupancy_pct     REAL
);
CREATE INDEX ix_adm_date ON admissions(admit_date);
CREATE INDEX ix_adm_patient ON admissions(patient_no);
CREATE INDEX ix_adm_dept ON admissions(department);

CREATE VIEW v_admissions AS
SELECT a.*, p.patient_name, p.gender, p.blood_group, d.consultant_name, w.ward,
       strftime('%Y', a.admit_date) AS adm_year, strftime('%Y-%m', a.admit_date) AS adm_month
FROM admissions a
JOIN patients p ON p.patient_no = a.patient_no
JOIN doctors  d ON d.consultant_id = a.consultant_id
JOIN wards    w ON w.ward_code  = a.ward_code;
"""
    (W3D / "mercyland_hospital_schema.sql").write_text(schema, encoding="utf-8")
    dbp = W3D / "mercyland_hospital.db"
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
    a = adm.drop(columns=["Patient_Name", "Gender", "Blood_Group", "Ward", "Consultant_Name"])
    ins("admissions", a)
    ins("bed_census", census)
    con.commit()
    assert not con.execute("PRAGMA foreign_key_check").fetchall()
    con.execute("VACUUM")
    con.close()


# ======================================================================================
# 4. MERCYLAND COMMUNITY DIABETES SCREENING (original data, neworgs.py)
# ======================================================================================
def build_diabetes():
    out = NG.gen_screening()
    out.to_csv(CORE / "Mercyland_Diabetes_Screening.csv", index=False)
    stage("lc_diab", out)
    LOG.append(dict(Dataset="Mercyland_Diabetes_Screening.csv", Organisation="Mercyland Specialist Hospital - community screening programme",
                    Rows=len(out), Columns=out.shape[1], Grain="One screening visit (adults)",
                    Period=f"{out.Screen_Date.min()} to {out.Screen_Date.max()}",
                    Main_Use="Distributions, z-scores, correlation, thresholds & confusion matrix, student project 3",
                    Built_From="Original ETEH teaching data (neworgs.py): six Abuja outreach sites; risk rises with age, BMI and blood pressure"))
    for c_, t_, d_, e_ in [("Screen_ID", "Text", "Unique screening key", "MSS-000001"),
                           ("Screen_Date / Site", "Date / Text", "Visit date and outreach site", "2023-03-14 / Kubwa"),
                           ("Gender / Age", "Text / Integer", "Sex and age in years (adults 18+)", "Female / 54"),
                           ("Hypertension / Heart_Condition", "Yes/No", "Known comorbidities", "No"),
                           ("Smoking_Status", "Text", "Never, Former, Current, Not Recorded", "Never"),
                           ("BMI", "Decimal", "Body-mass index kg/m²", 27.3),
                           ("HbA1c_Pct", "Decimal %", "Glycated haemoglobin; >= 6.5 is the diagnostic threshold", 6.6),
                           ("Glucose_mg_dL", "Integer", "Blood glucose at the visit", 140),
                           ("Diabetes", "Yes/No", "Confirmed diabetes diagnosis after follow-up", "No")]:
        note("Mercyland_Diabetes_Screening", c_, t_, d_, e_)
    return out


# ======================================================================================
# 5. KINGSWAY UNIVERSITY, ILORIN (original data, neworgs.py)
# ======================================================================================
def build_crestview():
    out = NG.gen_students()
    out.to_csv(CORE / "Kingsway_Student_Results.csv", index=False)
    stage("cv_perf", out)
    LOG.append(dict(Dataset="Kingsway_Student_Results.csv", Organisation="Kingsway University, Ilorin", Rows=len(out),
                    Columns=out.shape[1], Grain="One student, 2024/2025 first semester", Period="2024/2025 session",
                    Main_Use="Formulas, lookups, pivots, correlation & regression, capstone 2",
                    Built_From="Original ETEH teaching data (neworgs.py): scores driven by attendance and study hours, a small "
                               "high-achiever group; grade, attendance band and at-risk flag derived"))
    for c_, t_, d_, e_ in [("Matric_No", "Text", "Matric number KU/<entry year>/<dept code>/<serial>", "KU/2024/CSC/0001"),
                           ("Faculty / Department / Level", "Text / Integer", "Academic home and level (100-400)", "Science / Computer Science / 100"),
                           ("Study_Hours_Wk", "Integer", "Self-reported weekly private study hours", 19),
                           ("Lecture_Attendance_Pct", "Integer %", "Lecture attendance this semester", 84),
                           ("Sleep_Hrs_Night", "Integer", "Average nightly sleep", 7),
                           ("Last_Sem_Score", "Integer", "Previous semester average score", 73),
                           ("Tutorials_Attended", "Integer", "Tutorial classes attended per month", 1),
                           ("Exercise_Hrs_Wk", "Integer", "Sport / exercise hours per week", 3),
                           ("Guardian_Support ... Commute_Band", "Text", "Low/Medium/High style factors; gaps shown as 'Not Recorded'", "Medium"),
                           ("Final_Score", "Integer 0-100", "Semester examination average", 65),
                           ("Grade", "Text", "A 70-100, B 60-69, C 50-59, D 45-49, F <45 (Nigerian 5-point scale)", "B"),
                           ("Attendance_Band", "Text", "<70%, 70-79%, 80-89%, 90-100%", "80-89%"),
                           ("At_Risk", "Yes/No", "Final_Score < 62 OR Lecture_Attendance_Pct < 65", "No")]:
        note("Kingsway_Student_Results", c_, t_, d_, e_)
    dep_out = NG.gen_welfare()
    dep_out.to_csv(CORE / "Kingsway_Student_Welfare_Survey_SENSITIVE.csv", index=False)
    stage("cv_wellbeing", dep_out)
    LOG.append(dict(Dataset="Kingsway_Student_Welfare_Survey_SENSITIVE.csv", Organisation="Kingsway University - Student Affairs",
                    Rows=len(dep_out), Columns=dep_out.shape[1], Grain="One survey respondent",
                    Period="2024/2025 welfare survey", Main_Use="Data ethics (special-category data), aggregation & small-cell suppression, correlation",
                    Built_From="Original ETEH teaching data (neworgs.py): PHQ-9 screen linked to study pressure, money worries and sleep"))
    note("Kingsway_Student_Welfare_Survey_SENSITIVE", "(all)", "Survey", "Contains mental-health answers - special-category data. Only aggregate results may leave the analytics team; suppress cells with < 10 respondents.")
    sm = NG.gen_digital()
    sm.to_csv(CORE / "Kingsway_Digital_Habits_Survey.csv", index=False)
    stage("cv_social", sm)
    LOG.append(dict(Dataset="Kingsway_Digital_Habits_Survey.csv", Organisation="Kingsway University - digital habits survey",
                    Rows=len(sm), Columns=sm.shape[1], Grain="One student respondent", Period="2025 survey",
                    Main_Use="Week 1 data types & first formulas; Week 4 correlation (screen time vs sleep vs dependence score)",
                    Built_From="Original ETEH teaching data (neworgs.py)"))
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
# 7. ATLANTIC PEARL HOTEL, LEKKI (original data, neworgs.py)
# ======================================================================================
def build_hotel():
    out = NG.gen_hotel()
    out["Booked_On"] = pd.to_datetime(out.Booked_On)
    out["Check_In"] = pd.to_datetime(out.Check_In)
    out.assign(Booked_On=out.Booked_On.dt.date, Check_In=out.Check_In.dt.date).to_csv(
        CORE / "AtlanticPearl_Reservations_2023_2024.csv", index=False)
    out["Booked_On"] = out.Booked_On.dt.date
    out["Check_In"] = out.Check_In.dt.date
    stage("hotel", out)
    LOG.append(dict(Dataset="AtlanticPearl_Reservations_2023_2024.csv", Organisation="Atlantic Pearl Hotel, Lekki (320 rooms)",
                    Rows=len(out), Columns=out.shape[1], Grain="One room reservation",
                    Period=f"{out.Check_In.min()} to {out.Check_In.max()}",
                    Main_Use="Pivots, what-if (Goal Seek, Data Tables, Solver), cancellation analysis, student project 2",
                    Built_From="Original ETEH teaching data (neworgs.py): cancellations rise with days booked ahead; Deluxe rooms "
                               "reopened after a 2024 refurbishment"))
    for c_, t_, d_, e_ in [("Reservation_ID", "Text", "Reservation key", "APH-00001"), ("Booked_On / Check_In", "Date", "When booked / when due to arrive", "2024-10-02"),
                           ("Adults / Children", "Integer", "Guests", 2), ("Weekend_Nights / Week_Nights / Nights", "Integer", "Stay length", 3),
                           ("Board_Basis", "Text", "Room Only, Bed & Breakfast, Half Board, Full Board", "Bed & Breakfast"),
                           ("Room_Category", "Text", "Standard ... Presidential Suite", "Deluxe"),
                           ("Days_Ahead", "Integer", "Days between booking and check-in", 64),
                           ("Sales_Channel", "Text", "Booking channel", "Online Travel Agent"),
                           ("Room_Rate_NGN", "Currency ₦", "Average nightly rate", 95000),
                           ("Reservation_Status", "Text", "Stayed or Cancelled", "Stayed"),
                           ("Reservation_Value_NGN", "Currency ₦", "Room_Rate_NGN x Nights", 285000)]:
        note("AtlanticPearl_Reservations_2023_2024", c_, t_, d_, e_)
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
    """2023 EMR export as received: the clean 2023 episodes plus re-sent duplicates and bill reversals, then keying defects."""
    r = np.random.default_rng(SEED + 2)
    yr = adm[pd.to_datetime(adm.Admit_Date).dt.year == 2023].copy().reset_index(drop=True)
    def scramble(n, k):
        return "".join(ch.upper() if h(k + str(i), 3) == 0 else ch.lower() for i, ch in enumerate(n))
    d = pd.DataFrame({
        "Patient_Name": [scramble(nm, k) for nm, k in zip(yr.Patient_Name, yr.Patient_No)], "Age": yr.Age.astype(object),
        "Gender": yr.Gender.astype(object), "Blood_Group": yr.Blood_Group, "Diagnosis_Group": yr.Diagnosis_Group,
        "Admit_Date": pd.to_datetime(yr.Admit_Date).dt.strftime("%Y-%m-%d"),
        "Discharge_Date": pd.to_datetime(yr.Discharge_Date).dt.strftime("%Y-%m-%d"),
        "Consultant_Name": yr.Consultant_Name, "Ward": yr.Ward.astype(object), "Bed_No": yr.Bed_No,
        "HMO": yr.HMO.astype(object), "Bill_Amount_NGN": yr.Bill_Amount_NGN.round(2).astype(object),
        "Admission_Route": yr.Admission_Route.astype(object), "Medication": yr.Medication, "Lab_Result": yr.Lab_Result,
    })
    n = len(d)
    dup = r.choice(n, 112, replace=False)
    neg = r.choice(n, 34, replace=False)
    rev = d.loc[neg].copy()
    rev["Bill_Amount_NGN"] = [-float(v) for v in rev.Bill_Amount_NGN]
    d = pd.concat([d, d.loc[dup], rev], ignore_index=True)
    d = d.sample(frac=1, random_state=SEED + 2).sort_values("Admit_Date", kind="stable").reset_index(drop=True)
    n = len(d)
    log = [("Random casing", "Patient_Name in mixed case e.g. 'cHIoma okAfor' (as exported by the old EMR)", n, "PROPER / Capitalize Each Word"),
           ("Source duplicates", "Exact duplicate episodes re-sent by the EMR interface", int(d.duplicated().sum()), "Remove Duplicates"),
           ("Negative bills", "Bill_Amount_NGN below zero (reversals keyed as episodes)", 34, "Filter out & report to Finance")]
    idx = r.choice(n, int(n * 0.06), replace=False)
    wv = {"Iroko Ward": ["Iroko", "IROKO WARD", "Iroko Ward "], "Mahogany Ward": ["Mahogany", "Mahogony Ward", "MAHOGANY WARD"],
          "Obeche Ward": ["Obeche", "Obeche  Ward", "OBECHE WARD"], "Baobab Ward": ["Baobab", "Baobab ward", "BAOBAB WARD "]}
    d.loc[idx, "Ward"] = [wv[w][h(i, 3)] for i, w in zip(idx, d.loc[idx, "Ward"])]
    log.append(("Ward label variants", "Iroko / IROKO WARD / Mahogony Ward / trailing spaces", len(idx), "TRIM + mapping table"))
    idx = r.choice(n, int(n * 0.05), replace=False)
    d.loc[idx, "Gender"] = [{"Male": "M", "Female": "F"}[g] if i % 2 else g.lower() for i, g in zip(idx, d.loc[idx, "Gender"])]
    log.append(("Coded vs full labels", "Gender recorded as M/F/male/female", len(idx), "SWITCH / Replace Values"))
    idx = r.choice(n, int(n * 0.03), replace=False)
    d.loc[idx, "Admission_Route"] = [{"Emergency": "EMERG", "Urgent": "urgent", "Elective": "Elective "}[a] for a in d.loc[idx, "Admission_Route"]]
    log.append(("Abbreviations", "Admission_Route EMERG / urgent / 'Elective '", len(idx), "Standardise to Emergency/Urgent/Elective"))
    idx = r.choice(n, 70, replace=False); d.loc[idx, "Discharge_Date"] = None
    log.append(("Missing discharge date", "Patient still flagged in-house or date not keyed", 70, "Exclude from LOS; report count"))
    idx = r.choice(n, 25, replace=False)
    d.loc[idx, "Discharge_Date"] = [(pd.Timestamp(a) - pd.Timedelta(days=int(2 + h(i, 5)))).strftime("%Y-%m-%d") for i, a in zip(idx, d.loc[idx, "Admit_Date"])]
    log.append(("Discharge before admission", "Negative length of stay", 25, "Flag; correct from ward register"))
    idx = r.choice(n, 12, replace=False); d.loc[idx, "Age"] = [0 if i % 2 else 150 for i in idx]
    log.append(("Impossible ages", "Age 0 or 150 keyed in error", 12, "Data Validation 0-120; query records"))
    idx = r.choice(n, 40, replace=False); d.loc[idx, "HMO"] = None
    log.append(("Missing payer", "HMO blank", 40, "Label 'Self-pay / Unknown' only if policy allows"))
    idx = r.choice(n, int(n * 0.04), replace=False)
    idx = np.union1d(idx, np.flatnonzero(pd.to_numeric(d.Bill_Amount_NGN, errors="coerce").values < 0)[::5][:6])
    d.loc[idx, "Bill_Amount_NGN"] = [f"{float(v):,.2f}" for v in d.loc[idx, "Bill_Amount_NGN"]]
    log.append(("Thousands separator in text", "Bill_Amount_NGN stored as text '1,234,567.89'", len(idx), "Change Type with Locale / VALUE"))
    return d, pd.DataFrame(log, columns=["Defect", "What was injected", "Rows affected", "Expected fix"])


def dirty_crestview(perf):
    r = np.random.default_rng(SEED + 3)
    d = perf.drop(columns=["Grade", "Attendance_Band", "At_Risk"]).copy().astype(object)
    n = len(d)
    bad = r.choice(np.flatnonzero(perf.Final_Score.values == 100), 3, replace=False)
    d.loc[bad, "Final_Score"] = [101, 104, 110]                    # true score 100, keyed with an extra digit
    d = d.replace({"Not Recorded": None})
    blanks = int(d[["Lecturer_Rating", "Guardian_Education", "Commute_Band"]].isna().sum().sum())
    log = [("Out-of-range score", "Final_Score above 100 (keying errors on scores of 100)", 3, "Validate 0-100; cap at 100 after checking"),
           ("Missing categories", "Lecturer_Rating / Guardian_Education / Commute_Band blank", blanks, "Label 'Not Recorded'")]
    idx = r.choice(n, int(n * 0.06), replace=False)
    d.loc[idx, "Home_Internet"] = [{"Yes": ["Y", "yes", "YES"], "No": ["N", "no", "NO"]}[v][h(i, 3)] for i, v in zip(idx, d.loc[idx, "Home_Internet"])]
    log.append(("Yes/No variants", "Home_Internet as Y / yes / YES / N / no", len(idx), "UPPER+LEFT or mapping"))
    idx = r.choice(n, int(n * 0.05), replace=False); d.loc[idx, "Lecture_Attendance_Pct"] = [f"{v}%" for v in d.loc[idx, "Lecture_Attendance_Pct"]]
    log.append(("Percent as text", "Lecture_Attendance_Pct typed '84%'", len(idx), "VALUE(SUBSTITUTE())"))
    idx = r.choice(n, int(n * 0.04), replace=False); d.loc[idx, "Gender"] = [{"Male": "M", "Female": "F"}[g] for g in d.loc[idx, "Gender"]]
    log.append(("Coded gender", "Gender as M/F", len(idx), "Standardise"))
    idx = r.choice(n, int(n * 0.03), replace=False)
    typo = {"Computer Science": "Computer Sci.", "Economics": "Econs", "Accounting": "Acct", "Microbiology": "Micro-biology",
            "Mass Communication": "Mass Comm", "Civil Engineering": "Civil Engr", "Electrical Engineering": "Elect. Engineering"}
    d.loc[idx, "Department"] = [typo.get(v, v.upper()) for v in d.loc[idx, "Department"]]
    log.append(("Department variants", "Abbreviated / upper-case department names", len(idx), "Lookup to department master"))
    dup = r.choice(n, 45, replace=False)
    d = pd.concat([d, d.loc[dup]], ignore_index=True)
    log.append(("Duplicate students", "Same Matric_No appears twice (double registration)", 45, "COUNTIF(ID)>1 / Remove Duplicates"))
    idx = r.choice(len(d), 25, replace=False); d.loc[idx, "Matric_No"] = [" " + v + " " for v in d.loc[idx, "Matric_No"]]
    log.append(("Padded keys", "Matric_No with leading/trailing spaces (breaks lookups)", 25, "TRIM before VLOOKUP/XLOOKUP"))
    d = d.sample(frac=1, random_state=SEED).reset_index(drop=True)
    return d, pd.DataFrame(log, columns=["Defect", "What was injected", "Rows affected", "Expected fix"])


def dirty_hotel(hotel):
    r = np.random.default_rng(SEED + 4)
    d = hotel.copy().astype(object)
    n = len(d)
    d["Check_In"] = [x.strftime("%d/%m/%Y") for x in hotel.Check_In]
    d["Booked_On"] = [x.strftime("%d/%m/%Y") for x in hotel.Booked_On]
    log = [("Day-first text dates", "Check_In & Booked_On as dd/mm/yyyy text", n, "Change Type with Locale en-GB / DATE()")]
    idx = r.choice(n, 4, replace=False); d.loc[idx, "Check_In"] = "29/02/2023"
    log.append(("Invalid calendar date", "29/02/2023 (2023 is not a leap year)", 4, "Flag & query front desk"))
    idx = r.choice(n, int(n * 0.03), replace=False)
    d.loc[idx, "Sales_Channel"] = [{"Online Travel Agent": "OTA", "Walk-in / Travel Agent": "Offline", "Corporate": "corporate",
                                    "Complimentary": "Comp", "Airline Crew": "Aviation"}[v] for v in d.loc[idx, "Sales_Channel"]]
    log.append(("Channel codes", "OTA / Offline / corporate / Comp / Aviation", len(idx), "Mapping table"))
    idx = r.choice(n, int(n * 0.02), replace=False); d.loc[idx, "Reservation_Status"] = [v.upper() for v in d.loc[idx, "Reservation_Status"]]
    log.append(("Status casing", "CANCELLED / STAYED", len(idx), "PROPER"))
    zero = int((hotel.Room_Rate_NGN == 0).sum())
    log.append(("Zero room rate", "Room_Rate_NGN = 0 (complimentary stays) - genuine values", zero, "Separate complimentary stays before averaging"))
    idx = r.choice(n, 6, replace=False); d.loc[idx, "Children"] = [9, 10, 9, 10, 9, 10]
    log.append(("Extreme child count", "Children keyed as 9 or 10 in one room", 6, "Validate occupancy rules"))
    dup = r.choice(n, 120, replace=False)
    d = pd.concat([d, d.loc[dup]], ignore_index=True)
    log.append(("Duplicate reservations", "Same Reservation_ID twice (channel manager re-sent)", 120, "Remove duplicates on Reservation_ID"))
    d = d.drop(columns=["Reservation_Value_NGN", "Nights"])
    log.append(("Derived columns removed", "Nights and Reservation_Value_NGN must be rebuilt", 0, "Calculated columns"))
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
    idx = r.choice(n, int(n * 0.03), replace=False); d.loc[idx, "HbA1c_Pct"] = [f"{v}%" for v in d.loc[idx, "HbA1c_Pct"]]
    log.append(("Unit in value", "HbA1c_Pct typed '6.6%'", len(idx), "Strip % and convert"))
    log.append(("Not Recorded smoking", "Smoking_Status = 'Not Recorded' (genuine gap)", int((base.Smoking_Status == "Not Recorded").sum()), "Keep as its own category"))
    dup = r.choice(n, 80, replace=False)
    d = pd.concat([d, d.loc[dup]], ignore_index=True)
    log.append(("Duplicate screening IDs", "Screen_ID repeated", 80, "Remove duplicates"))
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
    lc23 = adm[pd.to_datetime(adm.Admit_Date).dt.year == 2023]
    lc23.to_csv(W2D / "Mercyland_Episodes_2023_pipe.txt", sep="|", index=False)
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
    dlc, llc = dirty_lifecare_2023(None, adm)
    stage("dirty_lc", dlc); stage("dirty_lc_log", llc)
    dcv, lcv = dirty_crestview(pd.read_pickle(STAGE / "cv_perf.pkl"))
    stage("dirty_cv", dcv); stage("dirty_cv_log", lcv)
    dho, lho = dirty_hotel(hotel)
    stage("dirty_hotel", dho); stage("dirty_hotel_log", lho)
    ddi, ldi = dirty_diabetes(pd.read_pickle(STAGE / "lc_diab.pkl"))
    stage("dirty_diab", ddi); stage("dirty_diab_log", ldi)
    week2_files(sales, stores, adm, dsm)
    dlc.to_csv(W2D / "Mercyland_Episodes_2023_DIRTY.csv", index=False)
    write_dictionary()
    print("Datasets built.")
    for r in LOG:
        print(f"  {r['Dataset'][:70]:70s} rows={r['Rows']:>7,}")


if __name__ == "__main__":
    main()
