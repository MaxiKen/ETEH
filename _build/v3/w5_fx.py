"""Capstone (Volume 5) edition-3 figures. Every number on the capstone slides is computed here, mostly by importing the
instructor capstone builders (04_proj_i1/i2/i3, whose module-level code is the reference analysis behind the SOLUTION
workbooks) and the capstone section-examples data. Nothing is typed in by hand."""
from __future__ import annotations

import hashlib
import importlib
import sqlite3
import sys
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
from common import CORE, W3D  # noqa: E402
from wk_common import L, lc_adm  # noqa: E402

C1 = importlib.import_module("04_proj_i1")      # SuperMart margin review
C2 = importlib.import_module("04_proj_i2")      # Kingsway progression
C3 = importlib.import_module("04_proj_i3")      # Mercyland board pack


def _neg(s):
    return "−" + s.replace("-", "", 1) if s.startswith("₦-") or s.startswith("-") else s


def N(v):
    return _neg(f"₦{v:,.0f}")


def Nm(v, d=1):
    return _neg(f"₦{v / 1e6:,.{d}f}m")


def Nb(v, d=2):
    return _neg(f"₦{v / 1e9:,.{d}f}bn")


def P(v, d=1):
    return _neg(f"{v * 100:.{d}f}%")


def I(v):
    return f"{int(v):,}"


def pts(v, d=2):
    return f"{v:+.{d}f}".replace("-", "−") + " pts"


# ================================================================ Capstone 1: SuperMart
SM = C1.sm
A24, B25 = C1.a24, C1.b25
S24, S25, P24, P25 = (float(x) for x in (C1.S24, C1.S25, C1.P24, C1.P25))
G25, M24, M25 = float(C1.g25), float(C1.m24), float(C1.m25)
DM = (M25 - M24) * 100
ORD25 = int(C1.ORD25)
LINES, ORDERS_ALL = len(SM), SM.Order_ID.nunique()
SKUS, SUBCATS = SM.Product_ID.nunique(), SM.Sub_Category.nunique()
STORES = C1.stores
NSTORES, NZONES = len(STORES), STORES.Zone.nunique()
TARGET = float(C1.TARGET)
DIRTY, CTRL, DLOG = C1.dirty, C1.ctrl, C1.dlog.set_index("Defect")["Rows affected"]
BLANK = int(DIRTY.isna().all(axis=1).sum())
DUPS = len(DIRTY) - BLANK - int(CTRL["Clean row count"])
CLEAN = int(CTRL["Clean row count"])
CUST25 = int(CTRL["Distinct customers"])
LG = C1.lg.assign(att=lambda d: d.S25 / d.T25, m=lambda d: d.P25 / d.S25)
LOSS_ST = C1.loss_stores
N_ON_TGT = int((LG.att >= 1).sum())

# discount bands, FY2025
BANDS = C1.BANDS
BD25 = {k: C1.bd[(2025, k)] for k in BANDS}
BD24 = {k: C1.bd[(2024, k)] for k in BANDS}
DEEP = float(C1.deep25_loss)
N_DEEP = BD25["21-40%"][2] + BD25[">40%"][2]
# rate / mix
MIX = pd.DataFrame(C1.mix, columns=["Category", "sh24", "sh25", "mg24", "mg25", "rate", "mix"]).set_index("Category")
RATE, MIXE = float(MIX.rate.sum()), float(MIX.mix.sum())
# quarters
QM = C1.qm


def _depth(d):
    return float((d.Discount * d.Sales).sum() / d.Sales.sum())


QD = {(y, q): _depth(d[d.Quarter == q]) for y, d in ((2024, A24), (2025, B25)) for q in ("Q1", "Q2", "Q3", "Q4")}
DEPTH24, DEPTH25 = _depth(A24), _depth(B25)
FURN = MIX.loc["Furniture"]
# what-if: cap Furniture discounts at 20%
CAP = C1.CAP
WHATIF = {loss: C1.whatif(CAP, loss)[1] - P25 for loss in (0.0, 0.25, 0.5)}
N_CAP = C1.whatif(CAP, 0.0)[2]
CAP_M = float(C1.FACTS["cap_m"])
ALL_LO, ALL_HI = float(C1.all_lo), float(C1.all_hi)
FURN_DEEP = B25[(B25.Category == "Furniture") & (B25.Discount > CAP)]
# SKUs
SKU = C1.allsku
N_UNSOLD, N_DROPPED, N_LOSSSKU, LOSSVAL = C1.n_unsold25, C1.n_dropped, C1.n_loss, float(C1.loss_val)
_ps = SKU[SKU.Sales_2025 > 0].Sales_2025.sort_values()
K30 = int(len(_ps) * 0.3)
TAIL30 = float(_ps.iloc[:K30].sum() / _ps.sum())
SOLD25 = len(_ps)
# baskets
O25 = B25.groupby("Order_ID").agg(Lines=("Row_ID", "count"), Basket_Value=("Sales", "sum"), Categories=("Category", "nunique"))
BASK = dict(mean_lines=O25.Lines.mean(), mode_lines=int(O25.Lines.mode()[0]), mean_val=O25.Basket_Value.mean(),
            med_val=O25.Basket_Value.median(), multi=(O25.Categories > 1).mean())
SUBL = C1.subl
CATS = ["Furniture", "Office Supplies", "Technology"]
ATTACH = {}
for _c1 in CATS:
    _h1 = set(B25[B25.Category == _c1].Order_ID)
    ATTACH[_c1] = (len(_h1), {c2: len(_h1 & set(B25[B25.Category == c2].Order_ID)) / len(_h1) for c2 in CATS if c2 != _c1})
Q4M24, Q4M25 = QM[(2024, "Q4")], QM[(2025, "Q4")]

# ================================================================ Capstone 2: Kingsway
CV = C2.cv
RAW2 = C2.raw
NCV, NRAW2 = len(CV), len(RAW2)
DLOG2 = C2.dlog.set_index("Defect")["Rows affected"]
ST, AT, TGT = C2.ST, C2.AT, 0.90
RISK = C2.risk
NRISK = int(RISK.sum())
PROG = 1 - RISK.mean()
NPROG = NCV - NRISK
TRIG = pd.Series(C2.trig[RISK.values]).value_counts()
FAC = C2.FAC
FAC_N = CV.groupby("Faculty").size()
LVL = C2.LVL
BAND_RATE, BAND_N = C2.band_rate, C2.band_n
CORR = C2.corr
BETA = dict(zip(C2.XP + ["const"], C2.beta))
R2 = float(C2.r2)
LIFT, RESCUED, MECH = C2.LIFT, int(C2.rescued), int(C2.mech)
PROG_AFTER = (NPROG + RESCUED) / NCV
NLOWATT = int((CV.Lecture_Attendance_Pct < AT).sum())
TOP_SHARE = float(C2.TOP_SHARE)
N_DEPT2 = CV.Department.nunique()
N_FAC2 = CV.Faculty.nunique()
SLOPE_ATT = float(np.polyfit(CV.Lecture_Attendance_Pct, CV.Final_Score, 1)[0])
LST = C2.lst

# ================================================================ Capstone 3: Mercyland
ADM, CEN = C3.adm, C3.cen
YEAR = C3.YEAR
C23 = C3.c23
OCC, D95, D100 = float(C3.occ), C3.d95, C3.d100
BEDS = int(C23.open_beds.mean())
BEDS_NEED = int(np.ceil(C3.beds_needed))
ALOS, MEDLOS, RR = float(C3.alos), C3.med_los, float(C3.rr)
BILL = float(C3.bill)
MOCC = C3.mocc
N_MOVER = C3.n_over
DEPT_CI = C3.dept_ci
ABOVE, BELOW = C3.above, C3.below
LOSB = [(lab, int(ADM.los_days.between(a, z).sum())) for a, z, lab in C3.LOSB]
LONG, N_LONG, SH_LONG, BD_LONG = C3.LONG, C3.n_long, float(C3.sh_long), float(C3.bd_long)
LOSMAX = C3.LOSMAX
WARD_LOS = ADM.groupby("ward").los_days.agg(["count", "mean", "median"]).sort_values("mean", ascending=False)
WARD_P90 = ADM.groupby("ward").los_days.quantile(0.9)
TYPE_LOS = ADM.groupby("admission_route").los_days.agg(["count", "mean"]).reindex(C3.TYPES)
HMO = ADM.groupby("hmo").bill_amount_ngn.sum().sort_values(ascending=False) / BILL
CTL = C3.ctl
QROWS = {k: len(v) for k, v in C3.R.items()}
QNAMES = list(C3.Q)
NWARDS = len(C3.WARDS)
WARDS_CSV = pd.read_csv(CORE / "Mercyland_Wards.csv")
NCONS = len(pd.read_csv(CORE / "Mercyland_Consultants.csv"))
EP_ALL = len(lc_adm())
CEN_DAYS = len(L("lc_census"))
PK_MON = max(MOCC, key=MOCC.get)
# escalation rule (capstone workbook P.5.7): first matching reason, Q4 of the year
_lc = lc_adm()
Q4ADM = _lc[(_lc.Admit_Date >= f"{YEAR}-10-01") & (_lc.Admit_Date <= f"{YEAR}-12-31")]
_reason = np.where(Q4ADM.Readmitted_30d == "Yes", "Readmitted", np.where(Q4ADM.LOS_Days >= 28, "Long stay",
                                                                          np.where(Q4ADM.Age >= 75, "Age 75+", "")))
ESC = pd.Series(_reason).value_counts().drop("", errors="ignore")
NQ4 = len(Q4ADM)

# ================================================================ P.2 source register (fingerprints of the core files)
REG_FILES = {"SuperMart_Stores_Sales_2022_2025.csv": "Finance", "SuperMart_Stores_Directory.csv": "Operations",
             "Mercyland_Inpatient_Episodes_2019_2024.csv": "Health Records", "Mercyland_Daily_Bed_Census.csv": "Bed Management",
             "Kingsway_Student_Results.csv": "Registry", "Kingsway_Student_Welfare_Survey_SENSITIVE.csv": "Student Affairs"}
REG = []
for _fn, _own in REG_FILES.items():
    _raw = (CORE / _fn).read_bytes()
    REG.append((_fn, _own, _raw.count(b"\n") - 1, hashlib.sha256(_raw).hexdigest()[:12]))

# before/after evidence from the dirty SuperMart returns (first example of each defect type)
_smi = SM.set_index("Row_ID")
_dd = DIRTY.dropna(subset=["Row_ID"])
BA = []
for _col, _test in [("Sales", lambda v: isinstance(v, str)), ("Discount", lambda v: isinstance(v, str)),
                    ("Order_Date", lambda v: isinstance(v, str) and not str(v)[:4].isdigit()),
                    ("Store_Name", lambda v: isinstance(v, str) and (v.isupper() or v.islower()))]:
    _hit = _dd[_dd[_col].map(_test)].head(1)
    if len(_hit):
        _rid = int(_hit.Row_ID.iloc[0])
        _after = _smi.loc[_rid, _col]
        _after = _after.strftime("%Y-%m-%d") if _col == "Order_Date" else (f"{_after:g}" if _col == "Discount" else _after)
        BA.append((_col, str(_hit[_col].iloc[0]), str(_after) if _col != "Sales" else f"{int(_after)} (number)"))

# ================================================================ P.9 SuperMart Online SQL screen + the STAR story
DB = W3D / "supermart_online.db"
SQL_EASY = ("SELECT payment_method, COUNT(*) AS orders,\n       SUM(total_price) AS revenue\nFROM transactions\n"
            "WHERE order_status = 'Delivered'\nGROUP BY payment_method\nORDER BY revenue DESC;")
SQL_MED = ("SELECT p.brand_name, COUNT(*) AS orders,\n       SUM(t.total_price) AS net_revenue\nFROM transactions t\n"
           "JOIN products p ON p.product_id = t.product_id  -- many-to-one\nWHERE t.order_status IN ('Delivered','Completed')\n"
           "GROUP BY p.brand_name\nORDER BY net_revenue DESC\nLIMIT 5;")
SQL_HARD = ("WITH s AS (\n  SELECT sl.store_zone, sl.store_name, SUM(t.total_price) AS revenue,\n"
            "         RANK() OVER (PARTITION BY sl.store_zone\n                      ORDER BY SUM(t.total_price) DESC) AS rnk\n"
            "  FROM transactions t JOIN sellers sl ON sl.seller_id = t.seller_id\n  GROUP BY sl.store_zone, sl.store_name)\n"
            "SELECT store_zone, store_name, revenue FROM s WHERE rnk = 1;")
with sqlite3.connect(DB) as _con:
    EASY = pd.read_sql_query(SQL_EASY, _con)
    MED = pd.read_sql_query(SQL_MED, _con)
    HARD = pd.read_sql_query(SQL_HARD, _con)
    TX_N, TX_TOT = _con.execute("SELECT COUNT(*), SUM(total_price) FROM transactions").fetchone()
    BAD_N, BAD_TOT = _con.execute("SELECT COUNT(*), SUM(t.total_price) FROM transactions t "
                                  "JOIN deliveries d ON d.customer_id = t.customer_id").fetchone()
    OK_N, OK_TOT = _con.execute("SELECT COUNT(*), SUM(t.total_price) FROM transactions t "
                                "JOIN deliveries d ON d.delivery_id = t.delivery_id").fetchone()
    SMO_TABLES = [r[0] for r in _con.execute("SELECT name FROM sqlite_master WHERE type='table'")]
with sqlite3.connect(C3.DB) as _con:
    MH_TABLES = [r[0] for r in _con.execute("SELECT name FROM sqlite_master WHERE type='table'")]

# live Excel test (capstone workbook P.9.2)
_st25 = B25.groupby("Store_Name").agg(S=("Sales", "sum"), P=("Profit", "sum"))
_st25["m"] = _st25.P / _st25.S
XL_SS = float(B25[B25.Zone == "South-South"].Sales.sum())
XL_TOPM = (_st25.m.idxmax(), float(_st25.m.max()))
XL_D40 = float((B25.Discount >= 0.4).mean())

# P.6 reconciliation tests (capstone workbook tRecon)
RECON = [("SuperMart lines, all years", LINES, LINES), ("FY2025 sales (₦)", S25, float(CTRL["Total Sales (₦)"])),
         ("FY2025 profit (₦)", P25, float(CTRL["Total Profit (₦)"])), ("FY2025 quantity", float(B25.Quantity.sum()), float(CTRL["Total Quantity"])),
         ("Kingsway cohort", NCV, len(L("cv_perf"))), (f"Mercyland {YEAR} admissions", len(ADM), float(CTL.loc[str(YEAR), "admissions"])),
         ("Bed census days", CEN_DAYS, CEN_DAYS)]
SMALL_STORE = B25.Store_Name.value_counts()
