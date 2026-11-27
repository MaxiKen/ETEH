"""Helpers for Week 2 slide copy: every number on a slide is computed here from the data (never typed)."""
from __future__ import annotations

import os
import re

import numpy as np
import pandas as pd

import w2_data as D

S = D.s25()                       # clean FY2025 lines (tSales25)
TS = D.table("tSales25")
ST = D.table("tStores")
A = D.table("tAdm23")
DR = D.dirty_raw()                # SuperMart_Sales_2025_DIRTY.csv as text
LC = D.lc_dirty_raw()             # Lifecare_Admissions_2023_DIRTY.csv as text
PERF = D.table("tStorePerf").set_index("Store_ID")


def N(v, d=0):
    v = float(v)
    return (f"-₦{-v:,.{d}f}" if v < 0 else f"₦{v:,.{d}f}")


def P(v, d=1):
    return f"{v:.{d}%}"


def I(v):
    return f"{int(round(v)):,}"


def Nm(v, d=1):
    return f"₦{v / 1e6:,.{d}f}m"


# ---------------------------------------------------------------- dirty SuperMart facts (from the file itself)
def _is(fmt):
    return DR.Order_Date.map(lambda v: bool(re.fullmatch(fmt, v)))


DF = dict(
    rows=len(DR), cols=DR.shape[1], dup_all=int(DR.duplicated().sum()), blank_rows=int((DR == "").all(axis=1).sum()),
    sales_text=int(DR.Sales.str.startswith("₦").sum()), disc_text=int(DR.Discount.str.endswith("%").sum()),
    qty_blank=int(((DR.Quantity == "") & (DR.Order_ID != "")).sum()), qty_neg=int(DR.Quantity.str.startswith("-").sum()),
    na_cust=int((DR.Customer_Name == "N/A").sum()), nbsp=int(DR.Customer_Name.str.contains("\u00a0").sum()),
    lf=int(DR.Customer_Name.str.contains("\n").sum()),
    zone_space=int(((DR.Zone != DR.Zone.str.strip()) & (DR.Zone != "")).sum()),
    store_case=int((~DR.Store_Name.isin(ST.Store_Name) & (DR.Store_Name != "")).sum()),
    cat_bad=int((~DR.Category.isin(["Furniture", "Office Supplies", "Technology"]) & (DR.Category != "")).sum()),
    d_dots=int(_is(r"\d{4}\.\d{2}\.\d{2}").sum()), d_slash=int(_is(r"\d{2}/\d{2}/\d{4}").sum()),
    d_mon=int(_is(r"\d{2}-[A-Za-z]{3}-\d{4}").sum()),
)
DF["text_dates"] = DF["d_dots"] + DF["d_slash"] + DF["d_mon"]
DF["cat_counts"] = DR.Category.value_counts().to_dict()


def _parse(v):
    for f in ("%Y-%m-%d", "%Y.%m.%d", "%d/%m/%Y", "%d-%b-%Y"):
        try:
            return pd.to_datetime(v, format=f)
        except (ValueError, TypeError):
            pass
    return pd.NaT


_od = DR.Order_Date.map(_parse)
_sd = pd.to_datetime(DR.Ship_Date, format="%Y-%m-%d", errors="coerce")
DF["ship_before"] = int((_sd < _od).sum())
_sv = pd.to_numeric(DR.Sales, errors="coerce")
_truth = S.set_index("Row_ID").Sales
_rid = pd.to_numeric(DR.Row_ID, errors="coerce")
_x100 = (_sv.notna() & _rid.notna()) & (_sv == _rid.map(_truth) * 100)
DF["x100"] = int(_x100.sum())
DF["x100_rows"] = sorted(set(int(r) for r in _rid[_x100]))
DF["top_sales"] = [int(v) for v in _sv.nlargest(3)]

# ---------------------------------------------------------------- dirty Lifecare facts
_bill = pd.to_numeric(LC.Billing_Amount.str.replace(",", ""), errors="coerce")
LF = dict(
    rows=len(LC), cols=LC.shape[1], dups=int(LC.duplicated().sum()), bill_text=int(LC.Billing_Amount.str.contains(",").sum()),
    neg_bill=int((_bill < 0).sum()), age0=int((LC.Age == "0").sum()), age150=int((LC.Age == "150").sum()),
    dis_blank=int((LC.Discharge_Date == "").sum()), hmo_blank=int((LC.HMO == "").sum()),
    dis_before=int((pd.to_datetime(LC.Discharge_Date.replace("", None)) < pd.to_datetime(LC.Admission_Date)).sum()),
    gender=LC.Gender.value_counts().to_dict(), ward=LC.Ward.value_counts().to_dict(), atype=LC.Admission_Type.value_counts().to_dict(),
)


def _std(d):
    d = d.copy()
    d["Patient_Name"] = d.Patient_Name.str.title()
    wm = {"emerald": "Emerald Ward", "sapphire": "Sapphire Ward", "saphire": "Sapphire Ward", "ruby": "Ruby Ward", "topaz": "Topaz Ward"}
    d["Ward"] = d.Ward.map(lambda w: wm[w.strip().lower().split()[0]])
    d["Gender"] = d.Gender.str.strip().str.upper().str[0].map({"M": "Male", "F": "Female"})
    d["Admission_Type"] = d.Admission_Type.str.strip().str.lower().map(
        {"emerg": "Emergency", "emergency": "Emergency", "urgent": "Urgent", "elective": "Elective"})
    d["Billing_Amount"] = pd.to_numeric(d.Billing_Amount.str.replace(",", ""))
    return d


LCS = _std(LC)
LF["dups_std"] = int(LCS.duplicated().sum())
LF["after_dedupe"] = len(LCS.drop_duplicates())
LCC = LCS.drop_duplicates()
LCC = LCC[LCC.Billing_Amount >= 0]
LF["clean_rows"] = len(LCC)
LF["ward_variants"] = int((~LC.Ward.isin(["Emerald Ward", "Sapphire Ward", "Ruby Ward", "Topaz Ward"])).sum())
LF["gender_variants"] = int((~LC.Gender.isin(["Male", "Female"])).sum())
LF["atype_variants"] = int((~LC.Admission_Type.isin(["Emergency", "Urgent", "Elective"])).sum())

CTRL = dict(rows=len(S), sales=int(S.Sales.sum()), profit=int(S.Profit.sum()), qty=int(S.Quantity.sum()),
            orders=S.Order_ID.nunique(), customers=D.sales()[D.sales().Year == 2025].Customer_ID.nunique())


# ---------------------------------------------------------------- extra facts used by modules 2.5-2.7
_nb = DR[~(DR == "").all(axis=1)]
DF["after_blank"] = len(_nb)
DF["dup_lines"] = int(_nb.duplicated().sum())
_dupv = _nb[_nb.duplicated()]
DF["dup_sales"] = float(_dupv.Row_ID.astype(int).map(_truth).sum())
_txt = _nb[_nb.Sales.str.startswith("₦")]
DF["text_sales_value"] = float(_txt.Sales.str.replace("[₦,]", "", regex=True).astype(float).sum())
DF["text_sales_example"] = _txt.Sales.iloc[0]
DF["text_sales_clean"] = float(_txt.drop_duplicates().Sales.str.replace("[₦,]", "", regex=True).astype(float).sum())
DF["tech_partial"] = int(DR.Category.str.lower().str.contains("tech").sum())
DF["zone_distinct"] = int(_nb.Zone.nunique())
DF["zone_distinct_trim"] = int(_nb.Zone.str.strip().nunique())
_sl = _nb[_nb.Order_Date.str.fullmatch(r"\d{2}/\d{2}/\d{4}")].Order_Date
_d, _m = _sl.str[:2].astype(int), _sl.str[3:5].astype(int)
DF["us_err"] = int((_d > 12).sum())                       # US reads mm/dd: a day above 12 is an invalid month -> error
DF["us_wrong"] = int(((_m <= 12) & (_d <= 12) & (_d != _m)).sum())
DF["us_ok"] = int((_d == _m).sum())
loss = S[S.Profit < 0]
F = dict(loss_lines=len(loss), loss_sales=float(loss.Sales.sum()), loss_profit=float(loss.Profit.sum()))
_band = pd.cut(S.Discount, [-1, 0, 0.2, 1], labels=["None", "1-20%", "Above 20%"])
F["bands"] = S.groupby(_band, observed=True).agg(Lines=("Sales", "size"), Sales=("Sales", "sum"), Profit=("Profit", "sum"))
F["zone"] = S.groupby("Zone").agg(Sales=("Sales", "sum"), Profit=("Profit", "sum")).sort_values("Sales", ascending=False)
F["zone"]["Margin"] = F["zone"].Profit / F["zone"].Sales
_c = S[S.Customer_ID == "SMC-0388"] if "Customer_ID" in S else None
F["smc0388"] = (len(_c), _c.Order_Date.min(), _c.Order_Date.max()) if _c is not None else None
F["first_rows"] = S.sort_values(["Order_Date", "Row_ID"]).head(3)[["Row_ID", "Order_Date", "Sales", "Profit"]].to_dict("records")
F["last_row"] = S.sort_values(["Order_Date", "Row_ID"]).iloc[-1][["Row_ID", "Order_Date"]].to_dict()


# ---------------------------------------------------------------- module 2.9: the Lifecare cleaning query, replayed in pandas
def _lc_query():
    d = LC.drop_duplicates().copy()
    out = dict(dedup=len(d))
    w = d.Ward.str.replace("  ", " ").str.replace("Saphire", "Sapphire").str.strip().str.lower()
    d["Ward"] = np.select([w.str.startswith(k) for k in ("emerald", "sapphire", "ruby", "topaz")],
                          ["Emerald Ward", "Sapphire Ward", "Ruby Ward", "Topaz Ward"], "CHECK")
    d["Gender"] = np.where(d.Gender.str.lower().isin(["m", "male"]), "Male", "Female")
    t = d.Admission_Type.str.strip().str.lower()
    d["Admission_Type"] = np.where(t.str.startswith("emerg"), "Emergency", t.str.title())
    d["Billing_Amount"] = pd.to_numeric(d.Billing_Amount.str.replace(",", ""))
    d["Age"] = d.Age.astype(int)
    ad = pd.to_datetime(d.Admission_Date)
    dd = pd.to_datetime(d.Discharge_Date.replace("", None))
    d["LOS"] = (dd - ad).dt.days
    exc = np.select([d.Billing_Amount < 0, dd.isna(), d.LOS < 0, (d.Age == 0) | (d.Age > 120)],
                    ["Negative bill", "No discharge date", "Discharge before admission", "Impossible age"], "")
    out["exceptions"] = pd.Series(exc[exc != ""]).value_counts().to_dict()
    out["clean"] = int((exc == "").sum())
    out["ward_check"] = int((d.Ward == "CHECK").sum())
    out["ruby_raw"] = int((LC.drop_duplicates().Ward == "Ruby Ward").sum())
    out["ruby_map"] = int((d.Ward == "Ruby Ward").sum())
    out["gender"] = d.Gender.value_counts().to_dict()
    out["atype"] = d.Admission_Type.value_counts().to_dict()
    out["ward_raw_n"] = LC.Ward.nunique()
    out["gender_raw_n"] = LC.Gender.nunique()
    out["atype_raw_n"] = LC.Admission_Type.nunique()
    nm = LC.Patient_Name[LC.Patient_Name.str.contains(r"[a-z][A-Z]")]
    out["name_example"] = nm.iloc[0] if len(nm) else None
    out["bill_example"] = LC.Billing_Amount[LC.Billing_Amount.str.contains(",")].iloc[0]
    return out


LQ = _lc_query()

CV = pd.read_csv(os.path.join(os.path.dirname(D.__file__), "..", "ETEH_Bootcamp_Pack", "01_Section_Examples", "00_Core_Datasets",
                              "Crestview_Student_Performance.csv"))
CVF = dict(n=len(CV), bands=pd.cut(CV.Attendance_Pct, [-1, 74.999, 89.999, 1000], labels=["Below 75%", "75-89%", "90%+"])
           .value_counts().to_dict(), mentor=int(((CV.Exam_Score < 60) & (CV.Attendance_Pct < 75)).sum()),
           grades=CV.Grade.value_counts(normalize=True).to_dict())

# ---------------------------------------------------------------- modules 2.10-2.11: the full SuperMart table tSM
SM = D.sales()
SMF = dict(rows=len(SM), total=float(SM.Sales.sum()), cols=SM.shape[1])
SMF["zy"] = SM.pivot_table(index="Zone", columns="Year", values="Sales", aggfunc="sum")
_old = SM[SM.Year < 2025]
SMF["fixed"] = (len(_old), float(_old.Sales.sum()))
SMF["q25"] = S.groupby(S.Order_Date.dt.quarter).Sales.sum().to_dict()
SMF["sizes"] = pd.cut(S.Sales, [0, 49999.99, 199999.99, 999999.99, 1e12],
                      labels=["Under ₦50k", "₦50k-200k", "₦200k-1m", "₦1m and above"]).value_counts().sort_index().to_dict()
SMF["zone_cnt"] = S.groupby("Zone").Sales.agg(["count", "sum"])
_zy = SMF["zy"]
SMF["share25"] = (_zy[2025] / _zy.sum(axis=1)).sort_values(ascending=False)
SMF["diff25"] = (_zy[2025] - _zy[2024])
_sp = S.groupby("Store_Name").agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"))
SMF["store"] = _sp.sort_values("Sales", ascending=False)
SMF["loss_stores"] = _sp.sort_values("Profit").head(2)
_jn = S[S.Order_Date.dt.month < 12]
SMF["jan_nov"] = (len(_jn), float(_jn.Sales.sum()))
SMF["dec"] = (len(S) - len(_jn), float(S.Sales.sum() - _jn.Sales.sum()))
SMF["avg_line_margin"] = float((S.Profit / S.Sales).mean())
_q = SM.groupby(["Year", SM.Order_Date.dt.quarter]).Sales.sum().unstack()
SMF["q4"] = pd.DataFrame({"Q4": _q[4], "share": _q[4] / _q.sum(axis=1)})
SMF["year"] = SM.groupby("Year").Sales.sum()
SMF["sw_year"] = _zy.loc["South-West"]
SMF["store_month"] = S.groupby(["Store_Name", S.Order_Date.dt.strftime("%Y-%m")]).agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"))
SMF["orders"] = S.Order_ID.nunique()
