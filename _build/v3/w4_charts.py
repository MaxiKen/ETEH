"""Week 4 edition-3 charts: real PowerPoint charts on the topics where a picture is the point (distributions, scatter,
residuals, seasonality, forecast bands, scenarios, storytelling). Every series reuses the figures the slide copy computes."""
import numpy as np
import pandas as pd

import w4_m1_2 as A
import w4_m3_4 as B
import w4_m5_6 as C
import w4_m7_10 as G
from slidekit import GREEN, NAVY, ORANGE, RED
from w4_fx import CV, EX, OV, Nm, P, SAL, skew

FM = '₦#,##0,,"m"'
FM1 = '₦#,##0.0,,"m"'

# 4.1.8 order values: equal ₦100k bins up to ₦1m, then one tail bin
_edges = list(range(0, 1_000_001, 100_000)) + [1e12]
_cnt = [int(((OV >= a) & (OV < b)).sum()) for a, b in zip(_edges[:-1], _edges[1:])]
_lab = [f"{int(a / 1000)}–{int(b / 1000)}k" for a, b in zip(_edges[:-2], _edges[1:-1])] + ["₦1m+"]

# 4.2.4 Pareto: customers in 20 equal groups (5% each), revenue per group and the running share
_ng = 20
_grp = np.array_split(A.CUST.values, _ng)
_gsum = [float(g.sum()) for g in _grp]
_gcum = list(np.cumsum(_gsum) / sum(_gsum))

# 4.3.6 scatter: hours studied against score
# 4.4.8 residuals against experience (simple model)
_res = pd.Series(B.S1["resid"]).values

# 4.5.11 forecast chart: 48 actual months, then 12 forecast months with the 95% band
_act = C.S.tolist()
_lab_m = [pd.Timestamp(m).strftime("%b %y") for m in C.MON.Month] + \
    [pd.Timestamp(year=2026, month=k, day=1).strftime("%b %y") for k in range(1, 13)]
_fc, _ci = C.HWF["fc"], C.HWF["ci"]
_nA, _nF = len(_act), len(_fc)
S_ACT = _act + [None] * _nF
S_FC = [None] * (_nA - 1) + [_act[-1]] + list(_fc)
S_HI = [None] * _nA + list(_fc + _ci)
S_LO = [None] * _nA + list(_fc - _ci)

# 4.7.x store attainment, sorted
_st = G.ST.sort_values("att")
_st_names = [n.split(" - ")[0] if n.count(" - ") else n for n in _st.Store_Name]
_below = [i for i, a in enumerate(_st.att) if a < 1]
_wi = list(_st.Store_ID).index(G.WORST.Store_ID)

CHARTS = {
    "4.1.8": dict(kind="histogram", cats=_lab, series=[("Orders", _cnt)], fmt="#,##0", legend=False, colors=[NAVY],
                  x_title="Order line value (₦)", title=f"Order values: long right tail (SKEW {skew(OV):.2f})"),
    "4.2.3": dict(kind="histogram", cats=[str(k) for k in A.EBINS.index], series=[("Students", A.EBINS.tolist())], fmt="#,##0", legend=False,
                  labels=True, title=f"Exam scores, bin width 5: peak at {A.EPEAK}"),
    "4.2.4": dict(kind="combo", cats=[f"{5 * (i + 1)}%" for i in range(_ng)], series=[("Net revenue", _gsum), ("Running share", _gcum)],
                  fmt=FM, y2_fmt="0%", x_title="Customers, largest first (5% groups)", gap=20,
                  title=f"{P(A.K80 / len(A.CUST), 0)} of customers bring 80% of revenue"),
    "4.3.6": dict(kind="scatter", points=[("Students", CV.Study_Hours_Wk.tolist(), EX.tolist())], trend=True, marker=3,
                  x_title="Hours studied per week", y_title="Final score", title=f"Hours studied and score: r = {B.R['Study_Hours_Wk']:.2f}"),
    "4.3.9": dict(kind="column", cats=["2023", "2024", "Both years"],
                  series=[(rc, [float(B.SPY.loc[(2023, rc), "mean"]), float(B.SPY.loc[(2024, rc), "mean"]), float(B.SPA[rc])])
                          for rc in ("Deluxe", "Standard Twin")],
                  fmt="0%", labels=True, label_fmt="0.0%", legend=True, colors=[NAVY, ORANGE],
                  title="Deluxe cancels less in each year, more overall"),
    "4.4.8": dict(kind="scatter", points=[("Residual", SAL.Years_of_Experience.tolist(), _res.tolist())], marker=3, fmt=FM,
                  x_title="Years of experience", y_title="Actual − predicted", title="Residuals bend: low at both ends, high in the middle"),
    "4.5.3": dict(kind="column", cats=C.MSHORT, series=[("Seasonal index", [float(C.SI[m]) for m in range(1, 13)])], fmt="0.00",
                  legend=False, point_colors=[(ORANGE if C.SI[m] >= 1.2 else None) for m in range(1, 13)],
                  title="Seasonal index, 2022-24 (1.00 = a normal month)"),
    "4.5.11": dict(kind="line", cats=_lab_m, series=[("Actual", S_ACT), ("Forecast", S_FC), ("Upper 95%", S_HI), ("Lower 95%", S_LO)],
                   fmt=FM, legend=True, colors=[NAVY, ORANGE, "B8C2CC", "B8C2CC"], h=2.6,
                   note="Actuals navy, forecast orange, band light grey; the forecast starts Jan 2026.",
                   title=f"2026 forecast {Nm(float(_fc.sum()), 0)}, with its 95% band"),
    "4.6.7": dict(kind="column", cats=list(C.SCP), series=[("Profit", list(C.SCP.values()))], fmt=FM, labels=True, label_fmt=FM1, legend=False,
                  point_colors=[RED if v < 0 else GREEN for v in C.SCP.values()], title="2025 profit under three scenarios"),
    "4.7.6": dict(kind="bar", cats=_st_names, series=[("Attainment", _st.att.tolist())], fmt="0%", legend=False,
                  point_colors=[RED if i == _wi else None for i in range(len(_st))],
                  title=f"{G.NBEL} of {len(G.ST)} stores missed target; {G.WORST.Store_Name.split(' - ')[0]} is furthest behind", h=3.0),
    "4.7.7": dict(kind="bar", cats=_st_names, series=[("Attainment", _st.att.tolist())], fmt="0%", legend=False, labels=_below, label_fmt="0%",
                  point_colors=[RED if a < 1 else "B8C2CC" for a in _st.att], title="Grey for context, red for the point", h=3.0),
    "4.7.8": dict(kind="column", cats=G.ZN.index.tolist(), series=[("Sales", G.ZN.tolist())], fmt=FM, legend=False, y_min=G.AX, colors=[RED],
                  title=f"Axis from {Nm(G.AX, 0)}: the gap looks {(G.ZN.iloc[0] - G.AX) / (G.ZN.iloc[-1] - G.AX):.0f}×", label="THE MISLEADING VERSION"),
}
ORDER = {"4.2.3": ["chart", "grid"], "4.3.6": ["chart", "grid"], "4.5.11": ["chart", "steps"]}
SPLIT = {"4.6.7": 0.6, "4.5.11": 0.56}


def attach(modules):
    n = 0
    for m in modules:
        for t in m["topics"]:
            c = CHARTS.get(t["tag"])
            if c:
                t["chart"] = c
                if t["tag"] in SPLIT:
                    t["split"] = SPLIT[t["tag"]]
                if t["tag"] in ORDER:
                    t["order"] = ORDER[t["tag"]]
                n += 1
    return n
