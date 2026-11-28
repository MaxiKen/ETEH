"""Week 2 edition-3 charts: Module 2.11 (PivotCharts) gets the charts it describes, drawn from the same pivots as the copy."""
import w2_m9_11 as M
from slidekit import NAVY, ORANGE
from w2_fx import CVF, SMF

FM = '₦#,##0,,"m"'
ZY, ZORD, YEARS, DEPT = M.ZY, M.ZORD, M.YEARS, M.DEPT
from w2_fx import SM  # noqa: E402  (the full 2022-2025 SuperMart table behind the 2.10-2.11 pivots)
_q = SM.assign(Q=SM.Order_Date.dt.quarter).pivot_table(index=SM.Order_Date.dt.year, columns="Q", values="Sales", aggfunc="sum").reindex(YEARS)
assert _q.notna().all().all() and abs(_q[4] - SMF["q4"].Q4).max() < 1
_d = DEPT.sort_values("Admissions", ascending=False)
_g = sorted(CVF["grades"].items())
Z25 = [float(ZY.loc[z, 2025]) for z in ZORD]

CH_ZY = dict(kind="column", cats=ZORD, series=[(str(y), [float(ZY.loc[z, y]) for z in ZORD]) for y in YEARS], fmt=FM, legend=True,
             title="Sales by zone and year", gap=40)
CH_DEPT = dict(kind="bar", cats=_d.Department.tolist(), series=[("Admissions", _d.Admissions.tolist())], fmt="#,##0", legend=False, labels=True,
               title="Mercyland admissions by department, 2023")
CH_GR = dict(kind="pie", cats=[g for g, _ in _g], series=[("Students", [v for _, v in _g])], title="Kingsway grades", legend=True)
CHARTS = {
    "2.11.1": dict(CH_ZY, label="THE PIVOTCHART: AXIS = ZONE, LEGEND = YEAR"),
    "2.11.2": dict(kind="column", cats=[str(y) for y in YEARS], series=[(f"Q{q}", _q[q].tolist()) for q in range(1, 5)], fmt=FM, legend=True,
                   colors=["B8C2CC", "93A3B3", "6B7B8C", ORANGE], title="Q4 (orange) is the biggest quarter every year"),
    "2.11.4": dict(kind="column", cats=ZORD, series=[("Sales", Z25)], fmt=FM, legend=False, title="FY2025 sales by zone (Clustered Column)"),
    "2.11.7": dict(kind="column", cats=[str(y) for y in YEARS], series=[("South-West", [float(SMF["sw_year"][y]) for y in YEARS])], fmt=FM,
                   labels=True, label_fmt='₦#,##0.0,,"m"', legend=False, title="Filtered to South-West: the pivot follows"),
    "2.11.8": dict(kind="column", cats=ZORD, series=[("Sales", Z25)], fmt=FM, labels=True, label_fmt='₦#,##0.0,,"m"', legend=False,
                   title="South-West leads FY2025 sales", label="FORMATTED FOR THE BOARD PACK"),
    "2.11.9": dict(CH_DEPT, label="RANKED DEPARTMENTS → BAR CHART"),
    "2.11.10": dict(kind="bar", cats=ZORD, series=[("Sales", Z25)], fmt=FM, labels=True, label_fmt='₦#,##0.0,,"m"', legend=False,
                    point_colors=[NAVY] * len(ZORD), title="FY2025, sorted largest first"),
}
ORDER = {"2.11.2": ["chart", "points"], "2.11.4": ["steps", "chart"], "2.11.7": ["chart", "points"], "2.11.10": ["chart", "steps"]}
MORE = {"2.11.11": [dict(title="Three PivotCharts, one per dataset", pill="IN PRACTICE",
                         chart=dict(kind="dashboard", full=True, h=4.95, label="ONE PIVOTCHART PER DATASET",
                                    charts=[dict(CH_ZY, title="SuperMart: sales by zone and year"), CH_DEPT, CH_GR], widths=[1.2, 1.1, 0.9]),
                         notes="Each chart type is chosen for its data: zones over years (column), ranked departments (bar), five grades of one whole (pie).")]}


def attach(modules):
    n = 0
    for m in modules:
        for t in m["topics"]:
            tag = t["tag"]
            if tag in CHARTS:
                t["chart"] = CHARTS[tag]
                if tag in ORDER:
                    t["order"] = ORDER[tag]
                n += 1
            if tag in MORE:
                t["more"] = MORE[tag]
    return n
