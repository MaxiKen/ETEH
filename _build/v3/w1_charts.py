"""Week 1 edition-3 chart: 1.3.9 shows its 'change over time' example as a real line chart (Mercyland daily census, 2023)."""
import pandas as pd

import w3_data

_c = w3_data.table("tCensus")
_c["Census_Date"] = pd.to_datetime(_c.Census_Date)
_c = _c[_c.Census_Date.dt.year == 2023]
_m = _c.groupby(_c.Census_Date.dt.month).Occupancy_Pct.mean()
_MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
_pk = int(_m.values.argmax())
CHARTS = {"1.3.9": dict(kind="line", cats=_MON, series=[("Occupancy", _m.tolist())], fmt="0%", legend=False, labels=[_pk], label_fmt="0.0%",
                        title=f"Mercyland occupancy peaked in {_MON[_pk]} 2023 ({_m.iloc[_pk]:.1%})", label="CHANGE OVER TIME → LINE CHART")}


def attach(modules):
    n = 0
    for m in modules:
        for t in m["topics"]:
            if t["tag"] in CHARTS:
                t["chart"] = CHARTS[t["tag"]]
                n += 1
    return n
