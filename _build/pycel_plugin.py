"""Extra Excel functions for pycel so verify_xlsx.py can check more formulas independently.
Only used for verification - never part of the delivered workbooks."""
from __future__ import annotations

import datetime as dt
import math
import statistics

import numpy as np


def _flat(x):
    if isinstance(x, (tuple, list)):
        out = []
        for v in x:
            out.extend(_flat(v))
        return out
    return [x]


def _nums(x):
    return [v for v in _flat(x) if isinstance(v, (int, float)) and not isinstance(v, bool)]


def _col(x):
    return [r[0] if isinstance(r, (tuple, list)) else r for r in x] if isinstance(x, (tuple, list)) else [x]


def proper(text):
    return str(text).title()


def exact(a, b):
    return str(a) == str(b)


def rows(x):
    return len(x) if isinstance(x, (tuple, list)) else 1


def columns(x):
    return len(x[0]) if isinstance(x, (tuple, list)) and isinstance(x[0], (tuple, list)) else 1


def countblank(x):
    return sum(1 for v in _flat(x) if v is None or v == "")


def median(*args):
    v = _nums(args)
    return statistics.median(v)


def mode_sngl(*args):
    v = _nums(args)
    return max(sorted(set(v)), key=v.count)


def textbefore(text, delim, instance=1, *a):
    s = str(text)
    parts = s.split(delim)
    if instance < 0:
        return delim.join(parts[:instance])
    return delim.join(parts[:instance])


def textafter(text, delim, instance=1, *a):
    s = str(text)
    parts = s.split(delim)
    if instance < 0:
        return delim.join(parts[instance:])
    return delim.join(parts[instance:])


def textjoin(delim, ignore_empty, *args):
    v = [str(x) for x in _flat(args) if not (ignore_empty and (x is None or x == ""))]
    return delim.join(v)


def xlookup(key, look, ret, if_not_found=None, *a):
    lk, rt = _col(look), _col(ret)
    for i, v in enumerate(lk):
        if str(v).lower() == str(key).lower():
            return rt[i]
    return if_not_found if if_not_found is not None else "#N/A"


def numbervalue(text, dec=".", grp=","):
    s = str(text).replace(grp, "").replace("\u00a0", "").replace(" ", "")
    return float(s.replace(dec, "."))


def _d(x):
    if isinstance(x, (int, float)):
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(x))
    return x


def _ser(d):
    return (d - dt.date(1899, 12, 30)).days


def datedif(a, b, unit):
    a, b = _d(a), _d(b)
    if unit.lower() == "d":
        return (b - a).days
    raise ValueError


def networkdays(a, b, *h):
    a, b = _d(a), _d(b)
    return int(np.busday_count(a, b + dt.timedelta(days=1)))


def workday(a, n, *h):
    return _ser(np.busday_offset(_d(a), int(n), roll="backward").astype(dt.date))


def datevalue(s):
    for f in ("%d-%b-%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return _ser(dt.datetime.strptime(str(s), f).date())
        except ValueError:
            pass
    return "#VALUE!"


def forecast_linear(x, ys, xs):
    yv, xv = _nums(ys), _nums(xs)
    b, a = np.polyfit(xv, yv, 1)
    return float(a + b * x)


def formulatext(ref):
    raise NotImplementedError


def isna(x):
    return isinstance(x, str) and x == "#N/A"


def correl(a, b):
    return float(np.corrcoef(_nums(a), _nums(b))[0, 1])


def stdev_s(*args):
    return statistics.stdev(_nums(args))


def stdev_p(*args):
    return statistics.pstdev(_nums(args))


def var_s(*args):
    return statistics.variance(_nums(args))


def percentile_inc(arr, k):
    return float(np.percentile(_nums(arr), k * 100))


def quartile_inc(arr, q):
    return float(np.percentile(_nums(arr), q * 25))


def skew(*args):
    v = np.array(_nums(args), dtype=float)
    n = len(v)
    s = v.std(ddof=1)
    return float(n / ((n - 1) * (n - 2)) * (((v - v.mean()) / s) ** 3).sum())


def kurt(*args):
    v = np.array(_nums(args), dtype=float)
    n = len(v)
    s = v.std(ddof=1)
    m4 = (((v - v.mean()) / s) ** 4).sum()
    return float(n * (n + 1) / ((n - 1) * (n - 2) * (n - 3)) * m4 - 3 * (n - 1) ** 2 / ((n - 2) * (n - 3)))


def standardize(x, mean, sd):
    return (x - mean) / sd


def rsq(ys, xs):
    return correl(ys, xs) ** 2


def slope(ys, xs):
    return float(np.polyfit(_nums(xs), _nums(ys), 1)[0])


def intercept(ys, xs):
    return float(np.polyfit(_nums(xs), _nums(ys), 1)[1])


def norm_dist(x, m, s, cum):
    z = (x - m) / s
    return 0.5 * (1 + math.erf(z / math.sqrt(2))) if cum else math.exp(-z * z / 2) / (s * math.sqrt(2 * math.pi))


def norm_s_dist(z, cum=True):
    return norm_dist(z, 0, 1, cum)


def rank_avg(x, ref, order=0):
    v = _nums(ref)
    gt = sum(1 for a in v if (a > x if not order else a < x))
    eq = sum(1 for a in v if a == x)
    return gt + (eq + 1) / 2


def rank_eq(x, ref, order=0):
    v = _nums(ref)
    return 1 + sum(1 for a in v if (a > x if not order else a < x))


def percentrank_inc(arr, x, sig=3):
    v = sorted(_nums(arr))
    r = sum(1 for a in v if a < x)
    f = 10 ** int(sig)
    return math.floor(r / (len(v) - 1) * f) / f


def t_dist_2t(x, df):
    from scipy import stats
    return float(2 * stats.t.sf(abs(x), df))


def confidence_t(alpha, sd, n):
    from scipy import stats
    return float(stats.t.ppf(1 - alpha / 2, n - 1) * sd / math.sqrt(n))


def stdev(*args):
    return float(np.std(_nums(args), ddof=1))
