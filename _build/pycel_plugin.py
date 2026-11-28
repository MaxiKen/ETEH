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


def _hol(h):
    out = []
    for x in h:
        for v in (x if isinstance(x, tuple) else (x,)):
            for w in (v if isinstance(v, tuple) else (v,)):
                if w not in (None, ""):
                    out.append(_d(w))
    return out


def networkdays(a, b, *h):
    a, b = _d(a), _d(b)
    return int(np.busday_count(a, b + dt.timedelta(days=1), holidays=_hol(h)))


def workday(a, n, *h):
    return _ser(np.busday_offset(_d(a), int(n), roll="forward" if int(n) < 0 else "backward", holidays=_hol(h)).astype(dt.date))


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


# ---------------------------------------------------------------- Week 3 additions (verification only)
def counta(*args):
    """COUNTA: non-empty cells (pycel has no COUNTA)."""
    return sum(1 for v in _flat(args) if v is not None and v != "")


def unique(arr, by_col=False, exactly_once=False):
    """UNIQUE on a single column or row: first-seen order, returned as a column."""
    vals = [v for v in _flat(arr)]
    seen, out = {}, []
    for v in vals:
        k = v.lower() if isinstance(v, str) else v
        seen[k] = seen.get(k, 0) + 1
    done = set()
    for v in vals:
        k = v.lower() if isinstance(v, str) else v
        if k in done or (exactly_once and seen[k] != 1):
            continue
        done.add(k)
        out.append(v)
    return tuple((v,) for v in out)


def _is_arr(x):
    return isinstance(x, (tuple, list)) and len(x) and isinstance(x[0], (tuple, list))


def countif(rng, criteria):
    """COUNTIF; with an array of plain values as criteria (COUNTIF(A, B) inside SUMPRODUCT) returns an array of counts,
    worked out with one pass (case-insensitive, as Excel). Other forms fall back to pycel's COUNTIF."""
    from pycel.lib.stats import countifs as _countifs
    if _is_arr(criteria):
        from collections import Counter
        cnt = Counter((v.lower() if isinstance(v, str) else v) for v in _flat(rng) if v is not None and v != "")
        if all(not (isinstance(c, str) and c[:1] in "<>=*?") and not (isinstance(c, str) and ("*" in c or "?" in c))
               for c in _flat(criteria)):
            return tuple(tuple(cnt.get(c.lower() if isinstance(c, str) else c, 0) for c in row) for row in criteria)
    return _countifs(rng, criteria)


def sumifs(sum_range, *args):
    """SUMIFS; when a criteria argument is an array (SUMIFS(..., tLines[seller_id], tSellers[seller_id], ...) inside SUMPRODUCT),
    returns the array of sums. Equality criteria are grouped in one pass; other forms fall back to pycel's SUMIFS."""
    from pycel.excellib import sumifs as _sumifs
    crits = list(zip(args[0::2], args[1::2]))
    arr = [i for i, (_, c) in enumerate(crits) if _is_arr(c)]
    if not arr:
        return _sumifs(sum_range, *args)
    k = arr[0]
    shape = crits[k][1]
    out = []
    for row in shape:
        out_row = []
        for c in row:
            a = list(args)
            a[2 * k + 1] = c
            out_row.append(_sumifs(sum_range, *a))
        out.append(tuple(out_row))
    return tuple(out)


def countifs(*args):
    """COUNTIFS; when a criteria argument is an array, returns the array of counts (one per criteria value)."""
    from pycel.lib.stats import countifs as _countifs
    crits = list(zip(args[0::2], args[1::2]))
    arr = [i for i, (_, c) in enumerate(crits) if _is_arr(c)]
    if not arr:
        return _countifs(*args)
    k = arr[0]
    out = []
    for row in crits[k][1]:
        out_row = []
        for c in row:
            a = list(args)
            a[2 * k + 1] = c
            out_row.append(_countifs(*a))
        out.append(tuple(out_row))
    return tuple(out)


# ---------------------------------------------------------------- Week 4 additions (verification only)
pearson = correl


def _rows(x):
    """2-D range -> list of lists of floats (one per row)."""
    if not isinstance(x, (tuple, list)):
        return [[float(x)]]
    return [[float(v) for v in (r if isinstance(r, (tuple, list)) else (r,))] for r in x]


def steyx(ys, xs):
    y, x = np.array(_nums(ys), float), np.array(_nums(xs), float)
    b, a = np.polyfit(x, y, 1)
    return float(np.sqrt(((y - (a + b * x)) ** 2).sum() / (len(y) - 2)))


def norm_s_inv(p):
    from scipy import stats
    return float(stats.norm.ppf(p))


def norm_inv(p, m, s):
    from scipy import stats
    return float(stats.norm.ppf(p, m, s))


def var_p(*args):
    return statistics.pvariance(_nums(args))


def linest(known_y, known_x, const=True, want_stats=False):
    """LINEST with intercept; returns Excel's layout (coefficients right-to-left, intercept last; 5 stat rows if asked)."""
    y = np.array(_nums(known_y), float)
    X = np.array(_rows(known_x), float)
    if X.shape[0] != len(y) and X.shape[1] == len(y):
        X = X.T
    n, k = X.shape
    Xm = np.column_stack([X, np.ones(n)])
    beta, *_ = np.linalg.lstsq(Xm, y, rcond=None)
    fit = Xm @ beta
    sse = float(((y - fit) ** 2).sum())
    sst = float(((y - y.mean()) ** 2).sum())
    df = n - k - 1
    se = np.sqrt(np.diag(sse / df * np.linalg.inv(Xm.T @ Xm)))
    row1 = tuple(float(v) for v in list(beta[:k][::-1]) + [beta[k]])
    if not want_stats:
        return (row1,)
    na = "#N/A"
    r2 = 1 - sse / sst
    rows_ = [row1, tuple(float(v) for v in list(se[:k][::-1]) + [se[k]]),
             (r2, float(np.sqrt(sse / df))) + (na,) * (k - 1), (float(r2 / k / ((1 - r2) / df)), float(df)) + (na,) * (k - 1),
             (sst - sse, sse) + (na,) * (k - 1)]
    return tuple(rows_)


def mode_sngl(*args):
    """MODE.SNGL: most frequent number; ties go to the value that appears first (Excel's rule)."""
    v = _nums(args)
    cnt = {}
    for a in v:
        cnt[a] = cnt.get(a, 0) + 1
    top = max(cnt.values())
    if top < 2:
        return "#N/A"
    return next(a for a in v if cnt[a] == top)


mode = mode_sngl
