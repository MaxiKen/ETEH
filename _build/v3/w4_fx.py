"""Helpers for Week 4 slide copy: every number on a slide is computed here from the w4_data tables (never typed).

The statistics use Excel's definitions so the workbook formulas reproduce them: sample SD (STDEV.S), QUARTILE.INC / PERCENTILE.INC
(linear interpolation), SKEW and KURT with the bias-corrected formulas, LINEST-style OLS with standard errors and two-sided
p-values (T.DIST.2T).
"""
from __future__ import annotations

import numpy as np
import pandas as pd
from scipy import optimize, stats

import w4_data as D

ROOMS = D.ROOMS


def N(x, d=0):
    x = float(x)
    return f"-₦{-x:,.{d}f}" if x < 0 else f"₦{x:,.{d}f}"


def Nm(x, d=1):
    x = float(x)
    return f"-₦{-x / 1e6:,.{d}f}m" if x < 0 else f"₦{x / 1e6:,.{d}f}m"


def Nk(x):
    return f"₦{float(x) / 1e3:,.0f}k"


def P(x, d=1):
    return f"{x:.{d}%}"


def I(x):
    return f"{int(round(float(x))):,}"


def F(x, d=1):
    return f"{float(x):,.{d}f}"


def F2(x):
    return f"{float(x):.2f}"


# ---------------------------------------------------------------- tables
SAL = D.table("tSal")
ORD = D.table("tOrders")
CV = D.table("tCV")
LOS = D.table("tLOS")
C23 = D.table("tCensus23")
HOT = D.table("tHotel")
D25 = D.table("tDash25")
MON = D.table("tMonthly")
STO = D.table("tStores")

Y = SAL.Annual_Salary_NGN
OV = ORD.total_price
EX = CV.Exam_Score
LS = LOS.Length_of_Stay


# ---------------------------------------------------------------- Excel-matching statistics
def skew(x):
    return float(stats.skew(x, bias=False))


def kurt(x):
    return float(stats.kurtosis(x, bias=False))


def q(x, p):
    return float(np.quantile(np.asarray(x, float), p))      # = PERCENTILE.INC / QUARTILE.INC


def pct_rank_inc(x, v):
    """PERCENTRANK.INC(range, v) for a value in the range (3-digit truncation as Excel)."""
    a = np.sort(np.asarray(x, float))
    below = (a < v).sum()
    r = below / (len(a) - 1)
    return np.floor(r * 1000) / 1000


def ols(y, X, names):
    """OLS with intercept. Returns dict: coef, se, t, p, r2, adj, sey, n, k, fitted, resid (LINEST equivalents)."""
    y = np.asarray(y, float)
    Xm = np.column_stack([np.ones(len(y))] + [np.asarray(X[c], float) for c in names])
    beta, *_ = np.linalg.lstsq(Xm, y, rcond=None)
    fit = Xm @ beta
    res = y - fit
    n, k = len(y), len(names)
    sse = float((res ** 2).sum())
    sst = float(((y - y.mean()) ** 2).sum())
    df = n - k - 1
    sigma2 = sse / df
    cov = sigma2 * np.linalg.inv(Xm.T @ Xm)
    se = np.sqrt(np.diag(cov))
    t = beta / se
    p = 2 * stats.t.sf(np.abs(t), df)
    r2 = 1 - sse / sst
    adj = 1 - (1 - r2) * (n - 1) / df
    lab = ["Intercept"] + list(names)
    return dict(coef=dict(zip(lab, beta)), se=dict(zip(lab, se)), t=dict(zip(lab, t)), p=dict(zip(lab, p)), r2=r2, adj=adj,
                sey=float(np.sqrt(sigma2)), n=n, k=k, df=df, fitted=fit, resid=res, F=(r2 / k) / ((1 - r2) / df))


# ---------------------------------------------------------------- Holt-Winters additive (the method behind FORECAST.ETS)
def hw_fit(y_, m=12, h=12):
    y_ = np.asarray(y_, float)

    def run(params, h_=0):
        a, b_, g_ = params
        lvl = y_[:m].mean()
        tr = (y_[m:2 * m].mean() - y_[:m].mean()) / m
        se_ = list(y_[:m] - lvl)
        fitted, err = [], []
        for t in range(len(y_)):
            f_ = lvl + tr + se_[t]
            fitted.append(f_)
            e = y_[t] - f_
            err.append(e)
            nl = a * (y_[t] - se_[t]) + (1 - a) * (lvl + tr)
            tr = b_ * (nl - lvl) + (1 - b_) * tr
            se_.append(g_ * (y_[t] - nl) + (1 - g_) * se_[t])
            lvl = nl
        fc = [lvl + (k + 1) * tr + se_[len(y_) - m + (k % m)] for k in range(h_)]
        return np.array(err), np.array(fc), np.array(fitted)

    res = optimize.minimize(lambda p: (run(p)[0][m:] ** 2).sum(), x0=[.3, .05, .2], bounds=[(0, 1)] * 3, method="L-BFGS-B")
    err, fc, fitted = run(res.x, h)
    sig = np.sqrt((err[m:] ** 2).mean())
    a, b_, g_ = res.x
    ci = np.array([1.96 * sig * np.sqrt(1 + sum((a + j * a * b_ + (g_ if j % m == 0 else 0)) ** 2 for j in range(1, hh)))
                   for hh in range(1, h + 1)])
    return dict(fc=fc, ci=ci, params=res.x, fitted=fitted, sigma=sig)


def mape(a, f_):
    a, f_ = np.asarray(a, float), np.asarray(f_, float)
    return float(np.mean(np.abs(a - f_) / a))


def mae(a, f_):
    return float(np.mean(np.abs(np.asarray(a, float) - np.asarray(f_, float))))


def rmse(a, f_):
    return float(np.sqrt(np.mean((np.asarray(a, float) - np.asarray(f_, float)) ** 2)))
