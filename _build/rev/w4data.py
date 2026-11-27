"""Load Week 4 workbook D_ sheets as DataFrames (cached in _stage). Usage: from w4data import D; D('D_Salary')"""
import os, pickle, openpyxl, pandas as pd
WB = "/home/user/ETEH/ETEH_Bootcamp_Pack/01_Section_Examples/ETEH_Week4_Section_Examples.xlsx"
CACHE = "/home/user/ETEH/_build/_stage/w4data.pkl"
_c = None
def _load():
    global _c
    if _c is None:
        if os.path.exists(CACHE):
            _c = pickle.load(open(CACHE, "rb"))
        else:
            wb = openpyxl.load_workbook(WB, read_only=True, data_only=True)
            _c = {}
            for ws in wb.worksheets:
                if not ws.title.startswith("D_"): continue
                rows = list(ws.iter_rows(min_row=5, values_only=True))
                hdr = [h for h in rows[0][1:] if h is not None]
                _c[ws.title] = pd.DataFrame([r[1:1+len(hdr)] for r in rows[1:] if r[1] is not None], columns=hdr)
            pickle.dump(_c, open(CACHE, "wb"))
    return _c
def D(name): return _load()[name].copy()
