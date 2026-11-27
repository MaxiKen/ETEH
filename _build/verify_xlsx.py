"""Independent check of every formula in a generated workbook.

1. Structured references (tSales[Col], [@Col], [[#This Row],Col], SUM([Col]) inside a table) are
   translated to plain A1 references.
2. The translated copy is evaluated with pycel (a Python Excel-formula engine).
3. Each evaluated result is compared with the cached value the build script wrote.

Usage: python3 _build/verify_xlsx.py file.xlsx [more.xlsx ...]
Cells whose functions pycel does not implement are reported as 'unsupported' (not failures).
"""
from __future__ import annotations

import datetime as dt
import math
import re
import os
import sys
import tempfile
import logging
import warnings
from collections import Counter

import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries

warnings.filterwarnings("ignore")
logging.disable(logging.CRITICAL)


def table_map(wb):
    T = {}
    for ws in wb.worksheets:
        for t in ws.tables.values():
            name = t.name
            c1, r1, c2, r2 = range_boundaries(t.ref)
            hdr = [ws.cell(r1, c).value for c in range(c1, c2 + 1)]
            tot = 1 if t.totalsRowCount or t.totalsRowShown and any(
                getattr(col, "totalsRowFunction", None) or getattr(col, "totalsRowLabel", None) for col in t.tableColumns) else 0
            T[name] = dict(sheet=ws.title, c1=c1, r1=r1, c2=c2, r2=r2 - tot, hdr=[str(h) for h in hdr], tot=tot)
    return T


def qs(s):
    return "'" + s.replace("'", "''") + "'"


def colrange(t, col, sheet_prefix=True):
    j = t["c1"] + t["hdr"].index(col)
    L = get_column_letter(j)
    pre = qs(t["sheet"]) + "!" if sheet_prefix else ""
    return f"{pre}${L}${t['r1'] + 1}:${L}${t['r2']}"


def find_table(T, sheet, r, c):
    for n, t in T.items():
        if t["sheet"] == sheet and t["c1"] <= c <= t["c2"] and t["r1"] <= r <= t["r2"] + t["tot"]:
            return n, t
    return None, None


RE_THISROW = re.compile(r"(\b[A-Za-z_][\w.]*)?\[\[#This Row\],\[?([^\]]+?)\]?\]")
RE_AT = re.compile(r"(\b[A-Za-z_][\w.]*)?\[@\[?([^\]]+?)\]?\]")
RE_COL = re.compile(r"(\b[A-Za-z_][\w.]*)?\[([^\[\]@#]+)\]")
RE_BARE = re.compile(r"(?<![\w'!\[])(t[A-Z][\w]*)(?![\w\[(])")


def translate(f, T, sheet, r, c):
    own_name, own = find_table(T, sheet, r, c)

    def this_row(m):
        tn = m.group(1) or own_name
        t = T.get(tn) if tn else None
        if not t:
            return m.group(0)
        j = t["c1"] + t["hdr"].index(m.group(2))
        pre = "" if t["sheet"] == sheet else qs(t["sheet"]) + "!"
        return f"{pre}{get_column_letter(j)}{r}"

    f = RE_THISROW.sub(this_row, f)
    f = RE_AT.sub(this_row, f)

    def col(m):
        tn = m.group(1)
        if tn and tn not in T:
            return m.group(0)
        t = T.get(tn) if tn else own
        if not t or m.group(2) not in t["hdr"]:
            return m.group(0)
        return colrange(t, m.group(2))

    f = RE_COL.sub(col, f)

    def bare(m):
        t = T.get(m.group(1))
        if not t:
            return m.group(0)
        return f"{qs(t['sheet'])}!${get_column_letter(t['c1'])}${t['r1'] + 1}:${get_column_letter(t['c2'])}${t['r2']}"

    f = RE_BARE.sub(bare, f)
    return f


def _ser(x):
    if isinstance(x, dt.datetime):
        return (x - dt.datetime(1899, 12, 30)).total_seconds() / 86400
    if isinstance(x, dt.date):
        return (dt.datetime.combine(x, dt.time()) - dt.datetime(1899, 12, 30)).days
    return x


def same(a, b):
    if a is None and b in ("", 0, None):
        return True
    if a == "" and b is None:
        return True
    a, b = _ser(a), _ser(b)
    if isinstance(b, str) and re.match(r"^\d{4}-\d{2}-\d{2}$", b or ""):
        b = (dt.datetime.fromisoformat(b) - dt.datetime(1899, 12, 30)).days
    if isinstance(a, bool) or isinstance(b, bool):
        return str(a).upper() == str(b).upper()
    try:
        fa, fb = float(a), float(b)
        return math.isclose(fa, fb, rel_tol=1e-6, abs_tol=0.01)
    except (TypeError, ValueError):
        return str(a).strip() == str(b).strip()


def verify(path, show=25):
    from pycel import ExcelCompiler
    wb = openpyxl.load_workbook(path)
    cached = openpyxl.load_workbook(path, data_only=True)
    T = table_map(wb)
    names = {}
    try:
        for n, dn in wb.defined_names.items():
            names[n] = dn.attr_text
    except Exception:  # noqa: BLE001
        pass
    cells = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if cell.data_type == "f" and isinstance(v, str) and v.startswith("="):
                    nf = translate(v, T, ws.title, cell.row, cell.column)
                    for n, ref in names.items():
                        nf = re.sub(rf"(?<![\w.]){re.escape(n)}(?![\w(])", ref, nf)
                    cell.value = nf
                    cells.append((ws.title, cell.coordinate, v, nf))
                elif hasattr(v, "text") and str(getattr(v, "text", "")).startswith("="):
                    nf = translate(v.text, T, ws.title, cell.row, cell.column)
                    cell.value = nf
                    cells.append((ws.title, cell.coordinate, v.text, nf))
    for ws in wb.worksheets:
        for k in list(ws.tables.keys()):
            del ws.tables[k]
    try:
        for k in list(wb.defined_names.keys()):
            del wb.defined_names[k]
    except Exception:  # noqa: BLE001
        pass
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False).name
    wb.save(tmp)
    sys.path.insert(0, __import__('os').path.dirname(__file__))
    xc = ExcelCompiler(filename=tmp, plugins=['pycel_plugin'])
    ok = bad = unsup = 0
    fails, unsup_f = [], Counter()
    for sh, coord, orig, nf in cells:
        exp = cached[sh][coord].value
        try:
            got = xc.evaluate(f"{qs(sh)}!{coord}")
        except Exception as e:  # noqa: BLE001
            msg = str(e)
            fn = re.findall(r"([A-Z][A-Z0-9.]+)\(", nf)
            unsup += 1
            unsup_f[",".join(sorted(set(fn)))[:60] or msg[:40]] += 1
            if os.environ.get("VDEBUG") and not fn:
                print("  UNSUP", sh, coord, orig[:60], "|", msg.strip().splitlines()[-1][:160])
            continue
        if isinstance(got, str) and got.startswith("#") and not (isinstance(exp, str) and exp.startswith("#")):
            fn = re.findall(r"([A-Z][A-Z0-9.]+)\(", nf)
            if any(x in nf for x in ("FORMULATEXT", "TYPE(", "ISTEXT", "LEN(", "UNIQUE", "FILTER", "SORT", "TEXTBEFORE", "TEXTAFTER", "XLOOKUP", "LET", "SEQUENCE", "FORECAST", "TEXTJOIN", "CONCAT", "SWITCH", "IFS", "EXACT", "OFFSET", "AGGREGATE", "DATEDIF", "NETWORKDAYS", "WORKDAY", "EOMONTH", "PERCENTRANK", "RANK", "CORREL", "RSQ", "SLOPE", "INTERCEPT", "STDEV", "VAR", "SKEW", "KURT", "PERCENTILE", "QUARTILE", "STANDARDIZE", "NORM", "MODE", "MEDIAN", "LINEST", "TREND", "FREQUENCY", "DATEVALUE")):
                unsup += 1
                unsup_f[",".join(sorted(set(fn)))[:60]] += 1
                continue
        if same(got, exp):
            ok += 1
        else:
            bad += 1
            fails.append((sh, coord, orig[:90], got, exp))
    print(f"\n{path.split('/')[-1]}: {len(cells)} formulas | match {ok} | MISMATCH {bad} | not evaluable {unsup}")
    for f in fails[:show]:
        print("  MISMATCH", f)
    if unsup_f:
        print("  not evaluable by pycel (by function set):", dict(unsup_f.most_common(12)))
    return bad


if __name__ == "__main__":
    tot = sum(verify(p) for p in sys.argv[1:])
    sys.exit(1 if tot else 0)
