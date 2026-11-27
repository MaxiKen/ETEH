"""Dump a section-example sheet: formulas with cached values. Usage: dump_sheet.py <week> "<sheet>" [maxlen]"""
import sys, warnings, openpyxl
warnings.filterwarnings("ignore")
wk, sh = sys.argv[1], sys.argv[2]
ml = int(sys.argv[3]) if len(sys.argv) > 3 else 140
p = f"/home/user/ETEH/ETEH_Bootcamp_Pack/01_Section_Examples/ETEH_Week{wk}_Section_Examples.xlsx"
ws = openpyxl.load_workbook(p)[sh]; wv = openpyxl.load_workbook(p, data_only=True)[sh]
print(sh, ws.max_row, "rows,", len(ws._charts), "charts, tables:", {t.name: t.ref for t in ws.tables.values()})
for r in range(1, ws.max_row + 1):
    out = []
    for c in range(1, min(ws.max_column, 16) + 1):
        v = ws.cell(r, c).value
        if v is None: continue
        v = getattr(v, "text", v); cv = wv.cell(r, c).value
        s = f"{ws.cell(r, c).coordinate}={str(v)[:ml]}"
        if isinstance(v, str) and v.startswith("=") and "FORMULATEXT" not in v: s += f" [{cv}]"
        if "FORMULATEXT" in str(v): continue
        out.append(s)
    if out: print(" | ".join(out))
