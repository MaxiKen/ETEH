"""Inventory every 'Your turn' exercise in the section workbooks: sheet, Ex n, task, key formula."""
import sys, re, json, warnings, openpyxl
warnings.filterwarnings("ignore")
P = "/home/user/ETEH/ETEH_Bootcamp_Pack/01_Section_Examples/"
out = []
for wk in sys.argv[1:]:
    p = P + ("ETEH_Capstone_Section_Examples.xlsx" if wk == "C" else f"ETEH_Week{wk}_Section_Examples.xlsx")
    wb = openpyxl.load_workbook(p)
    for ws in wb.worksheets:
        if ws.title.startswith(("Data_", "D_", "00_")): continue
        ex, key, tag = {}, {}, ""
        for r in range(1, ws.max_row + 1):
            b = ws.cell(r, 2).value
            b = getattr(b, "text", b)
            if isinstance(b, str) and re.fullmatch(r"\d+\.\d+[AB]?\.\d+(-\d+)?", b.strip()): tag = b.strip()
            if isinstance(b, str) and re.fullmatch(r"Ex \d+", b):
                ex[b] = dict(sheet=ws.title, ex=b, tag=tag, task=ws.cell(r, 3).value, chk=str(getattr(ws.cell(r, 10).value, "text", ws.cell(r, 10).value))[:80])
            if isinstance(b, str) and re.match(r"Ex \d+  ·", b):
                n = b.split("  ·")[0]
                key[n] = (b, str(getattr(ws.cell(r, 3).value, "text", ws.cell(r, 3).value)))
        for k, v in ex.items():
            v["key"] = key.get(k, ("", ""))[1]
            out.append(v)
json.dump(out, open("/tmp/a/ex.json", "w"), indent=1, default=str)
import collections
c = collections.Counter(o["sheet"] for o in out)
for k, v in c.items(): print(k, v)
print(len(out))
