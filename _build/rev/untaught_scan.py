import openpyxl, re, sys
from openpyxl.worksheet.formula import ArrayFormula
def fns(path, skip=()):
    wb=openpyxl.load_workbook(path); out={}
    for ws in wb.worksheets:
        if ws.title in skip: continue
        for row in ws.iter_rows():
            for c in row:
                v=c.value
                if isinstance(v,ArrayFormula): v=v.text
                if isinstance(v,str) and v.startswith('=') and c.data_type=='f':
                    for f in re.findall(r'([A-Z][A-Z0-9.]*)\(', v.replace('_xlfn.','').replace('_xlws.','')):
                        out.setdefault(f, f"{ws.title}!{c.coordinate}")
    return out
P='/home/user/ETEH/ETEH_Bootcamp_Pack/'
import glob
def taught_upto(w):
    t=set()
    for i in range(1,w+1): t|=set(fns(P+f'01_Section_Examples/ETEH_Week{i}_Section_Examples.xlsx'))
    return t
for k in sys.argv[1:]:
    mw=re.search(r'Week(\d)',k); w=int(mw.group(1)) if mw else 4
    taught=taught_upto(w); d=fns(k, skip=('Function Guide',)); g=set(fns(k))-set(d); print(k.split('/')[-1], 'untaught:', {f:a for f,a in d.items() if f not in taught}, '| guide-only:', sorted(g-taught))
