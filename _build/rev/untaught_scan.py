import openpyxl, re, sys
from openpyxl.worksheet.formula import ArrayFormula
def fns(path):
    wb=openpyxl.load_workbook(path); out={}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v=c.value
                if isinstance(v,ArrayFormula): v=v.text
                if isinstance(v,str) and v.startswith('=') and c.data_type=='f':
                    for f in re.findall(r'([A-Z][A-Z0-9.]*)\(', v.replace('_xlfn.','').replace('_xlws.','')):
                        out.setdefault(f, f"{ws.title}!{c.coordinate}")
    return out
P='/home/user/ETEH/ETEH_Bootcamp_Pack/'
taught=set(fns(P+'01_Section_Examples/ETEH_Week1_Section_Examples.xlsx'))|set(fns(P+'01_Section_Examples/ETEH_Week2_Section_Examples.xlsx'))|(set(fns(P+'01_Section_Examples/ETEH_Week3_Section_Examples.xlsx')) if 'Week3' in ' '.join(sys.argv) else set())
for k in sys.argv[1:]:
    d=fns(k); print(k.split('/')[-1], 'untaught:', {f:a for f,a in d.items() if f not in taught})
