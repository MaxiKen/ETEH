"""Find formulas that use [@col] (this-row) references outside an Excel table: Excel cannot resolve them and repairs the file."""
import re
import sys
import zipfile
from openpyxl.utils.cell import range_boundaries, coordinate_from_string, column_index_from_string

for path in sys.argv[1:]:
    z = zipfile.ZipFile(path)
    rels = {}
    bad = 0
    sheets = [n for n in z.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml$", n)]
    for sh in sheets:
        x = z.read(sh).decode("utf8")
        if "[@" not in x:
            continue
        relp = sh.replace("worksheets/", "worksheets/_rels/") + ".rels"
        refs = []
        if relp in z.namelist():
            for t in re.findall(r'Target="\.\./tables/(table\d+\.xml)"', z.read(relp).decode()):
                refs.append(range_boundaries(re.search(r'ref="([A-Z0-9:]+)"', z.read("xl/tables/" + t).decode()).group(1)))
        for addr, f in re.findall(r'<c r="([A-Z]+\d+)"[^>]*>(?:<f[^>]*>)([^<]*\[@[^<]*)</f>', x):
            col, row = coordinate_from_string(addr)
            ci = column_index_from_string(col)
            if not any(c0 <= ci <= c1 and r0 <= row <= r1 for c0, r0, c1, r1 in refs):
                bad += 1
                if bad <= 5:
                    print("  ", sh, addr, f[:70])
    print(path.split("/")[-1], "this-row refs outside a table:", bad)
