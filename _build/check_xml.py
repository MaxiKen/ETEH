"""Structural sanity check of generated .xlsx files: numeric cells must hold numbers, XML must parse."""
import re
import sys
import zipfile
from xml.dom import minidom

bad_total = 0
for path in sys.argv[1:]:
    z = zipfile.ZipFile(path)
    bad = 0
    for n in z.namelist():
        if not n.endswith(".xml"):
            continue
        data = z.read(n).decode("utf-8")
        try:
            if len(data) < 3_000_000:
                minidom.parseString(data)
        except Exception as e:  # noqa: BLE001
            print("  XML PARSE ERROR", n, e); bad += 1
        if n.startswith("xl/worksheets/sheet"):
            for m in re.finditer(r'<c r="([A-Z]+\d+)"([^>]*)>(?:<f[^>]*>[^<]*</f>|<f[^>]*/>)?<v>([^<]*)</v>', data):
                attrs, v = m.group(2), m.group(3)
                t = re.search(r'\bt="(\w+)"', attrs)
                if (t is None or t.group(1) == "n") and v != "":
                    try:
                        float(v)
                    except ValueError:
                        print("  NON-NUMERIC VALUE", n, m.group(1), v[:40]); bad += 1
    print(f"{path.split('/')[-1]}: {'OK' if not bad else str(bad) + ' problems'}")
    bad_total += bad
sys.exit(1 if bad_total else 0)
