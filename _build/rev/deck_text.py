"""Print slide text. Usage: deck_text.py <vol 1-5> <slide numbers or ranges e.g. 82,102-104> [maxchars]"""
import sys, glob
from pptx import Presentation
vol = sys.argv[1]; spec = sys.argv[2]; mx = int(sys.argv[3]) if len(sys.argv) > 3 else 1500
f = glob.glob(f"/home/user/ETEH/ETEH_Bootcamp_Pack/ETEH_Vol{vol}_*.pptx")[0]
p = Presentation(f); want = []
for part in spec.split(","):
    a, _, b = part.partition("-"); want += list(range(int(a), int(b or a) + 1))
for i in want:
    s = p.slides[i - 1]; out = []
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip(): out.append(sh.text_frame.text.replace("\n", " / "))
        elif sh.has_table: out.append("TABLE: " + " | ".join(c.text for r in sh.table.rows for c in r.cells))
    t = " ¶ ".join(x for x in out if "Eaglet2EagleHub" not in x)
    print(f"--- {i}: {t[:mx]}")
