"""Coherence lint over the v2 topic copy (t1_*..t5_*).  python3 _build/rev/lint_topics.py

Checks per topic
  SO-NUM     a figure in the 'so' line that is not in case/steps/table (and not a simple derivation)
  TBL-SHAPE  table rows with a different number of cells from the header
  TBL-EMPTY  empty table cells
  TRY        try-it workbook or sheet that does not exist
  DUP-CASE   identical case text used by two topics
  V1-TEXT    leftovers from known-invented v1 figures
"""
import re
import sys
import warnings
from collections import defaultdict
from pathlib import Path

import openpyxl

warnings.filterwarnings("ignore")
REV = Path(__file__).resolve().parent
sys.path.insert(0, str(REV))
from copy_of import load_all  # noqa: E402
from qa_trace import tokens, PACK  # noqa: E402

BANNED = ["Tobacco", "Ward 3 31%", "Triage Band", "41 million", "Weighted CGPA", "41-record", "fourteen measures",
          "3 min 20", "J. Adeyemi", "45 minutes to 6 seconds", "8% missing", "96 unmatched", "36%", "Every zone grew",
          "surge beds", "26 of 86", "refurbishments", "seniority gap"]


def nums(text):
    out = []
    for raw, v, tol, suf, naira, nd in tokens(text):
        out.append((raw, v, tol, suf, nd))
    return out


def present(v, tol, pool):
    for p, pt in pool:
        if abs(abs(v) - abs(p)) <= max(tol, pt) * 1.01:
            return True
    return False


def derivable(v, tol, pool):
    vals = [abs(p) for p, _ in pool if p]
    for a in vals:
        for b in vals:
            for r in (a - b, a + b, a / b * 100 if b else None, (a / b - 1) * 100 if b else None, a / b if b else None):
                if r is not None and abs(abs(v) - abs(r)) <= max(tol, abs(v) * 0.006) + 1e-9:
                    return True
    return False


_wb_cache = {}


def sheets_of(name):
    if name not in _wb_cache:
        hits = list(PACK.rglob(name))
        if not hits:
            _wb_cache[name] = None
        else:
            wb = openpyxl.load_workbook(hits[0], read_only=True)
            _wb_cache[name] = set(wb.sheetnames)
            wb.close()
    return _wb_cache[name]


def main():
    allt = load_all()
    issues = defaultdict(list)
    cases = defaultdict(list)
    for tag, (fn, t) in allt.items():
        table = t.get("table") or []
        steps = t.get("steps") or []
        flat_tbl = " ".join(" ".join(map(str, r)) for r in table)
        src = " ".join([t.get("case") or "", " ".join(steps), flat_tbl])
        pool = [(v, tol) for _, v, tol, suf, nd in nums(src)]
        pool += [(v / 100, tol / 100) for _, v, tol, suf, nd in nums(src) if suf == "%"]
        for raw, v, tol, suf, nd in nums(t.get("so") or ""):
            if suf in ("", "x", "×") and nd == 0 and v <= 12:
                continue
            if 1990 <= v <= 2035 and not suf:
                continue
            if not present(v, tol, pool) and not derivable(v, tol, pool):
                issues["SO-NUM"].append(f"{tag} ({fn}): '{raw}' in so-line")
        if table:
            w = len(table[0])
            for r in table[1:]:
                if len(r) != w:
                    issues["TBL-SHAPE"].append(f"{tag}: row {r} has {len(r)} cells, header {w}")
                if any(str(c).strip() == "" for c in r):
                    issues["TBL-EMPTY"].append(f"{tag}: {r}")
        tr = t.get("tryit") or ""
        m = re.search(r"([\w\-]+\.xlsx)", tr)
        if m:
            sh = sheets_of(m.group(1))
            if sh is None:
                issues["TRY"].append(f"{tag}: workbook {m.group(1)} not found")
            else:
                for q in re.findall(r"'([^']+)'", tr):
                    if q not in sh:
                        issues["TRY"].append(f"{tag}: sheet '{q}' not in {m.group(1)}")
        elif t.get("case"):
            issues["TRY"].append(f"{tag}: no workbook in tryit: {tr[:60]}")
        if t.get("case"):
            cases[t["case"].strip()].append(tag)
        blob = " ".join(str(x) for x in t.values())
        for b in BANNED:
            if b in blob:
                issues["V1-TEXT"].append(f"{tag}: '{b}'")
    for c, tags in cases.items():
        if len(tags) > 1:
            issues["DUP-CASE"].append(f"{tags}: {c[:80]}")
    print(f"topics: {len(allt)}")
    for k in ["SO-NUM", "TBL-SHAPE", "TBL-EMPTY", "TRY", "DUP-CASE", "V1-TEXT"]:
        print(f"\n== {k}: {len(issues[k])}")
        for x in issues[k]:
            print("  ", x)


if __name__ == "__main__":
    main()
