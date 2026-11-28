"""Self-test for a learner workbook: type every correct answer, then every wrong answer, and recalculate the checks.

    python3 _build/learn/selftest.py <workbook.xlsx> [...]

Uses the answer map written by lkit (_build/_stage/<name>.answers.json). Each check cell must show
"✓ Correct" for the right answer and must NOT show it for a wrong one. Formulas are evaluated with pycel
after structured references are translated (same engine as _build/verify_xlsx.py).
"""
from __future__ import annotations

import json
import os
import re
import sys
import tempfile
import warnings
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
sys.path[:0] = [str(HERE.parent)]
warnings.filterwarnings("ignore")
import verify_xlsx as V  # noqa: E402
from common import STAGE  # noqa: E402


def wrong(a, kind):
    if kind == "num":
        return float(a) * 1.5 + 7 if a else 99
    if kind == "mc":
        return "D" if a != "D" else "A"
    if kind in ("yn",):
        return "No" if a == "Yes" else "Yes"
    if kind == "tf":
        return "False" if a == "True" else "True"
    return "zzz wrong"


def run(path, mode):
    ans = json.loads((STAGE / (Path(path).stem + ".answers.json")).read_text())
    wb = openpyxl.load_workbook(path)
    for a in ans:
        v = a["answer"] if mode == "right" else wrong(a["answer"], a["kind"])
        if a["kind"] == "num" and mode == "right":
            v = float(v)
        wb[a["sheet"]][a["cell"]].value = v
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False).name
    wb.save(tmp)
    from pycel import ExcelCompiler
    wb2 = openpyxl.load_workbook(tmp)
    T = V.table_map(wb2)
    for ws in wb2.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                v = getattr(v, "text", v)
                if c.data_type == "f" and isinstance(v, str) and v.startswith("="):
                    c.value = V.translate(v, T, ws.title, c.row, c.column)
        for k in list(ws.tables.keys()):
            del ws.tables[k]
    wb2.save(tmp)
    sys.path.insert(0, str(HERE.parent))
    xc = ExcelCompiler(filename=tmp, plugins=["pycel_plugin"])
    bad, skipped = [], 0
    for a in ans:
        col = "F"
        r = re.match(r"[A-Z]+(\d+)", a["cell"]).group(1)
        try:
            got = xc.evaluate(f"{V.qs(a['sheet'])}!{col}{r}")
        except Exception as e:  # noqa: BLE001
            got = f"ERR {type(e).__name__}"
            xc.graph_todos, xc.range_todos = [], []
        if isinstance(got, str) and got.startswith("ERR") and a.get("dynamic"):
            skipped += 1          # dynamic-array Solutions formula (FILTER/UNIQUE): pycel cannot spill it; Excel 365 can
            continue
        ok = (got == "✓ Correct") if mode == "right" else (got != "✓ Correct")
        if not ok:
            bad.append((a["sheet"], a["cell"], a["kind"], a["answer"], got))
    os.unlink(tmp)
    return len(ans), bad, skipped


if __name__ == "__main__":
    total_bad = 0
    for p in sys.argv[1:]:
        for mode in ("right", "wrong"):
            n, bad, sk = run(p, mode)
            note = f" ({sk} dynamic-array item(s) not evaluable by pycel)" if sk else ""
            print(f"{Path(p).name} [{mode} answers]: {n} questions, {len(bad)} problems{note}")
            for b in bad[:20]:
                print("   ", b)
            total_bad += len(bad)
    sys.exit(1 if total_bad else 0)
