"""Traceability audit: every figure on a worked-example slide must appear in the workbook sheet that its
'TRY IT YOURSELF' line points to (cached values or numbers inside text cells).

  python3 _build/rev/qa_trace.py [deck-glob] > _build/_stage/trace.txt

Output: one line per slide with the figures that could not be traced. These are then checked by hand
against the raw data (or the sheet pointer / figure is fixed).
"""
import glob
import re
import sys
import warnings
from functools import lru_cache
from pathlib import Path

import openpyxl
from pptx import Presentation

warnings.filterwarnings("ignore")
ROOT = Path(__file__).resolve().parents[2]
PACK = ROOT / "ETEH_Bootcamp_Pack"
NUM = re.compile(r"(?<![\w.])([−\-+]?)(₦?)(\d{1,3}(?:,\d{3})+|\d+)(\.\d+)?\s?(bn|m|k|%|x|×| pts| points| pp)?(?![\w])")


def tokens(text):
    """yield (raw, value, tol, kind)"""
    for m in NUM.finditer(text):
        sign, naira, whole, dec, suf = m.groups()
        raw = m.group(0).strip()
        v = float(whole.replace(",", "") + (dec or ""))
        nd = len(dec) - 1 if dec else 0
        unit = 10 ** -nd
        suf = (suf or "").strip()
        mult = {"bn": 1e9, "m": 1e6, "k": 1e3}.get(suf, 1)
        yield raw, v * mult, 0.5 * unit * mult + 1e-9, suf, naira, nd


def skip(raw, v, suf, naira, nd):
    if suf in ("x", "×"):
        return False
    if not naira and not suf and nd == 0:
        if v <= 12 or (1990 <= v <= 2035):
            return True          # step numbers, counts like '3 stores', years
    return False


@lru_cache(None)
def find_wb(name):
    hits = list(PACK.rglob(name))
    return hits[0] if hits else None


@lru_cache(None)
def sheet_values(path, sheet):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if sheet not in wb.sheetnames:
        return None
    vals = []
    for row in wb[sheet].iter_rows(values_only=True):
        for c in row:
            if isinstance(c, bool) or c is None:
                continue
            if isinstance(c, (int, float)):
                vals.append(float(c))
            elif isinstance(c, str):
                for _, v, _, suf, _, _ in tokens(c):
                    vals.append(v)
                    if suf == "%":
                        vals.append(v / 100)
    wb.close()
    return sorted(set(abs(v) for v in vals))


@lru_cache(None)
def all_index():
    idx = []
    for wbp in sorted(PACK.rglob("*.xlsx")):
        wb = openpyxl.load_workbook(wbp, read_only=True)
        names = wb.sheetnames
        wb.close()
        for sn in names:
            v = sheet_values(str(wbp), sn)
            if v:
                idx.append((wbp.name, sn, v))
    return idx


def locate(v, tol, suf, wbname):
    same, other = None, None
    for wn, sn, vals in all_index():
        if traced(v, tol, suf, vals):
            if wn == wbname and same is None:
                same = sn
            elif other is None:
                other = f"{wn}:{sn}"
    return same, other


def traced(v, tol, suf, vals):
    cands = [abs(v)]
    if suf == "%" or suf in ("pts", "points", "pp"):
        cands += [abs(v) / 100]
        tol_c = [tol, tol / 100]
    else:
        tol_c = [tol]
    import bisect
    for c, t in zip(cands, tol_c):
        i = bisect.bisect_left(vals, c - t)
        if i < len(vals) and vals[i] <= c + t:
            return True
    return False


def parse_try(text):
    m = re.search(r"TRY IT YOURSELF\s*›\s*(.+)", text)
    if not m:
        return None
    parts = [p.strip() for p in m.group(1).split("›")]
    wbname = next((p for p in parts if p.endswith(".xlsx")), None)
    sheet = None
    for p in parts:
        q = re.search(r"'([^']+)'", p)
        if q:
            sheet = q.group(1)
    return wbname, sheet


def slide_text(s):
    out = {}
    for sh in s.shapes:
        if sh.has_text_frame:
            if sh.text_frame.text.strip().isdigit():
                continue            # page number
            out[sh.name] = sh.text_frame.text
        if sh.has_table:
            out[sh.name] = "\n".join(" | ".join(c.text for c in r.cells) for r in sh.table.rows)
    return out


def main():
    pat = sys.argv[1] if len(sys.argv) > 1 else str(PACK / "ETEH_Vol*.pptx")
    tot = flagged = 0
    for f in sorted(glob.glob(pat)):
        p = Presentation(f)
        for i, s in enumerate(p.slides, 1):
            parts = slide_text(s)
            full = "\n".join(parts.values())
            if "WORKED EXAMPLE" not in full:
                continue
            tr = parse_try(full)
            tag = re.search(r"(\S+)\s+·\s+2/2", full)
            tag = tag.group(1) if tag else "?"
            if not tr or not tr[0]:
                print(f"{Path(f).name[5:9]} s{i} {tag}: NO TRY-IT POINTER")
                continue
            path = find_wb(tr[0])
            if not path:
                print(f"{Path(f).name[5:9]} s{i} {tag}: WORKBOOK NOT FOUND {tr[0]}")
                continue
            vals = sheet_values(str(path), tr[1]) if tr[1] else None
            if vals is None:
                print(f"{Path(f).name[5:9]} s{i} {tag}: SHEET NOT FOUND {tr}")
                continue
            body = "\n".join(t for k, t in parts.items() if "TRY IT" not in t and "Next" not in t[:6]
                             and "ETEH Ltd" not in t and "Module" not in t[:7])
            miss = {"same-wb": [], "other-wb": [], "NOWHERE": []}
            for raw, v, tol, suf, naira, nd in tokens(body):
                tot += 1
                if skip(raw, v, suf, naira, nd):
                    continue
                if not traced(v, tol, suf, vals):
                    same, other = locate(v, tol, suf, tr[0])
                    if same:
                        miss["same-wb"].append(f"{raw}@{same[:4]}")
                    elif other:
                        miss["other-wb"].append(f"{raw}@{other.split(':')[0][:22]}")
                    else:
                        miss["NOWHERE"].append(raw)
            if any(miss.values()):
                flagged += 1
                print(f"{Path(f).name[5:9]} s{i} {tag} [{tr[1]}] " + " || ".join(f"{k}: " + "; ".join(dict.fromkeys(v)) for k, v in miss.items() if v))
    print(f"# figures scanned {tot}; slides with untraced figures {flagged}", file=sys.stderr)


if __name__ == "__main__":
    main()
