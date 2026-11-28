"""Apply rebrand_map to the active builders (and the README). Word-exact, case-sensitive, idempotent.

    python3 _build/rebrand.py            # rewrite files, print a count per file
    python3 _build/rebrand.py --check    # list remaining old names without writing
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

from rebrand_map import pairs

HERE = Path(__file__).parent
SKIP = {"neworgs.py", "rebrand_map.py", "rebrand.py"}


def files():
    out = [p for p in sorted(HERE.glob("*.py")) + sorted((HERE / "v3").glob("*.py")) + sorted((HERE / "v3").glob("*.json"))
           if p.name not in SKIP]
    out.append(HERE.parent / "ETEH_Bootcamp_Pack" / "README.md")
    return out


def pattern():
    P = sorted(pairs(), key=lambda kv: -len(kv[0]))
    alts = []
    for old, _ in P:
        a = re.escape(old)
        if re.match(r"\w", old[0]):
            a = r"(?<![A-Za-z0-9_])" + a
        if re.match(r"\w", old[-1]):
            a = a + r"(?![A-Za-z0-9_])"
        alts.append(a)
    return re.compile("|".join(alts)), dict(P)


def main():
    rx, M = pattern()
    check = "--check" in sys.argv
    tot = 0
    for p in files():
        s = p.read_text(encoding="utf-8")
        n = len(rx.findall(s))
        if n:
            tot += n
            print(f"{n:5d}  {p.relative_to(HERE.parent)}")
            if not check:
                p.write_text(rx.sub(lambda m: M[m.group(0)], s), encoding="utf-8")
    print("total", tot)


if __name__ == "__main__":
    main()
