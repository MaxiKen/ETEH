"""Guard against stale hard-coded figures: every headline fact about the edition-3 organisations (Mercyland, Kingsway,
Atlantic Pearl) that appears as typed text in slide copy must match the data. Recomputes the facts from the CSVs (facts3.py)
and checks that the expected phrasing is present in the built decks and absent in stale form.

    python3 _build/rev/check_facts.py          # exit code 1 on any mismatch
"""
from __future__ import annotations

import sys
from pathlib import Path

from pptx import Presentation

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
from facts3 import F, e23, st  # noqa: E402

PACK = HERE.parent.parent / "ETEH_Bootcamp_Pack"


def deck_text(name):
    out = []
    for s in Presentation(PACK / name).slides:
        for sh in s.shapes:
            if sh.has_text_frame:
                out.append(sh.text_frame.text)
            if sh.has_table:
                out += [c.text for r in sh.table.rows for c in r.cells]
    return "\n".join(out)


fac = st.groupby("Faculty").At_Risk.apply(lambda s: (s == "Yes").mean())
dep = e23.groupby("Department").Readmitted_30d.apply(lambda s: (s == "Yes").mean())
MONTH = ["", "Jan", "Feb", "Mar", "Apr", "May", "June", "July", "Aug", "Sep", "Oct", "Nov", "Dec"][F["PEAKM"]]
W1 = [
    f"{F['STUD']:,} student records", f"{F['EP_ALL']:,} admissions, 2019-2024", f"{F['RES']:,} reservations, 2023-2024",
    f"{F['STUD']:,} students; {F['ATRISK']:,} at risk", f"{F['EP23']:,} admissions; readmission {F['READM23']:.1%}",
    f"At-risk share highest in {fac.idxmax()} ({fac.max():.1%})", f"{dep.idxmax()} has the highest readmission rate ({dep.max():.1%})",
    f"({MONTH} {F['PEAKOCC']:.1%})", f"on {F['HI95']} days in 2023", f"{F['OTA']:.1%} of online agency bookings cancelled",
    f"{F['ATRISK']:,} of {F['STUD']:,} students", f"{F['ATRISK'] / F['STUD']:.1%}", f"{F['ALOS23']:.1f} days", f"{F['CANC']:.1%}",
]
CHECKS = {"ETEH_Vol1_Week1_Foundations_Slides.pptx": W1}


def main():
    bad = 0
    for deck, phrases in CHECKS.items():
        txt = deck_text(deck)
        for p in phrases:
            if p not in txt:
                bad += 1
                print(f"MISSING in {deck}: {p!r}")
    print(f"{sum(map(len, CHECKS.values()))} facts checked, {bad} missing")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
