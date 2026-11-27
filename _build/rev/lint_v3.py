"""Lint for the v3 Week 1 deck.  python3 _build/rev/lint_v3.py

  PAGE     page-number box does not match the slide index
  MODULE   footer module number does not match the topic tag
  SEQ      topic tags missing or out of order (must be 1.1.1 ... 1.9.12, 111 topics)
  GLYPH    characters the Carlito font cannot draw
  BANNED   invented or superseded figures that must never reappear
  EXCEL    'open the workbook' style instruction on a visible slide other than How This Week Works
  SPACE    double spaces or trailing spaces inside a run of text
"""
import re
from pathlib import Path

from pptx import Presentation

DECK = Path(__file__).resolve().parents[2] / "ETEH_Bootcamp_Pack" / "ETEH_Vol1_Week1_Foundations_Slides.pptx"
GLYPHS = set("α≥≤≈≠✔√▲▼")
BANNED = ["74 duplicate", "Tobacco", "Ward 3", "Triage Band", "41 million", "Weighted CGPA", "mean CGPA",
          "J. Adeyemi", "45 minutes to 6 seconds", "8% missing", "96 unmatched", "Every zone grew", "surge beds",
          "26 of 86", " 4x", "OTA 36%", "3 min 20 s"]
EXPECT = {"1.1": 14, "1.2": 6, "1.3": 12, "1.4": 10, "1.5": 17, "1.6": 14, "1.7": 14, "1.8": 12, "1.9": 12}


def texts(s):
    out = []
    for sh in s.shapes:
        if sh.has_text_frame:
            out.append((sh.name, sh.text_frame.text))
            for para in sh.text_frame.paragraphs:
                for r in para.runs:
                    out.append(("run:" + sh.name, r.text))
        if sh.has_table:
            for r in sh.table.rows:
                for c in r.cells:
                    out.append(("cell", c.text))
    return out


def main():
    p = Presentation(DECK)
    issues, tags = [], []
    for i, s in enumerate(p.slides, 1):
        tx = texts(s)
        allt = "\n".join(t for _, t in tx)
        tag = next((t.strip() for n, t in tx if n == "TextBox 8"), "")
        page = next((t.strip() for n, t in tx if n == "TextBox 6" and t.strip().isdigit()), None)
        mod = next((re.match(r"Module\s+(\S+)", t) for n, t in tx if n == "TextBox 5" and t.startswith("Module")), None)
        if page is not None and int(page) != i:
            issues.append(f"PAGE   slide {i}: box says {page}")
        if re.fullmatch(r"\d\.\d\.\d+", tag):
            tags.append(tag)
            if mod and not tag.startswith(mod.group(1) + "."):
                issues.append(f"MODULE slide {i}: tag {tag} footer {mod.group(1)}")
        bad = GLYPHS & set(allt)
        if bad:
            issues.append(f"GLYPH  slide {i}: {''.join(sorted(bad))}")
        for b in BANNED:
            if b in allt:
                issues.append(f"BANNED slide {i}: {b!r}")
        title = next((t for n, t in tx if n.startswith("Title")), "")
        if re.search(r"open (the |your )?(week ?1 )?(section )?(examples|workbook|excel file)", allt, re.I) \
                and "How This Week Works" not in title and "How This Week Works" not in allt:
            issues.append(f"EXCEL  slide {i}: workbook instruction on slide")
        for n, t in tx:
            if not n.startswith(("run:", "cell")):
                continue
            if re.search(r"[a-z,;]  +[a-z]| [,.;:](?!\w)| \n", t) or (n == "cell" and re.search(r"\S +$", t)):
                issues.append(f"SPACE  slide {i} [{n}]: {t[:50]!r}")
    # sequence
    want = [f"{m}.{k}" for m, c in EXPECT.items() for k in range(1, c + 1)]
    if tags != want:
        miss = sorted(set(want) - set(tags))
        extra = [t for t in tags if t not in want]
        issues.append(f"SEQ    {len(tags)} topics; missing {miss[:10]} extra {extra[:10]}"
                      + ("" if miss or extra else " (order differs)"))
    print(f"{len(p.slides)} slides, {len(tags)} topics")
    print("\n".join(issues) if issues else "0 issues")


if __name__ == "__main__":
    main()
