"""Lint for the v3 decks.  python3 _build/rev/lint_v3.py [week]   (default week 1)

  PAGE     page-number box does not match the slide index
  MODULE   footer module number does not match the topic tag
  SEQ      topic tags missing or out of order (must be 1.1.1 ... 1.9.12, 111 topics)
  GLYPH    characters the Carlito font cannot draw
  BANNED   invented or superseded figures that must never reappear
  EXCEL    'open the workbook' style instruction on a visible slide other than How This Week Works
  SPACE    double spaces or trailing spaces inside a run of text
"""
import re
import sys
from pathlib import Path

from pptx import Presentation

WEEK = int(sys.argv[1]) if len(sys.argv) > 1 else 1
DECKS = {1: "ETEH_Vol1_Week1_Foundations_Slides.pptx", 2: "ETEH_Vol2_Week2_PowerQuery_Pivots_Slides.pptx",
         3: "ETEH_Vol3_Week3_SQL_DAX_Dashboards_Slides.pptx", 4: "ETEH_Vol4_Week4_Stats_Forecast_AI_Slides.pptx"}
DECK = Path(__file__).resolve().parents[2] / "ETEH_Bootcamp_Pack" / DECKS[WEEK]
GLYPHS = set("α≥≤≈≠✔√▲▼")
BANNED = ["74 duplicate", "Tobacco", "Ward 3", "Triage Band", "41 million", "Weighted CGPA", "mean CGPA",
          "J. Adeyemi", "45 minutes to 6 seconds", "8% missing", "96 unmatched", "Every zone grew", "surge beds",
          "26 of 86", " 4x", "OTA 36%", "3 min 20 s"]
EXPECTS = {1: {"1.1": 14, "1.2": 6, "1.3": 12, "1.4": 10, "1.5": 17, "1.6": 14, "1.7": 14, "1.8": 12, "1.9": 12},
           2: {"2.1": 8, "2.2": 10, "2.3": 8, "2.4": 5, "2.5": 12, "2.6": 20, "2.7": 6, "2.8": 65, "2.9": 11, "2.10": 19,
               "2.11": 11},
           3: {"3.1": 15, "3.2": 17, "3.3": 9, "3.4": 16, "3.5": 7, "3.6": 41, "3.7": 26, "3.8": 19, "3.9": 3, "3.10": 5},
           4: {"4.1": 14, "4.2": 9, "4.3": 10, "4.4": 12, "4.5": 13, "4.6": 12, "4.7": 9, "4.8": 10, "4.9": 10, "4.10": 8,
               "4.11": 10, "4.12": 8, "4.13": 9, "4.14": 5}}
EXPECT = EXPECTS[WEEK]


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
        if re.fullmatch(r"\d\.\d+\.\d+", tag):
            tags.append(tag)
            if mod and not tag.startswith(mod.group(1).rstrip("AB") + "."):
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
            if not n.startswith(("run:", "cell")) or n == "run:Code":   # code panels keep their own spacing
                continue
            if re.search(r"[a-z,;]  +[a-z]| [,.;:](?![\w=])| \n", t) or (n == "cell" and re.search(r"\S +$", t)):
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
