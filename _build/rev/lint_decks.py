"""Structural coherence lint over the built decks.  python3 _build/rev/lint_decks.py

  PAIR      a 1/2 concept slide not followed by its 2/2 worked example (same tag, same title)
  NEXT      'Next ›' link that does not name the topic that actually follows
  MODULE    footer module number that does not match the topic tag
  EMPTY     text boxes with a section label but no body text, or tables with empty cells
  SEQ       topic tags out of order inside a module
  PAGE      page-number box that does not match the slide index
"""
import glob
import re
from pathlib import Path

from pptx import Presentation

PACK = Path(__file__).resolve().parents[2] / "ETEH_Bootcamp_Pack"
LABELS = {"IN SIMPLE TERMS", "THINK OF IT LIKE THIS", "HOW TO DO IT", "REMEMBER", "AVOID", "STEP BY STEP",
          "WHAT THIS TELLS US", "WHAT YOU SHOULD SEE"}


def info(s):
    d = {"texts": [], "title": "", "tag": None, "part": None, "next": None, "module": None, "page": None, "tables": []}
    for sh in s.shapes:
        if sh.has_table:
            d["tables"].append([[c.text for c in r.cells] for r in sh.table.rows])
        if not sh.has_text_frame:
            continue
        t = sh.text_frame.text.strip()
        if sh.name.startswith("Title"):
            d["title"] = t
        m = re.fullmatch(r"(\S+)\s+·\s+([12])/2", t)
        if m:
            d["tag"], d["part"] = m.group(1), int(m.group(2))
        m = re.match(r"Next\s+›\s+(\S+)\s+(.*)", t)
        if m:
            d["next"] = (m.group(1), m.group(2))
        m = re.match(r"Module\s+(\S+)\s+·", t)
        if m:
            d["module"] = m.group(1)
        if t.isdigit():
            d["page"] = int(t)
        d["texts"].append((sh.name, t))
    return d


def main():
    total = 0
    for f in sorted(glob.glob(str(PACK / "ETEH_Vol*.pptx"))):
        p = Presentation(f)
        S = [info(s) for s in p.slides]
        issues = []
        topics = [(i, d) for i, d in enumerate(S) if d["tag"]]
        for i, d in enumerate(S):
            n = i + 1
            if d["page"] is not None and d["page"] != n and n > 6:
                issues.append(f"PAGE s{n}: box says {d['page']}")
            if d["tag"] and d["module"]:
                mod = d["module"]
                mod = mod.rstrip("AB") if re.fullmatch(r"\d+\.\d+[AB]", mod) else mod
                if not (d["tag"].startswith(mod + ".") or d["tag"] == mod):
                    issues.append(f"MODULE s{n}: tag {d['tag']} under Module {mod}")
            if d["part"] == 1:
                nx = S[i + 1] if i + 1 < len(S) else None
                if not nx or nx["tag"] != d["tag"] or nx["part"] != 2:
                    issues.append(f"PAIR s{n}: {d['tag']} 1/2 not followed by its 2/2")
                elif nx["title"] != d["title"]:
                    issues.append(f"PAIR s{n}: title differs 1/2 '{d['title'][:40]}' vs 2/2 '{nx['title'][:40]}'")
            # labels with empty bodies: label box followed directly by an empty box
            names = d["texts"]
            for k, (nm, t) in enumerate(names):
                if t in LABELS:
                    body = names[k + 1][1] if k + 1 < len(names) else ""
                    if t != "WHAT YOU SHOULD SEE" and (not body or body in LABELS):
                        issues.append(f"EMPTY s{n}: '{t}' has no body")
            for tb in d["tables"]:
                for r in tb[1:]:
                    if any(not c.strip() for c in r):
                        issues.append(f"EMPTY s{n}: table row {r}")
            if d["next"]:
                # next topic after this one (skip the 2/2 partner)
                later = [e for j, e in topics if j > i and e["tag"] != d["tag"]]
                exp = later[0] if later else None
                if exp is None:
                    issues.append(f"NEXT s{n}: link to {d['next'][0]} but no later topic")
                elif d["next"][0] == "Module":
                    mm = re.match(r"(\S+)", d["next"][1]).group(1).rstrip("AB")
                    if not exp["tag"].startswith(mm + "."):
                        issues.append(f"NEXT s{n}: says Module {mm}, next topic is {exp['tag']}")
                elif d["next"][0] != exp["tag"]:
                    issues.append(f"NEXT s{n}: says {d['next'][0]}, next topic is {exp['tag']}")
                else:
                    shown = d["next"][1].rstrip("…").strip()
                    if not exp["title"].startswith(shown[: max(5, len(shown) - 1)]):
                        issues.append(f"NEXT s{n}: name '{shown}' vs title '{exp['title'][:50]}'")
        # order of tags inside modules
        def key(tag):
            return [int(x) if x.isdigit() else x for x in re.split(r"[.]", tag)]
        seen = []
        for j, d in topics:
            if d["part"] == 1:
                seen.append((j, d["tag"]))
        for (j1, a), (j2, b) in zip(seen, seen[1:]):
            try:
                if a.split(".")[:2] == b.split(".")[:2] and key(b) <= key(a):
                    issues.append(f"SEQ s{j2 + 1}: {b} after {a}")
            except TypeError:
                pass
        print(f"{Path(f).name}: {len(S)} slides, {sum(1 for _, d in topics if d['part'] == 1)} topics, {len(issues)} issues")
        for x in issues:
            print("   ", x)
        total += len(issues)
    print("TOTAL issues", total)


if __name__ == "__main__":
    main()
