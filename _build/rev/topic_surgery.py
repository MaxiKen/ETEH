"""One-off source surgery for the decks: drop topics, insert new topics, renumber tags, rewrite cross-references.

    python3 _build/rev/topic_surgery.py 3|4 [--apply]

Topic blocks are recognised by lines that start with `        T("x.y.z",` inside the v3/w{week}_m*.py copy files. A block
ends where the next topic starts or at the line `    ],`. New topics (rev/new_topics_w{week}.py: INSERT = {after_tag: [src, ...]})
carry a placeholder tag such as "3.1.5n1"; after deletion and insertion every module family (3.1, 3.6, ...) is renumbered
1..n in file order. The old -> new map is then applied to every x.y.z reference in the week's copy files and the files that
cite them (workbook builders, chart attachers, capstone copy). References to dropped tags are reported, not changed.
"""
import importlib.util
import re
import sys
from pathlib import Path

B = Path(__file__).resolve().parents[1]
WEEK = sys.argv[1]
APPLY = "--apply" in sys.argv
PARTS = {"3": ["w3_m1_2", "w3_m3_5", "w3_m6a", "w3_m6b", "w3_m7", "w3_m8_10"],
         "4": ["w4_m1_2", "w4_m3_4", "w4_m5_6", "w4_m7_10", "w4_m11_14"]}[WEEK]
FILES = [B / "v3" / f"{p}.py" for p in PARTS]
spec = importlib.util.spec_from_file_location("nt", B / "rev" / f"new_topics_w{WEEK}.py")
NT = importlib.util.module_from_spec(spec)
spec.loader.exec_module(NT)
DROP, INSERT = set(NT.DROP), NT.INSERT
CITERS = [B / "v3" / f"w{WEEK}_charts.py", B / "v3" / f"w{WEEK}_extra.py", B / "v3" / f"w{WEEK}_fx.py",
          B / "learn" / f"w{WEEK}.py", B / "learn" / f"a{WEEK}.py", B / "learn" / "w5.py",
          *sorted((B / "v3").glob("w5_*.py")), *sorted(B.glob("04_proj_*.py")), B / "proj_common.py", B / "rev" / "check_facts.py"]
CITERS = [c for c in CITERS if c.exists()]

START = re.compile(r'^        T\("(\d+\.\d+\.\d+(?:n\d+)?)",')
END = re.compile(r"^    \],")
TAG = re.compile(r"(?<![\w.])(" + WEEK + r"\.\d{1,2}\.\d{1,2}(?:n\d+)?)(?!\d|\.\d|n\d)")


def blocks(lines):
    """[(tag, start, end)] with end exclusive."""
    out, cur = [], None
    for i, ln in enumerate(lines):
        m = START.match(ln)
        if m or END.match(ln):
            if cur:
                out.append((cur[0], cur[1], i))
                cur = None
            if m:
                cur = (m.group(1), i)
    return out


texts = {f: f.read_text().split("\n") for f in FILES}
seen_drop, seen_ins = set(), set()
for f in FILES:
    lines = texts[f]
    for tag, s, e in reversed(blocks(lines)):
        ins = INSERT.get(tag)
        if ins:
            seen_ins.add(tag)
            add = []
            for src in ins:
                add += ["        " + ln if ln.strip() else "" for ln in src.strip("\n").split("\n")] + [""]
            lines[e:e] = add
        if tag in DROP:
            seen_drop.add(tag)
            del lines[s:e]
    texts[f] = lines
miss = (DROP - seen_drop) | (set(INSERT) - seen_ins)
assert not miss, f"tags not found: {sorted(miss)}"

# renumber in file order
mapping, count = {}, {}
for f in FILES:
    for tag, s, e in blocks(texts[f]):
        fam = ".".join(tag.split(".")[:2])
        count[fam] = count.get(fam, 0) + 1
        mapping[tag] = f"{fam}.{count[fam]}"
dropped = DROP


def sub_all(text, where, report):
    def rep(m):
        t = m.group(1)
        if t in mapping:
            return mapping[t]
        if t in dropped:
            report.append(f"{where}: reference to dropped {t}")
        return t
    return TAG.sub(rep, text)


report = []
new_texts = {f: sub_all("\n".join(texts[f]), f.name, report) for f in FILES}
cit = {}
for c in CITERS:
    old = c.read_text()
    new = sub_all(old, c.name, report)
    if new != old:
        cit[c] = new
print("families:", count)
print("changed citers:", [c.name for c in cit])
moved = {k: v for k, v in mapping.items() if k != v}
print(f"{len(moved)} tags renumbered; e.g.", list(moved.items())[:12])
print("\n".join(sorted(set(report))) or "no references to dropped tags")
if APPLY:
    for f, t in new_texts.items():
        f.write_text(t)
    for c, t in cit.items():
        c.write_text(t)
    print("APPLIED")
