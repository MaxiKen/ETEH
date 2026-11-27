"""Update the five ETEH slide volumes (Arena build step 5).

* removes the last two (lab / extension) modules of every volume
* fixes duplicated module numbers on dividers, footers and roadmaps (aligned with the Section Examples sheets)
* rewrites modules 2.4, 3.10, 4.14 and P.2 so they use the bootcamp's own practice files - no external web sources
* rebuilds the P.3-P.5 capstone walkthroughs on the real capstone results (03_Instructor_Projects)
* fixes Vol 1 currency ($ -> ₦) and the external-source examples in 1.3.2 / 1.3.4
* aligns recurring illustrative figures with the real datasets (see facts.py)

Originals in the repository root are left untouched; updated decks are written to ETEH_Bootcamp_Pack/.
"""
import re
import sqlite3
import sys
from pathlib import Path

import pandas as pd
from pptx import Presentation

sys.path.insert(0, str(Path(__file__).parent))
from common import ROOT, PACK, CORE, W3D  # noqa: E402
import slidelib as L  # noqa: E402
from slide_content import MODULES, CAPSTONE, V1_FIXES  # noqa: E402
from slide_facts import FACT_SUBS  # noqa: E402

DECKS = [
    (1, "ETEH_Vol1_Week1_Foundations_Slides-2.pptx", "ETEH_Vol1_Week1_Foundations_Slides.pptx", (178, 202), 1),
    (2, "ETEH_Vol2_Week2_PowerQuery_Pivots_Slides-1.pptx", "ETEH_Vol2_Week2_PowerQuery_Pivots_Slides.pptx", (229, 252), 2),
    (3, "ETEH_Vol3_Week3_SQL_DAX_Dashboards_Slides-1.pptx", "ETEH_Vol3_Week3_SQL_DAX_Dashboards_Slides.pptx", (223, 242), 3),
    (4, "ETEH_Vol4_Week4_Stats_Forecast_AI_Slides-1.pptx", "ETEH_Vol4_Week4_Stats_Forecast_AI_Slides.pptx", (203, 225), 4),
    (5, "ETEH_Vol5_Capstone_Career_Slides-1.pptx", "ETEH_Vol5_Capstone_Career_Slides.pptx", (125, 152), None),
]
# roadmap boxes hold ~50 characters; the template truncated longer titles mid-word
SHORT_TITLES = {
    "Data Entry Best Practices: The Twelve Tidy Data Laws": "Data Entry Best Practices: Twelve Tidy Data Laws",
    "Power Pivot and DAX: Time Intelligence, KPIs and Performance": "DAX: Time Intelligence, KPIs and Performance",
}
COMPANION = {
    1: "Companion files: ETEH_Week1_Section_Examples.xlsx and ETEH_Week1_Assignment.xlsx.",
    2: "Companion files: ETEH_Week2_Section_Examples.xlsx and ETEH_Week2_Assignment.xlsx.",
    3: "Companion files: ETEH_Week3_Section_Examples.xlsx and ETEH_Week3_Assignment.xlsx.",
    4: "Companion files: ETEH_Week4_Section_Examples.xlsx and ETEH_Week4_Assignment.xlsx.",
    5: "Companion files: ETEH_Capstone_Section_Examples.xlsx and the four student projects.",
}


def tag_of(slide):
    t = L.get_text(slide, "TextBox 8")
    if not t:
        return None
    t = t.split("·")[0].strip()
    return t if re.match(r"^[0-9P]+(\.\d+)+$", t) else None


def is_divider(slide):
    t = L.get_text(slide, "TextBox 3")
    return bool(t and t.startswith("MODULE "))


def structure(prs, vol, cut):
    """Remove lab modules, renumber dividers/footers, update roadmap and the IN THIS SESSION box."""
    slides = list(prs.slides)
    divs = [i for i, s in enumerate(slides) if is_divider(s)]
    info = []
    for i in divs:
        nxt = slides[i + 1]
        tg = tag_of(nxt)
        num = tg.rsplit(".", 1)[0]
        ntop = int(re.search(r"(\d+)\s+TOPICS", L.get_text(slides[i], "TextBox 3")).group(1))
        info.append(dict(idx=i, num=num, n=ntop, title=L.get_text(slides[i], "TextBox 4")))
    counts = pd.Series([d["num"] for d in info]).value_counts()
    seen = {}
    for d in info:
        if counts[d["num"]] > 1:
            seen[d["num"]] = seen.get(d["num"], 0) + 1
            d["label"] = d["num"] + "AB"[seen[d["num"]] - 1]
        else:
            d["label"] = d["num"]
        d["keep"] = not (cut[0] - 1 <= d["idx"] <= cut[1] - 1)
    # new titles from MODULES
    for d in info:
        m = MODULES.get(d["label"])
        if m:
            d["title"] = m["title"]
    # dividers + footers
    cur = None
    for i, s in enumerate(slides):
        if is_divider(s):
            cur = next(d for d in info if d["idx"] == i)
            L.put(s, "TextBox 3", f"MODULE {cur['label']}  ·  {cur['n']} TOPICS", f"V{vol}S{i + 1}")
            continue
        if cur is None:
            continue
        foot = L.get_text(s, "TextBox 5")
        if foot and foot.startswith("Module "):
            L.set_text(L.shapes_by_name(s)["TextBox 5"].text_frame, f"Module {cur['label']}  ·  {cur['title']}"[:60], check=False)
    # roadmap (slide 3)
    rm = slides[2]
    sh = L.shapes_by_name(rm)
    rects = sorted([x for x in rm.shapes if x.name.startswith("Rounded Rectangle")], key=lambda x: (x.left, x.top))
    entries = []
    for r in rects:
        row = [x for x in rm.shapes if x.has_text_frame and abs(x.top - r.top) < 5000 and r.left - 5000 <= x.left <= r.left + 6 * L.EMU and x.text_frame.text.strip()]
        row = sorted(row, key=lambda x: x.left)
        entries.append((r, row))
    assert len(entries) == len(info), (vol, len(entries), len(info))
    for (r, row), d in zip(entries, info):
        if not d["keep"]:
            L.delete_shape(r)
            for x in row:
                L.delete_shape(x)
            continue
        lab, ttl, cnt = row
        L.set_text(lab.text_frame, d["label"], check=False)
        L.set_text(ttl.text_frame, SHORT_TITLES.get(d["title"], d["title"])[:50], check=False)
        L.set_text(cnt.text_frame, str(d["n"]), check=False)
    # IN THIS SESSION
    kept = [d for d in info if d["keep"]]
    s2 = slides[1]
    L.put(s2, "TextBox 13",
          f"{sum(d['n'] for d in kept)} topics across {len(kept)} modules.\n\n"
          "Every topic is presented in full: concept, worked workplace example, tables and best-practice guidance.\n\n"
          "Datasets: SuperMart Retail, Crestview University, Lifecare General Hospital.\n\n" + COMPANION[vol], f"V{vol}S2")
    # remove slides last
    L.remove_slides(prs, cut[0], cut[1])
    return info


def rewrite_module(prs, vol, label):
    m = MODULES[label]
    slides = list(prs.slides)
    i = next(k for k, s in enumerate(slides) if is_divider(s) and L.get_text(s, "TextBox 3").startswith(f"MODULE {label} "))
    L.put(slides[i], "TextBox 4", m["title"], f"V{vol}S{i + 1}")
    L.put(slides[i], "TextBox 5", m["desc"], f"V{vol}S{i + 1}")
    k = i + 1
    for t in m["topics"]:
        s1, s2 = slides[k], slides[k + 1]
        w1, w2 = f"V{vol}S{k + 1}", f"V{vol}S{k + 2}"
        assert tag_of(s1) == t["tag"], (tag_of(s1), t["tag"])
        names = L.shapes_by_name(s1)
        L.put(s1, "Title 1", t["title"], w1)
        L.put(s2, "Title 1", t["title"] + " (continued)", w2)
        L.put(s1, "TextBox 12", t["key"], w1)
        has_src = names["TextBox 15"].text_frame.text.strip() == "LIVE SOURCE"
        if has_src:
            L.put(s1, "TextBox 15", t["src"][0], w1)
            L.put(s1, "TextBox 16", t["src"][1], w1)
            L.put(s1, "TextBox 20", "BEFORE STATE: " + t["before"], w1)
            L.put(s1, "TextBox 24", "AFTER STATE: " + t["after"], w1)
            L.put(s1, "TextBox 28", t["best"], w1)
            L.put(s2, "TextBox 12", t["watch"], w2)
            L.put(s2, "TextBox 15", "WORKED EXAMPLE · " + t.get("sector", m["sector"]), w2)
            L.put(s2, "TextBox 16", t["example"], w2)
            L.set_table(s2, "Table 17", t["table"], t.get("widths"), w2)
        else:
            L.put(s1, "TextBox 16", "BEFORE STATE: " + t["before"], w1)
            L.put(s1, "TextBox 20", "AFTER STATE: " + t["after"], w1)
            L.put(s1, "TextBox 24", t["best"], w1)
            L.put(s1, "TextBox 28", t["watch"], w1)
            L.put(s2, "TextBox 11", "WORKED EXAMPLE · " + t.get("sector", m["sector"]), w2)
            L.put(s2, "TextBox 12", t["example"], w2)
            L.set_table(s2, "Table 13", t["table"], t.get("widths"), w2)
        k += 2


def rewrite_capstones(prs):
    slides = list(prs.slides)
    for sno, spec in CAPSTONE.items():
        s = slides[sno - 1]
        w = f"V5S{sno}"
        for name, val in spec.items():
            if name.startswith("Table"):
                if isinstance(val, dict):
                    L.set_table(s, name, val["rows"], val.get("widths"), w)
                else:
                    L.set_table(s, name, val, None, w)
            else:
                L.put(s, name, val, w)


MONEY = re.compile(r"(?<![A-Za-z$:])(-?)\$(\d[\d,]*(?:\.\d+)?)(k?)(?![\d,]*:\$)")


def naira_scaled(m):
    sign, num, k = m.group(1), m.group(2), m.group(3)
    v = float(num.replace(",", ""))
    if k:
        txt = f"{v:g}m"
    else:
        n = v * 1000
        if n >= 1e9 and round(n / 1e7) == n / 1e7:
            txt = f"{n / 1e9:g}bn"
        else:
            txt = f"{n:,.0f}"
    return f"{sign}₦{txt}"


def v1_currency(prs):
    slides = list(prs.slides)
    changed = 0
    for sno in V1_FIXES["currency_slides"]:
        s = slides[sno - 1]
        for sh, tf in L.iter_text_frames(s):
            if sno in V1_FIXES["currency_unscaled"]:
                f = lambda t: MONEY.sub(lambda m: f"{m.group(1)}₦{m.group(2)}{m.group(3)}", t).replace("10-cent", "10-kobo")
            else:
                f = lambda t: MONEY.sub(naira_scaled, t)
            f2 = lambda t: f(t).replace("($ / %)", "(₦ / %)").replace("CAC ($)", "CAC (₦)").replace("Unit Price ($)", "Unit Price (₦)").replace("% and $!", "% and ₦!")
            changed += L.replace_in_runs(tf, f2)
    return changed


def v1_sources(prs):
    slides = list(prs.slides)
    for sno, spec in V1_FIXES["slides"].items():
        s = slides[sno - 1]
        for name, val in spec.items():
            if name.startswith("Table"):
                L.set_table(s, name, val, None, f"V1S{sno}")
            else:
                L.put(s, name, val, f"V1S{sno}")


def facts(prs, vol, log):
    n = 0
    slides = list(prs.slides)
    for scope, subs in FACT_SUBS.get(vol, {}).items():
        targets = range(len(slides)) if scope == "*" else [x - 1 for x in scope]
        hits = {o: 0 for o, _ in subs}
        for si in targets:
            for sh, tf in L.iter_text_frames(slides[si]):
                before_txt = tf.text
                for old, new in subs:
                    assert old not in new or new.count(old) == 0, old
                    c = L.para_replace(tf, old, new)
                    if c:
                        n += c
                        hits[old] += c
                        log.append(f"V{vol} S{si + 1}: {old!r} -> {new!r} x{c}")
                if tf.text != before_txt and sh.has_text_frame:
                    lines, cpl = L.capacity(sh, L._font_size(tf))
                    need_new, need_old = L.needed_lines(tf.text, cpl), L.needed_lines(before_txt, cpl)
                    if need_new > max(lines, need_old):
                        L.WARN.append(f"OVERFLOW V{vol}S{si + 1}/{sh.name} (facts): needs {need_new}, box {lines}, "
                                      f"was {need_old}: {tf.text[:60]}")
        for o, c in hits.items():
            if c == 0 and scope != "*":
                L.WARN.append(f"FACT NO MATCH V{vol} {scope if len(scope) < 4 else 'multi'}: {o[:70]!r}")
    return n


def main():
    log = []
    for vol, src, dst, cut, wk in DECKS:
        prs = Presentation(str(ROOT / src))
        before = len(prs.slides)
        # content edits use ORIGINAL slide numbers, so do them before slides are removed
        if vol == 1:
            v1_sources(prs)
            c = v1_currency(prs)
            log.append(f"V1 currency runs changed: {c}")
        nf = facts(prs, vol, log)
        log.append(f"V{vol} fact substitutions: {nf}")
        if vol == 5:
            rewrite_capstones(prs)
        info = structure(prs, vol, cut)
        for label in MODULES:
            if any(d["label"] == label and d["keep"] for d in info):
                rewrite_module(prs, vol, label)
        out = PACK / dst
        prs.save(str(out))
        log.append(f"V{vol}: {before} -> {len(prs.slides)} slides -> {out.relative_to(ROOT)}")
    print("\n".join(log))
    if L.WARN:
        print("\nWARNINGS:")
        print("\n".join(L.WARN))


if __name__ == "__main__":
    main()
