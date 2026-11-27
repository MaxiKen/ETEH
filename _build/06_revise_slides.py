"""Revise the five decks into beginner-friendly, self-study slides.

Pipeline:  05_update_slides.py (v1 decks)  ->  this script (v2 decks, same file names, same folder).

Each topic becomes exactly two slides built in the ETEH design:
  1. Concept slide   - IN SIMPLE TERMS · THINK OF IT LIKE THIS (analogy) · HOW TO DO IT · REMEMBER · AVOID
  2. Example slide   - green "WORKED EXAMPLE" badge (marks the end of every topic group) · THE SITUATION ·
                       STEP BY STEP · WHAT YOU SHOULD SEE (table) · WHAT THIS TELLS US · TRY IT YOURSELF · Next
Topic copy lives in _build/rev/t*.py (TOPICS lists). Topics without copy keep their v1 slides.

Usage:
  python3 _build/06_revise_slides.py                 # full build (runs 05 first)
  python3 _build/06_revise_slides.py --skip05        # revise the current v1 decks in the pack (testing only)
  python3 _build/06_revise_slides.py --preview 1.6.1 2.4.3   # also write PNG previews to _build/_stage/preview
"""
import copy
import shutil
import importlib.util
import re
import subprocess
import sys
from pathlib import Path

import openpyxl
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Emu
from pptx.opc.packuri import PackURI

sys.path.insert(0, str(Path(__file__).resolve().parent))
import slidekit as K  # noqa: E402
from common import ROOT, PACK, SEC, STAGE  # noqa: E402

DECKS = [
    (1, "ETEH_Vol1_Week1_Foundations_Slides.pptx", "ETEH_Week1_Section_Examples.xlsx", "Week 1", "ETEH_Week1_Assignment.xlsx"),
    (2, "ETEH_Vol2_Week2_PowerQuery_Pivots_Slides.pptx", "ETEH_Week2_Section_Examples.xlsx", "Week 2", "ETEH_Week2_Assignment.xlsx"),
    (3, "ETEH_Vol3_Week3_SQL_DAX_Dashboards_Slides.pptx", "ETEH_Week3_Section_Examples.xlsx", "Week 3", "ETEH_Week3_Assignment.xlsx"),
    (4, "ETEH_Vol4_Week4_Stats_Forecast_AI_Slides.pptx", "ETEH_Week4_Section_Examples.xlsx", "Week 4", "ETEH_Week4_Assignment.xlsx"),
    (5, "ETEH_Vol5_Capstone_Career_Slides.pptx", "ETEH_Capstone_Section_Examples.xlsx", "Capstone", None),
]
TAG = re.compile(r"^([0-9P]\.\d+[AB]?\.\d+)(?:\s+·\s+(\d)/(\d))?$")
CONTENT_BOTTOM = 6.70     # concept slide: cards must end above this line
EXAMPLE_BOTTOM = 6.56     # example slide: cards must end above the TRY IT line
PREVIEW = STAGE / "preview"
V1 = STAGE / "v1_decks"   # pristine output of 05_update_slides.py; this script always revises from here


# ------------------------------------------------------------------ topic copy
def load_topics():
    topics = {}
    for f in sorted((ROOT / "_build" / "rev").glob("t*.py")):
        spec = importlib.util.spec_from_file_location(f.stem, f)
        m = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(m)
        for t in m.TOPICS:
            assert t["tag"] not in topics, f"duplicate topic {t['tag']} in {f.name}"
            t["_file"] = f.name
            topics[t["tag"]] = t
    return topics


def sheet_map(xlsx):
    wb = openpyxl.load_workbook(SEC / xlsx, read_only=True)
    names = wb.sheetnames
    wb.close()
    out = {}
    for n in names:
        code = n.split(" ")[0]
        out.setdefault(code, []).append(n)
    return out


# ------------------------------------------------------------------ deck scanning
def names(slide):
    d = {}
    for sh in slide.shapes:
        d.setdefault(sh.name, sh)
    return d


def scan(prs):
    """Return a list of segments in slide order: ('topic', tag, [slides]) or ('other', [slide])."""
    segs, cur_mod = [], None
    for s in prs.slides:
        n = names(s)
        t3 = n.get("TextBox 3")
        if t3 is not None and t3.has_text_frame and t3.text_frame.text.startswith("MODULE"):
            cur_mod = (t3.text_frame.text.split()[1], n["TextBox 4"].text_frame.text)
            segs.append(("other", [s], cur_mod))
            continue
        t8 = n.get("TextBox 8")
        m = TAG.match(t8.text_frame.text.strip()) if (t8 is not None and t8.has_text_frame) else None
        if m:
            tag = m.group(1)
            if segs and segs[-1][0] == "topic" and segs[-1][1] == tag:
                segs[-1][2].append(s)
            else:
                segs.append(("topic", tag, [s], cur_mod))
        else:
            segs.append(("other", [s], cur_mod))
    return segs


def orig_text(slides, labels):
    """Body text that follows a card label (e.g. BEST PRACTICE) on the original slides."""
    for s in slides:
        shapes = [sh for sh in s.shapes if sh.has_text_frame]
        for i, sh in enumerate(shapes):
            if sh.text_frame.text.strip() in labels and i + 1 < len(shapes):
                return shapes[i + 1].text_frame.text.strip()
    return None


# ------------------------------------------------------------------ slide factory
TEMPLATE_PARTS = ("Rectangle 2", "TextBox 4", "TextBox 5", "TextBox 6", "Rounded Rectangle 7", "TextBox 8")


def new_slide(prs, tmpl):
    s = prs.slides.add_slide(tmpl.slide_layout)
    tree = s.shapes._spTree
    for el in list(tree):
        if el.tag.endswith("}sp") or el.tag.endswith("}pic") or el.tag.endswith("}graphicFrame"):
            tree.remove(el)
    bg = tmpl._element.cSld.find(qn("p:bg"))
    if bg is not None:
        s._element.cSld.insert(0, copy.deepcopy(bg))
    tn = names(tmpl)
    K.copy_shape(tn["Title 1"], s)
    for nm in TEMPLATE_PARTS:
        if nm in tn:
            K.copy_shape(tn[nm], s)
    if "Picture 3" in tn:
        K.copy_picture(tn["Picture 3"], tmpl, s)
    return s


def set_title(slide, text, where):
    t = names(slide)["Title 1"]
    for size in (28, 25, 22, 20):
        n = len(K.wrap(text, size, t.width / K.EMU - 0.15, True))
        if n == 1 or (size <= 22 and n <= 2):
            break
    if n > 2:
        K.WARN.append(f"TITLE {where}: '{text}' does not fit")
    K.set_runs_text(t, text, size)
    return size


def set_pill(slide, text):
    n = names(slide)
    w = K.text_width_pt(text, 12, True) / 72 + 0.34
    for nm in ("Rounded Rectangle 7", "TextBox 8"):
        if nm in n:
            n[nm].width = Emu(int(w * K.EMU))
    K.set_runs_text(n["TextBox 8"], text)


def ghost_template(cv, slide, title_size):
    """Display-only ops so previews show the copied template elements."""
    n = names(slide)
    cv.ops.insert(0, ("rect", 0, 0, 13.333, 0.14, K.ORANGE, None, False))
    cv.ops.append(("rect", 12.05, 0.38, 0.69, 0.62, "E8E4DA", None, True))
    t = n["Title 1"]
    cv.ops.append(("text", 0.70, 0.82, 10.9, 0.9, [t.text_frame.text], title_size, K.NAVY, True, "l", "t", 0))
    p = n["TextBox 8"]
    w = p.width / K.EMU
    cv.ops.append(("rect", 0.70, 0.40, w, 0.30, K.GREEN, None, True))
    cv.ops.append(("text", 0.70, 0.45, w, 0.25, [p.text_frame.text], 12, K.WHITE, True, "c", "t", 0))
    cv.ops.append(("text", 0.70, 7.05, 5.0, 0.3, ["ETEH Ltd  ·  Eaglet2EagleHub Limited"], 11, K.GREY, True, "l", "t", 0))
    cv.ops.append(("text", 5.20, 7.05, 5.8, 0.3, [n["TextBox 5"].text_frame.text], 11, K.GREY, False, "l", "t", 0))


# ------------------------------------------------------------------ the two topic slides
def numbered(items):
    return [f"**{i}.**  {x}" for i, x in enumerate(items, 1)]


def build_concept(prs, tmpl, t, title, tag, where, tip, trap):
    s = new_slide(prs, tmpl)
    ts = set_title(s, title, where)
    set_pill(s, f"{tag}  ·  1/2")
    cv = K.Canvas(s, where + " concept")
    x0, W = 0.70, 11.90
    y = 1.72
    cv.text(x0, y, W, 0.26, ["IN SIMPLE TERMS"], 11.5, K.GREEN, bold=True)
    y += 0.30
    what = t["what"] if isinstance(t["what"], list) else [t["what"]]
    h, _ = K.measure(what, 15, W, False, 0.08)
    cv.text(x0, y, W, h, what, 15, K.INK, gap=0.08)
    y += h + 0.20
    like = [t["like"]]
    h = cv.card_height(W, like, 14)
    cv.card(x0, y, W, h, K.AMBER, "THINK OF IT LIKE THIS", like, 14)
    y += h + 0.18
    how = t.get("how")
    tip = t.get("tip", tip)
    trap = t.get("trap", trap)
    side = [(K.GREEN, "REMEMBER", tip), (K.RED, "AVOID", trap)]
    side = [x for x in side if x[2]]
    if len(side) < 2:
        K.WARN.append(f"NOTE {where}: no {'REMEMBER' if not tip else 'AVOID'} text (add tip=/trap= to the topic copy)")
    if how:
        lw = 6.55 if side else W
        hl = cv.card_height(lw, numbered(how), 13)
        rw = W - lw - 0.20
        hs = [cv.card_height(rw, [txt], 13) for _, _, txt in side]
        row = max(hl, sum(hs) + 0.14 * (len(hs) - 1)) if side else hl
        cv.card(x0, y, lw, row, K.GREEN, "HOW TO DO IT", numbered(how), 13)
        if side:
            extra = (row - (sum(hs) + 0.14 * (len(hs) - 1))) / len(hs)
            yy = y
            for (acc, lab, txt), hh in zip(side, hs):
                cv.card(x0 + lw + 0.20, yy, rw, hh + extra, acc, lab, [txt], 13)
                yy += hh + extra + 0.14
        y += row
    elif side:
        cw = (W - 0.20) / len(side)
        hs = [cv.card_height(cw, [txt], 13) for _, _, txt in side]
        row = max(hs)
        for i, (acc, lab, txt) in enumerate(side):
            cv.card(x0 + i * (cw + 0.20), y, cw, row, acc, lab, [txt], 13)
        y += row
    if y > CONTENT_BOTTOM + 0.005:
        K.WARN.append(f"OVERFLOW {where} concept: ends at {y:.2f}in (limit {CONTENT_BOTTOM})")
    cv.text(7.60, 6.74, 5.00, 0.26, ["Worked example on the next slide  »"], 11.5, K.ORANGE, bold=True, align="r")
    ghost_template(cv, s, ts)
    return s, cv


def build_example(prs, tmpl, t, title, tag, where, tryit, nxt):
    s = new_slide(prs, tmpl)
    ts = set_title(s, title, where)
    set_pill(s, f"{tag}  ·  2/2")
    cv = K.Canvas(s, where + " example")
    # the marker: green badge with a tick, top right - the last slide of every topic group
    bx, by, bw, bh = 9.05, 0.38, 2.80, 0.40
    cv.rect(bx, by, bw, bh, K.GREEN, None, rounded=True, name="Worked Example Badge")
    cv.check(bx + 0.20, by + 0.11, 0.20, 0.18)
    cv.text(bx + 0.50, by + 0.075, bw - 0.60, 0.26, ["WORKED EXAMPLE"], 13, K.WHITE, bold=True, name="Worked Example Label")
    x0, W = 0.70, 11.90
    y = 1.72
    case = [t["case"]]
    h = cv.card_height(W, case, 14)
    cv.card(x0, y, W, h, K.NAVY, f"THE SITUATION  ·  {t['who']}", case, 14)
    y += h + 0.16
    steps = numbered(t["steps"])
    table = t.get("table")
    if table:
        lw = t.get("split", 5.55)
        rw = W - lw - 0.20
        hl = cv.card_height(lw, steps, 13)
        # table height is only known after building; measure with a throw-away estimate first
        tx = x0 + lw + 0.20
        cv.text(tx, y, rw, 0.26, ["WHAT YOU SHOULD SEE"], 11.5, K.NAVY, bold=True)
        th = cv.table(tx, y + 0.30, rw, table, 12, t.get("widths"))
        row = max(hl, th + 0.30)
        cv.card(x0, y, lw, row, K.GREEN, "STEP BY STEP", steps, 13)
        y += row + 0.16
    else:
        hl = cv.card_height(W, steps, 13)
        cv.card(x0, y, W, hl, K.GREEN, "STEP BY STEP", steps, 13)
        y += hl + 0.16
    so = [t["so"]]
    h = cv.card_height(W, so, 13.5)
    cv.card(x0, y, W, h, K.ORANGE, "WHAT THIS TELLS US", so, 13.5)
    y += h
    if y > EXAMPLE_BOTTOM + 0.005:
        K.WARN.append(f"OVERFLOW {where} example: ends at {y:.2f}in (limit {EXAMPLE_BOTTOM})")
    cv.text(x0, 6.68, 8.2, 0.26, [f"**TRY IT YOURSELF  ›**  {tryit}"], 11, K.GREY)
    cv.text(9.00, 6.68, 3.60, 0.26, [nxt], 11, K.GREY, align="r")
    ghost_template(cv, s, ts)
    return s, cv


# ------------------------------------------------------------------ deck-level slides
def build_guide(prs, tmpl, week, footer, companion):
    s = new_slide(prs, tmpl)
    ts = set_title(s, "How to Study With This Deck", "guide")
    set_pill(s, "START HERE")
    K.set_runs_text(names(s)["TextBox 5"], f"{footer}  ·  How to use this deck")
    cv = K.Canvas(s, f"{week} guide")
    routine = [
        "Read the module divider first: it tells you what the module is about.",
        "On every first slide, read **IN SIMPLE TERMS**, then the analogy. Say the idea back in your own words.",
        "Open Excel and follow **HOW TO DO IT** while you read. Note the REMEMBER and AVOID tips.",
        "The second slide, marked with the green **WORKED EXAMPLE** badge, ends every topic. Repeat the steps on the real file and check that you get the same result.",
        "After each module, practise on the matching sheet of the Section Examples workbook, then do the weekly assignment.",
    ]
    h = cv.card_height(6.0, numbered(routine), 13, gap=0.10)
    cv.card(0.70, 1.72, 6.0, h, K.GREEN, "YOUR STUDY ROUTINE", numbered(routine), 13, gap=0.10)
    lx, lw = 6.95, 5.65
    cv.text(lx, 1.72, lw, 0.26, ["WHAT EACH PART OF A SLIDE MEANS"], 11.5, K.NAVY, bold=True)
    legend = [
        (K.GREEN, "IN SIMPLE TERMS", "the idea in plain, everyday words"),
        (K.AMBER, "THINK OF IT LIKE THIS", "a comparison from daily life in Nigeria"),
        (K.GREEN, "HOW TO DO IT  ·  REMEMBER", "the clicks or formula, and the habit to keep"),
        (K.RED, "AVOID", "the mistake beginners make most often"),
        (K.NAVY, "THE SITUATION  ·  STEP BY STEP", "a real task on the bootcamp data files"),
        (K.ORANGE, "WHAT THIS TELLS US", "the lesson to take away from the example"),
    ]
    y = 2.05
    for acc, lab, txt in legend:
        cv.rect(lx, y, lw, 0.56, K.WHITE, K.BORDER)
        cv.rect(lx, y, 0.09, 0.56, acc)
        cv.text(lx + 0.25, y + 0.07, lw - 0.35, 0.22, [lab], 11, acc, bold=True)
        cv.text(lx + 0.25, y + 0.29, lw - 0.35, 0.22, [txt], 11.5, K.INK)
        y += 0.64
    cv.rect(lx, y + 0.06, 2.60, 0.40, K.GREEN, None, rounded=True)
    cv.check(lx + 0.20, y + 0.17, 0.20, 0.18)
    cv.text(lx + 0.50, y + 0.135, 2.0, 0.26, ["WORKED EXAMPLE"], 13, K.WHITE, bold=True)
    cv.text(lx + 2.75, y + 0.02, lw - 2.75, 0.50, ["This badge marks the last slide of every topic: the worked example."], 11.5, K.INK)
    data = ("**Data used in every example:** SuperMart Retail (12 stores), SuperMart Online, Crestview University, "
            "Lifecare General Hospital, EkoSuites Hotel and the ETEH Talent Salary Survey, all in "
            "01_Section_Examples/00_Core_Datasets. " + companion)
    dy = 1.72 + h + 0.20
    hh, _ = K.measure([data], 12.5, 5.70)
    cv.rect(0.70, dy, 6.0, hh + 0.30, "FFF4EC", None)
    cv.rect(0.70, dy, 0.09, hh + 0.30, K.ORANGE)
    cv.text(0.95, dy + 0.15, 5.70, hh, [data], 12.5, K.INK)
    if dy + hh + 0.30 > 6.85:
        K.WARN.append(f"OVERFLOW {week} guide: data note ends at {dy + hh + 0.30:.2f}in")
    ghost_template(cv, s, ts)
    return s, cv


def renumber(prs):
    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            if sh.has_text_frame and sh.top > Emu(int(6.9 * K.EMU)) and sh.text_frame.text.strip().isdigit():
                K.set_runs_text(sh, str(i))


def reorder(prs, order):
    lst = prs.slides._sldIdLst
    ids = {el.get("id"): el for el in lst}
    for el in list(lst):
        lst.remove(el)
    for sid in order:
        lst.append(ids[sid])


def drop(prs, slide):
    lst = prs.slides._sldIdLst
    for el in list(lst):
        if prs.part.related_part(el.rId) is slide.part:
            prs.part.drop_rel(el.rId)
            lst.remove(el)
            return


def sld_id(prs, slide):
    for el in prs.slides._sldIdLst:
        if prs.part.related_part(el.rId) is slide.part:
            return el.get("id")


def check_package(path, where):
    import zipfile
    with zipfile.ZipFile(path) as z:
        nm = z.namelist()
        dup = {n for n in nm if nm.count(n) > 1}
        if dup:
            K.WARN.append(f"PACKAGE {where}: duplicate parts {sorted(dup)[:4]}")
        bad = z.testzip()
        if bad:
            K.WARN.append(f"PACKAGE {where}: corrupt member {bad}")


def fit_next(head, title, width_in=3.55, size=11):
    words = title.split()
    for n in range(len(words), 0, -1):
        s = head + " ".join(words[:n]) + ("" if n == len(words) else "…")
        if K.text_width_pt(s, size) / 72 <= width_in:
            return s
    return head.strip()


def short(s, n=46):
    return s if len(s) <= n else s[: n - 1].rsplit(" ", 1)[0] + "…"


# ------------------------------------------------------------------ main
def revise(vol, fname, xlsx, week, asg, topics, preview_tags, log):
    path = PACK / fname
    prs = Presentation(str(V1 / fname))
    segs = scan(prs)
    sheets = sheet_map(xlsx)
    tmpl = next(sg[2][0] for sg in segs if sg[0] == "topic")
    topic_segs = [sg for sg in segs if sg[0] == "topic"]
    nxt_of = {}
    for i, sg in enumerate(topic_segs):
        if i + 1 < len(topic_segs):
            n2 = topic_segs[i + 1]
            same = n2[3][0] == sg[3][0]
            title2 = names(n2[2][0])["Title 1"].text_frame.text.replace(" (continued)", "")
            nxt_of[sg[1]] = fit_next(f"Next  ›  {n2[1]} ", title2) if same else f"Next  ›  Module {n2[3][0]}"
        else:
            nxt_of[sg[1]] = f"End of {week} topics"
    order, done, kept = [], 0, 0
    companion = (f"Practise with {xlsx}" + (f" and {asg}." if asg else " and the capstone projects."))
    for sg in segs:
        if sg[0] == "other":
            order.append(sld_id(prs, sg[1][0]))
            if sg[1][0] is prs.slides[1]:
                g, cv = build_guide(prs, tmpl, week, f"{week}", companion)
                order.append(sld_id(prs, g))
                if "guide" in preview_tags:
                    PREVIEW.mkdir(parents=True, exist_ok=True)
                    K.render(cv.ops, PREVIEW / f"V{vol}_guide.png")
            continue
        _, tag, slides, mod = sg
        t = topics.get(tag)
        if t is None:
            order += [sld_id(prs, s) for s in slides]
            kept += 1
            continue
        title = t.get("title") or names(slides[0])["Title 1"].text_frame.text.replace(" (continued)", "")
        where = f"V{vol} {tag}"
        tip = orig_text(slides, ("BEST PRACTICE", "PROFESSIONAL BEST PRACTICES"))
        trap = orig_text(slides, ("WATCH OUT", "COMMON MISTAKES TO AVOID"))
        code = tag.rsplit(".", 1)[0]
        sh = sheets.get(code, [])
        tryit = t.get("tryit") or (f"{xlsx}  ›  sheet '{sh[0]}'" if sh else xlsx)
        c, cvc = build_concept(prs, slides[0], t, title, tag, where, tip, trap)
        e, cve = build_example(prs, slides[0], t, title, tag, where, tryit, nxt_of[tag])
        order += [sld_id(prs, c), sld_id(prs, e)]
        if mod:
            foot = f"Module {mod[0]}  ·  {mod[1]}"
            for s_ in (c, e):
                tb5 = names(s_).get("TextBox 5")
                if tb5 is not None and K.text_width_pt(foot, 11) / 72 <= tb5.width / K.EMU - 0.1:
                    K.set_runs_text(tb5, foot)
        for s in slides:
            drop(prs, s)
        done += 1
        if tag in preview_tags or "all" in preview_tags or f"V{vol}" in preview_tags:
            PREVIEW.mkdir(parents=True, exist_ok=True)
            K.render(cvc.ops, PREVIEW / f"{tag}_1.png")
            K.render(cve.ops, PREVIEW / f"{tag}_2.png")
    reorder(prs, order)
    # IN THIS SESSION panel (slide 2)
    s2 = names(prs.slides[1]).get("TextBox 13")
    n_top = len(topic_segs)
    n_mod = len({sg[3][0] for sg in topic_segs})
    if s2 is not None:
        tf = s2.text_frame
        txt = (f"{n_top} topics across {n_mod} modules.\n\nEvery topic has two slides: the idea in simple terms with an "
               f"everyday analogy, then a worked example on real Nigerian business data, marked WORKED EXAMPLE.\n\n"
               f"Start with 'How to Study With This Deck' on the next slide.")
        paras = txt.split("\n")
        p0 = tf.paragraphs[0]
        for p in tf.paragraphs[1:]:
            p._p.getparent().remove(p._p)
        r0 = p0.runs[0]
        r0.text = paras[0]
        for r in p0.runs[1:]:
            r._r.getparent().remove(r._r)
        for ptxt in paras[1:]:
            np_ = copy.deepcopy(p0._p)
            tf._txBody.append(np_)
            runs = np_.findall(qn("a:r"))
            runs[0].find(qn("a:t")).text = ptxt
        need, _ = K.measure(paras, 17, s2.width / K.EMU)
        if need > s2.height / K.EMU + 0.05:
            K.WARN.append(f"OVERFLOW V{vol} S2 panel: {need:.2f}in")
    renumber(prs)
    # new slides reuse part names of dropped ones: give every slide (and its notes) a fresh, sequential name
    for i, s in enumerate(prs.slides, 1):
        s.part.partname = PackURI(f"/ppt/slides/slide{i}.xml")
        if s.has_notes_slide:
            s.notes_slide.part.partname = PackURI(f"/ppt/notesSlides/notesSlide{i}.xml")
    prs.save(str(path))
    check_package(path, where=f"V{vol}")
    log.append(f"V{vol}: {done} topics revised, {kept} kept as v1, {len(prs.slides)} slides -> {path.relative_to(ROOT)}")


def main():
    args = sys.argv[1:]
    preview = []
    if "--preview" in args:
        preview = args[args.index("--preview") + 1:]
    if "--skip05" not in args or not V1.exists():
        subprocess.run([sys.executable, str(ROOT / "_build" / "05_update_slides.py")], check=True, stdout=subprocess.DEVNULL)
        V1.mkdir(parents=True, exist_ok=True)
        for _, fname, *_ in DECKS:
            shutil.copy2(PACK / fname, V1 / fname)
    topics = load_topics()
    only = [a for a in args if a.startswith("--vol=")]
    vols = {int(a.split("=")[1]) for a in only} if only else {1, 2, 3, 4, 5}
    log = []
    for vol, fname, xlsx, week, asg in DECKS:
        if vol in vols:
            revise(vol, fname, xlsx, week, asg, topics, set(preview), log)
    print("\n".join(log))
    print(f"topic copy loaded: {len(topics)}")
    if K.WARN:
        print("\nWARNINGS:")
        print("\n".join(K.WARN))


if __name__ == "__main__":
    main()
