"""Edition-3 deck engine shared by every week (config-driven).

A week script builds a Config and calls build(cfg). Copy lives in _build/v3/wN_*.py; example tables come from
the week's data module (wN_data.py). With --preview every slide is also rendered to cfg.preview/NNN.png for QA.
"""
from __future__ import annotations

import copy
import sys
from dataclasses import dataclass, field
from pathlib import Path

from pptx import Presentation
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn
from pptx.util import Emu

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE / "v3"))
import slidekit as K  # noqa: E402
import v3kit as V  # noqa: E402
from slidekit import AMBER, BORDER, GREEN, GREY, INK, NAVY, ORANGE, RED, WHITE, measure  # noqa: E402

ROOT = HERE.parent
TEMPLATE = HERE / "templates" / "ETEH_deck_template.pptx"


@dataclass
class Config:
    week: int
    week_title: str                 # "Week 1: Data Analytics Foundations & Core Microsoft Excel Mastery"
    title_sub: str                  # subtitle on the title slide
    out: Path
    workbook: str                   # examples workbook file name
    modules: list
    sheets: dict                    # module id -> workbook sheet name
    data: object                    # week data module (excerpt/table)
    summary: dict
    glossary: list
    roadmap_label: str
    roadmap_notes: str
    title_notes: str
    assignment: str
    volume: str = ""                # "Volume 1 of 5"
    modules_text: str = ""          # welcome slide wording, default "<n> modules"
    example_table: str = "tStores25"  # named table quoted on the How This Week Works slide
    unit: str = ""                  # "Capstone": replaces "Week N" in the standard slides (default "Week <week>")
    assessment: str = ""            # replaces the "Assessment: the Week N assignment ..." line
    preview: Path = field(default_factory=lambda: HERE / "_stage" / "v3_preview")


CFG: Config = None   # set by build()
D = None
MODULES = []
WEEK = ""
WORKBOOK = ""


DARK = "00213F"
FRAME_PARTS = ("Rectangle 2", "TextBox 4", "TextBox 5", "TextBox 6", "Rounded Rectangle 7", "TextBox 8")


# ---------------------------------------------------------------- slide plumbing
def names(slide):
    return {sh.name: sh for sh in slide.shapes}


class Deck:
    def __init__(self):
        self.prs = Presentation(str(TEMPLATE))
        s = list(self.prs.slides)
        self.t_title, self.t_welcome, self.t_div, self.t_frame, self.t_end = s
        self.n_template = len(s)
        self.previews = []          # (ops, dark) per slide in final order
        self.order = []

    def clone(self, src, keep=None):
        s = self.prs.slides.add_slide(src.slide_layout)
        tree = s.shapes._spTree
        for el in list(tree):
            if el.tag.endswith("}sp") or el.tag.endswith("}pic") or el.tag.endswith("}graphicFrame"):
                tree.remove(el)
        bg = src._element.cSld.find(qn("p:bg"))
        if bg is not None:
            s._element.cSld.insert(0, copy.deepcopy(bg))
        for sh in src.shapes:
            if keep is not None and sh.name not in keep:
                continue
            if sh.shape_type == 13:
                K.copy_picture(sh, src, s)
            else:
                K.copy_shape(sh, s)
        return s

    def finish(self, path):
        # drop the template slides, renumber, give parts sequential names
        lst = self.prs.slides._sldIdLst
        for el in list(lst)[: self.n_template]:
            self.prs.part.drop_rel(el.rId)
            lst.remove(el)
        for i, s in enumerate(self.prs.slides, 1):
            for sh in s.shapes:
                if sh.has_text_frame and sh.top > Emu(int(6.9 * K.EMU)) and sh.text_frame.text.strip() == "#":
                    K.set_runs_text(sh, str(i))
            s.part.partname = PackURI(f"/ppt/slides/slide{i}.xml")
            if s.has_notes_slide:
                s.notes_slide.part.partname = PackURI(f"/ppt/notesSlides/notesSlide{i}.xml")
        self.prs.save(str(path))


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def fit_title(shape, text, where, sizes=(28, 26, 24, 22)):
    w = shape.width / K.EMU - 0.15
    for size in sizes:
        n = len(K.wrap(text, size, w, True))
        if n == 1 or (size <= 24 and n <= 2):
            break
    if n > 2:
        K.WARN.append(f"TITLE {where}: '{text}' does not fit")
    K.set_runs_text(shape, text, size)
    return size, n


def pill(slide, text):
    n = names(slide)
    w = K.text_width_pt(text, 12, True) / 72 + 0.34
    for nm in ("Rounded Rectangle 7", "TextBox 8"):
        n[nm].width = Emu(int(w * K.EMU))
    K.set_runs_text(n["TextBox 8"], text)
    return w


def frame(deck, title, tag, footer, where):
    s = deck.clone(deck.t_frame, keep=("Title 1",) + FRAME_PARTS)
    n = names(s)
    size, lines = fit_title(n["Title 1"], title, where)
    pw = pill(s, tag)
    K.set_runs_text(n["TextBox 5"], footer)
    K.set_runs_text(n["TextBox 6"], "#")
    cv = K.Canvas(s, where)
    # display-only ops for previews (template elements are copied XML, not canvas shapes)
    cv.ops.append(("rect", 0, 0, 13.333, 0.14, ORANGE, None, False))
    cv.ops.append(("rect", 12.05, 0.38, 0.69, 0.62, "E8E4DA", None, True))
    cv.ops.append(("text", 0.70, 0.82, 10.9, 0.9, [title], size, NAVY, True, "l", "t", 0))
    cv.ops.append(("rect", 0.70, 0.40, pw, 0.30, GREEN, None, True))
    cv.ops.append(("text", 0.70, 0.45, pw, 0.25, [tag], 12, WHITE, True, "c", "t", 0))
    cv.ops.append(("text", 0.70, 7.05, 5.0, 0.3, ["ETEH Ltd  ·  Eaglet2EagleHub Limited"], 11, GREY, True, "l", "t", 0))
    cv.ops.append(("text", 5.20, 7.05, 5.8, 0.3, [footer], 11, GREY, False, "l", "t", 0))
    top = 1.55 if lines == 1 else 1.55 + size * K.LINE / 72
    deck.previews.append((cv, False))
    return s, cv, top


def dark_preview(texts, logo):
    ops = [("rect", 0, 0, 13.333, 7.5, DARK, None, False), ("rect", 0, 0, 13.333, 0.14, ORANGE, None, False),
           ("rect", *logo, "E8E4DA", None, True)]
    for x, y, w, h, t, size, col, bold in texts:
        ops.append(("text", x, y, w, h, [t], size, col, bold, "l", "t", 0))
    return ops


class Ops:
    def __init__(self, ops):
        self.ops = ops


# ---------------------------------------------------------------- standard slides
def U():
    return CFG.unit or f"Week {CFG.week}"


def HOW():
    return "How This Week Works" if not CFG.unit else f"How the {CFG.unit} Works"


def title_slide(deck):
    s = deck.clone(deck.t_title)
    n = names(s)
    K.set_runs_text(n["TextBox 5"], WEEK)
    K.set_runs_text(n["TextBox 6"], CFG.title_sub)
    if CFG.volume:
        K.set_runs_text(n["TextBox 4"], f"FOUR-WEEK EXCEL DATA ANALYTICS BOOTCAMP  ·  {CFG.volume}")
    K.set_runs_text(n["TextBox 7"], "Eaglet2EagleHub Limited  ·  Online Tutorial Slides  ·  Edition 3")
    notes(s, CFG.title_notes)
    deck.previews.append((Ops(dark_preview([
        (2.5, 1.0, 9, .4, "ETEH Ltd  ·  TRAINING PROGRAMME", 16, ORANGE, True),
        (2.5, 1.4, 9, .4, n["TextBox 4"].text_frame.text, 14, "A9B6C4", False),
        (0.8, 2.9, 11.7, 2.2, WEEK, 46, WHITE, True),
        (0.8, 5.2, 11.5, .9, n["TextBox 6"].text_frame.text, 20, "E6EAEF", False),
        (0.8, 6.7, 10, .4, n["TextBox 7"].text_frame.text, 13, "A9B6C4", False)], (0.8, 0.8, 1.44, 1.3))), True))


def welcome_slide(deck, n_topics):
    s = deck.clone(deck.t_welcome)
    n = names(s)
    K.set_runs_text(n["TextBox 6"], "#")
    txt = (f"{n_topics} topics across {CFG.modules_text or f'{len(MODULES)} modules'}.\n\nEach topic gets the parts it needs: a standard "
           f"definition, a plain-English explanation, an analogy, and a small example from our datasets where it helps.\n\n"
           f"Practical work happens in the {U()} examples workbook at the end of each module.")
    tb = n["TextBox 13"]
    tf = tb.text_frame
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
        np_.findall(qn("a:r"))[0].find(qn("a:t")).text = ptxt
    need, _ = measure(paras, 17, tb.width / K.EMU)
    if need > tb.height / K.EMU + 0.05:
        K.WARN.append(f"OVERFLOW welcome panel: {need:.2f}in")
    ops = [("rect", 0, 0, 13.333, 0.14, ORANGE, None, False),
           ("text", 0.70, 0.82, 10.9, 0.9, ["Welcome and Your Lecturer"], 28, NAVY, True, "l", "t", 0),
           ("rect", 0.7, 2.1, 5.6, 3.6, WHITE, None, False), ("rect", 0.7, 2.1, 0.1, 3.6, ORANGE, None, False),
           ("text", 1.1, 2.4, 5, .4, ["YOUR LECTURER"], 13, ORANGE, True, "l", "t", 0),
           ("text", 1.1, 2.85, 5, .6, ["Abdulrahaman Kehinde"], 26, NAVY, True, "l", "t", 0),
           ("text", 1.1, 3.6, 5, 1.8, n["TextBox 11"].text_frame.text.split("\n"), 16, INK, False, "l", "t", 0),
           ("text", 6.9, 2.1, 5.7, .4, ["IN THIS SESSION"], 13, ORANGE, True, "l", "t", 0),
           ("text", 6.9, 2.55, 5.7, 3.3, paras, 17, INK, False, "l", "t", 0)]
    deck.previews.append((Ops(ops), False))


def roadmap_slide(deck):
    W_ = U().upper()
    s, cv, top = frame(deck, f"{U()} Roadmap", W_, f"{U()} overview", "roadmap")
    rows = [["Module", "Title", "Topics", "Examples sheet"]]
    for m in MODULES:
        rows.append([m["id"], m["title"], str(len(m["topics"])), m["sheet"]])
    cv.text(V.X0, top, V.W, 0.26, [CFG.roadmap_label], 11.5, ORANGE, bold=True)
    size = 14 if len(rows) <= 10 else 12.5
    h = cv.table(V.X0, top + 0.32, V.W, rows, size, widths=[1.0, 5.4, 0.9, 2.6])
    if top + 0.32 + h > V.BOTTOM + 0.01:
        K.WARN.append(f"OVERFLOW roadmap: {top + 0.32 + h:.2f}")
    notes(s, CFG.roadmap_notes)


def how_slide(deck):
    s, cv, top = frame(deck, HOW(), U().upper(), f"{U()} overview", "how")
    y = top
    wk = "Week_5_Capstone" if CFG.week == 5 else f"Week_{CFG.week}"
    msg = (f"After each module, open that module's sheet in **01_Section_Examples › {wk} › {WORKBOOK}**. Each sheet "
           f"asks only about what you have learned so far. The gold box names the data (File › Sheet › Table, or the "
           f"Ex_Files folder next to the workbook); column F checks your answer at once; the Solutions sheet explains every answer.")
    h = V.card_h(V.W, msg, 15)
    V.card(cv, V.X0, y, V.W, h, ORANGE, f"PRACTICAL WORK: THE {U().upper()} EXAMPLES WORKBOOK", msg, 15)
    y += h + 0.22
    cv.text(V.X0, y, V.W, 0.26, ["HOW EACH TOPIC IS PRESENTED (ONLY THE PARTS A TOPIC NEEDS)"], 11.5, NAVY, bold=True)
    y += 0.34
    legend = [(NAVY, "DEFINITION", "The standard, textbook definition."),
              (GREEN, "IN SIMPLE TERMS", "The same idea in plain English."),
              (AMBER, "ANALOGY", "A comparison from everyday Nigerian life."),
              (ORANGE, "EXAMPLE FROM OUR DATA", f"A few rows of a named table, e.g. {CFG.example_table}."),
              (V.TEAL, "COMPARISON / FORMULAS", "Side-by-side facts or formulas with results."),
              (RED, "WATCH OUT", "The most common mistake to avoid.")]
    wc = (V.W - 0.5) / 3
    for i, (c, lab, txt) in enumerate(legend):
        r, k = divmod(i, 3)
        x, yy = V.X0 + k * (wc + 0.25), y + r * 1.15
        V.card(cv, x, yy, wc, 1.0, c, lab, txt, 13)
    y += 2.35
    t = CFG.assessment or f"Assessment: the Week {CFG.week} assignment (02_Assignments › Week_{CFG.week} › {CFG.assignment}) is set at the end of the week."
    cv.text(V.X0, y, V.W, 0.3, [t], 13, GREY)
    notes(s, "This is the only place the deck mentions the workbook. After every module, switch to Excel and open the "
             "module's sheet: go through any worked example, then let students answer; open the Solutions sheet only after they have tried.")


def divider(deck, m):
    s = deck.clone(deck.t_div)
    n = names(s)
    lab = f"MODULE {m['id']}  ·  {len(m['topics'])} TOPICS"
    K.set_runs_text(n["TextBox 3"], lab)
    K.set_runs_text(n["TextBox 4"], m["title"])
    K.set_runs_text(n["TextBox 5"], m["subtitle"])
    K.set_runs_text(n["TextBox 8"], "#")
    need, _ = measure([m["title"]], 44, 10.8, True)
    if need > 2.2:
        K.WARN.append(f"DIVIDER {m['id']} title too long")
    need, _ = measure([m["subtitle"]], 18, 10.5)
    if need > 1.2:
        K.WARN.append(f"DIVIDER {m['id']} subtitle too long")
    deck.previews.append((Ops(dark_preview([
        (0.8, 1.8, 8, .5, lab, 16, ORANGE, True), (0.8, 2.4, 10.8, 2.2, m["title"], 44, WHITE, True),
        (0.8, 4.7, 10.5, 1.2, m["subtitle"], 18, "E6EAEF", False)], (11.6, 0.5, 1.0, 0.9))), True))


def objectives_slide(deck, m):
    foot = f"Module {m['id']}  ·  {m['title']}"
    s, cv, top = frame(deck, "What You Will Learn", f"MODULE {m['id']}", foot, f"{m['id']} objectives")
    obj = [f"**{i}.**  {o}" for i, o in enumerate(m["objectives"], 1)]
    wl = 6.2
    h, _ = measure(obj, 15, wl - 0.5, False, 0.12)
    h += 0.56
    cv.rect(V.X0, top, wl, h, WHITE, BORDER)
    cv.rect(V.X0, top, 0.09, h, GREEN)
    cv.text(V.X0 + 0.3, top + 0.12, wl - 0.45, 0.26, ["BY THE END OF THIS MODULE YOU CAN"], 11.5, GREEN, bold=True)
    cv.text(V.X0 + 0.3, top + 0.44, wl - 0.5, h - 0.5, obj, 15, INK, gap=0.12)
    xr = V.X0 + wl + 0.3
    wr = V.W - wl - 0.3
    tl = [f"**{t['tag']}**  {t['title']}" for t in m["topics"]]
    size = 12.5 if len(tl) <= 12 else 11.5
    if len(tl) <= 20:
        hh, _ = measure(tl, size, wr - 0.5, False, 0.03)
        cols = [tl]
    else:                                   # long lists (e.g. 33 functions): three columns
        per = -(-len(tl) // 3)
        cols = [tl[i:i + per] for i in range(0, len(tl), per)]
        cw = (wr - 0.5) / 3
        hh = max(measure(c, size, cw - 0.05, False, 0.03)[0] for c in cols)
    cv.rect(xr, top, wr, hh + 0.56, WHITE, BORDER)
    cv.rect(xr, top, 0.09, hh + 0.56, NAVY)
    cv.text(xr + 0.3, top + 0.12, wr - 0.45, 0.26, ["TOPICS"], 11.5, NAVY, bold=True)
    if len(cols) == 1:
        cv.text(xr + 0.3, top + 0.44, wr - 0.5, hh, tl, size, INK, gap=0.03)
    else:
        for i, c in enumerate(cols):
            cv.text(xr + 0.3 + i * cw, top + 0.44, cw - 0.05, hh, c, size, INK, gap=0.03)
    if top + max(h, hh + 0.56) > V.BOTTOM + 0.01:
        K.WARN.append(f"OVERFLOW {m['id']} objectives")


def topic_slide(deck, m, t):
    foot = f"Module {m['id']}  ·  {m['title']}"
    s, cv, top = frame(deck, t["title"], t["tag"], foot, t["tag"])
    V.TOP = top
    V.layout_topic(cv, t, D, t["tag"])
    if t.get("notes"):
        notes(s, t["notes"])
    # optional continuation slides (e.g. a chart or dashboard mock that needs the full width); their pill is not a bare
    # topic tag, so the deck lint still sees exactly one slide per mini-section
    for extra in t.get("more", []):
        tt = dict(extra)
        s2, cv2, top2 = frame(deck, tt.get("title", t["title"]), f"{t['tag']}  ·  {tt.get('pill', 'IN PRACTICE')}", foot, t["tag"] + "+")
        V.TOP = top2
        V.layout_topic(cv2, tt, D, t["tag"] + "+")
        if tt.get("notes"):
            notes(s2, tt["notes"])


def takeaways_slide(deck, m):
    foot = f"Module {m['id']}  ·  {m['title']}"
    s, cv, top = frame(deck, f"Key Takeaways: {m['title']}", f"MODULE {m['id']}  ·  RECAP", foot, f"{m['id']} recap")
    y = top
    for i, k in enumerate(m["takeaways"], 1):
        h, _ = measure([k], 16, V.W - 1.0)
        h = max(h, 0.3) + 0.30
        cv.rect(V.X0, y, V.W, h, WHITE, BORDER)
        cv.rect(V.X0, y, 0.55, h, GREEN)
        cv.check(V.X0 + 0.16, y + h / 2 - 0.11, 0.24, 0.22)
        cv.text(V.X0 + 0.80, y + 0.15, V.W - 1.0, h - 0.3, [k], 16, INK)
        y += h + 0.14
    if y > V.BOTTOM + 0.01:
        K.WARN.append(f"OVERFLOW {m['id']} takeaways")
    notes(s, f"Now switch to Excel: open {WORKBOOK}, sheet '{m['sheet']}'. Go through any worked example "
             f"together, then let students answer the questions (column F checks each answer). Open the Solutions sheet "
             f"only after students have tried.")


def summary_slide(deck):
    s, cv, top = frame(deck, f"{U()} Summary", f"{U().upper()}  ·  SUMMARY", f"{U()} summary", "summary")
    rows = [["Module", "What you can now do"]] + [[f"{m['id']}  {m['title']}", CFG.summary[m["id"]]] for m in MODULES]
    size = 13.5 if len(rows) <= 10 else 12
    h = cv.table(V.X0, top, V.W, rows, size, widths=[3.9, 8.0])
    if top + h > V.BOTTOM + 0.01:
        K.WARN.append(f"OVERFLOW summary: {top + h:.2f}")
    notes(s, f"Recap the week in two minutes, then point students to the Week {CFG.week} assignment." if not CFG.unit else
          "Recap the capstone modules in two minutes, then confirm the capstone deliverables and defence dates.")


def glossary_slides(deck):
    per = 15
    chunks = [CFG.glossary[i:i + per] for i in range(0, len(CFG.glossary), per)]
    for k, ch in enumerate(chunks, 1):
        s, cv, top = frame(deck, f"{U()} Glossary" + (f" ({k} of {len(chunks)})" if len(chunks) > 1 else ""),
                           f"{U().upper()}  ·  GLOSSARY", f"{U()} glossary", f"glossary {k}")
        rows = [["Term", "Meaning"]] + [[f"**{a}**", b] for a, b in ch]
        h = cv.table(V.X0, top, V.W, rows, 13, widths=[2.6, 9.3])
        if top + h > V.BOTTOM + 0.01:
            K.WARN.append(f"OVERFLOW glossary {k}: {top + h:.2f}")


def end_slide(deck):
    s = deck.clone(deck.t_end)
    n = names(s)
    K.set_runs_text(n["TextBox 4"], WEEK)
    deck.previews.append((Ops(dark_preview([(1.0, 3.1, 11.3, .9, "Thank You", 48, WHITE, True),
                                            (1.0, 4.1, 11.3, .9, WEEK, 18, "E6EAEF", False),
                                            (1.0, 6.6, 11.3, .4, "ETEH Ltd  ·  Eaglet2EagleHub Limited", 13, ORANGE, True)],
                                           (5.9, 1.2, 1.66, 1.5))), True))


# ---------------------------------------------------------------- main
def build(cfg: Config, argv=None):
    global CFG, D, MODULES, WEEK, WORKBOOK
    argv = sys.argv if argv is None else argv
    CFG, D, MODULES, WEEK, WORKBOOK = cfg, cfg.data, cfg.modules, cfg.week_title, cfg.workbook
    for m in MODULES:
        m["sheet"] = cfg.sheets[m["id"]]
    K.WARN.clear()
    preview = "--preview" in argv
    deck = Deck()
    n_topics = sum(len(m["topics"]) for m in MODULES)
    title_slide(deck)
    welcome_slide(deck, n_topics)
    roadmap_slide(deck)
    how_slide(deck)
    for m in MODULES:
        divider(deck, m)
        objectives_slide(deck, m)
        for t in m["topics"]:
            topic_slide(deck, m, t)
        takeaways_slide(deck, m)
    summary_slide(deck)
    glossary_slides(deck)
    end_slide(deck)
    deck.finish(cfg.out)
    print(f"{len(deck.prs.slides)} slides, {n_topics} topics -> {cfg.out.relative_to(ROOT)}")
    if preview:
        cfg.preview.mkdir(parents=True, exist_ok=True)
        for f in cfg.preview.glob("*.png"):
            f.unlink()
        for i, (cv, _) in enumerate(deck.previews, 1):
            K.render(cv.ops, cfg.preview / f"{i:03d}.png", dpi=80)
    if K.WARN:
        print("\nWARNINGS:")
        print("\n".join(K.WARN))
    return deck
