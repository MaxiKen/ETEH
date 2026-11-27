"""Edition 3: build the Week 1 deck from scratch.

    python3 _build/07_v3_week1_deck.py [--preview]

Starts from _build/templates/ETEH_deck_template.pptx (5 ETEH template slides: title, welcome, module divider,
content frame, closing), and writes ETEH_Bootcamp_Pack/ETEH_Vol1_Week1_Foundations_Slides.pptx.
Copy lives in _build/v3/w1_*.py; example tables come from _build/w1_data.py.
With --preview every slide is also rendered to _build/_stage/v3_preview/NNN.png for visual QA.
"""
from __future__ import annotations

import copy
import sys
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
import w1_data as D  # noqa: E402
from slidekit import AMBER, BORDER, GREEN, GREY, INK, NAVY, ORANGE, RED, WHITE, measure  # noqa: E402
from w1_m1_4 import MODULES as MA  # noqa: E402
from w1_m5_6 import MODULES as MB  # noqa: E402
from w1_m7_9 import MODULES as MC  # noqa: E402
from w1_extra import GLOSSARY, SUMMARY  # noqa: E402

ROOT = HERE.parent
TEMPLATE = HERE / "templates" / "ETEH_deck_template.pptx"
OUT = ROOT / "ETEH_Bootcamp_Pack" / "ETEH_Vol1_Week1_Foundations_Slides.pptx"
PREVIEW = HERE / "_stage" / "v3_preview"
WORKBOOK = "ETEH_Week1_Section_Examples.xlsx"
WEEK = "Week 1: Data Analytics Foundations & Core Microsoft Excel Mastery"
MODULES = MA + MB + MC
SHEETS = {"1.1": "1.1 Intro to Analytics", "1.2": "1.2 Types of Analytics", "1.3": "1.3 Analytics Process",
          "1.4": "1.4 Industry Applications", "1.5": "1.5 Excel Interface", "1.6": "1.6 Working with Data",
          "1.7": "1.7 Formulas & Functions", "1.8": "1.8 Excel Tables", "1.9": "1.9 Tidy Data Laws"}
for _m in MODULES:
    _m["sheet"] = SHEETS[_m["id"]]
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
def title_slide(deck):
    s = deck.clone(deck.t_title)
    n = names(s)
    K.set_runs_text(n["TextBox 5"], WEEK)
    K.set_runs_text(n["TextBox 6"], "Analytics thinking, the Excel workspace, formulas, Excel tables and tidy data")
    K.set_runs_text(n["TextBox 7"], "Eaglet2EagleHub Limited  ·  Online Tutorial Slides  ·  Edition 3")
    notes(s, "Welcome students. This week builds the foundations every later week depends on.")
    deck.previews.append((Ops(dark_preview([
        (2.5, 1.0, 9, .4, "ETEH Ltd  ·  TRAINING PROGRAMME", 16, ORANGE, True),
        (2.5, 1.4, 9, .4, "FOUR-WEEK EXCEL DATA ANALYTICS BOOTCAMP  ·  Volume 1 of 5", 14, "A9B6C4", False),
        (0.8, 2.9, 11.7, 2.2, WEEK, 46, WHITE, True),
        (0.8, 5.2, 11.5, .9, n["TextBox 6"].text_frame.text, 20, "E6EAEF", False),
        (0.8, 6.7, 10, .4, n["TextBox 7"].text_frame.text, 13, "A9B6C4", False)], (0.8, 0.8, 1.44, 1.3))), True))


def welcome_slide(deck, n_topics):
    s = deck.clone(deck.t_welcome)
    n = names(s)
    K.set_runs_text(n["TextBox 6"], "#")
    txt = (f"{n_topics} topics across {len(MODULES)} modules.\n\nEach topic gets the parts it needs: a standard "
           f"definition, a plain-English explanation, an analogy, and a small example from our datasets where it helps.\n\n"
           f"Practical work happens in the Week 1 examples workbook at the end of each module.")
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
    s, cv, top = frame(deck, "Week 1 Roadmap", "WEEK 1", "Week 1 overview", "roadmap")
    rows = [["Module", "Title", "Topics", "Examples sheet"]]
    for m in MODULES:
        rows.append([m["id"], m["title"], str(len(m["topics"])), m["sheet"]])
    cv.text(V.X0, top, V.W, 0.26, ["NINE MODULES THIS WEEK"], 11.5, ORANGE, bold=True)
    cv.table(V.X0, top + 0.32, V.W, rows, 14, widths=[1.0, 5.4, 0.9, 2.6])
    notes(s, "Walk through the nine modules. Modules 1.1-1.4 are thinking skills; 1.5-1.9 are hands-on Excel.")


def how_slide(deck):
    s, cv, top = frame(deck, "How This Week Works", "WEEK 1", "Week 1 overview", "how")
    y = top
    msg = (f"After each module, your tutor opens that module's sheet in **{WORKBOOK}** and works through the "
           f"prepared dataset, the step-by-step solutions and the exercises with you. Keep the workbook open "
           f"alongside these slides; each sheet starts with the module's dataset and ends with exercises and answers.")
    h = V.card_h(V.W, msg, 15)
    V.card(cv, V.X0, y, V.W, h, ORANGE, "PRACTICAL WORK: THE WEEK 1 EXAMPLES WORKBOOK", msg, 15)
    y += h + 0.22
    cv.text(V.X0, y, V.W, 0.26, ["HOW EACH TOPIC IS PRESENTED (ONLY THE PARTS A TOPIC NEEDS)"], 11.5, NAVY, bold=True)
    y += 0.34
    legend = [(NAVY, "DEFINITION", "The standard, textbook definition."),
              (GREEN, "IN SIMPLE TERMS", "The same idea in plain English."),
              (AMBER, "ANALOGY", "A comparison from everyday Nigerian life."),
              (ORANGE, "EXAMPLE FROM OUR DATA", "A few rows of a named table, e.g. tStores25."),
              (V.TEAL, "COMPARISON / FORMULAS", "Side-by-side facts or formulas with results."),
              (RED, "WATCH OUT", "The most common mistake to avoid.")]
    wc = (V.W - 0.5) / 3
    for i, (c, lab, txt) in enumerate(legend):
        r, k = divmod(i, 3)
        x, yy = V.X0 + k * (wc + 0.25), y + r * 1.15
        V.card(cv, x, yy, wc, 1.0, c, lab, txt, 13)
    y += 2.35
    t = "Assessment: the Week 1 assignment (02_Assignments/ETEH_Week1_Assignment.xlsx) is set at the end of the week."
    cv.text(V.X0, y, V.W, 0.3, [t], 13, GREY)
    notes(s, "This is the only place the deck mentions the workbook. After every module, switch to Excel and open the "
             "module's sheet: work the guided example, then let students try the exercises before revealing answers.")


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
    hh, _ = measure(tl, size, wr - 0.5, False, 0.03)
    cv.rect(xr, top, wr, hh + 0.56, WHITE, BORDER)
    cv.rect(xr, top, 0.09, hh + 0.56, NAVY)
    cv.text(xr + 0.3, top + 0.12, wr - 0.45, 0.26, ["TOPICS"], 11.5, NAVY, bold=True)
    cv.text(xr + 0.3, top + 0.44, wr - 0.5, hh, tl, size, INK, gap=0.03)
    if top + max(h, hh + 0.56) > V.BOTTOM + 0.01:
        K.WARN.append(f"OVERFLOW {m['id']} objectives")


def topic_slide(deck, m, t):
    foot = f"Module {m['id']}  ·  {m['title']}"
    s, cv, top = frame(deck, t["title"], t["tag"], foot, t["tag"])
    V.TOP = top
    V.layout_topic(cv, t, D, t["tag"])
    if t.get("notes"):
        notes(s, t["notes"])


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
    notes(s, f"Now switch to Excel: open {WORKBOOK}, sheet '{m['sheet']}'. Work through the module dataset and the "
             f"step-by-step solutions, then set the exercises; reveal the answers after students have tried.")


def summary_slide(deck):
    s, cv, top = frame(deck, "Week 1 Summary", "WEEK 1  ·  SUMMARY", "Week 1 summary", "summary")
    rows = [["Module", "What you can now do"]] + [[f"{m['id']}  {m['title']}", SUMMARY[m["id"]]] for m in MODULES]
    cv.table(V.X0, top, V.W, rows, 13.5, widths=[3.9, 8.0])
    notes(s, "Recap the week in two minutes, then point students to the Week 1 assignment.")


def glossary_slides(deck):
    per = 15
    chunks = [GLOSSARY[i:i + per] for i in range(0, len(GLOSSARY), per)]
    for k, ch in enumerate(chunks, 1):
        s, cv, top = frame(deck, "Week 1 Glossary" + (f" ({k} of {len(chunks)})" if len(chunks) > 1 else ""),
                           "WEEK 1  ·  GLOSSARY", "Week 1 glossary", f"glossary {k}")
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
def main():
    preview = "--preview" in sys.argv
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
    deck.finish(OUT)
    print(f"{len(deck.prs.slides)} slides, {n_topics} topics -> {OUT.relative_to(ROOT)}")
    if preview:
        PREVIEW.mkdir(parents=True, exist_ok=True)
        for f in PREVIEW.glob("*.png"):
            f.unlink()
        for i, (cv, _) in enumerate(deck.previews, 1):
            K.render(cv.ops, PREVIEW / f"{i:03d}.png", dpi=80)
    if K.WARN:
        print("\nWARNINGS:")
        print("\n".join(K.WARN))


if __name__ == "__main__":
    main()
