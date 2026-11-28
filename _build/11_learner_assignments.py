"""Edition 3.1 assignments: one per week, built twice from the same questions.

    python3 _build/11_learner_assignments.py [1 2 3 4]

Student file:  02_Assignments/Week_N/ETEH_WeekN_Assignment.xlsx           (Marks column, no Solutions sheet)
               02_Assignments/Week_N/Assignment_Files/                   (only when a task links external files)
Answer key:    03_Instructor_Projects/Week_N/ETEH_WeekN_Assignment_ANSWER_KEY.xlsx   (✓ checks + Solutions sheet)

Every task sheet stays inside the week's teaching: Week 1 tasks use Week 1 skills only, and so on.
"""
from __future__ import annotations

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path[:0] = [HERE, os.path.join(HERE, "v3"), os.path.join(HERE, "learn")]

import lkit  # noqa: E402
import paths  # noqa: E402

WEEKS = {
    1: ("Week 1 Assignment · SuperMart Q1 2025 Store Review", "a1"),
    2: ("Week 2 Assignment · Mercyland Admissions: Import, Clean and Summarise", "a2"),
    3: ("Week 3 Assignment · Mercyland Hospital: SQL, Data Model and PivotTables", "a3"),
    4: ("Week 4 Assignment · Statistics, Forecasting and Reporting", "a4"),
}


def student_intro(n):
    return [f"Do this assignment after you finish all the Week {n} modules and the Week {n} Section Examples. "
            f"Every task uses only what you learned in Week {n} (and earlier weeks).",
            "The gold box at the top of every task sheet tells you exactly which data to use: the File, the Sheet and "
            "the Table (or the folder that holds the file).",
            "Answer in the white cell with the gold border (column E). Multiple choice, Yes/No and True/False have a "
            "drop-down list. Number answers: type the number only (no ₦ sign); percentages can be typed as 25% or 0.25.",
            "Column F shows the marks for each question: 2 for a number or Excel task, 1 for every other question.",
            f"Save the file as ETEH_Week{n}_Assignment_YourName.xlsx and hand it in. Your tutor marks it with the answer key."]


def tutor_intro(n):
    files = (f"The students' outside files are in 02_Assignments › Week_{n} › Assignment_Files; the questions on this copy "
             "name them the same way." if n in (2, 3) else "All the data is inside this workbook (the Data_ tabs), exactly as in the students' copy.")
    return [f"TUTOR COPY - do not share. Same tasks, questions, order and data as the students' ETEH_Week{n}_Assignment.xlsx "
            "(02_Assignments › Week_" + str(n) + ").",
            "Solutions (last tab) lists every answer with a one-line 'How to get it'. Number answers are live formulas on this "
            "workbook's own tables: click one to see the exact calculation.",
            "To mark: open the student's file beside this one. Copy their column E answers (Paste Values) into column E of the "
            "same task sheet here; column F then shows ✓ Correct or ✗ Try again for every question.",
            "Award 2 marks for each correct number or 'Do it in Excel' question and 1 for every other question (the student's "
            "own column F shows each question's marks). Number checks already allow small rounding differences.",
            files]


def build_week(n):
    title, modname = WEEKS[n]
    mod = __import__(modname)
    out = []
    for key in (False, True):
        folder = paths.ipr(n) if key else paths.asg(n)
        folder.mkdir(parents=True, exist_ok=True)
        fname = f"ETEH_Week{n}_Assignment" + ("_ANSWER_KEY" if key else "") + ".xlsx"
        b = lkit.LBook(folder / fname, fname, n, title, title_suffix="Assignment" + (" · Answer key" if key else ""),
                       assignment=key, with_solutions=key, intro=tutor_intro(n) if key else student_intro(n),
                       files_home=paths.rel(paths.asg(n)) if key else None)
        mod.build(b)
        b.close()
        q = sum(m.nq for m in b.modules)
        marks = sum(getattr(m, "marks", 0) for m in b.modules)
        print(f"wrote {paths.rel(folder / fname)}: {len(b.modules)} task sheets, {q} questions, {marks} marks")
        out.append(b)
    return out


if __name__ == "__main__":
    for w in (sys.argv[1:] or ["1", "2", "3", "4"]):
        build_week(int(w))
