"""Edition 3.1 folder layout: every top folder holds one sub-folder per week, nothing else.

    01_Section_Examples/Week_N/        the week's workbook (+ one Ex<module>_Files folder per exercise that links files)
    02_Assignments/Week_N/             the week's assignment (+ Assignment_Files folder when it links files)
    03_Instructor_Projects/Week_N/     answer keys for that week (assignment key; capstones and project keys in Week 5)
    04_Student_Projects/Week_5_Capstone/ProjectN_.../   the four portfolio projects (done after Week 4)
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from common import PACK  # noqa: E402

WEEK_DIR = {1: "Week_1", 2: "Week_2", 3: "Week_3", 4: "Week_4", 5: "Week_5_Capstone"}
SEC = PACK / "01_Section_Examples"
ASG = PACK / "02_Assignments"
IPR = PACK / "03_Instructor_Projects"
SPR = PACK / "04_Student_Projects"


def sec(n):
    return SEC / WEEK_DIR[n]


def asg(n):
    return ASG / WEEK_DIR[n]


def ipr(n):
    return IPR / WEEK_DIR[n]


def spr(n=5):
    return SPR / WEEK_DIR[n]


def rel(p: Path) -> str:
    """Path as shown to learners, relative to the pack root, with › separators."""
    return " › ".join(Path(p).resolve().relative_to(PACK.resolve()).parts)
