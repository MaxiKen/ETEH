# Superseded builders (Edition 3.0)

These scripts built the Edition 3.0 Section Examples (`02_sec_*`) and assignments (`03_asg_*`) in the old flat
folder layout. Edition 3.1 replaces them:

- `_build/10_learner_workbooks.py` + `_build/learn/w1.py … w5.py`  → `01_Section_Examples/Week_N/`
- `_build/11_learner_assignments.py` + `_build/learn/a1.py … a4.py` → `02_Assignments/Week_N/` and
  `03_Instructor_Projects/Week_N/`

They are kept for reference only and are not part of the build. Do not run them: they write to the old paths.
Note: v3/wb2_*.py, wb1*, wb3*, wb4* are edition 3.0 workbook modules used only by the builders in this folder.
