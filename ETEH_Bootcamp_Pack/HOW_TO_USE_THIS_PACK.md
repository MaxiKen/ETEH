# How to use the ETEH Bootcamp Pack

## Before you start (once)
1. Unzip the **whole** pack to one place, for example `C:\ETEH\`. Keep the folders exactly as they are. Every workbook
   looks for its files in the folder **next to it**, so never move a workbook out of its `Week_N` folder.
2. You need Excel 365 (or Excel 2021) for all weeks. For Week 3 also install DB Browser for SQLite (free), and the SQLite
   ODBC driver if you want to try module 3.3.
3. The pack has five slide decks (one per week) and four folders:

| Item | What it is | Who uses it | When |
|---|---|---|---|
| `ETEH_Vol1` … `ETEH_Vol5` slides | The lessons, module by module | Everyone | First, for every module |
| `01_Section_Examples` | Practice workbook: one sheet per module | Students | **After each module** |
| `02_Assignments` | Graded assignment | Students | **At the end of each week** |
| `03_Instructor_Projects` | Answer keys, capstone solutions, project keys | **Tutors only** | Marking and class demos |
| `04_Student_Projects` | The four final projects | Students | **Week 5** |

---

## Every module (Weeks 1 to 4): slides, then the 01 folder
1. **Learn:** go through the module in that week's slide deck (for example Week 2 › Module 2.3).
2. **Open the practice workbook:** `01_Section_Examples › Week_N › ETEH_WeekN_Section_Examples.xlsx`.
   - The first time each week, read the **Start_Here** tab. It lists every sheet in teaching order, the data each sheet
     uses and its question types.
3. **Go to the sheet with the module's number** (for example `2.3 Connect & Combine`). Do only that sheet. Each sheet asks
   about that module and the ones before it, nothing later.
4. **Read the gold box at the top of the sheet:**
   - **You learned:** a reminder of the module.
   - **Data to use:** exactly where the data is, as **File › Sheet › Table**, for example
     `ETEH_Week2_Section_Examples.xlsx › Data_Stores › tStores`. Click the data tab (`Data_…`, near the end) to see the
     table. "No data" means you answer from the slides.
   - **Files folder** (some modules only): the folder of attached files, for example `Ex2.3_Files`. It sits **next to the
     workbook** inside the same `Week_N` folder. Open or import the named file from there. Questions name each file the
     same way, `Ex2.3_Files › SuperMart_Stores_Sales_2022_2025.csv`.
5. **Answer the questions** in the white cell with the gold border (column E).
   - Multiple choice, Yes/No and True/False: pick from the drop-down list.
   - Short answer: type one word or phrase.
   - Number or "Do it in Excel": do the task (a formula, a Power Query step, a query), then type the result.
6. **Check yourself:** column F shows ✓ Correct or ✗ Try again straight away. Try again before looking anything up.
7. **Only when you have tried:** open the **Solutions** tab (the last tab). It gives every answer, a one-line explanation
   and, for numbers, the formula that produces it.
8. Move on to the next module's slides, then the next sheet. Repeat.

### Which modules have a files folder
| Week | Attached-file folders (next to the workbook) | All other sheets use |
|---|---|---|
| 1 | none | the `Data_…` tabs inside the workbook, or the slides |
| 2 | `Ex2.1_Files`, `Ex2.2_Files`, `Ex2.3_Files`, `Ex2.4_Files`, `Ex2.6_Files` | the `Data_…` tabs, or the slides |
| 3 | `Ex3.1_Files` (schema), `Ex3.2_Files` (the two `.db` databases and the SQL examples), `Ex3.4_Files`, `Ex3.10_Files` | the `Data_…` tabs, or the slides |
| 4 | `Ex4.14_Files` | the `Data_…` tabs, or the slides |

---

## End of each week: the 02 folder (assignment)
1. Finish all of that week's module sheets first.
2. Open `02_Assignments › Week_N › ETEH_WeekN_Assignment.xlsx` and **save a copy with your name** (for example
   `ETEH_Week2_Assignment_AdaObi.xlsx`) in the same folder.
3. Read **Start_Here**, then do the tasks in order (`T1`, `T2`, …). Each task's gold box tells you the data to use, in the
   same **File › Sheet › Table** form.
   - **Week 2:** task T1 imports `Assignment_Files › Mercyland_EMR_Export_2023.csv`.
   - **Week 3:** you query `Assignment_Files › mercyland_hospital.db` in DB Browser for SQLite.
   - **Weeks 1 and 4:** all the data is inside the workbook.
   - Later tasks use a clean copy of the data that is already inside the workbook, so a slip in an early task does not
     spoil the rest.
4. Column F shows the **marks** for each question (2 for a number or Excel task, 1 for anything else). The assignment has
   no Solutions tab. Your tutor marks it with the answer key.
5. Submit the saved file to your tutor.

---

## Week 5: capstone and projects (01, then 04)
1. **Learn and practise:** work through the Capstone slides (`ETEH_Vol5`) together with
   `01_Section_Examples › Week_5_Capstone › ETEH_Capstone_Section_Examples.xlsx`. It has one sheet per module (P.1 to P.9)
   and works exactly like Weeks 1 to 4, including the Solutions tab.
2. **Do a project:** open `04_Student_Projects › Week_5_Capstone`. Your tutor tells you which of the four projects to do.
   Each project folder holds everything for that project:

| Open in this order | What it is |
|---|---|
| 1. `ProjectN_Brief.pdf` | The client, the questions, the rules and the marking scheme. Read it fully first. |
| 2. `Data` folder | The raw data you must clean and analyse. Project 1 also has the `supermart_online.db` database. |
| 3. `ProjectN_…_Workbook.xlsx` | Where you answer: `00_Cover` (your name), then milestones `M1` to `M4` in order, and `99_Marks` (your score summary). |
| 4. `Project1_SQL_Template.sql` (Project 1 only) | Write your SQL here, one query under each question. |

3. Do the milestones in order: **M1** cleans or audits the data, **M2** and **M3** analyse it, **M4** produces the
   dashboard, the memo and the slides. Submit the workbook, plus your dashboard file, memo, deck and (Project 1) the SQL
   file, then prepare the ten-minute defence (module P.7).

---

## For tutors: the 03 folder
| Folder | Use |
|---|---|
| `03_Instructor_Projects › Week_1` … `Week_4` | `ETEH_WeekN_Assignment_ANSWER_KEY.xlsx`: the same sheets as the student assignment with every answer filled in, plus a **Solutions** tab. Use it to mark. |
| `Week_5_Capstone › Capstone1_SuperMart_Stores`, `Capstone2_Kingsway_University`, `Capstone3_Mercyland_Hospital` | The three worked capstones you demonstrate in class (modules P.3, P.4 and P.5): instructor brief (PDF), `SOLUTION.xlsx`, and the `Data` folder (plus the extraction SQL for Capstone 3). |
| `Week_5_Capstone › Student_Project_Keys` | The marking keys for the four student projects (plus the SQL key for Project 1). |

Do not give students the 03 folder. If you share the pack with students, remove it first.

---

## One-week routine at a glance
```
For each module:   slides  ->  01_Section_Examples/Week_N  (that module's sheet)  ->  check column F  ->  Solutions only after trying
End of week:       02_Assignments/Week_N  ->  save a copy with your name  ->  submit
Tutor:             marks with 03_Instructor_Projects/Week_N answer key
Week 5:            Vol5 slides + 01/Week_5_Capstone  ->  04_Student_Projects/Week_5_Capstone/ProjectN (Brief -> Data -> Workbook)
```
