"""Week 2 assignment (Edition 3.1): Mercyland 2023 admissions export. Week 1-2 skills only.

    Task 1  import and profile the raw CSV with Power Query (2.1 - 2.4)
    Task 2  clean: TRIM / PROPER, a mapping table, duplicates, text numbers (2.5 - 2.7, 2.9)
    Task 3  functions on the cleaned table: SUMIFS / COUNTIFS / AVERAGEIFS, IF, XLOOKUP, TEXT (2.8A - 2.8B)
    Task 4  PivotTables (2.10 - 2.11)
Task 1 and Task 2 use the raw file in Assignment_Files; Tasks 3 and 4 use tAdm (the cleaned result) so a mistake in
cleaning does not spoil the later tasks.
"""
from __future__ import annotations

import pandas as pd

from asg_lc_clean import clean
from common import STAGE
from wk_common import L

CSV = "Mercyland_EMR_Export_2023.csv"


def build(b):
    raw = L("dirty_lc").copy()
    src = STAGE / "learn_files" / CSV
    src.parent.mkdir(parents=True, exist_ok=True)
    raw.to_csv(src, index=False)
    cl, _steps, _neg = clean(raw)
    adm = cl[["Patient_Name", "Age", "Gender", "Diagnosis_Group", "Admit_Date", "Discharge_Date", "Ward", "HMO",
              "Bill_Amount_NGN", "Admission_Route", "LOS_Days"]].copy()
    adm["Age"] = pd.to_numeric(adm.Age, errors="coerce")
    adm.insert(0, "Episode_No", range(1, len(adm) + 1))
    wnorm = raw.Ward.astype(str).str.strip().str.replace(r"\s+", " ", regex=True).str.title()
    wmap = pd.DataFrame({"Raw_Ward": sorted(wnorm.unique())})
    wmap["Ward"] = [("Mahogany Ward" if w.startswith("Mahog") else (w if w.endswith("Ward") else w + " Ward"))
                    for w in wmap.Raw_Ward]

    t1 = b.module("T1", "T1 Import & Profile", "Import and profile the raw export",
                  "Modules 2.1 - 2.4: file formats, Get Data › From Text/CSV, the Power Query editor, data types and "
                  "column quality / distribution / profile.")
    t2 = b.module("T2", "T2 Clean", "Clean the export", "Modules 2.5 - 2.7 and 2.9: TRIM, PROPER, Replace Values, "
                  "mapping tables, Remove Duplicates, changing types and reading the M code.")
    t3 = b.module("T3", "T3 Functions", "Answer questions with functions", "Modules 2.8A - 2.8B: SUMIFS, COUNTIFS, "
                  "AVERAGEIFS, ROUND, IF, IFERROR, TEXT, XLOOKUP.")
    t4 = b.module("T4", "T4 PivotTables", "Summarise with PivotTables", "Modules 2.10 - 2.11: PivotTables, grouping "
                  "dates, Show Values As, sorting, slicers.")
    f = t1.link_files([src], folder="Assignment_Files")
    t2.link_files([src], folder="Assignment_Files")
    r_map = b.data("WardMap", wmap, "One ward spelling (after TRIM and PROPER) and the standard ward name", used_in="T2")
    r_adm = b.data("Adm", adm, "One Mercyland admission in 2023, already cleaned (the result of Task 2)",
                   fmts={"Admit_Date": "date", "Discharge_Date": "date", "Bill_Amount_NGN": "naira2"},
                   widths={"HMO": 30}, used_in="T3, T4")
    t2.set_refs([r_map])
    t3.set_refs([r_adm])
    t4.set_refs([r_adm])

    # ================================================================== T1
    m = t1
    m.info("Scenario", "Mercyland Specialist Hospital has exported its 2023 admissions from the patient system (EMR). "
                       "The file is messy. Import it, measure how messy it is, clean it, then answer the Medical "
                       "Director's questions.")
    m.part("Part A · Ideas")
    m.mc("What kind of file is a .csv?", ["A plain-text file with values separated by commas", "An Excel workbook",
                                          "A PDF", "A database"], "A", "Comma-separated values.")
    m.mc("Where do you import a CSV with Power Query?", ["Data › Get Data › From File › From Text/CSV", "Insert › Table",
                                                         "File › Print", "Home › Paste"], "A", "Then click Transform Data.")
    m.yn("Does Power Query change the original CSV file?", "No", "It records steps and loads a cleaned copy.")
    m.part("Part B · Import and profile")
    m.num(f"Import {f} › {CSV} into Power Query. How many rows does it have? (status bar of the editor, bottom left, "
          "after View › Column profiling based on entire data set)", None, len(raw), f"{len(raw):,} rows.", do=True)
    m.num("How many columns does it have?", None, len(raw.columns), ", ".join(raw.columns), do=True)
    m.num("Tick View › Column distribution. How many DISTINCT values does the Ward column have?", None, int(raw.Ward.nunique()),
          "Spelling, case and spaces all make 'different' values.", do=True)
    m.num("How many rows have a blank Discharge_Date? (View › Column quality shows Empty, or filter the column)", None,
          int(raw.Discharge_Date.isna().sum()), "Patients still in hospital or missing data.", do=True)
    m.mc("Bill_Amount_NGN contains values like '754,146.44' (in quotes). What is the problem?",
         ["They are stored as text, so they will not add up", "They are too large", "They are dates", "Nothing"], "A",
         "Text numbers must be converted before summing.")

    # ================================================================== T2
    m = t2
    m.part("Part A · Text clean-up")
    m.mc("Which function removes extra spaces (leading, trailing and doubles)?", ["TRIM", "PROPER", "UPPER", "LEN"], "A",
         "TRIM. In Power Query: Transform › Format › Trim.")
    m.short("Patient names look like 'yetuNde p. garBa'. Type the result of =PROPER(\"yetuNde p. garBa\").",
            "Yetunde P. Garba", "PROPER capitalises the first letter of each word.")
    m.num("After =PROPER(TRIM(Ward)) how many different ward spellings are left? (they are listed in tWardMap)", None,
          len(wmap), "Case and spaces are fixed, but short names and a typo remain.")
    m.short("In tWardMap, which standard name does the spelling 'Mahogony Ward' map to?", "Mahogany Ward",
            "A typo: map it with the table, then use XLOOKUP or a Power Query merge.")
    m.num("After mapping, how many wards does the hospital have?", None, int(wmap.Ward.nunique()), "Four wards.")
    m.part("Part B · Duplicates, types and M")
    m.num("In Power Query on the raw data, Home › Remove Rows › Remove Duplicates (all columns). How many rows remain?", None,
          len(raw) - int(raw.duplicated().sum()), f"{int(raw.duplicated().sum())} exact duplicates removed.", do=True)
    m.mc("Why should you standardise text BEFORE removing duplicates?", ["It is faster",
                                                                        "'IROKO WARD' and 'Iroko Ward' only match after cleaning",
                                                                        "Power Query needs it", "It is not needed"], "B",
         "Clean first, then de-duplicate, or near-duplicates survive.")
    m.mc("Which M function does the Trim step use?", ["Text.Trim", "Table.Sort", "List.Sum", "Date.Year"], "A",
         "Home › Advanced Editor shows Table.TransformColumns(..., Text.Trim).")
    m.yn("After cleaning, should you check that the row count and the total bill still make sense?", "Yes",
         "Control totals.")

    # ================================================================== T3
    m = t3
    A = "tAdm"
    m.part("Answer each question with one formula on tAdm")
    m.num("How many admissions are in the cleaned table?", f"=ROWS({A})", len(adm), "ROWS of the table.")
    m.num("How many admissions came through Emergency?", f'=COUNTIFS({A}[Admission_Route],"Emergency")',
          int((adm.Admission_Route == "Emergency").sum()), "COUNTIFS.")
    m.num("Total bills for Iroko Ward? (nearest naira)", f'=ROUND(SUMIFS({A}[Bill_Amount_NGN],{A}[Ward],"Iroko Ward"),0)',
          round(float(adm.loc[adm.Ward == "Iroko Ward", "Bill_Amount_NGN"].sum())), "SUMIFS by Ward.", fmt="naira")
    m.num("Average length of stay (LOS_Days) in Mahogany Ward? (1 decimal place; AVERAGEIFS skips blanks)",
          f'=AVERAGEIFS({A}[LOS_Days],{A}[Ward],"Mahogany Ward")',
          float(adm.loc[adm.Ward == "Mahogany Ward", "LOS_Days"].mean()), "AVERAGEIFS by Ward.", fmt="dec1")
    m.num("How many Female patients were admitted to Obeche Ward?",
          f'=COUNTIFS({A}[Gender],"Female",{A}[Ward],"Obeche Ward")',
          int(((adm.Gender == "Female") & (adm.Ward == "Obeche Ward")).sum()), "COUNTIFS with two conditions.")
    m.num("Long stay = LOS_Days of 14 or more. Add a column =IF([@LOS_Days]>=14,\"Long\",\"Normal\"). How many are Long?",
          f'=COUNTIFS({A}[LOS_Days],">=14")', int((adm.LOS_Days >= 14).sum()), "Count the Long rows (or COUNTIFS >= 14).")
    m.num("How many LOS_Days values are negative (discharge before admission: a data error to report)?",
          f'=COUNTIFS({A}[LOS_Days],"<0")', int((adm.LOS_Days < 0).sum()), "COUNTIFS < 0.")
    m.short("Use XLOOKUP to find the Ward of Episode_No 100.", str(adm.loc[adm.Episode_No == 100, "Ward"].iloc[0]),
            "=XLOOKUP(100,tAdm[Episode_No],tAdm[Ward])")
    m.short("What does =TEXT(DATE(2023,3,15),\"mmm\") return?", "Mar", "TEXT with \"mmm\" gives the short month name.")

    # ================================================================== T4
    m = t4
    adm["Mon"] = adm.Admit_Date.dt.month
    bym = adm.groupby("Mon").size()
    byw = adm.groupby("Ward").size().sort_values(ascending=False)
    import calendar
    m.part("Build a PivotTable on tAdm (Insert › PivotTable)")
    m.short("Rows = Ward, Values = Count of Episode_No. Which ward has the most admissions?", byw.index[0],
            "Sort the pivot largest to smallest.")
    m.num("What is that ward's share of all admissions? (Show Values As › % of Grand Total)", None,
          float(byw.iloc[0] / len(adm)), f"{byw.iloc[0]:,} ÷ {len(adm):,}.", fmt="pct")
    m.short("Rows = Admit_Date grouped by Months. Which month has the most admissions? (type the 3-letter month, e.g. Jan)",
            calendar.month_abbr[int(bym.idxmax())], "Group the dates by Months, then sort.")
    m.num("Rows = Admission_Route, Values = Sum of Bill_Amount_NGN. What is the Elective total? (nearest naira)", None,
          round(float(adm.loc[adm.Admission_Route == "Elective", "Bill_Amount_NGN"].sum())), "Read the Elective row.",
          fmt="naira", do=True)
    m.mc("Which tool lets the Medical Director click a Ward to filter the pivot?", ["Slicer", "Flash Fill", "Goal Seek",
                                                                                     "Freeze Panes"], "A",
         "PivotTable Analyze › Insert Slicer.")
    m.yn("After adding new rows to tAdm, do you need to Refresh the pivot?", "Yes", "Right-click › Refresh.")
