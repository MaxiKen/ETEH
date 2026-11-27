"""Week 2 assignment - Lifecare 2023 EMR export: profile, clean (Excel + Power Query), analyse with functions and pivots."""
from __future__ import annotations

import shutil

import numpy as np
import pandas as pd

from asg_common import Asg
from asg_lc_clean import GENDER_MAP, TYPE_MAP, WARD_MAP, clean, ward_key
from common import ASG
from wk_common import DATASETS_TEXT, L

DATA_DIR = ASG / "Week2_Assignment_Data"
DATA_DIR.mkdir(parents=True, exist_ok=True)
CSV_NAME = "Lifecare_EMR_Export_2023.csv"

raw = L("dirty_lc").copy()
raw.to_csv(DATA_DIR / CSV_NAME, index=False)
cl, steps, neg = clean(raw)
raw_x = raw.copy()
raw_x["Admission_Date"] = pd.to_datetime(raw_x.Admission_Date)
raw_x["Discharge_Date"] = pd.to_datetime(raw_x.Discharge_Date)

n_dist_raw = len(raw) - int(raw.duplicated().sum())                 # PQ Table.Distinct is case-sensitive
n_blank_dis = int(raw.Discharge_Date.isna().sum())
n_text_bill = int(raw.Billing_Amount.map(lambda x: isinstance(x, str)).sum())
n_neg_num = int(raw.Billing_Amount.map(lambda x: (not isinstance(x, str)) and x < 0).sum())
n_age_bad = int(((raw.Age < 1) | (raw.Age > 120)).sum())
n_gender_bad = int((~raw.Gender.isin(["Male", "Female"])).sum())
n_ward_unique_excel = raw.Ward.str.lower().nunique()               # UNIQUE / Remove Duplicates ignore case, not spaces
assert len(neg) == 18 and n_neg_num < 18, (n_neg_num, len(neg))

# near-duplicates (identical except one blank field)
cols_cmp = [c for c in cl.columns if c not in ("Discharge_Date", "HMO", "Length_of_Stay", "LOS_Valid")]
near = int(cl.duplicated(subset=cols_cmp).sum())
TRUE_N = len(cl) - near

valid = cl[cl.LOS_Valid == "Yes"]
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# 30-row practice sample with every defect type represented
rng = np.random.default_rng(2026)
pick = set()
for mask in [raw.Billing_Amount.map(lambda x: isinstance(x, str)), ~raw.Ward.isin(["Ruby Ward", "Sapphire Ward", "Topaz Ward", "Emerald Ward"]),
             ~raw.Gender.isin(["Male", "Female"]), raw.Discharge_Date.isna(), ~raw.Admission_Type.isin(["Urgent", "Emergency", "Elective"])]:
    idx = raw.index[mask].tolist()
    pick.update(rng.choice(idx, 4, replace=False).tolist())
neg_los = raw.index[pd.to_datetime(raw.Discharge_Date) < pd.to_datetime(raw.Admission_Date)].tolist()
pick.update(neg_los[:2])
others = [i for i in raw.index if i not in pick]
pick.update(rng.choice(others, 30 - len(pick), replace=False).tolist())
smp = raw_x.loc[sorted(pick), ["Patient_Name", "Gender", "Ward", "Admission_Type", "Admission_Date", "Discharge_Date", "Billing_Amount"]].reset_index(drop=True)
smp.insert(0, "Row", range(1, len(smp) + 1))

s_name = smp.Patient_Name.str.strip().str.title().tolist()
s_ward = [WARD_MAP[ward_key(x)] for x in smp.Ward]
s_gen = [GENDER_MAP[str(x).strip().lower()] for x in smp.Gender]
s_type = [TYPE_MAP[str(x).strip().lower()] for x in smp.Admission_Type]
s_bill = [float(str(x).replace(",", "")) if isinstance(x, str) else float(x) for x in smp.Billing_Amount]
s_los = []
for a, d in zip(smp.Admission_Date, smp.Discharge_Date):
    s_los.append("" if pd.isna(d) or d < a else (d - a).days)

wardmap = pd.DataFrame(sorted(WARD_MAP.items()), columns=["Raw_Key", "Clean_Ward"])

M_SCRIPT = r'''let
    // 1  Connect (Data > Get Data > From Text/CSV) - UTF-8, comma delimited
    Source      = Csv.Document(File.Contents("C:\ETEH\Week2_Assignment_Data\Lifecare_EMR_Export_2023.csv"),
                    [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Promoted    = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    // 2  Standardise text BEFORE removing duplicates (PROPER names, collapse spaces, lower-case keys)
    TextFix     = Table.TransformColumns(Promoted, {
                    {"Patient_Name",   each Text.Proper(Text.Trim(_)), type text},
                    {"Ward",           each Text.Lower(Text.Combine(List.Select(Text.Split(Text.Trim(_), " "), each _ <> ""), " ")), type text},
                    {"Gender",         each Text.Lower(Text.Trim(_)), type text},
                    {"Admission_Type", each Text.Lower(Text.Trim(_)), type text}}),
    // 3  Ward labels -> mapping table tWardMap (Merge Queries, Left Outer)
    WardMap     = Excel.CurrentWorkbook(){[Name = "tWardMap"]}[Content],
    Merged      = Table.NestedJoin(TextFix, {"Ward"}, WardMap, {"Raw_Key"}, "Map", JoinKind.LeftOuter),
    Expanded    = Table.ExpandTableColumn(Merged, "Map", {"Clean_Ward"}),
    WardDone    = Table.RenameColumns(Table.RemoveColumns(Expanded, {"Ward"}), {{"Clean_Ward", "Ward"}}),
    // 4  Coded values -> full labels
    Gender      = Table.TransformColumns(WardDone, {{"Gender", each if Text.StartsWith(_, "m") then "Male" else "Female", type text}}),
    AdmType     = Table.TransformColumns(Gender, {{"Admission_Type", each if Text.StartsWith(_, "emerg") then "Emergency" else Text.Proper(_), type text}}),
    // 5  Text numbers, blanks and types
    Bill        = Table.TransformColumns(AdmType, {{"Billing_Amount", each Number.From(Text.Replace(_, ",", ""), "en-US"), type number}}),
    HMO         = Table.ReplaceValue(Bill, "", "Self-pay / Unknown", Replacer.ReplaceValue, {"HMO"}),
    Blanks      = Table.ReplaceValue(HMO, "", null, Replacer.ReplaceValue, {"Discharge_Date"}),
    Typed       = Table.TransformColumnTypes(Blanks, {{"Age", Int64.Type}, {"Room_Number", Int64.Type},
                    {"Admission_Date", type date}, {"Discharge_Date", type date}}),
    // 6  Remove exact duplicates (case-sensitive) and finance reversals
    Dedup       = Table.Distinct(Typed),
    NoReversals = Table.SelectRows(Dedup, each [Billing_Amount] >= 0),
    // 7  Derived columns and quality flags (flag, do not delete)
    LOS         = Table.AddColumn(NoReversals, "Length_of_Stay", each if [Discharge_Date] = null then null
                    else Duration.Days([Discharge_Date] - [Admission_Date]), Int64.Type),
    LOSValid    = Table.AddColumn(LOS, "LOS_Valid", each if [Length_of_Stay] <> null and [Length_of_Stay] >= 0 then "Yes" else "No", type text),
    AgeValid    = Table.AddColumn(LOSValid, "Age_Valid", each if [Age] >= 1 and [Age] <= 120 then "Yes" else "No", type text),
    Quarter     = Table.AddColumn(AgeValid, "Quarter", each "Q" & Text.From(Date.QuarterOfYear([Admission_Date])), type text),
    Month       = Table.AddColumn(Quarter, "Month", each Date.ToText([Admission_Date], "MMM", "en-GB"), type text)
in
    Month'''

CLEAN_COLS = ["Patient_Name", "Age", "Gender", "Blood_Type", "Medical_Condition", "Admission_Date", "Discharge_Date", "Doctor_Name", "Ward", "Room_Number",
              "HMO", "Billing_Amount", "Admission_Type", "Medication", "Test_Result", "Length_of_Stay", "LOS_Valid", "Age_Valid", "Quarter", "Month"]


def build(key: bool):
    A = Asg(2, "ETEH Week 2 Assignment - Lifecare 2023 Admissions: from EMR export to pivot report", key,
            "Week 2 · Power Query, cleaning, functions and PivotTables · Lifecare General Hospital 2023")
    b = A.b

    # ---------------------------------------------------------------- T1 import & connect
    s = A.task("T1 Import & Connect", "Task 1 - Getting the EMR export into Excel", "2.1-2.4",
               f"Lifecare's old EMR system exports admissions as {CSV_NAME} (folder Week2_Assignment_Data). The same rows are pasted, as received, "
               "on sheet D_Raw (table tRaw) so you can profile them with formulas. Import the CSV with Data > Get Data > From Text/CSV and answer.")
    A.q(s, "1.1", "Which delimiter separates the fields in the CSV?", "Comma", marks=1, choices=["Comma", "Semicolon", "Pipe", "Tab"], header=True)
    A.q(s, "1.2", "How many data rows (excluding the header) does the export contain? (formula on tRaw)", "=ROWS(tRaw)", len(raw), "int", 2)
    A.q(s, "1.3", "Which column does Power Query type as TEXT (ABC) instead of a number because some values contain thousands separators?", "Billing_Amount",
        marks=2, choices=["Age", "Room_Number", "Billing_Amount", "Admission_Date"], guide="441 values look like '2,507,101.35'.")
    A.q(s, "1.4", "Which file encoding should the Source step use so ₦ and accented names load correctly?", "UTF-8 (65001)", marks=2,
        choices=["UTF-8 (65001)", "Windows-1252", "ASCII", "UTF-16"])
    A.q(s, "1.5", "Before cleaning, a Merge on Ward to the 4-row ward master returns null for many rows. Why?", "Ward labels are inconsistent", marks=2,
        choices=["Ward labels are inconsistent", "The master table is too small", "Merge needs a numeric key", "The CSV is not sorted"],
        guide="661 rows use variants (EMERALD WARD, 'Topaz ward', 'Saphire Ward', trailing spaces).")
    A.q(s, "1.6", "The EMR team will send one CSV per month from 2025. Which Power Query connector combines every file in a folder?", "From Folder (Combine)",
        marks=2, choices=["From Folder (Combine)", "From Table/Range", "From Web", "From PDF"])
    A.q(s, "1.7", "HMO tariff schedules arrive as tables inside PDF letters. Which connector extracts them?", "From PDF", marks=2,
        choices=["From PDF", "From Web", "From Text/CSV", "From XML"])
    A.long(s, "1.8", "Name TWO risks of copy-pasting the CSV into Excel instead of connecting with Power Query.",
           "• No refresh: next month's file means repeating every manual step (and repeating mistakes).\n"
           "• Excel auto-converts on paste: '1,234.56' stays text, some IDs/dates are re-interpreted, leading zeros lost.\n"
           "• No audit trail of the cleaning steps (Applied Steps document every change).", marks=2, guide="1 mark per valid risk.", rows=3)
    A.end_task(s)

    # ---------------------------------------------------------------- T2 profiling + Excel cleaning
    s = A.task("T2 Profile & Clean in Excel", "Task 2 - Profile the defects, then clean a sample with worksheet functions", "2.5",
               "Part A: measure every defect with a formula on tRaw - these counts go into the data-quality log sent to the EMR team. "
               "Part B: clean the 30-row sample below using worksheet functions only (no Power Query).")
    s.band("2.5 A", "Data-quality profile (formulas on tRaw)")
    A.q(s, "2.1", "How many rows have a blank Discharge_Date?", "=COUNTBLANK(tRaw[Discharge_Date])", n_blank_dis, "int", 2, "", header=True)
    A.q(s, "2.2", "How many Billing_Amount values are stored as TEXT?", "=SUMPRODUCT(--ISTEXT(tRaw[Billing_Amount]))", n_text_bill, "int", 2,
        "COUNTA - COUNT on the column is also correct.")
    A.q(s, "2.3", "How many bills are negative (finance reversals keyed as admissions)?", '=COUNTIF(tRaw[Billing_Amount],"<0")', n_neg_num, "int", 2,
        f"COUNTIF only sees numbers: {18 - n_neg_num} more negative bills are stored as text and only surface after conversion (18 removed in Task 3).")
    A.q(s, "2.4", "How many ages are impossible (below 1 or above 120)?", '=COUNTIF(tRaw[Age],"<1")+COUNTIF(tRaw[Age],">120")', n_age_bad, "int", 2)
    A.q(s, "2.5", "How many Gender values are NOT exactly 'Male' or 'Female' (case matters)?",
        '=ROWS(tRaw)-SUMPRODUCT(--EXACT(tRaw[Gender],"Male"))-SUMPRODUCT(--EXACT(tRaw[Gender],"Female"))', n_gender_bad, "int", 2,
        "COUNTIF is case-insensitive and would wrongly count 'male' as valid - EXACT is required.")
    A.q(s, "2.6", "How many distinct Ward labels does =ROWS(UNIQUE(tRaw[Ward])) return?", "=ROWS(UNIQUE(tRaw[Ward]))", n_ward_unique_excel, "int", 2,
        "UNIQUE ignores case but not spaces: 16 raw spellings → 12. Either way, 4 real wards.")
    s.band("2.5 B", "Clean the sample (one formula per column, copied down)")
    s.para("Rules: Name → PROPER(TRIM()). Ward → look up LOWER(TRIM()) in tWardMap (sheet D_WardMap). Gender/Type → full labels "
           "(Male/Female; Emergency/Urgent/Elective). Bill → a number even when typed with commas. LOS → Discharge − Admission, blank if discharge "
           "is missing or earlier than admission.", "note")
    top = s.r + 1
    A.fill(s, smp, {
        "Name_Clean": ("=PROPER(TRIM(C{r}))", s_name, "gen"),
        "Ward_Clean": ("=XLOOKUP(LOWER(TRIM(E{r})),tWardMap[Raw_Key],tWardMap[Clean_Ward])", s_ward, "gen"),
        "Gender_Clean": ('=IF(LEFT(LOWER(TRIM(D{r})),1)="m","Male","Female")', s_gen, "gen"),
        "Type_Clean": ('=IF(LEFT(LOWER(TRIM(F{r})),3)="eme","Emergency",PROPER(TRIM(F{r})))', s_type, "gen"),
        "Bill_Clean": ('=IF(ISNUMBER(I{r}),I{r},VALUE(SUBSTITUTE(I{r},",","")))', s_bill, "naira2"),
        "LOS": ('=IF(OR(H{r}="",H{r}<G{r}),"",H{r}-G{r})', s_los, "int"),
    }, fmts={"Admission_Date": "date", "Discharge_Date": "date", "Billing_Amount": "naira2"}, marks=10,
        guide="6 columns: Name 1, Ward lookup 2 (TRIM inside LOWER collapses double spaces), Gender 2, Type 2, Bill 2, LOS 1. Accept SWITCH/mapping alternatives.")
    s.ws.set_column(2, 2, 22)
    s.ws.set_column(9, 13, 15)
    r1, r2 = top + 1, top + len(smp)
    A.q(s, "2.7", "Total of Bill_Clean for the sample (₦, 2 dp).", f"=SUM(N{r1}:N{r2})", sum(s_bill), "naira2", 3,
        "SUM of the raw Billing_Amount column silently skips the text values - compare the two totals.", header=True)
    A.end_task(s)

    # ---------------------------------------------------------------- T3 power query pipeline
    s = A.task("T3 Power Query Pipeline", "Task 3 - Build the repeatable cleaning query", "2.6, 2.7, 2.9",
               f"Build ONE Power Query query on {CSV_NAME} that applies the cleaning rules from Task 2 to all 11,026 rows, removes duplicates and finance "
               "reversals, and adds Length_of_Stay, LOS_Valid, Age_Valid, Quarter and Month. Load it to a new sheet as a Table named tClean "
               "(Tasks 4 and 5 use it). Report the control totals below.")
    s.grid([["1", "Standardise text first: PROPER names; TRIM + lower-case Ward, Gender, Admission_Type"],
            ["2", "Merge Ward with tWardMap (Left Outer) and expand Clean_Ward"],
            ["3", "Gender → Male/Female; Admission_Type → Emergency/Urgent/Elective"],
            ["4", "Billing_Amount: remove ',' then type as Decimal; blank HMO → 'Self-pay / Unknown'; blank Discharge_Date → null"],
            ["5", "Remove Duplicates (all columns) and filter out negative bills"],
            ["6", "Add Length_of_Stay (days), LOS_Valid (Yes if ≥ 0 and not null), Age_Valid (1-120), Quarter (Q1-Q4), Month (Jan-Dec)"]],
           header=["Step", "Rule"], merge_last=True)
    A.q(s, "3.1", "Rows left if you run Remove Duplicates in Power Query on the RAW rows, before any text clean-up.", n_dist_raw, None, "int", 2,
        "Table.Distinct is case-sensitive: only 63 exact copies. (Excel's Remove Duplicates ignores case and gives a different count - another reason to clean first.)", header=True)
    A.q(s, "3.2", "Rows after standardising text and THEN removing duplicates.", steps[1][1], None, "int", 2, "88 duplicates once casing/labels agree.")
    A.q(s, "3.3", "Rows after removing negative-bill reversals (= rows in tClean).", "=ROWS(tClean)", len(cl), "int", 2)
    A.q(s, "3.4", "Total Billing_Amount in tClean (₦, 2 dp).", "=SUM(tClean[Billing_Amount])", cl.Billing_Amount.sum(), "naira2", 2, "Accept ±₦1.")
    A.q(s, "3.5", "How many rows are flagged LOS_Valid = \"No\"?", '=COUNTIF(tClean[LOS_Valid],"No")', int((cl.LOS_Valid == "No").sum()), "int", 2,
        "70 missing discharge + 25 discharge-before-admission.")
    A.q(s, "3.6", "Average length of stay (days, 2 dp) over valid rows only.", '=ROUND(AVERAGEIFS(tClean[Length_of_Stay],tClean[LOS_Valid],"Yes"),2)',
        round(valid.Length_of_Stay.mean(), 2), "dec2", 2)
    A.long(s, "3.7", "Paste the M code of your query (Home > Advanced Editor). The Ward and Billing_Amount steps will be marked.",
           M_SCRIPT, marks=4, guide="2 marks Ward standardisation (trim/lower + merge or mapping), 1 Billing text→number, 1 correct step order (clean before Distinct).", rows=10)
    A.long(s, "3.8", "Excel formulas (Task 2B) or Power Query (Task 3) for this file next month? Justify in two sentences.",
           "Power Query: the steps are recorded once and re-run on every new file with Refresh, with an audit trail of Applied Steps. "
           "Worksheet formulas suit one-off fixes on small samples but must be re-pasted and re-checked each month.", marks=2, rows=3)
    A.q(s, "3.9", "ADVANCED: two pairs in tClean are identical except that one copy has a blank field. What is the true admission count after removing them?",
        TRUE_N, None, "int", 2, "Pairs differ only by a blank Discharge_Date / HMO - Remove Duplicates only catches exact matches. Find with COUNTIFS on name + admission date.", header=True)
    A.end_task(s)

    # ---------------------------------------------------------------- T4 functions
    s = A.task("T4 Functions on tClean", "Task 4 - Analyse the clean table with functions", "2.8",
               "Use worksheet functions on YOUR tClean table (loaded in Task 3). Each answer must be a live formula.")
    nh = cl.HMO.eq("NHIA (National Health Insurance)").mean()
    ix = cl.Billing_Amount.idxmax()
    cancer = valid[valid.Medical_Condition == "Cancer"].Length_of_Stay.median()
    wk = int((cl.Admission_Date.dt.dayofweek >= 5).sum())
    mc = cl.Month.value_counts()
    arr = "{" + ",".join(f'"{m}"' for m in MONTHS) + "}"
    rows = [
        ("Admissions to Emerald Ward.", '=COUNTIF(tClean[Ward],"Emerald Ward")', int((cl.Ward == "Emerald Ward").sum()), "int", 2, ""),
        ("Average bill for Emergency admissions (₦, nearest naira).", '=ROUND(AVERAGEIFS(tClean[Billing_Amount],tClean[Admission_Type],"Emergency"),0)',
         round(cl[cl.Admission_Type == "Emergency"].Billing_Amount.mean()), "naira", 2, ""),
        ("Name of the patient with the highest bill.", "=XLOOKUP(MAX(tClean[Billing_Amount]),tClean[Billing_Amount],tClean[Patient_Name])",
         cl.loc[ix, "Patient_Name"], "gen", 2, "INDEX/MATCH accepted."),
        ("Share of admissions paid by NHIA (National Health Insurance).", '=COUNTIF(tClean[HMO],"NHIA*")/ROWS(tClean)', nh, "pct", 2, "Wildcard or full label."),
        ("Patients aged 75+ with a valid age.", '=COUNTIFS(tClean[Age],">=75",tClean[Age_Valid],"Yes")', int(((cl.Age >= 75) & (cl.Age_Valid == "Yes")).sum()),
         "int", 2, "Age 150 must not count."),
        ("Median length of stay for Cancer (valid rows only).",
         '=AGGREGATE(17,6,tClean[Length_of_Stay]/((tClean[Medical_Condition]="Cancer")*(tClean[LOS_Valid]="Yes")),2)', cancer, "dec1", 2,
         "MEDIAN(FILTER(...)) or MEDIAN(IF(...)) equally valid."),
        ("Admissions on a Saturday or Sunday.", "=SUMPRODUCT(--(WEEKDAY(tClean[Admission_Date],2)>5))", wk, "int", 2, ""),
        ("Admissions in the busiest month.", f"=MAX(COUNTIF(tClean[Month],{arr}))", int(mc.max()), "int", 1, f"Busiest month: {mc.idxmax()}."),
    ]
    for i, (t, f, v, k, m, gd) in enumerate(rows, 1):
        A.q(s, f"4.{i}", t, f, v, k, m, gd, header=(i == 1))
    A.end_task(s)

    # ---------------------------------------------------------------- T5 pivots
    s = A.task("T5 Pivots & PivotCharts", "Task 5 - Management pivot report for the Medical Director", "2.10-2.11",
               "Build PivotTables from tClean on a new sheet called My_Pivots and answer from them. In the answer key the same numbers are reproduced "
               "with formulas (and a live PivotTable) so the marker can check them.")
    wq = cl.groupby(["Ward", "Quarter"]).size()
    top_wq = wq.idxmax()
    hmo = cl.groupby("HMO").Billing_Amount.sum().sort_values(ascending=False)
    topaz = cl[cl.Ward == "Topaz Ward"]
    los_c = valid.groupby("Medical_Condition").Length_of_Stay.mean().sort_values(ascending=False)
    band = int((cl.Age.between(60, 74) & (cl.Age_Valid == "Yes")).sum())
    elect = cl[cl.Admission_Type == "Elective"].Billing_Amount.sum()
    A.q(s, "5.1", "Rows = Ward, Columns = Quarter, Values = Count. Which ward-quarter has the most admissions? (e.g. 'Ruby Ward Q3')",
        f"{top_wq[0]} {top_wq[1]}", None, "gen", 3, f"{int(wq.max())} admissions.", header=True)
    A.q(s, "5.2", "...and how many admissions is that?", f'=COUNTIFS(tClean[Ward],"{top_wq[0]}",tClean[Quarter],"{top_wq[1]}")', int(wq.max()), "int", 2)
    A.q(s, "5.3", "Rows = HMO, Values = Sum of Billing_Amount, sorted largest first. Which payer is top?", hmo.index[0], None, "gen", 3,
        f"₦{hmo.iloc[0]:,.0f}. 'Self-pay / Unknown' must appear as its own row.")
    A.q(s, "5.4", "Show Values As % of Row Total: what share of Topaz Ward admissions were Emergency?",
        '=COUNTIFS(tClean[Ward],"Topaz Ward",tClean[Admission_Type],"Emergency")/COUNTIF(tClean[Ward],"Topaz Ward")',
        (topaz.Admission_Type == "Emergency").mean(), "pct", 3)
    A.q(s, "5.5", "Filter LOS_Valid = Yes, Rows = Medical_Condition, Values = Average of Length_of_Stay. Which condition has the longest average stay?",
        los_c.index[0], None, "gen", 3, f"{los_c.iloc[0]:.2f} days - the differences between conditions are small (all ≈ 15.5).")
    A.q(s, "5.6", "Group Age into bands (1-17, 18-39, 40-59, 60-74, 75-120). How many valid-age admissions fall in 60-74?",
        '=COUNTIFS(tClean[Age],">=60",tClean[Age],"<=74",tClean[Age_Valid],"Yes")', band, "int", 3, "Grouping with 'By' intervals or a helper band column.")
    A.q(s, "5.7", "Add a slicer on Admission_Type and select Elective. Total billing shown?", '=SUMIFS(tClean[Billing_Amount],tClean[Admission_Type],"Elective")',
        elect, "naira", 3)
    A.long(s, "5.8", "Insert a clustered-column PivotChart of admissions by Ward and Quarter. Write the one-sentence headline you would put above it.",
           f"e.g. 'All four wards admitted similar volumes in 2023 (about 2,600-2,800 each); {top_wq[0]} peaked in {top_wq[1]} with {int(wq.max())} admissions.' "
           "A headline states the finding, not the chart type.", marks=3, guide="Chart present 1; headline is a finding with a number 2.", rows=3)
    A.q(s, "5.9", "The EMR team sends a corrected CSV. Which single command refreshes the query AND every pivot?", "Data > Refresh All", marks=2,
        choices=["Data > Refresh All", "PivotTable Analyze > Refresh", "Ctrl + S", "Close and Load"], header=True)
    if key:
        s.band("KEY", "Reference PivotTable (Excel builds it from tClean when the file opens)")
        s.pivot(tc, rows=["Ward"], columns=["Quarter"], values=[{"field": "Patient_Name", "function": "count", "name": "Admissions"}], height=10)
    A.end_task(s)

    # ---------------------------------------------------------------- data
    b.data_sheet("D_Raw", raw_x, "tRaw", fmts={"Admission_Date": "date", "Discharge_Date": "date"},
                 title="Lifecare EMR export, 2023 - AS RECEIVED (table tRaw). Do not clean this sheet.", widths={"Patient_Name": 22, "HMO": 28, "Doctor_Name": 20})
    b.data_sheet("D_WardMap", wardmap, "tWardMap", title="Ward label mapping table (lower-case trimmed key → official ward name)")
    return A


def main():
    for k in (False, True):
        print("Week 2", "KEY" if k else "student", "->", _build(k))


def _build(key):
    global tc
    # create book via build(), but tClean must exist before T5 pivot -> build in two passes is not possible with XlsxWriter,
    # so the key's K_Clean sheet is created lazily: we pre-compute the TRef layout (header row 4, col B) and add the sheet at the end.
    from xlhelp import TRef
    tc = TRef("tClean", "K_Clean", 4, 1, CLEAN_COLS, len(cl))
    A = build(key)
    if key:
        b = A.b
        t = b.data_sheet("K_Clean", cl[CLEAN_COLS], "tClean", fmts={"Admission_Date": "date", "Discharge_Date": "date", "Billing_Amount": "naira2",
                                                                    "Length_of_Stay": "int"},
                         title="KEY - tClean as produced by the reference Power Query (10,920 rows)", widths={"Patient_Name": 22, "HMO": 28})
        assert (t.r0, t.c0) == (tc.r0, tc.c0)
        ks = b.sheet("K_PQ_Script", "KEY - reference Power Query (M) script", "Paste into Home > Advanced Editor; change the file path", "Key")
        ks.code(M_SCRIPT)
        ks.h2("Control totals - cleaning log")
        ks.grid([[a, v] for a, v in steps] + [["Near-duplicate pairs (blank field only)", near], ["True admissions 2023", TRUE_N]],
                header=["Stage", "Rows"], fmts=[None, "int"])
    path = A.finish("Week 2 Assignment - Lifecare 2023 Admissions",
                    ["Scenario: Lifecare General Hospital's Medical Director wants a 2023 admissions report. The only source is a messy export from the old "
                     "EMR system - mixed-case names, ward-label variants, coded genders, text-formatted bills, duplicates and finance reversals.",
                     "Profile the defects (Task 2), build a repeatable Power Query clean-up (Task 3), then answer the Director's questions with functions "
                     "(Task 4) and PivotTables (Task 5). Tasks 4-5 depend on the table tClean that YOU load in Task 3.",
                     "Time guide: 6-7 hours."],
                    [("Lifecare", DATASETS_TEXT["Lifecare"] + " - this assignment uses the deliberately dirty 2023 EMR export (11,026 rows).")],
                    files=f"02_Assignments/Week2_Assignment_Data/{CSV_NAME} (same rows as sheet D_Raw).",
                    rubric=[["Live formulas", "Typed numbers score zero where a formula is required."],
                            ["Follow-through", "If your tClean differs from the key, marks for Tasks 4-5 are awarded when the method is right for YOUR table."],
                            ["Tolerance", "Counts exact; money ±₦1; averages ±0.01."]])
    return path


if __name__ == "__main__":
    main()
