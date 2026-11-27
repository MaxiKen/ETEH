# -*- coding: utf-8 -*-
"""Slide copy for the rewritten modules. Every figure is taken from the bootcamp's own files:
Section Examples (2.4, 3.10, 4.14, P.2 sheets) and the instructor capstone solutions (P.3-P.5)."""

MODULES = {
    # ------------------------------------------------------------------ 2.4
    "2.4": dict(
        title="Web and PDF Data Sources (Practice Files)",
        sector="Supermart Retail",
        desc="Import SuperMart's store directory from a saved web page and the FY2025 sales bulletin from PDF, "
             "then tidy, reconcile to the ledger and refresh with evidence.",
        topics=[
            dict(tag="2.4.1", title="Beginner: inspect the store directory web page",
                 key="Open SuperMart_Store_Directory.html in a browser first: it holds two tables (the 12-store directory and opening hours) "
                     "and a footnote on floor area. Then use Data > From Web with the file path and pick the 'stores' table in the Navigator.",
                 src=("DATA FILE", "Week2_Data_Files/SuperMart_Store_Directory.html  (12 stores × 8 columns)"),
                 before="a web page whose tables, units and footnotes have not been read.",
                 after="the analyst knows the grain (one row per store), the fields and the floor-area caveat.",
                 best="Record the file, the table chosen, the date and the owner (Head Office) before shaping; "
                      "set Opened to Date using Locale English (United Kingdom).",
                 watch="Using Floor_area_sq_m for sales per m² without saying that it excludes car parks.",
                 example="The Navigator shows Table 1 (stores) and Table 2 (opening hours). Table 1 loads 12 rows: "
                         "ST01 Lagos - Ikeja is the largest store (2,400 m²) and ST09 Jos - Rayfield the smallest (1,050 m²).",
                 table=[["Check", "Result"], ["Rows loaded", "12 stores (ST01 – ST12)"],
                        ["Opened column", "Text '02 Mar 2015' → Date (en-GB)"], ["Footnote", "Floor area excludes car parks"]]),
            dict(tag="2.4.2", title="Beginner: import and extract the PDF bulletin",
                 key="Use Get Data > From File > From PDF on SuperMart_Monthly_Sales_Bulletin_2025.pdf. The Navigator shows Table001 "
                     "(store × month, ₦m to one decimal) and Table002 (zone summary); compare cells with the printed page.",
                 src=("DATA FILE", "Week2_Data_Files/SuperMart_Monthly_Sales_Bulletin_2025.pdf  (Table001, Table002)"),
                 before="the PDF is human-readable but not analysis-ready.",
                 after="Table001 is loaded as 12 stores × 12 months plus FY 2025, checked against the page.",
                 best="In the PDF Navigator, compare headings with the printed page and choose the table you need.",
                 watch="Loading the first detected table automatically: PDF connectors also pick up summaries and repeated headers.",
                 example="Spot-check three printed cells: Lagos - Ikeja Dec = 19.3, Port Harcourt - GRA Oct = 24.7 and "
                         "Jos - Rayfield FY = 66.2. All three match the extracted Table001, so the extraction is accepted.",
                 table=[["Bulletin table", "Content"], ["Table001", "12 stores × Jan–Dec + FY 2025 (₦m, 1 dp)"],
                        ["Table002", "Zone summary – not loaded as data"]]),
            dict(tag="2.4.3", title="Intermediate: tidy and normalise the bulletin",
                 key="Promote headers, filter out the Total row, remove the FY 2025 column, then Unpivot Other Columns on Store to get "
                     "Store | Month | Sales_m. Type Sales_m with locale en-GB and add a Month_No column.",
                 src=("QUERY STEP", "Table.UnpivotOtherColumns(NoFY, {\"Store\"}, \"Month\", \"Sales_m\")"),
                 before="a 12 × 12 cross-tab with a total row and total column cannot be filtered.",
                 after="144 tidy rows (12 stores × 12 months), typed, with every step named.",
                 best="Keep the imported PDF in a RAW query and make corrections in a separate query.",
                 watch="Leaving the Total row or the FY column in the data, so every PivotTable double-counts.",
                 example="Before: 12 rows × 14 columns. After Unpivot Other Columns: 144 rows × 4 columns. Unpivot OTHER columns means "
                         "a new month column in next year's bulletin is picked up automatically.",
                 table=[["Step", "Rows × columns"], ["Promoted, Total row removed", "12 × 14"], ["FY 2025 column removed", "12 × 13"],
                        ["Unpivoted + Month_No", "144 × 4"]]),
            dict(tag="2.4.4", title="Intermediate: reconcile the PDF to the ledger",
                 key="Merge the bulletin's FY 2025 figures with the ledger (SuperMart_Stores_Sales_2022_2025.csv) on Store_Name. Check that "
                     "all 12 keys match, round the ledger to ₦0.1m and flag any difference above ₦0.1m.",
                 src=("DATA FILES", "Bulletin PDF  +  00_Core_Datasets/SuperMart_Stores_Sales_2022_2025.csv"),
                 before="nobody knows whether the PDF agrees with the system of record.",
                 after="a 12-row reconciliation shows 0 exceptions, explained as rounding.",
                 best="Normalise the store key (TRIM, case) and inspect unmatched rows before trusting the merge.",
                 watch="Rebuilding the FY total from the 144 rounded monthly cells - rounding drift makes the PDF look wrong.",
                 example="Lagos - Ikeja: PDF 146.4 vs ledger ₦146,388,345 → 146.4, difference 0.0. All 12 stores are OK, yet the 144 "
                         "rounded monthly cells add to ₦732.9m against a ledger of ₦733.2m.",
                 table=[["Store", "PDF ₦m", "Ledger ₦m", "Status"], ["Lagos - Ikeja", "146.4", "146.4", "OK"],
                        ["Port Harcourt - GRA", "93.9", "93.9", "OK"], ["Owerri - New Owerri", "31.6", "31.6", "OK"],
                        ["All 12 stores", "733.0", "733.2", "0 exceptions"]], widths=[3, 1.4, 1.4, 1.8]),
            dict(tag="2.4.5", title="Advanced practical: audit-ready refresh",
                 key="Add a Source column (file, page, Table001), load Row_Count and Total_Check queries to a 99_Checks sheet and save "
                     "a baseline. After each refresh compare rows and totals with the baseline before publishing.",
                 src=("CHECKS SHEET", "99_Checks:  Row_Count = 144  ·  Total_Check = ledger FY 2025 ± ₦0.1m"),
                 before="refreshes may change results without anyone noticing.",
                 after="changed rows and totals are traceable and another analyst can rerun the flow.",
                 best="Never overwrite the ledger with PDF numbers: the PDF is a copy, the ledger is the record.",
                 watch="Assuming a clean-looking extraction proves the PDF table was read correctly.",
                 example="Refresh test: swap in next month's bulletin. Row_Count stays 144 and every store's Total_Check stays within "
                         "₦0.1m of the ledger, so the refresh is accepted and logged with date and initials.",
                 table=[["Control", "Expected after refresh"], ["Row_Count", "144 rows"], ["Total_Check", "Each store within ₦0.1m of ledger"],
                        ["Unmatched stores", "0"]]),
        ]),
    # ------------------------------------------------------------------ 3.10
    "3.10": dict(
        title="File Extracts into SQL, DAX and Dashboards",
        sector="Lifecare Hospital",
        desc="Profile Lifecare's 2023 admissions CSV, stage it in SQLite, expose exceptions with LEFT JOIN and UNION ALL, "
             "define DAX measures and reconcile the dashboard to source.",
        topics=[
            dict(tag="3.10.1", title="Beginner: profile the extract before loading",
                 key="Before loading Lifecare_Admissions_2023_DIRTY.csv, profile every column for rows, blanks and distinct values: 11,026 rows; "
                     "Discharge_Date has 70 blanks, HMO 40, and Ward shows 16 spellings of 4 wards.",
                 before="an extract arrives with unknown blanks, codes and variants.",
                 after="a 15-column profile lists blanks and distinct counts before any SQL.",
                 best="Profile the whole file (Column profiling based on entire data set), not the first 1,000 rows.",
                 watch="Loading straight into production and finding the blanks after the dashboard is built.",
                 example="The profile shows Gender with 6 distinct values (M, F, Male, female…) and Admission_Type with 6 spellings of "
                         "3 real types. Each becomes a mapping rule before the data is loaded.",
                 table=[["Column", "Blank", "Distinct", "Action"], ["Discharge_Date", "70", "392", "Flag open stays"],
                        ["HMO", "40", "5", "Keep NULL, not 'Unknown'"], ["Ward", "0", "16", "Map to 4 wards"],
                        ["Gender", "0", "6", "Map to Male / Female"]], widths=[2.2, 1, 1, 2.6]),
            dict(tag="3.10.2", title="Beginner: ingest the extract into a staging table",
                 key="Import the CSV into SQLite as stg_admissions_2023 (DB Browser: File > Import > Table from CSV, or .import --skip 1). "
                     "Staging keeps the raw text; SELECT COUNT(*) must equal the file's 11,026 data rows.",
                 before="file rows live outside the database with no governed landing place.",
                 after="stg_admissions_2023 holds all 11,026 rows exactly as delivered.",
                 best="Land raw data in a stg_ table; promote to production only after the checks pass.",
                 watch="Letting the import guess types, so '07/03/2023' or '₦1,200' silently become text or NULL.",
                 example="After the import, SELECT COUNT(*) FROM stg_admissions_2023 returns 11,026, matching the file, so no row was "
                         "lost to a stray comma or quote inside a name.",
                 table=[["Check", "Result"], ["CSV data rows", "11,026"], ["SELECT COUNT(*) on staging", "11,026"],
                        ["Status", "PASS – load accepted"]]),
            dict(tag="3.10.3", title="Intermediate: SQL joins and an exception report",
                 key="LEFT JOIN admissions to doctors and keep WHERE d.doctor_id IS NULL to find orphans. Then build one UNION ALL report: "
                     "discharge before admission, bill below zero, stay over 30 days, room outside ward range.",
                 before="an INNER JOIN silently hides admissions with no matching doctor.",
                 after="one exception table lists every rule with its count.",
                 best="Check that doctor_id is unique first, then use LEFT JOIN to discover unmatched rows.",
                 watch="Using INNER JOIN for checks: unmatched rows vanish and the report looks perfect.",
                 example="On lifecare_hospital.db the report returns 0 for all four rules and 0 orphan admissions. Run the same SQL "
                         "on stg_admissions_2023 to watch the dirty extract fail.",
                 table=[["Rule (lifecare_hospital.db)", "Rows"], ["Orphan admissions (no doctor)", "0"], ["Discharge before admission", "0"],
                        ["Bill below zero", "0"], ["Stay over 30 days", "0"], ["Room outside ward range", "0"]]),
            dict(tag="3.10.4", title="Intermediate: DAX measures and filter context",
                 key="Define explicit measures: Admissions = COUNTROWS(Admissions); ALOS = AVERAGE(Admissions[length_of_stay]); "
                     "Readmission % = DIVIDE(CALCULATE([Admissions], Admissions[readmit_30d] = \"Yes\"), [Admissions]).",
                 before="implicit sums and averages change meaning when slicers change.",
                 after="named measures with definitions answer consistently in every visual.",
                 best="Build ratios with DIVIDE and CALCULATE so they recalculate in each filter context.",
                 watch="Averaging a stored readmission % column instead of dividing counts in context.",
                 example="Across all years the model returns 54,860 admissions, ALOS 15.5 days and readmission 10.4%. Filter to Emergency "
                         "and the same Readmission % measure returns 13.7%: same formula, new context.",
                 table=[["Measure", "Result (all years)"], ["Admissions", "54,860"], ["ALOS", "15.5 days"], ["Readmission %", "10.4%"],
                        ["Emergency readmission %", "13.7%"]]),
            dict(tag="3.10.5", title="Advanced practical: reconcile a governed dashboard",
                 key="Before publishing, compare three numbers under the same filters: rows and billing in the source CSV, COUNT(*) and "
                     "SUM(billing_amount) in SQL, and the dashboard cards. Publish only when all three agree.",
                 before="dashboard totals may differ from source through joins or filters.",
                 after="source, SQL and dashboard agree and the check shows PASS.",
                 best="Reconcile SQL and dashboard measures under the same filters and date range.",
                 watch="Comparing a filtered dashboard total with an unfiltered SQL total.",
                 example="Lifecare_Admissions_2019_2024.csv holds 54,860 rows and ₦70,206,080,066.78 of billing. SQL returns the same "
                         "count and the same sum to the kobo, so the reconciliation cell shows PASS.",
                 table=[["Control", "Source CSV", "SQL", "Result"], ["Admissions", "54,860", "54,860", "PASS"],
                        ["Billing (₦)", "70,206,080,066.78", "70,206,080,066.78", "PASS"]], widths=[1.6, 2.2, 2.2, 1]),
        ]),
    # ------------------------------------------------------------------ 4.14
    "4.14": dict(
        title="Internal Benchmarks, PDF Checks and Reconciliation",
        sector="Supermart Retail",
        desc="Benchmark each SuperMart store against the chain median, verify the PDF bulletin against the ledger, align periods "
             "and definitions, show uncertainty and write a decision memo.",
        topics=[
            dict(tag="4.14.1", title="Beginner: understand benchmark fitness",
                 key="Benchmark each store against the chain, not an outside source: FY2025 net sales (after discount, excluding VAT) "
                     "per m² of floor area and margin points, compared with the chain median rather than the mean.",
                 before="stores ranked on total sales, so the biggest always looks best.",
                 after="same definition, period and unit (sales per m²) for all 12 stores.",
                 best="Match definition, period and unit; use the median so one big store cannot drag it.",
                 watch="Ranking on total sales: Lagos - Ikeja is big (2,400 m²), not necessarily efficient.",
                 example="The chain median is ₦37,517 of sales per m². Jos - Rayfield leads at ₦63,074 (+68%); Abuja - Wuse is lowest "
                         "at ₦19,737 (−47%) although it is the second-largest store (2,200 m²).",
                 table=[["Store", "Sales per m²", "vs median"], ["Jos - Rayfield", "₦63,074", "+68%"], ["Lagos - Ikeja", "₦60,995", "+63%"],
                        ["Ibadan - Ring Road", "₦20,576", "−45%"], ["Abuja - Wuse", "₦19,737", "−47%"]]),
            dict(tag="4.14.2", title="Beginner: extract and verify the PDF",
                 key="Extract Table001 from SuperMart_Monthly_Sales_Bulletin_2025.pdf (Week 2, 2.4), keep the FY 2025 column and compare "
                     "each store with the ledger. Check the ₦m unit, one-decimal rounding and the total row.",
                 before="extracted numbers may have shifted columns or lost their ₦m unit.",
                 after="all 12 stores agree with the ledger within ±₦0.05m: 'Rounding only'.",
                 best="Reconcile to the ledger, not to a sum of rounded months; log any manual fix.",
                 watch="Treating PDF extraction as exact without a visual cross-check of a few cells.",
                 example="Benin City - Ring Road prints FY 62.1, but its 12 printed months add to 61.9. The ledger says ₦62,116,000 → 62.1, "
                         "so the FY figure is right and the gap is rounding in the months.",
                 table=[["Store", "Bulletin FY", "Sum of months", "Ledger", "Status"],
                        ["Lagos - Ikeja", "146.4", "146.5", "146.4", "Rounding only"],
                        ["Benin City - Ring Road", "62.1", "61.9", "62.1", "Rounding only"],
                        ["Owerri - New Owerri", "31.6", "31.5", "31.6", "Rounding only"]], widths=[2.6, 1.3, 1.5, 1.1, 1.6]),
            dict(tag="4.14.3", title="Intermediate: reconcile time periods and definitions",
                 key="Align periods and definitions first: a calendar year and an April year-end cover different months, and SuperMart "
                     "Online gross bookings include cancelled and returned orders while net revenue does not.",
                 before="mixed periods and definitions create false growth or false gaps.",
                 after="each figure carries its period and definition; variance has a stated base.",
                 best="Label every number with its period and definition, and state the variance formula.",
                 watch="Comparing one company's gross bookings with another's net revenue.",
                 example="Calendar FY2025 sales are ₦733.2m, but April 2024 – March 2025 gives ₦639.1m. SuperMart Online 2025 gross bookings "
                         "are ₦432.0m against ₦386.7m net: gross overstates net by 11.7%.",
                 table=[["Measure", "Value", "Definition"], ["SuperMart FY2025", "₦733.2m", "Calendar year, net of discount"],
                        ["SuperMart Apr 2024 – Mar 2025", "₦639.1m", "April year-end"],
                        ["Online gross bookings 2025", "₦432.0m", "All orders"], ["Online net revenue 2025", "₦386.7m", "Delivered + Completed"]],
                 widths=[2.6, 1.2, 2.4]),
            dict(tag="4.14.4", title="Intermediate: uncertainty and missing data", sector="Lifecare Hospital",
                 key="Distinguish missing from zero: the Week 2 dirty extract has 46 lines with no Quantity, so report total units as a range "
                     "until stores confirm. For rates, show 95% confidence intervals before ranking anything.",
                 before="blanks treated as zero and small rate gaps presented as real.",
                 after="missing values disclosed; only non-overlapping intervals called different.",
                 best="Keep missing, not applicable and zero distinct; publish the interval with each rate.",
                 watch="Ranking departments on point estimates whose 95% intervals overlap.",
                 example="Lifecare FY2023 readmission: Oncology 13.2% (11.7–14.8%) sits above Endocrinology, Orthopaedics & Rheumatology "
                         "and Respiratory. Cardiology's 11.4% overlaps it, so the two are not ranked.",
                 table=[["Department", "Rate", "95% CI"], ["Cardiology", "11.4%", "10.0% – 12.9%"], ["Endocrinology", "9.9%", "9.0% – 10.9%"],
                        ["Oncology", "13.2%", "11.7% – 14.8%"], ["Orthopaedics & Rheumatology", "8.3%", "7.0% – 9.5%"],
                        ["Respiratory", "9.2%", "7.9% – 10.6%"]], widths=[3, 1.2, 2]),
            dict(tag="4.14.5", title="Advanced practical: benchmark decision memo",
                 key="Write a one-page memo: decision needed, benchmark used, finding, uncertainty and recommendation. Keep observed fact "
                     "(sales per m²) separate from inference and from the action proposed.",
                 before="a benchmark chart overstates certainty and implies the action.",
                 after="the decision is evidence-led and its limits and next check are stated.",
                 best="Separate observed fact, inference and recommendation; add sign-off lines.",
                 watch="Presenting a benchmark rank as proof that refurbishment will work.",
                 example="Finding: Abuja - Wuse has the lowest sales per m² (₦19,737), 47% below the chain median. Recommendation: shortlist "
                         "the bottom three (Wuse, Ibadan, Garki) for a site review before the FY2026 budget.",
                 table=[["Memo section", "Content"], ["Decision needed", "Which stores get the FY2026 refurbishment budget?"],
                        ["Benchmark", "Chain median sales per m² and margin, FY2025 ledger"],
                        ["Uncertainty", "One year; floor area excludes car parks; unaudited"],
                        ["Recommendation", "Site review of the bottom three before funding"]], widths=[1.6, 4]),
        ]),
    # ------------------------------------------------------------------ P.2
    "P.2": dict(
        title="Source Governance, Citation and Capstone Delivery",
        sector="Multi-Sector",
        desc="Register every capstone source, keep before and after states, protect personal data, make the work reproducible "
             "and hand over evidence a reviewer can follow.",
        topics=[
            dict(tag="P.2.1", title="Beginner: source register essentials",
                 key="For each source record title, owner, file path, extraction date, period, units, row count and intended use. "
                     "Capstone 1 registers the ledger, the store directory and the dirty FY2025 store returns.",
                 src=("REGISTER ENTRY", "SuperMart_Sales_2025_DIRTY.csv · Store Ops · FY2025 · 3,387 rows · ₦ · cleaning"),
                 before="values circulate without provenance.",
                 after="every dataset can be traced to a file, an owner and a date.",
                 best="Capture title, owner, file, extraction date, period, unit, row count and intended use.",
                 watch="Keeping only a copied value with no source or date.",
                 example="Capstone 1's register has three rows: the ledger (9,994 rows, 2022–2025), the directory (12 stores) and the "
                         "store returns (3,387 rows). A reviewer opens each file and sees the same counts.",
                 table=[["Source", "Rows", "Owner · period"], ["SuperMart_Stores_Sales_2022_2025.csv", "9,994", "Finance · 2022–2025"],
                        ["SuperMart_Stores_Directory.csv", "12", "Head Office · current"],
                        ["SuperMart_Sales_2025_DIRTY.csv", "3,387", "Store Ops · FY2025"]], widths=[3.4, 0.9, 2.2]),
            dict(tag="P.2.2", title="Beginner: preserve before and after states",
                 key="Keep the RAW file untouched. The BEFORE state records the raw file and its known issues; the AFTER state records each "
                     "transformation, the output and its checks. Never overwrite source evidence.",
                 src=("BEFORE → AFTER", "SuperMart_Sales_2025_DIRTY.csv (3,387 rows)  →  tClean25 (3,312 rows)"),
                 before="only a modified file remains and its changes cannot be explained.",
                 after="raw, cleaned and reported states are separate, named and documented.",
                 best="Keep an untouched before copy, a separately named after output and a change log with counts.",
                 watch="Editing the only source copy and losing the starting state.",
                 example="The change log reads: 3,387 raw rows − 15 spacer rows − 60 exact duplicates = 3,312 clean rows, with text "
                         "trimmed and case fixed first so near-duplicates are caught.",
                 table=[["State", "Rows", "Evidence"], ["Before (raw file)", "3,387", "Untouched CSV in 01_RAW"],
                        ["Spacer rows removed", "15", "Change log"], ["Duplicates removed", "60", "Change log"],
                        ["After (clean table)", "3,312", "Reconciles to ₦733,215,255"]], widths=[2.4, 1, 3]),
            dict(tag="P.2.3", title="Intermediate: citation, licence and privacy",
                 key="Cite each dataset by name and origin (e.g. built from 'Sample - Superstore'). Remove names and IDs before anything "
                     "leaves the organisation, and keep sensitive files such as the wellbeing survey out of portfolios.",
                 src=("RESTRICTED", "Crestview_Wellbeing_Survey_SENSITIVE.csv  ·  Capstone 2 at-risk list (student IDs)"),
                 before="reuse terms and privacy risks are overlooked.",
                 after="attribution and data protection are built into delivery.",
                 best="Drop direct identifiers and aggregate small groups before sharing outputs (NDPA 2023).",
                 watch="Assuming a file is shareable because you were able to download or query it.",
                 example="Capstone 1 drops Customer_Name before analysis; Capstone 3 never extracts patient names; Capstone 2's at-risk "
                         "list carries student IDs, so that workbook is marked restricted.",
                 table=[["Capstone", "Personal data", "Control"], ["1 SuperMart", "Customer names", "Dropped before analysis"],
                        ["2 Crestview", "Student IDs on at-risk list", "Restricted distribution"],
                        ["3 Lifecare", "Patient names", "Never extracted by the SQL"]]),
            dict(tag="P.2.4", title="Intermediate: reproducibility and review",
                 key="Name every transformation step, keep thresholds in input cells, test a refresh, reconcile key totals and ask "
                     "another person to follow the README unaided. Record exceptions and sign-off.",
                 src=("CONTROL TOTALS", "Capstone 1: 3,312 rows · ₦733,215,255 sales · ₦93,439,242 profit"),
                 before="only the analyst can reproduce the result.",
                 after="another analyst can refresh, audit and get the same totals.",
                 best="Keep rules such as 'exam below 62 or attendance below 65%' in input cells.",
                 watch="Hard-coding a threshold inside hundreds of formulas, so a policy change means a rebuild.",
                 example="Change Crestview's attendance threshold from 65% to 70% in its input cell and the at-risk list, dashboard and "
                         "senate figures all recalculate - no formula is edited.",
                 table=[["Check", "Evidence"], ["Named steps", "Applied Steps / formula map"], ["Inputs", "Yellow cells for every threshold"],
                        ["Control totals", "All OK on the reconcile sheet"], ["Peer refresh", "Colleague reproduces it unaided"]]),
            dict(tag="P.2.5", title="Advanced capstone: evidence-to-decision handoff",
                 key="Hand over the source register, before/after samples, the workbook or SQL, quality checks, findings with "
                     "uncertainty, an owner and a follow-up date. Keep observed facts apart from recommendations.",
                 src=("HANDOFF PACK", "Solution workbook + extraction SQL + instructor brief PDF + README"),
                 before="a polished presentation may lack evidence.",
                 after="the chain from source file to decision can be reviewed end to end.",
                 best="Ask a peer to reproduce one key claim from the source file to the final number.",
                 watch="Calling work reproducible because the author can rerun it.",
                 example="Capstone 3's handoff is the SQL file, the solution workbook and the brief. A peer reruns Q1 and gets 10,918 "
                         "FY2023 admissions and ₦13.98bn billing: claim verified.",
                 table=[["Deliverable", "Purpose"], ["Source register", "Where each number came from"], ["Workbook / SQL", "How it was produced"],
                        ["Findings + uncertainty", "What it means and how sure we are"], ["Owner + follow-up date", "Who acts, and when"]]),
        ]),
}

# ---------------------------------------------------------------------------------------------- P.3 - P.5
# keys: original Vol 5 slide numbers
CAPSTONE = {
    35: {"TextBox 5": "A complete retail build on the FY2025 store ledger: brief, data, cleaning, model, basket, margin and SKU "
                      "analysis, dashboard, insight memo and verification."},
    36: {"TextBox 9": "The SuperMart board asks why FY2025 sales grew 20.4% to ₦733.2m across 12 stores while margin fell from 13.4% "
                      "to 12.7%, which stores and products are responsible, and what to do in FY2026. Success: reconciled numbers, "
                      "a quantified cause and a costed recommendation.",
         "TextBox 13": "The signed brief commits to a ledger-reconciled model, a one-screen dashboard and a memo costing a 20% "
                       "Furniture discount cap, worth ₦17.7m–₦18.0m of extra profit depending on volume loss.",
         "Table 14": [["Brief Element", "Content", "Acceptance"], ["Question", "Why margin fell while sales grew", "Answered with evidence"],
                      ["Output", "Model, dashboard, memo", "All reconciled to the ledger"]]},
    37: {"TextBox 9": "Three sources: the store ledger (9,994 lines, 2022–2025), the store directory (12 stores with FY2025 targets) "
                      "and the FY2025 store returns as submitted (3,387 rows, dirty). Each is copied to 01_RAW before any "
                      "transformation.",
         "TextBox 17": "Treating line count as order count: 3,312 FY2025 lines are only 1,687 orders.",
         "TextBox 25": "The inventory records ledger 9,994 rows (FY2025: 3,312), directory 12 rows and returns 3,387 rows, each with "
                       "a named owner. Any row-count change at refresh is immediately meaningful."},
    38: {"Table 9": {"rows": [["Source", "Rows", "Owner"], ["SuperMart_Stores_Sales_2022_2025.csv", "9,994", "Finance"],
                              ["SuperMart_Stores_Directory.csv", "12", "Operations"],
                              ["SuperMart_Sales_2025_DIRTY.csv", "3,387", "Store managers"]], "widths": [3, 1, 1.6]}},
    39: {"TextBox 9": "Split the combined column, TRIM and CLEAN text, fix case and vocabulary, convert ₦-text and %-text to numbers, "
                      "standardise dates, flag impossible ship dates, then remove 15 spacer rows and 60 exact duplicates. The "
                      "cleaning log records every count.",
         "TextBox 21": "BEFORE STATE: defects and valid detail are mixed. AFTER STATE: cleaning rules preserve the line grain.",
         "TextBox 25": "Cleaning takes 3,387 raw rows to 3,312 valid lines across 16 defect types, and the result reconciles to "
                       "₦733,215,255 sales and ₦93,439,242 profit - each figure stated in the log."},
    40: {"Table 9": [["Cleaning Step", "Rows Removed", "Disclosed"], ["Spacer rows", "15", "Yes"], ["Exact duplicates", "60", "Yes"],
                     ["Text, number and date fixes", "0 (values corrected)", "Defect log"]]},
    41: {"TextBox 9": "The FY2024–FY2025 ledger relates to the store table on Store_ID. Measures: Sales, Profit, Margin % "
                      "(Σprofit ÷ Σsales), Orders, Growth % and Share of Total. Every measure reconciles to the ledger control "
                      "totals before analysis starts.",
         "TextBox 17": "Averaging line margins instead of dividing total profit by total sales.",
         "TextBox 25": "The model reports FY2025 sales ₦733,215,255 and margin 12.74% (FY2024: ₦609,205,599 at 13.43%), matching "
                       "the ledger exactly - the reconciliation that qualifies the model for analysis."},
    42: {"Table 9": [["Measure", "Value", "Reconciles"], ["FY2025 sales", "₦733,215,255", "Ledger: exact"],
                     ["FY2025 margin %", "12.74%", "Σprofit ÷ Σsales"], ["Growth vs FY2024", "+20.4%", "Derived, documented"]]},
    43: {"TextBox 9": "Basket analysis asks how big a typical order is and which categories travel together. FY2025: 1,687 orders "
                      "from 3,312 lines (1.96 lines per order); mean order ₦434,627 but median only ₦148,257.",
         "TextBox 21": "BEFORE STATE: the output uses line rather than order grain. AFTER STATE: order values use distinct "
                       "Order_IDs.",
         "TextBox 25": "Office Supplies appear in 75.4% of orders, but only 18.7% of orders combine Office Supplies with Technology "
                       "- the natural target for a cross-category bundle."},
    44: {"Table 9": [["Basket Metric", "Value", "Action"], ["Single-line orders", "51.9%", "Add-on prompts at checkout"],
                     ["Median order value", "₦148,257", "Promotion threshold"], ["Orders with Office Supplies", "75.4%", "Bundle anchor"]]},
    45: {"TextBox 9": "Explain the 0.7-point margin fall by splitting it into rate and mix effects by category and by discount "
                      "band. Lines discounted above 20% lost ₦47.6m in FY2025, and Furniture's margin fell from 3.5% to 1.4%.",
         "TextBox 25": "Rate effects: Office Supplies −0.98 pts, Furniture −0.62, Technology +0.40; mix +0.51. Capping discounts at "
                       "20% lifts modelled margin to 20.9%, assuming 25% volume loss on the capped lines."},
    46: {"Table 9": [["Discount band", "FY2025 margin", "Action"], ["0%", "29.9%", "Protect full-price lines"], ["1–20%", "13.3%", "Monitor"],
                     ["21–40%", "−17.6%", "Cap at 20%"], [">40%", "−92.5%", "Stop"]]},
    47: {"TextBox 9": "Compare the SKU master (2022–2025) with FY2025 sales using a right anti join: 337 of 1,862 SKUs sold nothing in "
                      "FY2025, 230 of them sold in FY2024. The bottom 30% of selling SKUs bring only 1.3% of revenue.",
         "TextBox 21": "BEFORE STATE: coverage and window are unknown. AFTER STATE: the stock recommendation states its window and "
                       "threshold.",
         "TextBox 25": "271 SKUs sold at a loss in FY2025, costing ₦38.3m; Tables alone lost ₦8.1m on ₦60.9m of sales. The 230 "
                       "lapsed SKUs go to Merchandising for a stock check, not automatic delisting."},
    48: {"Table 9": [["Segment", "SKUs", "FY2025 effect"], ["No FY2025 sales", "337", "₦0 revenue"],
                     ["Lapsed (sold FY2024, not FY2025)", "230", "Stock check"], ["Loss-making in FY2025", "271", "−₦38.3m profit"]]},
    49: {"TextBox 9": "One screen: four KPI cards (Sales, Profit, Margin %, Growth), sales and margin by store, a discount-band margin "
                      "chart and a quarterly margin trend. Year and zone slicers drive every visual through report connections.",
         "TextBox 13": "The dashboard puts the five loss-making stores (Abuja - Wuse, Onitsha, Ibadan, Abuja - Garki, Benin City) beside "
                       "the discount bands that explain them.",
         "Table 14": [["Zone", "Visual", "Question Answered"], ["Top", "4 KPI cards", "How are we doing overall?"],
                      ["Middle", "Store sales and margin bars", "Which stores are responsible?"],
                      ["Bottom", "Discount bands, quarterly trend", "Why did margin fall?"]]},
    50: {"TextBox 9": "A one-page memo: headline (sales +20.4%, margin down 0.7 points, driven by deep discounts), three supporting "
                      "numbers, the recommendation with owner and expected effect, and the data-quality note.",
         "TextBox 13": "The memo's ask is specific: cap Furniture discounts at 20% from Q1 FY2026; expected profit uplift ₦17.7m–₦18.0m; owner: "
                       "Head of Merchandising; loss stores reviewed monthly."},
    51: {"TextBox 9": "Verification: sales and profit reconcile to the ledger, row counts are documented, slicers tested and a "
                      "second-machine render confirmed. Handover: README with refresh steps, measure definitions and known "
                      "limitations.",
         "TextBox 25": "The handover pack records: 3,312 FY2025 lines; sales ₦733,215,255 = ledger; profit ₦93,439,242; margin "
                       "12.74%; limitation: profit is line-level, store overheads are not included."},
    # ---------------------------------------------------------------- P.4 Crestview
    53: {"TextBox 5": "A complete higher-education build: registry export cleaning, cohort and progression rule, drivers, at-risk "
                      "list, intervention estimate and a senate paper."},
    54: {"TextBox 9": "The senate's target is 90% progression; in 2024/25 only 86.4% of 6,607 students met the rule (exam ≥ 62 and "
                      "attendance ≥ 65%). It asks which faculties are failing, what drives performance and who needs Week-5 support.",
         "TextBox 13": "The brief fixes the cohort (one row per registered student after de-duplication, Student_ID as key) before "
                       "any data is touched, preventing progression measured on whoever happens to be present.",
         "Table 14": [["Brief Element", "Content", "Risk Controlled"], ["Cohort definition", "2024/25, de-duplicated", "Denominator drift"],
                      ["Deliverable", "Paper plus at-risk list", "Actionable output"]]},
    55: {"TextBox 9": "One registry export, Crestview_Registry_Export_RAW.csv (6,652 rows), carries faculty, department, level, "
                      "attendance, study habits and exam score. The published cohort file (6,607 students) is the control total.",
         "TextBox 13": "The inventory records 6,652 raw rows, 14 departments in 4 faculties and 8 defect types, with the registry as "
                       "owner of the export and the examinations office as owner of the scores.",
         "Table 14": [["Source", "Rows", "Cleaning Complexity"], ["Registry export (raw)", "6,652", "High – 8 defect types"],
                      ["Published cohort (control)", "6,607", "None – reconciliation target"],
                      ["Department master", "14 departments", "Low – mapping table"]]},
    56: {"TextBox 9": "Trim padded Student_IDs, map M/F codes, convert '84%' text to numbers, map department abbreviations to 14 "
                      "canonical names, standardise Yes/No variants, flag one out-of-range score, then remove 45 double "
                      "registrations.",
         "TextBox 13": "Cleaning keeps 6,607 unique students from 6,652 rows. Remove duplicates only after trimming the ID: 25 padded "
                       "IDs would otherwise survive as 'different' students.",
         "Table 14": [["Cleaning Action", "Rows Affected", "Outcome"], ["Duplicate registrations", "45", "6,607 students"],
                      ["Department variants", "198", "14 canonical names"], ["Percent stored as text", "330", "Numeric attendance"]]},
    57: {"TextBox 9": "Progression % = students meeting the rule ÷ cohort students, where the rule is exam score ≥ 62 and attendance "
                      "≥ 65%. Rule, thresholds and cohort sit in input cells before the calculation, so the number is "
                      "defensible rather than convenient.",
         "TextBox 25": "Progression of 86.4% = 5,706 of 6,607 students; 901 (13.6%) are at risk. Against the 90% target the gap is "
                       "about 240 students - the number the senate paper must address."},
    58: {"Table 9": [["Component", "Definition", "Value"], ["Numerator", "Exam ≥ 62 and attendance ≥ 65%", "5,706"],
                     ["Denominator", "2024/25 cohort, de-duplicated", "6,607"]]},
    59: {"Title 1": "Where At-Risk Students Concentrate",
         "TextBox 9": "Break the at-risk rate down by faculty and by attendance band, because concentration is where intervention "
                      "pays. Faculty rates are flat (12.7%–14.6%), but attendance bands separate students sharply.",
         "TextBox 13": "Students below 70% attendance have a 53.5% at-risk rate, against 3.1% at 70–79% and 0% at 90%+. Attendance "
                       "correlates 0.58 with exam score - the strongest single signal in the data.",
         "Table 14": [["Segment", "At-risk rate", "Interpretation"], ["Institution", "13.6%", "Baseline"],
                      ["Faculties (range)", "12.7% – 14.6%", "No faculty is the problem"],
                      ["Attendance below 70%", "53.5%", "Primary early warning"]]},
    60: {"TextBox 9": "Define at-risk as exam score below 62 or attendance below 65%. The list must be actionable: student ID, "
                      "faculty, department, level, the specific trigger and the recommended support route.",
         "TextBox 25": "The list returns 901 students: 585 on low attendance only, 161 on low score only and 155 on both. Each row "
                       "names the trigger so the adviser knows what to address in the meeting."},
    61: {"Table 9": [["Trigger", "Students", "Support Route"], ["Attendance below 65% only", "585", "Welfare and attendance check"],
                     ["Exam below 62 only", "161", "Academic adviser"], ["Both", "155", "Escalated case"]]},
    62: {"TextBox 9": "Estimate a Week-5 attendance programme from the regression (attendance, hours studied, previous score, "
                      "tutoring; R² 0.60), applied to students below 65% attendance. State the assumption and a range.",
         "TextBox 25": "Each attendance point is worth about 0.20 exam points. Lifting low attenders by 10 points moves 652 of the "
                       "901 at-risk students over the rule, raising progression to 96.2% (range 95.2%–96.9%)."},
    63: {"Table 9": [["Assumption", "Value", "Sensitivity"], ["Attendance uplift", "+10 points", "Range +5 to +15"],
                     ["Students moved over the rule", "652", "Range 585 – 694"], ["Institutional progression", "96.2%", "Range 95.2% – 96.9%"]]},
    64: {"TextBox 9": "One screen for the senate: KPI cards (Cohort, Progression %, At-Risk, Mean Exam Score), a faculty bar chart, "
                      "at-risk rate by attendance band and the at-risk list. Slicers: faculty, department, level.",
         "TextBox 13": "Flat faculty bars beside a steep attendance-band chart show the senate at once that the problem is "
                       "behaviour, not one faculty - with the 901-student list one click away.",
         "Table 14": [["Zone", "Visual", "Senate Question"], ["Top", "4 KPI cards", "How are we performing?"],
                      ["Middle", "Faculty bars, attendance bands", "Where is the risk?"], ["Bottom", "At-risk list", "Who needs help?"]]},
    65: {"TextBox 9": "A two-page paper: context (90% target), finding (86.4%; risk driven by attendance, not faculty), evidence "
                      "(53.5% at risk below 70% attendance), recommendation (Week-5 attendance programme, owner Dean) and risks.",
         "TextBox 13": "The recommendation states the estimate as a range (585–694 students) and calls it an association, not proof "
                       "of cause - so it survives scrutiny if the pilot result differs.",
         "Table 14": [["Section", "Content", "Length"], ["Context and target", "90% progression target", "Half page"],
                      ["Finding", "86.4%; attendance-driven", "Half page"], ["Recommendation", "Programme, owner, range", "One page"]]},
    66: {"TextBox 9": "Verify the cohort against the published file (6,607), the at-risk count against the rule (901), and that every "
                      "slicer drives every visual. Handover includes the definitions document, the most contested artefact in any "
                      "university report.",
         "TextBox 13": "Handover records: cohort 6,607 = published file; at risk 901 = rule; 45 duplicates removed; owner: Analytics "
                       "Unit. The list holds student IDs, so the workbook is restricted.",
         "Table 14": [["Verification", "Control Source", "Result"], ["Cohort count", "Published cohort file", "6,607 exact"],
                      ["At-risk count", "Rule in input cells", "901 exact"]]},
    # ---------------------------------------------------------------- P.5 Lifecare
    67: {"TextBox 5": "A complete healthcare build from lifecare_hospital.db: SQL extraction, occupancy, length of stay, readmission "
                      "safety with confidence intervals, payer mix and a board pack."},
    68: {"TextBox 9": "The board asks how full the hospital is, which patients stay longest, whether any department's readmission rate "
                      "is genuinely worse and whether the payer mix carries risk. Success: bed-day occupancy, intervals, no patient "
                      "identifiers.",
         "Table 14": [["Brief Element", "Content", "Constraint"], ["Question", "FY2023 capacity and safety", "Operational focus"],
                      ["Data scope", "Non-identifying columns only", "No PHI in outputs"]]},
    69: {"TextBox 9": "One SQLite database, lifecare_hospital.db: admissions (54,860 rows, 2019–2024), patients (48,803), doctors (28), "
                      "wards (4 × 130 beds) and bed_census (1,797 days). Five SQL queries extract only what the pack needs.",
         "TextBox 13": "The inventory notes 10,918 FY2023 admissions, 520 staffed beds and 365 census days, with the database owned "
                       "by the Data Office and the census by the Nursing Directorate.",
         "Table 14": [["Source", "Volume", "Access Method"], ["admissions", "54,860 rows", "SQL query (Q2)"],
                      ["bed_census", "1,797 days", "SQL query (Q3)"], ["patients", "48,803 rows", "Not extracted – identifiers"]]},
    70: {"Title 1": "Extracting and Reconciling Clinical Data",
         "TextBox 9": "Extract with SQL, not copy-paste: Q1 returns control totals by year, Q2 the FY2023 admissions without names, "
                      "Q3 the daily census. Reconcile each extract's rows and billing to Q1 before any analysis.",
         "TextBox 13": "Q2 returns 10,918 FY2023 admissions and ₦13,984,047,685.44 of billing; Q1 agrees to the kobo. The dirty 2023 CSV "
                       "(11,026 rows) stays in the 3.10 cleaning exercise, never in the board pack.",
         "Table 14": [["Control", "Extract (Q2)", "Q1 control"], ["FY2023 admissions", "10,918", "10,918"],
                      ["FY2023 billing", "₦13.98bn", "₦13.98bn"]]},
    71: {"TextBox 9": "Occupancy % = occupied bed-days ÷ available bed-days from the daily census (patients in house ÷ staffed beds), "
                      "never an average of daily percentages or a count of admissions. The denominator choice is documented.",
         "TextBox 25": "FY2023 occupancy is 89.5% = 169,923 ÷ 189,800 bed-days: above the 85% target in 11 of 12 months and above "
                       "95% on 43 days. Reaching 85% needs about 28 more staffed beds."},
    72: {"Table 9": [["Component", "Definition", "Value"], ["Numerator", "Occupied bed-days", "169,923"],
                     ["Denominator", "Available bed-days (520 × 365)", "189,800"]]},
    73: {"TextBox 9": "Report mean (15.5 days) and median (15.0 days) together, plus the 90th percentile (28 days) for discharge "
                      "planning. Stays are spread evenly from 1 to 30 days and the four wards differ by under 0.3 days.",
         "TextBox 13": "The tail matters: 3,319 stays of 22+ days use 51% of all bed-days. Discharge planning aimed at that group "
                       "frees far more capacity than shaving a day off the median.",
         "Table 14": [["Measure", "Value", "Planning Use"], ["Median LOS", "15.0 days", "Typical experience"],
                      ["90th percentile", "28 days", "Discharge planning"], ["Stays of 22+ days", "3,319", "51% of bed-days"]]},
    74: {"TextBox 9": "A 30-day readmission rate of 10.3% is computed on 10,918 FY2023 admissions and broken down by department with "
                      "95% intervals. Readmission is a safety metric: findings need clinical validation before publication.",
         "TextBox 25": "Only Oncology (13.2%, CI 11.7–14.8%) sits above the hospital rate; Orthopaedics & Rheumatology (8.3%) sits "
                       "below. Both go to the clinical governance committee before the pack is issued."},
    75: {"Table 9": [["Metric", "Value", "Route"], ["Readmission rate", "10.3%", "Clinical validation first"],
                     ["Above hospital rate", "Oncology (13.2%)", "Governance review"]]},
    76: {"TextBox 9": "Define escalation rules explicitly: stay of 22 days or more, readmission within 30 days, or age 75 and over. "
                      "Encode them as one conditional column so every report applies the same rule, with the trigger stated.",
         "TextBox 25": "In FY2023 the three triggers flag 5,099 admissions (47%): 3,319 long stays, 1,129 readmissions, 1,764 aged 75+ "
                       "(with overlap). A rule that flags half the hospital needs clinical tightening."},
    77: {"Table 9": [["Trigger", "Count", "Action"], ["LOS 22+ days", "3,319", "Discharge review"], ["Readmitted within 30 days", "1,129", "Clinical review"],
                     ["Age 75+", "1,764", "Care plan check"]]},
    78: {"TextBox 9": "One screen for the board: KPI cards (Admissions, Occupancy %, ALOS, Readmission %), monthly occupancy against "
                      "the 85% target, admissions by ward, LOS bands and payer mix. No patient identifiers anywhere.",
         "TextBox 13": "The board sees 10,918 admissions, 89.5% occupancy, 15.5-day ALOS and 10.3% readmission, all driven by month and "
                       "ward slicers - and nothing on screen can identify a patient.",
         "Table 14": [["Zone", "Visual", "Safeguard"], ["Top", "KPI cards", "Aggregates only"],
                      ["Middle", "Occupancy trend, ward bars", "Minimum group size applied"], ["Bottom", "LOS bands, payer mix", "No identifiers"]]},
    79: {"TextBox 9": "The pack: one-page summary, capacity, length of stay, a readmission safety note with intervals, payer mix "
                      "(five HMOs at about 20% each) and a data-quality appendix, with the confidentiality statement.",
         "TextBox 13": "The pack recommends 28 staffed beds or a discharge team for 22+ day stays (owner: COO) and an Oncology "
                       "readmission review (owner: Medical Director), flagged as under clinical review.",
         "Table 14": [["Page", "Content", "Purpose"], ["1", "Summary and recommendations", "The decision"],
                      ["2–4", "Capacity, LOS, readmissions", "The evidence"], ["Appendix", "Data quality and SQL controls", "Honesty about limits"]]},
    80: {"TextBox 9": "Verify admissions against the database control totals (10,918), billing against Q1 (₦13,984,047,685.44), "
                      "occupancy against the census, and confirm no PHI in any output. Handover documents the definitions clinicians "
                      "will contest.",
         "TextBox 13": "Handover records: 10,918 admissions = Q1; ₦13.98bn billing = Q1 to the kobo; occupancy 89.5% = census; PHI "
                       "scan: none present; owner: Analytics Unit.",
         "Table 14": [["Verification", "Control", "Result"], ["Admissions", "Q1 control totals", "10,918 exact"],
                      ["Billing", "Q1 control totals", "Exact to the kobo"], ["PHI scan", "Column review", "None present"]]},
}

V1_FIXES = {
    "currency_slides": [5, 7, 9, 11, 14, 21, 24, 26, 28, 34, 35, 36, 38, 41, 44, 46, 48, 50, 85, 86, 140],
    "currency_unscaled": [85, 86],
    "slides": {
        29: {"Table 14": [["Source Category", "Bootcamp Example", "Strengths & Watch-outs"],
                          ["Internal Operational (POS / EMR / SIS)", "SuperMart_Stores_Sales_2022_2025.csv , Lifecare_Admissions_2023_pipe.txt",
                           "Complete transaction coverage; may contain system codes & nulls."],
                          ["Primary Survey / Field Log", "Crestview_Digital_Wellbeing_Survey.csv (self-rated scores)",
                           "Captures 'why' sentiment; watch out for small sample response bias."],
                          ["Database / System Extract", "lifecare_hospital.db , supermart_online.db (SQLite)",
                           "Keyed, joinable tables; check row counts and keys before analysis."]]},
        31: {"Title 1": "Connecting to Files, Folders and Databases",
             "TextBox 9": "Rather than copying and pasting exports every Monday morning, Excel's Data > Get Data (Power Query, mastered "
                          "in Weeks 2–3) creates refreshable connections to CSV files, folders of monthly files, other workbooks and "
                          "SQLite databases, so Refresh All ( Ctrl+Alt+F5 ) pulls updated records cleanly.",
             "TextBox 13": "SuperMart's analyst connects Excel to the Week2_Data_Files folder; when next month's sales CSV lands there, "
                           "Ctrl+Alt+F5 appends it to the 2025 sales table with zero copy-paste."},
    },
}
