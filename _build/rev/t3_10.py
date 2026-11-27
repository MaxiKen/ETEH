from revlib import T

TRY = "ETEH_Week3_Section_Examples.xlsx  ›  sheet '3.10 Files to SQL & DAX'  ·  lifecare_hospital.db"
DIRTY = "Lifecare_Admissions_2023_DIRTY.csv"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", "Lifecare Hospital")
    return T(tag, **k)


TOPICS = [
    O("3.10.1",
      what="Before loading any file into a database, **profile** it: for every column count the rows, blanks and "
           "distinct values. Problems found now are cheap to fix; problems found in a dashboard are expensive.",
      like="Profiling is like checking a delivery at the door before you sign for it.",
      how=["Load the file into Power Query (or pandas) without changing anything.",
           "Count rows, blanks and distinct values per column.", "Write down a fix for every problem."],
      tip="A column with more distinct values than expected usually has spelling variants.",
      trap="Loading straight into the database and discovering problems later.",
      case=DIRTY + " (Week2_Data_Files) must be loaded into SQLite. The analyst profiles it first.",
      steps=["Turn on Column profile for all 15 columns.", "Note blanks and distinct counts.", "List the fix for each issue."],
      table=[["Column", "Blank", "Distinct", "Action"], ["Discharge_Date", "70", "392", "Flag; don't guess dates"],
             ["HMO", "40", "5 + blank", "Label 'Unknown'"], ["Ward", "0", "16", "Clean to 4 wards"],
             ["Gender", "0", "6", "Map to Female / Male"], ["All columns", "11,026 rows", "15 columns", "Record as control"]],
      so="11,026 rows arrived with blanks and spelling variants. Each one now has a planned fix.",
      split=4.20),

    O("3.10.2",
      what="Load the file into a **staging table** first: all columns as **TEXT**, exactly as received. Then check "
           "the row count matches the file. Cleaning happens in the next step, not during the import.",
      like="Staging is like a hospital's reception area: everyone is booked in exactly as they arrive, then sent "
           "to the right ward.",
      how=["DB Browser: File > Import > Table from CSV > stg_admissions_2023.",
           "Keep every column as TEXT.", "SELECT COUNT(*) and compare with the file."],
      tip="Name staging tables stg_… so nobody reports from them.",
      trap="Converting text to numbers during import: '2,507,101.35' becomes 2.",
      case="The analyst imports " + DIRTY + " into SQLite as stg_admissions_2023.",
      steps=["Import as TEXT into stg_admissions_2023.", "Compare COUNT(*) with the file's rows.",
             "Check Billing_Amount values that contain commas."],
      table=[["Check", "Result"], ["CSV data rows", "11,026"], ["Staging rows", "11,026 (match)"],
             ["Bills written with commas", "441, e.g. '2,507,101.35'"], ["CAST('2,507,101.35' AS REAL)", "2 (wrong); clean first"]],
      so="All rows landed. Keeping text exposed 441 comma-formatted bills that a direct number import would ruin.",
      split=4.90),

    O("3.10.3",
      what="Use SQL to find problems: a **LEFT JOIN … WHERE IS NULL** finds orphans (e.g. admissions with an "
           "unknown doctor), and a **UNION ALL** report counts rows breaking each business rule.",
      like="An exception report is like a teacher's list of missing homework: one line per problem, with a count.",
      how=["LEFT JOIN to the lookup table; keep rows where the lookup key IS NULL.",
           "Write one SELECT per rule with COUNT(*).", "Stack them with UNION ALL."],
      tip="Run the same report on staging and on production.",
      trap="Using an INNER JOIN, which silently drops the orphans you're looking for.",
      case="The analyst runs the exception report on stg_admissions_2023 and on the production admissions table.",
      steps=["Orphan check: Doctor_Name not in doctors.", "Rule checks: dates, bills, length of stay.",
             "Compare staging with production."],
      table=[["Rule", "Staging (dirty)", "Production"], ["Doctor not in doctors table", "0", "0"],
             ["Discharge before admission", "25", "0"], ["Bill below zero", "18", "0"], ["Exact duplicate rows", "63", "0"]],
      so="The extract breaks three rules; production breaks none. These rows must be fixed before loading.",
      split=4.20),

    O("3.10.4",
      what="On the hospital model, define **explicit measures**: Admissions = COUNTROWS, ALOS = AVERAGE of "
           "length_of_stay, Readmission % with DIVIDE and CALCULATE. Then filters (department, admission type) "
           "change the answers correctly.",
      like="Explicit measures are like agreed recipes: every kitchen (PivotTable) makes the dish the same way.",
      how=["Admissions := COUNTROWS ( Admissions ).", "ALOS := AVERAGE ( Admissions[length_of_stay] ).",
           "Readmission % := DIVIDE ( CALCULATE ( [Admissions], readmit_30d = \"Yes\" ), [Admissions] )."],
      tip="Put admission_type and department on Rows to test the filter context.",
      trap="Using implicit 'Average of readmit' style fields.",
      case="The quality team wants readmission rates overall and for emergency patients.",
      steps=["Create the three measures.", "Read the grand totals.", "Rows: admission_type; read Emergency."],
      table=[["Measure / filter", "Result (2019–2024)"], ["Admissions", "54,860"], ["ALOS", "15.5 days"],
             ["Readmission %", "10.4%"], ["Readmission %, Emergency", "13.7%"], ["Readmission %, Oncology", "13.5%"]],
      so="Emergency and Oncology patients come back more often. The measure answered both through filter context.",
      split=4.90),

    O("3.10.5",
      what="Before publishing, **reconcile** three numbers under the same filters: the **source file**, the "
           "**SQL** count and sum, and the **dashboard card**. Publish only when all three agree.",
      like="It's like three people counting the same cash box: only when all three agree do you bank it.",
      how=["Count rows and sum billing in the source CSV.", "Run COUNT(*) and SUM(billing_amount) in SQL.",
           "Compare with the dashboard cards; record PASS or FAIL."],
      tip="Keep the reconciliation table on the dashboard workbook's QA sheet.",
      trap="Publishing because the numbers 'look about right'.",
      case="The analyst reconciles the Lifecare dashboard against Lifecare_Admissions_2019_2024.csv and lifecare_hospital.db.",
      steps=["Count and sum the source CSV.", "Count and sum in SQL.", "Compare with the dashboard cards."],
      table=[["Control", "Source CSV", "SQL", "Result"], ["Admissions", "54,860", "54,860", "PASS"],
             ["Billing (₦)", "70,206,080,066.78", "70,206,080,066.78", "PASS"]],
      so="Source and SQL agree to the kobo. Once the dashboard cards show the same two numbers, it can be published.",
      split=4.20),
]
