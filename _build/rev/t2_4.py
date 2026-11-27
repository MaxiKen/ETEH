from revlib import T

HTML = "SuperMart_Store_Directory.html"
PDF = "SuperMart_Monthly_Sales_Bulletin_2025.pdf"

TOPICS = [
    T("2.4.1",
      what="A web page can hold data tables. **Data › From Web** reads a page (online, or a saved .html file) and "
           "the **Navigator** lists every table it finds. Before importing, **read the page in a browser**: check "
           "which tables exist, what one row means, the units, and any footnotes.",
      like="It is like reading a menu before ordering. You check what each dish is, its size and its price, before you "
           "tell the waiter what you want.",
      how=["Open the page in a browser and read it fully.",
           "Data › From Web › paste the address (or file:/// path).",
           "Pick the right table in the Navigator › Transform Data."],
      tip="Write down the page, the table chosen, the date and the owner.",
      trap="Ignoring footnotes; they often change what a number means.",
      who="SuperMart Retail",
      case="Head Office publishes its store directory as a web page, " + HTML + ". The analyst wants the 12 stores' "
           "floor areas.",
      steps=["Open the file in a browser; note the 2 tables and the footnote.",
             "Data › From Web › file path › choose Table 1 (stores).",
             "Set Opened to Date using Locale English (United Kingdom)."],
      table=[["Check", "Result"],
             ["Tables on the page", "2: stores, opening hours"],
             ["Store table size", "12 rows × 8 columns"],
             ["Largest / smallest store", "Ikeja 2,400 m² / Jos 1,050 m²"],
             ["Footnote", "Floor area excludes car parks"]],
      split=4.90,
      so="Reading the page first revealed the footnote. Sales per m² must say 'excluding car parks'."),

    T("2.4.2",
      what="**Get Data › From File › From PDF** finds tables inside a PDF and lists them in the Navigator as "
           "**Table001, Table002…** (and whole pages as Page001…). PDFs are made for reading, not analysis, so always "
           "**compare a few extracted numbers with the printed page** before trusting them.",
      like="It is like copying notes from a friend's exercise book. Afterwards you check a few lines against the "
           "original to make sure you copied correctly.",
      how=["Get Data › From File › From PDF › choose the file.",
           "Click each Table in the Navigator and compare with the PDF.",
           "Spot-check three numbers against the printed page."],
      tip="Spot-check at least three extracted numbers against the page.",
      trap="Loading the first table found without checking; PDFs often contain several tables.",
      who="SuperMart Retail",
      case="Finance sends " + PDF + ": a table of store sales by month in ₦ millions. The analyst extracts it.",
      steps=["From PDF › the bulletin › click Table001 and Table002.",
             "Choose Table001 (12 stores × 12 months plus FY).",
             "Check three cells against the printed page."],
      table=[["Spot-check cell", "Printed page", "Extracted"],
             ["Lagos - Ikeja, Dec", "19.3", "19.3"],
             ["Port Harcourt - GRA, Oct", "24.7", "24.7"],
             ["Jos - Rayfield, FY 2025", "66.2", "66.2"]],
      split=4.70,
      so="All three checks match, so the extraction can be trusted for the next steps."),

    T("2.4.3",
      what="PDF and report tables are usually **cross-tabs**: one column per month. That is easy to read but hard "
           "to analyse. **Unpivot** turns the month columns into rows: Store | Month | Sales. Use **Unpivot Other "
           "Columns** on Store, so any new month column is picked up automatically next time.",
      like="It is like taking a class timetable grid and writing it as a list: 'Monday, 8am, Maths', one line per "
           "lesson. The list is longer, but you can sort and count it easily.",
      how=["Promote headers; filter out the Total row.",
           "Remove the FY total column.",
           "Select Store › Transform › Unpivot Other Columns."],
      tip="Remove total rows and total columns before unpivoting.",
      trap="Leaving totals in, so every PivotTable counts them twice.",
      who="SuperMart Retail",
      case="Table001 from the bulletin has stores down the side and months across the top. The analyst needs it "
           "tidy for a PivotTable.",
      steps=["Remove the Total row and the FY_2025 column.",
             "Select Store › Unpivot Other Columns.",
             "Rename to Month and Sales_m; set Sales_m to Decimal."],
      table=[["Stage", "Rows × columns"],
             ["As extracted (with FY column)", "12 × 14"],
             ["After removing FY_2025", "12 × 13"],
             ["After Unpivot Other Columns", "144 × 3"],
             ["Columns now", "Store | Month | Sales_m"]],
      split=4.90,
      so="144 rows (12 stores × 12 months) can be filtered, sorted and pivoted any way. The cross-tab could not."),

    T("2.4.4",
      what="Data copied from a PDF or report must be **reconciled**: compared with the **system of record** (the "
           "ledger). Merge the two on a shared key, round to the same precision, and flag any difference. Never "
           "rebuild totals by adding rounded figures; small rounding errors add up.",
      like="It is like checking your bank alert against your passbook. If they agree, you trust both. If not, you "
           "find out why before spending.",
      how=["Merge the PDF table with the ledger on Store_Name.",
           "Round the ledger to ₦0.1m; subtract to find the difference.",
           "Flag any store with a difference above ₦0.1m."],
      tip="Treat the ledger as the truth; the PDF is only a copy.",
      trap="Adding up rounded monthly figures and calling it the total.",
      who="SuperMart Retail",
      case="Is the bulletin trustworthy? The analyst compares each store's FY figure with the ledger "
           "(SuperMart_Stores_Sales_2022_2025.csv).",
      steps=["Merge Queries on Store_Name: all 12 stores match.",
             "Compare each PDF FY figure with the rounded ledger total.",
             "Count exceptions, then check the sum of the 144 rounded cells."],
      table=[["Check", "PDF", "Ledger"],
             ["Lagos - Ikeja, FY 2025", "₦146.4m", "₦146.4m (OK)"],
             ["Owerri - New Owerri, FY 2025", "₦31.6m", "₦31.6m (OK)"],
             ["Stores that differ (of 12)", "0", "0"],
             ["Sum of 144 rounded monthly cells", "₦732.9m", "₦733.2m"]],
      split=4.70,
      so="Every store agrees, so the PDF is reliable. But adding its rounded cells loses ₦0.3m; use ledger totals."),

    T("2.4.5",
      what="An **audit-ready refresh** means anyone can rerun the import and prove it is correct. Keep the source "
           "name and table ID in a Source column, add **control checks** (row count and a total check), and build "
           "steps that survive layout changes, such as **Unpivot Other Columns**.",
      like="It is like a pilot's checklist before take-off. The same checks, ticked every time, by whoever is flying.",
      how=["Add a Source column: file name and table ID (Table001).",
           "Create Row_Count and Total_Check queries; load to a 99_Checks sheet.",
           "After each refresh, check both before publishing."],
      tip="Log every refresh with the date, the checks and your initials.",
      trap="Assuming a clean-looking result means the extraction was correct.",
      who="SuperMart Retail",
      case="Next month a new bulletin replaces the old one in the folder. The analyst refreshes and runs the "
           "controls before publishing.",
      steps=["Replace the PDF in the folder (same file name).",
             "Data › Refresh All.",
             "Read the 99_Checks sheet and log the result."],
      table=[["Control", "Expected", "Why"],
             ["Row_Count", "144", "12 stores × 12 months"],
             ["Total_Check", "Within ₦0.1m of ledger", "PDF agrees with the record"],
             ["Source column", "Bulletin file + Table001", "Every row is traceable"]],
      split=4.70,
      so="With controls in place, a refresh is proven, not assumed, and another analyst can repeat it exactly."),
]
