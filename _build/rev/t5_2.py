from revlib import T

TRY = "ETEH_Capstone_Section_Examples.xlsx  ›  sheet 'P.2 Source Register'"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", "Bootcamp datasets")
    return T(tag, **k)


TOPICS = [
    O("P.2.1",
      what="A **source register** lists every dataset you use: file name, **owner**, row count, date registered, "
           "permitted use and a **fingerprint** (a SHA-256 code computed from the file). If the file changes, the "
           "fingerprint changes.",
      like="Like a land registry: every plot has an owner, a size and a reference number.",
      how=["One row per file.", "Record owner, rows and permitted use.", "Add the fingerprint: Get-FileHash file.csv (PowerShell)."],
      tip="Recompute the fingerprint before analysis. A different code means someone changed the RAW file.",
      trap="Using a file nobody can say the source of.",
      case="Part of the bootcamp source register.",
      steps=["List the files.", "Record owner and rows.", "Compute the fingerprint."],
      table=[["File", "Owner", "Rows", "SHA256 (first 16)"], ["SuperMart_Stores_Sales_2022_2025.csv", "Finance", "9,994", "f9a5223dd25cc15b"],
             ["Lifecare_Admissions_2019_2024.csv", "Health Records", "54,860", "afe4541aa292d473"],
             ["Crestview_Student_Performance.csv", "Registry", "6,607", "8b6762a22e5ddc1e"],
             ["EkoSuites_Hotel_Bookings_2023_2024.csv", "Revenue Management", "36,275", "d65b5902a213fc7f"]],
      so="Each file has an owner, a size and a fingerprint, so you can prove it's the file that was registered.",
      split=4.20),

    O("P.2.2",
      what="Keep the **BEFORE** file untouched and produce the **AFTER** table only through named cleaning steps. "
           "A small before/after sample shows reviewers exactly what changed.",
      like="Like before-and-after photos from a mechanic: you can see what was fixed.",
      how=["Store the raw file in 01_RAW, unchanged.", "Clean with named Power Query steps.", "Show a few before/after examples."],
      tip="Anyone should be able to regenerate AFTER from BEFORE by refreshing.",
      trap="Overwriting the original so the changes can't be explained.",
      who="SuperMart Retail",
      case="Before/after sample from SuperMart_Sales_2025_DIRTY.csv.",
      steps=["Pick lines that show each fix.", "Record the value received.", "Record the cleaned value."],
      table=[["Field", "BEFORE (as received)", "AFTER (cleaned)"], ["Sales", "₦239,240 (text)", "239240 (number)"],
             ["Order_Date", "14/05/2025", "2025-05-14"], ["Store_Name", "LAGOS - IKEJA", "Lagos - Ikeja"],
             ["Quantity", "-2", "2"], ["Customer_Name", "N/A", "Blank, flagged for the store"]],
      so="Every change is visible and repeatable. The raw evidence is never lost.",
      split=4.70),

    O("P.2.3",
      what="**Cite** each dataset by name, origin and date. Check the **licence** (who may reuse it). Protect "
           "**privacy**: never publish row-level personal data, and keep sensitive files out of portfolios.",
      like="Like quoting a newspaper in an essay: name the source, and don't publish private letters.",
      how=["Write a full citation for each source.", "Check permitted use with the owner.", "Aggregate or remove personal data."],
      tip="The names in these files are fictional, but treat them as real.",
      trap="Posting a sensitive file in a public portfolio.",
      case="Where each bootcamp dataset came from (lineage).",
      steps=["Name the organisation.", "Name the original source file.", "Note the rule for reuse."],
      table=[["Dataset", "Built from", "Rule"], ["SuperMart Stores (9,994 lines)", "Sample - Superstore.csv", "Training only"],
             ["Lifecare admissions (54,860)", "healthcare_dataset.csv", "No row-level publishing"],
             ["Crestview wellbeing survey (27,870)", "Student Depression Dataset.csv", "Aggregates only, min cell 5"],
             ["EkoSuites bookings (36,275)", "Hotel Reservations.csv", "Training only"]],
      so="Every dataset is cited, and the privacy rule for each is written down.",
      split=4.20),

    O("P.2.4",
      what="Work is **reproducible** when someone else can rerun it and get the same totals. Name every step, "
           "store assumptions in cells, test a refresh, and ask a colleague to follow the README unaided.",
      like="Like a written recipe: anyone following it should get the same jollof.",
      how=["Rename every Power Query step.", "Put assumptions on an Assumptions sheet.", "Have a colleague refresh it on another PC."],
      tip="The second-person test is the strongest proof.",
      trap="Thresholds typed inside formulas where nobody can see them.",
      who="SuperMart Retail",
      case="Reproducibility checks for Capstone 1.",
      steps=["Name the steps.", "Store the assumptions.", "Run the refresh and second-person tests."],
      table=[["Practice", "Example"], ["Named steps", "'Removed spacer rows', 'Fixed ₦ text'"],
             ["Assumptions stored", "FY = calendar year; net sales after discount, excl. VAT"],
             ["Refresh test", "Next month's export → Refresh All → totals still reconcile"],
             ["Second-person test", "Colleague follows README and gets ₦733,215,255"]],
      so="A colleague reproduces the same ₦733,215,255, so the result doesn't depend on you.",
      split=4.90),

    O("P.2.5",
      what="The **handoff pack** links evidence to decision: source register, before/after sample, workbook or SQL, "
           "quality checks, findings with uncertainty, an owner and a follow-up date. Keep **facts** separate from "
           "**recommendations**.",
      like="Like a court case file: evidence first, then the argument, so the judge can check both.",
      how=["Collect the evidence items.", "Write findings with their uncertainty.", "Put recommendations in their own section."],
      tip="A fact can be checked; a recommendation is a judgement. Label which is which.",
      trap="A polished presentation with no evidence behind it.",
      who="SuperMart Retail",
      case="Capstone 1 handoff pack.",
      steps=["Assemble the pack.", "Separate facts and recommendations.", "Name the owner and follow-up date."],
      table=[["Item", "Example"], ["Evidence", "Register + before/after + workbook + QA checks"],
             ["Fact", "Q4 2025 margin fell to 9.8%"], ["Recommendation", "Cap Furniture discounts at 20%"],
             ["Owner / follow-up", "Commercial Director; review monthly"]],
      so="A reviewer can follow the chain from source file to decision.",
      split=4.90),
]
