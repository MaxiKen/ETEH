from revlib import T

TRY = "ETEH_Capstone_Section_Examples.xlsx  ›  sheet 'P.6 Workbook QA'"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", "SuperMart Retail · Capstone 1 QA")
    return T(tag, **k)


TOPICS = [
    O("P.6.1",
      what="**Quality assurance (QA)** proves three things: the numbers are **right**, the workbook **behaves "
           "predictably**, and **someone else can maintain it**. It's a separate stage from building, and "
           "ideally a different person does it.",
      like="Like a building inspector: the builder says it's safe; the inspector checks before anyone moves in.",
      how=["Check the numbers against controls.", "Try to break the workbook.", "Have someone else follow the README."],
      tip="Ask a peer to test it: the builder is the worst person to spot their own mistakes.",
      trap="Treating 'it looks right' as a test.",
      case="The three QA questions applied to Capstone 1.",
      steps=["Run the reconciliation table.", "Run the stress tests.", "Hand the README to a peer."],
      table=[["QA question", "Test", "Capstone 1 evidence"], ["Are the numbers right?", "Reconciliation", "7 tests, all PASS"],
             ["Does it behave?", "Stress and edge cases", "Survives 10x rows, extra column, empty month"],
             ["Can others maintain it?", "README handover", "Peer refreshes with no help"]],
      so="QA gives evidence, not opinion, that the workbook is ready.",
      split=4.20),

    O("P.6.2",
      what="Every workbook opens with a **00_README** sheet: purpose, audience, sources and owners, refresh steps, "
           "measures, exclusions, limitations, and version/date/contact. **No README means the workbook isn't finished.**",
      like="Like the manual that comes with a generator: without it, only the person who installed it can run it.",
      how=["Fill in all eight sections.", "Write the refresh steps as numbered actions.", "State the limitations honestly."],
      tip="Test it: can a colleague refresh the file using only the README?",
      trap="'Limitations: none.' Every analysis has some.",
      case="The Capstone 1 README.",
      steps=["Write each section.", "Ask a peer to follow it.", "Fix any gap they hit."],
      table=[["Section", "Content"], ["Purpose", "Explain FY2025 margin decline; support Q1 2026 pricing"],
             ["Sources and owners", "POS ledger (Finance); store directory (Operations)"],
             ["Refresh", "New CSV in 01_RAW → Refresh All → reconciliation = 0"],
             ["Limitations", "Calendar-year view; targets not adjusted for store size"],
             ["Version / contact", "v1.0, 2026-01-20, analyst@supermart.ng"]],
      so="A colleague can take over the workbook without a single phone call.",
      split=4.90),

    O("P.6.3",
      what="**Reconciliation testing**: compare each headline number with an independent **control total** and show "
           "the difference as a live formula. If there's a difference, explain it. Never tweak the model until it goes away.",
      like="Like balancing a cash book against the bank statement: every naira is accounted for.",
      how=["List the controls and their owners.", "Difference = workbook − control.", "PASS only if the difference is 0."],
      tip="Use live formulas, so the test re-runs after every refresh.",
      trap="Deleting rows or typing an adjustment to force a match.",
      case="The live reconciliation table on the P.6 sheet.",
      steps=["Type each control total.", "Formula for the workbook value.", "IF(ABS(diff) < 0.5, \"PASS\", \"FAIL\")."],
      table=[["Test", "Control", "Workbook", "Result"], ["SuperMart lines", "9,994", "9,994", "PASS"],
             ["FY2025 sales", "₦733,215,255", "₦733,215,255", "PASS"], ["Crestview cohort", "6,607", "6,607", "PASS"],
             ["Lifecare 2023 admissions", "10,918", "10,918", "PASS"], ["Bed census days", "1,797", "1,797", "PASS"]],
      so="All 7 tests pass with a difference of 0. That's the evidence the panel will ask for.",
      split=4.20),

    O("P.6.4",
      what="**Stress testing** feeds the workbook bigger and messier data on purpose: **10 times the rows**, an "
           "**extra column** in the middle, and an **empty month**. If it survives all three, it's ready for real use.",
      like="Like testing a bridge with heavier lorries than it will normally carry.",
      how=["Append the data to itself 10 times.", "Insert a column in the raw file.", "Delete one month of rows."],
      tip="The extra-column test catches formulas that use column positions instead of names.",
      trap="Testing only on the small, clean sample you built with.",
      case="Capstone 1 stress tests.",
      steps=["Make each distorted copy.", "Refresh.", "Check the pass criterion."],
      table=[["Stress test", "How", "Pass criterion"], ["10x rows", "Append the 9,994 lines 10 times", "Refresh under 60 s"],
             ["Extra column", "Insert 'Store_Location' in the RAW CSV", "Queries use names, not positions"],
             ["Empty month", "Remove all December rows", "No #DIV/0!; cards show blank"]],
      so="A workbook that passes all three won't break the first time real data misbehaves.",
      split=4.20),

    O("P.6.5",
      what="Test the **whole refresh chain** from a clean start: source file → Power Query → Data Model → dashboard. "
           "Row counts must match, measures must recalculate, and **nothing may depend on a pasted value**.",
      like="Like a fire drill: you only know the exits work when everyone walks through them.",
      how=["Close and reopen the file.", "Drop in the new source and Refresh All.", "Check row counts and reconciliation."],
      tip="A refresh that needs a manual step isn't a refresh. Write that step into a query instead.",
      trap="A KPI typed in by hand that never changes.",
      case="Capstone 1 refresh test.",
      steps=["Reopen from closed.", "Refresh All.", "Compare with expected values."],
      table=[["Refresh check", "Pass condition", "Failure signal"], ["Row count", "9,994 lines loaded", "Count unchanged after new data"],
             ["Measures", "Recalculate after refresh", "Same number every month"],
             ["Reconciliation", "All 7 tests still PASS", "Any FAIL or #REF!"]],
      so="If the refresh runs without help and still passes, the workbook can be handed over.",
      split=4.20),

    O("P.6.6",
      what="An **edge-case audit** hunts for errors, blanks in required fields and possible division by zero. "
           "Then it tests the extremes: the **smallest group**, an **empty filter**, a **missing month**.",
      like="Like checking a school timetable for the class that has only one student.",
      how=["Ctrl+G › Special › Formulas › Errors.", "COUNTBLANK on required fields.", "Test the smallest group and an empty filter."],
      tip="Wrap ratios in DIVIDE (DAX) or IFERROR so an empty group shows blank, not an error.",
      trap="Only testing the 'happy path' with every filter cleared.",
      case="Edge checks on the P.6 sheet.",
      steps=["Count blank sales.", "Count error values.", "Find the smallest store."],
      table=[["Check", "Formula", "Result"], ["Blank Sales cells", "COUNTBLANK(tSM[Sales])", "0"],
             ["Error values in Sales", "SUMPRODUCT(--ISERROR(…))", "0"],
             ["Smallest store, FY2025 lines", "MIN(COUNTIFS(…))", "134"], ["Any empty quarter in 2025?", "MIN(COUNTIFS(…)) = 0", "No"]],
      so="No errors or blanks, and the smallest store still has 134 lines: every visual has data to show.",
      split=4.20),

    O("P.6.7",
      what="**Performance testing** times the refresh, checks the file size and how fast visuals respond. Typical "
           "targets for a monthly workbook: **refresh under 60 seconds**, **file under 30 MB**.",
      like="Like timing a delivery route: if it takes too long, change the route, not the driver's speed.",
      how=["Time Refresh All.", "Check the file size.", "If slow, summarise earlier in Power Query."],
      tip="Load detail rows to the Data Model only, not to worksheets.",
      trap="Polishing formulas when the real problem is too many rows.",
      case="Capstone 1 performance check.",
      steps=["Time the refresh at normal size.", "Repeat with 10x rows.", "Compare with the targets."],
      table=[["Metric", "Target", "Remedy if missed"], ["Refresh time", "Under 60 s", "Group by month/store in Power Query"],
             ["File size", "Under 30 MB", "Load to Data Model only"], ["10x stress (about 100k rows)", "Still under 60 s", "Remove unused columns early"]],
      so="Measure first, then fix the biggest cause. Usually that means fewer rows, earlier.",
      split=4.20),

    O("P.6.8",
      what="**Release sign-off** records who tested, what was tested, the control totals used, the version and the "
           "date. Once signed, the file is **frozen** and archived read-only.",
      like="Like a signed delivery note: proof of what left the warehouse, when, and who checked it.",
      how=["A peer (not the builder) tests.", "Record the results and controls.", "Freeze and archive the version."],
      tip="Keep the previous version, so you can roll back if needed.",
      trap="Overwriting the approved file with 'final_final_v2.xlsx'.",
      case="The Capstone 1 sign-off record.",
      steps=["Complete each field.", "Save as read-only.", "Archive it."],
      table=[["Sign-off item", "Record"], ["Tested by", "Peer analyst (not the builder)"],
             ["What was tested", "Reconciliation (all PASS), stress, slicers, second machine"],
             ["Control totals", "₦733,215,255 sales; 9,994 lines"], ["Version", "v1.0, frozen, read-only"], ["Date", "2026-01-20"]],
      so="Anyone can see exactly which version was approved, and check it.",
      split=4.90),
]
