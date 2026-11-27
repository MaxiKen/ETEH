from revlib import T

TRY = "ETEH_Week4_Section_Examples.xlsx  ›  sheet '4.12 Verify AI Output'"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", "Bootcamp datasets")
    return T(tag, **k)


TOPICS = [
    O("4.12.1",
      what="A **hallucination** is AI output that is fluent, confident and **wrong**: a made-up number, column, "
           "function or source. It looks just as convincing as a correct answer.",
      like="Like a friend giving directions to a place they've never been, with total confidence.",
      how=["Expect it: every AI answer can contain one.", "Know the types: numbers, names, functions, sources.", "Check before you use."],
      tip="The more precise a number sounds, the more it needs checking.",
      trap="Believing an answer because it sounds expert.",
      case="Six claims an AI assistant made about our data.",
      steps=["List each claim.", "Recompute it from the data.", "Mark VERIFIED or WRONG."],
      table=[["AI claim", "Real value", "Verdict"], ["Median salary is ₦12.5m", "₦11.0m", "WRONG"],
             ["EkoSuites cancellation rate is 22%", "32.8%", "WRONG"],
             ["Attendance–exam correlation is 0.81", "0.58", "WRONG"],
             ["FY2025 sales were ₦733.2m", "₦733,215,255", "VERIFIED"]],
      so="Wrong and right claims look identical until you check them.",
      split=4.70),

    O("4.12.2",
      what="Check AI output in four steps: **recompute** every number from the data, **reconcile** totals with a "
           "control total, **confirm** every function and column exists, and **test** edge cases (blanks, cancelled rows).",
      like="Like a pilot's pre-flight checklist: the same steps every time, no matter how sure you feel.",
      how=["Put the claim and the recomputed value side by side.", "Set a tolerance (how close is close enough).", "IF(ABS(diff) <= tolerance, \"VERIFIED\", \"WRONG\")."],
      tip="Use a relative tolerance for money (0.1%) and an absolute one for rates (0.5 points).",
      trap="Checking only the claims that look suspicious.",
      case="The tClaims table checks each claim with a formula.",
      steps=["Claimed value in one column.", "Verified value from a formula.", "Verdict formula with tolerance."],
      table=[["Claim", "Claimed", "Verified", "Verdict"], ["Lagos stores = 30% of FY2025 sales", "30%", "29.6%", "VERIFIED (±0.5)"],
             ["Lifecare 2023 readmission 10.3%", "10.3%", "10.34%", "VERIFIED (±0.05)"],
             ["FY2025 sales ₦733.2m", "₦733,200,000", "₦733,215,255", "VERIFIED (0.1%)"],
             ["Median salary ₦12.5m", "₦12,500,000", "₦11,000,000", "WRONG"]],
      so="3 of 6 claims pass. Rounded-but-true claims pass; wrong ones are flagged automatically.",
      split=4.20),

    O("4.12.3",
      what="A **control total** is a figure from an **independent source** (the ledger, the raw table) that any "
           "summary must add up to. If the summary total differs, something was lost or duplicated.",
      like="Like counting the cash in the till against the day's receipts.",
      how=["Take the total from the source data.", "Sum the summary table.", "Difference not 0? Find out why."],
      tip="Calculate the control total from the source, never from the summary itself.",
      trap="Checking a summary against itself.",
      who="SuperMart Retail",
      case="An AI-built zone summary. Does it match the ledger?",
      steps=["Ledger: SUM(tDash25[Sales]).", "Sum the AI zone table.", "Subtract."],
      table=[["Measure", "Value"], ["Control total (ledger)", "₦733,215,255"], ["Sum of the AI zone table", "₦610,309,398"],
             ["Difference", "₦122,905,857"], ["Explanation", "South-East silently dropped"]],
      so="The difference is exactly South-East's sales. Discard the summary and redo it.",
      split=4.90),

    O("4.12.4",
      what="Warning signs of wrong output: **suspiciously round** numbers, parts that **don't add up** to the total, "
           "**unfamiliar** functions or columns, and results that **change** when you ask again.",
      like="Like a fake banknote: it feels slightly wrong if you know what to touch.",
      how=["Look for round numbers.", "Add the parts; compare with the total.", "Ask the same question twice."],
      tip="If the answer changes when you ask again, neither answer is evidence.",
      trap="Ignoring a small gut feeling that a number is odd.",
      case="Spot the warning signs in AI answers about our data.",
      steps=["Read the output.", "Match each sign.", "Recompute the suspect."],
      table=[["Warning sign", "Seen in", "Truth"], ["Round number", "'22% cancellations'", "32.8%"],
             ["Parts do not add up", "3-zone summary", "₦122.9m missing"], ["Too-neat statistic", "'r = 0.81'", "0.58"],
             ["Huge total", "SQL revenue ₦7.19bn", "₦1.39bn"]],
      so="Each warning sign pointed at a real error.",
      split=4.70),

    O("4.12.5",
      what="A **provenance log** records which parts used AI: **when**, **which tool**, **the prompt** (no data), "
           "**the check done**, **the outcome** and **who**. Reviewers can then retrace your steps.",
      like="Like a lab notebook: anyone can repeat the experiment.",
      how=["Add a log table to the workbook.", "One row per AI use.", "Record the check and the outcome."],
      tip="Log discarded output too. It shows the checks are working.",
      trap="Keeping no record, so nobody can tell which numbers came from AI.",
      case="The analyst's provenance log.",
      steps=["Record the time and tool.", "Write the prompt and check.", "Record the outcome."],
      table=[["When", "Prompt", "Check", "Outcome"], ["2024-01-12 10:40", "Zone revenue SQL", "56,725 rows; ₦7.19bn (~5.2x)", "Discarded"],
             ["2024-01-12 11:05", "Explain RANK.EQ ties", "Matched Microsoft docs", "Used in notes"]],
      so="One row shows a caught error, the other a checked explanation.",
      split=4.20),

    O("4.12.6",
      what="**Discard** AI output completely, rather than patching it, when it **invents** columns or functions, "
           "**fails a control total twice**, or **can't explain its own logic**.",
      like="Like a cake with salt instead of sugar: you don't add more icing, you bake again.",
      how=["Check the discard rules first.", "If one applies, delete and restart.", "Write it yourself or re-prompt from the schema."],
      tip="Patching broken output often hides a second error.",
      trap="Fixing the visible error and keeping the rest.",
      who="SuperMart Online",
      case="The zone revenue query: patch or discard?",
      steps=["Check the control total: fails.", "Find the cause: wrong join.", "Discard; rewrite the join yourself."],
      table=[["Condition", "Action"], ["Invents columns or functions", "Discard"], ["Fails control total twice", "Discard"],
             ["Cannot explain its logic", "Discard"], ["Our query: wrong join, 5.2x total", "Discarded; rewrote join"]],
      so="The rewritten query matches the ₦1,394,713,700 control total.",
      split=4.90),

    O("4.12.7",
      what="The analyst who publishes a number **owns it**, whatever tool produced it. 'The AI said so' is never a "
           "defence. Sign-off means you checked it.",
      like="Like a doctor signing a prescription: the pharmacist relies on that signature.",
      how=["Run the checklist.", "Get a reviewer to check.", "Initial and date the sign-off."],
      tip="Ask a reviewer who didn't build the work.",
      trap="Sending work 'for review' that you haven't checked yourself.",
      case="Sign-off block for the AI-assisted claims check.",
      steps=["Complete the checks.", "Reviewer re-runs 2 claims.", "Both initial."],
      table=[["Sign-off item", "Entry"], ["Prepared by", "A. Analyst"], ["Checks", "6 claims recomputed; control total matched"],
             ["Reviewer", "Initials and date (example)"], ["Statement", "I have verified every number in this file"]],
      so="The signature means the numbers were checked, not just produced.",
      split=4.90),

    O("4.12.8",
      what="Before release, tick every item: **schema only** shared, **re-ran** on real data, **control total** "
           "matches, **edge cases** tested, **logged**, **reviewer initials**.",
      like="Like checking your pockets before leaving: phone, keys, wallet.",
      how=["Put the checklist on the sheet.", "Tick each box honestly.", "Release only when all are ticked."],
      tip="Edge cases: blanks, cancelled orders, zero values, ties.",
      trap="Ticking boxes without doing the checks.",
      who="SuperMart Online",
      case="Final checklist for the zone revenue analysis.",
      steps=["Go through each item.", "Write the evidence.", "Sign off."],
      table=[["Check", "Evidence", "Pass?"], ["Schema only", "No data rows pasted", "Yes"],
             ["Re-ran on data", "Query run on the database", "Yes"], ["Control total", "₦1,394,713,700 matches", "Yes"],
             ["Edge cases", "All 11,000 rows; status not filtered", "Yes"], ["Logged and reviewed", "Provenance row + initials", "Yes"]],
      so="Every check passes, so the analysis can be released.",
      split=4.70),
]
