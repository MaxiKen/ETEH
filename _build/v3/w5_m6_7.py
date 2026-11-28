"""Capstone edition-3 slide copy: Modules P.6 (workbook QA and stress testing) and P.7 (the ten-minute executive defence).
Figures from w5_fx."""
from slidekit import NAVY, ORANGE
from v3kit import T
from w5_fx import (BLANK, CLEAN, CTRL, DEEP, DIRTY, DM, DUPS, G25, LINES, LOSS_ST, M24, M25, MIXE, N_CAP, ORD25, RATE, RECON,
                   S25, SMALL_STORE, WHATIF, I, N, Nm, P, pts)

MINS = [("Question and headline", 2), ("Evidence and live demo", 4), ("Recommendation", 2), ("Limitations and next steps", 2)]


def _v(x):
    return N(x) if x > 100000 else I(x)


M_P6 = dict(
    id="P.6", title="Workbook QA and Stress Testing",
    subtitle="Proving a workbook is correct, fast, portable and maintainable before anyone else opens it.",
    objectives=["Explain what workbook QA protects against.", "Write a 00_README to a standard.",
                "Build live reconciliation tests.", "Stress-test size, structure, refresh and edge cases.",
                "Measure performance and sign off a frozen release."],
    topics=[
        T("P.6.1", "The Purpose of Workbook QA",
          defn="Workbook QA is the set of checks that proves a workbook gives the right numbers, keeps giving them after a "
               "refresh, and can be run by someone else.",
          like="Like a pilot's pre-flight checklist: the plane usually flies fine, but nobody skips the list.",
          points=["Correct: totals reconcile to the source", "Robust: survives new, bigger and messier data",
                  "Portable: opens and refreshes on another machine", "Maintainable: a README explains it"],
          points_label="FOUR THINGS QA PROVES",
          watch="An error found by the board costs your credibility; the same error found in QA costs ten minutes."),

        T("P.6.2", "The 00_README Standard",
          defn="The 00_README sheet is the first sheet of every deliverable: purpose, audience, sources, refresh steps, "
               "measures, exclusions, limitations, version and contact.",
          grid=dict(label="00_README, Capstone 1", widths=[1.6, 4.3], size=12, rows=[
              ["Section", "Content"], ["Purpose", "Explain the FY2025 margin decline; support the Q1 2026 pricing decision"],
              ["Sources and owners", "POS ledger (Finance); store directory (Operations); see the P.2 register"],
              ["Refresh", "New CSV in 01_RAW › Data › Refresh All › reconciliation cells = 0"],
              ["Exclusions", f"None from FY totals; {BLANK} blank and {DUPS} duplicate rows removed from the returns"],
              ["Version and contact", "v1.0, 2026-01-20, analyst@supermart.ng"]])),

        T("P.6.3", "Reconciliation Testing",
          defn="A reconciliation test compares a workbook value with an independent control total; the difference must be 0.",
          fx=[("=[@Workbook_value]-[@Control_total]", "0", "Difference per test"),
              ('=IF(ABS([@Difference])<0.5,"PASS","FAIL")', "PASS", "Verdict")],
          fx_widths=[3.6, 0.8, 1.6],
          grid=dict(label="tRecon: seven live tests", widths=[2.4, 1.6, 0.7], size=11.5, rows=[
              ["Test", "Workbook = control", "Result"]] + [[t, _v(a), "PASS" if abs(a - b) < .5 else "FAIL"] for t, a, b in RECON]),
          order=["grid", "fx"], stack=True),

        T("P.6.4", "Stress Testing with Larger and Messier Data",
          defn="A stress test runs the workbook on bigger or messier data than it was built on, to find where it breaks.",
          grid=dict(label="Stress tests and pass criteria", widths=[1.6, 2.6, 2.0], size=12, rows=[
              ["Test", "How", "Pass criterion"],
              ["10× rows", f"Append the {I(LINES)} lines 10 times", "Refresh under 60 s"],
              ["Extra column", "Insert 'Store_Location' mid-file", "Queries use names, not positions"],
              ["Empty month", "Delete all December rows", "No #DIV/0!; cards show 0 cleanly"],
              ["Dirty rows", "Re-run the dirty returns", f"Still {I(CLEAN)} rows after cleaning"]])),

        T("P.6.5", "Testing Refreshes and Dependencies",
          defn="Test the whole refresh chain from a clean start: source file › Power Query › Data Model › dashboard.",
          steps=["Close the file and start Excel fresh", "Replace the RAW file with next month's export", "Data › Refresh All",
                 "Check row counts and every reconciliation cell"],
          steps_label="REFRESH TEST",
          grid=dict(label="Failure signals", widths=[1.8, 2.6], size=12, rows=[
              ["Symptom", "Likely cause"], ["Same number every month", "A pasted value, not a measure"],
              ["Row count unchanged", "Query points at the old file"], ["#REF! after refresh", "A column renamed in the source"]]),
          order=["steps", "grid"], split=0.48),

        T("P.6.6", "Error and Edge Case Audit",
          defn="An edge-case audit hunts for errors, blanks in required fields and possible division by zero, then tests the "
               "extremes: the smallest group, an empty period, a zero denominator.",
          fx=[("=COUNTBLANK(tSM[Sales])", "0", "Blanks in a required field"),
              ("=SUMPRODUCT(--ISERROR(tSM[Sales]))", "0", "Error values"),
              ('=MIN(COUNTIFS(tSM[Year],2025,tSM[Store_ID],{"ST01",…}))', I(SMALL_STORE.min()), "Smallest store's FY2025 lines")],
          fx_widths=[3.9, 0.7, 2.0], stack=True,
          watch=f"Check the smallest group ({SMALL_STORE.idxmin()}) renders in every visual before the board filters to it."),

        T("P.6.7", "Performance Testing",
          defn="Performance testing times the refresh, checks file size and visual response, against targets set before testing.",
          grid=dict(label="Targets for a monthly workbook", widths=[1.4, 1.3, 3.0], size=12, rows=[
              ["Metric", "Target", "Remedy if missed"], ["Refresh", "Under 60 s", "Group by month and store in Power Query"],
              ["File size", "Under 30 MB", "Load to the Data Model only"], ["Slicer response", "Under 2 s", "Measures, not calculated columns"],
              ["Formulas", "No volatile functions", "Replace OFFSET/INDIRECT with INDEX"]])),

        T("P.6.8", "Release Sign-Off",
          defn="Release sign-off records who tested, what was tested, the control totals used, the version and the date. The "
               "signed file is frozen.",
          grid=dict(label="Capstone 1 sign-off", widths=[1.6, 4.3], size=12.5, rows=[
              ["Item", "Record"], ["Tested by", "A peer analyst, not the builder"],
              ["What was tested", "Reconciliation (all PASS), stress, refresh, slicers, second machine"],
              ["Control totals", f"{N(S25)} sales; {I(LINES)} lines; {I(CTRL['Distinct orders'])} FY2025 orders"],
              ["Version", "v1.0, frozen and archived read-only; changes go into v1.1"]])),
    ],
    takeaways=["QA proves correct, robust, portable and maintainable.", "Every deliverable opens on a 00_README.",
               "Reconciliation tests are live formulas that must show PASS.", "Stress, refresh and edge cases before release.",
               "Sign off, freeze and version the file."],
)

M_P7 = dict(
    id="P.7", title="The Ten-Minute Executive Defence",
    subtitle="Presenting and defending an analytical build to a panel: structure, evidence, question handling and the "
             "demonstration.",
    objectives=["Know what the panel assesses.", "Structure ten minutes and a short live demo.", "Prepare answers to the predictable questions.",
                "Handle challenge and translate technical choices.", "Close for a decision and reflect afterwards."],
    topics=[
        T("P.7.1", "What the Defence Assesses",
          defn="The defence checks that you understand your own work: the definitions, the decisions, the limitations and the "
               "evidence.",
          grid=dict(label="Four areas, Capstone 1 examples", widths=[1.2, 2.2, 2.6], size=12, rows=[
              ["Area", "Likely question", "Evidence"], ["Definition", "How is margin defined?", "Profit ÷ net sales after discount"],
              ["Decision", "Is it price or mix?", f"Rate {pts(RATE)} vs mix {pts(MIXE)}"],
              ["Limitation", "What can't this tell us?", "No competitor prices; one year compared"],
              ["Evidence", "Does it reconcile?", f"{N(S25)} = ledger, difference ₦0"]])),

        T("P.7.2", "Structuring the Ten Minutes",
          defn="Plan the ten minutes in four segments and rehearse to the clock.",
          chart=dict(kind="bar", cats=[m for m, _ in MINS], series=[("Minutes", [v for _, v in MINS])], fmt="0", labels=True,
                     legend=False, point_colors=[NAVY, ORANGE, NAVY, NAVY], title="Ten minutes, four segments", h=2.4),
          grid=dict(label="What goes in each", widths=[1.8, 3.0], size=12, rows=[
              ["Segment", "Capstone 1 content"], ["Headline", f"Margin fell {-DM:.1f} points while sales grew {P(G25, 0)}"],
              ["Evidence", "Slicer › all visuals; exception list; reconciliation"],
              ["Recommendation", "Cap Furniture discounts at 20%; owner, date, effect"],
              ["Limits", "Calendar-year view; targets not size-adjusted"]]),
          order=["chart", "grid"], split=0.45),

        T("P.7.3", "The Live Demonstration",
          defn="Demonstrate three interactions only, each proving one thing, and rehearse them until they take 90 seconds.",
          grid=dict(label="Three demo steps", widths=[2.0, 2.6, 0.7], size=12.5, rows=[
              ["Step", "What it proves", "Time"], ["Slicer on one zone", "Every visual is connected", "30 s"],
              ["Loss-making stores list", f"The exceptions are real ({len(LOSS_ST)} stores)", "30 s"],
              ["Reconciliation cell", "The total equals the ledger", "30 s"]]),
          watch="Never demo something you have not rehearsed on the presentation laptop."),

        T("P.7.4", "Anticipating Panel Questions",
          defn="Most panel questions are predictable: source, definition, exclusions, reconciliation, sample size, assumptions, "
               "limitations, what you'd change, refresh and ownership.",
          grid=dict(label="Prepared answers", widths=[2.0, 3.9], size=11.5, rows=[
              ["Question", "Prepared answer"], ["Source and date?", "POS ledger export, 05-Jan-2026, Finance; fingerprint registered"],
              ["What did you exclude?", f"Nothing from FY totals; {BLANK} blank and {DUPS} duplicate rows from the returns"],
              ["Is the sample big enough?", f"All {I(CLEAN)} FY2025 lines and {I(ORD25)} orders: a census, not a sample"],
              ["What is the assumption?", f"Volume lost on re-priced lines; the gain stays {Nm(WHATIF[0.5])}-{Nm(WHATIF[0.0])} from 0-50%"],
              ["Biggest limitation?", "One year of comparison; no competitor pricing"]])),

        T("P.7.5", "Handling Challenge and Disagreement",
          defn="When challenged: restate the question, answer with evidence, and admit what is genuinely uncertain.",
          steps=["Restate: 'You think the margin fall is a mix effect?'",
                 f"Evidence: 'The bridge shows rate {pts(RATE)} and mix {pts(MIXE)}'",
                 "Admit: 'Competitor pricing is outside our data'", "Offer: 'I can test it with next quarter's prices'"],
          steps_label="FOUR MOVES",
          watch="Defensiveness reads as weakness. 'Let me show you' beats 'You're wrong'."),

        T("P.7.6", "Explaining Technical Choices to Non-Technical Panels",
          defn="Translate every technical choice into a business consequence.",
          grid=dict(label="Technical › business", widths=[2.4, 3.5], size=12.5, rows=[
              ["Technical statement", "Business translation"],
              ["Star schema with DAX measures", "New months drop in and every number updates itself"],
              ["Rate/mix decomposition", "The fall came from deeper discounts, not from selling different things"],
              ["Reconciliation to control totals", "Finance's numbers and ours are the same, to the naira"],
              ["Sales-weighted discount depth", "Big orders count more than small ones, as they do in the P&L"]])),

        T("P.7.7", "Closing for a Decision",
          defn="End with a specific ask: what to approve, who decides, by when, and the expected effect.",
          grid=dict(label="The Capstone 1 close", widths=[1.4, 4.5], size=12.5, rows=[
              ["Element", "Content"], ["Decision", "Cap Furniture discounts at 20% from Q1 FY2026"],
              ["Owner", "Commercial Director"], ["By when", "Today, so January prices can change"],
              ["Expected effect", f"About {Nm(WHATIF[0.25])} extra profit a year ({I(N_CAP)} lines re-priced), tracked monthly"]]),
          watch="'Any questions?' is not a close. Ask for the decision."),

        T("P.7.8", "Post-Defence Reflection",
          defn="Within a day, write down what the panel probed, what you could not answer well, and what you will change.",
          grid=dict(label="Reflection note", widths=[1.8, 4.1], size=12.5, rows=[
              ["Item", "Example"], ["Panel probed", "The definition of discount depth"],
              ["Couldn't answer well", "Why it is weighted by sales"],
              ["Will change", "Add the weighting to the README and one slide"],
              ["Deep-discount loss", f"Have the figure ready: {Nm(-DEEP)} on lines above 20%"]])),
    ],
    takeaways=["The panel tests definitions, decisions, limitations and evidence.", "Two, four, two, two: rehearse to the clock.",
               "Prepare the ten predictable answers with numbers.", "Restate, evidence, admit; translate to business terms.",
               "Close with a decision, an owner and a date."],
)

MODULES = [M_P6, M_P7]
