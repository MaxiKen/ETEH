"""Capstone edition-3 slide copy: Modules P.1 (the seven-stage project lifecycle, applied to Capstone 1) and P.2 (source
governance). Every figure comes from w5_fx."""
from slidekit import NAVY, ORANGE
from v3kit import T
from w5_fx import (BA, BLANK, CLEAN, CTRL, DIRTY, DLOG, DUPS, G25, LINES, M24, M25, NSTORES, ORD25, P25, Q4M25, QD, REG, S25,
                   I, N, P)

STAGES = ["Scope", "Acquire", "Clean", "Model", "Visualise", "Verify", "Hand over"]
SHARE = [.10, .10, .30, .20, .15, .10, .05]
HRS = 40

M_P1 = dict(
    id="P.1", title="The Seven-Stage Project Lifecycle",
    subtitle="The repeatable structure behind all three capstones: scope, acquire, clean, model, visualise, verify and hand "
             "over, each stage closed by a quality gate.",
    objectives=["Name the seven stages and the output of each.", "Write a one-page brief that fixes the question.",
                "Record provenance and reconcile cleaning to a control total.", "Set quality gates and a time box.",
                "Keep stakeholders informed and close with a post-project review."],
    topics=[
        T("P.1.1", "Why a Lifecycle Matters",
          defn="A project lifecycle is a fixed order of stages that every analysis follows, each ending in a checked output.",
          simple="The order stops the classic failure: a beautiful dashboard built on unreconciled data that answers the "
                 "wrong question.",
          diagram=("flow", dict(steps=STAGES, sub=["Signed brief", "Provenance", "Cleaning log", "Measures", "Dashboard",
                                                   "Reconciliation", "README"])),
          diagram_label="Seven stages, seven outputs (Capstone 1, SuperMart)",
          watch="Skipping a stage does not save time; it moves the cost to the end, where it is most expensive."),

        T("P.1.2", "Stage 1: Scoping the Question",
          defn="Scoping turns a vague request into a precise question, tied to a decision, with measures, dimensions, a period "
               "and an owner.",
          like="Like a tailor taking measurements before cutting the cloth: cut first and the agbada will not fit.",
          grid=dict(label="The Capstone 1 one-page brief", widths=[1.3, 4.6], size=12, rows=[
              ["Brief item", "Capstone 1 entry"],
              ["Decision", "Which pricing and discount actions SuperMart takes before Q1 2026"],
              ["Question", f"Why did FY2025 sales grow {P(G25)} while gross margin fell from {P(M24)} to {P(M25)}?"],
              ["Measures", "Sales, profit, margin %, discount depth, orders, basket value, target attainment"],
              ["Period", "FY2024 vs FY2025 (calendar years)"],
              ["Sign-off", "Commercial Director (requester) and Head of Finance (control totals)"]])),

        T("P.1.3", "Stage 2: Data Acquisition",
          defn="Acquisition means getting the data and recording its provenance: source system, extraction date, row count, "
               "owner and where the raw copy is stored.",
          simple="Write the provenance table before you open the file in Power Query. Raw files are read-only.",
          grid=dict(label="Provenance, recorded before touching the data", widths=[2.9, 1.5, 0.8, 1.2], size=11.5, rows=[
              ["File", "Source", "Rows", "Owner"],
              ["SuperMart_Stores_Sales_2022_2025.csv", "POS ledger", I(LINES), "Finance"],
              ["SuperMart_Stores_Directory.csv", "Store master", I(NSTORES), "Operations"],
              ["SuperMart_Sales_2025_DIRTY.csv", "Store returns", I(len(DIRTY)), "Store admin"]]),
          watch="If you cannot say where a file came from, you cannot defend any number built on it."),

        T("P.1.4", "Stage 3: Data Cleaning",
          defn="Cleaning fixes defects with named, repeatable steps and reconciles the result to an independent control total.",
          simple="Count rows before and after every removal; the last line must equal Finance's figure.",
          grid=dict(label="Row reconciliation: the dirty 2025 store returns", widths=[2.6, 1.3, 1.1], size=12.5, rows=[
              ["Step", "Rows removed", "Rows left"],
              ["Rows received", "", I(len(DIRTY))],
              ["Remove blank spacer rows", I(BLANK), I(len(DIRTY) - BLANK)],
              ["Remove exact duplicates", I(DUPS), I(CLEAN)],
              ["Control total (POS ledger, FY2025)", "", I(CLEAN)]]),
          note="Difference 0: Gate 3 passes."),

        T("P.1.5", "Stage 4: Modelling",
          defn="Modelling builds the calculation layer: a star schema (one fact table linked to dimension tables) and explicit "
               "measures.",
          simple="Write each measure once, in DAX, and every visual reuses it.",
          fx=[("DIVIDE ( [Profit], [Revenue] )", P(M25, 2), "Gross margin %"),
              ("DISTINCTCOUNT ( FactSales[Order_ID] )", I(ORD25), "Orders, not lines"),
              ("DIVIDE ( [Revenue] - [Revenue LY], [Revenue LY] )", P(G25), "Growth vs last year")],
          fx_head=["DAX measure", "FY2025", "Meaning"], fx_widths=[3.6, 1.2, 1.6],
          watch="Margin is total profit ÷ total sales. An average of line margins gives a different, wrong answer."),

        T("P.1.6", "Stage 5: Visualisation and Dashboard",
          defn="The dashboard shows the answers: KPI cards, one chart per question and slicers connected to every visual.",
          simple="Map each visual to a question in the brief. A visual with no question is decoration.",
          grid=dict(label="Visual-to-question map", widths=[2.0, 2.6, 1.6], size=12, rows=[
              ["Visual", "Question it answers", "Chart"],
              ["KPI cards", "How big, how profitable, vs last year?", "Cards"],
              ["Margin by quarter", "When did margin fall?", "Clustered column"],
              ["Store ranking", "Which stores missed target?", "Sorted bar"],
              ["Discount bands", "Where is profit lost?", "Column"]])),

        T("P.1.7", "Stage 6: Verification",
          defn="Verification proves the numbers before anyone sees them: totals against control totals, every slicer tested, "
               "and the file opened on a second machine.",
          grid=dict(label="Capstone 1 verification", widths=[2.4, 1.7, 1.7, 0.8], size=12, rows=[
              ["Test", "Control total", "Workbook", "Result"],
              ["SuperMart lines (all years)", I(LINES), I(LINES), "PASS"],
              ["FY2025 sales", N(S25), N(S25), "PASS"],
              ["FY2025 profit", N(P25), N(P25), "PASS"],
              ["FY2025 orders", I(CTRL["Distinct orders"]), I(ORD25), "PASS"]]),
          watch="Check against a figure from the source owner, never against another number you calculated."),

        T("P.1.8", "Stage 7: Handover and Documentation",
          defn="Handover means someone else can run and maintain the work without you.",
          simple="Hand over the workbook with a 00_README sheet, the cleaning log, measure definitions and refresh steps.",
          grid=dict(label="00_README for Capstone 1", widths=[1.4, 4.4], size=12, rows=[
              ["Section", "Content"],
              ["Purpose", "Explain the FY2025 margin decline; support the Q1 2026 pricing decision"],
              ["Refresh", "New CSV in 01_RAW › Refresh All › reconciliation cells = 0"],
              ["Measures", "Each DAX measure with a plain-English meaning"],
              ["Limitations", "Calendar-year view; targets not adjusted for store size"]])),

        T("P.1.9", "Quality Gates at Each Stage",
          defn="A quality gate is a checkpoint with a pass rule at the end of each stage; work does not move on until it passes.",
          like="Like the checkpoints on the Lagos-Ibadan expressway: papers checked before you drive on.",
          grid=dict(label="Seven gates, seven pass rules", widths=[1.4, 2.8, 1.8], size=11.5, rows=[
              ["Gate", "Pass rule", "Evidence"],
              ["1 Scope", "Brief signed by the requester", "Signed brief"],
              ["2 Provenance", "Source, date, rows, owner for every file", "Provenance table"],
              ["3 Cleaning", "Rows after cleaning = control total", f"{I(CLEAN)} = {I(CLEAN)}"],
              ["4 Model", "Relationships many-to-one; date table marked", "Model diagram"],
              ["5 Visuals", "Every visual answers a brief question", "Visual map"],
              ["6 Verify", f"Dashboard total = {N(S25)}", "Reconciliation cell"],
              ["7 Handover", "README, log, measures, refresh steps", "00_README"]])),

        T("P.1.10", "Time-Boxing the Project",
          defn="Time-boxing gives each stage a fixed share of the available hours, so no stage silently eats the others.",
          simple=f"A {HRS}-hour capstone: cleaning gets the biggest share because it always takes longest.",
          chart=dict(kind="bar", cats=STAGES, series=[("Hours", [s * HRS for s in SHARE])], fmt="0", labels=True, label_fmt="0",
                     legend=False, point_colors=[ORANGE if st == "Clean" else NAVY for st in STAGES],
                     title=f"A {HRS}-hour time box (hours per stage)", h=2.7),
          points=["Cleaning 30%: the biggest single block", "Verification is planned, not squeezed in",
                  "Review the box at each gate; move hours, not the deadline"],
          points_label="RULES OF THUMB", order=["chart", "points"], split=0.55),

        T("P.1.11", "Managing Stakeholders",
          defn="Stakeholder management keeps the requester and the data owners informed at fixed moments, so there are no "
               "surprises at the end.",
          grid=dict(label="Three touchpoints in Capstone 1", widths=[0.9, 2.6, 2.4], size=12, rows=[
              ["When", "What", "Purpose"],
              ["Week 1", "One-page scope to the Commercial Director", "Sign-off on question and period"],
              ["Week 2", "Draft dashboard to Finance", "Agree control totals before polish"],
              ["Week 3", "Final pack and QA note to the board", "Decision on the discount policy"]]),
          watch="A surprise in the final meeting is a stakeholder failure, even when the analysis is right."),

        T("P.1.12", "Post-Project Review",
          defn="A post-project review asks, after delivery: did it change a decision, what took longer, what to automate, and "
               "what data problem to raise with the owner.",
          grid=dict(label="Capstone 1 review (training scenario)", widths=[2.0, 3.9], size=12, rows=[
              ["Question", "Answer recorded"],
              ["Did it change a decision?", "Board capped Furniture discounts at 20% from Q1 2026"],
              ["What took longer?", f"Cleaning the {I(DLOG['Mixed date formats'])} mixed-format dates in the returns"],
              ["Automate next time", "Monthly folder import (Week 2) instead of copy and paste"],
              ["Raise with the data owner", "Stores typing '₦' into the Sales column of the template"]])),
    ],
    takeaways=["Seven stages, each with an output and a quality gate.", "Scope fixes the question before any data is touched.",
               "Cleaning reconciles to a control total from the source owner.",
               "Time-box the stages; cleaning takes the biggest share.", "Close every project with a review."],
)

M_P2 = dict(
    id="P.2", title="Source Governance, Citation and Capstone Delivery",
    subtitle="Register every capstone source, keep before and after states, protect personal data, make the work "
             "reproducible and hand over evidence that supports a decision.",
    objectives=["Build a source register with owners and fingerprints.", "Preserve before and after states of the data.",
                "Cite sources and respect licence and privacy rules.", "Make the work reproducible and reviewable.",
                "Assemble an evidence-to-decision handoff pack."],
    topics=[
        T("P.2.1", "Beginner: source register essentials",
          defn="A source register lists every dataset used: file, owner, rows, date registered, permitted use and a SHA-256 "
               "fingerprint.",
          simple="The fingerprint proves the file you analysed is the file that was registered: one changed byte changes it.",
          grid=dict(label="Part of the capstone source register", widths=[3.3, 1.3, 0.8, 1.3], size=11, rows=[
              ["File", "Owner", "Rows", "SHA-256 (first 12)"]] + [[f, o, I(n), h] for f, o, n, h in REG]),
          note="Recompute with PowerShell: Get-FileHash file.csv"),

        T("P.2.2", "Beginner: preserve before and after states",
          defn="Keep the BEFORE file untouched and produce the AFTER table only through named cleaning steps.",
          simple="A small before/after sample shows reviewers exactly what changed and why.",
          grid=dict(label="Before and after: the dirty 2025 store returns", widths=[1.2, 2.0, 2.2], size=12.5, rows=[
              ["Field", "BEFORE (as received)", "AFTER (cleaned)"]] + [list(r) for r in BA]),
          watch="Never 'fix' the raw file by hand. Nobody can then show what was changed."),

        T("P.2.3", "Intermediate: citation, licence and privacy",
          defn="Cite each dataset by name, origin and date; check the licence (who may reuse it); and protect privacy.",
          grid=dict(label="Rules applied to the bootcamp data", widths=[1.6, 4.3], size=12, rows=[
              ["Rule", "Applied as"],
              ["Cite precisely", "'SuperMart POS ledger, SuperMart_Stores_Sales_2022_2025.csv, extracted 05-Jan-2026, Finance'"],
              ["Licence", "Training copies only; ask the owner before any public portfolio use"],
              ["Privacy", "Names are fictional but treated as real: never publish row-level data"],
              ["Sensitive files", "Welfare survey: aggregate tables only, minimum cell size 5"]])),

        T("P.2.4", "Intermediate: reproducibility and review",
          defn="Work is reproducible when someone else can rerun it from the raw files and get the same totals.",
          steps=["Name every Power Query step ('Removed spacer rows')", "Store assumptions in cells, not in your head",
                 "Test a refresh with next month's file", f"Ask a colleague to follow the README and get {N(S25)}"],
          steps_label="FOUR PRACTICES",
          watch="If only you can refresh it, it is not finished."),

        T("P.2.5", "Advanced capstone: evidence-to-decision handoff",
          defn="The handoff pack links evidence to the decision: register, before/after sample, workbook or SQL, QA checks, "
               "findings with their uncertainty, an owner and a follow-up date.",
          grid=dict(label="Facts and recommendations kept apart", widths=[1.5, 4.4], size=12, rows=[
              ["Item", "Capstone 1"],
              ["Evidence", "Register, before/after, workbook, reconciliation (all PASS)"],
              ["Fact", f"Q4 2025 margin fell to {P(Q4M25)} as discount depth rose to {P(QD[(2025, 'Q4')])}"],
              ["Recommendation", "Cap Furniture discounts at 20% from Q1 2026"],
              ["Owner and follow-up", "Commercial Director; margin reviewed monthly on the dashboard"]]),
          watch="Mixing facts and opinions in one sentence makes both look weaker."),
    ],
    takeaways=["Register every source with an owner and a fingerprint.", "Raw files stay untouched; cleaning is named steps.",
               "Cite precisely and never publish personal rows.", "Reproducible means someone else gets the same totals.",
               "Hand over evidence, facts and recommendations separately."],
)

MODULES = [M_P1, M_P2]
