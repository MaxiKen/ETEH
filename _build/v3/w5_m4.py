"""Capstone edition-3 slide copy: Module P.4, Capstone 2 (Kingsway University progression). Figures from w5_fx, i.e. from
the reference analysis behind Capstone2_Kingsway_SOLUTION.xlsx."""
from slidekit import GREEN, NAVY, ORANGE, RED
from v3kit import T
from w5_fx import (AT, BAND_N, BAND_RATE, BETA, CORR, CV, DLOG2, FAC, FAC_N, LIFT, LVL, MECH, N_DEPT2, N_FAC2, NCV, NLOWATT,
                   NPROG, NRAW2, NRISK, PROG, PROG_AFTER, R2, RESCUED, SLOPE_ATT, ST, TGT, TOP_SHARE, TRIG, I, P)

BANDS = list(BAND_RATE)
GAP = (PROG - TGT) * 100
DASH = dict(kind="dashboard", full=True, h=4.95, banner="Kingsway 2024/25: progression below the Senate target",
            kpis=[("Cohort", I(NCV), "students"), ("Progression", P(PROG), f"target {P(TGT, 0)}", RED),
                  ("At risk", I(NRISK), f"{P(NRISK / NCV)} of cohort", RED),
                  ("Attendance", f"{CV.Lecture_Attendance_Pct.mean():.1f}%", "cohort average"),
                  ("If attendance lifts", P(PROG_AFTER), f"+{I(RESCUED)} students", GREEN)],
            charts=[dict(kind="bar", cats=FAC.index.tolist()[::-1], series=[("At-risk rate", FAC.tolist()[::-1])], fmt="0%", legend=False,
                         point_colors=[RED if i == len(FAC) - 1 else NAVY for i in range(len(FAC))], title="At-risk rate by faculty"),
                    dict(kind="column", cats=BANDS, series=[("At-risk rate", [BAND_RATE[b] for b in BANDS])], fmt="0%", legend=False,
                         colors=[ORANGE], title="At-risk rate by attendance band"),
                    dict(kind="column", cats=[str(k) for k in LVL.index], series=[("At-risk rate", LVL.tolist())], fmt="0%", legend=False,
                         title="At-risk rate by level")],
            widths=[1.1, 1.1, 0.9])

M_P4 = dict(
    id="P.4", title="Capstone 2: Kingsway University Progression",
    subtitle="A complete higher-education build: registry export cleaning, a fixed progression definition, where risk "
             "concentrates, an actionable at-risk list and a sized intervention.",
    objectives=["Fix the cohort and the progression rule before calculating.", "Clean the registry export in a safe order.",
                "Find where at-risk students concentrate.", "Produce an actionable, restricted at-risk list.",
                "Size an intervention with a stated assumption and verify the paper."],
    topics=[
        T("P.4.1", "Capstone 2 Brief and Success Criteria",
          defn=f"The Senate wants {P(TGT, 0)} of students in good standing. Which faculties are behind, what drives "
               "performance, and who needs support now?",
          simple="Success means a fixed definition, a reconciled cohort, an at-risk list advisers can use and a costed action.",
          grid=dict(label="The headline numbers", widths=[2.4, 1.4], size=13, rows=[
              ["Measure", "Value"], ["Cohort students", I(NCV)], ["Progressed (good standing)", I(NPROG)],
              ["Progression rate", P(PROG)], ["Gap to the 90% target", f"{GAP:+.1f} points".replace("-", "−")],
              ["At-risk students", I(NRISK)]])),

        T("P.4.2", "Kingsway Data Inventory",
          defn="The inventory for Capstone 2 is one registry export: faculty, department, level, attendance, study habits and "
               "the final score, one row per student.",
          grid=dict(label="Kingsway_Registry_Export_RAW.csv", widths=[2.6, 2.6], size=12.5, rows=[
              ["Item", "Value"], ["Raw rows", I(NRAW2)], ["Duplicate registrations", I(DLOG2["Duplicate students"])],
              ["Students after cleaning (control)", I(NCV)],
              ["Faculties / departments", f"{N_FAC2} / {N_DEPT2}"],
              ["Grain", "One student (Matric_No), semester 1 2024/25"]]),
          watch="The registry's published cohort is the control total: your clean table must match it exactly."),

        T("P.4.3", "Cleaning the Registry Extract",
          defn="Clean the registry in a safe order: TRIM the IDs first, then de-duplicate, then standardise codes and text.",
          grid=dict(label="Defect log (rows affected)", widths=[2.0, 0.7, 2.6], size=11.5, rows=[
              ["Defect", "Rows", "Fix"]] + [[d, I(n), f] for d, n, f in [
              ("Padded keys", DLOG2["Padded keys"], "TRIM before any lookup or de-duplication"),
              ("Duplicate students", DLOG2["Duplicate students"], "Remove Duplicates on the trimmed ID"),
              ("Percent as text", DLOG2["Percent as text"], 'VALUE(SUBSTITUTE(x,"%",""))'),
              ("Department variants", DLOG2["Department variants"], "Lookup to the department master"),
              ("Missing categories", DLOG2["Missing categories"], "Label 'Not Recorded'; never delete"),
              ("Out-of-range score", DLOG2["Out-of-range score"], "Query, then cap at 100")]]),
          watch="De-duplicate before trimming and ' KU/…' and 'KU/…' look like two students."),

        T("P.4.4", "Defining the Progression Measure",
          defn=f"Progression % = students meeting the rule ÷ cohort. The rule: final score of {ST} or more AND attendance of "
               f"{AT}% or more. At risk is everyone else.",
          fx=[(f'=COUNTIFS(tCV[Final_Score],">="&{ST},tCV[Lecture_Attendance_Pct],">="&{AT})', I(NPROG), "Progressed"),
              ("=ROWS(tCV)", I(NCV), "Cohort"),
              ("=progressed / cohort", P(PROG), "Progression %")],
          fx_widths=[4.4, 0.9, 1.3], stack=True,
          watch="Fix the definition before calculating. Changing the rule after seeing the result is moving the goalposts."),

        T("P.4.5", "Where At-Risk Students Concentrate",
          defn="Break the at-risk rate down by group to find where risk concentrates; support pays off most there.",
          grid=dict(label="At-risk rate by attendance band", widths=[1.4, 1.1, 1.2], size=12.5, rows=[
              ["Attendance", "Students", "At-risk rate"]] + [[b, I(BAND_N[b]), P(BAND_RATE[b])] for b in BANDS]),
          chart=dict(kind="bar", cats=FAC.index.tolist()[::-1], series=[("At-risk rate", FAC.tolist()[::-1])], fmt="0%", labels=True,
                     label_fmt="0.0%", legend=False, point_colors=[RED if i == len(FAC) - 1 else NAVY for i in range(len(FAC))],
                     title=f"{FAC.index[0]} is furthest behind", h=2.5),
          order=["grid", "chart"], split=0.47,
          watch=f"{FAC.index[0]} holds {P(TOP_SHARE, 0)} of all at-risk students. Attendance below 70% is the sharpest signal."),

        T("P.4.6", "The At-Risk List",
          defn="An at-risk list must be usable: student ID, faculty, level, the trigger (why flagged) and the support route.",
          grid=dict(label="Triggers and support routes", widths=[2.0, 0.9, 2.6], size=12.5, rows=[
              ["Trigger", "Students", "Support route"],
              ["Low score only", I(TRIG["Low score"]), "Week-5 tutorial programme"],
              ["Low attendance only", I(TRIG["Low attendance"]), "Pastoral call and attendance contract"],
              ["Score and attendance", I(TRIG["Score and attendance"]), "Adviser meeting first"],
              ["Total", I(NRISK), ""]]),
          ex=dict(t="tAtRiskTop20", cols=["Matric_No", "Faculty", "Level", "Final_Score", "Lecture_Attendance_Pct", "Trigger"], n=3, size=11),
          order=["grid", "ex"], stack=True,
          watch="Named lists are restricted: advisers only, never in the Senate paper."),

        T("P.4.7", "Quantifying the Intervention",
          defn="Size the effect of an intervention with a stated assumption, and report it as an estimate.",
          simple=f"Assume the tutorial programme lifts attendance by {LIFT} points for students below {AT}%; the regression says each "
                 f"attendance point is worth {BETA['Lecture_Attendance_Pct']:.2f} exam points.",
          grid=dict(label="Re-applying the rule to the lifted cohort", widths=[2.6, 1.2], size=12.5, rows=[
              ["Measure", "Value"], [f"Students below {AT}% attendance", I(NLOWATT)], ["At-risk students who would progress", I(RESCUED)],
              ["…already passing the exam bar", I(MECH)], ["Progression after the lift", P(PROG_AFTER)],
              ["Target", P(TGT, 0)]]),
          chart=dict(kind="column", cats=["Now", "After the lift", "Target"], series=[("Progression", [PROG, PROG_AFTER, TGT])],
                     fmt="0%", labels=True, label_fmt="0.0%", legend=False, point_colors=[RED, ORANGE, GREEN], y_min=0.5,
                     title="The lift narrows the gap but does not close it", h=2.5, y_max=1.0),
          order=["grid", "chart"], split=0.5,
          watch=f"Association, not proof: R² is {R2:.2f}. Present the estimate with its assumption."),

        T("P.4.8", "Building the Progression Dashboard",
          defn="The progression dashboard shows the cohort, the rate against target, where risk sits and the effect of the "
               "intervention; the named list sits on a restricted page.",
          grid=dict(label="Drivers of the final score (correlation r)", widths=[2.4, 1.0], size=12.5, rows=[
              ["Variable", "r"]] + [[k.replace("_", " "), f"{v:.2f}"] for k, v in sorted(CORR.items(), key=lambda kv: -abs(kv[1]))]),
          watch="Sleep and exercise show almost no link to the score here; do not build the story on them.",
          more=[dict(title="The Capstone 2 Dashboard", pill="DASHBOARD", chart=DASH,
                     notes="Five KPI cards and three at-risk views. The axis-free comparison of faculty, attendance band and level "
                           "shows that attendance is the strongest lever.")]),

        T("P.4.9", "The Senate Paper",
          defn="The Senate paper runs: context, finding, evidence, recommendation, risks, in that order and on two pages.",
          grid=dict(label="The paper's spine", widths=[1.4, 4.5], size=12, rows=[
              ["Section", "Content"], ["Context", f"Target {P(TGT, 0)} of students in good standing"],
              ["Finding", f"{P(PROG)} progressed; {I(NRISK)} at risk, most in {FAC.index[0]}"],
              ["Evidence", f"At-risk rate {P(BAND_RATE[BANDS[0]])} below 70% attendance vs {P(BAND_RATE[BANDS[-1]])} at 90%+"],
              ["Recommendation", f"Attendance-led tutorial programme; owner: Deans; estimate +{I(RESCUED)} students"],
              ["Risks", "Association is not causation; monitor the effect each semester"]])),

        T("P.4.10", "Kingsway Verification and Handover",
          defn="Verify the paper before it goes: cohort = registry, progressed = examinations office figure, every slicer "
               "drives every visual, and no names leave the restricted page.",
          grid=dict(label="Verification checks", widths=[2.4, 2.2, 0.8], size=12.5, rows=[
              ["Check", "Evidence", "Status"], ["Cohort = registry", f"{I(NCV)} = {I(NCV)}", "Pass"],
              ["Progressed = exams office", I(NPROG), "Pass"], ["Triggers add up", f"{I(sum(TRIG))} = {I(NRISK)}", "Pass"],
              ["Attendance slope", f"{SLOPE_ATT:.3f} points per point", "Logged"],
              ["Named list restricted", "Adviser page only", "Pass"]])),
    ],
    takeaways=["Fix the cohort and the progression rule before any calculation.", "TRIM IDs before de-duplicating.",
               "Attendance is the sharpest signal of risk.", "An at-risk list needs a trigger and a support route.",
               "Size interventions with a stated assumption and say it is an estimate."],
)

MODULES = [M_P4]
