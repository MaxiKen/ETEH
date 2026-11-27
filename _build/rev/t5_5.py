from revlib import T

TRY = "ETEH_Capstone_Section_Examples.xlsx  ›  sheet 'P.5 Capstone3 Lifecare'"
SOL = "03_Instructor_Projects  ›  Capstone3_Lifecare_SOLUTION.xlsx  ›  "


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", "Lifecare Hospital · Capstone 3")
    return T(tag, **k)


TOPICS = [
    O("P.5.1",
      what="Capstone 3: the board asks how full the hospital is, who stays longest, whether any department's "
           "readmissions are truly worse, and whether the payer mix is risky. A rule is built into the brief: "
           "**ward-level reporting only, no patient identifiers**.",
      like="Like a hospital matron's report: beds, stays and returns, without naming a single patient.",
      how=["Write the board's four questions.", "Add the privacy constraint to the brief.", "Pick the measures: bed-days, LOS, rates."],
      tip="Put ethical limits in the scope, before the first query.",
      trap="Pulling patient names 'just in case'.",
      case="Headline capacity numbers from the bed census (all years).",
      steps=["MAX of quarterly occupancy.", "COUNTIFS days above 95%.", "COUNTIFS days above 100%."],
      table=[["Measure", "Value"], ["Peak quarterly occupancy", "94.3% (2019-Q4)"], ["Days above 95%", "259"],
             ["Days above 100% (patients on trolleys)", "44"], ["Census days", "1,797"]],
      so="The hospital often runs above safe levels. That's the board's concern.",
      split=4.90),

    O("P.5.2",
      what="The **inventory** for Capstone 3 is one SQLite database, lifecare_hospital.db, with five tables. "
           "Record each table's size and what one row means.",
      like="Like listing the departments of a hospital before walking the wards.",
      how=["List the tables.", "Count rows in each.", "Note the period and owner."],
      tip="Extract only the columns the questions need.",
      trap="Exporting the whole database to Excel.",
      case="Tables in lifecare_hospital.db.",
      steps=["SELECT COUNT(*) per table.", "Check the date range.", "Record owners."],
      table=[["Table", "Rows", "Grain"], ["admissions", "54,860", "One admission, May 2019 – May 2024"],
             ["patients", "48,803", "One patient"], ["doctors", "28", "One doctor"],
             ["wards", "4", "130 staffed beds each = 520"], ["bed_census", "1,797", "One day"]],
      so="Five tables, one per thing. FY2023 has 10,918 admissions and 365 census days.",
      split=4.70),

    O("P.5.3",
      what="Extract with **SQL**, not copy-paste. Q1 gives **control totals** by year; Q2 pulls FY2023 admissions "
           "**without names or patient IDs**. Reconcile the extract to Q1 before analysing.",
      like="Like a bank teller counting cash against the deposit slip before banking it.",
      how=["Run Q1: totals by year.", "Run Q2: the minimised FY2023 extract.", "Compare every total; difference must be 0."],
      tip="Reconcile money to the kobo, counts to the row.",
      trap="Analysing an extract you haven't reconciled.",
      case="FY2023 extract vs the database control totals.",
      steps=["Q1 for 2023.", "Sum the Excel extract.", "Difference = 0?"],
      table=[["Control", "Database (Q1)", "Excel extract", "Status"], ["Admissions", "10,918", "10,918", "OK"],
             ["Billing", "₦13,984,047,685.44", "₦13,984,047,685.44", "OK"], ["Readmissions", "1,129", "1,129", "OK"],
             ["Census days", "365", "365", "OK"]],
      so="The extract matches the database exactly, so analysis can begin.",
      tryit=SOL + "'2 SQL & Reconcile'",
      split=4.20),

    O("P.5.4",
      what="**Occupancy %** = occupied bed-days ÷ available bed-days, from the daily census. Don't average daily "
           "percentages, and never use admission counts.",
      like="Like measuring how full a hotel was all month: room-nights sold ÷ room-nights available.",
      how=["Sum patients in house over the period.", "Sum staffed beds over the period.", "Divide."],
      tip="Write the denominator choice in the README.",
      trap="Averaging daily percentages when capacity changes.",
      case="FY2023 occupancy from the census.",
      steps=["SUM patients_in_house for 2023.", "SUM staffed_beds for 2023.", "Divide; compare with 85%."],
      table=[["Item", "Value"], ["Occupied bed-days", "169,923"], ["Available bed-days", "189,800 (520 × 365)"],
             ["FY2023 occupancy", "89.5%"], ["Months above the 85% target", "11 of 12 (November 82.5%)"],
             ["Days above 95%", "43"]],
      so="The hospital ran above target nearly all year.",
      split=4.90),

    O("P.5.5",
      what="Report **length of stay (LOS)** as mean, median and **90th percentile** (P90: 9 in 10 stays are this "
           "long or shorter). P90 is what discharge planning needs.",
      like="Like planning a bus timetable: the average trip matters, but so does how long the slow trips take.",
      how=["AVERAGEIFS for the mean.", "Median and P90 with AGGREGATE.", "Compare wards."],
      tip="Look at the long tail: how many bed-days come from long stays?",
      trap="Reporting only the mean.",
      case="2023 length of stay by ward.",
      steps=["Mean, median, P90 per ward.", "Count stays of 22+ days.", "Share of bed-days they use."],
      table=[["Ward", "Admissions", "Mean", "Median", "P90"], ["Emerald", "2,631", "15.6", "16", "28"],
             ["Ruby", "2,798", "15.3", "15", "28"], ["Sapphire", "2,755", "15.6", "16", "28"],
             ["Topaz", "2,734", "15.5", "15", "28"]],
      so="Wards differ by under 0.3 days. The 3,319 stays of 22+ days use 51% of all bed-days.",
      split=4.20),

    O("P.5.6",
      what="The 30-day **readmission rate** is a **safety** measure. Show it by department with a **95% interval**, "
           "and send any finding to the clinical lead before it reaches the board.",
      like="Like a referee reviewing the video before giving a red card.",
      how=["Rate = readmitted ÷ discharges.", "Add the 95% interval.", "Flag only clear differences."],
      tip="Case mix matters: cancer patients often return for planned treatment.",
      trap="Ranking departments whose intervals overlap.",
      case="2023 readmission by department.",
      steps=["COUNTIFS readmitted.", "Divide by discharges.", "Compare intervals with 10.3%."],
      table=[["Department", "Discharges", "Rate", "95% interval"], ["Oncology", "1,824", "13.2%", "11.7% – 14.8%"],
             ["Cardiology", "1,828", "11.4%", "10.0% – 12.9%"], ["Endocrinology", "3,682", "9.9%", "9.0% – 10.9%"],
             ["Orthopaedics & Rheumatology", "1,806", "8.3%", "7.0% – 9.5%"]],
      tryit=SOL + "'6 Readmission Safety'",
      so="Only Oncology sits clearly above the 10.3% hospital rate. It's a question for the clinical lead, not a verdict.",
      split=4.20),

    O("P.5.7",
      title="Escalation Logic: One Rule for Every Report",
      what="An **escalation rule** flags admissions for review. Write it once as a conditional column so every "
           "report uses the same rule. The **order** matters: the first matching reason is recorded.",
      like="Like a hospital gate: emergencies first, then appointments, then visitors. Same order every time.",
      how=["Readmitted within 30 days?", "Else stay of 28+ days?", "Else age 75+? Otherwise blank."],
      tip="Clinical owners must approve the thresholds. The tool advises; people decide.",
      trap="Treating a classroom rule as a clinical decision.",
      case="The rule applied to Q4 2023 admissions (Admission_ID only).",
      steps=["Nested IF (or Power Query if/else).", "Count each reason.", "Total the escalations."],
      table=[["Reason", "Admissions"], ["Readmitted within 30 days", "274"], ["Stay of 28+ days", "220"],
             ["Age 75+", "320"], ["Total escalations", "814 of 2,666"]],
      so="About 3 in 10 Q4 admissions are flagged, all by one consistent rule.",
      split=4.90),

    O("P.5.8",
      what="The board dashboard fits one screen: KPI cards (admissions, occupancy, mean LOS, readmission), the "
           "occupancy trend against 85%, ward bars and payer mix. **Nothing on screen identifies a patient.**",
      like="Like a hospital's lobby screen: waiting times and bed status, no patient names.",
      how=["Four KPI cards.", "Occupancy trend with an 85% line.", "Ward bars and payer mix."],
      tip="Check every visual and tooltip for IDs before sharing.",
      trap="A drill-through that reveals patient rows.",
      case="FY2023 KPI cards and payer mix.",
      steps=["Calculate the four KPIs.", "Count admissions by HMO.", "Check for identifiers."],
      table=[["Item", "Value"], ["Admissions", "10,918"], ["Occupancy", "89.5%"], ["Mean LOS", "15.5 days"],
             ["Readmission", "10.3%"], ["Payer mix", "5 HMOs, 18.9% to 20.6% each"]],
      so="The board sees everything it needs, and nothing that names a patient.",
      split=4.90),

    O("P.5.9",
      what="The **board pack**: one-page summary, capacity, length of stay, a readmission safety note, payer mix, a "
           "data-quality appendix and a confidentiality statement. Each recommendation has an **owner**.",
      like="Like a ward-round summary for the chief medical officer: short, factual, with clear actions.",
      how=["Summary page first.", "One page per question.", "Recommendations with owners."],
      tip="Flag clinical findings as 'under clinical review'.",
      trap="Recommendations with no owner.",
      case="The Capstone 3 recommendations.",
      steps=["Size the capacity gap.", "Write each recommendation.", "Assign the owner."],
      table=[["Recommendation", "Owner"], ["Fund about 28 staffed beds, or a discharge team for 22+ day stays", "COO"],
             ["Review Oncology readmissions", "Medical Director"], ["Monthly dashboard from the same SQL queries", "Analytics"]],
      so="About 28 extra beds would bring average occupancy to 85%; shorter long stays free similar capacity.",
      tryit=SOL + "'8 Board Pack'",
      split=4.90),

    O("P.5.10",
      what="Verify admissions and billing against the database control totals, occupancy against the census, and "
           "**confirm no patient identifiers** appear anywhere. Document the definitions clinicians may dispute.",
      like="Like a final medication check before the patient goes home.",
      how=["Reconcile counts and billing.", "Rebuild occupancy from the census.", "Scan outputs for names and IDs."],
      tip="Search every sheet for 'Patient' before release.",
      trap="Releasing without the privacy scan.",
      case="Capstone 3 verification (all years).",
      steps=["Compare with each control.", "Record the result.", "Sign off."],
      table=[["Check", "Control", "Result"], ["Admissions", "Health Records: 54,860", "Match"],
             ["Billing", "Finance ledger: ₦70,206,080,067", "Match"], ["Occupancy", "Bed census", "Rebuilt exactly"],
             ["Identifiers", "No names or Patient_IDs", "None found"]],
      so="The numbers reconcile and the outputs are safe to share.",
      split=4.70),
]
