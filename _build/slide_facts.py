"""Fact substitutions: fictitious figures in the original decks -> real figures from the bootcamp datasets.

FACT_SUBS = {vol: {scope: [(old, new), ...]}}; scope is "*" or a tuple of ORIGINAL slide numbers.
Substitutions run in insertion order, so slide-specific edits come before the deck-wide ones.

Sources for every figure (all verified with pandas against ETEH_Bootcamp_Pack/01_Section_Examples/00_Core_Datasets):
  SuperMart FY2025      3,312 lines · 1,687 orders · ₦733,215,255 sales · ₦93,439,242 profit · margin 12.74% (FY2024 13.43%)
                        growth +20.4% · mean order ₦434,627 · median order ₦148,257 · 50.2% of orders < ₦150k · 4.0% > ₦2m
                        1,525 selling SKUs: top 5% (76) = 44.6% of revenue, top 324 (21%) = 80%, bottom 30% (457) = 1.3%
                        lines discounted > 20%: 455, loss ₦47.6m (FY2024 ₦31.4m); Tables + Machines loss ₦11.0m
                        20% Furniture discount cap (167 lines) -> profit uplift ₦17.7m–₦18.0m (Instructor Capstone 1); all-category ₦59.6m–₦71.5m is phase 2 only
  SuperMart dirty file  SuperMart_Sales_2025_DIRTY.csv: 3,387 rows = 3,312 + 60 exact duplicates + 15 spacer rows;
                        duplicates carry ₦13.2m (1.8% of sales); ledger 2022-25 9,994 lines; store-category-month = 1,540 rows
  Crestview             6,607 students; Engineering 1,400; at risk 901 (progression 86.4%); mean exam score 67.2
  Lifecare FY2023       10,918 admissions; occupancy 89.5%; mean LOS 15.5, median 15; readmission 10.3%;
                        billing ₦13,984,047,685.44 (₦1,280,825 per admission); escalations 5,099; +28 staffed beds;
                        full extract 2019-2024 = 54,860 rows; DIRTY 2023 file has 70 blank discharge dates
"""

# ------------------------------------------------------------------ helpers
_ALL_V2_BUT_207 = tuple(i for i in range(1, 229) if i != 207)

# Crestview headcount appears in many generic CGPA illustrations. Use the real headcount (6,607) everywhere it
# denotes students; derived figures on the same slides are fixed first.
CRESTVIEW = [("4,212", "6,607")]

FACT_SUBS = {
    1: {
        (125,): [("reach row 4,212, then clicks M4212", "reach row 6,608, then clicks M6608"),
                 ("C4212:G4212,H4212:L4212", "C6608:G6608,H6608:L6608"),
                 ("SUM(H4212:L4212)", "SUM(H6608:L6608)")],
        (150,): [("recalculates live to 648 ,", "recalculates live to 1,400 ,")],
        (175,): [("A typed occupancy figure of 84% on", "A typed occupancy figure of 89% on"),
                 ("gives 84.2% and", "gives 89.5% and")],
        "*": CRESTVIEW,
    },
    2: {
        (72,): [("qrySales filters out Sale_Type = \"QUOTE\" and Store = \"TEST-HQ\" , removing 1,842 non-revenue rows from 41,654 "
                 "so that every downstream total reconciles to finance.",
                 "qrySales on SuperMart_Sales_2025_DIRTY.csv filters out rows where Order_ID is null, removing the 15 blank "
                 "spacer rows from 3,387 so that every downstream total reconciles to finance.")],
        (99,): [("demonstrates that filtering TEST-HQ removes 1,842 rows that would otherwise have inflated revenue by 4.3%.",
                 "demonstrates that removing 60 duplicate lines takes out ₦13.2m that would otherwise have inflated revenue by 1.8%.")],
        (101,): [("Same 41,654-row POS extract. Manual: 47 minutes, 41,654 rows retained including 1,842 quotes and 340 duplicates, "
                  "revenue overstated by ₦2.4m. Power Query: 6 seconds, 39,472 clean rows, revenue matching the finance ledger to the cent.",
                  "Same 3,387-row SuperMart_Sales_2025_DIRTY.csv. Manual: 47 minutes, 3,372 rows retained including 60 duplicates, "
                  "revenue overstated by ₦13.2m. Power Query: 6 seconds, 3,312 clean rows, revenue matching the ₦733,215,255 control total."),
                 ("41,654 (duplicates and quotes retained)", "3,372 (duplicates retained)"),
                 ("39,472 verified", "3,312 verified"),
                 ("+₦2.4m overstated", "+₦13.2m overstated")],
        (126,): [("returns 315 scholarship places from 4,212 students; ROUNDUP would award 316",
                  "returns 495 scholarship places from 6,607 students; ROUNDUP would award 496"),
                 ("4212*0.075, 0", "6607*0.075, 0"), ("315 places", "495 places"), ("316 places", "496 places")],
        (193,): [("totals row at A4213", "totals row at A6609")],
        (90,): [("AMT LC ($)", "AMT LC (₦)")],
        (210,): [("still shows 4,116 until refreshed", "still shows 6,511 until refreshed")],
        (211,): [("returns the true aggregate margin of 22.8%. Averaging a row-level margin column would have returned 19.4%, "
                  "because it weights a ₦50 sale equally with a ₦50,000 sale.",
                  "returns the true FY2025 margin of 12.7%. Averaging a row-level margin column would have returned 11.6%, "
                  "because it weights a ₦5,000 sale equally with a ₦5,000,000 sale.")],
        _ALL_V2_BUT_207: CRESTVIEW,
    },
    3: {
        (118,): [("and 648 when the Engineering slicer", "and 1,400 when the Engineering slicer")],
        (143,): [("Revenue 18,420,915.40 , Gross Margin 22.8% , Basket Value 8,412 , Transactions 41,654 , Like-for-like +6.25% "
                  "— five measures that answer every board question asked about the month.",
                  "Revenue ₦733,215,255 , Gross Margin 12.74% , Order Value ₦434,627 , Orders 1,687 , Growth +20.4% "
                  "— five FY2025 measures that answer every board question about the year."),
                 ("₦8,500", "₦450,000"), ("24%", "14%")],
        (144,): [("Admissions 500 , Occupancy 84.2% , Mean LOS 4.6 days , Median LOS 3.0 days , Readmissions 6.1% , "
                  "Cost per admission ₦8,424 — the six figures the clinical board reviews monthly.",
                  "Admissions 10,918 , Occupancy 89.5% , Mean LOS 15.5 days , Median LOS 15 days , Readmissions 10.3% , "
                  "Billing per admission ₦1,280,825 — the six FY2023 figures the clinical board reviews.")],
        (145,): [("Enrolment 4,212 , Weighted CGPA 3.41 , Progression 76% , At Risk 37 , Completion 89% — the set that lets "
                  "the senate see both institutional performance and the intervention list.",
                  "Enrolment 6,607 , Mean Exam Score 67.2 , Progression 86.4% , At Risk 901 , Engineering 1,400 — the set that "
                  "lets the senate see both institutional performance and the intervention list.")],
        (168,): [("such as 84% occupancy", "such as 90% occupancy"), ("84.2% in the centre", "89.5% in the centre")],
        (174,): [("A revenue-to-margin waterfall steps from ₦18.4m revenue through cost of goods, freight, shrinkage and discount "
                  "to arrive at ₦4.2m gross profit, making clear that discount — not cost of goods — caused the margin shortfall.",
                  "A FY2025 waterfall steps from ₦915.5m at list price through ₦182.2m of discounts to ₦733.2m net sales, then "
                  "₦639.8m of cost to ₦93.4m profit, making clear that discount, not cost, drives the margin shortfall.")],
        (209,): [("On the finished model, revenue reconciles to the finance ledger at ₦18,420,915.40 across 41,654 clean lines. "
                  "The board uses it to delist two slow categories and to reprice 148 SKUs whose margin fell below 15%.",
                  "On the finished model, revenue reconciles to the ledger at ₦733,215,255 across 3,312 clean FY2025 lines. "
                  "The board uses it to cap Furniture discounts at 20% and to review the five loss-making stores.")],
        (175,): [("runs 12,400 enquiries, 6,100 applications, 5,200 offers, 4,212 enrolments.",
                  "runs 19,600 enquiries, 9,600 applications, 8,200 offers, 6,607 enrolments."),
                 ("12,400", "19,600"), ("6,100", "9,600"), ("69% of offers", "81% of offers")],
        (189,): [("The KPI tile reads 18.4m at 24pt with Revenue · +6.3% vs LY at 8pt",
                  "The KPI tile reads ₦733.2m at 24pt with Revenue · +20.4% vs LY at 8pt"),
                 ("18.4m", "₦733.2m"), ("Revenue up 6.3% on last year", "Revenue up 20.4% on last year")],
        (190,): [("Crestview's four cards read 4,212 Enrolment · +2.4% , 3.41 Weighted CGPA · flat , 76% Progression · -3pts , "
                  "37 At Risk · +9 .",
                  "Crestview's four cards read 6,607 Enrolment , 67.2 Mean Exam Score , 86.4% Progression · target 90% , "
                  "901 At Risk · 13.6% .")],
        (210,): [("The dashboard reports 4,212 students, weighted CGPA 3.41, progression 76% (down 3 points) and 37 students at risk. "
                  "It directly triggered the Week-5 tutorial programme that lifted Year-1 progression from 76% to 89%.",
                  "The dashboard reports 6,607 students, mean exam score 67.2, progression 86.4% and 901 students at risk "
                  "(585 on attendance alone). It directly triggered the Week-5 attendance programme.")],
        (211,): [("The dashboard shows 500 admissions, 84.2% occupancy, mean LOS 4.6 days (median 3.0), readmissions 6.1% and "
                  "41 escalations. It exposed a Q4 respiratory surge in Wards 3 and 4 that led to staggered- shift staffing.",
                  "The FY2023 dashboard shows 10,918 admissions, 89.5% occupancy, mean LOS 15.5 days (median 15), readmissions "
                  "10.3% and 5,099 escalations. Occupancy above target all year made the case for 28 more staffed beds.")],
        "*": CRESTVIEW,
    },
    4: {
        (6,): [("SuperMart's mean basket of ₦8,412 is the correct figure for planning till throughput, because the total revenue "
                "divided by transactions is what the business actually experiences.",
                "SuperMart's FY2025 mean order of ₦434,627 (₦733.2m ÷ 1,687 orders) is the correct figure for revenue planning, "
                "because total revenue divided by orders is what the business actually experiences.")],
        (29,): [("A histogram of 41,654 basket values shows a strong peak at ₦6,000–₦9,000 and a thin tail beyond ₦50,000 — "
                 "revealing that the average of ₦8,412 sits above the modal shopper, which the raw average had concealed.",
                 "A histogram of 1,687 FY2025 orders shows half below ₦150,000 and a thin tail beyond ₦2m — so the ₦434,627 "
                 "average sits far above the ₦148,257 median order, which the raw average had concealed.")],
        (30,): [("Eighteen percent of SuperMart's 1,842 SKUs generate 82% of revenue. The Pareto chart makes that visible in one "
                 "image and justifies concentrating stock investment on those 331 lines while delisting the bottom decile.",
                 "21% of SuperMart's 1,525 selling SKUs generate 80% of FY2025 revenue. The Pareto chart makes that visible in "
                 "one image and justifies focusing stock on those 324 lines while reviewing the bottom 30%."),
                ("Top 18%", "Top 21%"), ("331", "324"), ("82%", "80%"), ("92", "76"), ("48%", "45%"),
                ("553", "457"), ("2%", "1.3%")],
        (99,): [("Optimistic (4,600, fees +5%) and Pessimistic (3,900, fees -3%)", "Optimistic (7,200, fees +5%) and Pessimistic (6,100, fees -3%)")],
        (108,): [("Margin fell 1.8 points; 148 SKUs priced below 15%; repricing them recovers ₦1.1m",
                  "Margin fell 0.7 points; lines discounted above 20% lost ₦47.6m; a 20% Furniture cap recovers ₦17.7m–₦18.0m")],
        (110,): [("Revenue grew 6.3% but margin fell 1.8 points because discounting doubled in Q3.",
                  "Revenue grew 20.4% but margin fell 0.7 points because deep-discount losses rose from ₦31.4m to ₦47.6m.")],
        (111,): [("Revenue up 6.3%, margin down 1.8pts", "Revenue up 20.4%, margin down 0.7pts"),
                 ("340 duplicate records inflate revenue 2.4%", "60 duplicate lines inflate revenue 1.8%")],
        (112,): [("Supporting: discount depth doubled, 148 SKUs under 15% margin, two categories account for 71% of the erosion. "
                  "Appendix: the full SKU table.",
                  "Supporting: deep-discount losses up from ₦31.4m to ₦47.6m, 455 lines discounted above 20%, Tables and Machines "
                  "lost ₦11.0m. Appendix: the full SKU table.")],
        (121,): [("SuperMart's summary: revenue +6.3% to ₦18.4m, margin -1.8 points to 22.8%, 148 SKUs below the 15% threshold; "
                  "recommend repricing (owner: category managers) to recover ₦1.1m by Q1.",
                  "SuperMart's summary: revenue +20.4% to ₦733.2m, margin -0.7 points to 12.7%, 455 lines discounted above 20% "
                  "lost ₦47.6m; recommend a 20% Furniture cap (owner: category managers) to recover ₦17.7m–₦18.0m.")],
        (146,): [("rows 4000–4212", "rows 6000–6608")],
        (147,): [("Watching the margin total, revenue total and 148-SKU flag count while fixing a broken IFERROR shows immediately "
                  "that the correction restores ₦1.1m of margin without changing revenue.",
                  "Watching the profit total (₦93,439,242), revenue total and the 455-line discount flag while fixing a broken "
                  "IFERROR shows immediately that the correction restores profit without changing revenue.")],
        "*": CRESTVIEW,
    },
    5: {
        (8,): [("clinical.admissions, extracted 2026-11-15, 2,104,882 rows, owner: Data Office",
                "Lifecare_Admissions_2019_2024.csv, extracted 2026-11-15, 54,860 rows, owner: Data Office")],
        (10,): [("The POS extract is cleaned to 39,472 rows from 41,654: 1,842 quotes and 340 duplicates removed,",
                 "SuperMart_Sales_2025_DIRTY.csv is cleaned to 3,312 rows from 3,387: 15 spacer rows and 60 duplicates removed,")],
        (15,): [("cost reconciles to the finance ledger at ₦4,212,226.50, admissions match the clinical report at 500,",
                 "FY2023 billing reconciles at ₦13,984,047,685.44, admissions match the clinical report at 10,918,")],
        (91,): [("reduced the model from 39,472 detail rows per month to 2,160 summary rows, cutting refresh from 52 to 6 seconds "
                 "and the file from 84 MB to 9 MB.",
                 "reduced the four-year ledger from 9,994 detail lines to 1,540 summary rows, cutting refresh time and file size "
                 "while every total still reconciles.")],
        (95,): [("why is the denominator 4,212?", "why is the denominator 6,607?"),
                ("Why 41 records removed?", "Why 75 rows removed?")],
        (97,): [("showing ₦4,212,226.50 against the ledger", "showing ₦13,984,047,685.44 against the ledger")],
        (100,): [("conceded that the 41 open admissions could not be assessed",
                  "conceded that the 70 admissions with no discharge date could not be assessed")],
        (106,): [("question (margin fell 1.8 points), data (39,472 clean lines), method (star schema plus twelve measures), "
                  "finding (148 SKUs below 15%), recommendation (repricing, ₦1.1m), outcome (approved).",
                  "question (margin fell 0.7 points), data (3,312 clean lines), method (star schema, 12 measures), "
                  "finding (deep discounts lost ₦47.6m), action (20% Furniture cap, ₦17.7m–₦18.0m), outcome (approved).")],
        (110,): [("Built a star-schema Excel model over 41,654 transactions, cutting reporting time from 45 minutes to 6 seconds "
                  "and identifying 148 SKUs for repricing worth ₦1.1m.",
                  "Built a star-schema Excel model over 9,994 transactions, cutting reporting time from 45 minutes to 6 seconds "
                  "and sizing a Furniture discount cap worth ₦17.7m–₦18.0m."),
                 ("across 39,472 verified records", "across 3,312 verified records"),
                 ("39,472 records, 6-second refresh", "3,312 records, 6-second refresh")],
        (114,): [("the 41 open admissions that could not be included in length of stay",
                  "the 70 admissions with no discharge date that could not be included in length of stay")],
        (117,): [("the labs double as live-test practice", "the weekly assignments double as live-test practice")],
    },
}
