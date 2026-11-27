from revlib import T

TRY = "ETEH_Week4_Section_Examples.xlsx  ›  sheet '4.14 Benchmarks & Recon'"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", "SuperMart Retail")
    return T(tag, **k)


TOPICS = [
    O("4.14.1",
      what="A **benchmark** is a fair yardstick. Compare each store with the **chain median**, using the **same "
           "period**, the **same definition** and a **size-adjusted** measure such as sales per m².",
      like="Like comparing runners by pace per km, not by total distance.",
      how=["Sales per m² = Sales / Floor_Area_sqm.", "Benchmark = MEDIAN of all stores.", "vs median = store / median − 1."],
      tip="Use the median: one big store would pull a mean up.",
      trap="Ranking stores on total sales: big stores always win.",
      case="FY2025 sales per m² for the 12 stores. Chain median is about ₦37,517.",
      steps=["Compute sales per m².", "MEDIAN across stores.", "Compare each store."],
      table=[["Store", "Sales per m²", "vs median"], ["Jos - Rayfield", "₦63,074", "+68.1%"],
             ["Lagos - Ikeja", "₦60,995", "+62.6%"], ["Abuja - Garki", "₦24,966", "−33.5%"],
             ["Ibadan - Ring Road", "₦20,576", "−45.2%"], ["Abuja - Wuse", "₦19,737", "−47.4%"]],
      so="Jos is the most efficient store, not Ikeja. Ikeja is big, but Jos sells more per m².",
      split=4.70),

    O("4.14.2",
      what="Figures in a **PDF report** must be **reconciled** to the ledger. Extract the table (Get Data > From PDF), "
           "then compare each figure with a SUMIFS on the source data.",
      like="Like checking a printed bank statement against your own records.",
      how=["Data > Get Data > From File > From PDF.", "SUMIFS the ledger for each store.", "Flag differences above rounding."],
      tip="Compare with the ledger, not with a sum of rounded printed numbers.",
      trap="Adding rounded monthly figures and calling the difference an error.",
      case="SuperMart_Monthly_Sales_Bulletin_2025.pdf (₦m, 1 decimal) vs the ledger.",
      steps=["Extract the FY column.", "Ledger SUMIFS / 1,000,000.", "Status: rounding (within ₦0.05m) or investigate."],
      table=[["Store", "Bulletin FY", "Months added", "Ledger", "Status"], ["Lagos - Ikeja", "146.4", "146.5", "146.39", "Rounding"],
             ["Benin City", "62.1", "61.9", "62.12", "Rounding"], ["Owerri", "31.6", "31.5", "31.59", "Rounding"],
             ["All 12 stores", "733.0", "n/a", "733.22", "All rounding"]],
      so="All 12 stores agree with the ledger. Added-up months can be ₦0.2m out purely from rounding.",
      split=4.20),

    O("4.14.3",
      what="Before comparing two figures, make sure they cover the **same period** and use the **same definition**. "
           "A calendar year and an April year-end are different numbers. So are gross and net.",
      like="Like comparing prices: one includes delivery and VAT, one doesn't.",
      how=["Label every figure with its period.", "Label it gross or net.", "Compare like with like only."],
      tip="Put the definition next to every number in a table.",
      trap="Comparing one company's gross with another's net.",
      case="Same company, different numbers.",
      steps=["Calendar 2025 vs April year-end.", "Online gross vs net.", "Work out the gap."],
      table=[["Measure", "Value", "Definition"], ["SuperMart FY2025", "₦733,215,255", "Jan-Dec 2025"],
             ["SuperMart Apr 2024 - Mar 2025", "₦639,113,278", "April year-end"],
             ["Online 2025 gross", "₦431,986,000", "All orders"], ["Online 2025 net", "₦386,719,300", "Delivered + Completed"],
             ["Gross ÷ net", "1.12", "Gross is about 12% higher"]],
      so="Without labels, these could all be called '2025 sales'. Label period and definition.",
      split=4.70),

    O("4.14.4",
      what="Every figure has **uncertainty**. Show rates with a **95% confidence interval**, and treat **missing** "
           "values as unknown, not zero.",
      like="Like a weather forecast: '25–28°C' is more honest than '26.5°C'.",
      how=["Margin = 1.96 × SQRT(rate × (1 − rate) / n).", "Interval = rate ± margin.", "Overlapping intervals: don't rank."],
      tip="The Week 2 dirty extract has 46 lines with no Quantity. Report units as a range.",
      trap="Ranking groups whose intervals overlap.",
      who="Lifecare Hospital",
      case="2023 readmission rate by department, with 95% intervals.",
      steps=["Rate = readmitted / admissions.", "Compute the margin.", "Compare the intervals."],
      table=[["Department", "Admissions", "Rate", "95% interval"], ["Oncology", "1,824", "13.2%", "11.7% – 14.8%"],
             ["Cardiology", "1,828", "11.4%", "10.0% – 12.9%"], ["Endocrinology", "3,682", "9.9%", "9.0% – 10.9%"],
             ["Respiratory", "1,778", "9.2%", "7.9% – 10.6%"], ["Orthopaedics & Rheumatology", "1,806", "8.3%", "7.0% – 9.5%"]],
      so="Oncology is clearly above three departments. The others overlap, so don't rank them.",
      split=4.20),

    O("4.14.5",
      what="A **benchmark memo** fits on one page: **decision needed**, **benchmark used**, **finding**, "
           "**uncertainty**, **recommendation**, **sign-off**.",
      like="Like a doctor's referral letter: short, specific, and enough to act on.",
      how=["Start with the decision.", "State the benchmark and that it was reconciled.", "Admit the limits, then recommend."],
      tip="Write the uncertainty honestly. It builds trust.",
      trap="A recommendation with no stated limits.",
      case="Memo: which stores to review for the FY2026 budget (example decision).",
      steps=["Fill each section.", "Check each number.", "Sign off."],
      table=[["Section", "Content"], ["Decision", "Which stores get the FY2026 refurbishment budget?"],
             ["Benchmark", "Chain median sales per m², FY2025 (reconciled to PDF)"],
             ["Finding", "Abuja - Wuse lowest: ₦19,737 per m², 47% below median"],
             ["Uncertainty", "One year; store age (opened 2015–2021) not adjusted"],
             ["Recommendation", "Site review of bottom 3: Wuse, Ibadan, Garki"]],
      so="One page: a decision, the evidence, its limits and a clear next step.",
      split=4.90),
]
