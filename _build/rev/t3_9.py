from revlib import T

TRY = "ETEH_Week3_Section_Examples.xlsx  ›  sheet '3.9 Dashboard Blueprints'"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    return T(tag, **k)


TOPICS = [
    O("3.9.1",
      what="A **blueprint** is the plan for a dashboard: who it's for, which KPIs, how each is defined, and which "
           "chart shows it. The retail blueprint has **KPI cards** on top, a **store attainment bar**, a **zone "
           "combo chart** and a **monthly line**, with Zone and Month slicers.",
      like="A blueprint is like an architect's drawing: everyone agrees the plan before any building starts.",
      how=["List each KPI with its definition and source.", "Choose one visual per KPI.",
           "Model: Sales fact + Stores directory + date table."],
      tip="Write every KPI's definition next to it, so nobody argues about meaning later.",
      trap="Building the dashboard first and deciding what the KPIs mean afterwards.",
      who="SuperMart Retail",
      case="The retail director signs off the SuperMart Stores dashboard blueprint (tBlueprint1).",
      steps=["Read each KPI's definition.", "Check its current value against the data.", "Note the visual for each."],
      table=[["KPI", "Current value", "Visual"], ["Net sales FY2025", "₦733,215,255", "KPI card"],
             ["Sales growth vs FY2024", "20.4%", "KPI card with arrow"], ["Target attainment", "98.2%", "Card + bar by store"],
             ["Margin %", "12.7%", "Card + combo by zone"], ["Loss-making lines", "620", "Drill-down table"]],
      so="Sales grew 20.4% but missed target slightly, and 620 lines lost money: the dashboard must show all three.",
      split=4.90),

    O("3.9.2",
      what="The university blueprint serves the **senate and faculty deans**: KPI cards for enrolment, average "
           "score, attendance, Grade A share and at-risk count, a **bar chart by faculty**, and a **restricted** "
           "list of at-risk students. Slicers: Faculty, Department, Level.",
      like="It's like a school's end-of-term report, but for the whole university on one page.",
      how=["Define each KPI (e.g. at-risk = score below 62 or attendance below 65%).", "Show faculty comparisons as sorted bars.",
           "Restrict the named at-risk list to authorised staff."],
      tip="Keep named student data on a protected page, separate from the summary.",
      trap="Putting named at-risk students on a dashboard everyone can see.",
      who="Crestview University",
      case="The registrar reviews the Crestview progression dashboard blueprint (tBlueprint2).",
      steps=["Check each KPI's current value.", "Check the at-risk count by faculty.", "Confirm who may see the named list."],
      table=[["KPI", "Current value", "Visual"], ["Students enrolled", "6,607", "KPI card"], ["Average exam score", "67.2", "Card + bar by faculty"],
             ["Average attendance", "80.0%", "Gauge-style card"], ["Grade A share (70+)", "24.6%", "Stacked bar by faculty"],
             ["At-risk students", "901", "Card + restricted list"]],
      so="901 students need support: Social Sciences has the most (285), Arts & Education the fewest (148).",
      split=4.90),

    O("3.9.3",
      what="The hospital blueprint supports **daily operations**: KPI cards for admissions, occupancy, ALOS, "
           "readmission rate and average bill; an **occupancy line with 85% and 95% bands**; **ALOS by "
           "department**; and **average bill by HMO**. Slicers: Department, Month, HMO.",
      like="It's like a hospital ward's whiteboard, but for the whole hospital and updated automatically.",
      how=["Define each KPI with its source (EMR, bed census, billing).", "Draw occupancy as a line with target bands.",
           "Use RAG status (with text labels) for readmissions."],
      tip="Show the 85% and 95% bands so managers see risk at a glance.",
      trap="Showing occupancy without the safe-limit bands.",
      who="Lifecare Hospital",
      case="The medical director reviews the Lifecare operations dashboard blueprint for 2023 (tBlueprint3).",
      steps=["Check each 2023 KPI value.", "Check the occupancy line's bands.", "Confirm each KPI's source system."],
      table=[["KPI (2023)", "Current value", "Visual"], ["Admissions", "10,918", "Card + monthly line"],
             ["Average occupancy", "89.5%", "Line with 85% / 95% bands"], ["ALOS", "15.5 days", "Bar by department"],
             ["30-day readmission rate", "10.3%", "Card with RAG + text"], ["Average bill", "₦1,280,825", "Bar by HMO"]],
      so="Occupancy averaged 89.5%, above the 85% comfort line, so bed pressure is the headline for managers.",
      split=4.90),
]
