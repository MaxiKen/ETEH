from revlib import T

TRY = "ETEH_Week4_Section_Examples.xlsx  ›  sheet '4.6 What-If & Solver'"
HT = "EkoSuites Hotel"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", HT)
    return T(tag, **k)


TOPICS = [
    O("4.6.1",
      what="**What-if analysis** changes the **inputs** of a model to see how the **outputs** respond, without "
           "touching real data. Build a small model first: yellow input cells, green formula outputs.",
      like="Like a recipe calculator: change the number of guests and the shopping list updates.",
      how=["Put every assumption in its own input cell.", "Outputs are formulas that use those cells.", "Then apply Goal Seek, data tables, scenarios or Solver."],
      tip="Never type a number inside an output formula.",
      trap="Hard-coding assumptions, so nothing updates.",
      case="A 2025 profit model for EkoSuites (modelled as 320 rooms), built from 2024 honoured bookings.",
      steps=["2024 actuals: 55,078 room-nights; ADR ₦102,924; occupancy 47.0%.", "Inputs: 47%, ₦102,900, ₦28,000 variable cost, ₦3.6bn fixed (assumptions).", "Profit = rooms × 365 × occupancy × (ADR − variable) − fixed."],
      table=[["Output", "Base result"], ["Room-nights sold", "54,896"], ["Room revenue", "₦5,648,798,400"],
             ["Operating profit", "₦511,710,400"], ["RevPAR (revenue per available room)", "₦48,363"]],
      so="One clear model, with every assumption visible, ready for every what-if tool.",
      split=4.90),

    O("4.6.2",
      what="**Goal Seek** finds the **one input** that makes a formula hit a **target**. Data > What-If Analysis "
           "> Goal Seek: Set cell (the output), To value (the target), By changing cell (one input).",
      like="Like asking 'how fast must I drive to arrive by 9 am?' and letting the sat-nav work it out.",
      how=["Data > What-If Analysis > Goal Seek.", "Set cell = profit; To value = 1,000,000,000.", "By changing = the ADR input."],
      tip="Check the answer with algebra or a second method.",
      trap="Pointing 'Set cell' at a typed number instead of a formula.",
      case="What average daily rate (ADR) gives ₦1bn operating profit, with occupancy at 47%?",
      steps=["Goal Seek on profit, target ₦1bn, changing ADR.", "Read the new ADR.", "Check: (target + fixed) ÷ room-nights + variable cost."],
      table=[["Measure", "Result"], ["Current ADR input", "₦102,900"], ["ADR needed for ₦1bn", "₦111,795"],
             ["Increase needed", "₦8,895 per night"], ["Algebra check", "Matches Goal Seek"]],
      so="EkoSuites needs about ₦8,900 more per room-night. The next question: will guests pay it?",
      split=4.90),

    O("4.6.3",
      what="Goal Seek changes **only one input**, **can't respect limits** (like 'occupancy at most 100%'), and gives "
           "only one answer. For several inputs or limits, use Solver (4.6.9).",
      like="Like a calculator that tells you the answer, even when the answer is impossible.",
      how=["Check the answer is realistic.", "Use Solver when there are limits.", "Use data tables to see many answers at once."],
      tip="Always ask: could this input really take this value?",
      trap="Reporting an impossible Goal Seek answer.",
      case="What occupancy would deliver different profit targets at ₦102,900 ADR?",
      steps=["Goal Seek: profit to ₦1bn, changing occupancy.", "Repeat for ₦5bn and ₦6bn.", "Check each answer is possible."],
      table=[["Profit target", "Occupancy Goal Seek returns", "Realistic?"], ["₦1bn", "52.6%", "Yes"],
             ["₦5bn", "98.3%", "Barely"], ["₦6bn", "109.7%", "Impossible"]],
      so="Goal Seek happily returns 109.7% occupancy: it doesn't know a hotel can't be more than full.",
      split=4.70),

    O("4.6.4",
      what="A **one-variable data table** recalculates an output for a **list of values of one input**. Data > "
           "What-If Analysis > Data Table. It shows many answers at once, not just one.",
      like="Like a price list that shows the cost for 1, 2, 5 and 10 items.",
      how=["List input values down a column.", "Put a link to the output formula above and right of the list.",
           "Select both; Data Table; Column input cell = the input."],
      tip="Data tables recalculate live; set calculation to 'Automatic except tables' if the workbook slows.",
      trap="Picking the wrong cell as the column input.",
      case="Profit at ADR ₦100,000 for occupancy from 40% to 60%.",
      steps=["List 40%, 45%, 50%, 55%, 60%.", "Link the profit formula.", "Data Table with Column input = occupancy."],
      table=[["Occupancy", "Profit at ₦100k ADR"], ["40%", "-₦236,160,000"], ["45%", "₦184,320,000"],
             ["50%", "₦604,800,000"], ["55%", "₦1,025,280,000"], ["60%", "₦1,445,760,000"]],
      so="Each 5 points of occupancy adds about ₦420m of profit at a ₦100k ADR.",
      split=4.90),

    O("4.6.5",
      what="A **two-variable data table** varies **two inputs at once**: one across the top row, one down the "
           "first column. The output formula sits in the top-left corner. The result is a grid of every combination.",
      like="Like a bus fare chart: find your starting stop down the side and your destination across the top.",
      how=["ADR values across the top, occupancy down the side.", "Profit formula link in the corner.",
           "Data Table: Row input = ADR, Column input = occupancy."],
      tip="Keep the grid small (about 6 × 7) so it reads at a glance.",
      trap="Swapping the row and column input cells.",
      case="Profit for occupancy 40%–70% × ADR ₦80k–₦130k.",
      steps=["Lay out the grid.", "Run Data Table.", "Read a few cells."],
      table=[["Occupancy", "ADR ₦80k", "ADR ₦100k", "ADR ₦120k"], ["40%", "-₦1,170,560,000", "-₦236,160,000", "₦698,240,000"],
             ["50%", "-₦563,200,000", "₦604,800,000", "₦1,772,800,000"], ["60%", "₦44,160,000", "₦1,445,760,000", "₦2,847,360,000"]],
      so="Profit ranges from a ₦1.2bn loss to a ₦2.8bn profit across these combinations.",
      split=4.20),

    O("4.6.6",
      what="Read a data table like a **heat map**: add a red-to-green colour scale, find the **break-even edge** "
           "(where loss turns to profit), and note which input moves profit more.",
      like="Like a weather map: colours show at a glance where it's safe and where it's stormy.",
      how=["Home > Conditional Formatting > Color Scales.", "Trace where values cross zero.", "Compare a step in each input."],
      tip="Report the break-even edge, not every cell.",
      trap="Presenting a grid of numbers with no reading.",
      case="Where is break-even in the occupancy × ADR grid?",
      steps=["Colour the grid.", "Find the first profitable cell in each ADR column.", "Compare steps."],
      table=[["ADR", "Lowest occupancy in the grid with profit", "Profit there"], ["₦80k", "60%", "₦44,160,000"],
             ["₦90k", "50%", "₦20,800,000"], ["₦100k", "45%", "₦184,320,000"], ["₦110k", "40%", "₦231,040,000"]],
      so="At 47% occupancy, break-even ADR is about ₦93,600 (₦3.6bn ÷ 54,896 room-nights + ₦28,000).",
      split=4.20),

    O("4.6.7",
      what="**Scenario Manager** (Data > What-If Analysis) stores **named sets of inputs**, such as Pessimistic, "
           "Base and Optimistic, and swaps them into the model on demand.",
      like="Like saved profiles on a phone: work, home and travel, each with different settings.",
      how=["Data > What-If Analysis > Scenario Manager > Add.", "Name it; choose the changing cells; enter values.", "Show to swap the values in."],
      tip="Change several inputs together; real scenarios move more than one thing.",
      trap="Scenarios that change only one input (use a data table instead).",
      case="Three 2025 scenarios for EkoSuites (tScen).",
      steps=["Pessimistic: 41%, ₦94,668, fixed ₦3.78bn.", "Base: 47%, ₦102,900, ₦3.6bn.", "Optimistic: 53%, ₦111,132, ₦3.6bn."],
      table=[["Scenario", "Revenue", "Profit"], ["Pessimistic", "₦4,533,461,184", "-₦587,402,816"],
             ["Base", "₦5,648,798,400", "₦511,710,400"], ["Optimistic", "₦6,879,515,328", "₦1,546,203,328"]],
      so="The bad case is a ₦587m loss; the good case is a ₦1.5bn profit.",
      split=4.70),

    O("4.6.8",
      what="**Scenario Manager > Summary** builds a new sheet comparing every scenario's inputs and results side "
           "by side. Name your cells first, so the report shows names instead of cell addresses.",
      like="Like a comparison chart on a shopping site: the options in columns, the features in rows.",
      how=["Name the input and output cells.", "Scenario Manager > Summary.", "Choose the result cells (profit)."],
      tip="The summary is static: rerun it after changing the model.",
      trap="Reports full of $E$15-style addresses nobody understands.",
      case="Management wants one table comparing the three scenarios.",
      steps=["Create the summary.", "Read profit for each scenario.", "Compute the spread."],
      table=[["Measure", "Result"], ["Lowest profit (Pessimistic)", "-₦587,402,816"], ["Highest profit (Optimistic)", "₦1,546,203,328"],
             ["Spread to plan for", "₦2,133,606,144"]],
      so="Management must plan for a ₦2.1bn swing between the bad and good cases.",
      split=4.90),

    O("4.6.9",
      what="**Solver** finds the best answer when **several inputs** must change within **limits**. Enable it "
           "once: File > Options > Add-ins > Excel Add-ins > Go, tick **Solver Add-in**. It appears on the Data tab.",
      like="Like a planner who juggles several bookings at once while respecting every rule.",
      how=["File > Options > Add-ins.", "Manage Excel Add-ins > Go.", "Tick Solver Add-in > OK."],
      tip="On a Mac: Tools > Excel Add-ins.",
      trap="Using Goal Seek for a problem with limits.",
      case="EkoSuites wants to split its 320 rooms across 4 market segments to earn the most each night.",
      steps=["Enable Solver.", "Open Data > Solver.", "Identify the three parts of the problem."],
      table=[["Solver part", "In this example"], ["Objective", "Maximise expected net value per night"],
             ["Changing cells", "Rooms allocated to each segment"], ["Constraints", "Total <= 320; within demand; OTA <= 50%"]],
      so="One setting unlocks a tool that solves problems Goal Seek can't.",
      split=4.90),

    O("4.6.10",
      what="A Solver model needs an **objective** (a formula to maximise or minimise), **decision cells** (inputs "
           "Solver may change) and **constraints** (limits). Each segment's **net value** = ADR × (1 − cancel rate) × (1 − commission).",
      like="Like packing a suitcase: the most value inside, without breaking the weight limit.",
      how=["Net value per room for each segment.", "Nightly value = rooms × net value; objective = SUM.", "Write each limit as a constraint."],
      tip="Check the model by hand with a simple allocation first.",
      trap="An objective cell that doesn't depend on the decision cells.",
      case="Net value per room for each segment (tSegments).",
      steps=["ADR, cancel rate and commission per segment.", "Net_value_per_room formula.", "Compare segments."],
      table=[["Segment", "Headline ADR", "Net value per room"], ["Corporate", "₦84,800", "₦74,919"],
             ["Airline Crew", "₦100,704", "₦70,896"], ["Online Travel Agent", "₦114,340", "₦56,872"],
             ["Walk-in / Travel Agent", "₦92,882", "₦54,675"]],
      so="OTA has the highest ADR but, after about 39% cancellations (2024) and 18% commission, earns far less than Corporate.",
      split=4.70),

    O("4.6.11",
      what="Solver offers **Simplex LP** for straight-line (linear) models, **GRG Nonlinear** for smooth curves, "
           "and **Evolutionary** for rough or jumpy models. Choose the method, run, then **Keep Solver Solution** "
           "and ask for an Answer report.",
      like="Like choosing a route planner for roads, rail or a mix: the right tool for the shape of the problem.",
      how=["Data > Solver: objective, Max, changing cells.", "Add constraints, including 'integer'.", "Method Simplex LP > Solve > Keep."],
      tip="If Solver says 'not linear', switch to GRG Nonlinear.",
      trap="Accepting a result without checking every constraint holds.",
      case="Solve the room allocation with Simplex LP.",
      steps=["Objective: max total nightly value.", "Constraints: min, max demand, integer, <= 320, OTA <= 160.", "Solve and read the allocation."],
      table=[["Segment", "Rooms allocated", "Why"], ["Corporate", "70", "All demand: highest value"],
             ["Airline Crew", "10", "Minimum and maximum are both 10"], ["Online Travel Agent", "160", "Stopped at the 50% cap"],
             ["Walk-in / Travel Agent", "80", "Takes the remaining rooms"]],
      so="Best plan: about ₦19.4m expected net value per night, filling Corporate first.",
      split=4.70),

    O("4.6.12",
      what="Document every what-if analysis: the **baseline inputs** and their source, the **assumptions**, the "
           "**tools** used, the **results** and the **limits**. Others must be able to repeat and challenge it.",
      like="Like a recipe card: ingredients, method and notes, so someone else can cook it too.",
      how=["List inputs with source and date.", "List assumptions and tools used.", "State the limits."],
      tip="Keep the notes on the same sheet as the model.",
      trap="A Solver answer with no record of the constraints used.",
      case="The analyst documents the EkoSuites revenue plan.",
      steps=["Inputs and source.", "Assumptions and tools.", "Limits."],
      table=[["Item", "Recorded"], ["Inputs", "2024 honoured bookings from D_Hotel"],
             ["Assumptions", "320 rooms, ₦28,000 variable cost, ₦3.6bn fixed"], ["Tools used", "Goal Seek, data table, 3 scenarios, Solver"],
             ["Limits", "Annual averages; ignores seasonality"]],
      so="Anyone can rerun or challenge the plan, because nothing is hidden.",
      split=4.90),
]
