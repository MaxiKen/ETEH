from revlib import T

W, LC, CV, MS = "SuperMart Retail", "Lifecare Hospital", "Crestview University", "Multi-Sector"
TRY = "ETEH_Week2_Section_Examples.xlsx  ›  sheets '2.11 PivotCharts' and '2.10 Live Pivots'"


def C(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", W)
    return T(tag, **k)


ZONE_2025 = [["Zone (2025)", "Sales"], ["South-West", "₦250.1m"], ["South-South", "₦213.1m"],
             ["North-Central", "₦147.1m"], ["South-East", "₦122.9m"]]

TOPICS = [
    C("2.11.1",
      what="A **PivotChart** is a chart that is **linked to a PivotTable**. It draws whatever the pivot shows. "
           "Change the pivot's fields or filters and the chart changes too, with no need to reselect any data.",
      like="A PivotChart is like a shadow of the PivotTable: wherever the table moves, the shadow follows.",
      how=["Click inside a PivotTable.", "PivotTable Analyze > PivotChart.", "Choose a chart type and click OK."],
      tip="Build the pivot first, then chart it.",
      trap="Copying pivot numbers into a normal chart; it won't follow later changes.",
      case="Head Office wants Pivot 1 (sales by zone and year, 2.10 Live Pivots) as a chart.",
      steps=["Click inside Pivot 1.", "PivotTable Analyze > PivotChart > Clustered Column.",
             "Check the bars match the pivot's numbers."],
      table=[["Zone", "2022", "2025"], ["South-West", "₦147.9m", "₦250.1m"], ["South-South", "₦128.7m", "₦213.1m"],
             ["North-Central", "₦103.8m", "₦147.1m"], ["South-East", "₦103.8m", "₦122.9m"]],
      so="The chart shows the same numbers as the pivot, as bars. South-West's growth is obvious at a glance.",
      split=4.70),

    C("2.11.2",
      what="Charts make a summary **easy to read in one glance**. Trends, gaps and leaders that take a minute to "
           "find in a table become obvious in a picture. Because PivotCharts follow the pivot, they stay correct "
           "after every refresh.",
      like="A table is like a list of match scores; a chart is like the league table on TV: you instantly see "
           "who is on top.",
      how=["Decide the one message the chart should show.", "Pick the chart type that shows it best.",
           "Give it a title that states the message."],
      tip="Write the message as the chart title, e.g. 'Q4 is our peak quarter'.",
      trap="Making a chart with no clear message.",
      case="Pivot 4 on the Live Pivots sheet shows sales by year and quarter. Is there a seasonal pattern?",
      steps=["Read the 16 numbers in the pivot table.", "Insert a PivotChart (clustered column).",
             "Look at which bar is tallest in each year."],
      table=[["Year", "Q4 sales", "Q4 share of year"], ["2022", "₦179.6m", "37.1%"], ["2023", "₦182.3m", "38.7%"],
             ["2024", "₦236.1m", "38.8%"], ["2025", "₦280.1m", "38.2%"]],
      so="In the chart, Q4 towers over the other quarters every year. The table hid that pattern in 16 numbers.",
      split=4.20),

    C("2.11.3",
      what="A **normal chart** points at fixed cells; if the data grows, you must reselect the range. A "
           "**PivotChart** points at a PivotTable, so it grows and shrinks with it, and has **field buttons** for "
           "filtering on the chart itself.",
      like="A normal chart is like a printed map; a PivotChart is like a phone map that updates as roads change.",
      how=["Use a normal chart for a small, fixed table.", "Use a PivotChart for data that grows or needs filtering.",
           "Look for field buttons to tell them apart."],
      tip="If the data will grow, use a PivotChart.",
      trap="Charting a fixed range and forgetting to extend it when new data arrives.",
      case="A chart of FY2025 sales by month must keep working as each month's file arrives.",
      steps=["Draw a normal chart on Jan–Nov.", "Add December and refresh.", "Compare a PivotChart built on the same data."],
      table=[["After December is added", "Normal chart", "PivotChart"], ["Bars shown", "11 (Jan–Nov)", "12"],
             ["Total behind the bars", "₦649.4m", "₦733.2m"], ["Field buttons", "No", "Yes"]],
      so="The normal chart missed December's ₦83.8m until someone reselected its range. The PivotChart didn't.",
      split=4.20),

    C("2.11.4",
      what="Three ways to create a PivotChart: (1) click inside a pivot and choose **PivotTable Analyze > "
           "PivotChart**; (2) click inside a Table and choose **Insert > PivotChart**, which creates the pivot too; "
           "(3) use **Insert > PivotChart & PivotTable** to get both side by side.",
      like="It's like ordering a meal and a drink: order them separately, or as a combo that brings both.",
      how=["Route 1: from an existing pivot.", "Route 2: Insert > PivotChart from a Table.",
           "Then drag fields as you would for a pivot."],
      tip="Route 1 is clearest for beginners: pivot first, then chart.",
      trap="Starting from a normal range instead of a Table.",
      case="The analyst creates a chart of FY2025 sales by zone.",
      steps=["Click inside Pivot 1 (2.10 Live Pivots).", "PivotTable Analyze > PivotChart > Clustered Column.",
             "Use the Year field button to keep only 2025."],
      table=ZONE_2025,
      so="Four bars, tallest to shortest from South-West to South-East, created in three clicks.",
      split=4.90),

    C("2.11.5",
      what="A PivotChart screen has **five parts**: (1) the **ribbon tabs** PivotChart Analyze, Design and Format; "
           "(2) the **filter pane**; (3) the **plot area** with bars and legend; (4) **field buttons** on the "
           "chart; (5) the **field list**, shared with the pivot.",
      like="It's like a car dashboard: steering (ribbon), mirrors (filters), the road ahead (plot area), quick "
           "switches (field buttons) and the gearbox (field list).",
      how=["Click the chart to show the PivotChart ribbon tabs.", "Use the field buttons for quick filters and sorts.",
           "Use the field list to change what the chart shows."],
      tip="Field buttons on the chart prove it is a PivotChart.",
      trap="Hiding field buttons and then forgetting a filter is applied.",
      case="The analyst explores the zone × year PivotChart from Pivot 1.",
      steps=["Click the chart: find the three ribbon tabs.", "Open the field list: Zone in Axis, Year in Legend.",
             "Click the Zone button on the chart and sort by value."],
      table=[["Part", "What it does here"], ["Ribbon tabs", "Refresh, Change Chart Type"],
             ["Field list", "Axis = Zone, Legend = Year"], ["Field buttons", "Zone and Year drop-downs"],
             ["Plot area", "16 bars, one colour per year"]],
      so="Knowing which part does what makes the chart predictable instead of confusing.",
      split=4.90),

    C("2.11.6",
      what="The PivotChart field list uses different names for the same areas: **Axis (Categories)** = Rows, "
           "**Legend (Series)** = Columns, **Values** = Values, **Filters** = Filters. Moving a field in one moves "
           "it in the other.",
      like="It's like a person with two names, one at home and one at work: different labels, same person.",
      how=["Axis: what goes along the bottom.", "Legend: what gets a different colour.", "Values: the height of the bars."],
      tip="Put time or categories on the Axis; a few groups in the Legend.",
      trap="Putting a field with many items in the Legend, giving dozens of colours.",
      case="The analyst sets up the zone × year chart from Pivot 1.",
      steps=["Axis (Categories) = Zone.", "Legend (Series) = Year.", "Values = Sum of Sales."],
      table=[["Chart area", "Pivot area", "Field used"], ["Axis (Categories)", "Rows", "Zone"],
             ["Legend (Series)", "Columns", "Year"], ["Values", "Values", "Sum of Sales"],
             ["Filters", "Filters", "(none)"]],
      so="Four zones along the bottom, four colours for the years: 16 bars that match the pivot's 16 cells.",
      split=4.20),

    C("2.11.7",
      what="A PivotChart and its PivotTable are linked **both ways**. Filter the chart and the table filters; sort "
           "the table and the chart reorders; refresh one and both update.",
      like="They are like two ends of a see-saw: move one end and the other moves too.",
      how=["Filter using the chart's field button.", "Look at the pivot: it is filtered too.",
           "Clear the filter in either place to reset both."],
      tip="Place the chart next to its pivot so readers see both change.",
      trap="Being surprised when filtering the chart changes the pivot.",
      case="The Lagos regional manager wants only South-West on the zone × year chart.",
      steps=["Click the Zone field button on the chart.", "Tick only South-West.", "Check Pivot 1: it now shows one row."],
      table=[["South-West", "Sales"], ["2022", "₦147.9m"], ["2023", "₦140.0m"], ["2024", "₦187.5m"],
             ["2025", "₦250.1m"]],
      so="One click on the chart filtered both. South-West dipped in 2023, then grew strongly.",
      split=4.90),

    C("2.11.8",
      what="Format a PivotChart with the **Design** tab (styles, layout) and the **Format** tab (colours). "
           "Right-click the value axis > **Format Axis** to set number format and bounds. For column and bar "
           "charts, start the value axis at **0**.",
      like="Formatting a chart is like setting a table: plates in the right place, nothing distracting, so the "
           "meal is the focus.",
      how=["Give the chart a title that states the message.", "Format the axis in millions: ₦#,##0,,\"m\".",
           "Remove clutter: extra gridlines, 3-D effects."],
      tip="Start bar and column axes at zero.",
      trap="Starting the axis at ₦100m, which makes small gaps look huge.",
      case="The FY2025 zone chart goes into a board pack.",
      steps=["Title: 'South-West leads FY2025 sales'.", "Axis number format: ₦#,##0,,\"m\"; minimum 0.",
             "Delete the legend (only one series)."],
      table=[["Setting", "Value used"], ["Title", "South-West leads FY2025 sales"], ["Axis format", "₦#,##0,,\"m\""],
             ["Axis minimum", "0"], ["Tallest bar", "South-West ₦250m"]],
      so="With a zero baseline, South-West (₦250m) looks about twice South-East (₦123m), which is true.",
      split=4.90),

    C("2.11.9",
      what="Use **PivotChart Analyze > Change Chart Type** to switch charts. **Column or line** suit time. **Bar** "
           "suits ranking many items. **Pie** works only for a few parts of one whole.",
      like="Choosing a chart is like choosing a vehicle: a bike for a short trip, a bus for many people. The "
           "wrong one makes the trip harder.",
      how=["Ask: is this over time, a ranking, or parts of a whole?", "Pick line/column, bar, or pie accordingly.",
           "Change Chart Type and check it reads well."],
      tip="Use a sorted bar chart for rankings.",
      trap="Using a pie for 12 stores; the slices are impossible to compare.",
      who=MS,
      case="The 2.11 sheet charts three datasets, each with a different shape of data.",
      steps=["SuperMart zone sales by year: clustered column.", "Lifecare admissions by department: bar chart.",
             "Crestview grade mix (3 grades used): pie."],
      table=[["Data shape", "Chart used", "Why"], ["Zones over 4 years", "Clustered column", "Compare and see trend"],
             ["5 departments ranked", "Bar", "Long labels, easy ranking"],
             ["3 grades of one whole", "Pie", "Few parts, adds to 100%"],
             ["12 stores' share", "Sorted bar, not pie", "Too many slices"]],
      so="The chart type follows the question. A pie only works with a few parts.",
      split=4.20),

    C("2.11.10",
      what="You can filter a PivotChart in three places: the **filter pane**, the **field buttons** on the chart, "
           "or a shared **Slicer/Timeline**. To sort, use the axis field button > **More Sort Options** > by value.",
      like="Filtering a chart is like zooming a camera: the picture is the same scene, focused on what matters.",
      how=["Field button: tick the items to keep.", "Sort by value from the axis button.", "Use a Slicer to control several charts at once."],
      tip="Sort bars by value, largest first, unless the axis is time.",
      trap="Sorting months by value, which scrambles the timeline.",
      case="Head Office wants the FY2025 zone chart, sorted from largest to smallest.",
      steps=["Year field button: keep only 2025.", "Zone button > More Sort Options > Descending by Sum of Sales.",
             "Add a Zone slicer to filter other charts too."],
      table=ZONE_2025,
      so="Sorted bars read like a ranking: South-West first, South-East last.",
      split=4.90),

    C("2.11.11",
      what="Each dataset suits a different chart. The 2.11 sheet has one example per sector: **SuperMart** zone "
           "sales by year (column), **Lifecare** admissions by department (bar), and **Crestview** grade "
           "distribution (pie).",
      like="Different datasets are like different stories; each needs the right picture to be told well.",
      how=["Pick the one question each chart answers.", "Choose the chart type for that question.",
           "Title the chart with the answer."],
      tip="One chart, one message.",
      trap="Squeezing several messages into one busy chart.",
      who=MS,
      case="The trainer shows one PivotChart per bootcamp dataset (2.11 sheet).",
      steps=["SuperMart: clustered column of sales by zone and year.",
             "Lifecare: bar chart of 2019–2024 admissions by department.", "Crestview: pie of grades."],
      table=[["Dataset", "Chart", "What it shows"], ["SuperMart", "Column", "South-West grew to ₦250.1m in 2025"],
             ["Lifecare", "Bar", "Endocrinology 18,324 admissions, twice any other"],
             ["Crestview", "Pie", "B 74.4%, A 24.6%, C 1.0%"]],
      so="Each chart answers one question at a glance. Lifecare's stays average about 15.5 days in every department.",
      split=4.20),
]
