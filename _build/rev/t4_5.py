from revlib import T

TRY = "ETEH_Week4_Section_Examples.xlsx  ›  sheet '4.5 Forecasting'"
RT = "SuperMart Retail"
ETS_NOTE = "Excel recalculates ETS on open; your figures may differ slightly."


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", RT)
    return T(tag, **k)


TOPICS = [
    O("4.5.1",
      what="A **time series** is one measure recorded at **regular intervals** (daily, monthly, yearly). It "
           "usually mixes a **trend** (long-run direction), **seasonality** (a repeating pattern) and **noise** "
           "(random ups and downs).",
      like="Like the Lagos traffic report: busier every year (trend), worst every weekday rush hour (season), with random jams (noise).",
      how=["Sort by date with no gaps.", "Plot a line chart.", "Look for direction, repeats and noise."],
      tip="Always use evenly spaced dates (one row per month).",
      trap="Missing months, which break most forecast tools.",
      case="SuperMart monthly sales, Jan 2022 to Dec 2025: 48 months.",
      steps=["One row per month with a Sales total.", "Line chart of Sales.", "Total each year."],
      table=[["Year", "Sales", "Change"], ["2022", "₦484,247,497", "First year"], ["2023", "₦470,532,511", "-2.8%"],
             ["2024", "₦609,205,599", "+29.5%"], ["2025", "₦733,215,255", "+20.4%"]],
      so="Sales grew strongly after 2023, and every year has the same busy months.",
      split=4.70),

    O("4.5.2",
      what="The **trend** is the underlying direction once short-term ups and downs are removed. A **12-month "
           "moving average** removes the yearly pattern and shows the trend clearly.",
      like="Like judging if a child is growing by comparing birthdays, not by measuring every morning.",
      how=["MA12 = AVERAGE of the last 12 months.", "Plot it over the raw sales.", "Read the direction."],
      tip="Compare the same month year on year to avoid seasonal effects.",
      trap="Calling one good month 'a new trend'.",
      case="Is SuperMart's trend really rising? (MA12 column)",
      steps=["MA12 from Dec 2022 onwards.", "Read it at each December.", "Compare."],
      table=[["Month", "12-month average", "Reading"], ["Dec 2022", "₦40,353,958", "Baseline"], ["Dec 2023", "₦39,211,043", "Flat to slightly down"],
             ["Dec 2024", "₦50,767,133", "Rising"], ["Dec 2025", "₦61,101,271", "Still rising"]],
      so="The trend was flat in 2023 and has risen steadily since 2024.",
      split=4.70),

    O("4.5.3",
      what="**Seasonality** is a pattern that repeats every cycle, here every 12 months. A **seasonal index** = "
           "average for that month ÷ overall monthly average. Above 1 = busier than usual; below 1 = quieter.",
      like="Like school fees season or Christmas shopping: you know it's coming every year.",
      how=["AVERAGEIFS of Sales for each month number.", "Divide by the overall monthly average.", "Read above / below 1."],
      tip="You need at least 2 to 3 full years to trust the index.",
      trap="Comparing December with November and calling it growth.",
      case="Seasonal index by month, from the training years 2022–2024 (tSeasonIdx).",
      steps=["AVERAGEIFS by Month_No.", "Divide by the 36-month average.", "Rank the months."],
      table=[["Month", "Seasonal index", "Meaning"], ["December", "1.85", "85% busier than average"],
             ["November", "1.80", "80% busier"], ["September", "1.69", "69% busier"], ["February", "0.30", "70% quieter"]],
      so="December sells about 6 times as much as February: plan stock and staff for Sep, Nov and Dec.",
      split=4.70),

    O("4.5.4",
      what="A **moving average** smooths a series by averaging the **last n periods**. A short window (3) "
           "follows changes quickly but stays bumpy; a long window (12) is smooth but **lags** behind turning points.",
      like="Like judging a footballer by their last 3 matches versus their last 12.",
      how=["MA3 = AVERAGE of the latest 3 months.", "MA12 = AVERAGE of the latest 12.", "Plot both with the sales."],
      tip="Use MA12 for monthly data with yearly seasons.",
      trap="Using a moving average to forecast seasonal months.",
      case="Compare the 3- and 12-month moving averages for late 2025.",
      steps=["MA3 and MA12 columns.", "Read Nov and Dec 2025.", "Compare with actual sales."],
      table=[["Month", "Actual", "MA3", "MA12"], ["Oct 2025", "₦77,776,923", "₦76,254,821", "₦58,945,760"],
             ["Nov 2025", "₦118,447,825", "₦94,697,133", "₦62,198,748"], ["Dec 2025", "₦83,829,319", "₦93,351,356", "₦61,101,271"]],
      so="MA3 jumps with the festive months; MA12 moves gently and shows the real trend.",
      split=4.20),

    O("4.5.5",
      what="**Simple exponential smoothing** gives more weight to recent months: next forecast = alpha × latest "
           "actual + (1 − alpha) × previous forecast. **Alpha** runs from 0 (very smooth) to 1 (copies last month).",
      like="Like updating your opinion of a restaurant: each new visit counts, but you don't forget the old ones.",
      how=["Put alpha in an input cell (0.3).", "First forecast = first actual.", "=alpha*actual + (1-alpha)*previous forecast."],
      tip="Try alpha = 0.1, 0.3 and 0.6 and compare the errors.",
      trap="Using simple smoothing on strongly seasonal data.",
      case="SES forecasts with alpha = 0.3 for the end of 2025 (SES_forecast column).",
      steps=["Set alpha = 0.3.", "Fill the SES formula down.", "Compare forecasts with actuals."],
      table=[["Month", "Actual", "SES forecast (alpha 0.3)"], ["Sep 2025", "₦87,866,652", "₦52,311,440"],
             ["Nov 2025", "₦118,447,825", "₦67,417,680"], ["Dec 2025", "₦83,829,319", "₦82,726,723"]],
      so="Simple smoothing lags behind the festive peaks, so seasonal data needs a seasonal method (4.5.6).",
      split=4.70),

    O("4.5.6",
      what="**=FORECAST.ETS(target_date, values, timeline, [seasonality], [data_completion], [aggregation])** "
           "forecasts with trend and seasonality automatically. Timeline must be evenly spaced dates. Seasonality 1 "
           "or blank = detect it; 12 = yearly for monthly data.",
      like="Like a sat-nav that learns from past journeys: it handles the regular rush-hour pattern for you.",
      how=["Sales and Month columns with no gaps.", "=FORECAST.ETS(month, sales, months, 12, 1).", "Fill for each future month."],
      tip="Needs at least 2 to 3 full seasons of history.",
      trap="Text dates or missing months in the timeline.",
      case="Forecast each month of 2026 from 48 months of history (tForecast26).",
      steps=["List Jan–Dec 2026.", "FORECAST.ETS with seasonality 12.", "Total the year."],
      table=[["Month", "Forecast (about)"], ["Feb 2026 (quietest)", "₦41.4m"], ["Dec 2026 (busiest)", "₦119.6m"],
             ["FY2026 total", "₦930.9m"], ["FY2025 actual", "₦733.2m"]],
      so="ETS expects about 27% growth in 2026 with the same seasonal shape. " + ETS_NOTE,
      split=4.90),

    O("4.5.7",
      what="**=FORECAST.ETS.CONFINT(...)** gives the **± margin** around each forecast at a chosen confidence "
           "(95% by default). Lower = forecast − margin, Upper = forecast + margin. The band widens further ahead.",
      like="Like a weather forecast: '30°C, give or take 2' tomorrow, but 'give or take 5' next week.",
      how=["Same arguments as FORECAST.ETS, plus 0.95.", "Lower and Upper columns.", "Plot as a band."],
      tip="Plan stock for the Upper value in peak months, cash for the Lower value.",
      trap="Presenting the forecast without the band.",
      case="95% bands for the 2026 forecast (CI_95, Lower, Upper).",
      steps=["CONFINT for each month.", "Lower and Upper.", "Compare January with December."],
      table=[["Month", "Forecast", "95% range (about)"], ["Jan 2026", "₦44.4m", "₦21.0m to ₦67.8m"],
             ["Sep 2026", "₦111.8m", "₦87.1m to ₦136.6m"], ["Dec 2026", "₦119.6m", "₦93.5m to ₦145.8m"]],
      so="The margin grows from about ±₦23.4m in January to ±₦26.2m in December. " + ETS_NOTE,
      split=4.70),

    O("4.5.8",
      what="**=FORECAST.ETS.SEASONALITY(values, timeline)** returns the **season length** Excel detected. 12 on "
           "monthly data = yearly cycle. 0 or 1 = no pattern found. Check it before you trust the forecast.",
      like="Like asking a new employee what pattern they've noticed: their answer shows if they understood the job.",
      how=["=FORECAST.ETS.SEASONALITY(sales, months).", "Compare with what you expect.", "If it's odd, set seasonality yourself."],
      tip="If the detected length makes no business sense, set 12 yourself.",
      trap="Letting auto-detect pick a strange cycle like 5.",
      case="Does Excel detect SuperMart's yearly pattern?",
      steps=["Run SEASONALITY on 48 months.", "Read the result.", "Match it to the seasonal index."],
      table=[["Check", "Result"], ["Detected season length", "12"], ["Meaning", "Yearly cycle in monthly data"],
             ["Agrees with seasonal index?", "Yes: Sep, Nov, Dec peaks each year"]],
      so="Excel sees the same yearly pattern as the seasonal index, so seasonality 12 is right.",
      split=4.90),

    O("4.5.9",
      what="Measure accuracy **after** the fact. **MAE** = average size of the miss; **MAPE** = average % miss; "
           "**RMSE** punishes big misses; **bias** = total forecast ÷ total actual − 1 (+ over, − under).",
      like="Like checking a weather app at the end of the month: how far off was it, and does it always guess too hot?",
      how=["Error = forecast − actual for each month.", "MAPE = SUMPRODUCT(ABS(A−F)/A)/n.", "Bias = SUM(F)/SUM(A) − 1."],
      tip="Use MAPE to compare methods; check bias for a steady lean.",
      trap="Judging a method on one month.",
      case="How accurate were three methods for 2025, trained on 2022–2024?",
      steps=["Forecast 2025 with each method.", "MAPE for each.", "Bias for each."],
      table=[["Method", "MAPE", "Bias"], ["Seasonal naive (same month last year)", "24.5%", "-16.9%"],
             ["Trend × seasonal index", "23.6%", "+9.7%"], ["FORECAST.ETS", "16.4%", "-5.1%"]],
      so="ETS was most accurate and least biased, yet still missed a typical month by 16%.",
      split=4.20),

    O("4.5.10",
      what="**Backtesting**: pretend it's the past. Fit the method on earlier data, forecast a period you "
           "already know, and compare. It's the only fair test of a forecast method.",
      like="Like a mock exam before the real one: you find out how you'd score without the risk.",
      how=["Hold back the last year (2025).", "Forecast it using 2022–2024 only.", "Compare forecasts with actuals."],
      tip="Never let the test period leak into training.",
      trap="Testing on the same data used to build the model.",
      case="Backtest: train on 2022–2024, forecast each month of 2025.",
      steps=["ETS on 36 months only.", "Forecast Jan–Dec 2025.", "Compare month by month."],
      table=[["Month", "Actual", "ETS forecast (about)"], ["Feb 2025", "₦20.3m", "₦24.3m"], ["Sep 2025", "₦87.9m", "₦87.8m"],
             ["Nov 2025", "₦118.4m", "₦100.4m"], ["FY2025 total", "₦733.2m", "₦695.7m"]],
      so="ETS caught the yearly pattern but under-forecast 2025 by about 5%. " + ETS_NOTE,
      split=4.70),

    O("4.5.11",
      what="Show forecasts as a **line chart**: actuals **solid**, forecast **dashed**, and a shaded "
           "**confidence band**. **Data > Forecast Sheet** draws this automatically.",
      like="Like a road map that shows the road you've driven and the fog ahead.",
      how=["Select Month and Sales.", "Data > Forecast Sheet; end Dec 2026.", "Options: confidence 95%, seasonality 12."],
      tip="Label the band: '95% range'.",
      trap="Drawing the forecast in the same style as the actuals.",
      case="The analyst builds the 2026 forecast chart for management.",
      steps=["Data > Forecast Sheet.", "Set end date and options.", "Format and label."],
      table=[["Element", "Shows", "Why"], ["Solid line", "Actual 2022–2025", "What happened"],
             ["Dashed line", "2026 forecast", "What we expect"], ["Shaded band", "95% range", "How unsure we are"]],
      so="Readers see the history, the forecast and the uncertainty at a glance.",
      split=4.70),

    O("4.5.12",
      what="Every forecast assumes the **future behaves like the past**. New stores, price changes, currency "
           "shocks or a pandemic aren't in the history. State assumptions and use ranges and scenarios.",
      like="Like planning a road trip with last year's map: fine unless a new bridge has opened or a road closed.",
      how=["List what the forecast assumes.", "Name events that would break it.", "Plan with ranges and scenarios (4.6)."],
      tip="Re-run the forecast every month as new actuals arrive.",
      trap="Presenting a forecast as a promise.",
      case="Management wants to use the ₦930.9m 2026 forecast as the target.",
      steps=["List the assumptions.", "Name the risks.", "Agree a range."],
      table=[["Assumption", "What could break it"], ["Growth continues like 2024–2025", "Price changes or a weaker naira"],
             ["Same yearly pattern", "A new promotion calendar"], ["Same 12 stores", "Store openings or closures"],
             ["Clean, complete data", "Missing or late months"]],
      so="Use the forecast as a planning range, and revisit it monthly.",
      split=4.90),

    O("4.5.13",
      what="Choose a method by **data and purpose**: seasonal naive for a quick check, moving average to show "
           "the trend, trend × seasonal index for an explainable model, FORECAST.ETS for the best accuracy with "
           "enough history. **Backtest** to decide.",
      like="Like choosing transport: okada for short trips, a bus for routine routes, a flight for long distances.",
      how=["Check how much history you have.", "Backtest the options.", "Pick the most accurate one you can explain."],
      tip="Show the backtest result when you present the forecast.",
      trap="Choosing the fanciest method without testing it.",
      case="Which method should SuperMart use for 2026 planning?",
      steps=["Compare backtest MAPE.", "Consider who must understand it.", "Choose and explain."],
      table=[["Method", "Use when", "Watch out for"], ["Seasonal naive", "Very short history; quick check", "Ignores growth"],
             ["Moving average", "Showing the trend", "Lags; no seasonality"], ["Trend × seasonal index", "Must explain to managers", "Assumes straight-line growth"],
             ["FORECAST.ETS", "2–3+ years of monthly data", "Black box; even dates needed"]],
      so="ETS won the backtest (16.4% MAPE), so use it, with trend × season as an explainable cross-check.",
      split=4.20),
]
