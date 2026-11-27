from revlib import T

TRY = "ETEH_Week4_Section_Examples.xlsx  ›  sheet '4.7 Data Storytelling'"
RT = "SuperMart Retail"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", RT)
    return T(tag, **k)


TOPICS = [
    O("4.7.1",
      what="Analysis changes nothing until someone **acts** on it, and people act on what they **understand and "
           "remember**. A data story turns numbers into a clear message, the reason it matters and what to do next.",
      like="Like a good news report: what happened, why it matters to you, and what happens next.",
      how=["Find the one message that matters.", "Show only the evidence that supports it.", "End with a decision or action."],
      tip="If you can't say the message in one sentence, you're not ready to present.",
      trap="Showing every chart you made.",
      case="The analyst must present SuperMart's FY2025 results to the board.",
      steps=["Compare a 'data dump' with a story.", "Pick the main message.", "Choose the action."],
      table=[["Without a story", "With a story"], ["12 store charts, no conclusion", "One headline: grew 20%, 7 stores missed target"],
             ["'Sales were ₦733,215,255'", "'Sales grew 20% to ₦733m'"], ["Board asks 'so what?'", "Board decides on Wuse and Owerri"]],
      so="The same numbers lead to a decision only when they're told as a story.",
      split=4.90),

    O("4.7.2",
      what="Use **three acts**: **Situation** (what's normal or expected), **Complication** (what changed or went "
           "wrong), **Resolution** (what we should do). Each act gets one clear piece of evidence.",
      like="Like a Nollywood plot: everyday life, then trouble, then how it's solved.",
      how=["Act 1: the situation in one line.", "Act 2: the problem, with one chart.", "Act 3: the action, owner and date."],
      tip="Most of your time belongs in Act 2: that's where the insight is.",
      trap="Ending on a chart with no Act 3.",
      case="The analyst plans the FY2025 story.",
      steps=["Write the situation.", "Write the complication.", "Write the resolution."],
      table=[["Act", "Message", "Evidence"], ["1. Situation", "Sales grew 20% in FY2025; North-Central was flat", "Monthly sales line"],
             ["2. Complication", "7 of 12 stores missed target", "Attainment bar chart"],
             ["3. Resolution", "Review Wuse and Owerri plans; learn from Lekki and Jos", "Action table"]],
      so="Three acts give the board context, the problem and a clear ask.",
      split=4.20),

    O("4.7.3",
      what="The **headline** is one sentence giving the **finding and its consequence**. Build it with a formula "
           "(using TEXT and cell values), so it can never disagree with the numbers.",
      like="Like a newspaper front page: the main story in one line you can't miss.",
      how=["Find the key numbers (growth, missed targets, worst gap).", "Join them with & and TEXT().", "Read it aloud: is it clear?"],
      tip="TEXT(value, \"0%\") formats numbers inside a sentence.",
      trap="Typing the headline by hand, so it goes stale when data changes.",
      case="Write the FY2025 headline from the store table (tStory).",
      steps=["Growth: SUM(Sales) ÷ ₦609,205,599 − 1.", "COUNTIFS(Attainment, \"<1\").", "INDEX/MATCH for the largest gap."],
      table=[["Part", "Value"], ["Growth vs FY2024", "20%"], ["FY2025 sales", "₦733m"], ["Stores below target", "7 of 12"],
             ["Largest gap", "Abuja - Wuse, ₦8.6m"]],
      so="'FY2025 sales grew 20% to ₦733m, but 7 of 12 stores missed target; Abuja - Wuse has the largest gap.'",
      split=4.90),

    O("4.7.4",
      what="Arrange the message in **three tiers**: **one headline**, **about three supporting points**, then the "
           "**detailed evidence**. Readers who stop at any level still get the main idea.",
      like="Like a pyramid: the point on top, the support underneath, the foundation at the bottom.",
      how=["Tier 1: the headline.", "Tier 2: three supporting points.", "Tier 3: evidence in charts, tables or an appendix."],
      tip="Each supporting point should back up the headline directly.",
      trap="Seven 'key points' with no ranking.",
      case="The analyst builds the message structure for the board pack.",
      steps=["Headline.", "Three supporting points.", "Evidence for each."],
      table=[["Tier", "Content"], ["Headline", "Grew 20% to ₦733m; 7 of 12 stores missed target"],
             ["Point 1", "Missed stores are ₦34.9m short of target in total"], ["Point 2", "Wuse (₦8.6m) and Owerri (₦7.4m) have the largest gaps"],
             ["Point 3", "Lekki (112%) and Jos (109%) beat target the most"]],
      so="A board member who reads only the headline still knows what matters.",
      split=4.90),

    O("4.7.5",
      what="Different audiences need different evidence. **Executives** want the headline, a few numbers and the "
           "decision. **Managers** want their own area against target. **Staff** want the detail they can act on.",
      like="Like giving directions: a visitor needs landmarks, a taxi driver needs street names.",
      how=["List the audiences.", "Ask what decision each one makes.", "Give each only the evidence they need."],
      tip="The same analysis can produce three different outputs.",
      trap="Giving the board the store managers' line-level detail.",
      case="Three audiences for the FY2025 results.",
      steps=["Board.", "Regional managers.", "Store managers."],
      table=[["Audience", "Evidence they need", "Format"], ["Board", "Headline, 3 numbers, decision needed", "1 slide"],
             ["Regional managers", "Their stores vs target, monthly trend", "Dashboard (3.8)"],
             ["Store managers", "Line-level detail, loss-making products", "Filtered table"]],
      so="Each group gets what it needs to act, and nothing that distracts.",
      split=4.20),

    O("4.7.6",
      what="Replace **label titles** ('Store attainment FY2025') with **finding titles** that state the message "
           "('7 of 12 stores missed their FY2025 target'). The reader gets the point before studying the chart.",
      like="Like a text message that says 'Running 10 minutes late' instead of 'Update'.",
      how=["Write what the chart shows in one sentence.", "Use it as the chart title.", "Check the chart really supports it."],
      tip="Link the title to a cell with a formula so it stays current.",
      trap="A finding title the chart doesn't support.",
      case="Retitle the store attainment bar chart (tAttain).",
      steps=["Sort stores by attainment.", "Colour bars below 100% red.", "Write the finding title."],
      table=[["Title type", "Example"], ["Label title (weak)", "Store attainment FY2025"],
             ["Finding title (strong)", "7 of 12 stores missed their FY2025 target"], ["Lowest store", "Owerri - New Owerri, 81.0%"],
             ["Highest store", "Lagos - Lekki, 112.4%"]],
      so="The strong title tells the story; the chart proves it.",
      split=4.90),

    O("4.7.7",
      what="**Annotations** point at the exact feature you're talking about: a target line, a callout on the "
           "worst bar, a label on the peak month. Use **one to three** per chart.",
      like="Like a tour guide pointing: 'look at that building on the left'.",
      how=["Add a target line at 100%.", "Colour the bars that matter.", "Add a short text callout to one or two of them."],
      tip="Use colour for meaning (red = below target), not decoration.",
      trap="So many labels that nothing stands out.",
      case="Annotate the attainment chart for the board.",
      steps=["Target line at 100%.", "Red for the 7 stores below target.", "Callouts on Owerri and Wuse."],
      table=[["Annotation", "Placement"], ["Target line", "100% across all bars"], ["Red bars", "The 7 stores below target"],
             ["Callout 1", "Owerri: 81.0% (lowest)"], ["Callout 2", "Wuse: ₦8.6m short (largest gap)"]],
      so="The eye goes straight to the two stores the board must discuss.",
      split=4.90),

    O("4.7.8",
      what="Never **cut the value axis** to exaggerate differences, avoid **3D** and odd shapes, and keep the same "
           "scale when comparing charts. Bar and column charts must start at **zero**.",
      like="Like a fairground mirror: the same person, but it makes them look twice as tall.",
      how=["Format Axis > Minimum = 0 for bars and columns.", "Use flat 2D charts.", "Same scale on charts that are compared."],
      tip="Line charts may zoom in, but say so on the axis.",
      trap="Letting Excel auto-pick an axis that doesn't start at zero.",
      case="FY2025 sales by zone, charted two ways (tZoneStory).",
      steps=["Column chart with the axis starting at ₦120m.", "Same chart with the axis at zero.", "Compare South-West with South-East."],
      table=[["Zone", "Sales", "Looks like (axis ₦120m)"], ["South-West", "₦250,128,366", "Huge"],
             ["South-South", "₦213,082,904", "Large"], ["North-Central", "₦147,098,128", "Small"],
             ["South-East", "₦122,905,857", "Almost nothing"]],
      so="With the axis at ₦120m, South-West looks over 40 times South-East; really it's about 2 times.",
      split=4.20),

    O("4.7.9",
      what="The **written report** follows the same story in short paragraphs: **headline**, **why it matters**, "
           "**evidence**, **recommendation** and **caveats**. Keep each to a few lines.",
      like="Like a well-packed bag: everything in its place, nothing extra.",
      how=["Paragraph 1: the headline.", "Paragraphs 2–3: why it matters and the evidence.", "Then the recommendation and the caveats."],
      tip="Put the caveats in, honestly; they build trust.",
      trap="Hiding the recommendation at the end of page 6.",
      case="The analyst writes the one-page FY2025 summary.",
      steps=["Write each paragraph.", "Check every number against the workbook.", "Add caveats."],
      table=[["Paragraph", "Content"], ["1. Headline", "Sales grew 20% to ₦733m; 7 of 12 stores missed target"],
             ["2. Why it matters", "Missed stores are ₦34.9m short of target"], ["3. Evidence", "Attainment chart, zone table, monthly trend"],
             ["4. Recommendation", "Three actions with owners and dates"], ["5. Caveats", "One target per store; FY2025 figures unaudited"]],
      so="A busy reader gets the whole story on one page.",
      split=4.90),
]
