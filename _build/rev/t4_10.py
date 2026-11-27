from revlib import T

TRY = "ETEH_Week4_Section_Examples.xlsx  ›  sheet '4.10 Audit & Protect'"
RT = "SuperMart Retail"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", RT)
    return T(tag, **k)


TOPICS = [
    O("4.10.1",
      what="Every workbook you inherit contains choices you didn't make and mistakes you can't see. **Auditing** "
           "means checking formulas, totals and references **before** you trust or share the numbers.",
      like="Like checking a second-hand car before buying: it looks fine until you open the bonnet.",
      how=["Compare totals with an independent check.", "Look for typed numbers where formulas should be.", "Trace the references of key cells."],
      tip="Add a check column that says OK or ERROR.",
      trap="Forwarding a colleague's figures without opening the formulas.",
      case="A colleague submits FY2025 zone sales shares. Three errors are hidden in it.",
      steps=["Recompute each share correctly.", "Compare with the submitted share.", "Flag the differences."],
      table=[["Zone", "Submitted share", "Correct share", "Check"], ["South-West", "41.0%", "34.1%", "ERROR"],
             ["South-South", "34.9%", "29.1%", "ERROR"], ["North-Central", "15% (typed)", "20.1%", "ERROR"],
             ["South-East", "100%", "16.8%", "ERROR"]],
      so="Every submitted share is wrong. The audit tools on the next slides show why.",
      split=4.20),

    O("4.10.2",
      what="**Formulas > Trace Precedents** draws arrows to the cells a formula **uses**. **Trace Dependents** "
           "shows which cells **use** this one. Remove Arrows clears them.",
      like="Like following the pipes to see where the water comes from and where it goes.",
      how=["Click the cell to check.", "Formulas > Trace Precedents.", "Read the blue box: does it cover the right range?"],
      tip="Double-click an arrow to jump to the other end.",
      trap="Assuming a SUM covers the whole column.",
      case="Trace the submitted Total cell.",
      steps=["Select the Total.", "Trace Precedents.", "Compare the box with the 4 zone rows."],
      table=[["Item", "Result"], ["Submitted total formula", "=SUM(C9:C11): 3 rows, not 4"], ["Submitted total", "₦610,309,398"],
             ["Correct total", "₦733,215,255"], ["Missing", "South-East, ₦122,905,857"]],
      so="The total misses South-East, so every share is too high.",
      split=4.90),

    O("4.10.3",
      what="**Show Formulas** (Ctrl + `) shows every formula instead of its result, so typed numbers stand out. "
           "**Evaluate Formula** steps through one formula piece by piece.",
      like="Like an X-ray: it shows what's underneath the surface.",
      how=["Press Ctrl + ` to toggle formula view.", "Look for numbers where formulas should be.", "Formulas > Evaluate Formula > Evaluate, step by step."],
      tip="Home > Find & Select > Go To Special > Constants selects every typed number.",
      trap="Trusting a cell because the result 'looks about right'.",
      case="Find the other two errors in the submitted share column.",
      steps=["Ctrl + ` on the sheet.", "Spot the typed value.", "Evaluate the South-East formula."],
      table=[["Cell", "What you see", "Problem"], ["North-Central share", "0.15", "Typed number, not a formula"],
             ["South-East share", "=C12/C12", "Divides by itself: 100%"], ["Fix", "=C12/SUM($C$9:$C$12)", "Absolute reference to the total"]],
      so="One typed value and one drifting reference: both invisible until you look at the formulas.",
      split=4.70),

    O("4.10.4",
      what="**Formulas > Error Checking** walks through every error cell in turn. The **Watch Window** keeps key "
           "cells in view while you edit other sheets.",
      like="Like a car dashboard: warning lights plus the gauges you watch while driving.",
      how=["Formulas > Error Checking: step through each flag.", "Formulas > Watch Window > Add Watch.", "Edit data elsewhere and watch the cell change."],
      tip="Watch the grand total while you clean data on another sheet.",
      trap="Ignoring the small green triangles.",
      case="Fix the three errors and watch the total.",
      steps=["Fix the total range.", "Replace the typed value; anchor the denominator.", "Add the total to the Watch Window."],
      table=[["After the fix", "Result"], ["Total", "₦733,215,255"], ["Shares add up to", "100%"],
             ["Check column", "OK on all 4 zones"]],
      so="With the fixes in place the check column shows OK and the total matches the ledger.",
      split=4.90),

    O("4.10.5",
      what="**Review > Protect Sheet** stops people overwriting **locked** cells. All cells are locked by default, so "
           "**unlock the input cells** first (Format Cells > Protection). Protection stops accidents; **only encryption** "
           "protects confidential data.",
      like="Like a glass case in a shop: you can look and change the price tag, but not move the goods.",
      how=["Unlock input cells: Format Cells > Protection > untick Locked.", "Review > Protect Sheet.", "For secrecy: File > Info > Protect Workbook > Encrypt."],
      tip="Colour unlocked input cells yellow so users know where to type.",
      trap="Thinking a sheet password keeps data secret.",
      case="The FY2026 plan sheet: only the growth input can change.",
      steps=["Unlock the growth cell.", "Protect the sheet.", "Try typing in the plan formula."],
      table=[["Cell", "Status", "Value"], ["Growth assumption", "Unlocked (yellow)", "12%"],
             ["FY2026 sales plan", "Locked formula", "₦821,201,086"], ["Typing in the plan cell", "Blocked", "Message shown"]],
      so="Users can try different growth rates but can't break the formula.",
      split=4.70),

    O("4.10.6",
      what="**Developer > Record Macro** records your clicks and keystrokes as VBA code you can replay. Record "
           "carefully, then tidy the code. Save as **.xlsm**.",
      like="Like recording a cooking video: anyone can repeat the steps exactly.",
      how=["Developer > Record Macro; name it.", "Do the task once, carefully.", "Stop; review the code (Alt+F11); save as .xlsm."],
      tip="Turn on the Developer tab: File > Options > Customize Ribbon.",
      trap="Recording with absolute clicks that break when data grows.",
      who="Lifecare Hospital",
      case="Automate the monthly refresh-and-export of the Lifecare report.",
      steps=["Record: Refresh All.", "Recalculate.", "Export the report sheet to PDF."],
      table=[["Step in the macro", "Line of code"], ["Refresh queries", "ThisWorkbook.RefreshAll"],
             ["Recalculate", "Application.CalculateFull"], ["Export PDF", "Worksheets(\"4.8 Report Writing\").ExportAsFixedFormat"]],
      so="One click does the monthly routine the same way every time.",
      split=4.90),

    O("4.10.7",
      what="**Office Scripts** automate Excel on the web (Automate > New Script) in TypeScript. **Power Automate** "
           "runs them on a schedule or trigger, for example every Monday or when a file arrives.",
      like="Like setting a washing machine timer: it runs while you're busy elsewhere.",
      how=["Excel for the web > Automate > New Script.", "Write or record the script.", "Schedule it in Power Automate."],
      tip="Scripts are stored in OneDrive, not inside the workbook.",
      trap="Using a VBA macro when users work in a browser.",
      case="Script: count the stores below target in tStory.",
      steps=["getTable(\"tStory\").", "Read the Attainment column.", "Count values below 1."],
      table=[["Tool", "Runs in", "Good for"], ["VBA macro", "Desktop Excel", "Local, complex routines"],
             ["Office Script", "Excel for the web", "Shared online workbooks"], ["Power Automate", "The cloud", "Scheduled or triggered runs"],
             ["Script result here", "tStory", "7 stores below target"]],
      so="Choose the tool by where people work: desktop, browser or schedule.",
      split=4.20),

    O("4.10.8",
      what="Every macro or script needs a short record: **name**, **purpose**, **trigger**, **inputs and "
           "outputs**, **owner** and **last tested**. Without it, automation becomes a mystery nobody dares touch.",
      like="Like the manual for a machine: without it, nobody knows how to fix it.",
      how=["Write the log next to the code.", "Fill every field.", "Update 'last tested' after each change."],
      tip="Put a one-line comment at the top of the code too.",
      trap="Automation that only its creator understands.",
      who="Lifecare Hospital",
      case="Log for the refresh-and-export macro.",
      steps=["Name and purpose.", "Trigger and inputs/outputs.", "Owner and test date."],
      table=[["Field", "Entry"], ["Name", "Refresh_And_Export_Report"], ["Purpose", "Monthly refresh + PDF of the Lifecare report"],
             ["Trigger", "Manual button, first working day"], ["Owner / last tested", "Analyst / 2024-01-12"],
             ["Saved as", ".xlsm (macro-enabled)"]],
      so="Anyone can find, run and maintain the macro.",
      split=4.90),
]
