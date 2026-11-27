from revlib import T

TRY = "ETEH_Week4_Section_Examples.xlsx  ›  sheet '4.8 Report Writing'"
LC = "Lifecare Hospital"


def O(tag, **k):
    k.setdefault("tryit", TRY)
    k.setdefault("who", LC)
    return T(tag, **k)


TOPICS = [
    O("4.8.1",
      what="A professional report has a fixed order: **title page**, **executive summary**, **method and data**, "
           "**findings**, **recommendations** and an **appendix**. Readers know where to look for each thing.",
      like="Like a well-organised supermarket: everyone knows where the bread is.",
      how=["Start from the six-section outline.", "Put details in the appendix, not the summary.", "Number the recommendations."],
      tip="Write the executive summary last, after the findings are final.",
      trap="Starting with method; busy readers stop before the findings.",
      case="The analyst outlines Lifecare's 2023 Operations Report.",
      steps=["List the six sections.", "Decide what goes in each.", "Move detail to the appendix."],
      table=[["Section", "Lifecare 2023 content"], ["Title page", "Title, version, date, author, classification"],
             ["Executive summary", "One page: findings, implications, recommendations"], ["Method and data", "Source, period, exclusions, ALOS and readmission definitions"],
             ["Findings", "Admissions, occupancy, stay, readmissions, billing"], ["Appendix", "Full tables, data dictionary, reconciliation"]],
      so="Every reader, from board member to auditor, finds what they need fast.",
      split=4.90),

    O("4.8.2",
      what="The **executive summary** fits on one page: the **headline**, about **three key numbers**, what they "
           "mean, and the **recommendations**. Build the number sentences with formulas so they always match the data.",
      like="Like the back cover of a book: enough to decide whether to act, without reading it all.",
      how=["Write each finding as a formula sentence (TEXT and &).", "Add what it means.", "Type the recommendations: they're judgement."],
      tip="If the board reads only this page, can they decide?",
      trap="A summary that only lists what was analysed.",
      case="Executive summary for 2023, built from tLOS (10,918 stays) and tCensus23 (365 days).",
      steps=["Finding sentences with formulas.", "Check each against the tables.", "Add the recommendation."],
      table=[["Part", "Text"], ["Finding 1", "10,918 patients admitted; Endocrinology was busiest"],
             ["Finding 2", "Average occupancy 89.5%, above the 85% safe level; 43 days above 95%"],
             ["Finding 3", "Average stay 15.5 days; 30-day readmission 10.3%"],
             ["Recommendation", "Plan Q2 surge capacity; pilot discharge planning in Oncology"]],
      so="Three facts and one ask. The numbers update themselves if the data is corrected.",
      split=4.90),

    O("4.8.3",
      what="Good layout uses **margins**, **alignment**, **consistent text sizes** and **white space**. Space "
           "groups related items and separates sections. Crowded pages are hard to read.",
      like="Like a tidy room: you see where everything is because nothing is piled together.",
      how=["Use one font and 2–3 text sizes.", "Align everything to the same left edge.", "Leave space between sections."],
      tip="Squint at the page: the structure should still be visible.",
      trap="Filling every gap with another chart.",
      case="The analyst lays out the findings page of the Lifecare report.",
      steps=["Set margins and a grid.", "Choose heading and body sizes.", "Add space between blocks."],
      table=[["Layout choice", "Recommended", "Effect"], ["Fonts", "One family, 2–3 sizes", "Calm, consistent"],
             ["Alignment", "One left edge", "Looks organised"], ["White space", "Between sections", "Groups ideas"],
             ["Charts per page", "1–2 with finding titles", "Clear focus"]],
      so="A clean page lets the findings stand out.",
      split=4.20),

    O("4.8.4",
      what="Report tables need a **header row**, **numbers right-aligned**, **text left-aligned**, **units in the "
           "headers**, the **same decimals** in a column, and a sensible sort order.",
      like="Like a bus timetable: clear columns, same format, easy to scan.",
      how=["Units in headers: 'ALOS (days)', 'Avg bill (₦)'.", "Right-align numbers; same decimals.", "Sort by the measure that matters; bold totals."],
      tip="Remove heavy gridlines; light lines are enough.",
      trap="Units typed inside cells ('15.3 days'), which stops sorting and maths.",
      case="Department summary table for 2023 (tLOSdep).",
      steps=["AVERAGEIFS for ALOS and average bill.", "COUNTIFS for readmission rate.", "Format and sort."],
      table=[["Department", "Admissions", "ALOS (days)", "Readmit (%)"], ["Endocrinology", "3,682", "15.3", "9.9%"],
             ["Cardiology", "1,828", "15.3", "11.4%"], ["Oncology", "1,824", "15.9", "13.2%"],
             ["Orthopaedics & Rheumatology", "1,806", "15.5", "8.3%"], ["Respiratory", "1,778", "15.7", "9.2%"]],
      so="Oncology stands out: highest readmission rate (13.2%) and the longest average stay.",
      split=4.20),

    O("4.8.5",
      what="Set up the page before printing: **orientation** (portrait for text, landscape for wide tables), "
           "**fit to 1 page wide**, **print titles** that repeat headings, and a **header and footer**.",
      like="Like setting up a printer for photos versus documents: the right paper and layout for the job.",
      how=["Page Layout > Orientation and Size (A4).", "Scale to fit: Width 1 page.", "Print Titles: rows to repeat; add header and footer."],
      tip="Check File > Print preview before exporting.",
      trap="A table that prints across 4 pages with columns cut off.",
      case="The 4.8 sheet is already set up to print as a clean A4 report.",
      steps=["Open File > Print preview.", "Check the settings below.", "Adjust if needed."],
      table=[["Setting", "Value"], ["Orientation / size", "Landscape A4, 1 page wide"], ["Print titles", "Rows 1–3 repeat on every page"],
             ["Header", "Report title"], ["Footer", "Confidential, page X of Y, file name"]],
      so="Every printed page carries the title, page number and classification.",
      split=4.90),

    O("4.8.6",
      what="**File > Export > Create PDF/XPS** makes the version you send out. In **Options**, choose 'Active "
           "sheet', 'Selection' or 'Entire workbook'. A PDF can't be edited by accident and looks the same everywhere.",
      like="Like laminating a notice before pinning it up: it stays exactly as you made it.",
      how=["Set the print area and page setup first.", "File > Export > Create PDF/XPS > Options.", "Choose what to publish; Standard quality."],
      tip="Open the PDF and check every page before sending.",
      trap="Exporting the entire workbook, including raw data sheets.",
      case="Export the Lifecare report sheet as a PDF for the board pack.",
      steps=["Check Print preview.", "Export with 'Active sheet'.", "Open and check the PDF."],
      table=[["Option", "Setting", "Why"], ["Publish what", "Active sheet", "Leaves out raw data sheets"],
             ["Quality", "Standard", "Sharp when printed"], ["Document properties", "Included", "Title and author travel with it"],
             ["After export", "Open and check", "Catch cut-off columns"]],
      so="The board gets a clean, fixed report, without the raw patient-level data.",
      split=4.20),

    O("4.8.7",
      what="A PDF that looks right on your computer can **swap fonts** on another, breaking layout or symbols. "
           "Use common fonts (Calibri, Arial) that contain the symbols you need, such as **₦**.",
      like="Like sending a recipe in a language the cook can't read: the meaning is lost.",
      how=["Use Calibri or Arial.", "Check that ₦ displays correctly.", "Open the PDF on another device."],
      tip="Test on a phone too: many board members read there.",
      trap="Decorative fonts that aren't available on other machines.",
      case="A board member reports that '₦' shows as an empty box in the PDF.",
      steps=["Find the font used for that text.", "Switch to Calibri.", "Re-export and check on another device."],
      table=[["Risk", "Symptom", "Prevention"], ["Font missing", "Layout shifts, text overflows", "Common fonts only"],
             ["Symbol missing", "₦ shows as a box", "Calibri contains ₦"], ["Not checked", "Errors found by the board", "Open on a second device"]],
      so="A quick check on another device prevents an embarrassing board pack.",
      split=4.20),

    O("4.8.8",
      what="Files carry hidden **metadata**: author, company, title, comments. Set the useful fields (**File > "
           "Info > Properties**) and remove what shouldn't travel (**Check for Issues > Inspect Document**).",
      like="Like the label on a parcel: right name and address on, old labels off.",
      how=["File > Info > Properties: Title, Author, Company.", "Check for Issues > Inspect Document.", "Remove comments and personal information."],
      tip="Inspect before every external release.",
      trap="Sending a file whose comments contain patient details.",
      case="Check the workbook's properties before release.",
      steps=["Open File > Info.", "Read the properties.", "Run Inspect Document."],
      table=[["Field", "Value in this workbook"], ["Title", "ETEH Week 4 - Section Examples"], ["Author", "ETEH Ltd - Eaglet2EagleHub Limited"],
             ["Comments / personal info", "Remove before external release"]],
      so="The file says who made it and what it is, and carries nothing it shouldn't.",
      split=4.90),

    O("4.8.9",
      what="Name files with a **date and version**, for example **2024-01-12_Lifecare_2023_Report_v1.0.pdf**, and "
           "keep a **change log**: version, date, who, what changed. Never call a file 'final'.",
      like="Like dated photos in an album: you always know which came first.",
      how=["Name: YYYY-MM-DD_Topic_vX.Y.", "Log every change.", "v1.0 = approved; v0.x = drafts."],
      tip="Dates written YYYY-MM-DD sort in the right order.",
      trap="Report_final_FINAL_v3_really.pdf",
      case="The change log for the Lifecare 2023 report.",
      steps=["Draft v0.1.", "Update after review.", "Approve as v1.0."],
      table=[["Version", "Date", "By", "Change"], ["v0.1", "2024-01-08", "Analyst", "First draft from extract LC_2023_v1"],
             ["v0.2", "2024-01-10", "Analyst", "Added readmission by department"],
             ["v1.0", "2024-01-12", "Head of Quality", "Approved for board pack"]],
      so="Anyone can see which version is approved and what changed.",
      split=4.20),

    O("4.8.10",
      what="Before release, run a **checklist**: numbers reconcile to the source, the headline matches the "
           "evidence, charts have finding titles and sources, no personal data, consistent formats, PDF checked.",
      like="Like a pilot's pre-flight checklist: even experts use it every time.",
      how=["Reconcile key totals to the source.", "Tick each item.", "Ask a second person to review."],
      tip="Keep the checklist in the workbook, next to the report.",
      trap="Releasing without a second pair of eyes.",
      case="Final check of the Lifecare report.",
      steps=["Reconcile admissions to the extract.", "Run the checklist.", "Reviewer signs off."],
      table=[["Check", "Pass condition"], ["Numbers reconcile", "Admissions = 10,918 = extract rows"],
             ["Charts", "Finding title and source line on each"], ["Privacy", "No patient names or IDs"],
             ["Formats", "Consistent ₦ and %; spell-checked"], ["PDF", "Opened and checked on another computer"]],
      so="Five ticks and a second reviewer, and the report is ready.",
      split=4.90),
]
