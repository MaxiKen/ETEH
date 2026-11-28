"""Week 4 edition-3 slide copy: Modules 4.7 (storytelling), 4.8 (report writing and PDF), 4.9 (slide presentations) and
4.10 (formula auditing, protection and light automation).
Data: tDash25 + tStores (SuperMart FY2025), tMonthly, tLOS + tCensus23 (Mercyland 2023), tHotel (Atlantic Pearl). Every number is
computed here."""
import pandas as pd

from v3kit import T
from w4_fx import C23, D25, HOT, I, LOS, MON, N, Nm, P, STO

# ---------------------------------------------------------------- SuperMart FY2025
ST = (D25.groupby(["Store_ID", "Store_Name", "Zone"]).agg(Sales=("Sales", "sum"), Profit=("Profit", "sum")).reset_index()
      .merge(STO[["Store_ID", "Annual_Target_2025"]]))
ST["att"] = ST.Sales / ST.Annual_Target_2025
ST["gap"] = ST.Sales - ST.Annual_Target_2025
TOT25 = ST.Sales.sum()
YR = MON.groupby("Year").Sales.sum()
GROW = TOT25 / YR[2024] - 1
BELOW = ST[ST.att < 1].sort_values("gap")
NBEL = len(BELOW)
WORST = BELOW.iloc[0]
BEST = ST.sort_values("att").iloc[-1]
LOSS = ST[ST.Profit < 0]
ZN = D25.groupby("Zone").Sales.sum().sort_values(ascending=False)
AX = 120e6
HEADLINE = (f"FY2025 sales grew {P(GROW, 0)} to {Nm(TOT25, 0)}, but {NBEL} of {len(ST)} stores missed target; "
            f"{WORST.Store_Name} has the largest gap ({Nm(WORST.gap)}).")

# ---------------------------------------------------------------- Mercyland 2023
DEP = (LOS.groupby("Department").agg(n=("Episode_ID", "size"), alos=("LOS_Days", "mean"),
                                    re=("Readmitted_30d", lambda x: (x == "Yes").mean())).sort_values("n", ascending=False))
RE_ALL = (LOS.Readmitted_30d == "Yes").mean()
OCC = C23.Occupancy_Pct
HI95 = int((OCC > 0.95).sum())
C23M = pd.to_datetime(C23.Census_Date).dt.month
HI95_Q2 = int(((OCC > 0.95) & C23M.between(4, 6)).sum())
PEAK = C23.loc[OCC.idxmax()]
TOPRE = DEP.re.idxmax()

# ---------------------------------------------------------------- Atlantic Pearl cancellations
H = HOT.assign(C=(HOT.Reservation_Status == "Cancelled").astype(int))
CANV = H.loc[H.C == 1, "Reservation_Value_NGN"].sum()
CR = H.C.mean()
SEG = H.groupby("Sales_Channel").agg(n=("C", "size"), r=("C", "mean")).sort_values("n", ascending=False)
LB = pd.cut(H.Days_Ahead, [-1, 7, 30, 90, 180, 10_000], labels=["0-7 days", "8-30 days", "31-90 days", "91-180 days", "181+ days"])
LEAD = H.groupby(LB, observed=True).C.mean()
T4 = f"Bookings made 6+ months ahead cancel {LEAD.iloc[-1]:.0%} of the time"

# ---------------------------------------------------------------- 4.10 audit case
ZO = ZN.reindex(["South-West", "South-South", "North-Central", "South-East"])
SUBTOT = ZO.iloc[:3].sum()
SUB = [ZO.iloc[0] / SUBTOT, ZO.iloc[1] / SUBTOT, 0.15, 1.0]
COR = list(ZO / ZO.sum())


M47 = dict(
    id="4.7", title="Data Storytelling",
    subtitle="Turning a finding into a decision: narrative structure, message hierarchy, evidence selection and the framing "
             "executives actually respond to.",
    objectives=["Explain why a story moves people to act when a table of numbers does not.",
                "Structure a presentation in three acts and write one headline message.",
                "Choose evidence that fits the audience.",
                "Write titles that state the finding and annotate charts.",
                "Avoid misleading charts and write a one-page narrative."],
    topics=[
        T("4.7.1", "Why Storytelling Matters",
          defn="Data storytelling links numbers into a narrative with a point, so the audience understands what happened, why it "
               "matters and what to do.",
          simple="People remember a story and a decision, not twelve charts.",
          grid=dict(label="SuperMart FY2025 board pack (tDash25, tStores)", widths=[2.4, 2.4], rows=[
              ["Without a story", "With a story"],
              [f"{len(ST)} store charts, no conclusion", f"One headline: +{P(GROW, 0)}; {NBEL} stores missed target"],
              ["Every number has the same weight", f"Focus on {WORST.Store_Name}"],
              ["Board asks 'so what?'", "Board decides on a recovery plan"]]),
          watch="A story is not spin: every claim must still be checkable against the data."),

        T("4.7.2", "The Three-Act Narrative",
          defn="The three-act narrative is situation (what is happening), complication (the problem or change) and resolution "
               "(what to do about it).",
          simple="Context, problem, ask: in that order.",
          grid=dict(label="FY2025 story in three acts", widths=[1.3, 2.5, 1.5], size=11.5, rows=[
              ["Act", "Message", "Evidence"],
              ["1. Situation", f"Sales grew {P(GROW, 0)} to {Nm(TOT25, 0)}", "Annual sales line"],
              ["2. Complication", f"{NBEL} of {len(ST)} stores missed target; {len(LOSS)} lost money", "Attainment bars"],
              ["3. Resolution", f"Recovery plans for the {NBEL} stores, starting with {WORST.Store_Name}", "Gap table"]]),
          watch="Skipping Act 1 leaves the audience without context; skipping Act 3 leaves them without a decision."),

        T("4.7.3", "The Headline Message",
          defn="The headline message is the one sentence the audience must remember, stating the finding and its size.",
          simple="Build it from formulas so it updates with the data: =\"Sales grew \"&TEXT(growth,\"0%\")&...",
          grid=dict(label="Parts of the FY2025 headline", widths=[2.2, 2], rows=[
              ["Part", "Value"],
              ["Growth vs FY2024", P(GROW)],
              ["FY2025 sales", N(TOT25)],
              ["Stores below target", f"{NBEL} of {len(ST)}"],
              ["Largest gap", f"{WORST.Store_Name}, {Nm(WORST.gap)}"]]),
          watch=f"Headline: '{HEADLINE}'"),

        T("4.7.4", "Message Hierarchy",
          defn="A message hierarchy puts the headline first, then two or three supporting points, then the detail.",
          simple="A reader who stops after any line still has the most important part.",
          grid=dict(label="FY2025 message hierarchy", widths=[1.4, 3.4], size=11.5, rows=[
              ["Tier", "Content"],
              ["Headline", f"Grew {P(GROW, 0)} to {Nm(TOT25, 0)}; {NBEL} of {len(ST)} stores missed target"],
              ["Support 1", f"{ZN.index[0]} brings {P(ZN.iloc[0] / TOT25, 0)} of sales"],
              ["Support 2", f"{len(LOSS)} stores made a loss, led by {LOSS.sort_values('Profit').Store_Name.iloc[0]} ({Nm(LOSS.Profit.min())})"],
              ["Detail", "Store table, monthly trend, method notes (appendix)"]]),
          watch="Burying the headline on page 5 means most readers never see it."),

        T("4.7.5", "Choosing Evidence for the Audience",
          defn="Evidence selection picks the few numbers and charts each audience needs to make its own decision.",
          simple="Ask: what will this person do with it?",
          grid=dict(label="Three audiences for the FY2025 results", widths=[1.3, 2.4, 1.1], rows=[
              ["Audience", "Evidence they need", "Format"],
              ["Board", "Headline, 3 numbers, decision needed", "1 slide"],
              ["Zone managers", "Their stores vs target, by month", "Dashboard"],
              ["Finance", "Reconciliation to the ledger, method", "Report"]]),
          watch="Giving everyone everything means nobody finds what they need."),

        T("4.7.6", "Titles That State the Finding",
          defn="An action title states the finding as a sentence; a label title only names the chart.",
          simple="Read only the titles of a deck: they should tell the whole story.",
          grid=dict(label="Retitling the store attainment chart", widths=[1.6, 3.2], rows=[
              ["Title type", "Example"],
              ["Label (weak)", "Store attainment FY2025"],
              ["Finding (strong)", f"{NBEL} of {len(ST)} stores missed target; {WORST.Store_Name.split(' - ')[0]} is furthest behind"]]),
          watch="Keep titles under about 12 words; the chart supplies the detail."),

        T("4.7.7", "Annotation and Callouts",
          defn="Annotations are labels, lines and colours added to a chart to point the eye at what matters.",
          simple="Grey for context, one strong colour for the point.",
          grid=dict(label="Annotating the attainment chart", widths=[1.6, 3.2], rows=[
              ["Annotation", "Placement"],
              ["Target line", "At 100% across all bars"],
              ["Red bars", f"The {NBEL} stores below target"],
              ["Callout", f"{WORST.Store_Name}: {P(WORST.att, 0)} of target"],
              ["Data labels", "On the red bars only"]]),
          watch="Annotate the point, not every bar; too many labels hide the message."),

        T("4.7.8", "Avoiding Misleading Presentations",
          defn="A chart misleads when its design changes the impression the numbers give: cut axes, 3-D, uneven intervals, "
               "cherry-picked periods.",
          simple="Bar charts must start at zero; bar length is the value.",
          grid=dict(label=f"FY2025 sales by zone: axis starting at {Nm(AX, 0)}", widths=[1.4, 1.4, 1.8], rows=[
              ["Zone", "Sales", "Visible bar"]] + [[z, Nm(v), Nm(max(v - AX, 0))] for z, v in ZN.items()]),
          watch=f"With the axis at {Nm(AX, 0)}, {ZN.index[0]} looks {(ZN.iloc[0] - AX) / (ZN.iloc[-1] - AX):.0f} times "
                f"{ZN.index[-1]}; really it is {ZN.iloc[0] / ZN.iloc[-1]:.1f} times."),

        T("4.7.9", "Structuring the Written Narrative",
          defn="A written narrative gives the headline, the supporting evidence, the implication and the recommendation, in "
               "short paragraphs.",
          simple="One page; one idea per paragraph; numbers that match the charts.",
          grid=dict(label="One-page FY2025 summary", widths=[1.4, 3.4], size=11.5, rows=[
              ["Paragraph", "Content"],
              ["1. Headline", f"Sales grew {P(GROW, 0)} to {Nm(TOT25, 0)}; {NBEL} of {len(ST)} stores missed target"],
              ["2. Evidence", f"{BEST.Store_Name} best at {P(BEST.att, 0)}; {WORST.Store_Name} worst at {P(WORST.att, 0)}"],
              ["3. Implication", f"Below-target stores cost {Nm(-BELOW.gap.sum())} against plan"],
              ["4. Ask", "Approve recovery plans for the stores below target"]]),
          watch="If a number in the text is not in a chart or table, the reader cannot check it."),
    ],
    takeaways=["A story turns numbers into a decision.",
               "Three acts: situation, complication, resolution.",
               "One headline, built from formulas, at the top.",
               "Fit evidence to the audience; titles state the finding.",
               "Annotate the point and never distort the axis."],
)


M48 = dict(
    id="4.8", title="Report Writing and PDF Production",
    subtitle="Structure, layout, page setup, PDF export and version control for a report someone will actually read.",
    objectives=["Structure a report and write an executive summary.",
                "Lay out pages and tables for reading.",
                "Set up pages for print and export clean PDFs.",
                "Check the exported PDF page by page before release.",
                "Version-control a report and review it with a checklist."],
    topics=[
        T("4.8.1", "Report Structure",
          defn="A standard report has a title page, executive summary, introduction and method, findings, recommendations and "
               "appendices.",
          simple="Readers jump: the structure lets each find their part.",
          grid=dict(label="Mercyland 2023 Operations Report", widths=[1.6, 3.2], rows=[
              ["Section", "Content"],
              ["Title page", "Title, version, date, author, classification"],
              ["Executive summary", "Findings and the ask, on one page"],
              ["Method", f"Data: tLOS ({I(len(LOS))} stays), tCensus23 ({len(C23)} days)"],
              ["Findings / Appendix", "Department table, occupancy chart / definitions"]]),
          watch="An appendix is for detail, not for findings nobody will see."),

        T("4.8.2", "Writing the Executive Summary",
          defn="The executive summary states the purpose, the three or four key findings with numbers, and the recommendation, "
               "on one page.",
          simple="Write it last; put it first.",
          grid=dict(label="Mercyland 2023 executive summary (tLOS, tCensus23)", widths=[1.2, 3.6], size=11.5, rows=[
              ["Part", "Text"],
              ["Finding 1", f"{I(len(LOS))} admissions; {DEP.index[0]} was busiest ({I(DEP.n.iloc[0])})"],
              ["Finding 2", f"Average occupancy {P(OCC.mean())}; above 95% on {HI95} days, {HI95_Q2} of them in April-June"],
              ["Finding 3", f"30-day readmission {P(RE_ALL)}; highest in {TOPRE} ({P(DEP.re[TOPRE])})"],
              ["Ask", f"Plan surge capacity for Q2; pilot discharge planning in {TOPRE}"]]),
          watch="Build each number with a formula from the data so a corrected extract updates the summary."),

        T("4.8.3", "Layout and White Space",
          defn="Layout is the arrangement of text, tables and charts on the page; white space is the empty space that separates "
               "and groups them.",
          simple="One font family, two or three sizes, aligned edges, generous margins.",
          grid=dict(label="Layout choices for the findings page", widths=[1.4, 1.9, 1.5], rows=[
              ["Choice", "Recommended", "Effect"],
              ["Fonts", "One family, 2-3 sizes", "Calm, consistent"],
              ["Alignment", "Left, on a grid", "Easy to scan"],
              ["White space", "Space between sections", "Groups ideas"],
              ["Colour", "One accent colour", "Points to the finding"]]),
          watch="Filling every gap makes a page look busy and hides the key number."),

        T("4.8.4", "Table Design for Reports",
          defn="A report table has clear headers with units, right-aligned numbers, consistent decimals, light lines and a sort "
               "order that makes the point.",
          simple="Sort by the measure that matters, not alphabetically.",
          grid=dict(label="Departments, 2023 (tLOS)", widths=[2.1, 1.1, 1, 1.1], size=11, rows=[
              ["Department", "Admissions", "ALOS (days)", "Readmit"]] + [
              [d, I(r.n), f"{r.alos:.1f}", P(r.re)] for d, r in DEP.iterrows()]),
          watch=f"{TOPRE} stands out on readmission ({P(DEP.re[TOPRE])}) and stay length; a sorted, clean table shows it at once."),

        T("4.8.5", "Page Setup for Reports",
          defn="Page setup controls orientation, paper size, margins, scaling, print titles, headers and footers.",
          simple="Page Layout › Page Setup; check in File › Print before exporting.",
          grid=dict(label="Settings for the report sheet", widths=[1.8, 3], rows=[
              ["Setting", "Value"],
              ["Orientation / size", "Landscape A4, fit to 1 page wide"],
              ["Print titles", "Rows 1-3 repeat on every page"],
              ["Footer", "Page &[Page] of &[Pages]; classification"],
              ["Print area", "The report block only"]]),
          watch="Without print titles, page 2 of a long table has no headers."),

        T("4.8.6", "Exporting to PDF",
          defn="File › Export › Create PDF/XPS (or Save As PDF) turns the sheet into a fixed document that looks the same "
               "everywhere.",
          simple="Options › Publish what: choose the active sheet or a selection, not the entire workbook.",
          grid=dict(label="Export settings", widths=[1.5, 1.4, 1.9], rows=[
              ["Option", "Setting", "Why"],
              ["Publish what", "Active sheet", "Leaves out raw data"],
              ["Optimise for", "Standard", "Print quality"],
              ["Document properties", "Include", "Title, author travel"]]),
          watch="'Entire workbook' can publish the patient-level data sheets: always check the page count."),

        T("4.8.7", "Version Control for Reports",
          defn="Version control records each version of a report with its number, date, author and what changed.",
          simple="A change log sheet and a clear file name: Report_v1.2_2024-01-15.",
          grid=dict(label="Change log, Mercyland 2023 report", widths=[0.8, 1.3, 2.7], rows=[
              ["Version", "Date", "Change"],
              ["v0.1", "2024-01-08", "First draft from the 2023 extract"],
              ["v0.2", "2024-01-11", "Readmission definition corrected"],
              ["v1.0", "2024-01-15", "Approved by the Head of Operations"]]),
          watch="'Final_v2_REALfinal.xlsx' is not version control."),

        T("4.8.8", "Review Checklist for Reports",
          defn="A review checklist lists the checks every report passes before release: numbers, text, format, file and approval.",
          simple="A second person ticks it, not the author.",
          grid=dict(label="Final check of the Mercyland report", widths=[1.7, 3.1], rows=[
              ["Check", "Pass condition"],
              ["Numbers reconcile", f"Admissions = {I(len(LOS))} = extract rows"],
              ["Text matches tables", "Every number in the summary is in a table"],
              ["PDF checked", "Opened on a second device"],
              ["Metadata and approval", "Inspected; sign-off recorded"]]),
          watch="The author is the worst person to proofread their own report."),
    ],
    takeaways=["Standard structure; executive summary first, written last.",
               "Clean layout and sorted tables make the point.",
               "Set up pages, then export only what should be shared.",
               "Open the PDF itself and check every page before sending.",
               "Version every release and review it with a checklist."],
)


M49 = dict(
    id="4.9", title="Slide Presentations",
    subtitle="Decks that argue rather than decorate: structure, one message per slide, charts built for projection, and "
             "delivery that survives questions.",
    objectives=["Plan a deck as a storyboard with timings.",
                "Write one message per slide and keep text light.",
                "Build charts and tables that read on a projector.",
                "Apply a consistent, accessible visual identity.",
                "Prepare for questions and deliver the analysis."],
    topics=[
        T("4.9.1", "Presentation Structure",
          defn="A presentation structure is the storyboard: the order of slides, the role of each, and the minutes each gets.",
          simple="Opening headline, evidence, implication, ask.",
          grid=dict(label=f"10-slide talk on Atlantic Pearl cancellations (tHotel, {I(len(H))} bookings)", widths=[0.8, 3.2, 0.9],
                    size=11.5, rows=[
              ["Slides", "Role", "Minutes"],
              ["1", f"Headline: cancellations cost {Nm(CANV / 1e3, 2).replace('m', 'bn')} of bookings", "0.5"],
              ["2-6", "Evidence: rate, segment, lead time, value", "6"],
              ["7-9", "Options and recommendation", "3.5"],
              ["10", "The decision needed", "1"]]),
          watch="Storyboard on paper first; building slides first produces a deck without a story."),

        T("4.9.2", "One Message per Slide",
          defn="Each slide makes one point, stated in its title and proved by one chart, table or number.",
          simple="If a slide needs two titles, it is two slides.",
          grid=dict(label="Titles and proof numbers, slides 2-4", widths=[3.4, 1.2], rows=[
              ["Slide title", "Proof"],
              ["One booking in three is cancelled", P(CR)],
              ["Online travel agent bookings cancel most", P(SEG.r["Online Travel Agent"])],
              [T4, P(LEAD.iloc[-1])]]),
          watch="A slide with no clear point invites the audience to find their own, often the wrong one."),

        T("4.9.3", "Building Charts for Projection",
          defn="A projection chart uses large fonts, few categories, thick bars, direct labels and one highlight colour.",
          simple="Test it from the back of the room: if you cannot read it, nobody can.",
          grid=dict(label="Cancel rate by lead time (slide 4)", widths=[1.4, 1.2, 1.2], rows=[
              ["Booked ahead", "Cancel rate", "Colour"]] + [[k, P(v), "Teal" if i == len(LEAD) - 1 else "Grey"] for i, (k, v) in enumerate(LEAD.items())]),
          watch="Gridlines, legends and 3-D effects that look fine on a laptop disappear or clutter on a projector."),

        T("4.9.4", "Text Density on Slides",
          defn="Text density is the amount of text per slide; slides support the speaker, they are not the script.",
          simple="Title under about 12 words; at most three short bullets; 24 pt or larger.",
          grid=dict(label="Check the slide 4 title with formulas", widths=[1.6, 3.2], rows=[
              ["Check", "Result"],
              ["Title", T4],
              ["Words", str(len(T4.split()))],
              ["Characters (LEN)", str(len(T4))]]),
          watch="If the audience is reading, it is not listening."),

        T("4.9.5", "Consistent Visual Identity",
          defn="Visual identity is the fixed set of colours, fonts, logo position and layouts used on every slide.",
          simple="Set it once in the Slide Master (View › Slide Master).",
          grid=dict(label="Visual rules for the deck", widths=[1.5, 3.3], rows=[
              ["Element", "Rule"],
              ["Main colour", "ETEH navy #0B2545"],
              ["Highlight", "One accent (teal) for the point"],
              ["Fonts", "One family; titles 28 pt, body 20 pt or larger"],
              ["Layout", "Title top-left; source bottom-left"]]),
          watch="A new colour on every slide makes the audience look for a meaning that is not there."),

        T("4.9.6", "Tables on Slides",
          defn="A slide table shows only the rows and columns needed for the point, in large type, with the key row highlighted.",
          simple="At most about five rows; round the numbers.",
          grid=dict(label="Cancel rate by market segment (slide 3)", widths=[2.2, 1.1, 1.1], rows=[
              ["Segment", "Bookings", "Cancel rate"]] + [[s, I(r.n), P(r.r)] for s, r in SEG.iterrows()]),
          watch="A 30-row table on a slide is a handout; send it as one."),

        T("4.9.7", "Handling Questions",
          defn="Handling questions means anticipating them, preparing short answers with evidence, and admitting what you do not "
               "know.",
          simple="Write the five hardest questions before the meeting; put backup slides in the appendix.",
          grid=dict(label="Prepared answers for the board", widths=[1.6, 3.2], rows=[
              ["Question", "Prepared answer"],
              ["Where is the data from?", f"{I(len(H))} bookings, 2023-24 (tHotel)"],
              ["Is the lead-time link causal?", "No: association; early bookers have more time to change plans"],
              ["What do we not know?", "Why guests cancel: no reason field"]]),
          watch="Guessing an answer under pressure damages trust more than 'I will check and come back to you'."),

        T("4.9.8", "Timing a Presentation",
          defn="Timing plans the minutes per slide so the talk fits the slot with room for questions.",
          simple="Rehearse aloud with a timer; plan to finish early.",
          grid=dict(label="Does the storyboard fit?", widths=[2.4, 1.4], rows=[
              ["Item", "Minutes"],
              ["Planned talk", "11"],
              ["Slot", "12 + 5 for questions"],
              ["Slide to drop if late", "Slide 8 (option detail)"]]),
          watch="Running over cuts the questions, which is where decisions are made."),

        T("4.9.9", "Accessible and Robust Decks",
          defn="An accessible deck works for everyone: good contrast, not colour alone, alt text, readable fonts; a robust deck "
               "works on any machine.",
          simple="Review › Check Accessibility; embed fonts; bring a PDF backup.",
          grid=dict(label="Accessibility and robustness check", widths=[2, 2.8], rows=[
              ["Risk", "Fix"],
              ["Colour-only meaning", "Add value labels"],
              ["Charts without alt text", "Alt text stating the finding"],
              ["Linked data breaks", "Paste charts as pictures in the final deck"]]),
          watch="About 1 in 12 men has a colour-vision deficiency: red-green alone fails them."),

        T("4.9.10", "Delivering the Analysis",
          defn="Delivery is how you open, move through and close the talk: headline first, evidence next, ask last.",
          simple="Open with the finding; close with the decision you need.",
          grid=dict(label="Open and close the Atlantic Pearl talk", widths=[1, 3.8], size=11.5, rows=[
              ["Moment", "What to say"],
              ["Open", f"Cancellations cost {Nm(CANV / 1e3, 2).replace('m', 'bn')} of bookings in 2023-24."],
              ["Close", "We ask the board to approve a deposit policy for bookings made 91+ days ahead."]]),
          watch="Ending on 'any questions?' wastes the strongest moment; end on the ask."),
    ],
    takeaways=["Storyboard first: roles and minutes per slide.",
               "One message per slide, stated in the title.",
               "Big, simple, highlighted charts and short tables.",
               "Consistent and accessible design.",
               "Prepare questions; open with the finding, close with the ask."],
)


M410 = dict(
    id="4.10", title="Formula Auditing, Protection and Light Automation",
    subtitle="Finding errors in a model inherited from someone else, protecting a workbook without breaking it, and automating "
             "repetition with a recorded macro.",
    objectives=["Audit an inherited workbook with independent checks.",
                "Trace precedents and dependents; show and evaluate formulas.",
                "Use Error Checking and the Watch Window.",
                "Protect sheets and workbooks at the right level.",
                "Record a simple macro and document it."],
    topics=[
        T("4.10.1", "Why Auditing Matters",
          defn="Auditing checks a workbook's formulas, totals and references before you trust or share its numbers.",
          simple="Recompute the key outputs a second way and compare.",
          grid=dict(label="A colleague's FY2025 zone shares, with three hidden errors (tDash25)", widths=[1.4, 1.2, 1.2, 0.9], rows=[
              ["Zone", "Submitted", "Correct", "Check"]] + [
              [z, P(s_) + (" (typed)" if i == 2 else ""), P(c), "ERROR"] for i, (z, s_, c) in enumerate(zip(ZO.index, SUB, COR))]),
          watch="Every submitted share is wrong, yet the sheet looks finished. The next three topics find why."),

        T("4.10.2", "Trace Precedents and Dependents",
          defn="Formulas › Trace Precedents draws arrows to the cells a formula uses; Trace Dependents shows the cells that use it.",
          simple="A blue box shows the range a formula reads: check it covers every row.",
          grid=dict(label="Tracing the submitted Total", widths=[2, 2.6], rows=[
              ["Item", "Result"],
              ["Total formula", "=SUM(C9:C11): 3 rows, not 4"],
              ["Submitted total", N(SUBTOT)],
              ["Correct total", N(ZO.sum())],
              ["Missing", f"South-East, {N(ZO['South-East'])}"]]),
          watch="Error 1: a SUM that stops one row short overstates every share."),

        T("4.10.3", "Show Formulas and Evaluate Formula",
          defn="Show Formulas (Ctrl+`) displays every formula instead of its result; Evaluate Formula steps through one formula "
               "calculation by calculation.",
          simple="Typed numbers stand out among formulas; Evaluate shows which cells a formula really uses.",
          grid=dict(label="The other two errors", widths=[1.5, 1.2, 2.3], rows=[
              ["Cell", "Shows", "Problem"],
              ["North-Central share", "0.15", "Error 2: a typed number, not a formula"],
              ["South-East share", "=C12/C12", "Error 3: no $ on the total, so the copy drifted"]]),
          watch="Fix error 3 with an absolute reference: =C12/$C$13."),

        T("4.10.4", "Error Checking and Watch Window",
          defn="Error Checking (Formulas tab) flags suspicious formulas with green triangles; the Watch Window shows chosen cells "
               "while you work elsewhere.",
          simple="Add the Total to the Watch Window, then fix the errors and watch it change.",
          grid=dict(label="After fixing the three errors", widths=[2.2, 2], rows=[
              ["Check", "Result"],
              ["Total", N(ZO.sum())],
              ["Shares add up to", P(sum(COR), 0)],
              ["Check column", "OK on all four rows"]]),
          watch="A green triangle 'Formula omits adjacent cells' would have caught error 1."),

        T("4.10.5", "Protecting Sheets and Workbooks",
          defn="Sheet protection stops locked cells being edited; workbook protection stops sheets being added, deleted or "
               "renamed; encryption with a password stops the file being opened.",
          simple="Unlock the input cells (Format Cells › Protection), then Review › Protect Sheet.",
          grid=dict(label="FY2026 plan: only the growth input can change", widths=[1.8, 1.6, 1.4], rows=[
              ["Cell", "Status", "Value"],
              ["Growth assumption", "Unlocked (yellow)", "12%"],
              ["FY2026 plan", "Locked formula", Nm(TOT25 * 1.12)]]),
          watch="Sheet protection prevents accidents, not theft; only File › Info › Encrypt with Password protects confidentiality."),

        T("4.10.6", "Recording a Macro",
          defn="A macro is a recorded sequence of actions saved as VBA code that can be replayed with one click.",
          simple="View › Macros › Record Macro, do the steps, Stop Recording; save as .xlsm.",
          code="Sub Refresh_And_Export_Report()\n    ThisWorkbook.RefreshAll\n    Sheets(\"4.8 Report Writing\").ExportAsFixedFormat _\n"
               "        Type:=xlTypePDF, Filename:=\"Mercyland_Report.pdf\"\nEnd Sub",
          code_label="VBA (Alt+F11 to view)",
          points=["Record with relative references if the steps should work anywhere.", "Test on a copy first.",
                  "Macros cannot be undone with Ctrl+Z."],
          watch="Only enable macros in files from a trusted source; macros are a common route for malware."),

        T("4.10.7", "Documenting Automation",
          defn="Automation documentation records each macro or script: name, purpose, owner, inputs, outputs, schedule and last "
               "test.",
          simple="Keep a log sheet in the workbook and comments in the code.",
          grid=dict(label="Log for the refresh-and-export macro", widths=[1.3, 3.5], rows=[
              ["Field", "Entry"],
              ["Name", "Refresh_And_Export_Report"],
              ["Purpose", "Monthly refresh and PDF of the Mercyland report"],
              ["Owner / tested", "Reporting analyst / 2024-01-15"]]),
          watch="An undocumented macro becomes an unmaintainable black box when its author leaves."),
    ],
    takeaways=["Audit before you trust: recompute key outputs independently.",
               "Trace, Show Formulas and Evaluate find wrong ranges, typed values and drifting references.",
               "Error Checking and the Watch Window keep you honest while fixing.",
               "Protection prevents accidents; encryption protects confidentiality.",
               "Record a macro for repeated steps, and document it."],
)

MODULES = [M47, M48, M49, M410]
