"""Week 1 workbook, modules 1.4 - 1.6 (edition 3)."""
from __future__ import annotations

import datetime as dt

import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell

from w1x import ModuleSheet
from xlhelp import LAST_COL

N = "naira"


def A(r, c):
    return xl_rowcol_to_cell(r, c)


# =====================================================================================================
def m14(b, D):
    PAY, FAC, DEPT, HSEG, SEN, COUR = D.PAY, D.FAC, D.DEPT, D.HSEG, D.SEN, D.COUR
    m = ModuleSheet(b, "1.4 Industry Applications", "1.4  Applications of Data Analytics Across Industries",
                    "Tutor walk-through after Module 1.4: the key measure of each industry, worked on real summary tables",
                    "1.4.1 - 1.4.10")
    m.dataset([("tPayMix", "Below", "SuperMart Online: orders and revenue by payment method (11,000 orders)."),
               ("tFaculty", "Below", "Kingsway University: students, average score and at-risk students by faculty."),
               ("tDept23", "Below", "Mercyland Specialist Hospital 2023: admissions, average stay and 30-day readmissions by department."),
               ("tHotelSeg", "Below", "Atlantic Pearl Hotel 2023-2024: bookings and cancellations by market segment."),
               ("tSeniority", "Below", "ETEH Talent Salary Survey: people, median salary and average experience by seniority."),
               ("tCouriers", "Below", "SuperMart Online: deliveries and delivered parcels by courier."),
               ("tCardTxn, tOreDaily", "Below", "Illustrative practice data (made up for the skill; not a bootcamp dataset).")])
    m.h("tPayMix")
    m.s.table(PAY[["Payment_Method", "Orders", "Revenue"]], "tPayMix", fmts={"Orders": "int", "Revenue": N},
              calc={"Share": ("=[@Orders]/SUM([Orders])", "pct", PAY.Share.tolist())}, widths={"Payment_Method": 17})
    m.h("tFaculty")
    m.s.table(FAC[["Faculty", "Students", "Avg_Score", "At_Risk"]], "tFaculty", fmts={"Students": "int", "Avg_Score": "dec1", "At_Risk": "int"},
              calc={"At_Risk_Rate": ("=[@At_Risk]/[@Students]", "pct", FAC.At_Risk_Rate.tolist())})
    m.h("tDept23")
    m.s.table(DEPT[["Department", "Admissions", "Avg_LOS", "Readmitted"]], "tDept23", fmts={"Admissions": "int", "Avg_LOS": "dec1", "Readmitted": "int"},
              calc={"Readmission_Rate": ("=[@Readmitted]/[@Admissions]", "pct", DEPT.Readmission_Rate.tolist())}, widths={"Department": 26})
    m.h("tHotelSeg")
    m.s.table(HSEG[["Sales_Channel", "Bookings", "Cancelled"]], "tHotelSeg", fmts={"Bookings": "int", "Cancelled": "int"},
              calc={"Cancel_Rate": ("=[@Cancelled]/[@Bookings]", "pct", HSEG.Cancel_Rate.tolist())}, widths={"Sales_Channel": 20})
    m.h("tSeniority")
    m.s.table(SEN, "tSeniority", fmts={"People": "int", "Median_Salary": N, "Avg_Experience": "dec1"})
    m.h("tCouriers")
    m.s.table(COUR[["Courier", "Deliveries", "Delivered"]], "tCouriers", fmts={"Deliveries": "int", "Delivered": "int"},
              calc={"Delivered_Rate": ("=[@Delivered]/[@Deliveries]", "pct", COUR.Delivered_Rate.tolist())}, widths={"Courier": 17})

    # ---------------------------------------------------------------- 1.4.1
    pm = PAY.set_index("Payment_Method")
    m.topic("1.4.1", "Retail and e-commerce",
            "Retailers track what sells, where, when and at what profit; online shops add baskets, payments and deliveries.")
    m.steps(["Average order value (AOV) = total revenue ÷ number of orders.",
             "Payment share = orders for one method ÷ all orders (the Share column).",
             "Add two shares together to see how much depends on two gateways."])
    m.worked([("Average order value", "=SUM(tPayMix[Revenue])/SUM(tPayMix[Orders])", PAY.Revenue.sum() / PAY.Orders.sum(), N, ""),
              ("Paystack + Flutterwave share", '=SUMIFS(tPayMix[Share],tPayMix[Payment_Method],"Paystack")+SUMIFS(tPayMix[Share],tPayMix[Payment_Method],"Flutterwave")',
               pm.Share["Paystack"] + pm.Share["Flutterwave"], "pct", "Over half of all orders.")])
    m.turn()
    m.ex("What share of orders were Cash on Delivery?", '=SUMIFS(tPayMix[Share],tPayMix[Payment_Method],"Cash on Delivery")', pm.Share["Cash on Delivery"],
         "pct", short="Cash on Delivery share")
    m.ex("What was the average order value for USSD orders?", '=SUMIFS(tPayMix[Revenue],tPayMix[Payment_Method],"USSD")/SUMIFS(tPayMix[Orders],tPayMix[Payment_Method],"USSD")',
         pm.Revenue["USSD"] / pm.Orders["USSD"], N, short="USSD average order value")

    # ---------------------------------------------------------------- 1.4.2
    fa = FAC.set_index("Faculty")
    m.topic("1.4.2", "Education and universities",
            "Universities track admission, attendance, results and progression, and use the data to support students before they fail.")
    m.steps(["At-risk rate = at-risk students ÷ students (the At_Risk_Rate column).",
             "University rate = total at-risk ÷ total students (from totals, not an average of rates).",
             "Compare each faculty with the university rate."])
    m.worked([("University at-risk rate", "=SUM(tFaculty[At_Risk])/SUM(tFaculty[Students])", FAC.At_Risk.sum() / FAC.Students.sum(), "pct", ""),
              ("Highest faculty rate", "=MAX(tFaculty[At_Risk_Rate])", FAC.At_Risk_Rate.max(), "pct", fa.At_Risk_Rate.idxmax())])
    m.turn()
    m.ex("What is the Engineering at-risk rate?", '=SUMIFS(tFaculty[At_Risk_Rate],tFaculty[Faculty],"Engineering")', fa.At_Risk_Rate["Engineering"], "pct",
         short="Engineering at-risk rate")
    m.ex("How many faculties have an at-risk rate above the university rate?", '=COUNTIF(tFaculty[At_Risk_Rate],">"&SUM(tFaculty[At_Risk])/SUM(tFaculty[Students]))',
         int((FAC.At_Risk_Rate > FAC.At_Risk.sum() / FAC.Students.sum()).sum()), "int", short="Faculties above the university rate")

    # ---------------------------------------------------------------- 1.4.3
    overall = DEPT.Readmitted.sum() / DEPT.Admissions.sum()
    m.topic("1.4.3", "Healthcare and hospitals",
            "Hospitals measure bed occupancy, average length of stay (ALOS) and 30-day readmissions to improve care and use scarce beds well.")
    m.steps(["Readmission rate = readmitted ÷ admissions.",
             "Hospital ALOS from department averages must be weighted by admissions: SUMPRODUCT(admissions, ALOS) ÷ total admissions.",
             "Flag departments above the hospital rate."])
    m.worked([("Hospital readmission rate, 2023", "=SUM(tDept23[Readmitted])/SUM(tDept23[Admissions])", overall, "pct", ""),
              ("Hospital ALOS (weighted), days", "=SUMPRODUCT(tDept23[Admissions],tDept23[Avg_LOS])/SUM(tDept23[Admissions])",
               float((DEPT.Admissions * DEPT.Avg_LOS).sum() / DEPT.Admissions.sum()), "dec1", "SUMPRODUCT multiplies pairs, then adds (Week 2).")])
    m.turn()
    m.ex("How many departments have a readmission rate above the hospital rate?", '=COUNTIF(tDept23[Readmission_Rate],">"&SUM(tDept23[Readmitted])/SUM(tDept23[Admissions]))',
         int((DEPT.Readmission_Rate > overall).sum()), "int", short="Departments above the hospital rate")
    m.ex("How many admissions did Surgery have in 2023?", '=SUMIFS(tDept23[Admissions],tDept23[Department],"Surgery")',
         int(DEPT.set_index("Department").Admissions["Surgery"]), "int", short="Surgery admissions")

    # ---------------------------------------------------------------- 1.4.4
    usual = 45000
    tx = pd.DataFrame({"Txn_ID": [f"T0{i}" for i in range(1, 9)], "Hour": [9, 13, 3, 18, 2, 11, 20, 16],
                       "Amount": [12500, 48000, 900000, 23750, 15000, 260000, 31200, 8900],
                       "Country": ["Nigeria", "Nigeria", "Ghana", "Nigeria", "Nigeria", "Nigeria", "United Kingdom", "Nigeria"]})
    m.topic("1.4.4", "Banking and finance",
            "Banks use transaction data to detect fraud, score credit, predict churn and place branches and ATMs. A fraud rule flags transactions that do not fit a customer's normal pattern.")
    m.steps(["Put the customer's usual spend in a yellow input cell.",
             "Write the rule as one IF with OR: review if the amount is more than 5 × usual, OR the country is not Nigeria, OR the hour is before 5 am.",
             "Count the flags with COUNTIF."])
    inp = m.s.inputs([("Customer's usual transaction, ₦ (illustrative)", usual, N, "Change it and watch the flags update.")])
    U = inp["Customer's usual transaction, ₦ (illustrative)"]
    flag = ["Review" if (a > 5 * usual or c != "Nigeria" or h < 5) else "OK" for h, a, c in zip(tx.Hour, tx.Amount, tx.Country)]
    m.h("tCardTxn  (illustrative practice data)")
    m.s.table(tx, "tCardTxn", fmts={"Amount": N},
              calc={"Flag": (f'=IF(OR([@Amount]>5*{U},[@Country]<>"Nigeria",[@Hour]<5),"Review","OK")', None, flag)})
    m.worked([("Transactions flagged for review", '=COUNTIF(tCardTxn[Flag],"Review")', flag.count("Review"), "int", "")])
    m.turn()
    m.ex("Change the usual transaction to ₦60,000. How many transactions are flagged now? (Answer key: worked with ₦60,000.)",
         '=SUMPRODUCT(--((tCardTxn[Amount]>5*60000)+(tCardTxn[Country]<>"Nigeria")+(tCardTxn[Hour]<5)>0))',
         sum(1 for h, a, c in zip(tx.Hour, tx.Amount, tx.Country) if a > 300000 or c != "Nigeria" or h < 5), "int", short="Flags at ₦60,000 usual spend")

    # ---------------------------------------------------------------- 1.4.5
    m.topic("1.4.5", "Government and public administration",
            "Governments use data to plan schools, clinics and roads, set budgets and track results. Rates are often per 1,000 people, and personal data must be protected.")
    m.steps(["Rate per 1,000 = count ÷ (population ÷ 1,000).",
             "Before publishing, remove anything that identifies a person (names, phone numbers, exact addresses).",
             "Publish groups, never individuals."])
    inp = m.s.inputs([("Population served (illustrative assumption)", 1200000, "int", "Not from a dataset: an assumption for practice.")])
    P = inp["Population served (illustrative assumption)"]
    m.worked([("Mercyland admissions per 1,000 people, 2023", f"=SUM(tDept23[Admissions])/({P}/1000)", DEPT.Admissions.sum() / 1200, "dec1", "")])
    m.turn()
    m.ex("With a population of 1,500,000, how many admissions per 1,000 people is that?", "=SUM(tDept23[Admissions])/(1500000/1000)",
         DEPT.Admissions.sum() / 1500, "dec1", tol=0.051, short="Admissions per 1,000 (1.5m people)")
    m.ex_text("The Mercyland admissions table has Episode_ID, Patient_Name, Age, Department and Bill_Amount_NGN. Which column must be removed before publishing?",
              "Patient_Name", alts=("patient name",))

    # ---------------------------------------------------------------- 1.4.6
    hs = HSEG.set_index("Sales_Channel")
    m.topic("1.4.6", "Marketing and customer management",
            "Marketing analytics finds which customers and channels bring the most value, and which cost more than they return.")
    m.steps(["Cancellation rate = cancelled ÷ bookings, for each segment.",
             "Overall rate from totals.",
             "Compare channels as a ratio: how many times higher is one rate than another?"])
    ota, corp = hs.Cancel_Rate["Online Travel Agent"], hs.Cancel_Rate["Corporate"]
    m.worked([("Overall cancellation rate", "=SUM(tHotelSeg[Cancelled])/SUM(tHotelSeg[Bookings])", HSEG.Cancelled.sum() / HSEG.Bookings.sum(), "pct", ""),
              ("Online travel agent ÷ corporate rate", '=SUMIFS(tHotelSeg[Cancel_Rate],tHotelSeg[Sales_Channel],"Online Travel Agent")/SUMIFS(tHotelSeg[Cancel_Rate],tHotelSeg[Sales_Channel],"Corporate")',
               ota / corp, "dec1", "About three times as likely to cancel.")])
    m.turn()
    m.ex("How many Online Travel Agent bookings were cancelled?", '=SUMIFS(tHotelSeg[Cancelled],tHotelSeg[Sales_Channel],"Online Travel Agent")',
         int(hs.Cancelled["Online Travel Agent"]), "int", short="OTA cancellations")
    m.ex("What share of all bookings came through Online Travel Agents?", '=SUMIFS(tHotelSeg[Bookings],tHotelSeg[Sales_Channel],"Online Travel Agent")/SUM(tHotelSeg[Bookings])',
         hs.Bookings["Online Travel Agent"] / HSEG.Bookings.sum(), "pct", short="OTA share of bookings")

    # ---------------------------------------------------------------- 1.4.7
    sn = SEN.set_index("Seniority")
    m.topic("1.4.7", "Human resources (people analytics)",
            "People analytics uses workforce data to improve decisions about hiring, pay, performance and retention. Report groups, never named individuals.")
    m.steps(["Compare medians across levels (medians are not pulled up by a few very high salaries).",
             "Express the gap as a ratio: Executive median ÷ Junior median.",
             "Check headcounts: small groups give less reliable medians."])
    m.worked([("Executive ÷ Junior median salary", '=SUMIFS(tSeniority[Median_Salary],tSeniority[Seniority],"Executive")/SUMIFS(tSeniority[Median_Salary],tSeniority[Seniority],"Junior")',
               sn.Median_Salary["Executive"] / sn.Median_Salary["Junior"], "dec1", "Times higher."),
              ("People in the survey", "=SUM(tSeniority[People])", int(SEN.People.sum()), "int", "")])
    m.turn()
    m.ex("How much higher is the Senior median than the Mid median, in naira?", '=SUMIFS(tSeniority[Median_Salary],tSeniority[Seniority],"Senior")-SUMIFS(tSeniority[Median_Salary],tSeniority[Seniority],"Mid")',
         float(sn.Median_Salary["Senior"] - sn.Median_Salary["Mid"]), N, short="Senior minus Mid median")
    m.ex_text("Which seniority level has the fewest people?", sn.People.idxmin(), formula="=INDEX(tSeniority[Seniority],MATCH(MIN(tSeniority[People]),tSeniority[People],0))",
              short="Smallest seniority group")

    # ---------------------------------------------------------------- 1.4.8
    m.topic("1.4.8", "Logistics and transportation",
            "Logistics analytics checks that goods arrive on time at the lowest cost: routes, couriers, stock and delivery success.")
    m.steps(["Delivered rate = delivered ÷ deliveries, per courier.",
             "Overall rate from totals.",
             "Spread = best rate minus worst rate: a small spread means no courier stands out."])
    m.worked([("Overall delivered rate", "=SUM(tCouriers[Delivered])/SUM(tCouriers[Deliveries])", COUR.Delivered.sum() / COUR.Deliveries.sum(), "pct", ""),
              ("Spread: best minus worst rate", "=MAX(tCouriers[Delivered_Rate])-MIN(tCouriers[Delivered_Rate])",
               COUR.Delivered_Rate.max() - COUR.Delivered_Rate.min(), "pct", "Under 3 points: couriers perform alike.")])
    m.turn()
    m.ex_text("Which courier has the highest delivered rate?", COUR.Courier[COUR.Delivered_Rate.idxmax()],
              formula="=INDEX(tCouriers[Courier],MATCH(MAX(tCouriers[Delivered_Rate]),tCouriers[Delivered_Rate],0))", short="Best courier")
    m.ex("How many parcels did Gokada deliver successfully?", '=SUMIFS(tCouriers[Delivered],tCouriers[Courier],"Gokada")',
         int(COUR.set_index("Courier").Delivered["Gokada"]), "int", short="Gokada delivered")

    # ---------------------------------------------------------------- 1.4.9
    ore = pd.DataFrame({"Day": ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"], "Plan_t": [1200] * 5 + [1000] * 2,
                        "Actual_t": [1150, 1240, 1180, 1265, 990, 1020, 1005]})
    m.topic("1.4.9", "Geology, mining and natural resources",
            "Mining companies collect drilling, sensor and production data. Analytics finds resources, controls cost and tracks production against plan (grade control).")
    m.steps(["Variance = actual tonnes - planned tonnes, per day.",
             "Met plan? = IF(actual >= plan, \"Yes\", \"No\").",
             "Week total variance and days below plan."])
    m.h("tOreDaily  (illustrative practice data)")
    m.s.table(ore, "tOreDaily", fmts={"Plan_t": "int", "Actual_t": "int"},
              calc={"Variance_t": ("=[@Actual_t]-[@Plan_t]", "int", (ore.Actual_t - ore.Plan_t).tolist()),
                    "Met_Plan": ('=IF([@Actual_t]>=[@Plan_t],"Yes","No")', None, ["Yes" if a >= p else "No" for a, p in zip(ore.Actual_t, ore.Plan_t)])})
    m.worked([("Week variance, tonnes", "=SUM(tOreDaily[Variance_t])", int((ore.Actual_t - ore.Plan_t).sum()), "int", "Slightly below plan.")])
    m.turn()
    m.ex("On how many days was production below plan?", '=COUNTIF(tOreDaily[Met_Plan],"No")', int((ore.Actual_t < ore.Plan_t).sum()), "int",
         short="Days below plan")
    m.ex("What share of planned tonnes was produced over the week?", "=SUM(tOreDaily[Actual_t])/SUM(tOreDaily[Plan_t])", ore.Actual_t.sum() / ore.Plan_t.sum(),
         "pct", short="Week attainment")

    # ---------------------------------------------------------------- 1.4.10
    m.topic("1.4.10", "Real-world data-driven decisions",
            "In each case below, data replaced a guess with evidence and pointed to a specific action.")
    m.steps(["State the evidence as a number (live formula).",
             "State the decision it supports.",
             "Say what you would measure next to check the decision worked."])
    m.worked([("Atlantic Pearl: online travel agent cancellation rate", '=SUMIFS(tHotelSeg[Cancel_Rate],tHotelSeg[Sales_Channel],"Online Travel Agent")', ota, "pct",
               "Decision: ask for deposits on agency bookings."),
              ("Kingsway: students at risk", "=SUM(tFaculty[At_Risk])", int(FAC.At_Risk.sum()), "int", "Decision: target tutoring at flagged students."),
              ("SuperMart Online: Paystack + Flutterwave share", '=SUMIFS(tPayMix[Share],tPayMix[Payment_Method],"Paystack")+SUMIFS(tPayMix[Share],tPayMix[Payment_Method],"Flutterwave")',
               pm.Share["Paystack"] + pm.Share["Flutterwave"], "pct", "Decision: keep both gateways reliable at checkout.")])
    m.turn()
    m.ex_open("Using tDept23, write one evidence → decision pair for Mercyland.",
              f"Evidence: Surgery had the highest 30-day readmission rate in 2023 ({DEPT.Readmission_Rate.max():.1%}). Decision: phone every Surgery patient within 7 days of discharge and track the rate monthly.")
    m.finish()


# =====================================================================================================
def m15(b, D):
    ST, SA = D.ST, D.SALES
    m = ModuleSheet(b, "1.5 Excel Interface", "1.5  Introduction to Microsoft Excel",
                    "Tutor walk-through after Module 1.5: editions, file formats, saving safely, the parts of the window and moving around fast",
                    "1.5.1 - 1.5.17")
    m.dataset([("tStores25", "Sheet Data_Stores25 (starts in A1)", "12 store rows: small enough to see every cell; used for the Name Box, formula bar and status bar."),
               ("tSales", "Sheet Data_Sales (starts in A1)", f"{len(SA):,} sales lines: big enough to practise navigation shortcuts."),
               ("VAT_Rate", "A named cell in 1.5.1, below", "A defined name used to practise the Name Box.")])

    # ---------------------------------------------------------------- 1.5.1
    m.topic("1.5.1", "What Microsoft Excel is",
            "Excel is a spreadsheet application: a grid of cells, organised into worksheets and workbooks, used to store, calculate, analyse and visualise data. "
            "Formulas recalculate instantly when an input changes.")
    m.steps(["Look at the yellow inputs: price, quantity and the VAT rate.",
             "The results are formulas: Total = Price × Quantity; VAT = Total × VAT_Rate.",
             "Change Quantity and watch every result update without retyping."])
    inp = m.s.inputs([("Unit price, ₦", 50000, N, ""), ("Quantity (change me)", 3, "int", ""), ("VAT rate", 0.075, "pct", "This cell is named VAT_Rate (see 1.5.11).")])
    Pc, Qc, Vc = inp["Unit price, ₦"], inp["Quantity (change me)"], inp["VAT rate"]
    b.wb.define_name("VAT_Rate", f"='1.5 Excel Interface'!{Vc}")
    vat_addr = Vc.replace("$", "")
    m.worked([("Total before VAT", f"={Pc}*{Qc}", 150000, N, ""),
              ("VAT", f"={Pc}*{Qc}*VAT_Rate", 11250, N, "Uses the name VAT_Rate instead of a cell address."),
              ("Total including VAT", f"={Pc}*{Qc}*(1+VAT_Rate)", 161250, N, "")])
    m.turn()
    m.ex("Change Quantity to 5. What is the new Total including VAT?", "=50000*5*(1+VAT_Rate)", 268750, N, short="Total incl. VAT for 5 units")

    # ---------------------------------------------------------------- 1.5.2
    m.topic("1.5.2", "Excel editions: 2016, 2019, 2021 and Microsoft 365",
            "One-off editions (2016, 2019, 2021, 2024) never gain new functions; Microsoft 365 is a subscription that keeps updating. Support for 2016 and 2019 ended in October 2025.")
    m.steps(["File › Account › About Excel shows your version.",
             "Test for a newer function: type the formula below in the white cell.",
             "If it shows #NAME?, your edition does not have that function."])
    m.turn()
    r = m.r
    inp_c = A(r, 8)
    m.ex_formula_live('Type =XLOOKUP("ST02",tStores25[Store_ID],tStores25[Store_Name]) in the white cell and press Enter.',
                      f'=IF(ISBLANK({inp_c}),"",IFERROR(IF({inp_c}="Lagos - Lekki","✓ Your Excel has XLOOKUP (2021 or 365)","✗ Check the formula"),"Your edition has no XLOOKUP (2016 or 2019)"))',
                      "", key="Lagos - Lekki, if your edition is Excel 2021 or Microsoft 365.")
    m.ex_text("Which one-off edition first added XLOOKUP, FILTER and UNIQUE? (2016, 2019 or 2021)", "2021", alts=("excel 2021",))
    m.ex_text("Which edition keeps receiving new functions? (Type: 365)", "365", alts=("microsoft 365", "office 365"))

    # ---------------------------------------------------------------- 1.5.3
    m.topic("1.5.3", "Excel platforms: web, desktop and Mac",
            "The same workbook opens on Windows, Mac and in a browser, but not every feature exists everywhere. This bootcamp assumes Excel for Windows (desktop).")
    m.steps(["Note which platform you are using.",
             "If you use the web or a Mac, check the platform table on the slide before each practical: Power Pivot and VBA are Windows features."])
    m.turn()
    m.ex_text("Which platform cannot run VBA macros? (Windows, Mac or Web)", "Web", alts=("excel for the web", "browser"))
    m.ex_text("Which platform has no Power Pivot data model? Type Mac or Windows.", "Mac")

    # ---------------------------------------------------------------- 1.5.4
    m.topic("1.5.4", "Excel file formats: .xlsx, .xlsm, .xls, .xlsb and .csv",
            "The file extension decides what a file can hold. Use .xlsx by default and .csv only for moving raw data between systems.")
    m.steps(["File › Save a Copy › choose CSV UTF-8 (Comma delimited) in the type list.",
             "Read the warning: only the active sheet is saved, as plain values.",
             "Open the .csv: formulas, formatting and the other sheets are gone. Keep the .xlsx as your working file."])
    m.turn()
    m.ex_text("Which format can store macros? (.xlsx, .xlsm, .xls, .xlsb or .csv)", ".xlsm", alts=("xlsm",))
    m.ex_text("Which format keeps only the values of one sheet?", ".csv", alts=("csv",))
    m.do([("Save a copy of Data_Stores25 as CSV, then open the CSV in Excel.", "One sheet, values only: the Margin column shows numbers, not formulas.")])

    # ---------------------------------------------------------------- 1.5.5
    m.topic("1.5.5", "Creating, opening and saving workbooks safely",
            "Save early, save often, use clear file names and never overwrite the only copy of raw data.")
    m.steps(["Ctrl+N creates a new workbook; Ctrl+O opens one.",
             "F12 opens Save As: name the file clearly, e.g. ETEH_Week1_Examples_YourName.xlsx.",
             "Ctrl+S saves; save after every exercise.",
             "A downloaded file opens in Protected View: click Enable Editing only if you trust the source."])
    m.turn()
    m.ex_text("Which key opens Save As?", "F12")
    m.do([("Save this workbook as ETEH_Week1_Examples_YourName.xlsx in your bootcamp folder.", "The new name appears in the title bar.")])

    # ---------------------------------------------------------------- 1.5.6
    m.topic("1.5.6", "AutoSave, AutoRecover and Version History",
            "AutoSave saves every change (files on OneDrive or SharePoint). AutoRecover keeps a recovery copy in case Excel crashes. Version History lets you restore an earlier version.")
    m.steps(["File › Options › Save: read the AutoRecover interval.",
             "If the file is on OneDrive, the AutoSave switch is at the top left.",
             "File › Info › Version History lists earlier versions you can open."])
    m.turn()
    m.ex_text("What is the default AutoRecover interval, in minutes?", "10")
    m.do([("Set AutoRecover to every 5 minutes (File › Options › Save).", "The box shows 5 minutes.")])

    # ---------------------------------------------------------------- 1.5.7
    m.topic("1.5.7", "Excel window anatomy",
            "Learn the names of the parts once: title bar, ribbon, Name Box, formula bar, grid, sheet tabs and status bar. Every later lesson refers to them.")
    m.steps(["Go to Data_Stores25 and click cell E2.",
             "Name Box (left of the formula bar): shows E2.",
             "Formula bar: shows what is really in the cell.",
             "Status bar (bottom): shows Ready, and quick totals when several numbers are selected."])
    m.turn()
    m.ex_text("Click cell I2 on Data_Stores25. Is it a typed value or a formula? (Type: typed or formula)", "formula")
    m.ex_text("What does the Name Box show when you click the Sales value of Lagos - Lekki on Data_Stores25?", "E3")

    # ---------------------------------------------------------------- 1.5.8
    m.topic("1.5.8", "The ribbon: tabs, groups and galleries",
            "The ribbon is the command bar across the top. Each tab holds groups of related commands; some open galleries of ready-made options.")
    m.steps(["Click each tab once and read its group names.",
             "Ctrl+F1 hides or shows the ribbon.",
             "Click inside a table: the Table Design tab appears only when needed (a contextual tab)."])
    m.turn()
    for cmd, tab in [("Remove Duplicates", "Data"), ("Freeze Panes", "View"), ("Conditional Formatting", "Home"), ("PivotTable", "Insert")]:
        m.ex_text(f"Which tab holds {cmd}?", tab)

    # ---------------------------------------------------------------- 1.5.9
    m.topic("1.5.9", "Quick Access Toolbar and KeyTips",
            "The Quick Access Toolbar (QAT) holds your most-used commands. KeyTips are letters that appear when you press Alt, so every command can be run from the keyboard.")
    m.steps(["Click the small arrow at the end of the QAT › More Commands.",
             "Choose Sort, click Add, OK.",
             "Press Alt: letters appear on each tab. Press H for Home, then follow the letters shown."])
    m.turn()
    m.ex_text("Press Alt, then W. Which tab's KeyTips appear?", "View")
    m.do([("Add Sort and Remove Duplicates to your Quick Access Toolbar.", "Two new icons in the QAT.")])

    # ---------------------------------------------------------------- 1.5.10
    m.topic("1.5.10", "Backstage view and workbook properties",
            "Backstage view is the full-screen File menu where you manage the workbook as a file: Info, New, Open, Save As, Print, Share, Export and Options.")
    m.steps(["Click File.", "Info › Properties: add a Title and Tags.", "Press Esc to return to the sheet."])
    m.turn()
    m.ex_text("Which Backstage page shows the author and the file's properties? (Info, New, Open, Print or Options)", "Info")
    m.do([("Add the Title 'Week 1 practice' under File › Info › Properties.", "The Title appears in the properties list.")])

    # ---------------------------------------------------------------- 1.5.11
    m.topic("1.5.11", "Name Box and formula bar",
            "The Name Box shows the address or name of the active cell and lets you jump to any cell. The formula bar shows the actual contents: the typed value or the formula behind a result.")
    m.steps(["Click in the Name Box, type K500 and press Enter: Excel jumps there.",
             "Open the Name Box drop-down: defined names (VAT_Rate) and table names are listed.",
             "Click VAT_Rate: Excel selects the named cell. Formulas can use the name instead of the address.",
             "Ctrl+Shift+U expands the formula bar for long formulas."])
    m.turn()
    m.ex_text("Choose VAT_Rate in the Name Box drop-down. Which cell address is selected?", vat_addr, alts=(Vc,))
    m.ex("In the white cell type =100000*VAT_Rate and press Enter. (The check reads the result.)", "=100000*VAT_Rate", 7500, N, short="₦100,000 × VAT_Rate")

    # ---------------------------------------------------------------- 1.5.12
    m.topic("1.5.12", "Column headers, row headers and the active cell",
            "Columns are lettered and rows numbered; a cell's address is its column then its row (C4). The active cell has the bold border.")
    m.steps(["Click a column letter to select the whole column (Ctrl+Space from the keyboard).",
             "Click a row number to select the whole row (Shift+Space).",
             "Double-click the right edge of a column letter to AutoFit its width."])
    m.turn()
    m.ex_text("On Data_Stores25, which column letter holds Profit?", "F")
    m.ex_text("On Data_Sales, what is in cell E2?", SA.Store_Name[0], formula="=Data_Sales!E2", short="Data_Sales E2")

    # ---------------------------------------------------------------- 1.5.13
    m.topic("1.5.13", "Status bar and view modes",
            "The status bar shows Excel's mode and instant statistics for whatever you select, with no formula needed. View buttons switch between Normal, Page Layout and Page Break Preview.")
    m.steps(["On Data_Stores25 select E2:E13 (the 12 Sales values).",
             "Read Average, Count and Sum on the status bar.",
             "Right-click the status bar to add Minimum and Maximum."])
    m.turn()
    m.ex("What Sum does the status bar show for E2:E13?", "=SUM(tStores25[Sales])", ST.Sales.sum(), N, short="Status bar Sum")
    m.ex("What Average does it show?", "=AVERAGE(tStores25[Sales])", ST.Sales.mean(), N, tol=1, short="Status bar Average")
    m.ex("Now select the whole column E by clicking its letter. What Count does the status bar show? (Count includes the header text.)",
         "=COUNTA(Data_Stores25!E:E)", len(ST) + 1, "int", short="Count, whole column E")

    # ---------------------------------------------------------------- 1.5.14
    m.topic("1.5.14", "The grid: 1,048,576 rows × 16,384 columns",
            "A worksheet has 1,048,576 rows and 16,384 columns (A to XFD). That is far more than most datasets need, but very large files become slow.")
    m.steps(["On Data_Sales click A1 and press Ctrl+Down: you jump to the last filled row.",
             "Press Ctrl+Down again: you reach row 1,048,576, the end of the sheet. Ctrl+Up goes back."])
    m.turn()
    m.ex("On Data_Sales, which row number holds the last sales line?", "=ROWS(tSales)+1", len(SA) + 1, "int", short="Last row of Data_Sales")
    m.ex_text("What is the last column letter of a worksheet?", "XFD")

    # ---------------------------------------------------------------- 1.5.15
    m.topic("1.5.15", "Navigating efficiently",
            "The mouse is slow on large data. Ctrl+Arrow, Ctrl+Home/End, Ctrl+PgUp/PgDn and F5 (Go To) move you anywhere instantly.")
    m.steps(["Ctrl+Home goes to A1; Ctrl+End to the last used cell.",
             "Ctrl+PgDn / Ctrl+PgUp move to the next / previous sheet.",
             "F5, type a name or address, Enter: Go To."])
    m.turn()
    m.ex_text("On Data_Sales press Ctrl+End. Which cell is selected?", f"K{len(SA) + 1}")
    m.ex_text("Press F5, type tStores25 and press Enter. Which sheet opens?", "Data_Stores25")

    # ---------------------------------------------------------------- 1.5.16
    m.topic("1.5.16", "Worksheets, sheet tabs and workbook structure",
            "A workbook is the file; a worksheet is one grid inside it. Organise it like a report: cover, raw data, workings, summary.")
    m.steps(["Shift+F11 (or the + button) inserts a sheet.",
             "Double-click a tab to rename it; right-click › Tab Color to colour it.",
             "Right-click › Move or Copy › tick Create a copy to duplicate a sheet."])
    m.turn()
    m.ex("How many worksheets does this workbook have? (Count the tabs.)", "=SHEETS()", 14, "int", short="Sheets in this workbook")
    m.do([("Make a copy of Data_Stores25 and rename the copy Practice_Stores.", "A new tab Practice_Stores; the original is untouched."),
          ("Colour the Practice_Stores tab green.", "The tab shows a green bar.")])

    # ---------------------------------------------------------------- 1.5.17
    m.topic("1.5.17", "Customising the Excel interface",
            "A few settings make Excel faster and safer for analysis work. Most live in File › Options.")
    m.steps(["File › Options › General: default font and number of sheets in a new workbook.",
             "File › Options › Customize Ribbon: add your own tab or group.",
             "View › Show: gridlines and headings on or off."])
    m.do([("Turn gridlines off on Practice_Stores, then back on.", "The grey cell lines disappear, then return."),
          ("Create a custom ribbon group 'My Analysis' on the Home tab with Remove Duplicates in it.", "A new group at the right end of the Home tab.")])
    m.finish()


# =====================================================================================================
def m16(b, D):
    ST, DI = D.ST, D.DIRTY
    F = b.F
    m = ModuleSheet(b, "1.6 Working with Data", "1.6  Working with Data in Excel",
                    "Tutor walk-through after Module 1.6: entering, editing, filling, formatting and checking data", "1.6.1 - 1.6.14")
    m.dataset([("Practice areas", "On this sheet", "Small editable ranges (white cells with a gold border) built for each skill. Their check cells "
                                                    "look at what you did and say ✓ when it is right."),
               ("tStores25", "Sheet Data_Stores25", "Store values copied into the practice areas."),
               ("tDirtySales", "Sheet Data_DirtySales", "Used to show the text-number trap in 1.6.2.")])

    # ---------------------------------------------------------------- 1.6.1
    m.topic("1.6.1", "Entering text, numbers and dates",
            "Type into the active cell and press Enter (down) or Tab (right). Excel decides the type from what you type: text aligns left, numbers and dates right.")
    m.steps(["Type each value from the left column into the white cell beside it and press Enter.",
             "Read what Excel stored: the live check uses ISNUMBER and ISTEXT.",
             "To keep leading zeros, type an apostrophe first: '007 is stored as text."])
    hdr = F["hdr"]
    m.ws.merge_range(m.r, 1, m.r, 2, "Type this", hdr)
    m.ws.write(m.r, 3, "Here", hdr)
    m.ws.merge_range(m.r, 4, m.r, 7, "What Excel stored (live)", hdr)
    m.ws.merge_range(m.r, 8, m.r, LAST_COL, "Expected", hdr)
    m.r += 1
    for typed, exp in [("Lagos - Ikeja", "Text"), ("474430", "Number"), ("2025-01-01", "Number: a date is stored as a number (45658)"),
                       ("007", "Number 7: the zeros are lost"), ("'007", "Text 007: the apostrophe keeps the zeros")]:
        c = A(m.r, 3)
        m.ws.merge_range(m.r, 1, m.r, 2, typed, F["lab"])
        m.ws.write_blank(m.r, 3, None, m.X["in"])
        m.ws.merge_range(m.r, 4, m.r, 7, "", m.X["chk"])
        m.ws.write_formula(m.r, 4, f'=IF(ISBLANK({c}),"",IF(ISNUMBER({c}),"Number: "&{c},IF(ISTEXT({c}),"Text: "&{c},"Other")))', m.X["chk"], "")
        m.ws.merge_range(m.r, 8, m.r, LAST_COL, exp, m.X["see"])
        m.r += 1
    m.r += 1
    m.turn()
    m.ex_text("Which keyboard shortcut enters today's date? (Type it as Ctrl+;)", "Ctrl+;", alts=("ctrl ;", "ctrl+semicolon"))

    # ---------------------------------------------------------------- 1.6.2
    nums = DI.Sales.map(lambda v: isinstance(v, (int, float)) and not pd.isna(v))
    m.topic("1.6.2", "The six Excel data types",
            "A data type is the kind of value a cell holds: Number, Text, Date/Time, Logical, Error or Array. The type decides what formulas can do with it.")
    m.steps(["Test a value with ISNUMBER, ISTEXT, ISLOGICAL or ISERROR.",
             "Remember: SUM silently skips numbers stored as text.",
             "COUNT counts numbers only; COUNTA counts every filled cell."])
    m.worked([("Is 474430 a number?", '=IF(ISNUMBER(474430),"Yes","No")', "Yes", None, ""),
              ("Is \"474430\" (in quotes = text) a number?", '=IF(ISNUMBER("474430"),"Yes","No")', "No", None, "Text that looks like a number."),
              ("Sales values SUM can add in tDirtySales", "=COUNT(tDirtySales[Sales])", int(nums.sum()), "int", ""),
              ("SUM of tDirtySales[Sales]", "=SUM(tDirtySales[Sales])", float(DI.Sales[nums].sum()), N, "Too low: 234 text values are skipped.")])
    m.turn()
    m.ex("How many Profit values in tDirtySales are numbers?", "=COUNT(tDirtySales[Profit])", int(DI.Profit.notna().sum()), "int", short="Numeric Profit values")
    m.ex_text("What data type is TRUE?", "Logical", alts=("boolean",))

    # ---------------------------------------------------------------- 1.6.3
    m.topic("1.6.3", "Editing, replacing and clearing cell contents",
            "Typing over a cell replaces it. F2 (or double-click) edits in place. Delete removes contents but keeps formats; Home › Editing › Clear lets you choose.")
    m.steps(["Click a white cell below and press F2 to edit in place; fix the typo; press Enter.",
             "For the number, edit to add the missing digit.",
             "For 'delete me', select it and press Delete.",
             "Made a mistake? Ctrl+Z undoes it."])
    m.turn()
    for wrong, right, kind in [("Lagso - Ikeja", "Lagos - Ikeja", "t"), ("Abuja - Wsue", "Abuja - Wuse", "t"),
                               (14638834, 146388345, "n"), ("delete me", None, "x")]:
        r = m.r
        c = A(r, 8)
        if kind == "t":
            fx = f'=IF(EXACT({c},"{right}"),"✓ Fixed","✗ Should read {right}")'
            task = f"Edit the white cell so it reads exactly: {right}"
            first = f"✗ Should read {right}"
        elif kind == "n":
            fx = f'=IF({c}={right},"✓ Fixed","✗ Should be {right:,}")'
            task = "A digit is missing: edit the number so it reads 146,388,345 (Lagos - Ikeja's FY2025 sales)."
            first = f"✗ Should be {right:,}"
        else:
            fx = f'=IF(ISBLANK({c}),"✓ Cleared","✗ Select the cell and press Delete")'
            task = "Clear the white cell with the Delete key."
            first = "✗ Select the cell and press Delete"
        m.ex_formula_live(task, fx, first)
        m.ws.write(r, 8, wrong, m.X["in"])

    # ---------------------------------------------------------------- 1.6.4
    m.topic("1.6.4", "Inserting, deleting, hiding and resizing rows and columns",
            "Insert and delete shift the cells around them and formulas adjust. Hiding keeps the data but takes it out of view, and SUM still includes hidden rows.")
    m.steps(["The small range below has one hidden row: look for the jump in row numbers.",
             "Compare SUM with SUBTOTAL(109,…), which ignores hidden rows.",
             "Select the rows either side of the gap, right-click › Unhide.",
             "Ctrl+Plus inserts and Ctrl+Minus deletes selected rows or columns."])
    five = ST.iloc[[0, 1, 4, 6, 11]][["Store_Name", "Sales"]].reset_index(drop=True)
    top = m.s.frame(five, fmts={"Sales": N})
    hid = top + 2
    m.ws.set_row(hid, None, None, {"hidden": True})
    rng = f"C{top + 1}:C{top + 5}"
    m.worked([("SUM of the five values (includes the hidden row)", f"=SUM({rng})", int(five.Sales.sum()), N, ""),
              ("SUBTOTAL(109) of the same range (skips hidden rows)", f"=SUBTOTAL(109,{rng})", int(five.Sales.sum() - five.Sales[2]), N,
               "The two differ until you unhide the row.")])
    m.turn()
    m.ex_text("Unhide the hidden row. Which store was hidden?", five.Store_Name[2], alts=(five.Store_Name[2].split(" - ")[0],))

    # ---------------------------------------------------------------- 1.6.5
    m.topic("1.6.5", "Copy, cut, paste and Paste Special",
            "Paste Special pastes only part of what was copied (values, formats, formulas) or transforms it as it pastes (transpose, add, multiply, divide).")
    m.steps(["Copy the five Source values (Ctrl+C).",
             "Click the first white Target cell › Ctrl+Alt+V › Values › OK.",
             "Type 1000000 in any empty cell and copy it.",
             "Select the five Target cells › Ctrl+Alt+V › Divide › OK. The values are now in millions. Delete the 1000000."])
    r0 = m.r
    m.ws.write(r0, 1, "Source (₦)", F["hdr"])
    m.ws.write(r0, 2, "Target", F["hdr"])
    for i, v in enumerate(five.Sales):
        m.ws.write_number(r0 + 1 + i, 1, int(v), b.nf(N))
        m.ws.write_blank(r0 + 1 + i, 2, None, m.X["in"])
    m.r = r0 + 7
    src, tgt = f"B{r0 + 2}:B{r0 + 6}", f"C{r0 + 2}:C{r0 + 6}"
    m.turn()
    m.ex_formula_live("Do steps 1-4. The check reads the Target cells.",
                      f'=IF(COUNT({tgt})<5,"Paste the five values first",IF(ABS(SUM({tgt})-SUM({src})/1000000)<0.001,"✓ Correct: now in millions",'
                      f'IF(SUM({tgt})=SUM({src}),"Values pasted: now divide by 1,000,000","✗ Not yet")))', "Paste the five values first",
                      key=f"Target total = {five.Sales.sum() / 1e6:,.6f} (millions).")

    # ---------------------------------------------------------------- 1.6.6
    m.topic("1.6.6", "AutoFill, series and custom lists",
            "AutoFill extends a pattern when you drag the fill handle (the small square at the bottom-right of the selection).")
    m.steps(["Click the first cell in each row below (Jan, ST01, 2025-01-01), in column C.",
             "Drag the fill handle right to column M (11 cells in all).",
             "Excel continues months, codes with numbers and dates by itself.",
             "Home › Fill › Series sets a step value; File › Options › Advanced › Edit Custom Lists adds your own lists."])
    m.turn()
    for start, fmt_, chk, see in [("Jan", None, '="Nov"', "Nov"), ("ST01", None, '="ST11"', "ST11"),
                                  (dt.datetime(2025, 1, 1), "date", "=DATE(2025,1,11)", "2025-01-11")]:
        m.n += 1
        r = m.r
        m.ws.write(r, 1, f"Ex {m.n}", m.X["exid"])
        if isinstance(start, dt.datetime):
            m.ws.write_datetime(r, 2, start, b.nf("date", style="cell", bold=True))
        else:
            m.ws.write(r, 2, start, F["lab"])
        for c in range(3, LAST_COL + 1):
            m.ws.write_blank(r, c, None, b.nf(fmt_, style="plain", border=2, border_color="#E0A526") if fmt_ else m.X["in"])
        last = A(r, LAST_COL)
        m.ws.merge_range(r + 1, 2, r + 1, LAST_COL, "", m.X["chk"])
        msg = f"Drag the fill handle from C{r + 1} to M{r + 1}"
        m.ws.write_formula(r + 1, 2, f'=IF({last}="","{msg}",IF({last}{chk},"✓ Correct: the series reaches {see} in column M","✗ Not yet"))',
                           m.X["chk"], msg)
        m.key_txt.append((m.n, "1.6.6", f"Column M shows {see}."))
        m.r += 3

    # ---------------------------------------------------------------- 1.6.7
    names = ["Lagos - Ikeja", "Port Harcourt - GRA", "Abuja - Wuse", "Jos - Rayfield", "Enugu - Independence Layout", "Owerri - New Owerri"]
    m.topic("1.6.7", "Flash Fill (Ctrl+E)",
            "Flash Fill detects a pattern from one or two examples you type and fills the rest of the column to match. Its results are typed values, not formulas.")
    m.steps(["In the first white cell type the city from the store name: Lagos.",
             "Move to the white cell below.",
             "Press Ctrl+E (or Data › Flash Fill).",
             "Check every row: Flash Fill guesses."])
    r0 = m.r
    m.ws.merge_range(r0, 1, r0, 3, "Store_Name", F["hdr"])
    m.ws.write(r0, 4, "City (type)", F["hdr"])
    m.ws.merge_range(r0, 5, r0, 6, "Check", F["hdr"])
    for i, nm in enumerate(names):
        r = r0 + 1 + i
        m.ws.merge_range(r, 1, r, 3, nm, F["cell"])
        m.ws.write_blank(r, 4, None, m.X["in"])
        m.ws.merge_range(r, 5, r, 6, "", m.X["chk"])
        m.ws.write_formula(r, 5, f'=IF({A(r, 4)}="","",IF({A(r, 4)}=LEFT({A(r, 1)},FIND(" - ",{A(r, 1)})-1),"✓","✗"))', m.X["chk"], "")
    m.r = r0 + 8
    m.turn()
    m.ex_formula_live("Fill all six cities with Flash Fill. The check counts the correct rows.",
                      f'=IF(COUNTA(E{r0 + 2}:E{r0 + 7})=0,"",IF(SUMPRODUCT(--(E{r0 + 2}:E{r0 + 7}=LEFT(B{r0 + 2}:B{r0 + 7},FIND(" - ",B{r0 + 2}:B{r0 + 7})-1)))=6,"✓ All six correct","✗ Check the rows marked ✗"))',
                      "", key="Lagos, Port Harcourt, Abuja, Jos, Enugu, Owerri.")

    # ---------------------------------------------------------------- 1.6.8
    m.topic("1.6.8", "Formatting cells with Ctrl+1",
            "Formatting changes how a value looks (font, borders, fill, number format, alignment) without changing the value stored. Ctrl+1 opens Format Cells.")
    m.steps(["Click the white cell holding 0.127438.",
             "Press Ctrl+1 › Number tab › Percentage › Decimal places 1 › OK.",
             "The cell shows 12.7%; the formula bar still shows 0.127438."])
    m.turn()
    r = m.r
    c = A(r, 8)
    m.ex_formula_live("Format the white cell as Percentage with 1 decimal place.",
                      f'=IF(CELL("format",{c})="P1","✓ Correct: shown as 12.7%","Format it as Percentage, 1 decimal (Ctrl+1)")',
                      "Format it as Percentage, 1 decimal (Ctrl+1)", key="Percentage, 1 decimal place: the cell shows 12.7%.")
    m.ws.write_number(r, 8, 0.127438, m.X["in"])
    m.ex_text("Which Format Cells tab adds lines around cells?", "Border")

    # ---------------------------------------------------------------- 1.6.9
    m.topic("1.6.9", "Formatting dates, currency, percentages and decimals",
            "The same stored number can look completely different. Choose the format that makes the meaning obvious; custom codes are typed in Ctrl+1 › Custom.")
    m.steps(["Each row below holds the same kind of stored value with a different number format.",
             "Click a result cell and read the formula bar: the stored value has not changed.",
             "The TEXT function applies a format code inside a formula: =TEXT(value, \"code\")."])
    r0 = m.r
    for j, h in enumerate(["Stored value", "Format", "Code", "Displays"]):
        m.ws.write(r0, 1 + 2 * j, h, F["hdr"])
        m.ws.write(r0, 2 + 2 * j, "", F["hdr"])
    rows = [(733215255, "Number, comma", "#,##0"), (733215255, "Currency (naira)", "₦#,##0"), (733215255, "Millions", '₦#,##0.0,,"m"'),
            (0.127438, "Percentage, 1 dp", "0.0%"), (45658, "Date, ISO", "yyyy-mm-dd"), (45658, "Date, long", "d mmmm yyyy")]
    for i, (v, nm, code) in enumerate(rows):
        r = r0 + 1 + i
        m.ws.merge_range(r, 1, r, 2, v, b.nf("gen"))
        m.ws.merge_range(r, 3, r, 4, nm, F["cell"])
        m.ws.merge_range(r, 5, r, 6, code, F["cell"])
        m.ws.merge_range(r, 7, r, 8, v, b.nf(code, style="result"))
    m.r = r0 + 8
    m.turn()
    m.ex_text('In the white cell type =TEXT(733215255,"₦#,##0.0,,""m""") and press Enter.', "₦733.2m",
              formula='=TEXT(733215255,"₦#,##0.0,,""m""")', short="TEXT with the millions code")
    m.ex_text('Type =TEXT(45658,"yyyy-mm-dd") in the white cell.', "2025-01-01", formula='=TEXT(45658,"yyyy-mm-dd")', short="TEXT with the ISO date code")

    # ---------------------------------------------------------------- 1.6.10
    m.topic("1.6.10", "Alignment, Wrap Text and the merge-cell trap",
            "Alignment positions text in the cell; Wrap Text splits long text over several lines. Merging looks neat but breaks sorting, filtering and copying: use Center Across Selection.")
    m.steps(["Select the title cell and the four cells to its right (B to F below).",
             "Ctrl+1 › Alignment › Horizontal: Center Across Selection › OK.",
             "The title is centred, but every cell is still separate: sorting and copying still work."])
    m.ws.write(m.r, 1, "FY2025 store results", m.X["in"])
    for c in range(2, 6):
        m.ws.write_blank(m.r, c, None, m.X["in"])
    m.r += 2
    m.do([("Centre the title above across B to F with Center Across Selection.", "The title sits in the middle; clicking C shows the cell is empty and separate."),
          ("Turn on Wrap Text for a long header (Home › Wrap Text).", "The header shows on two lines; the row gets taller.")])

    # ---------------------------------------------------------------- 1.6.11
    m.topic("1.6.11", "Cell Styles, themes and conditional formatting basics",
            "Cell Styles apply saved formatting in one click; themes change fonts and colours for the whole workbook. Conditional formatting changes a cell's format automatically according to its value.")
    m.steps(["Select the Margin values in the practice range below.",
             "Home › Conditional Formatting › Highlight Cells Rules › Less Than.",
             "Type 0, choose Light Red Fill, OK: the loss-making stores light up.",
             "The Attainment column already has data bars (Home › Conditional Formatting › Data Bars)."])
    prac = ST[["Store_Name", "Profit", "Margin", "Attainment"]].copy()
    top = m.s.frame(prac, fmts={"Profit": N, "Margin": "pct", "Attainment": "pct"})
    m.ws.conditional_format(top, 4, top + 11, 4, {"type": "data_bar", "bar_color": "#13807A"})
    mr = f"D{top + 1}:D{top + 12}"
    m.turn()
    m.ex("How many Margin cells turn red?", f'=COUNTIF({mr},"<0")', int((ST.Margin < 0).sum()), "int", short="Red Margin cells")
    m.do([("Apply the Cell Style 'Good' to the best Attainment value.", "A green fill on Lagos - Lekki's 112.4%.")])

    # ---------------------------------------------------------------- 1.6.12
    m.topic("1.6.12", "Freeze Panes and Split",
            "Freeze Panes keeps the header row (and first column) visible while you scroll. Split divides the window into panes that scroll separately.")
    m.steps(["Go to Data_Sales.", "View › Freeze Panes › Freeze Top Row.", "Scroll down: row 1 stays.", "To freeze row 1 and column A: click B2 › Freeze Panes."])
    m.turn()
    m.ex_text("After freezing on Data_Sales, scroll to row 5,000. Which row number stays visible at the top?", "1")
    m.do([("On Data_Sales freeze row 1 and column A together (click B2 first).", "Scrolling right keeps Row_ID visible; scrolling down keeps the headers.")])

    # ---------------------------------------------------------------- 1.6.13
    m.topic("1.6.13", "Printing and page setup",
            "Set the page up before printing or saving as PDF, and always check Print Preview first.")
    m.steps(["On Data_Sales press Ctrl+P: the preview shows the first page.",
             "Page Layout › Print Titles › Rows to repeat at top: $1:$1.",
             "Settings: Landscape, Fit All Columns on One Page.",
             "Print, or choose Microsoft Print to PDF."])
    m.turn()
    m.ex_text("Which Page Layout command repeats the header row on every printed page?", "Print Titles")
    m.do([("Set Print Titles to $1:$1 on Data_Sales and check page 2 of the preview.", "The column headers appear at the top of page 2.")])

    # ---------------------------------------------------------------- 1.6.14
    variants = ["LAGOS - IKEJA", "lagos - ikeja", "Lagos - Ikeja", "LAGOS - IKEJA", "Lagos - Ikeja", "lagos - ikeja", "LAGOS - IKEJA", "Lagos - Ikeja"]
    m.topic("1.6.14", "Find & Replace, Go To Special and data validation",
            "Find & Replace fixes many cells at once. Go To Special selects cells of one kind (blanks, formulas). Data validation stops bad values being typed in the first place.")
    m.steps(["Select the white store-name cells below. Ctrl+H › Find what: LAGOS - IKEJA › Replace with: Lagos - Ikeja › Options › tick Match case › Replace All.",
             "Repeat for lagos - ikeja.",
             "Go To Special: select the Discount cells › F5 › Special › Blanks › OK: only the empty cells are selected.",
             "Data validation: the Zone cell below accepts only the four zones; try typing Lagos."])
    r0 = m.r
    m.ws.write(r0, 1, "Store_Name", F["hdr"])
    m.ws.write(r0, 2, "", F["hdr"])
    m.ws.write(r0, 4, "Discount", F["hdr"])
    disc = [0.2, None, 0, 0.45, None, 0.2, 0, None]
    for i, (v, d_) in enumerate(zip(variants, disc)):
        m.ws.merge_range(r0 + 1 + i, 1, r0 + 1 + i, 2, v, m.X["in"])
        if d_ is None:
            m.ws.write_blank(r0 + 1 + i, 4, None, b.nf("pct0"))
        else:
            m.ws.write_number(r0 + 1 + i, 4, d_, b.nf("pct0"))
    srng, drng = f"B{r0 + 2}:B{r0 + 9}", f"E{r0 + 2}:E{r0 + 9}"
    m.r = r0 + 10
    m.turn()
    m.ex_formula_live("Fix all eight store names with Find & Replace (Match case).",
                      f'=IF(SUMPRODUCT(--EXACT({srng},"Lagos - Ikeja"))=8,"✓ All eight fixed",SUMPRODUCT(--EXACT({srng},"Lagos - Ikeja"))&" of 8 correct so far")',
                      f"{variants.count('Lagos - Ikeja')} of 8 correct so far", key="All eight read Lagos - Ikeja.")
    m.ex("How many blank Discount cells does Go To Special › Blanks select?", f"=COUNTBLANK({drng})", disc.count(None), "int", short="Blank Discount cells")
    r = m.r
    c = A(r, 8)
    m.ex_formula_live("Pick a zone in the white cell from its drop-down. Then try typing Lagos: validation refuses it.",
                      f'=IF({c}="","",IF(OR({c}="North-Central",{c}="South-East",{c}="South-South",{c}="South-West"),"✓ Valid zone","✗ Not a zone"))', "",
                      key="Any of North-Central, South-East, South-South, South-West.")
    m.ws.data_validation(r, 8, r, 8, {"validate": "list", "source": ["North-Central", "South-East", "South-South", "South-West"],
                                      "error_title": "Not a zone", "error_message": "Choose one of the four zones from the list."})
    m.finish()


def build(b, D):
    m14(b, D)
    m15(b, D)
    m16(b, D)
