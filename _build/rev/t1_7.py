from revlib import T

SALES = "SuperMart_Stores_Sales_2022_2025.csv"

TOPICS = [
    T("1.7.1",
      what="A **formula** is an instruction you write in a cell to make Excel calculate. It always starts with an **equal "
           "sign (=)**. A formula usually points to other cells, such as =R2-U2, instead of typing the numbers, so when "
           "those cells change the answer changes too.",
      like="A formula is like a recipe card that says 'two cups of whatever rice is in the bag'. If the rice changes, "
           "the instruction still works. You never rewrite the recipe.",
      how=["Click the cell where the answer should go.",
           "Type = then click the cells and type the operators (+ − * /).",
           "Press Enter, then check the Formula Bar."],
      tip="Point to cells in formulas; avoid typing numbers inside them.",
      trap="Forgetting the = sign; Excel then stores your formula as plain text.",
      who="SuperMart Retail",
      case="The SuperMart file shows Sales and Profit for each line but not the cost. For four 2025 lines the analyst "
           "wants cost (Sales − Profit) and margin (Profit ÷ Sales).",
      steps=["In V6685 type =R6685-U6685 (cost).",
             "In W6685 type =U6685/R6685 and format it as a percentage (margin).",
             "Copy both formulas to the other rows."],
      table=[["Row_ID", "Sales", "Profit", "Cost", "Margin"],
             ["6683", "₦48,896", "₦8,557", "₦40,339", "17.5%"],
             ["6684", "₦474,430", "₦199,261", "₦275,169", "42.0%"],
             ["6686", "₦454,560", "−₦107,958", "₦562,518", "−23.8%"],
             ["6688", "₦310,744", "−₦26,635", "₦337,379", "−8.6%"]],
      split=4.40,
      so="Two short formulas reveal what the file didn't show: two of these four lines sold below cost."),

    T("1.7.2",
      what="A **function** is a ready-made formula with a name, built into Excel. You write its name and give it "
           "**arguments** (the inputs) in brackets: =SUM(R2:R3313). Functions do jobs that would be long or impossible "
           "with + − * / alone, such as adding thousands of cells, finding the middle value, or making decisions with "
           "IF.",
      like="A function is like a blender. You put in the ingredients (arguments), press one button, and get the result. "
           "You don't need to know how the motor works.",
      how=["Type = then the function name, e.g. =SUM(",
           "Select or type the arguments, separated by commas.",
           "Close the bracket and press Enter."],
      tip="Read the yellow tooltip that appears as you type; it lists the arguments.",
      trap="Typing =R2+R3+R4+… for thousands of cells instead of =SUM(R2:R3313).",
      who="SuperMart Retail",
      case="The analyst has the 2025 sales lines in a sheet, with Sales in R2:R3313. She needs the total, the average, "
           "the middle value and the number of lines.",
      steps=["Type =SUM(R2:R3313).",
             "Type =AVERAGE(R2:R3313), =MEDIAN(R2:R3313) and =COUNT(R2:R3313).",
             "Compare the average with the median."],
      table=[["Function", "Returns (2025 lines)"],
             ["=SUM(R2:R3313)", "₦733,215,255"],
             ["=AVERAGE(R2:R3313)", "₦221,381"],
             ["=MEDIAN(R2:R3313)", "₦53,810"],
             ["=COUNT(R2:R3313)", "3,312"]],
      split=5.20,
      so="The average line is ₦221k, but half of all lines are below ₦54k. A few big sales pull the average up."),

    T("1.7.3",
      what="Every function is used inside a formula, but not every formula uses a function. **=R2-U2** is a formula "
           "with only operators. **=SUM(R2:R10)** is a formula that uses a function. Most real work combines both, for "
           "example =SUM(U2:U10)/SUM(R2:R10).",
      like="A formula is a full meal; a function is a ready-made ingredient, like stock cubes. Some meals use them, some "
           "don't, and most good meals mix ready-made and fresh ingredients.",
      how=["Use operators for simple sums between a few cells.",
           "Use functions for ranges, decisions and lookups.",
           "Combine them when you need to, e.g. =SUM(…)/SUM(…)."],
      tip="Prefer a function when a range has more than a few cells.",
      trap="Dividing two averages when you need the ratio of the two totals.",
      who="SuperMart Retail",
      case="The analyst needs three results for 2025: one line's cost, the total profit, and the overall margin. Which "
           "need functions?",
      steps=["Cost of one line: =R2-U2 (operators only).",
             "Total profit: =SUM(U2:U3313) (one function).",
             "Overall margin: =SUM(U2:U3313)/SUM(R2:R3313) (both together)."],
      table=[["Need", "Formula", "Type"],
             ["Cost of one line", "=R2-U2", "Operators only"],
             ["Total profit", "=SUM(U2:U3313)", "Function"],
             ["Overall margin", "=SUM(U2:U3313)/SUM(R2:R3313)", "Both: 12.7%"]],
      split=4.30,
      so="Use whichever fits: operators for simple sums, functions for ranges, and both together for ratios."),

    T("1.7.4",
      what="A formula always starts with **=**. After it come **cell references** (R2), **operators** (+ − * /), "
           "**functions** (SUM), **numbers** (1.075) and sometimes **text** in double quotes (\"Loss\"). Excel works "
           "it out and shows the result. The Formula Bar keeps showing the formula itself.",
      like="A formula is like a sentence with grammar rules. The = is the capital letter at the start. References and "
           "operators are the words. Break the rules and Excel cannot understand you.",
      how=["Start with =.",
           "Put text in double quotes: \"Loss\".",
           "Check that every opening bracket has a closing one."],
      tip="Put fixed values (like a tax rate) in their own cell and refer to it.",
      trap="Typing a space or an apostrophe before =; the formula shows as text.",
      who="SuperMart Retail",
      case="The analyst writes one formula that labels each 2025 line as 'Loss' or 'Profit' and wants to understand "
           "every part of it.",
      steps=["In V2 type =IF(U2<0,\"Loss\",\"Profit\").",
             "Read each part using the table.",
             "Copy the formula down and count the 'Loss' rows with the filter."],
      table=[["Part", "In =IF(U2<0,\"Loss\",\"Profit\")"],
             ["Equal sign", "="],
             ["Function", "IF"],
             ["Reference and operator", "U2<0"],
             ["Text results", "\"Loss\" and \"Profit\""]],
      split=5.20,
      so="Each part has a job. Copied down, the formula flags all 620 loss-making lines of 2025."),

    T("1.7.5",
      what="When a formula has several operators, Excel follows **BODMAS**: **B**rackets first, then **O**rders "
           "(powers, ^), then **D**ivision and **M**ultiplication, then **A**ddition and **S**ubtraction. Use brackets "
           "to make Excel calculate in the order you mean.",
      like="It is like getting dressed: shirt before tie, socks before shoes. Change the order and the result looks "
           "very wrong, even with the same clothes.",
      how=["Write the calculation in words first.",
           "Put brackets around the part that must happen first.",
           "Test it with simple numbers you can check in your head."],
      tip="Add brackets whenever there is any doubt; they never hurt.",
      trap="Writing =Price*1-Discount, which subtracts the discount rate itself instead of applying it.",
      who="SuperMart Retail",
      case="A list price is ₦100,000 and the discount is 20% (0.2). A new analyst writes =100000*1-0.2 and gets a strange "
           "answer. Why?",
      steps=["Type =100000*1-0.2 and note the result.",
             "Type =100000*(1-0.2) and note the result.",
             "Explain the difference using BODMAS."],
      table=[["Formula", "Order Excel uses", "Result"],
             ["=100000*1-0.2", "100000×1, then −0.2", "₦99,999.80"],
             ["=100000*(1-0.2)", "(1−0.2) first, then ×", "₦80,000"],
             ["=2+3*4", "3×4, then +2", "14"],
             ["=(2+3)*4", "2+3, then ×4", "20"]],
      split=4.70,
      so="Brackets turned a meaningless ₦99,999.80 into the correct ₦80,000. The order of operations matters."),

    T("1.7.6",
      what="A **cell reference** is a cell's address used in a formula: R2, a range like R2:R3313, or a cell on another "
           "sheet like Rates!B1. A **relative** reference (R2) moves when you copy the formula. An **absolute** "
           "reference ($B$1) stays fixed. That is how one tax-rate cell can serve thousands of rows.",
      like="Relative is like saying 'the house next to mine': it changes depending on where you stand. Absolute is a "
           "full address, '12 Allen Avenue, Ikeja', which is the same from anywhere.",
      how=["Put the fixed value (e.g. VAT 7.5%) in one cell, e.g. B1.",
           "In the formula, lock it: =A4*$B$1.",
           "Copy the formula down; A4 moves but $B$1 does not."],
      tip="Keep every fixed rate in one labelled cell and lock it with $.",
      trap="Forgetting the $; copied formulas then point at empty cells and give 0.",
      who="SuperMart Retail",
      case="SuperMart must show 7.5% VAT on some invoices. The rate sits in cell B1. The analyst calculates VAT for "
           "several sales without retyping 7.5%.",
      steps=["Type 7.5% in B1 and the sales amounts in A4:A6.",
             "In B4 type =A4*$B$1 and in C4 type =A4+B4.",
             "Copy B4:C4 down to row 6."],
      table=[["Sale (A)", "VAT =A4*$B$1", "Total =A4+B4"],
             ["₦474,430", "₦35,582.25", "₦510,012.25"],
             ["₦100,000", "₦7,500.00", "₦107,500.00"],
             ["₦48,896", "₦3,667.20", "₦52,563.20"]],
      split=4.90,
      so="If the VAT rate changes, you edit B1 once and every row updates. That is the power of references."),

    T("1.7.7",
      what="Press **F4** while your cursor is on a reference to cycle its locks: A1 → **$A$1** (both locked) → **A$1** "
           "(row locked) → **$A1** (column locked). The $ sign locks the part after it. Mixed references let one formula "
           "fill a whole grid, down and across.",
      like="Think of seats in a hall. $A$1 is a reserved seat: always the same one. A$1 means 'always the front row, any "
           "column'. $A1 means 'always the aisle seat, any row'.",
      how=["Click into the reference in the formula bar.",
           "Press F4 until the $ signs are where you need them.",
           "Fill the formula down and across, then check one corner cell."],
      tip="Test a corner cell of the grid after filling; it shows whether locks are right.",
      trap="Locking both parts ($A$1) when you need one to move.",
      who="SuperMart Retail",
      case="Management wants 2026 targets for three stores under growth rates of 5%, 10% and 15%. The 2025 targets are in "
           "B2:B4 and the growth rates in C1:E1.",
      steps=["In C2 type =$B2*(1+C$1): the column of B and the row of 1 are locked.",
             "Fill C2 right to E2, then fill C2:E2 down to row 4.",
             "Check E4: it should use B4 and E1."],
      table=[["Store (2025 target)", "+5%", "+10%", "+15%"],
             ["Lagos - Ikeja (₦142m)", "₦149.1m", "₦156.2m", "₦163.3m"],
             ["Port Harcourt - GRA (₦97m)", "₦101.8m", "₦106.7m", "₦111.5m"],
             ["Lagos - Lekki (₦63m)", "₦66.2m", "₦69.3m", "₦72.5m"]],
      split=4.70,
      so="One formula with mixed references filled all nine cells correctly. That is far faster and safer than typing "
         "nine formulas."),

    T("1.7.8",
      what="Most business sums use five basic **operators**: plus (+), minus (−), times (*), divide (/) and power "
           "(^). With these you can work out profit, margin, growth, share of total and average price. Learn these "
           "few patterns and you can answer most everyday questions.",
      like="These patterns are like the times tables you learnt in primary school. Once you know them by heart, you use "
           "them everywhere without thinking.",
      how=["Growth % = (new − old) / old.",
           "Share % = part / total.",
           "Margin % = profit / sales; average price = sales / quantity."],
      tip="Format results as % or ₦ so their meaning is obvious.",
      trap="Calculating growth as new/old and forgetting to subtract 1.",
      who="SuperMart Retail",
      case="The analyst prepares four headline numbers for 2025 using only basic operators and the yearly totals.",
      steps=["Growth: =(733215255-609205599)/609205599.",
             "Ikeja share: =146388345/733215255.",
             "Margin: =93439242/733215255; average order: =733215255/1687."],
      table=[["Calculation", "Pattern", "Result"],
             ["Sales growth 2024→25", "(new−old)/old", "+20.4%"],
             ["Ikeja share of sales", "part/total", "20.0%"],
             ["Profit margin", "profit/sales", "12.7%"],
             ["Average order value", "sales/orders", "₦434,627"]],
      split=4.90,
      so="Four patterns produce the four numbers a CEO asks for first. In a real sheet, use cell references, not "
         "typed numbers."),

    T("1.7.9",
      what="A small set of functions does most daily work. **SUM** adds, **AVERAGE** finds the mean, **MEDIAN** finds the "
           "middle value, **COUNT** counts cells with numbers, **COUNTA** counts cells that are not empty, **MAX** and "
           "**MIN** find the largest and smallest, and **ROUND** rounds a value.",
      like="These are like the basic tools in a mechanic's box: spanner, screwdriver, pliers. Most jobs need nothing "
           "more, so learn to use them well.",
      how=["Type =FUNCTION( and select the range.",
           "Close the bracket and press Enter.",
           "Compare COUNT with COUNTA to spot text in a number column."],
      tip="Use COUNTA on an ID column to count records reliably.",
      trap="Using COUNT on text columns; it returns 0 because it counts only numbers.",
      who="SuperMart Retail",
      case="The analyst summarises the 3,312 sales lines of 2025 (Sales in R2:R3313, Order_ID in B2:B3313) with the "
           "core functions.",
      steps=["Type =MAX(R2:R3313) and =MIN(R2:R3313).",
             "Type =COUNT(B2:B3313) and =COUNTA(B2:B3313).",
             "Explain why COUNT gives 0 for the Order_ID column."],
      table=[["Formula", "Result", "Why"],
             ["=MAX(R2:R3313)", "₦13,999,960", "Copier sale, Lekki"],
             ["=MIN(R2:R3313)", "₦444", "Smallest line"],
             ["=COUNT(B2:B3313)", "0", "Order IDs are text"],
             ["=COUNTA(B2:B3313)", "3,312", "Counts non-empty cells"]],
      split=4.70,
      so="Each function answers one simple question. Pick the right one: COUNT is for numbers, COUNTA for anything."),

    T("1.7.10",
      what="The **fx** button next to the Formula Bar opens **Insert Function**. You can search by describing what you "
           "want (for example 'middle value') and Excel suggests functions. Then a window shows each argument with a "
           "short explanation and a live preview of the result.",
      like="Insert Function is like asking a shop attendant: 'I need something to clean my shoes.' You describe the job "
           "and they bring you the right product.",
      how=["Click the cell, then click fx next to the Formula Bar.",
           "Type a description and click Go.",
           "Pick a function, fill in each argument box, check the preview, click OK."],
      tip="Use fx whenever you are unsure of a function's arguments.",
      trap="Guessing argument order; the dialog shows it clearly.",
      who="ETEH Talent Salary Survey",
      case="A learner wants the typical salary in ETEH_Talent_Salary_Survey.csv (salaries in G2:G1782) but cannot "
           "remember which function gives the middle value.",
      steps=["Click an empty cell and then fx.",
             "Search 'middle value' and choose MEDIAN.",
             "In Number1 type G2:G1782; the preview shows the answer. Click OK."],
      table=[["In the dialog", "What you do or see"],
             ["Search box", "middle value"],
             ["Suggested", "MEDIAN"],
             ["Number1", "G2:G1782"],
             ["Result", "₦11,000,000"]],
      split=5.20,
      so="Without remembering the name, the learner found MEDIAN and the answer: the typical salary is ₦11.0m."),

    T("1.7.11",
      what="**AutoSum** (**Alt+=**) inserts a SUM of the cells above or to the left in one step. Its drop-down arrow "
           "(Home › Editing) offers the other common functions: **Average**, **Count Numbers**, **Max** and **Min**. "
           "Always check the range Excel guesses before pressing Enter.",
      like="AutoSum is like the 'total' button on a POS machine: one press and it adds everything above. You still "
           "glance at the receipt to check it.",
      how=["Click the cell below a column of numbers.",
           "Press Alt+= and check the highlighted range.",
           "Press Enter, or use the drop-down for Average, Max or Min."],
      tip="Check the dotted border Excel draws before you press Enter.",
      trap="A blank cell in the column makes AutoSum stop early and miss rows.",
      who="SuperMart Retail",
      case="The analyst has a 12-row table of 2025 sales by store (B2:B13) and needs the total, average, highest and "
           "lowest store in seconds.",
      steps=["Click B14 and press Alt+= then Enter.",
             "In B15–B17 use the AutoSum drop-down: Average, Max, Min.",
             "Check that each range is B2:B13."],
      table=[["AutoSum choice", "Formula", "Result"],
             ["Sum", "=SUM(B2:B13)", "₦733.2m"],
             ["Average", "=AVERAGE(B2:B13)", "₦61.1m per store"],
             ["Max", "=MAX(B2:B13)", "₦146.4m (Ikeja)"],
             ["Min", "=MIN(B2:B13)", "₦31.6m (Owerri)"]],
      split=4.90,
      so="Four results in under a minute. The biggest store sells more than four times as much as the smallest."),

    T("1.7.12",
      what="Error values are Excel's way of telling you what went wrong. **#DIV/0!** means dividing by zero or an empty "
           "cell. **#VALUE!** means text was used where a number was needed. **#REF!** means a cell the formula used "
           "was deleted. **#NAME?** means a misspelt function. **#N/A** means a lookup found no match.",
      like="Errors are like warning lights on a car dashboard. Each light means something specific. Read it and fix that "
           "problem, don't just cover the light with tape.",
      how=["Click the error cell and read the formula.",
           "Match the error to its meaning in the table.",
           "Fix the cause; use IFERROR only when the error is expected."],
      tip="Find the cause of each error before hiding it.",
      trap="Wrapping everything in IFERROR(…,0); real problems disappear from view.",
      who="SuperMart Retail",
      case="An analyst's sheet built on " + "SuperMart_Sales_2025_DIRTY.csv" + " is full of errors. Each one has a "
           "different cause.",
      steps=["Click each error cell and read its formula.",
             "Identify the cause using the table.",
             "Apply the fix and check the result."],
      table=[["Error", "Cause in this sheet", "Fix"],
             ["#VALUE!", "='₦239,240'*2 (text sale)", "Convert to a number"],
             ["#DIV/0!", "Margin on a blank row", "Remove blank rows"],
             ["#NAME?", "=SUMM(R2:R10)", "Correct to SUM"],
             ["#REF!", "Column M deleted", "Rebuild the reference"]],
      split=4.90,
      so="Each error points to a specific cause. Fixing the cause is better than hiding the message."),

    T("1.7.13",
      what="Write a formula once, then **copy** it to every row. You can drag the fill handle, **double-click** it (fills "
           "down as far as the data next to it), or use **Ctrl+D** (fill down) and **Ctrl+R** (fill right). Relative "
           "references adjust automatically on each row.",
      like="It is like printing a flyer: you design one, then the printer copies it hundreds of times. You don't redesign "
           "each copy.",
      how=["Write and check the formula in the first row.",
           "Double-click the fill handle to fill to the last row.",
           "Press Ctrl+Down to check the last row has the formula."],
      tip="Double-click the fill handle to fill thousands of rows instantly.",
      trap="A blank cell in the next column stops the double-click fill early.",
      who="SuperMart Retail",
      case="The analyst wrote =U2/R2 (margin) in V2 of the 2025 sheet. She needs it on all 3,312 lines without "
           "dragging for minutes.",
      steps=["Click V2 and check the result.",
             "Double-click the small square at the bottom-right of V2.",
             "Press Ctrl+Down: the last cell, V3313, should show =U3313/R3313."],
      table=[["Method", "Action", "Best for"],
             ["Double-click handle", "Double-click the square", "Long columns"],
             ["Drag handle", "Drag down or right", "A few cells"],
             ["Ctrl+D", "Select range, press Ctrl+D", "Selected range"],
             ["Ctrl+R", "Select range, press Ctrl+R", "Filling across"]],
      split=4.90,
      so="One double-click filled 3,312 rows correctly, because each copy points to its own row."),

    T("1.7.14",
      what="Real work combines several operators and functions in one sheet to support a decision. A shop needs margin "
           "and a loss flag. A university needs a grade. A hospital needs occupancy. A hotel needs a cancellation "
           "rate. The building blocks are the same in every sector.",
      like="Like a tailor who uses the same few tools (tape, chalk, scissors) to make a shirt, an agbada or a gown. "
           "Different results, same skills.",
      how=["Write the business rule in words.",
           "Turn it into a formula with references.",
           "Test it on one row you can check by hand."],
      tip="Test every new formula on one row you can work out by hand.",
      trap="Copying a formula from another sector without checking that its rule fits.",
      who="Multi-Sector",
      case="Four bootcamp organisations each need one key calculation. The analyst writes each formula and tests it on "
           "real data.",
      steps=["Write each rule in words.",
             "Type the formula shown in the table.",
             "Check the result against the value given."],
      table=[["Sector", "Formula", "Result"],
             ["Retail", "=U2/R2 (margin, Row_ID 6684)", "42.0%"],
             ["University", "=IF(Exam>=70,\"A\",IF(Exam>=60,\"B\",\"C\"))", "Exam 61 → B"],
             ["Hospital", "=Patients/520 (beds)", "2023 average 89.5%"],
             ["Hotel", "=Cancelled/Bookings", "11,885/36,275 = 32.8%"]],
      split=4.20,
      so="The same skills (references, operators, IF) solve problems in every sector. Only the business rule "
         "changes."),
]
