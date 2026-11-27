from revlib import T

DIRTY = "SuperMart_Sales_2025_DIRTY.csv"

TOPICS = [
    T("1.6.1",
      what="When you press **Enter**, Excel decides what kind of value you typed. **Text** lines up on the left and is "
           "never calculated. **Numbers** line up on the right and can be added. **Dates** also line up right because "
           "Excel stores each date as a day number. If Excel guesses wrongly, your sums and sorts go wrong.",
      like="It is like a bank teller sorting what you hand over: cash goes in the till, cheques in another tray. Put a "
           "cheque in the cash till and the count at closing time is wrong.",
      how=["Type the value and press Enter.",
           "Check the alignment: left = text, right = number or date.",
           "For codes with leading zeros, format the column as Text first."],
      tip="Type dates as YYYY-MM-DD (2025-03-18); Excel always reads them correctly.",
      trap="Typing phone numbers like 08031234567 into a number cell; Excel drops the leading 0.",
      who="SuperMart Retail",
      case="A sales clerk types new records by hand into a copy of the SuperMart file. Some entries look fine, but they "
           "do not behave correctly in formulas.",
      steps=["Type each value from the table below into an empty cell.",
             "Note the alignment and what the Formula Bar shows.",
             "Apply the fix shown and type it again."],
      table=[["You type", "Excel stores", "Fix"],
             ["08031234567", "8031234567 (0 lost)", "Format as Text first"],
             ["04/05/2025", "4 May or 5 April?", "Type 2025-05-04"],
             ["₦16,448", "Number with ₦ format", "OK on most PCs"],
             ["16448 (with a space)", "Text: not summed", "Delete the space"]],
      split=4.90,
      so="Look at the alignment after you type: it shows at once whether Excel understood the value."),

    T("1.6.2",
      what="Every cell holds one of six **data types**: **Number** (including money and %), **Text**, **Date/Time**, "
           "**Logical** (TRUE or FALSE), **Error** (such as #DIV/0! or #N/A) and **Array** (a list of results that "
           "spills into nearby cells). The type decides what formulas can do. SUM adds numbers but silently skips "
           "text.",
      like="It is like a bus with seats for adults and children. If a child's ticket is recorded as 'luggage', the "
           "head count is wrong, even though the child is on the bus.",
      how=["Select a column and check the Status Bar's Count and Sum.",
           "Test a cell with =ISNUMBER(A2) or =ISTEXT(A2).",
           "Convert text numbers before summing (Week 2 shows how)."],
      tip="Use =ISNUMBER() on a column before you trust its total.",
      trap="Assuming a value that looks like a number is a number.",
      who="SuperMart Retail",
      case="In " + DIRTY + " some Sales values were exported as text, like '₦239,240'. They look like numbers, but "
           "SUM ignores them. The analyst checks the types.",
      steps=["In a new column type =ISNUMBER(R2) and fill it down.",
             "Filter that column to FALSE.",
             "Count the FALSE rows; those sales are missing from SUM."],
      table=[["Type", "Example in the pack", "Detect with"],
             ["Number", "16448", "ISNUMBER = TRUE"],
             ["Text", "'₦239,240' (Sales, dirty file)", "ISTEXT = TRUE"],
             ["Date", "2025-03-18", "Right-aligned, a number underneath"],
             ["Error", "#DIV/0! when Quantity is 0", "ISERROR = TRUE"]],
      split=4.70,
      so="234 sales in the dirty file are text, worth ₦57.1m that SUM would skip. Checking types catches this."),

    T("1.6.3",
      what="To **edit** a cell, press **F2** (or double-click) and change part of it. **Delete** removes the content "
           "but keeps the formatting. **Home › Clear** gives more choice: Clear All, Clear Formats, Clear Contents or "
           "Clear Comments. **Ctrl+Z** undoes your last actions if you make a mistake.",
      like="Editing is like correcting one word in your notebook with a pen. Clearing is like using a correction "
           "fluid: you choose whether to remove the writing, the highlight, or both.",
      how=["F2 to edit inside a cell; Enter to confirm, Esc to cancel.",
           "Delete clears content; Home › Clear › Clear All clears everything.",
           "Ctrl+Z to undo, Ctrl+Y to redo."],
      tip="Use F2 to fix one character rather than retyping the whole cell.",
      trap="Pressing Delete and expecting old formatting (e.g. red fill) to go too.",
      who="SuperMart Retail",
      case="In " + DIRTY + " a customer name is stored as 'Omolara Yusuf' followed by a hidden line break. Another "
           "cell has old yellow highlighting with no value. The analyst fixes both.",
      steps=["Click the name cell, press F2, then End, and Backspace to remove the hidden break.",
             "Select the highlighted empty cell: Home › Clear › Clear Formats.",
             "If you clear the wrong cell, press Ctrl+Z immediately."],
      table=[["Action", "Removes value", "Removes format"],
             ["Delete key", "Yes", "No"],
             ["Clear Contents", "Yes", "No"],
             ["Clear Formats", "No", "Yes"],
             ["Clear All", "Yes", "Yes"]],
      so="Choosing the right kind of clearing removes exactly what you want, and nothing else."),

    T("1.6.4",
      what="Right-click a row number or column letter to **insert** or **delete** a whole row or column. Formulas adjust "
           "automatically. **Hide** keeps data but takes it out of view, which is safer than deleting. Drag or "
           "double-click a header's border to **resize** it.",
      like="Inserting a row is like a church usher adding an extra chair in a row: everyone shifts along one seat, and "
           "nobody loses their place.",
      how=["Right-click a header › Insert, Delete or Hide.",
           "Ctrl+Plus inserts and Ctrl+Minus deletes the selected rows or columns.",
           "Double-click a header border to auto-fit the width."],
      tip="Hide helper columns instead of deleting them; you may need them later.",
      trap="Deleting a column that formulas elsewhere depend on; they turn into #REF!.",
      who="SuperMart Retail",
      case="The dirty SuperMart file has an empty column called Column1 (column M) left by an export. The analyst also "
           "wants a new Margin column next to Profit.",
      steps=["Right-click column header M (Column1) › Delete.",
             "Click the column after Profit, right-click › Insert, and type the header Margin.",
             "Double-click the border of Product_Name to fit the long names."],
      table=[["Operation", "Shortcut", "Effect on formulas"],
             ["Insert column", "Ctrl+Plus", "Adjust automatically"],
             ["Delete column", "Ctrl+Minus", "References to it become #REF!"],
             ["Hide column", "Ctrl+0", "Unchanged; still calculate"]],
      split=4.90,
      so="The empty export column is gone, and the new Margin column sits where it is needed. No data was lost."),

    T("1.6.5",
      what="**Ctrl+C** copies, **Ctrl+X** cuts and **Ctrl+V** pastes everything: values, formulas and formatting. "
           "**Paste Special** (**Ctrl+Alt+V**) lets you choose: **Values** only (removes formulas), **Formats** only, "
           "**Transpose** (turns rows into columns), or a calculation such as **Divide** applied to every cell.",
      like="Normal paste is like photocopying a whole page. Paste Special is like copying only the numbers onto a clean "
           "form: you choose exactly what to carry over.",
      how=["Copy the cells with Ctrl+C.",
           "Select the destination and press Ctrl+Alt+V.",
           "Choose Values, Formats, Transpose or an operation, then OK."],
      tip="Paste Values before sending a file, so totals can't change by accident.",
      trap="Pasting formulas into another workbook; links back to the old file appear.",
      who="SuperMart Retail",
      case="SuperMart's store targets are in full naira (₦142,000,000). For a board slide the analyst wants them in "
           "millions (142) without retyping them.",
      steps=["Type 1000000 in an empty cell and copy it (Ctrl+C).",
             "Select the 12 targets, press Ctrl+Alt+V, choose Divide, and click OK.",
             "Delete the helper cell."],
      table=[["Store", "Before", "After Paste Special › Divide"],
             ["Lagos - Ikeja", "142,000,000", "142"],
             ["Port Harcourt - GRA", "97,000,000", "97"],
             ["Lagos - Lekki", "63,000,000", "63"],
             ["Owerri - New Owerri", "39,000,000", "39"]],
      split=4.90,
      so="Twelve values were converted in one move, with no retyping and no typing errors."),

    T("1.6.6",
      what="The **fill handle** is the small square at the bottom-right of a selected cell. Drag it to copy a value or "
           "continue a pattern: Jan gives Feb, Mar; 2022, 2023 gives 2024, 2025; ST01 gives ST02. Double-click the "
           "handle to fill down as far as the data next to it goes.",
      like="It is like a pattern on Ankara fabric: once the first motifs are printed, the machine repeats them all the "
           "way down the cloth.",
      how=["Type the first value (or two, to show the step).",
           "Drag the fill handle across or down.",
           "Use the Auto Fill Options button to switch between copy and series."],
      tip="Type two values to show Excel the step you want (e.g. 2022, 2023).",
      trap="Dragging a single number copies it instead of counting up; hold Ctrl, or type two values.",
      who="SuperMart Retail",
      case="The analyst builds an empty monthly summary for 2025 and a list of store IDs, without typing every label.",
      steps=["Type Jan in A2 and drag the fill handle down to A13.",
             "Type ST01 in C2 and drag down to C13.",
             "Type 2022 and 2023 in two cells, select both and drag to 2025."],
      table=[["You type then drag", "Result"],
             ["Jan", "Feb, Mar … Dec"],
             ["ST01", "ST02, ST03 … ST12"],
             ["2022, 2023", "2024, 2025"],
             ["Q1", "Q2, Q3, Q4"]],
      so="Twelve month labels and twelve store IDs took two drags. AutoFill removes boring, error-prone typing."),

    T("1.6.7",
      what="**Flash Fill** (**Ctrl+E**) copies a pattern you show it. Type the result you want for the first one or two "
           "rows, press Ctrl+E, and Excel fills the rest of the column. It is good for splitting names, pulling out part "
           "of a code, or joining text, all without formulas.",
      like="It is like showing a new apprentice how to cut one plantain into slices. After watching once, he cuts the "
           "whole bunch the same way.",
      how=["Insert an empty column next to the source column.",
           "Type the result for the first row exactly as you want it.",
           "Press Ctrl+E, then check a few rows."],
      tip="Check the results; Flash Fill guesses from your example.",
      trap="Flash Fill does not update when the source changes; use formulas for live data.",
      who="SuperMart Retail",
      case="The marketing team wants SuperMart customers' first names for SMS greetings, and the city of each store for "
           "a map.",
      steps=["Next to Customer_Name, type the first person's first name and press Ctrl+E.",
             "Next to Store_Name, type 'Lagos' for 'Lagos - Ikeja' and press Ctrl+E.",
             "Scroll through a few rows to check the results."],
      table=[["Source", "You type once", "Flash Fill gives"],
             ["Adenike Ogunlesi", "Adenike", "first names for all rows"],
             ["Lagos - Ikeja", "Lagos", "Abuja, Ibadan, Jos …"],
             ["SM-2025-144463", "2025", "the year of every order"]],
      split=4.90,
      so="One example and Ctrl+E did a job that would take an hour by hand, with no formula to learn."),

    T("1.6.8",
      what="**Ctrl+1** opens **Format Cells**, where you control how a value looks: number style, alignment, font, "
           "border and fill. The key rule: formatting changes only the **appearance**, never the stored value. A cell "
           "showing 13% can still hold 0.127438, and every calculation uses the stored value.",
      like="Formatting is like a frame on a photo. A gold frame or a plastic frame changes how it looks on the wall, but "
           "the photo inside is the same.",
      how=["Select the cells and press Ctrl+1.",
           "On the Number tab choose Number, Currency, Percentage or Date.",
           "Check the Formula Bar: the true value has not changed."],
      tip="Use ROUND() in the formula when you really need a rounded value.",
      trap="Thinking you rounded a number because it now shows fewer decimals.",
      who="SuperMart Retail",
      case="The analyst calculates SuperMart's 2025 margin as =93439242/733215255 and formats it as a percentage with "
           "no decimals. The board pack says 13%; finance says 12.7%.",
      steps=["Type the formula; the cell shows 0.127438.",
             "Press Ctrl+1 › Percentage with 0 decimals: it shows 13%.",
             "Set 1 decimal place: it shows 12.7%. The stored value never changed."],
      table=[["Format", "Cell shows", "Cell stores"],
             ["General", "0.127438", "0.1274376…"],
             ["Percentage, 0 dp", "13%", "0.1274376…"],
             ["Percentage, 1 dp", "12.7%", "0.1274376…"]],
      split=4.90,
      so="Both reports used the same value; only the display differed. Agree decimal places before publishing."),

    T("1.6.9",
      what="Good number formats make reports easy to read. Use thousands separators for counts (1,687), the ₦ sign for "
           "money, a % sign for rates, and one clear date style. You can type **custom format codes** in Ctrl+1 › "
           "Custom, for example to show large amounts in millions.",
      like="It is like writing an amount on a cheque: '₦733,215,255' is quick to read and check. '733215255' makes "
           "everyone count the digits.",
      how=["Ctrl+1 › Number tab › Custom.",
           "Type the code in the Type box (see the table).",
           "Apply the same format to the whole column."],
      tip="Use one format per kind of value across the whole report.",
      trap="Mixing ₦m and full naira in the same table.",
      who="SuperMart Retail",
      case="The analyst formats the key 2025 figures for a board report so each one is clear at a glance.",
      steps=["Select each value and press Ctrl+1 › Custom.",
             "Type the code from the table and click OK.",
             "Check the Formula Bar: the stored values are unchanged."],
      table=[["Value stored", "Format code", "Shows"],
             ["1687", "#,##0", "1,687"],
             ["733215255", "\"₦\"#,##0", "₦733,215,255"],
             ["733215255", "\"₦\"#,##0.0,,\"m\"", "₦733.2m"],
             ["0.127438", "0.0%", "12.7%"]],
      split=4.90,
      so="The same numbers become instantly readable. Each comma in the code after 0 divides the display by 1,000."),

    T("1.6.10",
      what="**Alignment** controls where text sits in a cell. **Wrap Text** shows long text on several lines, and "
           "**Alt+Enter** starts a new line inside a cell. **Merge & Center** joins cells into one, but it breaks "
           "sorting, filtering, copying and PivotTables. Use **Center Across Selection** instead: it looks the same but "
           "keeps every cell separate.",
      like="Merging cells is like knocking down the walls between three shops to make one hall. It looks nice, but now "
           "you can't rent, number or lock the shops separately.",
      how=["Select the title cells, e.g. A1:F1.",
           "Ctrl+1 › Alignment › Horizontal › Center Across Selection.",
           "Use Wrap Text for long headers instead of merging."],
      tip="Use Center Across Selection for titles; never merge inside data.",
      trap="Merged cells inside a data table stop Sort and Filter from working.",
      who="SuperMart Retail",
      case="A store manager merges cells A1:F1 for the title 'Store Sales 2025' and merges some store names across two "
           "rows. Sorting now fails with an error.",
      steps=["Select the merged cells and click Merge & Center again to unmerge.",
             "Select A1:F1 › Ctrl+1 › Center Across Selection.",
             "Fill the store names into every row, then sort."],
      table=[["Technique", "Looks centred", "Sort & Filter work"],
             ["Merge & Center", "Yes", "No"],
             ["Center Across Selection", "Yes", "Yes"],
             ["Wrap Text", "n/a", "Yes"]],
      split=4.90,
      so="Center Across Selection gives the same look without breaking anything. Merged cells cause problems later."),

    T("1.6.11",
      what="**Cell Styles** apply a saved set of formatting (for example Heading, Total, Bad) in one click. **Themes** "
           "change fonts and colours across the whole workbook. **Conditional Formatting** colours cells automatically "
           "according to their value, for example red for losses, so problems stand out.",
      like="Conditional formatting is like a traffic light: you don't read numbers, you see red and stop. The colour "
           "changes by itself when the situation changes.",
      how=["Select the values.",
           "Home › Conditional Formatting › Highlight Cells Rules › Less Than.",
           "Type 0 and choose Light Red Fill."],
      tip="Use one or two colours with a clear meaning (red = loss).",
      trap="Rainbow colours everywhere; when everything is highlighted, nothing stands out.",
      who="SuperMart Retail",
      case="The CEO's store table for 2025 lists profit for 12 stores. The analyst wants loss-making stores to turn red "
           "automatically.",
      steps=["Select the Profit column of the store summary.",
             "Conditional Formatting › Highlight Cells Rules › Less Than 0 › Light Red Fill.",
             "Add Data Bars to the Sales column to show size."],
      table=[["Store (FY2025)", "Profit", "Colour"],
             ["Abuja - Wuse", "−₦8.8m", "Red"],
             ["Onitsha - Main Market", "−₦6.2m", "Red"],
             ["Ibadan - Ring Road", "−₦4.1m", "Red"],
             ["Lagos - Ikeja", "₦29.4m", "None"]],
      split=4.90,
      so="Five of the 12 stores turn red at once. If a store's profit changes, its colour updates too."),

    T("1.6.12",
      what="**Freeze Panes** keeps headers on screen while you scroll. Excel freezes the rows **above** and the columns "
           "**to the left of** the cell you click first. **Freeze Top Row** and **Freeze First Column** are one-click "
           "options. **Split** divides the window so you can look at two distant parts of a sheet at once.",
      like="Like the fixed scoreboard at a football match: however the game moves around, the team names and score "
           "stay visible.",
      how=["Click the cell below the header row and right of the ID column.",
           "View › Freeze Panes › Freeze Panes.",
           "To undo: View › Freeze Panes › Unfreeze Panes."],
      tip="Click the cell just below and right of what you want to freeze.",
      trap="Clicking a cell in the middle of the sheet first; half the screen gets stuck.",
      who="SuperMart Retail",
      case="When scrolling down to row 5,000 of the SuperMart file, the analyst forgets which column is Sales and which "
           "order is which. She wants row 1 and the Order_ID column fixed.",
      steps=["Click cell C2 (below row 1, right of column B, Order_ID).",
             "View › Freeze Panes › Freeze Panes.",
             "Scroll down and right: row 1 and columns A–B stay visible."],
      table=[["You want to freeze", "Click first", "Then choose"],
             ["Row 1 only", "anywhere", "Freeze Top Row"],
             ["Column A only", "anywhere", "Freeze First Column"],
             ["Row 1 and columns A–B", "C2", "Freeze Panes"]],
      split=4.90,
      so="Headers and order numbers stay in view across all 9,994 rows, so you never lose your place."),

    T("1.6.13",
      what="Before printing or saving to PDF, set up the page: **orientation** (portrait or landscape), **margins**, "
           "**Print Area** (only the cells you want), **Print Titles** (repeat the header row on every page) and **Fit to "
           "page**. Always check **Print Preview** (**Ctrl+P**) first.",
      like="It is like arranging a newspaper page before it goes to press. Headlines, columns and margins must fit the "
           "paper, or the story is cut off.",
      how=["Page Layout › Orientation › Landscape for wide tables.",
           "Page Layout › Print Titles › Rows to repeat at top: $1:$1.",
           "Ctrl+P › Fit All Columns on One Page; check the preview."],
      tip="Always repeat the header row on long printed lists.",
      trap="Printing without a preview; column U ends up alone on page 2.",
      who="SuperMart Retail",
      case="A store manager prints the 2025 sales lines for her store. Page 1 has headers, but pages 2–10 show numbers "
           "with no labels, and the Profit column falls onto separate pages.",
      steps=["Page Layout › Orientation › Landscape.",
             "Page Layout › Print Titles › Rows to repeat at top › $1:$1.",
             "Ctrl+P › Scaling › Fit All Columns on One Page."],
      table=[["Setting", "Choose", "Why"],
             ["Orientation", "Landscape", "Wide tables fit"],
             ["Print Titles", "$1:$1", "Headers on every page"],
             ["Scaling", "Fit all columns", "No orphan columns"],
             ["Footer", "Page 1 of ?", "Pages stay in order"]],
      split=4.90,
      so="Every page now has headers and all columns, so the printout can be read on its own."),

    T("1.6.14",
      what="Three tools fix data problems quickly. **Find & Replace** (**Ctrl+H**) changes every copy of a value at once. "
           "**Go To Special** (**Ctrl+G › Special**) selects all blanks, all formulas or all errors. **Data Validation** "
           "stops wrong entries being typed in the first place, for example with a drop-down list.",
      like="Find & Replace is like a teacher correcting the same spelling mistake in every exercise book at once. Data "
           "Validation is the rule that stops pupils making it again.",
      how=["Ctrl+H: type what to find and what to replace it with.",
           "Ctrl+G › Special › Blanks to select every empty cell.",
           "Data › Data Validation › List to allow only set values."],
      tip="Tick 'Match case' and 'Match entire cell contents' for exact replacements.",
      trap="Replace All without 'Match entire cell contents' can change parts of other words.",
      who="SuperMart Retail",
      case="In " + DIRTY + " the Ikeja store is typed three ways (642 correct, 21 lower case, 19 capitals), and there are "
           "15 blank rows. New clerks keep making the same mistakes.",
      steps=["Ctrl+H, tick Match case: replace 'LAGOS - IKEJA' and then 'lagos - ikeja' with 'Lagos - Ikeja'.",
             "Ctrl+G › Special › Blanks, then delete the selected empty rows.",
             "Add Data Validation › List with the 12 correct store names."],
      table=[["Tool", "Problem in the dirty file", "Result"],
             ["Find & Replace", "40 Ikeja spellings wrong", "All 682 Ikeja rows match"],
             ["Go To Special", "15 blank rows", "Selected in one step"],
             ["Data Validation", "New misspellings", "Only the 12 names allowed"]],
      split=4.30,
      so="Three tools clean up the past and protect the future. This is a big step towards trustworthy data."),
]
