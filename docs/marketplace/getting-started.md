# Getting started with Spreadsheets for YouTrack

Spreadsheets for YouTrack needs YouTrack 2026.2 or later. An administrator sets it up once for each project. After that, people open spreadsheets from the issues of those projects.

## Set up (administrators)

1. **Install the app.** Click **Install** on this page and enter the address of your YouTrack, or go to **Administration > Apps** in YouTrack and click **Add app**.
2. **Attach it to projects.** In **Administration > Apps**, open **Spreadsheets for YouTrack**, go to **Projects** and click **Manage projects**. Select the projects that should have spreadsheets. The Spreadsheet section and the **Open spreadsheet** menu item appear only in issues of these projects.
3. **Choose who can edit (optional).** By default, users with the **Update Issue** permission in the project can create and edit spreadsheets, whichever role grants it (global, project or custom roles), and users who can only read an issue can view its spreadsheet. To change this for a project, open the project's settings, go to **Apps**, select **Spreadsheets for YouTrack** and open its **Settings** tab. There you can choose **Groups that can edit spreadsheets** (members edit even without Update Issue), **Groups that can only view spreadsheets** (members never edit) and **Everyone who can see the issue can edit**. Click **Save**. You need permission to update the project. The settings apply to one project at a time.

## First use

4. **Open an issue** in an attached project. The **Spreadsheet** section is above the activity stream. If it is collapsed, click its header.
5. **Create or import a spreadsheet.** Click **Create spreadsheet** for an empty spreadsheet, or **Import .xlsx or .csv file…** to start from an .xlsx, .csv, .tsv or .txt file. You can also drop the file on the section. Users without edit rights see "No spreadsheet yet" instead.
6. **Enter data and formulas.** Type a value, or a formula that starts with `=`, and press Enter or Tab to save the cell. Other people who have the spreadsheet expanded, or open in the large window, see the change within a few seconds.
7. **Structure the data.** Select a range and choose **Data > Format as table…** for a header row, a total row and filter buttons. Use **Insert > Checkbox**, **Insert > Dropdown list**, **Format > Conditional formatting…** and **Insert > Chart** for checklists, status columns, highlights and charts.
8. **Open the large window.** In the issue's options menu, choose **Open spreadsheet**.

## Everyday tasks

- **Collapse or expand:** click the header of the **Spreadsheet** section. YouTrack remembers your choice, and a collapsed section loads nothing.
- **Change the height:** drag the handle under the grid, or double-click it to reset. Height and zoom apply to all issues you open in this browser.
- **Show amounts in a currency:** choose **Format > Number > Currency** and pick a symbol. The currency button on the toolbar uses the currency of your YouTrack or browser locale when the app knows it, and $ otherwise. Its tooltip shows the symbol, for example Currency (€).
- **See and restore changes:** choose **Change history** in the bar's **⋯** menu, or **View > Change history**. On an entry that changed cell values, click **Restore these values** to put the old values and formats back as a new change. The button is shown only to users who can edit. It is not offered for changes to rows, columns or sheets, or for entries that changed more than 60 cells.
- **Find the best values with Solver:** choose **Data > Solver…**. Enter the objective cell and choose **Max**, **Min** or **Value Of**, enter the variable cells, click **Add** for each constraint, select a solving method and click **Solve**. The button next to a reference field lets you select the cells on the sheet. In **Solver Results**, keep the solution or restore the original values, and select reports to add them as new sheets. The model is saved with the sheet.
- **Try values with Goal Seek and scenarios:** **Data > What-if analysis > Goal Seek…** finds the value of one cell that makes a formula reach a target. **Scenario Manager…** saves sets of values for up to 32 cells, shows them on the sheet and compares them in a summary sheet.
- **Use semicolons in formulas:** if your language writes decimals with a comma (1,5), formulas use ; between arguments and , in numbers, for example =ROUND(1,5; 0). To change this, choose **View > Formula separators** (also in **More tools**, the **⋯** button at the end of the toolbar) and pick **Automatic**, **Comma** or **Semicolon**. The choice applies only to you, in this browser; formulas are stored in one form, so everyone sees them in their own style and downloaded .xlsx files are correct.
- **Download:** in the bar's **⋯** menu, choose **Download as .xlsx** for the whole spreadsheet, or **Download sheet as CSV** for the current sheet. For a CSV file you choose the separator between values (comma, semicolon or tab) and the decimal separator of numbers. The same items are in the **Data** menu.
- **Delete the spreadsheet:** in the bar's **⋯** menu, choose **Delete spreadsheet…** and confirm. This deletes all sheets and the change history of the issue and cannot be undone. Download a copy first if you may need it.

Users with view-only access see a **View only** tag in the bar ("You can see this spreadsheet but not edit it.") and no editing controls. They can still search, copy, open the change history and download.

## Where to find things

- **Edit:** Undo, Redo, Cut, Copy, Paste, Paste special, Find and replace, Fill down, Fill right, and deleting values, rows, columns and notes.
- **View:** Freeze, Gridlines, Show formulas, hidden rows and columns, Zoom, Change history.
- **Insert:** rows, columns, Sheet, Table, Chart, Function, Link, Note, Checkbox, Dropdown list, Today's date.
- **Format:** Format cells…, Number, Text, Alignment, Wrap text, Borders, Merge cells, Conditional formatting…, Table style…, column width and row height, Clear formatting.
- **Data:** Sort A → Z, Sort Z → A, Sort range…, Filter, Format as table…, Remove duplicates…, Split text to columns…, Data validation…, Named ranges…, What-if analysis (Scenario Manager, Goal Seek), Solver…, import and download.

## Keyboard shortcuts

- **Move and select:** arrow keys; Ctrl+arrow to jump to the edge of the data; Shift to extend the selection; Ctrl+click or Ctrl+drag (Cmd on a Mac) to add more cells, ranges, rows or columns to the selection; Ctrl+A to select the data region, and Ctrl+A again to select the whole sheet.
- **Edit:** type, double-click a cell or press F2 to edit; Enter or Tab to save and move (add Shift to go back); Alt+Enter for a new line in a cell; Ctrl+Enter to fill the selection; Esc to cancel; F4 to switch a reference between relative and absolute.
- **Clipboard and undo:** Ctrl+C, Ctrl+X and Ctrl+V to copy, cut and paste; Ctrl+Shift+V to paste values only; Ctrl+Z and Ctrl+Y to undo and redo.
- **Format:** Ctrl+B, Ctrl+I, Ctrl+U and Ctrl+5 for bold, italic, underline and strikethrough.
- **Data:** Ctrl+D and Ctrl+R to fill down and right; Alt+= for AutoSum; Ctrl+Shift+L to filter; Alt+Down to open a dropdown list; Ctrl+F and Ctrl+H to find and replace.
- **Insert:** Ctrl+K for a link; Shift+F2 for a note; Ctrl+; for today's date; Ctrl+Shift+; for the current time.
- **View:** Ctrl+` to show or hide formulas; Alt+PageUp and Alt+PageDown to go to the previous or next sheet.

## Before you remove the app

Spreadsheets are stored in the app's extension properties. Download the spreadsheets you want to keep with **Download as .xlsx** before you remove the app.
