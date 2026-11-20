# Spreadsheets for YouTrack privacy statement

Version 1.0.0. Last updated: November 17, 2026.

## Summary

Spreadsheets for YouTrack stores spreadsheet data inside your YouTrack. Layout preferences are kept in each user's browser (see below). The app makes no network requests to servers outside your YouTrack. It sends nothing to Codeo or to any other third party, and Codeo has no access to your YouTrack or to the data the app stores.

## What the app stores and where

**On each issue that has a spreadsheet** (issue extension properties):

- `sheetDoc`: the spreadsheet content, including sheets, cell values and formulas, formats, notes, tables, rules, charts and named ranges. It also holds the login of the user who created the spreadsheet and the login and display name of the user who last changed it.
- `sheetLog`: the most recent changes, at most 300 batches (a batch is a group of changes that a browser sends together) and normally no more than 1,500,000 characters (a single larger batch, such as an imported file, is kept whole). Each batch holds the changed content and the login of the user who sent it. Open browsers use it to catch up after missing updates.
- `sheetHistory`: the change history, at most 500 entries. Each entry holds the login and display name of the user, the time, a short description of changes to rows, columns, sheets, tables, rules, charts and names, and the old and new value and format of up to 60 cells. Values longer than 2,000 characters are shortened to their first 2,000 characters.

**On each user who has an issue with the app open, in the issue or in the large window** (user extension properties):

- `sheetPresence` and `sheetPresenceIssue`: the issue, the sheet and the cells the user has selected, whether they are editing a cell, and a timestamp. Other users on the same issue see this as an avatar and a colored selection. Presence older than 30 seconds is ignored. The record stays on the user until it is replaced or cleared.

**In the user's browser, in the app storage that YouTrack provides to the widget (if available; it is not tied to the YouTrack account):**

- The panel height, the zoom level, the formula separators you chose and the options of your last CSV download (separator and decimal mark). No spreadsheet content is stored here. If this storage is not available, nothing is saved and the defaults apply.

The app does not set cookies and does not use the browser's local storage directly.

## Personal data

The app processes these kinds of personal data, all inside YouTrack:

- YouTrack user logins and display names, used for the "edited … by …" text in the bar, for presence and for the change history.
- Timestamps of edits and of presence updates.
- Any personal data that users type, paste or import into cells or notes.

The app does not collect any other personal data.

## Who can see the data

- Everyone who can see an issue can view its spreadsheet, its change history (including old values and the names of the people who made each change) and the presence of other people on that issue, and can download the spreadsheet as an .xlsx file or the current sheet as a .csv file.
- Users who can update the issue, as YouTrack's permissions decide it for that issue, can change the spreadsheet, delete change history entries and delete the spreadsheet. If the project setting **Everyone who can see the issue can edit** is on, everyone who can see the issue can do this.
- The data is part of your YouTrack and is hosted wherever your YouTrack is hosted.

## What leaves YouTrack

- **Network:** nothing. The widget calls only the app's own endpoints inside YouTrack. There is no telemetry, analytics, license check or content delivery network, and the app loads no code from the internet.
- **Downloads:** .xlsx and .csv files are created in the user's browser and saved to the user's device only when the user downloads them.
- **Clipboard:** the app reads the clipboard only when a user pastes, and writes to it only when a user copies or cuts.
- **Links:** a link in a cell opens only when a user opens it on purpose: Ctrl+click (Cmd+click on a Mac) or a middle-click. Only http, https and mailto links are shown as links. A web link opens in a new browser tab and a mailto link opens the user's email app.

## Formulas

Formulas are calculated in each user's browser. The server stores only what users enter and never runs formulas.

## Retention and deletion

- The change history keeps at most 500 entries. Older entries are removed. Editors can delete single entries or clear the whole history.
- A value that was removed from a cell stays in the change history until its entry is removed, and in the log of recent changes until newer changes push it out of the log or the spreadsheet is deleted.
- The log of recent changes keeps at most 300 batches. Older batches are removed.
- **Delete spreadsheet…** removes the spreadsheet content, the change log and the change history of the issue. This cannot be undone.
- Presence records are replaced as users move between cells and issues, and are ignored after 30 seconds.
- Download any spreadsheets you want to keep before you remove the app from YouTrack.
- The content that users enter into cells stays under the control of the owner of the YouTrack site.

## Contact

For questions about this statement, contact Codeo at the email address shown on the vendor profile of Spreadsheets for YouTrack on JetBrains Marketplace.
