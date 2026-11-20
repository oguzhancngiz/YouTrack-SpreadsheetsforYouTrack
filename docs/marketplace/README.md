# JetBrains Marketplace listing: Spreadsheets for YouTrack

Texts and screenshots for the Spreadsheets for YouTrack page on JetBrains Marketplace. They were written to follow the Marketplace rules
(approval guidelines, listing best practices, the YouTrack app submission checklist) and checked claim by claim against
the code. Keep them in sync with the app: the YouTrack team tests every advertised feature and checks that menu names and
labels match.

## Where each text goes

| Marketplace field | File | Notes |
| --- | --- | --- |
| Name | — | `Spreadsheets for YouTrack` (from `manifest.json` `title`; 25 characters). Note: Approval Guideline 1.2.h says a name must not include the name of a JetBrains product; other apps on Marketplace use "YouTrack" in their names, but the review may ask for a different name. The app ID is `spreadsheets-for-youtrack` (up to 1.2.2: `issue-sheets`); if the review asks for another name, change only `title`, not the ID. |
| Description | [description.html](description.html) | Paste into the description editor and set the description source to the Marketplace version so later uploads don't replace it. The first 40 characters, "A shared spreadsheet in YouTrack issues.", become the preview card. |
| Description in the ZIP | `manifest.json` `description` | Plain text, same first sentence. Shown in YouTrack under Administration > Apps. |
| Change notes | [change-notes.html](change-notes.html) | `manifest.json` `changeNotes` is a shorter plain-text version of the same release notes. |
| Tags | [tags.txt](tags.txt) | Issue View, Productivity, Data Visualization |
| Getting Started | [getting-started.html](getting-started.html) | HTML for the Getting Started field. [getting-started.md](getting-started.md) is the same text in Markdown. |
| Custom page "FAQ" | [faq.md](faq.md) | Custom pages take Markdown. |
| Custom page "Privacy" + Privacy Policy URL | [privacy.md](privacy.md) | Publish as a Custom page (or on the Codeo website) and put its URL in the Privacy Policy field. |
| Media | [screenshots/](screenshots/) | See below. |
| Short description for sharing | [short-description.txt](short-description.txt) | Not a Marketplace field; for the vendor website, posts and the GitHub repository description. |

## Screenshots

All screenshots are 1280 × 800 PNG (one aspect ratio); the Marketplace ones (screenshot-1 to screenshot-9) use the light
theme and extra/dark-theme.png the dark theme. They were taken from the development preview with
`docs/capture.py`, so they show the app but not the surrounding YouTrack page. Marketplace asks for in-app screenshots
of the YouTrack interface: replace them with captures from a real YouTrack issue page (cropped to the page, no browser
window) once the app runs on your test instance, using the same scenes. The demo data uses fictional names.

| File | Caption |
| --- | --- |
| screenshot-1-spreadsheet.png | A budget in an issue: a table with a total row, data bars, a status list and a chart. Two other people are editing; their cells are outlined in their color. |
| screenshot-2-filter.png | Filter a table column by a list of values or by a condition, and sort it. |
| screenshot-3-formulas.png | Formulas with argument hints. Table references such as Budget[Category] work in formulas. |
| screenshot-4-test-matrix.png | A test matrix with checkboxes, a pass rate per row and a color scale. |
| screenshot-5-large-window.png | The large window from the issue's options menu (Open spreadsheet). |
| screenshot-6-conditional-formatting.png | Conditional formatting rules: a data bar on the totals and value rules that color the status column. |
| screenshot-7-history.png | The change history: who changed which cells and when, with old and new values. |
| screenshot-8-solver.png | Solver on a capacity plan: the cheapest set of build agents for the required builds per hour, with integer constraints. |
| screenshot-9-solver-results.png | Solver Results: keep the solution or restore the original values, and add reports as new sheets. |
| extra/dark-theme.png | The dark theme, for the GitHub README. Marketplace screenshots should use the default (light) theme. |

To retake them: `python serve.py 8766`, then `python docs/capture.py`.

## Before you publish

Done in the code and package:

- [x] Description and change notes in `manifest.json` replaced; no third-party brand names in the listing or the menus
  (the import and download items are now "Import .xlsx or .csv file…" and "Download as .xlsx").
- [x] `HYPERLINK()` opens only `http`, `https` and `mailto` links.
- [x] Custom formula validation is checked; SUBTOTAL and AGGREGATE leave out filtered rows (table total rows follow the
  filter); `.xlsx` download writes table references as ranges and newer functions with the prefixes other apps expect.
- [x] The currency button follows the locale (₺ for Turkey and plain "tr"; $ when the locale is unknown).
- [x] No Ctrl+T hint (browsers keep that shortcut); sheets switch with Alt+PageUp/PageDown.
- [x] App icons (icon.svg, icon-dark.svg) declare 40 × 40 px (SVG); the widget icon is a 24 × 24 viewBox. `THIRD-PARTY-NOTICES.txt` carries the Apache 2.0 notice for the Ring UI variables.
- [x] README.md (with development URLs) is no longer in the ZIP.

Still to do (needs your YouTrack, your accounts or your decision):

- [ ] **Test on a live YouTrack** with the app hidden on Marketplace until it passes, as an admin and as a regular user,
  in light and dark themes. Check in particular: the issue widget's frame takes the panel's height; the dark theme
  class reaches the frame; layout preferences are remembered (host.storage); the project setting reaches the backend;
  history paging; .xlsx and .csv downloads inside the app frame; the Cut, Copy and Paste menu items; presence; saving
  large spreadsheets (several MB) in extension properties; and whether YouTrack shows the app's property changes in the
  issue's activity or updated time. Remove or fix any claim that fails. The listing says "YouTrack 2026.2 or later";
  add "Cloud and Server" only after testing both.
- [ ] **Confirm the paths in Getting Started** on YouTrack 2026.2: Install from the Marketplace page, Administration >
  Apps > Add app, Projects > Manage projects, and where the project setting "Everyone who can see the issue can edit"
  appears (step 3 says "open the Spreadsheets for YouTrack settings for that project").
- [ ] **Replace the screenshots** with in-app captures (see above).
- [ ] **Vendor profile:** a working website and email, and the trader declaration. The description and the privacy page send users to the
  vendor email.
- [x] **License:** Apache License 2.0. LICENSE and NOTICE are in the ZIP, and each source file has the Apache notice.
  Because it is an open-source license, the Source code URL field is required (JetBrains Marketplace approval
  guidelines): the public repository below.
- [ ] **Links:** the repository is https://github.com/delgarathor/YouTrack-SpreadsheetsforYouTrack, and
  `manifest.json` has it as `url`. All links on the listing must be reachable, so upload to the Marketplace only after
  the repository is public with Issues on. Source code URL: the repository; Issue tracker:
  https://github.com/delgarathor/YouTrack-SpreadsheetsforYouTrack/issues.
- [ ] **Price:** the listing doesn't mention one. If the app is free, you can add "Spreadsheets for YouTrack is free." If it is paid,
  add a licensing section and a support email (answers within 2 business days for critical requests).
- [ ] **Downloads by readers:** everyone who can see an issue can download its spreadsheet. The listing says so. If
  only editors should download, change the code first and update the description, FAQ and privacy page.
- [ ] **Removing the app:** confirm on a test instance what happens to the spreadsheets, then say it plainly in the
  description, FAQ and privacy page (they now only advise downloading first).
- [ ] Search Marketplace for an existing "Spreadsheets for YouTrack" before the first upload. Upload with Ads: none, Stable channel.
  If there is no answer within 3–4 working days, write to marketplace@jetbrains.com.

Keep out of the listing (not verified or not true): real-time, instant or conflict-free collaboration; audit trail;
speed or "handles large data"; accessibility compliance; "all spreadsheet features".
