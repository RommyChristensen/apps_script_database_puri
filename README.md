# AOG Youth Puri — Coach Dashboard

A Google Apps Script web app that reads the **DATABASE tab** (the first tab of
`[DATABASE] AOG YOUTH PURI 01`) and shows it grouped by **Coach → CG → Jemaat**.

- **Coach cards**: number of people, number of CGs, % active, position mix and average growth
- **Coach drawer** (slides in): tabs for CG (members per group), Jemaat (filter by position/status) and Growth (5M / MSJ / CGT completion, click to see who hasn't finished)
- **Person modal**: growth journey, path (Member → Sponsor → CGL → Coach), personal info, notes, plus WhatsApp / call / email buttons and a link to the person's row in the sheet
- **Global search** (press `/`), **Growth Gaps** heatmap and a composition chart per coach
- **Edit per person**: in the person modal tap **Edit**. You can check or uncheck 5M / MSJ / CGT
  (tap a chip, or use "Centang semua" for a whole group) and change the *Jenjang*: Posisi
  (VIP → Simpatisan → Member → Sponsor → CGL) plus target and actual dates for Member, Sponsor, CGL and Coach.
  The 📅 button fills in the current month. Changes are saved straight to the sheet.
  The **Data Diri** section edits Gender, No HP, Email, NIJ, birth city, birth date (date picker),
  Domisili, Kerja/Kuliah, Kampus/Sekolah, Pelayanan and Alamat. Text fields suggest values already
  in the database (e.g. campus names) so spelling stays consistent.

| File | Purpose |
| --- | --- |
| `Code.gs` | `doGet`, `getDashboardData`, and the sheet parser (finds columns by header name) |
| `Index.html` / `Stylesheet.html` / `JavaScript.html` | UI (white, blue and sky-blue theme, animated drawer and modal) |
| `appsscript.json` | Manifest: read-only Sheets scope, web app settings |
| `dev/build-preview.js` | Builds `dev/preview.html`: the real `Code.gs` running in the browser on a fake sheet |
| `dev/mock-apps-script.js`, `dev/test.js` | Mock SpreadsheetApp + tests for the parser and for saving |

## Deploy (copy and paste, about 5 minutes)

1. Open the spreadsheet → **Extensions → Apps Script**.
2. Create the files with the same names and paste in their contents:
   `Code.gs`, plus HTML files named `Index`, `Stylesheet` and `JavaScript`.
3. ⚙️ **Project Settings** → turn on *Show "appsscript.json"* → paste the contents of `appsscript.json` into it.
4. **Deploy → New deployment** → type **Web app**:
   - *Execute as*: **User accessing the web app**
   - *Who has access*: **Anyone with Google account**
5. **Deploy**, authorize, and copy the `/exec` URL. That's the dashboard link.

> 🔒 Because it runs as the *user accessing*, only people who already have
> access to the spreadsheet can see the data. Everyone else gets an error card.
> To give a coach access, share the spreadsheet with them (Viewer is enough).

> ✏️ **Editing** needs **Editor** access to the spreadsheet. A Viewer can still open the dashboard
> but gets the error "tidak punya akses edit" when saving. Because the app now has write
> permission, everyone is asked to **authorize again** the first time they open the new version.

### How saving works
- Only cells that **actually changed** are written. Other columns (including AGE / ELIGIBLE formulas) are never touched.
- Before writing, the server checks the name in that row still matches. If someone has added or removed
  rows in the meantime, the save is rejected and you're asked to refresh. Saves also take a lock so two
  people can't write at the same moment.
- Checked milestones are written as `✔` and unchecked as `-`, the same as the existing sheet. Cells that are
  real checkboxes stay as checkboxes. Date and target text is stored as plain text, so "Juni 2026" isn't turned
  into a date.
- Cells containing **formulas** are refused, not overwritten.
- Validation runs in the browser and again on the server: email format, No HP (digits, 6–20 characters),
  NIJ digits only, and Gender must be Male or Female. Mistakes are shown in red and Save stays disabled.
- Birth dates are written in the TTL column's format, e.g. `5 March 2004`. Age is recalculated automatically.
- No HP is written to the **first** `NO HP` column. The second `NO HP` column is left untouched.
- Every save is recorded in the **`Edit Log`** tab (created automatically as the last tab): time, who
  (email), row, name, and `old → new` for each change. To turn it off, set `CONFIG.LOG_SHEET_NAME = ''`.
- **Catatan / Keterangan** (coaching notes) can be edited too. Only columns that exist in the sheet
  are shown in the form. For example, if there's no EMAIL column, the Email field doesn't appear.
- If a person's `TARGET JD COACH` cell still holds notes (the old shifted header), the Target Coach
  and Catatan fields are locked for that person so the notes can't be wiped.

After you change the code: **Deploy → Manage deployments → ✏️ → Version: New version**
so the URL stays the same.

### Alternative: clasp
```bash
npm i -g @google/clasp && clasp login
clasp clone <SCRIPT_ID>   # or: clasp create --type sheets --title "AOG Dashboard"
clasp push && clasp deploy
```

## Merging the 5 coaches into one tab

The parser reads the **first tab** (set `CONFIG.SHEET_NAME` in `Code.gs` to use a different one).
Keep the layout of the current Angie tab:

- Main header row containing `NAMA` (row 4), then a sub-header row for
  `PERCAYA YESUS … CGT 3` (row 5). Data starts on the next row.
- Columns are found **by header name**, so column order doesn't matter, but the names must match.
  Accepted aliases: `KERJA/KULIAH` or `SEKOLAH/KULIAH`, and `KAMPUS` or `SEKOLAH`.

Mark which coach each person belongs to **in one of two ways** (both can be used together):

1. **A `COACH` column (recommended).** Add a header `COACH` in row 4 and fill it with
   `Angie`, `Lenita`, `Carla`, `Meysia`, and so on. Spell each name the same way on every row.
2. **Banner rows.** Paste each coach's block below the previous one, with one row that says
   `COACH LENITA` (in column A, and leave NAMA empty) above it. The current `COACH ANGIE` in row 1
   already works this way.

The first row of each block (no CG and no POSISI, e.g. "Angie Lie") is treated as the coach themself:
their full name is shown on the card, and they aren't counted as a member.
Rows without a name (like the summary block at the bottom) are skipped.

### Notes on the current tab ("Full Database AOG Puri")
- It has a `COACH` column (column D), so each person is grouped by that column. The `COACH ANGIE`
  banner in row 1 is no longer needed (it's harmless if left there).
- The `KETERANGAN` column has been deleted, and the notes ("Keterikatan dosa", "Time Management"…)
  now sit under the `TARGET JD COACH` header. The dashboard detects this: if there's no `KETERANGAN`
  column and most of the `TARGET JD COACH` cells look like notes, the whole column is read as
  KETERANGAN, and a yellow warning appears. **Fix:** rename that header to `KETERANGAN`. The warning
  disappears, and if you want a real "Target jadi Coach" column, add a new column with that header.
- Lenita/Meysia/Carla's tabs have **a different column order** (NAMA is in column G, not C) and Carla
  has no EMAIL column. Because matching is by header that's fine, as long as you paste each column
  under the header with the same name in this tab.

## Local preview
```bash
node dev/build-preview.js            # demo data (fake)
node dev/build-preview.js export.csv # CSV of the database tab (stays local, git-ignored)
# open dev/preview.html in a browser. Edits also work there (saved to a fake in-memory sheet)

node dev/test.js                     # tests for the parser and updatePerson
```
