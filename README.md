# AOG Youth Puri — Coach Dashboard

A Google Apps Script web app that reads the **DATABASE tab** (the first tab of
`[DATABASE] AOG YOUTH PURI 01`) and shows it grouped by **Coach → CG → Jemaat**.

- **Coach cards**: number of people, number of CGs, % active, position mix and average growth
- **Coach drawer** (slides in): tabs for CG (members per group), Jemaat (filter by position/status) and Growth (5M / MSJ / CGT completion, click to see who hasn't finished)
- **Person modal**: growth journey, path (Member → Sponsor → CGL → Coach), personal info, notes, plus WhatsApp / call / email buttons and a link to the person's row in the sheet
- **Global search** (press `/`), **Growth Gaps** heatmap and a composition chart per coach
- Read-only: the dashboard never writes to the spreadsheet

| File | Purpose |
| --- | --- |
| `Code.gs` | `doGet`, `getDashboardData`, and the sheet parser (finds columns by header name) |
| `Index.html` / `Stylesheet.html` / `JavaScript.html` | UI (white, blue and sky-blue theme, animated drawer and modal) |
| `appsscript.json` | Manifest: read-only Sheets scope, web app settings |
| `dev/build-preview.js` | Builds `dev/preview.html` with demo data so you can see the UI locally |

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

### Data notes found in the current tab
- Lenita/Meysia/Carla's tabs have **a different column order** (NAMA is in column G, not C) and Carla
  has no EMAIL column. Because matching is by header that's fine, as long as the header names match
  when everything is pasted under one header.
- In the Angie tab the headers at the end are shifted: notes ("Mau MSJ 2…") sit under
  `TARGET JD COACH`, and the `KETERANGAN` column holds ages. The dashboard
  recovers these notes automatically, but fixing the header in the sheet is better.

## Local preview
```bash
node dev/build-preview.js            # demo data (fake)
node dev/build-preview.js export.csv # CSV of the database tab (stays local, git-ignored)
# open dev/preview.html in a browser
```
