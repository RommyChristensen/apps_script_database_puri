# AOG Youth Puri — Coach Dashboard

A Google Apps Script web app that reads the **DATABASE tab** (the first tab of
`[DATABASE] AOG YOUTH PURI 01`) and shows it grouped by **Coach → CG → Jemaat**.

- **Login per coach**: each coach signs in with a username and password from the
  **Authentication** tab and only sees (and can only edit) the people of their own coach group.
  Admin accounts (the team leader) see every coach.

- **Menu**: *Dashboard*, *Pertemuan*, *LDP* and *Lokasi* (in the topbar on desktop, a bottom tab bar on phones).
- **Lokasi CG**: a map of where every CG meets (OpenStreetMap/CARTO tiles via Leaflet, no API key), one color per
  coach. Each pin has *Buka di Google Maps* and *Rute* buttons. Coaches see all CGs but set or change only their own
  CGs' location: paste a Google Maps link, type an address (geocoded automatically), or click the map. Online CGs and
  CGs without a location are listed separately; the *Domisili jemaat* layer shows how many members live in each area
  (counts only, and coaches only see their own members).
- **LDP (Leaders Development Program)**: every person's path Member → Sponsor → CGL → Coach → Team Leader
  with a *target* date and a *realization* date per step. Each month, record the result for the person's
  current step: **Berhasil** fills the realization date and moves POSISI up (Member/Sponsor/CGL);
  **Belum** needs a note and can move the target. Filters for *perlu update bulan ini*, overdue targets and
  people without a target; every update is kept in the **LDP Progress** tab.
- **Log Pertemuan**: coaches record every meeting with their CGLs / CG members: date, with whom (picked
  from their own people), a purpose category + description, and an action list where each task has a
  deadline and a done-checkmark the coach can tick. Filters for open / overdue / done tasks, purpose and
  text search; overdue tasks show a red badge on the menu. Each person's profile lists their meetings
  with a *Catat meeting* shortcut.
- **Coach cards**: number of people, number of CGs, % active, position mix and average growth
- **CG page**: in the coach drawer, click a CG card (or *Lihat CG*) to open a page for that Connect Group: schedule,
  status and area (taken from the **CGL's row**; without a CGL, the most common value among members), location from
  the Lokasi CG tab with Google Maps / Rute buttons, number of people, average attendance, leaders, average growth,
  position mix, a *Jemaat* tab (CGL → Sponsor → Member → Simpatisan → VIP, with filters) and a *Growth* tab.
  *Ubah info CG* changes WAKTU / STATUS / DAERAH CG on **every row of that CG** (the page keeps showing the CGL's
  row; coaches: own CGs; admin: all).
- **Restructuring** (buttons in a person's profile):
  - *Status* (Aktif / Non-Aktif) is editable in the Edit form.
  - *Pindah CG* moves a person to another CG: NO CG, CGL, COACH and the CG schedule fields on their row follow the
    target CG. Coaches move people between their own CGs; admin anywhere.
  - *Hapus* (admin only, re-type your password): the row is copied to the **Arsip Jemaat** tab (date, who, reason,
    original row number, key fields and every cell as `HEADER: value`) and then deleted from the database.
- **Coaches and the Team Leader** in Pertemuan and LDP: admins can pick them in Pertemuan and set their LDP
  (e.g. Coach → Team Leader). A coach sees the entries about themselves, read-only.
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
| `Code.gs` | `doGet`, `getDashboardData`, `updatePerson`, and the sheet parser (finds columns by header name) |
| `Auth.gs` | Login, password hashing, sessions, and the per-coach access checks |
| `Meetings.gs` | Log Pertemuan: read, save, tick a task, delete (per-coach access) |
| `MeetingsUi.html` | Menu, Pertemuan page, the form and the Pertemuan section in profiles |
| `Ldp.gs` | LDP: targets, realization dates and the monthly berhasil/belum updates |
| `LdpUi.html` | LDP page and the per-person LDP update form |
| `Restructure.gs` | CG info for all rows, Pindah CG, delete with archive (password check) |
| `RestructureUi.html` | Pindah CG and Hapus forms |
| `Multiply.gs` | Multiplikasi: new CG, leader promoted to CGL, history in the "Multiplikasi CG" tab |
| `MultiplyUi.html` | Multiplikasi menu (history) and the new-CG form |
| `Visits.gs` | Visit CG: schedules and visit logs in the "Visit CG" tab |
| `VisitsUi.html` | Visit CG menu (Jadwal, Log, Cakupan) and forms |
| `CgUi.html` | CG overview page and the CG info form (`updateCgInfo` in `Code.gs`) |
| `Locations.gs` | Lokasi CG: list, set location, Maps link parsing, geocoding (cached), domisili layer |
| `LocationsUi.html` | Lokasi page: Leaflet map, CG list, edit form with a mini map |
| `Index.html` / `Stylesheet.html` / `JavaScript.html` | UI (white, blue and sky-blue theme, animated drawer and modal) |
| `appsscript.json` | Manifest: Sheets scope, web app runs as the owner, open to anyone (login required inside) |
| `dev/build-preview.js` | Builds `dev/preview.html`: the real `Code.gs` running in the browser on a fake sheet |
| `dev/mock-apps-script.js`, `dev/test.js` | Mock Apps Script services + tests for the parser, saving and login |

## Deploy (copy and paste, about 5 minutes)

1. Open the spreadsheet → **Extensions → Apps Script**.
2. Create the files with the same names and paste in their contents:
   script files `Code`, `Auth`, `Meetings`, `Ldp`, `Locations`, `Restructure`, `Multiply` and `Visits`, plus HTML files named `Index`,
   `Stylesheet`, `JavaScript`, `MeetingsUi`, `LdpUi`, `LocationsUi`, `CgUi`, `RestructureUi`, `MultiplyUi` and `VisitsUi`.
3. ⚙️ **Project Settings** → turn on *Show "appsscript.json"* → paste the contents of `appsscript.json` into it.
4. **Deploy → New deployment** → type **Web app** (or *Manage deployments → ✏️* for the existing one):
   - *Execute as*: **Me** (the spreadsheet owner)
   - *Who has access*: **Anyone**
5. **Deploy**, authorize, and copy the `/exec` URL. That's the dashboard link for every coach.

## Logins (Authentication tab)

| USERNAME | PASSWORD | COACH | ROLE | CATATAN |
| --- | --- | --- | --- | --- |
| admin | *(type one)* | | ADMIN | sees every coach |
| rommy | *(type one)* | Rommy Christensen | TL | sees every coach whose rows have TL = Rommy Christensen |
| angie | *(type one)* | Angie | COACH | sees only COACH = Angie |

- **COACH** must match the value in the database's COACH column (capitals don't matter).
- **ROLE**: `COACH` (default), `TL` (Team Leader: put the TL's name in COACH, exactly as in the database's
  **TL** column; sees and edits every coach under them, and records Pertemuan / LDP about those coaches, which the coach sees read-only), `ADMIN` (sees everyone and gets the spreadsheet links),
  or `NONAKTIF` (blocked). A TL's list of coaches is refreshed every 5 minutes.
- **Passwords**: type them in plain text. After the person's first successful login the app replaces the
  password with a salted hash (`sha256$…`), so the tab never keeps readable passwords. To **reset** a
  password, type a new one over the hash. An empty password means the account can't sign in.
- Changing someone's password, COACH or ROLE **signs them out** on every device straight away.
- Sessions last 6 hours from the last activity. After 5 wrong passwords a username is locked for 15 minutes.
- Every edit in the **Edit Log** tab records who made it, e.g. `angie (coach Angie)`.

> 🔒 **Why the app runs as you now.** The server checks the login and only sends each coach their own
> people. That protection only works if coaches can't simply open the spreadsheet. **Remove coaches'
> direct access to the spreadsheet** (Share → remove, or keep only people who should see everything).
> They don't need it anymore: the app reads and writes the sheet on your behalf.

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

## TL column and the Visit CG tab

- The database has a **TL** column (right after COACH) with the Team Leader of each row. The TL of a coach is
  the most common value among that coach's rows. Pindah CG and Multiplikasi to another coach also update TL.
- **Visit CG** menu: coaches visit their own CGs, a TL every CG of their coaches, admin every CG.
  - **Jadwalkan visit**: CG, date, time, focus of the visit (status `Terjadwal`).
  - **Isi log** on a schedule (or **Catat visit** for an unplanned one): visit date (not in the future),
    attendance, evaluation and the action for that CG's CGL (status `Selesai`).
  - **Cakupan CG** lists every CG with its last and next visit; CGs not visited for 3 months stand out.
  - Everyone sees all visits to CGs they can access (a coach also sees the TL's visits), but only the author
    or an admin can change or delete a visit.
- Stored in the **Visit CG** tab (created automatically): ID, CG, COACH, CGL, TANGGAL, JAM, PENGUNJUNG, PERAN,
  RENCANA, STATUS, TGL VISIT, HADIR, EVALUASI, ACTION CGL, DIBUAT OLEH, DIBUAT PADA, DIPERBARUI.

## Multiplikasi CG tab

The **Multiplikasi** menu (and the "Multiplikasi" button on a CG page) creates a new CG out of an existing one:

- Choose the source CG, the new CG number (e.g. `aog 5` is saved as `AOG 05`, and it must not be in use yet),
  the coach of the new CG, the date, the leader (anyone you can see, from any CG), the schedule, an optional
  location, and tick the members who move.
- The rows of the leader and the ticked members get the new NO CG, CGL, COACH, WAKTU / STATUS / DAERAH CG.
- If the leader is below CGL, they are promoted: POSISI = CGL, TGL JD CGL = the date (if empty), plus a
  "Berhasil" CGL entry in **LDP Progress**. An existing CGL or higher is never demoted.
- A location goes into **Lokasi CG**. Every multiplikasi is recorded as one row in the **Multiplikasi CG** tab
  (created automatically): TANGGAL, CG ASAL, CG BARU, COACH, LEADER, JUMLAH PINDAH, NAMA, OLEH, DICATAT PADA.
- Coaches can only use their own people and their own coach group. Anything that crosses coaches needs an admin
  account. Everything is checked before the first cell is written.

## Lokasi CG tab

One row per CG: `NO CG | COACH | NAMA TEMPAT | ALAMAT | LINK MAPS | LATITUDE | LONGITUDE | DIPERBARUI OLEH | DIPERBARUI`.
You only need an address **or** a Google Maps link; coordinates are filled in automatically:

- **Google Maps link** (best): in Google Maps, long-press the place → *Share* → *Copy link* (`https://maps.app.goo.gl/…`).
  The app reads the pin's coordinates from the link (short links are expanded once and cached).
- **Address**: geocoded with Apps Script's built-in Maps service (no API key, no billing), biased to Greater Jakarta,
  and cached in Script Properties, so each address is looked up only once.
- **Pin**: in the app you can also click the map or drag the pin for an exact spot.

A CG whose `STATUS CG` is *Online* and has no place is listed under *CG online* instead of on the map.
The map tiles come from CARTO (© OpenStreetMap contributors), free for this kind of low-volume use.

## LDP columns and the LDP Progress tab

The database tab has a target and a realization column for each step (dates written like `15 June 2026`):

| Step | Target | Tercapai |
| --- | --- | --- |
| Member | TARGET JD MEMBER | TGL JD MEMBER |
| Sponsor | TARGET JD SPONSOR | TGL JD SPONSOR |
| CGL | TARGET JD CGL | TGL JD CGL |
| Coach | TARGET JD COACH | TGL JD COACH |
| Team Leader | TARGET JD TEAM LEADER | TGL JD TEAM LEADER |

A person's *current step* is the first one not reached yet (a step counts as reached when its TGL is filled,
or POSISI is already at or above it). Older month-only values such as `June 2025` still work and count as the
end of that month.

**LDP Progress** gets one row per monthly update: `TANGGAL UPDATE | COACH | NAMA | NIJ | TAHAP | TARGET | HASIL
(Berhasil/Belum) | TGL TERCAPAI | TARGET BARU | CATATAN | OLEH | DICATAT PADA`. Coaches update their own people;
admin updates everyone.

## Log Pertemuan tab

One row per task; a meeting with several action items has several rows with the same `MEETING ID`
(a meeting without tasks is one row with an empty `TO DO`).

| MEETING ID | TANGGAL | COACH | DENGAN | KATEGORI | TUJUAN | NO | TO DO | DEADLINE | SELESAI | DIBUAT OLEH | DIBUAT PADA | DIPERBARUI |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| M2611241030… | 2026-11-24 | Angie | Sherly Virginia; Clementine Liezel Susanto | Persiapan naik jenjang | Persiapan naik jadi CGL… | 1 | Daftar MSJ 3 | 2026-12-10 | ✔ | angie | 2026-11-24 10:30 | 2026-11-26 19:02 |

- Dates are text in `yyyy-mm-dd`. `DENGAN` holds names exactly as in the database, separated by `; `.
- A coach can only pick their own people and only sees/changes meetings whose `COACH` is theirs; admin sees all.
  All people in one meeting must belong to the same coach.
- Purpose categories are listed in `MEETING.CATEGORIES` in `Meetings.gs` (also the dropdown in the tab).
- The tab is created automatically on the first save if it doesn't exist.

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
# Offline? LEAFLET_DIR=path/to/leaflet/dist node dev/build-preview.js   (serves Leaflet from disk)
node dev/build-preview.js export.csv # CSV of the database tab (stays local, git-ignored)
# open dev/preview.html in a browser. Edits also work there (saved to a fake in-memory sheet)

node dev/test.js                     # tests for the parser and updatePerson
```
