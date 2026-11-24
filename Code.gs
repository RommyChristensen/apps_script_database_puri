/**
 * AOG Youth Puri — Coach Dashboard (Apps Script Web App)
 *
 * Reads the merged DATABASE tab (first tab by default) and serves a
 * read-only dashboard grouped by coach → CG → people.
 */

const CONFIG = {
  // Spreadsheet that holds the database. Leave blank to use the spreadsheet
  // this script is bound to (Extensions → Apps Script).
  SPREADSHEET_ID: '1VCzzxXv2qXvnoK5Fnrx5WxpD8owbIt0bnqzPdhsEZl0',
  // Tab to read. Leave blank to always read the FIRST tab.
  SHEET_NAME: '',
  APP_TITLE: 'AOG Youth Puri',
  // Every edit made from the dashboard is appended to this tab (created
  // automatically as the LAST tab). Leave blank to disable the log.
  LOG_SHEET_NAME: 'Edit Log',
};

/** Values allowed in the POSISI column when editing. */
const POSITION_OPTIONS = ['VIP', 'SIMPATISAN', 'MEMBER', 'SPONSOR', 'CGL'];

/** Free-text "Jenjang" fields that can be edited from the dashboard. */
const EDITABLE_TEXT_FIELDS = [
  { key: 'targetMember', label: 'Target jadi Member' },
  { key: 'dateMember', label: 'Tgl jadi Member' },
  { key: 'targetSponsor', label: 'Target jadi Sponsor' },
  { key: 'dateSponsor', label: 'Tgl jadi Sponsor' },
  { key: 'targetCgl', label: 'Target CGL' },
  { key: 'targetCoach', label: 'Target jadi Coach' },
];

/** Growth milestones, in the order they appear under 5M / MSJ / CGT. */
const MILESTONES = [
  { key: 'percayaYesus', label: 'Percaya Yesus', group: '5M', header: 'PERCAYA YESUS' },
  { key: 'baptisSelam', label: 'Baptis Selam', group: '5M', header: 'BAPTIS SELAM' },
  { key: 'baptisRk', label: 'Baptis RK', group: '5M', header: 'BAPTIS RK' },
  { key: 'rutinCg', label: 'Rutin CG', group: '5M', header: 'RUTIN CG' },
  { key: 'rutinIbadah', label: 'Rutin Ibadah', group: '5M', header: 'RUTIN IBADAH' },
  { key: 'msj1', label: 'MSJ 1', group: 'MSJ', header: 'MSJ 1' },
  { key: 'msj2', label: 'MSJ 2', group: 'MSJ', header: 'MSJ 2' },
  { key: 'msj3', label: 'MSJ 3', group: 'MSJ', header: 'MSJ 3' },
  { key: 'cgt1', label: 'CGT 1', group: 'CGT', header: 'CGT 1' },
  { key: 'cgt2', label: 'CGT 2', group: 'CGT', header: 'CGT 2' },
  { key: 'cgt3', label: 'CGT 3', group: 'CGT', header: 'CGT 3' },
];

/** Field → accepted header names (first match wins, so order matters). */
const FIELDS = {
  coach: ['COACH'],
  cg: ['NO CG'],
  name: ['NAMA', 'NAME'],
  cgStatus: ['STATUS CG'],
  cgTime: ['WAKTU CG'],
  cgArea: ['DAERAH CG'],
  domicile: ['DOMISILI (KECAMATAN)', 'DOMISILI'],
  email: ['EMAIL'],
  phone: ['NO HP'],
  cgl: ['CGL'],
  gender: ['GENDER'],
  nij: ['NIJ'],
  position: ['POSISI'],
  status: ['STATUS'],
  birthCity: ['KOTA KELAHIRAN'],
  birthDate: ['TTL'],
  address: ['ALAMAT'],
  occupation: ['KERJA/KULIAH', 'SEKOLAH/KULIAH'],
  school: ['KAMPUS', 'SEKOLAH'],
  ministry: ['PELAYANAN'],
  targetMember: ['TARGET JD MEMBER'],
  dateMember: ['TGL JD MEMBER'],
  targetSponsor: ['TARGET JD SPONSOR'],
  dateSponsor: ['TGL JD SPONSOR'],
  targetCgl: ['TARGET CGL'],
  targetCoach: ['TARGET JD COACH'],
  notes: ['KETERANGAN', 'NOTES'],
  age: ['AGE', 'UMUR'],
  eligibleTl: ['ELIGIBLE TL'],
  eligibleCoach: ['ELIGIBLE COACH'],
  eligibleCgl: ['ELIGIBLE CGL'],
};

// ───────────────────────────── Web app ─────────────────────────────

function doGet() {
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.appTitle = CONFIG.APP_TITLE;
  return tpl
    .evaluate()
    .setTitle(CONFIG.APP_TITLE + ' · Coach Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/** Inline another project file (used by Index.html for CSS/JS). */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSpreadsheet_() {
  return CONFIG.SPREADSHEET_ID
    ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function getDatabaseSheet_(ss) {
  const sheet = CONFIG.SHEET_NAME ? ss.getSheetByName(CONFIG.SHEET_NAME) : ss.getSheets()[0];
  if (!sheet) throw new Error('Tab "' + CONFIG.SHEET_NAME + '" not found.');
  return sheet;
}

/** Called from the browser via google.script.run. */
function getDashboardData() {
  const ss = getSpreadsheet_();
  const sheet = getDatabaseSheet_(ss);

  const parsed = parseDatabase_(sheet.getDataRange().getDisplayValues());
  return {
    generatedAt: new Date().toISOString(),
    sheetName: sheet.getName(),
    spreadsheetUrl: ss.getUrl() + '#gid=' + sheet.getSheetId(),
    milestones: MILESTONES.map(function (m) {
      return { key: m.key, label: m.label, group: m.group };
    }),
    positionOptions: POSITION_OPTIONS,
    editableFields: EDITABLE_TEXT_FIELDS,
    coaches: parsed.coaches,
    people: parsed.people,
    warnings: parsed.warnings,
  };
}

// ───────────────────────────── Editing ─────────────────────────────

/**
 * Saves edits for one person (called from the browser).
 *
 * payload = {
 *   id: <sheet row number>, name: <name shown in the dashboard>,
 *   milestones: { percayaYesus: true, msj1: false, ... },   // optional
 *   position: 'MEMBER',                                       // optional
 *   targetMember, dateMember, targetSponsor, dateSponsor, targetCgl, targetCoach  // optional
 * }
 *
 * Only cells whose value actually changed are written. Returns the refreshed
 * dashboard data plus a list of what changed.
 */
function updatePerson(payload) {
  if (!payload || !payload.id || !payload.name) throw new Error('Data edit tidak valid.');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Sedang ada yang menyimpan. Coba lagi sebentar.');
  try {
    const ss = getSpreadsheet_();
    const sheet = getDatabaseSheet_(ss);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const map = mapColumns_(sheet.getRange(1, 1, Math.min(25, lastRow), lastCol).getDisplayValues());
    const col = map.col;

    const row = Math.floor(Number(payload.id));
    const firstDataRow = map.headerRow + (map.hasSub ? 2 : 1) + 1; // 1-based
    if (!(row >= firstDataRow && row <= lastRow)) throw new Error('Baris ' + row + ' di luar area data.');

    const range = sheet.getRange(row, 1, 1, lastCol);
    const display = range.getDisplayValues()[0];
    const raw = range.getValues()[0];
    const formulas = range.getFormulas()[0];

    // Guard against rows that moved since the dashboard was loaded.
    if (norm_(display[col.name]) !== norm_(payload.name)) {
      throw new Error('Data di baris ' + row + ' sudah berubah (nama tidak cocok). Refresh dulu lalu coba lagi.');
    }

    const changes = [];
    const queue = function (j, value, label, from, to, asText) {
      if (formulas[j]) throw new Error('Kolom "' + label + '" berisi formula, tidak bisa diedit dari dashboard.');
      changes.push({ col: j, value: value, label: label, from: from, to: to, asText: asText });
    };

    // 5M / MSJ / CGT
    const ms = payload.milestones || {};
    MILESTONES.forEach(function (m) {
      if (!Object.prototype.hasOwnProperty.call(ms, m.key) || col[m.key] === undefined) return;
      const j = col[m.key];
      const want = ms[m.key] === true;
      const had = isChecked_(display[j]);
      if (want === had) return;
      // Keep the cell's style: real checkboxes stay booleans, text cells use ✔ / -.
      const value = typeof raw[j] === 'boolean' ? want : (want ? '✔' : '-');
      queue(j, value, m.label, had ? '✔' : '✗', want ? '✔' : '✗', false);
    });

    // POSISI
    if (Object.prototype.hasOwnProperty.call(payload, 'position') && col.position !== undefined) {
      const want = String(payload.position || '').trim().toUpperCase();
      if (want && POSITION_OPTIONS.indexOf(want) < 0) throw new Error('Posisi "' + want + '" tidak dikenal.');
      const had = clean_(display[col.position]).toUpperCase();
      if (want && want !== had) queue(col.position, want, 'Posisi', had, want, true);
    }

    // Jenjang text fields
    const notesInTargetCoach = col.targetCoach !== undefined &&
      targetCoachHoldsNotes_(col.notes === undefined ? '' : clean_(display[col.notes]), clean_(display[col.targetCoach]));
    EDITABLE_TEXT_FIELDS.forEach(function (f) {
      if (!Object.prototype.hasOwnProperty.call(payload, f.key) || col[f.key] === undefined) return;
      if (f.key === 'targetCoach' && notesInTargetCoach) {
        if (String(payload.targetCoach || '').trim()) {
          throw new Error('Kolom TARGET JD COACH di baris ini masih berisi catatan. Rapikan header di sheet dulu.');
        }
        return; // never wipe the misplaced notes
      }
      const want = String(payload[f.key] == null ? '' : payload[f.key]).replace(/\s+/g, ' ').trim().slice(0, 100);
      const had = clean_(display[col[f.key]]);
      if (want === had) return;
      queue(col[f.key], want, f.label, had, want, true);
    });

    changes.forEach(function (c) {
      const cell = sheet.getRange(row, c.col + 1);
      // Store as plain text so "Juni 2026" isn't turned into a date by Sheets.
      if (c.asText && c.value !== '') cell.setNumberFormat('@');
      cell.setValue(c.value);
    });
    SpreadsheetApp.flush();

    if (changes.length) logEdit_(ss, row, display[col.name], changes);

    return {
      changes: changes.map(function (c) { return { label: c.label, from: c.from, to: c.to }; }),
      data: getDashboardData(),
    };
  } finally {
    lock.releaseLock();
  }
}

function logEdit_(ss, row, name, changes) {
  if (!CONFIG.LOG_SHEET_NAME) return;
  try {
    let log = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
    if (!log) {
      log = ss.insertSheet(CONFIG.LOG_SHEET_NAME, ss.getSheets().length); // keep the database tab first
      log.appendRow(['Waktu', 'Oleh', 'Baris', 'Nama', 'Perubahan']);
      log.setFrozenRows(1);
      log.getRange('A1:E1').setFontWeight('bold');
    }
    let who = '';
    try { who = Session.getActiveUser().getEmail(); } catch (e) { /* not available */ }
    log.appendRow([
      new Date(), who || '(tidak diketahui)', row, name,
      changes.map(function (c) { return c.label + ': ' + (c.from || '—') + ' → ' + (c.to || '—'); }).join('\n'),
    ]);
  } catch (e) {
    console.warn('Edit log failed: ' + e);
  }
}

// ───────────────────────────── Parsing ─────────────────────────────

function norm_(v) {
  return String(v == null ? '' : v)
    .replace(/ /g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function clean_(v) {
  const s = String(v == null ? '' : v).replace(/ /g, ' ').replace(/\s+/g, ' ').trim();
  return s === '-' || s === '–' ? '' : s;
}

function titleCase_(s) {
  return String(s)
    .toLowerCase()
    .replace(/(^|[\s\-'])\S/g, function (c) { return c.toUpperCase(); });
}

/** "COACH ANGIE" → "Angie"; returns '' when the text is not a coach banner. */
function coachFromBanner_(text) {
  const m = String(text || '').trim().match(/^COACH\s*[:\-]?\s+(.+)$/i);
  return m ? titleCase_(m[1].trim()) : '';
}

function isChecked_(v) {
  const s = norm_(v);
  return s === '✔' || s === '✓' || s === '✅' || s === 'V' || s === 'TRUE' || s === 'YES' || s === 'Y' || s === 'SUDAH';
}

/**
 * Parses the sheet's display values into people + coaches.
 * Supports two ways of marking the coach after merging:
 *   1. A "COACH" column (recommended), and/or
 *   2. Banner rows such as "COACH ANGIE" above each coach's block.
 */
function parseDatabase_(values) {
  const map = mapColumns_(values);
  const warnings = map.warnings;
  const headerRow = map.headerRow;
  const hasSub = map.hasSub;
  const col = map.col;
  const extraPhoneCols = map.extraPhoneCols;

  const get = function (row, field) {
    return col[field] === undefined ? '' : clean_(row[col[field]]);
  };

  // A coach banner above the header (e.g. row 1 "COACH ANGIE") seeds the first block.
  let currentCoach = '';
  for (let i = 0; i < headerRow; i++) {
    for (let j = 0; j < values[i].length; j++) {
      const c = coachFromBanner_(values[i][j]);
      if (c) currentCoach = c;
    }
  }

  const people = [];
  const coachOrder = [];
  const coachFullName = {};
  const seenCoach = {};
  const start = headerRow + (hasSub ? 2 : 1);

  for (let i = start; i < values.length; i++) {
    const row = values[i];
    const name = get(row, 'name');

    // Banner row inside the data (merged blocks): "COACH LENITA".
    if (!name) {
      for (let j = 0; j < Math.min(row.length, 8); j++) {
        const c = coachFromBanner_(row[j]);
        if (c) { currentCoach = c; break; }
      }
      continue;
    }
    if (norm_(name) === 'NAMA' || norm_(name) === 'NAME') continue; // repeated header

    const coach = titleCase_(get(row, 'coach')) || currentCoach || 'Unassigned';
    if (!seenCoach[coach]) { seenCoach[coach] = true; coachOrder.push(coach); }

    const cg = get(row, 'cg').toUpperCase();
    let position = get(row, 'position').toUpperCase();
    const isCoach = position === 'COACH' || (!cg && !position);
    if (isCoach) {
      position = 'COACH';
      if (!coachFullName[coach]) coachFullName[coach] = name;
    }

    let phone = get(row, 'phone');
    for (let k = 0; !phone && k < extraPhoneCols.length; k++) phone = clean_(row[extraPhoneCols[k]]);

    // Notes sometimes sit one column to the left of KETERANGAN; recover them.
    let notes = get(row, 'notes');
    let targetCoach = get(row, 'targetCoach');
    const targetCoachHoldsNotes = targetCoachHoldsNotes_(notes, targetCoach);
    if (targetCoachHoldsNotes) {
      notes = targetCoach;
      targetCoach = '';
    } else if (isNoiseNote_(notes)) {
      notes = '';
    }

    const birthDate = get(row, 'birthDate');
    let age = ageFromDate_(parseDate_(birthDate));
    if (age === null) {
      const a = parseInt(get(row, 'age'), 10);
      age = a > 0 && a < 100 ? a : null;
    }

    const milestones = {};
    MILESTONES.forEach(function (m) {
      milestones[m.key] = col[m.key] === undefined ? false : isChecked_(row[col[m.key]]);
    });

    people.push({
      id: i + 1, // sheet row number
      coach: coach,
      isCoach: isCoach,
      cg: cg,
      name: name,
      cgStatus: get(row, 'cgStatus'),
      cgTime: get(row, 'cgTime'),
      cgArea: get(row, 'cgArea'),
      domicile: get(row, 'domicile'),
      email: get(row, 'email'),
      phone: phone,
      cgl: get(row, 'cgl'),
      gender: get(row, 'gender'),
      nij: get(row, 'nij'),
      position: position || 'UNKNOWN',
      status: get(row, 'status') || 'Unknown',
      birthCity: get(row, 'birthCity'),
      birthDate: birthDate,
      age: age,
      address: get(row, 'address'),
      occupation: get(row, 'occupation'),
      school: get(row, 'school'),
      ministry: get(row, 'ministry'),
      targetMember: get(row, 'targetMember'),
      dateMember: get(row, 'dateMember'),
      targetSponsor: get(row, 'targetSponsor'),
      dateSponsor: get(row, 'dateSponsor'),
      targetCgl: get(row, 'targetCgl'),
      targetCoach: targetCoach,
      targetCoachHoldsNotes: targetCoachHoldsNotes,
      notes: notes,
      eligibleTl: get(row, 'eligibleTl'),
      eligibleCoach: get(row, 'eligibleCoach'),
      eligibleCgl: get(row, 'eligibleCgl'),
      milestones: milestones,
    });
  }

  if (col.coach === undefined && coachOrder.length <= 1) {
    warnings.push('No "COACH" column found — everyone is grouped under "' + (coachOrder[0] || 'Unassigned') +
      '". Add a COACH column or "COACH <name>" banner rows when merging.');
  }

  return {
    people: people,
    coaches: coachOrder.map(function (c) {
      return { key: c, name: c, fullName: coachFullName[c] || '' };
    }),
    warnings: warnings,
  };
}

/**
 * Finds the header row (the one containing "NAMA"), the optional milestone
 * sub-header row below it, and maps every known field to a 0-based column.
 */
function mapColumns_(values) {
  const warnings = [];
  let headerRow = -1;
  for (let i = 0; i < Math.min(values.length, 25); i++) {
    if (values[i].some(function (c) { return norm_(c) === 'NAMA' || norm_(c) === 'NAME'; })) {
      headerRow = i;
      break;
    }
  }
  if (headerRow < 0) throw new Error('Could not find the header row (a cell with "NAMA").');

  const header = values[headerRow].map(norm_);
  const milestoneHeaders = MILESTONES.map(function (m) { return m.header; });
  const maybeSub = (values[headerRow + 1] || []).map(norm_);
  const hasSub = maybeSub.some(function (c) { return milestoneHeaders.indexOf(c) >= 0; });
  const sub = hasSub ? maybeSub : [];

  // Map columns.
  const col = {};
  const extraPhoneCols = [];
  header.forEach(function (h, j) {
    const s = sub[j] || '';
    const mi = milestoneHeaders.indexOf(s);
    if (mi >= 0) { col[MILESTONES[mi].key] = j; return; }
    Object.keys(FIELDS).forEach(function (field) {
      if (FIELDS[field].indexOf(h) < 0) return;
      if (col[field] === undefined) col[field] = j;
      else if (field === 'phone') extraPhoneCols.push(j);
    });
  });
  if (col.name === undefined) throw new Error('Column "NAMA" not found.');
  MILESTONES.forEach(function (m) {
    if (col[m.key] === undefined) warnings.push('Column "' + m.header + '" not found.');
  });

  return { headerRow: headerRow, hasSub: hasSub, col: col, extraPhoneCols: extraPhoneCols, warnings: warnings };
}

function isNoiseNote_(s) {
  return /^\d{1,3}$/.test(s) || /^january 1900$/i.test(s);
}

/** True when the TARGET JD COACH cell actually holds the person's notes (shifted header). */
function targetCoachHoldsNotes_(notes, targetCoach) {
  return (!notes || isNoiseNote_(notes)) && looksLikeNote_(targetCoach);
}

function looksLikeNote_(s) {
  if (!s || s.length < 4 || /^[\d\s\/\-.]+$/.test(s)) return false;
  return !/^[\d\s]*[A-Za-z]+\.?\s*(\/\s*[A-Za-z]+\s*)?\d{4}$/.test(s); // not "Nov 2024" / "Nov / Des 2024"
}

const MONTHS_ = {
  jan: 0, januari: 0, january: 0, feb: 1, februari: 1, february: 1, mar: 2, maret: 2, march: 2,
  apr: 3, april: 3, mei: 4, may: 4, jun: 5, juni: 5, june: 5, jul: 6, juli: 6, july: 6,
  agu: 7, agt: 7, agustus: 7, aug: 7, august: 7, sep: 8, sept: 8, september: 8,
  okt: 9, oktober: 9, oct: 9, october: 9, nov: 10, november: 10, des: 11, desember: 11, dec: 11, december: 11,
};

/** Parses "19 November 2002", "19 Nov 2002", "19/11/2002" or "2002-11-19". */
function parseDate_(s) {
  if (!s) return null;
  let m = s.match(/(\d{1,2})\s+([A-Za-z]+)\.?\s+(\d{4})/);
  if (m && MONTHS_[m[2].toLowerCase()] !== undefined) return new Date(+m[3], MONTHS_[m[2].toLowerCase()], +m[1]);
  m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  return null;
}

function ageFromDate_(d) {
  if (!d || isNaN(d.getTime())) return null;
  const now = new Date();
  let age = now.getFullYear() - d.getFullYear();
  if (now.getMonth() < d.getMonth() || (now.getMonth() === d.getMonth() && now.getDate() < d.getDate())) age--;
  return age > 0 && age < 100 ? age : null;
}
