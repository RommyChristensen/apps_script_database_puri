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
};

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

/** Called from the browser via google.script.run. */
function getDashboardData() {
  const ss = CONFIG.SPREADSHEET_ID
    ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  const sheet = CONFIG.SHEET_NAME ? ss.getSheetByName(CONFIG.SHEET_NAME) : ss.getSheets()[0];
  if (!sheet) throw new Error('Tab "' + CONFIG.SHEET_NAME + '" not found.');

  const parsed = parseDatabase_(sheet.getDataRange().getDisplayValues());
  return {
    generatedAt: new Date().toISOString(),
    sheetName: sheet.getName(),
    spreadsheetUrl: ss.getUrl() + '#gid=' + sheet.getSheetId(),
    milestones: MILESTONES.map(function (m) {
      return { key: m.key, label: m.label, group: m.group };
    }),
    coaches: parsed.coaches,
    people: parsed.people,
    warnings: parsed.warnings,
  };
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
    if ((!notes || /^\d{1,3}$/.test(notes) || /^january 1900$/i.test(notes)) && looksLikeNote_(targetCoach)) {
      notes = targetCoach;
      targetCoach = '';
    } else if (/^\d{1,3}$/.test(notes) || /^january 1900$/i.test(notes)) {
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
