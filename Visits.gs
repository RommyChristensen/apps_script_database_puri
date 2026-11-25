/**
 * Visit CG: coaches, TLs and admins plan visits to CGs and log how they went.
 *
 * Stored in the "Visit CG" tab (CONFIG.VISIT_SHEET_NAME), one row per visit:
 *
 *   ID | CG | COACH | CGL | TANGGAL | JAM | PENGUNJUNG | PERAN | RENCANA |
 *   STATUS | TGL VISIT | HADIR | EVALUASI | ACTION CGL | DIBUAT OLEH |
 *   DIBUAT PADA | DIPERBARUI
 *
 * - A schedule is a row with STATUS "Terjadwal" (TANGGAL / JAM / RENCANA).
 * - Filling the log sets STATUS "Selesai" plus TGL VISIT, HADIR, EVALUASI and
 *   ACTION CGL (the follow-up for that CG's CGL). A visit that wasn't planned
 *   can be logged directly (a new row that is "Selesai" straight away).
 * - Who may visit a CG: a coach their own CGs, a TL every CG of the coaches
 *   under them (TL column of the database), admin every CG.
 * - Everyone sees every visit to the CGs they can access (a coach also sees
 *   the TL's visits to their CGs), but only the author or an admin can change
 *   or delete a visit.
 */

const VISIT = {
  HEADERS: ['ID', 'CG', 'COACH', 'CGL', 'TANGGAL', 'JAM', 'PENGUNJUNG', 'PERAN', 'RENCANA', 'STATUS',
    'TGL VISIT', 'HADIR', 'EVALUASI', 'ACTION CGL', 'DIBUAT OLEH', 'DIBUAT PADA', 'DIPERBARUI'],
  PLANNED: 'Terjadwal',
  DONE: 'Selesai',
};

/** Called from the browser: the visits this session may see, newest first. */
function getVisits(token) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  return visitsPayload_(ss, session);
}

/** Creates (no id) or changes (with id) a planned visit. v = { id?, cg, date, time, plan } */
function scheduleVisit(token, v) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  if (!v) throw new Error('Data visit tidak valid.');
  const date = toIsoDate_(v.date);
  if (!date) throw new Error('Isi tanggal visit.');
  const time = visitText_(v.time, 40);
  const plan = visitText_(v.plan, 1000);

  return withScriptLock_(function () {
    const cg = visitCg_(ss, session, v.cg);
    const sheet = getVisitSheet_(ss);
    const rows = readVisitRows_(sheet);
    const now = visitNow_();
    let row;
    if (v.id) {
      row = visitById_(rows, v.id);
      requireVisitWrite_(session, row);
      if (row.status === VISIT.DONE) throw new Error('Visit ini sudah ada lognya. Ubah lewat log-nya.');
    } else {
      row = newVisitRow_(session, now);
    }
    Object.assign(row, { cg: cg.code, coach: cg.coach, cgl: cg.cgl, date: date, time: time, plan: plan,
      status: VISIT.PLANNED, updated: now });
    writeVisitRow_(sheet, row);
    SpreadsheetApp.flush();
    return visitsPayload_(ss, session);
  });
}

/**
 * Fills the log of a planned visit (with id) or records an unplanned one (no id).
 * v = { id?, cg, date, attendance, evaluation, actions }
 */
function logVisit(token, v) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  if (!v) throw new Error('Data visit tidak valid.');
  const date = toIsoDate_(v.date);
  if (!date) throw new Error('Isi tanggal visit.');
  const today = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd');
  if (date > today) throw new Error('Tanggal visit belum lewat. Untuk rencana, pakai "Jadwalkan visit".');
  const evaluation = visitText_(v.evaluation, 3000);
  if (evaluation.length < 3) throw new Error('Tulis evaluasi visit.');
  const actions = visitText_(v.actions, 2000);
  const att = String(v.attendance == null ? '' : v.attendance).trim();
  if (att && !/^\d{1,3}$/.test(att)) throw new Error('Jumlah hadir harus angka.');

  return withScriptLock_(function () {
    const sheet = getVisitSheet_(ss);
    const rows = readVisitRows_(sheet);
    const now = visitNow_();
    let row;
    if (v.id) {
      row = visitById_(rows, v.id);
      requireVisitWrite_(session, row);
      visitCg_(ss, session, row.cg, true); // still allowed to visit that CG (CG itself may be gone)
    } else {
      const cg = visitCg_(ss, session, v.cg);
      row = newVisitRow_(session, now);
      Object.assign(row, { cg: cg.code, coach: cg.coach, cgl: cg.cgl, date: date });
    }
    Object.assign(row, { status: VISIT.DONE, visitDate: date, attendance: att, evaluation: evaluation,
      actions: actions, updated: now });
    writeVisitRow_(sheet, row);
    SpreadsheetApp.flush();
    return visitsPayload_(ss, session);
  });
}

/** Deletes a visit (its author or an admin). */
function deleteVisit(token, id) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  return withScriptLock_(function () {
    const sheet = getVisitSheet_(ss);
    const row = visitById_(readVisitRows_(sheet), id);
    requireVisitWrite_(session, row);
    sheet.deleteRow(row.row);
    SpreadsheetApp.flush();
    return visitsPayload_(ss, session);
  });
}

// ───────────────────────────── helpers ─────────────────────────────

const VISIT_KEYS_ = ['id', 'cg', 'coach', 'cgl', 'date', 'time', 'visitor', 'role', 'plan', 'status',
  'visitDate', 'attendance', 'evaluation', 'actions', 'by', 'at', 'updated'];

function visitText_(v, max) {
  return String(v == null ? '' : v).replace(/\r\n?/g, '\n').replace(/[ \t]+/g, ' ').trim().slice(0, max);
}

function visitNow_() {
  return Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd HH:mm');
}

/** The CG as it is in the database now: { code, coach, cgl }; checks this session may visit it. */
function visitCg_(ss, session, code, allowMissing) {
  code = cgCode_(code);
  if (!code) throw new Error('Pilih CG.');
  const people = parseDatabase_(getDatabaseSheet_(ss).getDataRange().getDisplayValues()).people;
  const members = people.filter(function (p) { return p.cg === code && !p.isCoach; });
  if (!members.length) {
    if (allowMissing) return { code: code, coach: '', cgl: '' };
    throw new Error('CG "' + code + '" tidak ditemukan.');
  }
  if (!canAccessCoach_(session, members[0].coach)) throw new Error('Kamu hanya bisa visit CG di bawah kamu.');
  const cgl = members.filter(function (p) { return p.position === 'CGL'; })[0];
  return { code: code, coach: members[0].coach, cgl: cgl ? cgl.name : cgInfoOf_(members, 'cgl') };
}

function newVisitRow_(session, now) {
  return {
    id: 'V' + Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyMMddHHmmss') +
      Utilities.getUuid().replace(/-/g, '').slice(0, 4).toUpperCase(),
    visitor: session.role === 'ADMIN' ? session.username : session.coach,
    role: roleLabel_(session), plan: '', time: '', visitDate: '', attendance: '', evaluation: '', actions: '',
    by: session.username, at: now,
  };
}

function requireVisitWrite_(session, row) {
  if (session.role === 'ADMIN') return;
  if (row.by !== session.username) throw new Error('Hanya yang membuat visit ini (atau admin) yang bisa mengubahnya.');
}

function visitById_(rows, id) {
  const row = rows.filter(function (r) { return r.id === String(id || ''); })[0];
  if (!row) throw new Error('Visit ini sudah tidak ada. Refresh dulu.');
  return row;
}

function getVisitSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.VISIT_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.VISIT_SHEET_NAME, ss.getSheets().length);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(VISIT.HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, VISIT.HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

/** All visit rows as objects (with .row = sheet row number). */
function readVisitRows_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getDisplayValues();
  const header = values[0].map(norm_);
  const c = VISIT.HEADERS.map(function (h) { return header.indexOf(h); });
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = { row: i + 1 };
    VISIT_KEYS_.forEach(function (k, j) { r[k] = c[j] < 0 ? '' : String(values[i][c[j]] == null ? '' : values[i][c[j]]).trim(); });
    if (!r.id) continue;
    r.date = toIsoDate_(r.date) || r.date;
    r.visitDate = toIsoDate_(r.visitDate) || r.visitDate;
    out.push(r);
  }
  return out;
}

function writeVisitRow_(sheet, r) {
  const values = [VISIT_KEYS_.map(function (k) { return String(r[k] == null ? '' : r[k]); })];
  const range = sheet.getRange(r.row || sheet.getLastRow() + 1, 1, 1, VISIT_KEYS_.length);
  range.setNumberFormat('@');
  range.setValues(values);
}

function visitsPayload_(ss, session) {
  const sheet = ss.getSheetByName(CONFIG.VISIT_SHEET_NAME);
  const visits = readVisitRows_(sheet).filter(function (r) { return canAccessCoach_(session, r.coach); })
    .map(function (r) {
      const out = {};
      VISIT_KEYS_.forEach(function (k) { out[k] = r[k]; });
      out.attendance = r.attendance === '' ? null : Number(r.attendance);
      out.canEdit = session.role === 'ADMIN' || r.by === session.username;
      return out;
    });
  visits.sort(function (a, b) {
    return ((b.visitDate || b.date) + b.at).localeCompare((a.visitDate || a.date) + a.at);
  });
  return { visits: visits, today: Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd') };
}
