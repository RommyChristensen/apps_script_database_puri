/**
 * Restructuring a CG: change a CG's info for all its members, move a person
 * to another CG, and delete a person (archived first).
 *
 * - updateCgInfo: WAKTU / STATUS / DAERAH CG are written to every row of the
 *   CG (the app still displays the CGL's row). Coaches: own CGs; admin: all.
 * - movePerson ("Pindah CG"): sets NO CG, CGL, COACH and the CG info of one
 *   person to those of the target CG. Coaches: within their own CGs.
 * - deletePerson: admin only, re-typing their own password. The row is copied
 *   to the "Arsip Jemaat" tab (CONFIG.ARCHIVE_SHEET_NAME) and then deleted.
 */

const CG_INFO_FIELDS = ['cgTime', 'cgStatus', 'cgArea'];

const ARCHIVE_HEADERS = ['DIARSIPKAN PADA', 'OLEH', 'ALASAN', 'BARIS ASAL', 'NAMA', 'NO CG', 'COACH', 'POSISI',
  'NIJ', 'NO HP', 'DATA LENGKAP'];

/** Changes a CG's schedule / status / area on every row of that CG. info = { cg, cgTime?, cgStatus?, cgArea? } */
function updateCgInfo(token, info) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  const code = cgCode_(info && info.cg);
  if (!code) throw new Error('Pilih CG.');
  return withScriptLock_(function () {
    const db = readDatabase_(ss);
    const members = db.people.filter(function (p) { return p.cg === code && !p.isCoach; });
    if (!members.length) throw new Error('CG "' + code + '" tidak ditemukan.');
    if (!canAccessCoach_(session, members[0].coach)) throw new Error('Kamu hanya bisa mengubah CG milik coach kamu.');

    const want = {};
    CG_INFO_FIELDS.forEach(function (k) {
      if (!Object.prototype.hasOwnProperty.call(info, k)) return;
      const f = fieldDef_(k);
      const v = String(info[k] == null ? '' : info[k]).replace(/\s+/g, ' ').trim().slice(0, 100);
      validateField_(f, v);
      want[k] = v;
    });

    // What the CG showed before (CGL's row) → what it shows now; then write every member's row.
    const summary = CG_INFO_FIELDS.filter(function (k) { return k in want && want[k] !== cgInfoOf_(members, k); })
      .map(function (k) { return { label: fieldDef_(k).label, from: cgInfoOf_(members, k), to: want[k] }; });
    members.forEach(function (p) { setRowFields_(db, p.id, want); });
    SpreadsheetApp.flush();
    if (summary.length) {
      const cgl = members.filter(function (p) { return p.position === 'CGL'; })[0] || members[0];
      logEdit_(ss, cgl.id, code + ' (' + members.length + ' orang)', summary, session);
    }
    return { changes: summary, data: buildDashboard_(ss, session) };
  });
}

/** Moves one person to another CG. req = { id, name, cg } */
function movePerson(token, req) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  const code = cgCode_(req && req.cg);
  if (!req || !req.id || !req.name || !code) throw new Error('Data pindah CG tidak valid.');
  return withScriptLock_(function () {
    const db = readDatabase_(ss);
    const person = personAt_(db, req.id, req.name);
    if (person.isCoach) throw new Error('Coach / Team Leader tidak bisa dipindah CG.');
    const target = db.people.filter(function (p) { return p.cg === code && !p.isCoach; });
    if (!target.length) throw new Error('CG "' + code + '" tidak ditemukan.');
    if (person.cg === code) throw new Error(person.name + ' sudah ada di ' + code + '.');
    if (!canAccessCoach_(session, person.coach) || !canAccessCoach_(session, target[0].coach)) {
      throw new Error('Kamu hanya bisa memindahkan orang antar CG milik coach kamu.');
    }
    if (target[0].coach !== person.coach && db.col.coach === undefined) {
      throw new Error('Pindah ke CG coach lain butuh kolom COACH di database.');
    }
    const cgl = target.filter(function (p) { return p.position === 'CGL'; })[0];
    const want = { cg: code, cgl: cgl ? cgl.name : '', coach: target[0].coach };
    CG_INFO_FIELDS.forEach(function (k) { want[k] = cgInfoOf_(target, k); });
    const changes = setRowFields_(db, person.id, want);
    SpreadsheetApp.flush();
    if (changes.length) logEdit_(ss, person.id, person.name, changes, session);
    return { changes: changes, data: buildDashboard_(ss, session) };
  });
}

/** Admin only: archives the row, then deletes it. req = { id, name, reason, password } */
function deletePerson(token, req) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  if (session.role !== 'ADMIN') throw new Error('Hanya admin yang bisa menghapus data.');
  if (!req || !req.id || !req.name) throw new Error('Data hapus tidak valid.');
  const reason = String(req.reason || '').replace(/\s+/g, ' ').trim().slice(0, 500);
  if (reason.length < 3) throw new Error('Tulis alasan penghapusan.');
  confirmPassword_(ss, session, req.password);

  return withScriptLock_(function () {
    const db = readDatabase_(ss);
    const person = personAt_(db, req.id, req.name);
    const row = db.display[person.id - 1];
    const header = db.header;
    // Every non-empty cell, as "HEADER: value" (milestones use their sub-header, e.g. "MSJ 1: -").
    const details = header.map(function (h, j) {
      const v = String(row[j] == null ? '' : row[j]).replace(/\u00a0/g, ' ').trim();
      return h && v ? h + ': ' + v : '';
    }).filter(Boolean).join('\n');

    const archive = getArchiveSheet_(ss);
    const now = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd HH:mm');
    const values = [[now, session.username, reason, String(person.id), person.name, person.cg, person.coach,
      person.position, person.nij, person.phone, details]];
    const range = archive.getRange(archive.getLastRow() + 1, 1, 1, ARCHIVE_HEADERS.length);
    range.setNumberFormat('@');
    range.setValues(values);
    SpreadsheetApp.flush();

    // Only delete once the archive row is really there.
    const check = archive.getRange(archive.getLastRow(), 1, 1, 5).getDisplayValues()[0];
    if (check[4] !== person.name) throw new Error('Arsip gagal disimpan, data tidak dihapus.');
    db.sheet.deleteRow(person.id);
    SpreadsheetApp.flush();
    logEdit_(ss, person.id, person.name, [{ label: 'Dihapus (diarsipkan)', from: person.cg || '', to: reason }], session);
    return { deleted: person.name, data: buildDashboard_(ss, session) };
  });
}

// ───────────────────────────── helpers ─────────────────────────────

function cgCode_(v) {
  return String(v || '').toUpperCase().replace(/\s+/g, ' ').trim();
}

function fieldDef_(key) {
  return EDITABLE_FIELDS.filter(function (f) { return f.key === key; })[0];
}

/** The whole database tab, parsed once: { sheet, display, formulas, header, col, people }. */
function readDatabase_(ss) {
  const sheet = getDatabaseSheet_(ss);
  const range = sheet.getDataRange();
  const display = range.getDisplayValues();
  const map = mapColumns_(display);
  const header = display[map.headerRow].map(function (h, j) {
    const sub = map.hasSub ? clean_(display[map.headerRow + 1][j]) : '';
    return sub || clean_(h);
  });
  return {
    sheet: sheet, display: display, formulas: range.getFormulas(), header: header,
    col: map.col, people: parseDatabase_(display).people,
  };
}

function personAt_(db, id, name) {
  const p = db.people.filter(function (x) { return x.id === Math.floor(Number(id)); })[0];
  if (!p || norm_(p.name) !== norm_(name)) {
    throw new Error('Data di baris ' + id + ' sudah berubah. Refresh dulu lalu coba lagi.');
  }
  return p;
}

/** Same rule as the app: the CGL's value, else the most common one among the members. */
function cgInfoOf_(members, key) {
  const cgl = members.filter(function (p) { return p.position === 'CGL' && p[key]; })[0];
  if (cgl) return cgl[key];
  const counts = {};
  let best = '', n = 0;
  members.forEach(function (p) {
    if (!p[key]) return;
    counts[p[key]] = (counts[p[key]] || 0) + 1;
    if (counts[p[key]] > n) { n = counts[p[key]]; best = p[key]; }
  });
  return best;
}

const ROW_FIELD_LABELS_ = { cg: 'NO CG', cgl: 'CGL', coach: 'COACH', cgTime: 'Waktu CG', cgStatus: 'Status CG', cgArea: 'Daerah CG' };

/**
 * Writes { field: value } to one row (only cells that change, never over a
 * formula, as plain text). Returns the changes; keeps db.display in step.
 */
function setRowFields_(db, row, want) {
  const changes = [];
  Object.keys(want).forEach(function (k) {
    const j = db.col[k];
    if (j === undefined) return;
    const had = clean_(db.display[row - 1][j]);
    const v = want[k];
    if (had === v) return;
    if (db.formulas[row - 1] && db.formulas[row - 1][j]) {
      throw new Error('Kolom "' + (ROW_FIELD_LABELS_[k] || k) + '" di baris ' + row + ' berisi formula.');
    }
    const cell = db.sheet.getRange(row, j + 1);
    if (v !== '') cell.setNumberFormat('@');
    cell.setValue(v);
    db.display[row - 1][j] = v;
    changes.push({ label: ROW_FIELD_LABELS_[k] || k, from: had, to: v });
  });
  return changes;
}

function getArchiveSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.ARCHIVE_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.ARCHIVE_SHEET_NAME, ss.getSheets().length);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(ARCHIVE_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, ARCHIVE_HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

/** Re-checks the signed-in admin's own password (counts towards the login lock). */
function confirmPassword_(ss, session, password) {
  const cache = CacheService.getScriptCache();
  const failKey = 'fail:' + session.username;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= AUTH.MAX_FAILED) throw new Error('Terlalu banyak percobaan. Coba lagi dalam ' + AUTH.LOCK_MINUTES + ' menit.');
  const acc = readAccounts_(ss).filter(function (a) { return a.username === session.username; })[0];
  if (!acc || !passwordMatches_(String(password || ''), acc.stored)) {
    cache.put(failKey, String(fails + 1), AUTH.LOCK_MINUTES * 60);
    throw new Error('Password salah. Data tidak dihapus.');
  }
  cache.remove(failKey);
}
