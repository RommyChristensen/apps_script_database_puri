/**
 * Meeting tracker ("Pertemuan" in the app): coaches log meetings with their CGLs / CG members.
 *
 * Stored in the "Log Pertemuan" tab (CONFIG.MEETING_SHEET_NAME), one row per
 * task. A meeting with three action items is three rows sharing a MEETING ID;
 * a meeting without tasks is a single row with an empty TO DO.
 *
 *   MEETING ID | TANGGAL | COACH | DENGAN | KATEGORI | TUJUAN | NO | TO DO |
 *   DEADLINE | SELESAI | DIBUAT OLEH | DIBUAT PADA | DIPERBARUI
 *
 * Dates are stored as text yyyy-mm-dd. DENGAN holds the people's names
 * separated by "; " (names as written in the database tab).
 *
 * Access: a coach sees and changes only meetings whose COACH is their own and
 * may only pick people of their own coach group; ADMIN sees and edits all.
 */

const MEETING = {
  HEADERS: ['MEETING ID', 'TANGGAL', 'COACH', 'DENGAN', 'KATEGORI', 'TUJUAN', 'NO', 'TO DO',
    'DEADLINE', 'SELESAI', 'DIBUAT OLEH', 'DIBUAT PADA', 'DIPERBARUI'],
  CATEGORIES: ['Pembinaan', 'Follow up', 'Konseling', 'Persiapan naik jenjang', 'Evaluasi CG', 'Lainnya'],
  MAX_PEOPLE: 20,
  MAX_TASKS: 20,
  PEOPLE_SEPARATOR: '; ',
};

/** Called from the browser: the meetings this session may see. */
function getMeetings(token) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  return meetingsPayload_(ss, session);
}

/**
 * Creates (no id) or replaces (with id) a meeting.
 * meeting = { id?, date: 'yyyy-mm-dd', people: [names], category, purpose,
 *             tasks: [{ text, deadline: 'yyyy-mm-dd' | '', done: bool }] }
 */
function saveMeeting(token, meeting) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  const clean = validateMeeting_(meeting);

  // Resolve the people against the database: all must exist and belong to one coach.
  const people = parseDatabase_(getDatabaseSheet_(ss).getDataRange().getDisplayValues()).people;
  const coaches = {};
  const names = clean.people.map(function (name) {
    const hit = people.filter(function (p) {
      return (!p.isCoach || leadsCoaches_(session)) && norm_(p.name) === norm_(name);
    });
    const visible = hit.filter(function (p) { return canAccessCoach_(session, p.coach); });
    if (!visible.length) throw new Error('"' + name + '" tidak ditemukan di data coach kamu.');
    coaches[visible[0].coach] = true;
    return visible[0].name;
  });
  const coachList = Object.keys(coaches);
  if (coachList.length > 1) throw new Error('Pilih orang dari satu coach saja (sekarang: ' + coachList.join(', ') + ').');
  const coach = coachList[0];

  return withMeetingLock_(function () {
    const sheet = getMeetingSheet_(ss, true);
    const rows = readMeetingRows_(sheet);
    const now = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd HH:mm');
    let id = clean.id;
    let createdBy = session.username;
    let createdAt = now;

    if (id) {
      const old = rows.filter(function (r) { return r.id === id; });
      if (!old.length) throw new Error('Pertemuan ini sudah tidak ada. Refresh dulu.');
      requireMeetingWrite_(ss, session, old[0]);
      createdBy = old[0].createdBy || createdBy;
      createdAt = old[0].createdAt || createdAt;
      deleteMeetingRows_(sheet, old);
    } else {
      id = 'M' + Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyMMddHHmmss') +
        Utilities.getUuid().replace(/-/g, '').slice(0, 4).toUpperCase();
    }

    const tasks = clean.tasks.length ? clean.tasks : [{ text: '', deadline: '', done: false }];
    const values = tasks.map(function (t, i) {
      return [id, clean.date, coach, names.join(MEETING.PEOPLE_SEPARATOR), clean.category, clean.purpose,
        t.text ? String(i + 1) : '', t.text, t.deadline, t.text ? (t.done ? '✔' : '-') : '',
        createdBy, createdAt, now];
    });
    const start = sheet.getLastRow() + 1;
    const range = sheet.getRange(start, 1, values.length, MEETING.HEADERS.length);
    range.setNumberFormat('@'); // keep dates, NIJ-like text and "=..." as plain text
    range.setValues(values);
    SpreadsheetApp.flush();
    return Object.assign(meetingsPayload_(ss, session), { savedId: id });
  });
}

/** Ticks or unticks one task. */
function setTaskDone(token, meetingId, taskNo, done) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  return withMeetingLock_(function () {
    const sheet = getMeetingSheet_(ss, false);
    if (!sheet) throw new Error('Pertemuan ini sudah tidak ada. Refresh dulu.');
    const rows = readMeetingRows_(sheet);
    const row = rows.filter(function (r) { return r.id === String(meetingId) && r.no === String(taskNo) && r.task; })[0];
    if (!row) throw new Error('Tugas ini sudah tidak ada. Refresh dulu.');
    requireMeetingWrite_(ss, session, row);
    sheet.getRange(row.row, MEETING.HEADERS.indexOf('SELESAI') + 1).setValue(done ? '✔' : '-');
    sheet.getRange(row.row, MEETING.HEADERS.indexOf('DIPERBARUI') + 1)
      .setValue(Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd HH:mm'));
    SpreadsheetApp.flush();
    return meetingsPayload_(ss, session);
  });
}

function deleteMeeting(token, meetingId) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  return withMeetingLock_(function () {
    const sheet = getMeetingSheet_(ss, false);
    const rows = sheet ? readMeetingRows_(sheet).filter(function (r) { return r.id === String(meetingId); }) : [];
    if (!rows.length) throw new Error('Pertemuan ini sudah tidak ada. Refresh dulu.');
    requireMeetingWrite_(ss, session, rows[0]);
    deleteMeetingRows_(sheet, rows);
    SpreadsheetApp.flush();
    return meetingsPayload_(ss, session);
  });
}

// ───────────────────────────── helpers ─────────────────────────────

function withMeetingLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Sedang ada yang menyimpan. Coba lagi sebentar.');
  try { return fn(); } finally { lock.releaseLock(); }
}

function getMeetingSheet_(ss, create) {
  let sheet = ss.getSheetByName(CONFIG.MEETING_SHEET_NAME);
  if (!sheet && create) sheet = ss.insertSheet(CONFIG.MEETING_SHEET_NAME, ss.getSheets().length); // never before the database tab
  if (sheet && create && sheet.getLastRow() === 0) {
    sheet.appendRow(MEETING.HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, MEETING.HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

/** All task rows: [{ row, id, date, coach, people, category, purpose, no, task, deadline, done, ... }]. */
function readMeetingRows_(sheet) {
  if (sheet.getLastRow() === 0) return [];
  const values = sheet.getDataRange().getDisplayValues();
  if (!values.length) return [];
  const header = values[0].map(norm_);
  const c = {};
  MEETING.HEADERS.forEach(function (h) { c[h] = header.indexOf(h); });
  if (c['MEETING ID'] < 0) throw new Error('Tab "' + CONFIG.MEETING_SHEET_NAME + '" harus punya kolom MEETING ID di baris 1.');
  const get = function (r, h) { return c[h] < 0 ? '' : String(r[c[h]] || '').trim(); };
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const id = get(values[i], 'MEETING ID');
    if (!id) continue;
    out.push({
      row: i + 1,
      id: id,
      date: toIsoDate_(get(values[i], 'TANGGAL')),
      coach: titleCase_(get(values[i], 'COACH')),
      people: get(values[i], 'DENGAN').split(/\s*;\s*/).filter(Boolean),
      category: get(values[i], 'KATEGORI'),
      purpose: get(values[i], 'TUJUAN'),
      no: get(values[i], 'NO'),
      task: get(values[i], 'TO DO'),
      deadline: toIsoDate_(get(values[i], 'DEADLINE')),
      done: isChecked_(get(values[i], 'SELESAI')),
      createdBy: get(values[i], 'DIBUAT OLEH'),
      createdAt: get(values[i], 'DIBUAT PADA'),
      updatedAt: get(values[i], 'DIPERBARUI'),
    });
  }
  return out;
}

function deleteMeetingRows_(sheet, rows) {
  rows.map(function (r) { return r.row; })
    .sort(function (a, b) { return b - a; })
    .forEach(function (n) { sheet.deleteRow(n); });
}

/** Groups task rows into meetings the session may see, newest first. */
/**
 * Pertemuan with a coach or the team leader are written by admins and TLs; the coach
 * can see the ones about them, read-only. Returns a test for a meeting row/object.
 */
function leaderTest_(ss) {
  const leaders = {};
  parseDatabase_(getDatabaseSheet_(ss).getDataRange().getDisplayValues()).people.forEach(function (p) {
    if (p.isCoach) leaders[norm_(p.coach) + '|' + norm_(p.name)] = true;
  });
  return function (m) {
    return m.people.some(function (n) { return leaders[norm_(m.coach) + '|' + norm_(n)]; });
  };
}

function requireMeetingWrite_(ss, session, m) {
  if (!canAccessCoach_(session, m.coach)) throw new Error('Kamu tidak punya akses ke pertemuan ini.');
  if (!leadsCoaches_(session) && leaderTest_(ss)(m)) {
    throw new Error('Pertemuan dengan coach / team leader hanya bisa diubah oleh TL atau admin.');
  }
}

function meetingsPayload_(ss, session) {
  const sheet = getMeetingSheet_(ss, false);
  const rows = sheet ? readMeetingRows_(sheet) : [];
  const byId = {};
  const order = [];
  rows.forEach(function (r) {
    if (!canAccessCoach_(session, r.coach)) return;
    let m = byId[r.id];
    if (!m) {
      m = byId[r.id] = {
        id: r.id, date: r.date, coach: r.coach, people: r.people, category: r.category, purpose: r.purpose,
        createdBy: r.createdBy, createdAt: r.createdAt, updatedAt: r.updatedAt, tasks: [],
      };
      order.push(m);
    }
    if (r.updatedAt > m.updatedAt) m.updatedAt = r.updatedAt;
    if (r.task) m.tasks.push({ no: r.no, text: r.task, deadline: r.deadline, done: r.done });
  });
  order.sort(function (a, b) { return (b.date + b.createdAt).localeCompare(a.date + a.createdAt); });
  const aboutLeader = leaderTest_(ss);
  order.forEach(function (m) {
    m.aboutLeader = aboutLeader(m);
    m.readOnly = !leadsCoaches_(session) && m.aboutLeader;
  });
  return { meetings: order, categories: MEETING.CATEGORIES };
}

function validateMeeting_(m) {
  if (!m || typeof m !== 'object') throw new Error('Data pertemuan tidak valid.');
  const date = toIsoDate_(m.date);
  if (!date) throw new Error('Isi tanggal pertemuan.');
  const people = (Array.isArray(m.people) ? m.people : [])
    .map(function (p) { return String(p || '').trim(); })
    .filter(function (p, i, all) { return p && all.indexOf(p) === i; });
  if (!people.length) throw new Error('Pilih minimal satu orang.');
  if (people.length > MEETING.MAX_PEOPLE) throw new Error('Maksimal ' + MEETING.MAX_PEOPLE + ' orang per pertemuan.');
  const category = String(m.category || '').trim();
  if (MEETING.CATEGORIES.indexOf(category) < 0) throw new Error('Pilih kategori tujuan pertemuan.');
  const purpose = String(m.purpose || '').replace(/\s+/g, ' ').trim().slice(0, 1000);
  if (purpose.length < 3) throw new Error('Tulis tujuan pertemuan.');
  const tasks = (Array.isArray(m.tasks) ? m.tasks : []).map(function (t) {
    const text = String((t && t.text) || '').replace(/\s+/g, ' ').trim().slice(0, 300);
    const raw = String((t && t.deadline) || '').trim();
    const deadline = toIsoDate_(raw);
    if (raw && !deadline) throw new Error('Deadline "' + raw + '" tidak valid.');
    return { text: text, deadline: deadline, done: !!(t && t.done) };
  }).filter(function (t) { return t.text; });
  if (tasks.length > MEETING.MAX_TASKS) throw new Error('Maksimal ' + MEETING.MAX_TASKS + ' tugas per pertemuan.');
  const id = m.id ? String(m.id).replace(/[^A-Za-z0-9]/g, '') : '';
  return { id: id, date: date, people: people, category: category, purpose: purpose, tasks: tasks };
}

/** "2026-11-24" (or a date the sheet shows as 24/11/2026 or 24 November 2026) → "2026-11-24"; else ''. */
function toIsoDate_(s) {
  s = String(s || '').trim();
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
  let d = m ? new Date(+m[1], +m[2] - 1, +m[3]) : parseDate_(s);
  if (!d || isNaN(d.getTime())) return '';
  if (m && (d.getMonth() !== +m[2] - 1 || d.getDate() !== +m[3])) return ''; // e.g. 2026-02-31
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
}
