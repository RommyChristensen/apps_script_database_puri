/**
 * Jadwal Acara: a year-round event calendar per TL.
 *
 * Stored in the "Jadwal Acara" tab (CONFIG.EVENT_SHEET_NAME, created automatically):
 *
 *   ID | TL | JUDUL | MULAI | SELESAI | JAM | LOKASI | KATEGORI | CATATAN |
 *   DIBUAT OLEH | DIBUAT PADA | DIPERBARUI
 *
 * TL is the owner TL's name (as in the database's TL column), or SEMUA for an
 * event every TL and coach sees. Dates are yyyy-mm-dd; SELESAI is empty for a
 * one-day event.
 *
 * Access: a coach sees their TL's events and SEMUA (read-only); a TL sees and,
 * with the "events" permission, writes their own (plus sees SEMUA); an admin
 * sees all and writes for any TL or SEMUA.
 */

const EVENTS = {
  HEADERS: ['ID', 'TL', 'JUDUL', 'MULAI', 'SELESAI', 'JAM', 'LOKASI', 'KATEGORI', 'CATATAN', 'DIBUAT OLEH', 'DIBUAT PADA', 'DIPERBARUI'],
  KEYS: ['id', 'owner', 'title', 'start', 'end', 'time', 'place', 'category', 'notes', 'by', 'at', 'updated'],
  CATEGORIES: ['Ibadah', 'Camp / Retreat', 'Training / Kelas', 'Outreach', 'Rapat', 'Lainnya'],
  ALL: 'SEMUA',
};

/** Called from the browser: the events this session may see. */
function getEvents(token) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  return eventsPayload_(ss, session);
}

/** Creates (no id) or changes (with id) an event. e = { id?, owner?, title, start, end?, time?, place?, category, notes? } */
function saveEvent(token, e) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  requireEventWriter_(session);
  if (!e) throw new Error('Data acara tidak valid.');
  const title = visitText_(e.title, 120);
  if (title.length < 2) throw new Error('Isi nama acara.');
  const start = toIsoDate_(e.start);
  if (!start) throw new Error('Isi tanggal mulai.');
  const end = e.end ? toIsoDate_(e.end) : '';
  if (e.end && !end) throw new Error('Tanggal selesai tidak valid.');
  if (end && end < start) throw new Error('Tanggal selesai tidak boleh sebelum tanggal mulai.');
  const category = EVENTS.CATEGORIES.indexOf(e.category) >= 0 ? e.category : '';
  if (!category) throw new Error('Pilih kategori acara.');

  return withScriptLock_(function () {
    const sheet = visiSheetNamed_(ss, CONFIG.EVENT_SHEET_NAME, EVENTS.HEADERS);
    const rows = readRows_(sheet, EVENTS.HEADERS, EVENTS.KEYS);
    const now = visitNow_();
    let row;
    if (e.id) {
      row = rows.filter(function (r) { return r.id === String(e.id); })[0];
      if (!row) throw new Error('Acara ini sudah tidak ada. Refresh dulu.');
      requireEventOwner_(session, row.owner);
    } else {
      row = {
        id: 'EV' + Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyMMddHHmmss') +
          Utilities.getUuid().replace(/-/g, '').slice(0, 4).toUpperCase(),
        by: session.username, at: now,
      };
    }
    const owner = session.role === 'ADMIN' ? eventOwner_(ss, e.owner || row.owner) : session.coach;
    Object.assign(row, {
      owner: owner, title: title, start: start, end: end === start ? '' : end, time: visitText_(e.time, 40),
      place: visitText_(e.place, 120), category: category, notes: visitText_(e.notes, 1000), updated: now,
    });
    writeRow_(sheet, EVENTS.KEYS, row);
    SpreadsheetApp.flush();
    return Object.assign(eventsPayload_(ss, session), { savedId: row.id });
  });
}

function deleteEvent(token, id) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  requireEventWriter_(session);
  return withScriptLock_(function () {
    const sheet = visiSheetNamed_(ss, CONFIG.EVENT_SHEET_NAME, EVENTS.HEADERS);
    const row = readRows_(sheet, EVENTS.HEADERS, EVENTS.KEYS).filter(function (r) { return r.id === String(id || ''); })[0];
    if (!row) throw new Error('Acara ini sudah tidak ada. Refresh dulu.');
    requireEventOwner_(session, row.owner);
    sheet.deleteRow(row.row);
    SpreadsheetApp.flush();
    return eventsPayload_(ss, session);
  });
}

// ───────────────────────────── helpers ─────────────────────────────

function requireEventWriter_(session) {
  if (!leadsCoaches_(session)) throw new Error('Hanya TL atau admin yang bisa mengisi jadwal acara.');
  requirePerm_(session, 'events');
}

function requireEventOwner_(session, owner) {
  if (session.role === 'ADMIN') return;
  if (norm_(owner) !== norm_(session.coach)) throw new Error('Kamu hanya bisa mengubah acara milik kamu.');
}

/** Admin: the owner must be SEMUA or a TL from the database. */
function eventOwner_(ss, owner) {
  owner = String(owner || '').trim();
  if (!owner || norm_(owner) === EVENTS.ALL) return EVENTS.ALL;
  const tls = eventTls_(ss);
  const hit = tls.filter(function (t) { return norm_(t) === norm_(owner); })[0];
  if (!hit) throw new Error('TL "' + owner + '" tidak ditemukan di kolom TL database.');
  return hit;
}

function eventTls_(ss) {
  const seen = {};
  parseDatabase_(getDatabaseSheet_(ss).getDataRange().getDisplayValues()).people.forEach(function (p) {
    if (p.tl) seen[p.tl] = true;
  });
  return Object.keys(seen).sort();
}

/** The TL whose calendar this session follows ('' for admin). */
function myEventTl_(ss, session) {
  if (session.role === 'TL') return session.coach;
  if (session.role !== 'COACH') return '';
  const people = parseDatabase_(getDatabaseSheet_(ss).getDataRange().getDisplayValues()).people;
  const me = people.filter(function (p) { return !p.isLeader && norm_(p.coach) === norm_(session.coach); })[0];
  return me ? tlOfCoach_(people, me.coach) : '';
}

function eventsPayload_(ss, session) {
  const admin = session.role === 'ADMIN';
  const myTl = admin ? '' : myEventTl_(ss, session);
  const canWrite = leadsCoaches_(session) && hasPerm_(session, 'events');
  const events = readRows_(ss.getSheetByName(CONFIG.EVENT_SHEET_NAME), EVENTS.HEADERS, EVENTS.KEYS)
    .filter(function (r) {
      if (admin) return true;
      return norm_(r.owner) === EVENTS.ALL || (!!myTl && norm_(r.owner) === norm_(myTl));
    })
    .map(function (r) {
      const start = toIsoDate_(r.start) || r.start;
      const end = toIsoDate_(r.end) || '';
      return {
        id: r.id, owner: norm_(r.owner) === EVENTS.ALL ? EVENTS.ALL : r.owner, title: r.title, start: start,
        end: end && end !== start ? end : '', time: r.time, place: r.place, category: r.category, notes: r.notes,
        by: r.by, at: r.at, updated: r.updated,
        canEdit: canWrite && (admin || norm_(r.owner) === norm_(session.coach)),
      };
    })
    .filter(function (e) { return /^\d{4}-\d{2}-\d{2}$/.test(e.start); })
    .sort(function (a, b) { return a.start.localeCompare(b.start) || (a.time || '').localeCompare(b.time || '') || a.title.localeCompare(b.title); });
  return {
    events: events,
    categories: EVENTS.CATEGORIES,
    owners: admin ? [EVENTS.ALL].concat(eventTls_(ss)) : [],
    canWrite: canWrite,
    myTl: myTl,
    isAdmin: admin,
    today: Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd'),
  };
}
