/**
 * Jadwal Acara: a year-round event calendar for TLs and their coaches.
 *
 * Stored in the "Jadwal Acara" tab (CONFIG.EVENT_SHEET_NAME, created automatically):
 *
 *   ID | TL | JUDUL | MULAI | SELESAI | JAM | LOKASI | KATEGORI | CATATAN |
 *   DIBUAT OLEH | DIBUAT PADA | DIPERBARUI | COACH
 *
 * - TL event:    TL = the TL's name, COACH empty.
 * - Coach event: TL = that coach's TL, COACH = the coach's name.
 * - SEMUA:       TL = SEMUA (posted by an admin, everyone sees it).
 * Dates are yyyy-mm-dd; SELESAI is empty for a one-day event. Tabs from the
 * previous version get the COACH column added; their rows are TL events.
 *
 * Who sees what:
 *   coach — their own events, their TL's TL events and SEMUA
 *   TL    — their own events, the events of every coach under them, and SEMUA
 *   admin — everything
 * Everyone (with the "events" permission) writes only their own events; a TL
 * sees coach events read-only. Admins write for SEMUA, any TL or any coach.
 */

const EVENTS = {
  HEADERS: ['ID', 'TL', 'JUDUL', 'MULAI', 'SELESAI', 'JAM', 'LOKASI', 'KATEGORI', 'CATATAN', 'DIBUAT OLEH', 'DIBUAT PADA',
    'DIPERBARUI', 'COACH'],
  KEYS: ['id', 'owner', 'title', 'start', 'end', 'time', 'place', 'category', 'notes', 'by', 'at', 'updated', 'coach'],
  CATEGORIES: ['Ibadah', 'Camp / Retreat', 'Training / Kelas', 'Outreach', 'Rapat', 'Lainnya'],
  ALL: 'SEMUA',
};

/** Called from the browser: the events this session may see. */
function getEvents(token) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  return eventsPayload_(ss, session);
}

/**
 * Creates (no id) or changes (with id) an event.
 * e = { id?, owner? (admin: 'SEMUA' | 'TL:<name>' | 'COACH:<name>' | a TL name), title, start, end?, time?, place?, category, notes? }
 */
function saveEvent(token, e) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  requirePerm_(session, 'events');
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
    const people = parseDatabase_(getDatabaseSheet_(ss).getDataRange().getDisplayValues()).people;
    const sheet = visiSheetNamed_(ss, CONFIG.EVENT_SHEET_NAME, EVENTS.HEADERS);
    const rows = readRows_(sheet, EVENTS.HEADERS, EVENTS.KEYS);
    const now = visitNow_();
    let row;
    if (e.id) {
      row = rows.filter(function (r) { return r.id === String(e.id); })[0];
      if (!row) throw new Error('Acara ini sudah tidak ada. Refresh dulu.');
      requireEventOwner_(session, row);
    } else {
      row = {
        id: 'EV' + Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyMMddHHmmss') +
          Utilities.getUuid().replace(/-/g, '').slice(0, 4).toUpperCase(),
        by: session.username, at: now,
      };
    }
    const owner = eventOwnerFor_(session, people, e.owner, row);
    Object.assign(row, owner, {
      title: title, start: start, end: end === start ? '' : end, time: visitText_(e.time, 40),
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
  requirePerm_(session, 'events');
  return withScriptLock_(function () {
    const sheet = visiSheetNamed_(ss, CONFIG.EVENT_SHEET_NAME, EVENTS.HEADERS);
    const row = readRows_(sheet, EVENTS.HEADERS, EVENTS.KEYS).filter(function (r) { return r.id === String(id || ''); })[0];
    if (!row) throw new Error('Acara ini sudah tidak ada. Refresh dulu.');
    requireEventOwner_(session, row);
    sheet.deleteRow(row.row);
    SpreadsheetApp.flush();
    return eventsPayload_(ss, session);
  });
}

// ───────────────────────────── helpers ─────────────────────────────

/** True when this session wrote / owns the event (admins own everything). */
function ownsEvent_(session, row) {
  if (session.role === 'ADMIN') return true;
  if (session.role === 'TL') return !row.coach && norm_(row.owner) === norm_(session.coach);
  if (session.role === 'COACH') return !!row.coach && norm_(row.coach) === norm_(session.coach);
  return false;
}

function requireEventOwner_(session, row) {
  if (!ownsEvent_(session, row)) throw new Error('Kamu hanya bisa mengubah acara milik kamu sendiri.');
}

/** { owner (TL column), coach } for a new or edited event. Only admins choose; others write their own calendar. */
function eventOwnerFor_(session, people, wanted, row) {
  if (session.role === 'COACH') return { owner: coachTl_(people, session.coach), coach: coachName_(people, session.coach) };
  if (session.role === 'TL') return { owner: session.coach, coach: '' };
  // admin: keep the event's calendar unless another one is picked
  let w = String(wanted == null ? '' : wanted).trim();
  if (!w) return row.id && (row.owner || row.coach) ? { owner: row.owner, coach: row.coach || '' } : { owner: EVENTS.ALL, coach: '' };
  if (norm_(w) === EVENTS.ALL) return { owner: EVENTS.ALL, coach: '' };
  const m = /^(TL|COACH):(.+)$/i.exec(w);
  if (m && m[1].toUpperCase() === 'COACH') {
    const coach = coachName_(people, m[2]);
    if (!coach) throw new Error('Coach "' + m[2] + '" tidak ditemukan di database.');
    return { owner: coachTl_(people, coach), coach: coach };
  }
  w = m ? m[2] : w;
  const tl = eventTls_(people).filter(function (t) { return norm_(t) === norm_(w); })[0];
  if (!tl) throw new Error('TL "' + w + '" tidak ditemukan di kolom TL database.');
  return { owner: tl, coach: '' };
}

function coachName_(people, name) {
  const p = people.filter(function (x) { return !x.isLeader && x.coach && norm_(x.coach) === norm_(name); })[0];
  return p ? p.coach : '';
}
function coachTl_(people, coach) {
  const c = coachName_(people, coach);
  return c ? tlOfCoach_(people, c) : '';
}

function eventTls_(people) {
  const seen = {};
  people.forEach(function (p) { if (p.tl) seen[p.tl] = true; });
  return Object.keys(seen).sort();
}
function eventCoaches_(people) {
  const seen = {};
  people.forEach(function (p) { if (!p.isLeader && p.coach && p.coach !== 'Unassigned') seen[p.coach] = true; });
  return Object.keys(seen).sort();
}

function eventsPayload_(ss, session) {
  const people = parseDatabase_(getDatabaseSheet_(ss).getDataRange().getDisplayValues()).people;
  const admin = session.role === 'ADMIN';
  const myTl = session.role === 'TL' ? session.coach : session.role === 'COACH' ? coachTl_(people, session.coach) : '';
  const myCoach = session.role === 'COACH' ? (coachName_(people, session.coach) || session.coach) : '';
  const canWrite = hasPerm_(session, 'events');

  const events = readRows_(ss.getSheetByName(CONFIG.EVENT_SHEET_NAME), EVENTS.HEADERS, EVENTS.KEYS)
    .filter(function (r) {
      if (admin) return true;
      if (norm_(r.owner) === EVENTS.ALL && !r.coach) return true;
      if (session.role === 'TL') return norm_(r.owner) === norm_(session.coach);
      if (session.role === 'COACH') {
        if (r.coach) return norm_(r.coach) === norm_(session.coach);
        return !!myTl && norm_(r.owner) === norm_(myTl);
      }
      return false;
    })
    .map(function (r) {
      const start = toIsoDate_(r.start) || r.start;
      const end = toIsoDate_(r.end) || '';
      const all = norm_(r.owner) === EVENTS.ALL && !r.coach;
      return {
        id: r.id, owner: all ? EVENTS.ALL : r.owner, coach: r.coach || '', kind: r.coach ? 'coach' : all ? 'all' : 'tl',
        title: r.title, start: start, end: end && end !== start ? end : '', time: r.time, place: r.place,
        category: r.category, notes: r.notes, by: r.by, at: r.at, updated: r.updated,
        canEdit: canWrite && ownsEvent_(session, r),
      };
    })
    .filter(function (e) { return /^\d{4}-\d{2}-\d{2}$/.test(e.start); })
    .sort(function (a, b) { return a.start.localeCompare(b.start) || (a.time || '').localeCompare(b.time || '') || a.title.localeCompare(b.title); });

  return {
    events: events,
    categories: EVENTS.CATEGORIES,
    owners: admin ? [{ value: EVENTS.ALL, label: 'Semua TL', group: '' }]
      .concat(eventTls_(people).map(function (t) { return { value: 'TL:' + t, label: 'TL ' + t, group: 'TL' }; }))
      .concat(eventCoaches_(people).map(function (c) { return { value: 'COACH:' + c, label: 'Coach ' + c, group: 'Coach' }; })) : [],
    canWrite: canWrite,
    myTl: myTl,
    myCoach: myCoach,
    role: session.role,
    isAdmin: admin,
    today: Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd'),
  };
}
