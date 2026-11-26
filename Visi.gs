/**
 * Visi: the vision of a coach or a TL, with targets (jiwa in CG, number of
 * CG, number of sponsors, target date) and progress milestones.
 *
 * Tabs (created automatically):
 *   "Visi"            ID | PEMILIK | PERAN | VISI | TARGET JIWA | TARGET CG | TARGET SPONSOR |
 *                     TARGET TANGGAL | AWAL JIWA | AWAL CG | AWAL SPONSOR | DIBUAT OLEH |
 *                     DIBUAT PADA | DIPERBARUI
 *   "Visi Milestone"  ID | VISI ID | TANGGAL | CATATAN | JIWA | CG | SPONSOR | OLEH | DICATAT PADA
 *
 * Progress numbers are counted from the database, never typed in:
 *   jiwa    = people in a CG (not coach / TL rows), STATUS not Non-Aktif
 *   CG      = distinct NO CG among them
 *   sponsor = those with POSISI SPONSOR
 * for a coach: rows with that COACH; for a TL: rows with that TL. AWAL is
 * recorded when the visi is created, and every milestone stores the numbers
 * at the moment it is recorded.
 *
 * Access: only TLs and admins write visi and milestones (a TL for themselves
 * and the coaches under them, admin for anyone). A coach sees, read-only, the
 * visi of their TL and visi written about them, with their own contribution:
 * their share of the numbers now and their growth since the visi started
 * (AWAL PER COACH / PER COACH hold the numbers per coach as JSON).
 */

const VISI = {
  HEADERS: ['ID', 'PEMILIK', 'PERAN', 'VISI', 'TARGET JIWA', 'TARGET CG', 'TARGET SPONSOR', 'TARGET TANGGAL',
    'AWAL JIWA', 'AWAL CG', 'AWAL SPONSOR', 'DIBUAT OLEH', 'DIBUAT PADA', 'DIPERBARUI', 'AWAL PER COACH'],
  KEYS: ['id', 'owner', 'role', 'text', 'tPeople', 'tCgs', 'tSponsors', 'tDate', 'bPeople', 'bCgs', 'bSponsors',
    'by', 'at', 'updated', 'bByCoach'],
  MS_HEADERS: ['ID', 'VISI ID', 'TANGGAL', 'CATATAN', 'JIWA', 'CG', 'SPONSOR', 'OLEH', 'DICATAT PADA', 'PER COACH'],
  MS_KEYS: ['id', 'visiId', 'date', 'note', 'people', 'cgs', 'sponsors', 'by', 'at', 'byCoach'],
  ROLES: ['Coach', 'TL'],
};

/** Called from the browser: the visi this session may see, plus whose visi it may write. */
function getVisi(token) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  return visiPayload_(ss, session);
}

/** Creates (no id) or changes (with id) a visi. v = { id?, owner, role, text, people, cgs, sponsors, date } */
function saveVisi(token, v) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  if (!v) throw new Error('Data visi tidak valid.');
  const text = visitText_(v.text, 3000);
  if (text.length < 5) throw new Error('Tulis visinya.');
  const nums = {};
  [['people', 'jiwa'], ['cgs', 'CG'], ['sponsors', 'sponsor']].forEach(function (k) {
    const s = String(v[k[0]] == null ? '' : v[k[0]]).trim();
    if (!/^\d{1,5}$/.test(s) || +s === 0) throw new Error('Isi target jumlah ' + k[1] + ' (angka lebih dari 0).');
    nums[k[0]] = String(+s);
  });
  const date = toIsoDate_(v.date);
  if (!date) throw new Error('Isi tanggal target.');

  return withScriptLock_(function () {
    const people = parseDatabase_(getDatabaseSheet_(ss).getDataRange().getDisplayValues()).people;
    const sheet = visiSheet_(ss, 'Visi', VISI.HEADERS);
    const rows = readRows_(sheet, VISI.HEADERS, VISI.KEYS);
    const now = visitNow_();
    let row;
    if (v.id) {
      row = rows.filter(function (r) { return r.id === String(v.id); })[0];
      if (!row) throw new Error('Visi ini sudah tidak ada. Refresh dulu.');
      requireVisiOwner_(session, people, row.owner, row.role);
    } else {
      const owner = visiOwner_(session, people, v.owner, v.role);
      const base = visiCounts_(people, owner.role, owner.name);
      row = {
        id: 'VS' + Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyMMddHHmmss') +
          Utilities.getUuid().replace(/-/g, '').slice(0, 4).toUpperCase(),
        owner: owner.name, role: owner.role, bPeople: String(base.people), bCgs: String(base.cgs),
        bSponsors: String(base.sponsors), by: session.username, at: now,
        bByCoach: JSON.stringify(visiBreakdown_(people, owner.role, owner.name)),
      };
    }
    Object.assign(row, { text: text, tPeople: nums.people, tCgs: nums.cgs, tSponsors: nums.sponsors, tDate: date, updated: now });
    writeRow_(sheet, VISI.KEYS, row);
    SpreadsheetApp.flush();
    return Object.assign(visiPayload_(ss, session), { savedId: row.id });
  });
}

/** Deletes a visi and its milestones. */
function deleteVisi(token, id) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  return withScriptLock_(function () {
    const people = parseDatabase_(getDatabaseSheet_(ss).getDataRange().getDisplayValues()).people;
    const sheet = visiSheet_(ss, 'Visi', VISI.HEADERS);
    const row = readRows_(sheet, VISI.HEADERS, VISI.KEYS).filter(function (r) { return r.id === String(id || ''); })[0];
    if (!row) throw new Error('Visi ini sudah tidak ada. Refresh dulu.');
    requireVisiOwner_(session, people, row.owner, row.role);
    const ms = ss.getSheetByName(CONFIG.VISI_MILESTONE_SHEET_NAME);
    if (ms) {
      readRows_(ms, VISI.MS_HEADERS, VISI.MS_KEYS).filter(function (m) { return m.visiId === row.id; })
        .sort(function (a, b) { return b.row - a.row; })
        .forEach(function (m) { ms.deleteRow(m.row); });
    }
    sheet.deleteRow(row.row);
    SpreadsheetApp.flush();
    return visiPayload_(ss, session);
  });
}

/** Records a milestone with today's numbers. m = { visiId, date, note } */
function addVisiMilestone(token, m) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  if (!m) throw new Error('Data milestone tidak valid.');
  const date = toIsoDate_(m.date);
  if (!date) throw new Error('Isi tanggal milestone.');
  if (date > Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd')) throw new Error('Tanggal milestone tidak boleh di masa depan.');
  const note = visitText_(m.note, 2000);
  if (note.length < 3) throw new Error('Tulis catatan progresnya.');

  return withScriptLock_(function () {
    const people = parseDatabase_(getDatabaseSheet_(ss).getDataRange().getDisplayValues()).people;
    const visi = readRows_(visiSheet_(ss, 'Visi', VISI.HEADERS), VISI.HEADERS, VISI.KEYS)
      .filter(function (r) { return r.id === String(m.visiId || ''); })[0];
    if (!visi) throw new Error('Visi ini sudah tidak ada. Refresh dulu.');
    requireVisiOwner_(session, people, visi.owner, visi.role);
    const n = visiCounts_(people, visi.role, visi.owner);
    const sheet = visiSheet_(ss, 'Milestone', VISI.MS_HEADERS);
    writeRow_(sheet, VISI.MS_KEYS, {
      id: 'VM' + Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyMMddHHmmss') +
        Utilities.getUuid().replace(/-/g, '').slice(0, 4).toUpperCase(),
      visiId: visi.id, date: date, note: note, people: String(n.people), cgs: String(n.cgs), sponsors: String(n.sponsors),
      by: session.username, at: visitNow_(), byCoach: JSON.stringify(visiBreakdown_(people, visi.role, visi.owner)),
    });
    SpreadsheetApp.flush();
    return visiPayload_(ss, session);
  });
}

function deleteVisiMilestone(token, id) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  return withScriptLock_(function () {
    const people = parseDatabase_(getDatabaseSheet_(ss).getDataRange().getDisplayValues()).people;
    const sheet = visiSheet_(ss, 'Milestone', VISI.MS_HEADERS);
    const m = readRows_(sheet, VISI.MS_HEADERS, VISI.MS_KEYS).filter(function (r) { return r.id === String(id || ''); })[0];
    if (!m) throw new Error('Milestone ini sudah tidak ada. Refresh dulu.');
    const visi = readRows_(visiSheet_(ss, 'Visi', VISI.HEADERS), VISI.HEADERS, VISI.KEYS)
      .filter(function (r) { return r.id === m.visiId; })[0];
    if (visi) requireVisiOwner_(session, people, visi.owner, visi.role);
    else if (session.role !== 'ADMIN') throw new Error('Visi milestone ini sudah tidak ada.');
    sheet.deleteRow(m.row);
    SpreadsheetApp.flush();
    return visiPayload_(ss, session);
  });
}

// ───────────────────────────── helpers ─────────────────────────────

function isNonActive_(p) {
  return /^non/i.test(String(p.status || '').trim());
}

/** Live numbers for a coach ('Coach') or a TL ('TL'): { people, cgs, sponsors }. */
function visiCounts_(people, role, owner) {
  const key = norm_(owner);
  const list = people.filter(function (p) {
    if (p.isCoach || !p.cg || isNonActive_(p)) return false;
    return role === 'TL' ? norm_(p.tl) === key : norm_(p.coach) === key;
  });
  const cgs = {};
  list.forEach(function (p) { cgs[p.cg] = true; });
  return {
    people: list.length, cgs: Object.keys(cgs).length,
    sponsors: list.filter(function (p) { return p.position === 'SPONSOR'; }).length,
  };
}

/** The numbers per coach inside a visi's scope: { coach: { people, cgs, sponsors } }. */
function visiBreakdown_(people, role, owner) {
  const coaches = {};
  people.forEach(function (p) {
    if (p.isLeader || !p.coach || p.coach === 'Unassigned') return;
    if (role === 'TL' ? norm_(p.tl) === norm_(owner) : norm_(p.coach) === norm_(owner)) coaches[p.coach] = true;
  });
  const out = {};
  Object.keys(coaches).forEach(function (c) { out[c] = visiCounts_(people, 'Coach', c); });
  return out;
}

/** A coach sees the visi of their TL and visi about themselves; TL / admin see what they may write. */
function visiVisible_(session, people, row, writable) {
  if (session.role === 'ADMIN') return true;
  if (writable[row.role + '|' + norm_(row.owner)]) return true;
  if (session.role !== 'COACH') return false;
  if (row.role === 'Coach') return norm_(row.owner) === norm_(session.coach);
  const me = people.filter(function (p) { return !p.isLeader && norm_(p.coach) === norm_(session.coach); })[0];
  return !!me && norm_(tlOfCoach_(people, me.coach)) === norm_(row.owner);
}

function parseJson_(s) {
  try { const v = JSON.parse(s || ''); return v && typeof v === 'object' ? v : null; } catch (e) { return null; }
}

/** Whose visi this session may write (TL / admin only): [{ name, role }]. */
function visiOwners_(session, people) {
  const out = [], seen = {};
  if (!leadsCoaches_(session) || !hasPerm_(session, 'visi')) return out; // coaches only read
  const add = function (name, role) {
    const k = role + '|' + norm_(name);
    if (!name || seen[k]) return;
    seen[k] = true;
    out.push({ name: name, role: role });
  };
  people.forEach(function (p) {
    if (p.isLeader || !p.coach || p.coach === 'Unassigned') return;
    if (!canAccessCoach_(session, p.coach)) return;
    if (session.role === 'TL' && norm_(p.tl) !== norm_(session.coach)) return;
    add(p.coach, 'Coach');
    if (session.role === 'ADMIN' && p.tl) add(p.tl, 'TL');
  });
  if (session.role === 'TL') add(session.coach, 'TL');
  return out.sort(function (a, b) { return (a.role === b.role ? 0 : a.role === 'TL' ? -1 : 1) || a.name.localeCompare(b.name); });
}

function visiOwner_(session, people, name, role) {
  role = VISI.ROLES.indexOf(role) >= 0 ? role : (session.role === 'TL' ? 'TL' : 'Coach');
  if (!name) name = session.role === 'ADMIN' ? '' : session.coach;
  const hit = visiOwners_(session, people).filter(function (o) { return o.role === role && norm_(o.name) === norm_(name); })[0];
  if (!leadsCoaches_(session)) throw new Error('Hanya TL atau admin yang bisa mengisi visi dan milestone.');
  requirePerm_(session, 'visi');
  if (!hit) throw new Error('Kamu tidak bisa menulis visi untuk ' + (name || 'orang itu') + '.');
  return hit;
}

function requireVisiOwner_(session, people, name, role) {
  if (session.role === 'ADMIN') return requirePerm_(session, 'visi');
  visiOwner_(session, people, name, role);
}

function visiSheet_(ss, which, headers) {
  const name = which === 'Visi' ? CONFIG.VISI_SHEET_NAME : CONFIG.VISI_MILESTONE_SHEET_NAME;
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name, ss.getSheets().length);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  } else if (sheet.getLastColumn() < headers.length) {
    // A tab from an older version: add the new header cells at the end.
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
  return sheet;
}

/** Rows of a simple tab as objects keyed by `keys` (header order), with .row = sheet row. */
function readRows_(sheet, headers, keys) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getDisplayValues();
  const header = values[0].map(norm_);
  const c = headers.map(function (h) { return header.indexOf(h); });
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = { row: i + 1 };
    keys.forEach(function (k, j) { r[k] = c[j] < 0 ? '' : String(values[i][c[j]] == null ? '' : values[i][c[j]]).trim(); });
    if (r.id) out.push(r);
  }
  return out;
}

function writeRow_(sheet, keys, r) {
  const range = sheet.getRange(r.row || sheet.getLastRow() + 1, 1, 1, keys.length);
  range.setNumberFormat('@');
  range.setValues([keys.map(function (k) { return String(r[k] == null ? '' : r[k]); })]);
}

function visiPayload_(ss, session) {
  const people = parseDatabase_(getDatabaseSheet_(ss).getDataRange().getDisplayValues()).people;
  const owners = visiOwners_(session, people);
  const writable = {};
  owners.forEach(function (o) { writable[o.role + '|' + norm_(o.name)] = true; });
  const me = session.role === 'COACH' ? session.coach : '';
  const num = function (s) { return s === '' || isNaN(+s) ? null : +s; };
  const mineOf = function (obj) {
    if (!obj) return null;
    const k = Object.keys(obj).filter(function (c) { return norm_(c) === norm_(me); })[0];
    return k ? obj[k] : null;
  };

  const msByVisi = {};
  readRows_(ss.getSheetByName(CONFIG.VISI_MILESTONE_SHEET_NAME), VISI.MS_HEADERS, VISI.MS_KEYS).forEach(function (m) {
    (msByVisi[m.visiId] = msByVisi[m.visiId] || []).push({
      id: m.id, date: toIsoDate_(m.date) || m.date, note: m.note,
      people: num(m.people), cgs: num(m.cgs), sponsors: num(m.sponsors), by: m.by, at: m.at,
      mine: me ? mineOf(parseJson_(m.byCoach)) : null,
    });
  });

  const visions = readRows_(ss.getSheetByName(CONFIG.VISI_SHEET_NAME), VISI.HEADERS, VISI.KEYS)
    .filter(function (r) { return visiVisible_(session, people, r, writable); })
    .map(function (r) {
      const ms = (msByVisi[r.id] || []).sort(function (a, b) { return (a.date + a.at).localeCompare(b.date + b.at); });
      const nowBy = visiBreakdown_(people, r.role, r.owner);
      const baseBy = parseJson_(r.bByCoach) || {};
      // A coach only receives their own line; TL / admin get every coach in the visi.
      const contributions = Object.keys(nowBy)
        .filter(function (c) { return !me || norm_(c) === norm_(me); })
        .map(function (c) { return { coach: c, now: nowBy[c], base: baseBy[c] || null }; })
        .sort(function (a, b) { return b.now.people - a.now.people || a.coach.localeCompare(b.coach); });
      return {
        id: r.id, owner: r.owner, role: r.role, text: r.text,
        target: { people: num(r.tPeople), cgs: num(r.tCgs), sponsors: num(r.tSponsors), date: toIsoDate_(r.tDate) || r.tDate },
        base: { people: num(r.bPeople), cgs: num(r.bCgs), sponsors: num(r.bSponsors) },
        now: visiCounts_(people, r.role, r.owner),
        contributions: contributions, hasBaseByCoach: !!r.bByCoach,
        milestones: ms, by: r.by, at: r.at, updated: r.updated,
        canEdit: !!writable[r.role + '|' + norm_(r.owner)] || (session.role === 'ADMIN' && hasPerm_(session, 'visi')),
      };
    });
  visions.sort(function (a, b) { return (a.target.date || '').localeCompare(b.target.date || '') || a.owner.localeCompare(b.owner); });
  return {
    visions: visions,
    owners: owners.map(function (o) { return { name: o.name, role: o.role, now: visiCounts_(people, o.role, o.name) }; }),
    me: me,
    canCreate: owners.length > 0,
    today: Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd'),
  };
}
