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
 * Access: a coach works with their own visi, a TL with their own and those of
 * the coaches under them, admin with all (and can write one for anyone).
 */

const VISI = {
  HEADERS: ['ID', 'PEMILIK', 'PERAN', 'VISI', 'TARGET JIWA', 'TARGET CG', 'TARGET SPONSOR', 'TARGET TANGGAL',
    'AWAL JIWA', 'AWAL CG', 'AWAL SPONSOR', 'DIBUAT OLEH', 'DIBUAT PADA', 'DIPERBARUI'],
  KEYS: ['id', 'owner', 'role', 'text', 'tPeople', 'tCgs', 'tSponsors', 'tDate', 'bPeople', 'bCgs', 'bSponsors',
    'by', 'at', 'updated'],
  MS_HEADERS: ['ID', 'VISI ID', 'TANGGAL', 'CATATAN', 'JIWA', 'CG', 'SPONSOR', 'OLEH', 'DICATAT PADA'],
  MS_KEYS: ['id', 'visiId', 'date', 'note', 'people', 'cgs', 'sponsors', 'by', 'at'],
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
      by: session.username, at: visitNow_(),
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

/** Whose visi this session may see and write: [{ name, role }]. */
function visiOwners_(session, people) {
  const out = [], seen = {};
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
  if (!hit) throw new Error('Kamu tidak bisa menulis visi untuk ' + (name || 'orang itu') + '.');
  return hit;
}

function requireVisiOwner_(session, people, name, role) {
  if (session.role === 'ADMIN') return;
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
  const mine = {};
  owners.forEach(function (o) { mine[o.role + '|' + norm_(o.name)] = true; });
  const num = function (s) { return s === '' || isNaN(+s) ? null : +s; };

  const msByVisi = {};
  readRows_(ss.getSheetByName(CONFIG.VISI_MILESTONE_SHEET_NAME), VISI.MS_HEADERS, VISI.MS_KEYS).forEach(function (m) {
    (msByVisi[m.visiId] = msByVisi[m.visiId] || []).push({
      id: m.id, date: toIsoDate_(m.date) || m.date, note: m.note,
      people: num(m.people), cgs: num(m.cgs), sponsors: num(m.sponsors), by: m.by, at: m.at,
    });
  });

  const visions = readRows_(ss.getSheetByName(CONFIG.VISI_SHEET_NAME), VISI.HEADERS, VISI.KEYS)
    .filter(function (r) { return session.role === 'ADMIN' || mine[r.role + '|' + norm_(r.owner)]; })
    .map(function (r) {
      const ms = (msByVisi[r.id] || []).sort(function (a, b) { return (a.date + a.at).localeCompare(b.date + b.at); });
      return {
        id: r.id, owner: r.owner, role: r.role, text: r.text,
        target: { people: num(r.tPeople), cgs: num(r.tCgs), sponsors: num(r.tSponsors), date: toIsoDate_(r.tDate) || r.tDate },
        base: { people: num(r.bPeople), cgs: num(r.bCgs), sponsors: num(r.bSponsors) },
        now: visiCounts_(people, r.role, r.owner),
        milestones: ms, by: r.by, at: r.at, updated: r.updated, canEdit: true,
      };
    });
  visions.sort(function (a, b) { return (a.target.date || '').localeCompare(b.target.date || '') || a.owner.localeCompare(b.owner); });
  return {
    visions: visions,
    owners: owners.map(function (o) { return { name: o.name, role: o.role, now: visiCounts_(people, o.role, o.name) }; }),
    today: Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd'),
  };
}
