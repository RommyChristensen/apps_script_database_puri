/**
 * SWOT analysis of each coach, written by their TL (or an admin).
 *
 * Stored in the "SWOT Coach" tab (CONFIG.SWOT_SHEET_NAME), one row per saved
 * version, so the history is kept; the newest row of a coach is the current
 * SWOT:
 *
 *   ID | COACH | STRENGTHS | WEAKNESSES | OPPORTUNITIES | THREATS | OLEH | PERAN | DISIMPAN PADA
 *
 * Access: anyone who can see the coach can read it (the coach sees their own,
 * read-only); TLs write it for the coaches under them, admins for every coach.
 */

const SWOT = {
  HEADERS: ['ID', 'COACH', 'STRENGTHS', 'WEAKNESSES', 'OPPORTUNITIES', 'THREATS', 'OLEH', 'PERAN', 'DISIMPAN PADA'],
  KEYS: ['id', 'coach', 's', 'w', 'o', 't', 'by', 'role', 'at'],
  FIELDS: ['s', 'w', 'o', 't'],
  MAX: 3000,
};

/** Called from the browser: the current SWOT of a coach and its earlier versions. */
function getSwot(token, coach) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  const key = swotCoach_(ss, session, coach);
  return swotPayload_(ss, session, key);
}

/** Saves a new version. v = { coach, s, w, o, t } */
function saveSwot(token, v) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  if (!v) throw new Error('Data SWOT tidak valid.');
  const coach = swotCoach_(ss, session, v.coach);
  if (!leadsCoaches_(session)) throw new Error('Hanya TL atau admin yang bisa mengisi SWOT coach.');
  requirePerm_(session, 'visi');
  const clean = {};
  SWOT.FIELDS.forEach(function (k) { clean[k] = visitText_(v[k], SWOT.MAX); });
  if (!SWOT.FIELDS.some(function (k) { return clean[k]; })) throw new Error('Isi minimal satu bagian SWOT.');

  return withScriptLock_(function () {
    const sheet = visiSheetNamed_(ss, CONFIG.SWOT_SHEET_NAME, SWOT.HEADERS);
    const rows = readRows_(sheet, SWOT.HEADERS, SWOT.KEYS).filter(function (r) { return norm_(r.coach) === norm_(coach); });
    const last = rows[rows.length - 1];
    const same = last && SWOT.FIELDS.every(function (k) { return last[k] === clean[k]; });
    if (!same) {
      writeRow_(sheet, SWOT.KEYS, Object.assign({
        id: 'SW' + Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyMMddHHmmss') +
          Utilities.getUuid().replace(/-/g, '').slice(0, 4).toUpperCase(),
        coach: coach, by: session.role === 'ADMIN' ? session.username : session.coach, role: roleLabel_(session), at: visitNow_(),
      }, clean));
      SpreadsheetApp.flush();
    }
    return Object.assign(swotPayload_(ss, session, coach), { unchanged: !!same });
  });
}

// ───────────────────────────── helpers ─────────────────────────────

/** The coach as written in the database; throws when this session can't see them. */
function swotCoach_(ss, session, coach) {
  const people = parseDatabase_(getDatabaseSheet_(ss).getDataRange().getDisplayValues()).people;
  const hit = people.filter(function (p) { return !p.isLeader && p.coach && norm_(p.coach) === norm_(coach); })[0];
  if (!hit) throw new Error('Coach "' + coach + '" tidak ditemukan.');
  if (!canAccessCoach_(session, hit.coach)) throw new Error('Kamu tidak punya akses ke coach ini.');
  return hit.coach;
}

function visiSheetNamed_(ss, name, headers) {
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

function swotPayload_(ss, session, coach) {
  const versions = readRows_(ss.getSheetByName(CONFIG.SWOT_SHEET_NAME), SWOT.HEADERS, SWOT.KEYS)
    .filter(function (r) { return norm_(r.coach) === norm_(coach); })
    .sort(function (a, b) { return b.row - a.row; }) // newest version = last row in the tab
    .map(function (r) { return { id: r.id, s: r.s, w: r.w, o: r.o, t: r.t, by: r.by, role: r.role, at: r.at }; });
  return { coach: coach, current: versions[0] || null, history: versions.slice(1), canEdit: leadsCoaches_(session) && hasPerm_(session, 'visi') };
}
