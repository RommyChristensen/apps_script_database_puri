/**
 * Multiplikasi: the birth of a new CG out of an existing one.
 *
 * multiplyCg moves the chosen members (and the new leader) to the new CG:
 * NO CG, CGL, COACH and the schedule fields are set on their rows. The leader
 * becomes CGL (never demoted) with TGL JD CGL = the multiplikasi date and a
 * "Berhasil" CGL entry in LDP Progress. An optional location goes to the
 * Lokasi CG tab, and the event is recorded in the "Multiplikasi CG" tab.
 *
 * Access: coaches only with their own people and into their own coach group;
 * anything involving another coach's people (or another coach) needs admin.
 * Everything is validated before the first write.
 */

const MULTIPLY_HEADERS = ['TANGGAL', 'CG ASAL', 'CG BARU', 'COACH', 'LEADER', 'JUMLAH PINDAH', 'NAMA',
  'OLEH', 'DICATAT PADA'];

/** Called from the browser: multiplikasi history this session may see, newest first. */
function getMultiplications(token) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  return multiplyPayload_(ss, session);
}

/**
 * req = {
 *   sourceCg, newCg, coach, date: 'yyyy-mm-dd',
 *   leader: { id, name }, members: [{ id, name }],
 *   cgTime, cgStatus, cgArea,
 *   location: { place, address, link, lat, lng }   // optional
 * }
 */
function multiplyCg(token, req) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  if (!req) throw new Error('Data multiplikasi tidak valid.');
  const admin = session.role === 'ADMIN';

  const newCg = normalizeCgCode_(req.newCg);
  if (!newCg) throw new Error('Nomor CG baru tidak valid. Contoh: AOG 23 atau TEEN 13.');
  const date = toIsoDate_(req.date);
  if (!date) throw new Error('Isi tanggal multiplikasi.');
  const info = {};
  CG_INFO_FIELDS.forEach(function (k) {
    const v = String(req[k] == null ? '' : req[k]).replace(/\s+/g, ' ').trim().slice(0, 100);
    validateField_(fieldDef_(k), v);
    info[k] = v;
  });
  const loc = req.location || {};
  const wantsLocation = !!(loc.link || loc.address || (loc.lat != null && loc.lng != null));
  const resolved = wantsLocation ? resolveLocation_(loc) : null; // may throw before anything is written

  return withScriptLock_(function () {
    const db = readDatabase_(ss);
    const coaches = {};
    db.people.forEach(function (p) { if (p.coach) coaches[norm_(p.coach)] = p.coach; });

    const source = cgCode_(req.sourceCg);
    if (!db.people.some(function (p) { return p.cg === source; })) throw new Error('CG asal "' + source + '" tidak ditemukan.');
    if (db.people.some(function (p) { return p.cg === newCg; })) throw new Error(newCg + ' sudah dipakai. Pilih nomor lain.');

    const coach = coaches[norm_(req.coach)];
    if (!coach) throw new Error('Pilih coach untuk CG baru.');
    if (!canAccessCoach_(session, coach)) throw new Error('Kamu hanya bisa membuat CG baru di bawah coach kamu sendiri.');

    if (!req.leader || !req.leader.id) throw new Error('Pilih leader CG baru.');
    const leader = personAt_(db, req.leader.id, req.leader.name);
    if (leader.isCoach) throw new Error('Coach / Team Leader tidak bisa jadi CGL lewat multiplikasi.');
    const members = (req.members || []).map(function (m) { return personAt_(db, m.id, m.name); })
      .filter(function (p) { return p.id !== leader.id; });
    const everyone = [leader].concat(members);
    const seen = {};
    everyone.forEach(function (p) {
      if (seen[p.id]) throw new Error(p.name + ' dipilih dua kali.');
      seen[p.id] = true;
      if (p.isCoach) throw new Error('Coach / Team Leader tidak bisa dipindah lewat multiplikasi.');
      if (!admin && !canAccessCoach_(session, p.coach)) {
        throw new Error(p.name + ' ada di coach lain. Multiplikasi lintas coach hanya bisa oleh admin.');
      }
    });
    if (db.col.coach === undefined && everyone.some(function (p) { return p.coach !== coach; })) {
      throw new Error('Pindah ke coach lain butuh kolom COACH di database.');
    }

    // Write: move everyone, then promote the leader.
    const fields = { cg: newCg, cgl: leader.name, coach: coach };
    const tl = tlOfCoach_(db.people, coach);
    if (tl) fields.tl = tl;
    CG_INFO_FIELDS.forEach(function (k) { fields[k] = info[k]; });
    everyone.forEach(function (p) { setRowFields_(db, p.id, fields); });

    const rank = POSITION_RANK_[leader.position] || 0;
    const promote = {};
    if (rank < POSITION_RANK_.CGL) promote.position = 'CGL';
    if (!leader.dateCgl) promote.dateCgl = longDate_(date);
    const leaderChanges = setRowFields_(db, leader.id, promote);
    SpreadsheetApp.flush();

    if (rank < POSITION_RANK_.CGL) {
      const cglStep = LDP_STEPS.filter(function (s) { return s.key === 'cgl'; })[0];
      appendLdpProgress_(ss, session, { coach: coach, name: leader.name, nij: leader.nij }, {
        step: cglStep, result: 'berhasil', achieved: date, newTarget: '',
        note: 'Multiplikasi ' + source + ' → ' + newCg,
        target: toIsoDate_(leader.targetCgl) || leader.targetCgl || '',
      });
    }
    if (resolved) writeLocationRow_(ss, session, newCg, coach, resolved);

    const history = getMultiplySheet_(ss);
    const now = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd HH:mm');
    const row = [date, source, newCg, coach, leader.name, String(members.length),
      members.map(function (p) { return p.name; }).join('; '), session.username, now];
    const range = history.getRange(history.getLastRow() + 1, 1, 1, row.length);
    range.setNumberFormat('@');
    range.setValues([row]);
    SpreadsheetApp.flush();

    logEdit_(ss, leader.id, leader.name, [
      { label: 'Multiplikasi', from: source, to: newCg + ' (CGL, +' + members.length + ' orang)' },
    ].concat(leaderChanges), session);

    return Object.assign(multiplyPayload_(ss, session), { newCg: newCg, data: buildDashboard_(ss, session) });
  });
}

// ───────────────────────────── helpers ─────────────────────────────

/** "aog 5" → "AOG 05", "Teen 13" → "TEEN 13"; '' when it doesn't look like PREFIX + number. */
function normalizeCgCode_(v) {
  const m = /^([A-Za-z]{2,10})\s*0*(\d{1,3})$/.exec(String(v || '').trim());
  if (!m || +m[2] === 0) return '';
  return m[1].toUpperCase() + ' ' + ('0' + (+m[2])).slice(-2);
}

function getMultiplySheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.MULTIPLY_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.MULTIPLY_SHEET_NAME, ss.getSheets().length);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(MULTIPLY_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, MULTIPLY_HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

function multiplyPayload_(ss, session) {
  const sheet = ss.getSheetByName(CONFIG.MULTIPLY_SHEET_NAME);
  const out = [];
  if (sheet && sheet.getLastRow() > 1) {
    const values = sheet.getDataRange().getDisplayValues();
    const header = values[0].map(norm_);
    const c = {};
    MULTIPLY_HEADERS.forEach(function (h) { c[h] = header.indexOf(h); });
    const get = function (r, h) { return c[h] < 0 ? '' : String(r[c[h]] || '').trim(); };
    for (let i = 1; i < values.length; i++) {
      const coach = titleCase_(get(values[i], 'COACH'));
      if (!get(values[i], 'CG BARU') || !canAccessCoach_(session, coach)) continue;
      out.push({
        date: toIsoDate_(get(values[i], 'TANGGAL')), source: get(values[i], 'CG ASAL'), newCg: get(values[i], 'CG BARU'),
        coach: coach, leader: get(values[i], 'LEADER'), count: parseInt(get(values[i], 'JUMLAH PINDAH'), 10) || 0,
        names: get(values[i], 'NAMA').split(/\s*;\s*/).filter(Boolean), by: get(values[i], 'OLEH'), at: get(values[i], 'DICATAT PADA'),
      });
    }
  }
  out.sort(function (a, b) { return (b.date + b.at).localeCompare(a.date + a.at); });
  return { history: out };
}
