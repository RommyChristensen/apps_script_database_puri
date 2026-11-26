/**
 * LDP — Leaders Development Program.
 *
 * Each person climbs Member → Sponsor → CGL → Coach → Team Leader. For every
 * step the database tab has a TARGET column (planned date) and a TGL column
 * (when it actually happened), e.g. TARGET JD SPONSOR / TGL JD SPONSOR.
 * Dates in those columns are written as text like "15 June 2026".
 *
 * Once a month the coach (or admin) records the result for a person's current
 * step in the "LDP Progress" tab (CONFIG.LDP_SHEET_NAME), one row per check:
 *   - Berhasil → TGL JD <step> is filled and POSISI moves up (Member/Sponsor/CGL).
 *   - Belum    → a note is required; optionally a new target date is set.
 *
 * Access is the same as editing a person: coaches only for their own people.
 */

const LDP_STEPS = [
  { key: 'member', label: 'Member', target: 'targetMember', date: 'dateMember', position: 'MEMBER', rank: 1 },
  { key: 'sponsor', label: 'Sponsor', target: 'targetSponsor', date: 'dateSponsor', position: 'SPONSOR', rank: 2 },
  { key: 'cgl', label: 'CGL', target: 'targetCgl', date: 'dateCgl', position: 'CGL', rank: 3 },
  { key: 'coach', label: 'Coach', target: 'targetCoach', date: 'dateCoach', position: '', rank: 4 },
  { key: 'tl', label: 'Team Leader', target: 'targetTl', date: 'dateTl', position: '', rank: 5 },
];

const POSITION_RANK_ = { VIP: 0, SIMPATISAN: 0, MEMBER: 1, SPONSOR: 2, CGL: 3, COACH: 4, 'TEAM LEADER': 5 };

const LDP_HEADERS = ['TANGGAL UPDATE', 'COACH', 'NAMA', 'NIJ', 'TAHAP', 'TARGET', 'HASIL', 'TGL TERCAPAI',
  'TARGET BARU', 'CATATAN', 'OLEH', 'DICATAT PADA'];

const LDP_RESULTS = { berhasil: 'Berhasil', belum: 'Belum' };

/** Called from the browser: the progress history this session may see. */
function getLdpProgress(token) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  return ldpPayload_(ss, session);
}

/**
 * Saves targets / realization dates and (optionally) this month's result for one person.
 * payload = {
 *   id, name,                                    // row + name as shown (guards against moved rows)
 *   targets: { member: 'yyyy-mm-dd' | '', … },   // optional, per step key
 *   dates:   { member: 'yyyy-mm-dd' | '', … },   // optional, per step key
 *   progress: { step, result: 'berhasil' | 'belum', date: 'yyyy-mm-dd', note, newTarget: 'yyyy-mm-dd' }  // optional
 * }
 */
function saveLdp(token, payload) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  requirePerm_(session, 'edit');
  if (!payload || !payload.id || !payload.name) throw new Error('Data LDP tidak valid.');

  return withScriptLock_(function () {
    const people = parseDatabase_(getDatabaseSheet_(ss).getDataRange().getDisplayValues()).people;
    const person = people.filter(function (p) { return p.id === Math.floor(Number(payload.id)); })[0];
    if (!person || norm_(person.name) !== norm_(payload.name)) {
      throw new Error('Data di baris ' + payload.id + ' sudah berubah. Refresh dulu lalu coba lagi.');
    }

    const update = { id: person.id, name: person.name };
    ['targets', 'dates'].forEach(function (group) {
      const values = payload[group] || {};
      LDP_STEPS.forEach(function (step) {
        if (!Object.prototype.hasOwnProperty.call(values, step.key)) return;
        update[group === 'targets' ? step.target : step.date] = ldpDateCell_(values[step.key], step.label);
      });
    });

    let entry = null;
    const pr = payload.progress;
    if (pr) {
      const step = LDP_STEPS.filter(function (s) { return s.key === pr.step; })[0];
      if (!step) throw new Error('Tahap LDP tidak dikenal.');
      const result = String(pr.result || '');
      if (!LDP_RESULTS[result]) throw new Error('Pilih hasil: berhasil atau belum.');
      const note = String(pr.note || '').replace(/\s+/g, ' ').trim().slice(0, 1000);
      const achieved = result === 'berhasil' ? toIsoDate_(pr.date) : '';
      const newTarget = result === 'belum' && pr.newTarget ? toIsoDate_(pr.newTarget) : '';
      if (result === 'berhasil' && !achieved) throw new Error('Isi tanggal tercapai.');
      if (result === 'belum' && note.length < 3) throw new Error('Tulis catatan kenapa belum tercapai.');
      if (pr.newTarget && !newTarget) throw new Error('Target baru tidak valid.');

      if (result === 'berhasil') {
        update[step.date] = longDate_(achieved);
        // Move POSISI up (never down) for the steps that have a position.
        if (step.position && step.rank > (POSITION_RANK_[person.position] || 0)) update.position = step.position;
      } else if (newTarget) {
        update[step.target] = longDate_(newTarget);
      }
      entry = {
        step: step, result: result, achieved: achieved, newTarget: newTarget, note: note,
        target: toIsoDate_(person[step.target]) || person[step.target] || '',
      };
    }

    const res = applyPersonUpdate_(ss, session, update);
    if (entry) appendLdpProgress_(ss, session, person, entry);

    return Object.assign(ldpPayload_(ss, session), {
      changes: res.changes,
      data: buildDashboard_(ss, session),
    });
  });
}

// ───────────────────────────── helpers ─────────────────────────────

const MONTHS_EN_ = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August',
  'September', 'October', 'November', 'December'];

/** "2026-06-15" → "15 June 2026" (the style the database already uses). */
function longDate_(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
  return m ? (+m[3]) + ' ' + MONTHS_EN_[+m[2] - 1] + ' ' + m[1] : '';
}

/** Validates a yyyy-mm-dd from the browser and returns the cell text ('' clears). */
function ldpDateCell_(value, label) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return '';
  const iso = toIsoDate_(raw);
  if (!iso) throw new Error('Tanggal untuk ' + label + ' tidak valid.');
  return longDate_(iso);
}

function getLdpSheet_(ss, create) {
  let sheet = ss.getSheetByName(CONFIG.LDP_SHEET_NAME);
  if (!sheet && create) sheet = ss.insertSheet(CONFIG.LDP_SHEET_NAME, ss.getSheets().length);
  if (sheet && create && sheet.getLastRow() === 0) {
    sheet.appendRow(LDP_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, LDP_HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

function appendLdpProgress_(ss, session, person, e) {
  const sheet = getLdpSheet_(ss, true);
  const now = new Date();
  const row = [
    Utilities.formatDate(now, CONFIG.TIME_ZONE, 'yyyy-MM-dd'), person.coach, person.name, person.nij,
    e.step.label, e.target, LDP_RESULTS[e.result], e.achieved, e.newTarget, e.note,
    session.username, Utilities.formatDate(now, CONFIG.TIME_ZONE, 'yyyy-MM-dd HH:mm'),
  ];
  const range = sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length);
  range.setNumberFormat('@');
  range.setValues([row]);
}

/** Progress rows grouped per person ("coach|name" key), newest first, limited to what the session may see. */
function ldpPayload_(ss, session) {
  const sheet = getLdpSheet_(ss, false);
  const entries = [];
  if (sheet && sheet.getLastRow() > 1) {
    const values = sheet.getDataRange().getDisplayValues();
    const header = values[0].map(norm_);
    const c = {};
    LDP_HEADERS.forEach(function (h) { c[h] = header.indexOf(h); });
    const get = function (r, h) { return c[h] < 0 ? '' : String(r[c[h]] || '').trim(); };
    for (let i = 1; i < values.length; i++) {
      const r = values[i];
      const coach = titleCase_(get(r, 'COACH'));
      if (!get(r, 'NAMA') || !canAccessCoach_(session, coach)) continue;
      const step = LDP_STEPS.filter(function (s) { return norm_(s.label) === norm_(get(r, 'TAHAP')); })[0];
      entries.push({
        date: toIsoDate_(get(r, 'TANGGAL UPDATE')), coach: coach, name: get(r, 'NAMA'), nij: get(r, 'NIJ'),
        step: step ? step.key : '', stepLabel: get(r, 'TAHAP'), target: get(r, 'TARGET'),
        result: norm_(get(r, 'HASIL')) === 'BERHASIL' ? 'berhasil' : 'belum',
        achieved: get(r, 'TGL TERCAPAI'), newTarget: get(r, 'TARGET BARU'), note: get(r, 'CATATAN'),
        by: get(r, 'OLEH'), at: get(r, 'DICATAT PADA'),
      });
    }
  }
  entries.sort(function (a, b) { return (b.date + b.at).localeCompare(a.date + a.at); });
  return {
    steps: LDP_STEPS.map(function (s) {
      return { key: s.key, label: s.label, target: s.target, date: s.date, position: s.position, rank: s.rank };
    }),
    progress: entries,
  };
}
