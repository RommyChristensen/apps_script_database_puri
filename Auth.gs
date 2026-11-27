/**
 * Coach login for the dashboard.
 *
 * Accounts live in the "Authentication" tab (CONFIG.AUTH_SHEET_NAME):
 *
 *   USERNAME | PASSWORD | COACH | ROLE
 *   angie    | rahasia1 | Angie | COACH     → sees only people whose COACH is "Angie"
 *   rommy    | ••••••   |       | ADMIN     → sees every coach
 *   tl       | ••••••   | Rommy Christensen | TL → sees every coach whose rows
 *                                     have this name in the database's TL column
 *
 * - Type a password in plain text. After its first successful login it is
 *   replaced by a salted SHA-256 hash (sha256$<salt>$<hex>), so the tab never
 *   keeps readable passwords. To reset a password, type a new one over it.
 * - Clearing PASSWORD, setting ROLE to NONAKTIF or STATUS to NONAKTIF disables an account.
 * - Changing a password (or role/coach) signs that user out everywhere.
 * - Optional columns (added automatically by the "Kelola pengguna" page, see Users.gs):
 *   CATATAN | STATUS (AKTIF/NONAKTIF) | IZIN (empty = role default, "-" = none,
 *   else e.g. "edit,restructure") | WAJIB GANTI (YA = must set a new password
 *   before using the app) | TERAKHIR LOGIN | DIPERBARUI OLEH
 * - Logging in with the default password (set by an admin) also forces a change.
 *
 * Sessions are random tokens kept in the script cache for SESSION_HOURS
 * (sliding: every request extends it). After MAX_FAILED wrong passwords a
 * username is locked for LOCK_MINUTES.
 */

const AUTH = {
  SESSION_HOURS: 6, // CacheService maximum
  MAX_FAILED: 5,
  LOCK_MINUTES: 15,
  HEADERS: ['USERNAME', 'PASSWORD', 'COACH', 'ROLE'],
  ALL_HEADERS: ['USERNAME', 'PASSWORD', 'COACH', 'ROLE', 'CATATAN', 'STATUS', 'IZIN', 'WAJIB GANTI', 'TERAKHIR LOGIN',
    'DIPERBARUI OLEH'],
  MIN_PASSWORD: 8,
};

/**
 * Feature permissions. `roles`: who can be given it; `def`: who has it when IZIN
 * is empty. ADMIN always has everything (so an admin can never lock themselves out).
 */
const PERMS = [
  { key: 'edit', label: 'Edit data jemaat & LDP', roles: ['COACH', 'TL', 'ADMIN'], def: ['COACH', 'TL', 'ADMIN'] },
  { key: 'restructure', label: 'Pindah CG, multiplikasi, info & lokasi CG', roles: ['COACH', 'TL', 'ADMIN'], def: ['COACH', 'TL', 'ADMIN'] },
  { key: 'delete', label: 'Hapus data jemaat (diarsipkan)', roles: ['COACH', 'TL', 'ADMIN'], def: ['ADMIN'] },
  { key: 'visi', label: 'Isi Visi & SWOT coach', roles: ['TL', 'ADMIN'], def: ['TL', 'ADMIN'] },
  { key: 'events', label: 'Isi jadwal acara', roles: ['COACH', 'TL', 'ADMIN'], def: ['COACH', 'TL', 'ADMIN'] },
  { key: 'users', label: 'Kelola pengguna', roles: ['TL', 'ADMIN'], def: ['TL', 'ADMIN'] },
];

/** Effective permission keys of an account (or session-like { role, permsRaw }). */
function permsOf_(role, raw) {
  if (role === 'ADMIN') return PERMS.map(function (p) { return p.key; });
  raw = String(raw == null ? '' : raw).trim().toLowerCase();
  const list = raw === '' ? null : raw === '-' ? [] : raw.split(/[\s,;]+/);
  return PERMS.filter(function (p) {
    if (p.roles.indexOf(role) < 0) return false;
    return list ? list.indexOf(p.key) >= 0 : p.def.indexOf(role) >= 0;
  }).map(function (p) { return p.key; });
}

function hasPerm_(session, key) {
  return (session.perms || permsOf_(session.role, '')).indexOf(key) >= 0;
}

/** Throws a friendly error when the session lacks a permission. */
function requirePerm_(session, key) {
  if (hasPerm_(session, key)) return;
  const p = PERMS.filter(function (x) { return x.key === key; })[0];
  throw new Error('Akun kamu tidak punya izin "' + (p ? p.label : key) + '". Hubungi admin atau TL kamu.');
}

/** Thrown for anything that should send the browser back to the login screen. */
function authError_(msg) {
  return new Error('AUTH: ' + msg);
}

function getAuthSheet_(ss) {
  return ss.getSheetByName(CONFIG.AUTH_SHEET_NAME);
}

/** Reads all accounts: [{ row, username, stored, coach, role }]. */
function readAccounts_(ss) {
  const sheet = getAuthSheet_(ss);
  if (!sheet) throw new Error('Tab "' + CONFIG.AUTH_SHEET_NAME + '" belum ada. Buat tab itu dengan kolom ' +
    AUTH.HEADERS.join(', ') + '.');
  const values = sheet.getDataRange().getDisplayValues();
  const header = (values[0] || []).map(norm_);
  const c = {};
  AUTH.ALL_HEADERS.forEach(function (h) { c[h] = header.indexOf(h); });
  if (c.USERNAME < 0 || c.PASSWORD < 0) {
    throw new Error('Tab "' + CONFIG.AUTH_SHEET_NAME + '" harus punya kolom USERNAME dan PASSWORD di baris 1.');
  }
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const username = String(values[i][c.USERNAME] || '').trim().toLowerCase();
    if (!username) continue;
    out.push({
      row: i + 1,
      passwordCol: c.PASSWORD + 1,
      username: username,
      stored: String(values[i][c.PASSWORD] || ''),
      coach: c.COACH >= 0 ? titleCase_(String(values[i][c.COACH] || '').trim()) : '',
      role: c.ROLE >= 0 ? (norm_(values[i][c.ROLE]) || 'COACH') : 'COACH',
      status: c.STATUS >= 0 ? (norm_(values[i][c.STATUS]) || 'AKTIF') : 'AKTIF',
      permsRaw: c.IZIN >= 0 ? String(values[i][c.IZIN] || '').trim() : '',
      mustChange: c['WAJIB GANTI'] >= 0 ? /^(YA|Y|YES|TRUE|1)$/i.test(String(values[i][c['WAJIB GANTI']] || '').trim()) : false,
      note: c.CATATAN >= 0 ? String(values[i][c.CATATAN] || '').trim() : '',
      lastLogin: c['TERAKHIR LOGIN'] >= 0 ? String(values[i][c['TERAKHIR LOGIN']] || '').trim() : '',
      updatedBy: c['DIPERBARUI OLEH'] >= 0 ? String(values[i][c['DIPERBARUI OLEH']] || '').trim() : '',
    });
  }
  return out;
}

function sha256Hex_(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); })
    .join('');
}

function hashPassword_(password, salt) {
  return 'sha256$' + salt + '$' + sha256Hex_(salt + ':' + password);
}

/** Compares a typed password with the stored value (hash or plain text). */
function passwordMatches_(typed, stored) {
  if (!stored) return false;
  const m = /^sha256\$([^$]+)\$[0-9a-f]{64}$/.exec(stored);
  const candidate = m ? hashPassword_(typed, m[1]) : typed;
  return safeEqual_(candidate, stored);
}

/** String comparison that doesn't stop at the first different character. */
function safeEqual_(a, b) {
  let diff = a.length ^ b.length;
  for (let i = 0; i < Math.max(a.length, b.length); i++) diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}

/** Fingerprint of an account's current credentials; a change ends its sessions. */
function accountStamp_(acc) {
  return sha256Hex_([acc.stored, acc.coach, acc.role].join('|')).slice(0, 16);
}

function isUsable_(acc) {
  return acc && acc.stored && acc.role !== 'NONAKTIF' && acc.status !== 'NONAKTIF' && (acc.role === 'ADMIN' || acc.coach);
}

/** True when the typed password is the current default password (see Users.gs). */
function isDefaultPassword_(typed) {
  const stored = PropertiesService.getScriptProperties().getProperty('AUTH_DEFAULT_PW');
  return !!stored && passwordMatches_(typed, stored);
}

/** Called from the login form. Returns { token, user } or throws. */
function login(username, password) {
  username = String(username || '').trim().toLowerCase();
  password = String(password || '');
  if (!username || !password) throw new Error('Isi username dan password.');

  const cache = CacheService.getScriptCache();
  const failKey = 'fail:' + username;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= AUTH.MAX_FAILED) {
    throw new Error('Terlalu banyak percobaan. Coba lagi dalam ' + AUTH.LOCK_MINUTES + ' menit.');
  }

  const ss = getSpreadsheet_();
  let acc = readAccounts_(ss).filter(function (a) { return a.username === username; })[0];
  if (!isUsable_(acc) || !passwordMatches_(password, acc.stored)) {
    cache.put(failKey, String(fails + 1), AUTH.LOCK_MINUTES * 60);
    throw new Error('Username atau password salah.');
  }
  cache.remove(failKey);

  // First login with a plain-text password: replace it with a salted hash.
  if (!/^sha256\$/.test(acc.stored)) {
    const hashed = hashPassword_(password, Utilities.getUuid().replace(/-/g, '').slice(0, 16));
    const cell = getAuthSheet_(ss).getRange(acc.row, acc.passwordCol);
    cell.setNumberFormat('@');
    cell.setValue(hashed);
    acc = Object.assign({}, acc, { stored: hashed });
  }

  const mustChange = acc.mustChange || isDefaultPassword_(password);
  try { setAuthCells_(ss, acc.row, { 'TERAKHIR LOGIN': Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd HH:mm') }); } catch (e) { /* not essential */ }

  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const session = { username: acc.username, coach: acc.coach, role: acc.role, stamp: accountStamp_(acc), mustChange: mustChange };
  cache.put('session:' + token, JSON.stringify(session), AUTH.SESSION_HOURS * 3600);
  session.perms = permsOf_(acc.role, acc.permsRaw);
  if (session.role === 'TL') session.coaches = tlCoaches_(ss, session.coach);
  return { token: token, user: publicUser_(session), mustChange: mustChange };
}

/**
 * Sets a new password for the signed-in user. When the session must change
 * its password (default password / reset by admin) the current one isn't asked
 * again. req = { current?, next }
 */
function changePassword(token, req) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss, { allowPwChange: true });
  const next = String(req && req.next || '');
  const acc = readAccounts_(ss).filter(function (a) { return a.username === session.username; })[0];
  if (!session.mustChange) confirmPassword_(ss, session, req && req.current, 'Password lama salah.');
  if (next.length < AUTH.MIN_PASSWORD) throw new Error('Password baru minimal ' + AUTH.MIN_PASSWORD + ' karakter.');
  if (next.toLowerCase() === session.username) throw new Error('Password tidak boleh sama dengan username.');
  if (isDefaultPassword_(next)) throw new Error('Password baru tidak boleh sama dengan password default.');
  if (passwordMatches_(next, acc.stored)) throw new Error('Password baru harus berbeda dari yang lama.');

  const hashed = hashPassword_(next, Utilities.getUuid().replace(/-/g, '').slice(0, 16));
  setAuthCells_(ss, acc.row, { PASSWORD: hashed, 'WAJIB GANTI': '', 'DIPERBARUI OLEH': session.username + ' · ' + Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd HH:mm') });
  SpreadsheetApp.flush();
  // Keep this session (its stamp follows the new password); other devices are signed out.
  const fresh = Object.assign({}, acc, { stored: hashed });
  const stored = { username: session.username, coach: session.coach, role: session.role, stamp: accountStamp_(fresh), mustChange: false };
  CacheService.getScriptCache().put('session:' + token, JSON.stringify(stored), AUTH.SESSION_HOURS * 3600);
  session.mustChange = false;
  return { user: publicUser_(session) };
}

/** Writes { HEADER: value } on one row of the Authentication tab (adds missing columns). */
function setAuthCells_(ss, row, values) {
  const sheet = getAuthSheet_(ss);
  const col = ensureAuthColumns_(sheet);
  Object.keys(values).forEach(function (h) {
    const cell = sheet.getRange(row, col[h] + 1);
    cell.setNumberFormat('@');
    cell.setValue(values[h]);
  });
}

/** Makes sure every AUTH.ALL_HEADERS column exists (new ones go at the end). Returns { HEADER: index }. */
function ensureAuthColumns_(sheet) {
  const width = Math.max(sheet.getLastColumn(), 1);
  const header = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(norm_);
  const col = {};
  let next = header.length;
  while (next > 0 && !header[next - 1]) next--;
  AUTH.ALL_HEADERS.forEach(function (h) {
    let j = header.indexOf(h);
    if (j < 0) {
      j = next++;
      sheet.getRange(1, j + 1).setValue(h).setFontWeight('bold');
    }
    col[h] = j;
  });
  return col;
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('session:' + String(token));
  return true;
}

function publicUser_(s) {
  return {
    username: s.username, coach: s.coach, role: s.role || 'COACH',
    isAdmin: s.role === 'ADMIN', isTl: s.role === 'TL', coaches: s.role === 'TL' ? (s.coaches || []) : undefined,
    perms: s.perms || permsOf_(s.role, ''),
  };
}

/**
 * Resolves a token to its session, re-checking the account in the sheet so a
 * changed password, role or coach takes effect immediately. Throws AUTH: errors.
 */
function requireSession_(token, ss, opts) {
  if (!token || typeof token !== 'string') throw authError_('Silakan masuk dulu.');
  const cache = CacheService.getScriptCache();
  const raw = cache.get('session:' + token);
  if (!raw) throw authError_('Sesi berakhir. Silakan masuk lagi.');
  const s = JSON.parse(raw);
  const acc = readAccounts_(ss).filter(function (a) { return a.username === s.username; })[0];
  if (!isUsable_(acc) || accountStamp_(acc) !== s.stamp) {
    cache.remove('session:' + token);
    throw authError_('Akun berubah atau dinonaktifkan. Silakan masuk lagi.');
  }
  if (s.mustChange && !(opts && opts.allowPwChange)) {
    throw new Error('PWCHANGE: Buat password baru dulu sebelum melanjutkan.');
  }
  cache.put('session:' + token, raw, AUTH.SESSION_HOURS * 3600); // sliding expiry
  s.perms = permsOf_(acc.role, acc.permsRaw);
  if (s.role === 'TL') s.coaches = tlCoaches_(ss, s.coach);
  return s;
}

/**
 * Coaches under a TL: those with the TL's name in the TL column of the
 * database. Cached for a few minutes (one read of the database otherwise).
 */
function tlCoaches_(ss, tlName) {
  const cache = CacheService.getScriptCache();
  const key = 'tl:' + norm_(tlName).replace(/\s+/g, '_').slice(0, 200);
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  const people = parseDatabase_(getDatabaseSheet_(ss).getDataRange().getDisplayValues()).people;
  const out = [];
  people.forEach(function (p) {
    if (!p.isLeader && p.coach && norm_(p.tl) === norm_(tlName) && out.indexOf(p.coach) < 0) out.push(p.coach);
  });
  cache.put(key, JSON.stringify(out), 300);
  return out;
}

/** Admins and TLs record Pertemuan / LDP about coaches and edit coach rows (TL: only their own coaches). */
function leadsCoaches_(session) {
  return session.role === 'ADMIN' || session.role === 'TL';
}

/** 'Admin' | 'TL' | 'Coach' — how this session is labelled in logs. */
function roleLabel_(session) {
  return session.role === 'ADMIN' ? 'Admin' : session.role === 'TL' ? 'TL' : 'Coach';
}

/** True when the session may see/edit people of this coach. */
function canAccessCoach_(session, coach) {
  if (session.role === 'ADMIN') return true;
  if (session.role === 'TL') {
    return (session.coaches || []).some(function (c) { return norm_(c) === norm_(coach); });
  }
  return norm_(coach) === norm_(session.coach);
}
