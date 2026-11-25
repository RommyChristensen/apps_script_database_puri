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
 * - Clearing PASSWORD, or setting ROLE to NONAKTIF, disables an account.
 * - Changing a password (or role/coach) signs that user out everywhere.
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
};

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
  AUTH.HEADERS.forEach(function (h) { c[h] = header.indexOf(h); });
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
  return acc && acc.stored && acc.role !== 'NONAKTIF' && (acc.role === 'ADMIN' || acc.coach);
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

  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const session = { username: acc.username, coach: acc.coach, role: acc.role, stamp: accountStamp_(acc) };
  cache.put('session:' + token, JSON.stringify(session), AUTH.SESSION_HOURS * 3600);
  if (session.role === 'TL') session.coaches = tlCoaches_(ss, session.coach);
  return { token: token, user: publicUser_(session) };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('session:' + String(token));
  return true;
}

function publicUser_(s) {
  return {
    username: s.username, coach: s.coach, role: s.role || 'COACH',
    isAdmin: s.role === 'ADMIN', isTl: s.role === 'TL', coaches: s.role === 'TL' ? (s.coaches || []) : undefined,
  };
}

/**
 * Resolves a token to its session, re-checking the account in the sheet so a
 * changed password, role or coach takes effect immediately. Throws AUTH: errors.
 */
function requireSession_(token, ss) {
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
  cache.put('session:' + token, raw, AUTH.SESSION_HOURS * 3600); // sliding expiry
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
