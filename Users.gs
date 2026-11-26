/**
 * "Kelola pengguna": accounts in the Authentication tab, managed from the app.
 *
 * - Admins manage every account and the default password.
 * - TLs (with the "users" permission) manage COACH accounts of the coaches
 *   under them, and can only give permissions they have themselves.
 * - New accounts and "Reset password" get the default password, and the user
 *   must set a new one on their first login (WAJIB GANTI = YA).
 * - Revoke = STATUS NONAKTIF (signed out at once, can be re-activated).
 *
 * The default password is kept only as a salted hash in Script Properties
 * (AUTH_DEFAULT_PW), never in the sheet. New accounts receive that same hash.
 */

const USERS = {
  ROLES: ['COACH', 'TL', 'ADMIN'],
  USERNAME: /^[a-z0-9][a-z0-9._-]{2,29}$/,
};

/** Called from the browser: the accounts this session may manage. */
function getUsers(token) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  requirePerm_(session, 'users');
  return usersPayload_(ss, session);
}

/**
 * Creates (isNew) or updates an account.
 * u = { isNew, username, role, coach, perms: null (role default) | [keys], note }
 */
function saveUser(token, u) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  requirePerm_(session, 'users');
  if (!u) throw new Error('Data pengguna tidak valid.');
  const username = String(u.username || '').trim().toLowerCase();
  const admin = session.role === 'ADMIN';
  const role = admin ? String(u.role || '').toUpperCase() : 'COACH';
  if (USERS.ROLES.indexOf(role) < 0) throw new Error('Pilih role: COACH, TL atau ADMIN.');
  const coach = titleCase_(String(u.coach || '').replace(/\s+/g, ' ').trim()).slice(0, 80);
  if (role !== 'ADMIN' && !coach) throw new Error(role === 'TL' ? 'Isi nama TL (sama dengan kolom TL di database).' : 'Pilih coach.');
  if (!admin && !canAccessCoach_(session, coach)) throw new Error('Kamu hanya bisa mengelola akun coach di bawah kamu.');
  const note = String(u.note || '').replace(/\s+/g, ' ').trim().slice(0, 200);
  const permsRaw = permsInput_(session, role, u.perms);

  return withScriptLock_(function () {
    const accounts = readAccounts_(ss);
    const now = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd HH:mm');
    const stamp = session.username + ' · ' + now;
    const sheet = getAuthSheet_(ss);
    const col = ensureAuthColumns_(sheet);
    let acc = accounts.filter(function (a) { return a.username === username; })[0];

    if (u.isNew) {
      if (!USERS.USERNAME.test(username)) throw new Error('Username 3–30 karakter: huruf kecil, angka, titik, minus atau garis bawah.');
      if (acc) throw new Error('Username "' + username + '" sudah dipakai.');
      const def = defaultPasswordHash_();
      if (!def) throw new Error('Admin belum mengatur default password. Atur dulu di halaman ini.');
      const row = Math.max(sheet.getLastRow(), 1) + 1;
      const values = { USERNAME: username, PASSWORD: def, COACH: coach, ROLE: role, CATATAN: note, STATUS: 'AKTIF',
        IZIN: permsRaw, 'WAJIB GANTI': 'YA', 'TERAKHIR LOGIN': '', 'DIPERBARUI OLEH': stamp };
      Object.keys(values).forEach(function (h) {
        const cell = sheet.getRange(row, col[h] + 1);
        cell.setNumberFormat('@');
        cell.setValue(values[h]);
      });
    } else {
      acc = manageable_(session, accounts, username);
      if (acc.username === session.username && role !== acc.role) throw new Error('Kamu tidak bisa mengubah role akunmu sendiri.');
      if (acc.role === 'ADMIN' && role !== 'ADMIN' && activeAdmins_(accounts).length <= 1 && acc.status !== 'NONAKTIF') {
        throw new Error('Ini admin aktif terakhir. Buat admin lain dulu.');
      }
      setAuthCells_(ss, acc.row, { COACH: coach, ROLE: role, CATATAN: note, IZIN: permsRaw, 'DIPERBARUI OLEH': stamp });
    }
    SpreadsheetApp.flush();
    return usersPayload_(ss, session);
  });
}

/** Revoke (active = false) or re-activate an account. */
function setUserActive(token, username, active) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  requirePerm_(session, 'users');
  return withScriptLock_(function () {
    const accounts = readAccounts_(ss);
    const acc = manageable_(session, accounts, username);
    if (!active) {
      if (acc.username === session.username) throw new Error('Kamu tidak bisa menonaktifkan akunmu sendiri.');
      if (acc.role === 'ADMIN' && activeAdmins_(accounts).length <= 1) throw new Error('Ini admin aktif terakhir. Buat admin lain dulu.');
    }
    const cells = { STATUS: active ? 'AKTIF' : 'NONAKTIF', 'DIPERBARUI OLEH': session.username + ' · ' + visitNow_() };
    if (active && acc.role === 'NONAKTIF') cells.ROLE = 'COACH'; // old-style disabled row
    setAuthCells_(ss, acc.row, cells);
    SpreadsheetApp.flush();
    if (!active) clearLock_(acc.username);
    return usersPayload_(ss, session);
  });
}

/** Sets the account back to the default password; the user must change it on next login. */
function resetUserPassword(token, username) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  requirePerm_(session, 'users');
  return withScriptLock_(function () {
    const acc = manageable_(session, readAccounts_(ss), username);
    if (acc.username === session.username) throw new Error('Untuk akunmu sendiri, pakai "Ganti password".');
    const def = defaultPasswordHash_();
    if (!def) throw new Error('Admin belum mengatur default password.');
    setAuthCells_(ss, acc.row, { PASSWORD: def, 'WAJIB GANTI': 'YA', 'DIPERBARUI OLEH': session.username + ' · ' + visitNow_() });
    SpreadsheetApp.flush();
    CacheService.getScriptCache().remove('fail:' + acc.username); // lift a lock-out too
    return usersPayload_(ss, session);
  });
}

/** Admin only: removes the account row for good. */
function deleteUser(token, username) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  if (session.role !== 'ADMIN') throw new Error('Hanya admin yang bisa menghapus akun.');
  return withScriptLock_(function () {
    const accounts = readAccounts_(ss);
    const acc = manageable_(session, accounts, username);
    if (acc.username === session.username) throw new Error('Kamu tidak bisa menghapus akunmu sendiri.');
    if (acc.role === 'ADMIN' && acc.status !== 'NONAKTIF' && activeAdmins_(accounts).length <= 1) {
      throw new Error('Ini admin aktif terakhir.');
    }
    getAuthSheet_(ss).deleteRow(acc.row);
    SpreadsheetApp.flush();
    clearLock_(acc.username);
    return usersPayload_(ss, session);
  });
}

/** Admin only: sets the default password used for new accounts and resets. */
function setDefaultPassword(token, password) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  if (session.role !== 'ADMIN') throw new Error('Hanya admin yang bisa mengatur default password.');
  password = String(password || '');
  if (password.length < AUTH.MIN_PASSWORD) throw new Error('Default password minimal ' + AUTH.MIN_PASSWORD + ' karakter.');
  const props = PropertiesService.getScriptProperties();
  props.setProperty('AUTH_DEFAULT_PW', hashPassword_(password, Utilities.getUuid().replace(/-/g, '').slice(0, 16)));
  props.setProperty('AUTH_DEFAULT_META', JSON.stringify({ by: session.username, at: visitNow_() }));
  return usersPayload_(ss, session);
}

// ───────────────────────────── helpers ─────────────────────────────

function defaultPasswordHash_() {
  return PropertiesService.getScriptProperties().getProperty('AUTH_DEFAULT_PW') || '';
}

function activeAdmins_(accounts) {
  return accounts.filter(function (a) { return a.role === 'ADMIN' && a.status !== 'NONAKTIF' && a.stored; });
}

/** Lifts a login lock-out. (A revoked account is stopped by requireSession_ re-checking STATUS.) */
function clearLock_(username) {
  CacheService.getScriptCache().remove('fail:' + username);
}

/** The account if this session may manage it (admin: all; TL: COACH accounts of their coaches). */
function manageable_(session, accounts, username) {
  const acc = accounts.filter(function (a) { return a.username === String(username || '').trim().toLowerCase(); })[0];
  if (!acc) throw new Error('Akun "' + username + '" tidak ditemukan. Refresh dulu.');
  if (session.role === 'ADMIN') return acc;
  if ((acc.role === 'COACH' || acc.role === 'NONAKTIF') && acc.coach && canAccessCoach_(session, acc.coach)) return acc;
  throw new Error('Kamu hanya bisa mengelola akun coach di bawah kamu.');
}

/**
 * IZIN cell value from the form: null → '' (role default). A list is kept only
 * with permissions the role can have; a TL can only hand out what they have.
 */
function permsInput_(session, role, perms) {
  if (perms == null) return '';
  if (!Array.isArray(perms)) throw new Error('Izin tidak valid.');
  const allowed = PERMS.filter(function (p) { return p.roles.indexOf(role) >= 0; }).map(function (p) { return p.key; });
  const want = perms.map(function (k) { return String(k).toLowerCase(); }).filter(function (k) { return allowed.indexOf(k) >= 0; });
  if (session.role !== 'ADMIN') {
    want.forEach(function (k) { if (!hasPerm_(session, k)) throw new Error('Kamu tidak bisa memberi izin yang kamu sendiri tidak punya.'); });
  }
  if (role === 'ADMIN') return '';
  const def = permsOf_(role, '');
  const same = want.length === def.length && want.every(function (k) { return def.indexOf(k) >= 0; });
  return same ? '' : (want.length ? want.join(',') : '-');
}

function usersPayload_(ss, session) {
  const admin = session.role === 'ADMIN';
  const accounts = readAccounts_(ss).filter(function (a) {
    if (admin) return true;
    return (a.role === 'COACH' || a.role === 'NONAKTIF') && a.coach && canAccessCoach_(session, a.coach);
  });
  const people = parseDatabase_(getDatabaseSheet_(ss).getDataRange().getDisplayValues()).people;
  const coaches = {}, tls = {};
  people.forEach(function (p) {
    if (p.isLeader || !p.coach || p.coach === 'Unassigned') return;
    if (canAccessCoach_(session, p.coach)) coaches[p.coach] = true;
    if (admin && p.tl) tls[p.tl] = true;
  });
  let meta = null;
  try { meta = JSON.parse(PropertiesService.getScriptProperties().getProperty('AUTH_DEFAULT_META') || 'null'); } catch (e) { meta = null; }
  return {
    users: accounts.map(function (a) {
      const role = a.role === 'NONAKTIF' ? 'COACH' : a.role;
      return {
        username: a.username, coach: a.coach, role: role,
        active: a.status !== 'NONAKTIF' && a.role !== 'NONAKTIF', hasPassword: !!a.stored,
        perms: permsOf_(role, a.permsRaw), customPerms: !!a.permsRaw && role !== 'ADMIN',
        mustChange: a.mustChange, note: a.note, lastLogin: a.lastLogin, updatedBy: a.updatedBy,
        self: a.username === session.username,
      };
    }).sort(function (x, y) {
      return USERS.ROLES.indexOf(y.role) - USERS.ROLES.indexOf(x.role) || x.coach.localeCompare(y.coach) || x.username.localeCompare(y.username);
    }),
    perms: PERMS.map(function (p) { return { key: p.key, label: p.label, roles: p.roles, def: p.def, mine: hasPerm_(session, p.key) }; }),
    coaches: Object.keys(coaches).sort(),
    tls: Object.keys(tls).sort(),
    isAdmin: admin,
    defaultSet: !!defaultPasswordHash_(),
    defaultMeta: admin ? meta : null,
  };
}
