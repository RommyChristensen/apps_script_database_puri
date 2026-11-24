/**
 * Tests for the server-side code (Code.gs) against an in-memory sheet.
 *   node dev/test.js [optional-export.csv]
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');
const createMockServices = require('./mock-apps-script');
const { demoValues, parseCsv } = require('./build-preview');

const code = ['Code.gs', 'Auth.gs'].map((f) => fs.readFileSync(path.join(__dirname, '..', f), 'utf8')).join('\n');

const AUTH_HEADER = ['USERNAME', 'PASSWORD', 'COACH', 'ROLE'];
const defaultAccounts = () => [AUTH_HEADER, ['admin', 'admin-pass', '', 'ADMIN']];

/**
 * Loads Code.gs + Auth.gs against a mock spreadsheet. `api` is signed in as
 * `opts.as` (default: an admin) so the older tests see everything; `raw`
 * exposes the real server functions for the auth tests.
 */
function boot(values, opts) {
  opts = Object.assign({}, opts);
  const accounts = opts.accounts || defaultAccounts();
  const mocks = createMockServices(values, Object.assign({}, opts, { sheets: { Authentication: accounts } }));
  const ctx = {
    SpreadsheetApp: mocks.SpreadsheetApp, LockService: mocks.LockService, Session: mocks.Session,
    CacheService: mocks.CacheService, Utilities: mocks.Utilities, HtmlService: {}, console,
  };
  vm.createContext(ctx);
  vm.runInContext(code + '\nthis.raw = { getDashboardData, updatePerson, login, logout, MILESTONES };', ctx);
  const [user, pass] = opts.as || ['admin', 'admin-pass'];
  const token = ctx.raw.login(user, pass).token;
  mocks.writes.length = 0; // ignore the password-hash write from this setup login
  const api = {
    getDashboardData: () => ctx.raw.getDashboardData(token),
    updatePerson: (p) => ctx.raw.updatePerson(token, p),
  };
  return { api, raw: ctx.raw, token, mocks, grid: values, auth: accounts };
}
const clone = (v) => JSON.parse(JSON.stringify(v));
let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('  ✓ ' + name); }
  catch (e) { console.log('  ✗ ' + name + '\n    ' + e.message); process.exitCode = 1; }
}

console.log('Code.gs (demo data)');
const base = demoValues();
const H = base[3];
const C = (h) => H.indexOf(h); // 0-based col by header
const SUBC = (h) => base[4].indexOf(h);

test('parses 5 coaches from banner rows', () => {
  const { api } = boot(clone(base));
  const d = api.getDashboardData();
  assert.deepStrictEqual(clone(d.coaches.map((c) => c.name)), ['Angie', 'Lenita', 'Carla', 'Meysia', 'David']);
  assert.ok(d.positionOptions.includes('SPONSOR'));
  assert.strictEqual(d.editableFields.filter((f) => f.section === 'jenjang').length, 7);
  assert.strictEqual(d.editableFields.filter((f) => f.section === 'pribadi').length, 11);
});

test('updates milestones, position and jenjang; writes only changed cells; logs the edit', () => {
  const s = boot(clone(base));
  const p = s.api.getDashboardData().people.find((x) => x.position === 'MEMBER' && !x.milestones.msj3 && x.milestones.msj1);
  const res = s.api.updatePerson({
    id: p.id, name: p.name,
    milestones: { msj3: true, msj1: true /* unchanged */ },
    position: 'SPONSOR', dateSponsor: 'November 2026', targetCgl: '  Juni   2027 ',
  });
  assert.deepStrictEqual(clone(res.changes.map((c) => c.label)), ['MSJ 3', 'Posisi', 'Tgl jadi Sponsor', 'Target CGL']);
  const row = s.grid[p.id - 1];
  assert.strictEqual(row[SUBC('MSJ 3')], '✔');
  assert.strictEqual(row[C('POSISI')], 'SPONSOR');
  assert.strictEqual(row[C('TGL JD SPONSOR')], 'November 2026');
  assert.strictEqual(row[C('TARGET CGL')], 'Juni 2027');
  assert.strictEqual(s.mocks.writes.filter((w) => w.sheet === 'Database (demo)' || w.sheet === 'Database').length, 4);
  const updated = res.data.people.find((x) => x.id === p.id);
  assert.strictEqual(updated.position, 'SPONSOR');
  assert.strictEqual(updated.milestones.msj3, true);
  const log = s.mocks.spreadsheet.getSheetByName('Edit Log');
  assert.ok(log, 'log sheet created');
  assert.strictEqual(s.mocks.spreadsheet.getSheets()[0].getName() !== 'Edit Log', true, 'database stays first tab');
  assert.strictEqual(log._grid.length, 2);
  assert.match(log._grid[1][4], /MSJ 3: ✗ → ✔/);
});

test('unchecking writes "-"', () => {
  const s = boot(clone(base));
  const p = s.api.getDashboardData().people.find((x) => x.milestones.cgt1 && !x.isCoach);
  s.api.updatePerson({ id: p.id, name: p.name, milestones: { cgt1: false } });
  assert.strictEqual(s.grid[p.id - 1][SUBC('CGT 1')], '-');
});

test('no-op save writes nothing and does not log', () => {
  const s = boot(clone(base));
  const p = s.api.getDashboardData().people.find((x) => !x.isCoach);
  const res = s.api.updatePerson({ id: p.id, name: p.name, position: p.position, milestones: clone(p.milestones) });
  assert.strictEqual(res.changes.length, 0);
  assert.strictEqual(s.mocks.writes.length, 0);
  assert.strictEqual(s.mocks.spreadsheet.getSheetByName('Edit Log'), null);
});

test('rejects when the row no longer holds that person', () => {
  const s = boot(clone(base));
  const p = s.api.getDashboardData().people.find((x) => !x.isCoach);
  assert.throws(() => s.api.updatePerson({ id: p.id, name: 'Someone Else', position: 'CGL' }), /nama tidak cocok/);
  assert.strictEqual(s.mocks.writes.length, 0);
});

test('rejects rows outside the data and unknown positions', () => {
  const s = boot(clone(base));
  assert.throws(() => s.api.updatePerson({ id: 4, name: 'NAMA', position: 'CGL' }), /di luar area data/);
  const p = s.api.getDashboardData().people.find((x) => !x.isCoach);
  assert.throws(() => s.api.updatePerson({ id: p.id, name: p.name, position: 'BOSS' }), /tidak dikenal/);
});

test('refuses to overwrite formula cells', () => {
  const vals = clone(base);
  const tmp = boot(clone(base));
  const p = tmp.api.getDashboardData().people.find((x) => !x.isCoach);
  const s = boot(vals, { formulas: { [p.id + ',' + (C('POSISI') + 1)]: '=X1' } });
  assert.throws(() => s.api.updatePerson({ id: p.id, name: p.name, position: p.position === 'CGL' ? 'MEMBER' : 'CGL' }), /formula/);
  assert.strictEqual(s.mocks.writes.length, 0);
});

test('real checkbox cells stay booleans', () => {
  const vals = clone(base);
  const tmp = boot(clone(base)).api.getDashboardData().people.find((x) => !x.isCoach && !x.milestones.cgt3);
  vals[tmp.id - 1][SUBC('CGT 3')] = false;
  const s = boot(vals);
  s.api.updatePerson({ id: tmp.id, name: tmp.name, milestones: { cgt3: true } });
  assert.strictEqual(s.grid[tmp.id - 1][SUBC('CGT 3')], true);
});

test('edits personal data fields', () => {
  const s = boot(clone(base));
  const p = s.api.getDashboardData().people.find((x) => !x.isCoach);
  const addr = 'Jl. Panjang ' + 'x'.repeat(400);
  const res = s.api.updatePerson({
    id: p.id, name: p.name,
    gender: p.gender === 'Male' ? 'Female' : 'Male', phone: '0812-3456-7890', email: 'baru@example.com',
    nij: '24009999', birthCity: 'Bandung', birthDate: '5 March 2004', domicile: 'Kembangan Utara',
    occupation: 'Kerja Freelance', school: 'UPH Karawaci', ministry: 'Worship Team', address: addr,
  });
  assert.strictEqual(res.changes.length, 11);
  const row = s.grid[p.id - 1];
  assert.strictEqual(row[C('NO HP')], '0812-3456-7890', 'first NO HP column, leading zero kept');
  assert.strictEqual(row[C('EMAIL')], 'baru@example.com');
  assert.strictEqual(row[C('TTL')], '5 March 2004');
  assert.strictEqual(row[C('ALAMAT')].length, 300, 'address capped at 300 chars');
  const fresh = res.data.people.find((x) => x.id === p.id);
  assert.strictEqual(fresh.birthCity, 'Bandung');
  assert.ok(fresh.age >= 21, 'age recomputed from new birth date');
});

test('rejects invalid personal data without writing anything', () => {
  const s = boot(clone(base));
  const p = s.api.getDashboardData().people.find((x) => !x.isCoach);
  assert.throws(() => s.api.updatePerson({ id: p.id, name: p.name, ministry: 'Usher', email: 'bukan-email' }), /Email .* tidak valid/);
  assert.throws(() => s.api.updatePerson({ id: p.id, name: p.name, phone: 'abc' }), /No HP .* tidak valid/);
  assert.throws(() => s.api.updatePerson({ id: p.id, name: p.name, gender: 'Other' }), /Gender harus/);
  assert.throws(() => s.api.updatePerson({ id: p.id, name: p.name, nij: '24A' }), /hanya boleh angka/);
  assert.strictEqual(s.mocks.writes.length, 0);
});

test('clearing a personal field writes an empty cell', () => {
  const s = boot(clone(base));
  const p = s.api.getDashboardData().people.find((x) => !x.isCoach && x.school);
  s.api.updatePerson({ id: p.id, name: p.name, school: '' });
  assert.strictEqual(s.grid[p.id - 1][C('KAMPUS')], '');
});

// Fixtures for the two layouts seen in the real sheet.
const NOTE_COL = C('KETERANGAN');
const TCOACH_COL = C('TARGET JD COACH');
const noteRows = () => base.map((r, i) => i).filter((i) => i > 4 && base[i][NOTE_COL] && base[i][2] && base[i][1] !== '-');

/** Old layout: KETERANGAN holds ages, the notes sit one column left under TARGET JD COACH. */
function shiftedLayout() {
  const v = clone(base);
  noteRows().forEach((i) => { v[i][TCOACH_COL] = v[i][NOTE_COL]; v[i][NOTE_COL] = '22'; });
  return v;
}
/** Current layout (Nov 2026): KETERANGAN column deleted, notes left under TARGET JD COACH. */
function noKeteranganLayout() {
  const v = clone(base);
  noteRows().forEach((i) => { v[i][TCOACH_COL] = v[i][NOTE_COL]; });
  v.forEach((row) => row.splice(NOTE_COL, 1));
  return v;
}

console.log('Code.gs (sheet layouts)');

test('old layout: per-row notes in TARGET JD COACH are locked and never wiped', () => {
  const s = boot(shiftedLayout());
  const d = s.api.getDashboardData();
  const p = d.people.find((x) => x.targetCoachHoldsNotes);
  assert.ok(p && p.notes, 'note recovered');
  const before = s.grid[p.id - 1][TCOACH_COL];
  const res = s.api.updatePerson({ id: p.id, name: p.name, targetCoach: '', milestones: { msj3: !p.milestones.msj3 } });
  assert.deepStrictEqual(clone(res.changes.map((c) => c.label)), ['MSJ 3']);
  assert.throws(() => s.api.updatePerson({ id: p.id, name: p.name, targetCoach: 'Juni 2027' }), /masih berisi catatan/);
  assert.throws(() => s.api.updatePerson({ id: p.id, name: p.name, notes: 'baru' }), /masih berisi catatan/);
  assert.strictEqual(s.grid[p.id - 1][TCOACH_COL], before);
});

test('no KETERANGAN column: TARGET JD COACH is read as notes for every row', () => {
  const s = boot(noKeteranganLayout());
  const d = s.api.getDashboardData();
  const withNotes = d.people.filter((x) => x.notes);
  assert.strictEqual(withNotes.length, noteRows().length);
  assert.ok(d.people.every((x) => !x.targetCoachHoldsNotes && !x.targetCoach));
  const keys = d.editableFields.map((f) => f.key);
  assert.ok(keys.includes('notes'), 'notes editable');
  assert.ok(!keys.includes('targetCoach'), 'no separate Target Coach column to edit');
  assert.ok(d.warnings.some((w) => /dibaca sebagai KETERANGAN/.test(w)));
});

test('no KETERANGAN column: editing notes writes into that column', () => {
  const s = boot(noKeteranganLayout());
  const p = s.api.getDashboardData().people.find((x) => !x.isCoach && !x.notes);
  const res = s.api.updatePerson({ id: p.id, name: p.name, notes: 'Mulai rutin CG,  perlu follow up MSJ' });
  assert.deepStrictEqual(clone(res.changes.map((c) => c.label)), ['Catatan / Keterangan']);
  assert.strictEqual(s.grid[p.id - 1][TCOACH_COL], 'Mulai rutin CG, perlu follow up MSJ');
  assert.strictEqual(res.data.people.find((x) => x.id === p.id).notes, 'Mulai rutin CG, perlu follow up MSJ');
});

test('fields whose column is missing are not offered for editing', () => {
  const v = clone(base);
  const emailCol = C('EMAIL');
  v.forEach((row) => row.splice(emailCol, 1));
  const d = boot(v).api.getDashboardData();
  assert.ok(!d.editableFields.some((f) => f.key === 'email'));
  assert.ok(d.editableFields.some((f) => f.key === 'phone'));
});

test('team leader row is not grouped under a coach; attendance is parsed', () => {
  const v = clone(base);
  v[3] = v[3].concat(['KEHADIRAN', '% KEHADIRAN']);
  v[4] = v[4].concat(['', '']);
  const tl = Array(v[3].length).fill('');
  tl[2] = 'Rommy Christensen'; tl[1] = '-'; tl[C('POSISI')] = 'TEAM LEADER';
  v.splice(5, 0, tl);
  const target = v.findIndex((r, i) => i > 6 && r[C('POSISI')] === 'MEMBER');
  v[target][v[3].length - 2] = '15'; v[target][v[3].length - 1] = '94%';
  const d = boot(v).api.getDashboardData();
  assert.deepStrictEqual(clone(d.teamLeaders), ['Rommy Christensen']);
  assert.ok(!d.coaches.some((c) => c.name === ''), 'no empty coach group');
  const r = d.people.find((x) => x.name === 'Rommy Christensen');
  assert.ok(r.isLeader && r.isCoach && r.coach === '');
  const m = d.people.find((x) => x.id === target + 1);
  assert.strictEqual(m.attendance, 15); assert.strictEqual(m.attendancePct, 94);
});

console.log('Auth.gs');
const coachAccounts = () => [AUTH_HEADER,
  ['admin', 'admin-pass', '', 'ADMIN'],
  ['angie', 'angie-pass', 'Angie', 'COACH'],
  ['lenita', 'lenita-pass', 'lenita', ''],
  ['old', 'x', 'Carla', 'NONAKTIF'],
  ['nopass', '', 'Carla', 'COACH']];

test('login: wrong password, unknown user, disabled and empty-password accounts are refused', () => {
  const s = boot(clone(base), { accounts: coachAccounts() });
  assert.throws(() => s.raw.login('angie', 'nope'), /salah/);
  assert.throws(() => s.raw.login('ghost', 'x'), /salah/);
  assert.throws(() => s.raw.login('old', 'x'), /salah/);
  assert.throws(() => s.raw.login('nopass', ''), /Isi username/);
  assert.ok(s.raw.login('ANGIE ', 'angie-pass').token, 'username is case-insensitive and trimmed');
});

test('plain-text password is replaced by a salted hash after the first login, and still works', () => {
  const s = boot(clone(base), { accounts: coachAccounts() });
  s.raw.login('angie', 'angie-pass');
  const stored = s.auth[2][1];
  assert.match(stored, /^sha256\$[0-9a-f]{16}\$[0-9a-f]{64}$/);
  assert.ok(!stored.includes('angie-pass'));
  assert.ok(s.raw.login('angie', 'angie-pass').token);
  assert.throws(() => s.raw.login('angie', stored), /salah/, 'the hash itself is not a password');
});

test('5 wrong passwords lock the username', () => {
  const s = boot(clone(base), { accounts: coachAccounts() });
  for (let i = 0; i < 5; i++) assert.throws(() => s.raw.login('angie', 'bad'), /salah/);
  assert.throws(() => s.raw.login('angie', 'angie-pass'), /Terlalu banyak/);
});

test('a coach only receives their own people; admin receives everyone', () => {
  const s = boot(clone(base), { accounts: coachAccounts(), as: ['angie', 'angie-pass'] });
  const d = s.api.getDashboardData();
  assert.ok(d.people.length > 0);
  assert.ok(d.people.every((p) => p.coach === 'Angie'), 'only Angie');
  assert.deepStrictEqual(clone(d.coaches.map((c) => c.name)), ['Angie']);
  assert.strictEqual(d.spreadsheetUrl, '', 'no spreadsheet link for coaches');
  assert.deepStrictEqual(clone(d.user), { username: 'angie', coach: 'Angie', isAdmin: false });
  const all = s.raw.getDashboardData(s.raw.login('admin', 'admin-pass').token);
  assert.strictEqual(all.coaches.length, 5);
  assert.ok(all.people.length > d.people.length);
  assert.ok(all.user.isAdmin && all.spreadsheetUrl);
});

test('coach names in the account are matched case-insensitively', () => {
  const s = boot(clone(base), { accounts: coachAccounts(), as: ['lenita', 'lenita-pass'] });
  assert.ok(s.api.getDashboardData().people.every((p) => p.coach === 'Lenita'));
});

test('a coach cannot edit people of another coach, or coach rows', () => {
  const s = boot(clone(base), { accounts: coachAccounts(), as: ['angie', 'angie-pass'] });
  const all = s.raw.getDashboardData(s.raw.login('admin', 'admin-pass').token).people;
  const other = all.find((p) => p.coach === 'Carla' && !p.isCoach);
  assert.throws(() => s.api.updatePerson({ id: other.id, name: other.name, position: 'CGL' }), /tidak punya akses/);
  const coachRow = all.find((p) => p.coach === 'Angie' && p.isCoach);
  assert.throws(() => s.api.updatePerson({ id: coachRow.id, name: coachRow.name, gender: 'Male' }), /tidak punya akses/);
  const mine = all.find((p) => p.coach === 'Angie' && !p.isCoach && p.position !== 'CGL');
  const res = s.api.updatePerson({ id: mine.id, name: mine.name, position: 'CGL' });
  assert.strictEqual(res.changes.length, 1);
  assert.ok(res.data.people.every((p) => p.coach === 'Angie'), 'refreshed data is still filtered');
  const log = s.mocks.spreadsheet.getSheetByName('Edit Log');
  assert.match(log._grid[1][1], /angie \(coach Angie\)/);
});

test('missing, unknown and logged-out tokens are rejected with AUTH:', () => {
  const s = boot(clone(base), { accounts: coachAccounts() });
  assert.throws(() => s.raw.getDashboardData(''), /^Error: AUTH:/);
  assert.throws(() => s.raw.getDashboardData('made-up'), /AUTH:/);
  assert.throws(() => s.raw.updatePerson('made-up', { id: 7, name: 'x' }), /AUTH:/);
  const t = s.raw.login('angie', 'angie-pass').token;
  s.raw.logout(t);
  assert.throws(() => s.raw.getDashboardData(t), /AUTH:/);
});

test('changing the password, coach or role in the tab ends existing sessions', () => {
  const s = boot(clone(base), { accounts: coachAccounts() });
  const t = s.raw.login('angie', 'angie-pass').token;
  assert.ok(s.raw.getDashboardData(t));
  s.auth[2][2] = 'Carla'; // moved to another coach
  assert.throws(() => s.raw.getDashboardData(t), /AUTH:/);
  s.auth[2][2] = 'Angie';
  const t2 = s.raw.login('angie', 'angie-pass').token;
  s.auth[2][1] = 'new-pass'; // password reset
  assert.throws(() => s.raw.getDashboardData(t2), /AUTH:/);
  assert.ok(s.raw.login('angie', 'new-pass').token);
});

test('a missing Authentication tab gives a clear setup error', () => {
  const mocks = createMockServices(clone(base));
  const ctx = { SpreadsheetApp: mocks.SpreadsheetApp, LockService: mocks.LockService, Session: mocks.Session,
    CacheService: mocks.CacheService, Utilities: mocks.Utilities, HtmlService: {}, console };
  vm.createContext(ctx);
  vm.runInContext(code + '\nthis.raw = { login };', ctx);
  assert.throws(() => ctx.raw.login('a', 'b'), /Authentication" belum ada/);
});

const csv = process.argv[2];
if (csv) {
  console.log('Code.gs (' + path.basename(csv) + ')');
  const real = parseCsv(fs.readFileSync(csv, 'utf8'));
  test('real export parses and a no-op save writes nothing', () => {
    const s = boot(clone(real));
    const d = s.api.getDashboardData();
    assert.ok(d.people.length > 0 && d.coaches.length > 0);
    const p = d.people.find((x) => !x.isCoach);
    const res = s.api.updatePerson({ id: p.id, name: p.name, position: p.position, notes: p.notes, milestones: clone(p.milestones) });
    assert.strictEqual(res.changes.length, 0);
    assert.strictEqual(s.mocks.writes.length, 0);
    console.log('    (' + d.people.length + ' people, coaches: ' + d.coaches.map((c) => c.name).join(', ') +
      ', editable: ' + d.editableFields.length + ' fields, warnings: ' + d.warnings.length + ')');
  });
}

console.log(passed + ' passed' + (process.exitCode ? ', some FAILED' : ''));
