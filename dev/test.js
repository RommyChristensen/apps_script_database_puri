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

const code = ['Code.gs', 'Auth.gs', 'Meetings.gs', 'Ldp.gs', 'Locations.gs'].map((f) => fs.readFileSync(path.join(__dirname, '..', f), 'utf8')).join('\n');

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
    PropertiesService: mocks.PropertiesService, Maps: mocks.Maps, UrlFetchApp: mocks.UrlFetchApp,
  };
  vm.createContext(ctx);
  vm.runInContext(code + '\nthis.raw = { getDashboardData, updatePerson, login, logout, getMeetings, saveMeeting, setTaskDone, deleteMeeting, getLdpProgress, saveLdp, getCgLocations, setCgLocation, previewLocation, MILESTONES };', ctx);
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
  assert.strictEqual(d.editableFields.filter((f) => f.section === 'jenjang').length, 11);
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
  assert.deepStrictEqual(clone(res.changes.map((c) => c.label)), ['MSJ 3', 'Posisi', 'Tgl jadi Sponsor', 'Target jadi CGL']);
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

console.log('Meetings.gs');
function meetingSetup(as) {
  const s = boot(clone(base), { accounts: coachAccounts(), as: as || ['angie', 'angie-pass'] });
  const admin = s.raw.login('admin', 'admin-pass').token;
  const everyone = s.raw.getDashboardData(admin).people;
  const of = (coach) => everyone.filter((p) => p.coach === coach && !p.isCoach);
  const sheet = () => s.mocks.spreadsheet.getSheetByName('Log Pertemuan');
  return { s, admin, of, sheet };
}
const meeting = (people, extra) => Object.assign({
  date: '2026-11-20', people, category: 'Pembinaan', purpose: 'Persiapan naik sponsor',
  tasks: [{ text: 'Daftar MSJ 3', deadline: '2026-12-10' }, { text: 'Follow up CG', deadline: '' }],
}, extra);

test('saving a meeting creates the tab (last), one row per task, as text', () => {
  const { s, of, sheet } = meetingSetup();
  const [a, b] = of('Angie');
  const res = s.raw.saveMeeting(s.token, meeting([a.name, b.name]));
  assert.ok(/^M\d{12}[0-9A-F]{4}$/.test(res.savedId));
  const tabs = s.mocks.spreadsheet.getSheets().map((t) => t.getName());
  assert.strictEqual(tabs[0], 'Database', 'database tab stays first');
  assert.strictEqual(tabs[tabs.length - 1], 'Log Pertemuan');
  const g = sheet()._grid;
  assert.deepStrictEqual(clone(g[0].slice(0, 4)), ['MEETING ID', 'TANGGAL', 'COACH', 'DENGAN']);
  assert.strictEqual(g.length, 3);
  assert.deepStrictEqual(clone(g[1].slice(1, 10)), ['2026-11-20', 'Angie', a.name + '; ' + b.name, 'Pembinaan',
    'Persiapan naik sponsor', '1', 'Daftar MSJ 3', '2026-12-10', '-']);
  assert.strictEqual(g[2][7], 'Follow up CG');
  const m = res.meetings[0];
  assert.deepStrictEqual(clone(m.people), [a.name, b.name]);
  assert.strictEqual(m.tasks.length, 2);
  assert.strictEqual(m.createdBy, 'angie');
});

test('a meeting without tasks is one row; editing replaces its rows and keeps the creator', () => {
  const { s, of, sheet } = meetingSetup();
  const [a] = of('Angie');
  const id = s.raw.saveMeeting(s.token, meeting([a.name], { tasks: [] })).savedId;
  assert.strictEqual(sheet()._grid.length, 2);
  const admin = s.raw.login('admin', 'admin-pass').token;
  const res = s.raw.saveMeeting(admin, meeting([a.name], { id, purpose: 'Diubah admin', tasks: [{ text: 'A' }, { text: 'B' }, { text: 'C' }] }));
  assert.strictEqual(sheet()._grid.length, 4);
  assert.ok(sheet()._grid.slice(1).every((r) => r[0] === id && r[5] === 'Diubah admin' && r[10] === 'angie'));
  assert.strictEqual(res.meetings.length, 1);
});

test('ticking a task only changes that task', () => {
  const { s, of } = meetingSetup();
  const [a] = of('Angie');
  const id = s.raw.saveMeeting(s.token, meeting([a.name])).savedId;
  const res = s.raw.setTaskDone(s.token, id, '2', true);
  assert.deepStrictEqual(clone(res.meetings[0].tasks.map((t) => t.done)), [false, true]);
  assert.deepStrictEqual(clone(s.raw.setTaskDone(s.token, id, '2', false).meetings[0].tasks.map((t) => t.done)), [false, false]);
  assert.throws(() => s.raw.setTaskDone(s.token, id, '9', true), /sudah tidak ada/);
});

test('coaches only see their own meetings and cannot touch other coaches\' meetings or people', () => {
  const { s, of, admin } = meetingSetup();
  const carla = of('Carla')[0];
  assert.throws(() => s.raw.saveMeeting(s.token, meeting([carla.name])), /tidak ditemukan di data coach kamu/);
  const carlaMeeting = s.raw.saveMeeting(admin, meeting([carla.name])).savedId;
  s.raw.saveMeeting(s.token, meeting([of('Angie')[0].name]));
  const mine = s.raw.getMeetings(s.token).meetings;
  assert.strictEqual(mine.length, 1);
  assert.ok(mine.every((m) => m.coach === 'Angie'));
  assert.strictEqual(s.raw.getMeetings(admin).meetings.length, 2);
  assert.throws(() => s.raw.setTaskDone(s.token, carlaMeeting, '1', true), /tidak punya akses/);
  assert.throws(() => s.raw.deleteMeeting(s.token, carlaMeeting), /tidak punya akses/);
  assert.throws(() => s.raw.saveMeeting(s.token, meeting([of('Angie')[0].name], { id: carlaMeeting })), /tidak punya akses/);
});

test('admin cannot mix people of different coaches in one meeting', () => {
  const { s, of, admin } = meetingSetup();
  assert.throws(() => s.raw.saveMeeting(admin, meeting([of('Angie')[0].name, of('Carla')[0].name])), /satu coach saja/);
});

test('meeting validation', () => {
  const { s, of } = meetingSetup();
  const n = of('Angie')[0].name;
  assert.throws(() => s.raw.saveMeeting(s.token, meeting([n], { date: '' })), /tanggal/);
  assert.throws(() => s.raw.saveMeeting(s.token, meeting([n], { date: '2026-02-31' })), /tanggal/);
  assert.throws(() => s.raw.saveMeeting(s.token, meeting([])), /minimal satu orang/);
  assert.throws(() => s.raw.saveMeeting(s.token, meeting([n], { category: 'Ngobrol' })), /kategori/);
  assert.throws(() => s.raw.saveMeeting(s.token, meeting([n], { purpose: ' ' })), /tujuan/);
  assert.throws(() => s.raw.saveMeeting(s.token, meeting([n], { tasks: [{ text: 'x', deadline: 'besok' }] })), /Deadline/);
  assert.throws(() => s.raw.getMeetings('nope'), /AUTH:/);
});

test('deleting a meeting removes all its rows and leaves others alone', () => {
  const { s, of, sheet } = meetingSetup();
  const [a, b] = of('Angie');
  const keep = s.raw.saveMeeting(s.token, meeting([a.name])).savedId;
  const drop = s.raw.saveMeeting(s.token, meeting([b.name], { date: '2026-11-21' })).savedId;
  const res = s.raw.deleteMeeting(s.token, drop);
  assert.deepStrictEqual(clone(res.meetings.map((m) => m.id)), [keep]);
  assert.ok(sheet()._grid.slice(1).every((r) => r[0] === keep));
});

console.log('Ldp.gs');
function ldpSetup(as) {
  const s = boot(clone(base), { accounts: coachAccounts(), as: as || ['angie', 'angie-pass'] });
  const admin = s.raw.login('admin', 'admin-pass').token;
  const people = () => s.raw.getDashboardData(admin).people;
  const colOf = (h) => s.grid[3].indexOf(h);
  return { s, admin, people, colOf };
}

test('setting targets writes "15 June 2026" text into the database', () => {
  const { s, people, colOf } = ldpSetup();
  const p = people().find((x) => x.coach === 'Angie' && x.position === 'MEMBER');
  const res = s.raw.saveLdp(s.token, { id: p.id, name: p.name, targets: { sponsor: '2026-06-15', cgl: '2027-01-10', tl: '' } });
  assert.strictEqual(s.grid[p.id - 1][colOf('TARGET JD SPONSOR')], '15 June 2026');
  assert.strictEqual(s.grid[p.id - 1][colOf('TARGET CGL')], '10 January 2027');
  assert.strictEqual(s.grid[p.id - 1][colOf('TARGET JD TEAM LEADER')] || '', '');
  assert.strictEqual(res.data.people.find((x) => x.id === p.id).targetSponsor, '15 June 2026');
  assert.strictEqual(res.progress.length, 0, 'no monthly entry without progress');
  assert.throws(() => s.raw.saveLdp(s.token, { id: p.id, name: p.name, targets: { sponsor: '2026-02-30' } }), /tidak valid/);
});

test('berhasil fills TGL, moves POSISI up and logs the check', () => {
  const { s, people, colOf } = ldpSetup();
  const p = people().find((x) => x.coach === 'Angie' && x.position === 'MEMBER');
  s.raw.saveLdp(s.token, { id: p.id, name: p.name, targets: { sponsor: '2026-11-30' } });
  const res = s.raw.saveLdp(s.token, { id: p.id, name: p.name, progress: { step: 'sponsor', result: 'berhasil', date: '2026-11-22' } });
  assert.strictEqual(s.grid[p.id - 1][colOf('TGL JD SPONSOR')], '22 November 2026');
  assert.strictEqual(s.grid[p.id - 1][colOf('POSISI')], 'SPONSOR');
  const log = s.mocks.spreadsheet.getSheetByName('LDP Progress')._grid;
  assert.deepStrictEqual(clone(log[0]).slice(0, 4), ['TANGGAL UPDATE', 'COACH', 'NAMA', 'NIJ']);
  assert.deepStrictEqual(clone(log[1].slice(1, 11)), ['Angie', p.name, p.nij, 'Sponsor', '2026-11-30', 'Berhasil', '2026-11-22', '', '', 'angie']);
  assert.strictEqual(res.progress[0].result, 'berhasil');
  assert.strictEqual(res.progress[0].step, 'sponsor');
});

test('berhasil never moves POSISI down; Coach/TL steps only record the date', () => {
  const { s, admin, people, colOf } = ldpSetup();
  const cgl = people().find((x) => x.coach === 'Angie' && x.position === 'CGL');
  s.raw.saveLdp(s.token, { id: cgl.id, name: cgl.name, progress: { step: 'member', result: 'berhasil', date: '2026-11-01' } });
  assert.strictEqual(s.grid[cgl.id - 1][colOf('POSISI')], 'CGL');
  s.raw.saveLdp(admin, { id: cgl.id, name: cgl.name, progress: { step: 'coach', result: 'berhasil', date: '2026-11-02' } });
  assert.strictEqual(s.grid[cgl.id - 1][colOf('POSISI')], 'CGL');
  assert.strictEqual(s.grid[cgl.id - 1][colOf('TGL JD COACH')], '2 November 2026');
});

test('belum needs a note and can move the target', () => {
  const { s, people, colOf } = ldpSetup();
  const p = people().find((x) => x.coach === 'Angie' && x.position === 'SPONSOR');
  assert.throws(() => s.raw.saveLdp(s.token, { id: p.id, name: p.name, progress: { step: 'cgl', result: 'belum', note: '' } }), /catatan/);
  s.raw.saveLdp(s.token, { id: p.id, name: p.name, targets: { cgl: '2026-11-15' } });
  const res = s.raw.saveLdp(s.token, { id: p.id, name: p.name,
    progress: { step: 'cgl', result: 'belum', note: 'Belum ikut CGT 2', newTarget: '2027-02-01' } });
  assert.strictEqual(s.grid[p.id - 1][colOf('TARGET CGL')], '1 February 2027');
  assert.strictEqual(s.grid[p.id - 1][colOf('POSISI')], 'SPONSOR');
  const e = res.progress[0];
  assert.deepStrictEqual([e.result, e.target, e.newTarget, e.note], ['belum', '2026-11-15', '2027-02-01', 'Belum ikut CGT 2']);
});

test('LDP access: coaches only for their own people and progress', () => {
  const { s, admin, people } = ldpSetup();
  const carla = people().find((x) => x.coach === 'Carla' && !x.isCoach);
  assert.throws(() => s.raw.saveLdp(s.token, { id: carla.id, name: carla.name, targets: { sponsor: '2026-06-01' } }), /tidak punya akses/);
  s.raw.saveLdp(admin, { id: carla.id, name: carla.name, progress: { step: 'sponsor', result: 'belum', note: 'Belum siap' } });
  assert.strictEqual(s.raw.getLdpProgress(s.token).progress.length, 0, 'Angie does not see Carla progress');
  assert.strictEqual(s.raw.getLdpProgress(admin).progress.length, 1);
  assert.throws(() => s.raw.saveLdp(s.token, { id: carla.id, name: 'Someone else', targets: {} }), /sudah berubah/);
  assert.throws(() => s.raw.getLdpProgress('x'), /AUTH:/);
});

console.log('Locations.gs');
const GEO = {
  'jl. puri indah raya no. 1, jakarta': { lat: -6.1885, lng: 106.7338 },
  'kalideres': { lat: -6.1411, lng: 106.7012 },
  'kembangan': { lat: -6.1902, lng: 106.7409 },
};
function locSetup(as, extra) {
  const s = boot(clone(base), Object.assign({ accounts: coachAccounts(), as: as || ['angie', 'angie-pass'], geocode: GEO }, extra));
  const admin = s.raw.login('admin', 'admin-pass').token;
  const cgsOf = (coach) => s.raw.getCgLocations(admin).cgs.filter((c) => c.coach === coach);
  const tab = () => s.mocks.spreadsheet.getSheetByName('Lokasi CG');
  return { s, admin, cgsOf, tab };
}

test('every CG is listed for every coach; only own CGs are editable', () => {
  const { s } = locSetup();
  const res = s.raw.getCgLocations(s.token);
  const coaches = new Set(res.cgs.map((c) => c.coach));
  assert.ok(coaches.size === 5, 'coach sees every coach group');
  assert.ok(res.cgs.every((c) => c.canEdit === (c.coach === 'Angie')));
  const g = res.cgs[0];
  assert.ok(g.code && g.size > 0 && g.cgl, 'CG summary has code, size and CGL');
  assert.ok(!('people' in g), 'no member details in the CG list');
});

test('setting a location from an address geocodes it and writes the tab (last)', () => {
  const { s, cgsOf, tab } = locSetup();
  const cg = cgsOf('Angie')[0];
  const res = s.raw.setCgLocation(s.token, { cg: cg.code, place: 'Rumah Sherly', address: 'Jl. Puri Indah Raya No. 1, Jakarta' });
  const got = res.cgs.find((c) => c.code === cg.code);
  assert.deepStrictEqual([got.lat, got.lng, got.place], [-6.1885, 106.7338, 'Rumah Sherly']);
  const names = s.mocks.spreadsheet.getSheets().map((t) => t.getName());
  assert.strictEqual(names[names.length - 1], 'Lokasi CG');
  assert.deepStrictEqual(clone(tab()._grid[1].slice(0, 8)), [cg.code, 'Angie', 'Rumah Sherly', 'Jl. Puri Indah Raya No. 1, Jakarta', '', '-6.188500', '106.733800', 'angie']);
  // second save updates the same row
  s.raw.setCgLocation(s.token, { cg: cg.code, place: 'Gereja', lat: -6.2, lng: 106.75 });
  assert.strictEqual(tab()._grid.length, 2);
  assert.strictEqual(tab()._grid[1][2], 'Gereja');
});

test('coordinates are read from Google Maps links, including maps.app.goo.gl short links', () => {
  const short = 'https://maps.app.goo.gl/AbCdEf123';
  const { s, cgsOf } = locSetup(null, { redirects: {
    [short]: 'https://www.google.com/maps/place/Rumah/@-6.1700000,106.7400000,17z/data=!3m1!4b1!4m6!3m5!1s0x0:0x0!8m2!3d-6.1712345!4d106.7398765',
  } });
  const [a, b] = cgsOf('Angie');
  let res = s.raw.setCgLocation(s.token, { cg: a.code, link: 'https://www.google.com/maps/@-6.15,106.72,15z' });
  assert.deepStrictEqual([res.cgs.find((c) => c.code === a.code).lat, res.cgs.find((c) => c.code === a.code).lng], [-6.15, 106.72]);
  res = s.raw.setCgLocation(s.token, { cg: b.code, link: short });
  const got = res.cgs.find((c) => c.code === b.code);
  assert.deepStrictEqual([got.lat, got.lng], [-6.171234, 106.739876], 'uses the pin (!3d!4d), not the map centre (stored to 6 decimals)');
  assert.deepStrictEqual(clone(s.mocks.fetchCalls), [short]);
  s.raw.setCgLocation(s.token, { cg: b.code, link: short });
  assert.strictEqual(s.mocks.fetchCalls.length, 1, 'expanded short link is cached');
});

test('location input is validated and access-checked', () => {
  const { s, cgsOf, admin } = locSetup();
  const mine = cgsOf('Angie')[0], other = cgsOf('Carla')[0];
  assert.throws(() => s.raw.setCgLocation(s.token, { cg: other.code, lat: -6.2, lng: 106.7 }), /coach kamu/);
  assert.throws(() => s.raw.setCgLocation(s.token, { cg: 'AOG 99', lat: -6.2, lng: 106.7 }), /tidak ditemukan/);
  assert.throws(() => s.raw.setCgLocation(s.token, { cg: mine.code, link: 'https://evil.example/maps/@-6,106' }), /link Google Maps/);
  assert.throws(() => s.raw.setCgLocation(s.token, { cg: mine.code, address: 'Tempat yang tidak ada' }), /tidak ketemu/);
  assert.throws(() => s.raw.getCgLocations('x'), /AUTH:/);
  // admin can edit any CG
  const res = s.raw.setCgLocation(admin, { cg: other.code, lat: -6.2, lng: 106.7 });
  assert.strictEqual(res.cgs.find((c) => c.code === other.code).lat, -6.2);
});

test('rows typed in the tab with only an address are geocoded once and cached', () => {
  const { s, cgsOf, tab, admin } = locSetup();
  const cg = cgsOf('Angie')[0];
  s.raw.setCgLocation(s.token, { cg: cg.code, lat: -6.2, lng: 106.7 }); // creates the tab
  tab()._grid[1][5] = ''; tab()._grid[1][6] = ''; tab()._grid[1][3] = 'Kembangan';
  const before = s.mocks.geocodeCalls.length;
  const got = () => s.raw.getCgLocations(admin).cgs.find((c) => c.code === cg.code);
  assert.deepStrictEqual([got().lat, got().lng], [-6.1902, 106.7409]);
  got(); got();
  assert.strictEqual(s.mocks.geocodeCalls.filter((q) => /kembangan/i.test(q)).length, 1, 'geocoded only once');
  assert.ok(s.mocks.geocodeCalls.length >= before);
});

test('domisili layer only counts people the session may see', () => {
  const { s, admin } = locSetup();
  const mine = s.raw.getCgLocations(s.token).homes;
  const all = s.raw.getCgLocations(admin).homes;
  const total = (h) => h.reduce((n, x) => n + x.count, 0);
  assert.ok(mine.length > 0 && total(mine) < total(all));
  assert.ok(mine.every((h) => typeof h.lat === 'number' && !('names' in h)));
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
