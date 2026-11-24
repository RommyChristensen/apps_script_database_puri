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

const code = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');

function boot(values, opts) {
  const mocks = createMockServices(values, opts);
  const ctx = { SpreadsheetApp: mocks.SpreadsheetApp, LockService: mocks.LockService, Session: mocks.Session, HtmlService: {}, console };
  vm.createContext(ctx);
  vm.runInContext(code + '\nthis.api = { getDashboardData, updatePerson, MILESTONES };', ctx);
  return { api: ctx.api, mocks, grid: values };
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
