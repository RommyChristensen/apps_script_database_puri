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
  assert.strictEqual(d.editableFields.length, 6);
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

const csv = process.argv[2];
if (csv) {
  console.log('Code.gs (' + path.basename(csv) + ')');
  const real = parseCsv(fs.readFileSync(csv, 'utf8'));
  test('notes stored under TARGET JD COACH are flagged and not overwritten by an empty field', () => {
    const s = boot(clone(real));
    const p = s.api.getDashboardData().people.find((x) => x.targetCoachHoldsNotes);
    assert.ok(p, 'found a misplaced note');
    const before = s.grid[p.id - 1][real[3].indexOf('TARGET JD COACH')];
    const res = s.api.updatePerson({ id: p.id, name: p.name, targetCoach: '', milestones: { msj3: !p.milestones.msj3 } });
    assert.deepStrictEqual(clone(res.changes.map((c) => c.label)), ['MSJ 3']);
    assert.throws(() => s.api.updatePerson({ id: p.id, name: p.name, targetCoach: 'Juni 2027' }), /masih berisi catatan/);
    assert.strictEqual(s.grid[p.id - 1][real[3].indexOf('TARGET JD COACH')], before);
  });
}

console.log(passed + ' passed' + (process.exitCode ? ', some FAILED' : ''));
