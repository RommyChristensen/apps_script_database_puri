/**
 * Builds dev/preview.html — the web app rendered locally with fake demo data,
 * so the UI can be checked in a browser without deploying.
 *
 *   node dev/build-preview.js            # synthetic demo data
 *   node dev/build-preview.js sheet.csv  # a CSV export of the database tab
 *
 * The generated preview.html is git-ignored (it may contain real data).
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.join(__dirname, '..');
const read = (f) => fs.readFileSync(path.join(root, f), 'utf8');

// Load Code.gs into a sandbox so we reuse the real parser.
const ctx = {};
vm.createContext(ctx);
vm.runInContext(read('Code.gs') + '\nthis.parseDatabase_ = parseDatabase_; this.MILESTONES = MILESTONES;', ctx);

function parseCsv(t) {
  const rows = []; let row = [], f = '', q = false;
  for (let i = 0; i < t.length; i++) {
    const c = t[i];
    if (q) { if (c === '"') { if (t[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c; }
    else if (c === '"') q = true;
    else if (c === ',') { row.push(f); f = ''; }
    else if (c === '\n') { row.push(f); rows.push(row); row = []; f = ''; }
    else if (c !== '\r') f += c;
  }
  if (f || row.length) { row.push(f); rows.push(row); }
  return rows;
}

const HEADER = ['', 'NO CG', 'NAMA', 'STATUS CG', 'WAKTU CG', 'DAERAH CG', 'DOMISILI (KECAMATAN)', 'EMAIL', 'NO HP', 'CGL', 'GENDER', 'NIJ', 'POSISI', 'STATUS', 'KOTA KELAHIRAN', 'TTL', 'ALAMAT', 'NO HP', 'KERJA/KULIAH', 'KAMPUS', '5M', '', '', '', '', 'MSJ', '', '', 'CGT', '', '', 'PELAYANAN', 'TARGET JD MEMBER', 'TGL JD MEMBER', 'TARGET JD SPONSOR', 'TGL JD SPONSOR', 'TARGET CGL', 'TARGET JD COACH', 'KETERANGAN', 'AGE', 'ELIGIBLE TL', 'ELIGIBLE COACH', 'ELIGIBLE CGL'];
const SUB = Array(20).fill('').concat(['PERCAYA YESUS', 'BAPTIS SELAM', 'BAPTIS RK', 'RUTIN CG', 'RUTIN IBADAH', 'MSJ 1', 'MSJ 2', 'MSJ 3', 'CGT 1', 'CGT 2', 'CGT 3'], Array(12).fill(''));

function demoValues() {
  let seed = 7;
  const rnd = () => (seed = (seed * 9301 + 49297) % 233280) / 233280;
  const pick = (a) => a[Math.floor(rnd() * a.length)];
  const first = ['Nathan', 'Grace', 'Kevin', 'Jessica', 'Daniel', 'Michelle', 'Joshua', 'Clara', 'Samuel', 'Felicia', 'Ryan', 'Natasha', 'Hans', 'Vania', 'Steven', 'Cindy', 'Jonathan', 'Olivia', 'Timothy', 'Evelyn'];
  const last = ['Wijaya', 'Santoso', 'Halim', 'Tanoto', 'Gunawan', 'Lim', 'Hartono', 'Kusuma', 'Susanto', 'Salim'];
  const coaches = [['ANGIE', 'Demo Angie'], ['LENITA', 'Demo Lenita'], ['CARLA', 'Demo Carla'], ['MEYSIA', 'Demo Meysia'], ['DAVID', 'Demo David']];
  const rows = [['COACH ' + coaches[0][0]].concat(Array(42).fill('')), Array(43).fill(''), Array(43).fill(''), HEADER, SUB];
  let n = 1, cgNo = 1;
  coaches.forEach(([c, full], ci) => {
    if (ci > 0) rows.push(['COACH ' + c].concat(Array(42).fill('')));
    rows.push([String(n++), '-', full].concat(Array(40).fill('-')));
    const cgCount = 2 + Math.floor(rnd() * 3);
    for (let g = 0; g < cgCount; g++) {
      const code = 'AOG ' + String(cgNo++).padStart(2, '0');
      const time = pick(['Jumat 19.30', 'Minggu 11.00', 'Kamis 19.00', 'Sabtu 16.00']);
      const area = pick(['Puri', 'Cengkareng', 'Tandur', 'Kalideres']);
      const size = 4 + Math.floor(rnd() * 6);
      let cgl = '';
      for (let k = 0; k < size; k++) {
        const name = pick(first) + ' ' + pick(last);
        const pos = k === 0 ? 'CGL' : pick(['SPONSOR', 'MEMBER', 'MEMBER', 'MEMBER', 'SIMPATISAN', 'VIP']);
        if (k === 0) cgl = name;
        const lvl = { CGL: 11, SPONSOR: 8, MEMBER: 5, SIMPATISAN: 2, VIP: 1 }[pos];
        const ms = Array.from({ length: 11 }, (_, i) => (i < lvl || rnd() < 0.15 ? '✔' : '-'));
        const r = Array(43).fill('');
        Object.assign(r, {
          0: String(n++), 1: code, 2: name, 3: pick(['Onsite', 'Onsite', 'Online']), 4: time, 5: area,
          6: pick(['Kembangan', 'Kalideres', 'Cengkareng', 'Kebon Jeruk']), 7: name.split(' ')[0].toLowerCase() + '@example.com',
          8: '0812000' + String(1000 + n), 9: cgl, 10: rnd() < 0.5 ? 'Male' : 'Female', 11: String(24000000 + n), 12: pos,
          13: rnd() < 0.9 ? 'Aktif' : 'Non-Aktif', 14: 'Jakarta', 15: (1 + Math.floor(rnd() * 27)) + ' ' + pick(['January', 'March', 'June', 'October']) + ' ' + (1999 + Math.floor(rnd() * 8)),
          16: 'Jl. Contoh No. ' + n, 18: pick(['Kuliah', 'Kerja']), 19: pick(['Binus', 'UPH', 'Trisakti', 'Untar', '']),
          31: pos === 'MEMBER' && rnd() < 0.5 ? '' : pick(['Usher', 'Multimedia', 'Worship', '']),
          36: pos === 'SPONSOR' ? 'Juni 2026' : '', 38: rnd() < 0.3 ? 'Sedang dibimbing untuk naik level berikutnya' : '',
          40: 'NO', 41: 'NO', 42: pos === 'SPONSOR' && rnd() < 0.4 ? 'YES' : 'NO',
        });
        ms.forEach((v, i) => (r[20 + i] = v));
        rows.push(r);
      }
    }
  });
  return rows;
}

const csvArg = process.argv[2];
const values = csvArg ? parseCsv(fs.readFileSync(csvArg, 'utf8')) : demoValues();
const parsed = ctx.parseDatabase_(values);
const data = {
  generatedAt: new Date().toISOString(),
  sheetName: csvArg ? 'CSV preview' : 'Database (demo)',
  spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/demo/edit#gid=0',
  milestones: ctx.MILESTONES.map((m) => ({ key: m.key, label: m.label, group: m.group })),
  coaches: parsed.coaches, people: parsed.people, warnings: parsed.warnings,
};

let html = read('Index.html')
  .replace(/<\?!= include\('(\w+)'\); \?>/g, (_, f) => (f === 'JavaScript' ? '<script>window.__DEMO_DATA__ = ' + JSON.stringify(data) + ';</script>\n' : '') + read(f + '.html'))
  .replace(/<\?= appTitle \?>/g, 'AOG Youth Puri');
fs.writeFileSync(path.join(__dirname, 'preview.html'), html);
console.log('Wrote dev/preview.html —', data.coaches.length, 'coaches,', data.people.length, 'people,', data.warnings.length, 'warnings');
