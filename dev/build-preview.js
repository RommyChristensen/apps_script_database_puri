/**
 * Builds dev/preview.html — the web app running locally (real Code.gs on a mock sheet),
 * so the UI can be checked in a browser without deploying.
 *
 *   node dev/build-preview.js            # synthetic demo data
 *   node dev/build-preview.js sheet.csv  # a CSV export of the database tab
 *
 * The generated preview.html is git-ignored (it may contain real data).
 */
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const read = (f) => fs.readFileSync(path.join(root, f), 'utf8');

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

const HEADER = ['', 'NO CG', 'NAMA', 'STATUS CG', 'WAKTU CG', 'DAERAH CG', 'DOMISILI (KECAMATAN)', 'EMAIL', 'NO HP', 'CGL', 'GENDER', 'NIJ', 'POSISI', 'STATUS', 'KOTA KELAHIRAN', 'TTL', 'ALAMAT', 'NO HP', 'KERJA/KULIAH', 'KAMPUS', '5M', '', '', '', '', 'MSJ', '', '', 'CGT', '', '', 'PELAYANAN', 'TARGET JD MEMBER', 'TGL JD MEMBER', 'TARGET JD SPONSOR', 'TGL JD SPONSOR', 'TARGET CGL', 'TARGET JD COACH', 'KETERANGAN', 'AGE', 'ELIGIBLE TL', 'ELIGIBLE COACH', 'ELIGIBLE CGL',
  'TGL JD CGL', 'TGL JD COACH', 'TARGET JD TEAM LEADER', 'TGL JD TEAM LEADER'];
const SUB = Array(20).fill('').concat(['PERCAYA YESUS', 'BAPTIS SELAM', 'BAPTIS RK', 'RUTIN CG', 'RUTIN IBADAH', 'MSJ 1', 'MSJ 2', 'MSJ 3', 'CGT 1', 'CGT 2', 'CGT 3'], Array(16).fill(''));

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

function main() {
  const csvArg = process.argv[2];
  const values = csvArg ? parseCsv(fs.readFileSync(csvArg, 'utf8')) : demoValues();

  // Demo logins: admin/admin123 sees everyone; <coach>/<coach>123 sees one coach.
  const coachNames = new Set();
  const coachCol = values[3] ? values[3].indexOf('COACH') : -1;
  values.forEach((r, i) => {
    const banner = /^COACH\s+(.+)$/i.exec(r[0] || '');
    if (banner) coachNames.add(banner[1].trim());
    if (i > 4 && coachCol >= 0 && r[coachCol] && r[coachCol] !== '-') coachNames.add(r[coachCol].trim());
  });
  const accounts = [['USERNAME', 'PASSWORD', 'COACH', 'ROLE'], ['admin', 'admin123', '', 'ADMIN']]
    .concat([...coachNames].map((c) => [c.toLowerCase(), c.toLowerCase() + '123', c, 'COACH']));

  // The page runs the REAL Code.gs in the browser against an in-memory mock of
  // SpreadsheetApp, and google.script.run is stubbed to call it with a delay.
  const server = `
  <script>${read('dev/mock-apps-script.js')}</script>
  <script>
  (function () {
    var mocks = createMockServices(${JSON.stringify(values)}, {
      sheetName: ${JSON.stringify(csvArg ? 'CSV preview' : 'Database (demo)')},
      sheets: { Authentication: ${JSON.stringify(accounts)} },
      geocodeAny: true,
    });
    var SpreadsheetApp = mocks.SpreadsheetApp, LockService = mocks.LockService, Session = mocks.Session,
      HtmlService = mocks.HtmlService, CacheService = mocks.CacheService, Utilities = mocks.Utilities,
      PropertiesService = mocks.PropertiesService, Maps = mocks.Maps, UrlFetchApp = mocks.UrlFetchApp;
    ${read('Code.gs')}
    ${read('Auth.gs')}
    ${read('Meetings.gs')}
    ${read('Ldp.gs')}
    ${read('Locations.gs')}
    window.__mock = mocks;
    var api = { getDashboardData: getDashboardData, updatePerson: updatePerson, login: login, logout: logout,
      getMeetings: getMeetings, saveMeeting: saveMeeting, setTaskDone: setTaskDone, deleteMeeting: deleteMeeting,
      getLdpProgress: getLdpProgress, saveLdp: saveLdp,
      getCgLocations: getCgLocations, setCgLocation: setCgLocation, previewLocation: previewLocation };
    function runner(ok, fail) {
      var r = {
        withSuccessHandler: function (f) { return runner(f, fail); },
        withFailureHandler: function (f) { return runner(ok, f); },
      };
      Object.keys(api).forEach(function (name) {
        r[name] = function () {
          var args = Array.prototype.slice.call(arguments);
          setTimeout(function () {
            try { var res = JSON.parse(JSON.stringify(api[name].apply(null, args))); ok && ok(res); }
            catch (e) { fail && fail(e); }
          }, window.__mockDelay == null ? 450 : window.__mockDelay);
        };
      });
      return r;
    }
    window.google = { script: { run: runner() } };
  })();
  </script>
  `;

  let html = read('Index.html')
    .replace(/<\?!= include\('(\w+)'\); \?>/g, (_, f) => (f === 'JavaScript' ? server : '') + read(f + '.html'))
    .replace(/<\?= appTitle \?>/g, 'AOG Youth Puri');
  // Offline preview: LEAFLET_DIR=/path/to/leaflet/dist serves Leaflet from disk instead of cdnjs.
  if (process.env.LEAFLET_DIR) {
    html = html.split('https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/').join('file://' + path.resolve(process.env.LEAFLET_DIR) + '/');
  }
  fs.writeFileSync(path.join(__dirname, 'preview.html'), html);
  console.log('Wrote dev/preview.html (' + (csvArg ? 'CSV' : 'demo') + ' data, ' + values.length + ' rows)');
  console.log('Demo logins: ' + accounts.slice(1).map((a) => a[0] + '/' + a[1]).join(', '));
}

module.exports = { demoValues, parseCsv };
if (require.main === module) main();
