/**
 * Builds dev/preview.html from a fully INVENTED dataset (no real people) for
 * portfolio screenshots. Throwaway dev tool, not part of the Apps Script app.
 *
 *   LEAFLET_DIR=... node dev/portfolio.js [out.csv]
 *
 * Logins: admin/admin123, adrian/adrian123 (TL), bella/bella123 (TL),
 * naomi/naomi123 (coach), etc. (name + 123).
 */
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

let seed = 11;
const rnd = () => (seed = (seed * 9301 + 49297) % 233280) / 233280;
const pick = (a) => a[Math.floor(rnd() * a.length)];

const HEADER = ['', 'NO CG', 'NAMA', 'COACH', 'TL', 'STATUS CG', 'WAKTU CG', 'DAERAH CG', 'DOMISILI (KECAMATAN)', 'EMAIL', 'NO HP', 'CGL',
  'GENDER', 'NIJ', 'POSISI', 'STATUS', 'KOTA KELAHIRAN', 'TTL', 'ALAMAT', 'NO HP', 'KERJA/KULIAH', 'KAMPUS', '5M', '', '', '', '', 'MSJ', '', '',
  'CGT', '', '', '', 'PELAYANAN', 'TARGET JD MEMBER', 'TGL JD MEMBER', 'TARGET JD SPONSOR', 'TGL JD SPONSOR', 'TARGET JD CGL', 'TGL JD CGL',
  'TARGET JD COACH', 'TGL JD COACH', 'TARGET JD TEAM LEADER', 'TGL JD TEAM LEADER', 'ELIGIBLE TL', 'ELIGIBLE COACH', 'ELIGIBLE CGL',
  'KEHADIRAN', '% KEHADIRAN'];
const W = HEADER.length;
const SUB = Array(W).fill('');
['PERCAYA YESUS', 'BAPTIS SELAM', 'BAPTIS RK', 'RUTIN CG', 'RUTIN IBADAH', 'MSJ 1', 'MSJ 2', 'MSJ 3', 'CGT 1', 'CGT 2', 'CGT 3', 'CGT 4']
  .forEach((h, i) => { SUB[22 + i] = h; });

const FIRST = ['Alvaro', 'Bianca', 'Cleo', 'Dimas', 'Elisa', 'Farrel', 'Gita', 'Hendra', 'Indah', 'Jovan', 'Kirana', 'Lukas', 'Maya', 'Nico',
  'Olla', 'Putra', 'Queenie', 'Rafi', 'Salsa', 'Tegar', 'Una', 'Vino', 'Wulan', 'Xena', 'Yosef', 'Zahra', 'Arka', 'Bella', 'Cakra', 'Dara',
  'Erwin', 'Fiona', 'Galih', 'Hana', 'Ivan', 'Jelita', 'Kenzo', 'Lia', 'Mario', 'Nadia'];
const LAST = ['Aditama', 'Brata', 'Cahyadi', 'Darmawan', 'Efendi', 'Firmansyah', 'Gunadi', 'Hartanto', 'Irawan', 'Jayadi', 'Kurnia', 'Lesmana',
  'Mahendra', 'Nugraha', 'Oktavia', 'Pratama', 'Rahardjo', 'Saputra', 'Tirtayasa', 'Utomo', 'Wibowo', 'Yulianto'];
const used = new Set();
function fakeName() {
  for (;;) { const n = pick(FIRST) + ' ' + pick(LAST); if (!used.has(n)) { used.add(n); return n; } }
}

const TLS = [['Adrian Pratama', 'adrian'], ['Bella Anjani', 'bella']];
// coach short name, full name, TL index
const COACHES = [['Naomi', 'Naomi Kartika', 0], ['Samuel', 'Samuel Hartono', 0], ['Priska', 'Priska Wirawan', 0], ['Reza', 'Reza Mahardika', 1], ['Tiara', 'Tiara Lestari', 1]];
// CG plan per coach: [members, sponsors, hasCgl]  → mix of Healthy / Growing / At Risk, one very big CG
const PLAN = [
  [[7, 3, 1], [5, 2, 1], [15, 3, 1], [3, 1, 1]],
  [[6, 2, 1], [9, 4, 1], [2, 1, 1], [5, 2, 1]],
  [[8, 3, 1], [4, 1, 1], [6, 0, 1], [5, 2, 1]],
  [[7, 2, 1], [3, 2, 1], [10, 3, 1], [5, 3, 0]],
  [[6, 3, 1], [4, 2, 1], [1, 0, 1], [8, 2, 1]],
];
const TIMES = ['Jumat 19.30', 'Minggu 11.00', 'Kamis 19.00', 'Sabtu 16.00', 'Rabu 19.00'];
const AREAS = ['Kembangan', 'Kalideres', 'Cengkareng', 'Kebon Jeruk', 'Grogol'];

const now = new Date();
const isoShift = (n) => new Date(now.getFullYear(), now.getMonth(), now.getDate() + n);
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const ttl = (d, year) => d.getDate() + ' ' + MONTHS[d.getMonth()] + ' ' + year;
// birthdays: one tomorrow, the others spread over the next weeks (invented years)
const BIRTH_OFFSETS = [1, 4, 9, 12, 15, 18, 23, 27];
let birthIdx = 0;

const rows = [Array(W).fill(''), Array(W).fill(''), Array(W).fill(''), HEADER, SUB];
let n = 1, cgNo = 1;
const row = (vals) => { const r = Array(W).fill(''); Object.entries(vals).forEach(([k, v]) => { r[+k] = v; }); return r; };

// Team leaders
TLS.forEach(([name]) => rows.push(row({ 0: n++, 1: '-', 2: name, 4: name, 12: 'Male', 13: String(21000000 + n), 14: 'TEAM LEADER', 15: 'Aktif', 17: ttl(isoShift(40 + n), 1996) })));

COACHES.forEach(([short, full, ti], ci) => {
  const tl = TLS[ti][0];
  rows.push(row({ 0: n++, 1: '-', 2: full, 3: short, 4: tl, 12: ci % 2 ? 'Male' : 'Female', 13: String(22000000 + n), 14: 'COACH', 15: 'Aktif', 17: ttl(isoShift(60 + n), 1999) }));
  PLAN[ci].forEach(([members, sponsors, hasCgl], g) => {
    const code = (ci === 3 && g === 3 ? 'TEEN ' : 'AOG ') + String(cgNo++).padStart(2, '0');
    const time = TIMES[(ci + g) % TIMES.length], area = AREAS[(ci * 2 + g) % AREAS.length];
    const mode = g === 3 ? 'Online' : 'Onsite';
    const cgl = hasCgl ? fakeName() : '';
    const total = members + sponsors + (hasCgl ? 1 : 0);
    for (let k = 0; k < total; k++) {
      const name = k === 0 && hasCgl ? cgl : fakeName();
      const pos = k === 0 && hasCgl ? 'CGL' : k < sponsors + (hasCgl ? 1 : 0) ? 'SPONSOR' : (k % 11 === 0 ? 'VIP' : pick(['MEMBER', 'MEMBER', 'MEMBER', 'MEMBER', 'SIMPATISAN']));
      const lvl = { CGL: 12, SPONSOR: 9, MEMBER: 6, SIMPATISAN: 3, VIP: 1 }[pos];
      const bday = birthIdx < BIRTH_OFFSETS.length && k === 1 ? isoShift(BIRTH_OFFSETS[birthIdx++]) : isoShift(30 + Math.floor(rnd() * 330));
      const att = 40 + Math.floor(rnd() * 60);
      const r = row({
        0: n++, 1: code, 2: name, 3: short, 4: tl, 5: mode, 6: time, 7: area, 8: pick(['Kembangan', 'Kalideres', 'Cengkareng', 'Kebon Jeruk']),
        9: name.split(' ')[0].toLowerCase() + n + '@example.com', 10: '08120' + String(100000 + n * 37).slice(-6), 11: cgl,
        12: rnd() < 0.5 ? 'Male' : 'Female', 13: String(23000000 + n), 14: pos, 15: rnd() < 0.93 ? 'Aktif' : 'Non-Aktif', 16: 'Jakarta',
        17: ttl(bday, 2000 + Math.floor(rnd() * 8)), 18: 'Jl. Contoh No. ' + n, 20: pick(['Kuliah', 'Kerja']),
        21: pick(['Universitas Contoh', 'Politeknik Nusantara', 'Institut Harapan', '']),
        34: pos === 'MEMBER' && rnd() < 0.5 ? '' : pick(['Usher', 'Multimedia', 'Worship', 'Hospitality', '']),
        48: String(8 + Math.floor(rnd() * 8)), 49: att + '%',
      });
      for (let i = 0; i < 12; i++) r[22 + i] = i < lvl || rnd() < 0.12 ? '✔' : '-';
      rows.push(r);
    }
  });
});

const csv = rows.map((r) => r.map((c) => (/[",\n]/.test(String(c)) ? '"' + String(c).replace(/"/g, '""') + '"' : String(c))).join(',')).join('\n');
const out = process.argv[2] || path.join(__dirname, 'portfolio-data.csv');
fs.writeFileSync(out, csv);
console.log('Fake dataset: ' + (rows.length - 5) + ' rows, ' + (cgNo - 1) + ' CGs, written to ' + out);
execFileSync('node', [path.join(__dirname, 'build-preview.js'), out], { stdio: 'inherit', env: process.env });
// Friendlier label than "CSV preview" in the sidebar and footer.
const html = path.join(__dirname, 'preview.html');
fs.writeFileSync(html, fs.readFileSync(html, 'utf8').split('CSV preview').join('Database Jemaat'));
