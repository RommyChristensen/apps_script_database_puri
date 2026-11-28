/**
 * Draws the "Connect Group" icon (five people gathered around a ring) and renders the PNG sizes
 * the launcher needs, using the preinstalled Chromium (Playwright). Output goes to docs/icons/.
 *   node dev/build-icons.js
 */
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const OUT = path.join(__dirname, '..', 'docs', 'icons');
const RING = 230, SCALE = 1.8;

// One person: head + shoulders, drawn around (0, 0).
const person = '<circle cx="0" cy="-58" r="28"/><path d="M-48 44 v-18 a48 48 0 0 1 96 0 v18 z"/>';
// Five people around a ring: top, then clockwise. Upright, drawn back to front.
const angles = [-90, -18, 198, 54, 126];
const people = angles.map((a) => {
  const x = 512 + RING * Math.cos(a * Math.PI / 180), y = 512 + RING * Math.sin(a * Math.PI / 180);
  return `<g transform="translate(${x.toFixed(1)} ${y.toFixed(1)}) scale(${SCALE})">${person}</g>`;
}).join('');

function svg({ rounded }) {
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1024 1024" width="1024" height="1024">
  <defs>
    <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#1d4ed8"/><stop offset="0.55" stop-color="#2563eb"/><stop offset="1" stop-color="#38bdf8"/></linearGradient>
    <radialGradient id="glow" cx="0.78" cy="0.12" r="0.7"><stop offset="0" stop-color="#ffffff" stop-opacity="0.28"/><stop offset="1" stop-color="#ffffff" stop-opacity="0"/></radialGradient>
  </defs>
  <rect width="1024" height="1024" ${rounded ? 'rx="224" ' : ''}fill="url(#bg)"/>
  <rect width="1024" height="1024" ${rounded ? 'rx="224" ' : ''}fill="url(#glow)"/>
  <circle cx="512" cy="512" r="${RING}" fill="none" stroke="#fff" stroke-opacity="0.38" stroke-width="16"/>
  <circle cx="512" cy="512" r="58" fill="#fff" fill-opacity="0.16"/>
  <g fill="#ffffff">${people}</g>
</svg>`;
}

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  fs.writeFileSync(path.join(OUT, 'icon.svg'), svg({ rounded: true }));
  fs.writeFileSync(path.join(OUT, 'icon-maskable.svg'), svg({ rounded: false }));
  const browser = await chromium.launch();
  const page = await browser.newPage();
  const render = async (file, source, size) => {
    await page.setViewportSize({ width: size, height: size });
    await page.setContent(`<style>html,body{margin:0;background:transparent}svg{display:block;width:${size}px;height:${size}px}</style>${source}`);
    await page.screenshot({ path: path.join(OUT, file), omitBackground: true, clip: { x: 0, y: 0, width: size, height: size } });
    console.log('wrote', file, size + 'px');
  };
  const any = svg({ rounded: true }), maskable = svg({ rounded: false });
  await render('icon-192.png', any, 192);
  await render('icon-512.png', any, 512);
  await render('icon-maskable-512.png', maskable, 512);
  await render('apple-touch-icon.png', maskable, 180); // iOS rounds the corners itself
  await render('favicon-32.png', any, 32);
  await browser.close();
})();
