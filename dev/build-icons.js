/**
 * Renders the launcher icons in docs/icons/ from the Army of God logo (docs/icons/logo-source.jpg),
 * using the preinstalled Chromium (Playwright canvas). Replace logo-source.jpg and rerun to change the logo.
 *   node dev/build-icons.js
 *
 * - icon-192 / icon-512 ("any"): the full logo with rounded corners.
 * - icon-maskable-512: the logo at 88% on a matching dark gradient, edges feathered, so Android's round / squircle masks never crop the words.
 * - apple-touch-icon (180): same safe layout, square (iOS rounds the corners itself).
 * - favicon-32.
 */
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const DIR = path.join(__dirname, '..', 'docs', 'icons');
const SRC = 'data:image/jpeg;base64,' + fs.readFileSync(path.join(DIR, 'logo-source.jpg')).toString('base64');

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage();
  await page.setContent('<canvas id="c"></canvas>');
  const draw = (size, opts) => page.evaluate(async ([src, size, opts]) => {
    const img = new Image();
    img.src = src;
    await img.decode();
    const c = document.getElementById('c');
    c.width = c.height = size;
    const g = c.getContext('2d');
    g.imageSmoothingQuality = 'high';
    if (opts.radius) { g.beginPath(); g.roundRect(0, 0, size, size, size * opts.radius); g.clip(); }
    const sz = size * opts.scale, off = (size - sz) / 2;
    if (opts.scale < 1) {
      // Padding = a dark gradient matching the logo's edges (grey top, black bottom); the sharp logo gets
      // feathered edges so no seam shows.
      const grad = g.createLinearGradient(0, 0, 0, size);
      grad.addColorStop(0, '#272727'); grad.addColorStop(0.55, '#101010'); grad.addColorStop(1, '#000000');
      g.fillStyle = grad; g.fillRect(0, 0, size, size);
      const t = document.createElement('canvas');
      t.width = t.height = Math.round(sz);
      const tg = t.getContext('2d');
      tg.imageSmoothingQuality = 'high';
      tg.drawImage(img, 0, 0, sz, sz);
      const m = document.createElement('canvas');
      m.width = m.height = Math.round(sz);
      const mg = m.getContext('2d');
      const f = sz * 0.05;
      mg.filter = 'blur(' + f * 0.6 + 'px)';
      mg.fillStyle = '#fff';
      mg.fillRect(f, f, sz - 2 * f, sz - 2 * f);
      tg.globalCompositeOperation = 'destination-in';
      tg.drawImage(m, 0, 0);
      g.drawImage(t, off, off);
    } else {
      g.drawImage(img, 0, 0, size, size);
    }
    return c.toDataURL('image/png');
  }, [SRC, size, opts]);

  const write = async (file, size, opts) => {
    const url = await draw(size, opts);
    fs.writeFileSync(path.join(DIR, file), Buffer.from(url.split(',')[1], 'base64'));
    console.log('wrote', file, size + 'px');
  };
  await write('icon-192.png', 192, { scale: 1, radius: 0.22 });
  await write('icon-512.png', 512, { scale: 1, radius: 0.22 });
  await write('icon-maskable-512.png', 512, { scale: 0.88, radius: 0 });
  await write('apple-touch-icon.png', 180, { scale: 0.94, radius: 0 });
  await write('favicon-32.png', 32, { scale: 1, radius: 0.22 });
  await browser.close();
})();
