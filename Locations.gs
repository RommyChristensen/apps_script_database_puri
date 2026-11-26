/**
 * CG locations for the "Lokasi" menu.
 *
 * Stored in the "Lokasi CG" tab (CONFIG.LOCATION_SHEET_NAME), one row per CG:
 *   NO CG | COACH | NAMA TEMPAT | ALAMAT | LINK MAPS | LATITUDE | LONGITUDE | DIPERBARUI OLEH | DIPERBARUI
 *
 * Fill a row by hand (address and/or a Google Maps link; coordinates optional)
 * or from the app. Missing coordinates are worked out from the Maps link or,
 * failing that, geocoded from the address with Apps Script's built-in Maps
 * service (no API key needed); results are cached in Script Properties.
 *
 * Access: every signed-in coach sees all CGs (CG-level info only: code, coach,
 * CGL, schedule, place); only the CG's own coach or an admin may change its
 * location. The "domisili" layer only counts the people the session may see.
 */

const LOCATION = {
  HEADERS: ['NO CG', 'COACH', 'NAMA TEMPAT', 'ALAMAT', 'LINK MAPS', 'LATITUDE', 'LONGITUDE', 'DIPERBARUI OLEH', 'DIPERBARUI'],
  // Bias geocoding towards Greater Jakarta (south-west and north-east corners).
  BOUNDS: [-6.6, 106.3, -5.9, 107.2],
  MAX_GEOCODES_PER_REQUEST: 40,
  // Hosts whose short links may be expanded to read the coordinates.
  SHORT_LINK_HOSTS: ['maps.app.goo.gl', 'goo.gl', 'g.co'],
};

/** Called from the browser: all CGs with their location + the domisili layer for this session. */
function getCgLocations(token) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  return locationsPayload_(ss, session);
}

/**
 * Sets one CG's location. loc = { cg, place, address, link, lat, lng }.
 * Coordinates come from lat/lng (a pin dropped on the map), else the Maps
 * link, else the address. Clearing every field removes the location.
 */
function setCgLocation(token, loc) {
  const ss = getSpreadsheet_();
  const session = requireSession_(token, ss);
  requirePerm_(session, 'restructure');
  if (!loc || !loc.cg) throw new Error('Pilih CG.');
  const cgs = cgSummaries_(ss);
  const cg = cgs.filter(function (c) { return c.code === String(loc.cg).toUpperCase().trim(); })[0];
  if (!cg) throw new Error('CG "' + loc.cg + '" tidak ditemukan di database.');
  if (!canAccessCoach_(session, cg.coach)) throw new Error('Kamu hanya bisa mengubah lokasi CG milik coach kamu.');
  const resolved = resolveLocation_(loc);
  return withScriptLock_(function () {
    writeLocationRow_(ss, session, cg.code, cg.coach, resolved);
    SpreadsheetApp.flush();
    return locationsPayload_(ss, session);
  });
}

/**
 * Cleans the form input and works out coordinates (pin, else Maps link, else address).
 * Throws when a link/address was given but can't be placed. Returns { place, address, link, coords }.
 */
function resolveLocation_(loc) {
  const place = String(loc.place || '').replace(/\s+/g, ' ').trim().slice(0, 120);
  const address = String(loc.address || '').replace(/\s+/g, ' ').trim().slice(0, 300);
  const link = String(loc.link || '').trim().slice(0, 500);
  if (link && !/^https:\/\/([a-z0-9-]+\.)*(google\.[a-z.]+|goo\.gl|g\.co)\//i.test(link)) {
    throw new Error('Link harus link Google Maps (https://maps.app.goo.gl/… atau https://www.google.com/maps/…).');
  }
  let coords = validCoords_(loc.lat, loc.lng);
  if (!coords && link) coords = coordsFromMapsLink_(link);
  if (!coords && address) coords = geocode_(address);
  if (!coords && (link || address)) {
    throw new Error('Lokasi tidak ketemu dari link/alamat itu. Coba tempel link Google Maps lain, atau klik peta untuk menaruh pin.');
  }
  return { place: place, address: address, link: link, coords: coords };
}

/** Writes (or replaces) the Lokasi CG row of one CG. Caller holds the script lock. */
function writeLocationRow_(ss, session, code, coach, r) {
  const sheet = getLocationSheet_(ss, true);
  const existing = readLocationRows_(sheet).filter(function (x) { return x.code === code; })[0];
  const now = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd HH:mm');
  const values = [[code, coach, r.place, r.address, r.link,
    r.coords ? r.coords.lat.toFixed(6) : '', r.coords ? r.coords.lng.toFixed(6) : '', session.username, now]];
  const range = sheet.getRange(existing ? existing.row : sheet.getLastRow() + 1, 1, 1, LOCATION.HEADERS.length);
  range.setNumberFormat('@');
  range.setValues(values);
}

/** Called from the edit form to preview where an address or Maps link lands. */
function previewLocation(token, input) {
  const ss = getSpreadsheet_();
  requireSession_(token, ss);
  const link = String((input && input.link) || '').trim();
  const address = String((input && input.address) || '').trim();
  const coords = (link && coordsFromMapsLink_(link)) || (address && geocode_(address)) || null;
  return coords ? { lat: coords.lat, lng: coords.lng } : null;
}

// ───────────────────────────── helpers ─────────────────────────────

function getLocationSheet_(ss, create) {
  let sheet = ss.getSheetByName(CONFIG.LOCATION_SHEET_NAME);
  if (!sheet && create) sheet = ss.insertSheet(CONFIG.LOCATION_SHEET_NAME, ss.getSheets().length);
  if (sheet && create && sheet.getLastRow() === 0) {
    sheet.appendRow(LOCATION.HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, LOCATION.HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

function readLocationRows_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getDisplayValues();
  const header = values[0].map(norm_);
  const c = {};
  LOCATION.HEADERS.forEach(function (h) { c[h] = header.indexOf(h); });
  const get = function (r, h) { return c[h] < 0 ? '' : String(r[c[h]] || '').trim(); };
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const code = get(values[i], 'NO CG').toUpperCase().replace(/\s+/g, ' ');
    if (!code) continue;
    out.push({
      row: i + 1, code: code, place: get(values[i], 'NAMA TEMPAT'), address: get(values[i], 'ALAMAT'),
      link: get(values[i], 'LINK MAPS'), lat: get(values[i], 'LATITUDE'), lng: get(values[i], 'LONGITUDE'),
      updatedBy: get(values[i], 'DIPERBARUI OLEH'), updatedAt: get(values[i], 'DIPERBARUI'),
    });
  }
  return out;
}

/** One entry per CG in the database: code, coach, CGL, schedule, area, size. */
function cgSummaries_(ss) {
  const people = parseDatabase_(getDatabaseSheet_(ss).getDataRange().getDisplayValues()).people;
  const byCode = {};
  const order = [];
  people.forEach(function (p) {
    if (p.isCoach || !p.cg || p.cg === '-') return;
    let g = byCode[p.cg];
    if (!g) {
      g = byCode[p.cg] = { code: p.cg, coach: p.coach, cgl: '', time: p.cgTime, mode: p.cgStatus, area: p.cgArea, size: 0 };
      order.push(g);
    }
    g.size++;
    if (p.position === 'CGL' && !g.cglRow) {
      // The CGL's row is the source of truth for the CG's schedule, status and area.
      g.cglRow = true;
      g.cgl = p.name;
      if (p.cgTime) g.time = p.cgTime;
      if (p.cgStatus) g.mode = p.cgStatus;
      if (p.cgArea) g.area = p.cgArea;
    }
    if (!g.cgl && p.cgl) g.cgl = p.cgl;
    if (!g.time && p.cgTime) g.time = p.cgTime;
    if (!g.mode && p.cgStatus) g.mode = p.cgStatus;
    if (!g.area && p.cgArea) g.area = p.cgArea;
  });
  return order;
}

function locationsPayload_(ss, session) {
  const cgs = cgSummaries_(ss);
  const rows = readLocationRows_(getLocationSheet_(ss, false));
  const byCode = {};
  rows.forEach(function (r) { byCode[r.code] = r; });
  let budget = LOCATION.MAX_GEOCODES_PER_REQUEST;

  const list = cgs.map(function (g) {
    const r = byCode[g.code] || {};
    let coords = validCoords_(r.lat, r.lng);
    // Rows typed by hand may have only an address or link: work the coordinates out (cached).
    if (!coords && r.link) coords = coordsFromMapsLink_(r.link);
    if (!coords && r.address && budget-- > 0) coords = geocode_(r.address);
    return {
      code: g.code, coach: g.coach, cgl: g.cgl, time: g.time, mode: g.mode, area: g.area, size: g.size,
      online: /^online$/i.test(String(g.mode || '').trim()),
      place: r.place || '', address: r.address || '', link: r.link || '',
      lat: coords ? coords.lat : null, lng: coords ? coords.lng : null,
      updatedBy: r.updatedBy || '', updatedAt: r.updatedAt || '',
      canEdit: canAccessCoach_(session, g.coach) && hasPerm_(session, 'restructure'),
    };
  });

  // Domisili layer: how many of the people this session may see live in each area.
  const people = parseDatabase_(getDatabaseSheet_(ss).getDataRange().getDisplayValues()).people
    .filter(function (p) { return !p.isCoach && p.domicile && canAccessCoach_(session, p.coach); });
  const areas = {};
  people.forEach(function (p) {
    const key = norm_(p.domicile);
    const a = areas[key] || (areas[key] = { label: p.domicile, count: 0, cgs: {} });
    a.count++;
    a.cgs[p.cg || '-'] = (a.cgs[p.cg || '-'] || 0) + 1;
  });
  const homes = Object.keys(areas).map(function (k) {
    const a = areas[k];
    const coords = budget-- > 0 ? geocode_(a.label) : cachedGeocode_(a.label);
    return coords ? { label: a.label, count: a.count, cgs: a.cgs, lat: coords.lat, lng: coords.lng } : null;
  }).filter(Boolean);

  return { cgs: list, homes: homes, unplaced: Object.keys(areas).length - homes.length };
}

function validCoords_(lat, lng) {
  const a = parseFloat(lat), b = parseFloat(lng);
  if (!isFinite(a) || !isFinite(b) || Math.abs(a) > 90 || Math.abs(b) > 180 || (a === 0 && b === 0)) return null;
  return { lat: a, lng: b };
}

/** Reads coordinates out of the usual Google Maps URL shapes; expands maps.app.goo.gl short links. */
function coordsFromMapsLink_(link) {
  let url = String(link || '').trim();
  if (isShortLink_(url)) {
    const cacheKey = 'link:' + url;
    const props = PropertiesService.getScriptProperties();
    const cached = props.getProperty(cacheKey);
    if (cached) {
      url = cached;
    } else {
      try {
        // Only Google's short-link hosts are ever fetched; we stop at the first redirect elsewhere.
        for (let hop = 0; hop < 3 && isShortLink_(url); hop++) {
          const res = UrlFetchApp.fetch(url, { followRedirects: false, muteHttpExceptions: true });
          const loc = res.getHeaders().Location || res.getHeaders().location;
          if (!loc) break;
          url = String(loc);
        }
        props.setProperty(cacheKey, url.slice(0, 8000));
      } catch (e) {
        return null;
      }
    }
  }
  const text = decodeURIComponent(url);
  const patterns = [
    /!3d(-?\d+(?:\.\d+)?)!4d(-?\d+(?:\.\d+)?)/, // the dropped pin itself (most precise)
    /@(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)/, // map centre
    /[?&](?:q|query|ll|destination|center)=(-?\d+(?:\.\d+)?),\s*(-?\d+(?:\.\d+)?)/,
  ];
  for (let i = 0; i < patterns.length; i++) {
    const m = patterns[i].exec(text);
    if (m) {
      const c = validCoords_(m[1], m[2]);
      if (c) return c;
    }
  }
  return null;
}

function isShortLink_(url) {
  const host = ((/^https:\/\/([^/?#]+)/i.exec(String(url)) || [])[1] || '').toLowerCase();
  return LOCATION.SHORT_LINK_HOSTS.indexOf(host) >= 0;
}

function geocodeKey_(text) {
  return 'geo:' + norm_(text).slice(0, 200);
}

function cachedGeocode_(text) {
  const v = PropertiesService.getScriptProperties().getProperty(geocodeKey_(text));
  if (!v || v === 'x') return null;
  const p = v.split(',');
  return validCoords_(p[0], p[1]);
}

/** Address → coordinates with the built-in Maps service, cached forever ('x' = not found). */
function geocode_(text) {
  text = String(text || '').trim();
  if (!text) return null;
  const props = PropertiesService.getScriptProperties();
  const key = geocodeKey_(text);
  const hit = props.getProperty(key);
  if (hit) return hit === 'x' ? null : cachedGeocode_(text);
  let coords = null;
  try {
    const b = LOCATION.BOUNDS;
    const res = Maps.newGeocoder().setRegion('id').setLanguage('id').setBounds(b[0], b[1], b[2], b[3])
      .geocode(/indonesia/i.test(text) ? text : text + ', Indonesia');
    const r = res && res.status === 'OK' && res.results && res.results[0];
    if (r) coords = validCoords_(r.geometry.location.lat, r.geometry.location.lng);
  } catch (e) {
    console.warn('Geocode failed for "' + text + '": ' + e);
    return null; // don't cache errors (e.g. quota), try again later
  }
  props.setProperty(key, coords ? coords.lat.toFixed(6) + ',' + coords.lng.toFixed(6) : 'x');
  return coords;
}
