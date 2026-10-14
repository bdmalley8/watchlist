// Movie Watchlist — Google Sheet backend.
// Lives in the watchlist Google Sheet (Extensions → Apps Script). Only the Cloudflare Worker in
// worker/ calls it: every request must carry `key` matching the WORKER_KEY Script Property
// (Project Settings → Script Properties). Requests without it are refused, so the /exec URL is
// useless on its own.

const FAMILY_HEADERS = ['title','year','posterUrl','streaming','lastChecked','addedAt','rating','imdbId','rtScore','imdbScore','mcScore','releaseDate','digitalDate','runtime'];
const KIDS_HEADERS   = ['title','year','posterUrl','streaming','lastChecked','addedAt','rating','csmRating','parentsRating','imdbId','rtScore','imdbScore','mcScore','releaseDate','digitalDate','runtime'];
const TV_HEADERS     = ['title','firstAirYear','posterUrl','streaming','lastChecked','addedAt','tmdbId','imdbId','totalSeasons','seasonsWatched','seasons','status','watchedBy','rtScore','imdbScore','mcScore'];

const LISTS = { family: 'Family', kids: 'Kids', tv: 'TV' };

function getSheet(list) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LISTS[list] || 'Family');
}

function ensureHeaders(sheet, list) {
  const expected = list === 'kids' ? KIDS_HEADERS : list === 'tv' ? TV_HEADERS : FAMILY_HEADERS;
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(expected);
    return;
  }
  const current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(String);
  expected.forEach(h => {
    if (!current.includes(h)) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h);
    }
  });
}

function readRows(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { if (h) obj[String(h)] = row[i] !== undefined ? row[i] : ''; });
      return obj;
    })
    .filter(r => r.title);
}

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// True only for requests from the Worker. Refuses everything if WORKER_KEY isn't set.
function fromWorker(e) {
  const key = PropertiesService.getScriptProperties().getProperty('WORKER_KEY');
  return !!key && e.parameter.key === key;
}

function doGet(e) {
  if (!fromWorker(e)) return json({ error: 'unauthorized' });
  const action = e.parameter.action || 'get';
  const list   = (e.parameter.list || 'family').toLowerCase();

  if (action === 'csm') {
    const title = e.parameter.title || '';
    const year  = e.parameter.year  || '';
    return json(fetchCsm(title, year));
  }

  if (action === 'all') {
    const out = {};
    Object.keys(LISTS).forEach(k => {
      const sheet = getSheet(k);
      if (!sheet) { out[k] = []; return; }
      ensureHeaders(sheet, k);
      out[k] = readRows(sheet);
    });
    return json(out);
  }

  const sheet = getSheet(list);
  ensureHeaders(sheet, list);
  return json(readRows(sheet));
}

function doPost(e) {
  if (!fromWorker(e)) return json({ error: 'unauthorized' });
  const action = e.parameter.action;
  const list   = (e.parameter.list || 'family').toLowerCase();
  const sheet  = getSheet(list);
  ensureHeaders(sheet, list);
  const p = JSON.parse(e.postData.contents);

  if (action === 'add') {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const row = headers.map(h => h === 'addedAt' ? new Date().toISOString() : (p[String(h)] !== undefined ? p[String(h)] : ''));
    sheet.appendRow(row);
  }

  if (action === 'update') {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const titleCol = headers.map(String).indexOf('title');
    if (titleCol < 0) return json({ok: false});
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][titleCol]) === String(p.title)) {
        headers.forEach((h, j) => {
          const key = String(h);
          if (key && key !== 'title' && key !== 'addedAt' && p[key] !== undefined) {
            sheet.getRange(i + 1, j + 1).setValue(p[key]);
          }
        });
        break;
      }
    }
  }

  if (action === 'remove') {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const titleCol = headers.map(String).indexOf('title');
    if (titleCol < 0) return json({ok: false});
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][titleCol]) === String(p.title)) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
  }

  return json({ok: true});
}

// ── CSM Scraping ──────────────────────────────────────────────────────────────
function fetchCsm(title, year) {
  const slug = title.toLowerCase()
    .replace(/[^a-z0-9\s-]/g, '')
    .trim()
    .replace(/\s+/g, '-');

  const seen = new Set();
  const candidates = [slug, slug + '-' + year, slug + '-0'].filter(s => {
    if (seen.has(s)) return false;
    seen.add(s);
    return true;
  });

  for (const candidate of candidates) {
    try {
      const url = 'https://www.commonsensemedia.org/movie-reviews/' + candidate;
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) continue;
      const html = res.getContentText();

      let csmRating = '';
      const ldMatch = html.match(/"typicalAgeRange"\s*:\s*"([^"]+)"/);
      if (ldMatch) {
        const num = ldMatch[1].match(/\d+/);
        csmRating = num ? num[0] + '+' : ldMatch[1];
      }

      let parentsRating = '';
      const parentsMatch = html.match(/class="[^"]*rating--user[^"]*"[\s\S]*?(\d+)/);
      if (parentsMatch) parentsRating = parentsMatch[1] + '+';

      if (csmRating || parentsRating) return { csmRating, parentsRating };
    } catch(e) {}
  }

  return { csmRating: '', parentsRating: '' };
}
