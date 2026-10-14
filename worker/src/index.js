// Movie Watchlist API proxy (Cloudflare Worker).
//
// The public page calls this Worker instead of TMDB, OMDb, RapidAPI and the Apps Script directly,
// so no key or backend URL ships in the page. Changes to the lists require the family passphrase.
//
// Secrets (set with `npx wrangler secret put <NAME>`, never committed):
//   TMDB_KEY, OMDB_KEY, RAPIDAPI_KEY  — upstream API keys
//   APPS_SCRIPT_URL                   — the watchlist Apps Script /exec URL
//   WORKER_KEY                        — shared with the Apps Script's WORKER_KEY Script Property
//   WATCHLIST_PASSPHRASE              — the family passphrase for edits
//
// Routes:
//   GET  /sheet?action=get|all|csm&…   read the lists, or a Common Sense Media lookup
//   POST /sheet?action=add|update|remove&list=…   body {pass, data}
//   POST /auth                          body {pass} → 200 if correct, 401 if not
//   GET  /tmdb/<path>?…                 TMDB v3 (allowlisted paths only)
//   GET  /omdb?i=<imdbId>               OMDb scores
//   GET  /streaming?title=…             Streaming Availability (RapidAPI) title search

const ALLOWED_ORIGINS = [/^https:\/\/bdmalley8\.github\.io$/, /^http:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/];
const TMDB_PATHS = /^(search\/(movie|tv)|movie\/\d+(\/(watch\/providers|release_dates))?|tv\/\d+(\/(season\/\d+|watch\/providers|external_ids))?)$/;
const READ_ACTIONS = ['get', 'all', 'csm'];
const WRITE_ACTIONS = ['add', 'update', 'remove'];
const LISTS = ['family', 'kids', 'tv'];

export default {
  async fetch(request, env) {
    const origin = request.headers.get('Origin') || '';
    if (!ALLOWED_ORIGINS.some(re => re.test(origin))) return new Response('Forbidden', { status: 403 });
    const cors = { 'Access-Control-Allow-Origin': origin, 'Vary': 'Origin' };

    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: {
        ...cors, 'Access-Control-Allow-Methods': 'GET, POST', 'Access-Control-Allow-Headers': 'Content-Type', 'Access-Control-Max-Age': '86400',
      } });
    }

    const url = new URL(request.url);
    const ip = request.headers.get('CF-Connecting-IP') || 'unknown';
    try {
      const res = await route(request, url, env, ip);
      const headers = new Headers(res.headers);
      for (const [k, v] of Object.entries(cors)) headers.set(k, v);
      return new Response(res.body, { status: res.status, headers });
    } catch (e) {
      return json({ error: 'upstream request failed' }, 502, cors);
    }
  },
};

async function route(request, url, env, ip) {
  const path = url.pathname.replace(/^\/+/, '');
  const method = request.method;

  if (method === 'POST' && (path === 'auth' || path === 'sheet')) {
    if (!(await limited(env.WRITE_LIMITER, ip))) return json({ error: 'too many requests' }, 429);
    let body;
    try { body = JSON.parse(await request.text()); } catch (e) { return json({ error: 'invalid JSON body' }, 400); }
    if (!(await passOk(body && body.pass, env.WATCHLIST_PASSPHRASE))) return json({ error: 'wrong passphrase' }, 401);
    if (path === 'auth') return json({ ok: true });

    const action = url.searchParams.get('action');
    const list = url.searchParams.get('list') || 'family';
    if (!WRITE_ACTIONS.includes(action) || !LISTS.includes(list)) return json({ error: 'bad action or list' }, 400);
    return appsScript(env, { action, list }, JSON.stringify(body.data || {}));
  }

  if (method !== 'GET') return json({ error: 'method not allowed' }, 405);
  if (!(await limited(env.API_LIMITER, ip))) return json({ error: 'too many requests' }, 429);

  if (path === 'sheet') {
    const action = url.searchParams.get('action') || 'get';
    if (!READ_ACTIONS.includes(action)) return json({ error: 'bad action' }, 400);
    const params = { action };
    for (const k of ['list', 'title', 'year']) if (url.searchParams.has(k)) params[k] = url.searchParams.get(k);
    return appsScript(env, params);
  }

  if (path.startsWith('tmdb/')) {
    const tmdbPath = path.slice(5);
    if (!TMDB_PATHS.test(tmdbPath)) return json({ error: 'path not allowed' }, 404);
    const upstream = new URL('https://api.themoviedb.org/3/' + tmdbPath);
    for (const k of ['query', 'year', 'language', 'page']) if (url.searchParams.has(k)) upstream.searchParams.set(k, url.searchParams.get(k));
    upstream.searchParams.set('api_key', env.TMDB_KEY);
    return passthrough(await fetch(upstream));
  }

  if (path === 'omdb') {
    const imdbId = url.searchParams.get('i') || '';
    if (!/^tt\d+$/.test(imdbId)) return json({ error: 'bad imdb id' }, 400);
    return passthrough(await fetch('https://www.omdbapi.com/?i=' + imdbId + '&apikey=' + encodeURIComponent(env.OMDB_KEY)));
  }

  if (path === 'streaming') {
    // RapidAPI bills per call, so it gets its own tighter limit.
    if (!(await limited(env.STREAMING_LIMITER, ip))) return json({ error: 'too many requests' }, 429);
    const upstream = new URL('https://streaming-availability.p.rapidapi.com/shows/search/title');
    for (const k of ['title', 'country', 'show_type', 'output_language']) if (url.searchParams.has(k)) upstream.searchParams.set(k, url.searchParams.get(k));
    if (!upstream.searchParams.get('title')) return json({ error: 'title required' }, 400);
    return passthrough(await fetch(upstream, { headers: {
      'X-RapidAPI-Key': env.RAPIDAPI_KEY, 'X-RapidAPI-Host': 'streaming-availability.p.rapidapi.com',
    } }));
  }

  return json({ error: 'not found' }, 404);
}

// Apps Script answers POSTs with a 302 to a GET-able result URL; fetch follows it.
async function appsScript(env, params, body) {
  const target = new URL(env.APPS_SCRIPT_URL);
  for (const [k, v] of Object.entries(params)) target.searchParams.set(k, v);
  target.searchParams.set('key', env.WORKER_KEY);
  const res = await fetch(target, body === undefined ? {} : {
    method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body,
  });
  const text = await res.text();
  if (text.includes('"error":"unauthorized"')) return json({ error: 'sheet rejected the Worker key' }, 502);
  return new Response(text, { status: res.ok ? 200 : 502, headers: { 'Content-Type': 'application/json' } });
}

function passthrough(res) {
  return new Response(res.body, { status: res.status, headers: { 'Content-Type': res.headers.get('Content-Type') || 'application/json' } });
}

function json(data, status = 200, extra = {}) {
  return new Response(JSON.stringify(data), { status, headers: { 'Content-Type': 'application/json', ...extra } });
}

async function limited(limiter, ip) {
  if (!limiter) return true;
  const { success } = await limiter.limit({ key: ip });
  return success;
}

async function passOk(given, expected) {
  if (typeof given !== 'string' || !expected) return false;
  // Hash both sides so the comparison is constant-time and length-independent.
  const enc = new TextEncoder();
  const [a, b] = await Promise.all([given, expected].map(s => crypto.subtle.digest('SHA-256', enc.encode(s))));
  return crypto.subtle.timingSafeEqual(a, b);
}
