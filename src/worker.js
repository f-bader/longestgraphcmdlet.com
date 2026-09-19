import snapshot from './snapshot.js';

const escapeHtml = (value) => String(value ?? '').replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&#39;');
const dateFormatter = new Intl.DateTimeFormat('en', { year: 'numeric', month: 'short', day: 'numeric', timeZone: 'UTC' });
function dateLabel(value) {
  const date = new Date(typeof value === 'object' && value ? value['#text'] : value ?? '');
  return Number.isNaN(date.valueOf()) || date.getUTCFullYear() < 2000 ? 'Date unavailable' : dateFormatter.format(date);
}

export function renderPage(data) {
  const current = data.current;
  const topRows = (data.topLongest ?? []).map((item, index) => `<tr><td class="rank">${String(index + 1).padStart(2, '0')}</td><td><code>${escapeHtml(item.name)}</code><span class="detail">${escapeHtml(item.packageId)} · Graph ${escapeHtml(item.versionFound)}</span></td><td class="number">${escapeHtml(item.length)}</td></tr>`).join('');
  const historyRows = (data.history ?? []).map((item) => `<li><div class="history-date"><time>${escapeHtml(dateLabel(item.publishedAt))}</time><span class="detail">Graph ${escapeHtml(item.version)}</span></div><div><code>${escapeHtml(item.name)}</code><span class="detail">${escapeHtml(item.packageId)}</span></div></li>`).join('');
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Longest Microsoft Graph Cmdlet</title>
  <meta name="description" content="A small tribute to very long commands. Explore the longest Microsoft Graph PowerShell cmdlet, the all-time top ten, and the history behind them.">
  <meta property="og:title" content="Longest Microsoft Graph Cmdlet">
  <meta property="og:description" content="A small tribute to very long commands. The current longest, the all-time top ten, and their history.">
  <meta name="theme-color" content="#111313">
  <style>
    :root { color-scheme:dark; --bg:#111313; --text:#eff1ed; --muted:#a3aaa4; --line:#343934; --accent:#c9f77b; }
    * { box-sizing:border-box; }
    body { margin:0; background:var(--bg); color:var(--text); font:15px/1.6 system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif; }
    main,footer { width:min(100% - 48px,1080px); margin-inline:auto; }
    a { color:inherit; text-underline-offset:4px; }
    a:hover { color:var(--accent); }
    a:focus-visible { outline:2px solid var(--accent); outline-offset:5px; }
    .eyebrow { text-transform:uppercase; letter-spacing:.13em; font-size:11px; font-weight:650; }
    p { color:var(--muted); margin:0; }
    .hero { padding:40px 0; border-bottom:1px solid var(--line); }
    .hero-heading { display:flex; align-items:baseline; justify-content:space-between; gap:20px; margin-bottom:25px; }
    .eyebrow { color:var(--muted); margin:0; }
    .length { font-size:14px; color:var(--muted); white-space:nowrap; }
    .length strong { color:var(--accent); font-size:40px; font-weight:500; letter-spacing:-.06em; margin-right:7px; }
    code { font-family:ui-monospace,SFMono-Regular,Consolas,"Liberation Mono",monospace; overflow-wrap:anywhere; }
    .winner { display:block; color:var(--accent); font-size:clamp(20px,3.2vw,34px); line-height:1.5; letter-spacing:-.025em; }
    .metadata { display:flex; flex-wrap:wrap; gap:18px 50px; margin-top:28px; }
    .metadata p { font-size:13px; overflow-wrap:anywhere; }
    .metadata .label { display:block; font-size:10px; letter-spacing:.1em; text-transform:uppercase; margin-bottom:5px; }
    .metadata .value { color:var(--text); }
    section { padding-top:48px; }
    .section-heading { display:flex; justify-content:space-between; align-items:baseline; gap:16px; margin-bottom:20px; }
    h2 { margin:0; font-size:22px; font-weight:500; letter-spacing:-.025em; }
    .section-heading p { font-size:12px; }
    table { width:100%; table-layout:fixed; border-collapse:collapse; }
    th { color:var(--muted); font-size:10px; font-weight:500; letter-spacing:.08em; text-transform:uppercase; text-align:left; }
    th,td { padding:18px 0; border-bottom:1px solid var(--line); vertical-align:top; }
    td code { font-size:13px; }
    .rank { width:48px; color:var(--muted); font-size:12px; font-variant-numeric:tabular-nums; }
    .number { width:65px; text-align:right; font-variant-numeric:tabular-nums; }
    .detail { display:block; color:var(--muted); font-size:11px; margin-top:6px; overflow-wrap:anywhere; }
    .history { padding:0; margin:0; list-style:none; }
    .history li { display:grid; grid-template-columns:145px minmax(0,1fr); gap:25px; padding:23px 0; border-top:1px solid var(--line); }
    .history code { font-size:13px; }
    .history-date { font-size:12px; }
    footer { display:flex; justify-content:space-between; flex-wrap:wrap; gap:12px; margin-top:55px; padding:24px 0 36px; border-top:1px solid var(--line); color:var(--muted); font-size:11px; }
    @media(max-width:600px) { main,footer { width:calc(100% - 32px); } .hero-heading { gap:10px; } .length strong { font-size:30px; } .length { font-size:11px; } .section-heading { align-items:flex-start; flex-direction:column; gap:4px; } .rank { width:30px; } .number { width:46px; } td code,.history code { font-size:12px; } .history li { grid-template-columns:minmax(0,1fr); gap:12px; } .history-date { display:flex; align-items:baseline; gap:12px; } .history-date .detail { margin:0; } }
  </style>
</head>
<body>
  <main>
    <section class="hero" aria-labelledby="current-title">
      <div class="hero-heading"><h1 class="eyebrow" id="current-title">The current longest cmdlet</h1>${current?.name ? `<span class="length"><strong>${escapeHtml(current.length)}</strong> characters</span>` : ''}</div>
      ${current?.name ? `<code class="winner">${escapeHtml(current.name)}</code><div class="metadata"><p><span class="label">Module</span><span class="value">${escapeHtml(current.packageId)}</span></p><p><span class="label">Longest since</span><span class="value">${escapeHtml(dateLabel(current.sincePublished))} · Graph ${escapeHtml(current.sinceVersion)}</span></p></div>` : '<p>No command data is available yet.</p>'}
    </section>
    <section aria-labelledby="top-title"><div class="section-heading"><h2 id="top-title">The all-time top 10</h2><p>Unique commands, ranked by length</p></div>${topRows ? `<table><thead><tr><th class="rank" scope="col">Rank</th><th scope="col">Command / module</th><th class="number" scope="col">Chars</th></tr></thead><tbody>${topRows}</tbody></table>` : '<p>No rankings available.</p>'}</section>
    <section aria-labelledby="history-title"><div class="section-heading"><h2 id="history-title">A history of long commands</h2><p>The latest 10 changes to the lead</p></div>${historyRows ? `<ol class="history">${historyRows}</ol>` : '<p>No history available.</p>'}</section>
  </main>
  <footer><span>${escapeHtml(data.totalVersionsAnalyzed)} releases analyzed · Updated ${escapeHtml(dateLabel(data.generatedAt))}</span><span><a href="https://www.powershellgallery.com/packages/Microsoft.Graph">PowerShell Gallery</a> · <a href="/api/longest">JSON API</a></span></footer>
</body>
</html>`;
}

const htmlPage = renderPage(snapshot);
const snapshotJson = JSON.stringify(snapshot);
function jsonResponse(payload, status = 200, cacheControl = 'public, max-age=300, stale-while-revalidate=86400') {
  return new Response(typeof payload === 'string' ? payload : JSON.stringify(payload), {
    status,
    headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': cacheControl, 'x-content-type-options': 'nosniff' },
  });
}

export default {
  fetch(request) {
    const url = new URL(request.url);
    const readable = request.method === 'GET' || request.method === 'HEAD';
    if (url.pathname === '/api/longest' && readable) {
      const response = jsonResponse(snapshotJson);
      return request.method === 'HEAD' ? new Response(null, response) : response;
    }
    if (url.pathname === '/api/refresh') return jsonResponse({ error: 'Refresh is managed by the GitHub Actions publishing workflow.' }, 410, 'private, no-store');
    if (url.pathname === '/' && readable) {
      return new Response(request.method === 'HEAD' ? null : htmlPage, {
        headers: {
          'content-type': 'text/html; charset=utf-8',
          'cache-control': 'public, max-age=60',
          'x-content-type-options': 'nosniff',
          'content-security-policy': "default-src 'none'; style-src 'unsafe-inline'; base-uri 'none'; frame-ancestors 'none'; form-action 'none'",
          'referrer-policy': 'strict-origin-when-cross-origin',
        },
      });
    }
    return jsonResponse({ error: 'Not found.' }, 404, 'private, no-store');
  },
};

export const _internals = { jsonResponse };
