import snapshot from './snapshot.js';

const HISTORY_LIMIT = 10;

const htmlPage = `<!doctype html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Longest Microsoft Graph Cmdlet</title>
  <style>
    :root { color-scheme: light dark; --bg:#0b1020; --surface:#121a33; --surface-2:#1a254a; --text:#f5f7ff; --muted:#a9b3d9; --accent:#7aa2ff; --accent-2:#4de2c5; --error:#ff8f8f; }
    * { box-sizing: border-box; }
    body { margin:0; min-height:100vh; padding:2rem 1rem; font-family:Inter,system-ui,-apple-system,Segoe UI,Roboto,sans-serif; background:radial-gradient(circle at top right,#22356f 0%,var(--bg) 35%); color:var(--text); }
    .container { max-width:980px; margin:0 auto; display:grid; gap:1rem; }
    .card { padding:1.25rem; border:1px solid rgba(255,255,255,.08); border-radius:16px; background:linear-gradient(145deg,var(--surface),var(--surface-2)); box-shadow:0 16px 32px rgba(0,0,0,.3); }
    h1 { margin:0 0 .25rem; font-size:clamp(1.3rem,3vw,2rem); }
    p { margin:0; color:var(--muted); }
    .metric { margin:.7rem 0 .4rem; color:var(--accent-2); font-size:clamp(1.4rem,4vw,2.3rem); font-weight:700; overflow-wrap:anywhere; }
    .sub { color:var(--muted); font-size:.95rem; }
    table { width:100%; margin-top:.75rem; border-collapse:collapse; font-size:.95rem; }
    th,td { padding:.65rem .5rem; border-bottom:1px solid rgba(255,255,255,.08); text-align:left; vertical-align:top; }
    th { color:var(--muted); font-weight:600; }
    .badge { display:inline-block; padding:.2rem .55rem; border:1px solid rgba(255,255,255,.12); border-radius:999px; color:var(--accent); font-size:.8rem; }
    .error { color:var(--error); }
  </style>
</head>
<body>
  <main class="container">
    <section class="card">
      <h1>Longest Microsoft Graph PowerShell Function/Cmdlet</h1>
      <p>Calculated from Microsoft.Graph metadata and module dependencies on PowerShell Gallery.</p>
      <div id="status" class="sub" style="margin-top:.5rem">Loading latest data…</div>
      <div id="current"></div>
    </section>
    <section class="card">
      <h2 style="margin:0;font-size:1.15rem">Change History (latest ${HISTORY_LIMIT})</h2>
      <div id="history"></div>
    </section>
  </main>
  <script>
    const statusNode = document.getElementById('status');
    const currentNode = document.getElementById('current');
    const historyNode = document.getElementById('history');
    const esc = (value) => String(value ?? '').replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;').replaceAll("'",'&#39;');
    const toDate = (value) => { if (!value) return 'unknown'; const date = new Date(value); return Number.isNaN(date.valueOf()) ? value : date.toLocaleDateString(undefined,{year:'numeric',month:'short',day:'numeric'}); };
    async function run() {
      try {
        const response = await fetch('/api/longest');
        if (!response.ok) throw new Error('Failed with status ' + response.status);
        const data = await response.json();
        statusNode.textContent = 'Generated ' + toDate(data.generatedAt) + ' · Versions analyzed: ' + data.totalVersionsAnalyzed;
        if (!data.current?.name) {
          currentNode.innerHTML = '<p class="sub">No longest command could be resolved yet.</p>';
          historyNode.innerHTML = '<p class="sub">No history available.</p>';
          return;
        }
        currentNode.innerHTML = [
          '<div class="metric">' + esc(data.current.name) + '</div>',
          '<div class="sub">Length: ' + data.current.length + ' characters</div>',
          '<div class="sub">Package: <span class="badge">' + esc(data.current.packageId) + '</span></div>',
          '<div class="sub">Longest since: ' + esc(toDate(data.current.sincePublished)) + ' (Microsoft.Graph ' + esc(data.current.sinceVersion) + ')</div>',
          '<div class="sub" style="margin-top:.4rem">Previous longest: ' + esc(data.previous?.name || 'n/a') + '</div>',
        ].join('');
        if (!data.history?.length) { historyNode.innerHTML = '<p class="sub">No history available.</p>'; return; }
        const rows = data.history.map((item) => '<tr><th scope="row">' + esc(toDate(item.publishedAt)) + '<br><span class="sub">v' + esc(item.version) + '</span></th><td><strong>' + esc(item.name) + '</strong><br><span class="sub">' + esc(item.packageId) + '</span></td><td>' + esc(item.previousName || 'n/a') + '</td></tr>').join('');
        historyNode.innerHTML = '<table><thead><tr><th scope="col">When it changed</th><th scope="col">New longest</th><th scope="col">Previous longest</th></tr></thead><tbody>' + rows + '</tbody></table>';
      } catch (error) {
        statusNode.classList.add('error');
        statusNode.textContent = 'Unable to load the published snapshot.';
        currentNode.innerHTML = '<p class="error">' + esc(error.message) + '</p>';
      }
    }
    run();
  </script>
</body>
</html>`;

function jsonResponse(payload, status = 200, cacheControl = 'public, max-age=300, stale-while-revalidate=86400') {
  return new Response(JSON.stringify(payload, null, 2), {
    status,
    headers: {
      'content-type': 'application/json; charset=utf-8',
      'cache-control': cacheControl,
    },
  });
}

export default {
  fetch(request) {
    const url = new URL(request.url);

    if (url.pathname === '/api/longest' && request.method === 'GET') {
      return jsonResponse(snapshot);
    }

    if (url.pathname === '/api/refresh') {
      return jsonResponse(
        { error: 'Refresh is managed by the GitHub Actions publishing workflow.' },
        410,
        'private, no-store',
      );
    }

    if (url.pathname === '/' && request.method === 'GET') {
      return new Response(htmlPage, {
        headers: {
          'content-type': 'text/html; charset=utf-8',
          'cache-control': 'public, max-age=60',
        },
      });
    }

    return jsonResponse({ error: 'Not found.' }, 404, 'private, no-store');
  },
};

export const _internals = { jsonResponse };
