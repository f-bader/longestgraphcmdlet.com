import { XMLParser } from 'fast-xml-parser';

const GALLERY_BASE = 'https://www.powershellgallery.com/api/v2';
const ROOT_PACKAGE = 'Microsoft.Graph';
const HISTORY_LIMIT = 10;
const CACHE_TTL_MS = 6 * 60 * 60 * 1000;

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseTagValue: true,
  trimValues: true,
});

const memoryCache = {
  data: null,
  expiresAt: 0,
  inFlight: null,
};

const htmlPage = `<!doctype html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Longest Microsoft Graph Cmdlet</title>
  <style>
    :root {
      color-scheme: light dark;
      --bg: #0b1020;
      --surface: #121a33;
      --surface-2: #1a254a;
      --text: #f5f7ff;
      --muted: #a9b3d9;
      --accent: #7aa2ff;
      --accent-2: #4de2c5;
      --error: #ff8f8f;
    }

    * { box-sizing: border-box; }

    body {
      margin: 0;
      font-family: Inter, system-ui, -apple-system, Segoe UI, Roboto, sans-serif;
      background: radial-gradient(circle at top right, #22356f 0%, var(--bg) 35%);
      color: var(--text);
      min-height: 100vh;
      padding: 2rem 1rem;
    }

    .container {
      max-width: 980px;
      margin: 0 auto;
      display: grid;
      gap: 1rem;
    }

    .card {
      background: linear-gradient(145deg, var(--surface), var(--surface-2));
      border: 1px solid rgba(255, 255, 255, 0.08);
      border-radius: 16px;
      padding: 1.25rem;
      box-shadow: 0 16px 32px rgba(0, 0, 0, 0.3);
    }

    h1 {
      margin: 0 0 0.25rem;
      font-size: clamp(1.3rem, 3vw, 2rem);
    }

    p { margin: 0; color: var(--muted); }

    .metric {
      font-size: clamp(1.4rem, 4vw, 2.3rem);
      margin: 0.7rem 0 0.4rem;
      color: var(--accent-2);
      font-weight: 700;
      word-break: break-word;
    }

    .sub {
      color: var(--muted);
      font-size: 0.95rem;
    }

    table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 0.75rem;
      font-size: 0.95rem;
    }

    th, td {
      text-align: left;
      padding: 0.65rem 0.5rem;
      border-bottom: 1px solid rgba(255,255,255,0.08);
      vertical-align: top;
    }

    th { color: var(--muted); font-weight: 600; }

    .badge {
      display: inline-block;
      padding: 0.2rem 0.55rem;
      border-radius: 999px;
      font-size: 0.8rem;
      border: 1px solid rgba(255,255,255,0.12);
      color: var(--accent);
    }

    .error { color: var(--error); }
  </style>
</head>
<body>
  <main class="container">
    <section class="card">
      <h1>Longest Microsoft Graph PowerShell Function/Cmdlet</h1>
      <p>Calculated from <code>Microsoft.Graph</code> metadata and its module dependencies on PowerShell Gallery.</p>
      <div id="status" class="sub" style="margin-top:0.5rem;">Loading latest data…</div>
      <div id="current"></div>
    </section>

    <section class="card">
      <h2 style="margin:0; font-size:1.15rem;">Change History (latest ${HISTORY_LIMIT})</h2>
      <div id="history"></div>
    </section>
  </main>

  <script>
    const statusNode = document.getElementById('status');
    const currentNode = document.getElementById('current');
    const historyNode = document.getElementById('history');

    const toDate = (value) => {
      if (!value) return 'unknown';
      const date = new Date(value);
      if (Number.isNaN(date.valueOf())) return value;
      return date.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
    };

    async function run() {
      try {
        const response = await fetch('/api/longest');
        if (!response.ok) throw new Error('Failed with status ' + response.status);
        const data = await response.json();

        statusNode.textContent = 'Generated ' + toDate(data.generatedAt) + ' · Versions analyzed: ' + data.totalVersionsAnalyzed;

        currentNode.innerHTML = [
          '<div class="metric">' + data.current.name + '</div>',
          '<div class="sub">Length: ' + data.current.length + ' characters</div>',
          '<div class="sub">Package: <span class="badge">' + data.current.packageId + '</span></div>',
          '<div class="sub">Longest since: ' + toDate(data.current.sincePublished) + ' (Microsoft.Graph ' + data.current.sinceVersion + ')</div>',
          '<div class="sub" style="margin-top:0.4rem;">Previous longest: ' + (data.previous?.name || 'n/a') + '</div>',
        ].join('');

        if (!data.history.length) {
          historyNode.innerHTML = '<p class="sub">No history available.</p>';
          return;
        }

        const rows = data.history.map((item) =>
          '<tr>' +
            '<td>' + toDate(item.publishedAt) + '<br><span class="sub">v' + item.version + '</span></td>' +
            '<td><strong>' + item.name + '</strong><br><span class="sub">' + item.packageId + '</span></td>' +
            '<td>' + (item.previousName || 'n/a') + '</td>' +
          '</tr>'
        ).join('');

        historyNode.innerHTML =
          '<table>' +
            '<thead>' +
              '<tr>' +
                '<th>When it changed</th>' +
                '<th>New longest</th>' +
                '<th>Previous longest</th>' +
              '</tr>' +
            '</thead>' +
            '<tbody>' + rows + '</tbody>' +
          '</table>';
      } catch (error) {
        statusNode.classList.add('error');
        statusNode.textContent = 'Unable to load data right now.';
        currentNode.innerHTML = '<p class="error">' + error.message + '</p>';
      }
    }

    run();
  </script>
</body>
</html>`;

function asArray(value) {
  if (!value) return [];
  return Array.isArray(value) ? value : [value];
}

function pickKey(object, plainName) {
  if (!object || typeof object !== 'object') return null;

  for (const key of Object.keys(object)) {
    if (key === plainName || key.endsWith(`:${plainName}`)) {
      return object[key];
    }
  }

  return null;
}

function parseCsvCommands(rawValue) {
  if (!rawValue || typeof rawValue !== 'string') return [];

  return rawValue
    .split(/[|,;\s]+/g)
    .map((item) => item.trim())
    .filter((item) => item.length > 0);
}

function pickLongest(commands) {
  let best = null;

  for (const command of commands) {
    if (
      !best ||
      command.length > best.length ||
      (command.length === best.length && command < best)
    ) {
      best = command;
    }
  }

  return best;
}

function versionParts(version) {
  return String(version)
    .split('.')
    .map((part) => Number(part.replace(/\D.*$/, '')) || 0);
}

function compareVersions(a, b) {
  const left = versionParts(a);
  const right = versionParts(b);
  const length = Math.max(left.length, right.length);

  for (let i = 0; i < length; i += 1) {
    const l = left[i] ?? 0;
    const r = right[i] ?? 0;

    if (l !== r) {
      return l - r;
    }
  }

  return String(a).localeCompare(String(b));
}

function normalizeVersion(versionSpec) {
  if (!versionSpec) return null;
  const match = String(versionSpec).match(/\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?/);
  return match ? match[0] : null;
}

function parseDependencyString(raw) {
  if (!raw || typeof raw !== 'string') return [];

  return raw
    .split('|')
    .map((entry) => entry.trim())
    .filter(Boolean)
    .map((entry) => {
      const [id, versionSpec] = entry.split(':');
      return {
        id: id?.trim(),
        version: normalizeVersion(versionSpec?.trim()),
      };
    })
    .filter((dep) => dep.id && dep.version);
}

function extractProperties(document) {
  if (!document || typeof document !== 'object') return null;

  if (document.entry) {
    const content = document.entry.content;
    return (
      pickKey(content, 'properties') ||
      pickKey(document.entry, 'properties') ||
      null
    );
  }

  const entries = asArray(document.feed?.entry);
  if (entries.length === 0) return null;

  const content = entries[0]?.content;
  return pickKey(content, 'properties') || pickKey(entries[0], 'properties') || null;
}

function parsePackageProperties(xmlText) {
  const document = parser.parse(xmlText);
  const properties = extractProperties(document);
  if (!properties) {
    throw new Error('Unable to read package metadata properties.');
  }

  return {
    id: pickKey(properties, 'Id'),
    version: pickKey(properties, 'Version'),
    published: pickKey(properties, 'Published'),
    dependencies: pickKey(properties, 'Dependencies') ?? '',
    cmdlets: pickKey(properties, 'Cmdlets') ?? '',
    functions: pickKey(properties, 'Functions') ?? '',
  };
}

function parseFeed(xmlText) {
  const document = parser.parse(xmlText);
  const entries = asArray(document.feed?.entry).map((entry) => {
    const properties = pickKey(entry.content, 'properties') || pickKey(entry, 'properties') || {};

    return {
      id: pickKey(properties, 'Id'),
      version: pickKey(properties, 'Version'),
      published: pickKey(properties, 'Published'),
      dependencies: pickKey(properties, 'Dependencies') ?? '',
      cmdlets: pickKey(properties, 'Cmdlets') ?? '',
      functions: pickKey(properties, 'Functions') ?? '',
    };
  });

  const links = asArray(document.feed?.link);
  const nextLink = links.find((link) => link?.['@_rel'] === 'next')?.['@_href'] ?? null;

  return { entries, nextLink };
}

async function fetchText(url) {
  const response = await fetch(url, {
    headers: { accept: 'application/atom+xml,application/xml,text/xml' },
  });

  if (!response.ok) {
    throw new Error(`PowerShell Gallery request failed (${response.status}) for ${url}`);
  }

  return response.text();
}

function parseCommandsFromMetadata(metadata) {
  const cmdlets = parseCsvCommands(metadata.cmdlets);
  const functions = parseCsvCommands(metadata.functions);
  return [...new Set([...cmdlets, ...functions])];
}

async function getAllRootPackageVersions() {
  const all = [];
  let nextUrl = `${GALLERY_BASE}/FindPackagesById()?id='${ROOT_PACKAGE}'`;
  let guard = 0;

  while (nextUrl && guard < 50) {
    const xml = await fetchText(nextUrl);
    const parsed = parseFeed(xml);
    all.push(...parsed.entries);
    nextUrl = parsed.nextLink;
    guard += 1;
  }

  if (all.length === 0) {
    throw new Error('No Microsoft.Graph package versions were returned from PowerShell Gallery.');
  }

  return all
    .filter((item) => item.version)
    .sort((a, b) => compareVersions(a.version, b.version));
}

async function getPackageMetadata(packageId, version, memo) {
  const key = `${packageId}@${version}`;
  if (memo.has(key)) return memo.get(key);

  const url = `${GALLERY_BASE}/Packages(Id='${encodeURIComponent(packageId)}',Version='${encodeURIComponent(version)}')`;
  const promise = fetchText(url).then(parsePackageProperties);
  memo.set(key, promise);

  return promise;
}

function buildResponse(winnersByVersion) {
  const transitions = [];
  let last = null;

  for (const item of winnersByVersion) {
    if (!item.longestName) continue;

    if (!last || last.name !== item.longestName) {
      const transition = {
        version: item.version,
        publishedAt: item.published,
        name: item.longestName,
        packageId: item.packageId,
        previousName: last?.name ?? null,
      };

      transitions.push(transition);
      last = transition;
    }
  }

  const current = transitions[transitions.length - 1];
  const previous = transitions[transitions.length - 2] ?? null;

  return {
    generatedAt: new Date().toISOString(),
    source: 'PowerShell Gallery: Microsoft.Graph + dependencies',
    totalVersionsAnalyzed: winnersByVersion.length,
    current: {
      name: current?.name ?? null,
      length: current?.name?.length ?? 0,
      packageId: current?.packageId ?? null,
      sinceVersion: current?.version ?? null,
      sincePublished: current?.publishedAt ?? null,
    },
    previous,
    history: transitions.slice(-HISTORY_LIMIT).reverse(),
  };
}

async function computeLongestCmdletHistory() {
  const roots = await getAllRootPackageVersions();
  const packageMemo = new Map();
  const winners = [];

  for (const root of roots) {
    const rootCommands = parseCommandsFromMetadata(root);
    let candidates = rootCommands;
    let sourcePackage = root.id || ROOT_PACKAGE;

    if (candidates.length === 0) {
      const dependencies = parseDependencyString(root.dependencies)
        .filter((dep) => dep.id.startsWith('Microsoft.Graph'));

      const dependencyMetadata = await Promise.all(
        dependencies.map((dep) => getPackageMetadata(dep.id, dep.version, packageMemo))
      );

      candidates = dependencyMetadata.flatMap(parseCommandsFromMetadata);
      sourcePackage = ROOT_PACKAGE;
    }

    const longestName = pickLongest(candidates);
    winners.push({
      version: root.version,
      published: root.published,
      longestName,
      packageId: sourcePackage,
    });
  }

  return buildResponse(winners);
}

async function getCachedData() {
  const now = Date.now();

  if (memoryCache.data && memoryCache.expiresAt > now) {
    return memoryCache.data;
  }

  if (!memoryCache.inFlight) {
    memoryCache.inFlight = computeLongestCmdletHistory()
      .then((data) => {
        memoryCache.data = data;
        memoryCache.expiresAt = now + CACHE_TTL_MS;
        return data;
      })
      .finally(() => {
        memoryCache.inFlight = null;
      });
  }

  return memoryCache.inFlight;
}

function jsonResponse(payload, status = 200) {
  return new Response(JSON.stringify(payload, null, 2), {
    status,
    headers: {
      'content-type': 'application/json; charset=utf-8',
      'cache-control': 'public, max-age=300',
    },
  });
}

export default {
  async fetch(request) {
    const url = new URL(request.url);

    if (url.pathname === '/api/longest') {
      try {
        const data = await getCachedData();
        return jsonResponse(data);
      } catch (error) {
        return jsonResponse(
          {
            error: 'Unable to fetch PowerShell Gallery data.',
            details: error.message,
          },
          500,
        );
      }
    }

    if (url.pathname === '/api/refresh' && request.method === 'POST') {
      memoryCache.data = null;
      memoryCache.expiresAt = 0;
      return jsonResponse({ ok: true });
    }

    return new Response(htmlPage, {
      headers: {
        'content-type': 'text/html; charset=utf-8',
        'cache-control': 'public, max-age=60',
      },
    });
  },
};

export const _internals = {
  parseDependencyString,
  parseCommandsFromMetadata,
  pickLongest,
  compareVersions,
  buildResponse,
};
