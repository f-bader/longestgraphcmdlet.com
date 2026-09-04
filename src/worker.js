import { XMLParser } from 'fast-xml-parser';

const GALLERY_BASE = 'https://www.powershellgallery.com/api/v2';
const ROOT_PACKAGE = 'Microsoft.Graph';
const HISTORY_LIMIT = 10;
const CACHE_TTL_MS = 24 * 60 * 60 * 1000;

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
  inFlightGeneration: null,
  generation: 0,
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
    const esc = (value) =>
      String(value ?? '')
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('\"', '&quot;')
        .replaceAll("'", '&#39;');

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

        if (!data.current || !data.current.name) {
          currentNode.innerHTML = '<p class="sub">No longest command could be resolved yet.</p>';
          historyNode.innerHTML = '<p class="sub">No history available.</p>';
          return;
        }

        currentNode.innerHTML = [
          '<div class="metric">' + esc(data.current.name) + '</div>',
          '<div class="sub">Length: ' + data.current.length + ' characters</div>',
          '<div class="sub">Package: <span class="badge">' + esc(data.current.packageId) + '</span></div>',
          '<div class="sub">Longest since: ' + esc(toDate(data.current.sincePublished)) + ' (Microsoft.Graph ' + esc(data.current.sinceVersion) + ')</div>',
          '<div class="sub" style="margin-top:0.4rem;">Previous longest: ' + esc(data.previous?.name || 'n/a') + '</div>',
        ].join('');

        if (!data.history.length) {
          historyNode.innerHTML = '<p class="sub">No history available.</p>';
          return;
        }

        const rows = data.history.map((item) =>
          '<tr>' +
            '<th scope="row">' + esc(toDate(item.publishedAt)) + '<br><span class="sub">v' + esc(item.version) + '</span></th>' +
            '<td><strong>' + esc(item.name) + '</strong><br><span class="sub">' + esc(item.packageId) + '</span></td>' +
            '<td>' + esc(item.previousName || 'n/a') + '</td>' +
          '</tr>'
        ).join('');

        historyNode.innerHTML =
          '<table>' +
            '<thead>' +
              '<tr>' +
                '<th scope="col">When it changed</th>' +
                '<th scope="col">New longest</th>' +
                '<th scope="col">Previous longest</th>' +
              '</tr>' +
            '</thead>' +
            '<tbody>' + rows + '</tbody>' +
          '</table>';
      } catch (error) {
        statusNode.classList.add('error');
        statusNode.textContent = 'Unable to load data right now.';
        currentNode.innerHTML = '<p class="error">' + esc(error.message) + '</p>';
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
      const separator = entry.indexOf(':');
      const id = separator >= 0 ? entry.slice(0, separator) : entry;
      const versionSpec = separator >= 0 ? entry.slice(separator + 1) : '';
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
    throw new Error(`PowerShell Gallery request failed with status ${response.status}.`);
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

  while (nextUrl) {
    if (guard >= 50) {
      throw new Error('PowerShell Gallery pagination exceeded 50 pages while fetching Microsoft.Graph versions.');
    }
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

  const safePackageId = String(packageId).replaceAll("'", "''");
  const safeVersion = String(version).replaceAll("'", "''");
  const url = `${GALLERY_BASE}/Packages(Id='${safePackageId}',Version='${safeVersion}')`;
  const promise = fetchText(url).then(parsePackageProperties);
  memo.set(key, promise);

  return promise;
}

async function resolveWinnerForVersion(root, packageMemo, metadataLoader = getPackageMetadata) {
  const rootCommands = parseCommandsFromMetadata(root);
  let longestName = pickLongest(rootCommands);
  let sourcePackage = root.id || ROOT_PACKAGE;

  const dependencies = parseDependencyString(root.dependencies)
    .filter((dep) => dep.id.startsWith('Microsoft.Graph'));

  if (dependencies.length > 0) {
    const dependencyMetadata = await Promise.all(
      dependencies.map((dep) => metadataLoader(dep.id, dep.version, packageMemo))
    );

    for (const metadata of dependencyMetadata) {
      const candidate = pickLongest(parseCommandsFromMetadata(metadata));
      if (
        candidate && (
          !longestName ||
          candidate.length > longestName.length ||
          (candidate.length === longestName.length && candidate < longestName)
        )
      ) {
        longestName = candidate;
        sourcePackage = metadata.id || ROOT_PACKAGE;
      }
    }
  }

  return { longestName, packageId: sourcePackage };
}

function buildResponse(winnersByVersion) {
  const transitions = [];
  let last = null;

  for (const item of winnersByVersion) {
    if (!item.longestName) continue;

    if (!last || last.name !== item.longestName || last.packageId !== item.packageId) {
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
    const { longestName, packageId } = await resolveWinnerForVersion(root, packageMemo);

    winners.push({
      version: root.version,
      published: root.published,
      longestName,
      packageId,
    });
  }

  return buildResponse(winners);
}

async function getCachedData() {
  const now = Date.now();

  if (memoryCache.data && memoryCache.expiresAt > now) {
    return memoryCache.data;
  }

  if (memoryCache.inFlight && memoryCache.inFlightGeneration === memoryCache.generation) {
    return memoryCache.inFlight;
  }

  const generation = memoryCache.generation;
  const inFlightPromise = computeLongestCmdletHistory()
    .then((data) => {
      if (generation === memoryCache.generation) {
        memoryCache.data = data;
        memoryCache.expiresAt = Date.now() + CACHE_TTL_MS;
      }
      return data;
    })
    .finally(() => {
      if (memoryCache.inFlight === inFlightPromise) {
        memoryCache.inFlight = null;
        memoryCache.inFlightGeneration = null;
      }
    });

  memoryCache.inFlight = inFlightPromise;
  memoryCache.inFlightGeneration = generation;
  return inFlightPromise;
}

function jsonResponse(payload, status = 200, cacheControl = 'public, max-age=300') {
  return new Response(JSON.stringify(payload, null, 2), {
    status,
    headers: {
      'content-type': 'application/json; charset=utf-8',
      'cache-control': cacheControl,
    },
  });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (url.pathname === '/api/longest') {
      try {
        const data = await getCachedData();
        return jsonResponse(data);
      } catch {
        return jsonResponse(
          {
            error: 'Unable to fetch PowerShell Gallery data.',
          },
          500,
        );
      }
    }

    if (url.pathname === '/api/refresh' && request.method === 'POST') {
      const configuredToken = env?.REFRESH_TOKEN;
      const bearerToken = request.headers.get('authorization')?.replace(/^Bearer\s+/i, '').trim();
      const headerToken = request.headers.get('x-refresh-token')?.trim();
      const requestToken = headerToken || bearerToken;
      if (!configuredToken || !requestToken || requestToken !== configuredToken) {
        return jsonResponse({ error: 'Unauthorized.' }, 401, 'private, no-store');
      }

      memoryCache.data = null;
      memoryCache.expiresAt = 0;
      memoryCache.generation += 1;
      return jsonResponse({ ok: true }, 200, 'private, no-store');
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
  resolveWinnerForVersion,
  clearCache() {
    memoryCache.data = null;
    memoryCache.expiresAt = 0;
    memoryCache.generation += 1;
  },
  setCache(data, expiresAt = Date.now() + CACHE_TTL_MS) {
    memoryCache.data = data;
    memoryCache.expiresAt = expiresAt;
  },
  getCache() {
    return { data: memoryCache.data, expiresAt: memoryCache.expiresAt };
  },
};
