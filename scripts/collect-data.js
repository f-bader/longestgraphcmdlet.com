import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { XMLParser } from 'fast-xml-parser';

const GALLERY_BASE = 'https://www.powershellgallery.com/api/v2';
const ROOT_PACKAGE = 'Microsoft.Graph';
const HISTORY_LIMIT = 10;
const MAX_ROOT_PACKAGE_PAGES = 500;
const MAX_BATCH_PAGES = 500;
const DEPENDENCY_BATCH_SIZE = 25;
const REQUEST_CONCURRENCY = 4;
const RETRY_ATTEMPTS = 3;

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseTagValue: true,
  trimValues: true,
});

export class PowerShellGalleryError extends Error {
  constructor(message, { code = 'GALLERY_ERROR', retryable = false, status = null, cause = null } = {}) {
    super(message, cause ? { cause } : undefined);
    this.name = 'PowerShellGalleryError';
    this.code = code;
    this.retryable = retryable;
    this.status = status;
  }
}

function asArray(value) {
  if (!value) return [];
  return Array.isArray(value) ? value : [value];
}

function pickKey(object, plainName) {
  if (!object || typeof object !== 'object') return null;

  for (const key of Object.keys(object)) {
    if (key === plainName || key.endsWith(`:${plainName}`)) return object[key];
  }

  return null;
}

export function parseCsvCommands(rawValue) {
  if (!rawValue || typeof rawValue !== 'string') return [];

  return rawValue
    .split(/[|,;\s]+/g)
    .map((item) => item.trim())
    .filter(Boolean);
}

export function pickLongest(commands) {
  let best = null;

  for (const command of commands) {
    if (!best || command.length > best.length || (command.length === best.length && command < best)) {
      best = command;
    }
  }

  return best;
}

function parseCommandsFromMetadata(metadata) {
  return [...new Set([...parseCsvCommands(metadata.cmdlets), ...parseCsvCommands(metadata.functions)])];
}

function normalizeVersion(value) {
  const match = String(value ?? '').match(/\d+(?:\.\d+){1,3}(?:[-+][0-9A-Za-z.-]+)?/);
  return match ? match[0] : null;
}

function versionParts(version) {
  const match = String(version).match(/^(\d+(?:\.\d+){0,3})(?:-([0-9A-Za-z.-]+))?/);
  const numeric = (match?.[1] ?? '0').split('.').map(Number);
  return { numeric, prerelease: match?.[2] ?? null };
}

export function compareVersions(leftVersion, rightVersion) {
  const left = versionParts(leftVersion);
  const right = versionParts(rightVersion);
  const length = Math.max(left.numeric.length, right.numeric.length);

  for (let index = 0; index < length; index += 1) {
    const difference = (left.numeric[index] ?? 0) - (right.numeric[index] ?? 0);
    if (difference) return difference;
  }

  if (left.prerelease && !right.prerelease) return -1;
  if (!left.prerelease && right.prerelease) return 1;
  return String(left.prerelease ?? '').localeCompare(String(right.prerelease ?? ''));
}

export function parseVersionRange(rawSpec) {
  const spec = String(rawSpec ?? '').trim().replace(/:$/, '').trim();
  const interval = spec.match(/^(\[|\()\s*([^,]*)\s*,\s*([^\]\)]*)\s*(\]|\))$/);

  if (interval) {
    return {
      lower: normalizeVersion(interval[2]),
      lowerInclusive: interval[1] === '[',
      upper: normalizeVersion(interval[3]),
      upperInclusive: interval[4] === ']',
    };
  }

  const exact = normalizeVersion(spec);
  if (exact) {
    return { lower: exact, lowerInclusive: true, upper: exact, upperInclusive: true };
  }

  return { lower: null, lowerInclusive: false, upper: null, upperInclusive: false };
}

export function parseDependencyString(raw) {
  if (!raw || typeof raw !== 'string') return [];

  return raw
    .split('|')
    .map((entry) => entry.trim())
    .filter(Boolean)
    .map((entry) => {
      const separator = entry.indexOf(':');
      const id = separator >= 0 ? entry.slice(0, separator).trim() : entry;
      const spec = separator >= 0 ? entry.slice(separator + 1) : '';
      return { id, range: parseVersionRange(spec) };
    })
    .filter((dependency) => dependency.id);
}

function packageFromProperties(properties) {
  return {
    id: pickKey(properties, 'Id'),
    version: pickKey(properties, 'Version'),
    published: pickKey(properties, 'Published'),
    dependencies: pickKey(properties, 'Dependencies') ?? '',
    cmdlets: pickKey(properties, 'Cmdlets') ?? '',
    functions: pickKey(properties, 'Functions') ?? '',
  };
}

export function parseFeed(xmlText) {
  let document;
  try {
    document = parser.parse(xmlText);
  } catch (cause) {
    throw new PowerShellGalleryError('PowerShell Gallery returned invalid XML.', {
      code: 'FEED_PARSE_FAILED',
      cause,
    });
  }

  const entries = asArray(document.feed?.entry).map((entry) => {
    const properties = pickKey(entry.content, 'properties') || pickKey(entry, 'properties') || {};
    return packageFromProperties(properties);
  });
  const links = asArray(document.feed?.link);
  const nextLink = links.find((link) => link?.['@_rel'] === 'next')?.['@_href'] ?? null;

  return { entries, nextLink: nextLink?.replaceAll('&amp;', '&') ?? null };
}

function retryDelay(attempt) {
  return 250 * (2 ** attempt) + Math.floor(Math.random() * 100);
}

async function defaultSleep(milliseconds) {
  await new Promise((resolveSleep) => setTimeout(resolveSleep, milliseconds));
}

export async function fetchText(url, {
  fetchImpl = fetch,
  sleep = defaultSleep,
  retries = RETRY_ATTEMPTS,
} = {}) {
  let lastError;

  for (let attempt = 0; attempt <= retries; attempt += 1) {
    try {
      const response = await fetchImpl(url, {
        headers: { accept: 'application/atom+xml,application/xml,text/xml' },
      });

      if (response.ok) return response.text();

      const retryable = response.status === 429 || response.status >= 500;
      throw new PowerShellGalleryError(
        `PowerShell Gallery request failed with status ${response.status}.`,
        { code: 'UPSTREAM_REQUEST_FAILED', retryable, status: response.status },
      );
    } catch (error) {
      lastError = error instanceof PowerShellGalleryError
        ? error
        : new PowerShellGalleryError(`PowerShell Gallery request failed: ${error.message}`, {
          code: 'UPSTREAM_REQUEST_FAILED',
          retryable: true,
          cause: error,
        });

      if (!lastError.retryable || attempt === retries) throw lastError;
      await sleep(retryDelay(attempt));
    }
  }

  throw lastError;
}

export async function fetchAllPages(initialUrl, options = {}, maxPages = MAX_BATCH_PAGES) {
  const entries = [];
  const visitedUrls = new Set();
  let nextUrl = initialUrl;

  while (nextUrl) {
    if (visitedUrls.has(nextUrl)) {
      throw new PowerShellGalleryError(`PowerShell Gallery pagination loop detected for ${initialUrl}.`, {
        code: 'PAGINATION_LOOP',
      });
    }
    if (visitedUrls.size >= maxPages) {
      throw new PowerShellGalleryError(`PowerShell Gallery pagination exceeded ${maxPages} pages for ${initialUrl}.`, {
        code: 'PAGINATION_LIMIT_EXCEEDED',
      });
    }

    visitedUrls.add(nextUrl);
    const page = parseFeed(await fetchText(nextUrl, options));
    entries.push(...page.entries);
    nextUrl = page.nextLink;
  }

  return entries;
}

function rootFeedUrl() {
  const url = new URL(`${GALLERY_BASE}/FindPackagesById()`);
  url.searchParams.set('id', `'${ROOT_PACKAGE}'`);
  return url.toString();
}

function dependencyBatchUrl(packageIds) {
  const url = new URL(`${GALLERY_BASE}/Packages()`);
  const filter = packageIds.map((id) => `Id eq '${id.replaceAll("'", "''")}'`).join(' or ');
  url.searchParams.set('$filter', filter);
  return url.toString();
}

function chunk(values, size) {
  const result = [];
  for (let index = 0; index < values.length; index += size) result.push(values.slice(index, index + size));
  return result;
}

async function mapWithConcurrency(values, limit, mapper) {
  const results = new Array(values.length);
  let nextIndex = 0;

  async function worker() {
    while (nextIndex < values.length) {
      const index = nextIndex;
      nextIndex += 1;
      results[index] = await mapper(values[index], index);
    }
  }

  await Promise.all(Array.from({ length: Math.min(limit, values.length) }, worker));
  return results;
}

function isVersionInRange(version, range) {
  if (range.lower) {
    const lowerComparison = compareVersions(version, range.lower);
    if (lowerComparison < 0 || (lowerComparison === 0 && !range.lowerInclusive)) return false;
  }
  if (range.upper) {
    const upperComparison = compareVersions(version, range.upper);
    if (upperComparison > 0 || (upperComparison === 0 && !range.upperInclusive)) return false;
  }
  return true;
}

function publishedBefore(metadata, rootPublished) {
  const packageTimestamp = Date.parse(metadata.published ?? '');
  const rootTimestamp = Date.parse(rootPublished ?? '');
  return Number.isNaN(packageTimestamp) || Number.isNaN(rootTimestamp) || packageTimestamp <= rootTimestamp;
}

export function selectDependencyMetadata(metadata, range, rootPublished) {
  return metadata
    .filter((item) => isVersionInRange(item.version, range) && publishedBefore(item, rootPublished))
    .sort((left, right) => compareVersions(right.version, left.version))[0] ?? null;
}

function selectWinner(current, candidateName, packageId) {
  if (!candidateName) return current;
  if (!current || candidateName.length > current.name.length || (candidateName.length === current.name.length && candidateName < current.name)) {
    return { name: candidateName, packageId };
  }
  return current;
}

export function buildResponse(winnersByVersion, generatedAt = new Date().toISOString()) {
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

  const current = transitions.at(-1);
  const previous = transitions.at(-2) ?? null;
  return {
    generatedAt,
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

export async function collectSnapshot({ fetchImpl = fetch, sleep = defaultSleep, generatedAt } = {}) {
  const requestOptions = { fetchImpl, sleep };
  const roots = (await fetchAllPages(rootFeedUrl(), requestOptions, MAX_ROOT_PACKAGE_PAGES))
    .filter((item) => item.id === ROOT_PACKAGE && item.version)
    .sort((left, right) => compareVersions(left.version, right.version));

  if (roots.length === 0) {
    throw new PowerShellGalleryError('No Microsoft.Graph package versions were returned from PowerShell Gallery.', {
      code: 'ROOT_VERSIONS_EMPTY',
    });
  }

  const dependencyIds = new Set();
  for (const root of roots) {
    for (const dependency of parseDependencyString(root.dependencies)) {
      if (dependency.id.startsWith('Microsoft.Graph') && dependency.id !== ROOT_PACKAGE) dependencyIds.add(dependency.id);
    }
  }

  const batches = chunk([...dependencyIds].sort(), DEPENDENCY_BATCH_SIZE);
  const dependencyPages = await mapWithConcurrency(
    batches,
    REQUEST_CONCURRENCY,
    (batch) => fetchAllPages(dependencyBatchUrl(batch), requestOptions),
  );

  const metadataByPackage = new Map();
  for (const metadata of dependencyPages.flat()) {
    if (!metadata.id || !metadata.version || !dependencyIds.has(metadata.id)) continue;
    const condensed = {
      id: metadata.id,
      version: metadata.version,
      published: metadata.published,
      longestName: pickLongest(parseCommandsFromMetadata(metadata)),
    };
    const values = metadataByPackage.get(condensed.id) ?? [];
    values.push(condensed);
    metadataByPackage.set(condensed.id, values);
  }

  const winners = roots.map((root) => {
    let winner = selectWinner(null, pickLongest(parseCommandsFromMetadata(root)), root.id);

    for (const dependency of parseDependencyString(root.dependencies)) {
      if (!dependency.id.startsWith('Microsoft.Graph') || dependency.id === ROOT_PACKAGE) continue;
      const selected = selectDependencyMetadata(
        metadataByPackage.get(dependency.id) ?? [],
        dependency.range,
        root.published,
      );
      if (!selected) {
        throw new PowerShellGalleryError(
          `No dependency metadata matches ${dependency.id} for Microsoft.Graph ${root.version}.`,
          { code: 'DEPENDENCY_METADATA_MISSING' },
        );
      }
      winner = selectWinner(winner, selected.longestName, selected.id);
    }

    return {
      version: root.version,
      published: root.published,
      longestName: winner?.name ?? null,
      packageId: winner?.packageId ?? null,
    };
  });

  const snapshot = buildResponse(winners, generatedAt);
  if (!snapshot.current.name || snapshot.history.length === 0) {
    throw new PowerShellGalleryError('Collected metadata did not yield a complete command history.', {
      code: 'COLLECTION_INCOMPLETE',
    });
  }
  return snapshot;
}

export async function writeSnapshot(snapshot, outputPath) {
  const destination = resolve(outputPath);
  const contents = `// Generated by scripts/collect-data.js. Do not edit manually.\nexport default ${JSON.stringify(snapshot, null, 2)};\n`;
  await mkdir(dirname(destination), { recursive: true });
  await writeFile(destination, contents, 'utf8');
}

async function main() {
  const outIndex = process.argv.indexOf('--out');
  const outputPath = outIndex >= 0 ? process.argv[outIndex + 1] : 'src/snapshot.js';
  if (!outputPath) throw new Error('The --out option requires a path.');

  const snapshot = await collectSnapshot();
  await writeSnapshot(snapshot, outputPath);
  console.log(`Generated ${outputPath}: ${snapshot.totalVersionsAnalyzed} versions, current ${snapshot.current.name}.`);
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main().catch((error) => {
    console.error(error.stack || error.message);
    process.exitCode = 1;
  });
}
