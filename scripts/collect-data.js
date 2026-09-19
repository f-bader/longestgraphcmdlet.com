import { mkdir, readFile, writeFile, rename } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { inflateRawSync } from 'node:zlib';
import { XMLParser } from 'fast-xml-parser';

const GALLERY_BASE = 'https://www.powershellgallery.com/api/v2';
const ROOT_PACKAGE = 'Microsoft.Graph';
const HISTORY_LIMIT = 10;
const MAX_ROOT_PACKAGE_PAGES = 500;
const MAX_BATCH_PAGES = 500;
const DEPENDENCY_BATCH_SIZE = 25;
const REQUEST_CONCURRENCY = 4;
const RETRY_ATTEMPTS = 3;
const DEFAULT_RECENT_DAYS = 7;
export const STATE_VERSION = 2;
const COMMAND_CACHE_VERSION = 2;
const DAY_MS = 24 * 60 * 60 * 1000;

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
    .filter((name) => /^[A-Za-z][A-Za-z0-9]*-[A-Za-z0-9_]+$/.test(name));
}

export function pickLongest(commands) {
  let best = null;

  for (const command of commands) {
    if (!best || command.length > best.length || (command.length === best.length && command.localeCompare(best) < 0)) {
      best = command;
    }
  }

  return best;
}

export function pickTopTen(commands, packageId) {
  // Get the 10 longest commands from a list
  const sorted = [...new Set(commands)].sort((a, b) => b.length - a.length || a.localeCompare(b));
  return sorted.slice(0, 10).map(name => ({ name, length: name.length, packageId }));
}

function parseCommandsFromMetadata(metadata) {
  const tagged = String(metadata.tags ?? '').split(/\s+/).filter((tag) => /^PS(?:Cmdlet|Function)_/.test(tag)).map((tag) => tag.replace(/^PS(?:Cmdlet|Function)_/, ''));
  return [...new Set([...parseCsvCommands(metadata.cmdlets), ...parseCsvCommands(metadata.functions), ...parseCsvCommands(tagged.join(' '))])];
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
  const leftParts = String(left.prerelease ?? '').split('.');
  const rightParts = String(right.prerelease ?? '').split('.');
  for (let index = 0; index < Math.max(leftParts.length, rightParts.length); index += 1) {
    const a = leftParts[index];
    const b = rightParts[index];
    if (a === b) continue;
    if (a === undefined) return -1;
    if (b === undefined) return 1;
    const aNumeric = /^\d+$/.test(a);
    const bNumeric = /^\d+$/.test(b);
    if (aNumeric && bNumeric) return Number(a) - Number(b);
    if (aNumeric !== bNumeric) return aNumeric ? -1 : 1;
    return a.toLowerCase() < b.toLowerCase() ? -1 : 1;
  }
  return 0;
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

export function normalizeDate(value) {
  const text = typeof value === 'object' && value ? value['#text'] : value;
  if (typeof text !== 'string' || !/^\d{4}-\d{2}-\d{2}T/.test(text) || text.startsWith('1900-')) return null;
  const timestamp = Date.parse(/(?:Z|[+-]\d{2}:?\d{2})$/i.test(text) ? text : `${text}Z`);
  return Number.isFinite(timestamp) ? new Date(timestamp).toISOString() : null;
}

function propertyText(value) {
  return typeof value === 'object' && value ? value['#text'] ?? '' : value ?? '';
}

function packageFromProperties(properties) {
  return {
    id: propertyText(pickKey(properties, 'Id')),
    version: propertyText(pickKey(properties, 'Version')),
    published: normalizeDate(pickKey(properties, 'Published')) || normalizeDate(pickKey(properties, 'Created')),
    dependencies: propertyText(pickKey(properties, 'Dependencies')),
    cmdlets: propertyText(pickKey(properties, 'Cmdlets')),
    functions: propertyText(pickKey(properties, 'Functions')),
    tags: propertyText(pickKey(properties, 'Tags')),
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

  if (!document.feed) throw new PowerShellGalleryError('PowerShell Gallery returned no Atom feed.', { code: 'FEED_PARSE_FAILED' });
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

async function fetchResponse(url, {
  fetchImpl = fetch,
  sleep = defaultSleep,
  retries = RETRY_ATTEMPTS,
  timeoutMs = 60_000,
  consume = (response) => response,
} = {}) {
  let lastError;

  for (let attempt = 0; attempt <= retries; attempt += 1) {
    try {
      const response = await fetchImpl(url, {
        headers: { accept: 'application/atom+xml,application/xml,text/xml,application/octet-stream' },
        signal: AbortSignal.timeout(timeoutMs),
      });

      if (response.ok) return await consume(response);
      const retryAfter = response.headers.get('retry-after');
      await response.body?.cancel();

      const retryable = response.status === 429 || response.status >= 500;
      const error = new PowerShellGalleryError(
        `PowerShell Gallery request failed with status ${response.status}.`,
        { code: 'UPSTREAM_REQUEST_FAILED', retryable, status: response.status },
      );
      if (retryAfter) error.retryAfterMs = Math.min(60_000, Math.max(0, /^\d+$/.test(retryAfter) ? Number(retryAfter) * 1000 : Date.parse(retryAfter) - Date.now()));
      throw error;
    } catch (error) {
      lastError = error instanceof PowerShellGalleryError
        ? error
        : new PowerShellGalleryError(`PowerShell Gallery request failed: ${error.message}`, {
          code: 'UPSTREAM_REQUEST_FAILED',
          retryable: true,
          cause: error,
        });

      if (!lastError.retryable || attempt === retries) throw lastError;
      await sleep(Number.isFinite(lastError.retryAfterMs) ? lastError.retryAfterMs : retryDelay(attempt));
    }
  }

  throw lastError;
}

export async function fetchText(url, options = {}) {
  return fetchResponse(url, { ...options, consume: (response) => response.text() });
}

async function fetchBytes(url, options = {}) {
  return fetchResponse(url, { ...options, consume: async (response) => new Uint8Array(await response.arrayBuffer()) });
}

function readUint16(view, offset) {
  return view.getUint16(offset, true);
}

function readUint32(view, offset) {
  return view.getUint32(offset, true);
}

export function parseZipEntries(bytes, include = () => true) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let endOfCentralDirectory = -1;
  for (let offset = bytes.length - 22; offset >= 0; offset -= 1) {
    if (readUint32(view, offset) === 0x06054b50) {
      endOfCentralDirectory = offset;
      break;
    }
  }
  if (endOfCentralDirectory < 0) throw new PowerShellGalleryError('Package archive has no ZIP directory.', { code: 'PACKAGE_ARCHIVE_INVALID' });

  const entryCount = readUint16(view, endOfCentralDirectory + 10);
  let directoryOffset = readUint32(view, endOfCentralDirectory + 16);
  const entries = new Map();

  for (let index = 0; index < entryCount; index += 1) {
    if (readUint32(view, directoryOffset) !== 0x02014b50) {
      throw new PowerShellGalleryError('Package archive has an invalid ZIP directory entry.', { code: 'PACKAGE_ARCHIVE_INVALID' });
    }
    const method = readUint16(view, directoryOffset + 10);
    const compressedSize = readUint32(view, directoryOffset + 20);
    const nameLength = readUint16(view, directoryOffset + 28);
    const extraLength = readUint16(view, directoryOffset + 30);
    const commentLength = readUint16(view, directoryOffset + 32);
    const localHeaderOffset = readUint32(view, directoryOffset + 42);
    const name = new TextDecoder().decode(bytes.slice(directoryOffset + 46, directoryOffset + 46 + nameLength));
    directoryOffset += 46 + nameLength + extraLength + commentLength;
    if (!include(name)) continue;
    const localNameLength = readUint16(view, localHeaderOffset + 26);
    const localExtraLength = readUint16(view, localHeaderOffset + 28);
    const start = localHeaderOffset + 30 + localNameLength + localExtraLength;
    const compressed = bytes.slice(start, start + compressedSize);
    let contents;
    if (method === 0) contents = compressed;
    else if (method === 8) contents = new Uint8Array(inflateRawSync(compressed, { maxOutputLength: 32 * 1024 * 1024 }));
    else throw new PowerShellGalleryError(`Unsupported ZIP compression method ${method}.`, { code: 'PACKAGE_ARCHIVE_INVALID' });
    entries.set(name, contents);
  }

  return entries;
}

export function parseManifestCommands(manifestText) {
  // Tokenize literal data only. Never execute downloaded PowerShell code.
  const tokens = (manifestText.match(/<#[\s\S]*?#>|#[^\r\n]*|'(?:''|[^'])*'|"(?:`[\s\S]|[^"`])*"|[A-Za-z_]\w*|@\(|[^\s]/g) ?? [])
    .filter((token) => !token.startsWith('#') && !token.startsWith('<#'));
  const commands = [];
  for (let index = 0; index < tokens.length; index += 1) {
    if (!/^(CmdletsToExport|FunctionsToExport)$/i.test(tokens[index]) || tokens[index + 1] !== '=') continue;
    let cursor = index + 2;
    const array = tokens[cursor] === '@(' || tokens[cursor] === '(';
    if (array) cursor += 1;
    while (cursor < tokens.length) {
      const token = tokens[cursor];
      if (token === ',' || token === ';') { cursor += 1; continue; }
      if (!/^['"]/.test(token)) break;
      const value = token.slice(1, -1).replaceAll("''", "'");
      if (/^[A-Za-z][A-Za-z0-9]*-[A-Za-z0-9_]+$/.test(value)) commands.push(value);
      cursor += 1;
      if (!array && tokens[cursor] !== ',') break;
    }
  }
  return [...new Set(commands)];
}

export function parsePackageArchiveCommands(bytes) {
  const manifests = parseZipEntries(bytes, (name) => name.toLowerCase().endsWith('.psd1'));
  const commands = [];
  let wildcardExports = false;
  for (const content of manifests.values()) {
    const manifest = new TextDecoder().decode(content);
    commands.push(...parseManifestCommands(manifest));
    const withoutComments = manifest.replace(/<#[\s\S]*?#>|#[^\r\n]*/g, '');
    wildcardExports ||= /\bFunctionsToExport\s*=\s*(?:@?\(\s*)?['"]\*['"]/i.test(withoutComments);
  }
  // Early generated Graph modules export proxy functions from this directory
  // instead of listing names in the manifest. Ignore private/custom scripts.
  if (wildcardExports) {
    const proxies = parseZipEntries(bytes, (name) => /(?:^|\/)exports\/.*\.ps1$/i.test(name));
    for (const content of proxies.values()) {
      const script = new TextDecoder().decode(content).replace(/<#[\s\S]*?#>|#[^\r\n]*|'(?:''|[^'])*'|"(?:`[\s\S]|[^"`])*"/g, ' ');
      for (const match of script.matchAll(/^\s*function\s+([A-Za-z][A-Za-z0-9]*-[A-Za-z0-9_]+)\s*[{(]/gmi)) commands.push(match[1]);
    }
  }
  return [...new Set(commands)];
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

function rootFeedUrl(since = null) {
  if (since) {
    const recentUrl = new URL(`${GALLERY_BASE}/Packages()`);
    recentUrl.searchParams.set('$filter', `Id eq '${ROOT_PACKAGE}' and (Published ge datetime'${since.toISOString()}' or Created ge datetime'${since.toISOString()}')`);
    return recentUrl.toString();
  }

  const url = new URL(`${GALLERY_BASE}/FindPackagesById()`);
  url.searchParams.set('id', `'${ROOT_PACKAGE}'`);
  return url.toString();
}

function dependencyBatchUrl(packageIds, requirements = new Map()) {
  const url = new URL(`${GALLERY_BASE}/Packages()`);
  const filter = packageIds.map((id) => {
    const ranges = requirements.get(id) ?? [];
    const exact = ranges.length && ranges.every((range) => range.lower && range.lower === range.upper && range.lowerInclusive && range.upperInclusive);
    const versions = [...new Set(ranges.map((range) => range.lower))];
    const versionFilter = exact && versions.length <= 8 ? ` and (${versions.map((version) => `Version eq '${version}'`).join(' or ')})` : '';
    return `(Id eq '${id.replaceAll("'", "''")}'${versionFilter})`;
  }).join(' or ');
  url.searchParams.set('$filter', filter);
  return url.toString();
}

function packageArchiveUrl(packageId, version) {
  return `${GALLERY_BASE}/package/${encodeURIComponent(packageId)}/${encodeURIComponent(version)}`;
}

function dependencyBatches(packageIds, requirements) {
  const batches = [];
  let batch = [];
  let cost = 0;
  for (const id of packageIds) {
    const ranges = requirements.get(id) ?? [];
    const exact = ranges.length && ranges.every((range) => range.lower && range.lower === range.upper && range.lowerInclusive && range.upperInclusive);
    const versions = new Set(ranges.map((range) => range.lower)).size;
    // Stay below Gallery's 250-node OData filter limit, including version clauses.
    const itemCost = 8 + (exact && versions <= 8 ? versions * 8 : 0);
    if (batch.length && (batch.length >= DEPENDENCY_BATCH_SIZE || cost + itemCost > 200)) {
      batches.push(batch);
      batch = [];
      cost = 0;
    }
    batch.push(id);
    cost += itemCost;
  }
  if (batch.length) batches.push(batch);
  return batches;
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
  const packageDate = normalizeDate(metadata.published);
  const rootDate = normalizeDate(rootPublished);
  return Boolean(packageDate && rootDate && packageDate <= rootDate);
}

export function selectDependencyMetadata(metadata, range, rootPublished, rootVersion = null) {
  let best = null;
  for (const item of metadata) {
    if (isVersionInRange(item.version, range) && publishedBefore(item, rootPublished)
      && (!best || compareVersions(item.version, best.version) > 0)) best = item;
  }
  // Gallery publishes a release's modules sequentially. A dependency explicitly
  // named by this root release can appear just after the umbrella package.
  // Never apply that exception to an arbitrary newer version in an open range.
  if (!best && range.lower === rootVersion && range.lowerInclusive) {
    best = metadata.find((item) => item.version === rootVersion && isVersionInRange(item.version, range) && normalizeDate(item.published)) ?? null;
  }
  return best;
}

function selectWinner(current, candidateName, packageId) {
  if (!candidateName) return current;
  if (!current || candidateName.length > current.name.length || (candidateName.length === current.name.length && candidateName.localeCompare(current.name) < 0)) {
    return { name: candidateName, packageId };
  }
  return current;
}

export function buildResponse(winnersByVersion, generatedAt = new Date().toISOString()) {
  const transitions = [];
  let last = null;

  for (const item of winnersByVersion) {
    if (!item.longestName) continue;
    // A shorter winner in a newer release cannot displace the all-time record.
    if (!last || item.longestName.length > last.name.length
      || (item.longestName.length === last.name.length && item.longestName.localeCompare(last.name) < 0)) {
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
  const historyReversed = transitions.slice(-HISTORY_LIMIT).reverse();

  // Build topLongest by collecting top 10 from all versions and deduplicating
  const allTopTenCandidates = [];
  for (const item of winnersByVersion) {
    if (item.topTenLongest) {
      // Each item already has versionFound, just add the packageId if missing
      allTopTenCandidates.push(...item.topTenLongest.map(cmd => ({
        ...cmd,
        packageId: cmd.packageId || item.packageId,
      })));
    }
  }

  // Keep the first release in which each command appears, with deterministic ties.
  const topTenUnique = [];
  const seen = new Set();
  for (const item of allTopTenCandidates.sort((a, b) => b.length - a.length || a.name.localeCompare(b.name))) {
    if (!item.name || seen.has(item.name)) continue;
    topTenUnique.push(item);
    seen.add(item.name);
    if (topTenUnique.length >= HISTORY_LIMIT) break;
  }

  return {
    generatedAt,
    source: 'PowerShell Gallery: Microsoft.Graph + dependencies',
    totalVersionsAnalyzed: winnersByVersion.length,
    analyzedVersions: winnersByVersion.map((item) => item.version),
    current: {
      name: current?.name ?? null,
      length: current?.name?.length ?? 0,
      packageId: current?.packageId ?? null,
      sinceVersion: current?.version ?? null,
      sincePublished: current?.publishedAt ?? null,
    },
    previous,
    history: historyReversed,
    topLongest: topTenUnique,
  };
}

export function mergeIncrementalSnapshots(previous, recent) {
  const winners = new Map((previous.winners ?? []).map((winner) => [winner.version, winner]));
  for (const winner of recent.winners) winners.set(winner.version, winner);
  const ordered = [...winners.values()].sort((a, b) => compareVersions(a.version, b.version));
  return { ...buildResponse(ordered, recent.generatedAt), stateVersion: STATE_VERSION, winners: ordered };
}

export async function collectSnapshot({
  fetchImpl = fetch,
  sleep = defaultSleep,
  generatedAt,
  previousSnapshot = null,
  recentDays = null,
  commandCache = new Map(),
  saveCommandSummary = async () => {},
  onProgress = () => {},
} = {}) {
  const requestOptions = { fetchImpl, sleep };
  const effectiveGeneratedAt = generatedAt ?? new Date().toISOString();
  // Old snapshots contain lossy transitions and may have been produced by the date/parser bugs.
  if (previousSnapshot?.stateVersion !== STATE_VERSION || !Array.isArray(previousSnapshot?.winners)) previousSnapshot = null;
  const lastSuccessful = normalizeDate(previousSnapshot?.generatedAt);
  const since = lastSuccessful && recentDays > 0
    ? new Date(Math.min(Date.parse(lastSuccessful), Date.parse(effectiveGeneratedAt)) - recentDays * DAY_MS)
    : null;
  const roots = (await fetchAllPages(rootFeedUrl(since), requestOptions, MAX_ROOT_PACKAGE_PAGES))
    .filter((item) => item.id === ROOT_PACKAGE && item.version)
    .sort((left, right) => compareVersions(left.version, right.version));
  onProgress(`Found ${roots.length} root releases${since ? ' in the catch-up window' : ' for full collection'}.`);
  if (!roots.length) {
    if (previousSnapshot) return mergeIncrementalSnapshots(previousSnapshot, { winners: [], generatedAt: effectiveGeneratedAt });
    throw new PowerShellGalleryError('No Microsoft.Graph package versions were returned.', { code: 'ROOT_VERSIONS_EMPTY' });
  }

  const requirements = new Map();
  const dependenciesByRoot = new Map();
  for (const root of roots) {
    if (!root.published) throw new PowerShellGalleryError(`No usable publication date for Microsoft.Graph ${root.version}.`, { code: 'PUBLICATION_DATE_MISSING' });
    const dependencies = parseDependencyString(root.dependencies).filter((item) => item.id.startsWith('Microsoft.Graph.') );
    dependenciesByRoot.set(root.version, dependencies);
    for (const dependency of dependencies) {
      const ranges = requirements.get(dependency.id) ?? [];
      ranges.push(dependency.range);
      requirements.set(dependency.id, ranges);
    }
  }
  const metadataByPackage = new Map();
  const batches = dependencyBatches([...requirements.keys()].sort(), requirements);
  await mapWithConcurrency(batches, REQUEST_CONCURRENCY, async (batch) => {
    const entries = await fetchAllPages(dependencyBatchUrl(batch, requirements), requestOptions);
    for (const metadata of entries) {
      if (!metadata.id || !metadata.version || !requirements.has(metadata.id)) continue;
      const commands = parseCommandsFromMetadata(metadata);
      const summary = { id: metadata.id, version: metadata.version, published: metadata.published,
        longestName: pickLongest(commands), topTenLongest: pickTopTen(commands, metadata.id) };
      const values = metadataByPackage.get(metadata.id) ?? [];
      values.push(summary);
      metadataByPackage.set(metadata.id, values);
    }
    onProgress(`Read metadata for ${batch.length} dependency packages.`);
  });

  const selectedDependencies = new Map();
  const selectionsByRoot = new Map();
  for (const root of roots) {
    const selections = dependenciesByRoot.get(root.version).map((dependency) => {
      const selected = selectDependencyMetadata(metadataByPackage.get(dependency.id) ?? [], dependency.range, root.published, root.version);
      if (!selected) throw new PowerShellGalleryError(`No historical dependency matches ${dependency.id} for Microsoft.Graph ${root.version}.`, { code: 'DEPENDENCY_METADATA_MISSING' });
      const key = `${selected.id}@${selected.version}`;
      selectedDependencies.set(key, selected);
      return key;
    });
    selectionsByRoot.set(root.version, selections);
  }
  onProgress(`Resolving commands for ${selectedDependencies.size} distinct dependency versions.`);
  await mapWithConcurrency([...selectedDependencies], REQUEST_CONCURRENCY, async ([key, metadata]) => {
    const cached = commandCache.get(key);
    if (!metadata.longestName && cached?.cacheVersion === COMMAND_CACHE_VERSION) {
      metadata.longestName = cached.longestName;
      metadata.topTenLongest = cached.topTenLongest;
    }
    if (!metadata.longestName && cached?.cacheVersion !== COMMAND_CACHE_VERSION) {
      const commands = parsePackageArchiveCommands(await fetchBytes(packageArchiveUrl(metadata.id, metadata.version), requestOptions));
      metadata.longestName = pickLongest(commands);
      metadata.topTenLongest = pickTopTen(commands, metadata.id);
      onProgress(`Read exports from ${key}.`);
    }
    const summary = { cacheVersion: COMMAND_CACHE_VERSION, longestName: metadata.longestName, topTenLongest: metadata.topTenLongest };
    commandCache.set(key, summary);
    await saveCommandSummary(key, summary);
  });

  const winners = roots.map((root) => {
    const rootCommands = parseCommandsFromMetadata(root);
    let winner = selectWinner(null, pickLongest(rootCommands), root.id);
    const candidates = pickTopTen(rootCommands, root.id);
    for (const key of selectionsByRoot.get(root.version)) {
      const selected = selectedDependencies.get(key);
      winner = selectWinner(winner, selected.longestName, selected.id);
      candidates.push(...selected.topTenLongest);
    }
    if (!winner) throw new PowerShellGalleryError(`No winner for Microsoft.Graph ${root.version}.`, { code: 'COLLECTION_INCOMPLETE' });
    const unique = new Map();
    for (const item of candidates.sort((a, b) => b.length - a.length || a.name.localeCompare(b.name))) {
      if (!unique.has(item.name)) unique.set(item.name, { ...item, versionFound: root.version });
      if (unique.size === HISTORY_LIMIT) break;
    }
    return { version: root.version, published: root.published, longestName: winner.name, packageId: winner.packageId, topTenLongest: [...unique.values()] };
  });
  const recent = { ...buildResponse(winners, effectiveGeneratedAt), stateVersion: STATE_VERSION, winners };
  return previousSnapshot ? mergeIncrementalSnapshots(previousSnapshot, recent) : recent;
}

export function publicSnapshot(snapshot) {
  const { stateVersion, winners, ...data } = snapshot;
  return data;
}

export async function writeSnapshot(snapshot, outputPath) {
  const destination = resolve(outputPath);
  const contents = `// Generated by scripts/collect-data.js. Do not edit manually.\nexport default ${JSON.stringify(publicSnapshot(snapshot), null, 2)};\n`;
  await mkdir(dirname(destination), { recursive: true });
  await writeFile(`${destination}.tmp`, contents, 'utf8');
  await rename(`${destination}.tmp`, destination);
}

export async function readSnapshotState(statePath) {
  try {
    return JSON.parse(await readFile(resolve(statePath), 'utf8'));
  } catch (error) {
    if (error.code === 'ENOENT') return null;
    throw error;
  }
}

async function main() {
  const outIndex = process.argv.indexOf('--out');
  const outputPath = outIndex >= 0 ? process.argv[outIndex + 1] : 'src/snapshot.js';
  if (!outputPath) throw new Error('The --out option requires a path.');

  const statePath = process.env.SNAPSHOT_STATE ?? '.cache/graph-snapshot.json';
  const configuredDays = Number(process.env.COLLECTION_DAYS ?? DEFAULT_RECENT_DAYS);
  const forceFull = process.env.FULL_COLLECTION === 'true';
  const previousSnapshot = forceFull ? null : await readSnapshotState(statePath);
  const cacheDirectory = resolve('.cache/command-summaries');
  await mkdir(cacheDirectory, { recursive: true });
  const commandCache = new Map();
  const { readdir } = await import('node:fs/promises');
  for (const filename of await readdir(cacheDirectory)) {
    if (!filename.endsWith('.json')) continue;
    const value = JSON.parse(await readFile(resolve(cacheDirectory, filename), 'utf8'));
    commandCache.set(decodeURIComponent(filename.slice(0, -5)), value);
  }
  const snapshot = await collectSnapshot({
    commandCache,
    onProgress: console.log,
    saveCommandSummary: async (key, summary) => {
      const destination = resolve(cacheDirectory, `${encodeURIComponent(key)}.json`);
      await writeFile(`${destination}.tmp`, JSON.stringify(summary));
      await rename(`${destination}.tmp`, destination);
    },
    previousSnapshot,
    recentDays: Number.isFinite(configuredDays) ? configuredDays : DEFAULT_RECENT_DAYS,
  });
  await writeSnapshot(snapshot, outputPath);
  await mkdir(dirname(resolve(statePath)), { recursive: true });
  await writeFile(`${resolve(statePath)}.tmp`, `${JSON.stringify(snapshot, null, 2)}\n`, 'utf8');
  await rename(`${resolve(statePath)}.tmp`, resolve(statePath));
  console.log(`Generated ${outputPath}: ${snapshot.totalVersionsAnalyzed} versions, current ${snapshot.current.name}.`);
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main().catch((error) => {
    console.error(error.stack || error.message);
    process.exitCode = 1;
  });
}
