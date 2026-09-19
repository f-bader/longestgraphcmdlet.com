import { describe, expect, it, vi } from 'vitest';
import {
  collectSnapshot,
  buildResponse,
  mergeIncrementalSnapshots,
  normalizeDate,
  parseFeed,
  publicSnapshot,
  STATE_VERSION,
  compareVersions,
  fetchText,
  parseDependencyString,
  parseManifestCommands,
  parsePackageArchiveCommands,
  parseZipEntries,
  parseVersionRange,
  selectDependencyMetadata,
} from '../scripts/collect-data.js';

function feed(entries, nextLink = null) {
  const links = nextLink ? `<link rel="next" href="${nextLink}" />` : '';
  return `<?xml version="1.0"?><feed>${links}${entries.join('')}</feed>`;
}

function entry({ id, version, published, dependencies = '', cmdlets = '', functions = '' }) {
  return `<entry><content><properties><Id>${id}</Id><Version>${version}</Version><Published>${published}</Published><Dependencies>${dependencies}</Dependencies><Cmdlets>${cmdlets}</Cmdlets><Functions>${functions}</Functions></properties></content></entry>`;
}

function storedZip(filename, content) {
  const files = Array.isArray(filename) ? filename : [[filename, content]];
  const locals = [];
  const directories = [];
  let offset = 0;
  for (const [entryName, entryContent] of files) {
    const name = new TextEncoder().encode(entryName);
    const bytes = new TextEncoder().encode(entryContent);
    const local = new Uint8Array(30 + name.length + bytes.length);
    const localView = new DataView(local.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint32(18, bytes.length, true);
    localView.setUint32(22, bytes.length, true);
    localView.setUint16(26, name.length, true);
    local.set(name, 30);
    local.set(bytes, 30 + name.length);
    const central = new Uint8Array(46 + name.length);
    const centralView = new DataView(central.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint32(20, bytes.length, true);
    centralView.setUint32(24, bytes.length, true);
    centralView.setUint16(28, name.length, true);
    centralView.setUint32(42, offset, true);
    central.set(name, 46);
    locals.push(local);
    directories.push(central);
    offset += local.length;
  }
  const end = new Uint8Array(22);
  const endView = new DataView(end.buffer);
  endView.setUint32(0, 0x06054b50, true);
  endView.setUint16(8, files.length, true);
  endView.setUint16(10, files.length, true);
  endView.setUint32(12, directories.reduce((total, item) => total + item.length, 0), true);
  endView.setUint32(16, offset, true);
  return new Uint8Array(Buffer.concat([...locals, ...directories, end]));
}

describe('collector helpers', () => {
  it('parses exact, bounded, and unbounded dependency ranges', () => {
    expect(parseDependencyString('Microsoft.Graph.Users:[2.1.0, 2.1.0]:|Other:1.0.0')).toEqual([
      {
        id: 'Microsoft.Graph.Users',
        range: { lower: '2.1.0', lowerInclusive: true, upper: '2.1.0', upperInclusive: true },
      },
      {
        id: 'Other',
        range: { lower: '1.0.0', lowerInclusive: true, upper: '1.0.0', upperInclusive: true },
      },
    ]);
    expect(parseVersionRange('(, )')).toEqual({
      lower: null,
      lowerInclusive: false,
      upper: null,
      upperInclusive: false,
    });
  });

  it('uses semantic ordering and the root publication date for dependency resolution', () => {
    const metadata = [
      { version: '1.0.0', published: '2024-01-01T00:00:00Z' },
      { version: '1.10.0', published: '2024-02-01T00:00:00Z' },
      { version: '2.0.0', published: '2024-03-01T00:00:00Z' },
    ];

    expect(compareVersions('1.10.0', '1.9.0')).toBeGreaterThan(0);
    expect(selectDependencyMetadata(metadata, parseVersionRange('(, )'), '2024-02-15T00:00:00Z')).toMatchObject({ version: '1.10.0' });
    expect(selectDependencyMetadata(metadata, parseVersionRange('[1.0.0, 1.10.0]'), '2024-04-01T00:00:00Z')).toMatchObject({ version: '1.10.0' });
  });

  it('retries transient Gallery failures', async () => {
    const fetchImpl = vi.fn()
      .mockResolvedValueOnce(new Response('busy', { status: 503 }))
      .mockResolvedValueOnce(new Response('ok', { status: 200 }));

    await expect(fetchText('https://example.test', { fetchImpl, sleep: async () => {} })).resolves.toBe('ok');
    expect(fetchImpl).toHaveBeenCalledTimes(2);
  });

  it('reads command exports from a PowerShell module manifest in a package archive', () => {
    const manifest = `@{
      CmdletsToExport = @(
        'Get-MgUser',
        'Set-MgUser'
      )
      FunctionsToExport = @('Connect-MgGraph')
    }`;
    expect(parseManifestCommands(manifest)).toEqual(['Get-MgUser', 'Set-MgUser', 'Connect-MgGraph']);
    expect(parsePackageArchiveCommands(storedZip('Microsoft.Graph.Users.psd1', manifest))).toEqual([
      'Get-MgUser',
      'Set-MgUser',
      'Connect-MgGraph',
    ]);
    expect(parseZipEntries(storedZip('module.psd1', manifest)).has('module.psd1')).toBe(true);
  });
});

describe('collector integration', () => {
  it('follows root pagination, resolves dependencies, and builds a complete history', async () => {
    const initialRootUrl = 'https://www.powershellgallery.com/api/v2/FindPackagesById()?id=%27Microsoft.Graph%27';
    const nextRootUrl = 'https://example.test/root-page-2';
    const rootFirstPage = feed([
      entry({ id: 'Microsoft.Graph', version: '1.0.0', published: '2024-01-15T00:00:00Z', dependencies: 'Microsoft.Graph.Users:(, ):' }),
    ], nextRootUrl);
    const rootSecondPage = feed([
      entry({ id: 'Microsoft.Graph', version: '2.0.0', published: '2024-03-15T00:00:00Z', dependencies: 'Microsoft.Graph.Users:[2.0.0, 2.0.0]:' }),
    ]);
    const dependencyFeed = feed([
      entry({ id: 'Microsoft.Graph.Users', version: '1.0.0', published: '2024-01-01T00:00:00Z' }),
      entry({ id: 'Microsoft.Graph.Users', version: '2.0.0', published: '2024-03-01T00:00:00Z' }),
    ]);
    const directMetadata = {
      '1.0.0': storedZip('Microsoft.Graph.Users.psd1', "@{ CmdletsToExport = @('Get-MgFormerLongestCommand') }"),
      '2.0.0': storedZip('Microsoft.Graph.Users.psd1', "@{ CmdletsToExport = @('Get-MgCurrentLongestCommandName') }"),
    };
    const fetchImpl = vi.fn(async (url) => {
      if (url === initialRootUrl) return new Response(rootFirstPage);
      if (url === nextRootUrl) return new Response(rootSecondPage);
      if (url.includes('/Packages()')) return new Response(dependencyFeed);
      if (url.endsWith('/package/Microsoft.Graph.Users/1.0.0')) return new Response(directMetadata['1.0.0']);
      if (url.endsWith('/package/Microsoft.Graph.Users/2.0.0')) return new Response(directMetadata['2.0.0']);
      return new Response('missing', { status: 404 });
    });

    const snapshot = await collectSnapshot({
      fetchImpl,
      sleep: async () => {},
      generatedAt: '2026-01-01T00:00:00.000Z',
    });

    expect(snapshot.totalVersionsAnalyzed).toBe(2);
    expect(snapshot.current).toMatchObject({
      name: 'Get-MgCurrentLongestCommandName',
      packageId: 'Microsoft.Graph.Users',
      sinceVersion: '2.0.0',
    });
    expect(snapshot.history).toHaveLength(2);
    expect(fetchImpl).toHaveBeenCalledTimes(5);
  });

  it('uses a recent window after a previous snapshot and merges the result', async () => {
    const rootUrlPattern = 'https://www.powershellgallery.com/api/v2/Packages()?%24filter=Id+eq+%27Microsoft.Graph%27+and+%28Published+ge+datetime%27';
    const rootFeed = feed([
      entry({
        id: 'Microsoft.Graph',
        version: '2.0.0',
        published: '2024-03-15T00:00:00Z',
        dependencies: 'Microsoft.Graph.Users:[2.0.0, 2.0.0]:',
      }),
    ]);
    const dependencyFeed = feed([
      entry({ id: 'Microsoft.Graph.Users', version: '2.0.0', published: '2024-03-01T00:00:00Z' }),
    ]);
    const archive = storedZip(
      'Microsoft.Graph.Users.psd1',
      "@{ CmdletsToExport = @('Get-MgCurrentLongestCommandName') }",
    );
    const fetchImpl = vi.fn(async (url) => {
      if (url.startsWith(rootUrlPattern)) return new Response(rootFeed);
      if (url.includes('/Packages()')) return new Response(dependencyFeed);
      if (url.endsWith('/package/Microsoft.Graph.Users/2.0.0')) return new Response(archive);
      return new Response('missing', { status: 404 });
    });

    const snapshot = await collectSnapshot({
      fetchImpl,
      sleep: async () => {},
      generatedAt: '2024-03-20T00:00:00.000Z',
      recentDays: 7,
      previousSnapshot: {
        stateVersion: STATE_VERSION,
        winners: [{ version: '1.0.0', published: '2024-01-15T00:00:00Z', longestName: 'Get-MgFormerLongestCommand', packageId: 'Microsoft.Graph.Users', topTenLongest: [] }],
        generatedAt: '2024-03-10T00:00:00.000Z',
        source: 'PowerShell Gallery: Microsoft.Graph + dependencies',
        totalVersionsAnalyzed: 1,
        analyzedVersions: ['1.0.0'],
        current: {
          name: 'Get-MgFormerLongestCommand',
          length: 26,
          packageId: 'Microsoft.Graph.Users',
          sinceVersion: '1.0.0',
          sincePublished: '2024-01-15T00:00:00Z',
        },
        previous: null,
        history: [{
          version: '1.0.0',
          publishedAt: '2024-01-15T00:00:00Z',
          name: 'Get-MgFormerLongestCommand',
          packageId: 'Microsoft.Graph.Users',
          previousName: null,
        }],
      },
    });

    expect(fetchImpl.mock.calls[0][0]).toContain('Published+ge+datetime%272024-03-03T00%3A00%3A00.000Z%27');
    expect(snapshot.totalVersionsAnalyzed).toBe(2);
    expect(snapshot.current.name).toBe('Get-MgCurrentLongestCommandName');
    expect(snapshot.history.map((item) => item.version)).toEqual(['2.0.0', '1.0.0']);
  });
});

function winner(version, name, published = '2024-03-01T00:00:00Z') {
  return { version, published, longestName: name, packageId: 'Microsoft.Graph.Users', topTenLongest: [{ name, length: name.length, packageId: 'Microsoft.Graph.Users', versionFound: version }] };
}
function state(winners, generatedAt = '2024-03-10T00:00:00.000Z') {
  return { ...buildResponse(winners, generatedAt), stateVersion: STATE_VERSION, winners };
}

describe('collector regressions', () => {
  const record = 'Invoke-MgExtendDeviceManagementDeviceConfigurationGroupAssignmentDeviceConfigurationMicrosoftGraphWindowUpdateForBusinessConfigurationFeatureUpdatePause';
  const laterWinner = 'Remove-MgIdentityAuthenticationEventFlowAsOnGraphAPretributeCollectionExternalUserSelfServiceSignUpAttributeIdentityUserFlowAttributeByRef';

  it('keeps the all-time record when later releases have shorter winners', () => {
    const result = buildResponse([
      winner('1.0.0', 'Get-MgUser'),
      winner('1.4.0', record),
      winner('2.22.0', laterWinner),
    ]);
    expect(result.current).toMatchObject({ name: record, length: 152, sinceVersion: '1.4.0' });
    expect(result.current.name).toBe(result.topLongest[0].name);
    expect(result.history.map(item => item.version)).toEqual(['1.4.0', '1.0.0']);
    expect(result.previous.name).toBe('Get-MgUser');
  });

  it.each([false, true])('rebuilds cached records on refresh, empty window: %s', async (empty) => {
    const previous = state([winner('1.4.0', record), winner('2.22.0', laterWinner)]);
    // Simulate state persisted by the old latest-release selection logic.
    previous.current = { name: laterWinner, length: laterWinner.length };
    previous.history = [];
    const result = await collectSnapshot({
      previousSnapshot: previous, recentDays: 7,
      fetchImpl: async () => new Response(empty ? '<feed><title>Packages</title></feed>' : feed([entry({
        id: 'Microsoft.Graph', version: '2.40.0', published: '2024-03-12T00:00:00Z', functions: laterWinner,
      })])),
    });
    expect(result.current).toMatchObject({ name: record, sinceVersion: '1.4.0' });
    expect(result.history[0].name).toBe(record);
    expect(result.topLongest[0].name).toBe(record);
  });

  it('resolves equal-length records consistently with the ranking and preserves first appearance', () => {
    const result = buildResponse([winner('1.0.0', 'Get-Zzz'), winner('2.0.0', 'Get-Aaa'), winner('3.0.0', 'Get-Aaa')]);
    expect(result.current.name).toBe(result.topLongest[0].name);
    expect(result.current.sinceVersion).toBe('2.0.0');
  });

  it('normalizes typed XML dates and uses Created for unlisted packages', () => {
    const data = parseFeed(feed([
      '<entry><m:properties><d:Id>Microsoft.Graph.Users</d:Id><d:Version>1.0.0</d:Version><d:Published m:type="Edm.DateTime">1900-01-01T00:00:00</d:Published><d:Created m:type="Edm.DateTime">2024-01-01T03:00:00</d:Created></m:properties></entry>',
    ])).entries[0];
    expect(data.published).toBe('2024-01-01T03:00:00.000Z');
    expect(normalizeDate({ '#text': '2024-01-01T00:00:00', '@_m:type': 'Edm.DateTime' })).toBe('2024-01-01T00:00:00.000Z');
    expect(normalizeDate('1900-01-01T00:00:00')).toBeNull();
  });

  it('rejects future or undated dependencies even when XML dates are objects', () => {
    const metadata = [
      { version: '1.0.0', published: { '#text': '2020-01-01T00:00:00' } },
      { version: '9.0.0', published: { '#text': '2025-01-01T00:00:00' } },
      { version: '10.0.0', published: null },
    ];
    expect(selectDependencyMetadata(metadata, parseVersionRange('(, )'), { '#text': '2020-02-01T00:00:00' }).version).toBe('1.0.0');
    expect(selectDependencyMetadata(metadata, parseVersionRange('(, )'), null)).toBeNull();
  });

  it('reads only literal export values without leaking adjacent metadata or comments', () => {
    expect(parseManifestCommands(`@{
      # FunctionsToExport = @('Get-FakeFromComment')
      <# CmdletsToExport = @('Get-FakeFromBlock') #>
      FunctionsToExport = @('Get-MgUser', 'Set-MgUser'); PrivateData = @{ Tags = @('Get-NotACommand') }
      CmdletsToExport = 'Remove-MgUser'
      Description = 'Get-AnotherFakeCommand'
    }`)).toEqual(['Get-MgUser', 'Set-MgUser', 'Remove-MgUser']);
    expect(parseManifestCommands("@{ FunctionsToExport = @(); Description = 'Get-Fake' }")).toEqual([]);
    expect(parseManifestCommands("@{ FunctionsToExport = @('*'); Tags = @('Get-Fake') }")).toEqual([]);
  });

  it('does not inflate unrelated archive entries', () => {
    const zip = storedZip('large.dll', 'irrelevant');
    const view = new DataView(zip.buffer);
    const directory = view.getUint32(zip.length - 6, true);
    view.setUint16(directory + 10, 99, true);
    expect(parsePackageArchiveCommands(zip)).toEqual([]);
    expect(() => parseZipEntries(zip)).toThrow('Unsupported ZIP');
  });

  it('merges overlapping windows idempotently and replaces corrected release winners', () => {
    const a = winner('1.0.0', 'Get-First');
    const b = winner('2.0.0', 'Get-SecondLonger');
    const previous = state([a, b]);
    const merged = mergeIncrementalSnapshots(previous, state([a, b]));
    expect(merged.history.map((item) => item.version)).toEqual(['2.0.0', '1.0.0']);
    expect(mergeIncrementalSnapshots(merged, state([a, b]))).toEqual(merged);
    const corrected = mergeIncrementalSnapshots(previous, state([winner('2.0.0', 'Get-First')]));
    expect(corrected.history).toHaveLength(1);
    expect(corrected.current.sinceVersion).toBe('1.0.0');
  });

  it('retains complete state beyond the public ten-transition limit', () => {
    const releases = Array.from({ length: 15 }, (_, index) => winner(`${index + 1}.0.0`, `Get-Command${'X'.repeat(index)}`));
    const previous = state(releases);
    const merged = mergeIncrementalSnapshots(previous, state(releases.slice(0, 2)));
    expect(merged.winners).toHaveLength(15);
    expect(merged.history).toHaveLength(10);
    expect(merged.current.sinceVersion).toBe('15.0.0');
    expect(publicSnapshot(merged)).not.toHaveProperty('winners');
  });

  it('uses Gallery command tags without downloading archives and narrows exact dependency queries', async () => {
    const fetchImpl = vi.fn(async (url) => {
      if (url.includes('FindPackagesById')) return new Response(feed([entry({ id: 'Microsoft.Graph', version: '2.0.0', published: '2024-03-15T00:00:00Z', dependencies: 'Microsoft.Graph.Users:[2.0.0]:' })]));
      return new Response(feed([entry({ id: 'Microsoft.Graph.Users', version: '2.0.0', published: '2024-03-01T00:00:00Z' }).replace('</properties>', '<Tags>PSFunction_Get-MgUser PSCommand_Get-MgUser PSFunction_Remove-MgUserLonger</Tags></properties>')]));
    });
    const result = await collectSnapshot({ fetchImpl });
    expect(result.current.name).toBe('Remove-MgUserLonger');
    expect(fetchImpl).toHaveBeenCalledTimes(2);
    expect(new URL(fetchImpl.mock.calls[1][0]).searchParams.get('$filter')).toContain("Version eq '2.0.0'");
  });

  it('forces a full rebuild for legacy state', async () => {
    const fetchImpl = vi.fn(async () => new Response(feed([entry({ id: 'Microsoft.Graph', version: '1.0.0', published: '2024-01-01T00:00:00Z', functions: 'Get-MgUser' })])));
    const result = await collectSnapshot({ fetchImpl, previousSnapshot: { generatedAt: '2024-03-01T00:00:00Z', history: [] }, recentDays: 7 });
    expect(fetchImpl.mock.calls[0][0]).toContain('FindPackagesById');
    expect(result.stateVersion).toBe(STATE_VERSION);
  });

  it('does not treat an HTML error page as a successful empty refresh', async () => {
    await expect(collectSnapshot({ previousSnapshot: state([winner('1.0.0', 'Get-MgUser')]), recentDays: 7, fetchImpl: async () => new Response('<html>Gallery unavailable</html>') })).rejects.toThrow('no Atom feed');
  });

  it('honors Retry-After and retries failures while reading the body', async () => {
    const sleep = vi.fn(async () => {});
    const fetchImpl = vi.fn()
      .mockResolvedValueOnce(new Response('busy', { status: 429, headers: { 'retry-after': '2' } }))
      .mockResolvedValueOnce({ ok: true, text: async () => { throw new Error('connection reset'); } })
      .mockResolvedValueOnce(new Response('ok'));
    expect(await fetchText('https://example.test', { fetchImpl, sleep })).toBe('ok');
    expect(sleep.mock.calls[0][0]).toBe(2000);
    expect(fetchImpl).toHaveBeenCalledTimes(3);
  });

  it('orders numeric prerelease identifiers semantically', () => {
    expect(compareVersions('2.0.0-preview.10', '2.0.0-preview.2')).toBeGreaterThan(0);
    expect(compareVersions('2.0.0-preview', '2.0.0')).toBeLessThan(0);
  });
});

it('allows only the declared same-release dependency when module publishing follows the root', () => {
  const metadata = [
    { version: '1.0.0', published: '2024-01-01T00:00:05Z' },
    { version: '2.0.0', published: '2024-02-01T00:00:00Z' },
  ];
  expect(selectDependencyMetadata(metadata, parseVersionRange('[1.0.0, )'), '2024-01-01T00:00:00Z', '1.0.0')?.version).toBe('1.0.0');
  expect(selectDependencyMetadata(metadata, parseVersionRange('(, )'), '2024-01-01T00:00:00Z', '1.0.0')).toBeNull();
  expect(selectDependencyMetadata(metadata, parseVersionRange('[1.0.0, )'), '2024-01-01T00:00:00Z', '0.9.0')).toBeNull();
});


it('resolves legacy wildcard exports from proxy scripts without including private helpers', () => {
  const archive = storedZip([
    ['Microsoft.Graph.Users.psd1', "@{ FunctionsToExport = '*'; CmdletsToExport = @() }"],
    ['exports/ProxyCmdletDefinitions.ps1', `
      # function Get-FakeComment {}
      <# function Get-FakeBlock {} #>
      function Get-MgUser { 'function Get-FakeString {}' }
      function Remove-MgUser { }
    `],
    ['custom/private.ps1', 'function Get-PrivateHelperWithAVeryLongName {}'],
  ]);
  expect(parsePackageArchiveCommands(archive)).toEqual(['Get-MgUser', 'Remove-MgUser']);
});

it('splits version-filtered requests before exceeding the Gallery query limit', async () => {
  const ids = Array.from({ length: 26 }, (_, index) => `Microsoft.Graph.Module${index}`);
  const fetchImpl = vi.fn(async (url) => {
    if (url.includes('FindPackagesById')) return new Response(feed([entry({ id: 'Microsoft.Graph', version: '1.0.0', published: '2024-03-01T00:00:00Z', dependencies: ids.map((id) => `${id}:[1.0.0]:`).join('|') })]));
    const filter = new URL(url).searchParams.get('$filter');
    if ((filter.match(/Version eq/g) ?? []).length > 20) return new Response('OData node limit exceeded', { status: 400 });
    return new Response(feed(ids.filter((id) => filter.includes(`'${id}'`)).map((id) => entry({ id, version: '1.0.0', published: '2024-02-01T00:00:00Z', functions: 'Get-MgUser' }))));
  });
  const result = await collectSnapshot({ fetchImpl });
  expect(result.totalVersionsAnalyzed).toBe(1);
  expect(result.current.name).toBe('Get-MgUser');
  expect(fetchImpl.mock.calls.length).toBeGreaterThan(2);
});

it('caches verified empty legacy modules instead of downloading them again', async () => {
  const commandCache = new Map();
  const fetchImpl = vi.fn(async (url) => {
    if (url.includes('FindPackagesById')) return new Response(feed([entry({ id: 'Microsoft.Graph', version: '1.0.0', published: '2024-03-01T00:00:00Z', functions: 'Get-MgUser', dependencies: 'Microsoft.Graph.Empty:[1.0.0]:' })]));
    if (url.includes('/Packages()')) return new Response(feed([entry({ id: 'Microsoft.Graph.Empty', version: '1.0.0', published: '2024-02-01T00:00:00Z' })]));
    return new Response(storedZip('Microsoft.Graph.Empty.psd1', "@{ FunctionsToExport = '*'; CmdletsToExport = @() }"));
  });
  await collectSnapshot({ fetchImpl, commandCache });
  await collectSnapshot({ fetchImpl, commandCache });
  expect(fetchImpl.mock.calls.filter(([url]) => url.includes('/package/'))).toHaveLength(1);
});
