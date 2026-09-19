import { describe, expect, it, vi } from 'vitest';
import {
  collectSnapshot,
  compareVersions,
  fetchText,
  parseDependencyString,
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
      entry({ id: 'Microsoft.Graph.Users', version: '1.0.0', published: '2024-01-01T00:00:00Z', cmdlets: 'Get-MgFormerLongestCommand' }),
      entry({ id: 'Microsoft.Graph.Users', version: '2.0.0', published: '2024-03-01T00:00:00Z', cmdlets: 'Get-MgCurrentLongestCommandName' }),
    ]);
    const fetchImpl = vi.fn(async (url) => {
      if (url === initialRootUrl) return new Response(rootFirstPage);
      if (url === nextRootUrl) return new Response(rootSecondPage);
      if (url.includes('/Packages()')) return new Response(dependencyFeed);
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
    expect(fetchImpl).toHaveBeenCalledTimes(3);
  });
});
