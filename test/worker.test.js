import { describe, expect, it, vi } from 'vitest';
import worker, { _internals } from '../src/worker.js';

describe('internal helpers', () => {
  it('parses dependency strings with versions', () => {
    const value = "Microsoft.Graph.Users:[2.1.0, )|Microsoft.Graph.Identity:2.1.0|Other.Package:1.0.0";

    expect(_internals.parseDependencyString(value)).toEqual([
      { id: 'Microsoft.Graph.Users', version: '2.1.0' },
      { id: 'Microsoft.Graph.Identity', version: '2.1.0' },
      { id: 'Other.Package', version: '1.0.0' },
    ]);
  });

  it('computes transition history and current winner', () => {
    const response = _internals.buildResponse([
      {
        version: '1.0.0',
        published: '2024-01-01T00:00:00Z',
        longestName: 'Get-MgAlpha',
        packageId: 'Microsoft.Graph',
      },
      {
        version: '1.1.0',
        published: '2024-02-01T00:00:00Z',
        longestName: 'Get-MgAlpha',
        packageId: 'Microsoft.Graph',
      },
      {
        version: '1.2.0',
        published: '2024-03-01T00:00:00Z',
        longestName: 'Get-MgVeryVeryLongCommandName',
        packageId: 'Microsoft.Graph',
      },
    ]);

    expect(response.current.name).toBe('Get-MgVeryVeryLongCommandName');
    expect(response.current.sinceVersion).toBe('1.2.0');
    expect(response.previous.name).toBe('Get-MgAlpha');
    expect(response.history).toHaveLength(2);
  });

  it('records transition when package changes for same command', () => {
    const response = _internals.buildResponse([
      {
        version: '1.0.0',
        published: '2024-01-01T00:00:00Z',
        longestName: 'Get-MgVeryLongName',
        packageId: 'Microsoft.Graph.Users',
      },
      {
        version: '1.1.0',
        published: '2024-02-01T00:00:00Z',
        longestName: 'Get-MgVeryLongName',
        packageId: 'Microsoft.Graph.Identity',
      },
    ]);

    expect(response.history).toHaveLength(2);
    expect(response.current.packageId).toBe('Microsoft.Graph.Identity');
    expect(response.previous.packageId).toBe('Microsoft.Graph.Users');
  });

  it('selects dependency command when longer than root command', async () => {
    const root = {
      id: 'Microsoft.Graph',
      dependencies: 'Microsoft.Graph.Users:2.0.0',
      cmdlets: 'Get-MgShort',
      functions: '',
    };

    const winner = await _internals.resolveWinnerForVersion(
      root,
      new Map(),
      async () => ({
        id: 'Microsoft.Graph.Users',
        cmdlets: 'Get-MgReallyReallyLongCommandName',
        functions: '',
      }),
    );

    expect(winner).toEqual({
      longestName: 'Get-MgReallyReallyLongCommandName',
      packageId: 'Microsoft.Graph.Users',
    });
  });

  it('uses lexicographic tie-breaker across root and dependency commands', async () => {
    const root = {
      id: 'Microsoft.Graph',
      dependencies: 'Microsoft.Graph.Users:2.0.0',
      cmdlets: 'Set-MgTieCommand',
      functions: '',
    };

    const winner = await _internals.resolveWinnerForVersion(
      root,
      new Map(),
      async () => ({
        id: 'Microsoft.Graph.Users',
        cmdlets: 'Get-MgTieCommand',
        functions: '',
      }),
    );

    expect(winner).toEqual({
      longestName: 'Get-MgTieCommand',
      packageId: 'Microsoft.Graph.Users',
    });
  });

  it('keeps root winner when recoverable dependency metadata fetch fails', async () => {
    const root = {
      id: 'Microsoft.Graph',
      dependencies: 'Microsoft.Graph.Users:2.0.0',
      cmdlets: 'Get-MgLongestRootCommandName',
      functions: '',
    };

    const winner = await _internals.resolveWinnerForVersion(
      root,
      new Map(),
      async () => {
        throw new _internals.PowerShellGalleryError('fetch failed', 'UPSTREAM_REQUEST_FAILED');
      },
    );

    expect(winner).toEqual({
      longestName: 'Get-MgLongestRootCommandName',
      packageId: 'Microsoft.Graph',
    });
  });

  it('rethrows non-recoverable dependency metadata errors', async () => {
    const root = {
      id: 'Microsoft.Graph',
      dependencies: 'Microsoft.Graph.Users:2.0.0',
      cmdlets: 'Get-MgRootCommandName',
      functions: '',
    };

    await expect(
      _internals.resolveWinnerForVersion(
        root,
        new Map(),
        async () => {
          throw new Error('unexpected failure');
        },
      ),
    ).rejects.toThrow('unexpected failure');
  });

  it('skips uncached dependencies when fetch budget is exhausted', async () => {
    const root = {
      id: 'Microsoft.Graph',
      dependencies: 'Microsoft.Graph.Users:2.0.0|Microsoft.Graph.Identity:2.0.0',
      cmdlets: 'Get-MgRootCommandName',
      functions: '',
    };
    const packageMemo = new Map([
      [
        'Microsoft.Graph.Users@2.0.0',
        Promise.resolve({
          id: 'Microsoft.Graph.Users',
          cmdlets: 'Get-MgLongerCachedDependencyCommand',
          functions: '',
        }),
      ],
    ]);
    const metadataLoader = vi.fn(async (packageId, version, memo) => {
      const cached = memo.get(`${packageId}@${version}`);
      if (cached) return cached;
      return {
        id: packageId,
        cmdlets: 'Get-MgVeryVeryLongUncachedDependencyCommand',
        functions: '',
      };
    });

    const winner = await _internals.resolveWinnerForVersion(
      root,
      packageMemo,
      metadataLoader,
      { dependencyFetchBudget: { remaining: 0 } },
    );

    expect(metadataLoader).toHaveBeenCalledTimes(1);
    expect(winner).toEqual({
      longestName: 'Get-MgLongerCachedDependencyCommand',
      packageId: 'Microsoft.Graph.Users',
    });
  });
});

describe('worker endpoints', () => {
  it('returns HTML on root route', async () => {
    const response = await worker.fetch(new Request('https://example.com/'));

    expect(response.status).toBe(200);
    expect(response.headers.get('content-type')).toContain('text/html');
    const html = await response.text();
    expect(html).toContain('Longest Microsoft Graph PowerShell Function/Cmdlet');
  });

  it('returns refresh response', async () => {
    _internals.setCache({ seeded: true }, Date.now() + 60_000);
    const response = await worker.fetch(
      new Request('https://example.com/api/refresh', {
        method: 'POST',
        headers: { 'x-refresh-token': 'test-token' },
      }),
      { REFRESH_TOKEN: 'test-token' },
    );

    expect(response.status).toBe(200);
    const json = await response.json();
    expect(json.ok).toBe(true);
    expect(_internals.getCache()).toEqual({
      data: null,
      expiresAt: 0,
    });
  });

  it('rejects refresh without token', async () => {
    const response = await worker.fetch(new Request('https://example.com/api/refresh', { method: 'POST' }));
    expect(response.status).toBe(401);
  });

  it('accepts bearer authorization token', async () => {
    _internals.setCache({ seeded: true }, Date.now() + 60_000);
    const bearer = ['Bearer', 'test-token'].join(' ');
    const response = await worker.fetch(
      new Request('https://example.com/api/refresh', {
        method: 'POST',
        headers: { authorization: bearer },
      }),
      { REFRESH_TOKEN: 'test-token' },
    );

    expect(response.status).toBe(200);
    expect((await response.json()).ok).toBe(true);
  });

  it('rejects incorrect bearer authorization token', async () => {
    const bearer = ['Bearer', 'wrong-token'].join(' ');
    const response = await worker.fetch(
      new Request('https://example.com/api/refresh', {
        method: 'POST',
        headers: { authorization: bearer },
      }),
      { REFRESH_TOKEN: 'test-token' },
    );

    expect(response.status).toBe(401);
  });
});
