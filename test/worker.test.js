import { describe, expect, it } from 'vitest';
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
    const response = await worker.fetch(new Request('https://example.com/api/refresh', { method: 'POST' }));

    expect(response.status).toBe(200);
    const json = await response.json();
    expect(json.ok).toBe(true);
  });
});
