import { describe, expect, it } from 'vitest';
import worker from '../src/worker.js';
import snapshot from '../src/snapshot.js';

describe('worker endpoints', () => {
  it('returns the build-time snapshot from the public API', async () => {
    const response = await worker.fetch(new Request('https://example.com/api/longest'));

    expect(response.status).toBe(200);
    expect(response.headers.get('content-type')).toContain('application/json');
    expect(response.headers.get('cache-control')).toContain('stale-while-revalidate');
    expect(await response.json()).toEqual(snapshot);
  });

  it('returns HTML on the root route', async () => {
    const response = await worker.fetch(new Request('https://example.com/'));

    expect(response.status).toBe(200);
    expect(response.headers.get('content-type')).toContain('text/html');
    expect(await response.text()).toContain('Longest Microsoft Graph PowerShell Function/Cmdlet');
  });

  it('retires the runtime refresh endpoint', async () => {
    const response = await worker.fetch(new Request('https://example.com/api/refresh', { method: 'POST' }));

    expect(response.status).toBe(410);
    expect((await response.json()).error).toContain('GitHub Actions');
  });

  it('returns 404 for unknown routes', async () => {
    const response = await worker.fetch(new Request('https://example.com/unknown'));

    expect(response.status).toBe(404);
  });
});
