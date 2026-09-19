import { describe, expect, it } from 'vitest';
import worker, { renderPage } from '../src/worker.js';
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
    const html = await response.text();
    expect(html).toContain(snapshot.current.name);
    expect(html).toContain('The all-time top 10');
    expect(html).not.toContain('<script');
    expect(html).not.toContain('Loading latest');
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


describe('static rendering', () => {
  it('escapes untrusted snapshot content', () => {
    const html = renderPage({ current: { name: '<script>alert(1)</script>', packageId: '"<&', length: 1 }, history: [], topLongest: [] });
    expect(html).toContain('&lt;script&gt;alert(1)&lt;/script&gt;');
    expect(html).not.toContain('<script>');
  });

  it.each(['/', '/api/longest'])('supports HEAD for %s without a response body', async (path) => {
    const get = await worker.fetch(new Request(`https://example.com${path}`));
    const head = await worker.fetch(new Request(`https://example.com${path}`, { method: 'HEAD' }));
    expect(head.status).toBe(200);
    expect(head.headers.get('content-type')).toBe(get.headers.get('content-type'));
    expect(await head.text()).toBe('');
  });
});

it('publishes internally consistent command lengths, dates, and distinct rankings', () => {
  expect(snapshot.current.length).toBe(snapshot.current.name.length);
  expect(snapshot.totalVersionsAnalyzed).toBe(new Set(snapshot.analyzedVersions).size);
  expect(snapshot.history[0].name).toBe(snapshot.current.name);
  expect(new Set(snapshot.history.map((item) => item.version)).size).toBe(snapshot.history.length);
  expect(new Set(snapshot.topLongest.map((item) => item.name)).size).toBe(snapshot.topLongest.length);
  for (const item of snapshot.topLongest) {
    expect(item.length).toBe(item.name.length);
    expect(snapshot.analyzedVersions).toContain(item.versionFound);
  }
  for (const item of snapshot.history) {
    expect(typeof item.publishedAt).toBe('string');
    expect(new Date(item.publishedAt).getUTCFullYear()).toBeGreaterThan(2000);
    expect(snapshot.analyzedVersions).toContain(item.version);
  }
});
