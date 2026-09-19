# longestgraphcmdlet.com

Cloudflare Worker site that shows the longest Microsoft Graph PowerShell function/cmdlet and its recent history.

## Architecture

Visitors are served a generated data snapshot. The Worker never calls PowerShell Gallery, so Gallery timeouts, rate limits, and Cloudflare Worker subrequest limits cannot turn a page visit into a 500 response.

A GitHub Actions publishing workflow collects every historical `Microsoft.Graph` release and its `Microsoft.Graph*` dependency metadata, validates the complete result, then deploys the Worker. It runs after pushes to `main`, daily at 03:17 UTC, and on manual dispatch. A failed collection leaves the previous Worker version running.

For dependencies with an unbounded version range, the collector selects the newest matching package published no later than the root package release. Bounded ranges are resolved the same way, so historical results do not use dependency packages published later.

## Required GitHub secrets

Configure these repository or `production` environment secrets before enabling publishing:

- `CLOUDFLARE_API_TOKEN` — a least-privilege token allowed to deploy this Worker.
- `CLOUDFLARE_ACCOUNT_ID` — the Cloudflare account that owns the Worker.

The old Worker `REFRESH_TOKEN` secret is no longer used. `POST /api/refresh` now returns `410 Gone`; use **Actions → Publish data snapshot → Run workflow** to refresh immediately.

## Local development

```bash
npm install
npm test
npm run dev
```

`src/snapshot.js` is a small fixture for local development and pull-request builds. To create a real snapshot locally, run:

```bash
npm run collect
npm test
npm run dev
```

The collector replaces `src/snapshot.js` in the current workspace. Do not commit its generated output unless intentionally refreshing the development fixture.

## Verification

```bash
npm run check
```

This runs the deterministic unit/integration tests and validates the Worker deployment bundle without publishing it.
