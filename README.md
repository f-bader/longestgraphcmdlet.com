# longestgraphcmdlet.com

Cloudflare Worker site that shows the longest Microsoft Graph PowerShell function/cmdlet of all time and its record history. The headline and top ten consider every analyzed release, including commands removed from later releases; history tracks changes to the all-time record.

## Architecture

Visitors receive complete HTML rendered from a generated snapshot, with no client-side JavaScript or extra data request. The JSON API serves the same snapshot. The Worker never calls PowerShell Gallery, so Gallery timeouts, rate limits, and Cloudflare Worker subrequest limits cannot turn a page visit into a 500 response.

A GitHub Actions publishing workflow collects every historical `Microsoft.Graph` release on its first run, validates the complete result, then deploys the Worker. It runs after pushes to `main`, daily at 03:17 UTC, and on manual dispatch. Later runs restore versioned collection state and inspect releases since the last successful collection, with a seven-day overlap. This catches up after missed runs. Results are merged by release version before rebuilding history, so overlapping runs are repeatable. Legacy state automatically triggers a full rebuild. A failed collection leaves the previous Worker version running.

The collector normalizes Gallery dates to UTC and uses Created when an unlisted package has a sentinel Published date. Dependency ranges select the newest matching version published by the root release. If no such version exists, the root’s explicitly declared minimum version may be used only when it equals the root version: Gallery sometimes publishes modules just after the umbrella package in the same release. Missing dates or unresolved dependencies fail collection.

Exported command tags avoid archive downloads when available. Otherwise, only manifests and, for legacy wildcard exports, generated proxy scripts are decompressed from the archive, and compact command summaries are cached by package/version. Literal manifest exports are parsed without executing PowerShell. Exact dependency queries are narrowed to the required versions.

Collection state retains compact winners and top-ten candidates for every release. Only the public summary is bundled into the Worker.

## Required GitHub secrets

Configure these repository or `production` environment secrets before enabling publishing:

- `CLOUDFLARE_API_TOKEN` — a least-privilege token allowed to deploy this Worker.
- `CLOUDFLARE_ACCOUNT_ID` — the Cloudflare account that owns the Worker.

The old Worker `REFRESH_TOKEN` secret is no longer used. `POST /api/refresh` now returns `410 Gone`; use **Actions → Publish data snapshot → Run workflow** to refresh immediately.

The manual workflow has a `full_collection` input for rebuilding all history when needed. Snapshot state and command summaries are kept in the GitHub Actions cache after a successful deployment; if that cache is unavailable, the collector safely falls back to a full first run.

## Local development

```bash
npm install
npm test
npm run dev
```

`src/snapshot.js` is the checked-in snapshot for local development and pull-request builds. To create a real snapshot locally, run:

```bash
npm run collect
npm test
npm run dev
```

The collector replaces `src/snapshot.js` in the current workspace. Commit its generated output only when intentionally refreshing the development snapshot. Local state lives in the ignored `.cache/` directory. To rebuild every release, use `FULL_COLLECTION=true npm run collect`.

## Verification

```bash
npm run check
```

This runs the deterministic unit/integration tests and validates the Worker deployment bundle without publishing it.
