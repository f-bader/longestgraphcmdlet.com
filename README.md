# longestgraphcmdlet.com

Cloudflare Worker site that tracks the longest function/cmdlet name across `Microsoft.Graph` PowerShell module metadata and historical package versions from PowerShell Gallery.

## What it shows

- Current longest cmdlet/function name
- Since which `Microsoft.Graph` version/date it has been the longest
- Previous longest cmdlet
- Last 10 longest-name transitions

## Data source

- `https://www.powershellgallery.com/packages/Microsoft.Graph`
- PowerShell Gallery OData (`/api/v2`) package metadata
- `Microsoft.Graph*` dependencies are also evaluated per version, and the overall longest command is selected from root + dependency metadata

## How it works

- Fetches all historical `Microsoft.Graph` versions from PowerShell Gallery
- Reads command metadata (`Cmdlets` / `Functions`)
- Compares the longest command from the root package with the longest command from dependency packages
- Builds a transition timeline when the longest command changes
- Caches computed data in memory for 24 hours (daily refresh cadence) to minimize remote calls

## Local development

```bash
npm install
npm run dev
```

Then open the local Worker URL shown by Wrangler.

To use the protected refresh endpoint:

```bash
export REFRESH_TOKEN=your-refresh-token
```

## Test

```bash
npm test
```

## Deploy

```bash
npm run deploy
```

Set `REFRESH_TOKEN` as a Cloudflare Worker secret for production:

```bash
npx wrangler secret put REFRESH_TOKEN
```
