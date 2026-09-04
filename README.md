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
- `Microsoft.Graph` dependencies are used as fallback source when command metadata is not available on the root package entry

## How it works

- Fetches all historical `Microsoft.Graph` versions from PowerShell Gallery
- Reads command metadata (`Cmdlets` / `Functions`)
- Falls back to dependency package metadata when needed
- Builds a transition timeline when the longest command changes
- Caches computed data in memory for 6 hours to minimize remote calls

## Local development

```bash
npm install
npm run dev
```

Then open the local Worker URL shown by Wrangler.

## Test

```bash
npm test
```

## Deploy

```bash
npm run deploy
```
