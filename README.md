# MarketCleaner AI

MarketCleaner is a browser-based stock-data anomaly viewer. Load CSV or Excel data, inspect flagged prices, and export your results. The built-in sample data lets you try it without an upload.

## Run locally

With Node.js 22 or 24 and npm:

```sh
npm ci
npm run dev
```

For a production bundle, run `npm run build`. To serve that bundle locally, run `npm run preview`. Excel import loads SheetJS in the browser and requires access to its CDN.

## Scope

Outlier flags are a review aid, not a guarantee that data is clean or a reason to automatically discard rows. This repository is the maintained codebase for the cleaner.
