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

## Try the missing-close regression fixture

Upload [fixtures/missing-close-and-real-drop.csv](fixtures/missing-close-and-real-drop.csv) with the default 10% move threshold:

- November 17: `Missing close`. The row stays visible and the Close cell stays blank in the exported CSV.
- November 16 → 17 and November 17 → 18: no returns are computed. The app does not compare November 16 directly with November 18.
- November 20: `Black Swan Drop (-20.6%)`. Both endpoints are valid, so the drop remains available for review.
- **Export with Flags (All Rows)** keeps all five rows and separates validation issues from outlier flags. **Export Unflagged Rows Only** deliberately excludes both kinds of flagged rows.
