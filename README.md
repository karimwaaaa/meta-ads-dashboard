# Meta Ads Dashboard

A self-healing Meta (Facebook/Instagram) Ads reporting dashboard built on Google Apps Script and Google Sheets — no paid BI tool, no server to host. It pulls Campaign, Ad Set, Ad, and Creative-level data from the Meta Marketing API on a schedule, stores it in Sheets as a durable data layer, and serves an interactive web dashboard on top of it.

**[Live demo](#)** — fake data, fully interactive, no backend required. *(Replace this link once GitHub Pages is enabled — see below.)*

> This is a sanitized copy of a dashboard I built and maintain for a real advertiser. Sheet IDs, the Meta account ID/access token, and the original brand name have been replaced with placeholders; benchmark thresholds have been replaced with illustrative example values. Everything else — the pipeline logic, the rate-limit handling, the caching strategy, the dashboard itself — is the real, working implementation.

## Why this exists

Meta's own Ads Manager reporting is slow to filter, hard to customize, and doesn't do cross-level (Campaign → Ad Set → Ad → Creative) comparisons well. Paid BI tools solve this but cost money and add another vendor. This project reads directly from the Marketing API into Google Sheets on a schedule, then serves a purpose-built dashboard from the same Apps Script project — free to run, fully customizable, and the data lives somewhere you already control.

## What it does

- **Four-level reporting**: Campaign, Ad Set, Ad, and Creative (thumbnail-level) views, each with their own filters and date ranges.
- **Self-healing daily pipeline**: re-fetches and overwrites a rolling 7-day window on every run (matching Meta's 7-day click attribution window), so late-attributed conversions correct themselves automatically without manual backfills.
- **Creative resolution with real caching**: walks each ad's creative shape (`object_story_spec`, `asset_feed_spec` for Advantage+/dynamic creative, carousels, video) to find the highest-resolution image/thumbnail available, caches the result (including permanent misses) so it's never re-fetched, and falls back gracefully when nothing better exists.
- **Rate-limit and timeout aware by design**: every Marketing API call goes through shared retry/backoff helpers that distinguish transient errors (rate limits, timeouts — retry later) from permanent ones (deleted asset, bad hash — cache and stop asking), and large paginated pulls are sized to stay under both Apps Script's 6-minute execution ceiling and Meta's per-request response-size limit.
- **A benchmark scorecard**: each ad is scored Poor/Average/Good/Excellent against configurable per-objective thresholds (Messaging / Sales / Generic), with a plain-language verdict and "why."
- **Freshness you can see**: a status strip shows when each pipeline last *genuinely* succeeded (not just "when did this page load"), plus a clearly-separated "Today (Live)" widget that pulls today's numbers straight from the API — visually and textually marked as provisional so it's never confused with the settled historical data.
- **Diagnostic tooling**: standalone functions to find and classify creative-resolution failures (permanently dead vs. simply not-yet-attempted vs. cached-but-not-yet-synced-to-the-sheet) instead of guessing.

## Architecture

```
Meta Marketing API
        │
        ▼
┌─────────────────────────────────────────────┐
│  Apps Script time-driven triggers (daily)    │
│  DataPipeline.gs   → "data - by day"         │
│  AdSetPipeline.gs  → "data - adset - by day" │
│  AdPipeline.gs     → "data - ad - by day"     │
│                     + "ad_creatives" (cached  │
│                       thumbnails/full images) │
└─────────────────────────────────────────────┘
        │  (Google Sheets — durable data layer)
        ▼
┌─────────────────────────────────────────────┐
│  Code.gs  — Apps Script Web App backend      │
│  Reads/aggregates the sheets on demand,      │
│  scores creatives, and exposes a small       │
│  google.script.run API to the front end.     │
│  Also proxies a lightweight, cached "live    │
│  today" pull straight from the Meta API.     │
└─────────────────────────────────────────────┘
        │
        ▼
┌─────────────────────────────────────────────┐
│  index.html (Dashboard.html) — the web app   │
│  A single-file HTML/CSS/JS dashboard served  │
│  by the Web App deployment. Falls back to    │
│  fake demo data automatically when there's   │
│  no Apps Script bridge available (this is    │
│  what makes the GitHub Pages demo possible). │
└─────────────────────────────────────────────┘
```

## Tech stack

- **Google Apps Script** (V8 runtime) — scheduled data pipelines + the web app backend
- **Google Sheets** — the data layer (also gives you a free audit trail / manual-edit escape hatch)
- **Vanilla JS, HTML, CSS** for the dashboard — no framework, no build step, one file
- **Chart.js** and **Font Awesome** via CDN for charts/icons
- **Meta Marketing Graph API** (v18.0) as the data source

## Repo layout

```
├── index.html                    # the dashboard (Dashboard.html in the real project) —
│                                  #   open it directly, or deploy it as an Apps Script Web App
└── apps-script/
    ├── Code.gs                   # Web App backend: getDashboardData(), scoring, freshness API
    ├── DataPipeline.gs           # Campaign-level daily import + backfill
    ├── AdSetPipeline.gs          # Ad Set-level daily import + shared fetch/retry helpers
    ├── AdPipeline.gs             # Ad + Creative-level import, image/video thumbnail cache,
    │                              #   diagnostics
    └── Config.example.gs         # NOT real code — documents the 3 globals (ACCOUNT_ID,
                                   #   ACCESS_TOKEN, fetchCampaignSettings()) the pipeline
                                   #   files expect, so the repo is self-contained to read
```

## Running your own copy

1. **Create three Google Sheets** (Campaign, Ad Set, Ad workbooks — the real project splits them to keep each one's row count manageable) and copy each one's ID from its URL.
2. **Create a new Apps Script project** (script.google.com → New project, or bind one to a Sheet via Extensions → Apps Script).
3. Paste in `Code.gs`, `DataPipeline.gs`, `AdSetPipeline.gs`, `AdPipeline.gs`, and `Dashboard.html` (rename `index.html` back to `Dashboard.html` — Apps Script's HTML Service expects that file type).
4. Fill in the placeholders:
   - `[SHEET_ID]` in `Code.gs`, `[AD_SHEET_ID]` in `AdPipeline.gs`, `[ADSET_SHEET_ID]` in `AdSetPipeline.gs` → your three real Sheet IDs.
   - `[ACCOUNT_ID]` / `[ACCESS_TOKEN]` → your Meta ad account ID and a Marketing API access token (see `Config.example.gs` for the shape; you'll need to write your own version of this file since it isn't included).
   - `[COMPANY_NAME]` in setup comments, and the `AD_BENCHMARK_BANDS_` thresholds in `Code.gs` → your own vertical's real benchmark numbers, if you have them.
5. Set up daily time-driven triggers (Apps Script → Triggers) for `importDataFromJSON` / `copyDataOnceADay`, `importAdSetData`, `importAdData`, and `importAdCreatives`.
6. Deploy `Code.gs` + `Dashboard.html` as a Web App (Deploy → New deployment → Web app) to get a live, backend-connected dashboard.

Or skip all of that and just open `index.html` in a browser — it runs entirely on fake demo data with zero setup, which is exactly what the GitHub Pages link above does.

## A note on the redactions

Everything a reader would need to reconstruct the real advertiser's account, spreadsheets, or historical performance has been removed or replaced with an illustrative placeholder. Everything a reader would need to evaluate the actual engineering — the rate-limit handling, the caching design, the self-healing pipeline, the diagnostic tooling, the dashboard itself — is untouched and fully functional.
