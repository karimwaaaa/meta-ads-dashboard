/**
 * Config.example.gs — NOT part of the real project's source, and not auto-loaded by
 * anything else in this repo.
 *
 * The other four files (Code.gs, DataPipeline.gs, AdSetPipeline.gs, AdPipeline.gs) are the
 * real pipeline and dashboard code, lightly redacted for a public portfolio (sheet IDs,
 * a couple of internal object IDs used only as log examples, and the demo brand name are
 * placeholder'd — see each file's SETUP comments). They all assume three things exist as
 * global Apps Script variables/functions, defined in a small Config file that isn't
 * included here because it also carries a live access token:
 *
 *   - ACCOUNT_ID    (string)   e.g. 'act_1234567890'  — your Meta ad account ID.
 *   - ACCESS_TOKEN  (string)             — a long-lived Meta Marketing API access token.
 *   - fetchCampaignSettings()  (function) — a small helper that reads campaign-level
 *     settings (objective, bid strategy, daily budget, etc.) directly from the Marketing
 *     API, used to enrich the daily rows written by DataPipeline.gs.
 *
 * This file sketches the shape of that dependency so the repo is self-explanatory to a
 * reader — it is NOT the real implementation, just enough to show how the pieces fit
 * together and to get a fresh copy of this project running end to end.
 */

// ── Fill these in with your own values (Meta for Developers → your app → Marketing API) ──
var ACCOUNT_ID = '[ACCOUNT_ID]';       // e.g. 'act_1234567890' — no quotes around the numbers, just the string
var ACCESS_TOKEN = '[ACCESS_TOKEN]';   // generate via Graph API Explorer or a System User token; never commit a real one

// Illustrative shape only — the real dashboard's version pulls objective, bid_strategy,
// daily_budget, and a couple of other fields per campaign and merges them into each
// daily row DataPipeline.gs writes.
function fetchCampaignSettings() {
  var url = 'https://graph.facebook.com/v18.0/' + ACCOUNT_ID +
    '/campaigns?fields=id,name,objective,bid_strategy,daily_budget&limit=200';
  var resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ACCESS_TOKEN },
    muteHttpExceptions: true
  });
  var data = JSON.parse(resp.getContentText());
  var map = {};
  (data.data || []).forEach(function (c) {
    map[c.name] = { objective: c.objective, bidStrategy: c.bid_strategy, dailyBudget: c.daily_budget };
  });
  return map;
}
