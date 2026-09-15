/**
 * DataPipeline.gs  —  Meta Ads Dashboard (portfolio copy)
 *
 * UPDATED (Aug 2026):
 *   - "Results" is no longer hardcoded to purchases. It's now optimization-goal-aware
 *     (see OPTIMIZATION_GOAL_ACTION_MAP_ / generalizeResults_) so Awareness/Engagement/
 *     Lead campaigns stop showing 0 Results. "Purchases" stays purchases-only — the two
 *     are now genuinely different numbers for non-Sales campaigns, on purpose.
 *   - Added: Optimization goal, Landing page views, Initiate checkouts, ThruPlays
 *     (15-second video plays) — 4 new columns on both "data - by day" and
 *     "data - adset - by day" (AdSetPipeline.gs reuses buildRowFromInsight_).
 *   - Because "Results" changes definition, you MUST re-run the backfill functions
 *     once after installing this file (see BACKFILL GUIDE at the bottom) so historical
 *     rows reflect the corrected numbers instead of the old purchases-only figure.
 *
 * REPLACES your old importDataFromJSON() and copyDataOnceADay().
 * Before pasting this file in:
 *   1. In importDataFromJSON.gs, DELETE the old importDataFromJSON() and the old
 *      copyDataOnceADay(), and delete ONE of the two duplicate getYesterday()s.
 *   2. Keep ACCOUNT_ID, ACCESS_TOKEN, fetchCampaignSettings(), and one getYesterday()
 *      in that file — this file reuses them (globals are shared across .gs files).
 *   3. ALSO DELETE backfill12Months() and backfillBreakdownData() from
 *      importDataFromJSON.gs — they are legacy, don't clear-before-write like
 *      backfillDateRange()/importBreakdownRange() do, and will DUPLICATE rows if
 *      anyone ever runs them again by accident. backfillDateRange() below and
 *      importBreakdownRange() in BreakdownPipeline.gs are their idempotent replacements.
 */

// How many days back the daily job re-pulls and overwrites each run.
// Set >= your attribution window (you're on 7-day click) so late conversions get corrected.
// It also self-heals gaps: a missed night is refilled on the next successful run,
// as long as the outage was shorter than this many days.
var LOOKBACK_DAYS = 7;

// ── OPTIMIZATION-GOAL-AWARE "RESULTS" ──
// Meta's own dashboard defines "Results" as whatever action type matches the ad set's
// optimization goal. Your sheet was hardcoding it to purchases, so anything that wasn't
// a Sales campaign silently reported 0 Results. This map fixes that. Extend it if you
// run campaigns with an optimization goal not listed here — unmapped goals fall back to
// the old purchase-based behavior so nothing breaks, it just won't be "smart" for that goal.
var OPTIMIZATION_GOAL_ACTION_MAP_ = {
  'OFFSITE_CONVERSIONS':      ['omni_purchase', 'purchase'],
  'ONSITE_CONVERSIONS':       ['omni_purchase', 'purchase'],
  'VALUE':                    ['omni_purchase', 'purchase'],
  'LANDING_PAGE_VIEWS':       ['landing_page_view'],
  'LINK_CLICKS':              ['link_click'],
  'LEAD_GENERATION':          ['lead', 'onsite_conversion.lead_grouped', 'offsite_conversion.fb_pixel_lead'],
  'QUALITY_LEAD':             ['lead', 'onsite_conversion.lead_grouped'],
  'CONVERSATIONS':            ['onsite_conversion.messaging_conversation_started_7d'],
  'APP_INSTALLS':             ['app_install', 'mobile_app_install', 'omni_app_install'],
  'THRUPLAY':                 ['__thruplay__'],   // special-cased below (own field, not in `actions`)
  'POST_ENGAGEMENT':          ['post_engagement'],
  'PAGE_LIKES':               ['like'],
  'REACH':                    ['__impressions__'],       // no conversion concept — special-cased
  'IMPRESSIONS':              ['__impressions__']
};

function av_(arr, type) {
  if (!arr) return 0;
  var f = arr.filter(function (a) { return a.action_type === type; });
  return f.length ? (parseFloat(f[0].value) || 0) : 0;
}

// Sums every action_type mapped to this optimization goal. Falls back to the classic
// purchase-only formula for goals we haven't mapped (never regresses existing behavior).
function generalizeResults_(item, optimizationGoal, purchaseCount, thruplayCount) {
  var types = OPTIMIZATION_GOAL_ACTION_MAP_[optimizationGoal];
  if (!types) return purchaseCount; // unmapped goal — old behavior
  if (types[0] === '__thruplay__') return thruplayCount;
  if (types[0] === '__impressions__') return parseInt(item.impressions) || 0;
  var sum = 0;
  types.forEach(function (t) { sum += av_(item.actions, t); });
  return sum;
}

// One shared row builder so daily + backfill produce IDENTICAL columns.
function buildRowFromInsight_(item, campaignSettings) {
  var av = av_; // local alias, keeps the body below unchanged
  var roasValue = (item.purchase_roas && item.purchase_roas.length) ? (parseFloat(item.purchase_roas[0].value) || 0) : 0;
  var videoPlays      = av(item.actions, 'video_view');
  var postEngagements = av(item.actions, 'post_engagement');
  var purchases        = av(item.actions, 'omni_purchase') + av(item.actions, 'purchase');
  var blockedContacts = av(item.actions, 'onsite_conversion.messaging_block');
  var costPerMessage  = av(item.cost_per_action_type, 'onsite_conversion.messaging_conversation_started_7d');
  var conversionValue = av(item.action_values, 'omni_purchase') + av(item.action_values, 'purchase');

  // ── NEW: funnel + creative-spend fields ──
  var landingPageViews = av(item.actions, 'landing_page_view');
  var initiateCheckouts = av(item.actions, 'omni_initiated_checkout') + av(item.actions, 'initiate_checkout');
  var thruplays = av(item.video_thruplay_watched_actions, 'video_view');
  var optimizationGoal = item.optimization_goal || '';
  var results = generalizeResults_(item, optimizationGoal, purchases, thruplays);

  // Real video-milestone counts (these were hard-coded 0 before)
  var video25  = av(item.video_p25_watched_actions,  'video_view');
  var video50  = av(item.video_p50_watched_actions,  'video_view');
  var video75  = av(item.video_p75_watched_actions,  'video_view');
  var video95  = av(item.video_p95_watched_actions,  'video_view');
  var video100 = av(item.video_p100_watched_actions, 'video_view');

  var impressions = parseInt(item.impressions) || 0;
  // FIXED (Aug 2026, per your explicit go-ahead to touch campaign data for bug fixes):
  // item.clicks is ALL click types (likes, comments, photo expands — not just link
  // clicks). This used to compute all-clicks CTR and store it under the "CTR (link
  // click-through rate)" label — identical to CTR (all), just recomputed a second time.
  // inline_link_clicks is Meta's real link-click count; this is now genuine link CTR.
  // Same fix already applied to the breakdown sheets (BreakdownPipeline.gs) — this
  // closes the matching gap in the main "data - by day" / "data - adset - by day" tables.
  var linkClicks  = parseInt(item.inline_link_clicks) || 0;
  var ctrLink     = impressions > 0 ? (linkClicks / impressions) * 100 : 0;
  var frequency   = parseFloat(item.frequency) || 0;
  var ctrAll      = parseFloat(item.ctr) || 0;

  var name = item.campaign_name || '';
  var st = campaignSettings[name] || {};

  // Hook = 3-sec plays / impressions (how many people the ad hooked out of everyone who
  // saw it); Hold = 100% plays / 3-sec plays (of the people it hooked, how many it held to
  // the end). FIXED (Sep 2026, "hold rate is missing a whole number — Ads Manager says 33%,
  // dashboard says 3.3%"): this used to divide by impressions for BOTH, matching Hook Rate's
  // formula instead of Hold Rate's. That's the exact-10x bug — impressions is roughly 10x
  // 3-sec plays for typical creative, so dividing by the wrong (much bigger) denominator
  // produced a number ~10x too small. The benchmark PDF's own Hold Rate bounds ([20,30,40])
  // only make sense as a retention-of-hooked-viewers rate — a completions/impressions ratio
  // is almost never in that range — which is corroborating evidence this was the formula
  // actually intended, not just a wrong number.
  // RE-APPLIED (Sep 2026): your own correction — "hold rate formula is ThruPlays / 3 second
  // video plays" — got lost somewhere between an earlier round of this project and what
  // actually ended up in this file (this line was still on the OLD video100/videoPlays
  // formula when this file was re-read just now, not the ThruPlays version). Fixing it here
  // for real this time: ThruPlays (Meta's own 15-second-or-completion threshold) is a
  // materially easier, larger count than literal 100%-completion for any video over ~15
  // seconds, so this reads meaningfully HIGHER than the old formula for longer videos.
  var hookRate = impressions > 0 ? (videoPlays / impressions) * 100 : 0;
  var holdRate = videoPlays > 0 ? (thruplays / videoPlays) * 100 : 0;

  // 35 columns, in the exact order of the "data - by day" header row
  return [
    name,                                  // 1  Campaign name
    item.objective || '',                  // 2  Objective
    item.attribution_setting || '',        // 3  Attribution setting
    st.bid_strategy || '',                 // 4  Campaign bid strategy
    item.date_stop || '',                  // 5  Reporting ends
    st.daily_budget || 0,                  // 6  Daily budget
    item.engagement_rate_ranking || '',    // 7  Engagement rate ranking
    item.quality_ranking || '',            // 8  Quality ranking
    frequency,                             // 9  Frequency
    parseFloat(item.cpm) || 0,             // 10 CPM
    parseInt(item.reach) || 0,             // 11 Reach
    results,                               // 12 Results (NOW optimization-goal-aware)
    blockedContacts,                       // 13 Blocked messaging contacts
    parseFloat(item.spend) || 0,           // 14 Amount spent
    videoPlays,                            // 15 Video plays
    postEngagements,                       // 16 Post engagements
    ctrAll,                                // 17 CTR (all)
    ctrLink,                               // 18 CTR (link click-through rate)
    videoPlays,                            // 19 3-second video plays (= video_view)
    video25,                               // 20 Video plays at 25%
    video50,                               // 21 Video plays at 50%
    video75,                               // 22 Video plays at 75%
    video95,                               // 23 Video plays at 95%
    video100,                              // 24 Video plays at 100%
    costPerMessage,                        // 25 Cost per messaging conversation started
    roasValue,                             // 26 Purchase ROAS
    conversionValue,                       // 27 Purchases conversion value
    purchases,                             // 28 Purchases (STILL purchases-only, on purpose)
    item.conversion_rate_ranking || '',    // 29 Conversion rate ranking
    hookRate,                              // 30 Hook rate
    holdRate,                              // 31 Hold rate
    optimizationGoal,                      // 32 Optimization goal            ← NEW
    landingPageViews,                      // 33 Landing page views          ← NEW
    initiateCheckouts,                     // 34 Initiate checkouts          ← NEW
    thruplays                              // 35 ThruPlays (15-sec plays)    ← NEW
  ];
}

var INSIGHT_FIELDS_ = [
  'campaign_name', 'objective', 'attribution_setting', 'date_stop',
  'quality_ranking', 'engagement_rate_ranking', 'conversion_rate_ranking',
  'optimization_goal',
  'frequency', 'cpm', 'reach', 'impressions',
  'actions', 'action_values', 'cost_per_action_type',
  'spend', 'ctr', 'clicks', 'purchase_roas', 'inline_link_clicks',
  'video_p25_watched_actions', 'video_p50_watched_actions', 'video_p75_watched_actions',
  'video_p95_watched_actions', 'video_p100_watched_actions', 'video_thruplay_watched_actions'
].join(',');

function fetchOptions_() {
  return { headers: { Authorization: 'Bearer ' + ACCESS_TOKEN }, muteHttpExceptions: true, method: 'get' };
}

// ── HEADERS for test-fb / "data - by day" ──
// FIXED (Aug 2026, per your go-ahead to touch campaign-level bugs): unlike the Ad Set
// sheet (ensureAdSetHeaders_) and every breakdown sheet (ensureBreakdownHeaders_), these
// two campaign sheets never had a headers array or a self-healing header check — row 1
// was assumed to already be typed in by hand from before this pipeline existed. Every
// column added to buildRowFromInsight_() since then (Hook rate, Hold rate, Optimization
// goal, Landing page views, Initiate checkouts, ThruPlays) landed data in columns 30-35
// with nothing above it in row 1. This is the SAME 35-column shape ADSET_HEADERS_
// (AdSetPipeline.gs) uses, minus its leading "Ad set name" column — buildRowFromInsight_
// is the shared row builder for both, so keep these two lists in sync if either changes.
var CAMPAIGN_HEADERS_ = [
  'Campaign name', 'Objective', 'Attribution setting', 'Campaign bid strategy',
  'Reporting ends', 'Daily budget', 'Engagement rate ranking', 'Quality ranking', 'Frequency',
  'CPM (cost per 1,000 impressions)', 'Reach', 'Results', 'Blocked messaging contacts', 'Amount spent',
  'Video plays', 'Post engagements', 'CTR (all)', 'CTR (link click-through rate)', '3-second video plays',
  'Video plays at 25%', 'Video plays at 50%', 'Video plays at 75%', 'Video plays at 95%', 'Video plays at 100%',
  'Cost per messaging conversation started', 'Purchase ROAS (return on ad spend)', 'Purchases conversion value',
  'Purchases', 'Conversion rate ranking', 'Hook rate', 'Hold rate',
  'Optimization goal', 'Landing page views', 'Initiate checkouts', 'ThruPlays'
];

// Same "generalized so any past header state auto-heals" pattern as
// ensureBreakdownHeaders_()/ensureAdSetHeaders_(): blank row 1, or a row 1 whose last
// expected header doesn't match, gets (re)written. Never touches data rows below it.
function ensureCampaignHeaders_(sheet) {
  ensureColumns_(sheet, CAMPAIGN_HEADERS_.length);
  var first = sheet.getRange(1, 1, 1, CAMPAIGN_HEADERS_.length).getValues()[0];
  var blank = first.every(function (c) { return String(c).trim() === ''; });
  var lastExpected = CAMPAIGN_HEADERS_[CAMPAIGN_HEADERS_.length - 1];
  var matches = !blank && String(first[CAMPAIGN_HEADERS_.length - 1] || '').trim() === lastExpected;
  if (blank || !matches) sheet.getRange(1, 1, 1, CAMPAIGN_HEADERS_.length).setValues([CAMPAIGN_HEADERS_]);
}

// ── DAILY IMPORT (trigger: ~midnight–1am) → writes YESTERDAY to test-fb ──
function importDataFromJSON() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('test-fb');
  if (!sheet) { Logger.log('test-fb not found'); return; }
  ensureCampaignHeaders_(sheet);

  var until = getYesterday();
  var since = getDateNDaysAgo_(LOOKBACK_DAYS);   // rolling window: last LOOKBACK_DAYS ending yesterday
  var campaignSettings = fetchCampaignSettings();

  var params = {
    level: 'campaign',
    time_increment: 1,             // one row per campaign PER DAY across the window (required for multi-day)
    include_archived: true,
    fields: INSIGHT_FIELDS_,
    time_range: JSON.stringify({ since: since, until: until }),
    limit: 500
  };
  var qs = Object.keys(params).map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(String(params[k])); }).join('&');
  var base = 'https://graph.facebook.com/v18.0/' + ACCOUNT_ID + '/insights';

  // fetch (paginated — a multi-day window can exceed one page)
  var rows = [], nextPage = null;
  do {
    var url = nextPage ? nextPage : (base + '?' + qs);
    try {
      var resp = UrlFetchApp.fetch(url, fetchOptions_());
      var data = JSON.parse(resp.getContentText());
      if (data.error) { Logger.log('importDataFromJSON API error: ' + JSON.stringify(data.error)); return; }
      (data.data || []).forEach(function (item) { rows.push(buildRowFromInsight_(item, campaignSettings)); });
      nextPage = (data.paging && data.paging.next) ? data.paging.next : null;
      Utilities.sleep(150);
    } catch (e) {
      Logger.log('importDataFromJSON fetch error: ' + e.message);
      return;                       // bail without touching data - by day; next run re-pulls the window
    }
  } while (nextPage);

  // overwrite test-fb with the fresh window (35 cols now)
  ensureColumns_(sheet, 35);
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 35).clearContent();
  if (rows.length) {
    var need = 1 + rows.length;
    if (need > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), need - sheet.getMaxRows());
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    Logger.log('importDataFromJSON: wrote ' + rows.length + ' rows for ' + since + '..' + until);
  } else {
    Logger.log('importDataFromJSON: no data for ' + since + '..' + until);
  }

  // Guaranteed order + UPSERT: hand the window to the copy, in the same execution.
  SpreadsheetApp.flush();
  copyDataOnceADay(since, until);
}

// ── ONE-OFF BACKFILL for a date window → writes straight to "data - by day" ──
// Run this once from the editor to fix the Aug 5–11 gap:  backfillAug5to11()
function backfillAug5to11() { backfillDateRange('2026-08-05', '2026-08-11'); }

// ── ONE-OFF: re-backfill ALL history so "Results", Optimization goal, Landing Page
// Views, Initiate Checkouts and ThruPlays are correct/populated back to Nov 13, 2025.
// Safe to re-run — backfillDateRange() clears each window before writing (idempotent).
// Runs month-by-month with a pause between calls to be kind to the rate limit.
function backfillAllHistory_CampaignLevel() {
  var start = new Date('2025-11-13');
  var end = new Date(getYesterday());
  var cur = new Date(start.getFullYear(), start.getMonth(), 1);
  var months = 0;
  while (cur <= end) {
    var mStart = new Date(Math.max(new Date(cur.getFullYear(), cur.getMonth(), 1).getTime(), start.getTime()));
    var mEnd = new Date(Math.min(new Date(cur.getFullYear(), cur.getMonth() + 1, 0).getTime(), end.getTime()));
    backfillDateRange(dayKey_(mStart), dayKey_(mEnd));
    months++;
    if (cur < end) Utilities.sleep(5000);
    cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
  }
  Logger.log('backfillAllHistory_CampaignLevel: done ' + months + ' months, 2025-11-13 .. ' + dayKey_(end));
}

function backfillDateRange(sinceStr, untilStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dest = ss.getSheetByName('data - by day');
  if (!dest) { Logger.log('"data - by day" not found — run backfillAllHistory_CampaignLevel once to create it.'); return; }
  ensureCampaignHeaders_(dest);

  // 1) idempotency: remove any existing rows already in [since, until]
  removeRowsInDateRange_(dest, sinceStr, untilStr);

  // 2) fetch (daily granularity, paginated) and append
  var settings = fetchCampaignSettings();
  var baseUrl = 'https://graph.facebook.com/v18.0/' + ACCOUNT_ID + '/insights';
  var params = {
    level: 'campaign',
    time_increment: 1,
    include_archived: true,
    fields: INSIGHT_FIELDS_,
    time_range: JSON.stringify({ since: sinceStr, until: untilStr }),
    limit: 500
  };
  var qs = Object.keys(params).map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(String(params[k])); }).join('&');

  var nextPage = null, allRows = [];
  do {
    var fullUrl = nextPage ? nextPage : (baseUrl + '?' + qs);
    try {
      var resp = UrlFetchApp.fetch(fullUrl, fetchOptions_());
      var data = JSON.parse(resp.getContentText());
      if (data.error) { Logger.log('API error: ' + JSON.stringify(data.error)); break; }
      (data.data || []).forEach(function (item) { allRows.push(buildRowFromInsight_(item, settings)); });
      nextPage = (data.paging && data.paging.next) ? data.paging.next : null;
      Utilities.sleep(200);
    } catch (e) {
      Logger.log('backfillDateRange error: ' + e.message);
      break;
    }
  } while (nextPage);

  if (allRows.length) {
    ensureColumns_(dest, 35);
    dest.getRange(dest.getLastRow() + 1, 1, allRows.length, allRows[0].length).setValues(allRows);
  }
  Logger.log('backfillDateRange ' + sinceStr + '..' + untilStr + ': wrote ' + allRows.length + ' rows.');
}

// ═══════════════════════════════════════════════════════════════════════════
// ── LIGHTWEIGHT CTR (LINK) RECOMPUTE — patches ONLY column 18, no full re-backfill ──
// backfillDateRange()/backfillAllHistory_CampaignLevel() re-fetch and rewrite all 35
// columns per row — safe, but heavier than it needs to be for a fix that only touches
// one column. This fetches just 4 fields (campaign_name, date_stop, impressions,
// inline_link_clicks) per row instead of 25, and overwrites ONLY the "CTR (link
// click-through rate)" cell for rows that already exist — every other column
// (spend, results, video metrics, etc.) is left completely untouched. Use this instead
// of the full backfill when the only thing you need corrected is CTR (link).
// ═══════════════════════════════════════════════════════════════════════════

var CTR_LINK_COL_CAMPAIGN_ = 18; // 1-based column of "CTR (link click-through rate)" in "data - by day"

function recomputeCtrLinkOnly_CampaignLevel(sinceStr, untilStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dest = ss.getSheetByName('data - by day');
  if (!dest) { Logger.log('"data - by day" not found.'); return; }
  var lastRow = dest.getLastRow();
  if (lastRow < 2) { Logger.log('recomputeCtrLinkOnly_CampaignLevel: no data rows.'); return; }

  var lastCol = dest.getLastColumn();
  var all = dest.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var sinceTs = new Date(sinceStr).getTime();
  var untilTs = new Date(untilStr).getTime() + 86399999;

  // Snapshot the whole CTR(link) column so untouched rows keep their existing value,
  // and build a campaign+date -> row-index lookup restricted to the window.
  var ctrLinkVals = all.map(function (r) { return [r[CTR_LINK_COL_CAMPAIGN_ - 1]]; });
  var rowIndexByKey = {};
  all.forEach(function (r, i) {
    var dk = dayKey_(r[4]); // col 5 = Reporting ends
    var ts = new Date(dk).getTime();
    if (isNaN(ts) || ts < sinceTs || ts > untilTs) return;
    rowIndexByKey[String(r[0]).trim() + '||' + dk] = i;
  });

  var params = {
    level: 'campaign',
    time_increment: 1,
    include_archived: true,
    fields: 'campaign_name,date_stop,impressions,inline_link_clicks', // minimal — just what CTR(link) needs
    time_range: JSON.stringify({ since: sinceStr, until: untilStr }),
    limit: 500
  };
  var qs = Object.keys(params).map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(String(params[k])); }).join('&');
  var base = 'https://graph.facebook.com/v18.0/' + ACCOUNT_ID + '/insights';

  var next = null, touched = 0;
  do {
    var url = next ? next : (base + '?' + qs);
    try {
      var resp = UrlFetchApp.fetch(url, fetchOptions_());
      var data = JSON.parse(resp.getContentText());
      if (data.error) { Logger.log('recomputeCtrLinkOnly_CampaignLevel API error: ' + JSON.stringify(data.error)); break; }
      (data.data || []).forEach(function (item) {
        var key = String(item.campaign_name || '').trim() + '||' + dayKey_(item.date_stop || '');
        var idx = rowIndexByKey[key];
        if (idx === undefined) return; // no stored row for this campaign+date — leave as-is
        var impressions = parseInt(item.impressions) || 0;
        var linkClicks = parseInt(item.inline_link_clicks) || 0;
        ctrLinkVals[idx][0] = impressions > 0 ? (linkClicks / impressions) * 100 : 0;
        touched++;
      });
      next = (data.paging && data.paging.next) ? data.paging.next : null;
      Utilities.sleep(150);
    } catch (e) {
      Logger.log('recomputeCtrLinkOnly_CampaignLevel fetch error: ' + e.message);
      break;
    }
  } while (next);

  // ONE bulk write for the entire column — never touches any other column.
  dest.getRange(2, CTR_LINK_COL_CAMPAIGN_, ctrLinkVals.length, 1).setValues(ctrLinkVals);
  Logger.log('recomputeCtrLinkOnly_CampaignLevel ' + sinceStr + '..' + untilStr + ': updated ' + touched + ' of ' + ctrLinkVals.length + ' rows.');
}

// ═══════════════════════════════════════════════════════════════════════════
// ── RESULTS BACKFILL (Sep 2026) — same lightweight patch-one-column pattern as the
// CTR (link) recompute above, applied to "Results". This is the fix the OPTIMIZATION-
// GOAL-AWARE "RESULTS" comment at the top of this file has been asking for since it was
// written: "you MUST re-run the backfill functions once... so historical rows reflect the
// corrected numbers" — except a FULL backfill re-pulls and rewrites all 35 columns per row
// (heavy, and risks re-tripping the "reduce the amount of data" ceiling on a big window).
// This patches ONLY the "Results" column (12 here), fetching just the 3 fields
// generalizeResults_ actually needs (optimization_goal, actions, thruplays) — every other
// column (spend, video metrics, CTR, etc.) is left completely untouched.
// ═══════════════════════════════════════════════════════════════════════════

var RESULTS_COL_CAMPAIGN_ = 12; // 1-based column of "Results" in "data - by day"

function recomputeResultsOnly_CampaignLevel(sinceStr, untilStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dest = ss.getSheetByName('data - by day');
  if (!dest) { Logger.log('"data - by day" not found.'); return; }
  var lastRow = dest.getLastRow();
  if (lastRow < 2) { Logger.log('recomputeResultsOnly_CampaignLevel: no data rows.'); return; }

  var lastCol = dest.getLastColumn();
  var all = dest.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var sinceTs = new Date(sinceStr).getTime();
  var untilTs = new Date(untilStr).getTime() + 86399999;

  var resultsVals = all.map(function (r) { return [r[RESULTS_COL_CAMPAIGN_ - 1]]; });
  var rowIndexByKey = {};
  all.forEach(function (r, i) {
    var dk = dayKey_(r[4]); // col 5 = Reporting ends
    var ts = new Date(dk).getTime();
    if (isNaN(ts) || ts < sinceTs || ts > untilTs) return;
    rowIndexByKey[String(r[0]).trim() + '||' + dk] = i;
  });

  var params = {
    level: 'campaign',
    time_increment: 1,
    include_archived: true,
    fields: 'campaign_name,date_stop,optimization_goal,actions,video_thruplay_watched_actions', // minimal — just what generalizeResults_ needs
    time_range: JSON.stringify({ since: sinceStr, until: untilStr }),
    limit: 500
  };
  var qs = Object.keys(params).map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(String(params[k])); }).join('&');
  var base = 'https://graph.facebook.com/v18.0/' + ACCOUNT_ID + '/insights';

  var next = null, touched = 0;
  do {
    var url = next ? next : (base + '?' + qs);
    try {
      var resp = UrlFetchApp.fetch(url, fetchOptions_());
      var data = JSON.parse(resp.getContentText());
      if (data.error) { Logger.log('recomputeResultsOnly_CampaignLevel API error: ' + JSON.stringify(data.error)); break; }
      (data.data || []).forEach(function (item) {
        var key = String(item.campaign_name || '').trim() + '||' + dayKey_(item.date_stop || '');
        var idx = rowIndexByKey[key];
        if (idx === undefined) return; // no stored row for this campaign+date — leave as-is
        var purchases = av_(item.actions, 'omni_purchase') + av_(item.actions, 'purchase');
        var thruplays = av_(item.video_thruplay_watched_actions, 'video_view');
        resultsVals[idx][0] = generalizeResults_(item, item.optimization_goal || '', purchases, thruplays);
        touched++;
      });
      next = (data.paging && data.paging.next) ? data.paging.next : null;
      Utilities.sleep(150);
    } catch (e) {
      Logger.log('recomputeResultsOnly_CampaignLevel fetch error: ' + e.message);
      break;
    }
  } while (next);

  // ONE bulk write for the entire column — never touches any other column.
  dest.getRange(2, RESULTS_COL_CAMPAIGN_, resultsVals.length, 1).setValues(resultsVals);
  Logger.log('recomputeResultsOnly_CampaignLevel ' + sinceStr + '..' + untilStr + ': updated ' + touched + ' of ' + resultsVals.length + ' rows.');
}

// ── ONE-OFF: run this once to fix "Results" across your whole history (2025-11-13 ->
// yesterday) now that it's optimization-goal-aware instead of purchases-only. Month-by-
// month, same idempotent-per-window pattern as everything else here — safe to re-run, safe
// if it times out partway (already-patched months just get harmlessly redone).
function recomputeResultsHistory_CampaignLevel() {
  var start = new Date('2025-11-13');
  var end = new Date(getYesterday());
  var cur = new Date(start.getFullYear(), start.getMonth(), 1);
  var months = 0;
  while (cur <= end) {
    var mStart = new Date(Math.max(new Date(cur.getFullYear(), cur.getMonth(), 1).getTime(), start.getTime()));
    var mEnd = new Date(Math.min(new Date(cur.getFullYear(), cur.getMonth() + 1, 0).getTime(), end.getTime()));
    recomputeResultsOnly_CampaignLevel(dayKey_(mStart), dayKey_(mEnd));
    months++;
    if (cur < end) Utilities.sleep(3000);
    cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
  }
  Logger.log('recomputeResultsHistory_CampaignLevel: done ' + months + ' months, 2025-11-13 .. ' + dayKey_(end));
}

// ── ONE-OFF: run this once instead of backfillAllHistory_CampaignLevel() if CTR (link)
// is the only thing you need corrected. Month-by-month, same idempotent-per-window
// pattern as everything else here — safe to re-run, safe if it times out partway
// (already-patched months just get harmlessly redone).
function recomputeCtrLinkHistory_CampaignLevel() {
  var start = new Date('2025-11-13');
  var end = new Date(getYesterday());
  var cur = new Date(start.getFullYear(), start.getMonth(), 1);
  var months = 0;
  while (cur <= end) {
    var mStart = new Date(Math.max(new Date(cur.getFullYear(), cur.getMonth(), 1).getTime(), start.getTime()));
    var mEnd = new Date(Math.min(new Date(cur.getFullYear(), cur.getMonth() + 1, 0).getTime(), end.getTime()));
    recomputeCtrLinkOnly_CampaignLevel(dayKey_(mStart), dayKey_(mEnd));
    months++;
    if (cur < end) Utilities.sleep(3000);
    cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
  }
  Logger.log('recomputeCtrLinkHistory_CampaignLevel: done ' + months + ' months, 2025-11-13 .. ' + dayKey_(end));
}

// ── COPY / UPSERT → refreshes a date window in "data - by day" from test-fb ──
// Called (chained) by importDataFromJSON with the rolling window. It clears that
// window in "data - by day" and re-inserts the fresh pull, so recent days get
// corrected (late attribution) and any missed day self-heals. Idempotent.
function copyDataOnceADay(sinceStr, untilStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src = ss.getSheetByName('test-fb');
  var dest = ss.getSheetByName('data - by day');
  if (!src || !dest) { Logger.log('copyDataOnceADay: a sheet is missing'); return; }

  var srcLast = src.getLastRow();
  if (srcLast < 2) { Logger.log('copyDataOnceADay: test-fb empty'); return; }

  ensureCampaignHeaders_(dest);
  ensureColumns_(dest, 35);
  var data = src.getRange(2, 1, srcLast - 1, 35).getValues().filter(function (r) {
    return r.join('').trim() !== '';
  });
  if (!data.length) { Logger.log('copyDataOnceADay: nothing to copy'); return; }

  // If no window was passed, derive it from the dates present in test-fb.
  if (!sinceStr || !untilStr) {
    var minTs = Infinity, maxTs = -Infinity;
    data.forEach(function (r) {
      var ts = new Date(dayKey_(r[4])).getTime();
      if (!isNaN(ts)) { if (ts < minTs) minTs = ts; if (ts > maxTs) maxTs = ts; }
    });
    if (minTs !== Infinity) { sinceStr = dayKey_(new Date(minTs)); untilStr = dayKey_(new Date(maxTs)); }
  }

  // UPSERT: clear the window, then append the fresh rows (no duplicates possible).
  if (sinceStr && untilStr) removeRowsInDateRange_(dest, sinceStr, untilStr);

  var start = dest.getLastRow() + 1;
  var need = start + data.length - 1;
  if (need > dest.getMaxRows()) dest.insertRowsAfter(dest.getMaxRows(), need - dest.getMaxRows()); // auto-grow grid
  dest.getRange(start, 1, data.length, data[0].length).setValues(data);

  Logger.log('copyDataOnceADay: upserted ' + data.length + ' rows for ' + sinceStr + '..' + untilStr);
  // Sep 2026, "let users know if dashboard data is fresh or stale" — stamped ONLY here, at
  // the exact point "data - by day" (what the dashboard actually reads) was genuinely
  // rewritten — not in importDataFromJSON() itself, which can succeed at the fetch step and
  // still fail to reach this point (e.g. "a sheet is missing"). markPipelineRefreshed_ is a
  // shared helper (defined in AdSetPipeline.gs alongside the other fail-streak helpers).
  markPipelineRefreshed_('campaigns');
}

// ── SHARED HELPERS ──
// Timezone-safe: uses the script's timezone (Manila), NOT UTC. This is the only
// getYesterday() the project should have — delete any copies in other files.
function getYesterday() {
  var d = new Date();
  d.setDate(d.getDate() - 1);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getDateNDaysAgo_(n) {
  var d = new Date();
  d.setDate(d.getDate() - n);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// Grow a sheet's column count if it has fewer than `n` columns (setValues can't overflow the grid).
function ensureColumns_(sheet, n) {
  var max = sheet.getMaxColumns();
  if (max < n) sheet.insertColumnsAfter(max, n - max);
}

function keyOf_(row) { return String(row[0]).trim() + '||' + dayKey_(row[4]); } // campaign || date

function dayKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var s = String(v == null ? '' : v).trim();
  var m = s.match(/^(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : s;
}

function removeRowsInDateRange_(sheet, sinceStr, untilStr) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var sinceTs = new Date(sinceStr).getTime();
  var untilTs = new Date(untilStr).getTime() + 86399999;
  var kept = values.filter(function (row) {
    var k = dayKey_(row[4]);        // col 5 = Reporting ends
    var ts = new Date(k).getTime();
    if (isNaN(ts)) return true;     // keep undated rows
    return !(ts >= sinceTs && ts <= untilTs);
  });
  var removed = values.length - kept.length;
  sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  if (kept.length) sheet.getRange(2, 1, kept.length, kept[0].length).setValues(kept);
  Logger.log('removeRowsInDateRange_: cleared ' + removed + ' existing rows in ' + sinceStr + '..' + untilStr);
}

/**
 * BACKFILL GUIDE (do this once, after pasting the updated files in):
 *   1. Paste this file, the updated AdSetPipeline.gs, BreakdownPipeline.gs, Code.gs.
 *   2. In importDataFromJSON.gs: delete backfill12Months(), backfillBreakdownData(),
 *      and the old importDataFromJSON()/copyDataOnceADay() if any copies remain there.
 *   3. Delete CustomReport.gs entirely (superseded — see Code.gs notes).
 *   4. Run backfillAllHistory_CampaignLevel() once from the Apps Script editor.
 *      It re-pulls Nov 13, 2025 → yesterday, month by month, ~5 sec pause between
 *      months. For 9 months this can take a while — if Apps Script's 6-minute
 *      execution limit interrupts it, just run it again: already-filled months are
 *      cleared-and-rewritten (idempotent), so re-running is always safe.
 *   5. Run backfillAdSetHistory() (in AdSetPipeline.gs) the same way.
 *   6. Run backfillAllHistory_Breakdowns() (in BreakdownPipeline.gs) the same way.
 */