/**
 * AdPipeline.gs — pulls AD level insights + creative previews for the new "Ads" tab.
 *
 * NEW (Aug 2026), built per your "open the ad level" request. Decisions you made when
 * this was scoped (so future-you knows why it's built this way, not just how):
 *   1. FULL historical backfill, same start date as everything else: 2025-11-13.
 *      Ad-level rows multiply fast (every ad × every day it ran), so this WILL be your
 *      biggest table by row count. See the workbook-split note below.
 *   2. Creative preview = thumbnail image + a "View in Ads Manager" link, NOT inline
 *      video playback. Meta's thumbnail_url field is reliably returned for image AND
 *      video creatives, but embeddable video playback is not consistently available
 *      across ads/permissions — a thumbnail + link is the version that works every time.
 *   3. "Thumb Stop Rate" from the benchmark PDF is DROPPED from this version. It's not a
 *      field Meta's API exposes and the PDF doesn't define its formula, so anything I
 *      built for it would be a guess wearing your benchmark's name. Everything else in
 *      the PDF's Creative Performance Benchmark table (Hook Rate, Hold Rate, CTR Link)
 *      IS implemented.
 *
 * ══════════════════════════════════════════════════════════════════════════
 * WORKBOOK: Ad-level data lives in its OWN spreadsheet — a NEW workbook you need to
 * create (do NOT point this at the Campaign or Ad Set workbook; see AD_SHEET_ID below).
 * Same reasoning as the Ad Set split: this keeps any one workbook under Sheets' 10M-cell
 * limit. Ad-level daily rows are the largest of the three levels, so this one matters
 * most — don't skip creating a dedicated file for it.
 *
 * SETUP (do this once):
 *   1. Create a new blank Google Sheet. Name it something like "[COMPANY_NAME] — Ads Data".
 *   2. Copy its ID from the URL (the long string between /d/ and /edit).
 *   3. Paste that ID into AD_SHEET_ID below, replacing 'YOUR_AD_SPREADSHEET_ID'.
 *   4. Extensions → Apps Script → paste THIS file in (plus keep everything else you
 *      already have — this file reuses globals from DataPipeline.gs / AdSetPipeline.gs /
 *      importDataFromJSON.gs: ACCOUNT_ID, ACCESS_TOKEN, fetchCampaignSettings(),
 *      fetchWithRetry_, fetchOptions_, INSIGHT_FIELDS_, buildRowFromInsight_, dayKey_,
 *      ensureColumns_, getYesterday, getDateNDaysAgo_, LOOKBACK_DAYS — all must already
 *      be in the same Apps Script project as this file for it to run).
 *   5. Run diagnoseAd() once to confirm it resolves to the RIGHT workbook.
 *   6. Run backfillAdHistory() once — this is the big one. It goes month-by-month from
 *      2025-11-13, sleeping between months to be kind to the rate limit. For 9+ months
 *      of AD-level data this can take a long time and will very likely hit Apps Script's
 *      6-minute execution limit partway through — that's expected, not a bug. Just run
 *      it again: already-filled months are cleared-and-rewritten (idempotent), so
 *      re-running picks up where it left off with no duplicates. If a single month
 *      itself times out (very large ad count), call backfillAdData(sinceStr, untilStr)
 *      directly with half-month windows instead.
 *   7. Run importAdCreatives() once (and then daily) to populate the creative preview
 *      snapshot (thumbnail + type) that the Ads tab reads.
 *   8. Add DAILY triggers (~1–2am, after the Campaign/Ad Set imports) for:
 *        importAdData
 *        importAdCreatives
 * ══════════════════════════════════════════════════════════════════════════
 */

var AD_SHEET_ID = '[AD_SHEET_ID]'; // ← create a NEW Google Sheet for Ads and paste its ID here. See Config.example.gs.
var AD_SHEET_ID_PROP_ = 'AD_SHEET_ID_OVERRIDE'; // set by a future createNewYearAdWorkbook(), read here first

function getAdSpreadsheet_() {
  var override = PropertiesService.getScriptProperties().getProperty(AD_SHEET_ID_PROP_);
  var id = override || AD_SHEET_ID;
  if (!id || id === 'YOUR_AD_SPREADSHEET_ID') {
    throw new Error('AD_SHEET_ID is not set in AdPipeline.gs. Create a new Google Sheet for Ad-level data and put its ID there — see the SETUP comment at the top of this file.');
  }
  var ss = SpreadsheetApp.openById(id);
  if (!ss) throw new Error('Could not open the Ads workbook (id ' + id + '). Check the ID and that this account can access it.');
  return ss;
}

/**
 * DIAGNOSTIC — run this straight from the editor to confirm the ad-level side resolves
 * to the RIGHT workbook (mirrors diagnoseAdSet() in AdSetPipeline.gs).
 */
function diagnoseAd() {
  var out = [];
  var override = PropertiesService.getScriptProperties().getProperty(AD_SHEET_ID_PROP_);
  out.push('AD_SHEET_ID constant: ' + AD_SHEET_ID);
  out.push('Script Properties override (' + AD_SHEET_ID_PROP_ + '): ' + (override || '(none set — using the constant)'));
  try {
    var ss = getAdSpreadsheet_();
    out.push('Resolved workbook: OK -> "' + ss.getName() + '"  [id ' + ss.getId() + ']');
    var tabs = ss.getSheets().map(function (s) { return '"' + s.getName() + '"'; });
    out.push('Tabs in that file: ' + tabs.join(', '));
  } catch (e) {
    out.push('FAILED to open ad workbook: ' + e.message);
  }
  var msg = out.join('\n');
  Logger.log(msg);
  return msg;
}

// 38 columns: [Ad ID, Ad name, Ad set name] + the same 35-column shape CAMPAIGN_HEADERS_
// (DataPipeline.gs) / ADSET_HEADERS_ (AdSetPipeline.gs) use, starting at "Campaign name".
// FUNCTION, not a top-level var: Apps Script does not guarantee that DataPipeline.gs's
// top-level `var CAMPAIGN_HEADERS_ = [...]` has already run before this file's top-level
// code does (files aren't guaranteed to initialize in any particular order — alphabetically
// "AdPipeline.gs" actually sorts BEFORE "DataPipeline.gs"). Computing this lazily inside a
// function sidesteps that entirely, since by the time any function in this project actually
// RUNS, every file's top-level vars are already initialized.
function getAdHeaders_() {
  return ['Ad ID', 'Ad name', 'Ad set name'].concat(CAMPAIGN_HEADERS_);
}

// ── DAILY (add a ~1–2am trigger): rolling window upsert ──
// Same self-healing, fail-streak/alert pattern as importAdSetData() in AdSetPipeline.gs —
// a failed fetch NEVER clears existing rows, so a bad day self-heals within LOOKBACK_DAYS
// once the underlying issue clears, and you get an email if it doesn't.
var AD_MAIN_FAIL_STREAK_PROP_ = 'AD_MAIN_IMPORT_FAIL_STREAK';

// See the matching comment on importAdSetData() (AdSetPipeline.gs) — this trigger's
// error rate can't come from the graceful result.ok path below (that never throws), so it
// has to be an uncaught exception this function had no try/catch around. Wrapped now so
// the next one actually surfaces instead of just incrementing a percentage on the
// Triggers page with no explanation.
function importAdData() {
  try {
    var sheet = getAdSheet_();
    ensureAdHeaders_(sheet);
    var until = getYesterday();
    var since = getDateNDaysAgo_(LOOKBACK_DAYS);
    var result = fetchAdRows_(since, until);

    if (!result.ok) {
      var streak = bumpFailStreak_(AD_MAIN_FAIL_STREAK_PROP_);
      Logger.log('importAdData: fetch failed (' + result.errorMessage + ') for ' + since + '..' + until +
        ' — leaving existing rows untouched. Fail streak: ' + streak + ' day(s). Retries automatically tomorrow.');
      if (streak >= LOOKBACK_DAYS) alertAdImportFailing_(streak, result.errorMessage);
      return;
    }
    resetFailStreak_(AD_MAIN_FAIL_STREAK_PROP_);

    removeAdRowsInRange_(sheet, since, until);
    appendAdRowsSafe_(sheet, result.rows);
    Logger.log('importAdData: ' + result.rows.length + ' rows for ' + since + '..' + until);
    markPipelineRefreshed_('ads');
  } catch (e) {
    Logger.log('⚠️ importAdData: UNCAUGHT exception — ' + (e && e.message ? e.message : String(e)) +
      '. Existing rows are untouched (the write only happens after the try block above completes).');
    try {
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Meta Ads Dashboard: importAdData crashed',
        'importAdData() threw an uncaught exception: ' + (e && e.message ? e.message : String(e)) +
        '\n\nCheck View > Executions > importAdData in the Apps Script editor for the full stack trace.');
    } catch (mailErr) {
      Logger.log('(Could not send the crash email — Logger above still has it. Reason: ' + mailErr.message + ')');
    }
  }
}

function alertAdImportFailing_(streak, errorMessage) {
  var msg = 'importAdData() has failed ' + streak + ' day(s) in a row (latest error: ' + errorMessage + '). ' +
    'The rolling ' + LOOKBACK_DAYS + '-day window means it can no longer self-heal the oldest affected dates — ' +
    'they will start permanently dropping out of "data - ad - by day" unless you fix the underlying issue and ' +
    'then run a manual backfill, e.g. backfillAdData(sinceStr, untilStr) for the affected range. Check ' +
    'View > Executions > importAdData in the Apps Script editor for the exact API error.';
  Logger.log('⚠️ ' + msg);
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Meta Ads Dashboard: Ad data import failing', msg);
  } catch (e) {
    Logger.log('(Could not send the email warning — Logger above still has it. Reason: ' + e.message + ')');
  }
}

// ── ONE-OFF BACKFILL (run once, per your call to pull the full history) ──
function backfillAdHistory() {
  var sheet = getAdSheet_();
  ensureAdHeaders_(sheet);
  var start = new Date('2025-11-13');
  var end = new Date(getYesterday());
  var cur = new Date(start.getFullYear(), start.getMonth(), 1);
  var monthCount = 0;
  while (cur <= end) {
    var mStart = new Date(Math.max(new Date(cur.getFullYear(), cur.getMonth(), 1).getTime(), start.getTime()));
    var mEnd = new Date(Math.min(new Date(cur.getFullYear(), cur.getMonth() + 1, 0).getTime(), end.getTime()));
    backfillAdData(dayKey_(mStart), dayKey_(mEnd));
    monthCount++;
    if (cur < end) Utilities.sleep(5000);
    cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
  }
  Logger.log('backfillAdHistory: done ' + monthCount + ' months from 2025-11-13 .. ' + dayKey_(end));
}

// Same guard as importAdData(): only clear+write this window if the fetch succeeded.
// FIXED (Aug 2026): a level='ad' + time_increment=1 request over a full month is the
// heaviest query this project makes — every ad, every day, every field, all in one
// response. Meta rejects requests like that outright with a "please reduce the amount
// of data you're asking for" family of errors (comes back as code 2, same top-level code
// as a genuine transient hiccup, but a DIFFERENT thing: it fails the SAME way on every
// one of fetchWithRetry_'s 5 attempts, because retrying identical parameters can't fix a
// request that's just too big). Rather than giving up on the whole month (which used to
// just log and leave it un-backfilled until you manually picked a smaller window and
// re-ran), this now auto-splits the range in half and retries each half on its own — all
// the way down to single days if it has to — so one call site (backfillAdData /
// backfillAdHistory) self-heals without you needing to guess a chunk size by hand.
function backfillAdData(sinceStr, untilStr) {
  var sheet = getAdSheet_();
  ensureAdHeaders_(sheet);
  backfillAdRange_(sheet, sinceStr, untilStr);
}

function backfillAdRange_(sheet, sinceStr, untilStr) {
  var result = fetchAdRows_(sinceStr, untilStr);
  if (result.ok) {
    removeAdRowsInRange_(sheet, sinceStr, untilStr);
    appendAdRowsSafe_(sheet, result.rows);
    Logger.log('backfillAdData ' + sinceStr + '..' + untilStr + ': ' + result.rows.length + ' rows');
    return;
  }

  var since = new Date(sinceStr), until = new Date(untilStr);
  var spanDays = Math.round((until.getTime() - since.getTime()) / 86400000) + 1;
  if (spanDays <= 1) {
    // Can't split a single day any further — this is either a genuine outage or one
    // unusually heavy day (huge ad count). Existing rows are left untouched either way.
    Logger.log('backfillAdData ' + sinceStr + '..' + untilStr + ': still failing at 1-day granularity (' +
      result.errorMessage + ') — existing rows left untouched. Check View > Executions for the exact ' +
      'error text; if it\'s the same "reduce the amount of data" family, this single day itself has more ' +
      'ad-rows than Meta will return in one response (rare) — otherwise it\'s a real outage, re-run later.');
    return;
  }

  var midOffset = Math.floor((spanDays - 1) / 2);
  var mid = new Date(since.getTime() + midOffset * 86400000);
  var firstEnd = dayKey_(mid);
  var secondStart = dayKey_(new Date(mid.getTime() + 86400000));
  Logger.log('backfillAdData ' + sinceStr + '..' + untilStr + ': failed as one request (' + result.errorMessage +
    ') — very likely Meta\'s "reduce the amount of data" limit, not a rate-limit hiccup (identical retries ' +
    'don\'t fix it). Splitting into ' + sinceStr + '..' + firstEnd + ' and ' + secondStart + '..' + untilStr + '.');
  Utilities.sleep(1000);
  backfillAdRange_(sheet, sinceStr, firstEnd);
  Utilities.sleep(1000);
  backfillAdRange_(sheet, secondStart, untilStr);
}

// ── FETCH (level=ad, daily, paginated) with retry ──
// Returns { rows, ok, errorMessage } — same reasoning as fetchAdSetRows_ in
// AdSetPipeline.gs: callers must be able to tell "genuinely no data" from "the fetch
// broke partway", so a bad day never silently wipes good existing rows.
function fetchAdRows_(sinceStr, untilStr) {
  var settings = fetchCampaignSettings(); // reused global (importDataFromJSON.gs) — campaign bid strategy/daily budget, keyed by campaign name
  var params = {
    level: 'ad',
    time_increment: 1,
    include_archived: true,
    fields: 'ad_id,ad_name,adset_name,' + INSIGHT_FIELDS_,
    time_range: JSON.stringify({ since: sinceStr, until: untilStr }),
    limit: 500
  };
  var qs = Object.keys(params).map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(String(params[k])); }).join('&');
  var base = 'https://graph.facebook.com/v18.0/' + ACCOUNT_ID + '/insights';

  var next = null, out = [], errorMessage = '';
  do {
    var url = next ? next : base + '?' + qs;
    var data = fetchWithRetry_(url);
    if (!data || data.error) {
      errorMessage = (data && data.error) ? JSON.stringify(data.error) : 'no response from fetchWithRetry_';
      Logger.log('ad insights error: ' + errorMessage);
      return { rows: out, ok: false, errorMessage: errorMessage };
    }
    (data.data || []).forEach(function (item) { out.push(buildAdRow_(item, settings)); });
    next = (data.paging && data.paging.next) ? data.paging.next : null;
    if (next) Utilities.sleep(500);
  } while (next);
  return { rows: out, ok: true, errorMessage: '' };
}

// One 38-col row: [Ad ID, Ad name, Ad set name] + the same 35 columns the campaign/ad-set
// builders produce (buildRowFromInsight_ shared from DataPipeline.gs).
function buildAdRow_(item, campaignSettings) {
  var base = buildRowFromInsight_(item, campaignSettings); // 35 cols, campaign bid strategy/budget filled via item.campaign_name
  return [item.ad_id || '', item.ad_name || '', item.adset_name || ''].concat(base);
}

// ═══════════════════════════════════════════════════════════════════════════
// ── LIGHTWEIGHT CTR (LINK) RECOMPUTE — same pattern as the Campaign/Ad Set versions ──
// Not needed right now (this pipeline is new, so its CTR(link) is correct from day one),
// but included proactively so if Meta ever changes/you ever need to patch just this one
// column here too, you don't have to re-derive this pattern from scratch.
// ═══════════════════════════════════════════════════════════════════════════

var CTR_LINK_COL_AD_ = 21; // 1-based: 3 prepended cols (Ad ID, Ad name, Ad set name) + col 18 within the 35-col shape

function recomputeCtrLinkOnly_AdLevel(sinceStr, untilStr) {
  var sheet = getAdSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('recomputeCtrLinkOnly_AdLevel: no data rows.'); return; }
  var lastCol = sheet.getLastColumn();
  var all = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var sinceTs = new Date(sinceStr).getTime();
  var untilTs = new Date(untilStr).getTime() + 86399999;

  var ctrLinkVals = all.map(function (r) { return [r[CTR_LINK_COL_AD_ - 1]]; });
  var rowIndexByKey = {}; // ad_id + date is a genuinely unique key, unlike names
  all.forEach(function (r, i) {
    var dk = dayKey_(r[7]); // col 8 = Reporting ends (Ad ID, Ad name, Ad set name, Campaign name, Objective, Attribution, Bid strategy, Reporting ends)
    var ts = new Date(dk).getTime();
    if (isNaN(ts) || ts < sinceTs || ts > untilTs) return;
    rowIndexByKey[String(r[0]).trim() + '||' + dk] = i;
  });

  var params = {
    level: 'ad',
    time_increment: 1,
    include_archived: true,
    fields: 'ad_id,date_stop,impressions,inline_link_clicks',
    time_range: JSON.stringify({ since: sinceStr, until: untilStr }),
    limit: 500
  };
  var qs = Object.keys(params).map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(String(params[k])); }).join('&');
  var base = 'https://graph.facebook.com/v18.0/' + ACCOUNT_ID + '/insights';

  var next = null, touched = 0;
  do {
    var url = next ? next : (base + '?' + qs);
    var data = fetchWithRetry_(url);
    if (!data || data.error) {
      if (data && data.error) Logger.log('recomputeCtrLinkOnly_AdLevel error: ' + JSON.stringify(data.error));
      break;
    }
    (data.data || []).forEach(function (item) {
      var key = String(item.ad_id || '').trim() + '||' + dayKey_(item.date_stop || '');
      var idx = rowIndexByKey[key];
      if (idx === undefined) return;
      var impressions = parseInt(item.impressions) || 0;
      var linkClicks = parseInt(item.inline_link_clicks) || 0;
      ctrLinkVals[idx][0] = impressions > 0 ? (linkClicks / impressions) * 100 : 0;
      touched++;
    });
    next = (data.paging && data.paging.next) ? data.paging.next : null;
    if (next) Utilities.sleep(300);
  } while (next);

  sheet.getRange(2, CTR_LINK_COL_AD_, ctrLinkVals.length, 1).setValues(ctrLinkVals);
  Logger.log('recomputeCtrLinkOnly_AdLevel ' + sinceStr + '..' + untilStr + ': updated ' + touched + ' of ' + ctrLinkVals.length + ' rows.');
}

function recomputeCtrLinkHistory_AdLevel() {
  var start = new Date('2025-11-13');
  var end = new Date(getYesterday());
  var cur = new Date(start.getFullYear(), start.getMonth(), 1);
  var months = 0;
  while (cur <= end) {
    var mStart = new Date(Math.max(new Date(cur.getFullYear(), cur.getMonth(), 1).getTime(), start.getTime()));
    var mEnd = new Date(Math.min(new Date(cur.getFullYear(), cur.getMonth() + 1, 0).getTime(), end.getTime()));
    recomputeCtrLinkOnly_AdLevel(dayKey_(mStart), dayKey_(mEnd));
    months++;
    if (cur < end) Utilities.sleep(3000);
    cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
  }
  Logger.log('recomputeCtrLinkHistory_AdLevel: done ' + months + ' months, 2025-11-13 .. ' + dayKey_(end));
}

// ═══════════════════════════════════════════════════════════════════════════
// ── LIGHTWEIGHT RESULTS RECOMPUTE — patches ONLY the "Results" column ──
// Same pattern as recomputeCtrLinkOnly_AdLevel() above and
// recomputeResultsOnly_CampaignLevel() in DataPipeline.gs / recomputeResultsOnly_AdSetLevel()
// in AdSetPipeline.gs. "Results" became optimization-goal-aware (see DataPipeline.gs header
// comment) — run this once so historical rows reflect the corrected numbers instead of the
// old purchases-only figure.
// ═══════════════════════════════════════════════════════════════════════════

var RESULTS_COL_AD_ = 15; // 1-based: 3 prepended cols (Ad ID, Ad name, Ad set name) + col 12 within the 35-col shape

function recomputeResultsOnly_AdLevel(sinceStr, untilStr) {
  var sheet = getAdSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('recomputeResultsOnly_AdLevel: no data rows.'); return; }
  var lastCol = sheet.getLastColumn();
  var all = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var sinceTs = new Date(sinceStr).getTime();
  var untilTs = new Date(untilStr).getTime() + 86399999;

  var resultsVals = all.map(function (r) { return [r[RESULTS_COL_AD_ - 1]]; });
  var rowIndexByKey = {}; // ad_id + date is a genuinely unique key, unlike names
  all.forEach(function (r, i) {
    var dk = dayKey_(r[7]); // col 8 = Reporting ends
    var ts = new Date(dk).getTime();
    if (isNaN(ts) || ts < sinceTs || ts > untilTs) return;
    rowIndexByKey[String(r[0]).trim() + '||' + dk] = i;
  });

  var params = {
    level: 'ad',
    time_increment: 1,
    include_archived: true,
    fields: 'ad_id,date_stop,optimization_goal,actions,video_thruplay_watched_actions',
    time_range: JSON.stringify({ since: sinceStr, until: untilStr }),
    limit: 500
  };
  var qs = Object.keys(params).map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(String(params[k])); }).join('&');
  var base = 'https://graph.facebook.com/v18.0/' + ACCOUNT_ID + '/insights';

  var next = null, touched = 0;
  do {
    var url = next ? next : (base + '?' + qs);
    var data = fetchWithRetry_(url);
    if (!data || data.error) {
      if (data && data.error) Logger.log('recomputeResultsOnly_AdLevel error: ' + JSON.stringify(data.error));
      break;
    }
    (data.data || []).forEach(function (item) {
      var key = String(item.ad_id || '').trim() + '||' + dayKey_(item.date_stop || '');
      var idx = rowIndexByKey[key];
      if (idx === undefined) return;
      var purchases = av_(item.actions, 'omni_purchase') + av_(item.actions, 'purchase');
      var thruplays = av_(item.video_thruplay_watched_actions, 'video_view');
      resultsVals[idx][0] = generalizeResults_(item, item.optimization_goal || '', purchases, thruplays);
      touched++;
    });
    next = (data.paging && data.paging.next) ? data.paging.next : null;
    if (next) Utilities.sleep(300);
  } while (next);

  sheet.getRange(2, RESULTS_COL_AD_, resultsVals.length, 1).setValues(resultsVals);
  Logger.log('recomputeResultsOnly_AdLevel ' + sinceStr + '..' + untilStr + ': updated ' + touched + ' of ' + resultsVals.length + ' rows.');
}

function recomputeResultsHistory_AdLevel() {
  var start = new Date('2025-11-13');
  var end = new Date(getYesterday());
  var cur = new Date(start.getFullYear(), start.getMonth(), 1);
  var months = 0;
  while (cur <= end) {
    var mStart = new Date(Math.max(new Date(cur.getFullYear(), cur.getMonth(), 1).getTime(), start.getTime()));
    var mEnd = new Date(Math.min(new Date(cur.getFullYear(), cur.getMonth() + 1, 0).getTime(), end.getTime()));
    recomputeResultsOnly_AdLevel(dayKey_(mStart), dayKey_(mEnd));
    months++;
    if (cur < end) Utilities.sleep(3000);
    cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
  }
  Logger.log('recomputeResultsHistory_AdLevel: done ' + months + ' months, 2025-11-13 .. ' + dayKey_(end));
}

// ── SHEET HELPERS ──
function getAdSheet_() {
  var ss = getAdSpreadsheet_();
  var sh = ss.getSheetByName('data - ad - by day');
  if (!sh) sh = ss.insertSheet('data - ad - by day');
  return sh;
}

function ensureAdHeaders_(sheet) {
  var headers = getAdHeaders_();
  ensureColumns_(sheet, headers.length);
  var first = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var blank = first.every(function (c) { return String(c).trim() === ''; });
  var lastExpected = headers[headers.length - 1];
  var matches = !blank && String(first[headers.length - 1] || '').trim() === lastExpected;
  if (blank || !matches) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

// Date lives in column 8 ("Reporting ends") — index 7 — in the ad-level layout.
function removeAdRowsInRange_(sheet, sinceStr, untilStr) {
  var last = sheet.getLastRow();
  if (last < 2) return;
  var lastCol = sheet.getLastColumn();
  var vals = sheet.getRange(2, 1, last - 1, lastCol).getValues();
  var s = new Date(sinceStr).getTime();
  var u = new Date(untilStr).getTime() + 86399999;
  var kept = vals.filter(function (r) {
    var ts = new Date(dayKey_(r[7])).getTime();
    if (isNaN(ts)) return true;
    return !(ts >= s && ts <= u);
  });
  sheet.getRange(2, 1, last - 1, lastCol).clearContent();
  if (kept.length) sheet.getRange(2, 1, kept.length, kept[0].length).setValues(kept);
}

function appendAdRowsSafe_(sheet, rows) {
  if (!rows || !rows.length) return;
  ensureColumns_(sheet, getAdHeaders_().length);
  var start = sheet.getLastRow() + 1;
  var need = start + rows.length - 1;
  if (need > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), need - sheet.getMaxRows());
  sheet.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
}

// ═══════════════════════════════════════════════════════════════════════════
// ── CREATIVE PREVIEW SNAPSHOT (current state only, like importAdSetStatus) ──
// Meta's Insights API (what everything above pulls from) does NOT return creative
// assets — thumbnails/video/type come from the separate /ads and /adcreatives edges,
// which only expose the CURRENT creative, not a historical log of what an ad's creative
// looked like on a past date. So, same as Learning Phase Status in AdSetPipeline.gs:
// this is a full daily snapshot (old rows for ads that no longer come back are dropped),
// not something you backfill to Nov 13, 2025. Run importAdCreatives() daily; there is
// nothing to backfill here.
// ═══════════════════════════════════════════════════════════════════════════

var AD_CREATIVE_SHEET_ = 'ad_creatives';
// 'Fallback Thumbnail URL' added Sep 2026 (see the "not loading" fix below) — appended at
// the END so existing fixed-index reads elsewhere (getAdCreativeMap_ in Code.gs) don't shift.
var AD_CREATIVE_HEADERS_ = ['Ad ID', 'Ad name', 'Campaign name', 'Ad set name', 'Status', 'Creative type', 'Thumbnail URL', 'Ads Manager link', 'Last updated', 'Fallback Thumbnail URL'];

// Wrapped in try/catch (Sep 2026, trigger error rate investigation) — see the matching
// comment on importAdSetData() (AdSetPipeline.gs) for why: this function's own failure
// paths already degrade gracefully (the `ok` guard below), so an error rate showing up on
// the Triggers page has to be an uncaught exception/platform kill this had no safety net
// around at all. This is also the heaviest of the daily triggers (per-ad thumbnail image
// processing + a native-video-frame lookup on top of pagination), so it's the single most
// likely one to occasionally hit Apps Script's 6-minute execution ceiling — that shows up
// as exactly this kind of silent, unexplained error rate.
function importAdCreatives() {
 try {
  // TIMING INSTRUMENTATION (Sep 2026, after a real "Exceeded maximum execution time" run
  // with NO Meta error at all — a genuine 6-minute platform timeout, not a rate limit).
  // Rather than guess which phase is actually eating the budget (pagination+thumbnail-resize
  // vs. video native-frame lookup vs. image full-res lookup), this logs elapsed time at each
  // checkpoint. Logger.log output is flushed incrementally, so even if this run ALSO times
  // out, View > Executions > importAdCreatives will show every checkpoint reached before the
  // kill — that tells us exactly where the 6 minutes went instead of another guess.
  var __t0 = Date.now();
  var sheet = getOrCreateAdPlainSheet_(AD_CREATIVE_SHEET_);
  ensureColumns_(sheet, AD_CREATIVE_HEADERS_.length);
  var first = sheet.getRange(1, 1, 1, AD_CREATIVE_HEADERS_.length).getValues()[0];
  if (first.every(function (c) { return String(c).trim() === ''; })) {
    sheet.getRange(1, 1, 1, AD_CREATIVE_HEADERS_.length).setValues([AD_CREATIVE_HEADERS_]);
  } else {
    // Sep 2026: backfill header label(s) for any column appended to AD_CREATIVE_HEADERS_
    // AFTER this sheet already existed (e.g. 'Fallback Thumbnail URL' below) — without this,
    // the "brand new sheet" check above never re-triggers on an existing sheet, so a newly
    // added header cell would silently stay blank forever even though the data under it
    // populates fine (rows are written by fixed position, not by header lookup).
    first.forEach(function (c, i) {
      if (String(c).trim() === '' && AD_CREATIVE_HEADERS_[i]) sheet.getRange(1, i + 1).setValue(AD_CREATIVE_HEADERS_[i]);
    });
  }

  var base = 'https://graph.facebook.com/v18.0/' + ACCOUNT_ID + '/ads';
  // FIXED (Aug 2026): thumbnail_width/thumbnail_height are TOP-LEVEL request parameters
  // on the call, not something you chain onto the `creative` field itself. The original
  // `creative.thumbnail_width(400).thumbnail_height(400){...}` syntax was invalid —
  // that's a malformed request, which is why it failed identically on every one of
  // fetchWithRetry_'s 5 attempts (a genuine transient hiccup would eventually succeed on
  // a retry; a bad query never does). Confirmed against Meta's own AdCreative reference:
  // thumbnail_url "accept[s] thumbnail_width and thumbnail_height" as regular query
  // parameters alongside `fields`, not as chained field parameters.
  // object_story_spec requested BARE (no {link_data{...}} sub-selection) — that nested
  // field-expansion syntax was the other thing I was least sure of last round, and
  // object_story_spec is a polymorphic struct (its shape differs for link_data vs
  // photo_data vs video_data vs template_data) that Graph API's `{}` sub-selection
  // doesn't reliably support the same way it does for a plain object like campaign{name}.
  // Requesting it bare returns the FULL nested JSON object regardless, and
  // classifyCreativeType_() below already reads creative.object_story_spec.link_data...
  // directly off that — same data, zero guessing about expansion syntax.
  // FIXED (Aug 2026): confirmed via the real error text ("Please reduce the amount of
  // data you're asking for, then retry your request") — this is a hard data-volume
  // ceiling, not a rate limit, and Meta's own documented fix for it is exactly this:
  // request less per call. limit:200 per page was fine for plain field pulls elsewhere
  // in this project, but resizing a thumbnail image is real server-side work per ad, and
  // apparently 200 of those in one response is too much. Dropped to 25.
  // BUMPED (Aug 2026, per your "thumbnails look low quality" report): thumbnail_width/
  // height only default to 64px — Meta's docs list no documented maximum, so 400 was a
  // conservative first guess, not a ceiling. Raised to 600. Dropping `limit` further (25
  // -> 15) at the same time is a deliberate trade-off, not a separate fix: a bigger
  // thumbnail is more server-side image work per ad, and 25-per-page was already right at
  // the edge of the "too much data" error above — so pushing resolution up without also
  // pulling batch size down risks reintroducing that exact error. If 15 still trips it,
  // drop it further; if you want even sharper thumbnails and 15 holds up fine, you can
  // try nudging width/height up further from here.
  // FIXED (Sep 2026, per your "some image creatives aren't loading" report): the account
  // has Advantage+/dynamic creative ads whose image lives in asset_feed_spec.images instead
  // of object_story_spec — those ads have NO object_story_spec at all, and often no usable
  // thumbnail_url either, so they were rendering as a broken/blank thumbnail with nothing to
  // fall back to. asset_feed_spec requested here as a SUB-SELECTION (unlike object_story_spec
  // above), not bare: asset_feed_spec is a fixed, well-documented struct (not the
  // link_data/photo_data/video_data polymorphism that made bare-fetching the safer bet for
  // object_story_spec), and for an Advantage+ ad it can otherwise carry many bodies/titles/
  // description/call-to-action text combinations we don't need — bare-fetching that risked
  // tripping the exact "too much data" ceiling described below for no benefit. Only
  // images{hash} and videos{video_id} are pulled — the minimum needed for extractImageHash_
  // and the video-id resolution below to also cover this creative shape.
  // REVERTED (Sep 2026, after your account hit "too many calls to this ad-account" — code
  // 80004, a call-VOLUME rate limit, not the "reduce the amount of data" response-SIZE ceiling
  // this `limit` knob was originally tuned against): dropping `limit` further here was the
  // wrong direction for that failure — a smaller page means MORE pages, i.e. MORE round-trip
  // calls against the account's shared rate budget, which makes a call-volume limit worse, not
  // better. asset_feed_spec is requested as a small sub-selection (see the field comment
  // below), so it shouldn't meaningfully change the per-page response size the old limit:15
  // was already tuned against — put back to 15. If you start seeing the OLD "reduce the amount
  // of data you're asking for" message again (a different error text/code than the rate-limit
  // one — see isAccountRateLimited_ in AdSetPipeline.gs for that one), THAT'S the signal to drop
  // `limit` back down; don't preemptively trade one ceiling for the other without evidence.
  var params = {
    fields: 'id,name,effective_status,campaign{name},adset{name},' +
      // image_hash added Sep 2026 for the full-resolution image lookup (see extractImageHash_
      // and getImageFullMap_ below) — covers the older direct-image creative shape; the more
      // common link_data/photo_data/child_attachments shapes come along for free since
      // object_story_spec is already requested bare (full nested JSON, no sub-selection).
      // asset_feed_spec{images{hash},videos{video_id}} added Sep 2026 (SUB-selected, not
      // bare — see the comment above the OLD version of this line, preserved in git history/
      // your last delivered copy, for why bare would risk the response-size ceiling instead).
      'creative{thumbnail_url,video_id,object_type,object_story_spec,image_hash,asset_feed_spec{images{hash},videos{video_id}}}',
    thumbnail_width: 600,
    thumbnail_height: 600,
    limit: 15,
    'effective_status': JSON.stringify(['ACTIVE', 'PAUSED', 'CAMPAIGN_PAUSED', 'ADSET_PAUSED', 'IN_PROCESS', 'WITH_ISSUES', 'PENDING_REVIEW', 'DISAPPROVED'])
  };
  var qs = Object.keys(params).map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(String(params[k])); }).join('&');
  var next = base + '?' + qs, guard = 0;
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var adItems = [];
  var ok = true, errorMessage = '';

  // FIXED (Aug 2026): same disease as the breakdown importers fixed earlier this
  // session — this used to `break` on any mid-pagination error and then unconditionally
  // write whatever partial `rows` it had collected so far as if it were the complete
  // snapshot, silently truncating ad_creatives (e.g. the first few hundred ads present,
  // the rest quietly missing, with no signal beyond a log line). Now a failure partway
  // through leaves the EXISTING sheet untouched instead of overwriting it with a partial
  // result — a stale-but-complete snapshot beats a fresh-but-truncated one.
  do {
    var data = fetchWithRetry_(next);
    if (!data || data.error) {
      ok = false;
      errorMessage = (data && data.error) ? JSON.stringify(data.error) : 'no response from fetchWithRetry_';
      Logger.log('importAdCreatives error: ' + errorMessage);
      break;
    }
    (data.data || []).forEach(function (ad) {
      adItems.push({ ad: ad, creative: ad.creative || null });
    });
    next = (data.paging && data.paging.next) ? data.paging.next : null;
    if (next) Utilities.sleep(400);
    guard++;
    // Backstop against a runaway loop, not a realistic ceiling: at limit:15/page this
    // covers 30,000 ads. Apps Script's own 6-minute execution limit will kick in long
    // before this does for any account that's actually that large.
  } while (next && guard < 2000);

  if (!ok) {
    Logger.log('importAdCreatives: fetch failed partway (' + errorMessage + ') after ' + adItems.length +
      ' ads collected — existing ad_creatives sheet left untouched rather than overwritten with a partial result. Re-run once the underlying issue clears.');
    return;
  }

  // TIMING CHECKPOINT 1: this is the /ads pagination + per-page 600x600 thumbnail-resize
  // phase — the one whose cost scales with TOTAL AD COUNT, not with any per-run cap you can
  // tune. If this checkpoint alone is already close to 300s (half the 6-minute budget), the
  // fix is NOT lowering `limit` (fewer ads per page = MORE pages = MORE total round-trips/
  // sleep, i.e. worse) — it's reducing the per-image resize cost (thumbnail_width/height
  // below) or accepting that this phase alone needs its own trigger, separate from the
  // video/image enrichment below.
  Logger.log('importAdCreatives: [timing] pagination phase done after ' + ((Date.now() - __t0) / 1000).toFixed(1) +
    's — ' + adItems.length + ' ads across ' + guard + ' page(s).');

  // FIXED (Aug 2026, per your "quality is still bad" report + screenshot): confirmed via
  // Meta's own docs that thumbnail_width/height only resize whatever image is already
  // there — for a VIDEO creative, the source is an auto-extracted poster frame that Meta
  // generates at a low internal resolution regardless of what size you request, so asking
  // for 600x600 just upscales that low-res frame (exactly the soft/blurry look in your
  // screenshot — the static "SHARE"-type image ad next to it was sharp because its source
  // actually was full resolution).
  //
  // FIRST ATTEMPT at fetching each video's own "format" field (filter:"native") failed
  // account-wide: "(#10) Application does not have permission for this action" on all 579
  // unique videos. That was removed rather than left in as dead weight. Root cause turned
  // out to be exactly what the error implied: this app's token had ads_management/ads_read
  // scope only, not the pages_read_engagement needed to read a Page-attached video's own
  // fields. You re-authorized the token with that scope + a role on both Pages
  // ([VIDEO_ID_1], [VIDEO_ID_2]), confirmed via testVideoPermissionAfterReauth(),
  // and it now genuinely works — the "native" filter returns the real full-resolution
  // frame (900px+, no dst-jpg resize/crop parameter on the URL, unlike the 130/480/720
  // tiers which are all resized derivatives of it). This is a real fix this time, not
  // another guess — verified against your account's own data before rebuilding this.
  //
  // Priority order per ad, cheapest/most-reliable first:
  //   1. object_story_spec.video_data.image_url — a custom cover image, if one was ever
  //      set. Free (already in data fetched), full quality when present.
  //   2. The video's own native-resolution frame via getVideoNativePictureMap_() below —
  //      one extra call per UNIQUE video (deduped — many ads reuse the same video), capped
  //      to bound run time, with per-video failure falling back gracefully rather than
  //      failing the whole run (e.g. a video attached to some OTHER Page you haven't been
  //      granted a role on yet would still 403 — that's expected, not a bug, and just
  //      means that one ad falls through to tier 3 same as before).
  //   3. For an IMAGE/CAROUSEL ad: the creative's own full-resolution source asset via
  //      getImageFullMap_() below (Sep 2026, per your "image creatives/ads are bad quality"
  //      report) — same root cause as the video fix: thumbnail_url is a resized PREVIEW
  //      Meta generates for the Ads Manager UI, capped below the actual uploaded image even
  //      at thumbnail_width/height=600. Looked up by image_hash via the account's own
  //      /adimages endpoint — unlike video this needs no extra Page permission, since
  //      image_hash lookups are scoped to the ad account itself.
  //   4. creative.thumbnail_url at 600x600 (unchanged) — the original low-res fallback,
  //      only reached if none of the above produced anything.
  //
  // Sep 2026: video_id resolution now goes through resolveVideoId_() everywhere (not
  // creative.video_id directly) so Advantage+/dynamic creative ads — whose video lives at
  // asset_feed_spec.videos[0].video_id, not the top-level field — are correctly routed into
  // the video branch instead of being misread as an image with no picture at all.
  var videoIdSet = {};
  adItems.forEach(function (item) {
    var vid = resolveVideoId_(item.creative);
    if (vid) videoIdSet[vid] = true;
  });
  var uniqueVideoIds = Object.keys(videoIdSet);
  // CAPPED DOWN (Sep 2026, "Exceeded maximum execution time"): was uncapped here (silently
  // defaulting to 60 inside getVideoNativePictureMap_). Since the asset_feed_spec fix above
  // made Advantage+ video ads visible to this bookkeeping for the FIRST time, this run's
  // uniqueVideoIds pool includes an entire historical backlog that was never counted before
  // — up to 60 brand-new video lookups landing in the SAME run as the heaviest pagination
  // phase this function has. Dropped to 20 as a safety margin while that one-time backlog
  // clears; run enrichVideoThumbnailCache() (cheap pagination, no thumbnail resize, cap 200)
  // repeatedly first to burn down the backlog OUTSIDE this function's budget — once the
  // cache is mostly warm, this inline cap barely ever has anything left to do and 20 is
  // plenty for day-to-day drift.
  var nativePictureByVideoId = uniqueVideoIds.length ? getVideoNativePictureMap_(uniqueVideoIds, 20) : {};
  Logger.log('importAdCreatives: [timing] video native-frame phase done after ' + ((Date.now() - __t0) / 1000).toFixed(1) + 's total.');

  // Sep 2026 image-quality fix (see the priority-order comment above): only look this up
  // for ads that AREN'T video — a video ad's cover image is handled by customCover/nativeFrame.
  var imageHashSet = {};
  adItems.forEach(function (item) {
    var creative = item.creative;
    if (resolveVideoId_(creative)) return;
    var h = extractImageHash_(creative);
    if (h) imageHashSet[h] = true;
  });
  var uniqueImageHashes = Object.keys(imageHashSet);
  // CAPPED DOWN 200 -> 50 (Sep 2026, "Exceeded maximum execution time" — same reasoning as
  // the video cap above): the asset_feed_spec fix means uniqueImageHashes now ALSO includes
  // an entire never-before-extracted Advantage+ backlog, landing in the same run as the
  // heaviest phase this function has. Same remedy: run enrichImageFullCache() (cheap
  // pagination, cap 1000) repeatedly first to clear the backlog outside this budget; once
  // warm, this inline cap rarely has anything left to do.
  var fullImageByHash = uniqueImageHashes.length ? getImageFullMap_(uniqueImageHashes, 50) : {};
  Logger.log('importAdCreatives: [timing] image full-res phase done after ' + ((Date.now() - __t0) / 1000).toFixed(1) + 's total.');

  var rows = adItems.map(function (item) {
    var ad = item.ad, creative = item.creative;
    var vid = resolveVideoId_(creative);
    var customCover = creative && creative.object_story_spec && creative.object_story_spec.video_data &&
      creative.object_story_spec.video_data.image_url;
    var nativeFrame = vid ? nativePictureByVideoId[vid] : null;
    var imgHash = !vid ? extractImageHash_(creative) : null;
    var fullImage = imgHash ? fullImageByHash[imgHash] : null;
    var thumb = customCover || nativeFrame || fullImage || (creative && creative.thumbnail_url) || '';
    // Sep 2026, "some image creatives aren't loading": kept SEPARATE from `thumb` above so the
    // dashboard can retry a lower-quality-but-reliable URL client-side if the "best" one it
    // picked (a signed, time-limited CDN link — see the staleness comments on the video/image
    // caches above) has since expired between daily refreshes. Only meaningful when it
    // differs from `thumb` itself (i.e. thumb picked something better than raw thumbnail_url);
    // the dashboard checks for that before treating it as a real fallback.
    var rawThumbnailUrl = (creative && creative.thumbnail_url) || '';
    return [
      ad.id || '',
      ad.name || '',
      (ad.campaign && ad.campaign.name) || '',
      (ad.adset && ad.adset.name) || '',
      ad.effective_status || '',
      classifyCreativeType_(creative),
      thumb,
      adsManagerLinkFor_(ad.id),
      now,
      rawThumbnailUrl
    ];
  });

  // Full overwrite — this is a CURRENT snapshot, not a log. Old rows for ads that no
  // longer come back (deleted/fully archived) are intentionally dropped. Only reached
  // when the ENTIRE paginated fetch completed successfully (see the ok guard above).
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, AD_CREATIVE_HEADERS_.length).clearContent();
  if (rows.length) {
    ensureColumns_(sheet, AD_CREATIVE_HEADERS_.length);
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  Logger.log('importAdCreatives: snapshotted ' + rows.length + ' ads (' + uniqueVideoIds.length +
    ' unique videos, ' + Object.keys(nativePictureByVideoId).length + ' native-res frames found) at ' + now +
    '. [timing] total run time ' + ((Date.now() - __t0) / 1000).toFixed(1) + 's.');
  markPipelineRefreshed_('creatives');
 } catch (e) {
  Logger.log('⚠️ importAdCreatives: UNCAUGHT exception — ' + (e && e.message ? e.message : String(e)) +
    '. The ad_creatives sheet is untouched (the overwrite only happens at the very end, after everything above succeeds).');
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Meta Ads Dashboard: importAdCreatives crashed',
      'importAdCreatives() threw an uncaught exception: ' + (e && e.message ? e.message : String(e)) +
      '\n\nIf the message mentions "Exceeded maximum execution time": check the [timing] lines in View > ' +
      'Executions > importAdCreatives\'s log FIRST — they show elapsed time at (1) end of /ads pagination, ' +
      '(2) end of video native-frame lookup, (3) end of image full-res lookup, so you can see which phase ' +
      'actually ran out the clock instead of guessing.\n\n' +
      'CORRECTED (Sep 2026): a smaller `limit` in the pagination params does NOT help this — total per-image ' +
      'thumbnail-resize cost scales with your total ad count, not page size, so a smaller limit only means ' +
      'MORE pages (more round-trips + more Utilities.sleep time), making a timeout WORSE, not better. That ' +
      'knob is for the DIFFERENT "reduce the amount of data you\'re asking for" response-size error, not this one.\n\n' +
      'If checkpoint 1 (pagination) alone is eating most of the 6 minutes: that scales with total ad count and ' +
      'is not fixable by tuning caps here — it needs its own trigger, separate from the video/image enrichment.\n' +
      'If checkpoints 2/3 (video or image lookups) account for most of the time: the per-run caps just below ' +
      'each one (currently 20 videos / 50 hashes) are the lever — lower them further, and run ' +
      'enrichVideoThumbnailCache() / enrichImageFullCache() manually first to clear backlog outside this budget.\n\n' +
      'Otherwise check View > Executions > importAdCreatives in the Apps Script editor for the full stack trace.');
  } catch (mailErr) {
    Logger.log('(Could not send the crash email — Logger above still has it. Reason: ' + mailErr.message + ')');
  }
 }
}

// FIXED (Aug 2026): "Exceeded maximum execution time." The FIRST version of this function
// re-fetched EVERY unique video EVERY single run, capped at 150/run to survive Apps
// Script's 6-minute ceiling. That cap was set back when the permission wall meant every
// one of those calls failed instantly (a permission error returns fast, so 579 failed
// calls barely cost any time at all). Now that the permission actually works, each call
// does real work and takes real time — combined with importAdCreatives()'s own pagination
// (which processes thumbnail images for up to 3,048 ads), even 150 successful video calls
// pushed the whole run over 6 minutes. But the deeper problem wasn't the cap size, it was
// that this was RE-DOING already-finished work on every run: a video's native poster frame
// doesn't change once resolved, so re-fetching it tomorrow is pure waste. Fixed properly
// with a persistent cache sheet (ad_video_thumb_cache) — a video is only ever fetched
// ONCE across the life of this script, successful or not (a definitive failure, e.g. a
// video on a Page you're not granted, is cached too, as 'FAILED', so it's not retried
// forever). Each run only pays for videos it's never seen before, and the per-run cap
// (60, well under the 150 that caused the timeout) leaves comfortable headroom for
// importAdCreatives()'s own pagination to finish inside the 6-minute limit alongside it.
// Run enrichVideoThumbnailCache() (below) separately, as many times as you like, to grind
// through a large backlog faster than importAdCreatives()'s conservative per-run cap alone
// — it skips the heavy per-ad thumbnail-image pagination entirely, so it can afford a
// bigger batch per call.
var AD_VIDEO_THUMB_CACHE_SHEET_ = 'ad_video_thumb_cache';
var AD_VIDEO_THUMB_CACHE_HEADERS_ = ['Video ID', 'Native Picture URL (or FAILED)', 'Resolved At'];

// Returns { videoId: { value: <url or 'FAILED'>, resolvedAt: <'yyyy-MM-dd HH:mm' string or ''> } }.
// FIXED (Sep 2026, "some ads thumbnails aren't showing anymore"): this used to return just
// the raw value (url or 'FAILED') with the "Resolved At" column read from the sheet but
// never actually used for anything. That's the bug: Meta's picture/video CDN URLs are
// SIGNED links with a built-in expiry (the `oe=` param in the URL) — they are NOT permanent
// identifiers. Caching one forever was fine for a 'FAILED' entry (that's just a "don't
// bother retrying" marker, no URL involved) but wrong for a genuinely resolved URL: it
// looked fine on the day it was cached, then quietly started 404ing days later as its
// signature expired, with nothing in this pipeline ever re-checking it. Now the timestamp
// comes back with the value so getVideoNativePictureMap_ can tell "resolved" from "resolved
// a while ago, due for a refresh."
function readVideoThumbCache_() {
  var sheet = getOrCreateAdPlainSheet_(AD_VIDEO_THUMB_CACHE_SHEET_);
  var first = sheet.getRange(1, 1, 1, AD_VIDEO_THUMB_CACHE_HEADERS_.length).getValues()[0];
  if (first.every(function (c) { return String(c).trim() === ''; })) {
    sheet.getRange(1, 1, 1, AD_VIDEO_THUMB_CACHE_HEADERS_.length).setValues([AD_VIDEO_THUMB_CACHE_HEADERS_]);
    return {};
  }
  var lastRow = sheet.getLastRow();
  var map = {};
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, AD_VIDEO_THUMB_CACHE_HEADERS_.length).getValues().forEach(function (r) {
      if (r[0]) map[String(r[0])] = { value: r[1], resolvedAt: r[2] || '' };
    });
  }
  return map;
}

function appendVideoThumbCacheEntries_(entries) {
  if (!entries.length) return;
  var sheet = getOrCreateAdPlainSheet_(AD_VIDEO_THUMB_CACHE_SHEET_);
  ensureColumns_(sheet, AD_VIDEO_THUMB_CACHE_HEADERS_.length);
  sheet.getRange(sheet.getLastRow() + 1, 1, entries.length, 3).setValues(entries);
}

// One extra lightweight call per unique video NOT ALREADY IN THE CACHE (not per ad, and
// not per run) to get Meta's actual full-resolution frame grab instead of the ad
// creative's own low-res auto-thumbnail. CONFIRMED WORKING (Aug 2026) via
// testVideoPermissionAfterReauth() after re-authorizing the token with
// pages_read_engagement + a role on the relevant Pages.
// maxNewPerRun is deliberately small by default (60) because this normally runs INSIDE
// importAdCreatives(), sharing the 6-minute ceiling with that function's own pagination —
// pass a bigger number (from enrichVideoThumbnailCache()) when nothing else is competing
// for time in the same run.
// A resolved (non-FAILED) cache entry older than this is treated as due for a refresh —
// see the readVideoThumbCache_ comment above for why. 5 days is a judgment call, not a
// number Meta documents anywhere: it's a conservative guess at "comfortably inside
// whatever the CDN's signed-URL lifetime actually is." If thumbnails start going stale
// again even sooner than this, lower it; Meta's own thumbnail_url (the fallback used
// whenever a cache entry is missing/stale) still auto-refreshes independently, so nothing
// breaks either way — an ad is never stuck on a fully-dead image, at worst it's stuck on
// its lower-res fallback for a few extra hours until the next cache refresh sweeps it up.
var VIDEO_THUMB_STALE_DAYS_ = 5;

function isVideoThumbCacheEntryStale_(entry) {
  if (!entry) return false; // not in cache at all — that's "missing", handled separately
  if (entry.value === 'FAILED') return false; // a permanent give-up marker, not a URL — nothing to expire
  if (!entry.resolvedAt) return true; // pre-existing rows from before this timestamp fix — refresh once to backfill it
  var ageMs = new Date().getTime() - new Date(entry.resolvedAt).getTime();
  return isNaN(ageMs) ? true : ageMs > VIDEO_THUMB_STALE_DAYS_ * 86400000;
}

function getVideoNativePictureMap_(videoIds, maxNewPerRun) {
  maxNewPerRun = maxNewPerRun || 60;
  var cache = readVideoThumbCache_(); // videoId -> { value: <url or 'FAILED'>, resolvedAt }
  // "Missing" now means genuinely uncached OR cached-but-stale (see isVideoThumbCacheEntryStale_) —
  // a resolved URL doesn't get to sit in the cache forever just because it worked once.
  var missing = videoIds.filter(function (id) { return !(id in cache) || isVideoThumbCacheEntryStale_(cache[id]); });
  var toFetch = missing.slice(0, maxNewPerRun);
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var newEntries = [];

  // FIXED (Aug 2026): the first version of this loop cached ANY failure as permanent
  // 'FAILED', including errors Meta itself labels "is_transient":true (e.g. code 4,
  // "Application request limit reached"). That was a real correctness bug, not just bad
  // luck — a video that happened to hit a rate limit during a fetch got permanently stuck
  // on the old blurry thumbnail forever, because it would never be attempted again even
  // after the rate limit cleared. Now: only a confirmed NON-transient error (checked via
  // isTransientMetaError_, the same function used elsewhere in this codebase for this
  // exact distinction — e.g. code 10 permission-denied, which genuinely will never
  // resolve on its own) gets cached as FAILED. A transient error is left uncached
  // entirely, so it's naturally retried on the next run. If several transient errors hit
  // in a row, that's not bad luck with individual videos — it means Meta is rate-limiting
  // this ad account right now, and continuing to hammer it with more calls only makes that
  // worse. Stop early in that case rather than burning through the rest of this batch.
  var consecutiveTransientFailures = 0;
  for (var ti = 0; ti < toFetch.length; ti++) {
    var videoId = toFetch[ti];
    var url = 'https://graph.facebook.com/v18.0/' + videoId + '?fields=format,picture';
    // Only 2 retries here (not the default 5) — thumbnail_url is always a safe fallback,
    // so it's not worth burning a full 5-attempt backoff per video on a "nice to have" hi-res grab.
    var data = fetchWithRetry_(url, 2);
    if (data && !data.error) {
      var native = null, widest = null;
      (data.format || []).forEach(function (f) {
        if (!f || !f.picture) return;
        if (f.filter === 'native') native = f;
        if (!widest || (f.width || 0) > (widest.width || 0)) widest = f;
      });
      var chosen = native || widest;
      var resolved = (chosen && chosen.picture) || data.picture || null;
      var storedValue = resolved || 'FAILED'; // a real response with no picture at all is genuinely odd/unusual, not a rate limit — treat as failed
      cache[videoId] = { value: storedValue, resolvedAt: now };
      newEntries.push([videoId, storedValue, now]);
      consecutiveTransientFailures = 0;
    } else if (data && data.error && isTransientMetaError_(data.error)) {
      Logger.log('getVideoNativePictureMap_: transient error on video ' + videoId + ' (' + JSON.stringify(data.error) +
        ') — NOT caching as failed, this will be retried on a future run.');
      consecutiveTransientFailures++;
      if (consecutiveTransientFailures >= 3) {
        Logger.log('getVideoNativePictureMap_: ' + consecutiveTransientFailures + ' transient errors in a row — this is Meta ' +
          'rate-limiting this ad account, not a problem with these specific videos. Stopping early this run instead of making ' +
          'it worse. Wait at least 15-20 minutes before running this again.');
        break;
      }
    } else {
      Logger.log('getVideoNativePictureMap_: permanent failure for video ' + videoId + ' (' +
        (data && data.error ? JSON.stringify(data.error) : 'no response') + ') — caching as FAILED so it is not retried every run; falls back to thumbnail_url.');
      cache[videoId] = { value: 'FAILED', resolvedAt: now };
      newEntries.push([videoId, 'FAILED', now]);
      consecutiveTransientFailures = 0;
    }
    if (ti < toFetch.length - 1) Utilities.sleep(150);
  }
  appendVideoThumbCacheEntries_(newEntries);

  if (missing.length > toFetch.length) {
    Logger.log('getVideoNativePictureMap_: ' + (missing.length - toFetch.length) + ' videos still not cached after this run — ' +
      'they\'ll use thumbnail_url for now. Run enrichVideoThumbnailCache() (bigger per-run batch, no competing pagination) ' +
      'or just let subsequent daily importAdCreatives() runs keep chipping away at the backlog — already-cached videos never cost another API call.');
  }

  var resultMap = {};
  videoIds.forEach(function (id) {
    if (cache[id] && cache[id].value && cache[id].value !== 'FAILED') resultMap[id] = cache[id].value;
  });
  Logger.log('getVideoNativePictureMap_: ' + Object.keys(resultMap).length + ' of ' + videoIds.length +
    ' requested videos have a native frame (cache-wide: ' + Object.keys(cache).length + ' videos resolved or confirmed unavailable so far, ' +
    newEntries.length + ' fetched just now).');
  return resultMap;
}

// STANDALONE CATCH-UP (Aug 2026): run this manually as many times as you like to grind
// through the video-thumbnail cache backlog faster than importAdCreatives()'s conservative
// 60-per-run cap. This does its OWN much cheaper pagination (id + video_id only — no
// thumbnail_width/height, no image resizing, so a limit of 300 is safe here even though
// importAdCreatives() itself has to use 15), so nearly all its 6-minute budget is free for
// the actual video lookups: a bigger batch (200) per call. Safe to run repeatedly — every
// video it resolves is cached permanently and simply won't be re-fetched next time.
function enrichVideoThumbnailCache() {
  var base = 'https://graph.facebook.com/v18.0/' + ACCOUNT_ID + '/ads';
  var params = {
    fields: 'id,creative{video_id}',
    limit: 300,
    'effective_status': JSON.stringify(['ACTIVE', 'PAUSED', 'CAMPAIGN_PAUSED', 'ADSET_PAUSED', 'IN_PROCESS', 'WITH_ISSUES', 'PENDING_REVIEW', 'DISAPPROVED'])
  };
  var qs = Object.keys(params).map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(String(params[k])); }).join('&');
  var next = base + '?' + qs, guard = 0;
  var videoIdSet = {};
  do {
    var data = fetchWithRetry_(next);
    if (!data || data.error) {
      // FIXED (Aug 2026): code 80004 ("too many calls to this ad-account") is Meta
      // rate-limiting the WHOLE account, typically from running this function repeatedly
      // back-to-back — each call already makes 200+ requests on its own. This is NOT the
      // same thing as "0 videos found," and letting execution fall through to
      // getVideoNativePictureMap_([], 200) below made it LOOK like a successful run that
      // legitimately found nothing, which is actively misleading. Detect it and stop
      // clearly instead.
      var err = data && data.error;
      var isAccountRateLimit = err && (err.code === 80004 || /too many calls/i.test(err.message || ''));
      if (isAccountRateLimit) {
        Logger.log('enrichVideoThumbnailCache: STOPPED — Meta is rate-limiting this whole ad account right now (' +
          JSON.stringify(err) + '). This is not "0 videos," it means too many API calls have hit this account in a short ' +
          'window — almost certainly from running this function repeatedly back-to-back. Wait at least 15-20 minutes ' +
          'before running this again, and avoid re-running it multiple times in quick succession.');
      } else {
        Logger.log('enrichVideoThumbnailCache: fetch failed — ' + (err ? JSON.stringify(err) : 'no response'));
      }
      return; // stop here — do NOT proceed to resolve videos against an empty/partial list as if this were a normal pass
    }
    (data.data || []).forEach(function (ad) {
      var vid = ad.creative && ad.creative.video_id;
      if (vid) videoIdSet[vid] = true;
    });
    next = (data.paging && data.paging.next) ? data.paging.next : null;
    guard++;
  } while (next && guard < 500);

  var uniqueVideoIds = Object.keys(videoIdSet);
  Logger.log('enrichVideoThumbnailCache: ' + uniqueVideoIds.length + ' unique videos across your account. Resolving up to 200 uncached ones...');
  getVideoNativePictureMap_(uniqueVideoIds, 200);
  Logger.log('enrichVideoThumbnailCache: done this pass. Re-run this function again if the log above still shows a backlog, then run importAdCreatives() to apply the newly cached thumbnails to the sheet. Don\'t run it back-to-back repeatedly — space runs a few minutes apart to avoid tripping Meta\'s account-level rate limit.');
}

// ONE-TIME REPAIR (Aug 2026) — run this ONCE, now, before your next enrichVideoThumbnailCache()
// or importAdCreatives() call. The bug described above (in getVideoNativePictureMap_) means
// some videos already sitting in your ad_video_thumb_cache sheet as 'FAILED' were actually
// just transient rate-limit hits during an earlier run, not genuine permanent failures — under
// the OLD buggy logic they'd be stuck as FAILED forever even though they'd very likely resolve
// fine now. This clears every FAILED row so they all get a fair retry under the fixed logic.
// Safe to run even if some of those really were permanent failures (e.g. a genuinely
// unauthorized Page) — they'll just get re-marked FAILED again, correctly this time, on the
// next run. Only needs to be run once; after this, the fixed code won't create bad entries.
function resetFailedVideoThumbCacheEntries() {
  var sheet = getOrCreateAdPlainSheet_(AD_VIDEO_THUMB_CACHE_SHEET_);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log('resetFailedVideoThumbCacheEntries: cache is empty, nothing to reset.');
    return;
  }
  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var kept = data.filter(function (r) { return r[1] !== 'FAILED'; });
  var removedCount = data.length - kept.length;
  sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
  if (kept.length) sheet.getRange(2, 1, kept.length, 3).setValues(kept);
  Logger.log('resetFailedVideoThumbCacheEntries: removed ' + removedCount + ' FAILED entries out of ' + data.length +
    ' total cached videos. These will be retried fresh next time you run enrichVideoThumbnailCache() or importAdCreatives() ' +
    '— genuinely unresolvable ones will simply be marked FAILED again, correctly this time.');
}

// ═══════════════════════════════════════════════════════════════════════════
// ── IMAGE CREATIVE FULL-RESOLUTION LOOKUP (Sep 2026, "the image creatives/ads are bad
// quality") ── Same root cause as the video fix above: creative.thumbnail_url is a resized
// PREVIEW Meta generates for the Ads Manager UI, capped well below the actual uploaded
// image — even at thumbnail_width/height=600 (see importAdCreatives()'s params) it can look
// soft next to the original. The account's own /adimages endpoint, keyed by the creative's
// image_hash, returns the real original-resolution asset instead — the image-side
// equivalent of /{video_id}?fields=format,picture for video.
//
// Good news relative to the video fix: this does NOT need pages_read_engagement or a role
// on any Page. image_hash lookups are scoped to the ad ACCOUNT itself (the account already
// owns every image it uploaded), so there is no repeat of the permission saga above.
//
// Same signed-URL-expiry caveat as the video cache, though: adimages.url is also a
// time-limited CDN link, not a permanent identifier — see VIDEO_THUMB_STALE_DAYS_'s comment
// for why a "resolved a while ago" entry still needs a periodic refresh rather than being
// cached forever.
// ═══════════════════════════════════════════════════════════════════════════

var AD_IMAGE_FULL_CACHE_SHEET_ = 'ad_image_full_cache';
var AD_IMAGE_FULL_CACHE_HEADERS_ = ['Image Hash', 'Full URL (or FAILED)', 'Resolved At'];
var IMAGE_FULL_STALE_DAYS_ = 5; // same judgment call as VIDEO_THUMB_STALE_DAYS_ above, not a Meta-documented number

function readImageFullCache_() {
  var sheet = getOrCreateAdPlainSheet_(AD_IMAGE_FULL_CACHE_SHEET_);
  var first = sheet.getRange(1, 1, 1, AD_IMAGE_FULL_CACHE_HEADERS_.length).getValues()[0];
  if (first.every(function (c) { return String(c).trim() === ''; })) {
    sheet.getRange(1, 1, 1, AD_IMAGE_FULL_CACHE_HEADERS_.length).setValues([AD_IMAGE_FULL_CACHE_HEADERS_]);
    return {};
  }
  var lastRow = sheet.getLastRow();
  var map = {};
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, AD_IMAGE_FULL_CACHE_HEADERS_.length).getValues().forEach(function (r) {
      if (r[0]) map[String(r[0])] = { value: r[1], resolvedAt: r[2] || '' };
    });
  }
  return map;
}

function appendImageFullCacheEntries_(entries) {
  if (!entries.length) return;
  var sheet = getOrCreateAdPlainSheet_(AD_IMAGE_FULL_CACHE_SHEET_);
  ensureColumns_(sheet, AD_IMAGE_FULL_CACHE_HEADERS_.length);
  sheet.getRange(sheet.getLastRow() + 1, 1, entries.length, 3).setValues(entries);
}

function isImageFullCacheEntryStale_(entry) {
  if (!entry) return false; // not in cache at all — that's "missing", handled separately
  if (entry.value === 'FAILED') return false; // permanent give-up marker, not a URL — nothing to expire
  if (!entry.resolvedAt) return true; // pre-existing rows from before this timestamp existed — refresh once
  var ageMs = new Date().getTime() - new Date(entry.resolvedAt).getTime();
  return isNaN(ageMs) ? true : ageMs > IMAGE_FULL_STALE_DAYS_ * 86400000;
}

// Pulls an image_hash off of whichever shape this particular ad's creative actually uses.
// Only the FIRST card's hash is used for a carousel — good enough to fix "the thumbnail
// looks soft," not a full per-card gallery (the creative card only ever shows one image).
// NOT covered: Advantage+/dynamic creatives that source their image from asset_feed_spec
// instead of object_story_spec — that's a materially different, per-placement asset
// structure. Those ads simply fall back to thumbnail_url, exactly as they did before this
// fix (no regression, just an acknowledged gap).
function extractImageHash_(creative) {
  if (!creative) return null;
  if (creative.image_hash) return creative.image_hash;
  var spec = creative.object_story_spec || {};
  if (spec.link_data) {
    if (spec.link_data.image_hash) return spec.link_data.image_hash;
    if (spec.link_data.child_attachments && spec.link_data.child_attachments.length) {
      var first = spec.link_data.child_attachments[0];
      if (first && first.image_hash) return first.image_hash;
    }
  }
  if (spec.photo_data && spec.photo_data.image_hash) return spec.photo_data.image_hash;
  // Sep 2026, "some image creatives aren't loading": Advantage+/dynamic creative ads have NO
  // object_story_spec at all (that's WHY this was previously an acknowledged gap, not just an
  // oversight) — their image(s) live in asset_feed_spec.images instead. This is only reached
  // for ads that resolveVideoId_() has already determined are NOT video, so an asset_feed_spec
  // image here is genuinely this ad's picture, not a video's bundled cover option.
  var afs = creative.asset_feed_spec || {};
  if (afs.images && afs.images.length && afs.images[0] && afs.images[0].hash) return afs.images[0].hash;
  return null;
}

// One /adimages call per batch of up to 40 hashes NOT ALREADY IN THE CACHE (not per ad, and
// not per run) — Meta doesn't document a hard per-call cap on hashes, but batching keeps
// any single request small and safely retry-able. maxNewPerRun bounds how many hashes get
// resolved in one call to this function (default 400 = up to 10 batch calls); pass a bigger
// number from enrichImageFullCache() when nothing else is competing for time in the same run.
function getImageFullMap_(hashes, maxNewPerRun) {
  maxNewPerRun = maxNewPerRun || 400;
  var cache = readImageFullCache_();
  var missing = hashes.filter(function (h) { return !(h in cache) || isImageFullCacheEntryStale_(cache[h]); });
  var toFetch = missing.slice(0, maxNewPerRun);
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var newEntries = [];
  var actPath = String(ACCOUNT_ID || '');

  // Same transient-vs-permanent distinction as getVideoNativePictureMap_ above, and for the
  // same reason: a rate-limit hit during one run must NOT get permanently cached as FAILED,
  // or the image would be stuck on its lower-res fallback forever even after the limit clears.
  var consecutiveTransientFailures = 0;
  for (var i = 0; i < toFetch.length; i += 40) {
    var batch = toFetch.slice(i, i + 40);
    var url = 'https://graph.facebook.com/v18.0/' + actPath + '/adimages?fields=hash,url&hashes=' + encodeURIComponent(JSON.stringify(batch));
    var data = fetchWithRetry_(url, 2);
    if (data && !data.error) {
      var foundHashes = {};
      (data.data || []).forEach(function (img) {
        if (!img || !img.hash) return;
        foundHashes[img.hash] = true;
        var storedValue = img.url || 'FAILED';
        cache[img.hash] = { value: storedValue, resolvedAt: now };
        newEntries.push([img.hash, storedValue, now]);
      });
      // Any hash in this batch Meta didn't return at all (deleted image, stale/bad hash,
      // etc.) is a genuine permanent miss — cache as FAILED so it's not retried every run.
      batch.forEach(function (h) {
        if (!foundHashes[h]) { cache[h] = { value: 'FAILED', resolvedAt: now }; newEntries.push([h, 'FAILED', now]); }
      });
      consecutiveTransientFailures = 0;
    } else if (data && data.error && isTransientMetaError_(data.error)) {
      Logger.log('getImageFullMap_: transient error on a batch of ' + batch.length + ' hashes (' + JSON.stringify(data.error) +
        ') — NOT caching as failed, will be retried on a future run.');
      consecutiveTransientFailures++;
      if (consecutiveTransientFailures >= 3) {
        Logger.log('getImageFullMap_: 3 transient errors in a row — Meta is rate-limiting this ad account right now, not a ' +
          'problem with these specific images. Stopping early this run rather than making it worse.');
        break;
      }
    } else {
      Logger.log('getImageFullMap_: permanent failure for a batch of ' + batch.length + ' hashes (' +
        (data && data.error ? JSON.stringify(data.error) : 'no response') + ') — caching as FAILED; falls back to thumbnail_url.');
      batch.forEach(function (h) { cache[h] = { value: 'FAILED', resolvedAt: now }; newEntries.push([h, 'FAILED', now]); });
      consecutiveTransientFailures = 0;
    }
    if (i + 40 < toFetch.length) Utilities.sleep(150);
  }
  appendImageFullCacheEntries_(newEntries);

  if (missing.length > toFetch.length) {
    Logger.log('getImageFullMap_: ' + (missing.length - toFetch.length) + ' image hashes still not cached after this run — ' +
      'they will use thumbnail_url for now. Run enrichImageFullCache() to grind through the backlog faster, or let subsequent ' +
      'daily importAdCreatives() runs keep chipping away at it — already-cached hashes never cost another API call.');
  }

  var resultMap = {};
  hashes.forEach(function (h) {
    if (cache[h] && cache[h].value && cache[h].value !== 'FAILED') resultMap[h] = cache[h].value;
  });
  Logger.log('getImageFullMap_: ' + Object.keys(resultMap).length + ' of ' + hashes.length +
    ' requested image hashes have a full-res URL (cache-wide: ' + Object.keys(cache).length + ' resolved or confirmed ' +
    'unavailable so far, ' + newEntries.length + ' fetched just now).');
  return resultMap;
}

// STANDALONE CATCH-UP — same purpose as enrichVideoThumbnailCache() above: grind through
// the image-hash cache backlog without competing for time against importAdCreatives()'s own
// thumbnail_width/height pagination. This pagination avoids that function's execution-TIME
// cost (no per-image resizing), so nearly the whole 6-minute budget is free for /adimages
// calls — but it is NOT cheap on response SIZE (object_story_spec bare is a large blob; see
// the limit comment inside the function for the real ceiling this runs into and why). Safe
// to run repeatedly — every hash it resolves is cached and won't be re-fetched next time
// (until IMAGE_FULL_STALE_DAYS_ makes it due for a refresh).
function enrichImageFullCache() {
  var base = 'https://graph.facebook.com/v18.0/' + ACCOUNT_ID + '/ads';
  // Sep 2026: added asset_feed_spec{images{hash},videos{video_id}} (same sub-selected,
  // minimal-payload request as importAdCreatives() — see that function's field comment for
  // why this is sub-selected rather than bare) so this catch-up pass also finds Advantage+/
  // dynamic creative image hashes, which the old field list here silently missed entirely.
  //
  // FIXED (Sep 2026, real "Please reduce the amount of data you're asking for" failure — code
  // 1, a response-SIZE ceiling, not the account-rate-limit code 80004 or an execution timeout;
  // this one fails identically on all 5 retries because retrying identical params can't fix a
  // request that's just too big per page). The stale comment this replaces claimed this
  // pagination was "cheap... no image resizing, so limit 300 is safe here" — true for
  // execution TIME (no thumbnail resize cost) but that was never the relevant ceiling here.
  // object_story_spec requested BARE is itself a large nested blob, and this exact
  // combination (object_story_spec bare + a big per-page limit) already has a proven failure
  // precedent in THIS file: findMissingPageAccess() below hit the identical error with a
  // smaller field set (video_id, object_story_spec, effective_object_story_id) at limit 200
  // and had to drop to limit 30 to clear it. This field list carries the same object_story_spec
  // cost PLUS image_hash and the asset_feed_spec sub-selection on top, at limit 300 — nearly
  // 10x the proven-safe size. Matched to that same proven-working limit; if 30 still trips it,
  // drop further (object_story_spec is the dominant cost, not the fields added alongside it).
  var params = {
    fields: 'id,creative{video_id,image_hash,object_story_spec,asset_feed_spec{images{hash},videos{video_id}}}',
    limit: 30,
    'effective_status': JSON.stringify(['ACTIVE', 'PAUSED', 'CAMPAIGN_PAUSED', 'ADSET_PAUSED', 'IN_PROCESS', 'WITH_ISSUES', 'PENDING_REVIEW', 'DISAPPROVED'])
  };
  var qs = Object.keys(params).map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(String(params[k])); }).join('&');
  var next = base + '?' + qs, guard = 0;
  var hashSet = {};
  do {
    var data = fetchWithRetry_(next);
    if (!data || data.error) {
      Logger.log('enrichImageFullCache: fetch failed — ' + (data && data.error ? JSON.stringify(data.error) : 'no response'));
      return;
    }
    (data.data || []).forEach(function (ad) {
      var creative = ad.creative;
      if (resolveVideoId_(creative)) return;
      var h = extractImageHash_(creative);
      if (h) hashSet[h] = true;
    });
    next = (data.paging && data.paging.next) ? data.paging.next : null;
    // ADDED (Sep 2026, alongside the limit 300->30 fix): dropping the per-page limit 10x
    // means ~10x more pages for the same ad count, i.e. ~10x more calls against the
    // account's shared rate budget in the same short window — exactly the ingredient that
    // already caused a real code-80004 account-level rate-limit hit earlier this project.
    // A small sleep between pages costs little wall-clock time here (this function has no
    // image-resize work eating its budget) and meaningfully lowers that risk.
    if (next) Utilities.sleep(300);
    guard++;
  } while (next && guard < 500);

  var uniqueHashes = Object.keys(hashSet);
  Logger.log('enrichImageFullCache: ' + uniqueHashes.length + ' unique image hashes across your account. Resolving up to 1000 uncached ones...');
  getImageFullMap_(uniqueHashes, 1000);
  Logger.log('enrichImageFullCache: done this pass. Re-run again if the log above still shows a backlog, then run importAdCreatives() to apply the newly cached images to the sheet.');
}

// Exact same priority order as extractImageHash_() above, but returning WHICH branch fired
// instead of just the hash — used only by the diagnostic below to answer "where did this
// FAILED hash actually come from," not by any real import path.
function classifyImageHashSource_(creative) {
  if (!creative) return 'none';
  if (creative.image_hash) return 'creative.image_hash';
  var spec = creative.object_story_spec || {};
  if (spec.link_data) {
    if (spec.link_data.image_hash) return 'object_story_spec.link_data.image_hash';
    if (spec.link_data.child_attachments && spec.link_data.child_attachments.length &&
        spec.link_data.child_attachments[0] && spec.link_data.child_attachments[0].image_hash) {
      return 'object_story_spec.link_data.child_attachments[0].image_hash';
    }
  }
  if (spec.photo_data && spec.photo_data.image_hash) return 'object_story_spec.photo_data.image_hash';
  var afs = creative.asset_feed_spec || {};
  if (afs.images && afs.images.length && afs.images[0] && afs.images[0].hash) return 'asset_feed_spec.images[0].hash (Advantage+)';
  return 'none';
}

// ONE-OFF DIAGNOSTIC (Sep 2026) — same pattern as findMissingPageAccess() above, for the
// image-cache equivalent question: enrichImageFullCache() logged 175 of 270 hashes resolved,
// "0 fetched just now" — meaning the other 95 are ALREADY cached as FAILED and getImageFullMap_
// will never retry a FAILED entry (that's the whole point of the marker — see
// isImageFullCacheEntryStale_ above). So simply re-running enrichImageFullCache() again cannot
// change this number; something else has to give first. This cross-references every FAILED
// hash against the actual ad(s) referencing it to answer WHY, rather than guessing: if the
// failures cluster on asset_feed_spec (Advantage+/dynamic creative) hashes, that's a real
// structural gap — those hashes may not exist in the account's own /adimages library the same
// way a classic object_story_spec/image_hash creative's does, and needs a different fix (a
// different lookup endpoint, or accepting thumbnail_url as the ceiling for those ads). If
// they're scattered across ordinary object_story_spec-sourced hashes instead, that points to
// genuinely deleted/stale images instead, not a shape-specific gap.
function findFailedImageHashSource() {
  var cache = readImageFullCache_();
  var failedHashes = {};
  Object.keys(cache).forEach(function (h) { if (cache[h].value === 'FAILED') failedHashes[h] = true; });
  var failedCount = Object.keys(failedHashes).length;
  if (!failedCount) {
    Logger.log('findFailedImageHashSource: no FAILED hashes in the cache right now — nothing to chase.');
    return;
  }
  Logger.log('findFailedImageHashSource: looking up the source of ' + failedCount + ' FAILED image hash(es)...');

  var base = 'https://graph.facebook.com/v18.0/' + ACCOUNT_ID + '/ads';
  var params = {
    // Same field set (minus video_id, not needed here) as extractImageHash_ needs to cover
    // every branch it checks. limit:30 — same object_story_spec-bare size ceiling documented
    // on enrichImageFullCache() above; this requests the identical creative shape.
    fields: 'id,name,creative{image_hash,object_story_spec,asset_feed_spec{images{hash}}}',
    limit: 30,
    'effective_status': JSON.stringify(['ACTIVE', 'PAUSED', 'CAMPAIGN_PAUSED', 'ADSET_PAUSED', 'IN_PROCESS', 'WITH_ISSUES', 'PENDING_REVIEW', 'DISAPPROVED'])
  };
  var qs = Object.keys(params).map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(String(params[k])); }).join('&');
  var next = base + '?' + qs, guard = 0;
  var bySource = {}; // source label -> { count, example }
  var matchedHashes = {};

  do {
    var data = fetchWithRetry_(next);
    if (!data || data.error) {
      Logger.log('findFailedImageHashSource: fetch failed partway — ' + (data && data.error ? JSON.stringify(data.error) : 'no response') +
        '. Whatever was found before this point is still reported below; re-run later to fill in the rest.');
      break;
    }
    (data.data || []).forEach(function (ad) {
      var creative = ad.creative;
      if (!creative) return;
      var h = extractImageHash_(creative);
      if (!h || !failedHashes[h] || matchedHashes[h]) return;
      matchedHashes[h] = true;
      var source = classifyImageHashSource_(creative);
      if (!bySource[source]) bySource[source] = { count: 0, example: ad.name || ad.id };
      bySource[source].count++;
    });
    if (Object.keys(matchedHashes).length >= failedCount) {
      Logger.log('findFailedImageHashSource: matched all ' + failedCount + ' FAILED hashes — stopping early instead of scanning the rest of the account.');
      break;
    }
    next = (data.paging && data.paging.next) ? data.paging.next : null;
    if (next) Utilities.sleep(300);
    guard++;
  } while (next && guard < 500);

  var matchedCount = Object.keys(matchedHashes).length;
  if (!matchedCount) {
    Logger.log('findFailedImageHashSource: matched 0 of ' + failedCount + ' FAILED hashes to an ad — pagination may not have reached ' +
      'the relevant ads yet, or those ads have since been deleted/archived. Try re-running.');
    return;
  }
  Logger.log('findFailedImageHashSource: matched ' + matchedCount + ' of ' + failedCount + ' FAILED hash(es) to a source, grouped by which field produced the hash:');
  Object.keys(bySource).sort(function (a, b) { return bySource[b].count - bySource[a].count; }).forEach(function (source) {
    var info = bySource[source];
    Logger.log('  ' + source + ': ' + info.count + ' hash(es) — e.g. ad "' + info.example + '"');
  });
}

// ONE-OFF REPORT (Sep 2026, "catch leftover ads/creatives still low quality") — pure read
// of the ad_creatives sheet importAdCreatives() already wrote; makes ZERO Meta API calls,
// so it's safe to run any time, back-to-back, with no rate-limit risk at all. This is the
// FULL current picture, not just the 95 image hashes findFailedImageHashSource() already
// explained — it also catches VIDEO ads still waiting on their native-frame cache entry
// (576 were reported stale/uncached as of the last importAdCreatives() run), which that
// image-only diagnostic never looked at.
//
// "Low quality" is defined concretely off the same priority chain the dashboard itself uses
// (thumb = customCover || nativeFrame || fullImage || thumbnail_url): a row where Thumbnail
// URL and Fallback Thumbnail URL are IDENTICAL and non-blank means nothing beat
// creative.thumbnail_url — this ad is stuck on the lowest tier. A row where Thumbnail URL is
// blank entirely means nothing resolved at all — that one renders as the placeholder icon,
// not a low-res image.
function findLowQualityCreatives() {
  var sheet = getOrCreateAdPlainSheet_(AD_CREATIVE_SHEET_);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('findLowQualityCreatives: ad_creatives sheet is empty — run importAdCreatives() first.');
    return;
  }
  var rows = sheet.getRange(2, 1, lastRow - 1, AD_CREATIVE_HEADERS_.length).getValues();

  var blank = [];       // Thumbnail URL empty entirely — renders as the placeholder icon
  var stuckByType = {}; // creative type -> [{id, name}] stuck on tier-4 (raw thumbnail_url)

  rows.forEach(function (r) {
    var adId = r[0], adName = r[1], type = r[5] || 'Unknown', thumb = r[6], fallback = r[9];
    if (!thumb) {
      blank.push({ id: adId, name: adName, type: type });
      return;
    }
    if (fallback && thumb === fallback) {
      if (!stuckByType[type]) stuckByType[type] = [];
      stuckByType[type].push({ id: adId, name: adName });
    }
  });

  var totalStuck = Object.keys(stuckByType).reduce(function (sum, t) { return sum + stuckByType[t].length; }, 0);
  Logger.log('findLowQualityCreatives: ' + rows.length + ' total ads in ad_creatives.');
  Logger.log('  ' + blank.length + ' ad(s) with NO thumbnail at all (renders as the placeholder icon)' +
    (blank.length ? ' — e.g. ' + blank.slice(0, 5).map(function (b) { return '"' + b.name + '" (' + b.type + ')'; }).join(', ') : '') + '.');
  Object.keys(stuckByType).sort(function (a, b) { return stuckByType[b].length - stuckByType[a].length; }).forEach(function (type) {
    var list = stuckByType[type];
    Logger.log('  ' + list.length + ' ' + type + ' ad(s) stuck on the lowest-tier fallback (creative.thumbnail_url — ' +
      'no native-res upgrade ever found) — e.g. ' + list.slice(0, 5).map(function (l) { return '"' + l.name + '"'; }).join(', '));
  });
  Logger.log('findLowQualityCreatives: ' + totalStuck + ' of ' + rows.length + ' ads (' +
    (rows.length ? (totalStuck / rows.length * 100).toFixed(1) : '0') + '%) on the lowest-tier fallback, plus ' +
    blank.length + ' with nothing at all. For VIDEO ads in the list above: run enrichVideoThumbnailCache() a few ' +
    'more times (576 were still uncached as of the last importAdCreatives() run) — that number should shrink each ' +
    'time you do. For IMAGE ads: these are very likely the same permanently-FAILED hashes ' +
    'findFailedImageHashSource() already explained (72 creative.image_hash + 15 object_story_spec.link_data + 8 ' +
    'asset_feed_spec, as of that run) — re-run that diagnostic if this list looks meaningfully different, since a ' +
    'growing count here (vs. that known baseline) would mean something NEW started failing, worth chasing again.');
}

// ONE-OFF DIAGNOSTIC (Sep 2026, "930 image ads stuck — is that the same 95 known FAILED
// hashes, or something bigger?") — findLowQualityCreatives() above can only see the
// ad_creatives SHEET; "Thumbnail URL === Fallback Thumbnail URL" is true for THREE very
// different underlying reasons that each need a completely different fix:
//   1. FAILED       — Meta's /adimages genuinely has nothing for this hash. Permanent —
//                      running enrichImageFullCache() again will NEVER change this one.
//   2. MISSING      — this hash has simply never been attempted (getImageFullMap_ only
//                      resolves a capped number of NEW hashes per run — see maxNewPerRun).
//                      Just needs enrichImageFullCache() (then importAdCreatives()) again.
//   3. STALE SHEET  — the hash is ALREADY resolved in the cache (a real URL, not FAILED) —
//                      importAdCreatives() simply hasn't been re-run since it resolved.
//                      Zero new Meta calls needed, just a re-sync.
// Video ads are deliberately excluded (they're stuck on a completely separate cache,
// ad_video_thumb_cache — mixing them in would muddy this hash-specific picture). Stops
// early once every stuck ad ID from the sheet has been matched, same as
// findFailedImageHashSource() above, so this is far cheaper than a full account scan
// whenever the stuck list is a minority of your ads.
function diagnoseImageBacklog() {
  var sheet = getOrCreateAdPlainSheet_(AD_CREATIVE_SHEET_);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('diagnoseImageBacklog: ad_creatives sheet is empty — run importAdCreatives() first.');
    return;
  }
  var rows = sheet.getRange(2, 1, lastRow - 1, AD_CREATIVE_HEADERS_.length).getValues();

  var stuckIds = {};
  var stuckCount = 0;
  rows.forEach(function (r) {
    var adId = r[0], type = r[5] || 'Unknown', thumb = r[6], fallback = r[9];
    if (type === 'Video') return; // separate cache — see header comment
    if (thumb && fallback && thumb === fallback) { stuckIds[String(adId)] = true; stuckCount++; }
  });
  if (!stuckCount) {
    Logger.log('diagnoseImageBacklog: no Image/Carousel ads currently stuck on the lowest-tier fallback — nothing to diagnose.');
    return;
  }
  Logger.log('diagnoseImageBacklog: ' + stuckCount + ' stuck Image/Carousel ad(s) in the sheet — tracing each one\'s image hash...');

  var cache = readImageFullCache_();
  var base = 'https://graph.facebook.com/v18.0/' + ACCOUNT_ID + '/ads';
  var params = {
    // Same field set + proven-safe limit:30 as findFailedImageHashSource() above.
    fields: 'id,creative{image_hash,object_story_spec,asset_feed_spec{images{hash}}}',
    limit: 30,
    'effective_status': JSON.stringify(['ACTIVE', 'PAUSED', 'CAMPAIGN_PAUSED', 'ADSET_PAUSED', 'IN_PROCESS', 'WITH_ISSUES', 'PENDING_REVIEW', 'DISAPPROVED'])
  };
  var qs = Object.keys(params).map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(String(params[k])); }).join('&');
  var next = base + '?' + qs, guard = 0;

  var failedAds = 0, missingAds = 0, staleSheetAds = 0, noHashAds = 0;
  var failedHashes = {}, missingHashes = {}, staleHashes = {};
  var matchedIds = {};

  do {
    var data = fetchWithRetry_(next);
    if (!data || data.error) {
      Logger.log('diagnoseImageBacklog: fetch failed partway — ' + (data && data.error ? JSON.stringify(data.error) : 'no response') +
        '. Counts below only cover what was matched before this point; re-run later to fill in the rest.');
      break;
    }
    (data.data || []).forEach(function (ad) {
      var id = String(ad.id);
      if (!stuckIds[id] || matchedIds[id]) return;
      matchedIds[id] = true;
      var h = extractImageHash_(ad.creative);
      if (!h) { noHashAds++; return; }
      var entry = cache[h];
      if (!entry) { missingAds++; missingHashes[h] = true; }
      else if (entry.value === 'FAILED') { failedAds++; failedHashes[h] = true; }
      else { staleSheetAds++; staleHashes[h] = true; }
    });
    if (Object.keys(matchedIds).length >= stuckCount) {
      Logger.log('diagnoseImageBacklog: matched all ' + stuckCount + ' stuck ad(s) — stopping early instead of scanning the rest of the account.');
      break;
    }
    next = (data.paging && data.paging.next) ? data.paging.next : null;
    if (next) Utilities.sleep(300);
    guard++;
  } while (next && guard < 500);

  var matchedCount = Object.keys(matchedIds).length;
  Logger.log('diagnoseImageBacklog: matched ' + matchedCount + ' of ' + stuckCount + ' stuck ad(s) to a creative. Breakdown:');
  Logger.log('  ' + staleSheetAds + ' ad(s) — ' + Object.keys(staleHashes).length + ' unique hash(es) — are ALREADY RESOLVED in the ' +
    'cache (a real URL, not FAILED). Nothing to fetch — just run importAdCreatives() again to write them into the sheet.');
  Logger.log('  ' + missingAds + ' ad(s) — ' + Object.keys(missingHashes).length + ' unique hash(es) — have NEVER been attempted ' +
    '(not in the cache at all yet). Run enrichImageFullCache() to resolve up to 1000 of these per pass, then importAdCreatives().');
  Logger.log('  ' + failedAds + ' ad(s) — ' + Object.keys(failedHashes).length + ' unique hash(es) — are PERMANENTLY FAILED. Meta\'s ' +
    '/adimages has nothing for these no matter how many more times enrichImageFullCache() runs; the only real fix is re-uploading ' +
    'or replacing that image in Ads Manager itself — this pipeline cannot conjure a URL Meta doesn\'t have.');
  if (noHashAds) Logger.log('  ' + noHashAds + ' ad(s) had no extractable image hash at all (Advantage+/dynamic creative gap — see extractImageHash_\'s header comment).');
  if (matchedCount < stuckCount) {
    Logger.log('diagnoseImageBacklog: ' + (stuckCount - matchedCount) + ' stuck ad(s) not matched — likely deleted/archived since the ' +
      'sheet was last written, or pagination did not reach them. Re-run if this count seems off.');
  }
}

// ONE-OFF DIAGNOSTIC (Aug 2026) — run this to find exactly which Page(s) are behind your
// remaining FAILED videos, the same way diagnoseVideoPageAttachment() found the first two
// ([VIDEO_ID_1], [VIDEO_ID_2]). Cross-references every video cached as FAILED
// against object_story_spec.page_id from a lightweight /ads pass (no thumbnail image
// processing, so a limit of 200 is safe here), and groups the failures by Page so you know
// which one(s) are worth going through the Business Settings → Assign people flow for
// again, and roughly how many videos each one would unblock — no point chasing a Page
// that's only holding back one or two ads.
function findMissingPageAccess() {
  var cache = readVideoThumbCache_();
  var failedVideoIds = {};
  Object.keys(cache).forEach(function (id) { if (cache[id].value === 'FAILED') failedVideoIds[id] = true; });
  var failedCount = Object.keys(failedVideoIds).length;
  if (!failedCount) {
    Logger.log('findMissingPageAccess: no FAILED videos in the cache right now — nothing to chase.');
    return;
  }
  Logger.log('findMissingPageAccess: looking up the Page behind ' + failedCount + ' FAILED video(s)...');

  var base = 'https://graph.facebook.com/v18.0/' + ACCOUNT_ID + '/ads';
  // FIXED (Aug 2026): first version used limit:200, assuming this field set was "cheap"
  // since it has no thumbnail_width/image-resizing cost. Wrong — object_story_spec itself
  // is a large nested blob (video_data, potential child_attachments, etc.), and 200 of
  // those per page hit the same "Please reduce the amount of data you're asking for" data-
  // volume ceiling importAdCreatives() hit earlier for a different reason (image
  // processing there, raw payload size here). Dropped to 30/page — same fix pattern,
  // different cause. If this still trips the same error, drop it further.
  // ADDED effective_object_story_id (Aug 2026): the raw dump for the "UNKNOWN" case came
  // back as literally {"video_id":..., "id":...} — object_story_spec wasn't just shaped
  // differently, it was genuinely absent. That means these creatives don't define their
  // Page relationship through object_story_spec at all — they very likely point at an
  // EXISTING Page post via effective_object_story_id instead, a separate AdCreative field
  // formatted as "{page_id}_{post_id}". Requesting it now and extracting page_id from the
  // prefix before the underscore, as a second fallback after object_story_spec.
  var params = {
    fields: 'id,creative{video_id,object_story_spec,effective_object_story_id}',
    limit: 30,
    'effective_status': JSON.stringify(['ACTIVE', 'PAUSED', 'CAMPAIGN_PAUSED', 'ADSET_PAUSED', 'IN_PROCESS', 'WITH_ISSUES', 'PENDING_REVIEW', 'DISAPPROVED'])
  };
  var qs = Object.keys(params).map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(String(params[k])); }).join('&');
  var next = base + '?' + qs, guard = 0;
  var pageCounts = {}; // page_id -> { count, exampleAdName }
  var matchedVideoIds = {};

  do {
    var data = fetchWithRetry_(next);
    if (!data || data.error) {
      Logger.log('findMissingPageAccess: fetch failed partway — ' + (data && data.error ? JSON.stringify(data.error) : 'no response') +
        '. Whatever was found before this point is still reported below; re-run later to fill in the rest.');
      break;
    }
    (data.data || []).forEach(function (ad) {
      var creative = ad.creative;
      var vid = creative && creative.video_id;
      if (!vid || !failedVideoIds[vid] || matchedVideoIds[vid]) return;
      var oss = creative.object_story_spec || {};
      var eosi = creative.effective_object_story_id || '';
      var pageId = oss.page_id || (oss.video_data && oss.video_data.page_id) ||
        (eosi.indexOf('_') > 0 ? eosi.split('_')[0] : null) || 'UNKNOWN';
      if (!pageCounts[pageId]) pageCounts[pageId] = { count: 0, exampleAdName: ad.name || '' };
      pageCounts[pageId].count++;
      matchedVideoIds[vid] = true;
      // For the very first UNKNOWN case, dump the raw creative JSON — rather than guess why
      // page_id wasn't found here (a likely culprit: Advantage+/dynamic creative ads use
      // asset_feed_spec instead of object_story_spec entirely, which wouldn't have page_id
      // in this spot), just look at the actual structure directly.
      if (pageId === 'UNKNOWN' && !pageCounts.UNKNOWN.rawSample) {
        pageCounts.UNKNOWN.rawSample = JSON.stringify(creative);
      }
    });
    // Early exit once every FAILED video has been matched — no reason to keep paginating
    // through the rest of the account (which could be thousands more ads) once the answer
    // is already complete. This also means a normal run finishes well before the 500-page
    // guard even at limit:30.
    if (Object.keys(matchedVideoIds).length >= failedCount) {
      Logger.log('findMissingPageAccess: matched all ' + failedCount + ' FAILED videos — stopping early instead of scanning the rest of the account.');
      break;
    }
    next = (data.paging && data.paging.next) ? data.paging.next : null;
    if (next) Utilities.sleep(300);
    guard++;
  } while (next && guard < 500);

  var pageIds = Object.keys(pageCounts).sort(function (a, b) { return pageCounts[b].count - pageCounts[a].count; });
  if (!pageIds.length) {
    Logger.log('findMissingPageAccess: matched 0 of ' + failedCount + ' FAILED videos to a Page — the pagination may not have reached ' +
      'the relevant ads yet, or those ads have since been deleted/archived. Try re-running, or check the ad_video_thumb_cache sheet ' +
      'for the specific video IDs and look them up individually if this keeps coming back empty.');
    return;
  }
  Logger.log('findMissingPageAccess: matched ' + Object.keys(matchedVideoIds).length + ' of ' + failedCount +
    ' FAILED videos to ' + pageIds.length + ' distinct Page(s), ranked by how many videos each is blocking:');
  pageIds.forEach(function (pageId) {
    var info = pageCounts[pageId];
    if (pageId === 'UNKNOWN') {
      Logger.log('  UNKNOWN: ' + info.count + ' video(s) had no page_id in object_story_spec at all — this is NOT a Page-access ' +
        'issue like the others, something structurally different about these creatives. Raw creative JSON for one example (e.g. "' +
        info.exampleAdName + '") follows so we can see why: ' + info.rawSample);
    } else {
      Logger.log('  Page ' + pageId + ': blocking ' + info.count + ' video(s) — e.g. "' + info.exampleAdName + '". ' +
        'Go to Business Settings → Accounts → Pages, search this ID, and assign yourself a role on it the same way you did for the first two.');
    }
  });
}

// ONE-OFF DIAGNOSTIC (Aug 2026), step 2 of re-authorizing with pages_read_engagement.
// diagnoseVideoPageAttachment() already confirmed your video ads ARE Page-attached (26/26
// in the sample, only 2 distinct Pages: [VIDEO_ID_1] and [VIDEO_ID_2]) — that's the
// premise pages_read_engagement needs to matter at all. This function is the next gate:
// run it AFTER you've generated a new ACCESS_TOKEN that includes pages_read_engagement and
// pasted it into importDataFromJSON.gs, BEFORE asking me to rebuild the full native-frame
// lookup into importAdCreatives(). It hits the exact same call that returned "(#10)
// Application does not have permission for this action" last time, for one real video from
// your account — if the new token clears it, the log shows the actual format/picture data
// and rebuilding the full feature is worth doing. If it still fails with the same code 10,
// the extra permission alone wasn't the missing piece (e.g. the token's user/system-user
// account itself needs a role added on those 2 specific Pages in Business Settings, not
// just the OAuth scope) — better to find that out on one test call than after a full rebuild.
function testVideoPermissionAfterReauth() {
  var testVideoId = '[PAGE_ID]'; // from "BOFU 1-081926-L1 - Coffee Moments", page_id [VIDEO_ID_1]
  var url = 'https://graph.facebook.com/v18.0/' + testVideoId + '?fields=format,picture';
  var data = fetchWithRetry_(url, 2);
  if (data && data.error) {
    Logger.log('testVideoPermissionAfterReauth: STILL BLOCKED — ' + JSON.stringify(data.error) +
      '. The extra OAuth scope alone did not clear it. Next thing to check in Business Settings: ' +
      'does the account/system user that generated this token actually have a role (admin/editor/ ' +
      'analyst) assigned on Page [VIDEO_ID_1] itself, separate from the ad account access it ' +
      'already has? The permission scope and the Page role are two different things Meta checks.');
  } else if (data) {
    Logger.log('testVideoPermissionAfterReauth: SUCCESS — permission wall is cleared. Raw response: ' + JSON.stringify(data));
    Logger.log('testVideoPermissionAfterReauth: send me this log — I\'ll rebuild the native-frame lookup into importAdCreatives() using this now that it actually works.');
  } else {
    Logger.log('testVideoPermissionAfterReauth: no response from fetchWithRetry_ — check ACCESS_TOKEN was actually saved/redeployed.');
  }
}

// Image / Video / Carousel — a HEURISTIC (like classifyAudienceType_ in AdSetPipeline.gs),
// not an official Meta label. object_type is often just "SHARE" for a lot of ad creatives,
// so this leans on video_id (definitive: present only for video creatives) and
// child_attachments.length>1 (definitive: only carousels have more than one) first.
// FIXED (Aug 2026, visible in your screenshot): the old fallback `return ot || 'Image'`
// let Meta's raw object_type value (most commonly "SHARE") leak straight through as the
// badge text on the card — "SHARE" isn't a creative format, it's just Meta's generic
// label for an ordinary link/photo post, and it read like a bug rather than an intentional
// label. Anything that isn't definitively Video or Carousel is, in practice for this
// account, a static image ad — so everything else now collapses to 'Image'.
function classifyCreativeType_(creative) {
  if (!creative) return 'Unknown';
  if (resolveVideoId_(creative)) return 'Video';
  var ot = String(creative.object_type || '').toUpperCase();
  if (ot === 'VIDEO') return 'Video';
  var linkData = (creative.object_story_spec && creative.object_story_spec.link_data) || {};
  if (linkData.child_attachments && linkData.child_attachments.length > 1) return 'Carousel';
  return 'Image';
}

// Sep 2026: an Advantage+/dynamic creative ad has no top-level creative.video_id even when
// its actual asset IS a video — the video lives at asset_feed_spec.videos[0].video_id
// instead. Every place in this file that reads "this ad's video_id" (creative type
// classification, the native-frame lookup, and the image-vs-video branch below) goes
// through this helper so none of them can drift out of sync on which shape they check.
// Only the first bundled video is used if there's more than one — same "good enough, not a
// full per-variant gallery" call as the carousel/image-hash cases elsewhere in this file.
function resolveVideoId_(creative) {
  if (!creative) return null;
  if (creative.video_id) return creative.video_id;
  var afs = creative.asset_feed_spec || {};
  if (afs.videos && afs.videos.length && afs.videos[0] && afs.videos[0].video_id) return afs.videos[0].video_id;
  return null;
}

// Deep-links straight to this ad inside Ads Manager — always works if the viewer has
// Ads Manager access to this account, regardless of the post's own sharing permissions
// (unlike a facebook.com/<page>_<post> permalink, which can 404 for a logged-out viewer
// or a deleted/unpublished page post). This is the "link out" half of "thumbnail + link".
function adsManagerLinkFor_(adId) {
  if (!adId) return '';
  var actNum = String(ACCOUNT_ID || '').replace(/^act_/, '');
  return 'https://adsmanager.facebook.com/adsmanager/manage/ads?act=' + actNum + '&selected_ad_ids=' + adId;
}

function getOrCreateAdPlainSheet_(sheetName) {
  var ss = getAdSpreadsheet_();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  return sheet;
}

// ONE-OFF DIAGNOSTIC (Aug 2026) — run this BEFORE spending any effort re-authorizing the
// token with pages_read_engagement. That permission only unlocks anything if these video
// ads are actually attached to a Facebook Page post (object_story_spec.page_id present) —
// if they're plain videos uploaded straight into the ad account's own video library with
// no Page attachment, pages_read_engagement changes nothing and the "(#10) Application
// does not have permission" wall stays exactly as it is no matter what you re-authorize.
// This costs ZERO new permissions — object_story_spec is already part of the field set
// importAdCreatives() successfully pulls today, so this just samples a few video ads and
// logs whether page_id shows up. Check the execution log after running it.
function diagnoseVideoPageAttachment() {
  var url = 'https://graph.facebook.com/v18.0/' + ACCOUNT_ID + '/ads' +
    '?fields=' + encodeURIComponent('id,name,creative{video_id,object_story_spec}') +
    '&limit=50';
  var data = fetchWithRetry_(url);
  if (!data || data.error) {
    Logger.log('diagnoseVideoPageAttachment: fetch failed — ' + (data && data.error ? JSON.stringify(data.error) : 'no response'));
    return;
  }
  var videoAds = (data.data || []).filter(function (ad) { return ad.creative && ad.creative.video_id; });
  if (!videoAds.length) {
    Logger.log('diagnoseVideoPageAttachment: none of the first 50 ads returned had a video_id — re-run with a larger limit or filter, or just check importAdCreatives()\'s ad_creatives sheet for a "Video" row and use that ad\'s ID directly.');
    return;
  }
  Logger.log('diagnoseVideoPageAttachment: found ' + videoAds.length + ' video ad(s) in this sample. Checking each for a Page attachment...');
  videoAds.forEach(function (ad) {
    var oss = ad.creative.object_story_spec || {};
    var pageId = oss.page_id || (oss.video_data && oss.video_data.page_id) || null;
    Logger.log('  Ad "' + ad.name + '" (video_id ' + ad.creative.video_id + '): page_id = ' +
      (pageId || 'ABSENT') + (pageId ? ' — Page-attached, pages_read_engagement COULD help here.' :
        ' — no Page attachment found. pages_read_engagement would NOT unlock this one; re-authorizing wouldn\'t change anything for it.'));
  });
  Logger.log('diagnoseVideoPageAttachment: if most/all show ABSENT, don\'t bother re-authorizing — it won\'t fix the thumbnail quality. If most/all show a real page_id, re-authorizing is at least worth trying.');
}

/**
 * DAILY TRIGGERS TO ADD (Apps Script editor → Triggers → Add Trigger), ~1–2am,
 * time-driven, day timer, running AFTER importDataFromJSON() and importAdSetData():
 *   - importAdData
 *   - importAdCreatives
 */