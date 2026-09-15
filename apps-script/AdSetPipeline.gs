/**
 * AdSetPipeline.gs — pulls AD SET level insights into "data - adset - by day".
 *
 * UPDATED (Aug 2026):
 *   - Mirrors DataPipeline.gs's new columns: Optimization goal, Landing page views,
 *     Initiate checkouts, ThruPlays. "Results" is now optimization-goal-aware here too
 *     (buildRowFromInsight_ is shared) — re-backfill after installing (see bottom).
 *   - NEW: importAdSetStatus() — a separate, CURRENT-STATE-ONLY daily snapshot of
 *     Learning Phase Status and Audience Type Identifier (Broad / Interest-Based /
 *     Lookalike (LAL) / Retargeting), written to a new "adset_status" sheet.
 *     Per your call: these two fields are NOT backfillable to Nov 13, 2025 — Meta's
 *     API only exposes the CURRENT learning-phase/targeting state, never a historical
 *     log of what it was on a past date. This sheet starts recording from whenever you
 *     first run it and has no earlier history. getAdSetData() (Code.gs) joins it onto
 *     every ad set row by name, so the dashboard always shows "current status," clearly
 *     the same value regardless of which past date you're looking at.
 *   - NEW: importAdSetBreakdownData() / backfillAdSetBreakdownHistory() — ad-set-level
 *     Placement and Age×Gender breakdowns, written to "data_adset_placement" and
 *     "data_adset_agegender". These power the two new toggleable widgets on the Ad Set
 *     tab (Performance by Placement, Performance by Age & Gender) and ARE fully
 *     backfillable to Nov 13, 2025 like everything else in this file.
 *
 * The daily sheet mirrors "data - by day" exactly, with an "Ad set name" column prepended
 * (36 columns now), so the dashboard's existing reader/aggregator works unchanged.
 *
 * Reuses from DataPipeline.gs: buildRowFromInsight_, INSIGHT_FIELDS_, fetchOptions_,
 * dayKey_, ensureColumns_, getYesterday, getDateNDaysAgo_, LOOKBACK_DAYS, ACCOUNT_ID, ACCESS_TOKEN.
 *
 * SETUP:
 *   1. Sheet "data - adset - by day" exists (you added it) — headers are written automatically.
 *   2. Run  backfillAdSetHistory()  once (month-by-month; safe to re-run — idempotent per range).
 *   3. Run  backfillAdSetBreakdownHistory()  once for the new placement/age-gender widgets.
 *   4. Add a daily trigger for  importAdSetData()  AND  importAdSetStatus()  AND
 *      importAdSetBreakdownData()  (~1–2am), same as your other jobs.
 *
 * ══════════════════════════════════════════════════════════════════════════
 * WORKBOOK SPLIT (Aug 2026): Ad Set data now lives in its OWN spreadsheet,
 * separate from the Campaign workbook, to stay under Sheets' 10M-cell limit.
 * Every function below that touches an ad-set sheet goes through
 * getAdSetSpreadsheet_() (openById(ADSET_SHEET_ID)) instead of
 * SpreadsheetApp.getActiveSpreadsheet(). getActiveSpreadsheet() would silently
 * resolve to whatever workbook this script is bound to (the Campaign one) —
 * that's the bug that "just connect the sheet ID" would NOT have caught.
 *
 * If you ever need to point this at a DIFFERENT ad-set workbook (e.g. after a
 * year rollover — see WorkbookRollover.gs), you do not need to edit this
 * constant by hand: createNewYearAdSetWorkbook() writes the new ID into
 * Script Properties, and getAdSetSpreadsheet_() checks there FIRST before
 * falling back to the hardcoded constant below. Editing the constant to match
 * is still a good idea for clarity, but not required for the code to work.
 * ══════════════════════════════════════════════════════════════════════════
 */

var ADSET_SHEET_ID = '[ADSET_SHEET_ID]'; // Ad Set workbook — do NOT put the Campaign ID here. See Config.example.gs.
var ADSET_SHEET_ID_PROP_ = 'ADSET_SHEET_ID_OVERRIDE'; // set by createNewYearAdSetWorkbook(), read here first

function getAdSetSpreadsheet_() {
  var override = PropertiesService.getScriptProperties().getProperty(ADSET_SHEET_ID_PROP_);
  var id = override || ADSET_SHEET_ID;
  if (!id || id === 'YOUR_ADSET_SPREADSHEET_ID') {
    throw new Error('ADSET_SHEET_ID is not set in AdSetPipeline.gs. Put the Ad Set workbook\'s ID there.');
  }
  var ss = SpreadsheetApp.openById(id);
  if (!ss) throw new Error('Could not open the Ad Set workbook (id ' + id + '). Check the ID and that this account can access it.');
  return ss;
}

/**
 * DIAGNOSTIC — run this straight from the editor to confirm the ad-set side
 * resolves to the RIGHT workbook (mirrors Code.gs's diagnose() for Campaign).
 */
function diagnoseAdSet() {
  var out = [];
  var override = PropertiesService.getScriptProperties().getProperty(ADSET_SHEET_ID_PROP_);
  out.push('ADSET_SHEET_ID constant: ' + ADSET_SHEET_ID);
  out.push('Script Properties override (' + ADSET_SHEET_ID_PROP_ + '): ' + (override || '(none set — using the constant)'));
  try {
    var ss = getAdSetSpreadsheet_();
    out.push('Resolved workbook: OK -> "' + ss.getName() + '"  [id ' + ss.getId() + ']');
    var tabs = ss.getSheets().map(function (s) { return '"' + s.getName() + '"'; });
    out.push('Tabs in that file: ' + tabs.join(', '));
  } catch (e) {
    out.push('FAILED to open ad-set workbook: ' + e.message);
  }
  var msg = out.join('\n');
  Logger.log(msg);
  return msg;
}

var ADSET_HEADERS_ = [
  'Ad set name', 'Campaign name', 'Objective', 'Attribution setting', 'Campaign bid strategy',
  'Reporting ends', 'Daily budget', 'Engagement rate ranking', 'Quality ranking', 'Frequency',
  'CPM (cost per 1,000 impressions)', 'Reach', 'Results', 'Blocked messaging contacts', 'Amount spent',
  'Video plays', 'Post engagements', 'CTR (all)', 'CTR (link click-through rate)', '3-second video plays',
  'Video plays at 25%', 'Video plays at 50%', 'Video plays at 75%', 'Video plays at 95%', 'Video plays at 100%',
  'Cost per messaging conversation started', 'Purchase ROAS (return on ad spend)', 'Purchases conversion value',
  'Purchases', 'Conversion rate ranking', 'Hook rate', 'Hold rate',
  'Optimization goal', 'Landing page views', 'Initiate checkouts', 'ThruPlays'
];

// ── DAILY (add a ~1–2am trigger): rolling window upsert ──
// FIXED (Aug 2026): this is why "data - adset - by day" went stale while
// data_adset_placement / data_adset_agegender kept updating fine. The old version
// cleared the [since, until] window FIRST, then wrote whatever fetchAdSetRows_ returned
// — so on any day the Meta API returned an error this file didn't retry (anything other
// than the one specific rate-limit code+subcode it checked for), the run finished
// "successfully" (no thrown exception — still shows 0% error rate in the Apps Script
// Executions dashboard) but silently WIPED that day's window and wrote nothing back.
// Do that on enough consecutive days and the table's tail just... stops, exactly like
// you found it. The breakdown sheets use a lighter field list and apparently didn't hit
// whatever this was, which is why only this table went stale.
// Now: nothing gets cleared or written unless the fetch actually succeeded. A failed day
// leaves existing rows untouched, and the rolling LOOKBACK_DAYS-day window automatically
// retries those same dates on every subsequent run — so this self-heals within
// LOOKBACK_DAYS days of the underlying issue clearing, with no manual action needed. If
// it's STILL failing after LOOKBACK_DAYS days in a row (meaning the window is about to
// move past the affected dates for good), you get an email instead of a silent gap.
var ADSET_MAIN_FAIL_STREAK_PROP_ = 'ADSET_MAIN_IMPORT_FAIL_STREAK';

// FIXED (Sep 2026, trigger error rate investigation): every failure path INSIDE this
// function was already handled gracefully (result.ok checks, fail-streak email alerts) —
// none of those throw, so none of them were the cause of the error rate showing up on the
// Apps Script Triggers page. That page counts actual UNCAUGHT exceptions/platform kills
// (execution timeout, a transient SpreadsheetApp/UrlFetchApp service error, quota limits),
// which this function had NO try/catch around at all — so whatever is actually causing
// that error rate has been failing invisibly: no log line, no email, nothing except a
// percentage on a page you have to go look at. This wrapper doesn't fix the underlying
// cause (we don't know what it is yet), but it makes the NEXT occurrence visible — same
// email-alert treatment as a graceful fetch failure gets — so it can actually be diagnosed
// instead of guessed at.
function importAdSetData() {
  try {
    checkYearRolloverWarning_('adset'); // see WorkbookRollover.gs — logs/emails a heads-up ~14 days before Jan 1
    var sheet = getAdSetSheet_();
    ensureAdSetHeaders_(sheet);
    var until = getYesterday();
    var since = getDateNDaysAgo_(LOOKBACK_DAYS);
    var result = fetchAdSetRows_(since, until);

    if (!result.ok) {
      var streak = bumpFailStreak_(ADSET_MAIN_FAIL_STREAK_PROP_);
      Logger.log('importAdSetData: fetch failed (' + result.errorMessage + ') for ' + since + '..' + until +
        ' — leaving existing rows untouched. Fail streak: ' + streak + ' day(s). Retries automatically tomorrow.');
      if (streak >= LOOKBACK_DAYS) alertAdSetImportFailing_(streak, result.errorMessage);
      return;
    }
    resetFailStreak_(ADSET_MAIN_FAIL_STREAK_PROP_);

    removeAdSetRowsInRange_(sheet, since, until);
    appendAdSetRowsSafe_(sheet, result.rows);
    Logger.log('importAdSetData: ' + result.rows.length + ' rows for ' + since + '..' + until);
    markPipelineRefreshed_('adsets');
  } catch (e) {
    Logger.log('⚠️ importAdSetData: UNCAUGHT exception — ' + (e && e.message ? e.message : String(e)) +
      '. This is what was showing up as a silent error rate on the Triggers page with no explanation. ' +
      'Existing rows are untouched (the write only happens after the try block above completes).');
    try {
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Meta Ads Dashboard: importAdSetData crashed',
        'importAdSetData() threw an uncaught exception: ' + (e && e.message ? e.message : String(e)) +
        '\n\nCheck View > Executions > importAdSetData in the Apps Script editor for the full stack trace.');
    } catch (mailErr) {
      Logger.log('(Could not send the crash email — Logger above still has it. Reason: ' + mailErr.message + ')');
    }
  }
}

function bumpFailStreak_(prop) {
  var props = PropertiesService.getScriptProperties();
  var n = (parseInt(props.getProperty(prop), 10) || 0) + 1;
  props.setProperty(prop, String(n));
  return n;
}
function resetFailStreak_(prop) {
  PropertiesService.getScriptProperties().deleteProperty(prop);
}

// Sep 2026, "let users know if dashboard data is fresh or stale" — shared by all 4 daily
// pipelines (Campaigns via copyDataOnceADay in DataPipeline.gs, Ad Sets here, Ads and
// Creatives in AdPipeline.gs). Call ONLY on a genuine, complete success — never on a
// self-healed skip or a partial/failed run — so the dashboard's freshness indicator answers
// "when did this data last actually change," not "when did some code last attempt to update
// it." Placed here (not in Code.gs) because this file already holds the shared fail-streak
// helpers every pipeline reuses — same pattern, same file.
var PIPELINE_FRESHNESS_PROP_PREFIX_ = 'PIPELINE_LAST_SUCCESS_';
function markPipelineRefreshed_(pipelineKey) {
  PropertiesService.getScriptProperties().setProperty(
    PIPELINE_FRESHNESS_PROP_PREFIX_ + pipelineKey,
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss")
  );
}

function alertAdSetImportFailing_(streak, errorMessage) {
  var msg = 'importAdSetData() has failed ' + streak + ' day(s) in a row (latest error: ' + errorMessage + '). ' +
    'The rolling ' + LOOKBACK_DAYS + '-day window means it can no longer self-heal the oldest affected dates — ' +
    'they will start permanently dropping out of "data - adset - by day" unless you fix the underlying issue and ' +
    'then run a manual backfill, e.g. backfillAdSetData(sinceStr, untilStr) for the affected range. Check ' +
    'View > Executions > importAdSetData in the Apps Script editor for the exact API error.';
  Logger.log('⚠️ ' + msg);
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Meta Ads Dashboard: Ad Set data import failing', msg);
  } catch (e) {
    Logger.log('(Could not send the email warning — Logger above still has it. Reason: ' + e.message + ')');
  }
}

// ── ONE-OFF BACKFILL (run once) ──
function backfillAdSetHistory() {
  var sheet = getAdSetSheet_();
  ensureAdSetHeaders_(sheet);
  var start = new Date('2025-11-13');
  var end = new Date(getYesterday());
  var cur = new Date(start.getFullYear(), start.getMonth(), 1);
  var monthCount = 0;
  while (cur <= end) {
    var mStart = new Date(Math.max(new Date(cur.getFullYear(), cur.getMonth(), 1).getTime(), start.getTime()));
    var mEnd = new Date(Math.min(new Date(cur.getFullYear(), cur.getMonth() + 1, 0).getTime(), end.getTime()));
    backfillAdSetData(dayKey_(mStart), dayKey_(mEnd));
    monthCount++;
    // Wait 5 seconds between months to avoid rate limits
    if (cur < end) Utilities.sleep(5000);
    cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
  }
  Logger.log('backfillAdSetHistory: done ' + monthCount + ' months from 2025-11-13 .. ' + dayKey_(end));
}

// FIXED (Aug 2026): same guard as importAdSetData() above — only clear+write this
// window if the fetch actually succeeded. On failure this leaves the window exactly as
// it was (matches "nothing happened" rather than "silently emptied"), logs clearly, and
// returns — safe to just call again once the underlying issue is fixed.
function backfillAdSetData(sinceStr, untilStr) {
  var sheet = getAdSetSheet_();
  ensureAdSetHeaders_(sheet);
  var result = fetchAdSetRows_(sinceStr, untilStr);
  if (!result.ok) {
    Logger.log('backfillAdSetData ' + sinceStr + '..' + untilStr + ': fetch failed (' + result.errorMessage +
      ') — existing rows in this window left untouched. Re-run once the underlying issue clears.');
    return;
  }
  removeAdSetRowsInRange_(sheet, sinceStr, untilStr);   // idempotent
  appendAdSetRowsSafe_(sheet, result.rows);
  Logger.log('backfillAdSetData ' + sinceStr + '..' + untilStr + ': ' + result.rows.length + ' rows');
}

// ── ONE-OFF: closes the Aug 12 -> yesterday gap in "data - adset - by day" caused by
// the silent-failure bug fixed above. Safe to re-run (idempotent per range) — if it
// doesn't fully close the gap on the first try (the underlying Meta API issue may still
// be intermittent), just run it again.
function fixAdSetMainTableGapNow() {
  backfillAdSetData('2026-08-12', getYesterday());
}

// ── FETCH (level=adset, daily, paginated) with retry on rate limit ──
// Returns { rows, ok, errorMessage } instead of a bare array (Aug 2026) — callers need
// to know whether an empty/partial result means "genuinely no data" or "the fetch broke
// partway," so they can avoid destructively overwriting good existing rows with nothing.
function fetchAdSetRows_(sinceStr, untilStr) {
  var settings = fetchAdSetSettings_();   // also uses retry internally
  var params = {
    level: 'adset',
    time_increment: 1,
    include_archived: true,
    fields: 'adset_name,' + INSIGHT_FIELDS_,
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
      Logger.log('adset insights error: ' + errorMessage);
      return { rows: out, ok: false, errorMessage: errorMessage };   // stop pagination, flag the whole fetch as failed
    }
    (data.data || []).forEach(function (item) { out.push(buildAdSetRow_(item, settings)); });
    next = (data.paging && data.paging.next) ? data.paging.next : null;
    if (next) Utilities.sleep(500);
  } while (next);
  return { rows: out, ok: true, errorMessage: '' };
}

// ── GENERIC RETRY WRAPPER for any Meta API call ──
// FIXED (Aug 2026): this used to only retry ONE exact error (code 4 / subcode 1504022 —
// one specific rate-limit variant). Any other transient Graph API error — code 2
// ("please retry your request later," a generic transient service hiccup) or a
// different rate/throttling subcode under codes 4/17 — fell straight through to the
// "non-retryable" branch on the very first attempt, got logged, and was returned to the
// caller as a hard failure. That's the most likely root cause of "data - adset - by day"
// going stale while the lighter-weight breakdown queries kept working. Now retries the
// whole family of known-transient codes with the same backoff; still fails fast (no
// point retrying) on genuine errors like an invalid/expired token or a bad parameter.
function isTransientMetaError_(err) {
  if (!err) return false;
  var code = err.code;
  return code === 1 || code === 2 || code === 4 || code === 17;
}

// Sep 2026, "There have been too many calls to this ad-account" crashing importAdCreatives()
// partway through: code 80004 (and its documented sibling 613) is Meta's AD-ACCOUNT-level rate
// limit — see https://developers.facebook.com/docs/graph-api/overview/rate-limiting#ads-management
// — a budget SHARED across every function hitting this account (importAdCreatives(),
// enrichImageFullCache(), enrichVideoThumbnailCache(), all of it), not a per-request hiccup.
// It fell through isTransientMetaError_ above (not in {1,2,4,17}) and was treated as a hard,
// undistinguished failure on the very first attempt — the caller's "leave the sheet untouched"
// fallback already made that safe, but the log gave no hint this was Meta saying "slow down"
// versus an actual bug, and nothing stopped someone from re-running immediately into the same wall.
// Deliberately kept OUT of isTransientMetaError_'s retry loop rather than just added to it: that
// function's 2s-doubling backoff (2s/4s/8s/16s/32s, ~1 minute total) is sized for a brief
// per-request blip, not an account-wide throttle Meta expects real minutes of wall-clock wait
// for — retrying fast against THIS error just burns more calls against an already-exhausted
// budget and eats into Apps Script's 6-minute ceiling for zero benefit. fetchWithRetry_ below
// checks this FIRST and fails fast instead.
function isAccountRateLimited_(err) {
  if (!err) return false;
  return err.code === 80004 || err.code === 613;
}

// FIXED (Aug 2026): the old version below threw away Meta's ACTUAL error text on every
// retryable failure and replaced it with a synthetic "Transient Meta API error (code X)"
// message instead. That's exactly why a genuinely broken, non-retryable request (a
// malformed `fields` query, which will fail identically forever) was indistinguishable
// in the log from a real transient rate-limit hiccup (which eventually succeeds on
// retry) — both just said "code 1" with no further detail, five times, then gave up.
// Now the ORIGINAL error object Meta sent — message, type, error_subcode,
// error_user_msg, fbtrace_id, whatever's actually there — is preserved all the way back
// to the caller instead of being masked. This doesn't change WHEN something retries
// (still governed by isTransientMetaError_ below), only what you can see once it stops.
function fetchWithRetry_(url, maxRetries) {
  maxRetries = maxRetries || 5;
  var wait = 2000; // start with 2 seconds
  for (var attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      var resp = UrlFetchApp.fetch(url, fetchOptions_());
      var data = JSON.parse(resp.getContentText());
      if (data.error && isAccountRateLimited_(data.error)) {
        Logger.log('fetchWithRetry_: Meta AD-ACCOUNT rate limit hit (code ' + data.error.code + ') — ' +
          (data.error.message || '') + '. This is a budget shared across every function that calls ' +
          'this account, including any enrichImageFullCache()/enrichVideoThumbnailCache() catch-up ' +
          'runs fired around the same time as this one — NOT retrying fast against it (that only ' +
          'burns more of an already-exhausted budget). Whatever this run collected so far is ' +
          'discarded by the caller\'s own safe fallback (existing sheet data is untouched); wait ' +
          'several minutes — Meta does not publish an exact reset time — before re-running, and ' +
          'avoid running the catch-up functions in the same window as a daily import.');
        return { error: data.error };
      }
      if (data.error && isTransientMetaError_(data.error)) {
        if (attempt === maxRetries) {
          Logger.log('fetchWithRetry_ giving up after ' + maxRetries + ' attempts. Meta\'s actual error: ' + JSON.stringify(data.error));
          return { error: data.error };
        }
        Logger.log('fetchWithRetry_ attempt ' + attempt + ' failed (' + JSON.stringify(data.error) + '), retrying in ' + wait + 'ms');
        Utilities.sleep(wait);
        wait = wait * 2; // exponential backoff
        continue;
      }
      return data; // success, or a non-transient error — either way, return it as-is, no retry
    } catch (e) {
      // Network error, malformed URL, or JSON parse failure — not a Meta API response at
      // all. Log and fail fast, no point retrying.
      Logger.log('fetchWithRetry_ non-retryable error: ' + e.message);
      return { error: { message: e.message } };
    }
  }
  return { error: { message: 'Max retries exceeded' } };
}

// One 36-col row: [Ad set name] + the same 35 columns your campaign builder produces.
function buildAdSetRow_(item, adsetSettings) {
  var base = buildRowFromInsight_(item, {});     // 35 cols (bid/budget blank because we pass {})
  var adset = item.adset_name || '';
  var st = adsetSettings[adset] || {};
  base[3] = st.bid_strategy || '';               // col 5 in the sheet = ad-set bid strategy
  base[5] = st.daily_budget || 0;                // col 7 in the sheet = ad-set daily budget
  return [adset].concat(base);
}

// Ad-set names → {bid_strategy, daily_budget}. Retry‑aware.
function fetchAdSetSettings_() {
  var map = {};
  var base = 'https://graph.facebook.com/v18.0/' + ACCOUNT_ID + '/adsets';
  var params = { fields: 'name,bid_strategy,daily_budget', limit: 200 };
  var qs = Object.keys(params).map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(String(params[k])); }).join('&');
  var next = base + '?' + qs, guard = 0;
  do {
    var data = fetchWithRetry_(next);
    if (!data || data.error) {
      if (data && data.error) Logger.log('adset settings error: ' + JSON.stringify(data.error));
      break;
    }
    (data.data || []).forEach(function (a) {
      map[a.name] = { bid_strategy: a.bid_strategy || '', daily_budget: a.daily_budget ? (parseFloat(a.daily_budget) || 0) : 0 };
    });
    next = (data.paging && data.paging.next) ? data.paging.next : null;
    if (next) Utilities.sleep(500);
    guard++;
  } while (next && guard < 50);
  return map;
}

// ═══════════════════════════════════════════════════════════════════════════
// ── LIGHTWEIGHT CTR (LINK) RECOMPUTE — patches ONLY the CTR(link) column ──
// Same pattern as recomputeCtrLinkOnly_CampaignLevel() in DataPipeline.gs: fetches just
// 5 fields (adset_name, campaign_name, date_stop, impressions, inline_link_clicks)
// instead of the full 26-field row rebuild, and overwrites ONLY the "CTR (link
// click-through rate)" cell (column 19 here — Ad set name shifts everything +1 vs. the
// Campaign sheet) for rows that already exist. Every other column is left untouched.
// Use this instead of backfillAdSetHistory() when CTR (link) is the only thing wrong.
// ═══════════════════════════════════════════════════════════════════════════

var CTR_LINK_COL_ADSET_ = 19; // 1-based column of "CTR (link click-through rate)" in "data - adset - by day"

function recomputeCtrLinkOnly_AdSetLevel(sinceStr, untilStr) {
  var sheet = getAdSetSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('recomputeCtrLinkOnly_AdSetLevel: no data rows.'); return; }
  var lastCol = sheet.getLastColumn();
  var all = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var sinceTs = new Date(sinceStr).getTime();
  var untilTs = new Date(untilStr).getTime() + 86399999;

  var ctrLinkVals = all.map(function (r) { return [r[CTR_LINK_COL_ADSET_ - 1]]; });
  var rowIndexByKey = {};
  all.forEach(function (r, i) {
    var dk = dayKey_(r[5]); // col 6 = Reporting ends
    var ts = new Date(dk).getTime();
    if (isNaN(ts) || ts < sinceTs || ts > untilTs) return;
    // key on ad set name + campaign name + date (an ad set name could theoretically
    // repeat under a different campaign — cheap insurance, matches how the row was written)
    rowIndexByKey[String(r[0]).trim() + '||' + String(r[1]).trim() + '||' + dk] = i;
  });

  var params = {
    level: 'adset',
    time_increment: 1,
    include_archived: true,
    fields: 'adset_name,campaign_name,date_stop,impressions,inline_link_clicks',
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
      if (data && data.error) Logger.log('recomputeCtrLinkOnly_AdSetLevel error: ' + JSON.stringify(data.error));
      break;
    }
    (data.data || []).forEach(function (item) {
      var key = String(item.adset_name || '').trim() + '||' + String(item.campaign_name || '').trim() + '||' + dayKey_(item.date_stop || '');
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

  sheet.getRange(2, CTR_LINK_COL_ADSET_, ctrLinkVals.length, 1).setValues(ctrLinkVals);
  Logger.log('recomputeCtrLinkOnly_AdSetLevel ' + sinceStr + '..' + untilStr + ': updated ' + touched + ' of ' + ctrLinkVals.length + ' rows.');
}

// ── ONE-OFF: run this once instead of backfillAdSetHistory() if CTR (link) is the only
// thing you need corrected. Same month-by-month, idempotent-per-window pattern.
function recomputeCtrLinkHistory_AdSetLevel() {
  var start = new Date('2025-11-13');
  var end = new Date(getYesterday());
  var cur = new Date(start.getFullYear(), start.getMonth(), 1);
  var months = 0;
  while (cur <= end) {
    var mStart = new Date(Math.max(new Date(cur.getFullYear(), cur.getMonth(), 1).getTime(), start.getTime()));
    var mEnd = new Date(Math.min(new Date(cur.getFullYear(), cur.getMonth() + 1, 0).getTime(), end.getTime()));
    recomputeCtrLinkOnly_AdSetLevel(dayKey_(mStart), dayKey_(mEnd));
    months++;
    if (cur < end) Utilities.sleep(3000);
    cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
  }
  Logger.log('recomputeCtrLinkHistory_AdSetLevel: done ' + months + ' months, 2025-11-13 .. ' + dayKey_(end));
}

// ═══════════════════════════════════════════════════════════════════════════
// ── LIGHTWEIGHT RESULTS RECOMPUTE — patches ONLY the "Results" column ──
// Same pattern as recomputeCtrLinkOnly_AdSetLevel() above and
// recomputeResultsOnly_CampaignLevel() in DataPipeline.gs. "Results" became
// optimization-goal-aware (see DataPipeline.gs header comment) — run this once so
// historical rows reflect the corrected numbers instead of the old purchases-only figure.
// ═══════════════════════════════════════════════════════════════════════════

var RESULTS_COL_ADSET_ = 13; // 1-based column of "Results" in "data - adset - by day" (12 + 1 prepended col)

function recomputeResultsOnly_AdSetLevel(sinceStr, untilStr) {
  var sheet = getAdSetSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('recomputeResultsOnly_AdSetLevel: no data rows.'); return; }
  var lastCol = sheet.getLastColumn();
  var all = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var sinceTs = new Date(sinceStr).getTime();
  var untilTs = new Date(untilStr).getTime() + 86399999;

  var resultsVals = all.map(function (r) { return [r[RESULTS_COL_ADSET_ - 1]]; });
  var rowIndexByKey = {};
  all.forEach(function (r, i) {
    var dk = dayKey_(r[5]); // col 6 = Reporting ends
    var ts = new Date(dk).getTime();
    if (isNaN(ts) || ts < sinceTs || ts > untilTs) return;
    rowIndexByKey[String(r[0]).trim() + '||' + String(r[1]).trim() + '||' + dk] = i;
  });

  var params = {
    level: 'adset',
    time_increment: 1,
    include_archived: true,
    fields: 'adset_name,campaign_name,date_stop,optimization_goal,actions,video_thruplay_watched_actions',
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
      if (data && data.error) Logger.log('recomputeResultsOnly_AdSetLevel error: ' + JSON.stringify(data.error));
      break;
    }
    (data.data || []).forEach(function (item) {
      var key = String(item.adset_name || '').trim() + '||' + String(item.campaign_name || '').trim() + '||' + dayKey_(item.date_stop || '');
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

  sheet.getRange(2, RESULTS_COL_ADSET_, resultsVals.length, 1).setValues(resultsVals);
  Logger.log('recomputeResultsOnly_AdSetLevel ' + sinceStr + '..' + untilStr + ': updated ' + touched + ' of ' + resultsVals.length + ' rows.');
}

function recomputeResultsHistory_AdSetLevel() {
  var start = new Date('2025-11-13');
  var end = new Date(getYesterday());
  var cur = new Date(start.getFullYear(), start.getMonth(), 1);
  var months = 0;
  while (cur <= end) {
    var mStart = new Date(Math.max(new Date(cur.getFullYear(), cur.getMonth(), 1).getTime(), start.getTime()));
    var mEnd = new Date(Math.min(new Date(cur.getFullYear(), cur.getMonth() + 1, 0).getTime(), end.getTime()));
    recomputeResultsOnly_AdSetLevel(dayKey_(mStart), dayKey_(mEnd));
    months++;
    if (cur < end) Utilities.sleep(3000);
    cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
  }
  Logger.log('recomputeResultsHistory_AdSetLevel: done ' + months + ' months, 2025-11-13 .. ' + dayKey_(end));
}

// ── SHEET HELPERS ──
function getAdSetSheet_() {
  var ss = getAdSetSpreadsheet_();
  var sh = ss.getSheetByName('data - adset - by day');
  if (!sh) sh = ss.insertSheet('data - adset - by day');
  return sh;
}

function ensureAdSetHeaders_(sheet) {
  ensureColumns_(sheet, ADSET_HEADERS_.length);
  var first = sheet.getRange(1, 1, 1, ADSET_HEADERS_.length).getValues()[0];
  var blank = first.every(function (c) { return String(c).trim() === ''; });
  if (blank) sheet.getRange(1, 1, 1, ADSET_HEADERS_.length).setValues([ADSET_HEADERS_]);
}

// Date lives in column 6 ("Reporting ends") in the ad-set layout (index 5).
function removeAdSetRowsInRange_(sheet, sinceStr, untilStr) {
  var last = sheet.getLastRow();
  if (last < 2) return;
  var lastCol = sheet.getLastColumn();
  var vals = sheet.getRange(2, 1, last - 1, lastCol).getValues();
  var s = new Date(sinceStr).getTime();
  var u = new Date(untilStr).getTime() + 86399999;
  var kept = vals.filter(function (r) {
    var ts = new Date(dayKey_(r[5])).getTime();
    if (isNaN(ts)) return true;
    return !(ts >= s && ts <= u);
  });
  sheet.getRange(2, 1, last - 1, lastCol).clearContent();
  if (kept.length) sheet.getRange(2, 1, kept.length, kept[0].length).setValues(kept);
}

function appendAdSetRowsSafe_(sheet, rows) {
  if (!rows || !rows.length) return;
  ensureColumns_(sheet, ADSET_HEADERS_.length);
  var start = sheet.getLastRow() + 1;
  var need = start + rows.length - 1;
  if (need > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), need - sheet.getMaxRows());
  sheet.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
}

function backfillRemainingAdSetData() {
  var start = new Date('2026-06-01');
  var end = new Date(getYesterday());
  var cur = new Date(start.getFullYear(), start.getMonth(), 1);
  var count = 0;
  while (cur <= end) {
    var mStart = new Date(cur.getFullYear(), cur.getMonth(), 1);
    var mEnd = new Date(cur.getFullYear(), cur.getMonth() + 1, 0);
    if (mEnd > end) mEnd = new Date(end);
    backfillAdSetData(dayKey_(mStart), dayKey_(mEnd));
    count++;
    // Wait 5 seconds between months to be kind to the rate limit
    if (cur < end) Utilities.sleep(5000);
    cur.setMonth(cur.getMonth() + 1);
  }
  Logger.log('backfillRemainingAdSetData: done ' + count + ' months (Jun ' + dayKey_(start) + ' .. ' + dayKey_(end) + ')');
}

// ═══════════════════════════════════════════════════════════════════════════
// ── NEW: LEARNING PHASE STATUS + AUDIENCE TYPE (current-state snapshot) ──
// Not part of Insights, not historical — see the big comment at the top of this
// file. Run importAdSetStatus() daily; there is nothing to backfill here.
// ═══════════════════════════════════════════════════════════════════════════

var ADSET_STATUS_SHEET_ = 'adset_status';
var ADSET_STATUS_HEADERS_ = ['Ad set name', 'Campaign name', 'Learning phase status', 'Audience type', 'Last updated'];

function importAdSetStatus() {
  var sheet = getOrCreateAdSetPlainSheet_(ADSET_STATUS_SHEET_);
  var first = sheet.getRange(1, 1, 1, ADSET_STATUS_HEADERS_.length).getValues()[0];
  if (first.every(function (c) { return String(c).trim() === ''; })) {
    sheet.getRange(1, 1, 1, ADSET_STATUS_HEADERS_.length).setValues([ADSET_STATUS_HEADERS_]);
  }

  var base = 'https://graph.facebook.com/v18.0/' + ACCOUNT_ID + '/adsets';
  var params = {
    fields: 'name,campaign{name},effective_status,learning_stage_info,targeting{custom_audiences,flexible_spec}',
    limit: 200,
    // Only currently-active-ish ad sets matter for "current status" — but include
    // paused ones too, since you may still want to see where they landed.
    'effective_status': JSON.stringify(['ACTIVE', 'PAUSED', 'CAMPAIGN_PAUSED', 'ADSET_PAUSED', 'IN_PROCESS', 'WITH_ISSUES'])
  };
  var qs = Object.keys(params).map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(String(params[k])); }).join('&');
  var next = base + '?' + qs, guard = 0;
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var rows = [];

  do {
    var data = fetchWithRetry_(next);
    if (!data || data.error) {
      if (data && data.error) Logger.log('importAdSetStatus error: ' + JSON.stringify(data.error));
      break;
    }
    (data.data || []).forEach(function (a) {
      rows.push([
        a.name || '',
        (a.campaign && a.campaign.name) || '',
        mapLearningStatus_(a.learning_stage_info),
        classifyAudienceType_(a.targeting),
        now
      ]);
    });
    next = (data.paging && data.paging.next) ? data.paging.next : null;
    if (next) Utilities.sleep(400);
    guard++;
  } while (next && guard < 100);

  // Full overwrite — this is a CURRENT snapshot, not a log. Old rows for ad sets that
  // no longer come back (deleted/fully archived) are intentionally dropped.
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, ADSET_STATUS_HEADERS_.length).clearContent();
  if (rows.length) {
    ensureColumns_(sheet, ADSET_STATUS_HEADERS_.length);
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  Logger.log('importAdSetStatus: snapshotted ' + rows.length + ' ad sets at ' + now);
}

function mapLearningStatus_(info) {
  if (!info || !info.status) return 'Unknown';
  switch (String(info.status).toUpperCase()) {
    case 'LEARNING': return 'Learning';
    case 'SUCCESS':  return 'Active (exited learning)';
    case 'FAIL':     return 'Learning limited';
    default:         return info.status;
  }
}

// Broad / Interest-Based / Lookalike (LAL) / Retargeting — inferred from the ad set's
// CURRENT targeting spec. This is a heuristic, not an official Meta label: adjust the
// custom_audiences subtype checks below if your naming conventions differ.
function classifyAudienceType_(targeting) {
  if (!targeting) return 'Unknown';
  var cas = targeting.custom_audiences || [];
  var hasLookalike = cas.some(function (c) {
    return (c.subtype && String(c.subtype).toUpperCase().indexOf('LOOKALIKE') !== -1) ||
           (c.name && /lookalike|\bLAL\b/i.test(c.name));
  });
  if (hasLookalike) return 'Lookalike (LAL)';

  var hasRetargeting = cas.length > 0; // any other custom audience = a saved/custom pool = retargeting
  if (hasRetargeting) return 'Retargeting';

  var flex = targeting.flexible_spec || [];
  var hasInterest = flex.some(function (spec) {
    return (spec.interests && spec.interests.length) ||
           (spec.behaviors && spec.behaviors.length) ||
           (spec.demographics && spec.demographics.length);
  });
  if (hasInterest) return 'Interest-Based';

  return 'Broad';
}

// ═══════════════════════════════════════════════════════════════════════════
// ── NEW: AD-SET-LEVEL PLACEMENT + AGE×GENDER BREAKDOWNS ──
// Powers the toggleable "Performance by Placement" and "Performance by Age & Gender"
// widgets on the Ad Set tab. Fully backfillable to Nov 13, 2025 — same pattern as
// BreakdownPipeline.gs, just at level:'adset' with adset_name carried through.
// ═══════════════════════════════════════════════════════════════════════════

var ADSET_PLACEMENT_SHEET_ = 'data_adset_placement';
var ADSET_AGEGENDER_SHEET_ = 'data_adset_agegender';
// Pulled out as named constants (Aug 2026) so WorkbookRollover.gs can pre-create these
// sheets with the EXACT same headers, instead of a second hand-typed copy that could
// silently drift out of sync with these.
var ADSET_PLACEMENT_HEADERS_ = ['Date', 'Campaign name', 'Ad set name', 'Placement', 'Spend', 'Impressions', 'Clicks', 'Reach', 'CTR', 'Purchases', 'Purchase value'];
var ADSET_AGEGENDER_HEADERS_ = ['Date', 'Campaign name', 'Ad set name', 'Age', 'Gender', 'Spend', 'Impressions', 'Clicks', 'Reach', 'Purchases', 'Purchase value'];

// FIXED (Aug 2026): this is why data_adset_placement / data_adset_agegender went stale
// at Aug 17 while "data - adset - by day" kept updating fine — the exact same disease
// that table had, just never patched here. importAdSetPlacementRange_/
// importAdSetAgeGenderRange_ used to clear the [since, until] window FIRST, then call
// fetchAdSetBreakdownRows_, which on any API error just `break`d out of its pagination
// loop and returned whatever partial (possibly empty) array it had — so a transient
// error finished "successfully" (no thrown exception) but silently wiped that window and
// wrote nothing back. Now: fetchAdSetBreakdownRows_ returns {rows, ok, errorMessage}
// (same shape as fetchAdSetRows_), and importAdSetPlacementRange_/
// importAdSetAgeGenderRange_ only clear+write if the fetch actually succeeded — a failed
// day leaves existing rows untouched and self-heals within LOOKBACK_DAYS once the
// underlying issue clears, same as the main table. Fail-streak email alerting added too,
// so a persistent failure surfaces instead of quietly aging out of the rolling window.
var ADSET_PLACEMENT_FAIL_STREAK_PROP_ = 'ADSET_PLACEMENT_IMPORT_FAIL_STREAK';
var ADSET_AGEGENDER_FAIL_STREAK_PROP_ = 'ADSET_AGEGENDER_IMPORT_FAIL_STREAK';

// This is the ONE daily trigger that does two full LOOKBACK_DAYS-window fetches back to
// back in a single run (placement, then age/gender) — twice the work of any other trigger
// in this project. If its error rate is meaningfully higher than the others (worth
// checking View > Executions to confirm), that's the prime suspect: it's the one most
// likely to occasionally run past Apps Script's 6-minute execution ceiling. If that's
// confirmed, the real fix is splitting this into two separate triggers
// (importAdSetPlacementDaily / importAdSetAgeGenderDaily below) so each gets its own full
// 6-minute budget instead of splitting one budget two ways — wired up but not switched to
// automatically, since that means changing which functions your Apps Script triggers point
// at, which is your call to make, not something to silently change out from under you.
function importAdSetBreakdownData() {
 try {
  var until = getYesterday();
  var since = getDateNDaysAgo_(LOOKBACK_DAYS);

  var pResult = importAdSetPlacementRange_(since, until);
  if (!pResult.ok) {
    var pStreak = bumpFailStreak_(ADSET_PLACEMENT_FAIL_STREAK_PROP_);
    Logger.log('importAdSetBreakdownData (placement): fetch failed (' + pResult.errorMessage + ') — fail streak ' + pStreak + ' day(s).');
    if (pStreak >= LOOKBACK_DAYS) alertAdSetBreakdownImportFailing_('Placement', 'placement', pStreak, pResult.errorMessage);
  } else {
    resetFailStreak_(ADSET_PLACEMENT_FAIL_STREAK_PROP_);
  }

  var aResult = importAdSetAgeGenderRange_(since, until);
  if (!aResult.ok) {
    var aStreak = bumpFailStreak_(ADSET_AGEGENDER_FAIL_STREAK_PROP_);
    Logger.log('importAdSetBreakdownData (age/gender): fetch failed (' + aResult.errorMessage + ') — fail streak ' + aStreak + ' day(s).');
    if (aStreak >= LOOKBACK_DAYS) alertAdSetBreakdownImportFailing_('Age & Gender', 'agegender', aStreak, aResult.errorMessage);
  } else {
    resetFailStreak_(ADSET_AGEGENDER_FAIL_STREAK_PROP_);
  }
 } catch (e) {
  Logger.log('⚠️ importAdSetBreakdownData: UNCAUGHT exception — ' + (e && e.message ? e.message : String(e)) +
    '. If this says "Exceeded maximum execution time," see the comment above this function — switch your daily ' +
    'trigger to call importAdSetPlacementDaily() and importAdSetAgeGenderDaily() separately instead of this combined function.');
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Meta Ads Dashboard: importAdSetBreakdownData crashed',
      'importAdSetBreakdownData() threw an uncaught exception: ' + (e && e.message ? e.message : String(e)) +
      '\n\nCheck View > Executions > importAdSetBreakdownData in the Apps Script editor for the full stack trace. ' +
      'If it mentions execution time, switch your daily trigger to call importAdSetPlacementDaily() and ' +
      'importAdSetAgeGenderDaily() separately instead — each then gets its own full 6-minute budget.');
  } catch (mailErr) {
    Logger.log('(Could not send the crash email — Logger above still has it. Reason: ' + mailErr.message + ')');
  }
 }
}

// Split-out versions of the two halves above, for when importAdSetBreakdownData() itself
// is timing out from doing both in one run — point two SEPARATE daily triggers at these
// instead of one trigger at importAdSetBreakdownData(), and each gets its own full
// execution-time budget. Safe to add without removing the combined function — nothing
// else in this project calls importAdSetBreakdownData() directly, so switching your
// trigger is the only change needed.
function importAdSetPlacementDaily() {
 try {
  var until = getYesterday();
  var since = getDateNDaysAgo_(LOOKBACK_DAYS);
  var pResult = importAdSetPlacementRange_(since, until);
  if (!pResult.ok) {
    var pStreak = bumpFailStreak_(ADSET_PLACEMENT_FAIL_STREAK_PROP_);
    Logger.log('importAdSetPlacementDaily: fetch failed (' + pResult.errorMessage + ') — fail streak ' + pStreak + ' day(s).');
    if (pStreak >= LOOKBACK_DAYS) alertAdSetBreakdownImportFailing_('Placement', 'placement', pStreak, pResult.errorMessage);
  } else {
    resetFailStreak_(ADSET_PLACEMENT_FAIL_STREAK_PROP_);
  }
 } catch (e) {
  Logger.log('⚠️ importAdSetPlacementDaily: UNCAUGHT exception — ' + (e && e.message ? e.message : String(e)));
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Meta Ads Dashboard: importAdSetPlacementDaily crashed',
      'importAdSetPlacementDaily() threw an uncaught exception: ' + (e && e.message ? e.message : String(e)) +
      '\n\nCheck View > Executions > importAdSetPlacementDaily in the Apps Script editor for the full stack trace.');
  } catch (mailErr) {
    Logger.log('(Could not send the crash email — Logger above still has it. Reason: ' + mailErr.message + ')');
  }
 }
}

function importAdSetAgeGenderDaily() {
 try {
  var until = getYesterday();
  var since = getDateNDaysAgo_(LOOKBACK_DAYS);
  var aResult = importAdSetAgeGenderRange_(since, until);
  if (!aResult.ok) {
    var aStreak = bumpFailStreak_(ADSET_AGEGENDER_FAIL_STREAK_PROP_);
    Logger.log('importAdSetAgeGenderDaily: fetch failed (' + aResult.errorMessage + ') — fail streak ' + aStreak + ' day(s).');
    if (aStreak >= LOOKBACK_DAYS) alertAdSetBreakdownImportFailing_('Age & Gender', 'agegender', aStreak, aResult.errorMessage);
  } else {
    resetFailStreak_(ADSET_AGEGENDER_FAIL_STREAK_PROP_);
  }
 } catch (e) {
  Logger.log('⚠️ importAdSetAgeGenderDaily: UNCAUGHT exception — ' + (e && e.message ? e.message : String(e)));
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Meta Ads Dashboard: importAdSetAgeGenderDaily crashed',
      'importAdSetAgeGenderDaily() threw an uncaught exception: ' + (e && e.message ? e.message : String(e)) +
      '\n\nCheck View > Executions > importAdSetAgeGenderDaily in the Apps Script editor for the full stack trace.');
  } catch (mailErr) {
    Logger.log('(Could not send the crash email — Logger above still has it. Reason: ' + mailErr.message + ')');
  }
 }
}

function alertAdSetBreakdownImportFailing_(label, kind, streak, errorMessage) {
  var sheetName = kind === 'placement' ? ADSET_PLACEMENT_SHEET_ : ADSET_AGEGENDER_SHEET_;
  var msg = 'importAdSetBreakdownData() — the ' + label + ' breakdown ("' + sheetName + '") has failed ' + streak +
    ' day(s) in a row (latest error: ' + errorMessage + '). The rolling ' + LOOKBACK_DAYS + '-day window means it ' +
    'can no longer self-heal the oldest affected dates — they will start permanently dropping out of "' + sheetName +
    '" unless you fix the underlying issue and then run a manual backfill, e.g. ' +
    'backfillAdSetOneBreakdownCustomRange(\'' + kind + '\', sinceStr, untilStr) for the affected range. Check ' +
    'View > Executions > importAdSetBreakdownData in the Apps Script editor for the exact API error.';
  Logger.log('⚠️ ' + msg);
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Meta Ads Dashboard: Ad Set ' + label + ' breakdown import failing', msg);
  } catch (e) {
    Logger.log('(Could not send the email warning — Logger above still has it. Reason: ' + e.message + ')');
  }
}

// ── ONE-OFF: closes the Aug 18 -> yesterday gap in data_adset_placement /
// data_adset_agegender caused by the silent-failure bug fixed above. Safe to re-run
// (idempotent per range) — if it doesn't fully close the gap on the first try (the
// underlying Meta API issue may still be intermittent), just run it again.
function fixAdSetBreakdownGapNow() {
  backfillAdSetBreakdownsCustomRange('2026-08-18', getYesterday());
}

function backfillAdSetBreakdownHistory() {
  var start = new Date('2025-11-13');
  var end = new Date(getYesterday());
  var cur = new Date(start.getFullYear(), start.getMonth(), 1);
  var months = 0;
  while (cur <= end) {
    var mStart = new Date(Math.max(new Date(cur.getFullYear(), cur.getMonth(), 1).getTime(), start.getTime()));
    var mEnd = new Date(Math.min(new Date(cur.getFullYear(), cur.getMonth() + 1, 0).getTime(), end.getTime()));
    importAdSetPlacementRange_(dayKey_(mStart), dayKey_(mEnd));
    Utilities.sleep(3000);
    importAdSetAgeGenderRange_(dayKey_(mStart), dayKey_(mEnd));
    months++;
    if (cur < end) Utilities.sleep(5000);
    cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
  }
  Logger.log('backfillAdSetBreakdownHistory: done ' + months + ' months, 2025-11-13 .. ' + dayKey_(end));
}

// ═══════════════════════════════════════════════════════════════════════════
// ── GENERAL-PURPOSE RESUME TOOLS (Aug 2026) ──
// Ad-set-level Placement and Age×Gender pull noticeably more rows than the
// campaign-level breakdowns (there are more ad sets than campaigns), so
// backfillAdSetBreakdownHistory() running Nov 13, 2025 → yesterday in one go is
// very likely to hit Apps Script's 6-minute execution limit partway through —
// same issue you already hit twice with BreakdownPipeline.gs. Rather than wait for
// that to happen and patch it after the fact, use these two directly, one month
// (or any range you choose) at a time, from the start:
//
//   function runThisMonth() { backfillAdSetBreakdownsCustomRange('2026-05-01', '2026-05-31'); }
//
// Pick runThisMonth from the Run dropdown, edit the two dates, run again for the
// next month. Both functions are idempotent per range — safe to re-run, safe to
// overlap dates, never duplicates rows. If even a full month times out for your
// account (large ad-set count), pass a smaller window instead — e.g. 2026-05-01
// to 2026-05-15, then 2026-05-16 to 2026-05-31. Nothing about these functions
// requires a calendar month; that's just a convenient chunk size to start with.
// ═══════════════════════════════════════════════════════════════════════════

// BOTH ad-set breakdowns (Placement + Age×Gender) for a range you choose.
function backfillAdSetBreakdownsCustomRange(sinceStr, untilStr) {
  importAdSetPlacementRange_(sinceStr, untilStr);
  Utilities.sleep(3000);
  importAdSetAgeGenderRange_(sinceStr, untilStr);
  Logger.log('backfillAdSetBreakdownsCustomRange: done ' + sinceStr + ' .. ' + untilStr);
}

// ONE ad-set breakdown only ('placement' or 'agegender') for a range you choose —
// use this if, like data_placement did twice at the campaign level, only one of
// the two falls behind and you don't want to re-touch the one that already finished.
function backfillAdSetOneBreakdownCustomRange(kind, sinceStr, untilStr) {
  if (kind === 'placement') { importAdSetPlacementRange_(sinceStr, untilStr); }
  else if (kind === 'agegender') { importAdSetAgeGenderRange_(sinceStr, untilStr); }
  else { Logger.log('Unknown kind: "' + kind + '" — use "placement" or "agegender".'); return; }
  Logger.log('backfillAdSetOneBreakdownCustomRange (' + kind + '): ' + sinceStr + ' .. ' + untilStr);
}

function importAdSetPlacementRange_(sinceStr, untilStr) {
  var sheet = getOrCreateAdSetBreakdownSheet_(ADSET_PLACEMENT_SHEET_, ADSET_PLACEMENT_HEADERS_);
  var result = fetchAdSetBreakdownRows_(sinceStr, untilStr, ['publisher_platform', 'platform_position'], true);
  if (!result.ok) {
    Logger.log('importAdSetPlacementRange_ ' + sinceStr + '..' + untilStr + ': fetch failed (' + result.errorMessage + ') — existing rows in this window left untouched.');
    return { ok: false, errorMessage: result.errorMessage };
  }
  removeAdSetBreakdownRowsInRange_(sheet, sinceStr, untilStr);
  appendRowsSafe_(sheet, result.rows);
  Logger.log('importAdSetPlacementRange_ ' + sinceStr + '..' + untilStr + ': ' + result.rows.length + ' rows');
  return { ok: true, errorMessage: '' };
}

function importAdSetAgeGenderRange_(sinceStr, untilStr) {
  var sheet = getOrCreateAdSetBreakdownSheet_(ADSET_AGEGENDER_SHEET_, ADSET_AGEGENDER_HEADERS_);
  var result = fetchAdSetBreakdownRows_(sinceStr, untilStr, ['age', 'gender'], false);
  if (!result.ok) {
    Logger.log('importAdSetAgeGenderRange_ ' + sinceStr + '..' + untilStr + ': fetch failed (' + result.errorMessage + ') — existing rows in this window left untouched.');
    return { ok: false, errorMessage: result.errorMessage };
  }
  removeAdSetBreakdownRowsInRange_(sheet, sinceStr, untilStr);
  appendRowsSafe_(sheet, result.rows);
  Logger.log('importAdSetAgeGenderRange_ ' + sinceStr + '..' + untilStr + ': ' + result.rows.length + ' rows');
  return { ok: true, errorMessage: '' };
}

// Returns { rows, ok, errorMessage } — same reasoning as fetchAdSetRows_: callers need to
// know whether an empty/partial result means "genuinely no data" or "the fetch broke
// partway," so a bad day never silently wipes good existing rows (see the FIXED note on
// importAdSetBreakdownData() above for the bug this replaces).
function fetchAdSetBreakdownRows_(sinceStr, untilStr, breakdownFields, isPlacement) {
  var params = {
    level: 'adset',
    time_increment: 1,
    include_archived: true,
    breakdowns: breakdownFields.join(','),
    fields: 'adset_name,campaign_name,spend,impressions,clicks,reach,ctr,date_start,actions,action_values',
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
      Logger.log('adset breakdown error (' + breakdownFields.join(',') + '): ' + errorMessage);
      return { rows: out, ok: false, errorMessage: errorMessage };   // stop pagination, flag the whole fetch as failed
    }
    (data.data || []).forEach(function (item) {
      var purchases = av_(item.actions, 'omni_purchase') + av_(item.actions, 'purchase');
      var purchaseValue = av_(item.action_values, 'omni_purchase') + av_(item.action_values, 'purchase');
      if (isPlacement) {
        var placement = (item.publisher_platform || 'unknown') + '_' + (item.platform_position || 'unknown');
        out.push([
          item.date_start || '', item.campaign_name || '', item.adset_name || '', placement,
          parseFloat(item.spend) || 0, parseInt(item.impressions) || 0, parseInt(item.clicks) || 0,
          parseInt(item.reach) || 0, parseFloat(item.ctr) || 0, purchases, purchaseValue
        ]);
      } else {
        out.push([
          item.date_start || '', item.campaign_name || '', item.adset_name || '',
          item.age || 'unknown', item.gender || 'unknown',
          parseFloat(item.spend) || 0, parseInt(item.impressions) || 0, parseInt(item.clicks) || 0,
          parseInt(item.reach) || 0, purchases, purchaseValue
        ]);
      }
    });
    next = (data.paging && data.paging.next) ? data.paging.next : null;
    if (next) Utilities.sleep(400);
  } while (next);
  return { rows: out, ok: true, errorMessage: '' };
}

function getOrCreateAdSetBreakdownSheet_(name, headers) {
  var ss = getAdSetSpreadsheet_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  ensureAdSetBreakdownHeaders_(sheet, headers);   // FIXED (Aug 2026) — see note below
  return sheet;
}

// FIXED (Aug 2026): the old version above only ever wrote headers at the moment the
// sheet was FIRST created. If the sheet already existed without a header row (e.g. it
// got created some other way, or row 1 was cleared/overwritten by hand), every daily
// import and backfill since then has been silently appending data starting at row 1 —
// which is exactly what you ran into on data_adset_placement / data_adset_agegender.
// This runs on EVERY call now, same self-healing pattern already used for the other
// sheets (ensureBreakdownHeaders_ in BreakdownPipeline.gs, ensureAdSetHeaders_ above):
//   - row 1 truly blank            -> write headers into it, nothing else to do.
//   - row 1 holds a real DATA row (starts with a date) -> insert a fresh row above it
//     first, THEN write headers there, so no existing data row gets overwritten.
//   - row 1 holds some other/old header text           -> just rewrite it in place.
function ensureAdSetBreakdownHeaders_(sheet, headers) {
  ensureColumns_(sheet, headers.length);
  var first = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var blank = first.every(function (c) { return String(c).trim() === ''; });
  var lastExpected = headers[headers.length - 1];
  var matches = !blank && String(first[headers.length - 1] || '').trim() === lastExpected;
  if (blank || matches) {
    if (blank) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }
  var col0 = String(first[0] || '').trim();
  var looksLikeDataRow = first[0] instanceof Date || /^\d{4}-\d{2}-\d{2}/.test(col0);
  if (looksLikeDataRow) sheet.insertRowBefore(1);   // protect the existing data row
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

// ── ONE-OFF: run this once from the editor to fix the two sheets you already have —
// inserts a proper header row above row 1 (your existing data rows are untouched and
// simply shift down by one), instead of you doing it by hand and risking a
// column/header mismatch. Safe to run more than once — a no-op once headers are correct.
function fixAdSetBreakdownHeadersNow() {
  getOrCreateAdSetBreakdownSheet_(ADSET_PLACEMENT_SHEET_, ADSET_PLACEMENT_HEADERS_);
  getOrCreateAdSetBreakdownSheet_(ADSET_AGEGENDER_SHEET_, ADSET_AGEGENDER_HEADERS_);
  Logger.log('Header check complete for data_adset_placement and data_adset_agegender.');
}

// Ad-set-workbook equivalent of importDataFromJSON.gs's getOrCreateSheet() — kept as
// its OWN function (not a shared one) specifically so nothing here ever touches the
// Campaign-workbook version. Only used by importAdSetStatus() right now, generic
// enough that anything else new on the ad-set side can reuse it too.
function getOrCreateAdSetPlainSheet_(sheetName) {
  var ss = getAdSetSpreadsheet_();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  return sheet;
}

function removeAdSetBreakdownRowsInRange_(sheet, sinceStr, untilStr) {
  var last = sheet.getLastRow();
  if (last < 2) return;
  var lastCol = sheet.getLastColumn();
  var vals = sheet.getRange(2, 1, last - 1, lastCol).getValues();
  var s = new Date(sinceStr).getTime();
  var u = new Date(untilStr).getTime() + 86399999;
  var kept = vals.filter(function (r) {
    var ts = new Date(dayKey_(r[0])).getTime();
    if (isNaN(ts)) return true;
    return !(ts >= s && ts <= u);
  });
  sheet.getRange(2, 1, last - 1, lastCol).clearContent();
  if (kept.length) sheet.getRange(2, 1, kept.length, kept[0].length).setValues(kept);
}

/**
 * DAILY TRIGGERS TO ADD (Apps Script editor → Triggers → Add Trigger), all ~1–2am,
 * time-driven, day timer, running AFTER importDataFromJSON():
 *   - importAdSetData
 *   - importAdSetStatus
 *   - importAdSetBreakdownData
 */