/**
 * Meta Ads Dashboard — Google Apps Script backend
 * Meta Ads Dashboard (portfolio copy)
 *
 * UPDATED (Aug 2026):
 *   - DELETE CustomReport.gs from this project entirely. It defined a SECOND
 *     getCustomReport() that silently competed with the one in this file — only one
 *     can ever run and you had no way of knowing which. This file's version (reading
 *     the pre-aggregated breakdown sheets) is now the only one, per your call: faster,
 *     matches your backfilled historical data, and doesn't cost a live Graph API call
 *     per report click. Region breakdown stays disabled (no data_region sheet exists —
 *     that's a separate follow-up if you ever want it, not part of this round).
 *   - getCustomReport()/getBreakdownData() now expose Purchases, Purchase value, ROAS,
 *     Video plays, Hook rate and Hold rate for every breakdown, reading the 12-column
 *     breakdown sheets from the updated BreakdownPipeline.gs.
 *   - New fields end-to-end: Optimization goal, Landing page views, Initiate checkouts,
 *     ThruPlays, plus derived: CPC, Cost Per Result (CPR), Cost Per Purchase (CPA),
 *     Average Order Value (AOV), LPV Rate / LPV Drop-off Rate, Cost per 3-sec video
 *     play, Cost per ThruPlay.
 *   - "Results" is now optimization-goal-aware upstream (DataPipeline.gs) — nothing to
 *     change here, aggregateByCampaign_ just sums whatever's already correct in the sheet.
 *   - NOTE ON THE "LPV DROP-OFF RATE" FORMULA: you specified
 *     (Landing Page Views / Link Clicks) * 100 and called it a "drop-off rate," but that
 *     formula is actually the RETENTION rate (% of clickers whose page loaded) — a real
 *     drop-off rate is the inverse. Both are exposed below: `lpvRate` (your formula,
 *     healthy = high, matches your ">80% is fine" rule of thumb) and `lpvDropoffRate`
 *     (100 - lpvRate, healthy = low, matches what the column NAME implies). The ad-set
 *     table shows lpvDropoffRate under that label so the number matches the name.
 *   - NEW: getAdSetData() now joins the "adset_status" sheet (current-state Learning
 *     Phase Status + Audience Type, written by importAdSetStatus() in AdSetPipeline.gs)
 *     onto every ad set by name. These two fields reflect TODAY's state regardless of
 *     which past date range you're viewing — there is no historical version of them
 *     (see AdSetPipeline.gs for why).
 *   - NEW: getAdSetBreakdownData() — backs the two new toggleable widgets on the Ad Set
 *     tab (Performance by Placement, Performance by Age & Gender).
 */

// ── CONFIG ──────────────────────────────────────────────────────────────────
const SHEET_ID   = '[SHEET_ID]'; // paste your Campaign workbook's Google Sheet ID here. See Config.example.gs.
const SHEET_NAME = 'data - by day';
const HEADER_ROW = 1; // your column names live on row 1

// ── ENTRY POINT ───────────────────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('Meta Ads Dashboard — Acme Co')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ── HEADER RESOLUTION ─────────────────────────────────────────────────────────
function norm(s) {
  return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, '');
}

const FIELD_ALIASES = {
  campaignName:    ['campaignname'],
  adSetName:       ['adsetname', 'adset', 'adsetName'],
  objective:       ['objective'],
  attribution:     ['attributionsetting'],
  bidStrategy:     ['campaignbidstrategy'],
  reportingEnds:   ['reportingends', 'reportingstarts', 'day', 'date'],
  dailyBudget:     ['dailybudget'],
  engagementRank:  ['engagementrateranking'],
  qualityRank:     ['qualityranking'],
  conversionRank:  ['conversionrateranking'],
  frequency:       ['frequency'],
  cpm:             ['cpm', 'cpmcostper1000impressions', 'costper1000impressions'],
  reach:           ['reach'],
  results:         ['results'],
  blockedContacts: ['blockedmessagingcontacts'],
  spent:           ['amountspent', 'amountspentphp', 'amountspentusd', 'spend'],
  videoPlays:      ['videoplays'],
  postEngagements: ['postengagements'],
  ctrAll:          ['ctrall'],
  ctrLink:         ['ctrlinkclickthroughrate', 'ctrlinkclickthrough', 'ctrlink', 'linkctr'],
  roas:            ['purchaseroasreturnonadspend', 'purchaseroas', 'roas', 'returnonadspend'],
  purchaseValue:   ['purchasesconversionvalue', 'purchaseconversionvalue', 'conversionvalue'],
  purchases:       ['purchases', 'websitepurchases'],
  costPerMessage:  ['costpermessagingconversationstarted'],
  video3s:         ['3secondvideoplays', 'threesecondvideoplays'],
  video25:         ['videoplaysat25'],
  video50:         ['videoplaysat50'],
  video75:         ['videoplaysat75'],
  video95:         ['videoplaysat95'],
  video100:        ['videoplaysat100'],
  // ── NEW ──
  optimizationGoal:  ['optimizationgoal'],
  landingPageViews:  ['landingpageviews'],
  initiateCheckouts: ['initiatecheckouts'],
  thruplays:         ['thruplays']
};

const REQUIRED_FIELDS = ['campaignName', 'spent', 'results', 'reach', 'purchases', 'roas'];

function resolveHeaders(headers) {
  const normHeaders = headers.map(norm);
  const idx = {};
  Object.keys(FIELD_ALIASES).forEach(function (field) {
    const accepts = FIELD_ALIASES[field];
    let found = -1;
    for (let i = 0; i < normHeaders.length; i++) {
      if (accepts.indexOf(normHeaders[i]) !== -1) { found = i; break; }
    }
    idx[field] = found;
  });
  return idx;
}

function findHeaderRow_(range) {
  var configured = (typeof HEADER_ROW === 'number' ? HEADER_ROW : 1) - 1;
  if (configured >= 0 && configured < range.length) {
    if (resolveHeaders(range[configured]).campaignName !== -1) return configured;
  }
  var best = -1, bestScore = 0;
  var scan = Math.min(range.length, 10);
  for (var i = 0; i < scan; i++) {
    var idx = resolveHeaders(range[i]);
    var score = REQUIRED_FIELDS.filter(function (f) { return idx[f] !== -1; }).length;
    if (score > bestScore) { bestScore = score; best = i; }
  }
  return best === -1 ? 0 : best;
}

// ── VALUE HELPERS ─────────────────────────────────────────────────────────────
function toNumber(v) {
  if (v === '' || v == null) return 0;
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function toDateString(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(v == null ? '' : v).trim();
  if (!s) return '';
  const m = s.match(/^(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : s;
}

function tsOf(dateStr) {
  if (!dateStr) return NaN;
  const t = new Date(dateStr).getTime();
  return isNaN(t) ? NaN : t;
}

function rankToPoints(v) {
  const s = norm(v);
  if (!s || s === 'unknown' || s === 'na') return null;
  if (s.indexOf('above') !== -1) return 3;
  if (s.indexOf('below') !== -1) return 1;
  if (s.indexOf('average') !== -1) return 2;
  return null;
}

function qualityScore(q, e, c) {
  const dims = { q: q, e: e, c: c };
  const hasData = (q !== null || e !== null || c !== null);
  if (!hasData) return { hasData: false, score10: null, percent: null, dims: dims };
  const pts = [q, e, c].map(function (x) { return x === null ? 2 : x; });
  const avg = (pts[0] + pts[1] + pts[2]) / 3;
  const score10 = Math.round(avg * 3.33 * 10) / 10;
  const percent = Math.round(avg * 3.33 * 10);
  return { hasData: true, score10: score10, percent: percent, dims: dims };
}

function derive(spent, cpm, ctrLink) {
  const impressions = cpm > 0 ? (spent / cpm) * 1000 : 0;
  const clicks = impressions * (ctrLink / 100);
  const cpc = clicks > 0 ? spent / clicks : null;
  return { impressions: impressions, clicks: clicks, cpc: cpc };
}

// ── MAIN DATA CALL ──────────────────────────────────────────────────────────
function getDashboardData(filters) {
  filters = filters || {};
  try {
    const sheet = getSheet_();
    const range = sheet.getDataRange().getValues();
    if (range.length < 2) {
      return okEmpty_('Sheet "' + sheet.getName() + '" has no data rows.');
    }

    const headerRowIdx = findHeaderRow_(range);
    const headers = range[headerRowIdx];
    const idx = resolveHeaders(headers);
    const missing = REQUIRED_FIELDS.filter(function (f) { return idx[f] === -1; });

    // Map every data row to a canonical record
    const rows = range.slice(headerRowIdx + 1).map(function (r) {
      const g = function (field) { return idx[field] === -1 ? '' : r[idx[field]]; };
      const spent = toNumber(g('spent'));
      const cpm = toNumber(g('cpm'));
      const ctrLink = toNumber(g('ctrLink'));
      const ctrAllRaw = toNumber(g('ctrAll'));
      const d = derive(spent, cpm, ctrLink);
      // FIXED (Aug 2026): "clicks" reconstructed above is LINK clicks (derived from the
      // now-corrected ctrLink). Reconstruct an ALL-clicks count the same way, from the
      // separately-stored ctrAll percentage, so aggregateByCampaign_/buildSummary_ can
      // compute a genuine weighted CTR (all) instead of silently reusing link clicks for it.
      const allClicks = d.impressions > 0 ? d.impressions * (ctrAllRaw / 100) : 0;
      const dateStr = toDateString(g('reportingEnds'));
      return {
        campaignName: String(g('campaignName') || '').trim() || '(unnamed)',
        objective: String(g('objective') || '').trim() || 'Unknown',
        bidStrategy: String(g('bidStrategy') || '').trim(),
        reportingEnds: dateStr,
        reportingEndsTs: tsOf(dateStr),
        spent: spent,
        results: toNumber(g('results')),
        reach: toNumber(g('reach')),
        purchases: toNumber(g('purchases')),
        purchaseValue: toNumber(g('purchaseValue')),
        videoPlays: toNumber(g('videoPlays')),
        postEngagements: toNumber(g('postEngagements')),
        ctrAll: ctrAllRaw,
        ctrLink: ctrLink,
        cpm: cpm,
        impressions: d.impressions,
        clicks: d.clicks,
        allClicks: allClicks,
        cpc: d.cpc,
        roasSheet: toNumber(g('roas')),
        dailyBudget: toNumber(g('dailyBudget')),
        frequency: toNumber(g('frequency')),
        blockedContacts: toNumber(g('blockedContacts')),
        costPerMessage: toNumber(g('costPerMessage')),
        video3s: toNumber(g('video3s')),
        video25: toNumber(g('video25')),
        video50: toNumber(g('video50')),
        video75: toNumber(g('video75')),
        video95: toNumber(g('video95')),
        video100: toNumber(g('video100')),
        qRank: rankToPoints(g('qualityRank')),
        eRank: rankToPoints(g('engagementRank')),
        cRank: rankToPoints(g('conversionRank')),
        // ── NEW ──
        optimizationGoal: String(g('optimizationGoal') || '').trim(),
        landingPageViews: toNumber(g('landingPageViews')),
        initiateCheckouts: toNumber(g('initiateCheckouts')),
        thruplays: toNumber(g('thruplays'))
      };
    }).filter(function (row) {
      return row.campaignName !== '(unnamed)' || row.spent || row.results || row.reach;
    });

    if (!rows.length) {
      return okEmpty_('Found headers on row ' + (headerRowIdx + 1) + ', but there are no data rows beneath them.');
    }

    // 1) row-level filters
    let filteredRows = rows.filter(function (row) {
      if (filters.search) {
        if (row.campaignName.toLowerCase().indexOf(String(filters.search).toLowerCase()) === -1) return false;
      }
      if (filters.objective) {
        if (row.objective !== filters.objective) return false;
      }
      if (filters.dateFrom || filters.dateTo) {
        const t = row.reportingEndsTs;
        if (isNaN(t)) return false;
        if (filters.dateFrom && t < new Date(filters.dateFrom).getTime()) return false;
        if (filters.dateTo && t > new Date(filters.dateTo).getTime() + 86399999) return false;
      }
      return true;
    });

    // 2) aggregate to campaigns
    let campaigns = aggregateByCampaign_(filteredRows);

    // 3) campaign-level filters (spend / roas ranges apply to totals)
    campaigns = campaigns.filter(function (c) {
      const sMin = filters.spentMin !== '' && filters.spentMin != null ? parseFloat(filters.spentMin) : -Infinity;
      const sMax = filters.spentMax !== '' && filters.spentMax != null ? parseFloat(filters.spentMax) : Infinity;
      const rMin = filters.roasMin !== '' && filters.roasMin != null ? parseFloat(filters.roasMin) : -Infinity;
      const rMax = filters.roasMax !== '' && filters.roasMax != null ? parseFloat(filters.roasMax) : Infinity;
      if (c.spent < sMin || c.spent > sMax) return false;
      if (c.roas < rMin || c.roas > rMax) return false;
      return true;
    });

    // keep only the rows belonging to surviving campaigns
    const survivors = {};
    campaigns.forEach(function (c) { survivors[c.campaignName] = true; });
    const finalRows = filteredRows.filter(function (row) { return survivors[row.campaignName]; });

    const summary = buildSummary_(finalRows, campaigns);

    return {
      rows: finalRows,
      campaigns: campaigns,
      summary: summary,
      resolved: idx,
      missing: missing,
      sheetName: sheet.getName(),
      generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
      // All campaign names present in the sheet (unfiltered) — feeds the Search Campaign
      // autocomplete dropdown on the frontend, independent of whatever filters are active.
      allCampaignNames: distinctCampaignNames_(rows),
      error: null
    };

  } catch (err) {
    return {
      rows: [], campaigns: [], summary: emptySummary_(),
      resolved: {}, missing: [], sheetName: '',
      generatedAt: '', allCampaignNames: [], error: err.message || String(err)
    };
  }
}

function distinctCampaignNames_(rows) {
  var set = {};
  rows.forEach(function (r) { if (r.campaignName && r.campaignName !== '(unnamed)') set[r.campaignName] = true; });
  return Object.keys(set).sort();
}

// ── AGGREGATION ─────────────────────────────────────────────────────────────
// statusMap (optional): { [name]: {learningPhaseStatus, audienceType} } — only passed
// by getAdSetData(), so campaign-level aggregation is completely unaffected.
function aggregateByCampaign_(rows, statusMap) {
  const map = {};
  rows.forEach(function (row) {
    const key = row.campaignName;
    if (!map[key]) {
      map[key] = {
        campaignName: key,
        objective: row.objective,
        bidStrategy: row.bidStrategy || '',
        dailyBudget: row.dailyBudget || 0,
        spent: 0,
        results: 0,
        reach: 0,
        purchases: 0,
        purchaseValue: 0,
        videoPlays: 0,
        video3s: 0,
        video25: 0,
        video50: 0,
        video75: 0,
        video95: 0,
        video100: 0,
        postEngagements: 0,
        blockedContacts: 0,
        costPerMessage: null,
        impressions: 0,
        clicks: 0,
        allClicks: 0,        // NEW (Aug 2026) — see CTR (all) fix note near ctrAll below
        // ── NEW ──
        landingPageViews: 0,
        initiateCheckouts: 0,
        thruplays: 0,
        optimizationGoal: '',
        latestTs: -Infinity,
        latestDate: '',
        _rows: []
      };
    }
    const a = map[key];
    a.spent += row.spent;
    a.results += row.results;
    a.reach += row.reach;
    a.purchases += row.purchases;
    a.purchaseValue += row.purchaseValue;
    a.videoPlays += row.videoPlays;
    a.video3s += row.video3s;
    a.video25 += row.video25;
    a.video50 += row.video50;
    a.video75 += row.video75;
    a.video95 += row.video95;
    a.video100 += row.video100;
    a.postEngagements += row.postEngagements;
    a.blockedContacts += row.blockedContacts;
    a.impressions += row.impressions;
    a.clicks += row.clicks;
    a.allClicks += (row.allClicks || 0);
    a.landingPageViews += (row.landingPageViews || 0);
    a.initiateCheckouts += (row.initiateCheckouts || 0);
    a.thruplays += (row.thruplays || 0);

    if (row.costPerMessage > 0) a.costPerMessage = row.costPerMessage;
    // Keep bidStrategy and dailyBudget (they are campaign-level, so take the first non-empty)
    if (row.bidStrategy && !a.bidStrategy) a.bidStrategy = row.bidStrategy;
    if (row.dailyBudget > 0 && a.dailyBudget === 0) a.dailyBudget = row.dailyBudget;
    if (row.optimizationGoal && !a.optimizationGoal) a.optimizationGoal = row.optimizationGoal;

    if (!row.objective || a.objective === 'Unknown') a.objective = row.objective || a.objective;
    if (!isNaN(row.reportingEndsTs) && row.reportingEndsTs > a.latestTs) {
      a.latestTs = row.reportingEndsTs;
      a.latestDate = row.reportingEnds;
    }
    a._rows.push(row);
  });

  return Object.keys(map).map(function (key) {
    const a = map[key];
    const byDate = a._rows.slice().sort(function (x, y) {
      return (x.reportingEndsTs || 0) - (y.reportingEndsTs || 0);
    });
    const latestRank = function (prop) {
      let val = null;
      byDate.forEach(function (r) { if (r[prop] !== null) val = r[prop]; });
      return val;
    };
    const q = latestRank('qRank'), e = latestRank('eRank'), c = latestRank('cRank');
    const roas = a.spent > 0 ? a.purchaseValue / a.spent : 0;
    const ctrLink = a.impressions > 0 ? (a.clicks / a.impressions) * 100 : 0;
    const cpm = a.impressions > 0 ? (a.spent / a.impressions) * 1000 : 0;
    const cpc = a.clicks > 0 ? a.spent / a.clicks : null;   // cost per LINK click (matches its "CPC (cost per link click)" label on the frontend)
    // FIXED (Aug 2026): this used to reuse a.clicks (link clicks) for "CTR (all)" too —
    // harmless before the CTR(link) fix (both were the same all-clicks number back then),
    // but now that a.clicks is genuinely link-only, this would have silently made
    // "CTR (all)" duplicate "CTR (link)". Uses the separately-tracked a.allClicks instead.
    const ctrAll = a.impressions > 0 ? (a.allClicks / a.impressions) * 100 : 0;
    const frequency = a.reach > 0 ? a.impressions / a.reach : 0;
    const hookRate = a.impressions > 0 ? (a.video3s / a.impressions) * 100 : 0;
    // FIXED (Sep 2026): Hold Rate is completions as a share of the people who were HOOKED
    // (3-sec plays), not of total impressions — see the matching fix + full explanation in
    // DataPipeline.gs's buildRowFromInsight_. Dividing by impressions here produced numbers
    // ~10x too low versus Ads Manager.
    // RE-APPLIED (Sep 2026): the numerator was ALSO still wrong here — your own correction
    // ("hold rate formula is ThruPlays / 3 second video plays") had been applied to
    // DataPipeline.gs but this file was still on video100 when re-checked just now. ThruPlays
    // (Meta's 15-second-or-completion threshold) reads meaningfully HIGHER than literal
    // 100%-completion for any video over ~15 seconds.
    const holdRate = a.video3s > 0 ? (a.thruplays / a.video3s) * 100 : 0;

    // ── NEW DERIVED METRICS ──
    const cpr = a.results > 0 ? a.spent / a.results : null;                 // Cost Per Result
    const cpa = a.purchases > 0 ? a.spent / a.purchases : null;             // Cost Per Purchase
    const aov = a.purchases > 0 ? a.purchaseValue / a.purchases : null;     // Average Order Value
    const lpvRate = a.clicks > 0 ? (a.landingPageViews / a.clicks) * 100 : null;         // your formula (healthy = high)
    const lpvDropoffRate = lpvRate == null ? null : Math.max(0, 100 - lpvRate);          // matches the column's NAME (healthy = low)
    const costPer3sPlay = a.video3s > 0 ? a.spent / a.video3s : null;
    const costPerThruplay = a.thruplays > 0 ? a.spent / a.thruplays : null;

    const status = (statusMap && statusMap[key]) || {};

    return {
      campaignName: a.campaignName,
      objective: a.objective,
      bidStrategy: a.bidStrategy,
      dailyBudget: a.dailyBudget,
      reportingEnds: a.latestDate,
      spent: a.spent,
      results: a.results,
      reach: a.reach,
      purchases: a.purchases,
      purchaseValue: a.purchaseValue,
      videoPlays: a.videoPlays,
      impressions: a.impressions,
      clicks: a.clicks,
      ctrLink: ctrLink,
      cpm: cpm,
      cpc: cpc,
      ctrAll: ctrAll,
      frequency: frequency,
      hookRate: hookRate,
      holdRate: holdRate,
      v3s: a.video3s,
      v25: a.video25,
      v50: a.video50,
      v75: a.video75,
      v95: a.video95,
      v100: a.video100,
      costPerMsg: a.costPerMessage,
      blockedContacts: a.blockedContacts,
      postEngagements: a.postEngagements,
      roas: roas,
      eRankRaw: tierLabel(e),
      qRankRaw: tierLabel(q),
      cRankRaw: tierLabel(c),
      quality: qualityScore(q, e, c),
      // ── NEW ──
      optimizationGoal: a.optimizationGoal,
      landingPageViews: a.landingPageViews,
      initiateCheckouts: a.initiateCheckouts,
      thruplays: a.thruplays,
      cpr: cpr,
      cpa: cpa,
      aov: aov,
      lpvRate: lpvRate,
      lpvDropoffRate: lpvDropoffRate,
      costPer3sPlay: costPer3sPlay,
      costPerThruplay: costPerThruplay,
      learningPhaseStatus: status.learningPhaseStatus || '',
      audienceType: status.audienceType || ''
    };
  });
}

// ── SUMMARY ─────────────────────────────────────────────────────────────────
function buildSummary_(rows, campaigns) {
  const s = emptySummary_();
  s.totalCampaigns = campaigns.length;
  rows.forEach(function (r) {
    s.totalSpent += r.spent;
    s.totalResults += r.results;
    s.totalReach += r.reach;
    s.totalPurchases += r.purchases;
    s.totalPurchaseValue += r.purchaseValue;
    s.totalVideoPlays += r.videoPlays;
    s.totalVideo3s += r.video3s;
    s.totalVideo100 += r.video100;
    s.totalThruplays += (r.thruplays || 0);
    s.totalImpressions += r.impressions;
    s.totalClicks += r.clicks;
    s.totalAllClicks += (r.allClicks || 0);
    s.totalPostEngagements += r.postEngagements;
    if (r.spent > 0) s.activeCampaigns++;
  });
  s.roas = s.totalSpent > 0 ? s.totalPurchaseValue / s.totalSpent : 0;
  s.ctrLink = s.totalImpressions > 0 ? (s.totalClicks / s.totalImpressions) * 100 : 0;
  // FIXED (Aug 2026): was recomputed from totalClicks (link clicks) same as ctrLink above —
  // see the matching fix in aggregateByCampaign_. Now a genuine all-clicks CTR.
  s.ctrAll = s.totalImpressions > 0 ? (s.totalAllClicks / s.totalImpressions) * 100 : 0;
  s.frequency = s.totalReach > 0 ? s.totalImpressions / s.totalReach : 0;
  s.hookRate = s.totalImpressions > 0 ? (s.totalVideo3s / s.totalImpressions) * 100 : 0;
  // FIXED (Sep 2026) — same Hold Rate denominator fix as aggregateByCampaign_ above.
  // RE-APPLIED (Sep 2026): numerator fix (ThruPlays, not Video plays at 100%) re-applied
  // here too — see the matching comment on aggregateByCampaign_'s holdRate above.
  s.holdRate = s.totalVideo3s > 0 ? (s.totalThruplays / s.totalVideo3s) * 100 : 0;

  let wSum = 0, wTot = 0;
  campaigns.forEach(function (c) {
    if (c.quality.hasData && c.spent > 0) { wSum += c.quality.score10 * c.spent; wTot += c.spent; }
  });
  const accScore = wTot > 0 ? Math.round(wSum / wTot * 10) / 10 : null;

  const dimAvg = function (prop) {
    let ps = 0, pt = 0;
    campaigns.forEach(function (c) {
      const p = c.quality.dims[prop];
      if (p !== null && c.spent > 0) { ps += p * c.spent; pt += c.spent; }
    });
    return pt > 0 ? ps / pt : null;
  };
  s.accountQuality = {
    score10: accScore,
    dims: { q: dimAvg('q'), e: dimAvg('e'), c: dimAvg('c') }
  };
  return s;
}

function emptySummary_() {
  return {
    totalSpent: 0, totalResults: 0, totalReach: 0, totalPurchases: 0,
    totalPurchaseValue: 0, totalVideoPlays: 0, totalVideo3s: 0, totalVideo100: 0, totalThruplays: 0,
    totalImpressions: 0, totalClicks: 0, totalAllClicks: 0, totalPostEngagements: 0,
    totalCampaigns: 0, activeCampaigns: 0,
    roas: 0, ctrLink: 0, ctrAll: 0, frequency: 0,
    hookRate: 0, holdRate: 0,
    accountQuality: { score10: null, dims: { q: null, e: null, c: null } }
  };
}

function okEmpty_(msg) {
  return {
    rows: [], campaigns: [], summary: emptySummary_(),
    resolved: {}, missing: [], sheetName: '', generatedAt: '', allCampaignNames: [], error: msg
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// ── PIPELINE FRESHNESS (Sep 2026, "make sure users can distinguish accurate from
// inaccurate data") ── This is deliberately NOT the same thing as getDashboardData()'s own
// `generatedAt` field: that timestamp is just "when this function happened to read the
// sheet," which is effectively always "just now" no matter how stale the underlying import
// actually is — not useful for answering "can I trust this." This reads the REAL last-
// success timestamps each pipeline stamps via markPipelineRefreshed_ (in AdSetPipeline.gs,
// shared by all 4 daily pipelines: Campaigns/Ad Sets/Ads/Creatives), plus each pipeline's
// active fail-streak where one is tracked, so the dashboard can show an honest "data as of"
// per tab instead of a number that always looks fresh.
// ═══════════════════════════════════════════════════════════════════════════
function getPipelineFreshness() {
  var props = PropertiesService.getScriptProperties();
  function entry(key, failStreakProp) {
    var last = props.getProperty(PIPELINE_FRESHNESS_PROP_PREFIX_ + key) || null;
    // null = "this pipeline has no fail-streak counter at all" (Campaigns, Creatives) —
    // distinct from 0, which means "tracked, and currently healthy."
    var failStreak = failStreakProp ? (parseInt(props.getProperty(failStreakProp), 10) || 0) : null;
    return { lastSuccess: last, failStreak: failStreak };
  }
  return {
    campaigns: entry('campaigns', null),
    adsets:    entry('adsets', ADSET_MAIN_FAIL_STREAK_PROP_),
    ads:       entry('ads', AD_MAIN_FAIL_STREAK_PROP_),
    creatives: entry('creatives', null)
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// ── "TODAY (LIVE)" CAMPAIGN-LEVEL SNAPSHOT (Sep 2026) ── Deliberately narrow scope, agreed
// on explicitly rather than the original "call Ads Manager live for the whole dashboard"
// idea: campaign-level only, numbers only (spend/impressions/CTR link/purchases) — NO
// per-ad or per-ad-set breakdown, and absolutely no creative/thumbnail work, since that's
// the expensive, rate-limit-fragile part this whole project has spent so much effort
// taming. This stays cheap enough to run on a live page load.
//
// The part that actually keeps this safe: CacheService.getScriptCache() is a SINGLE cache
// shared by every viewer of this dashboard, not one per person. So no matter how many
// people have the dashboard open, or how often any one of them refreshes, Meta gets called
// at most once per TODAY_LIVE_CACHE_TTL_SECONDS_ (5 minutes, per your call) — a fixed,
// predictable ceiling, not something that scales with viewer count the way the original
// "call Ads Manager live per page load" idea would have.
// ═══════════════════════════════════════════════════════════════════════════
var TODAY_LIVE_CACHE_KEY_ = 'TODAY_LIVE_SNAPSHOT_V1';
var TODAY_LIVE_CACHE_TTL_SECONDS_ = 300; // 5 minutes

function getTodayLiveSnapshot() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(TODAY_LIVE_CACHE_KEY_);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      // corrupt/unparseable cache entry — fall through and refetch rather than error out
    }
  }

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var campResult = fetchTodayLiveFromMeta_(today);
  // Ad-level pulse (Sep 2026, "most needed metrics are creative level"): numbers only —
  // spend/impressions/CTR/purchases per ad — no live thumbnail fetch. See
  // fetchTodayLiveAdsFromMeta_ for why this stays cheap and does NOT reopen the
  // rate-limit/timeout risk this whole project has spent this session fixing.
  var adResult = fetchTodayLiveAdsFromMeta_(today);

  var result = {
    ok: !!campResult.ok,
    error: campResult.ok ? null : (campResult.errorMessage || 'Could not refresh live campaign data right now.'),
    campaigns: campResult.ok ? campResult.campaigns : [],
    asOf: campResult.ok ? campResult.asOf : null,
    adsOk: !!adResult.ok,
    adsError: adResult.ok ? null : (adResult.errorMessage || 'Could not refresh live ad data right now.'),
    ads: adResult.ok ? adResult.ads : []
  };

  // Only cache a genuine campaign-level success — a failed ad-level pulse alone shouldn't
  // force everyone to refetch both for the next 5 minutes, but a total failure must NOT be
  // cached, so the next viewer gets a fresh attempt instead of a frozen error.
  if (result.ok) {
    cache.put(TODAY_LIVE_CACHE_KEY_, JSON.stringify(result), TODAY_LIVE_CACHE_TTL_SECONDS_);
    return result;
  }

  // Meta call failed (rate limit, transient hiccup, etc.) — same self-healing philosophy as
  // the rest of this project: don't error the widget out. CacheService entries vanish once
  // expired though (there's no "stale but still there" read available), so there's genuinely
  // nothing to fall back to here beyond saying so honestly.
  return {
    ok: false,
    error: result.error,
    campaigns: [],
    asOf: null,
    adsOk: false,
    adsError: result.adsError,
    ads: []
  };
}

function fetchTodayLiveFromMeta_(dateStr) {
  var params = {
    level: 'campaign',
    time_range: JSON.stringify({ since: dateStr, until: dateStr }),
    fields: 'campaign_name,spend,impressions,inline_link_clicks,actions',
    limit: 200,
    // ACTIVE only — a paused campaign isn't spending today, and keeping this narrow keeps
    // the call cheap. This is a live PULSE, not a replacement for the full tables above.
    'effective_status': JSON.stringify(['ACTIVE'])
  };
  var qs = Object.keys(params).map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(String(params[k])); }).join('&');
  var url = 'https://graph.facebook.com/v18.0/' + ACCOUNT_ID + '/insights?' + qs;

  // Only 2 retries (not the default 5) — this is a "nice to have" live pulse with its own
  // graceful fallback above, not worth burning a long backoff on.
  var data = fetchWithRetry_(url, 2);
  if (!data || data.error) {
    return { ok: false, errorMessage: data && data.error ? JSON.stringify(data.error) : 'no response from fetchWithRetry_' };
  }

  var campaigns = (data.data || []).map(function (item) {
    var spend = parseFloat(item.spend) || 0;
    var impressions = parseInt(item.impressions) || 0;
    var linkClicks = parseInt(item.inline_link_clicks) || 0;
    var purchases = av_(item.actions, 'omni_purchase') + av_(item.actions, 'purchase');
    return {
      name: item.campaign_name || '(unnamed)',
      spend: spend,
      impressions: impressions,
      ctrLink: impressions > 0 ? (linkClicks / impressions) * 100 : 0,
      purchases: purchases
    };
  }).sort(function (a, b) { return b.spend - a.spend; });

  return {
    ok: true,
    campaigns: campaigns,
    asOf: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    error: null
  };
}

// ── "TODAY (LIVE)" AD-LEVEL PULSE (Sep 2026) ── Added after you flagged that the metrics
// you actually need most live at the CREATIVE level, not just campaign. Stress-tested
// before building: a full live creative view (fetching each ad's image/thumbnail on every
// cache miss) would reopen the exact rate-limit/timeout problem this whole session has been
// fixing — so this deliberately fetches NUMBERS ONLY from /insights (spend, impressions,
// CTR, purchases per ad), same cheap shape as the campaign pulse above, and borrows a
// thumbnail ONLY if importAdCreatives() has already resolved one into the ad_creatives
// sheet (getAdCreativeMap_ — a plain sheet read, zero extra Meta calls). A brand-new ad
// that hasn't been picked up by the self-healing pipeline yet will show with no thumbnail
// rather than trigger a live image fetch.
function fetchTodayLiveAdsFromMeta_(dateStr) {
  var fields = 'ad_id,ad_name,spend,impressions,inline_link_clicks,actions';
  var baseUrl = 'https://graph.facebook.com/v18.0/' + ACCOUNT_ID + '/insights?' +
    'level=ad&time_range=' + encodeURIComponent(JSON.stringify({ since: dateStr, until: dateStr })) +
    '&fields=' + encodeURIComponent(fields) +
    '&limit=100' +
    '&effective_status=' + encodeURIComponent(JSON.stringify(['ACTIVE']));

  var all = [];
  var next = baseUrl;
  var pages = 0;
  // Bounded at 3 pages (≤300 ad rows) — this is a live PULSE ("what's winning today"), not
  // an exhaustive report. Only ads with actual activity today return a row at all, so in
  // practice this is far smaller than the account's ~3000 total ads.
  while (next && pages < 3) {
    var data = fetchWithRetry_(next, 2);
    if (!data || data.error) {
      if (all.length) break; // keep whatever partial results we already have rather than discard them
      return { ok: false, errorMessage: data && data.error ? JSON.stringify(data.error) : 'no response from fetchWithRetry_' };
    }
    all = all.concat(data.data || []);
    next = (data.paging && data.paging.next) ? data.paging.next : null;
    pages++;
  }

  var creativeMap = getAdCreativeMap_();
  var ads = all.map(function (item) {
    var spend = parseFloat(item.spend) || 0;
    var impressions = parseInt(item.impressions) || 0;
    var linkClicks = parseInt(item.inline_link_clicks) || 0;
    var purchases = av_(item.actions, 'omni_purchase') + av_(item.actions, 'purchase');
    var id = String(item.ad_id || '');
    var creative = creativeMap[id];
    return {
      id: id,
      name: item.ad_name || '(unnamed)',
      spend: spend,
      impressions: impressions,
      ctrLink: impressions > 0 ? (linkClicks / impressions) * 100 : 0,
      purchases: purchases,
      // '' (not null) when nothing's resolved yet — client renders this as "no thumbnail
      // yet" rather than a broken image.
      thumbnailUrl: (creative && (creative.thumbnailUrl || creative.thumbnailFallbackUrl)) || ''
    };
  }).sort(function (a, b) { return b.spend - a.spend; }).slice(0, 15); // top 15 by spend — a pulse, not the full Ads tab

  return { ok: true, ads: ads };
}

// ── SHEET ACCESS ──────────────────────────────────────────────────────────────
function getSheet_() {
  var ss;
  if (SHEET_ID && SHEET_ID !== 'YOUR_SPREADSHEET_ID') {
    ss = SpreadsheetApp.openById(SHEET_ID);
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  if (!ss) {
    throw new Error('No spreadsheet resolved. Put a real SHEET_ID in Code.gs, or attach this script to your sheet (Extensions → Apps Script).');
  }
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    var names = ss.getSheets().map(function (sh) { return '"' + sh.getName() + '"'; });
    throw new Error('Tab "' + SHEET_NAME + '" not found. This spreadsheet has: ' + names.join(', ') +
      '. Set SHEET_NAME in Code.gs to match one of these EXACTLY (capitalization and spaces count).');
  }
  return sheet;
}

/**
 * DIAGNOSTIC — run this straight from the editor
 */
function diagnose() {
  var out = [];
  var active = null;
  try { active = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { active = null; }
  out.push('Bound/active spreadsheet: ' + (active ? active.getName() + '  [id ' + active.getId() + ']' : 'NONE (standalone script — SHEET_ID will be used)'));
  out.push('SHEET_ID in code: ' + (SHEET_ID === 'YOUR_SPREADSHEET_ID' ? '!! STILL THE PLACEHOLDER !!' : SHEET_ID));
  out.push('SHEET_NAME in code: "' + SHEET_NAME + '"');
  out.push('----');

  var ss;
  try {
    if (SHEET_ID && SHEET_ID !== 'YOUR_SPREADSHEET_ID') {
      ss = SpreadsheetApp.openById(SHEET_ID);
      out.push('openById(SHEET_ID): OK -> "' + ss.getName() + '"');
    } else {
      ss = active;
      out.push('No real SHEET_ID set, so using the active spreadsheet.');
    }
  } catch (e) {
    out.push('openById(SHEET_ID) FAILED: ' + e.message);
    out.push('>> Cause: the ID is wrong / you pasted the whole URL / this account cannot open that file.');
    Logger.log(out.join('\n')); return out.join('\n');
  }
  if (!ss) { out.push('>> Nothing to open. Set a real SHEET_ID or bind the script.'); Logger.log(out.join('\n')); return out.join('\n'); }

  var tabs = ss.getSheets().map(function (s) { return '"' + s.getName() + '"'; });
  out.push('Tabs in that file: ' + tabs.join(', '));

  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { out.push('>> No tab called "' + SHEET_NAME + '". Copy one of the tab names above into SHEET_NAME, exactly.'); Logger.log(out.join('\n')); return out.join('\n'); }

  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  out.push('"' + SHEET_NAME + '" has ' + lastRow + ' rows x ' + lastCol + ' cols of data.');
  if (lastRow < 2) {
    out.push('>> Fewer than 2 rows — nothing to read.');
  } else {
    var vals = sheet.getDataRange().getValues();
    var hr = findHeaderRow_(vals);
    out.push('Header row detected: row ' + (hr + 1) + '  (HEADER_ROW config = ' + HEADER_ROW + ').');
    out.push('Headers on that row: ' + vals[hr].join(' | '));
    var idx = resolveHeaders(vals[hr]);
    var missing = REQUIRED_FIELDS.filter(function (f) { return idx[f] === -1; });
    out.push('Required columns not found: ' + (missing.length ? missing.join(', ') : 'NONE — headers resolved correctly.'));
    out.push('Data rows beneath the header: ' + Math.max(0, vals.length - hr - 1));
    if (missing.length) out.push('>> Set HEADER_ROW at the top of Code.gs to the row your column names are on.');
  }
  var msg = out.join('\n');
  Logger.log(msg);
  return msg;
}

// Helper for ranking labels (used in aggregation)
function tierLabel(p) {
  if (p == null) return '–';
  if (p === 3) return 'Above average';
  if (p === 2) return 'Average';
  return 'Below average';
}

// ── BREAKDOWN DATA ──

/**
 * Get breakdown data for a specific breakdown type.
 * @param {string} breakdownType - 'age', 'gender', 'platform', or 'placement'
 * @param {object} filters - date filters (dateFrom, dateTo)
 */
function getBreakdownData(breakdownType, filters) {
  filters = filters || {};
  var sheetName = '';
  switch(breakdownType) {
    case 'age': sheetName = 'data_age'; break;
    case 'gender': sheetName = 'data_gender'; break;
    case 'platform': sheetName = 'data_platform'; break;
    case 'placement': sheetName = 'data_placement'; break;
    case 'region': sheetName = 'data_region'; break; // NEW (Aug 2026)
    default: return { error: 'Invalid breakdown type: ' + breakdownType };
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    return { error: 'Sheet "' + sheetName + '" not found. Run importBreakdownData() first.' };
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { rows: [] };

  var headers = data[0];
  var rows = data.slice(1);

  // Apply date filters if provided
  var from = filters.dateFrom ? new Date(filters.dateFrom) : null;
  var to = filters.dateTo ? new Date(filters.dateTo) : null;
  if (to) to.setHours(23, 59, 59, 999);

  var filtered = rows.filter(function(row) {
    if (!row[0]) return false;
    var d = new Date(row[0]);
    if (from && d < from) return false;
    if (to && d > to) return false;
    return true;
  });

  // Map to objects (cols 9-12 are the new Purchases/Purchase value/Video plays 3s/100% —
  // default to 0 for old rows that predate the schema upgrade, harmless)
  var result = filtered.map(function(row) {
    return {
      date: row[0],
      campaignName: row[1],
      breakdownValue: row[2],
      spend: row[3],
      impressions: row[4],
      clicks: row[5],
      reach: row[6],
      ctr: row[7],
      purchases: toNumber(row[8]),
      purchaseValue: toNumber(row[9]),
      video3s: toNumber(row[10]),
      video100: toNumber(row[11])
    };
  });

  // Aggregate by breakdown value (sum over dates)
  var summary = {};
  result.forEach(function(item) {
    var key = item.breakdownValue;
    if (!summary[key]) {
      summary[key] = { spend: 0, impressions: 0, clicks: 0, reach: 0, ctrSum: 0, count: 0,
        purchases: 0, purchaseValue: 0, video3s: 0, video100: 0 };
    }
    summary[key].spend += item.spend;
    summary[key].impressions += item.impressions;
    summary[key].clicks += item.clicks;
    summary[key].reach += item.reach;
    summary[key].ctrSum += item.ctr;
    summary[key].count++;
    summary[key].purchases += item.purchases;
    summary[key].purchaseValue += item.purchaseValue;
    summary[key].video3s += item.video3s;
    summary[key].video100 += item.video100;
  });

  var output = Object.keys(summary).map(function(key) {
    var s = summary[key];
    return {
      breakdownValue: key,
      spend: s.spend,
      impressions: s.impressions,
      clicks: s.clicks,
      reach: s.reach,
      ctr: s.count > 0 ? s.ctrSum / s.count : 0,
      purchases: s.purchases,
      purchaseValue: s.purchaseValue,
      roas: s.spend > 0 ? s.purchaseValue / s.spend : 0,
      videoPlays: s.video3s,
      hookRate: s.impressions > 0 ? (s.video3s / s.impressions) * 100 : 0,
      // FIXED (Sep 2026) — Hold Rate denominator should be 3-sec plays, not impressions;
      // see DataPipeline.gs's buildRowFromInsight_ for the full explanation.
      // NOT YET on the corrected numerator (ThruPlays, not Video plays at 100%) — the
      // breakdown sheets this reads from (age/gender/platform/placement/region, built by
      // BreakdownPipeline.gs) don't fetch video_thruplay_watched_actions, only the
      // video_p*_watched_actions milestones, so there's no ThruPlays figure available here
      // to use. The main Campaigns/Ad Sets tables (aggregateByCampaign_, buildSummary_ above)
      // ARE on the corrected formula. Fixing this one needs BreakdownPipeline.gs to fetch and
      // store ThruPlays too — flagged, not done.
      holdRate: s.video3s > 0 ? (s.video100 / s.video3s) * 100 : 0
    };
  });

  // Sort by spend descending
  output.sort(function(a, b) { return b.spend - a.spend; });

  return { rows: output, totalRows: result.length };
}

// ── REPORT BUILDER BACKEND ──
// SINGLE canonical implementation — CustomReport.gs's live-API version has been
// retired (delete that file from the project). Reads the breakdown sheets, which
// after the BreakdownPipeline.gs upgrade carry Purchases/Purchase value/Video plays
// alongside spend/impressions/clicks/reach/CTR.
// NOTE ON LIMITS:
//   • Sheets exist for age / gender / platform / placement. Region has NO sheet —
//     that stays out of scope for this round (would need its own daily+backfill
//     pipeline, same pattern as the other four, if you want it later).
//   • One breakdown dimension at a time (each sheet is single-dimension).
//   • "Results" here means Purchases specifically (breakdown sheets aren't sliced
//     per-campaign, so the optimization-goal-aware "Results" from the main table
//     can't be reconstructed at this granularity) — labeled "Results (purchases)"
//     on the frontend so it's not confused with the smarter campaign-level Results.

var BREAKDOWN_SHEETS_ = { age: 'data_age', gender: 'data_gender', platform: 'data_platform', placement: 'data_placement', region: 'data_region' };

function getCustomReport(cfg) {
  cfg = cfg || {};
  var dims = cfg.breakdowns || [];
  if (!dims.length) return { error: 'Pick a breakdown dimension.' };

  var dim = dims[0];
  var sheetName = BREAKDOWN_SHEETS_[dim];
  if (!sheetName) {
    return { error: 'The "' + dim + '" breakdown has no data. Available: Age, Gender, Platform, Placement, Region.' };
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return { error: 'Sheet "' + sheetName + '" not found. Run the breakdown backfill first (backfillRegionHistory() for Region).' };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { rows: [], campaigns: [], note: '' };
  return aggregateBreakdownData_(data, cfg, dim, dims.length > 1);
}

// Shared by getCustomReport() (live sheet) and ArchiveTools.gs's getCustomReportFromCsv()
// (uploaded CSV) — same 13-column row shape either way, so the exact same aggregation
// runs regardless of source. Pulled out (Aug 2026) specifically to support archived-year
// Report Builder without forking this logic into two copies that could drift apart.
function aggregateBreakdownData_(data, cfg, dim, multiDimsRequested) {
  var from = cfg.dateFrom ? new Date(cfg.dateFrom) : null;
  var to = cfg.dateTo ? new Date(cfg.dateTo) : null;
  if (to) to.setHours(23, 59, 59, 999);
  var campaign = (cfg.campaignName || '').trim();
  // Region-only free-text filter (e.g. "Metro Manila") — matches the frontend's
  // "Region filter" input, which was already wired up ahead of this backend support.
  var regionFilter = (dim === 'region' && cfg.region) ? String(cfg.region).trim().toLowerCase() : '';

  // columns: [date, campaign, breakdownValue, spend, impressions, clicks, reach, ctr, purchases, purchaseValue, video3s, video100, linkClicks]
  var active = {}, agg = {};
  data.slice(1).forEach(function (r) {
    if (!r[0]) return;
    var d = new Date(r[0]);
    if (from && d < from) return;
    if (to && d > to) return;
    var cname = String(r[1] || '');
    if (cname) active[cname] = true;                 // campaigns active in this window
    if (campaign && cname !== campaign) return;      // optional campaign filter
    var key = String(r[2] || '(none)');
    if (regionFilter && key.toLowerCase().indexOf(regionFilter) === -1) return;
    if (!agg[key]) agg[key] = { spend: 0, impressions: 0, clicks: 0, reach: 0, purchases: 0, purchaseValue: 0, video3s: 0, video100: 0, linkClicks: 0 };
    agg[key].spend += toNumber(r[3]);
    agg[key].impressions += toNumber(r[4]);
    agg[key].clicks += toNumber(r[5]);
    agg[key].reach += toNumber(r[6]);
    agg[key].purchases += toNumber(r[8]);
    agg[key].purchaseValue += toNumber(r[9]);
    agg[key].video3s += toNumber(r[10]);
    agg[key].video100 += toNumber(r[11]);
    agg[key].linkClicks += toNumber(r[12]); // NEW — real link clicks (col 13); blank/0 on rows from before this column existed
  });

  var out = Object.keys(agg).map(function (key) {
    var a = agg[key];
    return {
      label: key,
      spend: a.spend, impressions: a.impressions, clicks: a.clicks, reach: a.reach,
      // FIXED (Aug 2026): this used to recompute all-clicks CTR under the "CTR (link)"
      // label — same formula as ctrAll, just duplicated. Now genuinely link-click CTR,
      // using the new Link clicks column (inline_link_clicks from Meta).
      ctr: a.impressions > 0 ? (a.linkClicks / a.impressions) * 100 : 0,
      ctrAll: a.impressions > 0 ? (a.clicks / a.impressions) * 100 : 0,
      cpm: a.impressions > 0 ? (a.spend / a.impressions) * 1000 : 0,
      cpc: a.clicks > 0 ? a.spend / a.clicks : 0,
      frequency: a.reach > 0 ? a.impressions / a.reach : 0,
      purchases: a.purchases,
      results: a.purchases, // see note above — Results at this granularity = purchases
      purchaseValue: a.purchaseValue,
      roas: a.spend > 0 ? a.purchaseValue / a.spend : 0,
      videoPlays: a.video3s,
      hookRate: a.impressions > 0 ? (a.video3s / a.impressions) * 100 : 0,
      // FIXED (Sep 2026) — Hold Rate denominator should be 3-sec plays, not impressions;
      // see DataPipeline.gs's buildRowFromInsight_ for the full explanation.
      // NOT YET on the corrected numerator (ThruPlays, not Video plays at 100%) — same gap
      // as getBreakdownData above: these breakdown sheets don't store a ThruPlays figure.
      holdRate: a.video3s > 0 ? (a.video100 / a.video3s) * 100 : 0,
      linkClicks: a.linkClicks
    };
  });
  out.sort(function (a, b) { return b.spend - a.spend; });

  var note = '';
  if (multiDimsRequested) note += 'Only one breakdown at a time is supported with current data; showing "' + dim + '". ';
  return { rows: out, campaigns: Object.keys(active).sort(), note: note };
}

// Distinct campaigns active in a date range (for the report builder's campaign dropdown).
function getReportCampaigns(cfg) {
  cfg = cfg || {};
  var dim = (cfg.breakdowns && cfg.breakdowns[0]) || 'age';
  var sheetName = BREAKDOWN_SHEETS_[dim] || 'data_age';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return { campaigns: [] };
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { campaigns: [] };
  var from = cfg.dateFrom ? new Date(cfg.dateFrom) : null;
  var to = cfg.dateTo ? new Date(cfg.dateTo) : null;
  if (to) to.setHours(23, 59, 59, 999);
  var set = {};
  data.slice(1).forEach(function (r) {
    if (!r[0]) return;
    var d = new Date(r[0]);
    if (from && d < from) return;
    if (to && d > to) return;
    if (r[1]) set[String(r[1])] = true;
  });
  return { campaigns: Object.keys(set).sort() };
}

// ── AD SET LEVEL ──
// Reads "data - adset - by day" and returns the same shape as getDashboardData,
// but grouped by AD SET (each ad set is treated as a "campaign" internally so we
// reuse aggregateByCampaign_ + buildSummary_ unchanged). campaignName is REQUIRED
// (single-campaign view). Also returns campaignList for the picker.
function getAdSetData(campaignName, filters) {
  filters = filters || {};
  try {
    var sheet = getAdSetSpreadsheet_().getSheetByName('data - adset - by day');
    if (!sheet) {
      return { rows: [], campaigns: [], campaignList: [], summary: emptySummary_(),
               error: 'Sheet "data - adset - by day" not found. Add it, then run backfillAdSetHistory().' };
    }
    var range = sheet.getDataRange().getValues();
    if (range.length < 2) return { rows: [], campaigns: [], campaignList: [], summary: emptySummary_(), error: null };

    var headerRowIdx = findHeaderRow_(range);
    var headers = range[headerRowIdx];
    var idx = resolveHeaders(headers);

    var rows = range.slice(headerRowIdx + 1).map(function (r) {
      var g = function (field) { return idx[field] === -1 ? '' : r[idx[field]]; };
      var spent = toNumber(g('spent'));
      var cpm = toNumber(g('cpm'));
      var ctrLink = toNumber(g('ctrLink'));
      var ctrAllRaw = toNumber(g('ctrAll'));
      var d = derive(spent, cpm, ctrLink);
      // Same fix as getDashboardData() above — reconstruct real all-clicks count from the
      // stored ctrAll percentage, so aggregateByCampaign_ can compute a genuine CTR (all)
      // for the Ad Set view instead of reusing link clicks.
      var allClicks = d.impressions > 0 ? d.impressions * (ctrAllRaw / 100) : 0;
      var dateStr = toDateString(g('reportingEnds'));
      var adset = String(g('adSetName') || '').trim() || '(unnamed ad set)';
      return {
        campaignName: adset,                                   // relabel: group by ad set
        _realCampaign: String(g('campaignName') || '').trim(),
        objective: String(g('objective') || '').trim() || 'Unknown',
        bidStrategy: String(g('bidStrategy') || '').trim(),
        reportingEnds: dateStr, reportingEndsTs: tsOf(dateStr),
        spent: spent, results: toNumber(g('results')), reach: toNumber(g('reach')),
        purchases: toNumber(g('purchases')), purchaseValue: toNumber(g('purchaseValue')),
        videoPlays: toNumber(g('videoPlays')), postEngagements: toNumber(g('postEngagements')),
        ctrAll: ctrAllRaw, ctrLink: ctrLink, cpm: cpm,
        impressions: d.impressions, clicks: d.clicks, allClicks: allClicks, cpc: d.cpc,
        roasSheet: toNumber(g('roas')), dailyBudget: toNumber(g('dailyBudget')), frequency: toNumber(g('frequency')),
        blockedContacts: toNumber(g('blockedContacts')), costPerMessage: toNumber(g('costPerMessage')),
        video3s: toNumber(g('video3s')), video25: toNumber(g('video25')), video50: toNumber(g('video50')),
        video75: toNumber(g('video75')), video95: toNumber(g('video95')), video100: toNumber(g('video100')),
        qRank: rankToPoints(g('qualityRank')), eRank: rankToPoints(g('engagementRank')), cRank: rankToPoints(g('conversionRank')),
        // ── NEW ──
        optimizationGoal: String(g('optimizationGoal') || '').trim(),
        landingPageViews: toNumber(g('landingPageViews')),
        initiateCheckouts: toNumber(g('initiateCheckouts')),
        thruplays: toNumber(g('thruplays'))
      };
    }).filter(function (row) {
      return row.campaignName !== '(unnamed ad set)' || row.spent || row.results || row.reach;
    });

    // distinct campaigns present (for the picker) — before the campaign filter
    var campSet = {};
    rows.forEach(function (r) { if (r._realCampaign) campSet[r._realCampaign] = true; });
    var campaignList = Object.keys(campSet).sort();

    // date filters (optional)
    if (filters.dateFrom || filters.dateTo) {
      var fromTs = filters.dateFrom ? new Date(filters.dateFrom).getTime() : -Infinity;
      var toTs = filters.dateTo ? new Date(filters.dateTo).getTime() + 86399999 : Infinity;
      rows = rows.filter(function (r) { return !isNaN(r.reportingEndsTs) && r.reportingEndsTs >= fromTs && r.reportingEndsTs <= toTs; });
    }

    // REQUIRED single-campaign filter
    if (campaignName) rows = rows.filter(function (r) { return r._realCampaign === campaignName; });
    else return { rows: [], campaigns: [], campaignList: campaignList, summary: emptySummary_(), error: null };

    var statusMap = getAdSetStatusMap_();
    var adsets = aggregateByCampaign_(rows, statusMap);
    var summary = buildSummary_(rows, adsets);
    return {
      rows: rows, campaigns: adsets, campaignList: campaignList, summary: summary,
      generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
      error: null
    };
  } catch (e) {
    return { rows: [], campaigns: [], campaignList: [], summary: emptySummary_(), error: e.message || String(e) };
  }
}

// List of campaigns present in the ad-set sheet (for the picker).
// UPDATED (Aug 2026): accepts optional filters — { dateFrom, dateTo, activeOnly }.
// activeOnly restricts the list to campaigns that had at least one row with spend > 0
// within [dateFrom, dateTo] (or ever, if no range is given). Since fetchAdSetRows_ pulls
// with include_archived:true, this dropdown otherwise lists every campaign Meta ever ran
// on this account, including long-paused/archived ones — exactly what made it hard to
// search. Returns totalCampaigns too, so the frontend can show "(N of M)".
function getAdSetCampaigns(filters) {
  filters = filters || {};
  var sheet = getAdSetSpreadsheet_().getSheetByName('data - adset - by day');
  if (!sheet) return { campaigns: [], totalCampaigns: 0 };
  var range = sheet.getDataRange().getValues();
  if (range.length < 2) return { campaigns: [], totalCampaigns: 0 };
  var headerRowIdx = findHeaderRow_(range);
  var idx = resolveHeaders(range[headerRowIdx]);
  if (idx.campaignName === -1) return { campaigns: [], totalCampaigns: 0 };

  var fromTs = filters.dateFrom ? new Date(filters.dateFrom).getTime() : -Infinity;
  var toTs = filters.dateTo ? new Date(filters.dateTo).getTime() + 86399999 : Infinity;
  var activeOnly = !!filters.activeOnly;

  var all = {}, active = {};
  range.slice(headerRowIdx + 1).forEach(function (r) {
    var c = String(r[idx.campaignName] || '').trim();
    if (!c) return;
    all[c] = true;
    if (!activeOnly) return;
    var dateStr = idx.reportingEnds === -1 ? '' : toDateString(r[idx.reportingEnds]);
    var ts = tsOf(dateStr);
    if (isNaN(ts) || ts < fromTs || ts > toTs) return;   // outside the chosen window — doesn't count as "active" for it
    var spent = idx.spent === -1 ? 0 : toNumber(r[idx.spent]);
    if (spent > 0) active[c] = true;
  });

  var result = activeOnly ? Object.keys(active) : Object.keys(all);
  return { campaigns: result.sort(), totalCampaigns: Object.keys(all).length };
}

// ── NEW: Learning Phase Status + Audience Type lookup (current snapshot only) ──
// ADSET_STATUS_SHEET_ is declared in AdSetPipeline.gs and shared as a global.
function getAdSetStatusMap_() {
  var map = {};
  var sheet = getAdSetSpreadsheet_().getSheetByName(ADSET_STATUS_SHEET_);
  if (!sheet) return map; // importAdSetStatus() hasn't been run yet — fields just stay blank
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return map;
  data.slice(1).forEach(function (r) {
    var name = String(r[0] || '').trim();
    if (!name) return;
    map[name] = { learningPhaseStatus: r[2] || '', audienceType: r[3] || '' };
  });
  return map;
}

// ── NEW: Ad-set-level Placement / Age×Gender breakdown for the toggleable widgets ──
// breakdownType: 'placement' | 'agegender'
function getAdSetBreakdownData(campaignName, breakdownType, filters) {
  filters = filters || {};
  var sheetName = breakdownType === 'agegender' ? ADSET_AGEGENDER_SHEET_ : ADSET_PLACEMENT_SHEET_;
  var sheet = getAdSetSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) return { rows: [], error: 'Sheet "' + sheetName + '" not found. Run backfillAdSetBreakdownHistory() first.' };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { rows: [] };

  var from = filters.dateFrom ? new Date(filters.dateFrom) : null;
  var to = filters.dateTo ? new Date(filters.dateTo) : null;
  if (to) to.setHours(23, 59, 59, 999);

  var agg = {};
  data.slice(1).forEach(function (r) {
    if (!r[0]) return;
    var d = new Date(r[0]);
    if (from && d < from) return;
    if (to && d > to) return;
    var cname = String(r[1] || '');
    if (campaignName && cname !== campaignName) return;

    if (breakdownType === 'agegender') {
      // [date, campaign, adset, age, gender, spend, impressions, clicks, reach, purchases, purchaseValue]
      var key = String(r[3] || 'unknown') + ' / ' + String(r[4] || 'unknown');
      if (!agg[key]) agg[key] = { label: key, spend: 0, impressions: 0, clicks: 0, reach: 0, purchases: 0, purchaseValue: 0 };
      agg[key].spend += toNumber(r[5]);
      agg[key].impressions += toNumber(r[6]);
      agg[key].clicks += toNumber(r[7]);
      agg[key].reach += toNumber(r[8]);
      agg[key].purchases += toNumber(r[9]);
      agg[key].purchaseValue += toNumber(r[10]);
    } else {
      // [date, campaign, adset, placement, spend, impressions, clicks, reach, ctr, purchases, purchaseValue]
      var pkey = String(r[3] || 'unknown');
      if (!agg[pkey]) agg[pkey] = { label: pkey, spend: 0, impressions: 0, clicks: 0, reach: 0, ctrSum: 0, count: 0, purchases: 0, purchaseValue: 0 };
      agg[pkey].spend += toNumber(r[4]);
      agg[pkey].impressions += toNumber(r[5]);
      agg[pkey].clicks += toNumber(r[6]);
      agg[pkey].reach += toNumber(r[7]);
      agg[pkey].ctrSum += toNumber(r[8]);
      agg[pkey].count++;
      agg[pkey].purchases += toNumber(r[9]);
      agg[pkey].purchaseValue += toNumber(r[10]);
    }
  });

  var out = Object.keys(agg).map(function (key) {
    var a = agg[key];
    var row = {
      label: a.label, spend: a.spend, impressions: a.impressions, clicks: a.clicks, reach: a.reach,
      purchases: a.purchases, purchaseValue: a.purchaseValue,
      roas: a.spend > 0 ? a.purchaseValue / a.spend : 0
    };
    if (breakdownType !== 'agegender') row.ctr = a.count > 0 ? a.ctrSum / a.count : 0;
    return row;
  });
  out.sort(function (a, b) { return b.spend - a.spend; });
  return { rows: out };
}

// ═══════════════════════════════════════════════════════════════════════════
// ── AD LEVEL (Aug 2026) ──
// Reads "data - ad - by day" (AdPipeline.gs) and returns one scored record per AD,
// answering the three questions you asked for: which creatives are winning, why one ad
// differs from another (per-metric band breakdown), and what to test next (the Scale
// Decision recommendation). Reuses aggregateByCampaign_ the same trick getAdSetData()
// uses — group by ad instead of by campaign — so nothing about that shared function
// needed to change.
//
// DESIGN CALL: unlike the Ad Set tab (pick-one-campaign-first, load-on-demand), this
// returns campaigns/ad sets/ads in ONE call. An ad-level view is inherently about
// browsing/comparing many creatives at once (that's the whole point of a "which one is
// winning" checklist) rather than drilling into one entity, so a single round trip that
// feeds both the pickers and the grid is simpler and, if anything, faster than the
// two-call pattern would have been here.
//
// NOT built in this pass (said here plainly rather than left silently missing):
//   - Archive/CSV mode. Thumbnails need a LIVE call to Meta's /ads edge — a CSV export
//     can carry the metrics but not a working preview image, so wiring this into
//     ArchiveTools.gs would only get you half the feature. Live mode only, for now.
//   - A "select 2 ads and compare" view. The per-ad Metric-by-metric band breakdown
//     below already lets you SEE why two ads differ side by side on the page, but there's
//     no dedicated A/B comparison UI yet — a reasonable v2 if you want it.
// ═══════════════════════════════════════════════════════════════════════════

function getAdData(filters) {
  filters = filters || {};
  try {
    var sheet = getAdSpreadsheet_().getSheetByName('data - ad - by day');
    if (!sheet) {
      return { ads: [], campaignList: [], adsetList: [], summary: emptySummary_(),
               error: 'Sheet "data - ad - by day" not found. Add it, then run backfillAdHistory() (see AdPipeline.gs).' };
    }
    var range = sheet.getDataRange().getValues();
    if (range.length < 2) return { ads: [], campaignList: [], adsetList: [], summary: emptySummary_(), error: null };

    var headerRowIdx = findHeaderRow_(range);
    var headers = range[headerRowIdx];
    var idx = resolveHeaders(headers);
    var idxAdId = -1, idxAdName = -1;
    headers.forEach(function (h, i) {
      var n = norm(h);
      if (n === 'adid') idxAdId = i;
      if (n === 'adname') idxAdName = i;
    });

    var rows = range.slice(headerRowIdx + 1).map(function (r) {
      var g = function (field) { return idx[field] === -1 ? '' : r[idx[field]]; };
      var spent = toNumber(g('spent'));
      var cpm = toNumber(g('cpm'));
      var ctrLink = toNumber(g('ctrLink'));
      var ctrAllRaw = toNumber(g('ctrAll'));
      var d = derive(spent, cpm, ctrLink);
      var allClicks = d.impressions > 0 ? d.impressions * (ctrAllRaw / 100) : 0;
      var dateStr = toDateString(g('reportingEnds'));
      var adId = idxAdId === -1 ? '' : String(r[idxAdId] || '').trim();
      var adName = idxAdName === -1 ? '' : String(r[idxAdName] || '').trim();
      var groupKey = adId || adName || '(unnamed ad)';
      return {
        campaignName: groupKey,                                // group-by key (reuses aggregateByCampaign_ unchanged)
        _adId: adId, _adName: adName || '(unnamed ad)',
        _realCampaign: String(g('campaignName') || '').trim(),
        _realAdSet: String(g('adSetName') || '').trim(),
        objective: String(g('objective') || '').trim() || 'Unknown',
        bidStrategy: String(g('bidStrategy') || '').trim(),
        reportingEnds: dateStr, reportingEndsTs: tsOf(dateStr),
        spent: spent, results: toNumber(g('results')), reach: toNumber(g('reach')),
        purchases: toNumber(g('purchases')), purchaseValue: toNumber(g('purchaseValue')),
        videoPlays: toNumber(g('videoPlays')), postEngagements: toNumber(g('postEngagements')),
        ctrAll: ctrAllRaw, ctrLink: ctrLink, cpm: cpm,
        impressions: d.impressions, clicks: d.clicks, allClicks: allClicks, cpc: d.cpc,
        roasSheet: toNumber(g('roas')), dailyBudget: toNumber(g('dailyBudget')), frequency: toNumber(g('frequency')),
        blockedContacts: toNumber(g('blockedContacts')), costPerMessage: toNumber(g('costPerMessage')),
        video3s: toNumber(g('video3s')), video25: toNumber(g('video25')), video50: toNumber(g('video50')),
        video75: toNumber(g('video75')), video95: toNumber(g('video95')), video100: toNumber(g('video100')),
        qRank: rankToPoints(g('qualityRank')), eRank: rankToPoints(g('engagementRank')), cRank: rankToPoints(g('conversionRank')),
        optimizationGoal: String(g('optimizationGoal') || '').trim(),
        landingPageViews: toNumber(g('landingPageViews')),
        initiateCheckouts: toNumber(g('initiateCheckouts')),
        thruplays: toNumber(g('thruplays'))
      };
    }).filter(function (row) {
      return row._adName !== '(unnamed ad)' || row.spent || row.results || row.reach;
    });

    // distinct campaigns present (for the picker) — before any filtering
    var campSetAll = {};
    rows.forEach(function (r) { if (r._realCampaign) campSetAll[r._realCampaign] = true; });
    var campaignList = Object.keys(campSetAll).sort();

    // date filters (optional)
    if (filters.dateFrom || filters.dateTo) {
      var fromTs = filters.dateFrom ? new Date(filters.dateFrom).getTime() : -Infinity;
      var toTs = filters.dateTo ? new Date(filters.dateTo).getTime() + 86399999 : Infinity;
      rows = rows.filter(function (r) { return !isNaN(r.reportingEndsTs) && r.reportingEndsTs >= fromTs && r.reportingEndsTs <= toTs; });
    }
    // campaign filter (optional) — narrows BEFORE computing the ad-set list, so that
    // dropdown scopes itself to whatever campaign is currently chosen
    if (filters.campaignName) rows = rows.filter(function (r) { return r._realCampaign === filters.campaignName; });

    var scopedAdsetSet = {};
    rows.forEach(function (r) { if (r._realAdSet) scopedAdsetSet[r._realAdSet] = true; });
    var adsetList = Object.keys(scopedAdsetSet).sort();

    if (filters.adSetName) rows = rows.filter(function (r) { return r._realAdSet === filters.adSetName; });
    if (filters.search) {
      var needle = String(filters.search).toLowerCase();
      rows = rows.filter(function (r) { return r._adName.toLowerCase().indexOf(needle) !== -1; });
    }

    var ads = aggregateByCampaign_(rows, null); // no learning-phase/audience-type join at ad level

    var nameByKey = {};
    rows.forEach(function (r) {
      if (!nameByKey[r.campaignName]) {
        nameByKey[r.campaignName] = { adId: r._adId, adName: r._adName, realCampaign: r._realCampaign, realAdSet: r._realAdSet };
      }
    });

    var creativeMap = getAdCreativeMap_();
    ads = ads.map(function (a) {
      var meta = nameByKey[a.campaignName] || {};
      var creative = meta.adId ? creativeMap[meta.adId] : null;
      a.adId = meta.adId || '';
      a.adName = meta.adName || a.campaignName;
      a.realCampaignName = meta.realCampaign || '';
      a.adSetName = meta.realAdSet || '';
      a.creativeType = creative ? creative.creativeType : 'Unknown';
      a.thumbnailUrl = creative ? creative.thumbnailUrl : '';
      a.thumbnailFallbackUrl = creative ? creative.thumbnailFallbackUrl : '';
      a.status = creative ? creative.status : '';
      a.adsManagerLink = (creative && creative.adsManagerLink) || (meta.adId ? adsManagerLinkFor_(meta.adId) : '');
      a.ctrGap = (a.ctrAll || 0) - (a.ctrLink || 0); // Diagnostic Metrics: "CTR All vs CTR Link gap"
      a.scorecard = computeAdScorecard_(a);
      return a;
    });

    // Sep 2026, "add a filter to be able to look at images only or videos": a.creativeType
    // comes from ad_creatives (classifyCreativeType_ in AdPipeline.gs) — 'Video', 'Image',
    // 'Carousel', or 'Unknown' if that ad's creative row hasn't been synced yet.
    if (filters.creativeType) ads = ads.filter(function (a) { return a.creativeType === filters.creativeType; });
    if (filters.activeOnly) ads = ads.filter(function (a) { return a.spent > 0; });
    if (filters.winnersOnly) ads = ads.filter(function (a) { return a.scorecard.verdict === 'Winner'; });

    var sortKeyIn = filters.sort || 'score';
    ads.sort(function (x, y) {
      if (sortKeyIn === 'roas') return y.roas - x.roas;
      if (sortKeyIn === 'ctrLink') return y.ctrLink - x.ctrLink;
      if (sortKeyIn === 'spend') return y.spent - x.spent;
      var xp = x.scorecard.percent === null ? -1 : x.scorecard.percent;
      var yp = y.scorecard.percent === null ? -1 : y.scorecard.percent;
      return yp - xp; // default: best scorecard first
    });

    var summary = buildSummary_(rows, ads);
    return {
      ads: ads, campaignList: campaignList, adsetList: adsetList, summary: summary,
      generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
      error: null
    };
  } catch (e) {
    return { ads: [], campaignList: [], adsetList: [], summary: emptySummary_(), error: e.message || String(e) };
  }
}

// AD_CREATIVE_SHEET_ is declared in AdPipeline.gs and shared as a global.
function getAdCreativeMap_() {
  var map = {};
  var sheet = getAdSpreadsheet_().getSheetByName(AD_CREATIVE_SHEET_);
  if (!sheet) return map; // importAdCreatives() hasn't been run yet — creative fields just stay blank
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return map;
  data.slice(1).forEach(function (r) {
    var id = String(r[0] || '').trim();
    if (!id) return;
    // r[9] ('Fallback Thumbnail URL') added Sep 2026 — reads as '' on a sheet that hasn't
    // been re-synced by importAdCreatives() yet since the column was added, which is exactly
    // the fallback-not-available case the dashboard already has to handle gracefully.
    map[id] = { status: r[4] || '', creativeType: r[5] || 'Unknown', thumbnailUrl: r[6] || '', adsManagerLink: r[7] || '', thumbnailFallbackUrl: r[9] || '' };
  });
  return map;
}

// ── BENCHMARK SCORING ──
// Poor/Average/Good/Excellent bands, driven off a Poor/Average/Good/Excellent lookup table
// for three objective groups — "Messaging", "Sales (Purchase Objective)", and a "Generic"
// fallback for objectives that are neither (Awareness/Traffic/Leads/etc). In the real,
// deployed version of this dashboard these numbers came from the advertiser's own vertical-
// specific benchmark data; the ILLUSTRATIVE placeholder values below are for this portfolio
// copy only — swap them for your own market/vertical's real benchmarks (and currency).
var AD_BENCHMARK_BANDS_ = {
  messaging: {
    cpm:        { dir: 'lower',  bounds: [300, 200, 120] }, // [EXAMPLE_THRESHOLDS]
    ctrLink:    { dir: 'higher', bounds: [1.5, 2.5, 4.0] },
    ctrAll:     { dir: 'higher', bounds: [3.0, 5.0, 8.0] },
    frequency:  { dir: 'lower',  bounds: [4.0, 2.5, 1.5] },
    hookRate:   { dir: 'higher', bounds: [20, 30, 40] },
    holdRate:   { dir: 'higher', bounds: [20, 30, 40] },
    costPerMsg: { dir: 'lower',  bounds: [150, 90, 50] }
  },
  sales: {
    cpm:       { dir: 'lower',  bounds: [350, 220, 150] }, // [EXAMPLE_THRESHOLDS]
    ctrLink:   { dir: 'higher', bounds: [1.2, 2.0, 3.5] },
    ctrAll:    { dir: 'higher', bounds: [2.5, 4.0, 6.5] },
    frequency: { dir: 'lower',  bounds: [4.5, 3.0, 2.0] },
    hookRate:  { dir: 'higher', bounds: [20, 30, 40] },
    holdRate:  { dir: 'higher', bounds: [20, 30, 40] },
    roas:      { dir: 'higher', bounds: [1.5, 3.0, 4.5] }
  },
  // Used when the objective/optimization goal is neither Sales nor Messaging. The PDF's
  // universal table only gives 3 tiers (Poor/Good/Excellent) for these, so "Good" doubles
  // as the Average/Good split here — a reasonable reading, not something stated verbatim.
  generic: {
    hookRate: { dir: 'higher', bounds: [25, 35, 45] },
    holdRate: { dir: 'higher', bounds: [20, 30, 40] },
    ctrLink:  { dir: 'higher', bounds: [2, 3, 4] }
  }
};
var AD_BAND_LABELS_ = ['Poor', 'Average', 'Good', 'Excellent'];

function classifyBand_(value, spec) {
  if (value === null || value === undefined || value === '' || isNaN(value)) return null;
  var b = spec.bounds, pts;
  if (spec.dir === 'higher') {
    pts = value < b[0] ? 0 : value < b[1] ? 1 : value < b[2] ? 2 : 3;
  } else {
    pts = value > b[0] ? 0 : value > b[1] ? 1 : value > b[2] ? 2 : 3;
  }
  return { points: pts, label: AD_BAND_LABELS_[pts] };
}

// CONVERSATIONS optimization goal = a Messenger/WhatsApp/IG-DM ad, regardless of the
// campaign objective enum wrapping it — matches the PDF's "Messaging Campaign" bucket
// precisely, since Cost per Messaging Conversation only exists for that goal.
function classifyObjectiveGroup_(objective, optimizationGoal) {
  var og = String(optimizationGoal || '').toUpperCase();
  var obj = String(objective || '').toUpperCase();
  if (og === 'CONVERSATIONS') return 'messaging';
  if (obj.indexOf('SALES') !== -1 || og === 'OFFSITE_CONVERSIONS' || og === 'ONSITE_CONVERSIONS' || og === 'VALUE') return 'sales';
  return 'generic';
}

function computeAdScorecard_(a) {
  var group = classifyObjectiveGroup_(a.objective, a.optimizationGoal);
  var bands = AD_BENCHMARK_BANDS_[group];
  // FIXED (Sep 2026, "static image ads shouldn't have a hook rate and hold rate"): Hook Rate
  // and Hold Rate only mean anything for a creative with actual video-play data. A static
  // image ad's video3s/video100 counts are always 0 (there's no video to play), which used
  // to run straight through classifyBand_ and land as a hard "Poor" (0 points) on BOTH
  // metrics every single time — silently dragging down every image ad's overall score and
  // verdict for something that was never a fair thing to judge it on. Only score these two
  // when the creative is actually a video; classifyObjectiveGroup_'s a.creativeType is set
  // upstream in getAdData() before this runs.
  var isVideo = a.creativeType === 'Video';
  var metricValues = {
    cpm: a.cpm, ctrLink: a.ctrLink, ctrAll: a.ctrAll, frequency: a.frequency,
    hookRate: a.hookRate, holdRate: a.holdRate, roas: a.roas, costPerMsg: a.costPerMsg
  };
  var breakdown = [];
  var totalPts = 0, counted = 0;
  Object.keys(bands).forEach(function (key) {
    if (!isVideo && (key === 'hookRate' || key === 'holdRate')) return;
    var val = metricValues[key];
    var res = classifyBand_(val, bands[key]);
    if (!res) return; // metric not applicable/no data yet (e.g. costPerMsg with no messaging activity)
    breakdown.push({ metric: key, value: val, label: res.label, points: res.points });
    totalPts += res.points; counted++;
  });

  // Guards against mislabeling a brand-new/barely-spent ad as "Losing" just because its
  // early numbers are noisy — the framework's own "give it time before judging" instinct.
  var hasEnoughData = a.impressions >= 1000;
  var percent = counted > 0 ? Math.round((totalPts / (counted * 3)) * 100) : null;
  var verdict;
  if (!hasEnoughData || percent === null) verdict = 'Insufficient data';
  else if (percent >= 75) verdict = 'Winner';
  else if (percent >= 45) verdict = 'Needs optimization';
  else verdict = 'Losing';

  var decision = computeScaleDecision_(group, breakdown, isVideo);
  return { group: group, percent: percent, verdict: verdict, breakdown: breakdown, decision: decision.action, decisionWhy: decision.why };
}

function pointsFor_(breakdown, key) {
  for (var i = 0; i < breakdown.length; i++) if (breakdown[i].metric === key) return breakdown[i].points;
  return null;
}

// "Scale Decision Matrix" — this exact CPM/CTR/Hook/Hold/ROAS -> action table is VERBATIM
// from the benchmark PDF, but the PDF only gives it for the Sales (Purchase) objective.
// The 'messaging' and 'generic' branches below are THIS DASHBOARD'S OWN extrapolation of
// the same low/high logic — not stated in the PDF — flagged here (and in SETUP_GUIDE.md)
// so it's clear which part is your stated benchmark and which part is a judgment call
// layered on top of it.
function computeScaleDecision_(group, breakdown, isVideo) {
  var cpmPts = pointsFor_(breakdown, 'cpm');
  var ctrPts = pointsFor_(breakdown, 'ctrAll') !== null ? pointsFor_(breakdown, 'ctrAll') : pointsFor_(breakdown, 'ctrLink');
  var hookPts = pointsFor_(breakdown, 'hookRate');
  var holdPts = pointsFor_(breakdown, 'holdRate');
  var lowCpm = cpmPts !== null && cpmPts >= 2;

  // FIXED (Sep 2026): for a static image/carousel ad, hookPts/holdPts are always null now
  // (computeAdScorecard_ no longer scores them at all — see its comment). The old code read
  // "null" the same as "weak," which meant creativeStrong could never be true and
  // creativeWeak leaned true for every non-video ad regardless of how good its CTR actually
  // was. Now the "creative signal" is built only from whichever of CTR/Hook/Hold actually
  // apply to this ad — CTR alone for an image/carousel, all three for a video — so a
  // strong-CTR image ad can still be recognized as strong, and the "why" text below only
  // ever names metrics that were actually checked.
  var creativeMetricLabels = ['CTR'];
  var creativeSignals = [ctrPts !== null && ctrPts >= 2];
  if (hookPts !== null) { creativeMetricLabels.push('Hook Rate'); creativeSignals.push(hookPts >= 2); }
  if (holdPts !== null) { creativeMetricLabels.push('Hold Rate'); creativeSignals.push(holdPts >= 2); }
  var creativeStrong = creativeSignals.every(function (s) { return s; });
  var creativeWeak = creativeSignals.every(function (s) { return !s; });
  var creativeLabel = creativeMetricLabels.length > 1
    ? creativeMetricLabels.slice(0, -1).join(', ') + ', and ' + creativeMetricLabels[creativeMetricLabels.length - 1]
    : creativeMetricLabels[0];
  var creativeWeakWhy = isVideo
    ? creativeLabel + ' are all weak — the creative itself likely isn’t landing.'
    : creativeLabel + ' is weak — the creative itself likely isn’t landing (Hook/Hold Rate don’t apply to a static image/carousel ad).';
  var creativeStrongParen = isVideo ? '(CTR/Hook/Hold)' : '(CTR — Hook/Hold Rate don’t apply to a static image/carousel ad)';

  if (group === 'sales') {
    var roasPts = pointsFor_(breakdown, 'roas');
    var highRoas = roasPts !== null && roasPts >= 2;
    if (lowCpm && creativeStrong && highRoas) return { action: 'Scale aggressively', why: 'Efficient CPM, strong creative signals ' + creativeStrongParen + ', and healthy ROAS.' };
    if (!lowCpm && creativeStrong && !highRoas) return { action: 'Improve landing page or offer', why: 'The creative is clearly working (' + creativeLabel + ' are good), but ROAS is lagging — the leak is likely after the click.' };
    if (!lowCpm && creativeWeak) return { action: 'New audience + new creative', why: 'High CPM and weak engagement together usually means audience fatigue, not just a bad ad.' };
    if (lowCpm && creativeStrong && roasPts === 1) return { action: 'Optimize checkout or pricing', why: 'CPM and creative are both healthy, but ROAS is only average — check checkout friction or price/offer fit.' };
    if (creativeWeak) return { action: 'Replace creative', why: creativeWeakWhy };
    return { action: 'Keep testing', why: 'Mixed signals — no single lever stands out yet. Compare against your other ads in this set before deciding.' };
  }

  if (group === 'messaging') {
    var cpaPts = pointsFor_(breakdown, 'costPerMsg');
    var cheapConvo = cpaPts !== null && cpaPts >= 2;
    if (lowCpm && creativeStrong && cheapConvo) return { action: 'Scale aggressively', why: 'Efficient CPM, strong creative signals, and cheap cost per conversation.' };
    if (creativeStrong && cpaPts !== null && !cheapConvo) return { action: 'Speed up / simplify the messaging flow', why: 'People are engaging with the creative, but conversations are expensive — the drop-off is likely in the chat/response flow, not the ad.' };
    if (creativeWeak) return { action: 'Replace creative', why: creativeWeakWhy };
    if (!lowCpm && creativeStrong) return { action: 'New audience + new creative', why: 'High CPM despite decent engagement usually means the audience is saturated.' };
    return { action: 'Keep testing', why: 'Mixed signals — compare against your other ads in this set before deciding.' };
  }

  // generic (Awareness / Traffic / Leads / anything not Sales or Messaging) — the PDF has
  // no decision matrix for this bucket, so this stays deliberately conservative: only a
  // clear scale/replace call, "Keep testing" for everything in between.
  if (lowCpm && creativeStrong) return { action: 'Scale aggressively', why: 'Efficient CPM and strong ' + creativeLabel + ' signals.' };
  if (creativeWeak) return { action: 'Replace creative', why: creativeWeakWhy };
  return { action: 'Keep testing', why: 'This objective has no scale/replace matrix in your benchmark doc — judge this one against its sibling ads rather than an absolute rule.' };
}