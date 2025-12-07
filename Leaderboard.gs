/************************************************
 * LEADERBOARD – SERVER SIDE
 * Depends on:
 *   - CONFIG.SCORE_SHEET          (weekly scoreboard)
 *   - CONFIG.MASTER_SHEET         (masterlist)
 *   - CONFIG.COMPLETE_ROUTES_SHEET (optional) or "Complete route list"
 ************************************************/

// Logical keys for sheet config (must be explicitly set in the sidebar).
var LB_KEYS = {
  SCORE_SHEET: "scoreSheet",
  MASTER_SHEET: "masterSheet",
  COMPLETE_ROUTES_SHEET: "completeSheet",
};

// ScriptProperties key
var LB_PROP_KEY = "leaderboardConfig_v1";
// Quality rating thresholds
var QUALITY_THRESHOLDS = {
  percent: { fantastic: 0.96, onTarget: 0.85 }, // for DCR / POD / CC (fractions 0-1)
  dpmo: { fantastic: 700, onTarget: 1650 }, // for DNR DPMO (lower is better)
};

/************************************************
 * SHARED HELPERS (define once if missing)
 ************************************************/
// Round a number to one decimal place (fallback if not defined elsewhere).
if (typeof roundToOne_ !== "function") {
  function roundToOne_(n) {
    var num = Number(n);
    if (!isFinite(num)) return null;
    return Math.round(num * 10) / 10;
  }
}

// Safe date parser; defines parseDateCell_ if missing to avoid ReferenceErrors.
if (typeof parseDateCell_ !== "function") {
  function parseDateCell_(value) {
    return parseDateCellFallback_(value);
  }
}
function parseDateCellSafe_(cell) {
  if (typeof parseDateCell_ === "function") {
    try {
      return parseDateCell_(cell);
    } catch (e) {
      // fall through to fallback
    }
  }
  return parseDateCellFallback_(cell);
}
function parseDateCellFallback_(cell) {
  if (cell instanceof Date) {
    return new Date(cell.getTime());
  }
  if (typeof cell === "number" && isFinite(cell)) {
    // Sheets serial date (days since 1899-12-30)
    var epoch = new Date(Date.UTC(1899, 11, 30));
    var ms = cell * 24 * 60 * 60 * 1000;
    return new Date(epoch.getTime() + ms);
  }
  if (cell && typeof cell === "string") {
    var d = new Date(cell);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

// Rating helpers
function ratePercent_(fraction) {
  if (fraction == null || isNaN(fraction)) return "";
  if (fraction >= QUALITY_THRESHOLDS.percent.fantastic) return "Fantastic";
  if (fraction >= QUALITY_THRESHOLDS.percent.onTarget) return "On Target";
  return "Below Target";
}

function rateDpmo_(dpmo) {
  if (dpmo == null || isNaN(dpmo)) return "";
  if (dpmo < QUALITY_THRESHOLDS.dpmo.fantastic) return "Fantastic";
  if (dpmo < QUALITY_THRESHOLDS.dpmo.onTarget) return "On Target";
  return "Below Target";
}

function dnrQualityFraction_(dpmo) {
  if (dpmo == null || isNaN(dpmo)) return null;
  var best = QUALITY_THRESHOLDS.dpmo.fantastic;
  var okMax = QUALITY_THRESHOLDS.dpmo.onTarget;
  if (dpmo < best) return 1;
  if (dpmo > okMax) return 0;
  var span = okMax - best;
  if (span <= 0) return 0;
  var remaining = okMax - dpmo;
  return Math.max(0, Math.min(1, remaining / span));
}

/**
 * Build a Gaussian curve scaled to histogram counts.
 * y is scaled so that the peak roughly matches the histogram peak.
 */
function buildGaussianCurve_(scores, stats, bucketSize, minScore, maxScore, peakCount) {
  if (!stats || stats.stddev == null || stats.stddev === 0 || !scores || !scores.length) return [];
  var mean = (typeof stats.meanRaw === "number") ? stats.meanRaw : stats.mean;
  var sd = (typeof stats.stddevRaw === "number") ? stats.stddevRaw : stats.stddev;
  if (!sd || sd === 0) return [];
  var n = scores.length;
  var width = bucketSize && bucketSize > 0 ? bucketSize : Math.max(0.25, (maxScore - minScore) / 20 || 1);
  var start = (typeof minScore === "number" ? minScore : Math.min.apply(null, scores)) - width;
  var end = (typeof maxScore === "number" ? maxScore : Math.max.apply(null, scores)) + width;
  var steps = 80;
  var points = [];
  var normCoef = 1 / (sd * Math.sqrt(2 * Math.PI));
  var pdfPeak = normCoef; // at mean
  var targetPeak = (typeof peakCount === "number" && peakCount > 0) ? peakCount : n * width;
  var scale = pdfPeak > 0 ? targetPeak / pdfPeak : 1;
  for (var i = 0; i <= steps; i++) {
    var x = start + (i / steps) * (end - start);
    var expPart = -((x - mean) * (x - mean)) / (2 * sd * sd);
    var pdf = normCoef * Math.exp(expPart);
    points.push({ x: x, y: pdf * scale });
  }
  return points;
}

/************************************************
 * PUBLIC API – CALLED FROM HTML / SIDEBAR
 ************************************************/

/**
 * Return leaderboard configuration (weights etc.).
 */
function getLeaderboardConfig() {
  return getLeaderboardConfig_();
}

/**
 * Save leaderboard config from the sidebar sliders.
 * @param {Object} cfg
 */
function saveLeaderboardConfig(cfg) {
  cfg = cfg || {};

  var d = Number(cfg.dcrWeight || cfg.weightDcr) || 0;
  var p = Number(cfg.podWeight || cfg.weightPod) || 0;
  var c = Number(cfg.ccWeight || cfg.weightCc) || 0;
  var dn = Number(cfg.dnrWeight || cfg.weightDnr) || 0;
  var minWeeks = parseInt(cfg.minWeeks, 10);
  if (!isFinite(minWeeks) || minWeeks < 0) minWeeks = 0;

  // Normalise DCR/POD/CC/DNR weights so they sum to 1
  var total = d + p + c + dn;
  if (total <= 0) {
    d = 0.35;
    p = 0.35;
    c = 0.15;
    dn = 0.15;
    total = 1.0;
  }
  d = d / total;
  p = p / total;
  c = c / total;
  dn = dn / total;

  // Extra weights for volume / weeks / rescues
  var vol = Number(cfg.volumeWeight) || 0.5;
  var wWeeks = Number(cfg.weeksWeight) || 0.3;
  var wResGiven = Number(cfg.rescuesGivenWeight) || 1.0;
  var wResTaken = Number(cfg.rescuesTakenWeight) || 1.0;
  var volTarget = Number(cfg.volumeTarget) || 4000;
  var weeksTarget = Number(cfg.weeksTarget) || 12;
  var rescueCap = Number(cfg.rescueCap) || 5;
  // Adjust defaults to better scaling if not set
  if (!cfg || !cfg.hasOwnProperty("volumeTarget")) volTarget = 1500;
  if (!cfg || !cfg.hasOwnProperty("weeksTarget")) weeksTarget = 8;
  if (!cfg || !cfg.hasOwnProperty("rescueCap")) rescueCap = 3;

  var toSave = {
    dcrWeight: d,
    podWeight: p,
    ccWeight: c,
    dnrWeight: dn,
    minWeeks: minWeeks,
    volumeWeight: vol,
    weeksWeight: wWeeks,
    rescuesGivenWeight: wResGiven,
    rescuesTakenWeight: wResTaken,
    volumeTarget: volTarget,
    weeksTarget: weeksTarget,
    rescueCap: rescueCap,
  };

  PropertiesService.getScriptProperties().setProperty(
    LB_PROP_KEY,
    JSON.stringify(toSave)
  );

  return toSave;
}

/**
 * Build the leaderboard data model.
 * Called by leaderboard_page.html via google.script.run.
 * @param {Number|null} weekFilter Optional week number to filter by.
 */
function getLeaderboardData(weekFilter) {
  var scoreSheet = resolveLbSheet_(LB_KEYS.SCORE_SHEET);
  var masterSheet = resolveLbSheet_(LB_KEYS.MASTER_SHEET);
  var completeSheet = resolveLbSheet_(LB_KEYS.COMPLETE_ROUTES_SHEET, { optional: true });

  if (!scoreSheet || !masterSheet) {
    return {
      summary: {
        totalDrivers: 0,
        activeDrivers: 0,
        totalDeliveries: 0,
        avgDcr: null,
        avgPod: null,
      },
      rows: [],
    };
  }

  var cfg = getLeaderboardConfig_();
  var tz = scoreSheet.getParent().getSpreadsheetTimeZone() || "Etc/UTC";

  /********** MASTERLIST – DRIVER LIST + NAME→ID MAP **********/
  var mLast = masterSheet.getLastRow();
  var masterVals =
    mLast > 1 ? masterSheet.getRange(2, 1, mLast - 1, 8).getValues() : [];

  // masterlist layout:
  // A: Transporter ID
  // B: First Name
  // C: Last Name
  // F: Status (Active / Inactive)
  // G: Email
  // H: Nursery (Yes/No)
  var drivers = [];
  var nameToIds = {}; // "firstname lastname".toLowerCase() -> [transporterIds]

  for (var i = 0; i < masterVals.length; i++) {
    var row = masterVals[i];
    var trId = String(row[0] || "").trim();
    var first = String(row[1] || "").trim();
    var last = String(row[2] || "").trim();
    var status = String(row[5] || "").trim(); // F
    var email = String(row[6] || "").trim(); // G
    var nurseryRaw = String(row[7] || "").trim(); // H
    var isNurseryFlag = /^yes$/i.test(nurseryRaw);

    var name = (first + " " + last).trim();
    if (!name) name = "(Unnamed driver " + (i + 2) + ")";

    drivers.push({
      id: trId || null,
      name: name,
      status: status || "",
      email: email || "",
      isNursery: isNurseryFlag,
    });

    if (trId) {
      var key = name.toLowerCase();
      if (!nameToIds[key]) nameToIds[key] = [];
      if (nameToIds[key].indexOf(trId) === -1) {
        nameToIds[key].push(trId);
      }
    }
  }

  /********** SCOREBOARD – WEEKLY METRICS **********/
  var sLast = scoreSheet.getLastRow();
  var scoreVals =
    sLast > 1 ? scoreSheet.getRange(2, 1, sLast - 1, 11).getValues() : [];
  var weekOptions = buildWeekOptions_(scoreVals, tz);
  var selectedWeek =
    weekFilter == null || weekFilter === "" || isNaN(Number(weekFilter))
      ? null
      : Number(weekFilter);
  var prevWeek = findPreviousWeek_(selectedWeek, weekOptions);

  // Da_Scoreboard layout:
  // A: WK
  // B: WE Date
  // C: Driver Name
  // D: Status
  // E: spacer
  // F: Transporter ID
  // G: Delivered (number)
  // H: DCR  (percent)
  // I: DNR DPMO (ignored here)
  // J: POD  (percent)
  // K: CC   (percent)

  var statsById = {};
  var totalDeliveriesAll = 0;
  var dcrSumAll = 0,
    dcrCountAll = 0;
  var podSumAll = 0,
    podCountAll = 0;

  for (var r = 0; r < scoreVals.length; r++) {
    var sRow = scoreVals[r];
    var wk = sRow[0];
    if (selectedWeek !== null && wk !== selectedWeek) continue;
    var trId2 = String(sRow[5] || "").trim();
    if (!trId2) continue; // skip rows with no transporter ID

    var delivered = Number(sRow[6]) || 0;
    var dcrP = parsePercent_(sRow[7]);
    var dnrDpmo = sRow[8] != null && sRow[8] !== "" ? Number(sRow[8]) : null;
    var podP = parsePercent_(sRow[9]);
    var ccP = parsePercent_(sRow[10]);

    if (!statsById[trId2]) {
      statsById[trId2] = {
        deliveries: 0,
        weeksSet: {},
        dcrSum: 0,
        dcrCount: 0,
        dnrDefects: 0,
        dnrDelivered: 0,
        podSum: 0,
        podCount: 0,
        ccSum: 0,
        ccCount: 0,
        rescuesGiven: 0,
        rescuesTaken: 0,
      };
    }
    var s = statsById[trId2];
    s.deliveries += delivered;
    if (wk !== "" && wk != null) s.weeksSet[String(wk)] = true;

    if (dcrP != null) {
      s.dcrSum += dcrP;
      s.dcrCount++;
      dcrSumAll += dcrP;
      dcrCountAll++;
    }
    if (dnrDpmo != null && delivered) {
      s.dnrDefects += (dnrDpmo * delivered) / 1000000;
      s.dnrDelivered += delivered;
    }
    if (podP != null) {
      s.podSum += podP;
      s.podCount++;
      podSumAll += podP;
      podCountAll++;
    }
    if (ccP != null) {
      s.ccSum += ccP;
      s.ccCount++;
    }

    totalDeliveriesAll += delivered;
  }

  /********** RESCUES – FROM COMPLETE ROUTE LIST **********/
  // We assume:
  // - "Complete route list" has at least:
  //   D: Transporter ID  (col 4, index 3)
  //   M: Rescued by      (col 13, index 12) with names "Name1|Name2"
  var rescueById = {};
  var routeWindow = selectedWeek !== null ? findWeekDateWindow_(selectedWeek, scoreVals, tz) : null;
  if (completeSheet) {
    var cLast = completeSheet.getLastRow();
    if (cLast > 1) {
      var cVals = completeSheet.getRange(2, 1, cLast - 1, 13).getValues();
      for (var j = 0; j < cVals.length; j++) {
        var cRow = cVals[j];
        var routeTid = String(cRow[3] || "").trim(); // D = transporter ID
        var rescuedByRaw = String(cRow[12] || "").trim(); // M = "Rescued by"
        var routeDate = parseDateCellSafe_(cRow[1]); // B = Date

        if (routeWindow && (!routeDate || routeDate < routeWindow.start || routeDate > routeWindow.end)) {
          continue;
        }

        if (!routeTid && !rescuedByRaw) continue;

        if (!rescueById[routeTid] && routeTid) {
          rescueById[routeTid] = { given: 0, taken: 0 };
        }

        // If there are any rescues on this route, original driver "took" a rescue
        if (routeTid && rescuedByRaw) {
          rescueById[routeTid].taken++;
        }

        if (!rescuedByRaw) continue;

        var parts = rescuedByRaw.split("|");
        for (var k = 0; k < parts.length; k++) {
          var name = String(parts[k] || "").trim();
          if (!name) continue;
          var key = name.toLowerCase();
          var idList = nameToIds[key];
          if (!idList || !idList.length) continue;

          // Credit each matching transporter ID as having "given" a rescue
          for (var x = 0; x < idList.length; x++) {
            var rid = idList[x];
            if (!rid) continue;
            if (!rescueById[rid]) rescueById[rid] = { given: 0, taken: 0 };
            rescueById[rid].given++;
          }
        }
      }
    }
  }

  /********** FINAL PER-DRIVER AGGREGATE **********/
  var aggById = {};
  Object.keys(statsById).forEach(function (id) {
    var base = statsById[id];
    var weeks = Object.keys(base.weeksSet).length;

    var rescueStats = rescueById[id] || { given: 0, taken: 0 };

    aggById[id] = {
      deliveries: base.deliveries,
      weeks: weeks,
      dcr: base.dcrCount ? base.dcrSum / base.dcrCount : null,
      dnrDpmo:
        base.dnrDelivered > 0
          ? (base.dnrDefects / base.dnrDelivered) * 1000000
          : null,
      pod: base.podCount ? base.podSum / base.podCount : null,
      cc: base.ccCount ? base.ccSum / base.ccCount : null,
      rescuesGiven: rescueStats.given || 0,
      rescuesTaken: rescueStats.taken || 0,
    };
  });

  // Also ensure drivers that only appear in rescues (never in scoreboard)
  Object.keys(rescueById).forEach(function (id) {
    if (!id) return;
    if (!aggById[id]) {
      var rescueStats = rescueById[id];
      aggById[id] = {
        deliveries: 0,
        weeks: 0,
        dcr: null,
        dnrDpmo: null,
        pod: null,
        cc: null,
        rescuesGiven: rescueStats.given || 0,
        rescuesTaken: rescueStats.taken || 0,
      };
    }
  });

  var avgDcrAll = dcrCountAll ? dcrSumAll / dcrCountAll : null;
  var avgPodAll = podCountAll ? podSumAll / podCountAll : null;

  /********** BUILD ROWS FOR UI **********/
  var rows = [];
  var activeCount = 0;

  for (var d = 0; d < drivers.length; d++) {
    var drv = drivers[d];
    var id = drv.id;
    var stat = id ? aggById[id] : null;

    if (/^active$/i.test(drv.status)) activeCount++;

    var weeks = stat ? stat.weeks : 0;
    var deliveries = stat ? stat.deliveries : 0;
    var isNursery = drv && drv.isNursery ? true : false;
    var dcrAvg = stat && stat.dcr != null ? stat.dcr : null;
    var dnrDpmo = stat && stat.dnrDpmo != null ? stat.dnrDpmo : null;
    var podAvg = stat && stat.pod != null ? stat.pod : null;
    var ccAvg = stat && stat.cc != null ? stat.cc : null;
    var rescGiven = stat ? stat.rescuesGiven || 0 : 0;
    var rescTaken = stat ? stat.rescuesTaken || 0 : 0;

    var score = computeScore_(
      {
        deliveries: deliveries,
        weeks: weeks,
        dcr: dcrAvg,
        dnrDpmo: dnrDpmo,
        pod: podAvg,
        cc: ccAvg,
        rescuesGiven: rescGiven,
        rescuesTaken: rescTaken,
      },
      cfg
    );

    rows.push({
      name: drv.name,
      transporterId: id || "",
      status: drv.status || "",
      email: drv.email || "",
      weeks: weeks,
      isNursery: isNursery,
      deliveries: deliveries,
      dcr: dcrAvg != null ? roundPct_(dcrAvg) : null,
      dcrRating: ratePercent_(dcrAvg),
      dnrDpmo: dnrDpmo != null ? roundToOne_(dnrDpmo) : null,
      dnrRating: rateDpmo_(dnrDpmo),
      pod: podAvg != null ? roundPct_(podAvg) : null,
      podRating: ratePercent_(podAvg),
      cc: ccAvg != null ? roundPct_(ccAvg) : null,
      ccRating: ratePercent_(ccAvg),
      rescuesGiven: rescGiven,
      rescuesTaken: rescTaken,
      score: score,
    });
  }

  // Sort: ranked drivers with score first, then NA, by name.
  rows.sort(function (a, b) {
    var sa = a.score;
    var sb = b.score;
    if (sa == null && sb == null) {
      return a.name.localeCompare(b.name);
    }
    if (sa == null) return 1;
    if (sb == null) return -1;
    return sb - sa; // descending
  });

  // Assign rank numbers to rows with a score
  var rank = 1;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].score == null) {
      rows[i].rank = null;
    } else {
      rows[i].rank = rank++;
    }
  }

  // Rank deltas compared to previous week (only when a specific week is selected and a prior week exists)
  if (selectedWeek != null && prevWeek != null) {
    var currentRankMap = {};
    rows.forEach(function (r) {
      if (r.transporterId && r.rank != null) {
        currentRankMap[r.transporterId] = r.rank;
      }
    });
    var prevRankMap = computeRanksForWeek_(prevWeek, scoreVals, cfg, nameToIds, completeSheet, tz);
    rows.forEach(function (r) {
      var prevRank = prevRankMap[r.transporterId];
      if (prevRank != null && r.rank != null) {
        r.rankChange = prevRank - r.rank; // positive = moved up
      } else {
        r.rankChange = null;
      }
    });
  }

  var result = {
    summary: {
      totalDrivers: drivers.length,
      activeDrivers: activeCount,
      totalDeliveries: totalDeliveriesAll,
      avgDcr: avgDcrAll != null ? roundPct_(avgDcrAll) : null,
      avgPod: avgPodAll != null ? roundPct_(avgPodAll) : null,
    },
    rows: rows,
    weekOptions: [{ value: null, label: "Overall (all weeks)" }].concat(weekOptions),
    appliedWeek: selectedWeek,
    distribution: buildDistribution_(rows),
    distributionHistory: buildDistributionHistory_(
      selectedWeek,
      weekOptions,
      scoreVals,
      cfg,
      nameToIds,
      completeSheet,
      tz
    ),
  };

  return result;
}

/**
 * Open the leaderboard as a modal dialog from the sidebar.
 * This is what the "Open leaderboard" button calls.
 */
function openLeaderboardDialog() {
  var html = HtmlService.createHtmlOutputFromFile("leaderboard_page")
    .setWidth(1200)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, "Driver leaderboard");
}

/**
 * Optional: allow deploying as a standalone web app if you want.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("leaderboard_page");
}

/************************************************
 * INTERNAL HELPERS
 ************************************************/

function getLeaderboardConfig_() {
  var raw = PropertiesService.getScriptProperties().getProperty(LB_PROP_KEY);
  if (raw) {
    try {
      var cfg = JSON.parse(raw);
      if (
        typeof cfg.dcrWeight === "number" &&
        typeof cfg.podWeight === "number" &&
        typeof cfg.ccWeight === "number"
      ) {
        if (typeof cfg.dnrWeight !== "number") cfg.dnrWeight = 0.15;
        // Renormalise the four weights
        var total = cfg.dcrWeight + cfg.podWeight + cfg.ccWeight + cfg.dnrWeight;
        if (total <= 0) {
          cfg.dcrWeight = 0.35;
          cfg.podWeight = 0.35;
          cfg.ccWeight = 0.15;
          cfg.dnrWeight = 0.15;
        } else {
          cfg.dcrWeight = cfg.dcrWeight / total;
          cfg.podWeight = cfg.podWeight / total;
          cfg.ccWeight = cfg.ccWeight / total;
          cfg.dnrWeight = cfg.dnrWeight / total;
        }
        // Ensure extra weights exist with defaults
        if (typeof cfg.volumeWeight !== "number") cfg.volumeWeight = 0.5;
        if (typeof cfg.weeksWeight !== "number") cfg.weeksWeight = 0.3;
        if (typeof cfg.rescuesGivenWeight !== "number")
          cfg.rescuesGivenWeight = 1.0;
        if (typeof cfg.rescuesTakenWeight !== "number")
          cfg.rescuesTakenWeight = 1.0;
        if (typeof cfg.minWeeks !== "number") cfg.minWeeks = 1;
        if (typeof cfg.volumeTarget !== "number") cfg.volumeTarget = 4000;
        if (typeof cfg.weeksTarget !== "number") cfg.weeksTarget = 12;
        if (typeof cfg.rescueCap !== "number") cfg.rescueCap = 5;
        return cfg;
      }
    } catch (err) {
      // fall through to defaults
    }
  }
  // Defaults
  return {
    dcrWeight: 0.4,
    podWeight: 0.4,
    ccWeight: 0.2,
    dnrWeight: 0.15,
    minWeeks: 1,
    volumeWeight: 0.5,
    weeksWeight: 0.3,
    rescuesGivenWeight: 1.0,
    rescuesTakenWeight: 1.0,
    volumeTarget: 1500,
    weeksTarget: 8,
    rescueCap: 3,
  };
}

/**
 * Convert percent strings/numbers to 0–1.
 * Accepts "98.3%", 0.983, 98.3, etc.
 */
function parsePercent_(val) {
  if (val === null || val === "" || typeof val === "undefined") return null;

  if (typeof val === "number") {
    if (val > 2) return val / 100;
    return val; // assume already 0–1
  }

  var s = String(val).trim();
  if (!s) return null;
  if (s.indexOf("%") !== -1) {
    s = s.replace("%", "");
  }
  var n = parseFloat(s);
  if (!isFinite(n)) return null;
  if (n > 2) n = n / 100;
  return n;
}

/**
 * Round 0–1 fraction to one decimal percent (e.g. 0.985 → 98.5).
 */
function roundPct_(fraction) {
  return Math.round(fraction * 1000) / 10; // one decimal place
}

/**
 * Compute the overall leaderboard score from stats + config.
 * stats: {
 *   deliveries, weeks, dcr, pod, cc, rescuesGiven, rescuesTaken
 * }
 * dcr/pod/cc are 0–1 fractions.
 */
function computeScore_(stats, cfg) {
  if (!stats) return null;
  if (stats.dcr == null || stats.pod == null || stats.cc == null) return null;

  var weeks = stats.weeks || 0;
  var deliveries = stats.deliveries || 0;

  if (weeks < (cfg.minWeeks || 0)) {
    return null; // not enough history
  }

  // Quality component (0–100 each * normalised weights)
  var items = [];
  if (stats.dcr != null) items.push({ value: stats.dcr, weight: cfg.dcrWeight });
  if (stats.pod != null) items.push({ value: stats.pod, weight: cfg.podWeight });
  if (stats.cc != null) items.push({ value: stats.cc, weight: cfg.ccWeight });
  if (stats.dnrDpmo != null) {
    var qDnr = dnrQualityFraction_(stats.dnrDpmo);
    if (qDnr != null) items.push({ value: qDnr, weight: cfg.dnrWeight });
  }

  var weightSum = 0;
  var qualityFraction = 0;
  for (var i = 0; i < items.length; i++) {
    var w = items[i].weight || 0;
    if (w <= 0) continue;
    weightSum += w;
    qualityFraction += items[i].value * w;
  }
  if (weightSum <= 0) return null;
  var qualityScore = qualityFraction / weightSum; // 0-1

  // Volume (0-1, capped)
  var volTarget = cfg.volumeTarget || 1500;
  var volDen = Math.log(1 + Math.max(volTarget, 1));
  var volumeScore =
    volDen > 0 ? Math.min(1, Math.log(1 + Math.max(deliveries, 0)) / volDen) : 0;

  // Experience (0-1)
  var weeksTarget = cfg.weeksTarget || 8;
  var experienceScore =
    weeksTarget > 0 ? Math.min(1, Math.max(0, weeks / weeksTarget)) : 0;

  // Rescues (0-1), neutral at 0.5 for net zero; no rescues => no weight
  var rescueCap = cfg.rescueCap || 3;
  var g = stats.rescuesGiven || 0;
  var t = stats.rescuesTaken || 0;
  var netRescues =
    g * (cfg.rescuesGivenWeight || 0) -
    t * (cfg.rescuesTakenWeight || 0);
  var resActive = g !== 0 || t !== 0;
  var rescueScore =
    rescueCap > 0
      ? Math.min(1, Math.max(0, (netRescues + rescueCap) / (2 * rescueCap)))
      : 0.5;

  // Top-level weights derived from sliders and normalised to 1
  // Stronger bias to quality, shrink/zero rescue weight when no activity
  var qualityBase = 4;
  var volRaw = Math.max(cfg.volumeWeight || 0, 0);
  var expRaw = Math.max(cfg.weeksWeight || 0, 0);
  var resRawBase = Math.max(
    Math.max(cfg.rescuesGivenWeight || 0, 0),
    Math.max(cfg.rescuesTakenWeight || 0, 0)
  );
  var resRaw = resActive ? Math.min(resRawBase, 0.5) : 0; // cap to avoid overpowering
  var totalTop = qualityBase + volRaw + expRaw + resRaw;
  var wQuality = qualityBase / totalTop;
  var wVolume = volRaw / totalTop;
  var wExperience = expRaw / totalTop;
  var wRescue = resRaw / totalTop;

  var combined =
    wQuality * qualityScore +
    wVolume * volumeScore +
    wExperience * experienceScore +
    wRescue * rescueScore;

  var score = combined * 100;
  return Math.round(score * 10) / 10; // one decimal place, max 100
}

/**
 * Resolve a sheet for leaderboard usage based on sheet config.
 * Requires a sheet name; if not configured, returns null instead of falling back.
 */
function resolveLbSheet_(key, opts) {
  opts = opts || {};
  var cfg = loadSheetConfig_ ? loadSheetConfig_() : {};
  var entry = cfg && cfg[key] ? cfg[key] : null;
  if (!entry || (!entry.sheetName && !entry.spreadsheetId)) {
    if (opts.optional) return null;
    return null;
  }
  var ss = entry.spreadsheetId ? SpreadsheetApp.openById(entry.spreadsheetId) : SpreadsheetApp.getActive();
  var sheetName = entry.sheetName || "";
  if (!sheetName) {
    if (opts.optional) return null;
    return null;
  }
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    if (opts.optional) return null;
    throw new Error("Sheet not found for leaderboard: " + sheetName);
  }
  return sheet;
}

function buildWeekOptions_(scoreVals, tz) {
  var map = {};
  for (var i = 0; i < scoreVals.length; i++) {
    var row = scoreVals[i];
    var wk = row[0];
    if (wk === "" || wk == null) continue;
    var weekNum = Number(wk);
    if (!isFinite(weekNum)) continue;
    var weDate = parseDateCellSafe_(row[1]);
    var labelDate =
      weDate && !isNaN(weDate.getTime())
        ? Utilities.formatDate(weDate, tz, "d MMM yyyy")
        : null;
    if (!map[weekNum]) {
      map[weekNum] = {
        value: weekNum,
        label: labelDate ? "Week " + weekNum + " (" + labelDate + ")" : "Week " + weekNum,
        sortDate: weDate ? weDate.getTime() : 0,
      };
    }
  }
  return Object.keys(map)
    .map(function (k) {
      return map[k];
    })
    .sort(function (a, b) {
      if (a.sortDate !== b.sortDate) return b.sortDate - a.sortDate;
      return b.value - a.value;
    })
    .map(function (item) {
      return { value: item.value, label: item.label };
    });
}

function findWeekDateWindow_(weekNum, scoreVals, tz) {
  var target = null;
  for (var i = 0; i < scoreVals.length; i++) {
    var row = scoreVals[i];
    if (Number(row[0]) === weekNum) {
      var d = parseDateCellSafe_(row[1]);
      if (d && !isNaN(d.getTime())) {
        target = d;
        break;
      }
    }
  }
  if (!target) return null;
  var end = new Date(target.getTime());
  end.setHours(0, 0, 0, 0);
  var start = new Date(end.getTime());
  start.setDate(end.getDate() - 6); // include week window
  return { start: start, end: end, tz: tz };
}

function findPreviousWeek_(weekNum, weekOptions) {
  if (weekNum == null) return null;
  for (var i = 0; i < (weekOptions || []).length; i++) {
    if (weekOptions[i].value === weekNum && i + 1 < weekOptions.length) {
      return weekOptions[i + 1].value;
    }
  }
  return null;
}

function computeRanksForWeek_(weekNum, scoreVals, cfg, nameToIds, completeSheet, tz) {
  if (weekNum == null) return {};
  var statsById = {};
  for (var i = 0; i < scoreVals.length; i++) {
    var row = scoreVals[i];
    var wk = row[0];
    if (wk !== weekNum) continue;
    var id = String(row[5] || "").trim();
    if (!id) continue;
    var delivered = Number(row[6]) || 0;
    var dcrP = parsePercent_(row[7]);
    var dnrDpmo = row[8] != null && row[8] !== "" ? Number(row[8]) : null;
    var podP = parsePercent_(row[9]);
    var ccP = parsePercent_(row[10]);
    if (!statsById[id]) {
      statsById[id] = {
        deliveries: 0,
        weeksSet: {},
        dcrSum: 0,
        dcrCount: 0,
        dnrDefects: 0,
        dnrDelivered: 0,
        podSum: 0,
        podCount: 0,
        ccSum: 0,
        ccCount: 0,
        rescuesGiven: 0,
        rescuesTaken: 0,
      };
    }
    var s = statsById[id];
    s.deliveries += delivered;
    s.weeksSet[String(wk)] = true;
    if (dcrP != null) {
      s.dcrSum += dcrP;
      s.dcrCount++;
    }
    if (dnrDpmo != null && delivered) {
      s.dnrDefects += (dnrDpmo * delivered) / 1000000;
      s.dnrDelivered += delivered;
    }
    if (podP != null) {
      s.podSum += podP;
      s.podCount++;
    }
    if (ccP != null) {
      s.ccSum += ccP;
      s.ccCount++;
    }
  }

  // Rescues within the week window
  var rescueById = {};
  var routeWindow = findWeekDateWindow_(weekNum, scoreVals, tz);
  if (completeSheet && routeWindow) {
    var cLast = completeSheet.getLastRow();
    if (cLast > 1) {
      var cVals = completeSheet.getRange(2, 1, cLast - 1, 13).getValues();
      for (var j = 0; j < cVals.length; j++) {
        var cRow = cVals[j];
        var routeTid = String(cRow[3] || "").trim(); // D
        var rescuedByRaw = String(cRow[12] || "").trim(); // M
        var routeDate = parseDateCellSafe_(cRow[1]);
        if (!routeDate || routeDate < routeWindow.start || routeDate > routeWindow.end) continue;
        if (!routeTid && !rescuedByRaw) continue;
        if (routeTid && !rescueById[routeTid]) rescueById[routeTid] = { given: 0, taken: 0 };
        if (routeTid && rescuedByRaw) rescueById[routeTid].taken++;
        if (!rescuedByRaw) continue;
        var parts = rescuedByRaw.split("|");
        for (var k = 0; k < parts.length; k++) {
          var name = String(parts[k] || "").trim();
          if (!name) continue;
          var key = name.toLowerCase();
          var ids = nameToIds[key] || [];
          for (var x = 0; x < ids.length; x++) {
            var rid = ids[x];
            if (!rid) continue;
            if (!rescueById[rid]) rescueById[rid] = { given: 0, taken: 0 };
            rescueById[rid].given++;
          }
        }
      }
    }
  }

  var rows = [];
  Object.keys(statsById).forEach(function (idKey) {
    var base = statsById[idKey];
    var weeks = Object.keys(base.weeksSet).length;
    var rescueStats = rescueById[idKey] || { given: 0, taken: 0 };
    var dcrAvg = base.dcrCount ? base.dcrSum / base.dcrCount : null;
    var dnrDpmo =
      base.dnrDelivered > 0
        ? (base.dnrDefects / base.dnrDelivered) * 1000000
        : null;
    var podAvg = base.podCount ? base.podSum / base.podCount : null;
    var ccAvg = base.ccCount ? base.ccSum / base.ccCount : null;
    var score = computeScore_(
      {
        deliveries: base.deliveries,
        weeks: weeks,
        dcr: dcrAvg,
        dnrDpmo: dnrDpmo,
        pod: podAvg,
        cc: ccAvg,
        rescuesGiven: rescueStats.given || 0,
        rescuesTaken: rescueStats.taken || 0,
      },
      cfg
    );
    rows.push({ id: idKey, score: score });
  });

  rows.sort(function (a, b) {
    if (a.score == null && b.score == null) return 0;
    if (a.score == null) return 1;
    if (b.score == null) return -1;
    return b.score - a.score;
  });

  var rank = 1;
  var ranks = {};
  for (var r = 0; r < rows.length; r++) {
    if (rows[r].score == null) continue;
    ranks[rows[r].id] = rank++;
  }
  return ranks;
}

/**
 * Send scorecard PDFs for a list of transporter IDs.
 * If ids is empty/null, send for all drivers with email addresses.
 */
function sendLeaderboardEmails(ids, options) {
  options = options || {};
  ids = ids || [];
  var weekFilter = options.week != null ? options.week : null;
  var scopeLabel = weekFilter != null ? "Week " + weekFilter : "Overall";
  var subject = options.subject || ("Driver scorecard - " + scopeLabel);
  var body = options.body || "Please find attached the latest driver scorecard.";

  var data = getLeaderboardData(weekFilter);
  var rows = (data && data.rows) || [];
  var targetIds;
  if (!ids.length) {
    targetIds = rows
      .filter(function (r) { return r.email; })
      .map(function (r) { return r.transporterId; });
  } else {
    var set = {};
    ids.forEach(function (id) { if (id) set[id] = true; });
    targetIds = rows
      .filter(function (r) { return set[r.transporterId]; })
      .map(function (r) { return r.transporterId; });
  }

  var sent = [];
  var skipped = [];
  var errors = [];

  for (var i = 0; i < targetIds.length; i++) {
    var tid = targetIds[i];
    var row = rows.find(function (r) { return r.transporterId === tid; });
    if (!row || !row.email) {
      skipped.push({ id: tid, reason: "No email" });
      continue;
    }
    try {
      var pdf = downloadScorecardPdf(tid, weekFilter);
      if (!pdf || !pdf.base64) {
        errors.push({ id: tid, email: row.email, reason: "PDF generation failed" });
        continue;
      }
      var blob = Utilities.newBlob(
        Utilities.base64Decode(pdf.base64),
        pdf.mimeType || "application/pdf",
        pdf.fileName || ("scorecard_" + tid + ".pdf")
      );
      // Use MailApp to avoid Gmail-specific scopes.
      MailApp.sendEmail({
        to: row.email,
        subject: subject,
        body: body,
        attachments: [blob],
      });
      sent.push({ id: tid, email: row.email });
    } catch (err) {
      errors.push({ id: tid, email: row.email, reason: err && err.message ? err.message : String(err) });
    }
  }

  return {
    sent: sent,
    skipped: skipped,
    errors: errors,
    totalRequested: targetIds.length,
    totalSent: sent.length,
  };
}

function buildDistribution_(rows) {
  rows = rows || [];
  var scores = [];
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].score != null && rows[i].score === rows[i].score) {
      scores.push(rows[i].score);
    }
  }
  if (!scores.length) {
    return {
      buckets: [],
      smoothed: [],
      stats: {
        count: 0,
        mean: null,
        median: null,
        stddev: null,
        min: null,
        max: null,
        p25: null,
        p75: null,
        p90: null,
        p95: null,
      },
      bucketSize: null,
      bucketDrivers: [],
      minScore: null,
      maxScore: null,
    };
  }

  scores.sort(function (a, b) { return a - b; });
  var stats = computeStats_(scores);
  var histogram = computeHistogram_(scores, stats, rows);
  var maxCount = histogram.buckets.reduce(function(m, b) { return Math.max(m, b.count || 0); }, 0);
  var curvePoints = buildGaussianCurve_(scores, stats, histogram.bucketSize, histogram.minScore, histogram.maxScore, maxCount);

  return {
    buckets: histogram.buckets,
    smoothed: histogram.smoothed,
    stats: stats,
    bucketSize: histogram.bucketSize,
    bucketDrivers: histogram.bucketDrivers,
    minScore: histogram.minScore,
    maxScore: histogram.maxScore,
    curvePoints: curvePoints,
  };
}

/**
 * Build distribution history (ghost lines) for recent weeks to overlay on the chart.
 * Uses the last 3 weeks prior to the selected one (or last 3 available if none selected).
 */
function buildDistributionHistory_(selectedWeek, weekOptions, scoreVals, cfg, nameToIds, completeSheet, tz) {
  var weeks = [];
  var weekValues = [];
  (weekOptions || []).forEach(function (opt) {
    if (opt && opt.value != null) weekValues.push(opt.value);
  });
  if (!weekValues.length) return [];

  if (selectedWeek != null) {
    var idx = weekValues.indexOf(selectedWeek);
    var start = idx >= 0 ? idx + 1 : 0; // weeks after selected are older
    weeks = weekValues.slice(start, start + 3);
  } else {
    weeks = weekValues.slice(0, 3); // most recent weeks
  }

  var out = [];
  for (var i = 0; i < weeks.length; i++) {
    var wk = weeks[i];
    var scores = computeScoresForWeek_(wk, scoreVals, cfg, nameToIds, completeSheet, tz);
    var dist = buildDistributionFromScores_(scores);
    if (dist) {
      out.push({
        week: wk,
        label: findWeekLabel_(wk, weekOptions),
        smoothed: dist.smoothed,
        stats: dist.stats,
        minScore: dist.minScore,
        maxScore: dist.maxScore,
        curvePoints: dist.curvePoints,
      });
    }
  }
  return out;
}

function findWeekLabel_(weekValue, weekOptions) {
  for (var i = 0; i < (weekOptions || []).length; i++) {
    if (weekOptions[i].value === weekValue) {
      return weekOptions[i].label || ("Week " + weekValue);
    }
  }
  return "Week " + weekValue;
}

function buildDistributionFromScores_(scores) {
  if (!scores || !scores.length) return null;
  var sorted = scores.slice().sort(function (a, b) { return a - b; });
  var stats = computeStats_(sorted);
  var histogram = computeHistogramFromScores_(sorted, stats);
  var maxCount = histogram.buckets.reduce(function(m, b) { return Math.max(m, b.count || 0); }, 0);
  var curvePoints = buildGaussianCurve_(sorted, stats, histogram.bucketSize, histogram.minScore, histogram.maxScore, maxCount);
  return {
    smoothed: histogram.smoothed,
    stats: stats,
    minScore: histogram.minScore,
    maxScore: histogram.maxScore,
    curvePoints: curvePoints,
  };
}

/**
 * Compute histogram and smoothed curve from plain score numbers (no driver rows).
 */
function computeHistogramFromScores_(scores, stats) {
  if (!scores.length) return { buckets: [], smoothed: [], minScore: null, maxScore: null, bucketSize: null };
  var min = scores[0];
  var max = scores[scores.length - 1];
  var range = Math.max(1, max - min);
  var targetBuckets = Math.min(24, Math.max(12, Math.ceil(range / 2)));
  var bucketSize = Math.max(0.25, range / targetBuckets);
  var bucketsCount = Math.ceil(range / bucketSize) + 1;
  var buckets = [];
  for (var i = 0; i < bucketsCount; i++) {
    var start = min + i * bucketSize;
    var end = start + bucketSize;
    buckets.push({ from: start, to: end, count: 0 });
  }
  for (var j = 0; j < scores.length; j++) {
    var s = scores[j];
    var idx = Math.floor((s - min) / bucketSize);
    idx = Math.max(0, Math.min(idx, buckets.length - 1));
    buckets[idx].count++;
  }
  var smoothed = [];
  var window = 2;
  for (var k = 0; k < buckets.length; k++) {
    var startWin = Math.max(0, k - window);
    var endWin = Math.min(buckets.length - 1, k + window);
    var total = 0;
    var c = 0;
    for (var t = startWin; t <= endWin; t++) {
      total += buckets[t].count;
      c++;
    }
    smoothed.push({ x: (buckets[k].from + buckets[k].to) / 2, y: c ? total / c : 0 });
  }
  return {
    buckets: buckets,
    smoothed: smoothed,
    minScore: min,
    maxScore: max,
    bucketSize: bucketSize,
  };
}

/**
 * Compute score numbers for a specific week to use in ghost distributions.
 */
function computeScoresForWeek_(weekNum, scoreVals, cfg, nameToIds, completeSheet, tz) {
  if (weekNum == null) return [];
  var statsById = {};
  for (var i = 0; i < scoreVals.length; i++) {
    var row = scoreVals[i];
    var wk = row[0];
    if (wk !== weekNum) continue;
    var id = String(row[5] || "").trim();
    if (!id) continue;
    var delivered = Number(row[6]) || 0;
    var dcrP = parsePercent_(row[7]);
    var dnrDpmo = row[8] != null && row[8] !== "" ? Number(row[8]) : null;
    var podP = parsePercent_(row[9]);
    var ccP = parsePercent_(row[10]);
    if (!statsById[id]) {
      statsById[id] = {
        deliveries: 0,
        weeksSet: {},
        dcrSum: 0,
        dcrCount: 0,
        dnrDefects: 0,
        dnrDelivered: 0,
        podSum: 0,
        podCount: 0,
        ccSum: 0,
        ccCount: 0,
        rescuesGiven: 0,
        rescuesTaken: 0,
      };
    }
    var s = statsById[id];
    s.deliveries += delivered;
    s.weeksSet[String(wk)] = true;
    if (dcrP != null) {
      s.dcrSum += dcrP;
      s.dcrCount++;
    }
    if (dnrDpmo != null && delivered) {
      s.dnrDefects += (dnrDpmo * delivered) / 1000000;
      s.dnrDelivered += delivered;
    }
    if (podP != null) {
      s.podSum += podP;
      s.podCount++;
    }
    if (ccP != null) {
      s.ccSum += ccP;
      s.ccCount++;
    }
  }

  // Rescues within the week window
  var rescueById = {};
  var routeWindow = findWeekDateWindow_(weekNum, scoreVals, tz);
  if (completeSheet && routeWindow) {
    var cLast = completeSheet.getLastRow();
    if (cLast > 1) {
      var cVals = completeSheet.getRange(2, 1, cLast - 1, 13).getValues();
      for (var j = 0; j < cVals.length; j++) {
        var cRow = cVals[j];
        var routeTid = String(cRow[3] || "").trim(); // D
        var rescuedByRaw = String(cRow[12] || "").trim(); // M
        var routeDate = parseDateCellSafe_(cRow[1]);
        if (!routeDate || routeDate < routeWindow.start || routeDate > routeWindow.end) continue;
        if (!routeTid && !rescuedByRaw) continue;
        if (routeTid && !rescueById[routeTid]) rescueById[routeTid] = { given: 0, taken: 0 };
        if (routeTid && rescuedByRaw) rescueById[routeTid].taken++;
        if (!rescuedByRaw) continue;
        var parts = rescuedByRaw.split("|");
        for (var k = 0; k < parts.length; k++) {
          var name = String(parts[k] || "").trim();
          if (!name) continue;
          var key = name.toLowerCase();
          var ids = nameToIds[key] || [];
          for (var x = 0; x < ids.length; x++) {
            var rid = ids[x];
            if (!rid) continue;
            if (!rescueById[rid]) rescueById[rid] = { given: 0, taken: 0 };
            rescueById[rid].given++;
          }
        }
      }
    }
  }

  var scores = [];
  Object.keys(statsById).forEach(function (idKey) {
    var base = statsById[idKey];
    var weeks = Object.keys(base.weeksSet).length;
    var rescueStats = rescueById[idKey] || { given: 0, taken: 0 };
    var dcrAvg = base.dcrCount ? base.dcrSum / base.dcrCount : null;
    var dnrDpmo =
      base.dnrDelivered > 0
        ? (base.dnrDefects / base.dnrDelivered) * 1000000
        : null;
    var podAvg = base.podCount ? base.podSum / base.podCount : null;
    var ccAvg = base.ccCount ? base.ccSum / base.ccCount : null;
    var score = computeScore_(
      {
        deliveries: base.deliveries,
        weeks: weeks,
        dcr: dcrAvg,
        dnrDpmo: dnrDpmo,
        pod: podAvg,
        cc: ccAvg,
        rescuesGiven: rescueStats.given || 0,
        rescuesTaken: rescueStats.taken || 0,
      },
      cfg
    );
    if (score != null && score === score) scores.push(score);
  });
  return scores;
}

function computeStats_(arr) {
  var n = arr.length;
  var sum = 0;
  for (var i = 0; i < n; i++) sum += arr[i];
  var mean = sum / n;
  var median = n % 2 ? arr[(n - 1) / 2] : (arr[n / 2 - 1] + arr[n / 2]) / 2;
  var min = arr[0];
  var max = arr[n - 1];
  var p25 = percentile_(arr, 0.25);
  var p75 = percentile_(arr, 0.75);
  var variance = 0;
  for (var j = 0; j < n; j++) {
    var diff = arr[j] - mean;
    variance += diff * diff;
  }
  variance = variance / n;
  var stddev = Math.sqrt(variance);
  return {
    count: n,
    meanRaw: mean,
    medianRaw: median,
    stddevRaw: stddev,
    mean: roundToOne_(mean),
    median: roundToOne_(median),
    stddev: roundToOne_(stddev),
    min: roundToOne_(min),
    max: roundToOne_(max),
    p25: roundToOne_(p25),
    p75: roundToOne_(p75),
    p90: roundToOne_(percentile_(arr, 0.9)),
    p95: roundToOne_(percentile_(arr, 0.95)),
  };
}

function percentile_(sortedArr, p) {
  if (!sortedArr || !sortedArr.length) return null;
  var idx = (sortedArr.length - 1) * p;
  var lower = Math.floor(idx);
  var upper = Math.ceil(idx);
  if (lower === upper) return sortedArr[lower];
  return sortedArr[lower] + (sortedArr[upper] - sortedArr[lower]) * (idx - lower);
}

function computeHistogram_(scores, stats, rows) {
  if (!scores.length) return { buckets: [], smoothed: [], bucketSize: null, bucketDrivers: [], minScore: null, maxScore: null };
  var min = scores[0];
  var max = scores[scores.length - 1];
  var range = Math.max(1, max - min);
  // Finer buckets to reduce gaps; aim for 12–24 buckets with minimum width 0.25
  var targetBuckets = Math.min(24, Math.max(12, Math.ceil(range / 2)));
  var bucketSize = Math.max(0.25, range / targetBuckets);
  var bucketsCount = Math.ceil(range / bucketSize) + 1;
  var buckets = [];
  for (var i = 0; i < bucketsCount; i++) {
    var start = min + i * bucketSize;
    var end = start + bucketSize;
    var labelStart = Math.round(start * 10) / 10;
    var labelEnd = Math.round(end * 10) / 10;
    buckets.push({ label: labelStart + "-" + labelEnd, from: start, to: end, count: 0, drivers: [] });
  }
  for (var j = 0; j < rows.length; j++) {
    var row = rows[j];
    if (row.score == null) continue;
    var s = row.score;
    var idx = Math.floor((s - min) / bucketSize);
    idx = Math.max(0, Math.min(idx, buckets.length - 1));
    buckets[idx].count++;
    buckets[idx].drivers.push({
      name: row.name || "",
      transporterId: row.transporterId || "",
      score: row.score,
    });
  }
  // Simple smoothing (moving average over counts)
  var smoothed = [];
  var window = 2;
  for (var k = 0; k < buckets.length; k++) {
    var startWin = Math.max(0, k - window);
    var endWin = Math.min(buckets.length - 1, k + window);
    var total = 0;
    var c = 0;
    for (var t = startWin; t <= endWin; t++) {
      total += buckets[t].count;
      c++;
    }
    smoothed.push({ x: (buckets[k].from + buckets[k].to) / 2, y: c ? total / c : 0 });
  }
  return {
    buckets: buckets,
    smoothed: smoothed,
    bucketSize: bucketSize,
    bucketDrivers: buckets.map(function(b){ return b.drivers; }),
    minScore: min,
    maxScore: max
  };
}
