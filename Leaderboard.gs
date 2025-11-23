/************************************************
 * LEADERBOARD – SERVER SIDE
 * Depends on:
 *   - CONFIG.SCORE_SHEET          (weekly scoreboard)
 *   - CONFIG.MASTER_SHEET         (masterlist)
 *   - CONFIG.COMPLETE_ROUTES_SHEET (optional) or "Complete route list"
 ************************************************/

// Fallbacks in case CONFIG isn't defined
var LB_CONFIG = {
  SCORE_SHEET:
    (typeof CONFIG !== "undefined" && CONFIG.SCORE_SHEET) || "Da_Scoreboard",
  MASTER_SHEET:
    (typeof CONFIG !== "undefined" && CONFIG.MASTER_SHEET) || "masterlist",
  COMPLETE_ROUTES_SHEET:
    (typeof CONFIG !== "undefined" && CONFIG.COMPLETE_ROUTES_SHEET) ||
    "Complete route list",
};

// ScriptProperties key
var LB_PROP_KEY = "leaderboardConfig_v1";

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
  var minWeeks = parseInt(cfg.minWeeks, 10);
  if (!isFinite(minWeeks) || minWeeks < 0) minWeeks = 0;

  // Normalise DCR/POD/CC weights so they sum to 1
  var total = d + p + c;
  if (total <= 0) {
    d = 0.4;
    p = 0.4;
    c = 0.2;
    total = 1.0;
  }
  d = d / total;
  p = p / total;
  c = c / total;

  // Extra weights for volume / weeks / rescues
  var vol = Number(cfg.volumeWeight) || 0.5;
  var wWeeks = Number(cfg.weeksWeight) || 0.3;
  var wResGiven = Number(cfg.rescuesGivenWeight) || 1.0;
  var wResTaken = Number(cfg.rescuesTakenWeight) || 1.0;

  var toSave = {
    dcrWeight: d,
    podWeight: p,
    ccWeight: c,
    minWeeks: minWeeks,
    volumeWeight: vol,
    weeksWeight: wWeeks,
    rescuesGivenWeight: wResGiven,
    rescuesTakenWeight: wResTaken,
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
 */
function getLeaderboardData() {
  var ss = SpreadsheetApp.getActive();
  var scoreSheet = ss.getSheetByName(LB_CONFIG.SCORE_SHEET);
  var masterSheet = ss.getSheetByName(LB_CONFIG.MASTER_SHEET);
  var completeSheet = ss.getSheetByName(LB_CONFIG.COMPLETE_ROUTES_SHEET);

  if (!scoreSheet)
    throw new Error("Scoreboard sheet not found: " + LB_CONFIG.SCORE_SHEET);
  if (!masterSheet)
    throw new Error("Masterlist sheet not found: " + LB_CONFIG.MASTER_SHEET);

  var cfg = getLeaderboardConfig_();

  /********** MASTERLIST – DRIVER LIST + NAME→ID MAP **********/
  var mLast = masterSheet.getLastRow();
  var masterVals =
    mLast > 1 ? masterSheet.getRange(2, 1, mLast - 1, 6).getValues() : [];

  // masterlist layout:
  // A: Transporter ID
  // B: First Name
  // C: Last Name
  // F: Status (Active / Inactive)
  var drivers = [];
  var nameToIds = {}; // "firstname lastname".toLowerCase() -> [transporterIds]

  for (var i = 0; i < masterVals.length; i++) {
    var row = masterVals[i];
    var trId = String(row[0] || "").trim();
    var first = String(row[1] || "").trim();
    var last = String(row[2] || "").trim();
    var status = String(row[5] || "").trim(); // F

    var name = (first + " " + last).trim();
    if (!name) name = "(Unnamed driver " + (i + 2) + ")";

    drivers.push({
      id: trId || null,
      name: name,
      status: status || "",
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
    var trId2 = String(sRow[5] || "").trim();
    if (!trId2) continue; // skip rows with no transporter ID

    var delivered = Number(sRow[6]) || 0;
    var dcrP = parsePercent_(sRow[7]);
    var podP = parsePercent_(sRow[9]);
    var ccP = parsePercent_(sRow[10]);

    if (!statsById[trId2]) {
      statsById[trId2] = {
        deliveries: 0,
        weeksSet: {},
        dcrSum: 0,
        dcrCount: 0,
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
  if (completeSheet) {
    var cLast = completeSheet.getLastRow();
    if (cLast > 1) {
      var cVals = completeSheet.getRange(2, 1, cLast - 1, 13).getValues();
      for (var j = 0; j < cVals.length; j++) {
        var cRow = cVals[j];
        var routeTid = String(cRow[3] || "").trim(); // D = transporter ID
        var rescuedByRaw = String(cRow[12] || "").trim(); // M = "Rescued by"

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
    var dcrAvg = stat && stat.dcr != null ? stat.dcr : null;
    var podAvg = stat && stat.pod != null ? stat.pod : null;
    var ccAvg = stat && stat.cc != null ? stat.cc : null;
    var rescGiven = stat ? stat.rescuesGiven || 0 : 0;
    var rescTaken = stat ? stat.rescuesTaken || 0 : 0;

    var score = computeScore_(
      {
        deliveries: deliveries,
        weeks: weeks,
        dcr: dcrAvg,
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
      weeks: weeks,
      deliveries: deliveries,
      dcr: dcrAvg != null ? roundPct_(dcrAvg) : null,
      pod: podAvg != null ? roundPct_(podAvg) : null,
      cc: ccAvg != null ? roundPct_(ccAvg) : null,
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

  var result = {
    summary: {
      totalDrivers: drivers.length,
      activeDrivers: activeCount,
      totalDeliveries: totalDeliveriesAll,
      avgDcr: avgDcrAll != null ? roundPct_(avgDcrAll) : null,
      avgPod: avgPodAll != null ? roundPct_(avgPodAll) : null,
    },
    rows: rows,
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
        // Ensure extra weights exist with defaults
        if (typeof cfg.volumeWeight !== "number") cfg.volumeWeight = 0.5;
        if (typeof cfg.weeksWeight !== "number") cfg.weeksWeight = 0.3;
        if (typeof cfg.rescuesGivenWeight !== "number")
          cfg.rescuesGivenWeight = 1.0;
        if (typeof cfg.rescuesTakenWeight !== "number")
          cfg.rescuesTakenWeight = 1.0;
        if (typeof cfg.minWeeks !== "number") cfg.minWeeks = 1;
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
    minWeeks: 1,
    volumeWeight: 0.5,
    weeksWeight: 0.3,
    rescuesGivenWeight: 1.0,
    rescuesTakenWeight: 1.0,
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
  var qDcr = stats.dcr * 100 * cfg.dcrWeight;
  var qPod = stats.pod * 100 * cfg.podWeight;
  var qCc = stats.cc * 100 * cfg.ccWeight;
  var quality = qDcr + qPod + qCc;

  // Volume & experience – now fully weight-controlled
  var volume = Math.log(1 + deliveries) * (cfg.volumeWeight || 0);
  var experience = weeks * (cfg.weeksWeight || 0);

  // Rescues: given = positive, taken = negative
  var rescG = (stats.rescuesGiven || 0) * (cfg.rescuesGivenWeight || 0);
  var rescT = (stats.rescuesTaken || 0) * (cfg.rescuesTakenWeight || 0);

  var score = quality + volume + experience + rescG - rescT;
  return Math.round(score * 10) / 10; // one decimal place
}
