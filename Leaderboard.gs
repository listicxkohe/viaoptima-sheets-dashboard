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
  var routeWindow = selectedWeek !== null ? findWeekDateWindow_(selectedWeek, scoreVals, tz) : null;
  if (completeSheet) {
    var cLast = completeSheet.getLastRow();
    if (cLast > 1) {
      var cVals = completeSheet.getRange(2, 1, cLast - 1, 13).getValues();
      for (var j = 0; j < cVals.length; j++) {
        var cRow = cVals[j];
        var routeTid = String(cRow[3] || "").trim(); // D = transporter ID
        var rescuedByRaw = String(cRow[12] || "").trim(); // M = "Rescued by"
        var routeDate = parseDateCell_ ? parseDateCell_(cRow[1]) : null; // B = Date

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
    var weDate = typeof parseDateCell_ === "function" ? parseDateCell_(row[1]) : null;
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
      var d = typeof parseDateCell_ === "function" ? parseDateCell_(row[1]) : null;
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
    var podP = parsePercent_(row[9]);
    var ccP = parsePercent_(row[10]);
    if (!statsById[id]) {
      statsById[id] = {
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
    var s = statsById[id];
    s.deliveries += delivered;
    s.weeksSet[String(wk)] = true;
    if (dcrP != null) {
      s.dcrSum += dcrP;
      s.dcrCount++;
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
        var routeDate = parseDateCell_(cRow[1]);
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
    var podAvg = base.podCount ? base.podSum / base.podCount : null;
    var ccAvg = base.ccCount ? base.ccSum / base.ccCount : null;
    var score = computeScore_(
      {
        deliveries: base.deliveries,
        weeks: weeks,
        dcr: dcrAvg,
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

  return {
    buckets: histogram.buckets,
    smoothed: histogram.smoothed,
    stats: stats,
    bucketSize: histogram.bucketSize,
    bucketDrivers: histogram.bucketDrivers,
    minScore: histogram.minScore,
    maxScore: histogram.maxScore,
  };
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
  if (!sortedArr.length) return null;
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
  var window = 1;
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
