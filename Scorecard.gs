/************************************************
 * DRIVER SCORECARD - SERVER SIDE
 ************************************************/

var SCORECARD_CONFIG = {
  SCORE_SHEET:
    (typeof CONFIG !== "undefined" && CONFIG.SCORE_SHEET) || "Weekly-scoreboard",
  MASTER_SHEET:
    (typeof CONFIG !== "undefined" && CONFIG.MASTER_SHEET) || "masterlist",
  COMPLETE_ROUTES_SHEET:
    (typeof CONFIG !== "undefined" && CONFIG.COMPLETE_ROUTES_SHEET) ||
    "Complete route list",
};

function openDriverScorecardDialog(transporterId, driverName) {
  var id = sanitizeTransporterId_(transporterId);
  if (!id) {
    throw new Error("Transporter ID is required for the scorecard.");
  }

  var template = HtmlService.createTemplateFromFile("scorecard_page");
  template.transporterId = id;
  template.driverName = driverName || "";

  var html = template.evaluate().setWidth(1100).setHeight(780);
  SpreadsheetApp.getUi().showModalDialog(html, "Driver scorecard");
  return true;
}

function getDriverScorecardData(transporterId) {
  var id = sanitizeTransporterId_(transporterId);
  if (!id) {
    throw new Error("Transporter ID is required.");
  }

  var ss = SpreadsheetApp.getActive();
  var tz = ss.getSpreadsheetTimeZone() || "Etc/UTC";
  var scoreSheet = ss.getSheetByName(SCORECARD_CONFIG.SCORE_SHEET);
  var masterSheet = ss.getSheetByName(SCORECARD_CONFIG.MASTER_SHEET);
  var completeSheet = ss.getSheetByName(SCORECARD_CONFIG.COMPLETE_ROUTES_SHEET);

  if (!scoreSheet) {
    throw new Error(
      "Scoreboard sheet not found: " + SCORECARD_CONFIG.SCORE_SHEET
    );
  }
  if (!masterSheet) {
    throw new Error(
      "Masterlist sheet not found: " + SCORECARD_CONFIG.MASTER_SHEET
    );
  }

  var masterInfo = loadScorecardMaster_(masterSheet, id);
  var weeklyInfo = collectWeeklyScorecardStats_(scoreSheet, id, tz);
  var routeInfo = collectRouteScorecardStats_(
    completeSheet,
    id,
    buildDriverNameCandidates_(masterInfo, weeklyInfo),
    masterInfo.nameToIds,
    tz
  );
  var leaderboardInfo = getLeaderboardContextForScorecard_(id);
  var metricRanks = computeMetricRanks_(leaderboardInfo.rows, id);
  var comparisons = buildComparisonStats_(
    leaderboardInfo,
    metricRanks,
    weeklyInfo,
    routeInfo
  );
  var spotlight = buildSpotlightHighlights_(
    metricRanks,
    routeInfo.topRoutes,
    leaderboardInfo.row
  );

  var driverName =
    (masterInfo.driver && masterInfo.driver.name) ||
    weeklyInfo.lastDriverName ||
    "Driver " + id;
  var status =
    (masterInfo.driver && masterInfo.driver.status) ||
    weeklyInfo.lastStatus ||
    "Status unknown";

  return {
    driver: {
      name: driverName,
      transporterId: id,
      status: status,
      dspList: routeInfo.summary.dspList,
      weeks: weeklyInfo.weeksCount,
      rank: leaderboardInfo.row ? leaderboardInfo.row.rank : null,
      score:
        leaderboardInfo.row && leaderboardInfo.row.score != null
          ? leaderboardInfo.row.score
          : null,
      teamStanding: comparisons.teamStanding,
      totalDrivers: leaderboardInfo.totalDrivers || null,
      lastWeekLabel: weeklyInfo.lastWeekLabel || "",
      lastWeekNumber: weeklyInfo.lastWeekNumber || null,
      summaryNote: weeklyInfo.summaryNote || "",
    },
    metrics: {
      deliveries: weeklyInfo.totalDeliveries,
      avgDcr: weeklyInfo.avgDcr,
      avgPod: weeklyInfo.avgPod,
      avgCc: weeklyInfo.avgCc,
    },
    metricRanks: metricRanks,
    metricTrends: weeklyInfo.metricTrends,
    charts: {
      dailyDeliveries: routeInfo.dailyBuckets,
      dailyMax: routeInfo.dailyMax,
      rescuesTrend: routeInfo.weeklyTrend,
      weeklyScores: weeklyInfo.scoreTimeline,
    },
    rescues: {
      totalGiven: routeInfo.summary.totalGiven,
      totalTaken: routeInfo.summary.totalTaken,
      avgGivenPerWeek: routeInfo.summary.avgGivenPerWeek,
      avgTakenPerWeek: routeInfo.summary.avgTakenPerWeek,
    },
    comparisons: comparisons,
    additional: {
      avgStops: routeInfo.summary.avgStops,
      totalWeeks: weeklyInfo.weeksCount,
      totalRoutes: routeInfo.summary.totalRoutes,
    },
    routes: {
      top: routeInfo.topRoutes,
    },
    spotlight: spotlight,
    history: {
      weekly: weeklyInfo.weeklyRows,
    },
  };
}

function loadScorecardMaster_(sheet, transporterId) {
  var last = sheet.getLastRow();
  var values = last > 1 ? sheet.getRange(2, 1, last - 1, 6).getValues() : [];
  var driver = null;
  var nameToIds = {};

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var id = String(row[0] || "").trim();
    var first = String(row[1] || "").trim();
    var lastName = String(row[2] || "").trim();
    var status = String(row[5] || "").trim();
    var fullName = (first + " " + lastName).trim();
    var key = fullName.toLowerCase();

    if (fullName) {
      if (!nameToIds[key]) nameToIds[key] = [];
      if (id && nameToIds[key].indexOf(id) === -1) {
        nameToIds[key].push(id);
      }
    }

    if (id && id === transporterId) {
      driver = {
        id: id,
        name: fullName || "Driver " + id,
        status: status || "",
      };
    }
  }

  return { driver: driver, nameToIds: nameToIds };
}

function collectWeeklyScorecardStats_(sheet, transporterId, tz) {
  var last = sheet.getLastRow();
  var values = last > 1 ? sheet.getRange(2, 1, last - 1, 11).getValues() : [];
  var records = [];
  var weeksSet = {};
  var totalDeliveries = 0;
  var dcrSum = 0;
  var dcrCount = 0;
  var podSum = 0;
  var podCount = 0;
  var ccSum = 0;
  var ccCount = 0;
  var lastWeDate = null;
  var lastWeekNumber = null;
  var lastDriverName = "";
  var lastStatus = "";
  var nameVariants = {};

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var id = String(row[5] || "").trim();
    if (id !== transporterId) continue;

    var weekNumber = row[0] != null ? parseInt(row[0], 10) : null;
    var weDate = parseDateCell_(row[1]);
    if (weDate) weDate.setHours(0, 0, 0, 0);

    var delivered = Number(row[6]) || 0;
    var dcr = parsePercent_(row[7]);
    var pod = parsePercent_(row[9]);
    var cc = parsePercent_(row[10]);
    var driverName = String(row[2] || "").trim();
    var status = String(row[3] || "").trim();

    totalDeliveries += delivered;
    if (!isNaN(weekNumber)) {
      weeksSet[String(weekNumber)] = true;
    }

    if (dcr != null) {
      dcrSum += dcr;
      dcrCount++;
    }
    if (pod != null) {
      podSum += pod;
      podCount++;
    }
    if (cc != null) {
      ccSum += cc;
      ccCount++;
    }

    var entry = {
      week: !isNaN(weekNumber) ? weekNumber : null,
      weDateObj: weDate ? new Date(weDate.getTime()) : null,
      dateKey: weDate
        ? Utilities.formatDate(weDate, tz, "yyyy-MM-dd")
        : "",
      delivered: delivered,
      dcrValue: dcr,
      podValue: pod,
      ccValue: cc,
    };
    records.push(entry);

    if (!lastWeDate || (weDate && weDate.getTime() > lastWeDate.getTime())) {
      lastWeDate = weDate ? new Date(weDate.getTime()) : null;
      lastWeekNumber = !isNaN(weekNumber) ? weekNumber : null;
      if (driverName) lastDriverName = driverName;
      if (status) lastStatus = status;
    }

    if (driverName) {
      var key = driverName.toLowerCase();
      if (!nameVariants[key]) {
        nameVariants[key] = driverName;
      }
    }
  }

  records.sort(function (a, b) {
    var at = a.weDateObj ? a.weDateObj.getTime() : 0;
    var bt = b.weDateObj ? b.weDateObj.getTime() : 0;
    return at - bt;
  });

  var weeklyRows = [];
  var prevRecord = null;
  for (var j = 0; j < records.length; j++) {
    var current = records[j];
    var displayDate = current.weDateObj
      ? Utilities.formatDate(current.weDateObj, tz, "d MMM yyyy")
      : "";
    var dcrDisplay = current.dcrValue != null ? roundPct_(current.dcrValue) : null;
    var podDisplay = current.podValue != null ? roundPct_(current.podValue) : null;
    var ccDisplay = current.ccValue != null ? roundPct_(current.ccValue) : null;
    var entryRow = {
      week: current.week,
      weDate: displayDate,
      dateKey: current.dateKey,
      delivered: current.delivered,
      dcr: dcrDisplay,
      pod: podDisplay,
      cc: ccDisplay,
      deliveredChange: prevRecord ? current.delivered - prevRecord.delivered : null,
      deliveredChangePct:
        prevRecord && prevRecord.delivered
          ? ((current.delivered - prevRecord.delivered) / prevRecord.delivered) * 100
          : null,
      dcrChange: prevRecord && dcrDisplay != null && prevRecord.dcr != null
        ? roundToOne_(dcrDisplay - prevRecord.dcr)
        : null,
      podChange: prevRecord && podDisplay != null && prevRecord.pod != null
        ? roundToOne_(podDisplay - prevRecord.pod)
        : null,
      ccChange: prevRecord && ccDisplay != null && prevRecord.cc != null
        ? roundToOne_(ccDisplay - prevRecord.cc)
        : null,
    };
    weeklyRows.push(entryRow);
    prevRecord = {
      delivered: current.delivered,
      dcr: dcrDisplay,
      pod: podDisplay,
      cc: ccDisplay,
    };
  }

  var scoreTimeline = buildWeeklyScoreSeries_(records, tz);
  if (scoreTimeline.length) {
    var timelineMap = {};
    for (var s = 0; s < scoreTimeline.length; s++) {
      var point = scoreTimeline[s];
      var key = point.dateKey || (point.week != null ? "w" + point.week : String(s));
      timelineMap[key] = point;
    }
    for (var h = 0; h < weeklyRows.length; h++) {
      var row = weeklyRows[h];
      var match =
        timelineMap[row.dateKey] ||
        timelineMap["w" + row.week] ||
        null;
      if (match) {
        row.score = match.score;
        row.scoreChange = match.delta;
      }
    }
  }

  var avgDcr = dcrCount ? roundPct_(dcrSum / dcrCount) : null;
  var avgPod = podCount ? roundPct_(podSum / podCount) : null;
  var avgCc = ccCount ? roundPct_(ccSum / ccCount) : null;
  var weeksCount = Object.keys(weeksSet).length;

  var summaryNote;
  if (lastWeDate) {
    var label = Utilities.formatDate(lastWeDate, tz, "d MMM yyyy");
    if (lastWeekNumber != null) {
      summaryNote = "Latest week " + lastWeekNumber + " (" + label + ")";
    } else {
      summaryNote = "Latest week ending " + label;
    }
  } else {
    summaryNote =
      "No weekly scorecard rows have been imported for this driver yet.";
  }

  return {
    weeklyRows: weeklyRows,
    totalDeliveries: totalDeliveries,
    weeksCount: weeksCount,
    avgDcr: avgDcr,
    avgPod: avgPod,
    avgCc: avgCc,
    lastWeekLabel: lastWeDate
      ? Utilities.formatDate(lastWeDate, tz, "d MMM yyyy")
      : "",
    lastWeekNumber: lastWeekNumber,
    lastDriverName: lastDriverName,
    lastStatus: lastStatus,
    summaryNote: summaryNote,
    metricTrends: buildMetricTrendSummary_(weeklyRows),
    scoreTimeline: scoreTimeline,
    nameVariants: Object.keys(nameVariants).map(function (k) {
      return nameVariants[k];
    }),
  };
}

function buildWeeklyScoreSeries_(records, tz) {
  records = records || [];
  if (!records.length) return [];
  if (typeof computeScore_ !== "function") return [];

  var cfg =
    typeof getLeaderboardConfig_ === "function"
      ? getLeaderboardConfig_()
      : typeof getLeaderboardConfig === "function"
      ? getLeaderboardConfig()
      : null;
  if (!cfg) {
    cfg = {
      dcrWeight: 0.4,
      podWeight: 0.4,
      ccWeight: 0.2,
      minWeeks: 0,
      volumeWeight: 0.5,
      weeksWeight: 0.3,
      rescuesGivenWeight: 1.0,
      rescuesTakenWeight: 1.0,
    };
  } else {
    cfg = JSON.parse(JSON.stringify(cfg));
    cfg.minWeeks = 0;
  }

  var timeline = [];
  var running = {
    deliveries: 0,
    weeks: 0,
    dcrSum: 0,
    dcrCount: 0,
    podSum: 0,
    podCount: 0,
    ccSum: 0,
    ccCount: 0,
  };

  for (var i = 0; i < records.length; i++) {
    var rec = records[i];
    running.weeks++;
    running.deliveries += rec.delivered || 0;
    if (rec.dcrValue != null) {
      running.dcrSum += rec.dcrValue;
      running.dcrCount++;
    }
    if (rec.podValue != null) {
      running.podSum += rec.podValue;
      running.podCount++;
    }
    if (rec.ccValue != null) {
      running.ccSum += rec.ccValue;
      running.ccCount++;
    }

    var stats = {
      deliveries: running.deliveries,
      weeks: running.weeks,
      dcr: running.dcrCount ? running.dcrSum / running.dcrCount : null,
      pod: running.podCount ? running.podSum / running.podCount : null,
      cc: running.ccCount ? running.ccSum / running.ccCount : null,
      rescuesGiven: 0,
      rescuesTaken: 0,
    };
    var score = computeScore_(stats, cfg);
    timeline.push({
      week: rec.week,
      label: rec.weDateObj
        ? Utilities.formatDate(rec.weDateObj, tz, "dd/MM/yyyy")
        : rec.week != null
        ? "Week " + rec.week
        : "Week",
      score: score,
      dateKey: rec.dateKey || "",
    });
  }

  for (var j = 0; j < timeline.length; j++) {
    var prev = j > 0 ? timeline[j - 1] : null;
    timeline[j].delta =
      prev && prev.score != null && timeline[j].score != null
        ? roundToOne_(timeline[j].score - prev.score)
        : null;
  }

  return timeline;
}

function collectRouteScorecardStats_(sheet, transporterId, driverNameCandidates, nameToIds, tz) {
  driverNameCandidates = driverNameCandidates || [];
  var driverNameSet = {};
  for (var n = 0; n < driverNameCandidates.length; n++) {
    var nm = String(driverNameCandidates[n] || "").trim();
    if (nm) driverNameSet[nm.toLowerCase()] = true;
  }
  var dailyTotals = {};
  var driverRouteRecords = [];
  var dateKeySet = {};
  var weeklyMap = {};
  var routeTotals = {};
  var summary = {
    totalRoutes: 0,
    onTimeRoutes: 0,
    totalStops: 0,
    stopsComplete: 0,
    routesThisWeek: 0,
    totalGiven: 0,
    totalTaken: 0,
    dspSet: {},
  };

  if (sheet) {
    var last = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (last > 1) {
      var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
      var columnMap = buildRouteHeaderMap_(headerRow);
      var values = sheet.getRange(2, 1, last - 1, lastCol).getValues();
      for (var i = 0; i < values.length; i++) {
        var row = values[i];
        var dateIndex = columnMap.date != null ? columnMap.date : 1;
        var date = parseDateCell_(row[dateIndex]);
        if (!date) continue;
        date.setHours(0, 0, 0, 0);

        var transporterCell =
          columnMap.transporterId != null
            ? String(row[columnMap.transporterId] || "").trim()
            : "";
        var dsp =
          columnMap.dsp != null ? String(row[columnMap.dsp] || "").trim() : "";
        var progress =
          columnMap.progress != null
            ? String(row[columnMap.progress] || "")
            : "";
        var allStops =
          columnMap.allStops != null ? Number(row[columnMap.allStops] || 0) : 0;
        var stopsComplete =
          columnMap.stopsComplete != null
            ? Number(row[columnMap.stopsComplete] || 0)
            : 0;
        var rescuedBy =
          columnMap.rescuedBy != null
            ? String(row[columnMap.rescuedBy] || "").trim()
            : "";
        var driverCell =
          columnMap.driverName != null
            ? String(row[columnMap.driverName] || "").trim()
            : "";
        var routeCode =
          columnMap.routeCode != null
            ? String(row[columnMap.routeCode] || "").trim()
            : "";

        var dateKey = Utilities.formatDate(date, tz, "yyyy-MM-dd");
        var driverMatch =
          driverCell && driverNameSet[driverCell.toLowerCase()] ? true : false;
        var isDriverRow =
          (transporterCell && transporterCell === transporterId) || driverMatch;

        if (isDriverRow) {
          summary.totalRoutes++;
          if (dsp) summary.dspSet[dsp] = true;
          summary.totalStops += allStops;
          summary.stopsComplete += stopsComplete;

          var deliveredCount = stopsComplete > 0 ? stopsComplete : allStops;
          dailyTotals[keySafe_(dateKey)] =
            (dailyTotals[keySafe_(dateKey)] || 0) + (deliveredCount || 0);
          dateKeySet[dateKey] = true;
          if (routeCode) {
            routeTotals[routeCode] =
              (routeTotals[routeCode] || 0) + (deliveredCount || 0);
          }

          var success = /on[-_ ]?time|complete/i.test(progress || "");
          driverRouteRecords.push({
            date: new Date(date.getTime()),
            key: dateKey,
            onTime: success,
          });
          if (success) {
            summary.onTimeRoutes++;
          }

          if (rescuedBy) {
            summary.totalTaken++;
            incrementRescueWeek_(weeklyMap, date, tz, "taken");
          }
        }

        if (rescuedBy) {
          var ids = lookupIdsForNames_(rescuedBy, nameToIds || {});
          if (ids.indexOf(transporterId) !== -1) {
            summary.totalGiven++;
            incrementRescueWeek_(weeklyMap, date, tz, "given");
          }
        }
      }
    }
  }

  driverRouteRecords.sort(function (a, b) {
    return a.date.getTime() - b.date.getTime();
  });

  var baseDate = getTodayForTimezone_(tz);
  var bucketInfo = buildDailyBuckets_(tz, baseDate);
  var buckets = bucketInfo.buckets;
  var bucketKeyMap = bucketInfo.keyMap;
  var maxDailyValue = 0;

  for (var b = 0; b < buckets.length; b++) {
    var bucket = buckets[b];
    var key = keySafe_(bucket.key);
    var val = dailyTotals[key] || 0;
    bucket.value = val;
    if (val > maxDailyValue) maxDailyValue = val;
  }

  var routesThisWeek = 0;
  for (var r = 0; r < driverRouteRecords.length; r++) {
    if (bucketKeyMap[keySafe_(driverRouteRecords[r].key)]) {
      routesThisWeek++;
    }
  }
  summary.routesThisWeek = routesThisWeek;

  var weeklyTrend = [];
  for (var wk in weeklyMap) {
    if (weeklyMap.hasOwnProperty(wk)) {
      weeklyTrend.push(weeklyMap[wk]);
    }
  }
  weeklyTrend.sort(function (a, b) {
    return a.order - b.order;
  });

  var topRoutes = Object.keys(routeTotals)
    .map(function (code) {
      return { route: code, deliveries: routeTotals[code] };
    })
    .sort(function (a, b) {
      return b.deliveries - a.deliveries;
    })
    .slice(0, 3);

  var weeksWithRescues = weeklyTrend.length;
  var avgGivenPerWeek =
    weeksWithRescues > 0
      ? roundToOne_(summary.totalGiven / weeksWithRescues)
      : summary.totalGiven
      ? roundToOne_(summary.totalGiven)
      : 0;
  var avgTakenPerWeek =
    weeksWithRescues > 0
      ? roundToOne_(summary.totalTaken / weeksWithRescues)
      : summary.totalTaken
      ? roundToOne_(summary.totalTaken)
      : 0;

  var resultSummary = {
    totalRoutes: summary.totalRoutes,
    routesThisWeek: summary.routesThisWeek,
    totalGiven: summary.totalGiven,
    totalTaken: summary.totalTaken,
    avgGivenPerWeek: avgGivenPerWeek,
    avgTakenPerWeek: avgTakenPerWeek,
    onTimeRate:
      summary.totalRoutes
        ? roundToOne_((summary.onTimeRoutes / summary.totalRoutes) * 100)
        : null,
    avgStops:
      summary.totalRoutes
        ? roundToOne_(summary.totalStops / summary.totalRoutes)
        : null,
    completionRate:
      summary.totalStops
        ? roundToOne_((summary.stopsComplete / summary.totalStops) * 100)
        : null,
    dspList: Object.keys(summary.dspSet).sort(),
  };

  return {
    dailyBuckets: buckets,
    dailyMax: maxDailyValue,
    weeklyTrend: weeklyTrend,
    summary: resultSummary,
    topRoutes: topRoutes,
  };
}

function buildDailyBuckets_(tz, baseDate) {
  var anchor = baseDate ? new Date(baseDate.getTime()) : new Date();
  anchor.setHours(0, 0, 0, 0);
  var buckets = [];
  var map = {};

  for (var offset = 6; offset >= 0; offset--) {
    var d = new Date(anchor.getTime());
    d.setDate(anchor.getDate() - offset);
    var key = Utilities.formatDate(d, tz, "yyyy-MM-dd");
    var label = Utilities.formatDate(d, tz, "EEE");
    var bucket = { key: key, label: label, value: 0 };
    buckets.push(bucket);
    map[keySafe_(key)] = true;
  }

  return { buckets: buckets, keyMap: map };
}

function parseDateCell_(value) {
  if (value === null || typeof value === "undefined") return null;
  if (Object.prototype.toString.call(value) === "[object Date]") {
    var copy = new Date(value.getTime());
    return isNaN(copy.getTime()) ? null : copy;
  }
  if (typeof value === "number") {
    var excelEpoch = new Date(Date.UTC(1899, 11, 30));
    var ms = value * 24 * 60 * 60 * 1000;
    var numericDate = new Date(excelEpoch.getTime() + ms);
    if (!isNaN(numericDate.getTime())) {
      return numericDate;
    }
  }
  var str = String(value).trim();
  if (!str) return null;

  var match = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (match) {
    var day = parseInt(match[1], 10);
    var month = parseInt(match[2], 10) - 1;
    var year = parseInt(match[3], 10);
    if (isNaN(year)) {
      year = new Date().getFullYear();
    } else if (year < 100) {
      year += year > 70 ? 1900 : 2000;
    }
    var date = new Date(year, month, day);
    if (!isNaN(date.getTime())) {
      return date;
    }
  }

  var parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return parsed;
  }
  return null;
}

function sanitizeTransporterId_(value) {
  if (value === null || typeof value === "undefined") return "";
  return String(value).trim();
}

function lookupIdsForNames_(raw, nameToIds) {
  var ids = [];
  if (!raw) return ids;
  var parts = raw.split(/[,|]/);
  for (var i = 0; i < parts.length; i++) {
    var token = String(parts[i] || "").trim();
    if (!token) continue;
    var normalized = token.toLowerCase();
    var matches = nameToIds[normalized];
    if (matches && matches.length) {
      for (var j = 0; j < matches.length; j++) {
        var id = matches[j];
        if (ids.indexOf(id) === -1) {
          ids.push(id);
        }
      }
    }
  }
  return ids;
}

function incrementRescueWeek_(map, date, tz, field) {
  if (!date) return;
  var year = parseInt(Utilities.formatDate(date, tz, "yyyy"), 10);
  var weekStr = Utilities.formatDate(date, tz, "w");
  var week = parseInt(weekStr, 10);
  if (!year || !week) return;

  var key = year + "-W" + weekStr;
  if (!map[key]) {
    map[key] = {
      key: key,
      label: "Week " + weekStr,
      given: 0,
      taken: 0,
      order: year * 100 + week,
    };
  }
  map[key][field]++;
}

function getLeaderboardContextForScorecard_(transporterId) {
  try {
    var data = getLeaderboardData();
    var rows = (data && data.rows) || [];
    var match = null;
    var topScore = null;
    var sumScores = 0;
    var countScores = 0;

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (row.score != null) {
        sumScores += row.score;
        countScores++;
        if (topScore === null || row.score > topScore) {
          topScore = row.score;
        }
      }
      if (!match && row.transporterId === transporterId) {
        match = row;
      }
    }

    return {
      row: match,
      rows: rows,
      totalDrivers: data && data.summary ? data.summary.totalDrivers : rows.length,
      topScore: topScore,
      avgScore: countScores ? sumScores / countScores : null,
    };
  } catch (err) {
    return {
      row: null,
      rows: [],
      totalDrivers: null,
      topScore: null,
      avgScore: null,
    };
  }
}

function computeMetricRanks_(rows, transporterId) {
  rows = rows || [];
  var metrics = [
    { key: "score", label: "score" },
    { key: "deliveries", label: "deliveries" },
    { key: "dcr", label: "dcr" },
    { key: "pod", label: "pod" },
    { key: "cc", label: "cc" },
  ];
  var result = {};

  for (var i = 0; i < metrics.length; i++) {
    var metric = metrics[i];
    var list = [];
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      if (row[metric.key] != null && row[metric.key] !== "") {
        list.push({
          id: row.transporterId,
          value: row[metric.key],
        });
      }
    }
    list.sort(function (a, b) {
      return b.value - a.value;
    });
    var rank = null;
    for (var j = 0; j < list.length; j++) {
      if (list[j].id === transporterId) {
        rank = j + 1;
        break;
      }
    }
    result[metric.key] = { rank: rank, total: list.length };
  }

  return result;
}

function buildComparisonStats_(ctx, metricRanks, weeklyInfo, routeInfo) {
  var row = ctx && ctx.row ? ctx.row : null;
  var averages = computeTeamAverages_(ctx && ctx.rows ? ctx.rows : []);
  var driverQuality = row ? calculateQualityScore_(row) : null;
  var rescueBalance = row
    ? (row.rescuesGiven || 0) - (row.rescuesTaken || 0)
    : null;
  var deliveriesPerWeek =
    row && row.weeks ? row.deliveries / row.weeks : null;
  var percentile =
    row && ctx.totalDrivers
      ? Math.max(
          1,
          Math.round(((ctx.totalDrivers - row.rank + 1) / ctx.totalDrivers) * 100)
        )
      : null;

  return {
    vsTeamAvg:
      row && ctx.avgScore != null
        ? roundToOne_(row.score - ctx.avgScore)
        : null,
    vsTop:
      row && ctx.topScore != null
        ? roundToOne_(row.score - ctx.topScore)
        : null,
    teamStanding: ctx && ctx.row && ctx.totalDrivers
      ? ctx.row.rank + " of " + ctx.totalDrivers
      : ctx && ctx.totalDrivers
      ? "Unranked of " + ctx.totalDrivers
      : "",
    scoreDiff:
      row && ctx.avgScore != null
        ? roundToOne_(row.score - ctx.avgScore)
        : null,
    deliveriesDiff: diffNumbers_(row ? row.deliveries : null, averages.deliveries),
    qualityDiff: diffNumbers_(driverQuality, averages.quality),
    rescueDiff: diffNumbers_(rescueBalance, averages.rescueBalance),
    deliveriesPerWeekDiff: diffNumbers_(deliveriesPerWeek, averages.deliveriesPerWeek),
    weeksDiff: diffNumbers_(row ? row.weeks : null, averages.weeks),
    scoreVsTop:
      row && ctx.topScore != null && row.score != null
        ? roundToOne_(ctx.topScore - row.score)
        : null,
    averages: averages,
    totalDrivers: ctx.totalDrivers || (ctx.rows ? ctx.rows.length : null),
    metricRanks: metricRanks,
    percentile: percentile,
  };
}

function buildSpotlightHighlights_(metricRanks, topRoutes, leaderRow) {
  var highlights = [];
  if (!metricRanks) metricRanks = {};
  topRoutes = topRoutes || [];

  var metricNames = {
    dcr: "Avg DCR",
    pod: "Avg POD",
    cc: "Avg CC",
    deliveries: "Deliveries",
    score: "Overall score",
  };

  for (var key in metricNames) {
    if (!metricNames.hasOwnProperty(key)) continue;
    var info = metricRanks[key];
    if (info && info.rank && info.rank <= 5) {
      var descriptor = info.rank <= 3 ? "Top " + info.rank : "Top 5";
      highlights.push({
        title: "🏅 " + descriptor + " in " + metricNames[key],
        detail: "Ranked " + info.rank + " of " + (info.total || "-"),
      });
    }
  }

  if (leaderRow && leaderRow.rescuesGiven) {
    highlights.push({
      title: "🤝 Rescue hero",
      detail: (leaderRow.rescuesGiven || 0) + " rescues given",
    });
  }
  if (leaderRow && leaderRow.rescuesTaken) {
    highlights.push({
      title: "🆘 Support ready",
      detail: (leaderRow.rescuesTaken || 0) + " rescues received",
    });
  }

  if (topRoutes.length) {
    var topRoute = topRoutes[0];
    highlights.push({
      title: "🚚 Route " + (topRoute.route || ""),
      detail: (topRoute.deliveries || 0) + " deliveries handled",
    });
  }
  if (!highlights.length) {
    highlights.push({
      title: "Building history",
      detail: "Keep logging weeks to unlock achievements.",
    });
  }

  return highlights.slice(0, 6);
}

function computeTeamAverages_(rows) {
  var totals = {
    deliveries: 0,
    deliveriesCount: 0,
    dcr: 0,
    dcrCount: 0,
    pod: 0,
    podCount: 0,
    cc: 0,
    ccCount: 0,
    qualitySum: 0,
    qualityCount: 0,
    rescueSum: 0,
    rescueCount: 0,
    weeks: 0,
    weeksCount: 0,
    deliveriesPerWeek: 0,
    deliveriesPerWeekCount: 0,
    scoreSum: 0,
    scoreCount: 0,
  };

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (row.deliveries != null) {
      totals.deliveries += row.deliveries;
      totals.deliveriesCount++;
    }
    if (row.dcr != null) {
      totals.dcr += row.dcr;
      totals.dcrCount++;
    }
    if (row.pod != null) {
      totals.pod += row.pod;
      totals.podCount++;
    }
    if (row.cc != null) {
      totals.cc += row.cc;
      totals.ccCount++;
    }
    if (row.weeks != null) {
      totals.weeks += row.weeks;
      totals.weeksCount++;
    }
    if (row.deliveries != null && row.weeks) {
      totals.deliveriesPerWeek += row.deliveries / row.weeks;
      totals.deliveriesPerWeekCount++;
    }
    if (row.score != null) {
      totals.scoreSum += row.score;
      totals.scoreCount++;
    }
    var quality = calculateQualityScore_(row);
    if (quality != null) {
      totals.qualitySum += quality;
      totals.qualityCount++;
    }
    var balance =
      row.rescuesGiven != null || row.rescuesTaken != null
        ? (row.rescuesGiven || 0) - (row.rescuesTaken || 0)
        : null;
    if (balance != null) {
      totals.rescueSum += balance;
      totals.rescueCount++;
    }
  }

  return {
    deliveries:
      totals.deliveriesCount ? totals.deliveries / totals.deliveriesCount : null,
    dcr: totals.dcrCount ? totals.dcr / totals.dcrCount : null,
    pod: totals.podCount ? totals.pod / totals.podCount : null,
    cc: totals.ccCount ? totals.cc / totals.ccCount : null,
    quality: totals.qualityCount ? totals.qualitySum / totals.qualityCount : null,
    rescueBalance:
      totals.rescueCount ? totals.rescueSum / totals.rescueCount : null,
    weeks: totals.weeksCount ? totals.weeks / totals.weeksCount : null,
    deliveriesPerWeek:
      totals.deliveriesPerWeekCount
        ? totals.deliveriesPerWeek / totals.deliveriesPerWeekCount
        : null,
    score: totals.scoreCount ? totals.scoreSum / totals.scoreCount : null,
  };
}

function calculateQualityScore_(row) {
  if (!row) return null;
  var values = [];
  if (row.dcr != null) values.push(row.dcr);
  if (row.pod != null) values.push(row.pod);
  if (row.cc != null) values.push(row.cc);
  if (!values.length) return null;
  var sum = 0;
  for (var i = 0; i < values.length; i++) {
    sum += values[i];
  }
  return sum / values.length;
}

function buildDriverNameCandidates_(masterInfo, weeklyInfo) {
  var list = [];
  if (masterInfo && masterInfo.driver && masterInfo.driver.name) {
    list.push(masterInfo.driver.name);
  }
  if (weeklyInfo && weeklyInfo.nameVariants && weeklyInfo.nameVariants.length) {
    list = list.concat(weeklyInfo.nameVariants);
  }
  var seen = {};
  var result = [];
  for (var i = 0; i < list.length; i++) {
    var name = String(list[i] || "").trim();
    if (!name) continue;
    var key = name.toLowerCase();
    if (seen[key]) continue;
    seen[key] = true;
    result.push(name);
  }
  return result;
}

function buildRouteHeaderMap_(headerRow) {
  var map = {};
  for (var i = 0; i < headerRow.length; i++) {
    var label = String(headerRow[i] || "").trim().toLowerCase();
    if (!label) continue;
    if (!map.date && label === "date") map.date = i;
    else if (!map.transporterId && label.indexOf("transporter") !== -1)
      map.transporterId = i;
    else if (!map.driverName && label.indexOf("driver") !== -1)
      map.driverName = i;
    else if (!map.dsp && (label === "dsp" || label.indexOf("dsp") !== -1))
      map.dsp = i;
    else if (!map.routeCode && label.indexOf("route code") !== -1)
      map.routeCode = i;
    else if (!map.progress && label.indexOf("progress") !== -1)
      map.progress = i;
    else if (!map.allStops && label.indexOf("all stops") !== -1)
      map.allStops = i;
    else if (!map.stopsComplete && label.indexOf("stops complete") !== -1)
      map.stopsComplete = i;
    else if (!map.rescuedBy && label.indexOf("rescued by") !== -1)
      map.rescuedBy = i;
  }
  if (typeof map.date === "undefined") map.date = 1;
  if (typeof map.transporterId === "undefined") map.transporterId = 4;
  if (typeof map.driverName === "undefined") map.driverName = 5;
  if (typeof map.routeCode === "undefined") map.routeCode = 3;
  if (typeof map.progress === "undefined") map.progress = 6;
  if (typeof map.dsp === "undefined") map.dsp = 3;
  if (typeof map.allStops === "undefined") map.allStops = 9;
  if (typeof map.stopsComplete === "undefined") map.stopsComplete = 10;
  if (typeof map.rescuedBy === "undefined") map.rescuedBy = 12;
  return map;
}

function getTodayForTimezone_(tz) {
  var todayString = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  var parts = todayString.split("-");
  return new Date(
    parseInt(parts[0], 10),
    parseInt(parts[1], 10) - 1,
    parseInt(parts[2], 10)
  );
}

function buildMetricTrendSummary_(weeklyRows) {
  var last = weeklyRows.length ? weeklyRows[weeklyRows.length - 1] : null;
  var prev = weeklyRows.length > 1 ? weeklyRows[weeklyRows.length - 2] : null;
  return {
    deliveries: buildTrendData_(last ? last.delivered : null, prev ? prev.delivered : null),
    dcr: buildTrendData_(last ? last.dcr : null, prev ? prev.dcr : null),
    pod: buildTrendData_(last ? last.pod : null, prev ? prev.pod : null),
    cc: buildTrendData_(last ? last.cc : null, prev ? prev.cc : null),
  };
}

function buildTrendData_(current, previous) {
  if (current == null || previous == null) {
    return { delta: null, percent: null };
  }
  var delta = current - previous;
  var percent = previous !== 0 ? (delta / previous) * 100 : null;
  return {
    delta: roundToOne_(delta),
    percent: percent != null ? roundToOne_(percent) : null,
  };
}

function parseDateKey_(key) {
  var parts = String(key || "").split("-");
  if (parts.length !== 3) return new Date();
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

function keySafe_(key) {
  return String(key || "");
}

function diffNumbers_(current, baseline) {
  if (current == null || baseline == null) return null;
  return roundToOne_(current - baseline);
}

function roundToOne_(value) {
  if (value === null || typeof value === "undefined" || isNaN(value)) return null;
  return Math.round(value * 10) / 10;
}
