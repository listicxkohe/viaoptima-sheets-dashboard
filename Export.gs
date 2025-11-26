/************************************************
 * PDF EXPORT (SCREENSHOT-STYLE)
 * Builds a quick PDF with inline charts for sharing.
 ************************************************/

function downloadScorecardPdf(transporterId) {
  var id = sanitizeTransporterId_(transporterId);
  if (!id) {
    throw new Error("Transporter ID is required to export a scorecard.");
  }

  var data = getDriverScorecardData(id);
  if (!data || !data.driver) {
    throw new Error("Unable to load scorecard data for " + id + ".");
  }

  var doc = DocumentApp.create("scorecard_export_" + id + "_" + Date.now());
  var body = doc.getBody();
  body.clear();
  body.setMarginTop(18).setMarginBottom(18).setMarginLeft(24).setMarginRight(24);

  appendHeader_(body, data);
  appendMetricCards_(body, data);
  appendComparisonCards_(body, data);
  appendCharts_(body, data);
  appendRescues_(body, data);
  appendRoutes_(body, data);
  appendHistory_(body, data);
  appendSpotlight_(body, data);

  doc.saveAndClose();

  var blob = doc.getAs(MimeType.PDF);
  var fileName =
    (data.driver.name || "Driver " + id).replace(/[\\/:*?"<>|]/g, "") +
    " - scorecard.pdf";
  blob.setName(fileName);

  DriveApp.getFileById(doc.getId()).setTrashed(true);

  return {
    fileName: fileName,
    base64: Utilities.base64Encode(blob.getBytes()),
    mimeType: blob.getContentType(),
  };
}

function appendHeader_(body, data) {
  var driver = data.driver || {};
  var table = body.appendTable();
  table.setBorderWidth(0);
  var row = table.appendTableRow();
  var left = row.appendTableCell();
  left.setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(6).setPaddingRight(6);
  left.appendParagraph("Driver scorecard").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  left.appendParagraph(driver.name || "Driver").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  left.appendParagraph((driver.transporterId || "N/A") + " | " + (driver.status || "Unknown"));
  if (driver.dspList && driver.dspList.length) {
    left.appendParagraph("DSP: " + driver.dspList.join(", "));
  }
  left.appendParagraph(
    "Weeks tracked: " + (driver.weeks != null ? driver.weeks : "N/A")
  );
  var right = row.appendTableCell();
  right.setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(6).setPaddingRight(6);
  var scorePara = right.appendParagraph("Score");
  scorePara.setBold(true).setForegroundColor("#6b7280");
  var scoreVal = right.appendParagraph(driver.score != null ? driver.score.toFixed(1) : "N/A");
  scoreVal.setBold(true).setFontSize(26);
  right.appendParagraph(driver.teamStanding || "").setForegroundColor("#6b7280");
  right.setBackgroundColor("#f8fafc");
}

function appendMetricCards_(body, data) {
  var metrics = data.metrics || {};
  var ranks = data.metricRanks || {};
  var trends = data.metricTrends || {};
  var cardItems = [
    {
      label: "Total deliveries",
      value: formatScorecardNumber_(metrics.deliveries),
      note: formatScorecardRank_(ranks.deliveries),
      trend: buildTrendText_(trends.deliveries, true),
    },
    {
      label: "Avg DCR",
      value: formatScorecardPercent_(metrics.avgDcr),
      note: formatScorecardRank_(ranks.dcr),
      trend: buildTrendText_(trends.dcr, false),
    },
    {
      label: "Avg POD",
      value: formatScorecardPercent_(metrics.avgPod),
      note: formatScorecardRank_(ranks.pod),
      trend: buildTrendText_(trends.pod, false),
    },
    {
      label: "Avg CC",
      value: formatScorecardPercent_(metrics.avgCc),
      note: formatScorecardRank_(ranks.cc),
      trend: buildTrendText_(trends.cc, false),
    },
  ];
  appendCardGrid_(body, cardItems, 2);
}

function appendComparisonCards_(body, data) {
  var comps = data.comparisons || {};
  var items = [
    { label: "Score vs team", value: formatScorecardDelta_(comps.scoreDiff, "pts"), note: comps.teamStanding || "" },
    { label: "Gap to rank #1", value: formatScorecardDelta_(-comps.scoreVsTop, "pts"), note: "Points away from leader" },
    { label: "Deliveries vs avg", value: formatScorecardDelta_(comps.deliveriesDiff, ""), note: buildAvgText_(comps.averages && comps.averages.deliveries, "deliveries") },
    { label: "Quality vs avg", value: formatScorecardDelta_(comps.qualityDiff, "pts"), note: buildAvgText_(comps.averages && comps.averages.quality, "quality") },
    { label: "Rescue balance", value: formatScorecardDelta_(comps.rescueDiff, ""), note: buildAvgText_(comps.averages && comps.averages.rescueBalance, "balance") },
    { label: "Deliveries / week", value: formatScorecardDelta_(comps.deliveriesPerWeekDiff, ""), note: "Vs team average" },
    { label: "Experience", value: formatScorecardDelta_(comps.weeksDiff, "wks"), note: "Weeks vs average" },
    { label: "Percentile", value: comps.percentile != null ? "Top " + comps.percentile + "%" : "N/A", note: "" },
  ];
  appendCardGrid_(body, items, 3);
}

function appendCharts_(body, data) {
  var charts = data.charts || {};
  var rowTable = body.appendTable();
  rowTable.setBorderWidth(0);
  var row = rowTable.appendTableRow();

  var dailyBlob = buildDailyChartImage_(charts.dailyDeliveries || []);
  var dailyCell = row.appendTableCell();
  dailyCell.setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(6).setPaddingRight(6);
  dailyCell.appendParagraph("Daily deliveries (last 7 days)").setBold(true);
  if (dailyBlob) dailyCell.appendImage(dailyBlob).setWidth(320);

  var scoreBlob = buildScoreChartImage_(charts.weeklyScores || []);
  var scoreCell = row.appendTableCell();
  scoreCell.setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(6).setPaddingRight(6);
  scoreCell.appendParagraph("Weekly score progress").setBold(true);
  if (scoreBlob) scoreCell.appendImage(scoreBlob).setWidth(320);
}

function appendRescues_(body, data) {
  var rescues = data.rescues || {};
  body.appendParagraph("Rescues").setHeading(
    DocumentApp.ParagraphHeading.HEADING3
  );
  var items = [
    { label: "Given per week", value: formatScorecardNumber_(rescues.avgGivenPerWeek) },
    { label: "Taken per week", value: formatScorecardNumber_(rescues.avgTakenPerWeek) },
    { label: "Total given", value: formatScorecardNumber_(rescues.totalGiven) },
    { label: "Total taken", value: formatScorecardNumber_(rescues.totalTaken) },
  ];
  appendCardGrid_(body, items, 2);
}

function appendRoutes_(body, data) {
  var additional = data.additional || {};
  var topRoutes = (data.routes && data.routes.top) || [];
  body.appendParagraph("Route insights").setHeading(
    DocumentApp.ParagraphHeading.HEADING3
  );
  var items = [
    { label: "Avg stops per route", value: formatScorecardNumber_(additional.avgStops) },
    { label: "Total routes", value: formatScorecardNumber_(additional.totalRoutes) },
  ];
  appendCardGrid_(body, items, 2);

  if (topRoutes.length) {
    body.appendParagraph("Top routes").setHeading(
      DocumentApp.ParagraphHeading.HEADING4
    );
    var tableRows = [["Route", "Deliveries"]];
    for (var i = 0; i < topRoutes.length; i++) {
      tableRows.push([
        topRoutes[i].route || "(unknown)",
        formatScorecardNumber_(topRoutes[i].deliveries),
      ]);
    }
    var table = body.appendTable(tableRows);
    styleScorecardTableHeader_(table.getRow(0));
  }
}

function appendHistory_(body, data) {
  var weeks = (data.history && data.history.weekly) || [];
  if (!weeks.length) return;
  body.appendParagraph("Recent weekly highlights").setHeading(
    DocumentApp.ParagraphHeading.HEADING3
  );
  var lastWeeks = weeks.slice(-5);
  for (var i = 0; i < lastWeeks.length; i++) {
    var row = lastWeeks[i];
    var labelParts = [];
    if (row.week != null) labelParts.push("Week " + row.week);
    if (row.weDate) labelParts.push(row.weDate);
    var line = labelParts.join(" - ") || "Week detail";
    line += ": " + formatScorecardNumber_(row.delivered) + " deliveries";
    body.appendParagraph("• " + line);
  }
}

function appendSpotlight_(body, data) {
  var spotlight = data.spotlight || [];
  if (!spotlight.length) return;
  body.appendParagraph("Driver spotlight").setHeading(
    DocumentApp.ParagraphHeading.HEADING3
  );
  for (var i = 0; i < spotlight.length; i++) {
    var item = spotlight[i];
    body.appendParagraph("• " + (item.title || "") + " — " + (item.detail || ""));
  }
}

function appendCardGrid_(body, items, columns) {
  columns = columns || 2;
  var rows = Math.ceil(items.length / columns);
  var table = body.appendTable();
  table.setBorderWidth(0);

  var idx = 0;
  for (var r = 0; r < rows; r++) {
    var row = table.appendTableRow();
    for (var c = 0; c < columns; c++) {
      var cell = row.appendTableCell();
      cell.setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(8).setPaddingRight(8);
      if (idx < items.length) {
        var card = items[idx];
        var labelText = cell.appendParagraph(card.label || "");
        labelText.setBold(true).setFontSize(10).setForegroundColor("#6b7280");
        var valueText = cell.appendParagraph(card.value != null ? String(card.value) : "N/A");
        valueText.setFontSize(14).setBold(true);
        if (card.note) {
          var noteText = cell.appendParagraph(card.note);
          noteText.setFontSize(9).setForegroundColor("#6b7280");
        }
        if (card.trend) {
          var trendText = cell.appendParagraph(card.trend);
          trendText.setFontSize(9).setForegroundColor("#16a34a");
        }
        cell.setBackgroundColor("#f8fafc");
      } else {
        cell.setBackgroundColor("#ffffff");
      }
      idx++;
    }
  }
}

function buildDailyChartImage_(points) {
  if (!points || !points.length) return null;
  var dataTable = Charts.newDataTable();
  dataTable.addColumn(Charts.ColumnType.STRING, "Day");
  dataTable.addColumn(Charts.ColumnType.NUMBER, "Deliveries");
  for (var i = 0; i < points.length; i++) {
    dataTable.addRow([points[i].label || "", Number(points[i].value || 0)]);
  }
  var chart = Charts.newColumnChart()
    .setDataTable(dataTable)
    .setDimensions(520, 300)
    .setColors(["#60a5fa"])
    .setLegendPosition(Charts.Position.NONE)
    .setOption("chartArea", { width: "80%", height: "70%" })
    .build();
  return chart.getAs("image/png");
}

function buildScoreChartImage_(points) {
  if (!points || !points.length) return null;
  var dataTable = Charts.newDataTable();
  dataTable.addColumn(Charts.ColumnType.STRING, "Week");
  dataTable.addColumn(Charts.ColumnType.NUMBER, "Score");
  for (var i = 0; i < points.length; i++) {
    if (points[i].score == null) continue;
    dataTable.addRow([points[i].label || "", Number(points[i].score)]);
  }
  var chart = Charts.newLineChart()
    .setDataTable(dataTable)
    .setDimensions(520, 300)
    .setColors(["#a78bfa"])
    .setLegendPosition(Charts.Position.NONE)
    .setOption("chartArea", { width: "80%", height: "70%" })
    .build();
  return chart.getAs("image/png");
}

function buildTrendText_(trend, isPercent) {
  if (!trend || trend.delta == null) return "";
  var prefix = trend.delta > 0 ? "+" : "";
  if (isPercent && trend.percent != null) {
    return prefix + trend.percent.toFixed(1) + "% vs last week";
  }
  return prefix + trend.delta.toFixed(1) + " vs last week";
}

function styleScorecardTableHeader_(row) {
  if (!row) return;
  for (var i = 0; i < row.getNumCells(); i++) {
    row.getCell(i).editAsText().setBold(true);
  }
}

function formatScorecardNumber_(value) {
  if (value === null || typeof value === "undefined") return "N/A";
  return Number(value).toLocaleString();
}

function formatScorecardPercent_(value) {
  if (value === null || typeof value === "undefined") return "N/A";
  return Number(value).toFixed(1) + "%";
}

function formatScorecardRank_(info) {
  if (!info || !info.rank) return "N/A";
  return info.rank <= 3 ? "Top " + info.rank : "Rank " + info.rank;
}

function formatScorecardDelta_(value, suffix) {
  if (value === null || typeof value === "undefined") return "N/A";
  var prefix = value > 0 ? "+" : "";
  return prefix + value.toFixed(1) + (suffix || "");
}

function buildAvgText_(avg, label) {
  if (avg == null) return "";
  return "Team avg " + avg.toFixed(1) + (label ? " " + label : "");
}
