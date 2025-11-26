/************************************************
 * EXPORT HELPERS
 * Handles scorecard PDF downloads.
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
  body.setMarginTop(24).setMarginBottom(24).setMarginLeft(36).setMarginRight(36);

  appendHeader_(body, data);
  appendMetrics_(body, data);
  appendComparisons_(body, data);
  appendAdditional_(body, data);
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
  body.appendParagraph("Driver Scorecard").setHeading(
    DocumentApp.ParagraphHeading.HEADING1
  );
  body.appendParagraph(driver.name || "Driver").setHeading(
    DocumentApp.ParagraphHeading.HEADING2
  );
  body.appendParagraph(
    "ID: " +
      (driver.transporterId || "N/A") +
      "   |   Status: " +
      (driver.status || "Unknown")
  );
  if (driver.dspList && driver.dspList.length) {
    body.appendParagraph("DSP: " + driver.dspList.join(", "));
  }
  body.appendParagraph(
    "Score: " +
      (driver.score != null ? driver.score.toFixed(1) : "N/A") +
      (driver.teamStanding ? "   •   " + driver.teamStanding : "")
  );
  body.appendHorizontalRule();
}

function appendMetrics_(body, data) {
  var metrics = data.metrics || {};
  var ranks = data.metricRanks || {};
  body.appendParagraph("Key metrics").setHeading(
    DocumentApp.ParagraphHeading.HEADING3
  );
  var items = [
    {
      label: "Total deliveries",
      value: formatScorecardNumber_(metrics.deliveries),
      note: formatScorecardRank_(ranks.deliveries),
    },
    {
      label: "Avg DCR",
      value: formatScorecardPercent_(metrics.avgDcr),
      note: formatScorecardRank_(ranks.dcr),
    },
    {
      label: "Avg POD",
      value: formatScorecardPercent_(metrics.avgPod),
      note: formatScorecardRank_(ranks.pod),
    },
    {
      label: "Avg CC",
      value: formatScorecardPercent_(metrics.avgCc),
      note: formatScorecardRank_(ranks.cc),
    },
  ];
  appendCardGrid_(body, items, 2);
}

function appendComparisons_(body, data) {
  var comparisons = data.comparisons || {};
  body.appendParagraph("Performance vs team").setHeading(
    DocumentApp.ParagraphHeading.HEADING3
  );
  var items = [
    { label: "Score vs team", value: formatScorecardDelta_(comparisons.scoreDiff, "pts"), note: comparisons.teamStanding || "" },
    { label: "Gap to rank #1", value: formatScorecardDelta_(-comparisons.scoreVsTop, "pts"), note: "Points away from leader" },
    { label: "Deliveries vs avg", value: formatScorecardDelta_(comparisons.deliveriesDiff, ""), note: buildAvgText_(comparisons.averages && comparisons.averages.deliveries, "deliveries") },
    { label: "Quality vs avg", value: formatScorecardDelta_(comparisons.qualityDiff, "pts"), note: buildAvgText_(comparisons.averages && comparisons.averages.quality, "quality") },
    { label: "Rescue balance", value: formatScorecardDelta_(comparisons.rescueDiff, ""), note: buildAvgText_(comparisons.averages && comparisons.averages.rescueBalance, "balance") },
    { label: "Deliveries / week", value: formatScorecardDelta_(comparisons.deliveriesPerWeekDiff, ""), note: "Vs team average" },
    { label: "Experience", value: formatScorecardDelta_(comparisons.weeksDiff, "wks"), note: "Weeks vs average" },
    { label: "Percentile", value: comparisons.percentile != null ? "Top " + comparisons.percentile + "%" : "N/A", note: "" },
  ];
  appendCardGrid_(body, items, 3);
}

function appendAdditional_(body, data) {
  var additional = data.additional || {};
  body.appendParagraph("Additional metrics").setHeading(
    DocumentApp.ParagraphHeading.HEADING3
  );
  var items = [
    { label: "Avg stops per route", value: formatScorecardNumber_(additional.avgStops) },
    { label: "Total weeks", value: formatScorecardNumber_(additional.totalWeeks) },
    { label: "Total routes", value: formatScorecardNumber_(additional.totalRoutes) },
  ];
  appendCardGrid_(body, items, 3);
}

function appendRescues_(body, data) {
  var rescues = data.rescues || {};
  body.appendParagraph("Rescue activity").setHeading(
    DocumentApp.ParagraphHeading.HEADING3
  );
  var items = [
    { label: "Total given", value: formatScorecardNumber_(rescues.totalGiven) },
    { label: "Total taken", value: formatScorecardNumber_(rescues.totalTaken) },
    { label: "Given per week", value: formatScorecardNumber_(rescues.avgGivenPerWeek) },
    { label: "Taken per week", value: formatScorecardNumber_(rescues.avgTakenPerWeek) },
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
      cell.setBorderWidth(0).setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(8).setPaddingRight(8);
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
        cell.setBackgroundColor("#f8fafc");
      } else {
        cell.setBackgroundColor("#ffffff");
      }
      idx++;
    }
  }
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

function buildHistoryChangeLabel_(row) {
  if (!row || row.deliveredChange == null) {
    return "No prior data";
  }
  var pct =
    row.deliveredChangePct != null
      ? Math.abs(row.deliveredChangePct).toFixed(1) + "%"
      : "";
  if (row.deliveredChange > 0) {
    return "Up " + pct;
  }
  if (row.deliveredChange < 0) {
    return "Down " + pct;
  }
  return "Flat";
}

function buildAvgText_(avg, label) {
  if (avg == null) return "";
  return "Team avg " + avg.toFixed(1) + (label ? " " + label : "");
}
