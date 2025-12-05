/************************************************
 * PDF EXPORT (SCREENSHOT-STYLE)
 * Builds a quick PDF with inline charts for sharing.
 ************************************************/

function downloadScorecardPdf(transporterId, weekFilter) {
  var pdf = generateScorecardPdf_(transporterId, weekFilter);
  return {
    fileName: pdf.fileName,
    base64: Utilities.base64Encode(pdf.blob.getBytes()),
    mimeType: pdf.mimeType,
  };
}

/**
 * Send a scorecard PDF via email using the generated layout (no snapshots).
 */
function emailScorecardPdf(transporterId, weekFilter) {
  var pdf = generateScorecardPdf_(transporterId, weekFilter);
  var data = pdf.data;

  var weekLabel = buildScorecardWeekLabel_(data);
  var subject =
    "ViaOptima | " +
    (data.driver.name || data.driver.transporterId || "Driver") +
    " scorecard" +
    (weekLabel ? " - " + weekLabel : "");

  var bodyHtml = buildSnapshotEmailHtml_(data, weekLabel, pdf.fileName);
  var bodyText = buildSnapshotEmailText_(data, weekLabel);

  var to = "kiragod00@gmail.com";
  MailApp.sendEmail({
    to: to,
    subject: subject,
    body: bodyText,
    htmlBody: bodyHtml,
    attachments: [pdf.blob],
  });

  return { success: true, sentTo: to };
}

// Internal helper to build the PDF blob; shared by download/email.
function generateScorecardPdf_(transporterId, weekFilter) {
  var id = sanitizeTransporterId_(transporterId);
  if (!id) {
    throw new Error("Transporter ID is required to export a scorecard.");
  }
  var week = normalizeWeekFilter_ ? normalizeWeekFilter_(weekFilter) : null;

  var data = getDriverScorecardData(id, week);
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

  doc.saveAndClose();

  var blob = doc.getAs(MimeType.PDF);
  var fileName =
    (data.driver.name || "Driver " + id).replace(/[\\/:*?"<>|]/g, "") +
    " - scorecard.pdf";
  blob.setName(fileName);

  DriveApp.getFileById(doc.getId()).setTrashed(true);

  return {
    fileName: fileName,
    mimeType: blob.getContentType(),
    blob: blob,
    driverName: data.driver.name || "",
    transporterId: id,
    data: data,
  };
}

function buildScorecardWeekLabel_(data) {
  if (!data) return "";
  var tz =
    (data.weekWindow && data.weekWindow.tz) ||
    (SpreadsheetApp.getActive && SpreadsheetApp.getActive().getSpreadsheetTimeZone
      ? SpreadsheetApp.getActive().getSpreadsheetTimeZone()
      : Session.getScriptTimeZone && Session.getScriptTimeZone()) ||
    "Etc/UTC";

  var range = "";
  if (data.weekWindow && data.weekWindow.start && data.weekWindow.end) {
    var start = Utilities.formatDate(data.weekWindow.start, tz, "d MMM yyyy");
    var end = Utilities.formatDate(data.weekWindow.end, tz, "d MMM yyyy");
    range = start + " - " + end;
  } else if (data.driver && data.driver.lastWeekLabel) {
    range = "Through " + data.driver.lastWeekLabel;
  }

  if (data.appliedWeek != null) {
    return "Week " + data.appliedWeek + (range ? " (" + range + ")" : "");
  }
  return range || "Overall (all weeks)";
}

function buildScorecardEmailText_(data, weekLabel) {
  var driver = data && data.driver ? data.driver : {};
  var metrics = data && data.metrics ? data.metrics : {};
  var rescues = data && data.rescues ? data.rescues : {};
  var comparisons = data && data.comparisons ? data.comparisons : {};
  var lines = [];
  lines.push("Driver scorecard");
  lines.push("Driver: " + (driver.name || "Driver"));
  lines.push("Transporter ID: " + (driver.transporterId || "N/A"));
  lines.push("Status: " + (driver.status || "N/A"));
  lines.push("Scope: " + (weekLabel || "Overall"));
  lines.push("");
  lines.push(
    "Score: " +
      (driver.score != null ? driver.score.toFixed(1) : "N/A") +
      (driver.rank ? " | Rank #" + driver.rank : "")
  );
  lines.push("Deliveries: " + formatScorecardNumber_(metrics.deliveries));
  lines.push(
    "Avg DCR/POD/CC: " +
      [metrics.avgDcr, metrics.avgPod, metrics.avgCc]
        .map(formatScorecardPercent_)
        .join(" | ")
  );
  lines.push(
    "Rescues given/taken: " +
      formatScorecardNumber_(rescues.totalGiven) +
      " / " +
      formatScorecardNumber_(rescues.totalTaken)
  );
  lines.push(
    "Gap to leader: " +
      (comparisons && comparisons.scoreVsTop != null
        ? Math.abs(comparisons.scoreVsTop).toFixed(1) + " pts"
        : "N/A")
  );
  lines.push("");
  lines.push("PDF attached.");
  return lines.join("\n");
}

function buildScorecardEmailHtml_(data, weekLabel, fileName) {
  var driver = data && data.driver ? data.driver : {};
  var metrics = data && data.metrics ? data.metrics : {};
  var rescues = data && data.rescues ? data.rescues : {};
  var comparisons = data && data.comparisons ? data.comparisons : {};
  var dsp =
    driver.dspList && driver.dspList.length
      ? driver.dspList.join(", ")
      : "N/A";
  var rankText = driver.rank
    ? "#" + driver.rank + (driver.totalDrivers ? " of " + driver.totalDrivers : "")
    : "Not ranked yet";
  var scoreText =
    driver.score != null ? driver.score.toFixed(1) + " pts" : "N/A";
  var gapText =
    comparisons && comparisons.scoreVsTop != null
      ? Math.abs(comparisons.scoreVsTop).toFixed(1) + " pts"
      : "N/A";

  function pct(val) {
    return val == null ? "N/A" : Number(val).toFixed(1) + "%";
  }
  function num(val) {
    return formatScorecardNumber_(val);
  }

  var html =
    '<div style="font-family:\'Segoe UI\',Arial,sans-serif;background:#f8fafc;color:#0f172a;">' +
    '<div style="background:#0f172a;color:#e2e8f0;padding:16px 20px;font-size:16px;font-weight:600;">' +
    "ViaOptima | Driver scorecard" +
    "</div>" +
    '<div style="padding:20px;">' +
    '<div style="background:#ffffff;border:1px solid #e5e7eb;border-radius:12px;padding:16px 18px;margin-bottom:12px;">' +
    '<div style="font-size:18px;font-weight:700;margin-bottom:4px;">' +
    (driver.name || "Driver") +
    "</div>" +
    '<div style="font-size:13px;color:#6b7280;margin-bottom:6px;">' +
    "Transporter ID: " +
    (driver.transporterId || "N/A") +
    " | Status: " +
    (driver.status || "N/A") +
    "</div>" +
    '<div style="font-size:13px;color:#6b7280;margin-bottom:6px;">DSP: ' +
    dsp +
    "</div>" +
    '<div style="font-size:13px;color:#111827;font-weight:600;">' +
    "Scope: " +
    (weekLabel || "Overall") +
    "</div>" +
    (driver.summaryNote
      ? '<div style="font-size:12px;color:#6b7280;margin-top:4px;">' +
        driver.summaryNote +
        "</div>"
      : "") +
    "</div>" +
    '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:10px;margin-bottom:12px;">' +
    buildMetricCard_("Score", scoreText, rankText, "#2563eb") +
    buildMetricCard_("Deliveries", num(metrics.deliveries), "Total deliveries", "#0ea5e9") +
    buildMetricCard_("Avg DCR", pct(metrics.avgDcr), "Quality", "#16a34a") +
    buildMetricCard_("Avg POD", pct(metrics.avgPod), "Proof of delivery", "#6366f1") +
    buildMetricCard_("Avg CC", pct(metrics.avgCc), "Customer care", "#f97316") +
    buildMetricCard_(
      "Rescues",
      num(rescues.totalGiven) + " given / " + num(rescues.totalTaken) + " taken",
      "Balance across weeks",
      "#7c3aed"
    ) +
    buildMetricCard_("Gap to leader", gapText, "Delta to #1", "#0ea5e9") +
    "</div>" +
    '<div style="font-size:12px;color:#6b7280;margin-top:8px;">Attached: ' +
    (fileName || "scorecard.pdf") +
    "</div>" +
    "</div>" +
    "</div>";

  return html;
}

function buildSnapshotEmailText_(data, weekLabel) {
  var driver = data && data.driver ? data.driver : {};
  var lines = [];
  lines.push("Hello,");
  lines.push("");
  lines.push(
    "Attached is the latest scorecard for " +
      (driver.name || "your driver") +
      "."
  );
  lines.push("Transporter ID: " + (driver.transporterId || "N/A"));
  lines.push("Scope: " + (weekLabel || "Overall coverage"));
  lines.push("");
  lines.push("Thank you,");
  lines.push("ViaOptima Dashboard");
  return lines.join("\n");
}

function buildSnapshotEmailHtml_(data, weekLabel, fileName) {
  var driver = data && data.driver ? data.driver : {};
  var dsp =
    driver.dspList && driver.dspList.length
      ? driver.dspList.join(", ")
      : "N/A";
  var html =
    '<div style="font-family:\'Segoe UI\',Arial,sans-serif;background:#f6f7fb;color:#0f172a;padding:20px;">' +
    '<div style="max-width:640px;margin:0 auto;background:#ffffff;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden;">' +
    '<div style="background:#0f172a;color:#e2e8f0;padding:14px 18px;font-size:16px;font-weight:600;">' +
    "ViaOptima | Driver scorecard" +
    "</div>" +
    '<div style="padding:18px 18px 4px 18px;font-size:14px;color:#0f172a;">' +
    "<p style=\"margin:0 0 10px 0;\">Hello,</p>" +
    "<p style=\"margin:0 0 10px 0;\">Attached is the latest scorecard for <strong>" +
    (driver.name || "your driver") +
    "</strong>.</p>" +
    '<div style="background:#f8fafc;border:1px solid #e5e7eb;border-radius:10px;padding:12px 14px;margin:10px 0;">' +
    '<div style="font-size:12px;text-transform:uppercase;letter-spacing:0.05em;color:#6b7280;">Summary</div>' +
    '<div style="margin-top:6px;font-size:14px;font-weight:600;color:#0f172a;">' +
    (driver.name || "Driver") +
    "</div>" +
    '<div style="font-size:12px;color:#6b7280;margin-top:2px;">Transporter ID: ' +
    (driver.transporterId || "N/A") +
    "</div>" +
    '<div style="font-size:12px;color:#6b7280;margin-top:2px;">DSP: ' +
    dsp +
    "</div>" +
    '<div style="font-size:12px;color:#111827;margin-top:6px;">Scope: ' +
    (weekLabel || "Overall coverage") +
    "</div>" +
    (driver.status
      ? '<div style="font-size:12px;color:#6b7280;margin-top:2px;">Status: ' +
        driver.status +
        "</div>"
      : "") +
    "</div>" +
    "<p style=\"margin:6px 0 0 0;font-size:12px;color:#6b7280;\">Attachment: " +
    (fileName || "scorecard.pdf") +
    "</p>" +
    "<p style=\"margin:14px 0 6px 0;\">Thank you,<br>ViaOptima Dashboard</p>" +
    "</div>" +
    "</div>" +
    "</div>";
  return html;
}

function buildMetricCard_(label, value, sub, color) {
  return (
    '<div style="background:#ffffff;border:1px solid #e5e7eb;border-radius:12px;padding:12px 14px;">' +
    '<div style="font-size:11px;text-transform:uppercase;letter-spacing:0.05em;color:#6b7280;">' +
    label +
    "</div>" +
    '<div style="font-size:18px;font-weight:700;color:#0f172a;margin-top:4px;">' +
    value +
    "</div>" +
    (sub
      ? '<div style="font-size:12px;color:' +
        (color || "#6b7280") +
        ';margin-top:4px;">' +
        sub +
        "</div>"
      : "") +
    "</div>"
  );
}

function appendHeader_(body, data) {
  var driver = data.driver || {};
  var table = body.appendTable();
  table.setBorderWidth(0);
  var row = table.appendTableRow();
  var left = row.appendTableCell();
  left.setBackgroundColor("#ffffff");
  left.setPaddingTop(8).setPaddingBottom(8).setPaddingLeft(10).setPaddingRight(10);
  left.appendParagraph("Driver scorecard").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  left.appendParagraph(driver.name || "Driver").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  left.appendParagraph((driver.transporterId || "N/A") + " | " + (driver.status || "Unknown"));
  if (driver.dspList && driver.dspList.length) {
    left.appendParagraph("DSP: " + driver.dspList.join(", "));
  }
  left.appendParagraph(
    "Weeks tracked: " + (driver.weeks != null ? driver.weeks : "N/A")
  );
  var scopeLabel = data && data.appliedWeek != null
    ? "Week " + data.appliedWeek
    : "Overall (all weeks)";
  left.appendParagraph("Scope: " + scopeLabel).setForegroundColor("#475569");
  var right = row.appendTableCell();
  right.setPaddingTop(12).setPaddingBottom(12).setPaddingLeft(14).setPaddingRight(14);
  right.setBackgroundColor("#f1f5f9");
  var scorePara = right.appendParagraph("Score");
  scorePara.setBold(true).setForegroundColor("#6b7280");
  var scoreVal = right.appendParagraph(driver.score != null ? driver.score.toFixed(1) : "N/A");
  scoreVal.setBold(true).setFontSize(32).setForegroundColor("#111827");
  right.appendParagraph(driver.teamStanding || "").setForegroundColor("#475569");
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
    { label: "Deliveries / week", value: formatScorecardDelta_(comps.deliveriesPerWeekDiff, ""), note: "Vs team average" },
    { label: "Percentile", value: comps.percentile != null ? "Top " + comps.percentile + "%" : "N/A", note: "" },
  ];
  appendCardGrid_(body, items, 3);
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
