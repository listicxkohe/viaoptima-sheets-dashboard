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
  // Always send the scorecard PDF first, then any folder PDFs as separate attachments.
  var attachments = [pdf.blob];
  // Prefer extras already collected; fall back to loading from folder if missing.
  var extras = (pdf.extras && pdf.extras.length ? pdf.extras : []) || [];
  if (!extras.length) {
    var attachFolderId = getExportAttachmentFolderId_();
    if (attachFolderId) {
      extras = getFolderPdfBlobs_(attachFolderId);
    }
  }
  if (extras && extras.length) {
    attachments = attachments.concat(extras);
  }

  MailApp.sendEmail({
    to: to,
    subject: subject,
    body: bodyText,
    htmlBody: bodyHtml,
    attachments: attachments,
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

  var weekLabel = buildScorecardWeekLabel_(data);

  appendHeader_(body, data, weekLabel);
  appendMetricCards_(body, data);
  appendMetricsGuidePage_(body);

  doc.saveAndClose();

  var safeLabel = (weekLabel || "overall")
    .replace(/[\\/:*?"<>|]+/g, " ")
    .trim()
    .replace(/\s+/g, "_");
  var blob = doc.getAs(MimeType.PDF);
  var fileName =
    (data.driver.name || "Driver " + id).replace(/[\\/:*?"<>|]/g, "") +
    " - scorecard - " +
    safeLabel +
    ".pdf";
  blob.setName(fileName);

  // Optionally collect supplemental PDFs from a configured Drive folder (emailed as separate attachments).
  var attachFolderId = getExportAttachmentFolderId_();
  var extraPdfs = attachFolderId ? getFolderPdfBlobs_(attachFolderId) : [];

  DriveApp.getFileById(doc.getId()).setTrashed(true);

  return {
    fileName: fileName,
    mimeType: blob.getContentType(),
    blob: blob,
    extras: extraPdfs,
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

  function formatRange(startDate, endDate) {
    if (!startDate || !endDate) return "";
    var start = Utilities.formatDate(startDate, tz, "d MMM yyyy");
    var end = Utilities.formatDate(endDate, tz, "d MMM yyyy");
    return start + " - " + end;
  }

  var range = "";
  if (data.weekWindow && data.weekWindow.start && data.weekWindow.end) {
    range = formatRange(data.weekWindow.start, data.weekWindow.end);
  } else if (data.driver && data.driver.lastWeekLabel) {
    var endGuess = new Date(data.driver.lastWeekLabel);
    if (!isNaN(endGuess.getTime())) {
      endGuess.setHours(0, 0, 0, 0);
      // Align to Saturday–Friday cadence: find the most recent Saturday at/before endGuess.
      var startGuess = new Date(endGuess.getTime());
      while (startGuess.getDay() !== 6) {
        startGuess.setDate(startGuess.getDate() - 1);
      }
      var alignedEnd = new Date(startGuess.getTime());
      alignedEnd.setDate(startGuess.getDate() + 6);
      range = formatRange(startGuess, alignedEnd);
    }
  }

  if (data.appliedWeek != null) {
    return "Week " + data.appliedWeek + (range ? " | " + range : "");
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
    "Avg DNR DPMO: " +
      (metrics.avgDnrDpmo != null ? metrics.avgDnrDpmo.toFixed(1) : "N/A")
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
    buildMetricCard_("Avg DNR DPMO", metrics.avgDnrDpmo != null ? metrics.avgDnrDpmo.toFixed(1) : "N/A", "Lower is better", "#0f766e") +
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

function appendHeader_(body, data, weekLabel) {
  var driver = data.driver || {};
  var table = body.appendTable();
  table.setBorderWidth(0);
  var row = table.appendTableRow();

  var left = row.appendTableCell();
  left.setBackgroundColor("#ffffff");
  left.setPaddingTop(8).setPaddingBottom(8).setPaddingLeft(10).setPaddingRight(10);
  var title = left.appendParagraph("Driver scorecard");
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  left.appendParagraph(driver.name || "Driver").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  left.appendParagraph((driver.transporterId || "N/A") + " | " + (driver.status || "Unknown"));
  if (driver.dspList && driver.dspList.length) {
    left.appendParagraph("DSP: " + driver.dspList.join(", "));
  }
  var scopeLabel = weekLabel || (data && data.appliedWeek != null
    ? "Week " + data.appliedWeek
    : "Overall (all weeks)");
  left.appendParagraph("Scope: " + scopeLabel).setForegroundColor("#475569");

  var badgeRow = left.appendTable();
  badgeRow.setBorderWidth(0);
  var badgeCell = badgeRow.appendTableRow().appendTableCell();
  badgeCell.setBackgroundColor("#e2e8f0");
  badgeCell.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(8).setPaddingRight(8);
  badgeCell.appendParagraph("Performance snapshot").setBold(true).setForegroundColor("#334155");

  var right = row.appendTableCell();
  right.setPaddingTop(16).setPaddingBottom(16).setPaddingLeft(16).setPaddingRight(16);
  right.setBackgroundColor("#eef2ff");
  var scorePara = right.appendParagraph("Score");
  scorePara.setBold(true).setForegroundColor("#6b7280");
  var scoreVal = right.appendParagraph(driver.score != null ? driver.score.toFixed(1) : "N/A");
  scoreVal.setBold(true).setFontSize(36).setForegroundColor("#0f172a");
  var rankLine = driver.rank != null
    ? "Rank #" + driver.rank + (driver.totalDrivers ? " of " + driver.totalDrivers : "")
    : (driver.teamStanding || "");
  right.appendParagraph(rankLine || "").setForegroundColor("#475569").setBold(true);
}

function appendMetricCards_(body, data) {
  var metrics = data.metrics || {};
  var ranks = data.metricRanks || {};
  var trends = data.metricTrends || {};
  var heading = body.appendParagraph("Key metrics");
  heading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  heading.setFontSize(18);
  var cardItems = [
    {
      label: "Total deliveries",
      value: formatScorecardNumber_(metrics.deliveries),
      note: "",
      trend: null,
    },
    {
      label: "Avg DCR",
      value: formatScorecardPercent_(metrics.avgDcr),
      note: (metrics.ratings && metrics.ratings.dcr) || formatScorecardRank_(ranks.dcr),
      trend: buildTrendParts_(trends.dcr, { percent: true, betterHigh: true }),
    },
    {
      label: "Avg DNR DPMO",
      value: metrics.avgDnrDpmo != null ? metrics.avgDnrDpmo.toFixed(1) : "N/A",
      note: metrics.ratings && metrics.ratings.dnr ? metrics.ratings.dnr : "",
      trend: buildTrendParts_(trends.dnr, { percent: false, betterHigh: false, suffix: " pts" }),
    },
    {
      label: "Avg POD",
      value: formatScorecardPercent_(metrics.avgPod),
      note: (metrics.ratings && metrics.ratings.pod) || formatScorecardRank_(ranks.pod),
      trend: buildTrendParts_(trends.pod, { percent: true, betterHigh: true }),
    },
    {
      label: "Avg CC",
      value: formatScorecardPercent_(metrics.avgCc),
      note: (metrics.ratings && metrics.ratings.cc) || formatScorecardRank_(ranks.cc),
      trend: buildTrendParts_(trends.cc, { percent: true, betterHigh: true }),
    },
  ];
  appendCardGrid_(body, cardItems, 2);
}

// Add a second page with concise guidance on core metrics.
function appendMetricsGuidePage_(body) {
  body.appendPageBreak();
  var title = body.appendParagraph("AMAZON DSP METRICS");
  title
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setForegroundColor("#ffffff")
    .setBackgroundColor("#0f172a")
    .setSpacingAfter(8);

  var sections = [
    {
      label: "DNR (Delivered Not Received)",
      meaning: "Customer says it wasn't received after you marked delivered.",
      tips: [
        "Verify address, unit, entrance, and read notes.",
        "Deliver in-hand when possible (OTP in-person only).",
        "Unattended: hide from street view, weather-protect, and take POD.",
        "Stay in geofence; if the pin is wrong, escalate (SDS/DSP).",
      ],
    },
    {
      label: "POD (Photo on Delivery)",
      meaning: "Unattended deliveries need a usable photo.",
      tips: [
        "POD every unattended stop when prompted.",
        "No people; step back and show a landmark.",
        "Use flash; retake if blurry. If no photo option, in-app text exact location.",
      ],
    },
    {
      label: "CC (Contact Compliance)",
      meaning: "Required contact completed in-app (Rabbit).",
      tips: [
        "Use Rabbit call/text only; contact before undeliverable or OTP issues.",
        "Default ladder: Text -> Call -> SDS -> DSP.",
      ],
    },
    {
      label: "DCR (Delivery Completion Rate)",
      meaning: "% delivered vs returned to station.",
      tips: [
        "Do hard stops early (OTP, businesses with hours, access-problem buildings).",
        "Read notes before walking in; prevent failed attempts.",
        "If blocked, run CC flow fast and circle back; use accurate reason + required contact if return is unavoidable.",
      ],
    },
    {
      label: "Score",
      meaning: "A combined rating of all your metrics",
      tips: [
        "Score is out of 100; weighted mainly on quality: DCR, POD, CC, and DNR DPMO.",
        "Delivering more parcels helps a bit.",
        "Rescues: helping others adds a lot; needing help can shave a little.",
      ],
    },
  ];

  sections.forEach(function (section) {
    var heading = body.appendParagraph(section.label);
    heading.setHeading(DocumentApp.ParagraphHeading.HEADING3).setForegroundColor("#0f172a");
    if (section.label === "Score") {
      var meaningBullet = body.appendListItem("Meaning: " + section.meaning);
      meaningBullet.setGlyphType(DocumentApp.GlyphType.BULLET).setForegroundColor("#111827").setBold(false);
      if (section.tips && section.tips.length) {
        var improve = body.appendListItem("Improve by:");
        improve.setGlyphType(DocumentApp.GlyphType.BULLET).setForegroundColor("#111827").setBold(true);
        section.tips.forEach(function (tip) {
          var item = body.appendListItem(tip);
          item.setGlyphType(DocumentApp.GlyphType.BULLET).setForegroundColor("#111827");
        });
      }
      body.appendParagraph(""); // spacer
      return;
    }

    var meaning = body.appendParagraph("Meaning: " + section.meaning);
    meaning.setForegroundColor("#374151").setSpacingAfter(6);
    if (section.tips && section.tips.length) {
      var list = body.appendListItem("Improve by:");
      list.setBold(true).setForegroundColor("#111827");
      section.tips.forEach(function (tip) {
        var item = body.appendListItem(tip);
        item.setGlyphType(DocumentApp.GlyphType.BULLET).setForegroundColor("#111827");
      });
    }
    body.appendParagraph(""); // spacer
  });
}

function appendComparisonCards_(body, data) {
  // Comparison cards removed for a cleaner, focused scorecard.
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
        labelText.setBold(true).setFontSize(12).setForegroundColor("#475569");
        var valueText = cell.appendParagraph(card.value != null ? String(card.value) : "N/A");
        valueText.setFontSize(18).setBold(true);
        if (card.note) {
          var noteText = cell.appendParagraph(card.note);
          noteText.setFontSize(9).setForegroundColor("#6b7280");
        }
        if (card.trend && card.trend.text) {
          var trendText = cell.appendParagraph(card.trend.text);
          trendText.setFontSize(9).setForegroundColor(card.trend.color || "#6b7280");
        }
        cell.setBackgroundColor("#f8fafc");
      } else {
        cell.setBackgroundColor("#ffffff");
      }
      idx++;
    }
  }
}

function buildTrendParts_(trend, options) {
  if (!trend || trend.delta == null) return null;
  options = options || {};
  var betterHigh = options.betterHigh !== false; // default true
  var suffix = options.suffix || (options.percent ? "%" : "");
  var usePercent = options.percent && trend.percent != null;

  var value = usePercent ? trend.percent : trend.delta; // positive = improvement
  var displayVal = betterHigh ? value : -value; // show negative when lower is better and improved
  var prefix = displayVal > 0 ? "+" : "";
  var textValue = displayVal.toFixed(1);
  var text = prefix + textValue + suffix + " vs last week";

  var color = "#6b7280"; // neutral
  if (value > 0) color = "#16a34a";
  if (value < 0) color = "#dc2626";

  return { text: text, color: color };
}

// ---------------------------------------------------------------------------
// External PDF attachments (Drive folder)
// ---------------------------------------------------------------------------

var EXPORT_ATTACH_FOLDER_KEY = "exportAttachFolderId";

/**
 * Persist the Drive folder link/id that contains supplemental PDFs
 * to be merged after the scorecard. Returns true if a valid ID was saved.
 */
function setExportAttachmentFolder(urlOrId) {
  var id = parseDriveFolderId_(urlOrId);
  PropertiesService.getScriptProperties().setProperty(EXPORT_ATTACH_FOLDER_KEY, id || "");
  return !!id;
}

function getExportAttachmentConfig() {
  return { attachFolderId: getExportAttachmentFolderId_() };
}

function saveExportAttachmentConfig(input) {
  input = input || {};
  var raw = input.attachFolderId || input.attachFolder || input.folder || "";
  var ok = setExportAttachmentFolder(raw);
  return { success: ok, attachFolderId: getExportAttachmentFolderId_() };
}

// Compatibility wrappers for deployments that miss the main functions.
function getExportAttachmentConfigCompat() {
  return getExportAttachmentConfig();
}
function saveExportAttachmentConfigCompat(input) {
  return saveExportAttachmentConfig(input);
}

// Legacy getter used only as an emergency fallback from the sidebar.
function getExportAttachmentFolderIdLegacy() {
  return { attachFolderId: getExportAttachmentFolderId_() };
}

function getExportAttachmentFolderId_() {
  return PropertiesService.getScriptProperties().getProperty(EXPORT_ATTACH_FOLDER_KEY) || "";
}

function parseDriveFolderId_(input) {
  if (!input) return "";
  var trimmed = String(input).trim();
  var m = trimmed.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m && m[1]) return m[1];
  if (/^[a-zA-Z0-9_-]{20,}$/.test(trimmed)) return trimmed;
  return "";
}

// Collect all PDF blobs (non-recursive) from a Drive folder.
function getFolderPdfBlobs_(folderId) {
  var out = [];
  if (!folderId) return out;
  var folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (err) {
    return out;
  }
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    var mt = f.getMimeType();
    var name = (f.getName() || "").toLowerCase();
    if (mt === MimeType.PDF || name.endsWith(".pdf")) {
      out.push(f.getBlob().setName(f.getName()));
    }
  }
  return out;
}

/**
 * Merge the base scorecard PDF with all PDFs found in the given Drive folder.
 * Concatenates byte streams to build a single PDF blob.
 */
function mergeWithFolderPdfs_(baseBlob, folderId) {
  if (!baseBlob || !folderId) return baseBlob;
  var folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (err) {
    return baseBlob;
  }

  // Start with the base PDF bytes.
  var allBytes = baseBlob.getBytes() || [];

  var files = folder.getFiles();
  var hasExtra = false;
  while (files.hasNext()) {
    var f = files.next();
    var mt = f.getMimeType();
    var name = (f.getName() || "").toLowerCase();
    if (mt === MimeType.PDF || name.endsWith(".pdf")) {
      var extra = f.getBlob().getBytes() || [];
      // Merge byte arrays without using Function.apply to avoid stack issues on large PDFs.
      var merged = new Uint8Array(allBytes.length + extra.length);
      merged.set(allBytes);
      merged.set(extra, allBytes.length);
      allBytes = merged;
      hasExtra = true;
    }
  }
  if (!hasExtra) return baseBlob;

  // Utilities.newBlob accepts a byte[]; convert Uint8Array if needed.
  var finalBytes = Array.isArray(allBytes) ? allBytes : Array.from(allBytes);
  return Utilities.newBlob(finalBytes, "application/pdf", baseBlob.getName());
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
