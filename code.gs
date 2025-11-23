/*********** CONFIG ***********/
var CONFIG = {
  SCORE_SHEET: "Weekly-scoreboard",      // weekly scoreboard sheet
  PASTE_SHEET: "raw_paste",
  IMPORTED_SHEET: "ImportedWeeks",
  MASTER_SHEET: "masterlist",
  CURRENT_ROUTES_SHEET: "Current-Running list",
  COMPLETE_ROUTES_SHEET: "Complete route list"
};

/************************************************
 * WEEKLY – PUBLIC ENTRY (CALLED FROM SIDEBAR)
 * importFromPaste(rawText)
 ************************************************/
function importFromPaste(rawText) {
  if (!rawText || !String(rawText).trim()) {
    throw new Error("No text received from panel.");
  }
  // Write raw text into raw_paste sheet
  writeRawToPasteSheet_(String(rawText));
  // Run the existing weekly logic that reads from the paste sheet
  importWeeklyFromPasteSheet_();
}

/************************************************
 * WEEKLY – CORE LOGIC (READS FROM raw_paste)
 ************************************************/
function importWeeklyFromPasteSheet_() {
  var ss = SpreadsheetApp.getActive();
  var scoreSheet = ss.getSheetByName(CONFIG.SCORE_SHEET);
  var pasteSheet = ss.getSheetByName(CONFIG.PASTE_SHEET);
  var weekSheet = ss.getSheetByName(CONFIG.IMPORTED_SHEET);
  var masterSheet = ss.getSheetByName(CONFIG.MASTER_SHEET);

  if (!scoreSheet || !pasteSheet || !weekSheet || !masterSheet) {
    throw new Error("One of the sheets is missing: Weekly-scoreboard / raw_paste / ImportedWeeks / masterlist");
  }

  var lastPasteRow = pasteSheet.getLastRow();
  if (lastPasteRow === 0) {
    ss.toast("No data found in 'raw_paste' sheet.", "Import", 5);
    return;
  }
  var raw = pasteSheet.getRange(1, 1, lastPasteRow, 1).getValues();
  var lines = raw.map(function (r) { return String(r[0] || "").trim(); });
  if (!lines.length || (lines.length === 1 && lines[0] === "")) {
    ss.toast("No data found in 'raw_paste' sheet.", "Import", 5);
    return;
  }

  var allText = lines.join(" ");

  // -------- Week + Year --------
  var week = null;
  var year = null;

  var m = allText.match(/Week\s+(\d+)\s*[-–]\s*(20\d{2})/i);
  if (m) {
    week = parseInt(m[1], 10);
    year = parseInt(m[2], 10);
  } else {
    var mw = allText.match(/Week\s+(\d+)/i);
    if (mw) week = parseInt(mw[1], 10);
    var my = allText.match(/\b(20\d{2})\b/);
    if (my) year = parseInt(my[1], 10);
  }
  if (!week) throw new Error("Week number not found in pasted text.");
  if (!year) year = new Date().getFullYear();

  // -------- Check duplicate week --------
  var importedWeeks = getImportedWeeks_(weekSheet);
  if (importedWeeks.indexOf(week) !== -1) {
    ss.toast("Week " + week + " already imported.", "Import", 5);
    pasteSheet.clear();
    return;
  }

  // -------- Load masterlist map --------
  var masterMap = buildMasterMap_(masterSheet);

  // -------- Find header row in pasted text --------
  var headerRow = -1;
  for (var i = 0; i < lines.length; i++) {
    var lower = lines[i].toLowerCase();
    if (lower.indexOf("transporter id delivered dcr dnr dpmo pod cc") !== -1) {
      headerRow = i;
      break;
    }
  }
  if (headerRow === -1) {
    throw new Error("Transporter header row not found in 'raw_paste' sheet.");
  }

  // -------- Parse transporter rows --------
  var tz = ss.getSpreadsheetTimeZone() || "Etc/UTC";
  var weDate = Utilities.formatDate(getWeekEndDate_(week, year), tz, "dd/MM/yyyy");

  var rows = [];

  for (var r = headerRow + 1; r < lines.length; r++) {
    var line = lines[r];
    if (!line) continue;
    if (/^DSP WEEKLY SUMMARY/i.test(line)) break;
    if (/^Page\s+\d+\s+of\s+\d+/i.test(line)) break;

    var parts = line.split(/\s+/);
    if (parts.length < 6) continue;

    var id = parts[0];
    var delivered = parts[1];
    var dcr = parts[2];
    var dnr = parts[3];
    var pod = parts[4];
    var cc = parts[5];

    var info = masterMap[id] || {};
    var driverName = "";
    if (info.first || info.last) {
      driverName = (info.first || "") + " " + (info.last || "");
      driverName = driverName.trim();
    }
    var driverStatus = info.status || "";

    rows.push([
      week,
      weDate,
      driverName,
      driverStatus,
      "",
      id,
      delivered,
      dcr,
      dnr,
      pod,
      cc
    ]);
  }

  if (!rows.length) {
    ss.toast("No transporter rows parsed.", "Import", 5);
    return;
  }

  var startRow = scoreSheet.getLastRow() + 1;
  scoreSheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

  weekSheet.getRange(weekSheet.getLastRow() + 1, 1).setValue(week);

  pasteSheet.clear();

  ss.toast("Imported " + rows.length + " rows for week " + week, "Import Successful", 8);
}

/************************************************
 * WEEKLY HELPERS
 ************************************************/
function buildMasterMap_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  var vals = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var map = {};

  for (var i = 0; i < vals.length; i++) {
    var id = String(vals[i][0] || "").trim();
    if (!id) continue;
    map[id] = {
      first: String(vals[i][1] || "").trim(),
      last: String(vals[i][2] || "").trim(),
      status: String(vals[i][5] || "").trim()
    };
  }
  return map;
}

function getImportedWeeks_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var vals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return vals
    .map(function (r) { return r[0]; })
    .filter(function (v) { return v !== "" && v != null; })
    .map(function (v) { return Number(v); });
}

/** Week start = Saturday, week ends Friday */
function getWeekEndDate_(week, year) {
  var firstJan = new Date(year, 0, 1);
  var day = firstJan.getDay();         // 0=Sun..6=Sat
  var offsetToSaturday = (6 - day + 7) % 7;
  var firstSaturday = new Date(year, 0, 1 + offsetToSaturday);

  var weekStart = new Date(firstSaturday);
  weekStart.setDate(firstSaturday.getDate() + 7 * (week - 1));

  var weekEnd = new Date(weekStart);
  weekEnd.setDate(weekStart.getDate() + 6);
  return weekEnd;
}

/************************************************
 * DAILY ROUTES – SHARED PARSER (READS raw_paste)
 ************************************************/
function parseRoutesFromPaste_() {
  var ss = SpreadsheetApp.getActive();
  var pasteSheet = ss.getSheetByName(CONFIG.PASTE_SHEET);
  if (!pasteSheet) throw new Error(CONFIG.PASTE_SHEET + " sheet not found.");

  var range = pasteSheet.getDataRange();
  var data = range.getValues();
  if (!data.length) throw new Error("No data in " + CONFIG.PASTE_SHEET + " sheet.");

  var tz = ss.getSpreadsheetTimeZone() || "Etc/UTC";
  var today = new Date();
  var todayKey = Utilities.formatDate(today, tz, "yyyy-MM-dd");

  // Find header row
  var headerRowIndex = -1;
  for (var r = 0; r < data.length; r++) {
    var rowText = data[r].map(function (c) { return String(c).toLowerCase(); }).join(" ");
    if (rowText.indexOf("route code") !== -1 &&
        rowText.indexOf("route progress") !== -1 &&
        rowText.indexOf("transporter id") !== -1) {
      headerRowIndex = r;
      break;
    }
  }
  if (headerRowIndex === -1) {
    throw new Error("Route file header row not found in " + CONFIG.PASTE_SHEET + ".");
  }

  var headerRow = data[headerRowIndex].map(function (c) { return String(c).trim().toLowerCase(); });

  function findCol(regex) {
    for (var i = 0; i < headerRow.length; i++) {
      if (regex.test(headerRow[i])) return i;
    }
    return -1;
  }

  var colRouteCode     = findCol(/route\s*code/);
  var colDSP           = findCol(/\bdsp\b/);
  var colTransporterId = findCol(/transporter\s*id/);
  var colDriverName    = findCol(/driver\s*name/);
  var colProgress      = findCol(/route\s*progress/);
  var colDeliveryType  = findCol(/delivery service/i);
  var colDuration      = findCol(/route duration/i);
  var colAllStops      = findCol(/all stops/i);
  var colStopsComplete = findCol(/stops complete/i);
  var colNotStarted    = findCol(/not started/i);

  var requiredCols = [colRouteCode, colTransporterId, colDriverName, colProgress];
  for (var q = 0; q < requiredCols.length; q++) {
    if (requiredCols[q] === -1) {
      throw new Error("Some required daily columns not found (Route code/Transporter ID/Driver name/Route progress).");
    }
  }

  var routes = [];

  for (var r2 = headerRowIndex + 1; r2 < data.length; r2++) {
    var row = data[r2];
    var routeCode = String(row[colRouteCode] || "").trim();
    if (!routeCode) continue;

    var key = todayKey + "|" + routeCode;

    routes.push({
      key: key,
      dateKey: todayKey,
      dateValue: today,
      routeCode: routeCode,
      dsp: String(colDSP === -1 ? "" : row[colDSP] || "").trim(),
      transporterId: String(row[colTransporterId] || "").trim(),
      driverName: String(row[colDriverName] || "").trim(),
      progress: String(row[colProgress] || "").trim(),
      deliveryType: String(colDeliveryType === -1 ? "" : row[colDeliveryType] || "").trim(),
      routeDuration: String(colDuration === -1 ? "" : row[colDuration] || "").trim(),
      allStops: colAllStops === -1 ? "" : row[colAllStops],
      stopsComplete: colStopsComplete === -1 ? "" : row[colStopsComplete],
      notStarted: colNotStarted === -1 ? "" : row[colNotStarted],
      rescuedBy: ""
    });
  }

  if (!routes.length) throw new Error("No route rows found under header.");

  return routes;
}

/************************************************
 * DAILY ROUTES – PUBLIC ENTRY (CALLED FROM SIDEBAR)
 * importDailyFromRaw(rawText, isStartOfDay)
 ************************************************/
function importDailyFromRaw(rawText, isStartOfDay) {
  if (!rawText || !String(rawText).trim()) {
    throw new Error("No text received from panel.");
  }
  writeRawToPasteSheet_(String(rawText));
  if (isStartOfDay) {
    importDailyStartFromPaste();
  } else {
    importDailyEndFromPaste();
  }
}

/************************************************
 * DAILY ROUTES – START OF DAY
 ************************************************/
function importDailyStartFromPaste() {
  var ss = SpreadsheetApp.getActive();
  var currentSheet = ss.getSheetByName(CONFIG.CURRENT_ROUTES_SHEET);
  var completeSheet = ss.getSheetByName(CONFIG.COMPLETE_ROUTES_SHEET);
  if (!currentSheet || !completeSheet) {
    throw new Error("Current-Running list or Complete route list sheet missing.");
  }

  var routes = parseRoutesFromPaste_();

  var values = routes.map(function (r) {
    return [
      "",                // Duplicate Check (formula added later)
      r.dateValue,       // Date
      r.routeCode,
      r.dsp,
      r.transporterId,
      r.driverName,
      r.progress,
      r.deliveryType,
      r.routeDuration,
      r.allStops,
      r.stopsComplete,
      r.notStarted,
      ""                 // Rescued by
    ];
  });

  ensureRouteHeaders_(currentSheet);
  ensureRouteHeaders_(completeSheet);
  formatRouteDateColumn_(currentSheet);
  formatRouteDateColumn_(completeSheet);

  var lastCur = currentSheet.getLastRow();
  if (lastCur > 1) {
    currentSheet.getRange(2, 1, lastCur - 1, 13).clearContent();
  }
  currentSheet.getRange(2, 1, values.length, 13).setValues(values);
  applyDuplicateFormulas_(currentSheet, 2, values.length);

  upsertRoutesIntoComplete_(completeSheet, routes);

  sortCompleteByDateDesc_(completeSheet);

  ss.getSheetByName(CONFIG.PASTE_SHEET).clear();

  ss.toast("Daily START imported: " + routes.length + " routes.", "Daily Routes", 8);
}

/************************************************
 * DAILY ROUTES – END OF DAY
 ************************************************/
function importDailyEndFromPaste() {
  var ss = SpreadsheetApp.getActive();
  var completeSheet = ss.getSheetByName(CONFIG.COMPLETE_ROUTES_SHEET);
  var currentSheet = ss.getSheetByName(CONFIG.CURRENT_ROUTES_SHEET);
  if (!completeSheet || !currentSheet) {
    throw new Error("Complete route list or Current-Running list sheet missing.");
  }

  var routes = parseRoutesFromPaste_();

  var currentMap = buildCurrentRouteMap_(currentSheet);

  applyRescueLogic_(routes, currentMap);

  ensureRouteHeaders_(completeSheet);
  formatRouteDateColumn_(completeSheet);

  upsertRoutesIntoComplete_(completeSheet, routes);

  sortCompleteByDateDesc_(completeSheet);

  ss.getSheetByName(CONFIG.PASTE_SHEET).clear();

  ss.toast("Daily END imported (updated with rescues): " + routes.length + " routes.", "Daily Routes", 8);
}

/************************************************
 * DAILY ROUTES – HELPERS
 ************************************************/
function ensureRouteHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 13).setValues([[
      "Duplicate Check",
      "Date",
      "Route code",
      "DSP",
      "Transporter ID",
      "Driver name",
      "Route progress",
      "Delivery service type",
      "Route Duration",
      "All stops",
      "Stops complete",
      "not started stops",
      "Rescued by"
    ]]);
  }
}

function formatRouteDateColumn_(sheet) {
  sheet.getRange("B:B").setNumberFormat("dd/MM/yyyy");
}

/** Duplicate Check = how many times this Route code (column C) appears */
function applyDuplicateFormulas_(sheet, startRow, count) {
  for (var i = 0; i < count; i++) {
    var rowIndex = startRow + i;
    var formula = '=COUNTIF($C:$C,C' + rowIndex + ')';
    sheet.getRange(rowIndex, 1).setFormula(formula);
  }
}

function sortCompleteByDateDesc_(sheet) {
  var last = sheet.getLastRow();
  if (last > 1) {
    sheet.getRange(2, 1, last - 1, 13).sort({ column: 2, ascending: false });
  }
}

function buildCurrentRouteMap_(sheet) {
  var ss = SpreadsheetApp.getActive();
  var tz = ss.getSpreadsheetTimeZone() || "Etc/UTC";

  var last = sheet.getLastRow();
  var map = {};
  if (last <= 1) return map;

  var rows = sheet.getRange(2, 1, last - 1, 13).getValues();
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var dateCell = row[1]; // B
    var routeCode = String(row[2] || "").trim();
    if (!routeCode || !dateCell) continue;

    var dateKey;
    if (dateCell instanceof Date) {
      dateKey = Utilities.formatDate(dateCell, tz, "yyyy-MM-dd");
    } else {
      var parts = String(dateCell).split("/");
      if (parts.length === 3) {
        dateKey = parts[2] + "-" + parts[1] + "-" + parts[0];
      } else {
        continue;
      }
    }

    var key = dateKey + "|" + routeCode;
    map[key] = {
      transporterId: String(row[4] || "").trim(),
      driverName: String(row[5] || "").trim()
    };
  }
  return map;
}

function applyRescueLogic_(routes, currentMap) {
  for (var i = 0; i < routes.length; i++) {
    var r = routes[i];
    var cur = currentMap[r.key];

    var multi =
      (r.transporterId && r.transporterId.indexOf("|") !== -1) ||
      (r.driverName && r.driverName.indexOf("|") !== -1);

    if (!multi) {
      r.rescuedBy = "";
      continue;
    }

    var ids = r.transporterId ? r.transporterId.split("|") : [];
    var names = r.driverName ? r.driverName.split("|") : [];

    ids = ids.map(function (s) { return s.trim(); });
    names = names.map(function (s) { return s.trim(); });

    var origId = cur ? String(cur.transporterId || "").trim() : "";
    var origName = cur ? String(cur.driverName || "").trim() : "";

    var idx = -1;
    if (origId) idx = ids.indexOf(origId);
    if (idx === -1 && origName) idx = names.indexOf(origName);

    var rescuedNames = [];

    if (idx !== -1) {
      r.transporterId = ids[idx] || origId;
      r.driverName = names[idx] || origName;
      for (var j = 0; j < names.length; j++) {
        if (j !== idx && names[j]) rescuedNames.push(names[j]);
      }
    } else {
      // fallback – first treated as original
      r.transporterId = ids.length ? ids[0] : r.transporterId;
      r.driverName = names.length ? names[0] : r.driverName;
      for (var k = 1; k < names.length; k++) {
        if (names[k]) rescuedNames.push(names[k]);
      }
    }

    r.rescuedBy = rescuedNames.join(", ");
  }
}

function upsertRoutesIntoComplete_(sheet, routes) {
  var ss = SpreadsheetApp.getActive();
  var tz = ss.getSpreadsheetTimeZone() || "Etc/UTC";

  var last = sheet.getLastRow();
  var existingMap = {};
  var existingRows = [];

  if (last > 1) {
    existingRows = sheet.getRange(2, 1, last - 1, 13).getValues();
    for (var i = 0; i < existingRows.length; i++) {
      var row = existingRows[i];
      var dateCell = row[1]; // B
      var routeCode = String(row[2] || "").trim();
      if (!routeCode || !dateCell) continue;

      var dateKey;
      if (dateCell instanceof Date) {
        dateKey = Utilities.formatDate(dateCell, tz, "yyyy-MM-dd");
      } else {
        var parts = String(dateCell).split("/");
        if (parts.length === 3) {
          dateKey = parts[2] + "-" + parts[1] + "-" + parts[0];
        } else {
          continue;
        }
      }

      var key = dateKey + "|" + routeCode;
      existingMap[key] = { index: i + 2 };
    }
  }

  var rowsToAppend = [];

  for (var r = 0; r < routes.length; r++) {
    var rt = routes[r];

    var rowValues = [
      "",                 // Duplicate Check formula later
      rt.dateValue,
      rt.routeCode,
      rt.dsp,
      rt.transporterId,
      rt.driverName,
      rt.progress,
      rt.deliveryType,
      rt.routeDuration,
      rt.allStops,
      rt.stopsComplete,
      rt.notStarted,
      rt.rescuedBy || ""
    ];

    if (existingMap[rt.key]) {
      var rowIndex = existingMap[rt.key].index;
      sheet.getRange(rowIndex, 1, 1, 13).setValues([rowValues]);
      var formula = '=COUNTIF($C:$C,C' + rowIndex + ')';
      sheet.getRange(rowIndex, 1).setFormula(formula);
    } else {
      rowsToAppend.push(rowValues);
    }
  }

  if (rowsToAppend.length) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, 13).setValues(rowsToAppend);
    applyDuplicateFormulas_(sheet, startRow, rowsToAppend.length);
  }
}


/************************************************
 * GUI: HTML Import Sidebar
 ************************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("DNS6 Import")
    .addItem("Open Import Panel", "showImportPanel")
    .addToUi();
}

function showImportPanel() {
  var html = HtmlService.createHtmlOutputFromFile("ImportPanel")
    .setTitle("DNS6 Import");
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Take raw text (rows separated by newline, columns by TAB)
 * and write it into the 'raw_paste' sheet starting at A1.
 */
function writeRawToPasteSheet_(text) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.PASTE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.PASTE_SHEET);
  }

  sheet.clear();

  // Split into rows
  var rawLines = text.split(/\r?\n/);

  // Filter out leading/trailing completely empty rows
  var lines = rawLines.filter(function (line) {
    return String(line).trim() !== "";
  });
  if (!lines.length) return;

  // Build 2D array [row][col], columns split by TAB
  var rows = [];
  var maxCols = 0;

  lines.forEach(function (line) {
    var cols = line.split("\t"); // tab-separated when copied from Excel/Sheets
    rows.push(cols);
    if (cols.length > maxCols) maxCols = cols.length;
  });

  // Pad all rows to same length
  for (var i = 0; i < rows.length; i++) {
    while (rows[i].length < maxCols) {
      rows[i].push("");
    }
  }

  sheet.getRange(1, 1, rows.length, maxCols).setValues(rows);
}
