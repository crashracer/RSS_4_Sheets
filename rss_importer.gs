// ============================================================
//  RSS Feed Importer for Google Sheets  v3.1
//  Paste this entire file into Extensions > Apps Script
// ============================================================

const CONFIG = {
  CONTROL_SHEET:   "RSS Config",
  LOG_SHEET:       "Import Log",
  COMBINED_SHEET:  "All Feeds",    // Combined sheet name
  COMBINED_MODE:   true,           // true = write a combined "All Feeds" sheet
  INDIVIDUAL_MODE: true,           // true = also write individual sheets per feed
  ENABLED_COL:     1,   // A – Enabled checkbox
  URL_COL:         2,   // B – Feed URL
  SHEET_COL:       3,   // C – Destination sheet name
  MAX_ITEMS_COL:   4,   // D – Max items to import
  STATUS_COL:      5,   // E – Status (auto-written)
  FAIL_COUNT_COL:  6,   // F – Failure counter (hidden)
  LAST_FAIL_COL:   7,   // G – Last failure timestamp (hidden)
  HEADER_ROW:      1,
  DATA_START_ROW:  2,
  FEED_TIMEOUT_MS: 30000,    // 30 seconds per feed
  MAX_FAILS_24H:   2,        // Failures before auto-skip
  FAIL_WINDOW_MS:  86400000  // 24 hours in ms
};

// ─────────────────────────────────────────────────────────────
//  MENU
// ─────────────────────────────────────────────────────────────
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("📰 RSS Feeds")
    .addItem("▶ Import All Feeds",           "importAllFeeds")
    .addSeparator()
    .addSubMenu(ui.createMenu("⏱ Scheduler")
      .addItem("Set: Every 15 minutes",      "scheduleTrigger15min")
      .addItem("Set: Every hour",            "scheduleTrigger1hr")
      .addItem("Set: Every 6 hours",         "scheduleTrigger6hr")
      .addItem("Set: Daily (midnight)",      "scheduleTriggerDaily")
      .addItem("Set: Custom interval…",      "scheduleTriggerCustom")
      .addSeparator()
      .addItem("📋 View active triggers",    "viewTriggers")
      .addItem("🗑 Delete all triggers",     "deleteAllTriggers"))
    .addSeparator()
    .addSubMenu(ui.createMenu("📤 Export to Drive")
      .addItem("Export feeds CSV (today)",   "exportCombinedCSV")
      .addItem("Export log CSV (today)",     "exportLogCSV"))
    .addSeparator()
    .addItem("⚙ Setup Config Sheet",         "setupConfigSheet")
    .addItem("🔓 Reset auto-skipped feeds",  "resetAutoSkipped")
    .addItem("🗑 Clear Log",                 "clearLog")
    .addToUi();
}

// ─────────────────────────────────────────────────────────────
//  LOGGING
// ─────────────────────────────────────────────────────────────
var _logSheet = null;

function getLogSheet(ss) {
  if (_logSheet) return _logSheet;
  _logSheet = ss.getSheetByName(CONFIG.LOG_SHEET);
  if (!_logSheet) {
    _logSheet = ss.insertSheet(CONFIG.LOG_SHEET);
    _logSheet.getRange(1, 1, 1, 4)
      .setValues([["Timestamp", "Feed", "Level", "Message"]])
      .setFontWeight("bold")
      .setBackground("#37474f")
      .setFontColor("#ffffff");
    _logSheet.setColumnWidth(1, 160);
    _logSheet.setColumnWidth(2, 260);
    _logSheet.setColumnWidth(3, 80);
    _logSheet.setColumnWidth(4, 500);
    _logSheet.setFrozenRows(1);
  }
  return _logSheet;
}

function log(ss, feed, level, message) {
  var sheet   = getLogSheet(ss);
  var ts      = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  sheet.appendRow([ts, feed || "", level, message]);
  var lastRow = sheet.getLastRow();
  var lc = sheet.getRange(lastRow, 3);
  var mc = sheet.getRange(lastRow, 4);
  if      (level === "ERROR") { lc.setBackground("#ffcdd2").setFontColor("#b71c1c"); mc.setFontColor("#b71c1c"); }
  else if (level === "WARN")  { lc.setBackground("#fff9c4").setFontColor("#f57f17"); }
  else if (level === "OK")    { lc.setBackground("#c8e6c9").setFontColor("#1b5e20"); }
  else                        { lc.setBackground(null).setFontColor("#212121"); }
  if (lastRow > 501) sheet.deleteRow(2); // keep last 500 entries
  SpreadsheetApp.flush();
}

function clearLog() {
  _logSheet = null;
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.LOG_SHEET);
  if (sheet && sheet.getLastRow() > 1)
    sheet.deleteRows(2, sheet.getLastRow() - 1);
}

// ─────────────────────────────────────────────────────────────
//  SETUP CONFIG SHEET
// ─────────────────────────────────────────────────────────────
function setupConfigSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.CONTROL_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.CONTROL_SHEET, 0);

  sheet.clearContents();
  sheet.clearFormats();

  var headers = ["Enabled", "Feed URL", "Destination Sheet", "Max Items", "Status", "Fail Count", "Last Fail"];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#4a90d9")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  var examples = [
    ["https://feeds.bbci.co.uk/news/rss.xml",                     "BBC News", 10, "", "", 0, ""],
    ["https://rss.nytimes.com/services/xml/rss/nyt/HomePage.xml", "NYT",      10, "", "", 0, ""],
  ];
  sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.URL_COL, examples.length, 6).setValues(examples);

  var cbRange = sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.ENABLED_COL, examples.length, 1);
  cbRange.insertCheckboxes();
  cbRange.setValue(true);
  cbRange.setHorizontalAlignment("center");

  sheet.setColumnWidth(CONFIG.ENABLED_COL,   70);
  sheet.setColumnWidth(CONFIG.URL_COL,       360);
  sheet.setColumnWidth(CONFIG.SHEET_COL,     180);
  sheet.setColumnWidth(CONFIG.MAX_ITEMS_COL, 100);
  sheet.setColumnWidth(CONFIG.STATUS_COL,    280);
  sheet.hideColumns(CONFIG.FAIL_COUNT_COL);  // F – internal
  sheet.hideColumns(CONFIG.LAST_FAIL_COL);   // G – internal
  sheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert(
    "✅ Config sheet ready!\n\n" +
    "• Col A – Checkbox: enable/disable feed\n" +
    "• Col B – RSS or Atom feed URL\n" +
    "• Col C – Destination sheet name (also becomes the named range)\n" +
    "• Col D – Max items per import (default 25)\n\n" +
    "Cols F & G are hidden and managed internally for failure tracking."
  );
}

// ─────────────────────────────────────────────────────────────
//  AUTO-SKIP: FAILURE TRACKING
// ─────────────────────────────────────────────────────────────

// Called after any failure. Returns true if the feed was just auto-skipped.
function recordFailure(configSheet, sheetRow) {
  var now         = new Date().getTime();
  var failCount   = parseInt(configSheet.getRange(sheetRow, CONFIG.FAIL_COUNT_COL).getValue()) || 0;
  var lastFailRaw = configSheet.getRange(sheetRow, CONFIG.LAST_FAIL_COL).getValue();
  var lastFail    = lastFailRaw ? new Date(lastFailRaw).getTime() : 0;

  // Reset window if it's been more than 24h since last failure
  if (now - lastFail > CONFIG.FAIL_WINDOW_MS) failCount = 0;

  failCount++;
  configSheet.getRange(sheetRow, CONFIG.FAIL_COUNT_COL).setValue(failCount);
  configSheet.getRange(sheetRow, CONFIG.LAST_FAIL_COL).setValue(new Date().toISOString());

  if (failCount >= CONFIG.MAX_FAILS_24H) {
    configSheet.getRange(sheetRow, CONFIG.ENABLED_COL).setValue(false);
    // Record in script properties with the timestamp of auto-skip
    var props   = PropertiesService.getScriptProperties();
    var skipped = JSON.parse(props.getProperty("autoSkipped") || "{}");
    skipped[sheetRow] = now;
    props.setProperty("autoSkipped", JSON.stringify(skipped));
    return true;
  }
  return false;
}

// Called at the start of each import run.
// Silently re-enables any feed whose 24h window has passed and retries it this run.
function checkAndResetAutoSkipped(configSheet) {
  var props   = PropertiesService.getScriptProperties();
  var skipped = JSON.parse(props.getProperty("autoSkipped") || "{}");
  var now     = new Date().getTime();
  var changed = false;

  Object.keys(skipped).forEach(function(row) {
    if (now - skipped[row] > CONFIG.FAIL_WINDOW_MS) {
      var r = parseInt(row);
      configSheet.getRange(r, CONFIG.ENABLED_COL).setValue(true);   // re-check
      configSheet.getRange(r, CONFIG.FAIL_COUNT_COL).setValue(0);   // reset counter
      configSheet.getRange(r, CONFIG.LAST_FAIL_COL).setValue("");    // clear timestamp
      delete skipped[row];
      changed = true;
    }
  });

  if (changed) props.setProperty("autoSkipped", JSON.stringify(skipped));
}

// Manual reset from menu
function resetAutoSkipped() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(CONFIG.CONTROL_SHEET);
  if (!configSheet) { SpreadsheetApp.getUi().alert("No config sheet found."); return; }

  var props   = PropertiesService.getScriptProperties();
  var skipped = JSON.parse(props.getProperty("autoSkipped") || "{}");
  var count   = Object.keys(skipped).length;

  if (count === 0) { SpreadsheetApp.getUi().alert("No auto-skipped feeds to reset."); return; }

  Object.keys(skipped).forEach(function(row) {
    var r = parseInt(row);
    configSheet.getRange(r, CONFIG.ENABLED_COL).setValue(true);
    configSheet.getRange(r, CONFIG.FAIL_COUNT_COL).setValue(0);
    configSheet.getRange(r, CONFIG.LAST_FAIL_COL).setValue("");
    configSheet.getRange(r, CONFIG.STATUS_COL)
      .setValue("🔓 Manually reset – will retry on next import")
      .setFontColor("#1565c0");
  });

  props.setProperty("autoSkipped", "{}");
  SpreadsheetApp.getUi().alert("✅ Reset " + count + " auto-skipped feed(s). They will be retried on the next import.");
}

// ─────────────────────────────────────────────────────────────
//  NAMED RANGES
// ─────────────────────────────────────────────────────────────
function registerNamedRange(ss, sheet, sheetName) {
  var rangeName = sheetName.replace(/[^A-Za-z0-9_]/g, "_");
  var numRows   = Math.max(sheet.getLastRow() - 1, 1);
  var dataRange = sheet.getRange(2, 1, numRows, sheet.getLastColumn() || 1);

  // Remove any existing named range with the same name
  ss.getNamedRanges().forEach(function(nr) {
    if (nr.getName() === rangeName) nr.remove();
  });

  ss.setNamedRange(rangeName, dataRange);
  return rangeName;
}

// ─────────────────────────────────────────────────────────────
//  COMBINED SHEET  (source name as first column)
// ─────────────────────────────────────────────────────────────
function initCombinedSheet(ss) {
  var sheet = ss.getSheetByName(CONFIG.COMBINED_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.COMBINED_SHEET);
  else sheet.clearContents();

  var headers = ["Source", "Title", "Link", "Published", "Description", "Author", "Category"];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#37474f")
    .setFontColor("#ffffff");
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 140);  // Source
  sheet.setColumnWidth(2, 220);  // Title
  sheet.setColumnWidth(3, 280);  // Link
  sheet.setColumnWidth(4, 140);  // Published
  sheet.setColumnWidth(5, 380);  // Description
  return sheet;
}

function appendToCombined(combinedSheet, sourceName, items) {
  if (!items || items.length === 0) return;
  var rows = items.map(function(item) {
    return [sourceName, item.title, item.link, item.published, item.description, item.author, item.category];
  });
  var startRow = combinedSheet.getLastRow() + 1;
  combinedSheet.getRange(startRow, 1, rows.length, 7).setValues(rows);
}

// ─────────────────────────────────────────────────────────────
//  MAIN IMPORT
// ─────────────────────────────────────────────────────────────
function importAllFeeds() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  _logSheet = null;

  var configSheet = ss.getSheetByName(CONFIG.CONTROL_SHEET);
  if (!configSheet) { setupConfigSheet(); return; }

  // Silently re-enable any feeds whose 24h window has expired
  checkAndResetAutoSkipped(configSheet);

  var lastRow = configSheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return;

  var numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  var rows    = configSheet.getRange(CONFIG.DATA_START_ROW, 1, numRows, 7).getValues();

  log(ss, null, "INFO", "═══ Import run started ═══");

  var combinedSheet = CONFIG.COMBINED_MODE ? initCombinedSheet(ss) : null;

  rows.forEach(function(row, i) {
    var sheetRow = CONFIG.DATA_START_ROW + i;
    var enabled  = row[CONFIG.ENABLED_COL  - 1];
    var url      = String(row[CONFIG.URL_COL - 1]).trim();
    var destName = String(row[CONFIG.SHEET_COL - 1]).trim() || "Feed_" + sheetRow;
    var maxItems = parseInt(row[CONFIG.MAX_ITEMS_COL - 1]) || 25;

    if (!url) return;

    if (enabled === false || enabled === "FALSE" || enabled === "") {
      log(ss, url, "INFO", "Skipped (disabled)");
      configSheet.getRange(sheetRow, CONFIG.STATUS_COL)
        .setValue("⏸ Disabled").setFontColor("#757575");
      return;
    }

    log(ss, url, "INFO", "Starting → \"" + destName + "\", max: " + maxItems);
    var startTime = new Date().getTime();

    try {
      var result = importFeedWithTimeout(ss, url, destName, maxItems, startTime);

      if (result.timedOut) {
        var msg = "⏱ Timed out after 30s at: " + result.stage;
        log(ss, url, "WARN", msg);
        configSheet.getRange(sheetRow, CONFIG.STATUS_COL).setValue(msg).setFontColor("#e65100");
        if (recordFailure(configSheet, sheetRow)) {
          log(ss, url, "WARN", "Auto-skipped: " + CONFIG.MAX_FAILS_24H + " failures within 24h — will auto-retry after 24h");
          configSheet.getRange(sheetRow, CONFIG.STATUS_COL)
            .setValue("🚫 Auto-skipped (" + CONFIG.MAX_FAILS_24H + " failures in 24h) — auto-retries after 24h")
            .setFontColor("#b71c1c");
        }

      } else {
        var elapsed   = ((new Date().getTime() - startTime) / 1000).toFixed(1);
        var statusMsg = "✅ " + result.count + " items — " + new Date().toLocaleString() + " (" + elapsed + "s)";
        log(ss, url, "OK", "Done — " + result.count + " items in " + elapsed + "s (" + result.format + ")");
        configSheet.getRange(sheetRow, CONFIG.STATUS_COL).setValue(statusMsg).setFontColor("#2e7d32");
        // Clear failure counters on success
        configSheet.getRange(sheetRow, CONFIG.FAIL_COUNT_COL).setValue(0);
        configSheet.getRange(sheetRow, CONFIG.LAST_FAIL_COL).setValue("");
        // Append to combined sheet
        if (combinedSheet && result.items) appendToCombined(combinedSheet, destName, result.items);
      }

    } catch(err) {
      var elapsed = ((new Date().getTime() - startTime) / 1000).toFixed(1);
      log(ss, url, "ERROR", err.message + " (after " + elapsed + "s)");
      configSheet.getRange(sheetRow, CONFIG.STATUS_COL)
        .setValue("❌ " + err.message).setFontColor("#c62828");
      if (recordFailure(configSheet, sheetRow)) {
        log(ss, url, "WARN", "Auto-skipped: " + CONFIG.MAX_FAILS_24H + " failures within 24h — will auto-retry after 24h");
        configSheet.getRange(sheetRow, CONFIG.STATUS_COL)
          .setValue("🚫 Auto-skipped (" + CONFIG.MAX_FAILS_24H + " failures in 24h) — auto-retries after 24h")
          .setFontColor("#b71c1c");
      }
    }
  });

  // Register named range for the combined sheet
  if (combinedSheet && combinedSheet.getLastRow() > 1) {
    var nr = registerNamedRange(ss, combinedSheet, CONFIG.COMBINED_SHEET.replace(/\s/g, "_"));
    log(ss, null, "INFO", "Combined sheet named range registered: " + nr);
  }

  log(ss, null, "INFO", "═══ Import run finished ═══");
}

// ─────────────────────────────────────────────────────────────
//  CORE: FETCH + PARSE WITH TIMEOUT
//  Returns: { count, format, items, timedOut, stage }
// ─────────────────────────────────────────────────────────────
function importFeedWithTimeout(ss, url, destSheetName, maxItems, startTime) {
  var stage = "init";

  function timedOut(s) {
    stage = s;
    return (new Date().getTime() - startTime) > CONFIG.FEED_TIMEOUT_MS;
  }

  // ── 1. Fetch ────────────────────────────────────────────────
  log(ss, url, "INFO", "[fetch] Sending HTTP request…");
  if (timedOut("pre-fetch")) return { timedOut: true, stage: stage };

  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (timedOut("post-fetch")) return { timedOut: true, stage: stage };

  var httpCode = response.getResponseCode();
  log(ss, url, "INFO", "[fetch] HTTP " + httpCode);
  if (httpCode !== 200) throw new Error("HTTP " + httpCode);

  // ── 2. Parse XML ────────────────────────────────────────────
  log(ss, url, "INFO", "[parse] Parsing XML…");
  var xml      = response.getContentText();
  var doc      = XmlService.parse(xml);
  var root     = doc.getRootElement();
  var rootName = root.getName().toLowerCase();
  if (timedOut("XML parse")) return { timedOut: true, stage: stage };
  log(ss, url, "INFO", "[parse] Root element: <" + rootName + ">");

  var items  = [];
  var format = "unknown";

  // ── 3. Extract items ────────────────────────────────────────
  if (rootName === "rss") {
    format = "RSS 2.0";
    log(ss, url, "INFO", "[parse] Format: RSS 2.0");
    var channel = root.getChild("channel");
    if (!channel) throw new Error("No <channel> element in RSS feed");
    var rawItems = channel.getChildren("item");
    log(ss, url, "INFO", "[parse] Found " + rawItems.length + " items, importing up to " + maxItems);
    rawItems.slice(0, maxItems).forEach(function(item, idx) {
      if (timedOut("item #" + idx)) return;
      items.push({
        title:       getText(item, "title"),
        link:        getText(item, "link"),
        published:   getText(item, "pubDate"),
        description: stripHtml(getText(item, "description")),
        author:      getText(item, "author") || getText(item, "creator") || "",
        category:    item.getChildren("category").map(function(c) { return c.getValue(); }).join(", ")
      });
    });

  } else if (rootName === "feed") {
    format = "Atom";
    log(ss, url, "INFO", "[parse] Format: Atom");
    var atomNs     = XmlService.getNamespace("http://www.w3.org/2005/Atom");
    var rawEntries = root.getChildren("entry", atomNs);
    log(ss, url, "INFO", "[parse] Found " + rawEntries.length + " entries, importing up to " + maxItems);
    rawEntries.slice(0, maxItems).forEach(function(entry, idx) {
      if (timedOut("entry #" + idx)) return;
      var linkEl = entry.getChild("link", atomNs);
      var link   = linkEl && linkEl.getAttribute("href") ? linkEl.getAttribute("href").getValue() : "";
      items.push({
        title:       getTextNs(entry, "title",     atomNs),
        link:        link,
        published:   getTextNs(entry, "published", atomNs) || getTextNs(entry, "updated", atomNs),
        description: stripHtml(getTextNs(entry, "summary", atomNs) || getTextNs(entry, "content", atomNs)),
        author:      "",
        category:    ""
      });
    });

  } else {
    throw new Error("Unsupported feed format — root is <" + rootName + ">");
  }

  if (timedOut("post-parse")) return { timedOut: true, stage: stage };
  if (items.length === 0) throw new Error("Feed parsed OK but contained 0 items");

  // ── 4. Write individual destination sheet ───────────────────
  if (CONFIG.INDIVIDUAL_MODE) {
    log(ss, url, "INFO", "[write] Writing " + items.length + " rows to \"" + destSheetName + "\"…");
    var ss2  = SpreadsheetApp.getActiveSpreadsheet();
    var dest = ss2.getSheetByName(destSheetName);
    if (!dest) dest = ss2.insertSheet(destSheetName);
    else dest.clearContents();

    var hdrs = ["Title", "Link", "Published", "Description", "Author", "Category"];
    dest.getRange(1, 1, 1, hdrs.length)
      .setValues([hdrs])
      .setFontWeight("bold")
      .setBackground("#4a90d9")
      .setFontColor("#ffffff");

    var dataRows = items.map(function(it) {
      return [it.title, it.link, it.published, it.description, it.author, it.category];
    });
    dest.getRange(2, 1, dataRows.length, 6).setValues(dataRows);
    [1, 3, 5, 6].forEach(function(c) { dest.autoResizeColumn(c); });
    dest.setColumnWidth(2, 300);
    dest.setColumnWidth(4, 400);
    dest.setFrozenRows(1);

    var rangeName = registerNamedRange(ss2, dest, destSheetName);
    log(ss, url, "INFO", "[write] Named range registered: " + rangeName);

    if (timedOut("sheet write")) return { timedOut: true, stage: stage };
    log(ss, url, "INFO", "[write] Complete");
  }

  return { count: items.length, format: format, items: items, timedOut: false, stage: "done" };
}

// ─────────────────────────────────────────────────────────────
//  SCHEDULER
// ─────────────────────────────────────────────────────────────
function deleteAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "importAllFeeds") ScriptApp.deleteTrigger(t);
  });
  SpreadsheetApp.getUi().alert("✅ All import triggers deleted.");
}

function viewTriggers() {
  var triggers = ScriptApp.getProjectTriggers().filter(function(t) {
    return t.getHandlerFunction() === "importAllFeeds";
  });
  if (triggers.length === 0) {
    SpreadsheetApp.getUi().alert("No active import triggers.\n\nUse Scheduler menu to create one.");
    return;
  }
  var lines = triggers.map(function(t, i) {
    return (i + 1) + ". " + t.getHandlerFunction() +
           " [ID: " + t.getUniqueId().substring(0, 8) + "…]";
  });
  SpreadsheetApp.getUi().alert(
    triggers.length + " active trigger(s):\n\n" + lines.join("\n") +
    "\n\nUse Scheduler → Delete all triggers to remove."
  );
}

function _createMinuteTrigger(minutes) {
  deleteAllTriggers();
  ScriptApp.newTrigger("importAllFeeds").timeBased().everyMinutes(minutes).create();
  SpreadsheetApp.getUi().alert("✅ Trigger set: every " + minutes + " minute(s).");
}

function scheduleTrigger15min() { _createMinuteTrigger(15);  }
function scheduleTrigger1hr()   { _createMinuteTrigger(60);  }
function scheduleTrigger6hr()   { _createMinuteTrigger(360); }

function scheduleTriggerDaily() {
  deleteAllTriggers();
  ScriptApp.newTrigger("importAllFeeds").timeBased().everyDays(1).atHour(0).create();
  SpreadsheetApp.getUi().alert("✅ Trigger set: daily at midnight.");
}

function scheduleTriggerCustom() {
  var ui  = SpreadsheetApp.getUi();
  var res = ui.prompt("Custom Trigger", "Enter interval in minutes (1–1440):", ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var mins = parseInt(res.getResponseText());
  if (isNaN(mins) || mins < 1 || mins > 1440) {
    ui.alert("Invalid value. Enter a number between 1 and 1440.");
    return;
  }
  _createMinuteTrigger(mins);
}

// ─────────────────────────────────────────────────────────────
//  CSV EXPORT TO DRIVE
//  Naming convention:
//    Feeds-YYYY-MM-DD.csv   (combined all-feeds sheet)
//    Log-YYYY-MM-DD.csv     (import log)
//  If the file already exists for today it is overwritten.
// ─────────────────────────────────────────────────────────────
function getSheetFolder() {
  var file    = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var parents = file.getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}

function todayStamp() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function sheetToCsv(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return "";
  return sheet.getDataRange().getValues().map(function(row) {
    return row.map(function(cell) {
      var v = String(cell).replace(/"/g, '""');
      return (v.indexOf(",") !== -1 || v.indexOf("\n") !== -1 || v.indexOf('"') !== -1)
        ? '"' + v + '"' : v;
    }).join(",");
  }).join("\n");
}

function saveDailyCSV(baseName, csvContent) {
  var folder   = getSheetFolder();
  var fileName = baseName + "-" + todayStamp() + ".csv";

  // Overwrite today's file if it exists
  var existing = folder.getFilesByName(fileName);
  while (existing.hasNext()) existing.next().setTrashed(true);

  folder.createFile(fileName, csvContent, MimeType.CSV);
  return folder.getName() + "/" + fileName;
}

function exportCombinedCSV() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.COMBINED_SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("No combined feed data found. Run an import first.");
    return;
  }
  var path = saveDailyCSV("Feeds", sheetToCsv(sheet));
  SpreadsheetApp.getUi().alert("✅ Feeds exported to Drive:\n" + path);
}

function exportLogCSV() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.LOG_SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("Import Log is empty.");
    return;
  }
  var path = saveDailyCSV("Log", sheetToCsv(sheet));
  SpreadsheetApp.getUi().alert("✅ Log exported to Drive:\n" + path);
}

// ─────────────────────────────────────────────────────────────
//  HELPERS
// ─────────────────────────────────────────────────────────────
function getText(el, name) {
  try { return (el.getChild(name) || { getValue: function() { return ""; } }).getValue(); }
  catch(e) { return ""; }
}

function getTextNs(el, name, ns) {
  try { return (el.getChild(name, ns) || { getValue: function() { return ""; } }).getValue(); }
  catch(e) { return ""; }
}

function stripHtml(html) {
  if (!html) return "";
  return html
    .replace(/<[^>]+>/g, " ")
    .replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&nbsp;/g, " ")
    .replace(/\s+/g, " ").trim();
}
