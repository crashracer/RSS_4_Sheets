// ============================================================
//  RSS Feed Importer for Google Sheets
//  Paste this entire file into Extensions > Apps Script
// ============================================================

const CONFIG = {
  CONTROL_SHEET: "RSS Config",
  LOG_SHEET:     "Import Log",
  ENABLED_COL:    1,   // A – Enabled checkbox
  URL_COL:        2,   // B – Feed URL
  SHEET_COL:      3,   // C – Destination sheet name
  MAX_ITEMS_COL:  4,   // D – Max items to import (optional)
  STATUS_COL:     5,   // E – Status (written by script)
  HEADER_ROW:     1,
  DATA_START_ROW: 2,
  FEED_TIMEOUT_MS: 30000,  // 30 seconds per feed
};

// ──────────────────────────────────────────────
//  Menu
// ──────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📰 RSS Feeds")
    .addItem("▶ Import All Feeds",    "importAllFeeds")
    .addItem("⚙ Setup Config Sheet", "setupConfigSheet")
    .addItem("🗑 Clear Log",          "clearLog")
    .addToUi();
}

// ──────────────────────────────────────────────
//  Logging helpers
// ──────────────────────────────────────────────
var _logSheet = null;
var _logBuffer = [];

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
  var sheet = getLogSheet(ss);
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var row = [ts, feed || "", level, message];
  sheet.appendRow(row);

  // Colour-code the Level cell on the newly appended row
  var lastRow = sheet.getLastRow();
  var levelCell = sheet.getRange(lastRow, 3);
  var msgCell   = sheet.getRange(lastRow, 4);
  if (level === "ERROR") {
    levelCell.setBackground("#ffcdd2").setFontColor("#b71c1c");
    msgCell.setFontColor("#b71c1c");
  } else if (level === "WARN") {
    levelCell.setBackground("#fff9c4").setFontColor("#f57f17");
  } else if (level === "OK") {
    levelCell.setBackground("#c8e6c9").setFontColor("#1b5e20");
  } else {
    levelCell.setBackground(null).setFontColor("#212121");
  }

  // Keep log trimmed to last 500 rows (+ header)
  if (lastRow > 501) {
    sheet.deleteRow(2);
  }

  SpreadsheetApp.flush(); // write immediately so log is visible in real-time
}

function clearLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.LOG_SHEET);
  if (sheet && sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }
}

// ──────────────────────────────────────────────
//  Create / reset the config sheet
// ──────────────────────────────────────────────
function setupConfigSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.CONTROL_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.CONTROL_SHEET, 0);
  }

  sheet.clearContents();
  sheet.clearFormats();

  var headers = ["Enabled", "Feed URL", "Destination Sheet", "Max Items", "Status"];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#4a90d9")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  var examples = [
    ["https://feeds.bbci.co.uk/news/rss.xml",                     "BBC News", 10, ""],
    ["https://rss.nytimes.com/services/xml/rss/nyt/HomePage.xml", "NYT",      10, ""],
  ];
  sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.URL_COL, examples.length, 4).setValues(examples);

  var cbRange = sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.ENABLED_COL, examples.length, 1);
  cbRange.insertCheckboxes();
  cbRange.setValue(true);
  cbRange.setHorizontalAlignment("center");

  sheet.setColumnWidth(CONFIG.ENABLED_COL,   70);
  sheet.setColumnWidth(CONFIG.URL_COL,       360);
  sheet.setColumnWidth(CONFIG.SHEET_COL,     180);
  sheet.setColumnWidth(CONFIG.MAX_ITEMS_COL, 100);
  sheet.setColumnWidth(CONFIG.STATUS_COL,    260);
  sheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert(
    '✅ Config sheet ready!\n\n' +
    '• Column A – Check/uncheck to enable or disable a feed\n' +
    '• Column B – Paste your RSS/Atom feed URL\n' +
    '• Column C – Name for the destination sheet\n' +
    '• Column D – Max items to import (optional, default 25)\n\n' +
    'Then choose "Import All Feeds" from the RSS Feeds menu.'
  );
}

// ──────────────────────────────────────────────
//  Import all ENABLED feeds
// ──────────────────────────────────────────────
function importAllFeeds() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  _logSheet = null; // reset cached log sheet reference

  var configSheet = ss.getSheetByName(CONFIG.CONTROL_SHEET);
  if (!configSheet) {
    setupConfigSheet();
    return;
  }

  var lastRow = configSheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return;

  var numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  var rows = configSheet.getRange(CONFIG.DATA_START_ROW, 1, numRows, 5).getValues();

  log(ss, null, "INFO", "═══ Import run started ═══");

  rows.forEach(function(row, i) {
    var sheetRow = CONFIG.DATA_START_ROW + i;
    var enabled  = row[CONFIG.ENABLED_COL   - 1];
    var url      = String(row[CONFIG.URL_COL - 1]).trim();
    var destName = String(row[CONFIG.SHEET_COL - 1]).trim() || "Feed " + sheetRow;
    var maxItems = parseInt(row[CONFIG.MAX_ITEMS_COL - 1]) || 25;

    if (!url || url === "") return;

    if (enabled === false || enabled === "FALSE" || enabled === "") {
      log(ss, url, "INFO", "Skipped (disabled)");
      configSheet.getRange(sheetRow, CONFIG.STATUS_COL)
        .setValue("⏸ Disabled")
        .setFontColor("#757575");
      return;
    }

    log(ss, url, "INFO", "Starting import → sheet: \"" + destName + "\", max items: " + maxItems);

    var startTime = new Date().getTime();

    try {
      var result = importFeedWithTimeout(ss, url, destName, maxItems, startTime);

      if (result.timedOut) {
        var msg = "⏱ Timed out after 30s at stage: " + result.stage;
        log(ss, url, "WARN", msg);
        configSheet.getRange(sheetRow, CONFIG.STATUS_COL)
          .setValue(msg)
          .setFontColor("#e65100");
      } else {
        var elapsed = ((new Date().getTime() - startTime) / 1000).toFixed(1);
        var statusMsg = "✅ " + result.count + " items – " + new Date().toLocaleString() + " (" + elapsed + "s)";
        log(ss, url, "OK", "Done – " + result.count + " items in " + elapsed + "s (format: " + result.format + ")");
        configSheet.getRange(sheetRow, CONFIG.STATUS_COL)
          .setValue(statusMsg)
          .setFontColor("#2e7d32");
      }

    } catch (err) {
      var elapsed = ((new Date().getTime() - startTime) / 1000).toFixed(1);
      log(ss, url, "ERROR", err.message + " (after " + elapsed + "s)");
      configSheet.getRange(sheetRow, CONFIG.STATUS_COL)
        .setValue("❌ " + err.message)
        .setFontColor("#c62828");
    }
  });

  log(ss, null, "INFO", "═══ Import run finished ═══");
}

// ──────────────────────────────────────────────
//  Core: fetch + parse with timeout awareness
//  Returns: { count, format, timedOut, stage }
// ──────────────────────────────────────────────
function importFeedWithTimeout(ss, url, destSheetName, maxItems, startTime) {
  var stage = "fetch";

  // ── Timeout check helper ──────────────────────
  function checkTimeout(stageName) {
    stage = stageName;
    if (new Date().getTime() - startTime > CONFIG.FEED_TIMEOUT_MS) {
      return true;
    }
    return false;
  }

  // ── 1. Fetch ──────────────────────────────────
  log(ss, url, "INFO", "[fetch] Sending HTTP request…");
  if (checkTimeout("pre-fetch")) return { timedOut: true, stage: "pre-fetch" };

  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

  if (checkTimeout("post-fetch")) return { timedOut: true, stage: "HTTP fetch" };

  var httpCode = response.getResponseCode();
  log(ss, url, "INFO", "[fetch] HTTP " + httpCode);
  if (httpCode !== 200) throw new Error("HTTP " + httpCode);

  // ── 2. Parse XML ──────────────────────────────
  log(ss, url, "INFO", "[parse] Parsing XML…");
  var xml  = response.getContentText();
  var doc  = XmlService.parse(xml);
  var root = doc.getRootElement();
  var rootName = root.getName().toLowerCase();

  if (checkTimeout("XML parse")) return { timedOut: true, stage: "XML parse" };

  log(ss, url, "INFO", "[parse] Root element: <" + rootName + ">");

  var items  = [];
  var format = "unknown";

  // ── 3. Extract items ──────────────────────────
  if (rootName === "rss") {
    format = "RSS 2.0";
    log(ss, url, "INFO", "[parse] Format detected: RSS 2.0");
    var channel = root.getChild("channel");
    if (!channel) throw new Error("No <channel> element found in RSS feed");

    var rawItems = channel.getChildren("item");
    log(ss, url, "INFO", "[parse] Found " + rawItems.length + " items, importing up to " + maxItems);

    rawItems.slice(0, maxItems).forEach(function(item, idx) {
      if (checkTimeout("item parse #" + idx)) return;
      items.push({
        title:       getText(item, "title"),
        link:        getText(item, "link"),
        published:   getText(item, "pubDate"),
        description: stripHtml(getText(item, "description")),
        author:      getText(item, "author") || getText(item, "creator") || "",
        category:    item.getChildren("category").map(function(c){ return c.getValue(); }).join(", "),
      });
    });

  } else if (rootName === "feed") {
    format = "Atom";
    log(ss, url, "INFO", "[parse] Format detected: Atom");
    var atomNs   = XmlService.getNamespace("http://www.w3.org/2005/Atom");
    var rawEntries = root.getChildren("entry", atomNs);
    log(ss, url, "INFO", "[parse] Found " + rawEntries.length + " entries, importing up to " + maxItems);

    rawEntries.slice(0, maxItems).forEach(function(entry, idx) {
      if (checkTimeout("entry parse #" + idx)) return;
      var linkEl = entry.getChild("link", atomNs);
      var link   = linkEl ? (linkEl.getAttribute("href") ? linkEl.getAttribute("href").getValue() : "") : "";
      items.push({
        title:       getTextNs(entry, "title",     atomNs),
        link:        link,
        published:   getTextNs(entry, "published", atomNs) || getTextNs(entry, "updated", atomNs),
        description: stripHtml(getTextNs(entry, "summary", atomNs) || getTextNs(entry, "content", atomNs)),
        author:      "",
        category:    "",
      });
    });

  } else {
    throw new Error("Unsupported feed format – root element is <" + rootName + ">");
  }

  if (checkTimeout("post-parse")) return { timedOut: true, stage: "post-parse" };
  if (items.length === 0) throw new Error("Feed parsed OK but contained 0 items");

  // ── 4. Write to sheet ─────────────────────────
  log(ss, url, "INFO", "[write] Writing " + items.length + " rows to sheet \"" + destSheetName + "\"…");

  var dest = ss.getSheetByName(destSheetName);
  if (!dest) {
    dest = ss.insertSheet(destSheetName);
  } else {
    dest.clearContents();
  }

  var headers = ["Title", "Link", "Published", "Description", "Author", "Category"];
  dest.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#4a90d9")
    .setFontColor("#ffffff");

  var dataRows = items.map(function(item) {
    return [item.title, item.link, item.published, item.description, item.author, item.category];
  });
  dest.getRange(2, 1, dataRows.length, 6).setValues(dataRows);

  [1, 3, 5, 6].forEach(function(c){ dest.autoResizeColumn(c); });
  dest.setColumnWidth(2, 300);
  dest.setColumnWidth(4, 400);
  dest.setFrozenRows(1);

  if (checkTimeout("sheet write")) return { timedOut: true, stage: "sheet write" };

  log(ss, url, "INFO", "[write] Sheet write complete");

  return { count: items.length, format: format, timedOut: false, stage: "done" };
}

// ──────────────────────────────────────────────
//  Helpers
// ──────────────────────────────────────────────
function getText(el, name) {
  try { return (el.getChild(name) || { getValue: function(){ return ""; } }).getValue(); }
  catch (e) { return ""; }
}

function getTextNs(el, name, ns) {
  try { return (el.getChild(name, ns) || { getValue: function(){ return ""; } }).getValue(); }
  catch (e) { return ""; }
}

function stripHtml(html) {
  if (!html) return "";
  return html
    .replace(/<[^>]+>/g, " ")
    .replace(/&amp;/g,  "&")
    .replace(/&lt;/g,   "<")
    .replace(/&gt;/g,   ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g,  "'")
    .replace(/&nbsp;/g, " ")
    .replace(/\s+/g,    " ")
    .trim();
}
