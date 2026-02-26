# 📰 RSS Feed Importer for Google Sheets

A Google Apps Script that imports RSS 2.0 and Atom feeds directly into Google Sheets. Manage multiple feeds from a single config sheet, enable/disable feeds with a checkbox, and monitor imports in real-time via a live log sheet.

---

## Features

- Supports **RSS 2.0** and **Atom** feed formats
- **Checkbox toggle** per feed to enable or disable imports
- **30-second timeout** per feed — hangs are caught, logged, and skipped automatically
- **Real-time Import Log** sheet with colour-coded status levels
- Each feed writes to its own dedicated destination sheet
- Status column updated after every import with item count and elapsed time

---

## Installation

1. Open your Google Sheet
2. Go to **Extensions → Apps Script**
3. Delete any existing code in the editor
4. Paste the entire contents of `rss_importer.gs`
5. Save with `Ctrl+S` (or `Cmd+S`)
6. Close the Apps Script tab and **refresh your Google Sheet**
7. A **📰 RSS Feeds** menu will appear in the menu bar

---

## Setup

1. Click **📰 RSS Feeds → ⚙ Setup Config Sheet**
   - This creates an **RSS Config** tab with two example feeds pre-filled
2. Edit the config sheet with your own feeds (see column reference below)
3. Click **📰 RSS Feeds → ▶ Import All Feeds** to run

---

## RSS Config Sheet — Column Reference

| Column | Field | Description |
|--------|-------|-------------|
| A | **Enabled** | Checkbox — ✅ checked = import this feed, unchecked = skip |
| B | **Feed URL** | Full URL to the RSS or Atom feed |
| C | **Destination Sheet** | Name of the sheet where feed data will be written |
| D | **Max Items** | Maximum number of items to import (optional, defaults to 25) |
| E | **Status** | Written automatically after each import run |

### Status column values

| Status | Meaning |
|--------|---------|
| `✅ 10 items – 2/26/2026, 9:00:00 AM (2.3s)` | Import succeeded |
| `⏸ Disabled` | Feed was unchecked and skipped |
| `⏱ Timed out after 30s at stage: HTTP fetch` | Feed exceeded the 30s timeout |
| `❌ HTTP 404` | Feed returned an error response |

---

## Destination Sheets — Column Reference

Each feed gets its own sheet. Existing content is cleared and rewritten on every import.

| Column | Field |
|--------|-------|
| A | Title |
| B | Link |
| C | Published date |
| D | Description (HTML stripped) |
| E | Author |
| F | Category |

---

## Import Log Sheet

An **Import Log** tab is created automatically on first run. Every step of every import is written here in real-time so you can watch progress live or diagnose failures after the fact.

| Column | Field |
|--------|-------|
| Timestamp | Date and time to the second |
| Feed | URL being processed |
| Level | `INFO`, `OK`, `WARN`, or `ERROR` |
| Message | Verbose stage detail |

### Log levels

| Level | Colour | Meaning |
|-------|--------|---------|
| `INFO` | White | Normal progress step |
| `OK` | Green | Feed completed successfully |
| `WARN` | Yellow | Feed timed out or was skipped |
| `ERROR` | Red | Feed failed with an error |

### Example log output for a successful import

```
2026-02-26 09:00:01 | https://feeds.bbci.co.uk/... | INFO  | Starting import → sheet: "BBC News", max items: 10
2026-02-26 09:00:01 | https://feeds.bbci.co.uk/... | INFO  | [fetch] Sending HTTP request…
2026-02-26 09:00:02 | https://feeds.bbci.co.uk/... | INFO  | [fetch] HTTP 200
2026-02-26 09:00:02 | https://feeds.bbci.co.uk/... | INFO  | [parse] Parsing XML…
2026-02-26 09:00:02 | https://feeds.bbci.co.uk/... | INFO  | [parse] Root element: <rss>
2026-02-26 09:00:02 | https://feeds.bbci.co.uk/... | INFO  | [parse] Format detected: RSS 2.0
2026-02-26 09:00:02 | https://feeds.bbci.co.uk/... | INFO  | [parse] Found 40 items, importing up to 10
2026-02-26 09:00:03 | https://feeds.bbci.co.uk/... | INFO  | [write] Writing 10 rows to sheet "BBC News"…
2026-02-26 09:00:03 | https://feeds.bbci.co.uk/... | INFO  | [write] Sheet write complete
2026-02-26 09:00:03 | https://feeds.bbci.co.uk/... | OK    | Done – 10 items in 2.1s (format: RSS 2.0)
```

The log is automatically trimmed to the most recent 500 entries. Use **📰 RSS Feeds → 🗑 Clear Log** to wipe it manually.

---

## Menu Reference

| Menu Item | Function |
|-----------|----------|
| ▶ Import All Feeds | Runs import for all enabled feeds |
| ⚙ Setup Config Sheet | Creates or resets the RSS Config tab |
| 🗑 Clear Log | Clears all rows from the Import Log tab |

---

## Scheduled Auto-Import (Optional)

To run imports automatically on a schedule:

1. Open **Extensions → Apps Script**
2. Click the **clock icon** (Triggers) in the left sidebar
3. Click **+ Add Trigger**
4. Set function to `importAllFeeds`
5. Set event source to **Time-driven**
6. Choose your interval (e.g. every hour, every day)
7. Save

---

## Configuration

Advanced settings can be changed at the top of the script in the `CONFIG` object:

```javascript
const CONFIG = {
  CONTROL_SHEET:  "RSS Config",  // Name of the config tab
  LOG_SHEET:      "Import Log",  // Name of the log tab
  FEED_TIMEOUT_MS: 30000,        // Per-feed timeout in milliseconds (default: 30s)
  ...
};
```

---

## Troubleshooting

**Feed shows `❌ HTTP 403` or `❌ HTTP 401`**
The feed requires authentication or blocks automated requests. Try a different feed URL.

**Feed shows `❌ Unsupported feed format`**
The URL may point to an HTML page rather than an actual RSS/Atom feed. Check the URL in a browser and look for the raw XML feed link.

**Feed shows `⏱ Timed out after 30s`**
The feed server is too slow to respond. Try again later, reduce Max Items, or uncheck the feed to disable it. The Log sheet will show exactly which stage it hung at.

**Feed shows `✅` but the destination sheet is empty**
The feed was parsed successfully but contained 0 items. The feed may be empty or temporarily unavailable.

**"Exceeded maximum execution time" toast appears**
Google Apps Script has a 6-minute total execution limit. If you have many feeds, try splitting them across two config sheets and running them separately, or reduce Max Items per feed.