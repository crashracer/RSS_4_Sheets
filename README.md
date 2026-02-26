# 📰 RSS Feed Importer for Google Sheets — v3.1

A Google Apps Script that imports RSS 2.0 and Atom feeds into Google Sheets. Manage multiple feeds from a single config sheet, schedule automated imports, export daily CSVs to Drive, and reference any feed's data via named ranges.

---

## Features

| Feature | Details |
|---------|---------|
| Feed formats | RSS 2.0 and Atom |
| Per-feed enable/disable | Checkbox in column A |
| 30s timeout per feed | Hangs caught, logged, and skipped automatically |
| Auto-skip on failure | Disabled after 2 failures in 24h; silently auto-retried after 24h |
| Individual sheets | Each feed writes to its own named tab |
| Combined sheet | All feeds in one "All Feeds" tab, first column = source name |
| Named ranges | Every destination sheet's data registered as a named range |
| Scheduled imports | 15min / 1hr / 6hr / daily / custom — set from the menu |
| Daily CSV export | `Feeds-YYYY-MM-DD.csv` and `Log-YYYY-MM-DD.csv` saved to Drive |
| Real-time Import Log | Colour-coded verbose log written live during every import |

---

## Installation

1. Open your Google Sheet
2. Go to **Extensions → Apps Script**
3. Delete any existing code
4. Paste the entire contents of `rss_importer.gs`
5. Save (`Ctrl+S`)
6. Close the editor and **refresh your Sheet**
7. A **📰 RSS Feeds** menu will appear

---

## Quick Start

1. **📰 RSS Feeds → ⚙ Setup Config Sheet** — creates the RSS Config tab with example feeds
2. Add your feed URLs, destination names, and max item counts
3. **📰 RSS Feeds → ▶ Import All Feeds** — runs all enabled feeds

---

## RSS Config Sheet — Column Reference

| Col | Field | Description |
|-----|-------|-------------|
| A | **Enabled** | Checkbox — checked = import, unchecked = skip |
| B | **Feed URL** | Full RSS or Atom feed URL |
| C | **Destination Sheet** | Name of the tab where data is written (also becomes the named range) |
| D | **Max Items** | Items to import per run (default: 25) |
| E | **Status** | Auto-updated after every import |
| F | **Fail Count** | *(hidden)* Failure counter for auto-skip logic |
| G | **Last Fail** | *(hidden)* Timestamp of most recent failure |

### Status values

| Status | Meaning |
|--------|---------|
| `✅ 10 items — 2/26/2026 (2.3s)` | Import succeeded |
| `⏸ Disabled` | Manually unchecked, skipped |
| `⏱ Timed out after 30s at: HTTP fetch` | Exceeded the 30s per-feed timeout |
| `❌ HTTP 404` | Feed returned an HTTP error |
| `🚫 Auto-skipped (2 failures in 24h) — auto-retries after 24h` | Automatically disabled; will silently re-enable and retry after 24h |
| `🔓 Manually reset – will retry on next import` | Feed was manually re-enabled via the Reset menu item |

---

## Destination Sheets (Individual)

Each feed writes to its own sheet when `INDIVIDUAL_MODE: true` (default). Content is fully replaced on every import.

| Col | Field |
|-----|-------|
| A | Title |
| B | Link |
| C | Published |
| D | Description (HTML stripped) |
| E | Author |
| F | Category |

---

## Combined Sheet ("All Feeds")

When `COMBINED_MODE: true` (default), an **All Feeds** sheet is rebuilt on every import with all feed data in one place. The first column identifies the source.

| Col | Field |
|-----|-------|
| A | **Source** — destination sheet name of the originating feed |
| B | Title |
| C | Link |
| D | Published |
| E | Description |
| F | Author |
| G | Category |

---

## Named Ranges

After every successful import, each destination sheet's data (rows 2 onward) is registered as a **named range**. The name is derived from the sheet name with non-alphanumeric characters replaced by underscores.

| Sheet name | Named range |
|------------|-------------|
| `BBC News` | `BBC_News` |
| `Microsoft News` | `Microsoft_News` |
| `All Feeds` | `All_Feeds` |

Use named ranges in formulas anywhere in your spreadsheet:

```
=COUNTA(BBC_News)
=FILTER(All_Feeds, ISNUMBER(SEARCH("AI", All_Feeds)))
```

View all named ranges: **Data → Named ranges**.

---

## Auto-Skip on Failure

If a feed fails **2 or more times within a rolling 24-hour window**, the script automatically:

1. Unchecks the feed's checkbox (disabling it)
2. Sets the status to `🚫 Auto-skipped`
3. Records the skip timestamp internally

After **24 hours**, on the next import run the feed is **silently re-enabled and retried** — no manual action needed. The failure counter is also reset.

To manually re-enable all auto-skipped feeds immediately:
**📰 RSS Feeds → 🔓 Reset auto-skipped feeds**

---

## Scheduler

Set automatic imports directly from the menu — no Apps Script editor needed.

### Menu options

| Menu Item | Behaviour |
|-----------|-----------|
| Set: Every 15 minutes | Runs `importAllFeeds` every 15 min |
| Set: Every hour | Runs every 60 min |
| Set: Every 6 hours | Runs every 360 min |
| Set: Daily (midnight) | Runs once per day at 00:00 |
| Set: Custom interval… | Prompts for any value between 1–1440 min |
| View active triggers | Lists all active triggers with IDs |
| Delete all triggers | Removes all import triggers |

**Notes:**
- Only one trigger is active at a time — setting a new one removes the previous automatically.
- Triggers run under your Google account and consume your Apps Script daily quota.
- Google's **6-minute total execution limit** applies per run. If you have many slow feeds, increase the interval or reduce Max Items.

### Daily quota reference

| Account type | Max trigger runtime/day |
|--------------|------------------------|
| Free Google account | 90 minutes |
| Google Workspace | 6 hours |

---

## CSV Export to Drive

Exports are saved to the **same Google Drive folder that contains your Sheet**. Each file uses a date-based name and is **overwritten if re-exported on the same day**, so you always have one clean file per day.

### File naming

| Export | Filename |
|--------|----------|
| Combined feed data | `Feeds-YYYY-MM-DD.csv` |
| Import log | `Log-YYYY-MM-DD.csv` |

### Menu items

| Menu Item | What is exported |
|-----------|-----------------|
| Export feeds CSV (today) | The **All Feeds** combined sheet as a single CSV |
| Export log CSV (today) | The **Import Log** sheet as a CSV |

---

## Import Log Sheet

Created automatically on first run. Every step of every import is written in real-time.

| Col | Field |
|-----|-------|
| Timestamp | `yyyy-MM-dd HH:mm:ss` |
| Feed | URL being processed |
| Level | `INFO` / `OK` / `WARN` / `ERROR` |
| Message | Verbose stage detail |

### Log levels

| Level | Colour | When used |
|-------|--------|-----------|
| `INFO` | White | Normal progress step |
| `OK` | Green | Feed imported successfully |
| `WARN` | Yellow | Timeout or auto-skip |
| `ERROR` | Red | Exception or HTTP error |

The log is trimmed to the most recent **500 entries** automatically.
Clear manually via: **📰 RSS Feeds → 🗑 Clear Log**

### Example log output

```
2026-02-26 09:00:00 |                              | INFO  | ═══ Import run started ═══
2026-02-26 09:00:00 | https://feeds.bbci.co.uk/... | INFO  | Starting → "BBC News", max: 10
2026-02-26 09:00:00 | https://feeds.bbci.co.uk/... | INFO  | [fetch] Sending HTTP request…
2026-02-26 09:00:01 | https://feeds.bbci.co.uk/... | INFO  | [fetch] HTTP 200
2026-02-26 09:00:01 | https://feeds.bbci.co.uk/... | INFO  | [parse] Parsing XML…
2026-02-26 09:00:01 | https://feeds.bbci.co.uk/... | INFO  | [parse] Root element: <rss>
2026-02-26 09:00:01 | https://feeds.bbci.co.uk/... | INFO  | [parse] Format: RSS 2.0
2026-02-26 09:00:01 | https://feeds.bbci.co.uk/... | INFO  | [parse] Found 40 items, importing up to 10
2026-02-26 09:00:02 | https://feeds.bbci.co.uk/... | INFO  | [write] Writing 10 rows to "BBC News"…
2026-02-26 09:00:02 | https://feeds.bbci.co.uk/... | INFO  | [write] Named range registered: BBC_News
2026-02-26 09:00:02 | https://feeds.bbci.co.uk/... | INFO  | [write] Complete
2026-02-26 09:00:02 | https://feeds.bbci.co.uk/... | OK    | Done — 10 items in 2.1s (RSS 2.0)
2026-02-26 09:00:02 |                              | INFO  | Combined sheet named range registered: All_Feeds
2026-02-26 09:00:02 |                              | INFO  | ═══ Import run finished ═══
```

---

## Configuration Reference

Edit the `CONFIG` object at the top of `rss_importer.gs`:

```javascript
const CONFIG = {
  CONTROL_SHEET:   "RSS Config",  // Config tab name
  LOG_SHEET:       "Import Log",  // Log tab name
  COMBINED_SHEET:  "All Feeds",   // Combined sheet name
  COMBINED_MODE:   true,          // Write combined sheet?
  INDIVIDUAL_MODE: true,          // Write individual feed sheets?
  FEED_TIMEOUT_MS: 30000,         // Per-feed timeout in ms (default: 30s)
  MAX_FAILS_24H:   2,             // Failures before auto-skip
  FAIL_WINDOW_MS:  86400000       // Failure tracking window in ms (default: 24h)
};
```

---

## Troubleshooting

**`❌ HTTP 403` or `❌ HTTP 401`**
The feed blocks automated requests or requires authentication. Try a different URL or a public mirror of the feed.

**`❌ Unsupported feed format — root is <html>`**
The URL returns a webpage, not a feed. Find and copy the raw RSS/Atom URL (usually ends in `.xml` or `/rss` or `/feed`).

**`⏱ Timed out at: HTTP fetch`**
The feed server is too slow. Check the Import Log for which stage it hung at. Try again later, or reduce Max Items.

**Named range missing**
Named ranges require at least one data row below the header. Make sure the import succeeded and the destination sheet has content.

**"Exceeded maximum execution time" toast**
The 6-minute total limit was hit across all feeds in one run. Reduce the number of enabled feeds, lower Max Items, or switch to a longer trigger interval so fewer feeds need refreshing at once.

**CSV shows today's data was overwritten mid-day**
By design — each daily file is replaced on export. If you need point-in-time snapshots, rename the previous file before re-exporting.

---

## Suggested Improvements

- **Deduplication** — track item GUIDs/links in a hidden sheet and skip already-seen items so sheets grow incrementally rather than being fully replaced each run
- **Keyword filtering** — add a Filter column to the config sheet; only import items whose title or description matches a keyword or regex
- **Email / Slack digest** — after each scheduled import, send a summary with new item counts per feed via Gmail or a Slack webhook
- **Delta imports** — store the last successful import timestamp and pass an `If-Modified-Since` header to avoid re-downloading unchanged feeds
- **Error alerting** — send an email when a feed is auto-skipped or fails consecutively
- **Feed health dashboard** — a summary tab showing uptime %, average item count, and last-success time per feed
- **OPML import** — accept an OPML subscription file to bulk-populate the config sheet from an existing feed reader export
