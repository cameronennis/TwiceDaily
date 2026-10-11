function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Daily Log')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/* =========================================================================
   Sheet layout — one row per date. Everything lives in one "Daily" tab
   except Journal, which stays separate since it's long-form text rather
   than a quick trend-trackable data point. Column indexes are 1-based
   (Apps Script Range convention).

   Every option now writes real data — there's no more generic "Done
   HH:MM" placeholder anywhere. This schema is for brand-new sheets only;
   existing linked sheets from before this restructuring are not migrated.

   SHEET_FIELDS is the full granular reference — every individual data
   point, each in its actual column.
   ========================================================================= */
var SHEET_TABS = {
  Daily: [
    'Date',
    'Sleep Time', 'Wake Time', 'Sleep Duration', 'Wake-Up Energy', 'Day Energy',
    'Mood', 'Caffeine',
    'Weather Condition', 'High Temp',
    'High VX', '2-Rep Bench (lbs)', 'Hiking',
    'Calendar Events'
  ],
  Journal: ['Date', 'Journal']
};

var SHEET_FIELDS = {
  sleepTime:        { tab: 'Daily', col: 2 },  // fed by the "bedtime" half of the "sleep" tile
  wakeTime:         { tab: 'Daily', col: 3 },  // fed by the "wakeup" half of the "sleep" tile
  sleepDuration:    { tab: 'Daily', col: 4 },  // formula, calculated from sleepTime/wakeTime
  wakeUpEnergy:     { tab: 'Daily', col: 5 },  // integer 1-10
  dayEnergy:        { tab: 'Daily', col: 6 },  // integer 1-10

  mood:             { tab: 'Daily', col: 7 },  // integer 1-5
  caffeine:         { tab: 'Daily', col: 8 },  // "Before 2pm", "After 2pm", "Before 2pm, After 2pm", or ""

  weatherCondition: { tab: 'Daily', col: 9 },  // e.g. "Partly cloudy"
  weatherHighTemp:  { tab: 'Daily', col: 10 }, // integer, °F

  climbingGrade:    { tab: 'Daily', col: 11 }, // integer only: VB=-1, V0=0 .. V17=17
  benchMax:         { tab: 'Daily', col: 12 }, // integer, lbs
  hiking:           { tab: 'Daily', col: 13 }, // "Yes" or "No"

  calendarEvents:   { tab: 'Daily', col: 14 }, // "Title (time); Title (time); ..." — see submitCalendarEntry

  journal:          { tab: 'Journal', col: 2 }
};

// Drives serverSaveEntry: which cell each save target lands in. Most keys
// here match a CONFIG tile's id directly (one value per tile — mood,
// caffeine, climbing, gym, hiking, calendar, journal). Two tiles save more
// than one value at once (Sleep -> bedtime+wakeup; Weather ->
// weatherCondition+weatherHighTemp), so those use the granular field name
// as the key instead of the tile id — there is no "sleep" or "weather" key
// here, only their real underlying fields. `section` controls which half
// of the client's dayData ({morning}/{evening}) a loaded entry
// reconstructs into — it matches wherever that option's tile actually
// lives in CONFIG.
var SHEET_COLUMNS = {
  bedtime:          { tab: 'Daily',   col: 2,  section: 'morning' },
  wakeup:           { tab: 'Daily',   col: 3,  section: 'morning' },
  wakeUpEnergy:     { tab: 'Daily',   col: 5,  section: 'morning' },
  dayEnergy:        { tab: 'Daily',   col: 6,  section: 'evening' },
  mood:             { tab: 'Daily',   col: 7,  section: 'evening' },
  caffeine:         { tab: 'Daily',   col: 8,  section: 'evening' },
  weatherCondition: { tab: 'Daily',   col: 9,  section: 'morning' },
  weatherHighTemp:  { tab: 'Daily',   col: 10, section: 'morning' },
  climbing:         { tab: 'Daily',   col: 11, section: 'evening' },
  gym:              { tab: 'Daily',   col: 12, section: 'evening' },
  hiking:           { tab: 'Daily',   col: 13, section: 'evening' },
  calendar:         { tab: 'Daily',   col: 14, section: 'morning' },
  journal:          { tab: 'Journal', col: 2,  section: 'evening' }
};

/* ---------- Per-user spreadsheet selection ----------
   Stored in UserProperties, which Apps Script automatically scopes to
   whichever Google account is running the script — no manual per-user
   isolation needed, unlike the old client-side artifact version. */

function getSpreadsheetInfo() {
  var id = PropertiesService.getUserProperties().getProperty('SPREADSHEET_ID');
  if (!id) return null;
  if (!spreadsheetStillValid_(id)) {
    PropertiesService.getUserProperties().deleteProperty('SPREADSHEET_ID');
    return null;
  }
  var ss = SpreadsheetApp.openById(id);
  return { id: id, url: ss.getUrl() };
}

// SpreadsheetApp.openById() alone isn't enough to detect "the user deleted
// this sheet" — Drive's delete is normally just a move to Trash, and a
// trashed file still opens fine via the API. Check Drive's trashed flag
// too, so a deleted-from-Drive sheet is treated as gone here as well.
function spreadsheetStillValid_(id) {
  try {
    var file = DriveApp.getFileById(id);
    return !file.isTrashed();
  } catch (e) {
    return false; // permanently deleted, or no access at all
  }
}

function createOwnSpreadsheet(todayDateStr) {
  var ss = SpreadsheetApp.create('Daily Log');
  var tabOrder = Object.keys(SHEET_TABS); // ['Daily', 'Journal']
  tabOrder.forEach(function (tab) {
    ensureTab_(ss, tab, SHEET_TABS[tab]);
  });
  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet) ss.deleteSheet(defaultSheet);
  // insertSheet()'s default placement isn't guaranteed to be "at the end" —
  // don't rely on it. Force the left-to-right order explicitly instead.
  tabOrder.forEach(function (tab, i) {
    var sheet = ss.getSheetByName(tab);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(i + 1); // 1-based position
    }
  });

  // Give a brand-new sheet a running start: pre-populate the past 30 days
  // plus today as blank (date-only) rows, in chronological order — in
  // EVERY date-tab, not just Daily, so row N means the same calendar date
  // in all of them from the very first row onward. Creation-only; linking
  // an existing sheet self-heals alignment instead (see
  // ensureDateRowForAllTabs_ below) rather than backfilling history.
  if (todayDateStr) {
    getDateTabNames_().forEach(function (tabName) {
      var sheet = ss.getSheetByName(tabName);
      if (sheet) backfillPastDays_(sheet, todayDateStr, 30);
    });
    var dailySheet = ss.getSheetByName('Daily');
    if (dailySheet) backfillSleepDurationFormulas_(dailySheet); // rows didn't exist yet when ensureTab_ ran above
  }

  PropertiesService.getUserProperties().setProperty('SPREADSHEET_ID', ss.getId());
  return { id: ss.getId(), url: ss.getUrl(), maxSheetDate: getMaxDateForSheet_(ss) };
}

// Every tab whose first column is "Date" — currently Daily and Journal.
// Anything added to SHEET_TABS later with "Date" as its first header
// automatically joins the alignment guarantee below without further code
// changes.
function getDateTabNames_() {
  return Object.keys(SHEET_TABS).filter(function (tab) {
    return SHEET_TABS[tab][0] === 'Date';
  });
}

// One calendar day forward (or back, for a negative delta) from dateStr.
// JS's Date constructor normalizes day-of-month overflow/underflow across
// month and year boundaries on its own (e.g. day 31+1 correctly becomes
// day 1 of the next month), so this needs no special-casing for that.
function shiftDateStr_(dateStr, deltaDays) {
  var p = dateStr.split('-').map(Number);
  var d = new Date(p[0], p[1] - 1, p[2] + deltaDays);
  var y = d.getFullYear();
  var m = String(d.getMonth() + 1).padStart(2, '0');
  var day = String(d.getDate()).padStart(2, '0');
  return y + '-' + m + '-' + day;
}

// Every date from startInclusive through endInclusive, ascending. Plain
// string comparison is valid here since these are zero-padded ISO dates,
// which sort identically to chronological order.
function dateRangeInclusive_(startInclusive, endInclusive) {
  var result = [];
  var cur = startInclusive;
  var guard = 0; // safety net against a malformed date ever looping forever
  while (cur <= endInclusive && guard < 100000) {
    result.push(cur);
    cur = shiftDateStr_(cur, 1);
    guard++;
  }
  return result;
}

// `numDaysBack` days before todayDateStr through todayDateStr itself
// (inclusive), as blank rows, in order. Pure string/date-object
// arithmetic on the date the client says is "today" — deliberately not
// using the script's own timezone (Session.getScriptTimeZone()), which
// could differ from the visitor's and cause an off-by-one-day mismatch.
function backfillPastDays_(sheet, todayDateStr, numDaysBack) {
  var startDate = shiftDateStr_(todayDateStr, -numDaysBack);
  var dates = dateRangeInclusive_(startDate, todayDateStr);
  sheet.getRange(2, 1, dates.length, 1).setValues(dates.map(function (d) { return [d]; }));
}

function getExistingDatesSet_(sheet) {
  var lastRow = sheet.getLastRow();
  var set = {};
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(function (r) {
      if (r[0] !== '') set[String(r[0])] = true;
    });
  }
  return set;
}

// THE alignment guarantee: makes sure dateStr has a row in every date-tab,
// at the same row position in each — filling in any missing calendar days
// along the way (in every date-tab, not just the one that triggered this)
// so row N always means the same date everywhere. Empty rows for skipped
// days are expected and fine.
//
// Self-healing rather than assuming alignment already holds: it computes
// the union of every date across every date-tab, plus dateStr, then tops
// up whichever tab(s) are missing any date from that union. For a sheet
// that's already in sync (the normal case, since this is the only path
// that ever creates a date row), that's just dateStr itself if it's new —
// cheap. For a sheet linked from before this existed, or that drifted out
// of sync some other way, this brings every date-tab back into lockstep
// in one pass rather than leaving it broken.
//
// Performance note: only the rows this call actually inserts get the
// Sleep Duration formula / row-height fix applied — never the tab's whole
// existing history. This runs on every single save, so it has to stay
// O(rows added this time), not O(total rows ever logged), or every save
// would get slower forever as history grows.
//
// Returns the row number for dateStr — identical across every date-tab.
function ensureDateRowForAllTabs_(ss, dateStr) {
  var dateTabs = getDateTabNames_();
  var sheets = {};
  dateTabs.forEach(function (t) { sheets[t] = ensureTab_(ss, t, SHEET_TABS[t]); });

  var unionSet = {};
  dateTabs.forEach(function (t) {
    var existing = getExistingDatesSet_(sheets[t]);
    Object.keys(existing).forEach(function (d) { unionSet[d] = true; });
  });
  unionSet[dateStr] = true;

  var allDates = Object.keys(unionSet).sort();
  var fullRange = dateRangeInclusive_(allDates[0], allDates[allDates.length - 1]);

  var newRowsByTab = {};
  dateTabs.forEach(function (t) {
    var sheet = sheets[t];
    var existing = getExistingDatesSet_(sheet);
    var newRows = [];
    fullRange.forEach(function (d) {
      if (!existing[d]) newRows.push(insertDateOnlyRow_(sheet, d));
    });
    newRowsByTab[t] = newRows;
  });

  var dailySheet = sheets['Daily'];
  if (dailySheet && newRowsByTab['Daily']) {
    newRowsByTab['Daily'].forEach(function (row) { setSleepDurationFormula_(dailySheet, row); });
  }
  dateTabs.forEach(function (t) {
    var newRows = newRowsByTab[t];
    if (newRows.length && tabHasWrapColumn_(SHEET_TABS[t])) {
      newRows.forEach(function (row) { sheets[t].autoResizeRows(row, 1); });
    }
  });

  var refSheet = sheets['Daily'] || sheets[dateTabs[0]];
  var refLastRow = refSheet.getLastRow();
  var refDates = refSheet.getRange(2, 1, refLastRow - 1, 1).getValues().map(function (r) { return String(r[0]); });
  return refDates.indexOf(dateStr) + 2;
}

// Inserts one blank row for dateStr at its correct chronological position
// within this single sheet (or appends at the true end if it's later than
// everything already there). Doesn't touch the Sleep Duration formula
// itself — ensureDateRowForAllTabs_ backfills that once, in bulk, after
// all insertions for a given operation are done, rather than per-row here.
function insertDateOnlyRow_(sheet, dateStr) {
  var lastRow = sheet.getLastRow();
  var insertAt = null;

  if (lastRow >= 2) {
    var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]) > dateStr) { insertAt = i + 2; break; }
    }
  }

  if (insertAt !== null) {
    sheet.insertRowBefore(insertAt);
  } else {
    insertAt = lastRow + 1;
  }

  sheet.getRange(insertAt, 1).setValue(dateStr);
  return insertAt;
}

function linkExistingSpreadsheet(input, todayDateStr) {
  var id = extractSpreadsheetId_(input);
  if (!id) throw new Error('INVALID_LINK');
  var ss;
  try {
    ss = SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error('NO_ACCESS');
  }
  // ensureTab_ is safe to call on a sheet that already has real log data in
  // it (e.g. someone reconnecting the same template from an earlier
  // session) — it only ever writes the header row, and only when that row
  // doesn't already match, so existing rows are never touched.
  Object.keys(SHEET_TABS).forEach(function (tab) {
    ensureTab_(ss, tab, SHEET_TABS[tab]);
  });
  // Self-heals any drift between date-tabs (near-certain for a sheet
  // linked from before row alignment existed) and makes sure today has a
  // row in every date-tab right away, rather than waiting for the first
  // save.
  if (todayDateStr) ensureDateRowForAllTabs_(ss, todayDateStr);

  // One-time whole-tab catch-up for rows that predate either the Sleep
  // Duration formula or column wrapping — ordinary saves only fix rows
  // they just added (see ensureDateRowForAllTabs_), so a genuinely old
  // sheet's existing rows need this explicit pass to catch up. Linking is
  // the right (and only) place for this — it's the one moment a sheet's
  // history might actually predate either feature.
  var dailySheet = ss.getSheetByName('Daily');
  if (dailySheet) backfillSleepDurationFormulas_(dailySheet);
  Object.keys(SHEET_TABS).forEach(function (tab) {
    var sheet = ss.getSheetByName(tab);
    if (sheet) backfillWrapRowHeights_(sheet, SHEET_TABS[tab]);
  });

  PropertiesService.getUserProperties().setProperty('SPREADSHEET_ID', ss.getId());
  return { id: ss.getId(), url: ss.getUrl(), maxSheetDate: getMaxDateForSheet_(ss) };
}

function extractSpreadsheetId_(input) {
  var trimmed = String(input).trim();
  var m = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9-_]{20,}$/.test(trimmed)) return trimmed;
  return null;
}

// Creates the tab if it's missing entirely. If it already exists, only
// touches row 1: writes the header row if it's blank or doesn't already
// match, and never reads, writes, or otherwise disturbs row 2 onward. Safe
// to call repeatedly on the same sheet, including one already full of a
// previous session's log entries.
function ensureTab_(ss, tabName, headers) {
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) sheet = ss.insertSheet(tabName);

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  var existing = headerRange.getValues()[0];
  var matches = headers.every(function (h, i) { return String(existing[i] || '').trim() === h; });
  if (!matches) headerRange.setValues([headers]);

  sheet.getRange('A:A').setNumberFormat('@'); // keep dates as plain text for exact matching — safe, doesn't touch values
  sheet.setFrozenRows(1);
  applyColumnLayout_(sheet, headers);
  applyDateConditionalFormatting_(sheet, headers);

  // NOTE: this used to also re-run backfillSleepDurationFormulas_ (and a
  // whole-column row-height refresh) on every single call — since
  // ensureTab_ runs on every save, that meant rewriting every existing
  // row's formula and re-measuring every existing row's height on every
  // save, scaling with total history instead of what actually changed.
  // That's gone from here now. ensureDateRowForAllTabs_ handles both,
  // targeted only at rows it just inserted. The full whole-tab versions
  // still exist and get called explicitly, once, from
  // linkExistingSpreadsheet — the one place a genuinely stale sheet
  // (predating either feature) actually needs the expensive full pass.

  return sheet;
}

// Deliberate width per column, keyed by header text (not column index) so
// this stays correct automatically if columns get reordered again later —
// no need to re-map anything here when that happens.
//
// autoResizeColumns() turned out unreliable for this: it only sizes to
// whatever's actually IN the column right now, so any field nothing writes
// to yet (Wake-up Energy, Calendar Events, etc.) only had header text to
// go on, and Sheets' auto-fit sizing for that case is inconsistent. Setting
// explicit widths sidesteps that entirely.
var COLUMN_WIDTH_PX = {
  'Date': 100,
  'Sleep Time': 100,
  'Wake Time': 100,
  'Sleep Duration': 110,
  'Wake-Up Energy': 130,
  'Day Energy': 110,
  'Mood': 80,
  'Caffeine': 160,
  'Weather Condition': 140,
  'High Temp': 90,
  'High VX': 90,
  '2-Rep Bench (lbs)': 160,
  'Hiking': 90,
  'Calendar Events': 260,
  'Journal': 800
};
var DEFAULT_COLUMN_WIDTH_PX = 120; // fallback for any future column not listed above

// Only Journal wraps — entries can run several sentences, and wrapping is
// what makes a long entry actually readable in the grid instead of
// spilling off to the right. Row height grows automatically to fit.
// Journal being in its own tab (not mixed into Daily) is what makes that
// safe: a tall row there doesn't affect anything else.
//
// Everything else stays plain fixed-width, no wrap — those live in Daily
// alongside the rest of the trend-tracking data, where consistent, compact
// row heights matter more than seeing a long value in full at a glance.
var WRAP_COLUMNS = ['Journal'];

// setColumnWidth/setWrap are column-level properties — O(1) per column
// regardless of how many rows exist, so these stay cheap to re-run every
// time. What used to also live here — a whole-column autoResizeRows()
// re-measuring every existing row's height — did NOT scale the same way,
// and ran on every single save. That's been pulled out; see
// backfillWrapRowHeights_, called explicitly (once) only where a genuine
// whole-tab catch-up is actually warranted.
function applyColumnLayout_(sheet, headers) {
  headers.forEach(function (header, i) {
    var col = i + 1;
    sheet.setColumnWidth(col, COLUMN_WIDTH_PX[header] || DEFAULT_COLUMN_WIDTH_PX);
    if (WRAP_COLUMNS.indexOf(header) !== -1) {
      var letter = colToLetter_(col);
      sheet.getRange(letter + ':' + letter).setWrap(true);
    }
  });
}

function tabHasWrapColumn_(headers) {
  return headers.some(function (h) { return WRAP_COLUMNS.indexOf(h) !== -1; });
}

// The whole-tab version of the row-height fix — re-measures every
// existing row, so cost scales with total history. Only called from
// linkExistingSpreadsheet's one-time catch-up, not from the regular
// save path (see ensureDateRowForAllTabs_, which fixes only newly
// inserted rows there instead).
function backfillWrapRowHeights_(sheet, headers) {
  if (!tabHasWrapColumn_(headers)) return;
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.autoResizeRows(2, lastRow - 1);
}

// Highlights each row's Date cell by how it compares to today: gray for
// any past day, green for today, and nothing at all for anything else —
// future-dated rows shouldn't exist in the first place (the app never
// creates one ahead of time), so there's deliberately no rule for that
// case. DATEVALUE() converts the plain-text date string to a real date
// for comparison, since column A is deliberately kept as text elsewhere
// (see ensureTab_) for exact string matching — TODAY() and a bare text
// string wouldn't compare reliably otherwise.
//
// Only column A itself gets colored, not the whole row — this runs
// generically for any tab whose first column is "Date" (both Daily and
// Journal qualify), so it doesn't need separate handling per tab.
//
// Replaces (not adds to) this sheet's conditional formatting every time,
// so re-running this repeatedly never piles up duplicate rules. The range
// extends 1000 rows past the header — about 2.7 years of daily entries —
// so new rows are covered automatically without ever needing to reapply
// this when a row gets added.
function applyDateConditionalFormatting_(sheet, headers) {
  if (headers[0] !== 'Date') return;
  var range = sheet.getRange('A2:A1000');

  var pastRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=DATEVALUE($A2)<TODAY()')
    .setBackground('#D9D9D9')
    .setRanges([range])
    .build();

  var todayRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=DATEVALUE($A2)=TODAY()')
    .setBackground('#B7E1CD')
    .setRanges([range])
    .build();

  sheet.setConditionalFormatRules([pastRule, todayRule]);
}

// A1-style column letter for a 1-based column index (e.g. 2 -> "B",
// 27 -> "AA"). Used so the Sleep Duration formula always points at
// whatever columns SHEET_FIELDS.sleepTime/wakeTime currently say — if
// those columns ever move, this keeps the formula correct automatically
// without needing to hand-edit any cell references.
function colToLetter_(col) {
  var letter = '';
  while (col > 0) {
    var rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

// Sets a live formula in one row's Duration cell that computes elapsed
// time between that row's Sleep Time and Wake Time cells, using MOD to
// handle the overnight case correctly (waking up "before" the sleep-time
// value, on the next calendar day). Blank until both inputs are filled in.
function setSleepDurationFormula_(sheet, row) {
  var sleepCol = colToLetter_(SHEET_FIELDS.sleepTime.col);
  var wakeCol = colToLetter_(SHEET_FIELDS.wakeTime.col);
  var durationCol = SHEET_FIELDS.sleepDuration.col;
  var sleepRef = sleepCol + row;
  var wakeRef = wakeCol + row;
  var formula = '=IF(AND(' + sleepRef + '<>"", ' + wakeRef + '<>""), MOD(' + wakeRef + '-' + sleepRef + ', 1), "")';
  var cell = sheet.getRange(row, durationCol);
  cell.setFormula(formula);
  cell.setNumberFormat('[h]"h" mm"m"');
}

function backfillSleepDurationFormulas_(sheet) {
  var lastRow = sheet.getLastRow();
  for (var row = 2; row <= lastRow; row++) {
    setSleepDurationFormula_(sheet, row);
  }
}

function openUserSpreadsheet_() {
  var id = PropertiesService.getUserProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('NO_SPREADSHEET');
  if (!spreadsheetStillValid_(id)) {
    PropertiesService.getUserProperties().deleteProperty('SPREADSHEET_ID');
    throw new Error('NO_SPREADSHEET');
  }
  return SpreadsheetApp.openById(id);
}

/* ---------- Row lookup / write helpers ---------- */

function findRowForDate_(sheet, dateStr) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === dateStr) return i + 2; // +2: header row + 0-index
  }
  return null;
}

/* ---------- Called from the client ---------- */

// Options whose value is a clock time ("HH:MM", straight from an
// <input type="time">) rather than plain text. These need to land in the
// sheet as a genuine time-typed cell, not a text string — Sleep Duration's
// formula does real arithmetic on Sleep Time/Wake Time, and that only
// works reliably against actual time values. Relying on Sheets to
// auto-parse a plain "22:47" string is the kind of locale-dependent
// behavior that's bitten this project before (see the Date-column
// text-format workaround elsewhere in this file) — constructing the Date
// object explicitly sidesteps that entirely.
var TIME_VALUE_OPTIONS = ['bedtime', 'wakeup'];

// Dec 30 1899 is Sheets' own zero-date for time-only values — using it
// explicitly here keeps this independent of any particular date's
// meaning; only the time-of-day component is ever read from these cells.
function parseHHMM_(hhmm) {
  var parts = hhmm.split(':').map(Number);
  return new Date(1899, 11, 30, parts[0], parts[1], 0);
}

function formatTimeValue_(d) {
  var hh = String(d.getHours()).padStart(2, '0');
  var mm = String(d.getMinutes()).padStart(2, '0');
  return hh + ':' + mm;
}

// Reconstructs one day's dayData ({morning:{}, evening:{}}) by reading the
// matching row (if any) from each tab. A non-empty cell means that option
// is "done" for that date; the cell's raw value is carried along too, so
// once real popup content exists it comes through automatically. Takes an
// already-open spreadsheet so callers that need it alongside other data
// (see getInitialData) don't pay for opening it twice.
function loadDayFromSpreadsheet_(ss, dateStr) {
  var dayData = { morning: {}, evening: {} };
  var rowCache = {};

  Object.keys(SHEET_TABS).forEach(function (tab) {
    var sheet = ss.getSheetByName(tab);
    if (sheet) rowCache[tab] = findRowForDate_(sheet, dateStr);
  });

  Object.keys(SHEET_COLUMNS).forEach(function (optionId) {
    var meta = SHEET_COLUMNS[optionId];
    var row = rowCache[meta.tab];
    if (!row) return;
    var sheet = ss.getSheetByName(meta.tab);
    var rawValue = sheet.getRange(row, meta.col).getValue();
    if (rawValue !== '' && rawValue !== null && rawValue !== undefined) {
      // A time-typed cell reads back as a real Date object — format it as
      // "HH:MM" so it drops straight back into an <input type="time">,
      // rather than a Date's default (and much longer) string form.
      var displayValue = (TIME_VALUE_OPTIONS.indexOf(optionId) !== -1 && Object.prototype.toString.call(rawValue) === '[object Date]')
        ? formatTimeValue_(rawValue)
        : String(rawValue);
      dayData[meta.section][optionId] = { done: true, value: displayValue };
    }
  });

  return dayData;
}

// The Daily tab's last row is always its latest date, thanks to the
// chronological-insert/alignment guarantees elsewhere in this file —
// one cheap single-cell read, no need to scan anything.
function getMaxDateForSheet_(ss) {
  var dailySheet = ss.getSheetByName('Daily');
  if (!dailySheet) return null;
  var lastRow = dailySheet.getLastRow();
  if (lastRow < 2) return null;
  var v = dailySheet.getRange(lastRow, 1).getValue();
  return v ? String(v) : null;
}

function serverLoadDay(dateStr) {
  var ss = openUserSpreadsheet_();
  return loadDayFromSpreadsheet_(ss, dateStr);
}

// Startup-only: combines what used to be two separate client round-trips
// (getSpreadsheetInfo + serverLoadDay) into one, opening the spreadsheet
// (and checking Drive trash status) exactly once instead of twice. This
// noticeably cuts the delay between opening the page and the UI populating,
// since each of those was a full network round-trip on its own.
//
// maxSheetDate lets the client's calendar picker allow navigating to any
// date already in the sheet even if the visitor's own clock/timezone
// currently thinks that date is "in the future" — see renderCalendarPopover.
function getInitialData(dateStr) {
  var id = PropertiesService.getUserProperties().getProperty('SPREADSHEET_ID');
  if (!id || !spreadsheetStillValid_(id)) {
    if (id) PropertiesService.getUserProperties().deleteProperty('SPREADSHEET_ID');
    return { spreadsheetUrl: null, dayData: { morning: {}, evening: {} }, maxSheetDate: null };
  }
  var ss = SpreadsheetApp.openById(id);
  return {
    spreadsheetUrl: ss.getUrl(),
    dayData: loadDayFromSpreadsheet_(ss, dateStr),
    maxSheetDate: getMaxDateForSheet_(ss)
  };
}

// Writes a single option's value into its mapped cell, creating (and
// aligning across every date-tab — see ensureDateRowForAllTabs_) the date
// row if it doesn't exist yet.
function serverSaveEntry(dateStr, optionId, value) {
  var meta = SHEET_COLUMNS[optionId];
  if (!meta) throw new Error('Unknown option: ' + optionId);
  var ss = openUserSpreadsheet_();
  var row = ensureDateRowForAllTabs_(ss, dateStr);
  var sheet = ss.getSheetByName(meta.tab); // already ensured to exist above
  var cell = sheet.getRange(row, meta.col);
  if (TIME_VALUE_OPTIONS.indexOf(optionId) !== -1 && value) {
    cell.setValue(parseHHMM_(value));
    cell.setNumberFormat('HH:mm'); // otherwise displays as a full (meaningless) date+time
  } else {
    cell.setValue(value);
  }
  // ensureDateRowForAllTabs_'s row-height backfill (inside ensureTab_,
  // called from there) already ran before this write, so it can't know
  // about the value just written — fit this one row now if the column it
  // landed in wraps.
  var header = SHEET_TABS[meta.tab][meta.col - 1];
  if (WRAP_COLUMNS.indexOf(header) !== -1) sheet.autoResizeRows(row, 1);
  return true;
}

// startIso/endIso are full ISO timestamps for local midnight and end-of-day,
// computed client-side (which knows the visitor's timezone) — the server
// just uses them as-is, avoiding any script-timezone mismatch.
//
// Pulls from every calendar the account owns or is subscribed to (secondary
// calendars, shared/class calendars, the auto-generated Birthdays calendar,
// etc.) — getDefaultCalendar() alone only covers the primary one.
function getTodaysEvents(startIso, endIso) {
  var start = new Date(startIso);
  var end = new Date(endIso);
  var calendars = CalendarApp.getAllCalendars();
  var allEvents = [];

  calendars.forEach(function (cal) {
    var calendarName = cal.getName();
    cal.getEvents(start, end).forEach(function (ev) {
      allEvents.push({
        summary: ev.getTitle(),
        allDay: ev.isAllDayEvent(),
        start: ev.isAllDayEvent()
          ? ev.getAllDayStartDate().toISOString()
          : ev.getStartTime().toISOString(),
        calendarName: calendarName
      });
    });
  });

  // All-day events first, then chronological within each group.
  allEvents.sort(function (a, b) {
    if (a.allDay && !b.allDay) return -1;
    if (!a.allDay && b.allDay) return 1;
    return new Date(a.start) - new Date(b.start);
  });

  return allEvents;
}