/**
 * Family Calendar — Google Apps Script backend
 *
 * Setup:
 *  1. Create a new Google Sheet. Rename it e.g. "Family Calendar Data".
 *  2. Extensions → Apps Script. Paste this whole file in (replace Code.gs).
 *  3. Deploy → New deployment → Type: Web app
 *       Execute as: Me
 *       Who has access: Anyone
 *     Click Deploy. Copy the Web app URL.
 *  4. Paste that URL into the app's Settings tab → "Sync URL".
 *
 * The script auto-creates two sheets (Events, Bills) on first run with the right headers.
 * Don't rename the columns — the script reads them by header name.
 */

const EVENTS_SHEET = 'Events';
const BILLS_SHEET = 'Bills';
const REMINDERS_SHEET = 'Reminders'; // legacy, kept for migration

const EVENT_HEADERS = ['id', 'kind', 'date', 'time', 'end_date', 'end_time', 'from_location', 'to_location', 'from_tz', 'to_tz', 'travel_method', 'title', 'assignee', 'recurring', 'recur_until', 'exceptions', 'done', 'auto_roll', 'sync_to_ios', 'recipients', 'alarm_minutes', 'notes', 'booking_ref', 'created_at', 'updated_at'];
const BILL_HEADERS  = ['id', 'date', 'title', 'amount', 'assignee', 'recurring', 'recur_until', 'exceptions', 'paid', 'sync_to_ios', 'recipients', 'alarm_minutes', 'notes', 'created_at', 'updated_at'];
// ---- Travel app ----
const TRIPS_SHEET    = 'Trips';
const PACKING_SHEET  = 'Packing';
const EXPENSES_SHEET = 'Expenses';
const FILES_SHEET    = 'Files';
const TRIP_HEADERS    = ['id', 'name', 'destination', 'start_date', 'end_date', 'travellers', 'notes', 'created_at', 'updated_at'];
const PACKING_HEADERS = ['id', 'trip_id', 'item', 'category', 'who', 'packed', 'sort', 'created_at', 'updated_at'];
const EXPENSE_HEADERS = ['id', 'trip_id', 'date', 'title', 'category', 'amount', 'currency', 'paid_by', 'split_with', 'notes', 'created_at', 'updated_at'];
const FILE_HEADERS    = ['id', 'parent_id', 'trip_id', 'name', 'mime', 'size', 'drive_id', 'created_at', 'updated_at'];
const FILES_FOLDER_NAME = 'Family Calendar files';

const REMINDER_HEADERS_LEGACY = ['id', 'date', 'time', 'title', 'assignee', 'recurring', 'auto_roll', 'done', 'notes', 'created_at', 'updated_at'];

/**
 * One-time helper: copies any rows from a legacy Reminders sheet into the Events sheet
 * as events with `done` set, then drops the Reminders sheet. Safe to run multiple times.
 * Run this manually from the Apps Script editor (Run → migrateReminders) once after deploying.
 */
function migrateReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oldSh = ss.getSheetByName(REMINDERS_SHEET);
  if (!oldSh) return 'No Reminders sheet — nothing to migrate.';
  const last = oldSh.getLastRow();
  if (last < 2) {
    ss.deleteSheet(oldSh);
    return 'Empty Reminders sheet removed.';
  }
  const lastCol = oldSh.getLastColumn();
  const header = oldSh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const rows = oldSh.getRange(2, 1, last - 1, lastCol).getValues();
  const eventsSh = getSheet_(EVENTS_SHEET, EVENT_HEADERS);
  let moved = 0;
  rows.forEach(row => {
    if (!row[header.indexOf('id')]) return;
    const item = {};
    REMINDER_HEADERS_LEGACY.forEach(h => {
      const i = header.indexOf(h);
      if (i >= 0) item[h] = row[i];
    });
    if (item.date instanceof Date) item.date = Utilities.formatDate(item.date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    appendRowByHeaders_(eventsSh, item);
    moved++;
  });
  ss.deleteSheet(oldSh);
  return 'Migrated ' + moved + ' reminders into Events. Reminders sheet removed.';
}

// ---------- Web app entry points ----------

// Routing table — which iCal feeds each assignee's items appear in
const ICAL_ROUTING = {
  jess:  ['jess'],
  amus:  ['amus', 'jess'],
  mum:   ['jess', 'mum'],
  dad:   ['jess', 'dad'],
  anny:  ['anny'],
  eliot: ['eliot'],
  house: ['jess', 'amus'],
  milk:  ['jess'],
  bb:    ['jess'],
  cilla: ['jess']
};

function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    if (params.action === 'ics') {
      return buildIcsFeed_(params['for'] || '');
    }
    if (params.action === 'file') {
      return jsonOut(readFile_(params.id));
    }
    return jsonOut(allData_());
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action;
    let result;

    switch (action) {
      case 'add':    result = addRow_(body.kind, body.item); break;
      case 'update': result = updateRow_(body.kind, body.item); break;
      case 'delete': result = deleteRow_(body.kind, body.id); break;
      case 'bulk':   result = bulkApply_(body.changes); break;
      case 'upload': result = uploadFile_(body.item, body.data); break;
      default: throw new Error('Unknown action: ' + action);
    }

    const out = allData_();
    out.result = result;
    return jsonOut(out);
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

// ---------- Sheet helpers ----------

function getSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    return sh;
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    return sh;
  }
  // Existing sheet: ensure all expected headers are present, append missing ones to the right.
  const existing = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
  const missing = headers.filter(h => existing.indexOf(h) === -1);
  if (missing.length) {
    sh.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
    sh.getRange(1, 1, 1, existing.length + missing.length).setFontWeight('bold');
  }
  return sh;
}

// Read by header name regardless of column order, so adding new columns is safe.
function readSheet_(name, headers) {
  const sh = getSheet_(name, headers);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const lastCol = sh.getLastColumn();
  const headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const values = sh.getRange(2, 1, last - 1, lastCol).getValues();
  const tz = Session.getScriptTimeZone();
  return values
    .filter(row => row[headerRow.indexOf('id')]) // must have id
    .map(row => {
      const obj = {};
      headers.forEach(h => {
        const idx = headerRow.indexOf(h);
        let v = idx >= 0 ? row[idx] : '';
        if (v instanceof Date) {
          // Format based on the field's expected type, not blindly as date
          if (h === 'time' || h === 'end_time') {
            v = Utilities.formatDate(v, tz, 'HH:mm');
          } else if (h === 'recur_until' || h === 'date' || /_date$/.test(h)) {
            v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
          } else {
            // created_at, updated_at: ISO timestamp
            v = v.toISOString();
          }
        }
        obj[h] = v;
      });
      return obj;
    });
}

function findRowById_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const lastCol = sh.getLastColumn();
  const headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const idCol = headerRow.indexOf('id');
  if (idCol < 0) return -1;
  const ids = sh.getRange(2, idCol + 1, last - 1, 1).getValues();
  const target = String(id || '').trim();
  for (let i = 0; i < ids.length; i++) {
    const cell = String(ids[i][0] || '').trim();
    if (cell === target) return i + 2;
  }
  return -1;
}

// Write a row by header name, so column order in the sheet doesn't matter.
function writeRow_(sh, rowIdx, item) {
  const lastCol = sh.getLastColumn();
  const headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const out = headerRow.map(h => normaliseCell_(h, item[h]));
  sh.getRange(rowIdx, 1, 1, lastCol).setValues([out]);
}
function appendRowByHeaders_(sh, item) {
  const lastCol = sh.getLastColumn();
  const headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const out = headerRow.map(h => normaliseCell_(h, item[h]));
  sh.appendRow(out);
}

// Force string-type cells for fields that Sheets would otherwise auto-coerce
// (e.g. "17:00" becoming a Date object pinned to 1899-12-30).
function normaliseCell_(header, value) {
  if (value === undefined || value === null) return '';
  // Treat these fields as plain text so Sheets doesn't reinterpret them
  const TEXT_FIELDS = ['id', 'time', 'date', 'recur_until', 'recurring', 'assignee', 'title', 'notes', 'exceptions', 'created_at', 'updated_at',
    'name', 'destination', 'start_date', 'travellers', 'trip_id', 'parent_id', 'item', 'category', 'who', 'currency', 'paid_by', 'split_with', 'mime', 'drive_id'];
  // Booking refs and ticket numbers can be all digits (or start with 0): store them as text
  if (header === 'booking_ref') {
    const s = String(value);
    return /^[0-9+]/.test(s) ? "'" + s : s;
  }
  if (TEXT_FIELDS.indexOf(header) >= 0) return String(value);
  return value;
}

function addRow_(kind, item) {
  const { sh, headers } = sheetFor_(kind);
  if (!item.id) item.id = Utilities.getUuid();
  const now = new Date().toISOString();
  if (!item.created_at) item.created_at = now;
  item.updated_at = now;
  appendRowByHeaders_(sh, item);
  return item;
}

function updateRow_(kind, item) {
  const { sh, headers } = sheetFor_(kind);
  const row = findRowById_(sh, item.id);
  if (row < 0) throw new Error('Not found: ' + item.id);
  item.updated_at = new Date().toISOString();
  if (!item.created_at) {
    const lastCol = sh.getLastColumn();
    const headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    const ci = headerRow.indexOf('created_at');
    if (ci >= 0) {
      const existing = sh.getRange(row, ci + 1).getValue();
      item.created_at = existing || item.updated_at;
    } else {
      item.created_at = item.updated_at;
    }
  }
  writeRow_(sh, row, item);
  return item;
}

function deleteRow_(kind, id) {
  if (kind === 'files') trashDriveFile_(id);
  const { sh } = sheetFor_(kind);
  const row = findRowById_(sh, id);
  if (row < 0) return { id: id, deleted: false };
  sh.deleteRow(row);
  return { id: id, deleted: true };
}

function bulkApply_(changes) {
  const out = [];
  (changes || []).forEach(c => {
    if (c.action === 'add')    out.push(addRow_(c.kind, c.item));
    if (c.action === 'update') out.push(updateRow_(c.kind, c.item));
    if (c.action === 'delete') out.push(deleteRow_(c.kind, c.id));
  });
  return out;
}

function sheetFor_(kind) {
  if (kind === 'events')    return { sh: getSheet_(EVENTS_SHEET, EVENT_HEADERS), headers: EVENT_HEADERS };
  if (kind === 'bills')     return { sh: getSheet_(BILLS_SHEET,  BILL_HEADERS),  headers: BILL_HEADERS  };
  if (kind === 'trips')     return { sh: getSheet_(TRIPS_SHEET,    TRIP_HEADERS),    headers: TRIP_HEADERS };
  if (kind === 'packing')   return { sh: getSheet_(PACKING_SHEET,  PACKING_HEADERS), headers: PACKING_HEADERS };
  if (kind === 'expenses')  return { sh: getSheet_(EXPENSES_SHEET, EXPENSE_HEADERS), headers: EXPENSE_HEADERS };
  if (kind === 'files')     return { sh: getSheet_(FILES_SHEET,    FILE_HEADERS),    headers: FILE_HEADERS };
  throw new Error('Unknown kind: ' + kind);
}

/**
 * Build an iCalendar (.ics) feed for the given person ('jess'/'amus'/'mum'/'dad').
 * Filters items by sync_to_ios flag AND routing table.
 * Returns recurring items as a single VEVENT with RRULE so iOS expands natively.
 */
function buildIcsFeed_(forPerson) {
  forPerson = (forPerson || '').toLowerCase();
  if (!ICAL_ROUTING[forPerson] && forPerson !== '') {
    return ContentService.createTextOutput('Invalid recipient. Use ?action=ics&for=jess|amus|mum|dad|anny|eliot')
      .setMimeType(ContentService.MimeType.TEXT);
  }

  const events = readSheet_(EVENTS_SHEET, EVENT_HEADERS);
  const bills  = readSheet_(BILLS_SHEET,  BILL_HEADERS);

  const lines = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Family Calendar//EN',
    'CALSCALE:GREGORIAN',
    'METHOD:PUBLISH',
    'X-WR-CALNAME:Family Calendar (' + forPerson + ')',
    'X-WR-TIMEZONE:Europe/London'
  ];

  const tz = Session.getScriptTimeZone() || 'Europe/London';
  const stamp = Utilities.formatDate(new Date(), 'UTC', "yyyyMMdd'T'HHmmss'Z'");

  function shouldInclude(item) {
    if (!truthyCell_(item.sync_to_ios)) return false;
    if (forPerson === '') return true; // unrestricted

    // Preferred: explicit recipients list (comma-separated person IDs)
    const recipientsRaw = String(item.recipients || '').trim();
    if (recipientsRaw) {
      const recipients = recipientsRaw.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
      return recipients.indexOf(forPerson) >= 0;
    }

    // Back-compat: items created before recipients existed → use legacy routing table on assignee
    if (!item.assignee) return false;
    const assigneeIds = String(item.assignee).split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
    for (var i = 0; i < assigneeIds.length; i++) {
      const targets = ICAL_ROUTING[assigneeIds[i]] || [];
      if (targets.indexOf(forPerson) >= 0) return true;
    }
    return false;
  }

  events.forEach(ev => {
    if (!shouldInclude(ev)) return;
    if (!ev.date) return;
    pushEventLines_(lines, ev, 'event', stamp, tz);
  });

  bills.forEach(b => {
    if (!shouldInclude(b)) return;
    if (!b.date) return;
    pushEventLines_(lines, b, 'bill', stamp, tz);
  });

  lines.push('END:VCALENDAR');

  // iCalendar requires CRLF line endings
  const ics = lines.map(l => foldIcsLine_(l)).join('\r\n') + '\r\n';

  return ContentService
    .createTextOutput(ics)
    .setMimeType(ContentService.MimeType.ICAL);
}

function pushEventLines_(lines, item, kind, stamp, tz) {
  const isBill = kind === 'bill';
  const isTravel = !isBill && (item.kind === 'travel');

  let summary;
  if (isBill) {
    summary = '£' + (Number(item.amount) || 0).toFixed(2) + ' — ' + (item.title || 'Bill');
  } else if (isTravel) {
    const methodPrefix = travelMethodPrefix_(item.travel_method);
    const baseTitle = item.title || (item.from_location && item.to_location
      ? item.from_location + ' → ' + item.to_location
      : 'Travel');
    summary = methodPrefix + baseTitle;
  } else {
    summary = item.title || 'Event';
  }

  let description = (item.notes || '');
  if (!isBill && item.assignee) {
    const niceAssignees = String(item.assignee).split(',').map(s => s.trim()).filter(Boolean).join(', ');
    description += '\nFor: ' + niceAssignees;
  }
  if (isTravel) {
    if (item.from_location) description += '\nFrom: ' + item.from_location;
    if (item.to_location)   description += '\nTo: '   + item.to_location;
  }

  // Bills are all-day. Events use time if present, otherwise all-day.
  // Be defensive: time may arrive as 'HH:mm', or as a Date object (Sheets time cell),
  // or as a junk string like '1899-12-30...' if Sheets serialised oddly.
  const normalisedTime = normaliseTime_(item.time);
  const normalisedEnd  = normaliseTime_(item.end_time);
  const hasTime = !isBill && !!normalisedTime;

  let dtStart, dtEnd, allDay;
  // Multi-day entries (not travel, not bills): end_date later than date
  const multiDayEnd = (!isBill && !isTravel && item.end_date && String(item.end_date) > String(item.date))
    ? String(item.end_date) : '';
  let location = '';

  if (isTravel && hasTime) {
    // Travel: use per-end timezones; fall back to script tz if not specified
    const fromTz = String(item.from_tz || tz).trim() || tz;
    const toTz   = String(item.to_tz   || fromTz).trim() || fromTz;

    const [hh, mm] = normalisedTime.split(':').map(Number);
    const startD = parseIsoDate_(item.date);
    startD.setHours(hh, mm, 0, 0);

    const endDateStr = item.end_date || item.date;
    let endD;
    if (normalisedEnd) {
      const [eh, em] = normalisedEnd.split(':').map(Number);
      endD = parseIsoDate_(endDateStr);
      endD.setHours(eh, em, 0, 0);
    } else {
      endD = new Date(startD.getTime() + 60 * 60 * 1000);
    }

    dtStart = 'DTSTART;TZID=' + fromTz + ':' + formatIcsDateTime_(startD, fromTz);
    dtEnd   = 'DTEND;TZID='   + toTz   + ':' + formatIcsDateTime_(endD,   toTz);
    allDay = false;
    location = (item.from_location || '') + (item.to_location ? ' → ' + item.to_location : '');
  } else if (hasTime) {
    const [hh, mm] = normalisedTime.split(':').map(Number);
    const startD = parseIsoDate_(item.date);
    startD.setHours(hh, mm, 0, 0);
    let endD;
    if (normalisedEnd) {
      const [eh, em] = normalisedEnd.split(':').map(Number);
      endD = parseIsoDate_(multiDayEnd || item.date);
      endD.setHours(eh, em, 0, 0);
      if (endD.getTime() <= startD.getTime()) {
        endD.setDate(endD.getDate() + 1);
      }
    } else if (multiDayEnd) {
      // Multi-day with a start time but no end time: run to the end of the last day
      endD = parseIsoDate_(multiDayEnd);
      endD.setHours(23, 59, 0, 0);
    } else {
      endD = new Date(startD.getTime() + 60 * 60 * 1000);
    }
    dtStart = 'DTSTART;TZID=' + tz + ':' + formatIcsDateTime_(startD, tz);
    dtEnd   = 'DTEND;TZID='   + tz + ':' + formatIcsDateTime_(endD,   tz);
    allDay = false;
  } else {
    const startD = parseIsoDate_(item.date);
    // All-day DTEND is exclusive: the day after the last day shown
    const endD   = parseIsoDate_(multiDayEnd || item.date);
    endD.setDate(endD.getDate() + 1);
    dtStart = 'DTSTART;VALUE=DATE:' + formatIcsDate_(startD);
    dtEnd   = 'DTEND;VALUE=DATE:'   + formatIcsDate_(endD);
    allDay = true;
  }

  lines.push('BEGIN:VEVENT');
  lines.push('UID:' + item.id + '@family-calendar');
  lines.push('DTSTAMP:' + stamp);
  lines.push(dtStart);
  lines.push(dtEnd);
  lines.push('SUMMARY:' + escapeIcsText_(summary));
  if (location) lines.push('LOCATION:' + escapeIcsText_(location));
  if (description.trim()) lines.push('DESCRIPTION:' + escapeIcsText_(description.trim()));

  // Recurrence
  if (item.recurring && item.recurring !== 'none' && item.recurring !== '') {
    const rrule = buildRrule_(item.recurring, item.recur_until, allDay);
    if (rrule) lines.push(rrule);
    // Exceptions
    let exceptions = [];
    try {
      exceptions = item.exceptions ? (typeof item.exceptions === 'string' ? JSON.parse(item.exceptions) : item.exceptions) : [];
    } catch (err) { exceptions = []; }
    exceptions.forEach(d => {
      if (allDay) {
        lines.push('EXDATE;VALUE=DATE:' + d.replace(/-/g, ''));
      } else {
        // Same time as DTSTART
        const [hh, mm] = normalisedTime.split(':').map(Number);
        const exD = parseIsoDate_(d);
        exD.setHours(hh, mm, 0, 0);
        lines.push('EXDATE;TZID=' + tz + ':' + formatIcsDateTime_(exD, tz));
      }
    });
  }

  // Alarm
  // Any number of alarms: one value ("15") or several joined by "|" ("1440|60")
  const seenTriggers = {};
  String(item.alarm_minutes === null || item.alarm_minutes === undefined ? '' : item.alarm_minutes)
    .split(/[|,;\s]+/)
    .filter(s => s !== '' && !isNaN(Number(s)))
    .forEach(s => {
      const min = Number(s);
      const trigger = allDay ? formatAllDayTrigger_(min) : formatTrigger_(min);
      if (seenTriggers[trigger]) return;
      seenTriggers[trigger] = true;
      lines.push('BEGIN:VALARM');
      lines.push('ACTION:DISPLAY');
      lines.push('DESCRIPTION:' + escapeIcsText_(summary));
      lines.push('TRIGGER:' + trigger);
      lines.push('END:VALARM');
    });

  lines.push('END:VEVENT');
}

function buildRrule_(recurring, until, allDay) {
  let freq = '';
  let interval = 1;
  switch (recurring) {
    case 'daily':       freq = 'DAILY'; break;
    case 'weekly':      freq = 'WEEKLY'; break;
    case 'fortnightly': freq = 'WEEKLY'; interval = 2; break;
    case 'four-weekly': freq = 'WEEKLY'; interval = 4; break;
    case 'monthly':     freq = 'MONTHLY'; break;
    case 'quarterly':   freq = 'MONTHLY'; interval = 3; break;
    case 'yearly':      freq = 'YEARLY'; break;
    default: return '';
  }
  let r = 'RRULE:FREQ=' + freq;
  if (interval > 1) r += ';INTERVAL=' + interval;
  if (until) {
    // UNTIL must be the same kind of value as DTSTART: a plain date for
    // all-day items, a UTC date-time for timed ones.
    const u = parseIsoDate_(until);
    r += ';UNTIL=' + formatIcsDate_(u) + (allDay ? '' : 'T235959Z');
  }
  return r;
}

// All-day items (including bills) start at midnight, so a plain "1 day before"
// would ring at 00:00. Ring at 09:00 instead: on the day itself for alarms
// shorter than a day, otherwise 09:00 that many days before.
const ALL_DAY_ALARM_HOUR = 9;
function formatAllDayTrigger_(minBefore) {
  const days = Math.floor(minBefore / 1440);
  if (days <= 0) return 'PT' + ALL_DAY_ALARM_HOUR + 'H';
  const hoursBack = 24 - ALL_DAY_ALARM_HOUR;
  return '-P' + (days > 1 ? (days - 1) + 'D' : '') + 'T' + hoursBack + 'H';
}

function formatTrigger_(minBefore) {
  // VALARM TRIGGER as a duration before event start. 0 = at time of event.
  if (minBefore === 0) return 'PT0M';
  if (minBefore < 60) return '-PT' + minBefore + 'M';
  if (minBefore < 1440) {
    const h = Math.floor(minBefore / 60);
    const m = minBefore % 60;
    return '-PT' + h + 'H' + (m ? m + 'M' : '');
  }
  const days = Math.floor(minBefore / 1440);
  const rem = minBefore % 1440;
  let s = '-P' + days + 'D';
  if (rem) s = '-P' + days + 'DT' + Math.floor(rem / 60) + 'H';
  return s;
}

function parseIsoDate_(s) {
  // 'YYYY-MM-DD' → local Date at midnight
  if (s instanceof Date) return new Date(s.getTime());
  const parts = String(s).split('-').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2], 0, 0, 0, 0);
}

/**
 * Normalise whatever Sheets gives us back for a time field into 'HH:mm' or ''.
 * Handles:
 *   - 'HH:mm' or 'HH:mm:ss' strings (the happy path)
 *   - Date objects (Sheets time-formatted cells return Date with epoch 1899-12-30)
 *   - Junk strings like '1899-12-30T08:50:00.000Z' (Date.toISOString() output)
 *   - Empty / null
 */
function normaliseTime_(v) {
  if (v == null || v === '') return '';
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  }
  const s = String(v).trim();
  if (!s) return '';
  // Plain HH:mm or HH:mm:ss
  const m1 = s.match(/^(\d{1,2}):(\d{2})/);
  if (m1) return pad2_(m1[1]) + ':' + m1[2];
  // ISO timestamp like 1899-12-30T08:50:00.000Z — extract the time portion
  const m2 = s.match(/T(\d{2}):(\d{2})/);
  if (m2) return m2[1] + ':' + m2[2];
  return '';
}

function pad2_(n) {
  const s = String(n);
  return s.length < 2 ? '0' + s : s;
}

function travelMethodPrefix_(method) {
  switch (String(method || '').toLowerCase()) {
    case 'plane': return '✈ ';
    case 'train': return '🚆 ';
    case 'drive': case 'car': return '🚗 ';
    case 'bus':   return '🚌 ';
    case 'ferry': case 'boat': return '⛴ ';
    default: return '';
  }
}

function formatIcsDate_(d) {
  // YYYYMMDD
  const pad = n => (n < 10 ? '0' + n : '' + n);
  return d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate());
}

function formatIcsDateTime_(d, tz) {
  // YYYYMMDDTHHMMSS in script timezone
  return Utilities.formatDate(d, tz, "yyyyMMdd'T'HHmmss");
}

function escapeIcsText_(s) {
  return String(s)
    .replace(/\\/g, '\\\\')
    .replace(/\n/g, '\\n')
    .replace(/,/g, '\\,')
    .replace(/;/g, '\\;');
}

function foldIcsLine_(line) {
  // RFC 5545: lines >75 octets must be folded with CRLF + space
  if (line.length <= 75) return line;
  let out = '';
  let i = 0;
  while (i < line.length) {
    if (i === 0) {
      out += line.substring(0, 75);
      i = 75;
    } else {
      out += '\r\n ' + line.substring(i, i + 74);
      i += 74;
    }
  }
  return out;
}

function truthyCell_(v) {
  if (v === true) return true;
  if (v === false) return false;
  if (v == null || v === '') return false;
  if (typeof v === 'string') {
    const s = v.toLowerCase();
    return s === 'true' || s === '1' || s === 'yes';
  }
  return !!v;
}

// Everything both apps need, in one response
function allData_() {
  return {
    ok: true,
    events:   readSheet_(EVENTS_SHEET, EVENT_HEADERS),
    bills:    readSheet_(BILLS_SHEET, BILL_HEADERS),
    trips:    readSheet_(TRIPS_SHEET, TRIP_HEADERS),
    packing:  readSheet_(PACKING_SHEET, PACKING_HEADERS),
    expenses: readSheet_(EXPENSES_SHEET, EXPENSE_HEADERS),
    files:    readSheet_(FILES_SHEET, FILE_HEADERS),
    serverTime: new Date().toISOString()
  };
}

// ---- Files (tickets, QR codes, PDFs) ----
// Stored privately in the script owner's Google Drive. The apps fetch them back
// through this script, so files are never shared by public link.
function filesFolder_() {
  const it = DriveApp.getFoldersByName(FILES_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FILES_FOLDER_NAME);
}

// Run this once from the Apps Script editor (Run ▶) to grant Google Drive access.
function authoriseDrive() {
  return 'Drive folder ready: ' + filesFolder_().getName();
}

function uploadFile_(meta, base64) {
  if (!meta || !base64) throw new Error('Missing file');
  const bytes = Utilities.base64Decode(base64);
  if (bytes.length > 10 * 1024 * 1024) throw new Error('File too large (max 10 MB)');
  const name = String(meta.name || 'file');
  const mime = String(meta.mime || 'application/octet-stream');
  const file = filesFolder_().createFile(Utilities.newBlob(bytes, mime, name));
  return addRow_('files', {
    id: meta.id || Utilities.getUuid(),
    parent_id: meta.parent_id || '',
    trip_id: meta.trip_id || '',
    name: name,
    mime: mime,
    size: bytes.length,
    drive_id: file.getId()
  });
}

function readFile_(id) {
  const rows = readSheet_(FILES_SHEET, FILE_HEADERS).filter(r => String(r.id) === String(id));
  if (!rows.length) return { ok: false, error: 'File not found' };
  const blob = DriveApp.getFileById(rows[0].drive_id).getBlob();
  return { ok: true, name: rows[0].name, mime: rows[0].mime || blob.getContentType(),
           data: Utilities.base64Encode(blob.getBytes()) };
}

function trashDriveFile_(id) {
  try {
    const rows = readSheet_(FILES_SHEET, FILE_HEADERS).filter(r => String(r.id) === String(id));
    if (rows.length && rows[0].drive_id) DriveApp.getFileById(rows[0].drive_id).setTrashed(true);
  } catch (err) { /* already gone: still remove the row */ }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
