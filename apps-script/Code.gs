/**
 * Desk Dashboard - meeting data endpoint.
 *
 * Reads a Google Sheet of recurring weekly meetings and returns the
 * current and next meeting as JSON. Supports JSONP (callback param),
 * which is how the kiosk web page reads it cross-origin.
 *
 * Required Script Properties (Project Settings > Script Properties):
 *   DASH_KEY  - shared secret. Requests without a matching key get nothing.
 *   SHEET_ID  - id of the meetings spreadsheet.
 * Optional:
 *   SHEET_TAB - tab name. Default: the first tab.
 *
 * Sheet layout (a header row is fine - rows that don't parse are skipped):
 *   A: Day + start time, e.g. "Fri 10:00"      (day = Mon Tue Wed Thu Fri Sat Sun)
 *   B: Room / location, e.g. "Maple Room"
 *   C: Title, e.g. "Weekly Design Review"
 *   D: End time (optional), e.g. "11:00"       (default duration: 60 minutes)
 *
 * Deploy as a Web app: Execute as "Me", access "Anyone".
 * The deployment URL is unguessable and every request is key-checked,
 * so the endpoint is useless to anyone who finds the URL alone.
 */

var DEFAULT_DURATION_MIN_ = 60;
var DAY_INDEX_ = { sun: 0, mon: 1, tue: 2, wed: 3, thu: 4, fri: 5, sat: 6 };

function doGet(e) {
  var params = (e && e.parameter) || {};
  var props = PropertiesService.getScriptProperties();
  var expected = props.getProperty('DASH_KEY');
  if (!expected || !params.key || !timingSafeEqual_(String(params.key), String(expected))) {
    return respond_(params, { error: 'unauthorized' });
  }
  try {
    return respond_(params, buildPayload_(props));
  } catch (err) {
    return respond_(params, { error: 'server_error' });
  }
}

/** Constant-time-ish string comparison, so the key can't be probed byte by byte. */
function timingSafeEqual_(a, b) {
  if (a.length !== b.length) return false;
  var diff = 0;
  for (var i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function respond_(params, obj) {
  var json = JSON.stringify(obj);
  var cb = params.callback;
  if (cb && /^[A-Za-z_$][\w.$]*$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function buildPayload_(props) {
  var sheetId = props.getProperty('SHEET_ID');
  if (!sheetId) throw new Error('SHEET_ID script property is not set');
  var ss = SpreadsheetApp.openById(sheetId);
  var tabName = props.getProperty('SHEET_TAB');
  var sheet = tabName ? ss.getSheetByName(tabName) : ss.getSheets()[0];
  if (!sheet) throw new Error('Sheet tab not found');
  var tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();

  var rows = sheet.getDataRange().getValues();
  var meetings = [];
  for (var r = 0; r < rows.length; r++) {
    var m = parseRow_(rows[r]);
    if (m) meetings.push(m);
  }

  var now = new Date();
  var occurrences = [];
  meetings.forEach(function (m) {
    for (var w = 0; w <= 1; w++) {
      var start = occurrence_(now, m.day, m.startMin, w);
      occurrences.push({
        title: m.title,
        room: m.room,
        start: start,
        end: new Date(start.getTime() + m.durationMin * 60000)
      });
    }
  });
  occurrences.sort(function (a, b) { return a.start - b.start; });

  var current = null;
  var next = null;
  for (var i = 0; i < occurrences.length; i++) {
    var o = occurrences[i];
    if (now >= o.start && now < o.end) current = o;
    else if (o.start > now && !next) next = o;
  }

  return {
    serverNow: now.toISOString(),
    current: current ? describe_(current, tz) : null,
    next: next ? describe_(next, tz) : null
  };
}

function describe_(o, tz) {
  return {
    title: o.title,
    room: o.room,
    startISO: o.start.toISOString(),
    endISO: o.end.toISOString(),
    startLabel: Utilities.formatDate(o.start, tz, 'EEE h:mm a'),
    endLabel: Utilities.formatDate(o.end, tz, 'h:mm a'),
    dayLabel: Utilities.formatDate(o.start, tz, 'EEEE')
  };
}

/** Parses one sheet row; returns null for headers/blank/unparseable rows. */
function parseRow_(row) {
  var a = (row[0] === null || row[0] === undefined) ? '' : String(row[0]).trim();
  var match = a.match(/^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)[a-z]*\s+(\d{1,2}):(\d{2})\s*(AM|PM)?$/i);
  if (!match) return null;

  var day = DAY_INDEX_[match[1].substring(0, 3).toLowerCase()];
  var startMin = toMinutes_(match[2], match[3], match[4]);

  var duration = DEFAULT_DURATION_MIN_;
  var dCell = (row[3] === null || row[3] === undefined) ? '' : String(row[3]).trim();
  var em = dCell.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)?$/);
  if (em) {
    var endMin = toMinutes_(em[1], em[2], em[3]);
    if (endMin > startMin) duration = endMin - startMin;
  }

  return {
    day: day,
    startMin: startMin,
    durationMin: duration,
    room: (row[1] === null || row[1] === undefined) ? '' : String(row[1]).trim(),
    title: (row[2] === null || row[2] === undefined) ? '' : String(row[2]).trim()
  };
}

function toMinutes_(hh, mm, ampm) {
  var h = parseInt(hh, 10);
  var m = parseInt(mm, 10);
  if (ampm) {
    var ap = String(ampm).toUpperCase();
    if (ap === 'PM' && h < 12) h += 12;
    if (ap === 'AM' && h === 12) h = 0;
  }
  return h * 60 + m;
}

/** The date (at startMin) of weekday `day` in this week's cycle, + `weekOffset` weeks. */
function occurrence_(now, day, startMin, weekOffset) {
  var base = new Date(now.getTime());
  base.setHours(0, 0, 0, 0);
  var deltaDays = ((day - base.getDay() + 7) % 7) + weekOffset * 7;
  var d = new Date(base.getTime() + deltaDays * 86400000);
  d.setHours(Math.floor(startMin / 60), startMin % 60, 0, 0);
  return d;
}
