/**
 * Desk Dashboard - meeting data endpoint.
 *
 * Reads a Google Sheet of recurring weekly meetings and returns the
 * current and next meeting as JSON. Supports JSONP (callback param),
 * which is how the kiosk web page reads it cross-origin.
 *
 * Required Script Properties (Project Settings > Script Properties):
 *   DASH_KEY  - shared secret. Requests without a matching key get nothing.
 * Optional:
 *   SHEET_TAB - tab name. Default: the first tab.
 *   AI_API_KEY - Gemini API key (free at aistudio.google.com/apikey)
 *               enabling AI-written "Where's Mitch?" quips. Without it,
 *               built-in fallback lines are used.
 *   SHEET_ID  - only needed if this script is STANDALONE. A script bound
 *               to the spreadsheet (Extensions > Apps Script from the
 *               sheet) reads its own spreadsheet automatically.
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
  // Bound script: reads the spreadsheet it lives in.
  // Standalone script: falls back to the SHEET_ID property.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    var sheetId = props.getProperty('SHEET_ID');
    if (!sheetId) throw new Error('No spreadsheet: bind the script to the sheet, or set SHEET_ID');
    ss = SpreadsheetApp.openById(sheetId);
  }
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

  var nextDesc = next ? describe_(next, tz) : null;
  var currentDesc = current ? describe_(current, tz) : null;
  return {
    serverNow: now.toISOString(),
    current: currentDesc,
    next: nextDesc,
    // "Where's Mitch?" quip — only needed when he is not in a meeting.
    status: currentDesc ? null : buildStatus_(now, tz, nextDesc),
    // Commute times — null until TOMTOM_KEY + route endpoints are set.
    commute: commuteBlock_()
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


/* ================= "Where's Mitch?" status quips =================
 * When Mitch is not in a meeting, the dashboard shows a one-line,
 * AI-written guess at where he is, based on the situation (time of
 * day, day of week, next meeting). The line is generated by Gemini
 * when the AI_API_KEY script property is set, cached (AI lines for
 * 30 minutes — the page polls every minute, the model is not;
 * fallback lines only 2 minutes, so adding a key takes effect fast),
 * to built-in lines when no key is set or the call fails.
 * Get a free key at aistudio.google.com/apikey and add it as the
 * AI_API_KEY script property. Never put the key in code.
 */
var AI_MODEL_ = 'gemini-2.0-flash';
var QUIP_CACHE_MIN_ = 30;

var SITUATIONS_ = {
  night: {
    hint: 'the middle of the night — asleep, resting, collecting ZZZs',
    fallback: ['Fast asleep — dreaming in 4K',
               'Recharging. Do not disturb the Mitch.',
               'Somewhere under a blanket, collecting ZZZs',
               'Offline until further notice (asleep)']
  },
  early: {
    hint: 'very early morning — still asleep or just barely waking up',
    fallback: ['Probably still asleep — the day has not loaded yet',
               'Either asleep, or heroically making coffee',
               '5 AM club? Unlikely. Asleep.']
  },
  weekend: {
    hint: 'the weekend — family time, kids activities, errands, relaxing at home',
    fallback: ['Weekend mode: family time in progress',
               'Probably at a kid activity somewhere',
               'Off the grid — weekend edition',
               'Running the family circuit, most likely']
  },
  lunch: {
    hint: 'lunchtime — eating lunch somewhere',
    fallback: ['Probably eating lunch!',
               'On a lunch mission somewhere nearby',
               'Refueling — back after lunch']
  },
  evening: {
    hint: 'the end of the workday — heading home, dinner, relaxing, family time',
    fallback: ['Heading home — dinner and family time ahead',
               'Off duty: probably relaxing at home',
               'Commuting to his most important meeting: dinner']
  },
  break: {
    hint: 'a short break between meetings at the office — the break room, coffee, the library, a quick walk',
    fallback: ['Probably in the break room',
               'Out for a quick walk between meetings',
               'Getting coffee, statistically speaking',
               'Library run, or stretching his legs somewhere']
  }
};

function situationBucket_(now, tz) {
  var dayTxt = Utilities.formatDate(now, tz, 'EEE');
  var weekend = (dayTxt === 'Sat' || dayTxt === 'Sun');
  var mins = Number(Utilities.formatDate(now, tz, 'H')) * 60 +
             Number(Utilities.formatDate(now, tz, 'm'));
  if (mins >= 23 * 60 || mins < 5 * 60) return 'night';
  if (mins < 7 * 60) return 'early';
  if (weekend) return 'weekend';
  if (mins >= 11 * 60 + 30 && mins < 13 * 60 + 30) return 'lunch';
  if (mins >= 17 * 60) return 'evening';
  return 'break';
}

function buildStatus_(now, tz, next) {
  var bucket = situationBucket_(now, tz);
  var block = Math.floor(now.getTime() / (QUIP_CACHE_MIN_ * 60000));
  var cache = CacheService.getScriptCache();
  // The bucket and next meeting are in the key, so a situation change
  // refreshes immediately; the TTL does the time-based rotation.
  // AI lines are cached for the full window — but fallback lines only
  // 2 minutes, so adding the AI_API_KEY (or the API recovering) takes
  // effect almost at once instead of waiting out the block.
  var cacheKey = 'quip|' + bucket + '|' + (next ? next.startISO : 'none');
  var quip = cache.get(cacheKey);
  if (!quip) {
    var ai = aiQuip_(now, tz, bucket, next);
    quip = ai || fallbackQuip_(bucket, block);
    cache.put(cacheKey, quip, ai ? QUIP_CACHE_MIN_ * 60 : 120);
  }
  return { bucket: bucket, quip: quip };
}

function fallbackQuip_(bucket, block) {
  var lines = SITUATIONS_[bucket].fallback;
  return lines[Math.abs(block + bucket.length) % lines.length];
}

function aiQuip_(now, tz, bucket, next) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('AI_API_KEY');
  if (!apiKey) return null;
  var when = Utilities.formatDate(now, tz, 'EEEE h:mm a');
  var nextTxt = next
    ? 'His next meeting is ' + next.startLabel + ' ("' + next.title + '").'
    : 'He has no more meetings on the board.';
  var prompt = 'You write one-line status jokes for an office desk dashboard called "Where\'s Mitch?". ' +
    'Mitch is NOT in a meeting right now. It is ' + when + '. ' + nextTxt + ' ' +
    'Situation to riff on: ' + SITUATIONS_[bucket].hint + '. ' +
    'Reply with ONE playful, dry line, at most 12 words, third person, no quotes, no emoji, no hashtags.';
  try {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + AI_MODEL_ +
              ':generateContent?key=' + encodeURIComponent(apiKey);
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 1.1, maxOutputTokens: 40 }
      })
    });
    if (res.getResponseCode() !== 200) return null;
    var data = JSON.parse(res.getContentText());
    var text = (data.candidates[0].content.parts[0].text || '')
      .replace(/\s+/g, ' ').replace(/^["'“”]+|["'“”]+$/g, '').trim();
    if (!text) return null;
    return text.length > 110 ? text.substring(0, 107) + '…' : text;
  } catch (err) {
    return null;
  }
}


/* ================= Commute / traffic =================
 * Live drive times via the TomTom Routing API (free tier, 2,500
 * requests/day — no credit card). Both directions are fetched so the
 * card shows "to work" and "to home" at a glance.
 *
 * Required Script Properties (only if you want the commute card):
 *   TOMTOM_KEY  - free API key from developer.tomtom.com.
 *   COMMUTE_A   - one endpoint as "lat,lon" (e.g. home).
 *   COMMUTE_B   - the other endpoint as "lat,lon" (e.g. work).
 *                 Get coordinates by right-clicking in Google Maps.
 * Optional:
 *   COMMUTE_A_LABEL, COMMUTE_B_LABEL - names for the endpoints
 *                 (defaults: "Home" and "Work").
 *
 * Results are cached 5 minutes: traffic doesn't need fresher data
 * than that, and two calls per fetch is far inside the free tier.
 * No key or no endpoints => commute is null and the card hides.
 */
var COMMUTE_CACHE_SEC_ = 300;

function commuteBlock_() {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('TOMTOM_KEY');
  var a = props.getProperty('COMMUTE_A');
  var b = props.getProperty('COMMUTE_B');
  if (!apiKey || !a || !b) return null;
  var aLabel = props.getProperty('COMMUTE_A_LABEL') || 'Home';
  var bLabel = props.getProperty('COMMUTE_B_LABEL') || 'Work';
  var legs = [
    commuteLeg_(apiKey, a, b, bLabel),
    commuteLeg_(apiKey, b, a, aLabel)
  ];
  var ok = legs.filter(function (l) { return l !== null; });
  if (ok.length === 0) return null;
  return { legs: ok };
}

function commuteLeg_(apiKey, from, to, toLabel) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'commute|' + from + '|' + to;
  var cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);
  var leg = null;
  try {
    var url = 'https://api.tomtom.com/routing/1/calculateRoute/' +
      encodeURIComponent(from) + ':' + encodeURIComponent(to) +
      '/json?key=' + encodeURIComponent(apiKey) + '&traffic=true&travelMode=car';
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      var data = JSON.parse(res.getContentText());
      var summary = data.routes && data.routes[0] && data.routes[0].summary;
      if (summary) {
        var travelMin = Math.round(summary.travelTimeInSeconds / 60);
        var delayMin = Math.round((summary.trafficDelayInSeconds || 0) / 60);
        leg = {
          to: toLabel,
          minutes: travelMin,
          typicalMinutes: Math.max(1, travelMin - delayMin),
          delayMinutes: delayMin,
          level: delayMin <= 2 ? 'light' : (delayMin <= Math.max(5, (travelMin - delayMin) * 0.25) ? 'moderate' : 'heavy')
        };
        cache.put(cacheKey, JSON.stringify(leg), COMMUTE_CACHE_SEC_);
      }
    }
  } catch (err) {
    leg = null;
  }
  return leg;
}
