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
// Bump on every release — the dashboard debug modal shows this so a
// stale endpoint build is immediately visible. Keep in step with
// PAGE_VERSION in index.html.
var CODE_VERSION_ = '2026-10-25k';

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
    commute: commuteBlock_(),
    // Debug metadata for the dashboard's debug modal. Booleans only —
    // no keys or coordinates are ever sent to the page.
    meta: {
      codeVersion: CODE_VERSION_,
      aiConfigured: !!props.getProperty('AI_API_KEY'),
      tomtomConfigured: !!props.getProperty('TOMTOM_KEY')
    }
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
var AI_MODEL_ = 'gemini-3.8-flash';
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
  // v2: entries cached before the prompt/validation fix are orphaned,
  // so a bad cached line (e.g. a model fragment) doesn't linger for
  // the full 30-minute AI TTL.
  var cacheKey = 'quip|v2|' + bucket + '|' + (next ? next.startISO : 'none');
  var cached = cache.get(cacheKey);
  if (cached) {
    try {
      var obj = JSON.parse(cached);
      if (obj && obj.quip) return { bucket: bucket, quip: obj.quip, source: obj.source || 'cached', aiError: obj.aiError || null, aiRaw: obj.aiRaw || null };
    } catch (err) { /* legacy plain-string entry */ }
    return { bucket: bucket, quip: cached, source: 'cached', aiError: null, aiRaw: null };
  }
  var ai = aiQuip_(now, tz, bucket, next);
  var quip, source, aiError = null, aiRaw = null;
  if (ai && ai.quip) { quip = ai.quip; source = 'ai'; }
  else { quip = fallbackQuip_(bucket, block); source = 'fallback'; aiError = (ai && ai.error) || 'unknown'; aiRaw = (ai && ai.aiRaw) || null; }
  cache.put(cacheKey, JSON.stringify({ quip: quip, source: source, aiError: aiError, aiRaw: aiRaw }),
            source === 'ai' ? QUIP_CACHE_MIN_ * 60 : 120);
  return { bucket: bucket, quip: quip, source: source, aiError: aiError, aiRaw: aiRaw };
}

function fallbackQuip_(bucket, block) {
  var lines = SITUATIONS_[bucket].fallback;
  return lines[Math.abs(block + bucket.length) % lines.length];
}

function aiQuip_(now, tz, bucket, next) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('AI_API_KEY');
  if (!apiKey) return { quip: null, error: 'missing_key', aiRaw: null };
  var when = Utilities.formatDate(now, tz, 'EEEE h:mm a');
  var nextTxt = next
    ? 'His next meeting is ' + next.startLabel + ' ("' + next.title + '").'
    : 'He has no more meetings on the board.';
  var prompt = 'You write one-line status jokes for an office desk dashboard called "Where\'s Mitch?". ' +
    'Mitch is NOT in a meeting right now. It is ' + when + '. ' + nextTxt + ' ' +
    'Situation to riff on: ' + SITUATIONS_[bucket].hint + '. ' +
    'Reply with ONE complete, playful, dry sentence, at most 12 words, third person, plain text only: no markdown, no asterisks, no quotes, no emoji, no hashtags. ' +
    'For example, a good reply looks like: \'Probably out walking the dog before dinner.\' (Do not copy this; write a fresh one for the situation above.)';
  var bodyText = null;
  try {
    // AI_MODEL property overrides the built-in default, so a future
    // model retirement is a property edit, not a code change.
    var model = PropertiesService.getScriptProperties().getProperty('AI_MODEL') || AI_MODEL_;
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model +
              ':generateContent?key=' + encodeURIComponent(apiKey);
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        // 300, not 40: this model thinks before answering, and the
        // thinking tokens count against this budget. Observed
        // thoughtsTokenCount 37 on a 40-token budget left nothing for
        // the answer (empty content, MAX_TOKENS) — or a 1–2 word
        // fragment when thinking used slightly less.
        generationConfig: { temperature: 1.1, maxOutputTokens: 300 }
      })
    });
    bodyText = res.getContentText();
    if (res.getResponseCode() !== 200) {
      // Gemini error bodies are JSON with error.message — that's the
      // useful part (e.g. 'API key not valid'). Fall back to raw text.
      var msg = bodyText;
      try { msg = JSON.parse(bodyText).error.message || bodyText; } catch (e) { /* raw text */ }
      return { quip: null, error: 'gemini_' + res.getResponseCode() + ': ' + sanitizeErr_(msg), aiRaw: sanitizeErr_(bodyText, 2000) };
    }
    var data = JSON.parse(bodyText);
    var candidate = (data.candidates && data.candidates[0]) || null;
    var content = (candidate && candidate.content) || null;
    var parts = (content && content.parts) || [];
    if (!parts.length) {
      // A 200 with no usable text is usually a safety block or an
      // empty/degraded reply — surface the reason instead of a bare
      // label, or we're blind again one level down.
      var fb = data.promptFeedback || {};
      var reason = fb.blockReason || (candidate && candidate.finishReason) || 'empty';
      return { quip: null, error: 'no_candidates:' + reason, aiRaw: sanitizeErr_(bodyText, 2000) };
    }
    // Join all parts: some models split a reply across parts, and
    // reading only the first one used to surface bare fragments.
    var raw = parts.map(function (p) { return p.text || ''; }).join(' ');
    var text = raw
      .replace(/\s+/g, ' ')
      .replace(/[*_]/g, '')
      .replace(/^["'“”]+|["'“”]+$/g, '')
      .replace(/\s+/g, ' ').trim();
    if (!text) return { quip: null, error: 'empty_response', aiRaw: sanitizeErr_(bodyText, 2000) };
    // A fragment (e.g. a bare label the model emitted instead of a
    // line) is worse than a fallback — reject it and let the fallback
    // cover until the next generation. Include the rejected text so
    // the debug modal shows what the model actually emitted.
    if (text.split(/\s+/).length < 3) return { quip: null, error: 'short_response: "' + text.slice(0, 60) + '"', aiRaw: sanitizeErr_(bodyText, 2000) };
    return { quip: text.length > 110 ? text.substring(0, 107) + '…' : text, error: null, aiRaw: null };
  } catch (err) {
    return { quip: null, error: 'fetch_failed: ' + sanitizeErr_((err && err.message) || err), aiRaw: bodyText ? sanitizeErr_(bodyText, 2000) : null };
  }
}

// Strip anything URL-shaped (API keys, coordinates) from an error
// message, but keep Google scope names — they say which permission is
// missing. Shared by the commute and AI diagnostics. maxLen caps the
// output (short for error lines, longer for echoed response bodies).
function sanitizeErr_(msg, maxLen) {
  return String(msg)
    .replace(/https:\/\/www\.googleapis\.com\/auth\/([a-z0-9._-]+)/gi, 'scope:$1')
    .replace(/https?:\/\/\S+/g, '[url]')
    .slice(0, maxLen || 160);
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
  // The error field is a debugging aid for the endpoint's JSON:
  // missing_key / missing_endpoints / tomtom_<http code> / no_route /
  // fetch_failed:<reason>. It never contains the key or the coordinates.
  if (!apiKey) return { legs: [], error: 'missing_key' };
  if (!a || !b) return { legs: [], error: 'missing_endpoints' };
  var aLabel = props.getProperty('COMMUTE_A_LABEL') || 'Home';
  var bLabel = props.getProperty('COMMUTE_B_LABEL') || 'Work';
  var legs = [];
  var errors = [];
  var l1 = commuteLeg_(apiKey, a, b, bLabel);
  var l2 = commuteLeg_(apiKey, b, a, aLabel);
  if (l1.leg) legs.push(l1.leg); else if (l1.error) errors.push(l1.error);
  if (l2.leg) legs.push(l2.leg); else if (l2.error) errors.push(l2.error);
  return { legs: legs, error: legs.length ? null : (errors[0] || 'unknown') };
}

// Editor-runnable test: unlike doGet (which returns 'unauthorized'
// before reaching any external call when run without a web request),
// this goes straight to the TomTom call, so running it triggers the
// authorization prompt for the script.external_request scope. Run it,
// accept the prompt, run it again, and the execution log shows the legs.
function testCommute() {
  Logger.log(JSON.stringify(commuteBlock_()));
}

function commuteLeg_(apiKey, from, to, toLabel) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'commute|' + from + '|' + to;
  var cached = cache.get(cacheKey);
  if (cached) return { leg: JSON.parse(cached), error: null };
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
        var leg = {
          to: toLabel,
          minutes: travelMin,
          typicalMinutes: Math.max(1, travelMin - delayMin),
          delayMinutes: delayMin,
          level: delayMin <= 2 ? 'light' : (delayMin <= Math.max(5, (travelMin - delayMin) * 0.25) ? 'moderate' : 'heavy')
        };
        cache.put(cacheKey, JSON.stringify(leg), COMMUTE_CACHE_SEC_);
        return { leg: leg, error: null };
      }
      return { leg: null, error: 'no_route' };
    }
    return { leg: null, error: 'tomtom_' + res.getResponseCode() };
  } catch (err) {
    // Report the failure reason, but never the URL: it carries the API
    // key (query string) and the route coordinates (path).
    return { leg: null, error: 'fetch_failed: ' + sanitizeErr_((err && err.message) || err) };
  }
}
