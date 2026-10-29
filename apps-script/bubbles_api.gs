/**
 * Bubbles API – lets the Bubbles app read your weigh-ins, gym sessions and steps from this sheet.
 *
 * SETUP (once):
 *  1. In your Hevy sheet: Extensions > Apps Script. Click + next to "Files" > Script, name it bubbles_api,
 *     and paste all of this in. Save.
 *     (If you get "doGet already declared", the project already has a doGet – tell Claude.)
 *  2. Project Settings (gear icon) > Script properties > Add property:
 *       name  BUBBLES_TOKEN     value  a long random string you make up (30+ characters)
 *  3. Deploy > New deployment > type "Web app" > Execute as: Me > Who has access: Anyone > Deploy.
 *     Approve the permissions, then copy the Web app URL (it ends in /exec).
 *  4. In the Bubbles app: Habits > Settings, and paste into each habit's "Sync URL":
 *       Weigh in        <Web app URL>?what=weight&token=<your token>
 *       Go to the gym   <Web app URL>?what=gym&token=<your token>
 *       10k steps       <Web app URL>?what=steps&token=<your token>
 *
 * Other links (open them in a browser):
 *       ?what=exercises&token=…                  every exercise and which of Push / Pull / Legs it counts as
 *       ?what=steps_set&date=2026-10-28&steps=8123&token=…     what the iPhone Shortcut calls to save steps
 *       ?what=calendar&token=…                   the next 30 days of the family calendar (add &days=14 for fewer, &back=6 to include the last 6 days)
 *
 * CALENDAR SETUP (once): Project Settings > Script properties > Add property:
 *       name  FAMILY_CAL_URL    value  the calendar link from Family Daily (the https://app.familydaily.app/api/calendar/... one)
 *   Keep that link private: it shows your whole family calendar to anyone who has it. It lives only here,
 *   never in the app or on GitHub. Then Deploy > Manage deployments > pencil > New version > Deploy.
 *   The app picks the calendar up by itself: it reuses the address from your habit Sync URLs.
 *
 * After you edit this script later: Deploy > Manage deployments > pencil > Version: New version > Deploy.
 */

const BUBBLES_SHEET_ID = '1HcqxvQ08C5kr8PLAhzXowsy8WsC9nPwI7Vrz722LIGQ';   // your Hevy sheet
const BUBBLES_DAYS = 1100;       // how far back to report (about 3 years, for habit history)
const BUBBLES_STEP_GOAL = 8000;  // steps needed for the steps habit to tick (iPhone-only count runs ~10% under Health)

// How exercises are sorted into Push / Pull / Legs. Matching looks for these words inside the
// exercise name (lower-case). Lists are checked in this order and the FIRST match wins, so
// "Leg Curl" is Legs (checked first) and not Pull. 'skip' = warm-ups, core and mobility: not counted.
// Edit freely. Open  ?what=exercises  to see how each of your exercises is currently sorted.
const BUBBLES_PPL = [
  ['skip', ['warm up', 'warmup', 'stretch', 'walk', 'running', 'treadmill', 'cardio', 'mobility', 'cat cow', 'dead bug', 'plank', 'clamshell', 'crunch', 'sit up', 'sit-up']],
  ['legs', ['squat', 'leg press', 'leg extension', 'leg curl', 'lunge', 'calf', 'hip thrust', 'glute', 'romanian', 'rdl', 'stiff leg', 'hack', 'step up', 'step-up', 'adductor', 'abductor', 'good morning']],
  ['pull', ['pull up', 'pull-up', 'pullup', 'chin up', 'chin-up', 'chinup', 'pulldown', 'pull down', 'row', 'curl', 'face pull', 'shrug', 'deadlift', 'reverse fly', 'rear delt', 'hammer', 'pullover', 'back extension']],
  ['push', ['bench', 'press', 'push up', 'push-up', 'pushup', 'fly', 'flye', 'pec', 'crossover', 'chest', 'dip', 'tricep', 'pushdown', 'push down', 'skull', 'lateral raise', 'front raise', 'shoulder', 'overhead', 'arnold']],
];
const BUBBLES_PPL_MIN_SHARE = 0.6;   // a session is labelled Push/Pull/Legs if that share of its sets is one type; otherwise "Mixed"

function doGet(e) {
  const p = (e && e.parameter) || {};
  const secret = PropertiesService.getScriptProperties().getProperty('BUBBLES_TOKEN');
  if (!secret || p.token !== secret) return bubblesJson_({ error: 'unauthorised' });
  try {
    // The calendar does not need the sheet, so it is answered first (a sheet problem can't break it)
    if (p.what === 'calendar') return bubblesJson_(bubblesCalendar_(p));
    const ss = SpreadsheetApp.openById(BUBBLES_SHEET_ID);
    if (p.what === 'weight')    return bubblesJson_(bubblesWeight_(ss));
    if (p.what === 'gym')       return bubblesJson_(bubblesGym_(ss));
    if (p.what === 'steps')     return bubblesJson_(bubblesSteps_(ss));
    if (p.what === 'steps_set') return bubblesJson_(bubblesStepsSet_(ss, p));
    if (p.what === 'exercises') return bubblesJson_(bubblesExercises_(ss));
    return bubblesJson_({ error: 'unknown request' });
  } catch (err) {
    return bubblesJson_({ error: String(err) });
  }
}

function bubblesJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Turn a sheet cell (a real date, or text like 28/10/2026) into "2026-10-28"
function bubblesIso_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  return null;
}

function bubblesCutoff_(tz) {
  return Utilities.formatDate(new Date(Date.now() - BUBBLES_DAYS * 86400000), tz, 'yyyy-MM-dd');
}

const bubblesR1_ = n => Math.round(n * 10) / 10;

/* ---------------------------------------------------------------- weight */

// Find a column in a sheet's top row by a test on the header text (returns the 1-based column, or 0)
function bubblesCol_(sh, test) {
  const head = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
  for (let i = 0; i < head.length; i++) {
    if (test(String(head[i]).trim().toLowerCase(), i, head)) return i + 1;
  }
  return 0;
}

// Read a date column and a number column from one sheet into out[date] = number
function bubblesReadPair_(sh, dateCol, valCol, tz, out) {
  const n = sh.getLastRow() - 1;
  if (n < 1 || !dateCol || !valCol) return;
  const d = sh.getRange(2, dateCol, n, 1).getValues();
  const v = sh.getRange(2, valCol, n, 1).getValues();
  for (let i = 0; i < n; i++) {
    const iso = bubblesIso_(d[i][0], tz), x = Number(v[i][0]);
    if (iso && x > 0) out[iso] = bubblesR1_(x);
  }
}

// Trend weight for EVERY day, as date -> kg.
// Older history comes from the "Weight" tab (Date + Trend Weight); the newer "Calc" tab
// (the Date next to "Weigh-in", and its "Trend") then overrides it where both have a value.
function bubblesTrendSeries_(ss) {
  const tz = ss.getSpreadsheetTimeZone(), out = {};
  const w = ss.getSheetByName('Weight');
  if (w) bubblesReadPair_(w, bubblesCol_(w, h => h === 'date'), bubblesCol_(w, h => h === 'trend weight'), tz, out);
  const c = ss.getSheetByName('Calc');
  if (c) {
    const weighCol = bubblesCol_(c, h => h === 'weigh-in');
    bubblesReadPair_(c, weighCol ? weighCol - 1 : 0, bubblesCol_(c, h => h === 'trend'), tz, out);
  }
  return out;
}

// The latest trend number on the Dashboard tab (the heading "Trend weight (kg)" in column A, value below it)
function bubblesTrend_(ss) {
  const sh = ss.getSheetByName('Dashboard');
  if (!sh) return null;
  const colA = sh.getRange(1, 1, 20, 1).getValues();
  for (let i = 0; i < colA.length - 1; i++) {
    if (/^trend weight/i.test(String(colA[i][0]).trim())) {
      const v = Number(String(sh.getRange(i + 2, 1).getValue()).replace(',', '.'));
      return v > 0 ? bubblesR1_(v) : null;
    }
  }
  return null;
}

// Weigh-ins ("Withings" tab: C = unix timestamp, D = weight kg).
//   dates   = days you weighed in (these tick the habit)
//   values  = per day text shown in the app, e.g. "103.1 scale · 102.2 trend"
//   series  = trend weight on each weigh-in day, for the chart
//   latest  = your current TREND weight (falls back to the scale reading)
function bubblesWeight_(ss) {
  const tz = ss.getSpreadsheetTimeZone();
  const sh = ss.getSheetByName('Withings');
  const rows = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 4).getValues();
  const cutoff = bubblesCutoff_(tz);
  const byDay = {};
  rows.forEach(r => {
    const ts = Number(r[2]), kg = Number(r[3]);
    const iso = ts ? Utilities.formatDate(new Date(ts * 1000), tz, 'yyyy-MM-dd') : bubblesIso_(r[0], tz);
    if (!iso || !kg) return;
    if (!byDay[iso] || ts >= byDay[iso].ts) byDay[iso] = { ts: ts || 0, kg: kg };   // keep the latest reading of each day
  });
  const trends = bubblesTrendSeries_(ss);
  const dates = Object.keys(byDay).filter(d => d >= cutoff).sort();
  const values = {}, series = [];
  dates.forEach(d => {
    const scale = bubblesR1_(byDay[d].kg), tr = trends[d];
    values[d] = scale.toFixed(1) + ' scale' + (tr ? ' · ' + tr.toFixed(1) + ' trend' : '');
    if (tr) series.push({ d: d, v: tr });
  });
  let latest = null;
  if (dates.length) {
    const d = dates[dates.length - 1], scale = bubblesR1_(byDay[d].kg);
    const trend = bubblesTrend_(ss) || trends[d] || null;
    latest = { date: d, value: trend || scale, unit: 'kg', scale: scale, trend: !!trend };
    latest.label = latest.value.toFixed(1) + 'kg';
  }
  return { dates: dates, latest: latest, values: values, series: series };
}

/* ------------------------------------------------------------------- gym */

// 'push' | 'pull' | 'legs' | null (not counted) for one exercise name
function bubblesClassify_(name) {
  const n = String(name).toLowerCase();
  for (let i = 0; i < BUBBLES_PPL.length; i++) {
    const label = BUBBLES_PPL[i][0], words = BUBBLES_PPL[i][1];
    for (let j = 0; j < words.length; j++) if (n.indexOf(words[j]) >= 0) return label === 'skip' ? null : label;
  }
  return null;
}

// Read the Workouts tab: one object per set, found by header names (with sensible fallbacks)
function bubblesWorkoutRows_(ss) {
  const tz = ss.getSpreadsheetTimeZone();
  const sh = ss.getSheetByName('Workouts');
  const lastCol = Math.max(11, sh.getLastColumn());
  const all = sh.getRange(1, 1, Math.max(2, sh.getLastRow()), lastCol).getValues();
  const head = all[0].map(h => String(h).trim().toLowerCase());
  const find = (test, fallback) => { const i = head.findIndex(test); return i >= 0 ? i : fallback; };
  const cDate = find(h => h === 'date', 2);
  const cEx = find(h => /^exc?ersise$|^exercise/.test(h), 4);
  const cType = find(h => h === 'set type', 6);
  const out = [];
  for (let i = 1; i < all.length; i++) {
    const iso = bubblesIso_(all[i][cDate], tz), ex = String(all[i][cEx] || '').trim();
    if (!iso || !ex) continue;
    const type = String(all[i][cType] || '').toLowerCase();
    out.push({ date: iso, exercise: ex, warmup: type.indexOf('warm') >= 0 });
  }
  return out;
}

// Gym sessions: any day with a row on the Workouts tab counts as a session.
// values = per day label shown in the app (Push / Pull / Legs / Mixed), decided by which kind of sets dominate.
function bubblesGym_(ss) {
  const tz = ss.getSpreadsheetTimeZone();
  const cutoff = bubblesCutoff_(tz);
  const days = {};
  bubblesWorkoutRows_(ss).forEach(r => {
    if (r.date < cutoff) return;
    const d = days[r.date] || (days[r.date] = { push: 0, pull: 0, legs: 0 });
    if (r.warmup) return;
    const k = bubblesClassify_(r.exercise);
    if (k) d[k]++;
  });
  const values = {}, types = {};
  Object.keys(days).forEach(date => {
    const c = days[date], total = c.push + c.pull + c.legs;
    let kind = 'other';
    if (total > 0) {
      const top = ['push', 'pull', 'legs'].sort((a, b) => c[b] - c[a])[0];
      kind = c[top] / total >= BUBBLES_PPL_MIN_SHARE ? top : 'mixed';
    }
    types[date] = kind;
    values[date] = kind.charAt(0).toUpperCase() + kind.slice(1);
  });
  return { dates: Object.keys(days).sort(), latest: null, values: values, types: types };
}

// A list of every exercise in the last year, how many sets, and how it is sorted. Use it to tune BUBBLES_PPL.
function bubblesExercises_(ss) {
  const tz = ss.getSpreadsheetTimeZone();
  const cutoff = Utilities.formatDate(new Date(Date.now() - 365 * 86400000), tz, 'yyyy-MM-dd');
  const count = {};
  bubblesWorkoutRows_(ss).forEach(r => { if (r.date >= cutoff && !r.warmup) count[r.exercise] = (count[r.exercise] || 0) + 1; });
  const list = Object.keys(count).map(name => ({ name: name, sets: count[name], counts_as: bubblesClassify_(name) || 'not counted' }));
  list.sort((a, b) => b.sets - a.sets);
  return { exercises: list };
}

/* ----------------------------------------------------------------- steps */

// Steps live on a "Steps" tab (Date, Steps, Updated). The iPhone Shortcut sends today's total to
// ?what=steps_set a few times a day; the app reads it back with ?what=steps.
function bubblesStepsSheet_(ss, create) {
  let sh = ss.getSheetByName('Steps');
  if (!sh && create) {
    sh = ss.insertSheet('Steps');
    sh.getRange(1, 1, 1, 3).setValues([['Date', 'Steps', 'Updated']]);
    sh.getRange('A:A').setNumberFormat('@');            // keep dates as plain text
  }
  return sh;
}

function bubblesStepsSet_(ss, p) {
  const date = String(p.date || ''), steps = Math.round(Number(p.steps));
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { error: 'date must look like 2026-10-28' };
  if (!(steps >= 0 && steps <= 200000)) return { error: 'steps must be a number from 0 to 200000' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = bubblesStepsSheet_(ss, true);
    const n = sh.getLastRow() - 1;
    const dates = n > 0 ? sh.getRange(2, 1, n, 2).getValues() : [];
    let row = 0, best = steps;
    for (let i = dates.length - 1; i >= 0; i--) {
      if (String(dates[i][0]).trim() === date) { row = i + 2; best = Math.max(steps, Number(dates[i][1]) || 0); break; }
    }
    if (!row) row = sh.getLastRow() + 1;
    sh.getRange(row, 1).setNumberFormat('@').setValue(date);
    sh.getRange(row, 2, 1, 2).setValues([[best, new Date()]]);   // a day's total can only go up
    return { ok: true, date: date, steps: best };
  } finally {
    lock.releaseLock();
  }
}

function bubblesSteps_(ss) {
  const tz = ss.getSpreadsheetTimeZone(), cutoff = bubblesCutoff_(tz);
  const todayIso = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const sh = bubblesStepsSheet_(ss, false);
  const dates = [], values = {};
  let todaySteps = null;
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => {
      const iso = bubblesIso_(r[0], tz), n = Number(r[1]);
      if (!iso || !(n >= 0) || iso < cutoff) return;
      values[iso] = n.toLocaleString('en-US') + ' steps';
      if (n >= BUBBLES_STEP_GOAL) dates.push(iso);
      if (iso === todayIso) todaySteps = n;
    });
  }
  dates.sort();
  // only today's running total goes on the bubble (yesterday's number would be confusing)
  const latest = todaySteps === null ? null
    : { date: todayIso, value: todaySteps, unit: ' steps', label: todaySteps >= 1000 ? (todaySteps / 1000).toFixed(1) + 'k' : String(todaySteps) };
  return { dates: dates, latest: latest, values: values };
}

/* -------------------------------------------------------------- calendar */

// Reads the family calendar link (script property FAMILY_CAL_URL), works out repeating events
// and returns the next few days. Only the title, the time and a short place name leave this script:
// notes/descriptions and links (e.g. video-call passwords) are dropped on purpose.
const BUBBLES_CAL_TZ = 'Asia/Tokyo';     // the calendar's time zone (times in the feed are used as written)
const BUBBLES_CAL_DAYS = 30;             // how many days ahead to return by default
const BUBBLES_CAL_CACHE_SECS = 1800;     // the feed only refreshes about every 12 hours, so 30 minutes is plenty

function bubblesCalendar_(p) {
  let url = PropertiesService.getScriptProperties().getProperty('FAMILY_CAL_URL');
  if (!url) return { error: 'Script property FAMILY_CAL_URL is not set (see the setup notes at the top of the script)' };
  url = url.trim().replace(/^webcal:/i, 'https:');
  const days = Math.max(1, Math.min(90, Math.round(Number(p.days)) || BUBBLES_CAL_DAYS));
  const todayIso = Utilities.formatDate(new Date(), BUBBLES_CAL_TZ, 'yyyy-MM-dd');
  const back = Math.max(0, Math.min(7, Math.round(Number(p.back)) || 0));   // also return this many days BEFORE today (for a Monday-start week)
  const key = 'cal_' + todayIso + '_' + days + '_' + back;     // includes the date so "today" is never stale after midnight
  const cache = CacheService.getScriptCache();
  if (!p.refresh) {
    const hit = cache.get(key);
    if (hit) return JSON.parse(hit);
  }
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  if (res.getResponseCode() !== 200) return { error: 'Calendar link answered with code ' + res.getResponseCode() };
  const text = res.getContentText();
  if (text.indexOf('BEGIN:VCALENDAR') < 0) return { error: 'The calendar link did not return a calendar' };

  const todayDay = bubblesDay_(todayIso);
  const out = bubblesCalWindow_(bubblesIcsEvents_(text), todayDay - back, todayDay + days - 1);
  out.generated = Utilities.formatDate(new Date(), BUBBLES_CAL_TZ, "yyyy-MM-dd'T'HH:mm");
  out.from = bubblesDayIso_(todayDay - back);
  out.to = bubblesDayIso_(todayDay + days - 1);
  const json = JSON.stringify(out);
  if (json.length < 90000) cache.put(key, json, BUBBLES_CAL_CACHE_SECS);   // cache values are limited to ~100KB
  return out;
}

// "2026-10-29" <-> a whole number of days since 1970 (plain day maths, no time-zone surprises)
function bubblesDay_(iso) { const m = iso.split('-').map(Number); return Date.UTC(m[0], m[1] - 1, m[2]) / 86400000; }
function bubblesDayIso_(day) { return new Date(day * 86400000).toISOString().slice(0, 10); }
function bubblesHm_(min) { return ('0' + Math.floor(min / 60)).slice(-2) + ':' + ('0' + (min % 60)).slice(-2); }

// Calendar text:  20261029T200000  /  20261029  /  20261029T110000Z   ->  { day, min }  (min = null for all-day)
function bubblesIcsDate_(val) {
  const m = String(val).trim().match(/^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})?(Z)?)?$/);
  if (!m) return null;
  if (m[4] === undefined) return { day: Date.UTC(+m[1], +m[2] - 1, +m[3]) / 86400000, min: null };
  if (m[7] === 'Z') {   // a UTC time: convert to the calendar's own time zone
    const local = Utilities.formatDate(new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5])), BUBBLES_CAL_TZ, 'yyyy-MM-dd HH:mm');
    const t = local.split(/[- :]/).map(Number);
    return { day: Date.UTC(t[0], t[1] - 1, t[2]) / 86400000, min: t[3] * 60 + t[4] };
  }
  return { day: Date.UTC(+m[1], +m[2] - 1, +m[3]) / 86400000, min: +m[4] * 60 + +m[5] };   // times are used as written
}

// Calendar text escapes:  \,  \;  \\  \n
function bubblesIcsText_(s) {
  return String(s || '').replace(/\\(.)/g, (all, c) => (c === 'n' || c === 'N') ? ' ' : c).trim();
}

// The whole calendar file -> a list of raw events
function bubblesIcsEvents_(text) {
  const lines = text.replace(/\r\n?/g, '\n').replace(/\n[ \t]/g, '').split('\n');   // long lines are "folded" onto the next line
  const events = [];
  let cur = null;
  lines.forEach(line => {
    if (line === 'BEGIN:VEVENT') { cur = { exdays: {} }; return; }
    if (line === 'END:VEVENT')   { if (cur) events.push(cur); cur = null; return; }
    if (!cur) return;
    const i = line.indexOf(':');
    if (i < 0) return;
    const name = line.slice(0, i).split(';')[0].toUpperCase(), val = line.slice(i + 1);
    if (name === 'DTSTART') cur.start = bubblesIcsDate_(val);
    else if (name === 'DTEND') cur.end = bubblesIcsDate_(val);
    else if (name === 'SUMMARY') cur.title = bubblesIcsText_(val);
    else if (name === 'LOCATION') cur.place = bubblesPlace_(bubblesIcsText_(val));
    else if (name === 'RRULE') cur.rrule = val;
    else if (name === 'UID') cur.uid = val.trim();
    else if (name === 'STATUS') cur.cancelled = /cancel/i.test(val);
    else if (name === 'RECURRENCE-ID') cur.recurrenceDay = (bubblesIcsDate_(val) || {}).day;
    else if (name === 'EXDATE') val.split(',').forEach(v => { const d = bubblesIcsDate_(v); if (d) cur.exdays[d.day] = true; });
    // everything else (notes, links, attendees…) is deliberately ignored
  });
  return events;
}

// Keep only a short place NAME: no web links (they can hold meeting passwords) and no street addresses
function bubblesPlace_(loc) {
  if (!loc || /^https?:/i.test(loc) || /^〒/.test(loc)) return '';
  return loc.split(',')[0].trim().slice(0, 50);
}

// Which days does one (possibly repeating) event start on, up to lastDay?  Handles DAILY / WEEKLY (with BYDAY) /
// MONTHLY / YEARLY, INTERVAL, COUNT, UNTIL and cancelled days (EXDATE). Returns { days, unsupported }.
function bubblesCalRepeat_(ev, lastDay) {
  const base = ev.start.day;
  if (!ev.rrule) return { days: [base], unsupported: false };
  const r = {};
  ev.rrule.split(';').forEach(kv => { const k = kv.indexOf('='); if (k > 0) r[kv.slice(0, k).toUpperCase()] = kv.slice(k + 1); });
  const interval = Math.max(1, Number(r.INTERVAL) || 1), count = Number(r.COUNT) || 0;
  const until = r.UNTIL ? (bubblesIcsDate_(r.UNTIL) || {}).day : null;
  const freq = r.FREQ;
  const known = ['FREQ', 'INTERVAL', 'COUNT', 'UNTIL', 'WKST', 'BYDAY'];
  const unsupported = Object.keys(r).some(k => known.indexOf(k) < 0) || (r.BYDAY && freq !== 'WEEKLY') || ['DAILY', 'WEEKLY', 'MONTHLY', 'YEARLY'].indexOf(freq) < 0;
  const days = [];
  let made = 0;
  const push = day => {   // returns false once the series is over
    if (until !== null && until !== undefined && day > until) return false;
    if (day > lastDay) return false;
    made++;
    days.push(day);                       // a cancelled day still counts towards COUNT; it is removed below
    return !(count && made >= count);
  };
  const d0 = new Date(base * 86400000), y = d0.getUTCFullYear(), mo = d0.getUTCMonth(), dom = d0.getUTCDate();
  const weekdayOf = day => (new Date(day * 86400000).getUTCDay() + 6) % 7;   // Monday = 0
  const CODES = ['MO', 'TU', 'WE', 'TH', 'FR', 'SA', 'SU'];
  for (let k = 0; k < 5000; k++) {
    if (freq === 'DAILY') { if (!push(base + k * interval)) break; }
    else if (freq === 'WEEKLY' && r.BYDAY) {
      const weekStart = base - weekdayOf(base) + 7 * interval * k;
      const picks = r.BYDAY.split(',').map(s => CODES.indexOf(s.slice(-2))).filter(i => i >= 0).sort((a, b) => a - b);
      let stop = false;
      for (const i of picks) { const day = weekStart + i; if (day < base) continue; if (!push(day)) { stop = true; break; } }
      if (stop) break;
    }
    else if (freq === 'WEEKLY')  { if (!push(base + 7 * interval * k)) break; }
    else if (freq === 'MONTHLY') {
      const t = new Date(Date.UTC(y, mo + interval * k, dom));
      if (t.getUTCDate() !== dom) { if (Date.UTC(y, mo + interval * k, 1) / 86400000 > lastDay) break; continue; }   // e.g. the 31st in a short month: skipped
      if (!push(t.getTime() / 86400000)) break;
    }
    else if (freq === 'YEARLY') {
      const t = new Date(Date.UTC(y + interval * k, mo, dom));
      if (t.getUTCDate() !== dom) { if (t.getTime() / 86400000 > lastDay) break; continue; }
      if (!push(t.getTime() / 86400000)) break;
    }
    else break;
  }
  return { days: days.filter(d => !ev.exdays[d]), unsupported: unsupported };
}

// Raw events -> the events that touch firstDay..lastDay, as small plain objects, sorted
function bubblesCalWindow_(events, firstDay, lastDay) {
  // a one-off edit of a repeating event (RECURRENCE-ID) replaces that day of the series
  const replaced = {};
  events.forEach(e => { if (e.recurrenceDay !== undefined && e.uid) (replaced[e.uid] = replaced[e.uid] || {})[e.recurrenceDay] = true; });
  const found = [];
  let skippedRules = 0;
  events.forEach(ev => {
    if (!ev.start || ev.cancelled || !ev.title) return;
    const allDay = ev.start.min === null;
    // how long: all-day = whole days (the end date is the day AFTER); timed = minutes
    let spanDays = 1, durMin = 0;
    if (allDay) { if (ev.end && ev.end.day > ev.start.day) spanDays = ev.end.day - ev.start.day; }
    else if (ev.end) { durMin = (ev.end.day * 1440 + (ev.end.min || 0)) - (ev.start.day * 1440 + ev.start.min); if (durMin < 0) durMin = 0; }
    const rep = bubblesCalRepeat_(ev, lastDay);
    if (rep.unsupported) skippedRules++;
    const gone = (ev.recurrenceDay === undefined && ev.uid && replaced[ev.uid]) || {};
    rep.days.forEach(day => {
      if (gone[day]) return;
      let endDay, endMin = null;
      if (allDay) endDay = day + spanDays - 1;                                        // last day, inclusive
      else { const endAbs = day * 1440 + ev.start.min + durMin; endDay = Math.floor(endAbs / 1440); endMin = endAbs - endDay * 1440; }
      if (endDay < firstDay || day > lastDay) return;                                  // not in the window
      if (!allDay && endDay === firstDay && endMin === 0 && durMin > 0 && endDay > day) return;   // ended exactly at midnight before the window
      found.push({
        title: ev.title, start: bubblesDayIso_(day), time: allDay ? null : bubblesHm_(ev.start.min),
        end: bubblesDayIso_(endDay), end_time: allDay ? null : bubblesHm_(endMin), all_day: allDay, place: ev.place || ''
      });
    });
  });
  found.sort((a, b) => a.start < b.start ? -1 : a.start > b.start ? 1 : (a.time || '') < (b.time || '') ? -1 : (a.time || '') > (b.time || '') ? 1 : 0);
  return { events: found, skipped_rules: skippedRules, total_in_feed: events.length };
}
