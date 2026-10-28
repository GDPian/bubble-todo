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
 *
 * After you edit this script later: Deploy > Manage deployments > pencil > Version: New version > Deploy.
 */

const BUBBLES_SHEET_ID = '1HcqxvQ08C5kr8PLAhzXowsy8WsC9nPwI7Vrz722LIGQ';   // your Hevy sheet
const BUBBLES_DAYS = 1100;       // how far back to report (about 3 years, for habit history)
const BUBBLES_STEP_GOAL = 10000; // steps needed for the "10k steps" habit to tick

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
