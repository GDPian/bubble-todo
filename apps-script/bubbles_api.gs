/**
 * Bubbles API – lets the Bubbles app read your weigh-ins and gym sessions from this sheet.
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
 *
 * After you edit this script later: Deploy > Manage deployments > pencil > Version: New version > Deploy.
 */

const BUBBLES_SHEET_ID = '1HcqxvQ08C5kr8PLAhzXowsy8WsC9nPwI7Vrz722LIGQ';   // your Hevy sheet
const BUBBLES_DAYS = 1100;                                                  // how far back to report (about 3 years, for habit history)

function doGet(e) {
  const p = (e && e.parameter) || {};
  const secret = PropertiesService.getScriptProperties().getProperty('BUBBLES_TOKEN');
  if (!secret || p.token !== secret) return bubblesJson_({ error: 'unauthorised' });
  try {
    const ss = SpreadsheetApp.openById(BUBBLES_SHEET_ID);
    if (p.what === 'weight') return bubblesJson_(bubblesWeight_(ss));
    if (p.what === 'gym')    return bubblesJson_(bubblesGym_(ss));
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
  const m = String(v).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return null;
}

function bubblesCutoff_(tz) {
  return Utilities.formatDate(new Date(Date.now() - BUBBLES_DAYS * 86400000), tz, 'yyyy-MM-dd');
}

// Trend weight: the number under the "Trend weight (kg)" heading in column A of the Dashboard tab.
// (It finds the heading by its text, so it still works if you move things down a few rows.)
function bubblesTrend_(ss) {
  const sh = ss.getSheetByName('Dashboard');
  if (!sh) return null;
  const colA = sh.getRange(1, 1, 20, 1).getValues();
  for (let i = 0; i < colA.length - 1; i++) {
    if (/^trend weight/i.test(String(colA[i][0]).trim())) {
      const v = Number(String(sh.getRange(i + 2, 1).getValue()).replace(',', '.'));
      return v > 0 ? Math.round(v * 10) / 10 : null;
    }
  }
  return null;
}

// Weigh-ins: "Withings" tab. Column C = unix timestamp, D = weight (kg).
// "dates" = days you weighed in (these tick the habit); "latest.value" = your TREND weight,
// falling back to the raw scale reading if the trend can't be found.
function bubblesWeight_(ss) {
  const tz = ss.getSpreadsheetTimeZone();
  const sh = ss.getSheetByName('Withings');
  const rows = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 4).getValues();
  const cutoff = bubblesCutoff_(tz);
  const days = {};
  let latest = null;
  rows.forEach(r => {
    const ts = Number(r[2]), kg = Number(r[3]);
    const iso = ts ? Utilities.formatDate(new Date(ts * 1000), tz, 'yyyy-MM-dd') : bubblesIso_(r[0], tz);
    if (!iso || !kg) return;
    if (iso >= cutoff) days[iso] = true;
    if (!latest || iso > latest.date) latest = { date: iso, value: Math.round(kg * 10) / 10, unit: 'kg' };
  });
  if (latest) {
    latest.scale = latest.value;                 // what the scale said
    const trend = bubblesTrend_(ss);
    if (trend) { latest.value = trend; latest.trend = true; }
  }
  return { dates: Object.keys(days).sort(), latest: latest };
}

// Gym sessions: "Workouts" tab. Column C = workout date. Any day with a row counts as a session.
function bubblesGym_(ss) {
  const tz = ss.getSpreadsheetTimeZone();
  const sh = ss.getSheetByName('Workouts');
  const rows = sh.getRange(2, 3, Math.max(1, sh.getLastRow() - 1), 1).getValues();
  const cutoff = bubblesCutoff_(tz);
  const days = {};
  rows.forEach(r => {
    const iso = bubblesIso_(r[0], tz);
    if (iso && iso >= cutoff) days[iso] = true;
  });
  return { dates: Object.keys(days).sort(), latest: null };
}
