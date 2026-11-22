/**
 * Family Planner: edit the "This Week" grid directly.
 * - Typing or picking a value in a role row saves it to that day on Days.
 *   Clearing a cell goes back to the Usual Week default.
 * - Typing in a Plans cell adds a one-day event to the Events tab.
 * The grid cell then goes back to its formula so it keeps showing live values.
 */
const DASH = 'This Week';
const DAYS = 'Days';
const EVENTS = 'Events';
// Each block: the row holding the dates and the last row of the block.
const BLOCKS = [
  { hdr: 4, last: 13 },
  { hdr: 15, last: 24 },
];
// Rows (offset from the date row) that fall back to the Usual Week tab.
const TEMPLATE_OFFSETS = [2, 3, 4, 5, 6, 7];
// Notes row: shows typed notes plus any prep due that day from Events.
const NOTES_OFFSET = 9;

function gridFormula(letter, hdr, row) {
  const days = 'IFERROR(INDEX(Days!$A$2:$K,MATCH(' + letter + '$' + hdr +
    ',Days!$A$2:$A,0),MATCH($A' + row + ',Days!$A$1:$K$1,0))';
  if (row - hdr === NOTES_OFFSET) {
    return '=TEXTJOIN(CHAR(10),TRUE,' + days + '&"",""),IFERROR(FILTER("Prep: "&Events!$F$2:$F&" ("&Events!$C$2:$C&")",' +
      'Events!$G$2:$G=' + letter + '$' + hdr + ',Events!$F$2:$F<>""),""))';
  }
  if (TEMPLATE_OFFSETS.indexOf(row - hdr) === -1) {
    return '=' + days + ',"")';
  }
  return '=LET(v,' + days + '&"",""),IF(v<>"",v,IFERROR(INDEX(\'Usual Week\'!$B$5:$H$10,' +
    'MATCH($A' + row + ',\'Usual Week\'!$A$5:$A$10,0),WEEKDAY(' + letter + '$' + hdr + '))&"","")))';
}

function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== DASH) return;

  const ss = e.source;
  const days = ss.getSheetByName(DAYS);
  const dates = days.getRange(2, 1, days.getLastRow() - 1, 1).getValues()
    .map(r => (r[0] instanceof Date ? r[0].toDateString() : ''));
  const headers = days.getRange(1, 1, 1, days.getLastColumn()).getValues()[0];

  const r0 = e.range.getRow();
  const c0 = e.range.getColumn();
  const vals = e.range.getValues();

  for (let i = 0; i < vals.length; i++) {
    for (let j = 0; j < vals[i].length; j++) {
      const row = r0 + i;
      const col = c0 + j;
      if (col < 2 || col > 8) continue;
      const block = BLOCKS.find(b => row > b.hdr && row <= b.last);
      if (!block) continue;

      const cell = sh.getRange(row, col);
      const letter = String.fromCharCode(64 + col);
      const date = sh.getRange(block.hdr, col).getValue();
      let value = vals[i][j];
      if (row - block.hdr === NOTES_OFFSET) {
        // Drop the auto-added prep lines so they are not saved as typed notes.
        value = String(value).split('\n').filter(l => l.indexOf('Prep: ') !== 0).join('\n');
      }

      if (row === block.hdr + 1) {
        if (value !== '' && date instanceof Date) {
          ss.getSheetByName(EVENTS).appendRow([date, '', value, '', '', '', '', '']);
          ss.toast('Added to Events: ' + value);
        } else if (value === '') {
          ss.toast('To remove a plan, delete it on the Events tab.');
        }
        cell.setFormula(gridFormula(letter, block.hdr, row));
        continue;
      }

      const label = sh.getRange(row, 1).getValue();
      const dRow = dates.indexOf(date instanceof Date ? date.toDateString() : '');
      const dCol = headers.indexOf(label);

      if (dRow >= 0 && dCol >= 0) {
        days.getRange(dRow + 2, dCol + 1).setValue(value);
      } else {
        ss.toast('No matching day on the Days tab for that cell.');
      }
      cell.setFormula(gridFormula(letter, block.hdr, row));
    }
  }
}


// ---------------------------------------------------------------------------
// Web app: a phone-friendly planner with This week, Next week and Events tabs.
// Open the deployment URL and add it to the home screen.
// ---------------------------------------------------------------------------
// WEB_FIELDS, TEMPLATE_FIELDS, LIST_FIELDS and the rest of the family-specific
// schema live in Config.js, which is pushed to Apps Script but not committed.
function webFields_() { return FIELD_ORDER; }

const FAVICON = 'https://cdn.jsdelivr.net/gh/twitter/twemoji@14.0.2/assets/72x72/1f468-200d-1f469-200d-1f467-200d-1f466.png';

function lockedPage_() {
  return HtmlService.createHtmlOutput(
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<div style="font:16px -apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Arial,sans-serif;' +
    'max-width:30em;margin:3em auto;padding:0 1.5em;color:#111827">' +
    '<h1 style="font-size:19px">' + APP_TITLE + '</h1>' +
    '<p>This page needs the family access link. Open it from the app on your phone, ' +
    'or add <code>?t=YOUR_TOKEN</code> to the address.</p></div>'
  ).setTitle(APP_TITLE);
}

// The deployment is open to "Anyone", so the page itself checks the token too.
function doGet(e) {
  const t = (e && e.parameter && e.parameter.t) || '';
  if (!tokenOk_(t)) return lockedPage_();
  return HtmlService.createHtmlOutput(
      PAGE_HTML.replace(/__APP_TITLE__/g, APP_TITLE)
        .replace('__SCHEMA__', JSON.stringify(appSchema_())))
    .setTitle(APP_TITLE)
    .setFaviconUrl(FAVICON)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .addMetaTag('mobile-web-app-capable', 'yes');
}

function ymd_(d, tz) { return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); }
function nice_(d, tz) { return Utilities.formatDate(d, tz, 'EEE MMM d'); }
function toDate_(s, tz) { return s ? Utilities.parseDate(s, tz, 'yyyy-MM-dd') : ''; }

function listOptions_(range, fallback) {
  const dv = range.getDataValidation();
  if (dv && dv.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    return dv.getCriteriaValues()[0];
  }
  return fallback;
}

function listFrom_(dv, fallback) {
  if (dv && dv.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    return dv.getCriteriaValues()[0];
  }
  return fallback;
}

function getPlannerData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const days = ss.getSheetByName(DAYS);
  const vals = days.getRange(2, 1, days.getLastRow() - 1, 11).getValues();
  const headers = days.getRange(1, 1, 1, 11).getValues()[0];
  const byDate = {};
  vals.forEach(r => { if (r[0] instanceof Date) byDate[ymd_(r[0], tz)] = r; });

  const now = new Date();
  const todayStr = ymd_(now, tz);
  const dow = Number(Utilities.formatDate(now, tz, 'u')) % 7; // 0 = Sunday

  // Usual Week defaults: rows 5-10, columns B-H are Sun-Sat.
  const uw = ss.getSheetByName('Usual Week').getRange(5, 1, 6, 8).getValues();
  const template = {};
  uw.forEach(r => { template[String(r[0])] = r.slice(1).map(String); });

  // Events, with their sheet row so they can be edited.
  const ev = ss.getSheetByName(EVENTS);
  ensureEventIds_(ev);
  const evVals = ev.getLastRow() > 1 ? ev.getRange(2, 1, ev.getLastRow() - 1, EVENT_ID_COL).getValues() : [];
  const events = [];
  evVals.forEach((r, i) => {
    if (!(r[0] instanceof Date) || String(r[2]) === '') return;
    const end = r[1] instanceof Date ? r[1] : r[0];
    events.push({
      row: i + 2,
      id: String(r[EVENT_ID_COL - 1] || ''),
      start: ymd_(r[0], tz),
      end: r[1] instanceof Date ? ymd_(r[1], tz) : '',
      lastDay: ymd_(end, tz),
      label: nice_(r[0], tz) + (r[1] instanceof Date && ymd_(r[1], tz) !== ymd_(r[0], tz) ? ' to ' + nice_(r[1], tz) : ''),
      title: String(r[2]),
      category: String(r[3]),
      lead: String(r[4]),
      prep: String(r[5]),
      prepBy: r[6] instanceof Date ? ymd_(r[6], tz) : '',
      prepByLabel: r[6] instanceof Date ? nice_(r[6], tz) : '',
      notes: String(r[7]),
    });
  });
  events.sort((a, b) => (a.start < b.start ? -1 : a.start > b.start ? 1 : 0));

  // 14 days starting this Sunday. Days missing from the Days tab still show.
  const out = [];
  for (let k = 0; k < 14; k++) {
    const d = new Date(now.getTime() + (k - dow) * 86400000);
    const ds = ymd_(d, tz);
    const r = byDate[ds];
    const wd = k % 7;
    const fields = {};
    webFields_().forEach(f => {
      const c = headers.indexOf(f);
      const set = r && c >= 0 ? String(r[c]) : '';
      const usual = TEMPLATE_FIELDS.indexOf(f) >= 0 && template[f] ? (template[f][wd] || '') : '';
      fields[f] = { set: set, usual: usual };
    });
    out.push({
      date: ds,
      label: nice_(d, tz),
      weekday: wd,
      editable: !!r,
      plans: events.filter(e => e.start <= ds && e.lastDay >= ds).map(e => e.title),
      prep: events.filter(e => e.prepBy === ds && e.prep).map(e => e.prep + ' (' + e.title + ')'),
      fields: fields,
    });
  }

  // Dropdown choices come from the sheet, so editing them there updates the app.
  const options = {};
  const dvRow = days.getRange(2, 1, 1, 11).getDataValidations()[0];
  LIST_FIELDS.forEach(f => {
    const c = headers.indexOf(f);
    options[f] = c >= 0 ? listFrom_(dvRow[c], PEOPLE) : PEOPLE;
  });
  const meals = ss.getSheetByName('Meals');
  options['Dinner'] = meals.getLastRow() > 1
    ? meals.getRange(2, 1, meals.getLastRow() - 1, 1).getValues().map(r => String(r[0])).filter(String) : [];
  const evDv = ev.getRange(2, 4, 1, 2).getDataValidations()[0];
  options['Category'] = listFrom_(evDv[0], ['Holiday', 'Hosting', 'Birthday', 'Travel', 'Work', 'Kids', 'Money']);
  options['Lead'] = listFrom_(evDv[1], PEOPLE.concat(['Both']));

  const morningCol = headers.indexOf('Morning');
  const carpool = vals
    .filter(r => r[0] instanceof Date && ymd_(r[0], tz) >= todayStr && /carpool/i.test(String(r[morningCol])))
    .slice(0, 8)
    .map(r => ({ label: nice_(r[0], tz), who: String(r[morningCol]) }));

  return { today: todayStr, days: out, options: options, carpool: carpool, events: events,
    schema: appSchema_() };
}

function saveField(dateStr, field, value) {
  if (webFields_().indexOf(field) < 0) throw new Error('Unknown field: ' + field);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const days = ss.getSheetByName(DAYS);
  const headers = days.getRange(1, 1, 1, days.getLastColumn()).getValues()[0];
  const col = headers.indexOf(field) + 1;
  const dates = days.getRange(2, 1, days.getLastRow() - 1, 1).getValues();
  const i = dates.findIndex(r => r[0] instanceof Date && ymd_(r[0], tz) === dateStr);
  if (i < 0 || col < 1) throw new Error('That day is not on the Days tab.');
  days.getRange(i + 2, col).setValue(value);
  return true;
}

function addPlan(dateStr, text) {
  text = String(text || '').trim();
  if (!text) return false;
  return saveEvent(null, '', { start: dateStr, title: text });
}

// Create (row = null) or update an event. expectTitle guards against editing a row that moved.
function saveEvent(row, expectTitle, e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const ev = ss.getSheetByName(EVENTS);
  if (!e || !e.start || !String(e.title || '').trim()) throw new Error('An event needs a start date and a name.');
  const values = [[
    toDate_(e.start, tz), toDate_(e.end || '', tz), String(e.title).trim(), e.category || '',
    e.lead || '', e.prep || '', toDate_(e.prepBy || '', tz), e.notes || '',
  ]];
  if (row) {
    if (String(ev.getRange(row, 3).getValue()) !== expectTitle) throw new Error('This event changed in the sheet. Refresh and try again.');
    ev.getRange(row, 1, 1, 8).setValues(values);
  } else {
    ev.getRange(ev.getLastRow() + 1, 1, 1, 8).setValues(values);
  }
  return true;
}

function deleteEvent(row, expectTitle) {
  const ev = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EVENTS);
  if (String(ev.getRange(row, 3).getValue()) !== expectTitle) throw new Error('This event changed in the sheet. Refresh and try again.');
  ev.deleteRow(row);
  return true;
}

const PAGE_HTML = `<!DOCTYPE html>
<html>
<head>
<base target="_top">
<meta charset="utf-8">
<style>
:root{--navy:#1F3A5F;--bg:#F3F4F6;--card:#fff;--line:#E5E7EB;--muted:#6B7280;--amber:#FCD9A0;--amberText:#633806;--blue:#CBE1F7;--blueText:#0C447C;--today:#FFF2CC;--chip:#EEEDFE;--chipText:#3C3489}
*{box-sizing:border-box}
body{margin:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:var(--bg);color:#111827;font-size:15px}
header{position:sticky;top:0;z-index:5;background:var(--navy);color:#fff;padding:14px 16px 10px}
.top{display:flex;justify-content:space-between;align-items:baseline}
header h1{margin:0;font-size:19px;font-weight:600}
header .sub{font-size:13px;opacity:.85;margin-top:2px}
.refresh{background:none;border:0;color:#fff;opacity:.85;font-size:13px}
.tabs{display:flex;gap:6px;margin-top:10px}
.tab{flex:1;border:0;border-radius:8px;padding:8px 4px;font-size:14px;background:rgba(255,255,255,.15);color:#fff}
.tab.on{background:#fff;color:var(--navy);font-weight:600}
main{padding:12px;max-width:640px;margin:0 auto}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;margin-bottom:12px;overflow:hidden}
.card.today{border:2px solid #D99919}
.card.past{opacity:.55}
.dh{display:flex;justify-content:space-between;align-items:center;padding:10px 12px;border-bottom:1px solid var(--line)}
.card.today .dh{background:var(--today)}
.dh b{font-size:16px}
.badge{font-size:11px;background:#D99919;color:#fff;border-radius:999px;padding:2px 8px}
.plans{padding:8px 12px 0;display:flex;flex-wrap:wrap;gap:6px}
.chip{background:var(--chip);color:var(--chipText);border-radius:999px;padding:3px 10px;font-size:13px}
.prep{margin:8px 12px 0;padding:6px 10px;border-radius:8px;background:#FAEEDA;color:var(--amberText);font-size:13px}
.rows{padding:6px 12px 10px}
.row{display:flex;align-items:center;gap:10px;padding:5px 0}
.row label{width:96px;flex:none;font-size:13px;color:var(--muted)}
.row select,.row input,.f input,.f select{flex:1;min-width:0;width:100%;font-size:15px;padding:7px 9px;border:1px solid var(--line);border-radius:8px;background:#fff;color:#111827;-webkit-appearance:none;appearance:none;font-family:inherit}
.usual{color:#9CA3AF !important;font-style:italic}
.need{background:var(--amber) !important;color:var(--amberText) !important;border-color:#EF9F27 !important}
.pool{background:var(--blue) !important;color:var(--blueText) !important;font-weight:600}
.add{padding:0 12px 12px}
.add input{width:100%;font-size:14px;padding:7px 9px;border:1px dashed #C7CBD1;border-radius:8px}
.note{font-size:12px;color:var(--muted);padding:0 12px 10px}
h2{font-size:15px;margin:18px 4px 8px;color:var(--navy);display:flex;justify-content:space-between;align-items:center}
.list{background:#fff;border:1px solid var(--line);border-radius:12px;overflow:hidden}
.li{display:flex;gap:10px;padding:10px 12px;border-top:1px solid var(--line);font-size:14px;align-items:flex-start}
.li:first-child{border-top:0}
.li .d{width:96px;flex:none;color:var(--muted);font-size:13px}
.li .t{flex:1}
.li .s{display:block;color:var(--muted);font-size:12px;margin-top:2px}
.li.soon{background:var(--today)}
.li.poolrow{background:#EAF3FB}
.li.tap{cursor:pointer}
.empty{padding:10px 12px;color:var(--muted);font-size:14px}
.btn{border:0;border-radius:8px;padding:8px 12px;font-size:14px;background:var(--navy);color:#fff;font-family:inherit}
.btn.ghost{background:#fff;color:var(--navy);border:1px solid var(--line)}
.btn.danger{background:#fff;color:#A32D2D;border:1px solid #F09595}
.btn.small{padding:5px 10px;font-size:13px}
.editor{padding:12px;border-top:1px solid var(--line);background:#FAFAFB}
.f{margin-bottom:8px}
.f span{display:block;font-size:12px;color:var(--muted);margin-bottom:3px}
.two{display:flex;gap:8px}
.two .f{flex:1}
.actions{display:flex;gap:8px;margin-top:4px}
.actions .grow{flex:1}
#toast{position:fixed;left:50%;bottom:22px;transform:translateX(-50%);background:#111827;color:#fff;padding:8px 14px;border-radius:999px;font-size:13px;opacity:0;transition:opacity .2s;pointer-events:none;z-index:9}
#toast.show{opacity:.92}
.loading{padding:40px;text-align:center;color:var(--muted)}
main.wide{max-width:1180px}
.seg{display:flex;gap:4px;background:#E5E7EB;border-radius:8px;padding:3px;margin:0 0 10px;width:max-content}
.seg button{border:0;background:none;padding:5px 14px;border-radius:6px;font-size:13px;color:var(--muted);font-family:inherit}
.seg button.on{background:#fff;color:var(--navy);font-weight:600}
.tablewrap{overflow-x:auto;-webkit-overflow-scrolling:touch;background:#fff;border:1px solid var(--line);border-radius:12px}
table.grid{border-collapse:separate;border-spacing:0;width:100%}
.grid th,.grid td{border-bottom:1px solid var(--line);border-right:1px solid var(--line);padding:4px;vertical-align:top;min-width:124px}
.grid thead th{background:var(--navy);color:#fff;font-size:13px;font-weight:600;padding:8px 6px;text-align:left;white-space:nowrap}
.grid thead th.today{background:#D99919}
.grid .rl{position:sticky;left:0;z-index:2;background:#F1F3F5;min-width:78px;width:78px;font-size:12px;font-weight:600;color:#374151;padding:8px 6px}
.grid thead th.rl{background:var(--navy);color:#fff;z-index:3}
.grid td.today{background:#FFFBEB}
.grid td.past,.grid th.past{opacity:.5}
.grid select,.grid input{width:100%;font-size:13px;padding:5px 6px;border:1px solid var(--line);border-radius:6px;background:#fff;color:#111827;-webkit-appearance:none;appearance:none;font-family:inherit}
.grid .chip{display:inline-block;font-size:12px;margin:0 0 3px;padding:2px 8px}
.grid .prep{margin:0 0 4px;padding:3px 6px;font-size:12px}
.grid .plan-add{border-style:dashed;margin-top:2px}
main.stale select,main.stale input,main.stale button,main.stale .tap{pointer-events:none}

.tab{font-size:13px}
.gcard{background:#fff;border:1px solid var(--line);border-radius:12px;padding:10px 12px;margin-bottom:12px}
.grow{display:flex;gap:8px}
#gin{flex:1;min-width:0;font-size:16px;padding:10px;border:1px solid var(--line);border-radius:8px}
.gbtn{border:0;background:var(--navy);color:#fff;border-radius:8px;padding:8px 16px;font-size:15px}
.gchips{display:flex;flex-wrap:wrap;gap:6px;margin-top:8px;align-items:center}
.gat{margin:0 0 12px}
.gchip{border:1px solid var(--line);background:#fff;border-radius:999px;padding:6px 11px;font-size:13px;color:#374151}
.gchip.on{background:var(--navy);border-color:var(--navy);color:#fff}
.gchip.spec.on{background:#D99919;border-color:#D99919}
.glbl{font-size:12px;color:var(--muted);margin-right:2px}
.gwhere{background:#FFF7E6;border-color:#F3C77A}
.gh{font-size:12px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;margin:4px 0 6px}
.gw{margin:4px 0}
.gsmall{font-size:12px;color:var(--muted);margin-top:6px}
.gpad{padding:2px 12px 10px;margin:0}
.gsec{font-size:12px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;padding:10px 12px 4px;border-top:1px solid var(--line)}
.glist .gsec:first-child{border-top:0}
.gsec.spec{color:#8A5A00;background:#FFF7E6}
.glist{background:#fff;border:1px solid var(--line);border-radius:12px;margin-bottom:12px;overflow:hidden}
.gi{display:flex;align-items:center;gap:10px;padding:9px 12px;border-top:1px solid #F3F4F6;flex-wrap:wrap}
.gi.gspec{background:#FFFBF2}
.gi.gpend{opacity:.55}
.gck{width:24px;height:24px;border:2px solid #9CA3AF;border-radius:6px;background:#fff;flex:none;padding:0;font-size:14px;line-height:18px;color:#fff}
.gck.on{background:#9CA3AF}
.gn{flex:1;min-width:0}
.gnote{display:block;font-size:12px;color:var(--muted)}
.ggot .gnote{display:inline}
.gtag{font-size:11px;background:#FCD9A0;color:#633806;border-radius:999px;padding:2px 8px;white-space:nowrap}
.ggot .gn{text-decoration:line-through;color:var(--muted)}
.gedit{width:100%;padding:2px 0 4px 34px}
.gedit input{width:100%;font-size:16px;padding:8px;border:1px solid var(--line);border-radius:8px;margin-top:8px}
.gacts{display:flex;gap:12px;margin-top:8px;align-items:center}
.glink{background:none;border:0;color:var(--blueText);font-size:14px;padding:6px 0}
.gdanger{color:#B91C1C;margin-left:auto}
</style>
</head>
<body>
<header>
  <div class="top"><h1>👨‍👩‍👧‍👦 __APP_TITLE__</h1><button class="refresh" onclick="load()">Refresh</button></div>
  <div class="sub" id="sub">Loading...</div>
  <div class="tabs">
    <button class="tab on" id="t0" onclick="showTab(0)">This week</button>
    <button class="tab" id="t1" onclick="showTab(1)">Next week</button>
    <button class="tab" id="t2" onclick="showTab(2)">Events</button>
    <button class="tab" id="t3" onclick="showTab(3)">Groceries</button>
  </div>
</header>
<main id="main"><div class="loading">Loading the planner...</div></main>
<div id="toast"></div>
<script>
var DATA = null;
var TAB = 0;
var EDIT = null;       // event row being edited, or 'new'
var SHOWALL = false;
var VIEW = 'table';    // 'table' or 'cards' for the week tabs
var AUTO_SCROLL = true; // scroll the table to today on the next render
var CFG = __SCHEMA__;
var ORDER = CFG.order, LABELS = CFG.labels, SCHOOL = CFG.weekdayOnly;
var EVERYDAY = CFG.everyday, TEXT = CFG.text;

function esc(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function toast(msg) {
  var t = document.getElementById('toast');
  t.textContent = msg; t.className = 'show';
  clearTimeout(toast.h); toast.h = setTimeout(function () { t.className = ''; }, 1800);
}
var STALE = false;
function cacheGet() { try { var s = window.localStorage.getItem('rfa-cache'); return s ? JSON.parse(s) : null; } catch (e) { return null; } }
function cachePut(d) { try { window.localStorage.setItem('rfa-cache', JSON.stringify(d)); } catch (e) {} }
function load() {
  // Show the last saved copy right away, then refresh from the sheet in the background.
  var cached = DATA ? null : cacheGet();
  if (cached) { DATA = cached; STALE = true; render(); }
  else if (DATA) { STALE = true; render(); }
  else { document.getElementById('sub').textContent = 'Loading...'; }
  google.script.run
    .withSuccessHandler(function (d) { DATA = d; STALE = false; cachePut(d); render(); })
    .withFailureHandler(function (e) {
      STALE = false;
      if (DATA) { render(); toast('Could not refresh: ' + e.message); }
      else { document.getElementById('main').innerHTML = '<div class="loading">Could not load: ' + esc(e.message) + '</div>'; }
    })
    .getPlannerData();
}
function showTab(t) {
  TAB = t; EDIT = null; AUTO_SCROLL = true;
  try { localStorage.setItem('rfa-tab', String(t)); } catch (e) {}
  if (t === 3) loadGroceries();
  for (var i = 0; i < 4; i++) document.getElementById('t' + i).className = 'tab' + (i === t ? ' on' : '');
  render();
  window.scrollTo(0, 0);
}
function addDays(ds, n) {
  var p = ds.split('-');
  var d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]) + n);
  var m = d.getMonth() + 1, dd = d.getDate();
  return d.getFullYear() + '-' + (m < 10 ? '0' : '') + m + '-' + (dd < 10 ? '0' : '') + dd;
}

/* ---------- day cards ---------- */
function needed(day, f) {
  if (EVERYDAY.indexOf(f) >= 0) return true;
  if (SCHOOL.indexOf(f) >= 0) return day.weekday >= 1 && day.weekday <= 5;
  return false;
}
function shown(fv) { return fv.set !== '' ? fv.set : fv.usual; }
function openCount() {
  var n = 0;
  DATA.days.forEach(function (day) {
    if (day.date < DATA.today) return;
    ORDER.forEach(function (f) {
      if (!needed(day, f)) return;
      var v = shown(day.fields[f]);
      if (v === '' || v === 'Carpool') n++;
    });
  });
  return n;
}
function cls(day, f, fv) {
  var v = shown(fv);
  if (needed(day, f) && day.date >= DATA.today && (v === '' || v === 'Carpool')) return 'need';
  if (/[(]carpool[)]/i.test(v)) return 'pool';
  if (fv.set === '' && fv.usual !== '') return 'usual';
  return '';
}
function control(day, f) {
  var fv = day.fields[f];
  var id = day.date + '|' + f;
  var dis = day.editable ? '' : ' disabled';
  if (TEXT.indexOf(f) >= 0) {
    return '<input data-id="' + esc(id) + '" value="' + esc(fv.set) + '" placeholder="Add"' + dis + ' onchange="saveText(this)">';
  }
  var opts = (DATA.options[f] || []).slice();
  if (fv.set !== '' && opts.indexOf(fv.set) < 0) opts.unshift(fv.set);
  var first = fv.usual !== '' ? 'Usual: ' + fv.usual : 'Not set';
  var h = '<select data-id="' + esc(id) + '" class="' + cls(day, f, fv) + '"' + dis + ' onchange="saveSel(this)">';
  h += '<option value=""' + (fv.set === '' ? ' selected' : '') + '>' + esc(first) + '</option>';
  opts.forEach(function (o) {
    h += '<option value="' + esc(o) + '"' + (o === fv.set ? ' selected' : '') + '>' + esc(o) + '</option>';
  });
  h += '<option value="__other">Other...</option></select>';
  return h;
}
function dayCard(day) {
  var isToday = day.date === DATA.today;
  var past = day.date < DATA.today;
  var h = '<div class="card' + (isToday ? ' today' : '') + (past ? ' past' : '') + '">';
  h += '<div class="dh"><b>' + esc(day.label) + '</b>' + (isToday ? '<span class="badge">Today</span>' : '') + '</div>';
  if (day.plans.length) {
    h += '<div class="plans">' + day.plans.map(function (p) { return '<span class="chip">' + esc(p) + '</span>'; }).join('') + '</div>';
  }
  day.prep.forEach(function (p) { h += '<div class="prep">Prep: ' + esc(p) + '</div>'; });
  h += '<div class="rows">';
  ORDER.forEach(function (f) {
    h += '<div class="row"><label>' + esc(LABELS[f] || f) + '</label>' + control(day, f) + '</div>';
  });
  h += '</div>';
  if (day.editable) {
    h += '<div class="add"><input placeholder="+ Add a plan for this day" data-date="' + day.date + '" onkeydown="planKey(event, this)" onblur="planSave(this)"></div>';
  } else {
    h += '<div class="note">This day is before the planner starts, so it is view only.</div>';
  }
  return h + '</div>';
}
function viewTable() { VIEW = 'table'; AUTO_SCROLL = true; render(); }
function viewCards() { VIEW = 'cards'; render(); }
function tableView(days) {
  var h = '<div class="tablewrap"><table class="grid"><thead><tr><th class="rl"></th>';
  days.forEach(function (d) {
    var c = d.date === DATA.today ? 'today' : (d.date < DATA.today ? 'past' : '');
    h += '<th class="' + c + '">' + esc(d.label) + (d.date === DATA.today ? '  (today)' : '') + '</th>';
  });
  h += '</tr></thead><tbody>';
  // Plans row: chips, prep due that day, and a box to add a plan.
  h += '<tr><td class="rl">Plans</td>';
  days.forEach(function (d) {
    var c = d.date === DATA.today ? 'today' : (d.date < DATA.today ? 'past' : '');
    var cell = d.plans.map(function (p) { return '<span class="chip">' + esc(p) + '</span>'; }).join(' ');
    d.prep.forEach(function (p) { cell += '<div class="prep">Prep: ' + esc(p) + '</div>'; });
    if (d.editable) cell += '<input class="plan-add" placeholder="+ plan" data-date="' + d.date + '" onkeydown="planKey(event, this)" onblur="planSave(this)">';
    h += '<td class="' + c + '">' + cell + '</td>';
  });
  h += '</tr>';
  ORDER.forEach(function (f) {
    h += '<tr><td class="rl">' + esc(LABELS[f] || f) + '</td>';
    days.forEach(function (d) {
      var c = d.date === DATA.today ? 'today' : (d.date < DATA.today ? 'past' : '');
      h += '<td class="' + c + '">' + control(d, f) + '</td>';
    });
    h += '</tr>';
  });
  return h + '</tbody></table></div>';
}
function weekView(w) {
  var days = DATA.days.slice(w * 7, w * 7 + 7);
  var h = '<div class="seg"><button class="' + (VIEW === 'table' ? 'on' : '') + '" onclick="viewTable()">Table</button>' +
    '<button class="' + (VIEW === 'cards' ? 'on' : '') + '" onclick="viewCards()">Cards</button></div>';
  if (VIEW === 'table') {
    h += tableView(days);
  } else if (w === 0) {
    days.filter(function (d) { return d.date >= DATA.today; }).forEach(function (d) { h += dayCard(d); });
    days.filter(function (d) { return d.date < DATA.today; }).forEach(function (d) { h += dayCard(d); });
  } else {
    days.forEach(function (d) { h += dayCard(d); });
  }
  h += '<h2>Upcoming carpool</h2><div class="list">' + (DATA.carpool.length ? DATA.carpool.map(function (x) {
    return '<div class="li poolrow"><span class="d">' + esc(x.label) + '</span><span class="t">' + esc(x.who) + '</span></div>';
  }).join('') : '<div class="empty">None scheduled</div>') + '</div>';
  return h;
}

/* ---------- events ---------- */
function opts(list, cur) {
  var l = list.slice();
  if (cur && l.indexOf(cur) < 0) l.unshift(cur);
  return '<option value="">None</option>' + l.map(function (o) {
    return '<option value="' + esc(o) + '"' + (o === cur ? ' selected' : '') + '>' + esc(o) + '</option>';
  }).join('');
}
function editor(e) {
  e = e || { row: null, title: '', start: DATA.today, end: '', category: '', lead: '', prep: '', prepBy: '', notes: '' };
  var h = '<div class="editor" id="ed">';
  h += '<div class="f"><span>What</span><input id="e_title" value="' + esc(e.title) + '" placeholder="e.g. Hosting the Cohens"></div>';
  h += '<div class="two"><div class="f"><span>Starts</span><input type="date" id="e_start" value="' + esc(e.start) + '"></div>';
  h += '<div class="f"><span>Ends (optional)</span><input type="date" id="e_end" value="' + esc(e.end) + '"></div></div>';
  h += '<div class="two"><div class="f"><span>Category</span><select id="e_cat">' + opts(DATA.options.Category, e.category) + '</select></div>';
  h += '<div class="f"><span>Lead</span><select id="e_lead">' + opts(DATA.options.Lead, e.lead) + '</select></div></div>';
  h += '<div class="f"><span>Prep needed</span><input id="e_prep" value="' + esc(e.prep) + '" placeholder="e.g. Menu, invite guests"></div>';
  h += '<div class="f"><span>Prep by</span><input type="date" id="e_prepby" value="' + esc(e.prepBy) + '"></div>';
  h += '<div class="f"><span>Notes</span><input id="e_notes" value="' + esc(e.notes) + '"></div>';
  h += '<div class="actions"><button class="btn" onclick="saveEv()">Save</button><button class="btn ghost" onclick="cancelEv()">Cancel</button><span class="grow"></span>';
  if (e.row) h += '<button class="btn danger" onclick="delEv()">Delete</button>';
  return h + '</div></div>';
}
function findEvent(row) { return DATA.events.filter(function (e) { return e.row === row; })[0]; }
function eventRow(e, soonLimit) {
  var soon = e.start < soonLimit;
  var sub = [];
  if (e.category) sub.push(e.category);
  if (e.lead) sub.push(e.lead);
  if (e.prep) sub.push('Prep: ' + e.prep + (e.prepByLabel ? ' by ' + e.prepByLabel : ''));
  var h = '<div class="li tap' + (soon ? ' soon' : '') + '" onclick="editEv(' + e.row + ')"><span class="d">' + esc(e.label) + '</span><span class="t">' + esc(e.title) +
    (sub.length ? '<span class="s">' + esc(sub.join('  |  ')) + '</span>' : '') + '</span></div>';
  if (EDIT === e.row) h += editor(e);
  return h;
}
function eventsView() {
  var in14 = addDays(DATA.today, 14);
  var horizon = addDays(DATA.today, 84);
  var h = '';
  var prep = DATA.events.filter(function (e) { return e.prep && e.prepBy && e.prepBy >= DATA.today && e.prepBy < in14; })
    .sort(function (a, b) { return a.prepBy < b.prepBy ? -1 : 1; });
  h += '<h2>Prep due in the next 2 weeks</h2><div class="list">' + (prep.length ? prep.map(function (e) {
    return '<div class="li tap" onclick="editEv(' + e.row + ')"><span class="d">' + esc(e.prepByLabel) + '</span><span class="t">' + esc(e.prep) +
      '<span class="s">for ' + esc(e.title) + (e.lead ? ', ' + esc(e.lead) : '') + '</span></span></div>';
  }).join('') : '<div class="empty">Nothing due</div>') + '</div>';

  var upcoming = DATA.events.filter(function (e) { return e.lastDay >= DATA.today && (SHOWALL || e.start < horizon); });
  h += '<h2>Upcoming events <button class="btn small" onclick="newEv()">+ New event</button></h2>';
  h += '<div class="list">';
  if (EDIT === 'new') h += editor(null);
  h += upcoming.length ? upcoming.map(function (e) { return eventRow(e, in14); }).join('') : '<div class="empty">Nothing scheduled</div>';
  h += '</div>';
  var more = DATA.events.filter(function (e) { return e.start >= horizon; }).length;
  if (more && !SHOWALL) h += '<div style="text-align:center;margin-top:10px"><button class="btn ghost small" onclick="SHOWALL=true;render()">Show ' + more + ' later events</button></div>';
  h += '<div class="note" style="padding:10px 4px">Yellow = in the next 2 weeks. Tap any event to edit it.</div>';
  return h;
}
function editEv(row) { EDIT = (EDIT === row ? null : row); render(); }
function newEv() { EDIT = 'new'; render(); }
function cancelEv() { EDIT = null; render(); }
function val(id) { return document.getElementById(id).value.trim(); }
function saveEv() {
  var e = { title: val('e_title'), start: val('e_start'), end: val('e_end'), category: val('e_cat'), lead: val('e_lead'),
    prep: val('e_prep'), prepBy: val('e_prepby'), notes: val('e_notes') };
  if (!e.title || !e.start) { toast('Add a name and a start date'); return; }
  if (e.end && e.end < e.start) { toast('End date is before the start'); return; }
  var row = EDIT === 'new' ? null : EDIT;
  var old = row ? findEvent(row).title : '';
  toast('Saving...');
  google.script.run
    .withSuccessHandler(function () { EDIT = null; toast('Saved'); load(); })
    .withFailureHandler(function (err) { toast('Not saved: ' + err.message); })
    .saveEvent(row, old, e);
}
function delEv() {
  var e = findEvent(EDIT);
  if (!window.confirm('Delete "' + e.title + '"?')) return;
  google.script.run
    .withSuccessHandler(function () { EDIT = null; toast('Deleted'); load(); })
    .withFailureHandler(function (err) { toast('Not deleted: ' + err.message); })
    .deleteEvent(e.row, e.title);
}

/* ---------- render and save ---------- */
function render() {
  if (TAB === 3) { renderGroceries(); return; }
  if (!DATA) return;
  var c = openCount();
  var todayDay = DATA.days.filter(function (d) { return d.date === DATA.today; })[0];
  document.getElementById('sub').textContent = (todayDay ? todayDay.label + '  |  ' : '') +
    (c === 0 ? 'All covered for 2 weeks' : c + ' open slot' + (c === 1 ? '' : 's') + ' in the next 2 weeks') +
    (STALE ? '  |  Updating...' : '');
  var h = TAB === 2 ? eventsView() : weekView(TAB);
  // Keep the scroll position so saving a cell does not jump the view.
  var wrap = document.querySelector ? document.querySelector('.tablewrap') : null;
  var sx = wrap ? wrap.scrollLeft : 0;
  var sy = window.scrollY || 0;
  var main = document.getElementById('main');
  main.className = ((TAB !== 2 && VIEW === 'table') ? 'wide' : '') + (STALE ? ' stale' : '');
  main.innerHTML = h + '<div style="height:40px"></div>';
  var wrap2 = document.querySelector ? document.querySelector('.tablewrap') : null;
  if (wrap2) {
    var todayTh = AUTO_SCROLL ? wrap2.querySelector('th.today') : null;
    if (todayTh) {
      var label = wrap2.querySelector('th.rl');
      wrap2.scrollLeft = todayTh.offsetLeft - (label ? label.offsetWidth : 0);
      AUTO_SCROLL = false;
    } else {
      wrap2.scrollLeft = sx;
    }
  }
  if (window.scrollTo) window.scrollTo(0, sy);
}
function findDay(date) { return DATA.days.filter(function (d) { return d.date === date; })[0]; }
function save(id, value) {
  var parts = id.split('|');
  var day = findDay(parts[0]);
  var f = parts[1];
  var old = day.fields[f].set;
  day.fields[f].set = value;
  render();
  google.script.run
    .withSuccessHandler(function () { toast(value === '' ? 'Back to usual' : 'Saved'); cachePut(DATA); })
    .withFailureHandler(function (e) { day.fields[f].set = old; render(); toast('Not saved: ' + e.message); })
    .saveField(parts[0], f, value);
}
function saveSel(el) {
  var v = el.value;
  if (v === '__other') {
    v = window.prompt('Enter a name or meal');
    if (v === null) { render(); return; }
    v = v.trim();
  }
  save(el.getAttribute('data-id'), v);
}
function saveText(el) { save(el.getAttribute('data-id'), el.value.trim()); }
function planKey(ev, el) { if (ev.key === 'Enter') { el.blur(); } }
function planSave(el) {
  var text = el.value.trim();
  if (!text) return;
  el.value = '';
  var date = el.getAttribute('data-date');
  findDay(date).plans.push(text);
  render();
  google.script.run
    .withSuccessHandler(function () { toast('Plan added'); load(); })
    .withFailureHandler(function (e) { toast('Not saved: ' + e.message); load(); })
    .addPlan(date, text);
}
load();

// ---------- Groceries tab ----------
var G = null, GSTORE = '', GSEL = [], GTOUCH = false, GDRAFT = '', GEDIT = null, GESEL = [], GBUSY = 0, GSHOWGOT = false;
var GSECTIONS = ['Produce', 'Dairy & eggs', 'Meat & fish', 'Bakery', 'Pantry', 'Frozen', 'Household', 'Other'];
function gq(s) { return "'" + s + "'"; }
function gBy() { try { return localStorage.getItem('rfa-by') || ''; } catch (e) { return ''; } }
function gSetBy(n) { try { localStorage.setItem('rfa-by', n); } catch (e) {} renderGroceries(); }
function gCacheGet() { try { var s = localStorage.getItem('rfa-groc'); return s ? JSON.parse(s) : null; } catch (e) { return null; } }
function gCachePut(d) { try { localStorage.setItem('rfa-groc', JSON.stringify(d)); } catch (e) {} }
function gStores(w) { if (!w || w === 'Anywhere') return []; return String(w).split(',').map(function (s) { return s.trim(); }).filter(function (s) { return s; }); }
function gWhere(a) { return a.length ? a.join(', ') : 'Anywhere'; }
function gCat(item) {
  var k = String(item || '').trim().toLowerCase(), c = (G && G.catalog) || [];
  for (var i = 0; i < c.length; i++) if (c[i].item.toLowerCase() === k) return c[i];
  return null;
}
function gToday() { return (G && G.today) || ''; }
function gShort(d) {
  if (!d) return '';
  if (d === gToday()) return 'today';
  var p = d.split('-');
  return ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][+p[1] - 1] + ' ' + (+p[2]);
}
function loadGroceries(quiet) {
  if (!G) { G = gCacheGet(); if (G && TAB === 3) renderGroceries(); }
  google.script.run
    .withSuccessHandler(function (d) { if (GBUSY > 0) return; G = d; gCachePut(d); if (TAB === 3 && !GEDIT) renderGroceries(); })
    .withFailureHandler(function (e) { if (!quiet) toast('Could not load groceries: ' + e.message); })
    .getGroceries();
}
function gSave(call, optimistic) {
  if (optimistic) optimistic();
  GBUSY++;
  renderGroceries();
  call(google.script.run
    .withSuccessHandler(function (d) { GBUSY--; G = d; gCachePut(d); if (TAB === 3) renderGroceries(); })
    .withFailureHandler(function (e) { GBUSY--; toast('Not saved: ' + e.message); loadGroceries(true); }));
}
function gFind(id) { for (var i = 0; i < G.need.length; i++) if (G.need[i].id === id) return i; return -1; }
function gKey(e) { if (e.key === 'Enter') { e.preventDefault(); gAdd(); } }
function gTyped(v) {
  GDRAFT = v;
  if (GTOUCH) return;
  var c = gCat(v);
  GSEL = c ? gStores(c.where) : [];
  gPaintChips();
}
function gChips(sel, fn) {
  var h = '<span class="glbl">Only at</span><button class="gchip' + (sel.length ? '' : ' on') + '" onclick="' + fn + '(-1)">Anywhere</button>';
  G.stores.forEach(function (s, i) { h += '<button class="gchip spec' + (sel.indexOf(s) >= 0 ? ' on' : '') + '" onclick="' + fn + '(' + i + ')">' + esc(s) + '</button>'; });
  return h;
}
function gPaintChips() { var el = document.getElementById('gchips'); if (el) el.innerHTML = gChips(GSEL, 'gToggle'); }
function gToggle(i) {
  GTOUCH = true;
  if (i < 0) GSEL = [];
  else { var s = G.stores[i], k = GSEL.indexOf(s); if (k >= 0) GSEL.splice(k, 1); else GSEL.push(s); }
  gPaintChips();
}
function gAdd(name) {
  var item = String(name != null ? name : GDRAFT).trim();
  if (!item || !G) return;
  var c = gCat(item);
  if (c) item = c.item;
  var where = name != null ? (c ? c.where : 'Anywhere') : gWhere(GSEL);
  GDRAFT = ''; GSEL = []; GTOUCH = false;
  var inp = document.getElementById('gin'); if (inp) inp.value = '';
  for (var i = 0; i < G.need.length; i++) {
    if (G.need[i].item.toLowerCase() === item.toLowerCase()) { toast(item + ' is already on the list'); renderGroceries(); return; }
  }
  var by = gBy();
  gSave(function (r) { r.addGrocery(item, where, '', by); }, function () {
    G.need.push({ id: 'tmp' + Date.now(), item: item, where: where, note: '', addedBy: by, pending: true });
    if (!c) G.catalog.push({ item: item, where: where, section: '' });
  });
}
function gQuick(i) { gAdd(G.frequent[i]); }
function gCheck(id) {
  var k = gFind(id); if (k < 0 || G.need[k].pending) return;
  var r = G.need[k], at = GSTORE;
  gSave(function (s) { s.setGroceryStatus(id, 'Got', at); }, function () {
    G.need.splice(k, 1); r.status = 'Got'; r.gotOn = gToday(); r.gotAt = at; G.got.unshift(r);
  });
}
function gReadd(i) {
  var r = G.got[i]; if (!r) return;
  if (r.gotOn === gToday()) {
    gSave(function (s) { s.setGroceryStatus(r.id, 'Need', ''); }, function () {
      G.got.splice(i, 1); r.status = 'Need'; r.gotOn = ''; r.gotAt = ''; G.need.push(r);
    });
  } else gAdd(r.item);
}
function gAt(i) { GSTORE = i < 0 ? '' : G.stores[i]; renderGroceries(); window.scrollTo(0, 0); }
function gEdit(id) { var k = gFind(id); if (k < 0) return; GEDIT = GEDIT === id ? null : id; GESEL = gStores(G.need[k].where); renderGroceries(); }
function gEToggle(i) {
  if (i < 0) GESEL = [];
  else { var s = G.stores[i], k = GESEL.indexOf(s); if (k >= 0) GESEL.splice(k, 1); else GESEL.push(s); }
  var el = document.getElementById('gechips'); if (el) el.innerHTML = gChips(GESEL, 'gEToggle');
}
function gEditSave() {
  var id = GEDIT, k = gFind(id); if (k < 0) return;
  var note = (document.getElementById('gnote') || {}).value || '', where = gWhere(GESEL), item = G.need[k].item;
  GEDIT = null;
  gSave(function (s) { s.setGroceryDetails(id, where, note); }, function () {
    G.need.forEach(function (r) { if (r.item === item) r.where = where; });
    G.need[k].note = note;
    var c = gCat(item); if (c) c.where = where;
  });
}
function gRemove() {
  var id = GEDIT, k = gFind(id); if (k < 0) return;
  GEDIT = null;
  gSave(function (s) { s.setGroceryStatus(id, 'Removed', ''); }, function () { G.need.splice(k, 1); });
}
function gRow(r) {
  var s = gStores(r.where);
  var h = '<div class="gi' + (s.length ? ' gspec' : '') + (r.pending ? ' gpend' : '') + '">' +
    '<button class="gck" aria-label="Got it" onclick="gCheck(' + gq(r.id) + ')"></button>' +
    '<span class="gn" onclick="gEdit(' + gq(r.id) + ')">' + esc(r.item) + (r.note ? '<span class="gnote">' + esc(r.note) + '</span>' : '') + '</span>' +
    (s.length ? '<span class="gtag">' + esc(s.join(' / ')) + '</span>' : '');
  if (GEDIT === r.id) {
    h += '<div class="gedit"><div class="gchips" id="gechips">' + gChips(GESEL, 'gEToggle') + '</div>' +
      '<input id="gnote" placeholder="Note (brand, size, how many)" value="' + esc(r.note) + '">' +
      '<div class="gacts"><button class="gbtn" onclick="gEditSave()">Save</button>' +
      '<button class="glink" onclick="gEdit(' + gq(r.id) + ')">Cancel</button>' +
      '<button class="glink gdanger" onclick="gRemove()">Remove</button></div></div>';
  }
  return h + '</div>';
}
function gGroup(title, rows, spec) {
  if (!rows.length) return '';
  return '<div class="gsec' + (spec ? ' spec' : '') + '">' + esc(title) + '</div>' + rows.map(gRow).join('');
}
function gSecOf(r) { var c = gCat(r.item); return c && GSECTIONS.indexOf(c.section) >= 0 ? c.section : 'Other'; }
function gAnywhere(rows) {
  var h = '';
  GSECTIONS.forEach(function (sec) { h += gGroup(sec, rows.filter(function (r) { return gSecOf(r) === sec; }), false); });
  return h;
}
function renderGroceries() {
  var main = document.getElementById('main');
  main.className = '';
  if (!G) { main.innerHTML = '<div class="loading">Loading groceries...</div>'; return; }
  var focused = document.activeElement && document.activeElement.id === 'gin';
  var need = G.need, any = [], spec = [], rules = {}, order = [];
  need.forEach(function (r) {
    var s = gStores(r.where);
    if (!s.length) { any.push(r); return; }
    spec.push(r);
    var k = s.join(' or ');
    if (!rules[k]) { rules[k] = []; order.push(k); }
    rules[k].push(r);
  });
  document.getElementById('sub').textContent = need.length + ' to get' + (GSTORE ? '  |  Shopping at ' + GSTORE : '') + (GBUSY ? '  |  Saving...' : '');
  var h = '<div class="gcard"><div class="grow"><input id="gin" list="gcat" placeholder="Add an item" autocomplete="off" enterkeyhint="done" oninput="gTyped(this.value)" onkeydown="gKey(event)" value="' + esc(GDRAFT) + '">' +
    '<button class="gbtn" onclick="gAdd()">Add</button></div>' +
    '<div class="gchips" id="gchips">' + gChips(GSEL, 'gToggle') + '</div>' +
    '<datalist id="gcat">' + G.catalog.map(function (c) { return '<option value="' + esc(c.item) + '">'; }).join('') + '</datalist></div>';
  if (!GSTORE && need.length) {
    if (order.length) {
      h += '<div class="gcard gwhere"><div class="gh">Where to go</div>' +
        order.map(function (k) { return '<div class="gw"><b>' + esc(k) + ':</b> ' + esc(rules[k].map(function (r) { return r.item; }).join(', ')) + '</div>'; }).join('') +
        '<div class="gsmall">Everything else can come from whichever store you go to.</div></div>';
    } else {
      h += '<div class="gcard gwhere"><div class="gsmall" style="margin:0">Nothing store-specific. Any store works.</div></div>';
    }
  }
  h += '<div class="gchips gat"><span class="glbl">Shopping at</span><button class="gchip' + (GSTORE ? '' : ' on') + '" onclick="gAt(-1)">All</button>';
  G.stores.forEach(function (s, i) { h += '<button class="gchip' + (GSTORE === s ? ' on' : '') + '" onclick="gAt(' + i + ')">' + esc(s) + '</button>'; });
  h += '</div>';
  var list = '';
  if (GSTORE) {
    var here = spec.filter(function (r) { return gStores(r.where).indexOf(GSTORE) >= 0; });
    var elsewhere = spec.filter(function (r) { return gStores(r.where).indexOf(GSTORE) < 0; });
    list += gGroup('Get these here', here, true) + gAnywhere(any);
    if (elsewhere.length) list += '<div class="gsec">Not here</div><div class="gsmall gpad">' + esc(elsewhere.map(function (r) { return r.item + ' (' + gStores(r.where).join(' or ') + ')'; }).join(', ')) + '</div>';
  } else {
    order.forEach(function (k) { list += gGroup('Only at ' + k, rules[k], true); });
    list += gAnywhere(any);
  }
  h += list ? '<div class="glist">' + list + '</div>' : '<div class="gcard gsmall">The list is empty.</div>';
  if (G.frequent && G.frequent.length) {
    h += '<div class="gh">Add again</div><div class="gchips" style="margin-bottom:12px">' + G.frequent.map(function (f, i) { return '<button class="gchip" onclick="gQuick(' + i + ')">+ ' + esc(f) + '</button>'; }).join('') + '</div>';
  }
  if (G.got && G.got.length) {
    h += '<button class="glink" onclick="GSHOWGOT=!GSHOWGOT;renderGroceries()">' + (GSHOWGOT ? 'Hide' : 'Show') + ' got recently (' + G.got.length + ')</button>';
    if (GSHOWGOT) {
      h += '<div class="glist">' + G.got.map(function (r, i) {
        return '<div class="gi ggot"><button class="gck on" onclick="gReadd(' + i + ')">✓</button><span class="gn">' + esc(r.item) + '</span><span class="gnote">' + esc((r.gotAt ? r.gotAt + ', ' : '') + gShort(r.gotOn)) + '</span></div>';
      }).join('') + '<div class="gsmall gpad">Tap a checked item to put it back on the list.</div></div>';
    }
  }
  var by = gBy();
  h += '<div class="gchips" style="margin-top:14px"><span class="glbl">Adding as</span>' + (G.people || CFG.people).map(function (n) { return '<button class="gchip' + (by === n ? ' on' : '') + '" onclick="gSetBy(' + gq(n) + ')">' + n + '</button>'; }).join('') + '</div>';
  main.innerHTML = h + '<div style="height:40px"></div>';
  if (focused) { var inp = document.getElementById('gin'); if (inp) { inp.focus(); var n = inp.value.length; inp.setSelectionRange(n, n); } }
}
setInterval(function () { if (TAB === 3 && document.visibilityState === 'visible' && !GBUSY && !GEDIT) loadGroceries(true); }, 30000);
document.addEventListener('visibilitychange', function () { if (TAB === 3 && document.visibilityState === 'visible') loadGroceries(true); });
(function () { try { if (localStorage.getItem('rfa-tab') === '3') setTimeout(function () { showTab(3); }, 0); } catch (e) {} })();
</script>
</body>
</html>
`;


// ---------------------------------------------------------------------------
// Groceries: the shared list (Grocery List tab) with store rules (Grocery Items tab).
// Rows are never deleted. Checking off marks a row Got with the date and store,
// so the history can be analyzed later.
const G_LIST = 'Grocery List';
const G_ITEMS = 'Grocery Items';

function gCtx_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  return { list: ss.getSheetByName(G_LIST), items: ss.getSheetByName(G_ITEMS), tz: tz,
    today: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd') };
}
function gStr_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return String(v == null ? '' : v).trim();
}
function getGroceries() {
  const g = gCtx_();
  const ni = g.items.getLastRow();
  const iv = ni > 1 ? g.items.getRange(2, 1, ni - 1, 6).getValues() : [];
  const stores = iv.map(function (r) { return gStr_(r[5], g.tz); }).filter(String);
  const catalog = iv.filter(function (r) { return gStr_(r[0], g.tz); }).map(function (r) {
    return { item: gStr_(r[0], g.tz), where: gStr_(r[1], g.tz) || 'Anywhere', section: gStr_(r[2], g.tz) };
  });
  const nl = g.list.getLastRow();
  const lv = nl > 1 ? g.list.getRange(2, 1, nl - 1, 9).getValues() : [];
  const cutoff = Utilities.formatDate(new Date(Date.now() - 14 * 864e5), g.tz, 'yyyy-MM-dd');
  const need = [], got = [], counts = {};
  lv.forEach(function (r) {
    const item = gStr_(r[1], g.tz);
    if (!item) return;
    const o = { id: gStr_(r[0], g.tz), item: item, where: gStr_(r[2], g.tz) || 'Anywhere', note: gStr_(r[3], g.tz),
      status: gStr_(r[4], g.tz) || 'Need', addedBy: gStr_(r[5], g.tz), addedOn: gStr_(r[6], g.tz),
      gotOn: gStr_(r[7], g.tz), gotAt: gStr_(r[8], g.tz) };
    if (o.status !== 'Removed') {
      const k = item.toLowerCase();
      if (!counts[k]) counts[k] = { item: item, n: 0 };
      counts[k].n++;
    }
    if (o.status === 'Need') need.push(o);
    else if (o.status === 'Got' && o.gotOn >= cutoff) got.push(o);
  });
  got.sort(function (a, b) { return a.gotOn < b.gotOn ? 1 : a.gotOn > b.gotOn ? -1 : 0; });
  const onList = {};
  need.forEach(function (o) { onList[o.item.toLowerCase()] = true; });
  const frequent = Object.keys(counts).map(function (k) { return counts[k]; })
    .filter(function (c) { return c.n > 1 && !onList[c.item.toLowerCase()]; })
    .sort(function (a, b) { return b.n - a.n; }).slice(0, 12).map(function (c) { return c.item; });
  return { today: g.today, stores: stores, catalog: catalog, need: need, got: got.slice(0, 40),
    frequent: frequent, people: PEOPLE };
}
function gRowOf_(g, id) {
  const n = g.list.getLastRow();
  if (n > 1) {
    const ids = g.list.getRange(2, 1, n - 1, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === String(id)) return i + 2;
  }
  throw new Error('That item changed in the sheet. Refresh and try again.');
}
// Remember an item's store rule on Grocery Items (adds new items, updates changed rules).
function gLearn_(g, item, where) {
  const n = g.items.getLastRow();
  const v = n > 1 ? g.items.getRange(2, 1, n - 1, 2).getValues() : [];
  for (let i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim().toLowerCase() === item.toLowerCase()) {
      if (where && String(v[i][1]).trim() !== where) g.items.getRange(i + 2, 2).setValue(where);
      return String(v[i][0]).trim();
    }
  }
  let r = 2;
  while (r - 2 < v.length && String(v[r - 2][0]).trim()) r++;
  g.items.getRange(r, 1, 1, 2).setValues([[item, where || 'Anywhere']]);
  return item;
}
function addGrocery(item, where, note, by) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    addGrocery_(item, where, note, by);
  } finally {
    lock.releaseLock();
  }
  return getGroceries();
}
// Caller must hold the script lock. `id` lets a phone keep the ID it generated
// offline, so a replayed add lands on the same row instead of duplicating.
function addGrocery_(item, where, note, by, id) {
  item = String(item || '').trim();
  if (!item) throw new Error('Type an item first.');
  where = String(where || '').trim() || 'Anywhere';
  const g = gCtx_();
  const n = g.list.getLastRow();
  const v = n > 1 ? g.list.getRange(2, 1, n - 1, 5).getValues() : [];
  for (let i = 0; i < v.length; i++) {
    if (id && String(v[i][0]) === String(id)) return false;
    if (String(v[i][1]).trim().toLowerCase() === item.toLowerCase() && String(v[i][4]) === 'Need') return false;
  }
  item = gLearn_(g, item, where);
  g.list.getRange(n + 1, 1, 1, 9).setValues([[
    id || ('g' + Date.now()), item, where, note || '', 'Need', by || '', g.today, '', '',
  ]]);
  return true;
}
function setGroceryStatus(id, status, store) {
  setGroceryStatus_(id, status, store);
  return getGroceries();
}
function setGroceryStatus_(id, status, store) {
  if (['Need', 'Got', 'Removed'].indexOf(status) < 0) throw new Error('Unknown status');
  const g = gCtx_();
  const row = gRowOf_(g, id);
  g.list.getRange(row, 5).setValue(status);
  if (status === 'Got') g.list.getRange(row, 8, 1, 2).setValues([[g.today, store || '']]);
  if (status === 'Need') g.list.getRange(row, 8, 1, 2).setValues([['', '']]);
  return true;
}
function setGroceryDetails(id, where, note) {
  setGroceryDetails_(id, where, note);
  return getGroceries();
}
function setGroceryDetails_(id, where, note) {
  const g = gCtx_();
  where = String(where || '').trim() || 'Anywhere';
  const row = gRowOf_(g, id);
  const item = String(g.list.getRange(row, 2).getValue()).trim();
  g.list.getRange(row, 3, 1, 2).setValues([[where, String(note || '')]]);
  gLearn_(g, item, where);
  const n = g.list.getLastRow();
  const v = g.list.getRange(2, 2, n - 1, 4).getValues();
  for (let i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim().toLowerCase() === item.toLowerCase() && String(v[i][3]) === 'Need' && String(v[i][1]) !== where) {
      g.list.getRange(i + 2, 3).setValue(where);
    }
  }
  return true;
}
