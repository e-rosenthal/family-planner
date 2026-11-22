/**
 * Token-protected JSON API for the Family Planner PWA.
 *
 * The PWA on GitHub Pages cannot use google.script.run, so it POSTs JSON here
 * instead. Everything goes through one `sync` call: it applies a batch of
 * queued edits in order, then returns fresh state. That way a phone that has
 * been offline catches up and refreshes in a single round trip.
 *
 * CORS: the request must stay a "simple" request or the browser sends a
 * preflight OPTIONS, which Apps Script web apps cannot answer. So the client
 * posts text/plain with no custom headers, and the token travels in the body.
 *
 * Auth: a shared secret in Script Properties under API_TOKEN. Run
 * setupApiToken() once from the editor to generate it. It is never committed.
 */

const API_TOKEN_KEY = 'API_TOKEN';
const EVENT_ID_COL = 9;        // Events!I - stable IDs, added by ensureEventIds_
const MAX_OPS_PER_SYNC = 25;

/** Run once from the Apps Script editor. Prints the token to paste into each phone. */
function setupApiToken() {
  const props = PropertiesService.getScriptProperties();
  let t = props.getProperty(API_TOKEN_KEY);
  if (!t) {
    t = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    props.setProperty(API_TOKEN_KEY, t);
  }
  ensureEventIds_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EVENTS));
  Logger.log('API token (paste this into the app on each phone):\n\n' + t + '\n');
  return t;
}

/** Throw away the current token and issue a new one. Every phone must be re-paired. */
function rotateApiToken() {
  PropertiesService.getScriptProperties().deleteProperty(API_TOKEN_KEY);
  return setupApiToken();
}

function apiToken_() {
  return PropertiesService.getScriptProperties().getProperty(API_TOKEN_KEY) || '';
}

/** Length-independent compare, so a wrong token leaks nothing through timing. */
function tokenOk_(given) {
  const want = apiToken_();
  if (!want || typeof given !== 'string' || given.length !== want.length) return false;
  let diff = 0;
  for (let i = 0; i < want.length; i++) diff |= want.charCodeAt(i) ^ given.charCodeAt(i);
  return diff === 0;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function newId_(prefix) {
  return prefix + Date.now().toString(36) + Math.floor(Math.random() * 1e6).toString(36);
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

function doPost(e) {
  let req;
  try {
    req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return jsonOut_({ ok: false, error: 'Bad request body.' });
  }
  if (!tokenOk_(req.token)) {
    return jsonOut_({ ok: false, error: 'unauthorized', code: 'unauthorized' });
  }
  try {
    if (req.action === 'ping') {
      return jsonOut_({ ok: true, pong: true, serverTime: new Date().toISOString() });
    }
    return jsonOut_(apiSync_(req));
  } catch (err) {
    return jsonOut_({ ok: false, error: String((err && err.message) || err) });
  }
}

/**
 * Apply queued ops in order, then read back whatever the client asked for.
 * One op failing does not abandon the rest: each gets its own result so the
 * client can drop what succeeded and keep or surface what did not.
 */
function apiSync_(req) {
  const ops = Array.isArray(req.ops) ? req.ops.slice(0, MAX_OPS_PER_SYNC) : [];
  const results = [];

  if (ops.length) {
    const lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try {
      ops.forEach(function (op) {
        try {
          applyOp_(op);
          results.push({ opId: op.opId, ok: true });
        } catch (err) {
          results.push({ opId: op.opId, ok: false, error: String((err && err.message) || err) });
        }
      });
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
  }

  const want = Array.isArray(req.want) ? req.want : ['planner', 'groceries'];
  const out = {
    ok: true,
    results: results,
    queued: Array.isArray(req.ops) ? req.ops.length : 0,
    applied: ops.length,
    serverTime: new Date().toISOString(),
  };
  if (want.indexOf('planner') >= 0) out.planner = getPlannerData();
  if (want.indexOf('groceries') >= 0) out.groceries = getGroceries();
  return out;
}

function applyOp_(op) {
  switch (op && op.kind) {
    case 'setField':
      return saveField(op.date, op.field, op.value == null ? '' : String(op.value));
    case 'addPlan':
      return apiSaveEvent_({ id: op.id, fields: { start: op.date, title: op.text } });
    case 'saveEvent':
      return apiSaveEvent_(op);
    case 'deleteEvent':
      return apiDeleteEvent_(op.id);
    case 'addGrocery':
      return addGrocery_(op.item, op.where, op.note, op.by, op.id);
    case 'setGroceryStatus':
      return setGroceryStatus_(op.id, op.status, op.store);
    case 'setGroceryDetails':
      return setGroceryDetails_(op.id, op.where, op.note);
    default:
      throw new Error('Unknown action: ' + (op && op.kind));
  }
}

// ---------------------------------------------------------------------------
// Stable event IDs
//
// The Events tab was keyed by sheet row, which is fine for a live page but not
// for edits queued on a phone: by the time they sync, the other phone may have
// added or deleted a row and every number below it has moved. Column I holds
// an ID that stays with the event instead.
// ---------------------------------------------------------------------------

function ensureEventIds_(ev) {
  if (String(ev.getRange(1, EVENT_ID_COL).getValue()).trim() !== 'ID') {
    ev.getRange(1, EVENT_ID_COL).setValue('ID');
  }
  const n = ev.getLastRow();
  if (n < 2) return;
  const ids = ev.getRange(2, EVENT_ID_COL, n - 1, 1).getValues();
  const titles = ev.getRange(2, 3, n - 1, 1).getValues();
  let dirty = false;
  for (let i = 0; i < ids.length; i++) {
    if (!String(ids[i][0]).trim() && String(titles[i][0]).trim()) {
      ids[i][0] = newId_('e');
      dirty = true;
    }
  }
  if (dirty) ev.getRange(2, EVENT_ID_COL, n - 1, 1).setValues(ids);
}

function eventRowById_(ev, id) {
  const n = ev.getLastRow();
  if (n < 2 || !id) return 0;
  const ids = ev.getRange(2, EVENT_ID_COL, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

const EVENT_COLS = { start: 1, end: 2, title: 3, category: 4, lead: 5, prep: 6, prepBy: 7, notes: 8 };
const EVENT_DATE_FIELDS = { start: true, end: true, prepBy: true };

/**
 * Create or patch an event by stable ID.
 *
 * Patches touch only the fields present in op.fields, so two phones editing
 * different fields of the same event do not overwrite each other. A create
 * carries the ID the phone generated offline, which makes a retried create
 * land on the same row instead of adding a duplicate.
 */
function apiSaveEvent_(op) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const ev = ss.getSheetByName(EVENTS);
  ensureEventIds_(ev);

  const fields = op.fields || {};
  const id = String(op.id || '').trim();
  const row = eventRowById_(ev, id);

  if (!row) {
    const title = String(fields.title || '').trim();
    if (!fields.start || !title) throw new Error('An event needs a start date and a name.');
    const at = ev.getLastRow() + 1;
    ev.getRange(at, 1, 1, 8).setValues([[
      toDate_(fields.start, tz),
      toDate_(fields.end || '', tz),
      title,
      fields.category || '',
      fields.lead || '',
      fields.prep || '',
      toDate_(fields.prepBy || '', tz),
      fields.notes || '',
    ]]);
    ev.getRange(at, EVENT_ID_COL).setValue(id || newId_('e'));
    return true;
  }

  Object.keys(fields).forEach(function (k) {
    const col = EVENT_COLS[k];
    if (!col) return;
    const raw = fields[k];
    const value = EVENT_DATE_FIELDS[k]
      ? toDate_(raw || '', tz)
      : String(raw == null ? '' : raw);
    ev.getRange(row, col).setValue(value);
  });
  return true;
}

/** Deleting something already gone is a success, so a replayed op stays harmless. */
function apiDeleteEvent_(id) {
  const ev = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EVENTS);
  ensureEventIds_(ev);
  const row = eventRowById_(ev, id);
  if (!row) return true;
  ev.deleteRow(row);
  return true;
}
