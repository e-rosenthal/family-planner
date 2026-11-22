/**
 * Family Planner - offline-capable PWA.
 *
 * Ported from the Apps Script HtmlService page. google.script.run is replaced
 * by one POST endpoint, and every edit becomes a small queued action rather
 * than a whole-row write, so two phones editing different things never clobber
 * each other and anything typed with no signal syncs later.
 */
'use strict';

var CFG_KEY = 'rfa-cfg';
var QUEUE_KEY = 'rfa-queue';
var CACHE_KEY = 'rfa-cache';
var GROC_KEY = 'rfa-groc';

function lsGet(k) { try { var s = localStorage.getItem(k); return s ? JSON.parse(s) : null; } catch (e) { return null; } }
function lsPut(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } }

/**
 * Whether this browser will actually keep what we write. Private browsing and
 * "block all cookies" both make localStorage throw or silently forget, which
 * would otherwise show up only as the pairing screen returning every launch.
 */
function storageWorks() {
  try {
    localStorage.setItem('rfa-probe', '1');
    var ok = localStorage.getItem('rfa-probe') === '1';
    localStorage.removeItem('rfa-probe');
    return ok;
  } catch (e) { return false; }
}
var STORAGE_OK = storageWorks();
var STORAGE_NOTE = 'This browser is not letting the app save anything, so it has to ask every time. ' +
  'That is usually Private Browsing, or Settings > Safari > Block All Cookies. ' +
  'Turn those off, or open the app from its Home Screen icon rather than in the browser.';

var CFG = lsGet(CFG_KEY);              // { url, token }
var QUEUE = lsGet(QUEUE_KEY) || [];    // edits waiting to reach the sheet
var SETUP = !CFG;                      // showing the pairing screen
var SETUP_ERR = '';
var SETUP_BUSY = false;
var SYNCING = false;
var SYNCERR = '';
var AUTHFAIL = false;

function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }
function esc(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function toast(msg) {
  var t = document.getElementById('toast');
  t.textContent = msg; t.className = 'show';
  clearTimeout(toast.h); toast.h = setTimeout(function () { t.className = ''; }, 2200);
}

// ---------------------------------------------------------------------------
// Transport
// ---------------------------------------------------------------------------

/**
 * Apps Script web apps cannot answer a CORS preflight, so this has to stay a
 * "simple" request: text/plain, no custom headers. The token rides in the body
 * rather than an Authorization header for the same reason.
 */
function post(body) {
  return fetch(CFG.url, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(body),
    redirect: 'follow',
  }).then(function (r) {
    if (!r.ok) throw new Error('The script returned HTTP ' + r.status + '.');
    return r.text();
  }).then(function (text) {
    var d;
    try { d = JSON.parse(text); } catch (e) {
      throw new Error('That URL did not return JSON. Check it ends in /exec.');
    }
    if (!d.ok && d.code === 'unauthorized') {
      var err = new Error('The access code was rejected.');
      err.unauthorized = true;
      throw err;
    }
    if (!d.ok) throw new Error(d.error || 'Unknown error from the script.');
    return d;
  });
}

// ---------------------------------------------------------------------------
// Edit queue
// ---------------------------------------------------------------------------

function saveQueue() { lsPut(QUEUE_KEY, QUEUE); }

/**
 * Add an edit to the queue, folding it into a pending one where that is safe.
 * Setting the same cell twice offline should reach the sheet once, and editing
 * a new event before it has ever synced should still be a single create.
 */
function enqueue(op) {
  op.opId = uid();
  if (op.kind === 'setField') {
    QUEUE = QUEUE.filter(function (q) {
      return !(q.kind === 'setField' && q.date === op.date && q.field === op.field);
    });
  } else if (op.kind === 'setGroceryStatus' || op.kind === 'setGroceryDetails') {
    QUEUE = QUEUE.filter(function (q) { return !(q.kind === op.kind && q.id === op.id); });
  } else if (op.kind === 'saveEvent') {
    for (var i = 0; i < QUEUE.length; i++) {
      if (QUEUE[i].kind === 'saveEvent' && QUEUE[i].id === op.id) {
        Object.keys(op.fields).forEach(function (k) { QUEUE[i].fields[k] = op.fields[k]; });
        saveQueue(); flush(); return;
      }
    }
  } else if (op.kind === 'deleteEvent') {
    QUEUE = QUEUE.filter(function (q) { return !(q.kind === 'saveEvent' && q.id === op.id); });
  }
  QUEUE.push(op);
  saveQueue();
  flush();
}

function wants() {
  var w = ['planner'];
  if (G || TAB === 3) w.push('groceries');
  return w;
}

/** Send whatever is queued, then take the server's state as the new truth. */
function flush(manual) {
  if (!CFG || SETUP) return Promise.resolve();
  if (SYNCING) return Promise.resolve();
  if (!navigator.onLine && !manual) { paintWifi(); return Promise.resolve(); }

  var sending = QUEUE.slice(0, 25);
  SYNCING = true;
  paintWifi();

  return post({ token: CFG.token, action: 'sync', ops: sending, want: wants() })
    .then(function (d) {
      var handled = {}, errs = [];
      (d.results || []).forEach(function (r) {
        handled[r.opId] = true;
        if (!r.ok) errs.push(r.error);
      });
      // Drop both the applied and the permanently rejected: a rejected edit
      // will not start working on the next try, and retrying it forever would
      // block everything queued behind it.
      QUEUE = QUEUE.filter(function (q) { return !handled[q.opId]; });
      saveQueue();

      if (d.planner) { DATA = d.planner; applySchema(DATA.schema); lsPut(CACHE_KEY, DATA); }
      if (d.groceries) { G = d.groceries; lsPut(GROC_KEY, G); }
      SYNCERR = ''; AUTHFAIL = false; SYNCING = false;
      render();
      if (errs.length) toast(errs.length === 1 ? errs[0] : errs.length + ' edits could not be saved');
      if (QUEUE.length) return flush();   // more than one batch was waiting
    })
    .catch(function (e) {
      SYNCING = false;
      AUTHFAIL = !!e.unauthorized;
      SYNCERR = e.message || 'Could not reach the planner.';
      if (!DATA && !G) render(); else paintWifi();
    });
}

function paintWifi() {
  var bar = document.getElementById('wifi');
  var text = document.getElementById('wifitext');
  var btn = document.getElementById('wifibtn');
  if (!bar) return;
  var n = QUEUE.length;
  var cls = 'wifi', msg = '', show = true, label = 'Retry';

  if (AUTHFAIL) {
    cls += ' err'; msg = 'Access code rejected'; label = 'Settings';
  } else if (!navigator.onLine) {
    cls += ' off';
    msg = n ? n + (n === 1 ? ' edit' : ' edits') + ' saved on this phone' : 'Offline - showing the last update';
    show = true; label = 'Try now';
  } else if (SYNCING) {
    msg = n ? 'Syncing ' + n + (n === 1 ? ' edit' : ' edits') + '...' : 'Refreshing...';
    label = '';
  } else if (SYNCERR) {
    cls += ' err';
    msg = n ? n + ' waiting - ' + SYNCERR : SYNCERR;
  } else if (n) {
    msg = n + (n === 1 ? ' edit' : ' edits') + ' waiting';
  } else {
    show = false;
  }

  bar.className = cls;
  bar.hidden = !show;
  text.textContent = msg;
  btn.hidden = !label;
  btn.textContent = label || 'Retry';
  btn.onclick = AUTHFAIL ? openSetup : function () { flush(true); };
}

// ---------------------------------------------------------------------------
// Pairing screen - the script URL and access code live only on the phone
// ---------------------------------------------------------------------------

function openSetup() { SETUP = true; SETUP_ERR = ''; render(); }

function setupView() {
  var cur = CFG || { url: '', token: '' };
  return '<div class="setup">' +
    '<h2>Connect this phone</h2>' +
    '<p>These two values stay on this phone. They are not in the public code, so each phone is set up once.</p>' +
    (STORAGE_OK ? '' : '<div class="err">' + esc(STORAGE_NOTE) + '</div>') +
    '<label for="s_url">Web app URL</label>' +
    '<input id="s_url" type="url" inputmode="url" autocapitalize="off" autocorrect="off" spellcheck="false" ' +
    'placeholder="https://script.google.com/macros/s/.../exec" value="' + esc(cur.url) + '">' +
    '<label for="s_tok">Access code</label>' +
    '<input id="s_tok" type="text" autocapitalize="off" autocorrect="off" spellcheck="false" ' +
    'placeholder="Paste the code from the setup step" value="' + esc(cur.token) + '">' +
    (SETUP_ERR ? '<div class="err">' + esc(SETUP_ERR) + '</div>' : '') +
    '<div class="actions" style="margin-top:14px">' +
    '<button class="btn" onclick="saveSetup()"' + (SETUP_BUSY ? ' disabled' : '') + '>' +
    (SETUP_BUSY ? 'Checking...' : 'Connect') + '</button>' +
    (CFG ? '<button class="btn ghost" onclick="SETUP=false;render()">Cancel</button>' : '') +
    '</div>' +
    (QUEUE.length ? '<p style="margin-top:16px"><b>' + QUEUE.length + '</b> edit' +
      (QUEUE.length === 1 ? '' : 's') + ' made on this phone are still waiting to sync.</p>' : '') +
    '</div>';
}

function saveSetup() {
  var url = document.getElementById('s_url').value.trim();
  var token = document.getElementById('s_tok').value.trim();
  if (!url || !token) { SETUP_ERR = 'Both boxes are needed.'; render(); return; }
  if (!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec/.test(url)) {
    SETUP_ERR = 'That does not look like a web app URL. It should end in /exec.';
    render(); return;
  }
  var prev = CFG;
  CFG = { url: url, token: token };
  SETUP_BUSY = true; SETUP_ERR = ''; render();
  post({ token: token, action: 'ping' })
    .then(function () {
      lsPut(CFG_KEY, CFG);
      SETUP_BUSY = false;
      // Read it back: the connection worked, but if the phone will not keep
      // the settings, say so here rather than silently asking again later.
      var stored = lsGet(CFG_KEY);
      if (!stored || stored.token !== CFG.token) {
        STORAGE_OK = false;
        SETUP_ERR = STORAGE_NOTE;
        render();
        return;
      }
      SETUP = false; AUTHFAIL = false; SYNCERR = '';
      render();
      load();
    })
    .catch(function (e) {
      CFG = prev;
      SETUP_BUSY = false;
      SETUP_ERR = e.unauthorized
        ? 'The access code was not accepted. Check it for stray spaces.'
        : (e.message || 'Could not reach that URL.');
      render();
    });
}

// ---------------------------------------------------------------------------
// Planner state
// ---------------------------------------------------------------------------

var DATA = lsGet(CACHE_KEY);
var TAB = 0;
var EDIT = null;        // event id being edited, or 'new'
var SHOWALL = false;
var VIEW = 'table';
var AUTO_SCROLL = true;
// The field schema is family-specific, so it comes from the server with the
// data rather than being baked into this public file.
var ORDER = [], LABELS = {}, SCHOOL = [], EVERYDAY = [], TEXT = [];
function applySchema(sc) {
  if (!sc) return;
  ORDER = sc.order || []; LABELS = sc.labels || {};
  SCHOOL = sc.weekdayOnly || []; EVERYDAY = sc.everyday || []; TEXT = sc.text || [];
  if (sc.title) {
    document.title = sc.title;
    var h = document.querySelector('header h1');
    if (h) h.textContent = '\uD83D\uDC68\u200D\uD83D\uDC69\u200D\uD83D\uDC67\u200D\uD83D\uDC66 ' + sc.title;
  }
}
var MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
var DOW = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

/** Same shape as the sheet's 'EEE MMM d', so offline rows match synced ones. */
function niceDate(ds) {
  if (!ds) return '';
  var p = ds.split('-');
  var d = new Date(+p[0], +p[1] - 1, +p[2]);
  return DOW[d.getDay()] + ' ' + MON[d.getMonth()] + ' ' + d.getDate();
}
function addDays(ds, n) {
  var p = ds.split('-');
  var d = new Date(+p[0], +p[1] - 1, +p[2] + n);
  var m = d.getMonth() + 1, dd = d.getDate();
  return d.getFullYear() + '-' + (m < 10 ? '0' : '') + m + '-' + (dd < 10 ? '0' : '') + dd;
}

function load() { return flush(true); }

/** Build the client-side twin of a server event row. */
function mkEvent(id, f) {
  var end = f.end && f.end !== f.start ? f.end : '';
  return {
    id: id,
    row: null,
    start: f.start,
    end: end,
    lastDay: end || f.start,
    label: niceDate(f.start) + (end ? ' to ' + niceDate(end) : ''),
    title: String(f.title || '').trim(),
    category: f.category || '',
    lead: f.lead || '',
    prep: f.prep || '',
    prepBy: f.prepBy || '',
    prepByLabel: f.prepBy ? niceDate(f.prepBy) : '',
    notes: f.notes || '',
    pending: true,
  };
}

/**
 * Recompute each day's plans and prep from the event list, mirroring what
 * getPlannerData does on the server, so an event added offline shows up on the
 * right days straight away.
 */
function rebuildDays() {
  if (!DATA || !DATA.days) return;
  DATA.days.forEach(function (day) {
    day.plans = DATA.events.filter(function (e) {
      return e.start <= day.date && e.lastDay >= day.date;
    }).map(function (e) { return e.title; });
    day.prep = DATA.events.filter(function (e) {
      return e.prepBy === day.date && e.prep;
    }).map(function (e) { return e.prep + ' (' + e.title + ')'; });
  });
  DATA.events.sort(function (a, b) { return a.start < b.start ? -1 : a.start > b.start ? 1 : 0; });
  lsPut(CACHE_KEY, DATA);
}

// ---------------------------------------------------------------------------
// Day cards and the week table
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// Events
// ---------------------------------------------------------------------------

// Quote an id for an inline onclick inside a double-quoted HTML attribute.
function q(s) {
  return "'" + String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'")
    .replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;') + "'";
}
function opts(list, cur) {
  var l = (list || []).slice();
  if (cur && l.indexOf(cur) < 0) l.unshift(cur);
  return '<option value="">None</option>' + l.map(function (o) {
    return '<option value="' + esc(o) + '"' + (o === cur ? ' selected' : '') + '>' + esc(o) + '</option>';
  }).join('');
}
function editor(e) {
  e = e || { id: null, title: '', start: DATA.today, end: '', category: '', lead: '', prep: '', prepBy: '', notes: '' };
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
  if (e.id) h += '<button class="btn danger" onclick="delEv()">Delete</button>';
  return h + '</div></div>';
}
function findEvent(id) { return DATA.events.filter(function (e) { return e.id === id; })[0]; }
function eventRow(e, soonLimit) {
  var soon = e.start < soonLimit;
  var sub = [];
  if (e.category) sub.push(e.category);
  if (e.lead) sub.push(e.lead);
  if (e.prep) sub.push('Prep: ' + e.prep + (e.prepByLabel ? ' by ' + e.prepByLabel : ''));
  var h = '<div class="li tap' + (soon ? ' soon' : '') + '" onclick="editEv(' + q(e.id) + ')"><span class="d">' +
    esc(e.label) + '</span><span class="t">' + esc(e.title) +
    (sub.length ? '<span class="s">' + esc(sub.join('  |  ')) + '</span>' : '') + '</span></div>';
  if (EDIT === e.id) h += editor(e);
  return h;
}
function eventsView() {
  var in14 = addDays(DATA.today, 14);
  var horizon = addDays(DATA.today, 84);
  var h = '';
  var prep = DATA.events.filter(function (e) { return e.prep && e.prepBy && e.prepBy >= DATA.today && e.prepBy < in14; })
    .sort(function (a, b) { return a.prepBy < b.prepBy ? -1 : 1; });
  h += '<h2>Prep due in the next 2 weeks</h2><div class="list">' + (prep.length ? prep.map(function (e) {
    return '<div class="li tap" onclick="editEv(' + q(e.id) + ')"><span class="d">' + esc(e.prepByLabel) +
      '</span><span class="t">' + esc(e.prep) + '<span class="s">for ' + esc(e.title) +
      (e.lead ? ', ' + esc(e.lead) : '') + '</span></span></div>';
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
function editEv(id) { EDIT = (EDIT === id ? null : id); render(); }
function newEv() { EDIT = 'new'; render(); }
function cancelEv() { EDIT = null; render(); }
function val(id) { return document.getElementById(id).value.trim(); }

function saveEv() {
  var f = {
    title: val('e_title'), start: val('e_start'), end: val('e_end'), category: val('e_cat'),
    lead: val('e_lead'), prep: val('e_prep'), prepBy: val('e_prepby'), notes: val('e_notes'),
  };
  if (!f.title || !f.start) { toast('Add a name and a start date'); return; }
  if (f.end && f.end < f.start) { toast('End date is before the start'); return; }

  if (EDIT === 'new') {
    var id = 'e' + uid();
    enqueue({ kind: 'saveEvent', id: id, fields: f });
    DATA.events.push(mkEvent(id, f));
  } else {
    var cur = findEvent(EDIT);
    if (!cur) { EDIT = null; render(); return; }
    // Send only what actually changed, so the other phone's edit to a
    // different field of the same event survives.
    var patch = {};
    Object.keys(f).forEach(function (k) {
      if (String(cur[k] == null ? '' : cur[k]) !== f[k]) patch[k] = f[k];
    });
    if (!Object.keys(patch).length) { EDIT = null; render(); return; }
    enqueue({ kind: 'saveEvent', id: EDIT, fields: patch });
    var merged = mkEvent(EDIT, {
      start: f.start, end: f.end, title: f.title, category: f.category,
      lead: f.lead, prep: f.prep, prepBy: f.prepBy, notes: f.notes,
    });
    DATA.events[DATA.events.indexOf(cur)] = merged;
  }
  rebuildDays();
  EDIT = null;
  toast(navigator.onLine ? 'Saved' : 'Saved on this phone');
  render();
}

function delEv() {
  var e = findEvent(EDIT);
  if (!e) { EDIT = null; render(); return; }
  if (!window.confirm('Delete "' + e.title + '"?')) return;
  enqueue({ kind: 'deleteEvent', id: e.id });
  DATA.events.splice(DATA.events.indexOf(e), 1);
  rebuildDays();
  EDIT = null;
  toast('Deleted');
  render();
}

// ---------------------------------------------------------------------------
// Render and save
// ---------------------------------------------------------------------------

function showTab(t) {
  TAB = t; EDIT = null; AUTO_SCROLL = true;
  try { localStorage.setItem('rfa-tab', String(t)); } catch (e) {}
  for (var i = 0; i < 4; i++) document.getElementById('t' + i).className = 'tab' + (i === t ? ' on' : '');
  render();
  if (t === 3) flush();
  window.scrollTo(0, 0);
}

function render() {
  paintWifi();
  var main = document.getElementById('main');
  if (SETUP) {
    document.getElementById('sub').textContent = 'Setup';
    main.className = '';
    main.innerHTML = setupView();
    return;
  }
  if (TAB === 3) { renderGroceries(); return; }
  if (!DATA) {
    document.getElementById('sub').textContent = SYNCERR ? 'Could not load' : 'Loading...';
    main.className = '';
    main.innerHTML = '<div class="loading">' +
      (SYNCERR ? esc(SYNCERR) + '<br><br><button class="btn" onclick="load()">Try again</button>' +
        ' <button class="btn ghost" onclick="openSetup()">Settings</button>'
        : 'Loading the planner...') + '</div>';
    return;
  }
  var c = openCount();
  var todayDay = DATA.days.filter(function (d) { return d.date === DATA.today; })[0];
  document.getElementById('sub').textContent = (todayDay ? todayDay.label + '  |  ' : '') +
    (c === 0 ? 'All covered for 2 weeks' : c + ' open slot' + (c === 1 ? '' : 's') + ' in the next 2 weeks');

  var h = TAB === 2 ? eventsView() : weekView(TAB);
  var wrap = document.querySelector('.tablewrap');
  var sx = wrap ? wrap.scrollLeft : 0;
  var sy = window.scrollY || 0;
  main.className = (TAB !== 2 && VIEW === 'table') ? 'wide' : '';
  main.innerHTML = h + '<div style="height:40px"></div>';
  var wrap2 = document.querySelector('.tablewrap');
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
  window.scrollTo(0, sy);
}

function findDay(date) { return DATA.days.filter(function (d) { return d.date === date; })[0]; }

function save(id, value) {
  var parts = id.split('|');
  var day = findDay(parts[0]);
  var f = parts.slice(1).join('|');
  day.fields[f].set = value;
  lsPut(CACHE_KEY, DATA);
  enqueue({ kind: 'setField', date: parts[0], field: f, value: value });
  render();
  toast(value === '' ? 'Back to usual' : (navigator.onLine ? 'Saved' : 'Saved on this phone'));
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
  var id = 'e' + uid();
  enqueue({ kind: 'addPlan', id: id, date: date, text: text });
  DATA.events.push(mkEvent(id, { start: date, title: text }));
  rebuildDays();
  render();
  toast(navigator.onLine ? 'Plan added' : 'Plan saved on this phone');
}

// ---------------------------------------------------------------------------
// Groceries
// ---------------------------------------------------------------------------

var G = lsGet(GROC_KEY);
var GSTORE = '', GSEL = [], GTOUCH = false, GDRAFT = '', GEDIT = null, GESEL = [], GSHOWGOT = false;
var GSECTIONS = ['Produce', 'Dairy & eggs', 'Meat & fish', 'Bakery', 'Pantry', 'Frozen', 'Household', 'Other'];

function gBy() { try { return localStorage.getItem('rfa-by') || ''; } catch (e) { return ''; } }
function gSetBy(n) { try { localStorage.setItem('rfa-by', n); } catch (e) {} renderGroceries(); }
function gCachePut(d) { lsPut(GROC_KEY, d); }
function gStores(w) {
  if (!w || w === 'Anywhere') return [];
  return String(w).split(',').map(function (s) { return s.trim(); }).filter(Boolean);
}
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
  return MON[+p[1] - 1] + ' ' + (+p[2]);
}
/** Apply the change locally, queue it, and let the next sync confirm. */
function gSave(op, optimistic) {
  if (optimistic) optimistic();
  gCachePut(G);
  enqueue(op);
  renderGroceries();
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
  G.stores.forEach(function (s, i) {
    h += '<button class="gchip spec' + (sel.indexOf(s) >= 0 ? ' on' : '') + '" onclick="' + fn + '(' + i + ')">' + esc(s) + '</button>';
  });
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
  // The phone picks the row ID so a replayed add cannot create a second row.
  var id = 'g' + uid();
  gSave({ kind: 'addGrocery', id: id, item: item, where: where, note: '', by: by }, function () {
    G.need.push({ id: id, item: item, where: where, note: '', addedBy: by, status: 'Need', pending: true });
    if (!c) G.catalog.push({ item: item, where: where, section: '' });
  });
}
function gQuick(i) { gAdd(G.frequent[i]); }
function gCheck(id) {
  var k = gFind(id); if (k < 0) return;
  var r = G.need[k], at = GSTORE;
  gSave({ kind: 'setGroceryStatus', id: id, status: 'Got', store: at }, function () {
    G.need.splice(k, 1); r.status = 'Got'; r.gotOn = gToday(); r.gotAt = at; G.got.unshift(r);
  });
}
function gReadd(i) {
  var r = G.got[i]; if (!r) return;
  if (r.gotOn === gToday()) {
    gSave({ kind: 'setGroceryStatus', id: r.id, status: 'Need', store: '' }, function () {
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
  gSave({ kind: 'setGroceryDetails', id: id, where: where, note: note }, function () {
    G.need.forEach(function (r) { if (r.item === item) r.where = where; });
    G.need[k].note = note;
    var c = gCat(item); if (c) c.where = where;
  });
}
function gRemove() {
  var id = GEDIT, k = gFind(id); if (k < 0) return;
  GEDIT = null;
  gSave({ kind: 'setGroceryStatus', id: id, status: 'Removed', store: '' }, function () { G.need.splice(k, 1); });
}
function gRow(r) {
  var s = gStores(r.where);
  var h = '<div class="gi' + (s.length ? ' gspec' : '') + (r.pending ? ' gpend' : '') + '">' +
    '<button class="gck" aria-label="Got it" onclick="gCheck(' + q(r.id) + ')"></button>' +
    '<span class="gn" onclick="gEdit(' + q(r.id) + ')">' + esc(r.item) +
    (r.note ? '<span class="gnote">' + esc(r.note) + '</span>' : '') + '</span>' +
    (s.length ? '<span class="gtag">' + esc(s.join(' / ')) + '</span>' : '');
  if (GEDIT === r.id) {
    h += '<div class="gedit"><div class="gchips" id="gechips">' + gChips(GESEL, 'gEToggle') + '</div>' +
      '<input id="gnote" placeholder="Note (brand, size, how many)" value="' + esc(r.note) + '">' +
      '<div class="gacts"><button class="gbtn" onclick="gEditSave()">Save</button>' +
      '<button class="glink" onclick="gEdit(' + q(r.id) + ')">Cancel</button>' +
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
  if (!G) {
    main.innerHTML = '<div class="loading">' +
      (SYNCERR ? esc(SYNCERR) + '<br><br><button class="btn" onclick="load()">Try again</button>'
        : 'Loading groceries...') + '</div>';
    return;
  }
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
  document.getElementById('sub').textContent = need.length + ' to get' + (GSTORE ? '  |  Shopping at ' + GSTORE : '');
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
        return '<div class="gi ggot"><button class="gck on" onclick="gReadd(' + i + ')">&#10003;</button><span class="gn">' + esc(r.item) + '</span><span class="gnote">' + esc((r.gotAt ? r.gotAt + ', ' : '') + gShort(r.gotOn)) + '</span></div>';
      }).join('') + '<div class="gsmall gpad">Tap a checked item to put it back on the list.</div></div>';
    }
  }
  var by = gBy();
  h += '<div class="gchips" style="margin-top:14px"><span class="glbl">Adding as</span>' + ((G.people || []).map(function (n) { return '<button class="gchip' + (by === n ? ' on' : '') + '" onclick="gSetBy(' + q(n) + ')">' + esc(n) + '</button>'; })).join('') + '</div>';
  h += '<div style="margin-top:18px;text-align:center"><button class="glink" onclick="openSetup()">Settings</button></div>';
  main.innerHTML = h + '<div style="height:40px"></div>';
  if (focused) { var inp = document.getElementById('gin'); if (inp) { inp.focus(); var n = inp.value.length; inp.setSelectionRange(n, n); } }
}

// ---------------------------------------------------------------------------
// Start up
// ---------------------------------------------------------------------------

(function boot() {
  try {
    var t = localStorage.getItem('rfa-tab');
    if (t === '3') { TAB = 3; for (var i = 0; i < 4; i++) document.getElementById('t' + i).className = 'tab' + (i === 3 ? ' on' : ''); }
  } catch (e) {}

  // Ask the browser not to evict the queue if storage gets tight.
  if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(function () {});

  if (DATA) applySchema(DATA.schema);
  render();
  if (CFG) flush();

  window.addEventListener('online', function () { paintWifi(); flush(); });
  window.addEventListener('offline', paintWifi);
  document.addEventListener('visibilitychange', function () {
    if (document.visibilityState === 'visible') flush();
  });
  setInterval(function () {
    if (document.visibilityState === 'visible' && !GEDIT && EDIT === null) flush();
  }, 30000);

  if ('serviceWorker' in navigator) {
    window.addEventListener('load', function () {
      navigator.serviceWorker.register('./sw.js').catch(function () {});
    });
  }
})();
