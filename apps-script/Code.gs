/**
 * יומן התינוקת – גיבוי וסנכרון בגוגל שיטס
 * מדביקים את כל הקובץ הזה ב-Extensions ← Apps Script של הגיליון,
 * ומפרסמים כ-Web app (Execute as: Me, Who has access: Anyone).
 * לא צריך לשנות כלום בקוד.
 *
 * חשוב: לא למיין ולא למחוק שורות בלשונית "יומן התינוקת".
 * לקריאה נוחה יש לשונית "טבלה מסודרת": רק הגרסה האחרונה של כל רישום,
 * בלי מחוקים, לפי סדר הזמן. היא נבנית מחדש בכל סנכרון, אז לא כותבים בה.
 */
var TAB = 'יומן התינוקת';
var PHOTO_FOLDER = 'יומן התינוקת – תמונות';
var CLEAN_TAB = 'טבלה מסודרת';
var HEAD = ['מזהה','עודכן','מצב','תאריך','שעה','סוג','צד','משך (דק׳)','כמות (מ״ל)','סוג חלב','חיתול','צבע','איך הייתה','פריט / משקל','תגיות','הערה','תמונות','נתונים מלאים (לשחזור)'];

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TAB);
  if (!sh) {
    sh = ss.insertSheet(TAB);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function folder_() {
  var it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
}

function photoFile_(id) {
  var it = folder_().getFilesByName(id + '.jpg');
  return it.hasNext() ? it.next() : null;
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'photo') {
      var f = photoFile_(String(p.id || ''));
      if (!f) return json_({ ok: false, error: 'not_found' });
      return json_({ ok: true, data: 'data:image/jpeg;base64,' + Utilities.base64Encode(f.getBlob().getBytes()) });
    }
    var sh = sheet_();
    var total = Math.max(0, sh.getLastRow() - 1);
    if (p.action === 'ping') return json_({ ok: true, total: total });
    if (p.action === 'nstatus') return json_({ ok: true, total: total, last: +(PropertiesService.getScriptProperties().getProperty('nlast') || 0), trigger: hasTickTrigger_() });
    var since = Math.max(0, parseInt(p.since || '0', 10) || 0);
    if (since >= total) return json_({ ok: true, total: total, rows: [] });
    var vals = sh.getRange(2 + since, HEAD.length, total - since, 1).getValues();
    return json_({ ok: true, total: total, rows: vals.map(function (r) { return String(r[0] || ''); }) });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var b = JSON.parse(e.postData.contents);
    if (b.action === 'snooze' || b.action === 'done') { nAction_(b); return json_({ ok: true }); }
    if (b.action === 'append') {
      var sh = sheet_();
      var rows = (b.rows || []).map(function (r) {
        var row = r.slice(0, HEAD.length);
        while (row.length < HEAD.length) row.push('');
        return row.map(function (c) { return c === null || c === undefined ? '' : c; });
      });
      if (rows.length) {
        var start = sh.getLastRow() + 1;
        // keep the JSON column as plain text so Sheets never reformats it
        sh.getRange(start, HEAD.length, rows.length, 1).setNumberFormat('@');
        sh.getRange(start, 1, rows.length, HEAD.length).setValues(rows);
        try { rebuildClean_(sh); } catch (err) {}
        try { ingest_(rows.map(function (r) { return r[HEAD.length - 1]; })); } catch (err) {}
      }
      return json_({ ok: true, total: Math.max(0, sh.getLastRow() - 1) });
    }
    if (b.action === 'photo') {
      var id = String(b.id || '').replace(/[^A-Za-z0-9_-]/g, '');
      var data = String(b.data || '');
      if (!id || data.indexOf('base64,') < 0) return json_({ ok: false, error: 'bad_photo' });
      if (!photoFile_(id)) {
        var bytes = Utilities.base64Decode(data.split('base64,')[1]);
        folder_().createFile(Utilities.newBlob(bytes, 'image/jpeg', id + '.jpg'));
      }
      return json_({ ok: true });
    }
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* builds the readable tab: latest version of each record, without deleted ones, oldest first */
function rebuildClean_(sh) {
  var n = sh.getLastRow() - 1;
  if (n < 1) return;
  var vals = sh.getRange(2, 1, n, HEAD.length).getValues();
  var last = {};
  vals.forEach(function (r, i) { if (r[0]) last[String(r[0])] = i; });
  var out = [];
  Object.keys(last).forEach(function (id) {
    var r = vals[last[id]];
    if (id.indexOf('q:') === 0 || r[5] === 'הגדרות' || r[5] === 'טיימר') return;
    if (r[2] === 'נמחק' || r[2] === 'נמחקה') return;
    var t = 0;
    try { t = JSON.parse(r[HEAD.length - 1]).d.t || 0; } catch (e) {}
    out.push({ t: t, row: r.slice(3, 17) });
  });
  out.sort(function (a, b) { return a.t - b.t; });
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cl = ss.getSheetByName(CLEAN_TAB);
  if (!cl) { cl = ss.insertSheet(CLEAN_TAB); cl.setRightToLeft(true); }
  cl.clearContents();
  var head = HEAD.slice(3, 17);
  cl.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  cl.setFrozenRows(1);
  if (out.length) {
    cl.getRange(2, 1, out.length, head.length).setNumberFormat('@')
      .setValues(out.map(function (o) { return o.row.map(function (c) { return c === null || c === undefined ? '' : String(c); }); }));
  }
}

/* run once by hand (Run ← rebuildClean) to create the readable tab right away */
function rebuildClean() { rebuildClean_(sheet_()); }

/* ================= notifications to a locked phone (via the free ntfy app) =================
 * One-time setup: paste this file, then choose setupNotifications in the menu above and press Run
 * (approve the permissions), then Deploy ← Manage deployments ← edit ← New version.
 * A trigger then runs notifyTick every minute. It reads only a small saved summary, not the sheet. */
var NTFY = 'https://ntfy.sh/';

function setupNotifications() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'notifyTick') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('notifyTick').timeBased().everyMinutes(1).create();
  rebuildNotifyState_();
  notifyTick();
  return 'ok';
}
function hasTickTrigger_() {
  try { return ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'notifyTick'; }); } catch (e) { return null; }
}
function nget_(k, d) { try { var v = PropertiesService.getScriptProperties().getProperty(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } }
function nput_(k, v) { PropertiesService.getScriptProperties().setProperty(k, JSON.stringify(v)); }

/* keeps: latest settings, whether a feeding timer runs, and the last ~2 days of feeds / water / medicine */
function ingest_(jsons) {
  var set = nget_('nset', {}), ev = nget_('nev', {}), tim = nget_('ntim', { u: 0, feed: false }), now = Date.now(), changed = false;
  jsons.forEach(function (j) {
    var r; try { r = JSON.parse(j); } catch (e) { return; }
    if (!r || !r.d) return;
    var u = r.u || 0, d = r.d;
    if (r.k === 'set' && u >= (set.u || 0)) {
      set = { u: u, every: +d.feedEvery || 0, leads: (Array.isArray(d.notifyLeads) ? d.notifyLeads : [d.notifyLead == null ? 10 : d.notifyLead]).map(Number).filter(function (v) { return v > 0; }), end: d.notifyEnd !== false, due: d.notifyDue !== false, water: d.notifyWater !== false,
        pain: d.notifyPain === true, gap: d.momGap || {}, app: d.appUrl || '', meds: d.notifyMeds !== false,
        topic: d.ntfy || '', tz: d.tz || 'Asia/Jerusalem', w: d.water || {},
        daily: (d.daily || []).filter(function (x) { return x.notify !== false; }).map(function (x) { return { id: x.id, l: x.label, t: x.time, m: x.who === 'mom' ? 1 : 0 }; }) };
      changed = true;
    } else if (r.k === 'tim' && u >= (tim.u || 0)) { tim = { u: u, feed: !!d.feed }; changed = true; }
    else if (r.k === 'ev' && /^(feed|bottle|water|med|mmed)$/.test(d.type) && d.t > now - 2 * 864e5) {
      var cur = ev[d.id];
      if (!cur || (cur.u || 0) <= u) {
        ev[d.id] = { y: d.type, t: d.t, e: d.end || (d.dur ? d.t + d.dur * 6e4 : d.t), n: d.name || '', c: +d.n || 1, x: d.del ? 1 : 0, u: u };
        changed = true;
      }
    }
  });
  if (!changed) return;
  Object.keys(ev).forEach(function (id) { if (ev[id].t < now - 2 * 864e5) delete ev[id]; });
  nput_('nset', set); nput_('nev', ev); nput_('ntim', tim);
}
function rebuildNotifyState_() {
  var sh = sheet_(), n = sh.getLastRow() - 1;
  ['nset', 'nev', 'ntim'].forEach(function (k) { PropertiesService.getScriptProperties().deleteProperty(k); });
  if (n < 1) return;
  var from = Math.max(2, n - 1500 + 2);
  var vals = sh.getRange(from, HEAD.length, sh.getLastRow() - from + 1, 1).getValues();
  ingest_(vals.map(function (r) { return String(r[0] || ''); }));
}
/* btns: [[label, {action:'snooze'|'done', ...}], ...] become buttons on the notification that call this script */
function send_(topic, title, message, prio, tags, btns, click) {
  if (!topic) return;
  var msg = { topic: topic, title: title, message: message, priority: prio || 3, tags: tags || [] }, url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  if (click) msg.click = click;
  if (url && btns && btns.length) msg.actions = btns.map(function (b) {
    return { action: 'http', label: b[0], url: url, method: 'POST', headers: { 'Content-Type': 'text/plain' }, body: JSON.stringify(b[1]), clear: true }; });
  UrlFetchApp.fetch(NTFY, { method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify(msg) });
}
/* "snooze": forget that a reminder was sent, and stay quiet for N minutes; "done": log it in the sheet like the app does */
function nAction_(b) {
  var set = nget_('nset', {}), now = Date.now();
  if (b.action === 'snooze') {
    var sent = nget_('nsent', {}), snz = nget_('nsnz', {});
    Object.keys(sent).forEach(function (k) { if (k.indexOf(b.key) === 0) delete sent[k]; });
    snz[b.key] = now + (+b.min || 30) * 6e4;
    nput_('nsent', sent); nput_('nsnz', snz);
    return;
  }
  if (b.action === 'done') {
    var tz = set.tz || 'Asia/Jerusalem', id = 'n' + now.toString(36) + Math.random().toString(36).slice(2, 6), t = new Date(now);
    var d = b.kind === 'water' ? { id: id, type: 'water', n: 1, t: now } : { id: id, type: b.kind === 'mmed' ? 'mmed' : 'med', name: String(b.name || ''), t: now };
    var label = d.type === 'water' ? 'מים לאמא' : d.type === 'mmed' ? 'תרופה לאמא' : 'ויטמין / תרופה';
    var row = [id, Utilities.formatDate(t, tz, 'd.M.yyyy, HH:mm:ss'), 'פעיל', Utilities.formatDate(t, tz, 'd.M.yyyy'), Utilities.formatDate(t, tz, 'HH:mm'), label, '', '', '', '', '', '', '',
      d.type === 'water' ? 'בקבוק' : d.name, '', 'נרשם מההתראה', '', JSON.stringify({ k: 'ev', id: id, u: now, d: d })];
    var sh = sheet_(), start = sh.getLastRow() + 1;
    sh.getRange(start, HEAD.length, 1, 1).setNumberFormat('@');
    sh.getRange(start, 1, 1, HEAD.length).setValues([row]);
    try { rebuildClean_(sh); } catch (e) {}
    ingest_([row[HEAD.length - 1]]);
    send_(set.topic, 'נרשם ✓', d.type === 'water' ? 'בקבוק מים' : d.name, 1, ['white_check_mark']);
  }
}
function mins_(hhmm) { var a = String(hhmm || '00:00').split(':'); return (+a[0] || 0) * 60 + (+a[1] || 0); }

function notifyTick() {
  var P = PropertiesService.getScriptProperties();
  P.setProperty('nlast', String(Date.now()));
  var set = nget_('nset', null);
  if (!set || !set.topic) return;
  var ev = nget_('nev', {}), tim = nget_('ntim', {}), sent = nget_('nsent', {}), now = Date.now(), tz = set.tz || 'Asia/Jerusalem';
  var day = Utilities.formatDate(new Date(now), tz, 'yyyy-MM-dd'), hm = function (t) { return Utilities.formatDate(new Date(t), tz, 'HH:mm'); };
  var nowM = mins_(hm(now)), list = Object.keys(ev).map(function (k) { return ev[k]; }).filter(function (e) { return !e.x; });
  var today = list.filter(function (e) { return Utilities.formatDate(new Date(e.t), tz, 'yyyy-MM-dd') === day; });
  var snz = nget_('nsnz', {});
  var snoozed = function (key) { return Object.keys(snz).some(function (p) { return snz[p] > now && key.indexOf(p) === 0; }); };
  var once = function (key, fn) { if (sent[key] || snoozed(key)) return; sent[key] = now; try { fn(); } catch (e) {} };
  var app = set.app || '';

  // feeding: "ended" note, X minutes before the next feed, and when it is due
  var feeds = list.filter(function (e) { return e.y === 'feed' || e.y === 'bottle'; }).sort(function (a, b) { return b.e - a.e; });
  if (feeds.length && !tim.feed) {
    var end = feeds[0].e, due = set.every ? end + set.every * 36e5 : 0;
    // sent once per feed, any time before the next one is due (not only right after it ended), at normal priority so it shows on the lock screen
    if (set.end && (due ? now < due - 2 * 6e4 : now - end < 3 * 36e5)) once('e' + end, function () {
      send_(set.topic, 'הנקה הסתיימה ב-' + hm(end), due ? 'ההנקה הבאה ב-' + hm(due) + ' (בעוד ' + Math.max(1, Math.round((due - now) / 6e4)) + ' דק׳)' : 'נרשם ביומן', 3, ['baby_bottle'], null, app); });
    // several "minutes before" alerts; if more than one is already due (e.g. phone was offline), send only the closest one
    var hitL = (set.leads || []).filter(function (l) { return now >= due - l * 6e4; }).sort(function (a, b) { return a - b; });
    if (due && hitL.length && now < due) {
      var fresh = hitL.filter(function (l) { return !sent['p' + l + ':' + end]; });
      if (fresh.length) { hitL.forEach(function (l) { if (l !== fresh[0]) sent['p' + l + ':' + end] = now; });
        once('p' + fresh[0] + ':' + end, function () {
          send_(set.topic, 'הנקה בעוד ' + Math.max(1, Math.round((due - now) / 6e4)) + ' דק׳', 'הבאה ב-' + hm(due) + ' · האחרונה הסתיימה ב-' + hm(end), 4, ['baby_bottle', 'alarm_clock'], null, app); }); }
    }
    if (due && set.due && now >= due && now < due + 3 * 36e5) once('d' + end, function () {
      send_(set.topic, 'הגיע הזמן להנקה', 'האחרונה הסתיימה ב-' + hm(end), 4, ['baby_bottle'], [['עוד 15 דק׳', { action: 'snooze', key: 'd' + end, min: 15 }]], app); });
  }
  // daily vitamins / medicine, once a day each, until 3 hours after the set time
  if (set.meds) (set.daily || []).forEach(function (r) {
    var tm = mins_(r.t), key = 'r' + r.id + day; if (nowM < tm || (nowM > tm + 180 && !snz[key])) return;
    var given = today.some(function (e) { return e.y === (r.m ? 'mmed' : 'med') && e.n === r.l; });
    if (!given) once(key, function () { send_(set.topic, '💊 ' + r.l + (r.m ? ' לאמא' : ' לתינוקת'), 'תזכורת יומית · ' + r.t, 3, ['pill'],
      [[r.m ? 'נלקח ✓' : 'ניתן ✓', { action: 'done', kind: r.m ? 'mmed' : 'med', name: r.l }], ['עוד 30 דק׳', { action: 'snooze', key: key, min: 30 }]], app); });
  });
  // mother's painkiller: when the next dose is allowed (only if chosen and an interval was set)
  if (set.pain) Object.keys(set.gap || {}).forEach(function (n) {
    var g = +set.gap[n] || 0; if (!g) return;
    var last = list.filter(function (e) { return e.y === 'mmed' && e.n === n; }).sort(function (a, b) { return b.t - a.t; })[0];
    if (!last) return; var ready = last.t + g * 36e5;
    if (now >= ready && now < ready + 3 * 36e5) once('k' + last.t, function () { send_(set.topic, 'אפשר לקחת שוב ' + n, 'המנה הקודמת נלקחה ב-' + hm(last.t), 3, ['pill'], null, app); });
  });
  // water: behind the pace by at least one bottle (same rule as in the app)
  var w = set.w || {}, goal = +w.goal || 0;
  if (set.water && goal) {
    var a = mins_(w.start || '07:00'), b = mins_(w.end || '22:00'); if (b <= a) b = a + 16 * 60;
    if (nowM >= a + 60 && nowM <= b) {
      var exp = Math.floor(goal * (nowM - a) / (b - a) + 1e-9), cnt = today.filter(function (e) { return e.y === 'water'; }).reduce(function (s, e) { return s + (e.c || 1); }, 0);
      if (exp - cnt >= 1) once('w' + day + ':' + exp, function () { send_(set.topic, '💧 הגיע הזמן לבקבוק מים', 'שתית ' + cnt + ' מתוך ' + goal + '. לפי השעה כדאי להיות ב-' + exp + '.', 3, ['droplet'],
        [['+ בקבוק', { action: 'done', kind: 'water' }], ['עוד שעה', { action: 'snooze', key: 'w' + day, min: 60 }]], app); });
    }
  }
  Object.keys(sent).forEach(function (k) { if (sent[k] < now - 3 * 864e5) delete sent[k]; });
  Object.keys(snz).forEach(function (k) { if (snz[k] < now) delete snz[k]; });
  nput_('nsent', sent); nput_('nsnz', snz);
}
