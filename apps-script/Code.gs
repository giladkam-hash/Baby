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
