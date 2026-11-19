/**
 * יומן התינוקת – גיבוי וסנכרון בגוגל שיטס
 * מדביקים את כל הקובץ הזה ב-Extensions ← Apps Script של הגיליון,
 * ומפרסמים כ-Web app (Execute as: Me, Who has access: Anyone).
 * לא צריך לשנות כלום בקוד.
 *
 * חשוב: לא למיין ולא למחוק שורות בלשונית "יומן התינוקת".
 * לניתוחים אפשר ליצור לשונית אחרת שמפנה אליה.
 */
var TAB = 'יומן התינוקת';
var PHOTO_FOLDER = 'יומן התינוקת – תמונות';
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
