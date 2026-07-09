/**
 * אבחון רמת AI — אוסף תשובות לגיליון המשוב
 * מרכז הרכב שמעון ברזילי
 *
 * הסקריפט הזה יושב בתוך גיליון המשוב (Extensions → Apps Script),
 * ומקבל כל אבחון שמושלם בדף ומוסיף אותו כשורה חדשה.
 * מדריך הקמה מלא: ראו SETUP-DRIVE.md
 */

// שם הלשונית שאליה נכתבות התשובות
var SHEET_NAME = 'תשובות';

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // מונע דריסה כששניים שולחים באותו רגע

    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    // כותרות דינמיות: מוסיף עמודה לכל מפתח חדש שמגיע
    var lastCol = sh.getLastColumn();
    var headers = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    var changed = false;
    Object.keys(data).forEach(function (k) {
      if (headers.indexOf(k) === -1) { headers.push(k); changed = true; }
    });
    if (changed) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    }

    var row = headers.map(function (h) {
      var v = data[h];
      return (v === null || v === undefined) ? '' : v;
    });
    sh.appendRow(row);

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

// בדיקה מהירה שהשירות חי — פתחו את כתובת ה-Web App בדפדפן
function doGet() {
  return json({ ok: true, message: 'AI diagnostic collector is live' });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
