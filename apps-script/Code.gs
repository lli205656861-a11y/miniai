/**
 * אבחון רמת AI — מרכז הרכב שמעון ברזילי
 * מקבל כל אבחון שמושלם מדף הנחיתה ומוסיף אותו כשורה בגיליון המשוב.
 *
 * התקנה (מפורט ב-SETUP-DRIVE.md):
 * 1. פותחים את גיליון "משוב אבחון AI" בדרייב.
 * 2. Extensions ▸ Apps Script, מדביקים את הקוד הזה ושומרים.
 * 3. Deploy ▸ New deployment ▸ Web app ▸ Execute as: Me ▸ Who has access: Anyone.
 * 4. מעתיקים את כתובת ה-/exec ומדביקים ב-SUBMIT_URL שבתוך index.html.
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var data = JSON.parse(e.postData.contents);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('תשובות') || ss.getSheets()[0];

    // כותרות: יוצרים/מרחיבים לפי המפתחות שמגיעים
    var lastCol = sh.getLastColumn();
    var headers = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    Object.keys(data).forEach(function (k) {
      if (headers.indexOf(k) === -1) {
        headers.push(k);
        sh.getRange(1, headers.length).setValue(k);
      }
    });

    var row = headers.map(function (h) {
      var v = data[h];
      return (v === undefined || v === null) ? '' : v;
    });
    sh.appendRow(row);

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function doGet() {
  return ContentService.createTextOutput('AI diagnostic collector — OK');
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
