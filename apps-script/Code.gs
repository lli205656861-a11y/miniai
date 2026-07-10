/**
 * אבחון רמת AI — מרכז הרכב שמעון ברזילי
 * בכל אבחון שמושלם:
 *   1) מוסיף שורה לגיליון המרכזי (תמונת-על להשוואה בין העובדים).
 *   2) יוצר/מוצא תת-תיקייה על שם העובד בתוך "עובדי הקורס — אבחון AI".
 *   3) שומר בתיקיית העובד: מסמך אבחון אישי + קובץ התעודה המעוצבת (PNG).
 *
 * חשוב: אחרי עדכון הקוד, פרוס מחדש כ-**גרסה חדשה של אותה פריסה**
 * (Deploy ▸ Manage deployments ▸ ✏️ ▸ Version: New version ▸ Deploy)
 * כדי לשמור על אותה כתובת /exec. בפעם הראשונה אשר הרשאות דרייב.
 */

var SHEET_ID = '16RJdyAhl83WWhIRusFsdZAouTrqnP47sT9AjfOmfE40'; // גיליון "משוב אבחון AI"
var COURSE_FOLDER_ID = '1k99wtjGCUYui_GOTeAH7lXExf66TSZMA'; // תיקיית הקורס בדרייב
var EMPLOYEES_ROOT_NAME = 'עובדי הקורס — אבחון AI';

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var data = JSON.parse(e.postData.contents);

    // התעודה (base64) — נשמרת כקובץ, לא כתא בגיליון
    var certB64 = data['__cert_png'];
    delete data['__cert_png'];

    var ss = SpreadsheetApp.openById(SHEET_ID);

    // (2)+(3) תיק עובד — רץ קודם כדי לרשום את התוצאה/השגיאה לגיליון
    var info = '';
    try { info = saveToEmployeeFolder_(ss, data, certB64); } catch (fe) { info = 'folder_error: ' + String(fe); }
    data['_תיקייה'] = info;

    // (1) גיליון מרכזי
    var sh = ss.getSheetByName('תשובות') || ss.getSheets()[0];
    var lastCol = sh.getLastColumn();
    var headers = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    Object.keys(data).forEach(function (k) {
      if (headers.indexOf(k) === -1) { headers.push(k); sh.getRange(1, headers.length).setValue(k); }
    });
    sh.appendRow(headers.map(function (h) { return data[h] == null ? '' : data[h]; }));

    return json({ ok: true, folder: info });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function saveToEmployeeFolder_(ss, data, certB64) {
  var name = ((data['שם'] || '').toString().trim()) || 'ללא שם';
  var safe = name.replace(/[\\\/\[\]\*\?:<>|"]/g, ' ').replace(/\s+/g, ' ').trim() || 'ללא שם';

  var courseFolder = DriveApp.getFolderById(COURSE_FOLDER_ID);
  var root = getOrCreateFolder_(courseFolder, EMPLOYEES_ROOT_NAME);
  var emp = getOrCreateFolder_(root, safe);

  // מסמך אבחון אישי (מתוארך)
  var title = 'אבחון רמת AI — ' + (data['תאריך'] || '') + ' (' + (data['מזהה_אבחון'] || '') + ')';
  var doc = DocumentApp.create(title);
  var b = doc.getBody();
  b.appendParagraph('אבחון רמת AI — ' + name).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  b.appendParagraph('תפקיד: ' + (data['תפקיד'] || '—'));
  b.appendParagraph('מזהה אבחון: ' + (data['מזהה_אבחון'] || '') + '   ·   תאריך: ' + (data['תאריך'] || ''));
  b.appendParagraph('');
  b.appendParagraph('מדד מוכנות: ' + (data['מדד_מוכנות'] || '') + ' / 100     רמה: ' + (data['רמה'] || ''))
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  b.appendParagraph('ידע ' + (data['ידע_אחוז'] || '') + '%  ·  ניסיון ' + (data['ניסיון_אחוז'] || '') + '%  ·  מוכנות ' + (data['מוכנות_אחוז'] || '') + '%');
  b.appendParagraph('פירוט ידע — בסיסי ' + (data['ידע_בסיסי'] || '') + '  ·  בינוני ' + (data['ידע_בינוני'] || '') + '  ·  מתקדם ' + (data['ידע_מתקדם'] || ''));
  b.appendParagraph('');
  b.appendParagraph('כל התשובות').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  var skip = {'שם':1,'תפקיד':1,'מזהה_אבחון':1,'תאריך':1,'חותמת_זמן':1};
  Object.keys(data).forEach(function (k) { if (!skip[k]) b.appendParagraph('• ' + k + ':  ' + data[k]); });
  doc.saveAndClose();
  moveToFolder_(doc.getId(), emp);

  // התעודה המעוצבת (PNG)
  if (certB64) {
    var png = Utilities.newBlob(Utilities.base64Decode(certB64), 'image/png',
      'תעודת מוכנות AI — ' + name + ' — ' + (data['תאריך'] || '') + '.png');
    emp.createFile(png);
  }
  return emp.getUrl();
}

function getOrCreateFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function moveToFolder_(fileId, folder) {
  var f = DriveApp.getFileById(fileId);
  folder.addFile(f);
  try { DriveApp.getRootFolder().removeFile(f); } catch (e) {}
}
/**
 * הרץ פעם אחת מהעורך (Run ▸ setup) כדי לאשר את הרשאות Drive + Docs.
 * זה נדרש כי הרשאות מצטברות — אישור דרך doGet לא מבקש הרשאות שאין בו.
 */
function setup() {
  DriveApp.getFolderById(COURSE_FOLDER_ID).getName();
  var d = DocumentApp.create('אישור הרשאות — נא למחוק');
  DriveApp.getFileById(d.getId()).setTrashed(true);
  return 'authorized';
}

function doGet() { return ContentService.createTextOutput('AI diagnostic collector — OK'); }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
