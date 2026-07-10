/**
 * אבחון רמת AI — מרכז הרכב שמעון ברזילי
 * בכל אבחון שמושלם:
 *   1) מוסיף שורה לגיליון המרכזי (תמונת-על להשוואה בין כל העובדים).
 *   2) יוצר/מוצא תת-תיקייה על שם העובד בתוך "עובדי הקורס" שבתיקיית הקורס,
 *      ושומר בה מסמך אבחון אישי — נקודת ההתחלה של תיק העובד לאורך הקורס.
 *
 * חשוב: אחרי עדכון הקוד, פרוס מחדש כ-**גרסה חדשה של אותה פריסה**
 * (Deploy ▸ Manage deployments ▸ עיפרון ▸ Version: New version ▸ Deploy)
 * כדי לשמור על אותה כתובת /exec. בפעם הראשונה תתבקש לאשר הרשאות דרייב.
 */

var EMPLOYEES_ROOT_NAME = 'עובדי הקורס — אבחון AI';

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // (1) גיליון מרכזי
    var sh = ss.getSheetByName('תשובות') || ss.getSheets()[0];
    var lastCol = sh.getLastColumn();
    var headers = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    Object.keys(data).forEach(function (k) {
      if (headers.indexOf(k) === -1) { headers.push(k); sh.getRange(1, headers.length).setValue(k); }
    });
    sh.appendRow(headers.map(function (h) { return data[h] == null ? '' : data[h]; }));

    // (2) תיק עובד אישי (לא מפיל את השליחה אם משהו משתבש)
    var folderUrl = '';
    try { folderUrl = saveToEmployeeFolder_(ss, data); } catch (fe) { folderUrl = 'folder_error: ' + fe; }

    return json({ ok: true, folder: folderUrl });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function saveToEmployeeFolder_(ss, data) {
  var name = ((data['שם'] || '').toString().trim()) || 'ללא שם';
  var safe = name.replace(/[\\\/\[\]\*\?:<>|"]/g, ' ').replace(/\s+/g, ' ').trim() || 'ללא שם';

  // תיקיית הקורס = התיקייה שבה יושב הגיליון
  var courseFolder = DriveApp.getFileById(ss.getId()).getParents().next();
  var root = getOrCreateFolder_(courseFolder, EMPLOYEES_ROOT_NAME);
  var emp = getOrCreateFolder_(root, safe);

  // מסמך אבחון אישי (מתוארך — כך שגם אבחון פתיחה וגם אבחון סיום נשמרים)
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

  // העברת המסמך מ-My Drive אל תיקיית העובד
  var f = DriveApp.getFileById(doc.getId());
  emp.addFile(f);
  try { DriveApp.getRootFolder().removeFile(f); } catch (e3) {}

  return emp.getUrl();
}

function getOrCreateFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function doGet() { return ContentService.createTextOutput('AI diagnostic collector — OK'); }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
