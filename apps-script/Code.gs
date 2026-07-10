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
var EMPLOYEES_ROOT_ID = '1WBMvGbY374Kn9xjV524sIimiRpSeWkR8'; // תיקיית "עובדים" שיצרת

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var data = JSON.parse(e.postData.contents);

    if (data.type === 'course') return handleCourse_(data); // סנכרון מרחב הקורס

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

  var root = DriveApp.getFolderById(EMPLOYEES_ROOT_ID);
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
  DriveApp.getFolderById(EMPLOYEES_ROOT_ID).getName();
  var d = DocumentApp.create('אישור הרשאות — נא למחוק');
  DriveApp.getFileById(d.getId()).setTrashed(true);
  return 'authorized';
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'course' && e.parameter.code) {
    try {
      var ss = SpreadsheetApp.openById(SHEET_ID);
      var sh = ss.getSheetByName('מעקב קורס');
      if (!sh) return json({ ok: true, found: false });
      var v = sh.getDataRange().getValues();
      for (var i = 1; i < v.length; i++) {
        if (String(v[i][0]) === String(e.parameter.code)) {
          return json({ ok: true, found: true,
            progress: {
              sessions: String(v[i][2] || '').split(',').filter(Boolean).map(Number),
              exercises: String(v[i][3] || '').split(',').filter(Boolean)
            },
            notes: v[i][5] || '', ts: v[i][7] || 0 });
        }
      }
      return json({ ok: true, found: false });
    } catch (err) { return json({ ok: false, error: String(err) }); }
  }
  return ContentService.createTextOutput('AI diagnostic collector — OK');
}

// סנכרון התקדמות/הערות של משתתף לטאב "מעקב קורס" (שורה אחת לכל קוד)
function handleCourse_(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName('מעקב קורס');
    if (!sh) { sh = ss.insertSheet('מעקב קורס');
      sh.appendRow(['קוד','שם','מפגשים_שהושלמו','תרגולים_שהושלמו','אחוז','הערות','עודכן','ts']); }
    var sess = (data.progress && data.progress.sessions) || [];
    var exs  = (data.progress && data.progress.exercises) || [];
    var total = data.total || 6;
    var pct = Math.round(sess.length / total * 100) + '%';
    var row = [data.code||'', data.name||'', sess.join(','), exs.join(','), pct, data.notes||'', new Date(), data.ts||Date.now()];
    var codes = sh.getLastRow() > 1 ? sh.getRange(2,1,sh.getLastRow()-1,1).getValues() : [];
    var found = -1;
    for (var i=0;i<codes.length;i++){ if(String(codes[i][0])===String(data.code)){ found=i+2; break; } }
    if (found>0) sh.getRange(found,1,1,row.length).setValues([row]);
    else sh.appendRow(row);
    return json({ ok: true });
  } catch (err) { return json({ ok: false, error: String(err) }); }
  finally { try { lock.releaseLock(); } catch(e2){} }
}
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
