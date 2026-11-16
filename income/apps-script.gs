// סקריפט לטבלת גוגל של אפליקציית ההכנסות.
// מדביקים ב-Extensions > Apps Script של טבלה חדשה, ופורסים כ-Web app
// (Execute as: Me, Who has access: Anyone). את הכתובת מדביקים בהגדרות האפליקציה.
// אם רוצים סיסמה, כותבים אותה כאן ואותה מדביקים גם בהגדרות האפליקציה.
// את הסיסמה לא מעלים ל-GitHub.

var KEY = '';
var SHEET_NAME = 'Income';
var HEAD = ['id', 'date', 'src', 'income', 'tips', 'count', 'note', 'updated', 'deleted'];

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(SHEET_NAME);
  if (!s) {
    s = ss.insertSheet(SHEET_NAME);
    s.appendRow(HEAD);
    s.setFrozenRows(1);
    s.getRange('A:C').setNumberFormat('@');
    s.getRange('G:G').setNumberFormat('@');
  }
  return s;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function dateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Jerusalem', 'yyyy-MM-dd');
  return String(v);
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (KEY && p.key !== KEY) return out_({ ok: false, error: 'key' });
  var data = sheet_().getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0]) continue;
    list.push({
      id: String(r[0]),
      date: dateStr_(r[1]),
      src: String(r[2]),
      income: Number(r[3]) || 0,
      tips: Number(r[4]) || 0,
      count: Number(r[5]) || 0,
      note: String(r[6] || ''),
      updated: Number(r[7]) || 0,
      deleted: r[8] === true || r[8] === 'TRUE' || r[8] === 'true'
    });
  }
  return out_({ ok: true, entries: list });
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'json' }); }
  if (KEY && body.key !== KEY) return out_({ ok: false, error: 'key' });
  if (body.action !== 'upsertMany' || !Array.isArray(body.entries)) return out_({ ok: false, error: 'action' });
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var s = sheet_();
    var data = s.getDataRange().getValues();
    var idx = {};
    for (var i = 1; i < data.length; i++) idx[String(data[i][0])] = i + 1;
    var n = 0;
    body.entries.forEach(function (x) {
      if (!x || !x.id) return;
      var row = [String(x.id), String(x.date), String(x.src), Number(x.income) || 0, Number(x.tips) || 0, Number(x.count) || 0, String(x.note || ''), Number(x.updated) || 0, !!x.deleted];
      var at = idx[row[0]];
      if (at) {
        if (Number(data[at - 1][7]) <= row[7]) { s.getRange(at, 1, 1, 9).setValues([row]); n++; }
      } else {
        s.appendRow(row);
        idx[row[0]] = s.getLastRow();
        data[idx[row[0]] - 1] = row;
        n++;
      }
    });
    return out_({ ok: true, written: n });
  } finally {
    lock.releaseLock();
  }
}
