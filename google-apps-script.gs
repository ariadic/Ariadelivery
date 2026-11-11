/**
 * ארי משלוחים – גיבוי הזמנות לטבלה + מייל + סטטוס משלוח
 *
 * מדביקים את הקוד הזה ב-Apps Script של גיליון Google Sheets
 * (בגיליון: תוספים ← Apps Script), ופורסים כ-Web App.
 *  - כל הזמנה מהאתר נוספת כשורה בגיליון ונשלחת למייל של בעל הגיליון.
 *  - דף המנהל (admin.html) מציג את ההזמנות ומעדכן סטטוס (נאסף / נמסר).
 *  - הלקוח רואה את הסטטוס תחת "ההזמנות שלי".
 */

// לאן לשלוח את המייל. ריק = לחשבון הגוגל שפרס את הסקריפט.
const NOTIFY_EMAIL = "";

// סיסמה לדף המנהל – בחרו סיסמה משלכם (אותיות ומספרים, לפחות 6 תווים).
// אותה סיסמה מקלידים פעם אחת בדף המנהל.
const ADMIN_PASSWORD = "";

const STATUSES = ["חדשה", "נאסף", "נמסר"];
const HEADERS = [
  "תאריך", "שעה", "מס' הזמנה", "לקוח", "מקבל", "יעד", "כתובת מסירה",
  "טלפון מקבל", "מה", "מתי", "תשלום", "מחיר", "הערות", "כתובת איסוף", "טלפון איסוף", "סטטוס",
];
const COL = { date: 1, time: 2, no: 3, customer: 4, recipient: 5, dest: 6, address: 7, phone: 8,
  item: 9, when: 10, pay: 11, price: 12, notes: 13, pickup: 14, pickupPhone: 15, status: 16 };

/* ---------- קבלת הזמנה חדשה / עדכון סטטוס ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const o = JSON.parse(e.postData.contents);
    if (o && o.action === "setStatus") return setStatus(o);
    if (!o || !o.no || !o.message) return json({ ok: false, error: "bad request" });

    const sheet = getSheet();
    if (findRow(sheet, o.no)) return json({ ok: true, duplicate: true });

    const t = o.time ? new Date(o.time) : new Date();
    const tz = Session.getScriptTimeZone();
    sheet.appendRow([
      Utilities.formatDate(t, tz, "dd/MM/yyyy"), Utilities.formatDate(t, tz, "HH:mm"),
      "'" + o.no, clean(o.customer), clean(o.recipient), clean(o.dest), clean(o.address),
      "'" + clean(o.recipientPhone), clean(o.item), clean(o.when), clean(o.pay),
      typeof o.price === "number" ? o.price : clean(o.price), clean(o.notes),
      clean(o.pickup), "'" + clean(o.pickupPhone), STATUSES[0],
    ]);
    addStatusDropdown(sheet, sheet.getLastRow());

    const to = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
    MailApp.sendEmail({
      to: to,
      subject: "🛵 הזמנה חדשה #" + o.no + " – " + clean(o.customer) + " ← " + clean(o.dest),
      body: String(o.message).replace(/\*/g, ""),
    });
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function setStatus(o) {
  if (!checkPassword(o.key)) return json({ ok: false, error: "סיסמה שגויה" });
  if (STATUSES.indexOf(o.status) === -1) return json({ ok: false, error: "סטטוס לא חוקי" });
  const sheet = getSheet();
  const row = findRow(sheet, o.no);
  if (!row) return json({ ok: false, error: "הזמנה לא נמצאה" });
  sheet.getRange(row, COL.status).setValue(o.status);
  return json({ ok: true });
}

/* ---------- קריאה: סטטוסים ללקוח / רשימה למנהל ---------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const sheet = getSheet();
  const rows = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getDisplayValues() : [];

  // ללקוח: רק הסטטוס של מספרי ההזמנות שהוא שלח – בלי פרטים אישיים
  if (p.action === "status") {
    const wanted = String(p.nos || "").split(",").slice(0, 200);
    const out = {};
    rows.forEach(r => { if (wanted.indexOf(r[COL.no - 1]) !== -1) out[r[COL.no - 1]] = r[COL.status - 1]; });
    return json({ ok: true, statuses: out });
  }

  // למנהל: הזמנות שלא נמסרו + כל ההזמנות של 3 הימים האחרונים
  if (p.action === "list") {
    if (!checkPassword(p.key)) return json({ ok: false, error: "סיסמה שגויה" });
    const cutoff = Date.now() - 3 * 24 * 3600 * 1000;
    const orders = rows.map(r => ({
      date: r[0], time: r[1], no: r[2], customer: r[3], recipient: r[4], dest: r[5], address: r[6],
      phone: r[7], item: r[8], when: r[9], pay: r[10], price: r[11], notes: r[12],
      pickup: r[13], pickupPhone: r[14], status: r[15],
    })).filter(o => o.status !== STATUSES[2] || parseDate(o.date, o.time) > cutoff);
    return json({ ok: true, orders: orders.reverse() });
  }

  return json({ ok: true, service: "ari-delivery" });
}

/* ---------- עזר ---------- */
function getSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setRightToLeft(true);
  }
  return sheet;
}

function findRow(sheet, no) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const nos = sheet.getRange(2, COL.no, last - 1, 1).getDisplayValues().flat();
  const i = nos.indexOf(String(no));
  return i === -1 ? 0 : i + 2;
}

function addStatusDropdown(sheet, row) {
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build();
  sheet.getRange(row, COL.status).setDataValidation(rule);
}

function checkPassword(key) {
  return ADMIN_PASSWORD.length >= 6 && String(key || "") === ADMIN_PASSWORD;
}

function parseDate(d, t) {
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(d);
  if (!m) return 0;
  const hm = /^(\d{2}):(\d{2})$/.exec(t) || [0, 0, 0];
  return new Date(+m[3], +m[2] - 1, +m[1], +hm[1], +hm[2]).getTime();
}

// מונע הזרקת נוסחאות לגיליון ומגביל אורך
function clean(v) {
  let s = v === undefined || v === null ? "" : String(v).slice(0, 1000);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// להרצה ידנית מתוך העורך כדי לאשר הרשאות ולבדוק שהכול עובד
function testOrder() {
  const res = doPost({ postData: { contents: JSON.stringify({
    no: "TEST-" + Date.now(), time: new Date().toISOString(), customer: "בדיקה",
    recipient: "בדיקה", dest: "קרית אתא", address: "בדיקה 1, קרית אתא", recipientPhone: "0500000000",
    item: "זר", when: "רגיל", pay: "מזומן", price: 15, notes: "הזמנת בדיקה – אפשר למחוק",
    pickup: "העצמאות 69, קרית אתא", pickupPhone: "0546665375",
    message: "🛵 הזמנת בדיקה – אם קיבלת את המייל הזה, הגיבוי עובד ✅",
  }) } });
  Logger.log(res.getContent());
}
