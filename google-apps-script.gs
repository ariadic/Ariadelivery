/**
 * ארי משלוחים – גיבוי הזמנות לטבלה + מייל
 *
 * מדביקים את הקוד הזה ב-Apps Script של גיליון Google Sheets
 * (בגיליון: תוספים ← Apps Script), ופורסים כ-Web App.
 * כל הזמנה מהאתר נוספת כשורה בגיליון ונשלחת למייל של בעל הגיליון.
 */

// לאן לשלוח את המייל. ריק = לחשבון הגוגל שפרס את הסקריפט.
const NOTIFY_EMAIL = "";

const HEADERS = [
  "תאריך", "שעה", "מס' הזמנה", "לקוח", "מקבל", "יעד", "כתובת מסירה",
  "טלפון מקבל", "מה", "מתי", "תשלום", "מחיר", "הערות", "כתובת איסוף", "טלפון איסוף", "סטטוס",
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const o = JSON.parse(e.postData.contents);
    if (!o || !o.no || !o.message) return text("bad request");

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
      sheet.setFrozenRows(1);
      sheet.setRightToLeft(true);
    }

    // לא לרשום את אותה הזמנה פעמיים
    const last = sheet.getLastRow();
    if (last > 1) {
      const nos = sheet.getRange(2, 3, last - 1, 1).getDisplayValues().flat();
      if (nos.indexOf(String(o.no)) !== -1) return text("duplicate");
    }

    const t = o.time ? new Date(o.time) : new Date();
    const tz = Session.getScriptTimeZone();
    sheet.appendRow([
      Utilities.formatDate(t, tz, "dd/MM/yyyy"), Utilities.formatDate(t, tz, "HH:mm"),
      "'" + o.no, clean(o.customer), clean(o.recipient), clean(o.dest), clean(o.address),
      "'" + clean(o.recipientPhone), clean(o.item), clean(o.when), clean(o.pay),
      typeof o.price === "number" ? o.price : clean(o.price), clean(o.notes),
      clean(o.pickup), "'" + clean(o.pickupPhone), "חדשה",
    ]);

    const to = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
    MailApp.sendEmail({
      to: to,
      subject: "🛵 הזמנה חדשה #" + o.no + " – " + clean(o.customer) + " ← " + clean(o.dest),
      body: String(o.message).replace(/\*/g, ""),
    });
    return text("ok");
  } catch (err) {
    return text("error: " + err);
  } finally {
    lock.releaseLock();
  }
}

// מונע הזרקת נוסחאות לגיליון ומגביל אורך
function clean(v) {
  let s = v === undefined || v === null ? "" : String(v).slice(0, 1000);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function text(s) {
  return ContentService.createTextOutput(s);
}

// להרצה ידנית פעם אחת מתוך העורך כדי לאשר הרשאות ולבדוק שהכול עובד
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
