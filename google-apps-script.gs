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

const STATUSES = ["חדשה", "נאסף", "נמסר", "לא פנוי"];
const CLOSED = ["נמסר", "לא פנוי"];               // הזמנות סגורות – מוצגות בדף המנהל עד CLOSED_DAYS ימים אחורה
const CLOSED_DAYS = 90;
const HEADERS = [
  "תאריך", "שעה", "מס' הזמנה", "לקוח", "מקבל", "יעד", "כתובת מסירה",
  "טלפון מקבל", "מה", "מתי", "תשלום", "מחיר", "הערות", "כתובת איסוף", "טלפון איסוף", "סטטוס", "קוד לקוח", "שולם",
];
const COL = { date: 1, time: 2, no: 3, customer: 4, recipient: 5, dest: 6, address: 7, phone: 8,
  item: 9, when: 10, pay: 11, price: 12, notes: 13, pickup: 14, pickupPhone: 15, status: 16, code: 17, paid: 18 };

/* ---------- קבלת הזמנה חדשה / עדכון סטטוס ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const o = JSON.parse(e.postData.contents);
    if (o && o.action === "setStatus") return setStatus(o);
    if (o && o.action === "setAvailability") return setAvailability(o);
    if (o && o.action === "setPaid") return setPaid(o);
    if (o && o.action === "saveCustomer") return saveCustomer(o);
    if (o && o.action === "deleteCustomer") return deleteCustomer(o);
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
      clean(o.pickup), "'" + clean(o.pickupPhone), STATUSES[0], clean(o.code), "לא",
    ]);
    addStatusDropdown(sheet, sheet.getLastRow());

    // ההזמנה כבר שמורה בטבלה – תקלה במייל לא מבטלת אותה
    try {
      const to = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
      MailApp.sendEmail({
        to: to,
        subject: (o.resync ? "🔁 הזמנה שלא נקלטה (נשלחה שוב) #" : "🛵 הזמנה חדשה #") + o.no + " – " + clean(o.customer) + " ← " + clean(o.dest),
        body: String(o.message).replace(/\*/g, ""),
      });
    } catch (mailErr) {
      return json({ ok: true, saved: true, mailError: String(mailErr) });
    }
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

function setPaid(o) {
  if (!checkPassword(o.key)) return json({ ok: false, error: "סיסמה שגויה" });
  const sheet = getSheet();
  const row = findRow(sheet, o.no);
  if (!row) return json({ ok: false, error: "הזמנה לא נמצאה" });
  sheet.getRange(row, COL.paid).setValue(o.paid ? "כן" : "לא");
  return json({ ok: true });
}

/* ---------- לקוחות ומחירונים (לשונית "לקוחות" בטבלה) ----------
   כל לקוח נשמר בשורה: קוד | שם | הגדרות (JSON) | עודכן.
   הקוד "_default" הוא המחירון הכללי של הדף בלי לינק אישי. */
const CUST_SHEET = "לקוחות";
const DEFAULT_CODE = "_default";

function getCustSheet() {
  const ss = getSpreadsheet();
  let sh = ss.getSheetByName(CUST_SHEET);
  if (!sh) {
    getSheet();                                   // לוודא שגיליון ההזמנות נשאר ראשון
    sh = ss.insertSheet(CUST_SHEET);
    sh.appendRow(["קוד", "שם", "הגדרות (לא לערוך ידנית)", "עודכן"]);
    sh.getRange(1, 1, 1, 4).setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  }
  return sh;
}

function readCustomers() {
  const sh = getCustSheet();
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 3).getValues().map(r => {
    try { return Object.assign(JSON.parse(r[2]), { code: String(r[0]) }); } catch (e) { return null; }
  }).filter(Boolean);
}

function str(v, max) { return v === undefined || v === null ? "" : String(v).trim().slice(0, max); }
function num(v, max) { const n = Number(v); return isFinite(n) && n >= 0 ? Math.min(Math.round(n * 100) / 100, max) : 0; }

// מנקה ומגביל את מה שנשמר, כדי שהדף של הלקוח יקבל רק נתונים תקינים
function sanitizeCustomer(c) {
  const out = {
    name: str(c.name, 80), addr: str(c.addr, 150), phone: str(c.phone, 20), title: str(c.title, 80),
    fixedPickup: !!c.fixedPickup, pay: str(c.pay, 30), hidePayer: !!c.hidePayer, greeting: !!c.greeting,
    urgentExtra: num(c.urgentExtra, 1000),
    sizes: (Array.isArray(c.sizes) ? c.sizes : []).map(s => str(s, 30)).filter(String).slice(0, 10),
    prices: (Array.isArray(c.prices) ? c.prices : []).map(p => ({
      label: str(p.label, 60), city: str(p.city, 80) || str(p.label, 60), price: num(p.price, 10000),
    })).filter(p => p.label).slice(0, 150),
  };
  return out;
}

function saveCustomer(o) {
  if (!checkPassword(o.key)) return json({ ok: false, error: "סיסמה שגויה" });
  const code = str(o.code, 20);
  if (!/^[A-Za-z0-9_]{3,20}$/.test(code)) return json({ ok: false, error: "קוד לא תקין" });
  const data = sanitizeCustomer(o.customer || {});
  if (code !== DEFAULT_CODE && !data.name) return json({ ok: false, error: "חסר שם לקוח" });
  const sh = getCustSheet();
  const n = sh.getLastRow();
  const codes = n > 1 ? sh.getRange(2, 1, n - 1, 1).getDisplayValues().flat() : [];
  const i = codes.indexOf(code);
  const row = [code, code === DEFAULT_CODE ? "מחירון כללי" : data.name, JSON.stringify(data), new Date()];
  if (i === -1) sh.appendRow(row); else sh.getRange(i + 2, 1, 1, 4).setValues([row]);
  return json({ ok: true, code: code });
}

function deleteCustomer(o) {
  if (!checkPassword(o.key)) return json({ ok: false, error: "סיסמה שגויה" });
  const sh = getCustSheet();
  const n = sh.getLastRow();
  const codes = n > 1 ? sh.getRange(2, 1, n - 1, 1).getDisplayValues().flat() : [];
  const i = codes.indexOf(str(o.code, 20));
  if (i === -1) return json({ ok: false, error: "לקוח לא נמצא" });
  sh.deleteRow(i + 2);
  return json({ ok: true });
}

/* ---------- זמינות כללית (פנוי / לא פנוי) ---------- */
function getAvailability() {
  const props = PropertiesService.getScriptProperties();
  return { available: props.getProperty("AVAILABLE") !== "0", message: props.getProperty("AVAIL_MSG") || "" };
}
function setAvailability(o) {
  if (!checkPassword(o.key)) return json({ ok: false, error: "סיסמה שגויה" });
  const props = PropertiesService.getScriptProperties();
  props.setProperty("AVAILABLE", o.available ? "1" : "0");
  props.setProperty("AVAIL_MSG", clean(o.message).slice(0, 200));
  return json(Object.assign({ ok: true }, getAvailability()));
}

/* ---------- קריאה: סטטוסים ללקוח / רשימה למנהל ---------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const sheet = getSheet();
  const rows = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getDisplayValues() : [];

  // ללקוח: האם ארי פנוי כרגע
  if (p.action === "availability") return json(Object.assign({ ok: true }, getAvailability()));

  // לדף ההזמנה: ההגדרות של הלקוח מהלינק האישי + המחירון הכללי + זמינות
  if (p.action === "config") {
    const all = readCustomers();
    const code = str(p.c, 20);
    const pick = c => { if (!c) return null; const x = Object.assign({}, c); delete x.code; return x; };
    return json(Object.assign({ ok: true,
      customer: code && code !== DEFAULT_CODE ? pick(all.filter(c => c.code === code)[0]) : null,
      defaults: pick(all.filter(c => c.code === DEFAULT_CODE)[0]),
    }, getAvailability()));
  }

  // למנהל: כל הלקוחות והמחירונים
  if (p.action === "customers") {
    if (!checkPassword(p.key)) return json({ ok: false, error: "סיסמה שגויה" });
    return json({ ok: true, customers: readCustomers() });
  }

  // ללקוח: ההזמנות שלו – לפי קוד הלינק האישי, או לפי מספרי הזמנות שנשמרו במכשיר שלו.
  // מחזיר רק את מה שמוצג ב"ההזמנות שלי" (בלי טלפונים), 100 ימים אחורה.
  if (p.action === "history") {
    const code = String(p.c || "").trim();
    const byCode = code.length >= 4;
    const wanted = String(p.nos || "").split(",").filter(String).slice(0, 300);
    const cutoff = Date.now() - 100 * 24 * 3600 * 1000;
    const out = [];
    rows.forEach((r, i) => {
      const codeMatch = byCode && r[COL.code - 1] === code;
      const noMatch = wanted.indexOf(r[COL.no - 1]) !== -1;
      if (!codeMatch && !noMatch) return;
      if (parseDate(r[COL.date - 1], r[COL.time - 1]) < cutoff) return;
      // הזמנה ישנה בלי קוד – משייכים אותה ללקוח, כדי שתופיע לו גם במכשירים אחרים
      if (noMatch && byCode && !r[COL.code - 1]) sheet.getRange(i + 2, COL.code).setValue(code);
      out.push({ no: r[COL.no - 1], date: r[COL.date - 1], time: r[COL.time - 1], recipient: r[COL.recipient - 1],
        address: r[COL.address - 1], item: r[COL.item - 1], price: r[COL.price - 1], status: r[COL.status - 1],
        paid: r[COL.paid - 1] === "כן" });
    });
    return json({ ok: true, orders: out });
  }

  // ללקוח: רק הסטטוס של מספרי ההזמנות שהוא שלח – בלי פרטים אישיים
  if (p.action === "status") {
    const wanted = String(p.nos || "").split(",").slice(0, 200);
    const out = {};
    rows.forEach(r => { if (wanted.indexOf(r[COL.no - 1]) !== -1) out[r[COL.no - 1]] = r[COL.status - 1]; });
    return json({ ok: true, statuses: out });
  }

  // למנהל: הזמנות פתוחות + הזמנות שנסגרו ב-90 הימים האחרונים, ומצב הזמינות
  if (p.action === "list") {
    if (!checkPassword(p.key)) return json({ ok: false, error: "סיסמה שגויה" });
    const cutoff = Date.now() - CLOSED_DAYS * 24 * 3600 * 1000;
    const orders = rows.map(r => ({
      date: r[0], time: r[1], no: r[2], customer: r[3], recipient: r[4], dest: r[5], address: r[6],
      phone: r[7], item: r[8], when: r[9], pay: r[10], price: r[11], notes: r[12],
      pickup: r[13], pickupPhone: r[14], status: r[15], paid: r[COL.paid - 1] === "כן",
    })).filter(o => CLOSED.indexOf(o.status) === -1 || parseDate(o.date, o.time) > cutoff);
    return json(Object.assign({ ok: true, orders: orders.reverse() }, getAvailability()));
  }

  return json({ ok: true, service: "ari-delivery" });
}

/* ---------- עזר ---------- */
// עובד גם אם הסקריפט נוצר מתוך הגיליון וגם אם נוצר בנפרד ב-script.google.com:
// אם אין גיליון מחובר – נוצר גיליון "הזמנות ארי משלוחים" ב-Google Drive ונשמר לפעמים הבאות.
function getSpreadsheet() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty("SHEET_ID");
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create("הזמנות ארי משלוחים");
  props.setProperty("SHEET_ID", ss.getId());
  return ss;
}

function getSheet() {
  const sheet = getSpreadsheet().getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setRightToLeft(true);
  } else {
    // טבלה מגרסה קודמת – משלימים כותרות של עמודות חדשות
    [COL.code, COL.paid].forEach(c => {
      if (sheet.getRange(1, c).getDisplayValue() === "") sheet.getRange(1, c).setValue(HEADERS[c - 1]).setFontWeight("bold");
    });
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
  Logger.log("הטבלה: " + getSpreadsheet().getUrl());
}
