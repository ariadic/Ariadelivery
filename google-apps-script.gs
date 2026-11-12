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

// גרסת הקוד – מופיעה בבדיקת הכתובת, כדי לדעת שהפריסה באמת עודכנה
const CODE_VERSION = "2026-11-12d";

// סיסמה לדף המנהל – בחרו סיסמה משלכם (אותיות ומספרים, לפחות 6 תווים).
// אותה סיסמה מקלידים פעם אחת בדף המנהל.
const ADMIN_PASSWORD = "";

const STATUSES = ["חדשה", "נאסף", "נמסר", "לא פנוי", "אושרה"];   // "אושרה" = ארי לחץ "קיבלתי"
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
    if (o && o.action === "deleteOrder") return deleteOrder(o);
    if (o && o.action === "saveCustomer") return saveCustomer(o);
    if (o && o.action === "deleteCustomer") return deleteCustomer(o);
    if (!o || !o.no || !o.message) return json({ ok: false, error: "bad request" });

    const sheet = getSheet();
    if (findRow(sheet, o.no)) return json({ ok: true, duplicate: true });
    if (isDeleted(o.no)) return json({ ok: true, deleted: true });       // הזמנה שנמחקה – לא להחזיר אותה

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
        subject: (o.resync ? "🔁 הזמנה שלא נקלטה (נשלחה שוב) #" : "🚙 הזמנה חדשה #") + o.no + " – " + clean(o.customer) + " ← " + clean(o.dest),
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
  addStatusDropdown(sheet, row);            // שורות ישנות – מעדכנים את הרשימה הנפתחת לסטטוסים החדשים
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

/* ---------- לקוחות ומחירונים ----------
   שתי לשוניות בטבלה, שאפשר לערוך גם ידנית וגם מדף הניהול:
   "לקוחות"   – שורה לכל לקוח (קוד, שם, כתובת, טלפון, הגדרות)
   "מחירונים" – שורה לכל יעד: קוד לקוח | לקוח | יעד | מחיר | כתובת לווייז
   הקוד "כללי" הוא המחירון של הדף בלי לינק אישי. */
const CUST_SHEET = "לקוחות";
const PRICE_SHEET = "מחירונים";
const DEFAULT_CODE = "_default";
const DEFAULT_LABEL = "כללי";
const CUST_HEADERS = ["קוד", "שם העסק", "כתובת איסוף", "טלפון", "כותרת בדף", "איסוף קבוע", "אמצעי תשלום קבוע",
  "בלי \"מי משלם\"", "ברכה לכרטיס", "תוספת דחוף", "מה שולחים (מופרד בפסיקים)", "עודכן"];
const PRICE_HEADERS = ["קוד לקוח", "לקוח", "יעד", "מחיר", "כתובת לווייז (לא חובה)"];

function getTab(name, headers) {
  const ss = getSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    getSheet();                                   // לוודא שגיליון ההזמנות נשאר ראשון
    sh = ss.insertSheet(name, ss.getSheets().length);     // תמיד בסוף – גיליון ההזמנות נשאר ראשון
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  }
  // לשונית עם פחות עמודות ממה שצריך (נמחקו עמודות ריקות) – מוסיפים בחזרה
  const max = sh.getMaxColumns();
  if (max < headers.length) sh.insertColumnsAfter(max, headers.length - max);
  return sh;
}
function getCustSheet() {
  const sh = getTab(CUST_SHEET, CUST_HEADERS);
  migrateJsonCustomers(sh);
  return sh;
}

// גרסה קודמת שמרה כל לקוח כ-JSON בעמודה אחת – ממירים פעם אחת לטבלה קריאה
function migrateJsonCustomers(sh) {
  if (sh.getRange(1, 3).getDisplayValue() !== "הגדרות (לא לערוך ידנית)") return;
  const old = body(sh, 3).map(r => { try { return { code: String(r[0]), c: JSON.parse(r[2]) }; } catch (e) { return null; } }).filter(Boolean);
  const n = sh.getLastRow();
  sh.getRange(1, 1, Math.max(n, 1), CUST_HEADERS.length).clearContent();
  sh.getRange(1, 1, 1, CUST_HEADERS.length).setValues([CUST_HEADERS]).setFontWeight("bold");
  old.forEach(x => saveCustomerData(x.code, x.c));
}
const getPriceSheet = () => getTab(PRICE_SHEET, PRICE_HEADERS);

// קורא את השורות מתחת לכותרת. עובד גם אם בלשונית יש פחות עמודות ממה שהקוד מצפה
// (למשל אחרי מחיקת עמודות ריקות) – חסרים מושלמים כריקים.
function body(sh, cols) {
  const n = sh.getLastRow();
  if (n < 2) return [];
  const have = Math.min(cols, Math.max(sh.getMaxColumns ? sh.getMaxColumns() : cols, 1));
  return sh.getRange(2, 1, n - 1, have).getValues().map(r => r.length < cols ? r.concat(new Array(cols - r.length).fill("")) : r);
}
const toCode = v => { const s = String(v === undefined || v === null ? "" : v).trim(); return s === DEFAULT_LABEL ? DEFAULT_CODE : s; };
const fromCode = c => c === DEFAULT_CODE ? DEFAULT_LABEL : c;
const yes = v => v === true || /^(כן|v|✓|✔|true|1|x)$/i.test(String(v).trim());

function readCustomers() {
  const map = {};
  body(getCustSheet(), CUST_HEADERS.length).forEach(r => {
    const code = toCode(r[0]);
    if (!code) return;
    map[code] = sanitizeCustomer({
      name: code === DEFAULT_CODE ? "מחירון כללי" : r[1], addr: r[2], phone: r[3], title: r[4], fixedPickup: yes(r[5]),
      pay: r[6], hidePayer: yes(r[7]), greeting: yes(r[8]), urgentExtra: r[9],
      sizes: String(r[10] || "").split(",").map(s => s.trim()).filter(String), prices: [],
    });
    map[code].code = code;
  });
  body(getPriceSheet(), PRICE_HEADERS.length).forEach(r => {
    const code = toCode(r[0]), label = str(r[2], 60), price = Number(r[3]);
    if (!code || !label || String(r[3]).trim() === "" || !isFinite(price)) return;
    if (!map[code] && code === DEFAULT_CODE) map[code] = Object.assign(sanitizeCustomer({ name: "מחירון כללי" }), { code: code });
    if (!map[code] || map[code].prices.length >= 150) return;
    map[code].prices.push({ label: label, price: num(price, 10000), city: str(r[4], 80) || label });
  });
  return Object.keys(map).map(k => map[k]);
}

function str(v, max) { return v === undefined || v === null ? "" : String(v).trim().slice(0, max); }
function num(v, max) { const n = Number(v); return isFinite(n) && n >= 0 ? Math.min(Math.round(n * 100) / 100, max) : 0; }

// מנקה ומגביל את מה שנשמר, כדי שהדף של הלקוח יקבל רק נתונים תקינים
function sanitizeCustomer(c) {
  return {
    name: str(c.name, 80), addr: str(c.addr, 150), phone: str(c.phone, 20), title: str(c.title, 80),
    fixedPickup: !!c.fixedPickup, pay: str(c.pay, 30), hidePayer: !!c.hidePayer, greeting: !!c.greeting,
    urgentExtra: num(c.urgentExtra, 1000),
    sizes: (Array.isArray(c.sizes) ? c.sizes : []).map(s => str(s, 30)).filter(String).slice(0, 10),
    prices: (Array.isArray(c.prices) ? c.prices : []).map(p => ({
      label: str(p.label, 60), city: str(p.city, 80) || str(p.label, 60), price: num(p.price, 10000),
    })).filter(p => p.label).slice(0, 150),
  };
}

// כותב מחדש את כל לשונית המחירונים: מוריד את השורות של הקוד, ומוסיף את החדשות (אם יש)
function writePrices(code, name, prices) {
  const sh = getPriceSheet();
  const keep = body(sh, PRICE_HEADERS.length).filter(r => toCode(r[0]) && toCode(r[0]) !== code);
  const add = (prices || []).map(p => [fromCode(code), name, p.label, p.price, p.city && p.city !== p.label ? p.city : ""]);
  const all = keep.concat(add);
  const n = sh.getLastRow();
  if (n > 1) sh.getRange(2, 1, n - 1, PRICE_HEADERS.length).clearContent();
  if (all.length) sh.getRange(2, 1, all.length, PRICE_HEADERS.length).setValues(all);
}

function findCustRow(sh, code) {
  const rows = body(sh, 1).map(r => toCode(r[0]));
  const i = rows.indexOf(code);
  return i === -1 ? 0 : i + 2;
}

function saveCustomer(o) {
  if (!checkPassword(o.key)) return json({ ok: false, error: "סיסמה שגויה" });
  const code = str(o.code, 20);
  if (!/^[A-Za-z0-9_]{3,20}$/.test(code)) return json({ ok: false, error: "קוד לא תקין" });
  if (code !== DEFAULT_CODE && !str((o.customer || {}).name, 80)) return json({ ok: false, error: "חסר שם לקוח" });
  saveCustomerData(code, o.customer || {});
  return json({ ok: true, code: code });
}

function saveCustomerData(code, c) {
  const d = sanitizeCustomer(c);
  const isDef = code === DEFAULT_CODE;
  const name = isDef ? "מחירון כללי" : d.name;
  const yn = b => b ? "כן" : "לא";
  const row = [fromCode(code), name, d.addr, "'" + d.phone, d.title, yn(d.fixedPickup), d.pay, yn(d.hidePayer),
    yn(d.greeting), d.urgentExtra, d.sizes.join(", "), new Date()];
  if (d.phone === "") row[3] = "";
  const sh = getTab(CUST_SHEET, CUST_HEADERS);
  const r = findCustRow(sh, code);
  if (r) sh.getRange(r, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  writePrices(code, name, d.prices);
}

function deleteCustomer(o) {
  if (!checkPassword(o.key)) return json({ ok: false, error: "סיסמה שגויה" });
  const code = str(o.code, 20);
  const sh = getCustSheet();
  const r = findCustRow(sh, code);
  if (!r) return json({ ok: false, error: "לקוח לא נמצא" });
  sh.deleteRow(r);
  writePrices(code, "", []);
  return json({ ok: true });
}

/* ---------- מחיקת הזמנה ----------
   השורה נמחקת מהטבלה, ומספר ההזמנה נרשם בלשונית "נמחקו" – כדי שהאתר של הלקוח
   לא ישלח אותה שוב (הגיבוי האוטומטי) ושהיא תיעלם גם מ"ההזמנות שלי" שלו. */
const DELETED = "נמחקה";
const DELETED_SHEET = "נמחקו";
function getDeletedSheet() { return getTab(DELETED_SHEET, ["מס' הזמנה", "נמחקה ב", "לקוח", "מקבל", "כתובת", "מחיר"]); }
let deletedCache = null;
function isDeleted(no) {
  if (!deletedCache) deletedCache = body(getDeletedSheet(), 1).map(r => String(r[0]).replace(/^'/, ""));
  return deletedCache.indexOf(String(no)) !== -1;
}
function deleteOrder(o) {
  if (!checkPassword(o.key)) return json({ ok: false, error: "סיסמה שגויה" });
  const sheet = getSheet();
  const row = findRow(sheet, o.no);
  if (!row) return json({ ok: false, error: "הזמנה לא נמצאה" });
  const r = sheet.getRange(row, 1, 1, HEADERS.length).getDisplayValues()[0];
  getDeletedSheet().appendRow(["'" + o.no, new Date(), r[COL.customer - 1], r[COL.recipient - 1], r[COL.address - 1], r[COL.price - 1]]);
  sheet.deleteRow(row);
  deletedCache = null;
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
// כל שגיאה בקריאה חוזרת כהודעה ברורה (במקום מסך שגיאה של גוגל), כדי שיהיה קל לאתר אותה
function doGet(e) {
  try {
    return doGetInner(e);
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err), where: String(err && err.stack || "").split("\n").slice(0, 3).join(" | ") });
  }
}

function doGetInner(e) {
  const p = (e && e.parameter) || {};
  const sheet = getSheet();
  const rows = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getDisplayValues() : [];

  // ללקוח: האם ארי פנוי כרגע
  if (p.action === "availability") return json(Object.assign({ ok: true }, getAvailability()));

  // לדף ההזמנה: ההגדרות של הלקוח מהלינק האישי + המחירון הכללי + זמינות
  if (p.action === "config") {
    const all = readCustomers();
    const code = str(p.code, 20);          // לא "c" – גוגל שומר את השם הזה לעצמו וחוסם את הבקשה
    const pick = c => { if (!c) return null; const x = Object.assign({}, c); delete x.code; return x; };
    return json(Object.assign({ ok: true,
      customer: code && code !== DEFAULT_CODE ? pick(all.filter(c => c.code === code)[0]) : null,
      defaults: pick(all.filter(c => c.code === DEFAULT_CODE)[0]),
    }, getAvailability()));
  }

  // למנהל: כל הלקוחות והמחירונים
  if (p.action === "customers") {
    if (!checkPassword(p.key)) return json({ ok: false, error: "סיסמה שגויה" });
    return json({ ok: true, customers: readCustomers(), sheetUrl: getSpreadsheet().getUrl() });
  }

  // ללקוח: ההזמנות שלו – לפי קוד הלינק האישי, או לפי מספרי הזמנות שנשמרו במכשיר שלו.
  // מחזיר רק את מה שמוצג ב"ההזמנות שלי" (בלי טלפונים), 100 ימים אחורה.
  if (p.action === "history") {
    const code = String(p.code || "").trim();
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
    return json({ ok: true, orders: out, deleted: wanted.filter(isDeleted) });
  }

  // ללקוח: רק הסטטוס של מספרי ההזמנות שהוא שלח – בלי פרטים אישיים
  if (p.action === "status") {
    const wanted = String(p.nos || "").split(",").slice(0, 200);
    const out = {};
    rows.forEach(r => { if (wanted.indexOf(r[COL.no - 1]) !== -1) out[r[COL.no - 1]] = r[COL.status - 1]; });
    wanted.filter(isDeleted).forEach(n => { out[n] = DELETED; });
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
    return json(Object.assign({ ok: true, orders: orders.reverse(), mailQuota: mailQuota() }, getAvailability()));
  }

  return json({ ok: true, service: "ari-delivery", version: CODE_VERSION });
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

// גיליון ההזמנות: הלשונית הראשונה שאינה לשונית עזר (לקוחות / מחירונים / נמחקו)
function getSheet() {
  const aux = [CUST_SHEET, PRICE_SHEET, DELETED_SHEET];
  const all = getSpreadsheet().getSheets();
  const sheet = all.filter(s => aux.indexOf(s.getName()) === -1)[0] || all[0];
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
    message: "🚙 הזמנת בדיקה – אם קיבלת את המייל הזה, הגיבוי עובד ✅",
  }) } });
  Logger.log(res.getContent());
  Logger.log("הטבלה: " + getSpreadsheet().getUrl());
}

/* ---------- מכסת מיילים (לנורית בדף הניהול) ---------- */
function mailQuota() {
  try { return MailApp.getRemainingDailyQuota(); } catch (e) { return null; }
}

/* ---------- מייל בוקר יומי ----------
   מפעילים פעם אחת מתוך העורך: בוחרים installMorningReport ולוחצים "הפעלה".
   מאותו רגע נשלח כל בוקר בסביבות 07:30 מייל "המערכת תקינה" עם סיכום של אתמול.
   אם המייל לא הגיע – משהו לא עובד. */
const ADMIN_URL = "https://ariadic.github.io/Ariadelivery/admin.html";

function installMorningReport() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "morningReport")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("morningReport").timeBased().everyDays(1).atHour(7).nearMinute(30).create();
  morningReport();                                  // שולח מייל אחד עכשיו כדי לראות שזה עובד
  Logger.log("מייל הבוקר הופעל – יישלח כל יום בסביבות 07:30 (אזור זמן: " + Session.getScriptTimeZone() + ")");
}

function removeMorningReport() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "morningReport")
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function morningReport() {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const yest = Utilities.formatDate(new Date(now.getTime() - 24 * 3600 * 1000), tz, "dd/MM/yyyy");
  const month = Utilities.formatDate(now, tz, "MM/yyyy");
  const sheet = getSheet();
  const rows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getDisplayValues() : [];
  const price = r => { const n = Number(r[COL.price - 1]); return isFinite(n) ? n : 0; };
  const y = rows.filter(r => r[COL.date - 1] === yest && r[COL.status - 1] !== "לא פנוי");
  const notAvail = rows.filter(r => r[COL.date - 1] === yest && r[COL.status - 1] === "לא פנוי").length;
  const open = rows.filter(r => CLOSED.indexOf(r[COL.status - 1]) === -1);
  const openBy = {};
  open.forEach(r => { openBy[r[COL.customer - 1]] = (openBy[r[COL.customer - 1]] || 0) + 1; });
  const unpaid = rows.filter(r => r[COL.date - 1].slice(3) === month && r[COL.status - 1] === "נמסר" && r[COL.paid - 1] !== "כן");
  const sum = a => a.reduce((t, r) => t + price(r), 0);
  const quota = mailQuota();
  const lines = [
    "🟢 המערכת עובדת" + (quota !== null ? " · מיילים זמינים היום: " + quota : ""),
    "",
    "אתמול (" + yest + "): " + y.length + " משלוחים · ₪" + sum(y) + (notAvail ? " · " + notAvail + " לא פנוי" : ""),
    "פתוחות עכשיו: " + open.length + (open.length ? " (" + Object.keys(openBy).map(k => k + ": " + openBy[k]).join(", ") + ")" : ""),
    "לא שולם החודש: ₪" + sum(unpaid) + " · " + unpaid.length + " משלוחים",
    "",
    "אם המייל הזה לא הגיע עד 08:00 – משהו לא עובד. פתח את דף הניהול ובדוק את הנורית:",
    ADMIN_URL,
  ];
  MailApp.sendEmail({
    to: NOTIFY_EMAIL || Session.getEffectiveUser().getEmail(),
    subject: "☀️ בוקר טוב – ארי משלוחים: המערכת תקינה",
    body: lines.join("\n"),
  });
}

/* ---------- בדיקת מערכת ----------
   בעורך: בוחרים checkSystem ולוחצים "הפעלה". התוצאה מופיעה למטה ב"יומן ביצוע". */
function checkSystem() {
  const say = s => Logger.log(s);
  say("גרסת קוד: " + CODE_VERSION);
  try {
    const ss = getSpreadsheet();
    say("✅ טבלה: " + ss.getName() + " – " + ss.getSheets().map(s => s.getName() + " (" + s.getLastRow() + " שורות, " + s.getMaxColumns() + " עמודות)").join(" | "));
  } catch (e) { say("❌ לא מצליח לפתוח את הטבלה: " + e); return; }
  try { const o = getSheet(); say("✅ לשונית הזמנות: " + o.getName()); } catch (e) { say("❌ לשונית הזמנות: " + e); }
  try {
    const all = readCustomers();
    say("✅ לקוחות שנקראו: " + all.length);
    all.forEach(c => say("   • " + (c.code === DEFAULT_CODE ? "כללי" : c.code) + " – " + (c.name || "") + ": " + c.prices.length + " יעדים" +
      (c.prices.length ? " (" + c.prices.slice(0, 4).map(p => p.label + " " + p.price).join(", ") + (c.prices.length > 4 ? "…" : "") + ")" : "")));
    if (!all.some(c => c.code === "fl4x9")) say("⚠️ הלקוח fl4x9 (פרחי אודי ודורית) לא נמצא בלשונית \"לקוחות\" – בלי זה הדף שלה מציג מחירון ישן");
  } catch (e) { say("❌ שגיאה בקריאת לקוחות/מחירונים: " + e + (e.stack ? " | " + String(e.stack).split("\n")[1] : "")); }
  try {
    const r = JSON.parse(doGetInner({ parameter: { action: "config", c: "fl4x9" } }).getContent());
    say(r.customer ? "✅ הדף של דורית יקבל " + (r.customer.prices || []).length + " יעדים" : "⚠️ הדף של דורית לא יקבל לקוח (customer=null)");
  } catch (e) { say("❌ בקשת המחירון של דורית נכשלת: " + e + (e.stack ? " | " + String(e.stack).split("\n")[1] : "")); }
}
