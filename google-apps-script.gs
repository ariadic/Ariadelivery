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
const CODE_VERSION = "2026-11-12g";

// התראות לאייפון דרך אפליקציית Pushover (הכי אמין – עם צלצול חירום ו"קיבלתי" ישר מההתראה).
// מדביקים כאן את שני המפתחות מ-pushover.net. לא לשמור אותם באתר (הקוד באתר גלוי לכולם).
const PUSHOVER_TOKEN = "";   // API Token של האפליקציה "ארי משלוחים" (מתחיל ב-a)
const PUSHOVER_USER = "";    // Your User Key – מופיע בעמוד הראשי אחרי התחברות (מתחיל ב-u)

// גיבוי ישן: ntfy בחינם. משמש רק אם אין מפתחות Pushover. (בחינם הוא נחסם הרבה פעמים מגוגל)
const NTFY_TOPIC = "ari-orders-bawytzw56nm2wo";

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
    const q = (e && e.parameter) || {};
    if (q.action === "ack") return json(ackOrder(q));               // אישור התראת חירום ב-Pushover
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

    // התראה לאייפון (ntfy) – תקלה בהתראה לא מבטלת את ההזמנה
    notifyPhone(o);

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
   הקוד "מזדמנים" (או "כללי", השם הישן) הוא המחירון של לקוחות מזדמנים – הדף בלי לינק אישי. */
const CUST_SHEET = "לקוחות";
const PRICE_SHEET = "מחירונים";
const DEFAULT_CODE = "_default";
const DEFAULT_LABEL = "מזדמנים";
const DEFAULT_ALIASES = ["מזדמנים", "כללי", "לקוחות מזדמנים"];   // "כללי" – השם הקודם, עדיין נקרא
const DEFAULT_NAME = "לקוחות מזדמנים";
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
const toCode = v => { const s = String(v === undefined || v === null ? "" : v).trim(); return DEFAULT_ALIASES.indexOf(s) !== -1 ? DEFAULT_CODE : s; };
const fromCode = c => c === DEFAULT_CODE ? DEFAULT_LABEL : c;
const yes = v => v === true || /^(כן|v|✓|✔|true|1|x)$/i.test(String(v).trim());

function readCustomers() {
  const map = {};
  body(getCustSheet(), CUST_HEADERS.length).forEach(r => {
    const code = toCode(r[0]);
    if (!code) return;
    map[code] = sanitizeCustomer({
      name: code === DEFAULT_CODE ? DEFAULT_NAME : r[1], addr: r[2], phone: r[3], title: r[4], fixedPickup: yes(r[5]),
      pay: r[6], hidePayer: yes(r[7]), greeting: yes(r[8]), urgentExtra: r[9],
      sizes: String(r[10] || "").split(",").map(s => s.trim()).filter(String), prices: [],
    });
    map[code].code = code;
  });
  body(getPriceSheet(), PRICE_HEADERS.length).forEach(r => {
    const code = toCode(r[0]), label = str(r[2], 60), price = Number(r[3]);
    if (!code || !label || String(r[3]).trim() === "" || !isFinite(price)) return;
    if (!map[code] && code === DEFAULT_CODE) map[code] = Object.assign(sanitizeCustomer({ name: DEFAULT_NAME }), { code: code });
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
  const name = isDef ? DEFAULT_NAME : d.name;
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
    const q = (e && e.parameter) || {};
    if (q.action === "ack") return ackPage(ackOrder(q));               // "👍 קיבלתי" מההתראה באייפון
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

  // לדף ההזמנה: ההגדרות של הלקוח מהלינק האישי + מחירון לקוחות מזדמנים + זמינות
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
  const sheet = all.filter(s => aux.indexOf(s.getName()) === -1 && s.getName().indexOf(ARCHIVE_PREFIX) !== 0)[0] || all[0];
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

/* ---------- תחזוקה חודשית: ארכיון + מחיקת פרטים אישיים (רץ לבד ב-1 לכל חודש בלילה) ----------
   - הזמנות שנסגרו (נמסר / לא פנוי) לפני יותר מ-ARCHIVE_DAYS ימים עוברות ללשונית "ארכיון <שנה>". שום דבר לא נמחק.
     כך הלשונית הראשית נשארת קטנה והאתר נשאר מהיר.
   - בארכיון: אחרי ANONYMIZE_MONTHS חודשים נמחקים שם המקבל, הטלפון שלו, הרחוב וההערות (פרטיות).
     נשארים: תאריך, לקוח, יעד, מה, מחיר, תשלום, שולם – מה שצריך לצורכי מס.
   הפעלה: מריצים פעם אחת את installMaintenance מהעורך. */
const ARCHIVE_DAYS = 90;
const ANONYMIZE_MONTHS = 12;
const ARCHIVE_PREFIX = "ארכיון ";

function monthlyMaintenance() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const r = { archived: archiveOldOrders(), anonymized: anonymizeOld() };
    Logger.log("תחזוקה: הועברו לארכיון " + r.archived + " הזמנות · נמחקו פרטים אישיים מ-" + r.anonymized + " הזמנות ישנות");
    return r;
  } finally {
    lock.releaseLock();
  }
}

// ערכי טקסט שמתחילים בספרה (טלפון, מס' הזמנה) – שומרים כטקסט כדי שהאפס בהתחלה לא ייעלם
function asText(v) { return typeof v === "string" && /^[0-9+]/.test(v) ? "'" + v : v; }

function archiveOldOrders() {
  const sheet = getSheet();
  const n = sheet.getLastRow() - 1;
  if (n < 1) return 0;
  const rng = sheet.getRange(2, 1, n, HEADERS.length);
  const vals = rng.getValues(), disp = rng.getDisplayValues();
  const cutoff = Date.now() - ARCHIVE_DAYS * 24 * 3600 * 1000;
  const keep = [], move = {};
  const fix = row => row.map((v, j) => [COL.no, COL.phone, COL.pickupPhone].indexOf(j + 1) !== -1 ? asText(v) : v);
  disp.forEach((d, i) => {
    const t = parseDate(d[COL.date - 1], d[COL.time - 1]);
    if (CLOSED.indexOf(d[COL.status - 1]) !== -1 && t && t < cutoff) {
      const y = d[COL.date - 1].slice(6, 10);
      (move[y] = move[y] || []).push(fix(vals[i]));
    } else keep.push(fix(vals[i]));
  });
  const moved = n - keep.length;
  if (!moved) return 0;
  Object.keys(move).forEach(y => {
    const t = getTab(ARCHIVE_PREFIX + y, HEADERS);
    t.getRange(t.getLastRow() + 1, 1, move[y].length, HEADERS.length).setValues(move[y]);
  });
  rng.clearContent();
  if (keep.length) sheet.getRange(2, 1, keep.length, HEADERS.length).setValues(keep);
  return moved;
}

function anonymizeOld() {
  const cut = new Date();
  cut.setMonth(cut.getMonth() - ANONYMIZE_MONTHS);
  let count = 0;
  getSpreadsheet().getSheets().filter(sh => sh.getName().indexOf(ARCHIVE_PREFIX) === 0).forEach(sh => {
    const n = sh.getLastRow() - 1;
    if (n < 1) return;
    const disp = sh.getRange(2, 1, n, HEADERS.length).getDisplayValues();
    const who = sh.getRange(2, COL.recipient, n, 4).getValues();          // מקבל | יעד | כתובת | טלפון
    const notes = sh.getRange(2, COL.notes, n, 1).getValues();
    let changed = false;
    disp.forEach((d, i) => {
      const t = parseDate(d[COL.date - 1], d[COL.time - 1]);
      if (!t || t >= cut.getTime() || d[COL.recipient - 1] === "—") return;
      who[i] = ["—", who[i][1], who[i][1] || "—", ""];                     // הכתובת המלאה מתחלפת ביעד בלבד
      notes[i] = [""];
      changed = true; count++;
    });
    if (changed) {
      sh.getRange(2, COL.recipient, n, 4).setValues(who.map(r => r.map(asText)));
      sh.getRange(2, COL.notes, n, 1).setValues(notes);
    }
  });
  return count;
}

// להרצה פעם אחת מהעורך: מפעיל את התחזוקה החודשית ומריץ אותה עכשיו
function installMaintenance() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "monthlyMaintenance")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("monthlyMaintenance").timeBased().onMonthDay(1).atHour(3).create();
  const r = monthlyMaintenance();
  Logger.log("✅ תחזוקה חודשית הופעלה – תרוץ לבד ב-1 לכל חודש בלילה. עכשיו: " + r.archived + " הזמנות עברו לארכיון, " + r.anonymized + " הזמנות ישנות נוקו מפרטים אישיים.");
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
    const tr = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
    say((tr.indexOf("morningReport") !== -1 ? "✅" : "⚠️") + " מייל בוקר" + (tr.indexOf("morningReport") !== -1 ? " פעיל" : " לא פעיל – להריץ installMorningReport"));
    say((tr.indexOf("monthlyMaintenance") !== -1 ? "✅" : "⚠️") + " תחזוקה חודשית (ארכיון + פרטיות)" + (tr.indexOf("monthlyMaintenance") !== -1 ? " פעילה" : " לא פעילה – להריץ installMaintenance"));
  } catch (e) {}
  say(PUSHOVER_TOKEN && PUSHOVER_USER ? "✅ התראות לאייפון: Pushover" : NTFY_TOPIC ? "⚠️ התראות לאייפון: ntfy בחינם (לא אמין) – כדאי למלא מפתחות Pushover" : "⚠️ אין התראות לאייפון");
  try {
    const ss = getSpreadsheet();
    say("✅ טבלה: " + ss.getName() + " – " + ss.getSheets().map(s => s.getName() + " (" + s.getLastRow() + " שורות, " + s.getMaxColumns() + " עמודות)").join(" | "));
  } catch (e) { say("❌ לא מצליח לפתוח את הטבלה: " + e); return; }
  try { const o = getSheet(); say("✅ לשונית הזמנות: " + o.getName()); } catch (e) { say("❌ לשונית הזמנות: " + e); }
  try {
    const all = readCustomers();
    say("✅ לקוחות שנקראו: " + all.length);
    all.forEach(c => say("   • " + (c.code === DEFAULT_CODE ? DEFAULT_LABEL : c.code) + " – " + (c.name || "") + ": " + c.prices.length + " יעדים" +
      (c.prices.length ? " (" + c.prices.slice(0, 4).map(p => p.label + " " + p.price).join(", ") + (c.prices.length > 4 ? "…" : "") + ")" : "")));
    if (!all.some(c => c.code === "fl4x9")) say("⚠️ הלקוח fl4x9 (פרחי אודי ודורית) לא נמצא בלשונית \"לקוחות\" – בלי זה הדף שלה מציג מחירון ישן");
  } catch (e) { say("❌ שגיאה בקריאת לקוחות/מחירונים: " + e + (e.stack ? " | " + String(e.stack).split("\n")[1] : "")); }
  try {
    const r = JSON.parse(doGetInner({ parameter: { action: "config", code: "fl4x9" } }).getContent());
    say(r.customer ? "✅ הדף של דורית יקבל " + (r.customer.prices || []).length + " יעדים" : "⚠️ הדף של דורית לא יקבל לקוח (customer=null)");
  } catch (e) { say("❌ בקשת המחירון של דורית נכשלת: " + e + (e.stack ? " | " + String(e.stack).split("\n")[1] : "")); }
}

/* ---------- התראה לאייפון (ntfy) ---------- */
function notifyPhone(o) {
  if (PUSHOVER_TOKEN && PUSHOVER_USER) return notifyPushover(o);
  if (!NTFY_TOPIC) return false;
  try {
    const urgent = /דחוף/.test(String(o.when || ""));
    const casual = /^לקוח מזדמן/.test(String(o.customer || ""));
    const prepaid = /מראש/.test(String(o.pay || ""));
    const price = typeof o.price === "number" ? "₪" + o.price : String(o.price || "");
    const title = (o.resync ? "🔁 הזמנה שלא נקלטה – " : urgent ? "⚡ דחוף – " : casual ? "🚙 " : "🚙 הזמנה חדשה – ") + clean(o.customer);
    const message = [clean(o.recipient) + " ← " + clean(o.address), price, clean(o.pay)].filter(String).join(" · ") +
      (prepaid ? " – לבדוק תשלום" : "") + (o.when && !/^רגיל/.test(o.when) && !urgent ? "\n⏱ " + clean(o.when) : "");
    const res = UrlFetchApp.fetch("https://ntfy.sh/", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ topic: NTFY_TOPIC, title: title, message: message, priority: urgent ? 5 : 4, click: ADMIN_URL }),
      muteHttpExceptions: true,
    });
    return res.getResponseCode() < 300;
  } catch (e) {
    return false;
  }
}

/* ---------- Pushover: התראה לכל הזמנה, צלצול חירום לדחוף, "קיבלתי" מההתראה ---------- */
let lastPushError = "";
function notifyPushover(o) {
  lastPushError = "";
  try {
    const urgent = /דחוף/.test(String(o.when || ""));
    const casual = /^לקוח מזדמן/.test(String(o.customer || ""));
    const prepaid = /מראש/.test(String(o.pay || ""));
    const h = v => String(v === undefined || v === null ? "" : v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    const price = typeof o.price === "number" ? "₪" + o.price : String(o.price || "");
    const title = (o.resync ? "🔁 הזמנה שלא נקלטה – " : urgent ? "⚡ דחוף – " : casual ? "🚙 " : "🚙 הזמנה חדשה – ") + String(o.customer || "");
    const lines = [
      "<b>" + h(o.recipient) + "</b> ← " + h(o.address),
      [h(o.item), price, h(o.pay)].filter(String).join(" · ") + (prepaid ? " – <b>לבדוק תשלום</b>" : ""),
      o.when && !/^רגיל/.test(o.when) && !urgent ? "⏱ " + h(o.when) : "",
      o.notes ? "📝 " + h(o.notes) : "",
      '<a href="' + ADMIN_URL + '">פתיחת דף הניהול</a>',
    ].filter(String);
    const p = { token: PUSHOVER_TOKEN, user: PUSHOVER_USER, title: title.slice(0, 240), message: lines.join("\n").slice(0, 1000), html: "1" };
    const ack = o.no ? ackUrl(o.no) : "";
    if (ack) { p.url = ack; p.url_title = "👍 קיבלתי – הלקוח יראה שאישרת"; }
    if (urgent) {
      // צלצול חירום: חוזר כל דקה עד שמאשרים (עד 30 דקות). האישור מסמן "קיבלתי" לבד.
      p.priority = "2"; p.retry = "60"; p.expire = "1800";
      if (ack) p.callback = ack;
    } else {
      p.priority = "1";                                      // גבוה: מצלצל גם בשעות שקטות של האפליקציה
    }
    const res = UrlFetchApp.fetch("https://api.pushover.net/1/messages.json", { method: "post", payload: p, muteHttpExceptions: true });
    if (res.getResponseCode() >= 300) lastPushError = res.getResponseCode() + " " + res.getContentText().slice(0, 300);
    return res.getResponseCode() < 300;
  } catch (e) {
    lastPushError = String(e);
    return false;
  }
}

// קוד אישור לכל הזמנה (חתימה סודית) – רק מי שקיבל את ההתראה יכול לסמן "קיבלתי"
function ackSecret() {
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty("ACK_SECRET");
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props.setProperty("ACK_SECRET", s); }
  return s;
}
function ackToken(no) {
  const sig = Utilities.computeHmacSha256Signature(String(no), ackSecret());
  return Utilities.base64EncodeWebSafe(sig).replace(/=+$/, "").slice(0, 20);
}
function ackUrl(no) {
  let base = "";
  try { base = ScriptApp.getService().getUrl() || ""; } catch (e) {}
  if (!/\/exec$/.test(base)) return "";                    // רק מהפריסה הקבועה
  return base + "?action=ack&no=" + encodeURIComponent(no) + "&t=" + ackToken(no);
}
function ackOrder(q) {
  const no = str(q.no, 30);
  if (!no || String(q.t || "") !== ackToken(no)) return { ok: false, error: "קישור לא תקין" };
  const sheet = getSheet();
  const row = findRow(sheet, no);
  if (!row) return { ok: false, error: "ההזמנה לא נמצאה (אולי נמחקה)" };
  const cur = sheet.getRange(row, COL.status).getDisplayValue();
  if (cur === STATUSES[0]) { sheet.getRange(row, COL.status).setValue("אושרה"); addStatusDropdown(sheet, row); }
  return { ok: true, no: no, status: cur === STATUSES[0] ? "אושרה" : cur,
    recipient: sheet.getRange(row, COL.recipient).getDisplayValue(), dest: sheet.getRange(row, COL.dest).getDisplayValue() };
}
function ackPage(r) {
  const h = v => String(v || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const msg = r.ok
    ? (r.status === "אושרה" ? "סימנת <b>קיבלתי</b>.<br>הלקוח רואה עכשיו \"ארי אישר\"." : "ההזמנה כבר בסטטוס <b>" + h(r.status) + "</b> – לא שיניתי כלום.")
    : h(r.error);
  const html = '<!doctype html><html dir="rtl"><head><meta name="viewport" content="width=device-width,initial-scale=1"><meta charset="utf-8"></head>' +
    '<body style="font-family:-apple-system,Arial,sans-serif;text-align:center;padding:48px 20px;color:#111">' +
    '<div style="width:84px;height:84px;border-radius:50%;margin:0 auto 16px;display:flex;align-items:center;justify-content:center;font-size:44px;background:' + (r.ok ? "#e8f8ee" : "#fff1f2") + '">' + (r.ok ? "✅" : "⚠️") + '</div>' +
    (r.ok ? '<div style="font-size:15px;color:#70707a">' + h(r.recipient) + " · " + h(r.dest) + " · #" + h(r.no) + "</div>" : "") +
    '<p style="font-size:20px;line-height:1.5">' + msg + '</p>' +
    '<a href="' + ADMIN_URL + '" style="display:inline-block;margin-top:12px;background:#ff5a1f;color:#fff;text-decoration:none;padding:14px 22px;border-radius:14px;font-weight:600">פתיחת דף הניהול</a></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle("ארי משלוחים").addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// להרצה מהעורך: שולח התראת בדיקה לאייפון
function testNotification() {
  const ok = notifyPhone({ customer: "בדיקה", recipient: "התראת בדיקה", address: "אם קיבלת – זה עובד ✅", price: 0, pay: "", when: "רגיל" });
  Logger.log(ok ? "✅ התראה נשלחה – בדוק את האייפון"
    : PUSHOVER_TOKEN && PUSHOVER_USER ? "❌ Pushover לא שלח: " + lastPushError + " – בדוק את שני המפתחות (PUSHOVER_TOKEN / PUSHOVER_USER)"
    : "❌ ההתראה לא נשלחה – חסרים המפתחות של Pushover בראש הקוד");
}
// להרצה מהעורך: התראת חירום לבדיקה (מצלצלת כל דקה עד שמאשרים באפליקציה)
function testUrgentNotification() {
  const ok = notifyPhone({ customer: "בדיקה", recipient: "בדיקת חירום", address: "לאשר באפליקציה כדי להפסיק", price: 0, pay: "", when: "דחוף ⚡" });
  Logger.log(ok ? "✅ התראת חירום נשלחה – אשר אותה ב-Pushover" : "❌ לא נשלחה: " + lastPushError);
}
