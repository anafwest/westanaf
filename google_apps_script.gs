/**********************************************************************
 * خادم المزامنة — تطبيق تدقيق رخص البناء (إدارة رخص البناء – غرب الرياض)
 * -------------------------------------------------------------------
 * خطوات التركيب (مرة واحدة فقط):
 *
 * 1) افتح الرابط:  https://sheets.new
 *    وسمِّ الجدول مثلاً: "بيانات تدقيق رخص البناء".
 *
 * 2) من القائمة العلوية:  الإضافات (Extensions)  →  Apps Script.
 *
 * 3) امسح الكود الموجود في المحرر والصق هذا الملف كاملاً.
 *
 * 4) اضغط "حفظ" (أيقونة القرص) ثم زر "نشر" (Deploy) → "نشر جديد"
 *    (New deployment) → اختر النوع (Select type): "تطبيق ويب" (Web app) ثم:
 *        - الوصف (Description): تدقيق رخص البناء
 *        - التنفيذ باسم (Execute as):  أنا (Me)
 *        - من يملك الوصول (Who has access):  أي شخص (Anyone)
 *    ثم اضغط "نشر"، واسمح بالصلاحيات عند الطلب.
 *
 * 5) انسخ الرابط الناتج (ينتهي بـ /exec) وأرسله لي،
 *    أو ألصقه في التطبيق عبر زر ⚙️ إعدادات المزامنة.
 *
 * ملاحظة: الرمز SYNC_TOKEN يجب أن يطابق CLOUD_TOKEN في index.html.
 **********************************************************************/

const SYNC_TOKEN = 'rm-west-2026'; // رمز الأمان — يجب أن يطابق التطبيق

function doGet(e) {
  try {
    if (!e || !e.parameter || e.parameter.token !== SYNC_TOKEN) {
      return json_({ ok: false, error: 'رمز غير صحيح' });
    }
    return json_({ ok: true, auditors: readAuditors_(), history: readHistory_() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.token !== SYNC_TOKEN) return json_({ ok: false, error: 'رمز غير صحيح' });
    const action = body.action;
    const data = body.data || {};
    if (action === 'upsertHistory') { upsertHistory_(data); return json_({ ok: true }); }
    if (action === 'deleteHistory') { deleteRowById_('History', data.id); return json_({ ok: true }); }
    if (action === 'upsertAuditor') { upsertAuditor_(data); return json_({ ok: true }); }
    if (action === 'deleteAuditor') { deleteRowById_('Auditors', data.id); return json_({ ok: true }); }
    return json_({ ok: false, error: 'أمر غير معروف: ' + action });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (name === 'History') sh.appendRow(['id', 'name', 'date', 'auditorId', 'auditorName', 'content', 'structured', 'updatedAt']);
    if (name === 'Auditors') sh.appendRow(['id', 'name', 'role', 'createdAt', 'updatedAt']);
  }
  return sh;
}

// ==================== سجل التدقيقات ====================
function upsertHistory_(h) {
  const sh = sheet_('History');
  const values = sh.getDataRange().getValues();
  const idS = String(h.id);
  const row = [h.id, h.name || '', h.date || '', h.auditorId || '', h.auditorName || '', h.content || '', JSON.stringify(h.structured || null), new Date()];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === idS) {
      sh.getRange(i + 1, 1, 1, 8).setValues([row]);
      return;
    }
  }
  sh.appendRow(row);
}

function readHistory_() {
  const values = sheet_('History').getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (r[0] === '' || r[0] === null || r[0] === undefined) continue;
    let structured = null;
    try { structured = r[6] ? JSON.parse(r[6]) : null; } catch (e) {}
    out.push({
      id: (!isNaN(Number(r[0])) ? Number(r[0]) : String(r[0])),
      name: r[1], date: r[2], auditorId: r[3], auditorName: r[4],
      content: r[5], structured: structured
    });
  }
  return out;
}

// ==================== المدققون ====================
function upsertAuditor_(a) {
  const sh = sheet_('Auditors');
  const values = sh.getDataRange().getValues();
  const idS = String(a.id);
  const row = [a.id, a.name || '', a.role || 'auditor', a.createdAt || new Date(), new Date()];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === idS) {
      sh.getRange(i + 1, 1, 1, 5).setValues([row]);
      return;
    }
  }
  sh.appendRow(row);
}

function readAuditors_() {
  const values = sheet_('Auditors').getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (r[0] === '' || r[0] === null || r[0] === undefined) continue;
    out.push({ id: String(r[0]), name: r[1], role: r[2] || 'auditor', createdAt: r[3] });
  }
  return out;
}

// ==================== أدوات ====================
function deleteRowById_(sheetName, id) {
  const sh = sheet_(sheetName);
  const values = sh.getDataRange().getValues();
  const idS = String(id);
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][0]) === idS) sh.deleteRow(i + 1);
  }
}
