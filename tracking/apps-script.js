// ═══════════════════════════════════════════════════════════
// MOBICA TRACKING AGENT — Google Apps Script v1.0
// ═══════════════════════════════════════════════════════════
const TELEGRAM_BOT_TOKEN = '8209404725:AAH1zZ_QzS3nDW01VMd0TLLHyvC1EYFPdh0';
const SHEET_ID = '1qZSLK7zEsOxRactIxXVwvVR3tFYruR_CI8bU5rERGwI';
const SHEET_UPDATES = 'TrackingUpdates';

const NOTIFY_IDS = {
  'أشرف قاسم':  '7055250567',
  'محمد فتحي':  '8562342656',
  'خالد علي':   '8017935026',
  'محمد مختار': '1242309147',
};

const TEAM_MANAGER_IDS = {
  'P1':     '1453588787', // سامى فؤاد
  '109962': '5555531128', // هشام جمال
  'P4':     '979158814',  // حازم قاعود
};

// ══════════════════════════════════════════════════════════
// ROUTER
// ══════════════════════════════════════════════════════════
function doGet(e) {
  const action  = e.parameter.action  || '';
  const payload = e.parameter.payload ? JSON.parse(e.parameter.payload) : {};
  const act     = payload.action || action;

  let result = {};
  if (act === 'addUpdate')    result = addUpdate(payload);
  else if (act === 'getUpdates') result = getUpdates();
  else result = { error: 'unknown action: ' + act };

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════
// GET UPDATES
// ══════════════════════════════════════════════════════════
function getUpdates() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_UPDATES);
  if (!sheet) return { updates: [] };

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const updates = [];

  for (let i = 1; i < data.length; i++) {
    const row = {};
    headers.forEach((h, j) => row[h] = data[i][j]);
    updates.push({
      type:            row.type,
      order_no:        String(row.order_no),
      customer:        row.customer,
      status:          row.status,
      pct:             row.pct,
      issue_type:      row.issue_type,
      severity:        row.severity,
      delivery_status: row.delivery_status,
      note:            row.note,
      by:              row.by,
      ts:              new Date(row.ts).getTime(),
    });
  }

  return { updates: updates.reverse() };
}

// ══════════════════════════════════════════════════════════
// ADD UPDATE
// ══════════════════════════════════════════════════════════
function addUpdate(body) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_UPDATES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_UPDATES);
    sheet.appendRow(['type','order_no','customer','status','pct',
                     'issue_type','severity','delivery_status','note','by','ts']);
  }

  sheet.appendRow([
    body.type             || '',
    body.order_no         || '',
    body.customer         || '',
    body.status           || '',
    body.pct              || 0,
    body.issue_type       || '',
    body.severity         || '',
    body.delivery_status  || '',
    body.note             || '',
    body.by               || '',
    new Date(body.ts)     || new Date(),
  ]);

  // Send Telegram notifications
  if (body.type === 'status') {
    notifyStatus(body);
  } else if (body.type === 'issue') {
    notifyIssue(body);
  } else if (body.type === 'delivery') {
    notifyDelivery(body);
  }

  return { ok: true };
}

// ══════════════════════════════════════════════════════════
// TELEGRAM NOTIFICATIONS
// ══════════════════════════════════════════════════════════
function notifyStatus(b) {
  const emoji = b.status==='Completed'?'✅':b.status==='Being processed'?'🔵':b.status==='Stopped'?'🔴':'⏳';
  const msg = `${emoji} *تحديث أمر تركيب*\n`
    + `📋 الأمر: ${b.order_no}\n`
    + `🏢 العميل: ${b.customer}\n`
    + `📊 الحالة: ${b.status}${b.pct?' ('+b.pct+'%)':''}\n`
    + `${b.note?'📝 '+b.note+'\n':''}`
    + `👤 بواسطة: ${b.by}`;

  // Notify all admins
  Object.values(NOTIFY_IDS).forEach(id => sendTelegram(id, msg));
}

function notifyIssue(b) {
  const sev = b.severity==='urgent'?'🔴 عاجلة':b.severity==='medium'?'🟡 متوسطة':'🟢 بسيطة';
  const msg = `⚠️ *مشكلة في التركيب*\n`
    + `📋 الأمر: ${b.order_no}\n`
    + `🏢 العميل: ${b.customer}\n`
    + `🔍 النوع: ${b.issue_type}\n`
    + `⚡ الخطورة: ${sev}\n`
    + `📝 ${b.note}\n`
    + `👤 أبلغ: ${b.by}`;

  // Notify all admins
  Object.values(NOTIFY_IDS).forEach(id => sendTelegram(id, msg));
}

function notifyDelivery(b) {
  const msg = `📦 *تأكيد توريد*\n`
    + `📋 الأمر: ${b.order_no}\n`
    + `🏢 العميل: ${b.customer}\n`
    + `📊 حالة التوريد: ${b.delivery_status}\n`
    + `${b.note?'📝 '+b.note+'\n':''}`
    + `👤 بواسطة: ${b.by}`;

  Object.values(NOTIFY_IDS).forEach(id => sendTelegram(id, msg));
}

// ══════════════════════════════════════════════════════════
// SEND TELEGRAM
// ══════════════════════════════════════════════════════════
function sendTelegram(chatId, text) {
  if (!chatId) return;
  try {
    UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chatId, text, parse_mode: 'Markdown' })
    });
  } catch(e) {}
}
