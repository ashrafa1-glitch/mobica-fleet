// ═══════════════════════════════════════════════════════════
// MOBICA SNAG SYSTEM — Google Apps Script v1.0
// ═══════════════════════════════════════════════════════════
const TELEGRAM_BOT_TOKEN = '8209404725:AAH1zZ_QzS3nDW01VMd0TLLHyvC1EYFPdh0';
const SHEET_ID           = '1qZSLK7zEsOxRactIxXVwvVR3tFYruR_CI8bU5rERGwI'; // نفس الـ Sheet أو جديد

const NOTIFY_IDS = {
  'أشرف قاسم':   '7055250567',
  'محمد فتحي':   '8562342656',
  'خالد علي':    '8017935026',
  'محمد مختار':  '',             // أضف رقمه
  'أحمد حسن':    '8583850058',
  'حازم قاعود':  '979158814',
  'هشام جمال':   '5555531128',
};

const INSTALLATION_MGR_IDS = ['7055250567','8562342656','8017935026']; // أشرف + محمد فتحي + خالد

const SHEET_SNAGS   = 'Snags';
const SHEET_HISTORY = 'SnagHistory';

// ══════════════════════════════════════════════════════════
// SETUP
// ══════════════════════════════════════════════════════════
function setupSnagSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  if (!ss.getSheetByName(SHEET_SNAGS)) {
    const sh = ss.insertSheet(SHEET_SNAGS);
    sh.appendRow(['id','date','client','contract','item','defect_type','severity',
      'status','team','created_by','created_at','closed_at',
      'root_cause','factory','analysis_desc','proposed_solution',
      'analyzed_by','approved_by','approval_note','notes']);
  }

  if (!ss.getSheetByName(SHEET_HISTORY)) {
    const sh = ss.insertSheet(SHEET_HISTORY);
    sh.appendRow(['snag_id','step','by','at','note']);
  }
}

// ══════════════════════════════════════════════════════════
// SNR ID
// ══════════════════════════════════════════════════════════
function generateSnrId(ss) {
  const sh   = ss.getSheetByName(SHEET_SNAGS);
  const rows = sh.getLastRow() - 1; // بدون الـ header
  const year = new Date().getFullYear();
  const num  = String(rows + 1).padStart(4, '0');
  return `SNR-${year}-${num}`;
}

// ══════════════════════════════════════════════════════════
// HELPERS
// ══════════════════════════════════════════════════════════
function getSnagSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  if (!ss.getSheetByName(SHEET_SNAGS)) setupSnagSheets();
  return { ss, sh: ss.getSheetByName(SHEET_SNAGS), hist: ss.getSheetByName(SHEET_HISTORY) };
}

function sendTelegramSnag(chatId, text) {
  if (!chatId) return;
  try {
    UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chatId, text, parse_mode: 'Markdown' })
    });
  } catch(e) { Logger.log('Telegram error: ' + e.message); }
}

function notifyManagers(msg) {
  INSTALLATION_MGR_IDS.forEach(id => sendTelegramSnag(id, msg));
}

function corsResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════
// doGet
// ══════════════════════════════════════════════════════════
function doGet(e) {
  try {
    const raw  = e.parameter.payload;
    const body = raw ? JSON.parse(raw) : {};
    const action = body.action || e.parameter.action || '';

    if (action === 'ping')        return corsResponse({ ok: true, msg: 'Snag System online' });
    if (action === 'setup')       { setupSnagSheets(); return corsResponse({ ok: true }); }
    if (action === 'addSnag')     return corsResponse(addSnag(body));
    if (action === 'getSnags')    return corsResponse(getSnags(body));
    if (action === 'updateSnag')  return corsResponse(updateSnag(body));
    if (action === 'getHistory')  return corsResponse(getHistory(body));

    return corsResponse({ ok: false, error: 'Unknown action: ' + action });
  } catch(err) {
    return corsResponse({ ok: false, error: err.message });
  }
}

// ══════════════════════════════════════════════════════════
// ADD SNAG
// ══════════════════════════════════════════════════════════
function addSnag(body) {
  const { ss, sh, hist } = getSnagSheet();
  const id  = generateSnrId(ss);
  const now = new Date().toISOString();

  sh.appendRow([
    id,
    Utilities.formatDate(new Date(), 'Africa/Cairo', 'yyyy-MM-dd'),
    body.client||'', body.contract||'', body.item||'',
    body.defect_type||'', body.severity||'', 'new',
    body.team||'', body.created_by||'', now, '',
    '','','','','','','', body.notes||''
  ]);

  hist.appendRow([id, 'new', body.created_by||'', now, 'تم تسجيل العيب']);

  // إشعار لمديري المجموعة + إدارة التركيبات
  const sevLabel = body.severity==='urgent'?'🔴 عاجل':body.severity==='medium'?'🟡 متوسط':'🟢 بسيط';
  const msg = `🆕 *عيب جديد — ${id}*\n🏢 ${body.client||'—'}\n📄 ${body.contract||'—'} | ${body.item||'—'}\n🔧 ${body.defect_type||'—'} | ${sevLabel}\n👷 سجّله: ${body.created_by||'—'}\n\n⚠️ يحتاج تحليل من مدير المجموعة`;

  notifyManagers(msg);

  // إشعار محمد مختار
  if (NOTIFY_IDS['محمد مختار']) sendTelegramSnag(NOTIFY_IDS['محمد مختار'], msg);

  return { ok: true, id };
}

// ══════════════════════════════════════════════════════════
// GET SNAGS
// ══════════════════════════════════════════════════════════
function getSnags(body) {
  const { sh, hist } = getSnagSheet();
  const data    = sh.getDataRange().getValues();
  const headers = data[0];
  const histData = hist.getDataRange().getValues();
  const histHeaders = histData[0];

  const snags = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i] instanceof Date
      ? Utilities.formatDate(row[i], 'Africa/Cairo', 'yyyy-MM-dd')
      : String(row[i]||''));

    // إلحاق الـ history
    const snagHist = histData.slice(1)
      .filter(h => h[0] === obj.id)
      .map(h => {
        const o = {};
        histHeaders.forEach((k,i)=>o[k]=String(h[i]||''));
        return o;
      });
    obj.history = snagHist;
    return obj;
  }).filter(s => s.id);

  return { ok: true, snags };
}

// ══════════════════════════════════════════════════════════
// UPDATE SNAG
// ══════════════════════════════════════════════════════════
function updateSnag(body) {
  const { sh, hist } = getSnagSheet();
  const data    = sh.getDataRange().getValues();
  const headers = data[0];
  const idCol   = headers.indexOf('id');

  const rowIdx = data.findIndex((r,i) => i>0 && r[idCol]===body.id);
  if (rowIdx < 0) return { ok: false, error: 'SNR not found: ' + body.id };

  // تحديث الـ fields
  if (body.fields) {
    Object.entries(body.fields).forEach(([key, val]) => {
      const col = headers.indexOf(key);
      if (col >= 0) sh.getRange(rowIdx+1, col+1).setValue(val);
    });
  }

  // إضافة history
  if (body.history_entry) {
    const h = body.history_entry;
    hist.appendRow([body.id, h.step||'', h.by||'', h.at||new Date().toISOString(), h.note||'']);
  }

  // إشعارات بعد التحديث
  const status = body.fields?.status;
  const snag = data[rowIdx];
  const snagId = body.id;
  const client = String(snag[headers.indexOf('client')]||'');
  const sevCol = headers.indexOf('severity');
  const sev = sevCol>=0?String(snag[sevCol]||''):'';
  const sevLabel = sev==='urgent'?'🔴 عاجل':sev==='medium'?'🟡 متوسط':'🟢 بسيط';

  if (status === 'analyzed') {
    const msg = `🔍 *تم تحليل العيب — ${snagId}*\n🏢 ${client} | ${sevLabel}\n👷 حلّله: ${body.fields.analyzed_by||'—'}\n🎯 السبب: ${body.fields.root_cause||'—'}\n\n✅ يحتاج اعتماد من إدارة التركيبات`;
    notifyManagers(msg);
  }

  if (status === 'approved') {
    const msg = `✅ *تم اعتماد العيب — ${snagId}*\n🏢 ${client} | ${sevLabel}\n👤 اعتمده: ${body.fields.approved_by||'—'}\n\n📧 سيتم إرسال إيميل للجودة والبيع تلقائياً`;
    notifyManagers(msg);
    if (NOTIFY_IDS['محمد مختار']) sendTelegramSnag(NOTIFY_IDS['محمد مختار'], msg);
    // TODO: sendEmail to quality + sales
  }

  if (status === 'in_production') {
    const msg = `🏭 *بدأ التصنيع — ${snagId}*\n🏢 ${client}\n👤 بواسطة: ${body.history_entry?.by||'—'}\n⏱️ SLA: 7 أيام`;
    if (NOTIFY_IDS['محمد مختار']) sendTelegramSnag(NOTIFY_IDS['محمد مختار'], msg);
    notifyManagers(msg);
  }

  if (status === 'ready') {
    const msg = `📦 *جاهز للتوريد — ${snagId}*\n🏢 ${client}\n✅ يرجى إبلاغ إدارة التسليمات`;
    notifyManagers(msg);
    if (NOTIFY_IDS['محمد مختار']) sendTelegramSnag(NOTIFY_IDS['محمد مختار'], msg);
  }

  if (status === 'closed') {
    const msg = `🟢 *تم إغلاق العيب — ${snagId}*\n🏢 ${client}\n✅ اكتملت معالجة العيب بنجاح`;
    notifyManagers(msg);
  }

  if (status === 'rejected') {
    const msg = `❌ *تم رفض العيب — ${snagId}*\n🏢 ${client}\n📝 السبب: ${body.fields.approval_note||'—'}\n👤 رفضه: ${body.fields.approved_by||'—'}`;
    notifyManagers(msg);
  }

  return { ok: true };
}

// ══════════════════════════════════════════════════════════
// GET HISTORY
// ══════════════════════════════════════════════════════════
function getHistory(body) {
  const { hist } = getSnagSheet();
  const data    = hist.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1).map(r => {
    const obj = {};
    headers.forEach((h,i)=>obj[h]=String(r[i]||''));
    return obj;
  });
  const filtered = body.snag_id ? rows.filter(r=>r.snag_id===body.snag_id) : rows;
  return { ok: true, history: filtered };
}

// ══════════════════════════════════════════════════════════
// SLA CHECKER — شغّله كـ Time-based trigger كل يوم
// ══════════════════════════════════════════════════════════
function checkSLA() {
  const { sh } = getSnagSheet();
  const data    = sh.getDataRange().getValues();
  const headers = data[0];
  const now     = new Date();

  const SLA_DAYS = { urgent: 3, medium: 7, low: 14 };

  data.slice(1).forEach(row => {
    const status    = String(row[headers.indexOf('status')]||'');
    const severity  = String(row[headers.indexOf('severity')]||'');
    const createdAt = row[headers.indexOf('created_at')];
    const id        = String(row[headers.indexOf('id')]||'');
    const client    = String(row[headers.indexOf('client')]||'');

    if (!id || status === 'closed' || status === 'rejected') return;

    const created = new Date(createdAt);
    const days    = Math.floor((now - created) / (1000*60*60*24));
    const sla     = SLA_DAYS[severity] || 7;

    if (days >= sla) {
      const msg = `⚠️ *تجاوز SLA — ${id}*\n🏢 ${client}\n⏱️ مضى ${days} يوم (SLA: ${sla} يوم)\n📌 الحالة: ${status}\n\n🔴 يحتاج متابعة عاجلة!`;
      notifyManagers(msg);
    }
  });
}

// ══════════════════════════════════════════════════════════
// SETUP TRIGGERS (شغّله مرة واحدة)
// ══════════════════════════════════════════════════════════
function createSnagTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t=>t.getHandlerFunction()==='checkSLA')
    .forEach(t=>ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('checkSLA')
    .timeBased().everyDays(1).atHour(9).create();

  Logger.log('✅ SLA trigger created');
}
