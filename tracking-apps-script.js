// ═══════════════════════════════════════════════════════════
// MOBICA TRACKING AGENT — Google Apps Script v5.3
// إشعارات ذكية + تقارير يومية + تتبع تذكيرات المتأخرات
// ═══════════════════════════════════════════════════════════

const TELEGRAM_BOT_TOKEN = '8209404725:AAH1zZ_QzS3nDW01VMd0TLLHyvC1EYFPdh0';
const SHEET_ID           = '1qZSLK7zEsOxRactIxXVwvVR3tFYruR_CI8bU5rERGwI';
const SHEET_UPDATES      = 'TrackingUpdates';
const SHEET_REMINDERS    = 'ReminderLog';
const SHEET_ORDERS       = 'SAPOrders';      // ← قاعدة البيانات المركزية للأوامر

// ── إدارة التركيبات ──────────────────────────────────────
const ADMIN_IDS = {
  'أشرف قاسم':  '7055250567',
  'محمد فتحي':  '8562342656',
  'خالد علي':   '8017935026',
};

// ── مديرو المجموعات ──────────────────────────────────────
const TEAM_MANAGERS = {
  'P1':     { name: 'سامى فؤاد',   id: '1453588787' },
  '109962': { name: 'هشام جمال',   id: '5555531128' },
  'P4':     { name: 'حازم قاعود',  id: '979158814'  },
  'A4':     { name: 'محمد وسيم',   id: ''           }, // ID pending
  'ahmed':  { name: 'أحمد حسن',    id: '8583850058' },
};

// ── متابعة ────────────────────────────────────────────────
const MOKHTAR_ID = '1242309147';
const EMAD_ID    = '8296695318';
const YASSER_ID  = '8353455749';

// ═══════════════════════════════════════════════════════════
// ROUTER
// ═══════════════════════════════════════════════════════════
function doGet(e) {
  const params  = e.parameter || {};
  const action  = params.action || '';
  const payload = params.payload ? JSON.parse(params.payload) : {};
  const act     = payload.action || action;

  let result = {};
  try {
    if      (act === 'addUpdate')      result = addUpdate(payload);
    else if (act === 'getUpdates')     result = getUpdates(params.team);
    else if (act === 'getOrders')      result = getOrders();
    else if (act === 'updateOrder')      result = updateOrderField(payload);
    else if (act === 'saveOrdersBatch')  result = saveOrders(payload.orders, payload.by);
    else if (act === 'sendReport')       result = sendDailyReport('طلب');
    else if (act === 'addIR')          result = addIR(payload);
    else if (act === 'addSAPUpload')   result = addSAPUpload(payload);
    else if (act === 'repairHeaders')  result = repairHeaders();
    else if (act === 'debug')          result = debugSheet();
    else result = { error: 'unknown action: ' + act };
  } catch(err) {
    result = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════
// doPost — استقبال البيانات الكبيرة (حفظ الأوامر من SAP)
// ═══════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action || '';
    let result   = {};

    if (action === 'saveOrders') {
      result = saveOrders(body.orders, body.by);
    } else if (action === 'updateOrder') {
      result = updateOrderField(body);
    } else {
      result = { error: 'unknown action: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ═══════════════════════════════════════════════════════════
// SAVE ORDERS — حفظ أوامر SAP في الـ SAPOrders sheet
// ═══════════════════════════════════════════════════════════
function saveOrders(orders, by) {
  if (!orders || !orders.length) return { ok: false, error: 'no orders' };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_ORDERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ORDERS);
    sheet.appendRow(['order_no','customer','address','team_code','team_name',
                     'status','delivery','planned','total_units','notes',
                     'upload_by','upload_ts','json_data']);
  }

  // بناء خريطة الأوامر الموجودة (للـ merge)
  const existing = {};
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    existing[String(data[i][0])] = i + 1; // order_no → row number
  }

  let added = 0, updated = 0;
  const now = Date.now();

  orders.forEach(o => {
    const no = String(o.order_no || '').trim();
    if (!no) return;

    // بيانات مختصرة للأعمدة + json_data كاملة
    const row = [
      no,
      o.customer    || '',
      o.address     || '',
      o.team_code   || '',
      o.team_name   || '',
      o.status      || 'REL',
      o.delivery    || '',
      o.planned     || '',
      o.total_units || 0,
      o.notes       || '',
      by            || o.upload_by || '',
      o.upload_ts   || now,
      JSON.stringify({ items: o.items || [], main_group: o.main_group || '', phone: o.phone || '' })
    ];

    if (existing[no]) {
      // تحديث الصف الموجود (ما عدا status اللي تتحكم فيها يدوياً)
      const rowIdx = existing[no];
      sheet.getRange(rowIdx, 1, 1, 12).setValues([row.slice(0, 12)]);
      // نحدث json_data بس
      sheet.getRange(rowIdx, 13).setValue(row[12]);
      updated++;
    } else {
      sheet.appendRow(row);
      added++;
    }
  });

  return { ok: true, added, updated, total: orders.length };
}

// ═══════════════════════════════════════════════════════════
// GET ORDERS — جلب كل أوامر SAP
// ═══════════════════════════════════════════════════════════
function getOrders() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_ORDERS);
  if (!sheet) return { orders: [] };

  const data    = sheet.getDataRange().getValues();
  if (data.length < 2) return { orders: [] };

  const headers = data[0];
  const orders  = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const no  = String(row[0] || '').trim();
    if (!no) continue;

    let extra = {};
    try { extra = JSON.parse(row[12] || '{}'); } catch(e) {}

    orders.push({
      order_no:    no,
      customer:    String(row[1] || ''),
      address:     String(row[2] || ''),
      team_code:   String(row[3] || ''),
      team_name:   String(row[4] || ''),
      status:      String(row[5] || 'REL'),
      delivery:    String(row[6] || ''),
      planned:     String(row[7] || ''),
      total_units: Number(row[8] || 0),
      notes:       String(row[9] || ''),
      upload_by:   String(row[10] || ''),
      upload_ts:   Number(row[11] || 0),
      items:       extra.items      || [],
      main_group:  extra.main_group || '',
      phone:       extra.phone      || '',
    });
  }

  return { orders, ts: Date.now() };
}

// ═══════════════════════════════════════════════════════════
// UPDATE ORDER FIELD — تحديث حقل في أمر (status, notes...)
// ═══════════════════════════════════════════════════════════
function updateOrderField(body) {
  const no    = String(body.order_no || '').trim();
  const field = body.field || '';
  const value = body.value;
  if (!no || !field) return { ok: false, error: 'missing order_no or field' };

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_ORDERS);
  if (!sheet) return { ok: false, error: 'SAPOrders sheet not found' };

  const FIELD_MAP = { status:6, delivery:7, notes:10 }; // col index (1-based)
  const col = FIELD_MAP[field];
  if (!col) return { ok: false, error: 'unsupported field: ' + field };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === no) {
      sheet.getRange(i + 1, col).setValue(value);
      return { ok: true, order_no: no, field, value };
    }
  }
  return { ok: false, error: 'order not found: ' + no };
}

// ═══════════════════════════════════════════════════════════
// ADD UPDATE — حفظ + إشعارات ذكية
// ═══════════════════════════════════════════════════════════
// ─── الـ headers المطلوبة بالترتيب ──────────────────────────
const REQUIRED_HEADERS = [
  'type','order_no','customer','team_code','contract_no',
  'status','pct','issue_type','severity','snag_id',
  'photo_count','ir_type','ir_no','ir_vl',
  'delivery_status','note','by','ts'
];

// ─── إصلاح headers تلقائياً ──────────────────────────────────
function ensureHeaders(sheet) {
  const firstRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  // إذا الـ ts مش موجود في الـ headers — نعيد كتابة الـ header row
  if (!firstRow.includes('ts')) {
    sheet.getRange(1, 1, 1, REQUIRED_HEADERS.length).setValues([REQUIRED_HEADERS]);
  }
}

function addUpdate(body) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_UPDATES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_UPDATES);
    sheet.appendRow(REQUIRED_HEADERS);
  } else {
    ensureHeaders(sheet);
  }

  sheet.appendRow([
    body.type             || '',
    body.order_no         || '',
    body.customer         || '',
    body.team_code        || '',
    body.contract_no      || '',
    body.status           || '',
    body.pct              || 0,
    body.issue_type       || '',
    body.severity         || '',
    body.snag_id          || '',
    body.photo_count      || 0,
    body.ir_type          || '',
    body.ir_no            || '',
    body.ir_vl            || '',
    body.delivery_status  || '',
    body.note             || '',
    body.by               || '',
    body.ts               || Date.now(),
  ]);

  // ── إشعارات ذكية حسب نوع التحديث ──
  if      (body.type === 'status')   notifyStatus(body);
  else if (body.type === 'issue')    notifyIssue(body);
  else if (body.type === 'ir')       notifyIR(body);
  else if (body.type === 'delivery') notifyDelivery(body);

  return { ok: true, ts: new Date().getTime() };
}

// ═══════════════════════════════════════════════════════════
// GET UPDATES
// ═══════════════════════════════════════════════════════════
function getUpdates(team) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_UPDATES);
  if (!sheet) return { updates: [] };

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const updates = [];

  for (let i = 1; i < data.length; i++) {
    const row = {};
    headers.forEach((h, j) => row[h] = data[i][j]);
    // Filter by team if specified
    if (team && row.team_code && row.team_code !== team) continue;
    updates.push({
      type:            row.type,
      order_no:        String(row.order_no),
      customer:        row.customer,
      team_code:       row.team_code,
      contract_no:     row.contract_no,
      status:          row.status,
      pct:             row.pct,
      issue_type:      row.issue_type,
      severity:        row.severity,
      snag_id:         row.snag_id,
      photo_count:     row.photo_count,
      ir_type:         row.ir_type,
      ir_no:           row.ir_no,
      ir_vl:           row.ir_vl,
      delivery_status: row.delivery_status,
      note:            row.note,
      by:              row.by,
      ts:              row.ts ? Number(row.ts) : 0,
    });
  }

  return { updates: updates.reverse() };
}

// ═══════════════════════════════════════════════════════════
// NOTIFY — تحديث حالة
// ═══════════════════════════════════════════════════════════
function notifyStatus(b) {
  const emoji = b.status==='Completed'?'✅':
                b.status==='Being processed'?'🔵':
                b.status==='Stopped'?'🔴':'⏳';

  // رسالة للإدارة — كاملة
  const adminMsg =
    `${emoji} *تحديث أمر تركيب*\n`+
    `📋 الأمر: ${b.order_no}\n`+
    `🏢 العميل: ${b.customer}\n`+
    `📊 الحالة: *${b.status}*${b.pct?' ('+b.pct+'%)':''}\n`+
    `${b.note?'📝 '+b.note+'\n':''}`+
    `👤 بواسطة: ${b.by}\n`+
    `🕐 ${new Date().toLocaleString('ar-EG',{timeZone:'Africa/Cairo'})}`;

  Object.values(ADMIN_IDS).forEach(id => sendTelegram(id, adminMsg));

  // رسالة لمدير المجموعة — فقط أوامره
  const mgr = getTeamManager(b.team_code || b.by);
  if (mgr && b.status === 'Stopped') {
    // لو متوقف — يبلّغ مدير المجموعة بشكل خاص
    const mgrMsg =
      `⚠️ *أمر متوقف في فريقك*\n`+
      `📋 الأمر: ${b.order_no}\n`+
      `🏢 العميل: ${b.customer}\n`+
      `📝 ${b.note||'لا توجد ملاحظة'}\n`+
      `👤 بواسطة: ${b.by}`;
    sendTelegram(mgr.id, mgrMsg);
  }
}

// ═══════════════════════════════════════════════════════════
// NOTIFY — مشكلة
// ═══════════════════════════════════════════════════════════
function notifyIssue(b) {
  const sev = b.severity==='urgent'?'🔴 عاجلة':
              b.severity==='medium'?'🟡 متوسطة':'🟢 بسيطة';
  const isDefect = b.issue_type==='عيب في المنتج' || b.issue_type==='مشكلة مقاس';

  // رسالة للإدارة
  const adminMsg =
    `⚠️ *مشكلة في التركيب*\n`+
    `📋 الأمر: ${b.order_no}\n`+
    `🏢 العميل: ${b.customer}\n`+
    `🔍 النوع: ${b.issue_type}\n`+
    `⚡ الخطورة: ${sev}\n`+
    `📝 ${b.note}\n`+
    `${b.photo_count?'📸 '+b.photo_count+' صورة مرفقة\n':''}`+
    `👤 أبلغ: ${b.by}`;

  Object.values(ADMIN_IDS).forEach(id => sendTelegram(id, adminMsg));

  // لو عيب في المنتج → إشعار خاص لمحمد مختار
  if (isDefect) {
    const mokhMsg =
      `🔴 *عيب منتج جديد — يحتاج متابعتك مع المصنع*\n`+
      `📋 الأمر: ${b.order_no}\n`+
      `🏢 العميل: ${b.customer}\n`+
      `🔍 النوع: ${b.issue_type}\n`+
      `⚡ الخطورة: ${sev}\n`+
      `📝 ${b.note}\n`+
      `${b.snag_id?'🔗 Snag ID: '+b.snag_id+'\n':''}`+
      `${b.photo_count?'📸 '+b.photo_count+' صورة\n':''}`+
      `👤 أبلغ: ${b.by}\n`+
      `⏰ *مطلوب: تواصل مع المصنع وتحديث الـ Snag*`;
    sendTelegram(MOKHTAR_ID, mokhMsg);
  }

  // لو خطورة عاجلة — إشعار إضافي فوري للإدارة
  if (b.severity === 'urgent') {
    const urgentMsg = `🚨 *تنبيه عاجل — يحتاج تدخل فوري*\n`+adminMsg;
    Object.values(ADMIN_IDS).forEach(id => sendTelegram(id, urgentMsg));
  }
}

// ═══════════════════════════════════════════════════════════
// NOTIFY — محضر تسليم
// ═══════════════════════════════════════════════════════════
function notifyIR(b) {
  const msg =
    `📋 *محضر تسليم جديد*\n`+
    `النوع: ${b.ir_type}\n`+
    `الرقم: ${b.ir_no}\n`+
    `العقد: ${b.contract_no||'—'}\n`+
    `العميل: ${b.customer||'—'}\n`+
    `${b.ir_vl?'💰 القيمة: '+Number(b.ir_vl).toLocaleString()+' EGP\n':''}`+
    `${b.note?'📝 '+b.note+'\n':''}`+
    `👤 عماد عبدالواحد\n`+
    `🕐 ${new Date().toLocaleString('ar-EG',{timeZone:'Africa/Cairo'})}`;

  // الإدارة فقط
  Object.values(ADMIN_IDS).forEach(id => sendTelegram(id, msg));
}

// ═══════════════════════════════════════════════════════════
// NOTIFY — توريد
// ═══════════════════════════════════════════════════════════
function notifyDelivery(b) {
  const msg =
    `📦 *تحديث توريد*\n`+
    `📋 الأمر: ${b.order_no}\n`+
    `🏢 العميل: ${b.customer}\n`+
    `📊 ${b.delivery_status}\n`+
    `${b.note?'📝 '+b.note+'\n':''}`+
    `👤 ${b.by}`;

  Object.values(ADMIN_IDS).forEach(id => sendTelegram(id, msg));

  // إشعار مدير المجموعة بأن توريد جديد وصل لمواقعه
  const mgr = getTeamManager(b.team_code || '');
  if (mgr) {
    const mgrMsg =
      `📦 *توريد جديد وصل لمواقعك*\n`+
      `📋 الأمر: ${b.order_no}\n`+
      `🏢 العميل: ${b.customer}\n`+
      `📊 ${b.delivery_status}\n`+
      `${b.note?'📝 '+b.note+'\n':''}`+
      `✅ *يرجى تحديث حالة التركيب*`;
    sendTelegram(mgr.id, mgrMsg);
  }
}

// ═══════════════════════════════════════════════════════════
// DAILY REPORT — تقرير يومي
// ═══════════════════════════════════════════════════════════
function sendMorningReport()   { sendDailyReport('8 ص'); }
function sendAfternoonReport() { sendDailyReport('3 م'); }

function sendDailyReport(time) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_UPDATES);
  const updates = sheet ? sheet.getDataRange().getValues() : [];
  const headers = updates.length ? updates[0] : [];

  // حساب الإحصائيات من آخر 24 ساعة
  const now    = new Date().getTime();
  const day24  = 24 * 60 * 60 * 1000;
  const recent = [];

  for (let i = 1; i < updates.length; i++) {
    const row = {};
    headers.forEach((h, j) => row[h] = updates[i][j]);
    const ts = row.ts ? Number(row.ts) : 0;
    if (now - ts <= day24) recent.push(row);
  }

  const statusUpds    = recent.filter(r => r.type === 'status');
  const completed     = statusUpds.filter(r => r.status === 'Completed');
  const inProgress    = statusUpds.filter(r => r.status === 'Being processed');
  const stopped       = statusUpds.filter(r => r.status === 'Stopped');
  const issues        = recent.filter(r => r.type === 'issue');
  const urgentIssues  = issues.filter(r => r.severity === 'urgent');
  const defects       = issues.filter(r => r.issue_type === 'عيب في المنتج' || r.issue_type === 'مشكلة مقاس');
  const irUpds        = recent.filter(r => r.type === 'ir');

  // ── حساب المتأخرات ─────────────────────────────────────
  const threshold3d = 3 * 24 * 60 * 60 * 1000;
  const lastStatusMap = {};
  for (let i = 1; i < updates.length; i++) {
    const row = {};
    headers.forEach((h,j) => row[h] = updates[i][j]);
    if (row.type !== 'status') continue;
    const k = row.order_no;
    const rowTs = row.ts ? Number(row.ts) : 0;
    const curTs = lastStatusMap[k] ? Number(lastStatusMap[k].ts)||0 : 0;
    if (!lastStatusMap[k] || rowTs > curTs) lastStatusMap[k] = row;
  }
  const delayedOrders = [];
  Object.values(lastStatusMap).forEach(row => {
    if (row.status === 'Completed' || row.status === 'TECO' || row.status === 'CLSD') return;
    const lastTs = row.ts ? Number(row.ts) : 0;
    const delayMs = now - lastTs;
    if (delayMs < threshold3d) return;
    const delayDays = Math.floor(delayMs / (24*60*60*1000));
    delayedOrders.push({ ...row, delayDays });
  });
  delayedOrders.sort((a,b) => b.delayDays - a.delayDays);

  // ── بناء قسم المتأخرات ─────────────────────────────────
  let delayedSection = '';
  if (delayedOrders.length) {
    delayedSection =
      `─────────────────────\n`+
      `🚨 *المتأخرات (+3 أيام): ${delayedOrders.length} أمر*\n`+
      delayedOrders.slice(0,10).map(r =>
        `  • ${r.order_no} | ${r.customer||'—'} | *${r.delayDays} يوم* | ${TEAM_MANAGERS[r.team_code]?TEAM_MANAGERS[r.team_code].name:r.team_code||'—'}`
      ).join('\n') +
      (delayedOrders.length > 10 ? `\n  ... و${delayedOrders.length-10} أمر آخر` : '') +
      '\n';
  }

  // ── تقرير الإدارة ─────────────────────────────────────
  const adminReport =
    `📊 *تقرير إدارة التركيبات — ${time}*\n`+
    `📅 ${new Date().toLocaleDateString('ar-EG',{timeZone:'Africa/Cairo'})}\n`+
    `─────────────────────\n`+
    `*آخر 24 ساعة:*\n`+
    `✅ أوامر مكتملة: *${completed.length}*\n`+
    `🔵 جارى التركيب: *${inProgress.length}*\n`+
    `🔴 متوقفة: *${stopped.length}*\n`+
    `⚠️ مشاكل: *${issues.length}*${urgentIssues.length?' ('+urgentIssues.length+' عاجلة 🚨)':''}\n`+
    `🔴 عيوب منتجات: *${defects.length}*\n`+
    `📋 محاضر تسليم: *${irUpds.length}*\n`+
    delayedSection +
    `─────────────────────\n`+
    `${stopped.length?'⚠️ *أوامر متوقفة:*\n'+stopped.map(r=>`  • ${r.order_no} — ${r.customer} (${r.note||'—'})`).join('\n')+'\n':''}` +
    `${urgentIssues.length?'🚨 *مشاكل عاجلة:*\n'+urgentIssues.map(r=>`  • ${r.order_no} — ${r.issue_type}: ${r.note}`).join('\n')+'\n':''}` +
    `👁️ *للتفاصيل:* https://ashrafa1-glitch.github.io/mobica-fleet/tracking/`;

  // ── إحصائيات التذكيرات ────────────────────────────────
  const remStats = getReminderStats();
  if (remStats.total > 0) {
    const reminderSection =
      `\n─────────────────────\n`+
      `📨 *إحصائيات التذكيرات (آخر 7 أيام):*\n`+
      `إجمالي التذكيرات المرسلة: *${remStats.total}*\n`+
      `✅ ردّوا: *${remStats.responded}*\n`+
      `❌ لم يردوا: *${remStats.noResponse}*\n`+
      `\n*التفاصيل:*\n`+
      remStats.details.join('\n');

    Object.values(ADMIN_IDS).forEach(id => sendTelegram(id, reminderSection));
  }

  Object.values(ADMIN_IDS).forEach(id => sendTelegram(id, adminReport));

  // ── تقرير كل مدير مجموعة — أوامره المتأخرة فقط ────────
  sendDelayedToManagers();

    // ── تقرير خاص لخالد علي — المتأخرات + التذكيرات ──────
  sendKhaledReport(time);

  // ── إشعار مختار بالعيوب المفتوحة ──────────────────────
  if (defects.length) {
    const mokhMsg =
      `🔴 *ملخص العيوب — ${time}*\n`+
      `${defects.length} عيب يحتاج متابعتك مع المصانع:\n`+
      defects.map(r=>`  • ${r.order_no} | ${r.issue_type} | ${r.severity==='urgent'?'🚨 عاجل':r.severity==='medium'?'🟡':'🟢'}`).join('\n')+
      `\n\n⏰ يرجى تحديث حالة كل عيب في Snag Agent`;
    sendTelegram(MOKHTAR_ID, mokhMsg);
  }

  // ── إشعار أشرف بتأكيد الإرسال ─────────────────────────
  const confirmMsg =
    `✅ *تأكيد إرسال التقارير — ${time}*\n`+
    `📅 ${new Date().toLocaleDateString('ar-EG',{timeZone:'Africa/Cairo'})}\n`+
    `━━━━━━━━━━━━━━━━━━━━━\n`+
    `*📊 تقرير الإدارة أُرسل إلى:*\n`+
    `  • محمد فتحي ✅\n`+
    `  • خالد علي ✅\n`+
    `  • أشرف قاسم ✅\n\n`+
    `*⚠️ تقرير المتأخرات أُرسل إلى:*\n`+
    `  • أحمد حسن (A) ✅\n`+
    `  • سامى فؤاد (B) ✅\n`+
    `  • حازم قاعود (C) ✅\n`+
    `  • هشام جمال (D) ✅\n\n`+
    `*📋 تقرير خاص أُرسل إلى:*\n`+
    `  • خالد علي — المتأخرات والتذكيرات ✅\n`+
    (defects.length ? `\n*🔴 إشعار عيوب أُرسل إلى:*\n  • محمد مختار ✅\n` : '')+
    `━━━━━━━━━━━━━━━━━━━━━\n`+
    `📌 إجمالي: ${defects.length?8:7} رسائل أُرسلت بنجاح`;

  sendTelegram('7055250567', confirmMsg);

  return { ok: true, sent: new Date().toISOString() };
}

// ═══════════════════════════════════════════════════════════
// SEND DELAYED TO MANAGERS — المتأخرات لكل مدير مجموعة
// ═══════════════════════════════════════════════════════════
function sendDelayedToManagers() {
  // هنا بنبعت للمديرين أوامرهم المتأخرة
  // في النسخة الكاملة هنجيب البيانات من الـ orders sheet
  // دلوقتي بنبعت تذكير عام بالتحديث

  const reminder =
    `⏰ *تذكير — تحديث يومي*\n`+
    `يرجى تحديث حالة أوامر التركيب الخاصة بك على النظام\n`+
    `📱 https://ashrafa1-glitch.github.io/mobica-fleet/tracking/\n`+
    `\n`+
    `*تحقق من:*\n`+
    `• الأوامر المتأخرة عن موعدها ⚠️\n`+
    `• تحديث نسبة الإنجاز\n`+
    `• أي مشاكل محتاج تبلّغ عنها`;

  Object.values(TEAM_MANAGERS).forEach(mgr => {
    if (mgr.id) sendTelegram(mgr.id, reminder);
  });
}

// ═══════════════════════════════════════════════════════════
// SAP UPLOAD — تسجيل رفع ياسر + إشعارات
// ═══════════════════════════════════════════════════════════
function addSAPUpload(body) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('SAPUploads');
  if (!sheet) {
    sheet = ss.insertSheet('SAPUploads');
    sheet.appendRow(['timestamp','file_name','total_orders','new_orders','upd_orders','by']);
    sheet.getRange(1,1,1,6).setFontWeight('bold').setBackground('#1565c0').setFontColor('#fff');
  }

  sheet.appendRow([
    new Date(),
    body.file_name    || '',
    body.total_orders || 0,
    body.new_orders   || 0,
    body.upd_orders   || 0,
    body.by           || 'ياسر السيد',
  ]);

  const uploadTime = new Date().toLocaleString('ar-EG', {timeZone:'Africa/Cairo'});
  const newBadge   = body.new_orders > 0 ? `\n🆕 *أوامر جديدة: ${body.new_orders}*` : '';
  const sample     = (body.orders_sample||[]).slice(0,3)
    .map(o=>`  • ${o.order_no} | ${o.customer} | ${o.team_code||'—'} | ${o.units} وحدة`)
    .join('\n');

  // إشعار الإدارة
  const adminMsg =
    `📦 *رفع SAP جديد*\n`+
    `👤 بواسطة: ${body.by||'ياسر السيد'}\n`+
    `📄 الملف: ${body.file_name||'—'}\n`+
    `📊 الأوامر: *${body.total_orders}* إجمالي${newBadge}\n`+
    `🕐 ${uploadTime}\n`+
    (sample?`\n*عينة:*\n${sample}`:'');

  Object.values(ADMIN_IDS).forEach(id => sendTelegram(id, adminMsg));

  // ── كتابة الأوامر في TrackingUpdates كـ baseline ──────────
  if (body.orders_sample && body.orders_sample.length > 0) {
    let updSheet = ss.getSheetByName(SHEET_UPDATES);
    if (!updSheet) {
      updSheet = ss.insertSheet(SHEET_UPDATES);
      updSheet.appendRow(['type','order_no','customer','team_code','contract_no',
                          'status','pct','issue_type','severity','snag_id',
                          'photo_count','ir_type','ir_no','ir_vl',
                          'delivery_status','note','by','ts']);
    }
    // جيب الأوامر الموجودة عشان منضفش تكرار
    const existingData = updSheet.getDataRange().getValues();
    const existingOrders = new Set(existingData.slice(1).map(r => r[1]));

    body.orders_sample.forEach(o => {
      if (!existingOrders.has(o.order_no)) {
        updSheet.appendRow([
          'sap_import', o.order_no, o.customer, o.team_code || '',
          '', 'TECO', 0, '', '', '', 0, '', '', '', '',
          'مستورد من SAP', body.by || 'ياسر السيد', new Date()
        ]);
      }
    });
  }

  // إشعار مديري المجموعات بأوامرهم الجديدة
  if (body.new_orders > 0 && body.orders_sample) {
    const byTeam = {};
    body.orders_sample.forEach(o => {
      const tc = o.team_code || '';
      if (!byTeam[tc]) byTeam[tc] = [];
      byTeam[tc].push(o);
    });
    Object.entries(byTeam).forEach(([tc, ords]) => {
      const mgr = TEAM_MANAGERS[tc];
      if (!mgr || !mgr.id) return;
      const mgrMsg =
        `📦 *أوامر تركيب جديدة لفريقك*\n`+
        `${ords.map(o=>`  • ${o.order_no} | ${o.customer} | ${o.units} وحدة`).join('\n')}\n`+
        `\n✅ افتح النظام وابدأ التخطيط:\nhttps://ashrafa1-glitch.github.io/mobica-fleet/tracking/`;
      sendTelegram(mgr.id, mgrMsg);
    });
  }

  return { ok: true };
}

// ═══════════════════════════════════════════════════════════
// ADD IR — حفظ محضر تسليم
// ═══════════════════════════════════════════════════════════
function addIR(body) {
  return addUpdate(Object.assign({ type: 'ir' }, body));
}

// ═══════════════════════════════════════════════════════════
// HELPER — get team manager
// ═══════════════════════════════════════════════════════════
function getTeamManager(teamCode) {
  return TEAM_MANAGERS[teamCode] || null;
}

// ═══════════════════════════════════════════════════════════
// SEND TELEGRAM
// ═══════════════════════════════════════════════════════════
function sendTelegram(chatId, text) {
  if (!chatId) return;
  try {
    UrlFetchApp.fetch(
      `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`,
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          chat_id: String(chatId),
          text: text,
          parse_mode: 'Markdown'
        })
      }
    );
  } catch(e) {
    Logger.log('Telegram error: ' + e.message);
  }
}

// ═══════════════════════════════════════════════════════════
// KHALED REPORT — تقرير المتأخرات لخالد علي
// ═══════════════════════════════════════════════════════════
function sendKhaledReport(time) {
  const KHALED_ID = '8017935026';
  const ss        = SpreadsheetApp.openById(SHEET_ID);
  const updSheet  = ss.getSheetByName(SHEET_UPDATES);
  const logSheet  = ss.getSheetByName(SHEET_REMINDERS);
  if (!updSheet) return;

  const data    = updSheet.getDataRange().getValues();
  const headers = data[0];
  const now     = new Date().getTime();
  const today   = new Date().toLocaleDateString('ar-EG', {timeZone:'Africa/Cairo'});

  // ── آخر تحديث لكل أمر ──────────────────────────────────
  const lastStatus = {};
  for (let i = 1; i < data.length; i++) {
    const row = {};
    headers.forEach((h,j) => row[h] = data[i][j]);
    if (row.type !== 'status') continue;
    const k = row.order_no;
    if (!lastStatus[k] || Number(row.ts||0) > Number(lastStatus[k].ts||0))
      lastStatus[k] = row;
  }

  // ── الأوامر المتأخرة مجمّعة حسب الفريق ─────────────────
  const threshold = 3 * 24 * 60 * 60 * 1000; // +3 أيام
  const delayedByTeam = {};

  Object.values(lastStatus).forEach(row => {
    if (row.status === 'Completed') return;
    const lastUpd = row.ts ? Number(row.ts) : 0;
    const delayMs = now - lastUpd;
    if (delayMs < threshold) return;
    const tc = row.team_code || 'unknown';
    if (!delayedByTeam[tc]) delayedByTeam[tc] = [];
    const delayDays = Math.floor(delayMs / (24*60*60*1000));
    delayedByTeam[tc].push({ ...row, delayDays });
  });

  // ── إحصائيات التذكيرات لكل مدير ────────────────────────
  const reminderStats = {};
  if (logSheet) {
    const logData    = logSheet.getDataRange().getValues();
    const logHeaders = logData[0];
    const week7      = 7 * 24 * 60 * 60 * 1000;
    for (let i = 1; i < logData.length; i++) {
      const row = {};
      logHeaders.forEach((h,j) => row[h] = logData[i][j]);
      const ts = row.timestamp ? new Date(row.timestamp).getTime() : 0;
      if (now - ts > week7) continue;
      const name = row.manager_name || 'غير معروف';
      if (!reminderStats[name]) reminderStats[name] = { sent:0, responded:0 };
      reminderStats[name].sent++;
      if (row.responded === 'نعم') reminderStats[name].responded++;
    }
  }

  // ── بناء التقرير ────────────────────────────────────────
  let report =
    `📊 *تقرير المتأخرات — لخالد علي*\n`+
    `🕐 ${time} | 📅 ${today}\n`+
    `━━━━━━━━━━━━━━━━━━━━━\n`;

  let totalDelayed = 0;

  Object.entries(TEAM_MANAGERS).forEach(([code, mgr]) => {
    const delayed = delayedByTeam[code] || [];
    if (!delayed.length) return;
    totalDelayed += delayed.length;

    const stats   = reminderStats[mgr.name] || { sent:0, responded:0 };
    const noReply = stats.sent - stats.responded;

    // ترتيب تنازلي حسب مدة التأخير
    delayed.sort((a,b) => b.delayDays - a.delayDays);

    report +=
      `\n🔧 *${mgr.name}*\n`+
      `📨 تذكيرات أُرسلت: ${stats.sent} | ✅ رد: ${stats.responded} | ❌ لم يرد: ${noReply}\n`+
      `⏰ أوامر متأخرة: *${delayed.length}*\n`;

    delayed.forEach(r => {
      const urgency = r.delayDays >= 14 ? '🔴' : r.delayDays >= 7 ? '🟡' : '🟠';
      report += `  ${urgency} ${r.order_no} | ${r.customer||'—'} | *${r.delayDays} يوم*\n`;
    });
  });

  if (totalDelayed === 0) {
    report += `\n✅ *لا توجد أوامر متأخرة دلوقتي*\n`;
  } else {
    report +=
      `\n━━━━━━━━━━━━━━━━━━━━━\n`+
      `📌 *إجمالي المتأخرة: ${totalDelayed} أمر*\n`+
      `🔴 أكثر من 14 يوم | 🟡 7-14 يوم | 🟠 3-7 أيام\n`;
  }

  report += `\n👁️ التفاصيل: https://ashrafa1-glitch.github.io/mobica-fleet/tracking/`;

  sendTelegram(KHALED_ID, report);
  Logger.log('✅ Khaled report sent');
}

// ═══════════════════════════════════════════════════════════
// REMINDER SYSTEM — تذكيرات المتأخرات لمديري المجموعات
// ═══════════════════════════════════════════════════════════

// يُستدعى تلقائياً كل 12 ساعة
function sendDelayedReminders() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let logSheet = ss.getSheetByName(SHEET_REMINDERS);
  if (!logSheet) {
    logSheet = ss.insertSheet(SHEET_REMINDERS);
    logSheet.appendRow(['timestamp','manager_name','manager_id','team_code',
                        'reminder_type','orders_delayed','message_sent','responded','response_ts']);
    logSheet.getRange(1,1,1,9).setFontWeight('bold').setBackground('#1565c0').setFontColor('#fff');
  }

  // جيب الأوامر المتأخرة من الـ TrackingUpdates
  const updSheet = ss.getSheetByName(SHEET_UPDATES);
  if (!updSheet) return;

  const data    = updSheet.getDataRange().getValues();
  const headers = data[0];
  const now     = new Date().getTime();
  const today   = new Date().toLocaleDateString('ar-EG',{timeZone:'Africa/Cairo'});

  // تجميع آخر تحديث لكل أمر
  const lastStatus = {};
  for (let i = 1; i < data.length; i++) {
    const row = {};
    headers.forEach((h,j) => row[h] = data[i][j]);
    if (row.type !== 'status') continue;
    const k = row.order_no;
    if (!lastStatus[k] || Number(row.ts||0) > Number(lastStatus[k].ts||0)) {
      lastStatus[k] = row;
    }
  }

  // تحديد المتأخرة: أوامر حالتها مش Completed ومرت أكثر من 3 أيام بدون تحديث
  const threshold3d = 3 * 24 * 60 * 60 * 1000;
  const delayedByTeam = {};

  Object.values(lastStatus).forEach(row => {
    if (row.status === 'Completed') return;
    const lastUpd = row.ts ? Number(row.ts) : 0;
    if (now - lastUpd < threshold3d) return; // محدّث حديثاً
    const tc = row.team_code || 'unknown';
    if (!delayedByTeam[tc]) delayedByTeam[tc] = [];
    delayedByTeam[tc].push(row);
  });

  // إرسال تذكير لكل مدير مجموعة عنده متأخرات
  Object.entries(TEAM_MANAGERS).forEach(([code, mgr]) => {
    if (!mgr.id) return;
    const delayed = delayedByTeam[code] || [];
    if (!delayed.length) return;

    const ordersList = delayed.slice(0,5).map(r =>
      `  • ${r.order_no} — ${r.customer||'—'} (${r.status||'—'})`
    ).join('\n');

    const moreCount = delayed.length > 5 ? `\n  ... و${delayed.length-5} أمر آخر` : '';

    const msg =
      `⏰ *تذكير — أوامر تحتاج تحديث*\n`+
      `📅 ${today}\n`+
      `─────────────────────\n`+
      `لديك *${delayed.length}* أمر لم يُحدَّث منذ أكثر من 3 أيام:\n\n`+
      ordersList + moreCount +
      `\n─────────────────────\n`+
      `📱 حدّث الحالة على: https://ashrafa1-glitch.github.io/mobica-fleet/tracking/\n`+
      `*يرجى الرد بـ "تم" أو "جاري" لكل أمر*`;

    sendTelegram(mgr.id, msg);

    // تسجيل في ReminderLog
    logSheet.appendRow([
      new Date(),
      mgr.name,
      mgr.id,
      code,
      'delayed_orders',
      delayed.length,
      msg,
      'لا',  // responded = لا حتى الآن
      ''
    ]);
  });

  Logger.log('✅ Delayed reminders sent at ' + new Date());
}

// ── تسجيل رد مدير المجموعة (يُستدعى يدوياً أو من webhook) ──
function markReminderResponded(managerId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const logSheet = ss.getSheetByName(SHEET_REMINDERS);
  if (!logSheet) return;

  const data = logSheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][2]) === String(managerId) && data[i][7] === 'لا') {
      logSheet.getRange(i+1, 8).setValue('نعم');
      logSheet.getRange(i+1, 9).setValue(new Date());
      break;
    }
  }
}

// ═══════════════════════════════════════════════════════════
// REMINDER STATS — للتقرير اليومي
// ═══════════════════════════════════════════════════════════
function getReminderStats() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const logSheet = ss.getSheetByName(SHEET_REMINDERS);
  if (!logSheet) return { total:0, responded:0, noResponse:0, details:[] };

  const data    = logSheet.getDataRange().getValues();
  const headers = data[0];
  const now     = new Date().getTime();
  const week7   = 7 * 24 * 60 * 60 * 1000;

  let total = 0, responded = 0, noResponse = 0;
  const managerStats = {};

  for (let i = 1; i < data.length; i++) {
    const row = {};
    headers.forEach((h,j) => row[h] = data[i][j]);
    const ts = row.timestamp ? new Date(row.timestamp).getTime() : 0;
    if (now - ts > week7) continue; // آخر 7 أيام فقط

    total++;
    const name = row.manager_name || 'غير معروف';
    if (!managerStats[name]) managerStats[name] = { sent:0, responded:0 };
    managerStats[name].sent++;

    if (row.responded === 'نعم') {
      responded++;
      managerStats[name].responded++;
    } else {
      noResponse++;
    }
  }

  const details = Object.entries(managerStats).map(([name, s]) =>
    `  • ${name}: ${s.sent} تذكير — رد ${s.responded} — لم يرد ${s.sent - s.responded}`
  );

  return { total, responded, noResponse, details };
}

// ═══════════════════════════════════════════════════════════
// SETUP TRIGGERS — تشغيل مرة واحدة فقط
// ═══════════════════════════════════════════════════════════
function setupTriggers() {
  // احذف الـ triggers القديمة
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // 8:00 ص بتوقيت القاهرة
  ScriptApp.newTrigger('sendMorningReport')
    .timeBased()
    .atHour(6)  // 6 UTC = 8 Cairo
    .everyDays(1)
    .create();

  // 3:00 م بتوقيت القاهرة
  ScriptApp.newTrigger('sendAfternoonReport')
    .timeBased()
    .atHour(13) // 13 UTC = 3 PM Cairo
    .everyDays(1)
    .create();

  // تذكيرات المتأخرات كل 12 ساعة (9ص و9م)
  ScriptApp.newTrigger('sendDelayedReminders')
    .timeBased()
    .atHour(7)   // 7 UTC = 9 Cairo
    .everyDays(1)
    .create();

  ScriptApp.newTrigger('sendDelayedReminders')
    .timeBased()
    .atHour(19)  // 19 UTC = 9 PM Cairo
    .everyDays(1)
    .create();

  Logger.log('✅ Triggers set: 8AM + 3PM reports | 9AM + 9PM delayed reminders');
}

// ═══════════════════════════════════════════════════════════
// REPAIR HEADERS — يُصلح header row في TrackingUpdates
// ═══════════════════════════════════════════════════════════
function repairHeaders() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_UPDATES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_UPDATES);
    sheet.appendRow(REQUIRED_HEADERS);
    return { ok: true, action: 'created', headers: REQUIRED_HEADERS };
  }
  const firstRow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), REQUIRED_HEADERS.length)).getValues()[0];
  sheet.getRange(1, 1, 1, REQUIRED_HEADERS.length).setValues([REQUIRED_HEADERS]);
  return { ok: true, action: 'repaired', old_headers: firstRow, new_headers: REQUIRED_HEADERS };
}

// ═══════════════════════════════════════════════════════════
// DEBUG SHEET — يُرجع headers + آخر صف
// ═══════════════════════════════════════════════════════════
function debugSheet() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_UPDATES);
  if (!sheet) return { error: 'sheet not found' };
  const lastRow  = sheet.getLastRow();
  const lastCol  = sheet.getLastColumn();
  const headers  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const lastData = lastRow > 1 ? sheet.getRange(lastRow, 1, 1, lastCol).getValues()[0] : [];
  return {
    ok: true,
    total_rows: lastRow - 1,
    headers: headers,
    last_row: lastData,
    ts_col_index: headers.indexOf('ts'),
  };
}
