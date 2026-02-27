// ═══════════════════════════════════════════════════════════
// MOBICA FLEET AGENT — Google Apps Script
// ═══════════════════════════════════════════════════════════
// 1. افتح Google Sheets جديد
// 2. Extensions → Apps Script
// 3. انسخ الكود ده
// 4. Deploy → New Deployment → Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 5. انسخ الـ URL وحطّه في fleet-agent.html في متغير API_URL
// ═══════════════════════════════════════════════════════════

const TELEGRAM_BOT_TOKEN = '8209404725:AAH1zZ_QzS3nDW01VMd0TLLHyvC1EYFPdh0';
const EMAD_CHAT_ID       = '8416528954';
const ASHRAF_CHAT_ID     = '7055250567';  // مدير التخطيط
const SHEET_REQUESTS     = 'Requests';
const SHEET_CONFIG       = 'Config';
const SHEET_PLAN         = 'DailyPlan';

// ── خريطة المجموعات → Chat IDs (اضبطها في Config Sheet أو هنا) ──
// مدراء المجموعات — أضف الـ Telegram IDs الحقيقية
const TEAM_CHAT_IDS = {
  'A — أحمد حسن': '',   // أضف Chat ID هنا
  'B — سامى فؤاد': '',
  'C — حازم قاعود': '',
  'D — هشام جمال': '',
};

// ── Bootstrap: أنشئ الـ Sheets لو مش موجودة ──────────────
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let rs = ss.getSheetByName(SHEET_REQUESTS);
  if (!rs) {
    rs = ss.insertSheet(SHEET_REQUESTS);
    rs.appendRow(['id','date','team','from_loc','dest','techs','members','status',
                  'plate','driver','car_type','created_by','created_at','notes']);
    rs.setFrozenRows(1);
    rs.getRange('1:1').setBackground('#1565c0').setFontColor('#fff').setFontWeight('bold');
  }
  let cs = ss.getSheetByName(SHEET_CONFIG);
  if (!cs) {
    cs = ss.insertSheet(SHEET_CONFIG);
    cs.appendRow(['key','value']);
    cs.appendRow(['version','1.1']);
    // أضف Chat IDs للفرق هنا
    cs.appendRow(['team_A_chatId','']);
    cs.appendRow(['team_B_chatId','']);
    cs.appendRow(['team_C_chatId','']);
    cs.appendRow(['team_D_chatId','']);
  }
  // تأكد من وجود DailyPlan sheet
  if (!ss.getSheetByName(SHEET_PLAN)) {
    const ps = ss.insertSheet(SHEET_PLAN);
    ps.appendRow(['date','plate','driver','car_type','primary_from','all_froms','dests','techs','ids','approved_by','created_at']);
    ps.setFrozenRows(1);
    ps.getRange('1:1').setBackground('#0d47a1').setFontColor('#fff').setFontWeight('bold');
  }
}

// ── CORS helper ──────────────────────────────────────────
function corsResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET handler ──────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || '';
  const date   = e.parameter.date   || today();

  if (action === 'getRequests') return corsResponse(getRequests(date));
  if (action === 'ping')        return corsResponse({ ok: true, time: new Date().toISOString() });
  return corsResponse({ error: 'unknown action' });
}

// ── POST handler ─────────────────────────────────────────
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch(err) { return corsResponse({ ok: false, error: 'invalid JSON' }); }

  const action = body.action || '';

  if (action === 'addRequest')   return corsResponse(addRequest(body));
  if (action === 'assignVehicle')return corsResponse(assignVehicle(body));
  if (action === 'savePlan')     return corsResponse(savePlan(body));
  if (action === 'deleteRequest')return corsResponse(deleteRequest(body));
  return corsResponse({ error: 'unknown action' });
}

// ── Helpers ──────────────────────────────────────────────
function today() {
  return Utilities.formatDate(new Date(), 'Africa/Cairo', 'yyyy-MM-dd');
}

function uid() {
  return 'R' + Date.now().toString(36).toUpperCase();
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_REQUESTS) || (setupSheets(), ss.getSheetByName(SHEET_REQUESTS));
}

// ── getRequests ───────────────────────────────────────────
function getRequests(date) {
  const sh   = getSheet();
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1)
    .filter(r => r[1] === date || !date)
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = r[i]);
      return obj;
    });
  return { ok: true, requests: rows, date };
}

// ── addRequest ────────────────────────────────────────────
function addRequest(body) {
  const sh = getSheet();
  const id = uid();
  const now = new Date().toISOString();
  const membersStr = Array.isArray(body.members) ? body.members.join('،') : (body.members||'');
  sh.appendRow([
    id,
    body.date || today(),
    body.team        || '',
    body.from_loc    || '',
    body.dest        || '',
    body.techs       || 0,
    membersStr,              // أسماء الأفراد
    'pending',
    '', '', '',              // plate, driver, car_type
    body.created_by  || '',
    now,
    body.notes       || ''
  ]);

  // إشعار عماد على تيليجرام
  const teamLabel = body.team || '—';
  const membersLine = membersStr ? `\n👤 الأفراد: ${membersStr}` : '';
  const msg = `🚐 *طلب نقل جديد*\n`
    + `👷 الفريق: ${teamLabel}\n`
    + `📍 من: ${body.from_loc || '—'} ➜ ${body.dest || '—'}\n`
    + `👥 عدد الأفراد: ${body.techs || 0}${membersLine}\n`
    + `🗓 التاريخ: ${body.date || today()}\n`
    + `👤 أضافه: ${body.created_by || '—'}`;
  sendTelegram(EMAD_CHAT_ID, msg);

  return { ok: true, id };
}

// ── assignVehicle ─────────────────────────────────────────
function assignVehicle(body) {
  const sh   = getSheet();
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idCol    = headers.indexOf('id');
  const plateCol = headers.indexOf('plate');
  const driverCol= headers.indexOf('driver');
  const typeCol  = headers.indexOf('car_type');
  const statusCol= headers.indexOf('status');

  // body.ids = array of request IDs to assign
  const ids = body.ids || [];
  let updated = 0;
  data.forEach((row, i) => {
    if (i === 0) return;
    if (ids.includes(row[idCol])) {
      sh.getRange(i+1, plateCol+1) .setValue(body.plate    || '');
      sh.getRange(i+1, driverCol+1).setValue(body.driver   || '');
      sh.getRange(i+1, typeCol+1)  .setValue(body.car_type || '');
      sh.getRange(i+1, statusCol+1).setValue('assigned');
      updated++;
    }
  });

  // إشعار تأكيد
  if (updated > 0) {
    const msg = `✅ *تم تحديد السيارة*\n`
      + `🔑 رقم السيارة: ${body.plate || '—'}\n`
      + `👤 السائق: ${body.driver || '—'}\n`
      + `🚐 النوع: ${body.car_type_label || body.car_type || '—'}\n`
      + `📝 عدد الطلبات: ${updated}`;
    sendTelegram(EMAD_CHAT_ID, msg);
  }

  return { ok: true, updated };
}

// ── savePlan ──────────────────────────────────────────────
// يحفظ الخطة اليومية ويبعت إشعارات Telegram للفنيين
function savePlan(body) {
  const date        = body.date || today();
  const plan        = body.plan || [];
  const approvedBy  = body.approved_by || 'مسؤول الحملة';

  // حفظ في Sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ps = ss.getSheetByName(SHEET_PLAN);
  if (!ps) {
    ps = ss.insertSheet(SHEET_PLAN);
    ps.appendRow(['date','plate','driver','car_type','primary_from','all_froms','dests','techs','ids','approved_by','created_at']);
    ps.setFrozenRows(1);
    ps.getRange('1:1').setBackground('#0d47a1').setFontColor('#fff').setFontWeight('bold');
  }

  // حذف سجلات نفس اليوم (refresh)
  const data = ps.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === date) ps.deleteRow(i + 1);
  }

  // إضافة الخطة الجديدة
  plan.forEach(p => {
    ps.appendRow([
      date,
      p.plate || '',
      p.driver || '',
      p.car_type || '',
      p.primaryFrom || '',
      (p.allFroms || []).join(' | '),
      (p.dests || []).join(' | '),
      p.techs || 0,
      (p.ids || []).join(','),
      approvedBy,
      new Date().toISOString()
    ]);
  });

  // ── إشعارات Telegram ──────────────────────────────────
  const mapsBase = 'https://www.google.com/maps/search/?api=1&query=';

  plan.forEach(p => {
    const icon  = p.car_type === 'suzuki' ? '🚙' : p.car_type === 'micro14' ? '🚐' : p.car_type === 'bus26' ? '🚌' : '🛻';
    const dests = (p.dests || []).join(' + ');

    // إشعار لكل مجموعة في هذه الرحلة
    const teamsInvolved = [...new Set((p.rows || []).map(r => r.team).filter(Boolean))];

    teamsInvolved.forEach(team => {
      const teamReqs = (p.rows || []).filter(r => r.team === team);
      const teamFroms = [...new Set(teamReqs.map(r => r.from_loc).filter(Boolean))];
      const teamTechs = teamReqs.reduce((s, r) => s + parseInt(r.techs || 0), 0);

      // هل عندهم نقطة التقاء (مش نقطة التحرك الأساسية)؟
      const isMeeting = teamFroms.some(f => f !== p.primaryFrom);
      const meetingPoint = isMeeting ? teamFroms[0] : null;

      let msg = `${icon} *تأكيد الحملة — ${team}*\n`;
      msg += `🗓 ${date}\n`;
      msg += `👤 السائق: ${p.driver} | السيارة: ${p.plate}\n`;

      if (isMeeting && meetingPoint) {
        msg += `\n📌 *نقطة الالتقاء:* ${meetingPoint}\n`;
        msg += `🗺 [افتح في الخريطة](${mapsBase}${encodeURIComponent(meetingPoint + ' القاهرة مصر')})\n`;
        msg += `⏰ كون في نقطة الالتقاء قبل موعد التحرك\n`;
      } else {
        msg += `\n🟢 *نقطة التحرك:* ${p.primaryFrom}\n`;
        msg += `🗺 [افتح في الخريطة](${mapsBase}${encodeURIComponent(p.primaryFrom + ' القاهرة مصر')})\n`;
      }

      msg += `\n🔴 *الوجهة:* ${dests}\n`;
      msg += `👥 معك ${teamTechs} فرد\n`;
      msg += `\n✅ اعتمده: ${approvedBy}`;

      // إرسال لمدير المجموعة لو عنده Chat ID
      const chatId = TEAM_CHAT_IDS[team];
      if (chatId) sendTelegram(chatId, msg);
    });
  });

  // ملخص لعماد وأشرف
  const summary = plan.map((p,i) =>
    `${i+1}. ${p.plate} | ${p.primaryFrom}${p.mergedFroms?' + نقطة التقاء':''} ➜ ${(p.dests||[]).join('+')} | ${p.techs} فرد`
  ).join('\n');

  const summaryMsg = `✅ *خطة الحملة معتمدة*\n🗓 ${date}\n👤 اعتمدها: ${approvedBy}\n\n${summary}`;
  sendTelegram(EMAD_CHAT_ID, summaryMsg);
  sendTelegram(ASHRAF_CHAT_ID, summaryMsg);

  return { ok: true, notified: plan.length };
}

// ── deleteRequest ─────────────────────────────────────────
function deleteRequest(body) {
  const sh   = getSheet();
  const data = sh.getDataRange().getValues();
  const idCol = data[0].indexOf('id');
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][idCol] === body.id) {
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'not found' };
}

// ── Telegram ──────────────────────────────────────────────
function sendTelegram(chatId, text) {
  try {
    const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        chat_id: chatId,
        text,
        parse_mode: 'Markdown'
      })
    });
  } catch(e) {
    Logger.log('Telegram error: ' + e.message);
  }
}
