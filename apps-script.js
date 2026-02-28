// ═══════════════════════════════════════════════════════════
// MOBICA FLEET AGENT — Google Apps Script v2.0
// ═══════════════════════════════════════════════════════════
const TELEGRAM_BOT_TOKEN = '8209404725:AAH1zZ_QzS3nDW01VMd0TLLHyvC1EYFPdh0';
const EMAD_CHAT_ID       = '8296695318';
const ASHRAF_CHAT_ID     = '7055250567';
const SHEET_REQUESTS     = 'Requests';
const SHEET_CONFIG       = 'Config';
const SHEET_PLAN         = 'DailyPlan';
const SHEET_RATINGS      = 'Ratings';
const SHEET_DRIVER_LOG   = 'DriverLog';

// ── Chat IDs مديري المجموعات ──
const TEAM_CHAT_IDS = {
  'A — أحمد حسن': '8583850058',
  'B — سامى فؤاد': '',
  'C — حازم قاعود': '979158814',
  'D — هشام جمال': '5555531128',
};

// ── Chat IDs الأفراد ──
const MEMBER_CHAT_IDS = {
  'م. عبد العزيز عبده': '5421117602',
};

// ── Chat IDs السائقين ──
const DRIVER_CHAT_IDS = {
  // 'اسم السائق': 'chat_id',
};

// ══════════════════════════════════════════════════════════
// HELPERS
// ══════════════════════════════════════════════════════════
function today() {
  return 'D:' + Utilities.formatDate(new Date(), 'Africa/Cairo', 'yyyy-MM-dd');
}
function uid() { return 'R' + Date.now().toString(36).toUpperCase(); }

function getSheet() {
  const ss = SpreadsheetApp.openById("1qZSLK7zEsOxRactIxXVwvVR3tFYruR_CI8bU5rERGwI");
  return ss.getSheetByName(SHEET_REQUESTS) || (setupSheets(), ss.getSheetByName(SHEET_REQUESTS));
}

function normDate(v) {
  if (!v) return '';
  if (v instanceof Date) return 'D:' + Utilities.formatDate(v, 'Africa/Cairo', 'yyyy-MM-dd');
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return 'D:' + s;
  return s;
}

// ══════════════════════════════════════════════════════════
// SETUP
// ══════════════════════════════════════════════════════════
function setupSheets() {
  const ss = SpreadsheetApp.openById("1qZSLK7zEsOxRactIxXVwvVR3tFYruR_CI8bU5rERGwI");

  if (!ss.getSheetByName(SHEET_REQUESTS)) {
    const rs = ss.insertSheet(SHEET_REQUESTS);
    rs.appendRow(['id','date','team','client','from_loc','dest','techs','members','status',
                  'plate','driver','car_type','created_by','created_at','notes','stop_loc','rating','rating_note']);
    rs.setFrozenRows(1);
    rs.getRange('1:1').setBackground('#1565c0').setFontColor('#fff').setFontWeight('bold');
  }

  if (!ss.getSheetByName(SHEET_PLAN)) {
    const ps = ss.insertSheet(SHEET_PLAN);
    ps.appendRow(['date','plate','driver','car_type','primary_from','all_froms','dests','techs','ids','approved_by','created_at']);
    ps.setFrozenRows(1);
    ps.getRange('1:1').setBackground('#0d47a1').setFontColor('#fff').setFontWeight('bold');
  }

  if (!ss.getSheetByName(SHEET_RATINGS)) {
    const rat = ss.insertSheet(SHEET_RATINGS);
    rat.appendRow(['date','driver','plate','rating','note','rated_by','created_at']);
    rat.setFrozenRows(1);
    rat.getRange('1:1').setBackground('#4a148c').setFontColor('#fff').setFontWeight('bold');
  }

  if (!ss.getSheetByName(SHEET_DRIVER_LOG)) {
    const dl = ss.insertSheet(SHEET_DRIVER_LOG);
    dl.appendRow(['date','driver','plate','status','location','updated_at']);
    dl.setFrozenRows(1);
    dl.getRange('1:1').setBackground('#1b5e20').setFontColor('#fff').setFontWeight('bold');
  }
}

// ══════════════════════════════════════════════════════════
// CORS HELPER
// ══════════════════════════════════════════════════════════
function corsResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════
// GET HANDLER
// ══════════════════════════════════════════════════════════
function doGet(e) {
  const p      = e.parameter || {};
  const action = p.action || '';

  if (action === 'ping')        return corsResponse({ ok: true, time: new Date().toISOString() });
  if (action === 'getRequests') return corsResponse(getRequests(p.date || today()));
  if (action === 'getHistory')  return corsResponse(getHistory(p.days || 7));
  if (action === 'getRatings')  return corsResponse(getRatings());
  if (action === 'getDriverLog')return corsResponse(getDriverLog(p.date || today()));

  const body = parseParams(p);
  if (action === 'addRequest')    return corsResponse(addRequest(body));
  if (action === 'assignVehicle') return corsResponse(assignVehicle(body));
  if (action === 'savePlan')      return corsResponse(savePlan(body));
  if (action === 'deleteRequest') return corsResponse(deleteRequest(body));
  if (action === 'updateRequest') return corsResponse(updateRequest(body));
  if (action === 'cancelRequest') return corsResponse(cancelRequest(body));
  if (action === 'rateRequest')   return corsResponse(rateRequest(body));
  if (action === 'driverStatus')  return corsResponse(driverStatus(body));

  return corsResponse({ error: 'unknown action: ' + action });
}

function parseParams(p) {
  const body = {};
  for (const k in p) {
    if (k === 'action') continue;
    try { body[k] = JSON.parse(p[k]); } catch(e) { body[k] = p[k]; }
  }
  return body;
}

// ══════════════════════════════════════════════════════════
// GET REQUESTS
// ══════════════════════════════════════════════════════════
function getRequests(date) {
  const sh   = getSheet();
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return { ok: true, requests: [], date };
  const headers = data[0];
  const rows = data.slice(1)
    .filter(r => !date || normDate(r[1]) === normDate(date))
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = r[i] instanceof Date
          ? 'D:' + Utilities.formatDate(r[i], 'Africa/Cairo', 'yyyy-MM-dd') : r[i];
      });
      if (obj.members && typeof obj.members === 'string' && obj.members)
        obj.members = obj.members.split('،').map(s => s.trim()).filter(Boolean);
      if (obj.ids && typeof obj.ids === 'string')
        obj.ids = obj.ids.split(',').filter(Boolean);
      return obj;
    });
  return { ok: true, requests: rows, date };
}

// ══════════════════════════════════════════════════════════
// GET HISTORY (أيام سابقة)
// ══════════════════════════════════════════════════════════
function getHistory(days) {
  const sh   = getSheet();
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return { ok: true, requests: [] };
  const headers = data[0];
  const cutoff  = new Date();
  cutoff.setDate(cutoff.getDate() - parseInt(days));

  const rows = data.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = r[i] instanceof Date
        ? 'D:' + Utilities.formatDate(r[i], 'Africa/Cairo', 'yyyy-MM-dd') : r[i];
    });
    if (obj.members && typeof obj.members === 'string')
      obj.members = obj.members.split('،').map(s => s.trim()).filter(Boolean);
    return obj;
  });
  return { ok: true, requests: rows };
}

// ══════════════════════════════════════════════════════════
// ADD REQUEST
// ══════════════════════════════════════════════════════════
function addRequest(body) {
  const sh  = getSheet();
  const id  = uid();
  const now = new Date().toISOString();
  const membersStr = Array.isArray(body.members) ? body.members.join('،') : (body.members || '');
  sh.appendRow([
    id,
    body.date || today(),
    body.team     || '',
    body.client   || '',
    body.from_loc || '',
    body.dest     || '',
    body.techs    || 0,
    membersStr,
    'pending',
    '', '', '',
    body.created_by || '',
    now,
    body.notes    || '',
    body.stop_loc || '',
    '', ''   // rating, rating_note
  ]);

  // إشعار عماد
  const membersLine = membersStr ? `\n👤 الأفراد: ${membersStr}` : '';
  const stopLine    = body.stop_loc ? `\n🔸 محطة: ${body.stop_loc}` : '';
  sendTelegram(EMAD_CHAT_ID,
    `🚐 *طلب نقل جديد*\n🏢 العميل: ${body.client||'—'}\n👷 الفريق: ${body.team||'—'}\n` +
    `📍 من: ${body.from_loc||'—'}${stopLine} ➜ ${body.dest||'—'}\n` +
    `👥 الأفراد: ${body.techs||0}${membersLine}\n👤 أضافه: ${body.created_by||'—'}`
  );

  return { ok: true, id };
}

// ══════════════════════════════════════════════════════════
// UPDATE REQUEST (تعديل قبل الاعتماد)
// ══════════════════════════════════════════════════════════
function updateRequest(body) {
  const sh   = getSheet();
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === body.id && data[i][headers.indexOf('status')] === 'pending') {
      const fields = { client:3, from_loc:4, dest:5, techs:6, members:7, notes:14, stop_loc:15 };
      for (const [k, col] of Object.entries(fields)) {
        if (body[k] !== undefined) {
          const val = k === 'members' && Array.isArray(body[k]) ? body[k].join('،') : body[k];
          sh.getRange(i+1, col+1).setValue(val);
        }
      }
      return { ok: true };
    }
  }
  return { ok: false, error: 'not found or already approved' };
}

// ══════════════════════════════════════════════════════════
// CANCEL REQUEST
// ══════════════════════════════════════════════════════════
function cancelRequest(body) {
  const sh   = getSheet();
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idCol     = headers.indexOf('id');
  const statusCol = headers.indexOf('status');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === body.id) {
      sh.getRange(i+1, statusCol+1).setValue('cancelled');
      sendTelegram(EMAD_CHAT_ID, `❌ *تم إلغاء طلب*\nID: ${body.id}\nبواسطة: ${body.cancelled_by||'—'}`);
      return { ok: true };
    }
  }
  return { ok: false, error: 'not found' };
}

// ══════════════════════════════════════════════════════════
// DELETE REQUEST
// ══════════════════════════════════════════════════════════
function deleteRequest(body) {
  const sh   = getSheet();
  const data = sh.getDataRange().getValues();
  const idCol = data[0].indexOf('id');
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][idCol] === body.id) { sh.deleteRow(i+1); return { ok: true }; }
  }
  return { ok: false, error: 'not found' };
}

// ══════════════════════════════════════════════════════════
// ASSIGN VEHICLE
// ══════════════════════════════════════════════════════════
function assignVehicle(body) {
  const sh   = getSheet();
  const data = sh.getDataRange().getValues();
  const headers  = data[0];
  const idCol    = headers.indexOf('id');
  const plateCol = headers.indexOf('plate');
  const driverCol= headers.indexOf('driver');
  const typeCol  = headers.indexOf('car_type');
  const statusCol= headers.indexOf('status');

  const ids = Array.isArray(body.ids) ? body.ids : (body.ids||'').split(',');
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
  return { ok: true, updated };
}

// ══════════════════════════════════════════════════════════
// SAVE PLAN (اعتماد + إشعارات)
// ══════════════════════════════════════════════════════════
function savePlan(body) {
  const date       = body.date || today();
  const plan       = body.plan || [];
  const approvedBy = body.approved_by || 'مسؤول الحملة';

  // تحديث status الطلبات لـ approved
  const sh   = getSheet();
  const data = sh.getDataRange().getValues();
  const headers   = data[0];
  const idCol     = headers.indexOf('id');
  const statusCol = headers.indexOf('status');
  const plateCol  = headers.indexOf('plate');
  const driverCol = headers.indexOf('driver');
  const typeCol   = headers.indexOf('car_type');

  const allIds = plan.flatMap(p => Array.isArray(p.ids) ? p.ids : (p.ids||'').split(','));
  data.forEach((row, i) => {
    if (i === 0) return;
    const p = plan.find(pl => (Array.isArray(pl.ids)?pl.ids:(pl.ids||'').split(',')).includes(row[idCol]));
    if (p) {
      sh.getRange(i+1, statusCol+1).setValue('approved');
      sh.getRange(i+1, plateCol+1) .setValue(p.plate  || '');
      sh.getRange(i+1, driverCol+1).setValue(p.driver || '');
      sh.getRange(i+1, typeCol+1)  .setValue(p.car_type || '');
    }
  });

  // حفظ في DailyPlan
  const ss = SpreadsheetApp.openById("1qZSLK7zEsOxRactIxXVwvVR3tFYruR_CI8bU5rERGwI");
  let ps = ss.getSheetByName(SHEET_PLAN);
  if (!ps) { setupSheets(); ps = ss.getSheetByName(SHEET_PLAN); }
  const planData = ps.getDataRange().getValues();
  for (let i = planData.length - 1; i >= 1; i--) {
    if (normDate(planData[i][0]) === normDate(date)) ps.deleteRow(i+1);
  }
  plan.forEach(p => {
    ps.appendRow([date, p.plate||'', p.driver||'', p.car_type||'',
      p.primaryFrom||'', (p.allFroms||[]).join(' | '), (p.dests||[]).join(' | '),
      p.techs||0, (Array.isArray(p.ids)?p.ids:(p.ids||'').split(',')).join(','),
      approvedBy, new Date().toISOString()]);
  });

  // ── إشعارات ──
  const mapsBase = 'https://www.google.com/maps/search/?api=1&query=';

  plan.forEach(p => {
    const icon  = p.car_type==='bus45'?'🚌':p.car_type==='micro14'?'🚐':p.car_type==='van7'?'🚐':'🛻';
    const dests = (p.dests||[]).join(' + ');
    const ids   = Array.isArray(p.ids) ? p.ids : (p.ids||'').split(',');

    // إشعار مديري المجموعات + الأفراد
    const teams = [...new Set((p.rows||[]).map(r=>r.team).filter(Boolean))];
    teams.forEach(team => {
      const teamReqs  = (p.rows||[]).filter(r=>r.team===team);
      const teamFroms = [...new Set(teamReqs.map(r=>r.from_loc).filter(Boolean))];
      const teamTechs = teamReqs.reduce((s,r)=>s+parseInt(r.techs||0),0);
      const isMeeting = teamFroms.some(f=>f!==p.primaryFrom);
      const meetPt    = isMeeting ? teamFroms[0] : null;
      const allMembers= teamReqs.flatMap(r=>Array.isArray(r.members)?r.members:[]);

      let msg = `${icon} *✅ تأكيد الحملة — ${team}*\n🗓 ${date.replace('D:','')}\n`;
      msg += `🚗 السيارة: ${p.plate||'—'} | السائق: ${p.driver||'—'}\n`;
      if (isMeeting && meetPt) {
        msg += `\n📌 *نقطة الالتقاء:* ${meetPt}\n`;
        msg += `[📍 خريطة](${mapsBase}${encodeURIComponent(meetPt+' مصر')})\n`;
      } else {
        msg += `\n🟢 *نقطة التحرك:* ${p.primaryFrom||'—'}\n`;
        msg += `[📍 خريطة](${mapsBase}${encodeURIComponent((p.primaryFrom||'')+' مصر')})\n`;
      }
      msg += `🔴 *الوجهة:* ${dests}\n👥 ${teamTechs} فرد\n`;
      if (allMembers.length) msg += `👤 ${allMembers.join(' • ')}\n`;
      msg += `\n✅ اعتمده: ${approvedBy}`;

      // البحث عن chatId سواء كان الـkey هو 'A' أو 'A — اسم'
      const chatId = TEAM_CHAT_IDS[team] ||
        Object.entries(TEAM_CHAT_IDS).find(([k])=>k.startsWith(team+' '))?.[1] || '';
      if (chatId) sendTelegram(chatId, msg);

      // إشعار الأفراد المعروف chat IDs
      allMembers.forEach(m => {
        const mId = MEMBER_CHAT_IDS[m];
        if (mId) sendTelegram(mId, msg);
      });
    });

    // إشعار السائق
    const driverChatId = DRIVER_CHAT_IDS[p.driver];
    if (driverChatId) {
      const driverMsg = `🚐 *رحلة جديدة محددة لك*\n🗓 ${date.replace('D:','')}\n`
        + `🟢 التحرك من: ${p.primaryFrom||'—'}\n🔴 الوجهة: ${dests}\n`
        + `👥 عدد الأفراد: ${p.techs||0}\n✅ اعتمدها: ${approvedBy}`;
      sendTelegram(driverChatId, driverMsg);
    }
  });

  // ملخص لعماد وأشرف
  const summary = plan.map((p,i) =>
    `${i+1}. ${p.plate} | ${p.driver||'—'} | ${p.primaryFrom||'—'} ➜ ${(p.dests||[]).join('+')} | ${p.techs} فرد`
  ).join('\n');
  const sumMsg = `✅ *خطة الحملة معتمدة*\n🗓 ${date.replace('D:','')}\n👤 ${approvedBy}\n\n${summary}`;
  sendTelegram(EMAD_CHAT_ID, sumMsg);
  sendTelegram(ASHRAF_CHAT_ID, sumMsg);

  return { ok: true, notified: plan.length };
}

// ══════════════════════════════════════════════════════════
// RATE REQUEST
// ══════════════════════════════════════════════════════════
function rateRequest(body) {
  const sh   = getSheet();
  const data = sh.getDataRange().getValues();
  const headers    = data[0];
  const idCol      = headers.indexOf('id');
  const ratingCol  = headers.indexOf('rating');
  const noteCol    = headers.indexOf('rating_note');
  const driverCol  = headers.indexOf('driver');
  const plateCol   = headers.indexOf('plate');

  let driver = '', plate = '';
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === body.id) {
      sh.getRange(i+1, ratingCol+1).setValue(body.rating || '');
      sh.getRange(i+1, noteCol+1)  .setValue(body.ratingNote || '');
      driver = data[i][driverCol]; plate = data[i][plateCol];
      break;
    }
  }

  // حفظ في Ratings sheet
  const ss  = SpreadsheetApp.openById("1qZSLK7zEsOxRactIxXVwvVR3tFYruR_CI8bU5rERGwI");
  let rat = ss.getSheetByName(SHEET_RATINGS);
  if (!rat) { setupSheets(); rat = ss.getSheetByName(SHEET_RATINGS); }
  rat.appendRow([today().replace('D:',''), driver, plate, body.rating||'',
    body.ratingNote||'', body.rated_by||'', new Date().toISOString()]);

  return { ok: true };
}

// ══════════════════════════════════════════════════════════
// GET RATINGS (مقارنة التقييمات)
// ══════════════════════════════════════════════════════════
function getRatings() {
  const ss  = SpreadsheetApp.openById("1qZSLK7zEsOxRactIxXVwvVR3tFYruR_CI8bU5rERGwI");
  const rat = ss.getSheetByName(SHEET_RATINGS);
  if (!rat) return { ok: true, ratings: [], summary: [] };
  const data = rat.getDataRange().getValues();
  if (data.length <= 1) return { ok: true, ratings: [], summary: [] };
  const rows = data.slice(1).map(r => ({
    date: r[0], driver: r[1], plate: r[2], rating: r[3], note: r[4], rated_by: r[5]
  }));

  // ملخص لكل سائق
  const byDriver = {};
  rows.forEach(r => {
    if (!r.driver) return;
    if (!byDriver[r.driver]) byDriver[r.driver] = { driver:r.driver, plate:r.plate, total:0, count:0 };
    byDriver[r.driver].total += parseFloat(r.rating)||0;
    byDriver[r.driver].count++;
  });
  const summary = Object.values(byDriver).map(d => ({
    ...d, avg: d.count ? (d.total/d.count).toFixed(1) : 0
  })).sort((a,b) => b.avg - a.avg);

  return { ok: true, ratings: rows, summary };
}

// ══════════════════════════════════════════════════════════
// DRIVER STATUS (في الطريق)
// ══════════════════════════════════════════════════════════
function driverStatus(body) {
  const ss = SpreadsheetApp.openById("1qZSLK7zEsOxRactIxXVwvVR3tFYruR_CI8bU5rERGwI");
  let dl = ss.getSheetByName(SHEET_DRIVER_LOG);
  if (!dl) { setupSheets(); dl = ss.getSheetByName(SHEET_DRIVER_LOG); }
  dl.appendRow([today().replace('D:',''), body.driver||'', body.plate||'',
    body.status||'', body.location||'', new Date().toISOString()]);

  // إشعار عماد
  const statusLabel = body.status==='on_way'?'🚀 في الطريق':body.status==='arrived'?'✅ وصل':body.status;
  sendTelegram(EMAD_CHAT_ID,
    `🚗 *تحديث السائق*\n${statusLabel}\n👤 ${body.driver||'—'} | ${body.plate||'—'}`);

  return { ok: true };
}

function getDriverLog(date) {
  const ss = SpreadsheetApp.openById("1qZSLK7zEsOxRactIxXVwvVR3tFYruR_CI8bU5rERGwI");
  const dl = ss.getSheetByName(SHEET_DRIVER_LOG);
  if (!dl) return { ok: true, log: [] };
  const data = dl.getDataRange().getValues();
  const rows = data.slice(1)
    .filter(r => r[0] === date.replace('D:',''))
    .map(r => ({ date:r[0], driver:r[1], plate:r[2], status:r[3], location:r[4], time:r[5] }));
  return { ok: true, log: rows };
}

// ══════════════════════════════════════════════════════════
// WEEKLY REPORT (يُستدعى من Trigger)
// ══════════════════════════════════════════════════════════
function weeklyReport() {
  const sh   = getSheet();
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return;
  const headers = data[0];
  const rows    = data.slice(1);

  const total     = rows.length;
  const approved  = rows.filter(r=>r[headers.indexOf('status')]==='approved').length;
  const cancelled = rows.filter(r=>r[headers.indexOf('status')]==='cancelled').length;

  // أفضل سائق هذا الأسبوع
  const ratRes = getRatings();
  const topDriver = ratRes.summary.length ? ratRes.summary[0] : null;
  const topLine   = topDriver ? `\n⭐ أفضل سائق: ${topDriver.driver} (${topDriver.avg}/5)` : '';

  // عدد الأفراد الكلي
  const totalPeople = rows.reduce((s,r)=>s+parseInt(r[headers.indexOf('techs')]||0),0);

  const msg = `📊 *التقرير الأسبوعي — حملة التركيبات*\n`
    + `🗓 الأسبوع المنتهي: ${Utilities.formatDate(new Date(),'Africa/Cairo','yyyy-MM-dd')}\n\n`
    + `📋 إجمالي الطلبات: ${total}\n`
    + `✅ المعتمدة: ${approved}\n`
    + `❌ الملغاة: ${cancelled}\n`
    + `👥 إجمالي الأفراد المنقولين: ${totalPeople}`
    + topLine;

  sendTelegram(ASHRAF_CHAT_ID, msg);
  sendTelegram(EMAD_CHAT_ID, msg);
}

// ══════════════════════════════════════════════════════════
// TELEGRAM
// ══════════════════════════════════════════════════════════
function sendTelegram(chatId, text) {
  if (!chatId) return;
  try {
    UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chatId, text, parse_mode: 'Markdown' })
    });
  } catch(e) { Logger.log('Telegram error: ' + e.message); }
}

// ══════════════════════════════════════════════════════════
// WEEKLY TRIGGER SETUP (شغّله مرة واحدة)
// ══════════════════════════════════════════════════════════
function createWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('weeklyReport')
    .timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(17).create();
}
