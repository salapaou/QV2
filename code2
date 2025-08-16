/** ================== CONFIG & SHEETS ================== **/
const SH_QUEUE   = 'Queue';
const SH_TYPES   = 'Qconfig_Types';
const SH_COUNTER = 'Qconfig_Counters';
const SH_LOG     = 'ActionLog';
const SH_DISPLAY = 'Qconfig_Display';

const STATUS_WAITING = 'waiting';
const STATUS_CALLED  = 'called';
const STATUS_CANCEL  = 'cancelled'; // สำหรับปุ่ม Cancel (ใช้ในรายงาน ไม่แสดงบนจอประกาศ)

function onOpen(){
  SpreadsheetApp.getUi().createMenu('QueueV2')
    .addItem('🧰 Setup sheets', 'setupSheets')
    .addSeparator()
    .addItem('▶ เปิด Kiosk (ตัวอย่าง A)', 'openKioskA')
    .addItem('▶ เปิด Desk โต๊ะ 1',       'openDesk1')
    .addItem('▶ เปิด Central (จอกลาง)',   'openCentral')
    .addItem('▶ เปิด Display (จอประกาศ)', 'openDisplay')
    .addSeparator()
    .addItem('🔗 ดูลิงก์ Web App',        'showWebAppLinks')
    .addToUi();
}

/** เปิดเป็น dialog/sidebar ให้เห็นจริง */
function openKioskA(){
  const html = HtmlService.createHtmlOutputFromFile('kiosk')
    .setWidth(420).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Kiosk (ตัวอย่างประเภท A)');
}
function openDesk1(){
  const html = HtmlService.createHtmlOutputFromFile('desk')
    .setWidth(780).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Desk โต๊ะ 1');
}
function openCentral(){
  const html = HtmlService.createHtmlOutputFromFile('central')
    .setTitle('Central');
  SpreadsheetApp.getUi().showSidebar(html); // จอกลางเหมาะกับ sidebar กว้างสูง
}
function openDisplay(){
  const html = HtmlService.createHtmlOutputFromFile('display')
    .setWidth(1024).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Display (ตัวอย่าง)');
}

/** แสดงลิงก์ Web App ที่ deploy แล้ว (ใช้เปิดภายนอกได้) */
function showWebAppLinks(){
  const base = ScriptApp.getService().getUrl() || '(ยังไม่ได้ Deploy เป็น Web App)';
  const html = HtmlService.createHtmlOutput(`
    <div style="font:14px/1.4 system-ui;padding:10px 14px;">
      <div style="margin-bottom:6px;"><b>Web App URL:</b><br>${base}</div>
      <ol style="margin:8px 0 0 18px">
        <li><a target="_blank" href="${base}?page=kiosk&type=A">Kiosk (type=A)</a></li>
        <li><a target="_blank" href="${base}?page=desk&counter=1">Desk โต๊ะ 1</a></li>
        <li><a target="_blank" href="${base}?page=central">Central (iPad)</a></li>
        <li><a target="_blank" href="${base}?page=display">Display</a></li>
      </ol>
      <div style="margin-top:10px;color:#888">* ถ้า base เป็นข้อความในวงเล็บ แปลว่ายังไม่ได้ Deploy → ไปที่ Deploy → New deployment → Web app</div>
    </div>
  `).setWidth(420).setHeight(240);
  SpreadsheetApp.getUi().showModalDialog(html, 'Web App Links');
}

function openKioskA(){ HtmlService.createHtmlOutputFromFile('kiosk'); }
function openDesk1(){ HtmlService.createHtmlOutputFromFile('desk'); }
function openCentral(){ HtmlService.createHtmlOutputFromFile('central'); }
function openDisplay(){ HtmlService.createHtmlOutputFromFile('display'); }

function setupSheets(){
  const ss = SpreadsheetApp.getActive();

  // Queue
  let sh = ss.getSheetByName(SH_QUEUE);
  if(!sh){
    sh = ss.insertSheet(SH_QUEUE);
    sh.appendRow([
      'ts_created','type','raw_number','queue_number','student_name',
      'status','counter_no','ts_called','ts_finished','called_by',
      'priority','pin_counter','client_id','recall_count','note'
    ]);
  }

  // Types
  let t = ss.getSheetByName(SH_TYPES);
  if(!t){
    t = ss.insertSheet(SH_TYPES);
    t.appendRow(['type','prefix','color_hex','display_name','start_number','next_number']);
    t.appendRow(['A','A','#ef4444','สสวท',1,1]);
    t.appendRow(['B','B','#22c55e','BP',1,1]);
    t.appendRow(['C','C','#3b82f6','ทั่วไป',1,1]);
  }

  // Counters
  let c = ss.getSheetByName(SH_COUNTER);
  if(!c){
    c = ss.insertSheet(SH_COUNTER);
    c.appendRow(['counter_no','type','is_open']);
    for (let i=1;i<=10;i++) c.appendRow([i, i<=4?'A':(i<=7?'B':'C'), true]);
  }

  // Display config
  let d = ss.getSheetByName(SH_DISPLAY);
  if(!d){
    d = ss.insertSheet(SH_DISPLAY);
    d.appendRow(['video_url','blink_ms','max_slots']);
    d.appendRow(['https://storage.googleapis.com/gtv-videos-bucket/sample/BigBuckBunny.mp4', 7000, 5]);
  }

  // ActionLog
  let l = ss.getSheetByName(SH_LOG);
  if(!l){
    l = ss.insertSheet(SH_LOG);
    l.appendRow(['ts','actor_role','actor_id','action','from_counter','to_counter','queue_number','from_status','to_status']);
  }
}

/** ============ Utilities ============ **/
function headerMap_(sheet){
  const h = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0] || [];
  const m = {}; h.forEach((name,i)=> m[(name||'').toString().trim()] = i);
  return m;
}
function loadConfig_(){
  const ss = SpreadsheetApp.getActive();
  const tSh = ss.getSheetByName(SH_TYPES);
  const cSh = ss.getSheetByName(SH_COUNTER);
  const dSh = ss.getSheetByName(SH_DISPLAY);
  if(!tSh || !cSh) throw new Error('Missing config sheets');

  const tMap = headerMap_(tSh);
  const types = {};
  if (tSh.getLastRow()>1){
    const vals = tSh.getRange(2,1,tSh.getLastRow()-1,tSh.getLastColumn()).getValues();
    vals.forEach(r=>{
      const type=(r[tMap['type']]||'').toString().trim(); if(!type) return;
      types[type] = {
        prefix: (r[tMap['prefix']]||'').toString().trim(),
        color:  (r[tMap['color_hex']]||'#374151').toString().trim(),
        display:(r[tMap['display_name']]||type).toString().trim(),
        next:    parseInt(r[tMap['next_number']]||1,10),
        row:     2 + vals.indexOf(r),
        map:     tMap
      };
    });
  }

  const cMap = headerMap_(cSh);
  const counters = {};
  if (cSh.getLastRow()>1){
    const vals = cSh.getRange(2,1,cSh.getLastRow()-1,cSh.getLastColumn()).getValues();
    vals.forEach(r=>{
      const no = parseInt(r[cMap['counter_no']],10); if(!no) return;
      counters[no] = {
        type:   (r[cMap['type']]||'').toString().trim(),
        isOpen: (r[cMap['is_open']]==='TRUE' || r[cMap['is_open']]===true)
      };
    });
  }

  let display = {video_url:'', blink_ms:7000, max_slots:5};
  if (dSh){
    const dMap = headerMap_(dSh);
    if (dSh.getLastRow()>1){
      const r = dSh.getRange(2,1,1,dSh.getLastColumn()).getValues()[0]||[];
      display.video_url = (r[dMap['video_url']]||'').toString().trim();
      display.blink_ms  = parseInt(r[dMap['blink_ms']]||7000,10);
      display.max_slots = parseInt(r[dMap['max_slots']]||5,10);
    }
  }

  return {types, counters, display, tSh};
}
function loadQueue_(){
  const sh = SpreadsheetApp.getActive().getSheetByName(SH_QUEUE);
  if(!sh) throw new Error('Missing sheet: '+SH_QUEUE);
  const map = headerMap_(sh);
  const lastRow = sh.getLastRow();
  const vals = lastRow>1 ? sh.getRange(2,1,lastRow-1,sh.getLastColumn()).getValues() : [];
  return {sh, map, vals};
}
function logAction_(obj){
  const sh = SpreadsheetApp.getActive().getSheetByName(SH_LOG); if(!sh) return;
  sh.appendRow([new Date(), obj.actor_role||'system', obj.actor_id||Session.getActiveUser().getEmail(),
    obj.action||'', obj.from_counter||'', obj.to_counter||'', obj.queue_number||'', obj.from_status||'', obj.to_status||'']);
}

/** ============ doGet Router ============ **/
function doGet(e){
  const page = (e && e.parameter && e.parameter.page)||'central';
  const t = {
    kiosk:   'kiosk',
    desk:    'desk',
    central: 'central',
    display: 'display'
  }[page] || 'central';
  return HtmlService.createHtmlOutputFromFile(t)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('QueueV2 / '+t);
}

/** ============ Core helpers for queue ============ **/
function findCurrentCalled_(vals, map, counterNo){
  for (let i=vals.length-1;i>=0;i--){
    const r = vals[i];
    if (String(r[map['status']])===STATUS_CALLED && String(r[map['counter_no']])===String(counterNo)){
      return {rowIndex:i+2, row:r};
    }
  }
  return null;
}
function pickNextWaiting_(vals, map, counterNo, type){
  let pick=null;
  // 1) pin to counter
  for (let i=0;i<vals.length;i++){
    const r=vals[i];
    if (String(r[map['status']])!==STATUS_WAITING) continue;
    if (String(r[map['pin_counter']])===String(counterNo)) { pick={rowIndex:i+2,row:r}; break; }
  }
  if (pick) return pick;
  // 2) priority within type
  for (let i=0;i<vals.length;i++){
    const r=vals[i];
    if (String(r[map['status']])!==STATUS_WAITING) continue;
    if (String(r[map['type']])!==String(type)) continue;
    const pri = (r[map['priority']]===true || r[map['priority']]==='TRUE');
    if (pri) { pick={rowIndex:i+2,row:r}; break; }
  }
  if (pick) return pick;
  // 3) normal FIFO
  for (let i=0;i<vals.length;i++){
    const r=vals[i];
    if (String(r[map['status']])!==STATUS_WAITING) continue;
    if (String(r[map['type']])!==String(type)) continue;
    return {rowIndex:i+2,row:r};
  }
  return null;
}

/** ============ Public APIs for UIs ============ **/

/** ---- KIOSK ---- **/
function kioskTypeInfo(type){
  const {types} = loadConfig_();
  const t = types[type] || null;
  if(!t) throw new Error('Invalid type');
  return {type, display:t.display, color_hex:t.color};
}
function kioskTakeTicket(payload){
  // payload = {type, student_name, client_id}
  const lock = LockService.getScriptLock(); lock.tryLock(30000);
  try{
    const {types, tSh} = loadConfig_();
    const {sh, map} = loadQueue_();
    const type = (payload.type||'').trim();
    const name = (payload.student_name||'').trim();
    const cid  = (payload.client_id||'').trim();
    if(!types[type]) throw new Error('ไม่พบประเภทคิว');
    if(!name) throw new Error('กรุณากรอกชื่อ-นามสกุล');

    // rate-limit 5 นาที ต่อ 1 อุปกรณ์
    const now = Date.now();
    const cache = CacheService.getScriptCache();
    const key = `rate:${cid}`;
    const before = cache.get(key);
    if(before) throw new Error('รับคิวถี่เกินไป กรุณารอสักครู่ (≤5 นาที)');

    // สร้างเลขคิวจาก next_number ของประเภท
    const t = types[type];
    const raw = t.next;
    const qn  = `${t.prefix}${String(raw).padStart(3,'0')}`;

    sh.appendRow([new Date(), type, raw, qn, name, STATUS_WAITING, '', '', '', '',
                  false, '', cid, 0, '']);
    // update next_number
    const tMap = t.map;
    tSh.getRange(t.row, tMap['next_number']+1).setValue(raw+1);

    // set rate-limit TTL 300s
    cache.put(key, String(now), 300);

    logAction_({actor_role:'kiosk', action:'take', queue_number:qn, to_status:STATUS_WAITING});
    return {ok:true, queue_number:qn, type_display:t.display, color_hex:t.color};
  } finally { try{lock.releaseLock();}catch(_){ } }
}

/** ---- CENTRAL (iPad) ---- **/
function getCentralState(){
  const {types, counters} = loadConfig_();
  const {map, vals} = loadQueue_();

  const currentByCounter = {};
  for (let i=vals.length-1;i>=0;i--){
    const r=vals[i];
    if (String(r[map['status']])!==STATUS_CALLED) continue;
    const c = parseInt(r[map['counter_no']],10);
    if (c && currentByCounter[c]==null){
      currentByCounter[c] = {
        queue_number: (r[map['queue_number']]||'').toString(),
        type: (r[map['type']]||'').toString()
      };
    }
  }
  const out=[];
  for (let no=1; no<=10; no++){
    const cfg = counters[no] || {type:'',isOpen:false};
    const t = types[cfg.type] || {color:'#374151',display:cfg.type||'—'};
    const cur = currentByCounter[no] || {queue_number:'—', type: cfg.type};
    out.push({counter_no:no, type:cfg.type, type_display:t.display, color_hex:t.color, is_open:!!cfg.isOpen, current_queue:cur.queue_number});
  }
  return {counters:out, now:new Date()};
}
function next(counterNo){
  const lock = LockService.getScriptLock(); lock.tryLock(30000);
  try{
    const {counters} = loadConfig_();
    const cfg = counters[counterNo]; if(!cfg || !cfg.isOpen) throw new Error('โต๊ะนี้ปิดอยู่');
    const {sh, map, vals} = loadQueue_();

    // finish current if exist
    const cur = findCurrentCalled_(vals, map, counterNo);
    if (cur){ sh.getRange(cur.rowIndex, map['ts_finished']+1).setValue(new Date());
      logAction_({actor_role:'central', action:'finish-auto', from_counter:counterNo, queue_number: cur.row[map['queue_number']], from_status:STATUS_CALLED, to_status:STATUS_CALLED}); }

    const pick = pickNextWaiting_(vals, map, counterNo, cfg.type);
    if(!pick) throw new Error('ไม่มีคิวรอประเภทนี้');
    const r = pick.rowIndex, row = pick.row, now = new Date();
    sh.getRange(r, map['status']+1).setValue(STATUS_CALLED);
    sh.getRange(r, map['counter_no']+1).setValue(counterNo);
    sh.getRange(r, map['ts_called']+1).setValue(now);
    if (map['priority']!=null)    sh.getRange(r, map['priority']+1).setValue(false);
    if (map['pin_counter']!=null) sh.getRange(r, map['pin_counter']+1).setValue('');
    logAction_({actor_role:'central', action:'next', from_counter:counterNo, queue_number: row[map['queue_number']], from_status:STATUS_WAITING, to_status:STATUS_CALLED});
    return {ok:true};
  } finally { try{lock.releaseLock();}catch(_){ } }
}
function recall(counterNo){
  const lock = LockService.getScriptLock(); lock.tryLock(10000);
  try{
    const {sh, map, vals} = loadQueue_();
    const cur = findCurrentCalled_(vals, map, counterNo);
    if(!cur) throw new Error('ยังไม่มีคิวที่ถูกเรียก');
    const r = cur.rowIndex, row = cur.row;
    sh.getRange(r, map['ts_called']+1).setValue(new Date());
    if (map['recall_count']!=null){
      const c = parseInt(row[map['recall_count']]||0,10)+1;
      sh.getRange(r, map['recall_count']+1).setValue(c);
    }
    logAction_({actor_role:'central', action:'recall', from_counter:counterNo, queue_number: row[map['queue_number']], from_status:STATUS_CALLED, to_status:STATUS_CALLED});
    return {ok:true};
  } finally { try{lock.releaseLock();}catch(_){ } }
}
function back(counterNo){
  const lock = LockService.getScriptLock(); lock.tryLock(10000);
  try{
    const {sh, map, vals} = loadQueue_();
    const cur = findCurrentCalled_(vals, map, counterNo);
    if(!cur) throw new Error('ไม่มีคิวให้ถอย');
    const r=cur.rowIndex, row=cur.row;
    sh.getRange(r, map['status']+1).setValue(STATUS_WAITING);
    sh.getRange(r, map['counter_no']+1).setValue('');
    if (map['ts_called']!=null)   sh.getRange(r, map['ts_called']+1).setValue('');
    if (map['ts_finished']!=null) sh.getRange(r, map['ts_finished']+1).setValue('');
    if (map['pin_counter']!=null) sh.getRange(r, map['pin_counter']+1).setValue(counterNo); // กลับไปคิวถัดไปของโต๊ะเดิม
    logAction_({actor_role:'central', action:'back', from_counter:counterNo, queue_number: row[map['queue_number']], from_status:STATUS_CALLED, to_status:STATUS_WAITING});
    return {ok:true};
  } finally { try{lock.releaseLock();}catch(_){ } }
}

/** ---- DESK (per counter) ---- **/
function getDeskState(counterNo){
  const {types, counters} = loadConfig_();
  const {sh, map, vals} = loadQueue_();
  const cfg = counters[counterNo] || {type:'', isOpen:false};
  const tinfo = types[cfg.type] || {display:cfg.type, color:'#374151'};
  const cur = findCurrentCalled_(vals, map, counterNo);
  let current=null;
  if (cur){
    const r=cur.row;
    current = {
      queue_number: r[map['queue_number']]||'',
      student_name: r[map['student_name']]||'',
      ts_called:    r[map['ts_called']]||'',
    };
  }
  return {counter_no:counterNo, type:cfg.type, type_display:tinfo.display, color_hex:tinfo.color, is_open:cfg.isOpen, current};
}
function deskCall(counterNo){ return next(counterNo); } // alias
function deskRecall(counterNo){ return recall(counterNo); }
function deskFinish(counterNo){
  const lock = LockService.getScriptLock(); lock.tryLock(10000);
  try{
    const {sh, map, vals} = loadQueue_();
    const cur = findCurrentCalled_(vals, map, counterNo);
    if(!cur) throw new Error('ยังไม่มีคิวที่กำลังให้บริการ');
    const r=cur.rowIndex, row=cur.row;
    sh.getRange(r, map['ts_finished']+1).setValue(new Date());
    logAction_({actor_role:'desk', action:'finish', from_counter:counterNo, queue_number: row[map['queue_number']], from_status:STATUS_CALLED, to_status:STATUS_CALLED});
    return {ok:true};
  } finally { try{lock.releaseLock();}catch(_){ } }
}
function deskFinishAndNext(counterNo){ deskFinish(counterNo); return next(counterNo); }
function deskCancel(counterNo){
  const lock = LockService.getScriptLock(); lock.tryLock(10000);
  try{
    const {sh, map, vals} = loadQueue_();
    const cur = findCurrentCalled_(vals, map, counterNo);
    if(!cur) throw new Error('ไม่มีคิวให้ยกเลิก');
    const r=cur.rowIndex, row=cur.row;
    sh.getRange(r, map['status']+1).setValue(STATUS_CANCEL);
    sh.getRange(r, map['counter_no']+1).setValue('');
    sh.getRange(r, map['ts_finished']+1).setValue(new Date());
    logAction_({actor_role:'desk', action:'cancel', from_counter:counterNo, queue_number: row[map['queue_number']], from_status:STATUS_CALLED, to_status:STATUS_CANCEL});
    return {ok:true};
  } finally { try{lock.releaseLock();}catch(_){ } }
}
function transfer(payload){
  // payload = {fromCounter, toCounter?, toType?}
  const lock = LockService.getScriptLock(); lock.tryLock(30000);
  try{
    const {types, counters} = loadConfig_();
    const {sh, map, vals}   = loadQueue_();

    const from = parseInt(payload.fromCounter,10);
    const cur = findCurrentCalled_(vals, map, from);
    if(!cur) throw new Error('ไม่พบคิวที่กำลังให้บริการของโต๊ะต้นทาง');

    let toCounter = payload.toCounter ? parseInt(payload.toCounter,10) : null;
    let toType    = (payload.toType||'').trim();

    if (toCounter!=null){
      const cfg = counters[toCounter]; if(!cfg || !cfg.isOpen) throw new Error('โต๊ะปลายทางปิดหรือไม่พบ');
      toType = cfg.type;
    } else {
      if(!types[toType]) throw new Error('ประเภทปลายทางไม่ถูกต้อง');
    }

    const r = cur.rowIndex, row = cur.row;
    // เปลี่ยนเป็น waiting ที่ปลายทาง (ไม่ออกเลขใหม่)
    sh.getRange(r, map['status']+1).setValue(STATUS_WAITING);
    sh.getRange(r, map['counter_no']+1).setValue('');
    sh.getRange(r, map['type']+1).setValue(toType);
    if (map['ts_called']!=null)   sh.getRange(r, map['ts_called']+1).setValue('');
    if (map['pin_counter']!=null && toCounter!=null) sh.getRange(r, map['pin_counter']+1).setValue(toCounter);
    if (map['pin_counter']!=null && toCounter==null) sh.getRange(r, map['pin_counter']+1).setValue('');
    logAction_({actor_role:'desk/central', action:'transfer', from_counter:from, to_counter:toCounter||'', queue_number: row[map['queue_number']], from_status:STATUS_CALLED, to_status:STATUS_WAITING});
    return {ok:true};
  } finally { try{lock.releaseLock();}catch(_){ } }
}
function setCounterOpen(counterNo, isOpen){
  const ss = SpreadsheetApp.getActive();
  const cSh = ss.getSheetByName(SH_COUNTER);
  const map = headerMap_(cSh);
  const vals = cSh.getRange(2,1,cSh.getLastRow()-1,cSh.getLastColumn()).getValues();
  for (let i=0;i<vals.length;i++){
    const r=vals[i];
    if (parseInt(r[map['counter_no']],10)===parseInt(counterNo,10)){
      cSh.getRange(2+i, map['is_open']+1).setValue(!!isOpen);
      return {ok:true};
    }
  }
  throw new Error('ไม่พบโต๊ะ');
}

/** ---- DISPLAY ---- **/
function getDisplayFeed(){
  const {display, types} = loadConfig_();
  const {map, vals} = loadQueue_();

  // ล่าสุดสถานะ called ตามเวลา ts_called (ไม่รวม cancelled)
  const called = vals
    .filter(r=> String(r[map['status']])===STATUS_CALLED)
    .map(r=>({
      queue_number: (r[map['queue_number']]||'')+'',
      type: (r[map['type']]||'')+'',
      counter_no: r[map['counter_no']],
      ts_called: r[map['ts_called']] ? new Date(r[map['ts_called']]) : null
    }))
    .sort((a,b)=>(b.ts_called||0)-(a.ts_called||0))
    .slice(0, display.max_slots);

  // enrich color & display name
  called.forEach(x=>{
    const t = types[x.type] || {display:x.type, color:'#374151'};
    x.type_display = t.display;
    x.color_hex    = t.color;
  });

  return {slots:called, blink_ms:display.blink_ms, now:new Date()};
}

