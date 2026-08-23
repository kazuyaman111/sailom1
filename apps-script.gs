/* ===========================================================
   KruChat_Classroom_DB — Apps Script รับส่งเซฟเกม
   ห้องเรียนไร้ขีดจำกัด #ห้องเรียนครูชัช

   วิธีติดตั้ง
   1. เปิด Google Sheets ชื่อ KruChat_Classroom_DB ที่สร้างไว้
   2. เมนู ส่วนขยาย → Apps Script
   3. ลบโค้ดเดิมทั้งหมด แล้ววางไฟล์นี้ลงไป
   4. กดบันทึก แล้วกด ทำให้ใช้งานได้ → การทำให้ใช้งานได้ใหม่
        ประเภท: เว็บแอป
        ดำเนินการในชื่อ: ฉัน
        ผู้ที่มีสิทธิ์เข้าถึง: ทุกคน        ← สำคัญ ถ้าเลือกผิดนักเรียนจะเชื่อมต่อไม่ได้
   5. คัดลอก URL ที่ได้ (ลงท้ายด้วย /exec) ไปวางในเกมที่ตัวแปร CLOUD_API_URL
   6. ทุกครั้งที่แก้โค้ดนี้ ต้องกด ทำให้ใช้งานได้ใหม่ อีกครั้ง ไม่อย่างนั้นของเดิมจะยังทำงานอยู่
   =========================================================== */

const SHEET_NAME = 'SaveData';
const MQ_SHEET   = 'Missions';
const HIT_SHEET  = 'Visits';      // ตัวนับจำนวนครั้งที่เปิดเว็บ
const MQ_HEADERS = ['grade', 'subj', 'unit', 'need', 'window', 'note', 'until', 'createdAt'];
const HEADERS = ['studentId', 'spiritName', 'level', 'coins', 'stats',
                 'lastUpdated', 'saveRaw', 'grade', 'progress'];

/* รหัสผ่านสำหรับเปิดแดชบอร์ดครู เปลี่ยนเป็นอะไรก็ได้ที่เดายาก
   ถ้าไม่ตั้ง ใครที่รู้ URL ก็ดูข้อมูลทั้งห้องได้ */
const TEACHER_KEY = 'kruchat2569';

/* จำกัดขนาดก้อนเซฟ กันช่องเดียวยาวเกินที่ Google Sheets รับได้ (50,000 ตัวอักษร) */
const MAX_RAW = 45000;

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function getMqSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(MQ_SHEET);
  if (!sh) {
    sh = ss.insertSheet(MQ_SHEET);
    sh.getRange(1, 1, 1, MQ_HEADERS.length).setValues([MQ_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ชีตนับจำนวนครั้งที่เปิดเว็บ เก็บแค่ตัวเลขรวม ไม่เก็บว่าใครหรือเมื่อไหร่ */
function getHitSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HIT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(HIT_SHEET);
    sh.getRange('A1:B1').setValues([['รายการ', 'จำนวน']]).setFontWeight('bold');
    sh.getRange('A2:B2').setValues([['เปิดเว็บทั้งหมด (ครั้ง)', 0]]);
    sh.getRange('A3:B3').setValues([['เริ่มนับเมื่อ', new Date()]]);
    sh.setColumnWidth(1, 220);
  }
  return sh;
}

function bumpHit_() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const cell = getHitSheet_().getRange('B2');
    const n = Number(cell.getValue()) || 0;
    cell.setValue(n + 1);
    return n + 1;
  } catch (err) {
    return -1;
  } finally {
    try { lock.releaseLock(); } catch (err) {}
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* หาแถวของรหัสนักเรียน คืน 0 ถ้าไม่พบ */
function findRow_(sh, studentId) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  const key = String(studentId).trim().toLowerCase();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim().toLowerCase() === key) return i + 2;
  }
  return 0;
}

/* ---------- ดึงเซฟ ---------- */
function doGet(e) {
  try {
    const p = (e && e.parameter) ? e.parameter : {};
    /* ---------- นับจำนวนครั้งที่เปิดเว็บ: ?hit=1 ---------- */
    if (p.hit === '1') {
      return json_({ status: 'success', total: bumpHit_() });
    }

    const sh = getSheet_();

    /* ---------- โหมดแดชบอร์ดครู: ?all=1&key=รหัสผ่าน ---------- */
    if (p.all === '1') {
      if (String(p.key || '') !== TEACHER_KEY)
        return json_({ status: 'error', message: 'รหัสผ่านครูไม่ถูกต้อง' });

      const last = sh.getLastRow();
      const rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
      return json_({
        status: 'success',
        count: rows.length,
        serverTime: new Date().toISOString(),
        data: rows.map(function (r) {
          var st = {}, pg = {};
          try { st = JSON.parse(r[4]); } catch (err) {}
          try { pg = JSON.parse(r[8]); } catch (err) {}
          return {
            studentId: r[0], spiritName: r[1], level: r[2], coins: r[3],
            stats: st, lastUpdated: r[5], grade: r[7], progress: pg
          };
        })
      });
    }

    /* ---------- ภารกิจที่ครูสั่ง: ?missions=1&grade=p5 ---------- */
    if (p.missions === '1') {
      const ms = getMqSheet_();
      const last = ms.getLastRow();
      const rows = last < 2 ? [] : ms.getRange(2, 1, last - 1, MQ_HEADERS.length).getValues();
      const want = String(p.grade || '').trim();
      const now = new Date().getTime();
      const out = [];
      rows.forEach(function (r) {
        const g = String(r[0] || '').trim();
        if (g && want && g !== want) return;
        const until = r[6] ? new Date(r[6]).getTime() : 0;
        if (until && now > until) return;
        if (!r[2]) return;
        out.push({
          grade: g, subj: String(r[1] || 'hist').trim(), unit: String(r[2]).trim(),
          need: Number(r[3]) || 8, win: Number(r[4]) || 12,
          note: String(r[5] || ''), until: until
        });
      });
      return json_({ status: 'success', count: out.length, data: out.slice(0, 3) });
    }

    const studentId = String(p.studentId || '').trim();
    if (!studentId) return json_({ status: 'error', message: 'ไม่ได้ระบุรหัสนักเรียน' });

    const row = findRow_(sh, studentId);
    if (!row) return json_({ status: 'error', message: 'ไม่พบข้อมูลของรหัสนี้' });

    const v = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
    let raw = null;
    try { raw = JSON.parse(v[6]); } catch (err) { raw = null; }

    return json_({
      status: 'success',
      data: {
        studentId: v[0], spiritName: v[1], level: v[2], coins: v[3],
        stats: v[4], lastUpdated: v[5], saveRaw: raw
      }
    });
  } catch (err) {
    return json_({ status: 'error', message: String(err) });
  }
}

/* ---------- ครูสั่งภารกิจจากแดชบอร์ด ---------- */
function addMission_(body) {
  if (String(body.key || '') !== TEACHER_KEY)
    return json_({ status: 'error', message: 'รหัสผ่านครูไม่ถูกต้อง' });
  const ms = getMqSheet_();
  if (body.clear) {
    const last = ms.getLastRow();
    if (last > 1) ms.deleteRows(2, last - 1);
    return json_({ status: 'success', cleared: true });
  }
  if (!body.unit) return json_({ status: 'error', message: 'ไม่ได้ระบุหน่วยการเรียนรู้' });
  const until = new Date();
  until.setDate(until.getDate() + (Number(body.days) || 7));
  ms.appendRow([
    String(body.grade || ''), String(body.subj || 'hist'), String(body.unit),
    Number(body.need) || 8, Number(body.win) || 12,
    String(body.note || '').slice(0, 120), until, new Date()
  ]);
  return json_({ status: 'success', until: until.toISOString() });
}

/* ---------- บันทึกเซฟ ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    if (!e || !e.postData || !e.postData.contents)
      return json_({ status: 'error', message: 'ไม่มีข้อมูลส่งมา' });

    const body = JSON.parse(e.postData.contents);

    /* ครูสั่งภารกิจ ใช้ POST เดียวกันแต่ระบุ action */
    if (body.action === 'mission') return addMission_(body);

    const studentId = String(body.studentId || '').trim();
    if (!studentId) return json_({ status: 'error', message: 'ไม่ได้ระบุรหัสนักเรียน' });
    if (!/^[A-Za-z0-9ก-๙_-]{3,20}$/.test(studentId))
      return json_({ status: 'error', message: 'รูปแบบรหัสนักเรียนไม่ถูกต้อง' });

    const rawStr = JSON.stringify(body.saveRaw || {});
    if (rawStr.length > MAX_RAW)
      return json_({ status: 'error', message: 'ก้อนเซฟใหญ่เกินไป' });

    const sh = getSheet_();
    const rowData = [
      studentId,
      String(body.spiritName || '').slice(0, 40),
      Number(body.level || 0),
      Number(body.coins || 0),
      JSON.stringify(body.stats || {}),
      new Date(),
      rawStr,
      String(body.grade || ''),
      JSON.stringify(body.progress || {})
    ];

    const row = findRow_(sh, studentId);
    if (row) sh.getRange(row, 1, 1, HEADERS.length).setValues([rowData]);
    else sh.appendRow(rowData);

    return json_({ status: 'success', studentId: studentId, row: row || sh.getLastRow() });
  } catch (err) {
    return json_({ status: 'error', message: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (err) {}
  }
}
