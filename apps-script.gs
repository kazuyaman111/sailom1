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
const HEADERS = ['studentId', 'spiritName', 'level', 'coins', 'stats', 'lastUpdated', 'saveRaw'];

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
    const studentId = e && e.parameter ? String(e.parameter.studentId || '').trim() : '';
    if (!studentId) return json_({ status: 'error', message: 'ไม่ได้ระบุรหัสนักเรียน' });

    const sh = getSheet_();

    /* โหมดสรุปทั้งห้องสำหรับครู: เรียก ?all=1 */
    if (e.parameter.all === '1') {
      const last = sh.getLastRow();
      const rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, 6).getValues();
      return json_({
        status: 'success',
        count: rows.length,
        data: rows.map(r => ({
          studentId: r[0], spiritName: r[1], level: r[2],
          coins: r[3], stats: r[4], lastUpdated: r[5]
        }))
      });
    }

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

/* ---------- บันทึกเซฟ ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    if (!e || !e.postData || !e.postData.contents)
      return json_({ status: 'error', message: 'ไม่มีข้อมูลส่งมา' });

    const body = JSON.parse(e.postData.contents);
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
      rawStr
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
