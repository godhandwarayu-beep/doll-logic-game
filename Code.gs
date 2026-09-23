/**
 * ภารกิจตุ๊กตาแม่ลูกดก — ระบบบันทึกคะแนนลง Google Sheets
 * วิธีติดตั้ง: ส่วนขยาย > Apps Script > วางโค้ดนี้ > การทำให้ใช้งานได้ > การทำให้ใช้งานได้รายการใหม่
 *   ประเภท: เว็บแอป | เรียกใช้ในฐานะ: ฉัน | ผู้ที่มีสิทธิ์เข้าถึง: ทุกคน
 * แล้วนำ URL ที่ลงท้ายด้วย /exec ไปใส่ในตัวแปร API_URL ในไฟล์ index.html
 */
const SHEET_ID = '1b08EFGBe3EfDDifSIKG6saqtmtJJgliio88ZyVxiPOw';
const LOG_SHEET = 'บันทึกการเล่น';
const SUM_SHEET = 'สรุปรายคน';
const HEADERS = ['วันเวลา', 'ชื่อ-นามสกุล', 'ชั้น/ห้อง', 'เลขที่', 'ด่าน', 'ดาว', 'จำนวนครั้งที่ลอง', 'เวลา (วินาที)', 'จำนวนบล็อก'];

function getLog_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET);
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#ffe0b2');
    sh.setFrozenRows(1);
  }
  return sh;
}

function clean_(v, max) {
  let s = String(v == null ? '' : v).trim().slice(0, max || 60);
  if (/^[=+\-@]/.test(s)) s = "'" + s; // กันสูตรแปลกปลอม
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** รับคะแนนจากเกม */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents);
    const level = Math.max(1, Math.min(50, parseInt(d.level, 10) || 0));
    const stars = Math.max(1, Math.min(3, parseInt(d.stars, 10) || 0));
    getLog_().appendRow([
      new Date(), clean_(d.name), clean_(d.cls, 20), clean_(d.no, 5), level, stars,
      parseInt(d.attempts, 10) || 1, parseInt(d.time, 10) || 0, parseInt(d.blocks, 10) || 0
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** ผลดีที่สุดของแต่ละคน แต่ละด่าน */
function bestByPlayer_() {
  const rows = getLog_().getDataRange().getValues().slice(1);
  const players = {};
  rows.forEach(r => {
    const name = String(r[1]), cls = String(r[2]), no = String(r[3]);
    if (!name) return;
    const key = name + '|' + cls + '|' + no;
    const p = players[key] || (players[key] = { name, cls, no, levels: {}, tries: 0 });
    const lv = r[4], stars = Number(r[5]), time = Number(r[7]);
    p.tries += Number(r[6]) || 0;
    const b = p.levels[lv];
    if (!b || stars > b.stars || (stars === b.stars && time < b.time)) p.levels[lv] = { stars, time };
  });
  return Object.values(players).map(p => {
    const lv = Object.values(p.levels);
    return {
      name: p.name, cls: p.cls, no: p.no, levelMap: p.levels, tries: p.tries,
      levels: lv.length,
      stars: lv.reduce((s, x) => s + x.stars, 0),
      time: lv.reduce((s, x) => s + x.time, 0)
    };
  }).sort((a, b) => b.stars - a.stars || b.levels - a.levels || a.time - b.time);
}

/** ส่งข้อมูลให้เกม: ?action=leaderboard  หรือ  ?action=progress&name=..&cls=..&no=.. */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const all = bestByPlayer_();
  if (p.action === 'progress') {
    const me = all.find(x => x.name === String(p.name) && x.cls === String(p.cls) && String(x.no) === String(p.no));
    return json_({ ok: true, levels: me ? me.levelMap : {} });
  }
  const list = p.cls ? all.filter(x => x.cls === p.cls) : all;
  return json_({
    ok: true,
    rows: list.slice(0, 20).map(x => ({ name: x.name, cls: x.cls, no: x.no, levels: x.levels, stars: x.stars, time: x.time }))
  });
}

/** (ไม่บังคับ) ครูกดรันฟังก์ชันนี้เพื่อสร้างแผ่นงานสรุปรายคน */
function buildSummary() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SUM_SHEET) || ss.insertSheet(SUM_SHEET);
  sh.clear();
  const head = ['อันดับ', 'ชื่อ-นามสกุล', 'ชั้น/ห้อง', 'เลขที่', 'ผ่าน (ด่าน)', 'ดาวรวม', 'เวลารวม (วินาที)', 'ลองทั้งหมด (ครั้ง)'];
  for (let i = 1; i <= 10; i++) head.push('ด่าน ' + i + ' (ดาว)');
  const data = bestByPlayer_().map((x, i) => {
    const r = [i + 1, x.name, x.cls, x.no, x.levels, x.stars, x.time, x.tries];
    for (let k = 1; k <= 10; k++) r.push(x.levelMap[k] ? x.levelMap[k].stars : '');
    return r;
  });
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#c8e6c9');
  if (data.length) sh.getRange(2, 1, data.length, head.length).setValues(data);
  sh.setFrozenRows(1);
}
