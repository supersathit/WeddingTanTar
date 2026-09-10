// Paste into Extensions > Apps Script in the destination Google Sheet.
// Run setup once, then deploy as Web app: Execute as Me; access Anyone.
function setup() {
  const properties = PropertiesService.getScriptProperties();
  const book = SpreadsheetApp.getActiveSpreadsheet();
  if (!book) throw new Error('Open Apps Script from your Google Sheet first.');
  properties.setProperty('SPREADSHEET_ID', book.getId());
  if (!properties.getProperty('SLIP_FOLDER_ID')) {
    const folder = DriveApp.createFolder('TanTar Wedding - Slips');
    properties.setProperty('SLIP_FOLDER_ID', folder.getId());
  }
  ensureSheet_(book, 'RSVP', ['Request ID', 'วันที่ส่ง', 'ชื่อ–นามสกุล', 'การเข้าร่วม', 'จำนวนคน', 'ช่วงที่ร่วมงาน', 'คำอวยพร']);
  ensureSheet_(book, 'Slips', ['Request ID', 'วันที่ส่ง', 'ชื่อ–นามสกุล', 'ชื่อไฟล์', 'ลิงก์สลิป']);
  book.getSheetByName('RSVP').getRange(1, 8).setValue('รหัสอ้างอิง');
  book.getSheetByName('Slips').getRange(1, 6).setValue('รหัสอ้างอิง');
  book.getSheetByName('RSVP').getRange(1, 9, 1, 2).setValues([['ยินยอมเผยแพร่', 'ชื่อที่แสดง']]);
}

function ensureSheet_(book, name, headers) {
  const sheet = book.getSheetByName(name) || book.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E2B75C');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Run in the Apps Script editor (not a public endpoint) to identify the actual destination.
function checkConnection() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('ยังไม่ได้ตั้งค่า SPREADSHEET_ID กรุณารัน setup ก่อน');
  const book = SpreadsheetApp.openById(id);
  console.log('ชีตที่ระบบใช้งาน: ' + book.getName());
  console.log('เปิดชีตปลายทาง: ' + book.getUrl());
  ['RSVP', 'Slips'].forEach(name => {
    const sheet = book.getSheetByName(name);
    if (!sheet) throw new Error('ไม่พบแท็บ ' + name + ' กรุณารัน setup');
    console.log(name + ': ' + Math.max(0, sheet.getLastRow() - 1) + ' รายการ');
  });
}

function text_(value, limit, required) {
  const text = String(value == null ? '' : value).trim();
  if ((required && !text) || text.length > limit) throw new Error('กรุณาตรวจสอบข้อมูลที่กรอก');
  // Treat guest-supplied text as text, never as a spreadsheet formula.
  return /^[=+@-]/.test(text) ? "'" + text : text;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    if (!e || !e.postData || e.postData.contents.length > 14500000) throw new Error('ข้อมูลหรือไฟล์มีขนาดใหญ่เกินไป');
    const data = JSON.parse(e.postData.contents);
    if (!['rsvp', 'slip'].includes(data.type)) throw new Error('ประเภทฟอร์มไม่ถูกต้อง');
    if (!/^[a-zA-Z0-9-]{16,80}$/.test(data.requestId || '')) throw new Error('รหัสรายการไม่ถูกต้อง');
    const name = text_(data.name, 150, true);
    const props = PropertiesService.getScriptProperties();
    const book = SpreadsheetApp.openById(props.getProperty('SPREADSHEET_ID'));
    const sheet = book.getSheetByName(data.type === 'rsvp' ? 'RSVP' : 'Slips');
    if (!sheet) throw new Error('กรุณาตั้งค่าระบบก่อนใช้งาน');
    lock.waitLock(30000);
    locked = true;
    const receiptColumn = data.type === 'rsvp' ? 8 : 6;
    const existing = sheet.getLastRow() > 1 && sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
      .createTextFinder(data.requestId).matchEntireCell(true).findNext();
    if (existing) {
      const cell = sheet.getRange(existing.getRow(), receiptColumn);
      let receiptId = String(cell.getValue() || '');
      if (!receiptId) { receiptId = nextReceiptId_(props); cell.setValue(receiptId); }
      return json_({ok:true, requestId:data.requestId, receiptId});
    }
    const receiptId = nextReceiptId_(props);
    if (data.type === 'rsvp') {
      if (!['มา', 'ไม่มา'].includes(data.attend)) throw new Error('กรุณาเลือกการเข้าร่วมงาน');
      const attending = data.attend === 'มา';
      const guests = attending ? Number(data.guests) : 0;
      if (attending && (!Number.isInteger(guests) || guests < 1 || guests > 20)) throw new Error('จำนวนผู้ร่วมงานต้องอยู่ระหว่าง 1–20 คน');
      if (attending && !['ทั้งวัน', 'พิธีเช้า', 'งานเลี้ยง'].includes(data.session)) throw new Error('กรุณาเลือกช่วงที่ร่วมงาน');
      sheet.appendRow([data.requestId, new Date(), name, data.attend, guests,
        attending ? data.session : '', text_(data.wish, 3000, false), receiptId,
        data.publishWish === 'yes', data.publishWish === 'yes' ? text_(data.publicName, 80, true) : '']);
    } else {
      const extensions = {'image/jpeg':'jpg', 'image/png':'png', 'image/webp':'webp', 'application/pdf':'pdf'};
      if (!Object.prototype.hasOwnProperty.call(extensions, data.mimeType) || typeof data.base64 !== 'string') throw new Error('ชนิดไฟล์ไม่รองรับ');
      const bytes = Utilities.base64Decode(data.base64);
      if (!bytes.length || bytes.length > 10 * 1024 * 1024) throw new Error('ไฟล์ต้องมีขนาดไม่เกิน 10 MB');
      const head = bytes.slice(0, 12).map(b => b & 255);
      const valid = data.mimeType === 'image/jpeg' ? head[0] === 255 && head[1] === 216 && head[2] === 255
        : data.mimeType === 'image/png' ? head.slice(0, 8).join(',') === '137,80,78,71,13,10,26,10'
        : data.mimeType === 'application/pdf' ? String.fromCharCode.apply(null, head.slice(0, 5)) === '%PDF-'
        : String.fromCharCode.apply(null, head.slice(0, 4)) === 'RIFF' && String.fromCharCode.apply(null, head.slice(8, 12)) === 'WEBP';
      if (!valid) throw new Error('รูปแบบไฟล์ไม่ตรงกับชนิดไฟล์ที่เลือก');
      text_(data.fileName, 255, true);
      const safeName = String(data.name).trim().replace(/[\\/:*?"<>|\u0000-\u001f\u007f]/g, '_');
      const fileName = receiptId + '_' + safeName + '.' + extensions[data.mimeType];
      const folder = DriveApp.getFolderById(props.getProperty('SLIP_FOLDER_ID'));
      const file = folder.createFile(Utilities.newBlob(bytes, data.mimeType, fileName));
      // Keep Drive files private; do not enable public-link sharing.
      try {
        sheet.appendRow([data.requestId, new Date(), name, fileName, file.getUrl(), receiptId]);
      } catch (error) {
        file.setTrashed(true);
        throw error;
      }
    }
    SpreadsheetApp.flush();
    return json_({ok:true, requestId:data.requestId, receiptId});
  } catch (error) {
    console.error(error);
    return json_({ok:false, message:'บันทึกไม่สำเร็จ กรุณาตรวจสอบข้อมูลแล้วลองอีกครั้ง'});
  } finally {
    if (locked) lock.releaseLock();
  }
}

// Public endpoint: publish wishes immediately with explicit guest consent.
function doGet(e) {
  try {
    if (!e || !e.parameter || e.parameter.action !== 'wishes') return json_({ok:false});
    const book = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
    const sheet = book.getSheetByName('RSVP');
    const rows = sheet && sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues() : [];
    const enabled = value => value === true || value === 'TRUE';
    const wishes = rows.filter(row => enabled(row[8]) && String(row[9] || '').trim() && String(row[6] || '').trim())
      .reverse().slice(0, 60).map(row => ({name:String(row[9]), wish:String(row[6])}));
    return json_({ok:true, wishes});
  } catch (error) {
    console.error(error);
    return json_({ok:false, message:'โหลดคำอวยพรไม่สำเร็จ'});
  }
}

// Called only while holding the script lock. Reserve IDs even if a write fails.
function nextReceiptId_(props) {
  const next = Number(props.getProperty('LAST_RECEIPT_ID') || 9999) + 1;
  if (!Number.isInteger(next) || next < 10000 || next > 99999) throw new Error('รหัสอ้างอิงเต็มแล้ว');
  props.setProperty('LAST_RECEIPT_ID', String(next));
  return String(next);
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
