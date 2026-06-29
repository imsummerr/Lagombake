// ══════════════════════════════════════════════════════════════════
//  Lagombake — Google Apps Script  v5
// ══════════════════════════════════════════════════════════════════

const SHEET_NAME         = 'Orders';
const LINE_CHANNEL_TOKEN = '';   // ← ใส่ Channel Access Token
const LINE_TARGET_ID     = '';   // ← ใส่ User ID หรือ Group ID

// ── รับ POST ─────────────────────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    // Webhook จาก LINE (มี events) → ตอบ ID กลับ
    if (body.events) {
      handleLineWebhook(body.events);
      return jsonResponse({ status: 'ok' });
    }

    // ออเดอร์จากเว็บ
    const slipUrls = body.slip ? saveSlipToDrive(body.slip, body.customerName) : null;
    saveToSheet(body, slipUrls ? slipUrls.viewUrl : '');
    sendLineMessage(body, slipUrls);
    return jsonResponse({ status: 'ok' });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function doGet() {
  return ContentService
    .createTextOutput('✓ Lagombake Script is running')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ── บันทึกสลิปลง Google Drive ────────────────────────────────────
// คืนค่า { viewUrl, directUrl } หรือ null ถ้าล้มเหลว
function saveSlipToDrive(base64Data, customerName) {
  try {
    const matches = base64Data.match(/^data:([^;]+);base64,(.+)$/);
    if (!matches) return null;
    const mimeType = matches[1];
    const ext      = mimeType.split('/')[1] || 'jpg';
    const data     = Utilities.base64Decode(matches[2]);
    const blob     = Utilities.newBlob(data, mimeType);

    const timestamp = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss');
    const fileName  = `slip_${customerName}_${timestamp}.${ext}`;
    blob.setName(fileName);

    let folder;
    const folders = DriveApp.getFoldersByName('Lagombake Slips');
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('Lagombake Slips');

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return {
      viewUrl  : file.getUrl(),
      directUrl: 'https://drive.google.com/uc?export=view&id=' + file.getId()
    };
  } catch (err) {
    Logger.log('saveSlipToDrive error: ' + err);
    return null;
  }
}

// ── บันทึกลง Google Sheet ────────────────────────────────────────
function saveToSheet(data, slipUrl) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const hRange = sheet.getRange(1, 1, 1, 6);
    hRange.setValues([['วันที่/เวลา', 'ผู้สั่ง', 'รายการที่สั่ง', 'ราคา (฿)', 'หมายเหตุ', 'สลิป']]);
    hRange.setBackground('#4A3728');
    hRange.setFontColor('#FFFFFF');
    hRange.setFontWeight('bold');
    hRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 480);
    sheet.setColumnWidth(4, 90);
    sheet.setColumnWidth(5, 200);
    sheet.setColumnWidth(6, 200);
  }

  const timestamp = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
  sheet.appendRow([
    timestamp,
    data.customerName || '',
    data.items        || '',
    data.totalPrice   || 0,
    data.remark       || '',
    slipUrl           || ''
  ]);

  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1, 1, 6)
       .setBackground(lastRow % 2 === 0 ? '#F7F3EE' : '#FFFFFF');
  sheet.getRange(lastRow, 3).setWrap(true);
  sheet.getRange(lastRow, 4).setHorizontalAlignment('right').setFontWeight('bold');

  if (slipUrl) {
    sheet.getRange(lastRow, 6)
         .setFormula(`=HYPERLINK("${slipUrl}","ดูสลิป")`);
  }
}

// ── แจ้งเตือน LINE ───────────────────────────────────────────────
// slipUrls = { viewUrl, directUrl } หรือ null
function sendLineMessage(data, slipUrls) {
  if (!LINE_CHANNEL_TOKEN || !LINE_TARGET_ID) return;

  const now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'HH:mm');
  const itemLines = (data.items || '')
    .split(' | ')
    .map(item => '  • ' + item.trim())
    .join('\n');

  const slipLine = slipUrls
    ? '🧾 สลิป: ' + slipUrls.viewUrl + '\n'
    : '⚠️ ไม่มีสลิป\n';

  const textMsg =
    '━━━━━━━━━━━━━━━━━━\n' +
    '🧁 ออเดอร์ใหม่ Lagombake\n' +
    '━━━━━━━━━━━━━━━━━━\n' +
    '🕐 เวลา  : ' + now + ' น.\n' +
    '👤 ชื่อ   : ' + (data.customerName || '-') + '\n' +
    '━━━━━━━━━━━━━━━━━━\n' +
    '📋 รายการ\n' + itemLines + '\n' +
    '━━━━━━━━━━━━━━━━━━\n' +
    '💰 รวม   : ' + (data.totalPrice || 0) + ' ฿\n' +
    (data.remark ? '📝 หมายเหตุ: ' + data.remark + '\n' : '') +
    slipLine +
    '━━━━━━━━━━━━━━━━━━';

  // สร้าง messages array — ถ้ามีสลิปให้ส่ง image ต่อท้าย
  const messages = [{ type: 'text', text: textMsg }];
  if (slipUrls) {
    messages.push({
      type              : 'image',
      originalContentUrl: slipUrls.directUrl,
      previewImageUrl   : slipUrls.directUrl
    });
  }

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method  : 'post',
    headers : {
      'Content-Type' : 'application/json',
      'Authorization': 'Bearer ' + LINE_CHANNEL_TOKEN
    },
    payload : JSON.stringify({
      to      : LINE_TARGET_ID,
      messages: messages
    }),
    muteHttpExceptions: true
  });
}

// ── Webhook: ตอบ Group/User ID กลับ ─────────────────────────────
function handleLineWebhook(events) {
  events.forEach(event => {
    if (event.type !== 'message') return;
    const sourceType = event.source.type;
    const id = sourceType === 'group' ? event.source.groupId : event.source.userId;
    const msg = sourceType === 'group'
      ? '🆔 Group ID:\n' + id
      : '🆔 User ID:\n' + id;

    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
      method  : 'post',
      headers : {
        'Content-Type' : 'application/json',
        'Authorization': 'Bearer ' + LINE_CHANNEL_TOKEN
      },
      payload : JSON.stringify({
        replyToken: event.replyToken,
        messages  : [{ type: 'text', text: msg }]
      }),
      muteHttpExceptions: true
    });
  });
}

// ── ทดสอบ LINE ───────────────────────────────────────────────────
function testSendLine() {
  sendLineMessage({
    customerName: 'ทดสอบ',
    items       : 'อเมริกาโน่ (น้อย) | แซนด์วิช [มายองเนส, ไข่ดาว]',
    totalPrice  : 74,
    remark      : 'ไม่ใส่น้ำแข็ง'
  }, null);
  Logger.log('ส่งทดสอบแล้ว');
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
