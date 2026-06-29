const { google } = require('googleapis');

const SHEET_NAME = 'Orders';

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method Not Allowed' };
  }

  try {
    const body = JSON.parse(event.body);

    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT),
      scopes: [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive',
      ],
    });

    const slipUrls = body.slip
      ? await saveSlipToDrive(auth, body.slip, body.customerName)
      : null;

    await saveToSheet(auth, body, slipUrls ? slipUrls.viewUrl : '');
    await sendLineMessage(body, slipUrls);

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ status: 'ok' }),
    };
  } catch (err) {
    console.error(err);
    return {
      statusCode: 500,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ status: 'error', message: err.message }),
    };
  }
};

// ── บันทึกสลิปลง Google Drive ────────────────────────────────────
async function saveSlipToDrive(auth, base64Data, customerName) {
  const matches = base64Data.match(/^data:([^;]+);base64,(.+)$/);
  if (!matches) return null;

  const mimeType = matches[1];
  const ext      = mimeType.split('/')[1] || 'jpg';
  const buffer   = Buffer.from(matches[2], 'base64');

  const drive     = google.drive({ version: 'v3', auth });
  const timestamp = new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' })
    .replace(/[/:, ]/g, '_');
  const fileName  = `slip_${customerName}_${timestamp}.${ext}`;

  // หาหรือสร้างโฟลเดอร์ "Lagombake Slips"
  const folderSearch = await drive.files.list({
    q: "name='Lagombake Slips' and mimeType='application/vnd.google-apps.folder' and trashed=false",
    fields: 'files(id)',
  });

  let folderId;
  if (folderSearch.data.files.length > 0) {
    folderId = folderSearch.data.files[0].id;
  } else {
    const folder = await drive.files.create({
      requestBody: { name: 'Lagombake Slips', mimeType: 'application/vnd.google-apps.folder' },
      fields: 'id',
    });
    folderId = folder.data.id;
  }

  // อัปโหลดไฟล์
  const { Readable } = require('stream');
  const stream = new Readable();
  stream.push(buffer);
  stream.push(null);

  const file = await drive.files.create({
    requestBody: { name: fileName, parents: [folderId] },
    media: { mimeType, body: stream },
    fields: 'id',
  });

  const fileId = file.data.id;

  // ตั้งให้ทุกคนที่มีลิงก์เข้าได้
  await drive.permissions.create({
    fileId,
    requestBody: { role: 'reader', type: 'anyone' },
  });

  return {
    viewUrl  : `https://drive.google.com/file/d/${fileId}/view`,
    directUrl: `https://drive.google.com/uc?export=view&id=${fileId}`,
  };
}

// ── บันทึกลง Google Sheet ─────────────────────────────────────────
async function saveToSheet(auth, data, slipUrl) {
  const sheets      = google.sheets({ version: 'v4', auth });
  const spreadsheetId = process.env.GOOGLE_SHEET_ID;

  // ตรวจว่ามี sheet ชื่อ Orders หรือยัง ถ้าไม่มีให้สร้าง
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const exists = meta.data.sheets.some(s => s.properties.title === SHEET_NAME);

  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{ addSheet: { properties: { title: SHEET_NAME } } }],
      },
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${SHEET_NAME}!A1:F1`,
      valueInputOption: 'RAW',
      requestBody: { values: [['วันที่/เวลา', 'ผู้สั่ง', 'รายการที่สั่ง', 'ราคา (฿)', 'หมายเหตุ', 'สลิป']] },
    });
  }

  const timestamp = new Date().toLocaleString('th-TH', {
    timeZone: 'Asia/Bangkok',
    day: '2-digit', month: '2-digit', year: 'numeric',
    hour: '2-digit', minute: '2-digit', second: '2-digit',
  });

  const slipCell = slipUrl ? `=HYPERLINK("${slipUrl}","ดูสลิป")` : '';

  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `${SHEET_NAME}!A:F`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[
        timestamp,
        data.customerName || '',
        data.items        || '',
        data.totalPrice   || 0,
        data.remark       || '',
        slipCell,
      ]],
    },
  });
}

// ── แจ้งเตือน LINE ────────────────────────────────────────────────
async function sendLineMessage(data, slipUrls) {
  const token    = process.env.LINE_CHANNEL_TOKEN;
  const targetId = process.env.LINE_TARGET_ID;
  if (!token || !targetId) return;

  const now = new Date().toLocaleTimeString('th-TH', {
    timeZone: 'Asia/Bangkok', hour: '2-digit', minute: '2-digit',
  });

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

  const messages = [{ type: 'text', text: textMsg }];
  if (slipUrls) {
    messages.push({
      type              : 'image',
      originalContentUrl: slipUrls.directUrl,
      previewImageUrl   : slipUrls.directUrl,
    });
  }

  await fetch('https://api.line.me/v2/bot/message/push', {
    method : 'POST',
    headers: {
      'Content-Type' : 'application/json',
      'Authorization': 'Bearer ' + token,
    },
    body: JSON.stringify({ to: targetId, messages }),
  });
}
