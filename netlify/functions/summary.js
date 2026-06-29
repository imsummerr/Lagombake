const { schedule } = require('@netlify/functions');
const { google }   = require('googleapis');

const SHEET_NAME = 'Orders';

// ── รันทุกวัน 01:01 เวลาไทย (18:01 UTC วันก่อน) ─────────────────
// ตัดยอดออเดอร์ถึง 23:59 ของวันก่อน แล้วส่งสรุปตี 1 นาที 1
exports.handler = schedule('1 18 * * *', async () => {
  try {
    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT),
      scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
    });

    const rows      = await getYesterdayRows(auth);
    const summary   = buildSummary(rows);
    const dateLabel = getYesterdayLabel();

    await sendLineSummary(summary, dateLabel, rows.length);
    console.log('summary sent:', dateLabel, 'orders:', rows.length);
  } catch (err) {
    console.error('summary error:', err);
  }
});

// ── ดึงแถวของวันนี้จาก Sheet ─────────────────────────────────────
async function getYesterdayRows(auth) {
  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.GOOGLE_SHEET_ID,
    range: `${SHEET_NAME}!A2:E`,  // A=วันที่ B=ชื่อ C=รายการ D=ราคา E=หมายเหตุ
  });

  const rows = res.data.values || [];
  const dateStr = getYesterdayDateStr();  // "dd/MM/yyyy" ของเมื่อวาน

  return rows.filter(row => {
    const dateCell = row[0] || '';
    return dateCell.startsWith(dateStr);
  });
}

// ── parse และรวมยอดเมนู ───────────────────────────────────────────
function buildSummary(rows) {
  // map: "ชาไทย (ปกติ)" → total qty
  const counts = new Map();

  rows.forEach(row => {
    const itemsStr = row[2] || '';
    itemsStr.split(' | ').forEach(segment => {
      const seg = segment.trim();
      if (!seg) return;

      // ดึง x{n} ถ้ามี  เช่น "ชาไทย x2 (ปกติ)"  หรือ "แซนด์วิช x3 [...]"
      const qMatch = seg.match(/ x(\d+)/);
      const qty    = qMatch ? parseInt(qMatch[1]) : 1;
      const key    = seg.replace(/ x\d+/, '').trim();   // ตัด x{n} ออก

      counts.set(key, (counts.get(key) || 0) + qty);
    });
  });

  return counts;
}

// ── ส่ง LINE ─────────────────────────────────────────────────────
async function sendLineSummary(summary, dateLabel, orderCount) {
  const token    = process.env.LINE_CHANNEL_TOKEN;
  const targetId = process.env.LINE_TARGET_ID;
  if (!token || !targetId) return;

  if (summary.size === 0) {
    const msg = `📋 สรุปออเดอร์ ${dateLabel}\n━━━━━━━━━━━━━━━━━━\nไม่มีออเดอร์วันนี้`;
    await pushLine(token, targetId, msg);
    return;
  }

  // แยกหมวดหมู่
  const drinks = [];
  const foods  = [];
  const sweets = [];

  summary.forEach((qty, key) => {
    const line = `  • ${key}  →  ${qty} ${unitOf(key)}`;
    if (isFood(key))       foods.push(line);
    else if (isSweet(key)) sweets.push(line);
    else                   drinks.push(line);
  });

  let body = '';
  if (drinks.length) body += '☕ เครื่องดื่ม\n' + drinks.join('\n') + '\n';
  if (foods.length)  body += '\n🥪 แซนด์วิช / ปังปิ้ง\n' + foods.join('\n') + '\n';
  if (sweets.length) body += '\n🍞 ของหวานปังปิ้ง\n' + sweets.join('\n') + '\n';

  const totalQty = [...summary.values()].reduce((s, n) => s + n, 0);

  const msg =
    '━━━━━━━━━━━━━━━━━━\n' +
    '📊 สรุปออเดอร์ประจำวัน\n' +
    `📅 ${dateLabel}\n` +
    '━━━━━━━━━━━━━━━━━━\n' +
    body +
    '━━━━━━━━━━━━━━━━━━\n' +
    `📦 ออเดอร์ : ${orderCount} รายการ\n` +
    `🔢 รวมทั้งหมด : ${totalQty} ชิ้น/แก้ว`;

  await pushLine(token, targetId, msg);
}

async function pushLine(token, targetId, text) {
  await fetch('https://api.line.me/v2/bot/message/push', {
    method : 'POST',
    headers: {
      'Content-Type' : 'application/json',
      'Authorization': 'Bearer ' + token,
    },
    body: JSON.stringify({
      to      : targetId,
      messages: [{ type: 'text', text }],
    }),
  });
}

// ── helpers ──────────────────────────────────────────────────────
// ฟังก์ชันรันตี 1:01 ของวันถัดไป → ต้องสรุปออเดอร์ "เมื่อวาน" (เวลาไทย)
function getThaiYesterday() {
  const now = new Date(new Date().toLocaleString('en-US', { timeZone: 'Asia/Bangkok' }));
  now.setDate(now.getDate() - 1);
  return now;
}

function getYesterdayDateStr() {
  // คืน "dd/MM/yyyy" ของเมื่อวาน เวลาไทย
  const d  = getThaiYesterday();
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  return `${dd}/${mm}/${d.getFullYear()}`;
}

function getYesterdayLabel() {
  return getThaiYesterday().toLocaleDateString('th-TH', {
    day: 'numeric', month: 'long', year: 'numeric',
  });
}

const FOOD_KEYWORDS  = ['แซนด์วิช', 'ปังปิ้ง'];
const SWEET_KEYWORDS = ['ปังปิ้งโอริโอ้','ปังปิ้งช็อก','ปังปิ้งเนย','ปังปิ้งนูเทลล่า',
                        'ปังปิ้งน้ำผึ้ง','ปังปิ้งน้ำตาล','ปังปิ้งกระเทียม','ปังปิ้งกล้วย',
                        'ปังปิ้งอโวคาโด้','ปังปิ้งโกโก้'];

function isSweet(key) { return SWEET_KEYWORDS.some(k => key.startsWith(k)); }
function isFood(key)  { return !isSweet(key) && FOOD_KEYWORDS.some(k => key.startsWith(k)); }

function unitOf(key) {
  if (isFood(key))  return 'ชิ้น';
  if (isSweet(key)) return 'ชิ้น';
  return 'แก้ว';
}
