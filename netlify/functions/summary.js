const { schedule } = require('@netlify/functions');

// รันทุกวัน 01:01 เวลาไทย (18:01 UTC) → เรียก Apps Script ให้ส่งสรุป
exports.handler = schedule('1 18 * * *', async () => {
  try {
    await fetch(`${process.env.APPS_SCRIPT_URL}?action=summary`, {
      redirect: 'follow',
    });
    console.log('summary triggered');
  } catch (err) {
    console.error('summary error:', err);
  }
});
