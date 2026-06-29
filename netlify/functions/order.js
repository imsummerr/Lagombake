// Proxy: รับออเดอร์จากเว็บ → ส่งต่อไป Google Apps Script
exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method Not Allowed' };
  }

  try {
    await fetch(process.env.APPS_SCRIPT_URL, {
      method  : 'POST',
      headers : { 'Content-Type': 'application/json' },
      body    : event.body,
      redirect: 'follow',
    });

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
