const { sessions } = require('./session.js');

// Google Apps Script deployment URL
const GAS_URL = process.env.GAS_URL || 'https://script.google.com/macros/s/AKfycbzahp6g2Fkbet8ivCUVNIiVBgyJgnMA1jHrzpSpqgJ9UGlMeTfmkXubCOP22uMt257T/exec';

function getSession(req) {
  const cookie = (req.headers.cookie || '').split(';').map(c => c.trim());
  for (const c of cookie) {
    if (c.startsWith('sid=')) return c.slice(4);
  }
  const auth = req.headers.authorization;
  if (auth && auth.startsWith('Bearer ')) return auth.slice(7);
  return null;
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', req.headers.origin || '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type,Authorization');
  res.setHeader('Access-Control-Allow-Credentials', 'true');
  if (req.method === 'OPTIONS') return res.status(200).end();

  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Method not allowed' });

  // Check authentication
  const token = getSession(req);
  if (!token || !sessions.has(token)) {
    return res.status(401).json({ ok: false, error: 'กรุณายืนยัน OTP ก่อนเข้าใช้งาน' });
  }

  try {
    const gasRes = await fetch(GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(req.body),
      redirect: 'follow',
    });
    const data = await gasRes.json();
    return res.status(200).json(data);
  } catch (err) {
    return res.status(502).json({ ok: false, error: 'เรียก Apps Script ไม่สำเร็จ: ' + err.message });
  }
};
