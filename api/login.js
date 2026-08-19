const { generateToken } = require('./session.js');

// Credentials from env vars (set in Vercel dashboard)
const ADMIN_USER = process.env.ADMIN_USER || 'MYCENTER';
const ADMIN_PASS = process.env.ADMIN_PASS || '9999';

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', req.headers.origin || '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Allow-Credentials', 'true');
  if (req.method === 'OPTIONS') return res.status(200).end();

  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Method not allowed' });

  const { username, password } = req.body || {};
  if (!username || !password) {
    return res.status(200).json({ ok: false, error: 'กรุณากรอกชื่อผู้ใช้และรหัสผ่าน' });
  }

  if (username === ADMIN_USER && password === ADMIN_PASS) {
    const token = generateToken({ username, iat: Date.now() });
    res.setHeader('Set-Cookie', `sid=${token}; Path=/; HttpOnly; SameSite=Lax; Max-Age=86400`);
    return res.status(200).json({ ok: true, token });
  }

  return res.status(200).json({ ok: false, error: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' });
};
