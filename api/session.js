const crypto = require('crypto');

const SECRET = process.env.SESSION_SECRET || 'mycenter-kpi-secret-2026';

function generateToken(payload) {
  const data = Buffer.from(JSON.stringify(payload)).toString('base64url');
  const sig = crypto.createHmac('sha256', SECRET).update(data).digest('base64url');
  return `${data}.${sig}`;
}

function verifyToken(token) {
  try {
    const [data, sig] = token.split('.');
    if (!data || !sig) return null;
    const expected = crypto.createHmac('sha256', SECRET).update(data).digest('base64url');
    if (sig !== expected) return null;
    return JSON.parse(Buffer.from(data, 'base64url').toString());
  } catch (_) { return null; }
}

function getSession(req) {
  const cookie = (req.headers.cookie || '').split(';').map(c => c.trim());
  for (const c of cookie) {
    if (c.startsWith('sid=')) return c.slice(4);
  }
  const auth = req.headers.authorization;
  if (auth && auth.startsWith('Bearer ')) return auth.slice(7);
  return null;
}

module.exports = { generateToken, verifyToken, getSession };
