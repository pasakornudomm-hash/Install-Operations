const crypto = require('crypto');

// In-memory session store (resets on cold start, but that's fine for Vercel)
const sessions = new Map();

function generateToken() {
  return crypto.randomBytes(24).toString('hex');
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

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', req.headers.origin || '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type,Authorization');
  res.setHeader('Access-Control-Allow-Credentials', 'true');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const token = getSession(req);
  if (token && sessions.has(token)) {
    return res.status(200).json({ authenticated: true, user: sessions.get(token) });
  }
  return res.status(200).json({ authenticated: false });
};

// Export for login.js to use
module.exports.sessions = sessions;
module.exports.generateToken = generateToken;
