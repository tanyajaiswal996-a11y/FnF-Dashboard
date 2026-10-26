// Sign-in for the dashboard.
//
// Microsoft (Entra ID) sign-in using the company's existing app registration;
// the session is a signed cookie, so it needs no server memory and works the
// same on every server instance. Who may sign in, and as what, comes from three
// environment lists (comma-separated work emails):
//
//   AUTH_HR_EMAILS        full access (defaults to the app's sender mailbox)
//   AUTH_ACCOUNTS_EMAILS  Accounts team: TDS-applicable employees + their sheet
//   AUTH_PLI_EMAILS       PLI contact: Sales employees, no salary data
//
// Sign-in is only enforced when AUTH_REQUIRED=1, so nothing changes (and nobody
// is locked out) until Microsoft sign-in has been set up for the app.

const crypto = require('crypto');

const COOKIE = 'fnf_session';
const OAUTH_COOKIE = 'fnf_oauth';
const SESSION_SECONDS = 8 * 3600;

const authRequired = () => process.env.AUTH_REQUIRED === '1';

// ---- signed cookies ----
function key(purpose) {
  const base = process.env.SESSION_SECRET || process.env.LINK_SIGNING_SECRET || process.env.OUTLOOK_CLIENT_SECRET;
  if (!base) throw new Error('SESSION_SECRET is not configured');
  return crypto.createHash('sha256').update(`${purpose}:${base}`).digest();
}

const b64 = (buf) => Buffer.from(buf).toString('base64url');

function sign(purpose, payload) {
  const body = b64(JSON.stringify(payload));
  const mac = b64(crypto.createHmac('sha256', key(purpose)).update(body).digest());
  return `${body}.${mac}`;
}

function verify(purpose, token) {
  if (!token || !token.includes('.')) return null;
  const [body, mac] = token.split('.');
  const expected = crypto.createHmac('sha256', key(purpose)).update(body).digest();
  let given;
  try { given = Buffer.from(mac, 'base64url'); } catch { return null; }
  if (given.length !== expected.length || !crypto.timingSafeEqual(given, expected)) return null;
  try {
    const payload = JSON.parse(Buffer.from(body, 'base64url').toString('utf8'));
    return payload.exp && payload.exp > Date.now() / 1000 ? payload : null;
  } catch { return null; }
}

function parseCookies(req) {
  const out = {};
  for (const part of String(req.headers.cookie || '').split(';')) {
    const i = part.indexOf('=');
    if (i > 0) out[part.slice(0, i).trim()] = decodeURIComponent(part.slice(i + 1).trim());
  }
  return out;
}

function cookieHeader(name, value, { maxAge, secure }) {
  return `${name}=${encodeURIComponent(value)}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${maxAge}${secure ? '; Secure' : ''}`;
}

function addCookie(res, header) {
  const existing = res.getHeader('Set-Cookie');
  res.setHeader('Set-Cookie', existing ? [].concat(existing, header) : header);
}

// ---- roles ----
const list = (v) => new Set(String(v || '').split(',').map((s) => s.trim().toLowerCase()).filter(Boolean));

function roleFor(email) {
  const e = String(email || '').trim().toLowerCase();
  if (!e) return null;
  const hr = list(process.env.AUTH_HR_EMAILS || process.env.OUTLOOK_SENDER_UPN);
  if (hr.has(e)) return 'hr';
  if (list(process.env.AUTH_ACCOUNTS_EMAILS).has(e)) return 'accounts';
  if (list(process.env.AUTH_PLI_EMAILS).has(e)) return 'pli';
  return null;
}

// dev=true marks a test login (see /auth/dev); honoured only when DEV_LOGIN=1 on a machine that is not Vercel.
const devLoginAllowed = () => process.env.DEV_LOGIN === '1' && !process.env.VERCEL;

function startSession(res, { email, name, role, dev }, secure) {
  const token = sign('session', { email, name, role, dev: !!dev, exp: Math.floor(Date.now() / 1000) + SESSION_SECONDS });
  addCookie(res, cookieHeader(COOKIE, token, { maxAge: SESSION_SECONDS, secure }));
}

function endSession(res, secure) {
  addCookie(res, cookieHeader(COOKIE, '', { maxAge: 0, secure }));
}

// The signed-in user, or null. With sign-in switched off everyone is treated as HR.
function currentUser(req) {
  if (!authRequired()) return { email: '', name: 'HR (sign-in is off)', role: 'hr', signInOff: true };
  const s = verify('session', parseCookies(req)[COOKIE]);
  if (!s) return null;
  if (s.dev) return devLoginAllowed() ? { email: s.email, name: s.name || s.email, role: s.role, dev: true } : null;
  const role = roleFor(s.email); // re-checked every request, so removing someone from the list takes effect at once
  return role ? { email: s.email, name: s.name || s.email, role } : null;
}

// ---- Microsoft sign-in (authorization code flow) ----
const tenant = () => process.env.OUTLOOK_TENANT_ID;
const clientId = () => process.env.OUTLOOK_CLIENT_ID;

function beginLogin(res, redirectUri, secure) {
  const state = crypto.randomBytes(16).toString('hex');
  const nonce = crypto.randomBytes(16).toString('hex');
  addCookie(res, cookieHeader(OAUTH_COOKIE, sign('oauth', { state, nonce, exp: Math.floor(Date.now() / 1000) + 600 }), { maxAge: 600, secure }));
  const q = new URLSearchParams({
    client_id: clientId(), response_type: 'code', redirect_uri: redirectUri, response_mode: 'query',
    scope: 'openid profile email', state, nonce, prompt: 'select_account',
  });
  return `https://login.microsoftonline.com/${tenant()}/oauth2/v2.0/authorize?${q}`;
}

// Checks the claims of an ID token received straight from Microsoft's token
// endpoint (over TLS, using our client secret — so no separate signature check
// is needed). Throws if anything is off; returns who signed in.
function readIdToken(idToken, { tenantId, audience, nonce, nowSeconds }) {
  const parts = String(idToken || '').split('.');
  if (parts.length !== 3) throw new Error('Malformed ID token');
  const c = JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8'));
  const now = nowSeconds == null ? Math.floor(Date.now() / 1000) : nowSeconds;
  if (c.aud !== audience) throw new Error('ID token is for a different app');
  if (c.tid !== tenantId) throw new Error('Sign-in is from a different organisation');
  if (c.iss !== `https://login.microsoftonline.com/${tenantId}/v2.0`) throw new Error('Unexpected token issuer');
  if (!c.exp || c.exp < now - 60) throw new Error('ID token has expired');
  if (c.nonce !== nonce) throw new Error('Sign-in response did not match the request');
  const email = c.preferred_username || c.email || c.upn;
  if (!email) throw new Error('No email address in the sign-in response');
  return { email, name: c.name || email };
}

async function finishLogin(req, query, redirectUri) {
  const saved = verify('oauth', parseCookies(req)[OAUTH_COOKIE]);
  if (!saved || !query.get('state') || query.get('state') !== saved.state) throw new Error('Sign-in session expired. Please try again.');
  if (query.get('error')) throw new Error(query.get('error_description') || query.get('error'));
  const code = query.get('code');
  if (!code) throw new Error('No sign-in code was returned.');

  const res = await fetch(`https://login.microsoftonline.com/${tenant()}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: clientId(), client_secret: process.env.OUTLOOK_CLIENT_SECRET, grant_type: 'authorization_code',
      code, redirect_uri: redirectUri, scope: 'openid profile email',
    }),
  });
  const body = await res.json();
  if (!body.id_token) throw new Error(body.error_description || 'Microsoft did not return a sign-in token.');
  return readIdToken(body.id_token, { tenantId: tenant(), audience: clientId(), nonce: saved.nonce });
}

module.exports = {
  authRequired, devLoginAllowed, currentUser, roleFor, startSession, endSession, beginLogin, finishLogin, readIdToken,
  sign, verify, parseCookies, cookieHeader, addCookie,
};
