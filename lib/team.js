// Sign-in for the three people who use the dashboard directly:
//
//   admin     (TEAM_ADMIN_EMP_CODE)     HR — the whole main dashboard
//   accounts  (TEAM_ACCOUNTS_EMP_CODE)  Accounts team — enters the final TDS (/team)
//   pli       (TEAM_PLI_EMP_CODE)       PLI contact — enters the final PLI (/team)
//
// A person signs in with their official email and their employee code as the
// password. Who they are comes from PMS (the employee with that code), so no
// email addresses are kept here. The admin can instead set TEAM_ADMIN_PASSWORD
// to use a stronger password than the code. Each login has its own cookie and
// its own page: signing in as one never gives access to another. Sessions are
// signed cookies, so they work on every server instance. Because an employee
// code is a short number, a person is locked out for 15 minutes after 5 wrong
// tries.

const crypto = require('crypto');
const pms = require('./pms');
const kv = require('./kv');
const auth = require('./auth');

const COOKIES = { admin: 'fnf_admin', accounts: 'fnf_team', pli: 'fnf_team' };
const SESSION_SECONDS = 8 * 3600;
const MAX_FAILS = 5;
const LOCK_MS = 15 * 60 * 1000;
const IDENTITY_TTL_MS = 12 * 3600 * 1000;
const NS = process.env.KV_NAMESPACE || (process.env.VERCEL ? 'prod' : 'dev');

const roleCodes = () => ({
  admin: String(process.env.TEAM_ADMIN_EMP_CODE || '').trim(),
  accounts: String(process.env.TEAM_ACCOUNTS_EMP_CODE || '').trim(),
  pli: String(process.env.TEAM_PLI_EMP_CODE || '').trim(),
});

// When the admin login is configured, the main dashboard requires it.
const adminConfigured = () => !!roleCodes().admin;

// Who the employee with this code is, from PMS. A good answer is kept for 12h so a
// PMS hiccup (it sometimes answers "Forbidden") doesn't stop people signing in.
async function identityFor(role) {
  const code = roleCodes()[role];
  if (!code) return null;
  const key = `fnf:${NS}:team:${role}:${code}`;
  let cached = null;
  try { cached = await kv.get(key); } catch { /* fall through to PMS */ }
  if (cached && Date.now() - cached.at < IDENTITY_TTL_MS) return cached.identity;
  try {
    const rows = await pms.getEmployeeDetails(code);
    const e = rows.find((r) => r.email_address);
    if (e) {
      const identity = {
        empId: code,
        email: String(e.email_address).trim().toLowerCase(),
        name: [e.first_name, e.middle_name, e.last_name].filter(Boolean).join(' ').replace(/\s+/g, ' ').trim(),
      };
      try { await kv.set(key, { at: Date.now(), identity }); } catch { /* cache only */ }
      return identity;
    }
  } catch (err) {
    console.error(`Could not look up ${role} login ${code} in PMS:`, err.message);
  }
  return cached ? cached.identity : null; // stale is better than locking everyone out
}

const sameText = (a, b) => {
  const x = Buffer.from(String(a)), y = Buffer.from(String(b));
  return x.length === y.length && crypto.timingSafeEqual(x, y);
};

// The secret that proves it's them: their employee code (the admin may set a stronger password).
const secretFor = (role, who) => (role === 'admin' && process.env.TEAM_ADMIN_PASSWORD) || who.empId;

// `roles` limits which logins this page accepts, so the main dashboard's login
// accepts only the admin and /team accepts only Accounts and PLI.
// -> { ok: true, user } | { ok: false, locked?: true }
async function login(emailInput, passwordInput, roles) {
  const email = String(emailInput || '').trim().toLowerCase().slice(0, 120);
  const password = String(passwordInput || '').trim().slice(0, 60);
  if (!email || !password) return { ok: false };

  for (const role of roles) {
    const who = await identityFor(role);
    if (!who || who.email !== email) continue; // not this person

    // Only the real accounts can be guessed at, so only they are tracked.
    const lockKey = `fnf:${NS}:login:${email}`;
    const state = await kv.get(lockKey);
    const now = Date.now();
    if (state && state.fails >= MAX_FAILS && now - state.first < LOCK_MS) return { ok: false, locked: true };

    if (sameText(password, secretFor(role, who))) {
      if (state) await kv.del(lockKey);
      return { ok: true, user: { role, email: who.email, name: who.name, empId: who.empId } };
    }
    const fresh = !state || now - state.first >= LOCK_MS;
    await kv.set(lockKey, { fails: fresh ? 1 : state.fails + 1, first: fresh ? now : state.first });
    return { ok: false };
  }
  return { ok: false };
}

function startSession(res, user, secure) {
  const token = auth.sign(COOKIES[user.role], { role: user.role, email: user.email, name: user.name, empId: user.empId, exp: Math.floor(Date.now() / 1000) + SESSION_SECONDS });
  auth.addCookie(res, auth.cookieHeader(COOKIES[user.role], token, { maxAge: SESSION_SECONDS, secure }));
}

function endSession(res, secure, role) {
  auth.addCookie(res, auth.cookieHeader(COOKIES[role], '', { maxAge: 0, secure }));
}

// The signed-in person for one of the cookies, or null. The employee code is
// re-checked against the configured one on every request, so changing the
// setting signs people out.
function readSession(req, roles) {
  for (const cookie of new Set(roles.map((r) => COOKIES[r]))) {
    const s = auth.verify(cookie, auth.parseCookies(req)[cookie]);
    if (!s || !roles.includes(s.role) || COOKIES[s.role] !== cookie) continue;
    const code = roleCodes()[s.role];
    if (code && String(s.empId) === code) return { role: s.role, email: s.email, name: s.name, empId: s.empId };
  }
  return null;
}

const currentTeamUser = (req) => readSession(req, ['accounts', 'pli']);
const currentAdmin = (req) => { const s = readSession(req, ['admin']); return s ? { ...s, role: 'hr', admin: true } : null; };

module.exports = { adminConfigured, login, startSession, endSession, currentTeamUser, currentAdmin, identityFor };
