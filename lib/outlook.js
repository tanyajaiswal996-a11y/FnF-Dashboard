// Sends mail through Microsoft Graph using the app registration's
// client-credentials (app-only) flow. Mail.Send is an Application permission
// with admin consent granted, so this can send as any mailbox in the tenant
// via /users/{upn}/sendMail without per-user delegated consent.

const TENANT_ID = process.env.OUTLOOK_TENANT_ID;
const CLIENT_ID = process.env.OUTLOOK_CLIENT_ID;
const CLIENT_SECRET = process.env.OUTLOOK_CLIENT_SECRET;

let cachedToken = null; // { accessToken, expiresAt }

async function fetchToken() {
  const res = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials',
    }),
  });
  const body = await res.json();
  if (!body.access_token) {
    throw new Error(`Outlook token request failed: ${body.error_description || body.error || res.status}`);
  }
  cachedToken = { accessToken: body.access_token, expiresAt: Date.now() + (body.expires_in - 60) * 1000 };
  return cachedToken;
}

async function getToken() {
  if (cachedToken && cachedToken.expiresAt > Date.now()) return cachedToken;
  return fetchToken();
}

function toRecipients(list) {
  return (Array.isArray(list) ? list : [list]).filter(Boolean).map((address) => ({ emailAddress: { address } }));
}

// Sends an HTML email as `fromUpn`. `to` and `cc` may be a string or array of strings.
async function sendMail({ fromUpn, to, cc, subject, html }) {
  const token = await getToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(fromUpn)}/sendMail`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token.accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      message: {
        subject,
        body: { contentType: 'HTML', content: html },
        toRecipients: toRecipients(to),
        ccRecipients: toRecipients(cc),
      },
      saveToSentItems: true,
    }),
  });
  if (res.status !== 202) {
    const errBody = await res.text();
    throw new Error(`Outlook sendMail failed (${res.status}): ${errBody}`);
  }
}

module.exports = { sendMail };
