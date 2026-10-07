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

// Sends an HTML email as `fromUpn`. `to` and `cc` may be a string or array of
// strings. `attachments`, if given, is an array of { name, contentType, contentBase64 }.
async function sendMail({ fromUpn, to, cc, subject, html, attachments }) {
  const token = await getToken();
  const message = {
    subject,
    body: { contentType: 'HTML', content: html },
    toRecipients: toRecipients(to),
    ccRecipients: toRecipients(cc),
  };
  if (attachments && attachments.length) {
    message.attachments = attachments.map((a) => ({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: a.name,
      contentType: a.contentType,
      contentBytes: a.contentBase64,
    }));
  }
  const res = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(fromUpn)}/sendMail`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token.accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ message, saveToSentItems: true }),
  });
  if (res.status !== 202) {
    const errBody = await res.text();
    throw new Error(`Outlook sendMail failed (${res.status}): ${errBody}`);
  }
}

// Finds the newest reply (subject "RE: ... <subjectTag>") from `fromUpn` that
// reached `mailboxUpn`'s inbox after `since` (ISO). Only message headers are
// listed; the body of the single matching reply is fetched separately, using
// uniqueBody so the quoted original isn't included. Returns null if none.
async function findLatestReply({ mailboxUpn, since, subjectTag, fromUpns }) {
  const token = await getToken();
  const headers = { Authorization: `Bearer ${token.accessToken}` };
  const base = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailboxUpn)}`;
  const query = `$filter=${encodeURIComponent(`receivedDateTime ge ${since}`)}&$orderby=receivedDateTime desc&$top=50&$select=id,subject,receivedDateTime,from`;

  const listRes = await fetch(`${base}/mailFolders/inbox/messages?${query}`, { headers });
  if (!listRes.ok) throw new Error(`Outlook inbox read failed (${listRes.status}): ${await listRes.text()}`);
  const { value = [] } = await listRes.json();

  const tag = subjectTag.toLowerCase();
  const senders = (fromUpns || []).map((s) => s.trim().toLowerCase()).filter(Boolean);
  const match = value.find((m) => {
    const subject = (m.subject || '').toLowerCase();
    const from = ((m.from && m.from.emailAddress && m.from.emailAddress.address) || '').toLowerCase();
    return /^\s*re:/.test(subject) && subject.includes(tag) && senders.includes(from);
  });
  if (!match) return null;

  const bodyRes = await fetch(`${base}/messages/${match.id}?$select=uniqueBody`, {
    headers: { ...headers, Prefer: 'outlook.body-content-type="text"' },
  });
  if (!bodyRes.ok) throw new Error(`Outlook message read failed (${bodyRes.status}): ${await bodyRes.text()}`);
  const body = await bodyRes.json();
  return {
    id: match.id,
    receivedAt: match.receivedDateTime,
    text: (body.uniqueBody && body.uniqueBody.content) || '',
  };
}

module.exports = { sendMail, findLatestReply };
