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

// ---- Thread tracking (initial FnF email -> acknowledgement / reminder / forward) ----

const GRAPH = 'https://graph.microsoft.com/v1.0';

async function graph(method, pathAndQuery, body, extraHeaders) {
  const token = await getToken();
  const res = await fetch(`${GRAPH}${pathAndQuery}`, {
    method,
    headers: { Authorization: `Bearer ${token.accessToken}`, 'Content-Type': 'application/json', ...(extraHeaders || {}) },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`Outlook ${method} ${pathAndQuery.split('?')[0]} failed (${res.status}): ${await res.text()}`);
  const text = await res.text();
  return text ? JSON.parse(text) : null;
}

const mb = (upn) => `/users/${encodeURIComponent(upn)}`;

// Sends a mail via draft-then-send so we learn its conversationId and
// internetMessageId (sendMail returns neither). That conversation is the
// unique reference for everything that follows.
async function sendMailTracked({ fromUpn, to, cc, subject, html }) {
  const draft = await graph('POST', `${mb(fromUpn)}/messages`, {
    subject,
    body: { contentType: 'HTML', content: html },
    toRecipients: toRecipients(to),
    ccRecipients: toRecipients(cc),
  });
  await graph('POST', `${mb(fromUpn)}/messages/${draft.id}/send`);
  return { conversationId: draft.conversationId, internetMessageId: draft.internetMessageId };
}

async function listConversation(mailboxUpn, conversationId) {
  const q = `$filter=${encodeURIComponent(`conversationId eq '${conversationId}'`)}&$top=50&$select=id,subject,from,receivedDateTime,sentDateTime,internetMessageId,parentFolderId`;
  const data = await graph('GET', `${mb(mailboxUpn)}/messages?${q}`);
  return (data && data.value) || [];
}

// The first message sent in the conversation = the initial FnF email, in Sent Items.
async function findRootMessage(mailboxUpn, conversationId) {
  const q = `$filter=${encodeURIComponent(`conversationId eq '${conversationId}'`)}&$top=25&$select=id,subject,sentDateTime,internetMessageId`;
  const data = await graph('GET', `${mb(mailboxUpn)}/mailFolders/sentitems/messages?${q}`);
  const items = ((data && data.value) || []).sort((a, b) => new Date(a.sentDateTime) - new Date(b.sentDateTime));
  return items[0] || null;
}

async function getUniqueBodyText(mailboxUpn, messageId) {
  const m = await graph('GET', `${mb(mailboxUpn)}/messages/${messageId}?$select=uniqueBody`, undefined, { Prefer: 'outlook.body-content-type="text"' });
  return (m && m.uniqueBody && m.uniqueBody.content) || '';
}

function withComment(draftHtml, commentHtml) {
  const html = draftHtml || '';
  return /<body[^>]*>/i.test(html) ? html.replace(/<body[^>]*>/i, (m) => m + commentHtml) : commentHtml + html;
}

// Replies to / forwards `messageId` (the initial email) so the original stays
// quoted underneath our text. Recipients are set explicitly, never inferred.
async function respondFromMessage(kind, { mailboxUpn, messageId, to, subject, commentHtml }) {
  const draft = await graph('POST', `${mb(mailboxUpn)}/messages/${messageId}/${kind === 'reply' ? 'createReply' : 'createForward'}`, {});
  const patch = {
    toRecipients: toRecipients(to),
    body: { contentType: 'HTML', content: withComment(draft.body && draft.body.content, commentHtml) },
  };
  if (subject) patch.subject = subject;
  await graph('PATCH', `${mb(mailboxUpn)}/messages/${draft.id}`, patch);
  await graph('POST', `${mb(mailboxUpn)}/messages/${draft.id}/send`);
  return { internetMessageId: draft.internetMessageId };
}

const replyInThread = (args) => respondFromMessage('reply', args);
const forwardMessage = (args) => respondFromMessage('forward', args);

module.exports = { sendMail, findLatestReply, sendMailTracked, listConversation, findRootMessage, getUniqueBodyText, replyInThread, forwardMessage };
