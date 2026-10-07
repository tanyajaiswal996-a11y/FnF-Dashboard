try {
  process.loadEnvFile(require('path').join(__dirname, '.env')); // Node 20.6+
} catch {
  // no .env present — fine while running on mock data
}

const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { URL } = require('url');
const store = require('./lib/store');
const { calc, recStatus, recSla, settlementCalc, settlementMonthLabels, computeTdsCalc, isSalesEmployee, pliShareWindow } = require('./lib/calc');
const pms = require('./lib/pms');
const outlook = require('./lib/outlook');
const { parseTdsReply } = require('./lib/tds-reply');
const { classifyAck } = require('./lib/ack-reply');

const SENDER_UPN = process.env.OUTLOOK_SENDER_UPN;
const PAYMENT_CONTACT_UPN = process.env.OUTLOOK_PAYMENT_CONTACT_UPN;

const PORT = 4174;
const ROOT = __dirname;

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript',
  '.css': 'text/css',
  '.json': 'application/json',
};

function sendJson(res, status, body) {
  res.writeHead(status, { 'Content-Type': 'application/json' });
  res.end(JSON.stringify(body));
}

async function readBody(req) {
  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);
  const raw = Buffer.concat(chunks).toString('utf8');
  return raw ? JSON.parse(raw) : {};
}

function withDerived(rec) {
  return Object.assign({}, rec, {
    computed: calc(rec),
    status: recStatus(rec),
  });
}

function fmtInr(n) {
  const v = Math.round((Number(n) || 0) * 100) / 100;
  return '₹' + Math.abs(v).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function approvalEmailHtml(rec, baseUrl) {
  const reviewUrl = `${baseUrl}/api/records/${rec.id}/review?token=${rec.sheetApproval.reviewToken}`;
  return `
  <div style="font-family:Segoe UI,Arial,sans-serif;color:#0a1f3a;max-width:600px;">
    <p>Dear ${rec.name || 'Employee'},</p>
    <p>Your tentative Full &amp; Final settlement sheet is ready. Please review the detailed breakdown and Approve or Reject it.</p>
    <div style="margin:24px 0;">
      <a href="${reviewUrl}" style="background:#003b73;color:#fff;padding:12px 24px;border-radius:5px;text-decoration:none;font-weight:600;">Review FnF Sheet →</a>
    </div>
    <p style="font-size:12px;color:#4c6480;">If you did not expect this email, please contact HR Operations directly.</p>
  </div>`;
}

// "RT - Name - EmpCode" when the employee is a Rayontara employee, else just "Name - EmpCode".
function employeeSubjectTag(rec) {
  const parts = [];
  if (rec.isRayontara) parts.push('RT');
  parts.push(rec.name || '');
  parts.push(rec.empId || '');
  return parts.join(' - ');
}

function settlementRow(label, amount) {
  return `<tr><td style="padding:6px 8px;border-bottom:1px dashed #cfe0f0;">${label}</td><td style="padding:6px 8px;border-bottom:1px dashed #cfe0f0;text-align:right;font-family:monospace;">${fmtInr(amount)}</td></tr>`;
}

// Renders the exact HR FnF sheet template: header block + two-column earnings/deductions table.
function settlementSheetHtml(rec) {
  const s = rec.settlement || {};
  const e = s.earnings || {};
  const d = s.deductions || {};
  const months = settlementMonthLabels(rec.lwd);
  const totals = settlementCalc(rec);

  const earningsRows = [
    settlementRow(`Salary (${months.current})`, e.salaryCurrentMonth),
    settlementRow(`Salary (${months.previous})`, e.salaryPreviousMonth),
    settlementRow('Incentives', e.incentives),
    settlementRow('TA Bill (If figure is in Positive)', e.taBillPositive),
    settlementRow('Gratuity', e.gratuity),
    settlementRow('Any other Reimbursement', e.otherReimbursement),
    settlementRow('Notice Period Pay', e.noticePeriodPay),
  ].join('');

  const deductionRows = [
    settlementRow(`PF (${months.current})`, d.pfCurrentMonth),
    settlementRow(`PF (${months.previous})`, d.pfPreviousMonth),
    settlementRow('Exam Dues', d.examDues),
    settlementRow('WFH Dues', d.wfhDues),
    settlementRow('NPS', d.nps),
    settlementRow('VPF', d.vpf),
    settlementRow('TA Bill (If figure is negative after settlements)', d.taBillNegative),
    settlementRow('Loan & Advances', d.loanAdvances),
    settlementRow('Personal rides deduction', d.personalRides),
    settlementRow('Any other deductions (Pluxee Meal)', d.pluxeeMeal),
    settlementRow('ESI', d.esi),
    settlementRow('TDS', d.tds),
    settlementRow('PT', d.pt),
  ].join('');

  return `
  <table style="width:100%;border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:13px;margin-bottom:20px;">
    <tr>
      <td style="font-weight:600;padding:6px 8px;border:1px solid #cfe0f0;">Emp Code</td><td style="padding:6px 8px;border:1px solid #cfe0f0;">${rec.empId||''}</td>
      <td style="font-weight:600;padding:6px 8px;border:1px solid #cfe0f0;">Employee Name</td><td style="padding:6px 8px;border:1px solid #cfe0f0;">${rec.name||''}</td>
    </tr>
    <tr>
      <td style="font-weight:600;padding:6px 8px;border:1px solid #cfe0f0;">DOJ</td><td style="padding:6px 8px;border:1px solid #cfe0f0;">${s.doj||''}</td>
      <td style="font-weight:600;padding:6px 8px;border:1px solid #cfe0f0;">Designation</td><td style="padding:6px 8px;border:1px solid #cfe0f0;">${rec.designation||''}</td>
    </tr>
    <tr>
      <td style="font-weight:600;padding:6px 8px;border:1px solid #cfe0f0;">Email Id</td><td style="padding:6px 8px;border:1px solid #cfe0f0;">${rec.contact||''}</td>
      <td style="font-weight:600;padding:6px 8px;border:1px solid #cfe0f0;">DOR</td><td style="padding:6px 8px;border:1px solid #cfe0f0;">${rec.dor||''}</td>
    </tr>
    <tr>
      <td colspan="2" style="border:1px solid #cfe0f0;"></td>
      <td style="font-weight:600;padding:6px 8px;border:1px solid #cfe0f0;">LWD</td><td style="padding:6px 8px;border:1px solid #cfe0f0;">${rec.lwd||''}</td>
    </tr>
  </table>
  <table style="width:100%;border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:13px;">
    <tr>
      <th style="text-align:left;padding:6px 8px;border-bottom:2px solid #0a1f3a;">Addition</th><th style="text-align:right;padding:6px 8px;border-bottom:2px solid #0a1f3a;">Amount</th>
      <th style="text-align:left;padding:6px 8px;border-bottom:2px solid #0a1f3a;">Deductions</th><th style="text-align:right;padding:6px 8px;border-bottom:2px solid #0a1f3a;">Amount</th>
    </tr>
  </table>
  <table style="width:100%;border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:13px;">
    <tr><td style="width:50%;vertical-align:top;padding:0;"><table style="width:100%;border-collapse:collapse;">${earningsRows}</table></td>
        <td style="width:50%;vertical-align:top;padding:0;"><table style="width:100%;border-collapse:collapse;">${deductionRows}</table></td></tr>
  </table>
  <table style="width:100%;border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:13px;margin-top:8px;">
    <tr style="font-weight:600;">
      <td style="padding:8px;border-top:2px solid #0a1f3a;">Grand Total Earnings</td><td style="padding:8px;border-top:2px solid #0a1f3a;text-align:right;font-family:monospace;">${fmtInr(totals.earningsTotal)}</td>
      <td style="padding:8px;border-top:2px solid #0a1f3a;">Total Deductions</td><td style="padding:8px;border-top:2px solid #0a1f3a;text-align:right;font-family:monospace;">${fmtInr(totals.deductionsTotal)}</td>
    </tr>
    <tr style="font-weight:700;background:#fff9d6;">
      <td colspan="3" style="padding:10px;">${totals.netPayable >= 0 ? 'Total Amount to be Paid' : 'Total Amount to be Recovered'}</td><td style="padding:10px;text-align:right;font-family:monospace;">${fmtInr(totals.netPayable)}</td>
    </tr>
  </table>`;
}

function reviewPageHtml(rec) {
  const sheetHtml = settlementSheetHtml(rec);
  const token = rec.sheetApproval.reviewToken;
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>FnF Sheet Review</title></head>
  <body style="font-family:Segoe UI,Arial,sans-serif;max-width:700px;margin:40px auto;color:#0a1f3a;">
    <h2>Full &amp; Final Settlement — ${rec.name||''}</h2>
    <p>Please review the detailed sheet below before responding.</p>
    ${sheetHtml}
    <div style="display:flex;gap:24px;margin-top:32px;flex-wrap:wrap;">
      <form method="GET" action="/api/records/${rec.id}/respond">
        <input type="hidden" name="action" value="approve">
        <input type="hidden" name="token" value="${token}">
        <button type="submit" style="background:#0f9d58;color:#fff;padding:12px 24px;border:none;border-radius:5px;font-weight:600;font-size:14px;cursor:pointer;">Approve</button>
      </form>
      <form method="GET" action="/api/records/${rec.id}/respond" style="flex:1;min-width:260px;">
        <input type="hidden" name="action" value="reject">
        <input type="hidden" name="token" value="${token}">
        <label style="display:block;font-size:13px;margin-bottom:6px;">Reason for rejection (required)</label>
        <textarea name="reason" required rows="2" style="width:100%;padding:8px;border:1px solid #cfe0f0;border-radius:5px;font-family:inherit;"></textarea>
        <button type="submit" style="margin-top:8px;background:#d64550;color:#fff;padding:12px 24px;border:none;border-radius:5px;font-weight:600;font-size:14px;cursor:pointer;">Reject</button>
      </form>
    </div>
  </body></html>`;
}

// ---- PLI / incentive confirmation (Sales only) ----
// The PLI contact only ever sees identity + dates for the employee — never
// salary or any settlement figure — and answers with a single amount.
const PUBLIC_BASE_URL = process.env.PUBLIC_BASE_URL || `http://localhost:${PORT}`;

function pliContacts() {
  return (process.env.PLI_CONTACT_UPN || '').split(',').map((s) => s.trim()).filter(Boolean);
}

function pliLink(rec) {
  return `${PUBLIC_BASE_URL}/api/pli/${rec.id}?token=${rec.pli.token}`;
}

function pliDetailsTable(rec) {
  const cell = 'padding:6px 10px;border:1px solid #cfe0f0;';
  const row = (k, v) => `<tr><td style="${cell}font-weight:600;">${k}</td><td style="${cell}">${v || '—'}</td></tr>`;
  return `<table style="border-collapse:collapse;font-size:13px;">
    ${row('Employee Code', rec.empId)}${row('Employee Name', rec.name)}${row('Designation', rec.designation)}
    ${row('Department', rec.department)}${row('Date of Resignation', rec.dor)}${row('Last Working Day', rec.lwd)}
  </table>`;
}

function pliEmailHtml(rec) {
  return `
  <div style="font-family:Segoe UI,Arial,sans-serif;color:#0a1f3a;max-width:600px;">
    <p>Hi Monika,</p>
    <p>The employee below is leaving shortly. Please confirm their final PLI / incentive amount so it can be included in the Full &amp; Final settlement.</p>
    ${pliDetailsTable(rec)}
    <div style="margin:24px 0;">
      <a href="${pliLink(rec)}" style="background:#003b73;color:#fff;padding:12px 24px;border-radius:5px;text-decoration:none;font-weight:600;">Enter PLI Amount →</a>
    </div>
    <p style="font-size:13px;">Enter <b>0</b> (or NIL) if there is no PLI payable. You can reopen the link to update the amount until the settlement is closed.</p>
  </div>`;
}

function pliPageHtml(rec) {
  const submitted = rec.pli.status === 'submitted';
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>PLI Amount</title></head>
  <body style="font-family:Segoe UI,Arial,sans-serif;max-width:560px;margin:40px auto;padding:0 16px;color:#0a1f3a;">
    <h2>PLI / Incentive Amount</h2>
    <p style="color:#4c6480;">Please confirm the final PLI / incentive payable to this employee.</p>
    ${pliDetailsTable(rec)}
    ${submitted ? `<p style="margin-top:18px;padding:10px 12px;background:#eef8f1;border-radius:6px;">Currently recorded: <b>${rec.pli.amount === 0 ? 'NIL (0)' : fmtInr(rec.pli.amount)}</b>. You can update it below.</p>` : ''}
    <form id="f" style="margin-top:20px;">
      <label style="display:block;font-weight:600;margin-bottom:6px;">Final PLI amount (₹) — enter 0 or NIL if none</label>
      <input id="amt" required autocomplete="off" value="${submitted ? rec.pli.amount : ''}" style="width:100%;box-sizing:border-box;padding:10px;border:1px solid #cfe0f0;border-radius:5px;font-size:15px;">
      <button type="submit" style="margin-top:14px;background:#0f9d58;color:#fff;padding:12px 24px;border:none;border-radius:5px;font-weight:600;font-size:14px;cursor:pointer;">${submitted ? 'Update amount' : 'Submit amount'}</button>
      <p id="msg" style="margin-top:12px;font-weight:600;"></p>
    </form>
    <script>
      document.getElementById('f').addEventListener('submit', async (e) => {
        e.preventDefault();
        const msg = document.getElementById('msg');
        msg.style.color = '#4c6480'; msg.textContent = 'Saving…';
        const res = await fetch(location.pathname, { method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ token: new URLSearchParams(location.search).get('token'), amount: document.getElementById('amt').value }) });
        const data = await res.json();
        msg.style.color = res.ok ? '#0f9d58' : '#d64550';
        msg.textContent = res.ok ? 'Saved — recorded as ' + (data.amount === 0 ? 'NIL (0)' : '₹' + data.amount.toLocaleString('en-IN')) + '. Thank you.' : (data.error || 'Could not save.');
      });
    </script>
  </body></html>`;
}

// "NIL"/"0" -> 0, otherwise a plain non-negative amount. Returns null if unreadable.
function parsePliAmount(raw) {
  const s = String(raw == null ? '' : raw).trim().replace(/[₹,\s]/g, '').replace(/^rs\.?/i, '').replace(/\/-$/, '');
  if (/^(nil|none)$/i.test(s)) return 0;
  if (!/^\d+(\.\d+)?$/.test(s)) return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

async function sharePli(rec) {
  const to = pliContacts();
  if (!to.length) throw new Error('PLI_CONTACT_UPN is not configured');
  if (!rec.pli.token) rec.pli.token = crypto.randomBytes(24).toString('hex');
  await outlook.sendMail({
    fromUpn: SENDER_UPN,
    to,
    subject: `PLI amount needed: ${rec.name || ''} (${rec.empId || ''}) — LWD ${rec.lwd || ''}`,
    html: pliEmailHtml(rec),
  });
  await store.updateRecord(rec.id, {
    pli: { token: rec.pli.token, sharedAt: new Date().toISOString(), status: rec.pli.status === 'submitted' ? 'submitted' : 'shared' },
  });
  return to;
}

// Shares every Sales employee whose LWD is within the next 3 days (and not yet
// past), once. Anyone else is shared manually from the PLI tab.
async function runPliSharing() {
  if (!pliContacts().length) return;
  const today = new Date().toISOString().slice(0, 10);
  try {
    for (const rec of await store.listRecords()) {
      if (!isSalesEmployee(rec) || rec.pli.status !== 'not_shared') continue;
      if (!pliShareWindow(rec, today).windowOpen) continue;
      await sharePli(rec);
      console.log(`PLI details shared for ${rec.name} (${rec.empId})`);
    }
  } catch (err) {
    console.error('PLI auto-share failed:', err.message);
  }
}

// ---- Acknowledgement tracking on the initial FnF email thread ----
// Everything keys off the conversation of the initial email: the employee's
// clear "accepted/acknowledged" reply in that thread, one reminder sent as a
// reply in that same thread after ACK_REMINDER_HOURS, and one forward of that
// initial email to the payment contact (Bhargavi). Other threads never trigger.
const ACK_REMINDER_HOURS = Number(process.env.ACK_REMINDER_HOURS) || 24;
const ackBusy = new Set(); // guards against two overlapping checks sending duplicates

function ackReminderHtml(rec) {
  return `<div style="font-family:Segoe UI,Arial,sans-serif;color:#0a1f3a;">
    <p>Dear ${rec.name || 'Employee'},</p>
    <p>This is a gentle reminder about your Full &amp; Final settlement sheet shared below. Please review it and reply to this email confirming your acknowledgement.</p>
    <p>Regards,<br/>HR Operations</p>
  </div><hr/>`;
}

// Hands the settlement to the payment contact exactly once — as a forward of the
// initial FnF email (original stays quoted underneath), with Payment Done / Not
// Done buttons. Used for both the emailed acknowledgement and the review-link approval.
async function notifyPaymentContact(rec, baseUrl) {
  const key = `pay:${rec.id}`;
  if (ackBusy.has(key)) return false;
  ackBusy.add(key);
  try {
    const current = await store.getRecord(rec.id);
    if (current.ackTracking.bhargaviNotifiedAt) return false;

    const paymentToken = crypto.randomBytes(16).toString('hex');
    await store.updateRecord(rec.id, { payment: { status: 'pending', responseToken: paymentToken, respondedAt: '' } });
    const upd = await store.getRecord(rec.id);
    const t = upd.ackTracking;
    const subject = `FnF Approved — (${employeeSubjectTag(upd)})`;
    const html = paymentEmailHtml(upd, baseUrl);

    let ownId = null;
    if (t.conversationId) {
      const root = await outlook.findRootMessage(SENDER_UPN, t.conversationId);
      if (root) {
        ownId = (await outlook.forwardMessage({ mailboxUpn: SENDER_UPN, messageId: root.id, to: PAYMENT_CONTACT_UPN, subject, commentHtml: html })).internetMessageId;
      }
    }
    if (!ownId) await outlook.sendMail({ fromUpn: SENDER_UPN, to: PAYMENT_CONTACT_UPN, subject, html });

    await store.updateRecord(rec.id, {
      ackTracking: {
        bhargaviNotifiedAt: new Date().toISOString(),
        status: t.status === 'not_tracked' ? 'not_tracked' : 'sent_to_bhargavi',
        ownMessageIds: ownId ? [...t.ownMessageIds, ownId] : t.ownMessageIds,
      },
    });
    return true;
  } finally {
    ackBusy.delete(key);
  }
}

async function processAckTracking(rec) {
  const key = `ack:${rec.id}`;
  if (ackBusy.has(key)) return;
  ackBusy.add(key);
  try {
    let t = rec.ackTracking;
    if (!t.conversationId || !['awaiting', 'reminder_sent', 'acknowledged'].includes(t.status)) return;
    const reload = async () => (await store.getRecord(rec.id)).ackTracking;

    // 1. Has the employee clearly acknowledged, in this thread?
    if (t.status !== 'acknowledged') {
      const employee = t.employeeEmail.toLowerCase();
      const own = new Set(t.ownMessageIds);
      const seen = new Set(t.seenMessageIds);
      const messages = await outlook.listConversation(SENDER_UPN, t.conversationId);
      const replies = messages
        .filter((m) => ((m.from && m.from.emailAddress && m.from.emailAddress.address) || '').toLowerCase() === employee
          && !own.has(m.internetMessageId) && new Date(m.receivedDateTime) > new Date(t.sentAt))
        .sort((a, b) => new Date(a.receivedDateTime) - new Date(b.receivedDateTime));

      let ack = null;
      const notAck = [];
      for (const m of replies) {
        if (seen.has(m.id)) continue; // already read and judged not an acknowledgement
        const text = await outlook.getUniqueBodyText(SENDER_UPN, m.id);
        if (classifyAck(text) === 'ack') { ack = text; break; }
        notAck.push(m.id);
      }
      if (ack) {
        const at = new Date().toISOString();
        await store.updateRecord(rec.id, {
          ackTracking: { status: 'acknowledged', acknowledgedAt: at, ackPreview: ack.trim().replace(/\s+/g, ' ').slice(0, 160), lastCheckedAt: at },
          sheetApproval: { status: 'approved', respondedAt: at.slice(0, 10) },
          ack: { shared: true, acknowledged: true },
        });
      } else {
        await store.updateRecord(rec.id, { ackTracking: { seenMessageIds: [...t.seenMessageIds, ...notAck], lastCheckedAt: new Date().toISOString() } });
      }
      t = await reload();
    }

    // 2. No acknowledgement within the window: one reminder, as a reply on the initial thread.
    if (t.status === 'awaiting' && !t.reminderSentAt && Date.now() - new Date(t.sentAt).getTime() >= ACK_REMINDER_HOURS * 3600 * 1000) {
      const root = await outlook.findRootMessage(SENDER_UPN, t.conversationId);
      if (root) {
        const reminder = await outlook.replyInThread({ mailboxUpn: SENDER_UPN, messageId: root.id, to: t.employeeEmail, commentHtml: ackReminderHtml(rec) });
        await store.updateRecord(rec.id, {
          ackTracking: { status: 'reminder_sent', reminderSentAt: new Date().toISOString(), ownMessageIds: [...t.ownMessageIds, reminder.internetMessageId] },
        });
        t = await reload();
      }
    }

    // 3. Acknowledged: tell Bhargavi once, via the initial email.
    if (t.status === 'acknowledged' && !t.bhargaviNotifiedAt) {
      await notifyPaymentContact(await store.getRecord(rec.id), PUBLIC_BASE_URL);
    }
  } finally {
    ackBusy.delete(key);
  }
}

async function runAckTracking() {
  try {
    for (const rec of await store.listRecords()) {
      if (!rec.ackTracking || !['awaiting', 'reminder_sent', 'acknowledged'].includes(rec.ackTracking.status)) continue;
      try { await processAckTracking(rec); } catch (err) { console.error(`Ack tracking failed for ${rec.name}:`, err.message); }
    }
  } catch (err) {
    console.error('Ack tracking run failed:', err.message);
  }
}

function tdsSubject(rec) {
  return `TDS: ${rec.name || ''} (${rec.empId || ''})`;
}

function inr0(n) {
  return Math.round(Number(n) || 0).toLocaleString('en-IN');
}

// Accounts-team TDS working sheet: April -> LWD month, laid out like the HR
// reference sheet (gross in the corner, month columns, PF columns, Net Salary,
// then the TDS already deducted per month).
function tdsCalcHtml(tds, failedMonths) {
  if (!tds.applicable) {
    return `<div style="margin-top:28px;padding:14px 16px;border:1px solid #cfe0f0;border-radius:6px;background:#f4f9fe;font-size:13px;">
      <b>TDS calculation not required.</b> No TDS was deducted from this employee's Net Payable in ${tds.previousMonth || 'the previous month'}.
    </div>`;
  }
  const th = 'padding:6px 10px;border:1px solid #cfe0f0;background:#eef5fc;font-weight:600;white-space:nowrap;text-align:right;';
  const td = 'padding:6px 10px;border:1px solid #cfe0f0;text-align:right;white-space:nowrap;font-family:monospace;';
  const lbl = 'padding:6px 10px;border:1px solid #cfe0f0;white-space:nowrap;font-weight:600;';
  const cols = tds.columns;
  const prior = cols.filter((c) => !c.isLast);
  const cell = (v) => `<td style="${td}">${inr0(v)}</td>`;
  const blank = `<td style="${td}"></td>`;

  const head = `<tr>
    <td style="${th}">${inr0(tds.gross)}</td>
    ${cols.map((c) => `<th style="${th}">${c.label}</th>`).join('')}
    <th style="${th}">Total</th>
    ${cols.map((c) => `<th style="${th}">PF - ${c.label}</th>`).join('')}
    <th style="${th}">Net Salary</th>
    ${prior.map((c) => `<th style="${th}">TDS - ${c.label}</th>`).join('')}
    <th style="${th}">Total TDS</th>
  </tr>`;

  const row = (label, key, first) => `<tr>
    <td style="${lbl}">${label}</td>
    ${cols.map((c) => cell(c[key])).join('')}
    ${cell(tds.totals[key])}
    ${first ? cols.map((c) => cell(c.pf)).join('') : cols.map(() => blank).join('')}
    ${cell(tds.net[key])}
    ${first ? prior.map((c) => (c.tds == null ? `<td style="${td}">n/a</td>` : cell(c.tds))).join('') : prior.map(() => blank).join('')}
    ${first ? cell(tds.totals.tds) : blank}
  </tr>`;

  const netTotal = tds.net.basic + tds.net.hra + tds.net.other;
  const warn = failedMonths && failedMonths.length
    ? `<p style="color:#b3261e;font-size:12px;">Could not fetch TDS from PMS for: ${failedMonths.join(', ')} — shown as n/a. Reload to retry.</p>` : '';

  return `<h3 style="margin-top:36px;">TDS Calculation</h3>
  <p style="color:#4c6480;font-size:13px;margin-top:-6px;">Gross Salary from Appraisal Master. Basic = Gross ÷ 2, HRA = Basic ÷ 2, Other Allowance = Gross − Basic − HRA. PF is the same every full month; the LWD month is pro-rated on Present Days ÷ Working Days (${tds.columns[tds.columns.length - 1].label}).</p>
  <div style="overflow-x:auto;">
    <table style="border-collapse:collapse;font-size:12px;">
      ${head}
      ${row('Basic', 'basic', true)}
      ${row('HRA', 'hra', false)}
      ${row('Other Allowances', 'other', false)}
    </table>
  </div>
  <p style="font-size:13px;margin-top:12px;"><b>Total Net Salary (Apr – ${cols[cols.length - 1].label}):</b> <span style="font-family:monospace;">${fmtInr(netTotal)}</span> &nbsp;|&nbsp; <b>TDS already deducted:</b> <span style="font-family:monospace;">${fmtInr(tds.totals.tds)}</span></p>
  ${warn}`;
}

// Preview of the FnF sheet. HR sees exactly what the employee will see; the
// Accounts view (audience=accounts) appends the TDS calculation below it.
function previewPageHtml(rec, accountsExtraHtml) {
  const sheetHtml = settlementSheetHtml(rec);
  const accounts = accountsExtraHtml != null;
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>FnF Sheet Preview</title></head>
  <body style="font-family:Segoe UI,Arial,sans-serif;max-width:${accounts ? '1280px' : '700px'};margin:${accounts ? '16px' : '40px'} auto;padding:0 16px;color:#0a1f3a;">
    <h2>Full &amp; Final Settlement — ${rec.name||''}</h2>
    <p style="color:#4c6480;">${accounts ? 'Accounts view — the FnF sheet followed by its TDS calculation.' : 'Preview only — this is exactly what the employee will see once you send it. No action has been taken.'}</p>
    <div style="max-width:700px;">${sheetHtml}</div>
    ${accounts ? accountsExtraHtml : ''}
  </body></html>`;
}

// Sent to the payment contact (e.g. Bhargavi) once the employee approves —
// includes the full FnF sheet plus Payment Done / Not Done buttons, styled
// like the Accept/Reject Offer pattern.
function paymentEmailHtml(rec, baseUrl) {
  const doneUrl = `${baseUrl}/api/records/${rec.id}/payment-response?action=done&token=${rec.payment.responseToken}`;
  const notDoneUrl = `${baseUrl}/api/records/${rec.id}/payment-response?action=not_done&token=${rec.payment.responseToken}`;
  return `
  <div style="font-family:Segoe UI,Arial,sans-serif;color:#0a1f3a;max-width:700px;">
    <p>Hi Bhargavi,</p>
    <p>${rec.name || ''} (${rec.empId || ''}) has approved their Full &amp; Final settlement. Please review the sheet below and process the payment.</p>
    ${settlementSheetHtml(rec)}
    <div style="margin:28px 0 8px;">
      <a href="${doneUrl}" style="background:#0f9d58;color:#fff;padding:12px 24px;border-radius:5px;text-decoration:none;font-weight:600;margin-right:12px;">✓ Payment Done</a>
      <a href="${notDoneUrl}" style="background:#d64550;color:#fff;padding:12px 24px;border-radius:5px;text-decoration:none;font-weight:600;">✗ Payment Not Done</a>
    </div>
    <p style="font-size:12px;color:#4c6480;">Clicking a button above will record your response instantly. If the buttons don't work, reply to this email with your decision.</p>
  </div>`;
}

function responsePageHtml(title, message) {
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>${title}</title></head>
  <body style="font-family:Segoe UI,Arial,sans-serif;max-width:600px;margin:60px auto;text-align:center;color:#0a1f3a;">
    <h2>${title}</h2>
    <p>${message}</p>
  </body></html>`;
}

async function handleApi(req, res, urlPath) {
  const asOf = new URL(req.url, `http://${req.headers.host}`).searchParams.get('asOf') || new Date().toISOString().slice(0, 10);

  if (urlPath === '/api/records' && req.method === 'GET') {
    const records = await store.listRecords();
    sendJson(res, 200, records.map((r) => Object.assign(withDerived(r), { sla: Object.assign({}, r.sla, { window: recSla(r, asOf) }) })));
    return true;
  }

  if (urlPath === '/api/records' && req.method === 'POST') {
    const body = await readBody(req);
    const rec = await store.createRecord(body);
    sendJson(res, 201, withDerived(rec));
    return true;
  }

  const singleMatch = urlPath.match(/^\/api\/records\/(\d+)$/);
  if (singleMatch) {
    const id = singleMatch[1];
    if (req.method === 'GET') {
      const rec = await store.getRecord(id);
      if (!rec) return sendJson(res, 404, { error: 'Not found' }), true;
      sendJson(res, 200, Object.assign(withDerived(rec), { sla: Object.assign({}, rec.sla, { window: recSla(rec, asOf) }) }));
      return true;
    }
    if (req.method === 'PATCH') {
      const body = await readBody(req);
      // Owned by the server (set by emailed links / mailbox replies) — a stale
      // browser copy must never overwrite them with its older view.
      delete body.pli;
      delete body.tdsReply;
      delete body.ackTracking;
      const rec = await store.updateRecord(id, body);
      if (!rec) return sendJson(res, 404, { error: 'Not found' }), true;
      // The PLI contact's confirmed amount is the final Incentives figure.
      if (rec.pli.status === 'submitted') {
        rec.settlement.earnings.incentives = rec.pli.amount;
        rec.comp.pli = rec.pli.amount;
      }
      sendJson(res, 200, withDerived(rec));
      return true;
    }
    if (req.method === 'DELETE') {
      const ok = await store.deleteRecord(id);
      if (!ok) return sendJson(res, 404, { error: 'Not found' }), true;
      sendJson(res, 204, {});
      return true;
    }
  }

  // Sends the real FnF review email straight to the employee's personal
  // email, triggered directly by the "Send Email to Employee" button.
  const sendApprovalMatch = urlPath.match(/^\/api\/records\/(\d+)\/send-approval-email$/);
  if (sendApprovalMatch && req.method === 'POST') {
    const id = sendApprovalMatch[1];
    const rec = await store.getRecord(id);
    if (!rec) return sendJson(res, 404, { error: 'Not found' }), true;
    if (!rec.personalEmail) return sendJson(res, 400, { error: 'Personal email is required before sending' }), true;

    const reviewToken = crypto.randomBytes(16).toString('hex');
    await store.updateRecord(id, { emailSentDate: asOf });
    await store.updateRecord(id, {
      sheetApproval: { status: 'sent', sentAt: asOf, respondedAt: '', reviewToken, rejectionReason: '' },
    });
    const updated = await store.getRecord(id);
    const proto = req.headers['x-forwarded-proto'] || 'http';
    const baseUrl = `${proto}://${req.headers.host}`;

    // Sent as a tracked message: its conversation is the one reference for the
    // acknowledgement, 24h reminder and the hand-off to the payment contact.
    const sentAtIso = new Date().toISOString();
    const sent = await outlook.sendMailTracked({
      fromUpn: SENDER_UPN,
      to: updated.personalEmail,
      cc: SENDER_UPN,
      subject: `Full & Final Settlement — (${employeeSubjectTag(updated)})`,
      html: approvalEmailHtml(updated, baseUrl),
    });
    await store.updateRecord(id, {
      ackTracking: {
        status: 'awaiting', conversationId: sent.conversationId, employeeEmail: updated.personalEmail,
        sentAt: sentAtIso, reminderSentAt: '', acknowledgedAt: '', ackPreview: '', bhargaviNotifiedAt: '',
        ownMessageIds: [sent.internetMessageId], seenMessageIds: [], lastCheckedAt: '',
      },
    });

    sendJson(res, 200, withDerived(await store.getRecord(id)));
    return true;
  }

  // Looks at the initial FnF thread now (the dashboard and the background job both call this).
  const checkAckMatch = urlPath.match(/^\/api\/records\/(\d+)\/check-ack$/);
  if (checkAckMatch && req.method === 'POST') {
    const rec = await store.getRecord(checkAckMatch[1]);
    if (!rec) return sendJson(res, 404, { error: 'Not found' }), true;
    await processAckTracking(rec);
    sendJson(res, 200, withDerived(await store.getRecord(rec.id)));
    return true;
  }

  const previewMatch = urlPath.match(/^\/api\/records\/(\d+)\/preview$/);
  if (previewMatch && req.method === 'GET') {
    const id = previewMatch[1];
    const rec = await store.getRecord(id);
    if (!rec) {
      res.writeHead(404, { 'Content-Type': 'text/html' });
      return res.end(responsePageHtml('Not found', 'This settlement record could not be found.')), true;
    }
    let accountsExtra;
    if (new URL(req.url, `http://${req.headers.host}`).searchParams.get('audience') === 'accounts') {
      const { tdsByMonth, failed } = await store.resolveTdsByMonth(rec);
      accountsExtra = tdsCalcHtml(computeTdsCalc(rec, tdsByMonth), failed);
    }
    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end(previewPageHtml(rec, accountsExtra));
    return true;
  }

  // Emails the FnF sheet plus its TDS calculation to the accounts/tax contact.
  // ACCOUNTS_CONTACT_UPN is the test mailbox for now; point it at Sarika to go live.
  const sendAccountsMatch = urlPath.match(/^\/api\/records\/(\d+)\/send-to-accounts$/);
  if (sendAccountsMatch && req.method === 'POST') {
    const rec = await store.getRecord(sendAccountsMatch[1]);
    if (!rec) return sendJson(res, 404, { error: 'Record not found' }), true;
    const to = (process.env.ACCOUNTS_CONTACT_UPN || SENDER_UPN).split(',').map((s) => s.trim()).filter(Boolean);

    const { tdsByMonth, failed } = await store.resolveTdsByMonth(rec);
    const html = `
      <div style="font-family:Segoe UI,Arial,sans-serif;color:#0a1f3a;">
        <p>Hi,</p>
        <p>Please find the Full &amp; Final settlement sheet for ${rec.name || ''} (${rec.empId || ''}) below, followed by the TDS calculation.</p>
        ${settlementSheetHtml(rec)}
        ${tdsCalcHtml(computeTdsCalc(rec, tdsByMonth), failed)}
      </div>`;
    const sentAt = new Date().toISOString();
    await outlook.sendMail({
      fromUpn: SENDER_UPN,
      to,
      subject: tdsSubject(rec),
      html,
    });
    await store.updateRecord(rec.id, {
      sla: { taxEmailSentAt: sentAt.slice(0, 10) },
      tdsReply: { status: 'awaiting', sentAt, messageId: '', receivedAt: '', amount: null, preview: '' },
    });
    sendJson(res, 200, { sentTo: to.join(', ') });
    return true;
  }

  // HR (or the auto-share job) sends a Sales employee's details to the PLI contact.
  const sharePliMatch = urlPath.match(/^\/api\/records\/(\d+)\/share-pli$/);
  if (sharePliMatch && req.method === 'POST') {
    const rec = await store.getRecord(sharePliMatch[1]);
    if (!rec) return sendJson(res, 404, { error: 'Record not found' }), true;
    if (!isSalesEmployee(rec)) return sendJson(res, 400, { error: 'PLI is only confirmed for Sales employees' }), true;
    const to = await sharePli(rec);
    sendJson(res, 200, { sentTo: to.join(', ') });
    return true;
  }

  // The PLI contact's token-gated page: details only (no salary), one amount field.
  const pliPageMatch = urlPath.match(/^\/api\/pli\/(\d+)$/);
  if (pliPageMatch) {
    const rec = await store.getRecord(pliPageMatch[1]);
    const q = new URL(req.url, `http://${req.headers.host}`).searchParams;
    const body = req.method === 'POST' ? await readBody(req) : {};
    const token = req.method === 'POST' ? body.token : q.get('token');
    const valid = rec && rec.pli.token && token === rec.pli.token;

    if (req.method === 'GET') {
      res.writeHead(valid ? 200 : 403, { 'Content-Type': 'text/html' });
      res.end(valid ? pliPageHtml(rec) : responsePageHtml('Invalid link', 'This PLI link is invalid or has expired.'));
      return true;
    }
    if (req.method === 'POST') {
      if (!valid) return sendJson(res, 403, { error: 'This PLI link is invalid or has expired.' }), true;
      const amount = parsePliAmount(body.amount);
      if (amount == null) return sendJson(res, 400, { error: 'Enter a valid amount, or 0 / NIL if there is no PLI.' }), true;
      await store.updateRecord(rec.id, {
        pli: { status: 'submitted', amount, submittedAt: new Date().toISOString() },
        settlement: { earnings: { incentives: amount } },
        comp: { pli: amount },
      });
      sendJson(res, 200, { amount });
      return true;
    }
  }

  // Looks for the tax team's reply to the TDS email and, if it states an amount
  // or NIL, writes it into the settlement sheet's TDS deduction.
  const checkReplyMatch = urlPath.match(/^\/api\/records\/(\d+)\/check-tds-reply$/);
  if (checkReplyMatch && req.method === 'POST') {
    const rec = await store.getRecord(checkReplyMatch[1]);
    if (!rec) return sendJson(res, 404, { error: 'Record not found' }), true;
    if (!rec.tdsReply || !rec.tdsReply.sentAt) return sendJson(res, 400, { error: 'This sheet has not been sent to the Accounts team yet' }), true;

    const reply = await outlook.findLatestReply({
      mailboxUpn: SENDER_UPN,
      since: rec.tdsReply.sentAt,
      subjectTag: tdsSubject(rec),
      // Tanya (the verifier) and the Accounts contact(s) — ACCOUNTS_CONTACT_UPN may list several, comma-separated.
      fromUpns: [SENDER_UPN, ...(process.env.ACCOUNTS_CONTACT_UPN || '').split(',')],
    });

    if (reply && reply.id !== rec.tdsReply.messageId) {
      const parsed = parseTdsReply(reply.text);
      const preview = reply.text.trim().replace(/\s+/g, ' ').slice(0, 160);
      const base = { messageId: reply.id, receivedAt: reply.receivedAt, preview };
      if (parsed.kind === 'unclear') {
        await store.updateRecord(rec.id, { tdsReply: { ...base, status: 'unclear', amount: null } });
      } else {
        const amount = parsed.kind === 'nil' ? 0 : parsed.amount;
        await store.updateRecord(rec.id, {
          tdsReply: { ...base, status: 'applied', amount },
          settlement: { deductions: { tds: amount } },
        });
      }
    }
    const updated = await store.getRecord(rec.id);
    sendJson(res, 200, { tdsReply: updated.tdsReply, tds: updated.settlement.deductions.tds });
    return true;
  }

  const reviewMatch = urlPath.match(/^\/api\/records\/(\d+)\/review$/);
  if (reviewMatch && req.method === 'GET') {
    const id = reviewMatch[1];
    const token = new URL(req.url, `http://${req.headers.host}`).searchParams.get('token');
    const rec = await store.getRecord(id);

    if (!rec) {
      res.writeHead(404, { 'Content-Type': 'text/html' });
      return res.end(responsePageHtml('Not found', 'This settlement record could not be found.')), true;
    }
    if (rec.sheetApproval.status !== 'sent') {
      res.writeHead(200, { 'Content-Type': 'text/html' });
      return res.end(responsePageHtml('Already responded', `This settlement was already marked as "${rec.sheetApproval.status}".`)), true;
    }
    if (token !== rec.sheetApproval.reviewToken) {
      res.writeHead(403, { 'Content-Type': 'text/html' });
      return res.end(responsePageHtml('Invalid link', 'This review link is invalid or has expired.')), true;
    }

    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end(reviewPageHtml(rec));
    return true;
  }

  const respondMatch = urlPath.match(/^\/api\/records\/(\d+)\/respond$/);
  if (respondMatch && req.method === 'GET') {
    const id = respondMatch[1];
    const q = new URL(req.url, `http://${req.headers.host}`).searchParams;
    const action = q.get('action');
    const token = q.get('token');
    const reason = (q.get('reason') || '').trim();
    const rec = await store.getRecord(id);

    if (!rec) {
      res.writeHead(404, { 'Content-Type': 'text/html' });
      return res.end(responsePageHtml('Not found', 'This settlement record could not be found.')), true;
    }
    if (rec.sheetApproval.status !== 'sent') {
      res.writeHead(200, { 'Content-Type': 'text/html' });
      return res.end(responsePageHtml('Already responded', `This settlement was already marked as "${rec.sheetApproval.status}".`)), true;
    }
    if (token !== rec.sheetApproval.reviewToken || (action !== 'approve' && action !== 'reject')) {
      res.writeHead(403, { 'Content-Type': 'text/html' });
      return res.end(responsePageHtml('Invalid link', 'This approval link is invalid or has expired.')), true;
    }
    if (action === 'reject' && !reason) {
      res.writeHead(400, { 'Content-Type': 'text/html' });
      return res.end(responsePageHtml('Reason required', 'Please go back and enter a reason for rejecting the settlement.')), true;
    }

    const status = action === 'approve' ? 'approved' : 'rejected';
    await store.updateRecord(id, { sheetApproval: { status, respondedAt: asOf, rejectionReason: action === 'reject' ? reason : '' } });
    if (action === 'approve') {
      // Reflects in the dashboard's "Employee has reviewed & acknowledged" toggle.
      await store.updateRecord(id, { ack: { shared: true, acknowledged: true } });
    }

    if (action === 'approve') {
      const proto = req.headers['x-forwarded-proto'] || 'http';
      await notifyPaymentContact(await store.getRecord(id), `${proto}://${req.headers.host}`);
    } else {
      await outlook.sendMail({
        fromUpn: SENDER_UPN,
        to: SENDER_UPN,
        subject: `FnF Rejected — Action Needed (${rec.name})`,
        html: `<p>${rec.name} (${rec.empId}) has rejected their Full &amp; Final settlement.</p><p><b>Reason:</b> ${reason}</p>`,
      });
    }

    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end(responsePageHtml(
      action === 'approve' ? 'Settlement Approved' : 'Settlement Rejected',
      action === 'approve'
        ? 'Thank you — your response has been recorded and payment processing has been notified.'
        : 'Thank you — your response and reason have been recorded. HR will reach out regarding next steps.'
    ));
    return true;
  }

  // Bhargavi (or whoever the payment contact is) clicks Payment Done / Not
  // Done from the approval-notification email.
  const paymentResponseMatch = urlPath.match(/^\/api\/records\/(\d+)\/payment-response$/);
  if (paymentResponseMatch && req.method === 'GET') {
    const id = paymentResponseMatch[1];
    const q = new URL(req.url, `http://${req.headers.host}`).searchParams;
    const action = q.get('action');
    const token = q.get('token');
    const rec = await store.getRecord(id);

    if (!rec) {
      res.writeHead(404, { 'Content-Type': 'text/html' });
      return res.end(responsePageHtml('Not found', 'This settlement record could not be found.')), true;
    }
    if (rec.payment.status !== 'pending') {
      res.writeHead(200, { 'Content-Type': 'text/html' });
      return res.end(responsePageHtml('Already responded', `Payment for this settlement was already marked as "${rec.payment.status}".`)), true;
    }
    if (token !== rec.payment.responseToken || (action !== 'done' && action !== 'not_done')) {
      res.writeHead(403, { 'Content-Type': 'text/html' });
      return res.end(responsePageHtml('Invalid link', 'This payment response link is invalid or has expired.')), true;
    }

    await store.updateRecord(id, {
      payment: { status: action, respondedAt: asOf, apConfirmed: action === 'done', apInformed: true },
    });

    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end(responsePageHtml(
      action === 'done' ? 'Payment Recorded as Done' : 'Payment Recorded as Not Done',
      action === 'done'
        ? 'Thank you — payment has been marked as completed for this settlement.'
        : 'Thank you — payment has been marked as not yet done. HR has been notified.'
    ));
    return true;
  }

  // Lazy-loaded on demand from the "Net Payable — Previous Months" panel —
  // fetching this for every employee on every dashboard load pushed cold-start
  // latency past what's acceptable, since each month is a sequential PMS call.
  const netPayHistoryMatch = urlPath.match(/^\/api\/records\/(\d+)\/netpay-history$/);
  if (netPayHistoryMatch && req.method === 'GET') {
    const id = netPayHistoryMatch[1];
    const rec = await store.getRecord(id);
    if (!rec) return sendJson(res, 404, { error: 'Record not found' }), true;

    if (!rec.settlement.netPayableHistory || !rec.settlement.netPayableHistory.length) {
      const history = await store.resolveNetPayableHistory(rec.empId, rec.lwd);
      await store.updateRecord(id, { settlement: { netPayableHistory: history } });
    }
    const updated = await store.getRecord(id);
    sendJson(res, 200, { netPayableHistory: updated.settlement.netPayableHistory });
    return true;
  }

  if (urlPath === '/api/pms/notice-period-summary' && req.method === 'GET') {
    const asOf = new URL(req.url, `http://${req.headers.host}`).searchParams.get('asOf') || new Date().toISOString().slice(0, 10);
    const employees = await pms.getEmployeeDetails('');
    const resigned = employees.filter((e) => e.date_of_resigantion);
    const servingNotice = resigned.filter((e) => !e.last_working_day || e.last_working_day >= asOf);
    sendJson(res, 200, {
      resignedCount: resigned.length,
      servingNoticeCount: servingNotice.length,
      employees: servingNotice.map((e) => ({
        name: [e.first_name, e.middle_name, e.last_name].filter(Boolean).join(' '),
        department: e.deparment_name,
        designation: e.designation_name,
        dateOfResignation: e.date_of_resigantion,
        lastWorkingDay: e.last_working_day,
      })),
    });
    return true;
  }

  if (urlPath === '/api/pms/exit-list' && req.method === 'GET') {
    const q = new URL(req.url, `http://${req.headers.host}`).searchParams;
    const fromDate = q.get('fromDate') || '2026-04-01';
    const toDate = q.get('toDate') || new Date().toISOString().slice(0, 10);
    const exits = await pms.getExitEmployeeList(fromDate, toDate);
    sendJson(res, 200, {
      fromDate,
      toDate,
      count: exits.length,
      employees: exits.map((e) => ({
        empCode: e.EmpCode,
        name: e.Name,
        department: e['Deptt.'],
        dor: e.DOR ? e.DOR.slice(0, 10) : '',
        lwd: e.LWD ? e.LWD.slice(0, 10) : '',
        baseLocation: e.BaseLocation,
      })),
    });
    return true;
  }

  return false;
}

const server = http.createServer((req, res) => {
  const parsed = new URL(req.url, `http://${req.headers.host}`);
  const urlPath = decodeURIComponent(parsed.pathname);

  if (urlPath.startsWith('/api/')) {
    handleApi(req, res, urlPath)
      .then((handled) => {
        if (!handled) sendJson(res, 404, { error: 'Not found' });
      })
      .catch((err) => {
        console.error(`${req.method} ${urlPath} failed:`, err.message);
        sendJson(res, 502, { error: err.message || 'Internal error' });
      });
    return;
  }

  const filePath = path.join(ROOT, 'public', urlPath === '/' ? '/index.html' : urlPath);
  if (!filePath.startsWith(path.join(ROOT, 'public'))) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }
  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end('Not found: ' + urlPath);
      return;
    }
    const ext = path.extname(filePath);
    res.writeHead(200, { 'Content-Type': MIME[ext] || 'application/octet-stream' });
    res.end(data);
  });
});

if (require.main === module) {
  server.listen(PORT, () => {
    console.log(`FnF Settlement app serving at http://localhost:${PORT}`);
    // Local/long-running only: serverless hosts need a scheduled job instead.
    setTimeout(runPliSharing, 30 * 1000);
    setInterval(runPliSharing, 30 * 60 * 1000);
    setInterval(runAckTracking, 5 * 60 * 1000);
  });
}

module.exports = server;
