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
const { calc, recStatus, recSla, settlementCalc, settlementMonthLabels } = require('./lib/calc');
const pms = require('./lib/pms');
const outlook = require('./lib/outlook');

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

// HR-facing preview — same sheet the employee will see, no Approve/Reject actions.
function previewPageHtml(rec) {
  const sheetHtml = settlementSheetHtml(rec);
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>FnF Sheet Preview</title></head>
  <body style="font-family:Segoe UI,Arial,sans-serif;max-width:700px;margin:40px auto;color:#0a1f3a;">
    <h2>Full &amp; Final Settlement — ${rec.name||''}</h2>
    <p style="color:#4c6480;">Preview only — this is exactly what the employee will see once you send it. No action has been taken.</p>
    ${sheetHtml}
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
      const rec = await store.updateRecord(id, body);
      if (!rec) return sendJson(res, 404, { error: 'Not found' }), true;
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

    await outlook.sendMail({
      fromUpn: SENDER_UPN,
      to: updated.personalEmail,
      cc: SENDER_UPN,
      subject: `Full & Final Settlement — (${employeeSubjectTag(updated)})`,
      html: approvalEmailHtml(updated, baseUrl),
    });

    sendJson(res, 200, withDerived(updated));
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
    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end(previewPageHtml(rec));
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
      const paymentToken = crypto.randomBytes(16).toString('hex');
      await store.updateRecord(id, { payment: { status: 'pending', responseToken: paymentToken, respondedAt: '' } });
      const updatedForPayment = await store.getRecord(id);
      const proto = req.headers['x-forwarded-proto'] || 'http';
      const baseUrl = `${proto}://${req.headers.host}`;
      await outlook.sendMail({
        fromUpn: SENDER_UPN,
        to: PAYMENT_CONTACT_UPN,
        subject: `FnF Approved — (${employeeSubjectTag(updatedForPayment)})`,
        html: paymentEmailHtml(updatedForPayment, baseUrl),
      });
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

server.listen(PORT, () => {
  console.log(`FnF Settlement app serving at http://localhost:${PORT}`);
});
