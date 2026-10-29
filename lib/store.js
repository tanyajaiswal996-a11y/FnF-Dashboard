// Data access layer for FnF records.
//
// Backed by live PMS data (lib/pms.js) — on first access this fetches every
// employee with a recorded resignation date and seeds one FnF record per
// person. The resulting array is cached in memory for the life of the
// process; edits made in the app (checklist, comp figures, etc.) mutate that
// cached array directly. Swap loadRecords() for a real DB read when Turso
// access is granted.

const pms = require('./pms');
const kv = require('./kv');
const { computeWorkingPresentDays, computeSalarySplit, computeGratuity, tdsCalcStartIndex, monthIndex, isTdsApplicable, pfForMonth, pfMonthsForLwd } = require('./calc');

let nextId = 1;

// PMS only exposes the manager's first name (e.g. "Sakshi"), which is
// ambiguous across the company. Scoped per-empId overrides let HR confirm the
// real full name case by case, without guessing wrong for everyone who shares
// that first name with a different manager.
const MANAGER_FULL_NAME_OVERRIDES = {
  '1285': 'Sakshi Gaba Dhawan', // Shubhrangshu Saha
};

function newRecord(seed) {
  seed = seed || {};
  const base = {
    id: nextId++,
    empId: '', name: '', department: '', designation: '', location: '', dor: '', lwd: '', contact: '', personalEmail: '', emailCc: '', emailBcc: '',
    isRayontara: false, // set once the Rayontara-affiliation check is wired up
    started: false,
    manager: '', managerApproved: false, pmsUpdated: false,
    checklist: [
      { label: 'ID card & access card returned', done: false },
      { label: 'Laptop / IT asset return confirmed', done: false },
      { label: 'Handover note submitted to manager', done: false },
      { label: 'Exit interview scheduled', done: false },
    ],
    dues: {
      exam: { required: 'not_required', amount: 0 },
      wfh: { balance: 0 },
      loanAdvance: { amount: 0 },
      dcTeam: { assetIssued: 'No' },
      noticePay: { applicable: false, date: '', amount: 0 },
    },
    fnfDue: { pendingFrom: 'none' },
    comp: { grossSalary: 0, epf: 0, workingDays: 26, presentDays: 26, tenureYears: 0, taBill: 0, nps: 0, vpf: 0, tds: 0, pli: 0 },
    // Detailed settlement breakdown — mirrors the HR FnF sheet template exactly,
    // independent of the simplified `comp`/`dues` figures used elsewhere in the app.
    settlement: {
      doj: '',
      earnings: {
        salaryCurrentMonth: 0, salaryPreviousMonth: 0, incentives: 0, taBillPositive: 0,
        gratuity: 0, otherReimbursement: 0, noticePeriodPay: 0,
      },
      deductions: {
        pfCurrentMonth: 0, pfPreviousMonth: 0, examDues: 0, wfhDues: 0, nps: 0, vpf: 0,
        taBillNegative: 0, pluxeeMeal: 0, esi: 0, tds: 0, pt: 0,
      },
      // Flagged from PMS "Net Payable Details" for the month before LWD — surfaced
      // to HR as a heads-up, not auto-merged into the deduction totals above.
      previousMonthDues: { month: '', tds: 0, nps: 0 },
      netPayableHistory: [],
      tdsByMonth: {}, // 'YYYY-MM' -> TDS deducted that month; filled on demand for the Accounts TDS sheet
    },
    // Tax team's emailed answer to the TDS sheet: status is 'not_sent' | 'awaiting' |
    // 'applied' (amount or NIL written into settlement.deductions.tds) | 'unclear'.
    tdsReply: { status: 'not_sent', sentAt: '', messageId: '', receivedAt: '', amount: null, preview: '' },
    // Final PLI/incentive amount confirmed by the PLI contact via a token link.
    // status: 'not_shared' | 'shared' | 'submitted'. No salary data is ever shared.
    pli: { status: 'not_shared', sharedAt: '', amount: null, submittedAt: '' },
    // Acknowledgement tracking for the initial FnF email thread (server-owned).
    // status: 'not_tracked' | 'awaiting' | 'reminder_sent' | 'acknowledged' | 'sent_to_bhargavi'
    ackTracking: {
      status: 'not_tracked', conversationId: '', employeeEmail: '', sentAt: '',
      reminderSentAt: '', acknowledgedAt: '', ackPreview: '', bhargaviNotifiedAt: '',
      ownMessageIds: [], seenMessageIds: [], lastCheckedAt: '',
    },
    ack: { shared: false, acknowledged: false },
    emailSentDate: '',
    sheetApproval: {
      status: 'not_sent', // 'not_sent' | 'internal_review' | 'sent' | 'approved' | 'rejected'
      sentAt: '',
      respondedAt: '',
      reviewToken: '',
      rejectionReason: '',
    },
    payment: {
      apInformed: false, apConfirmed: false,
      status: 'pending', // 'pending' | 'done' | 'not_done'
      responseToken: '', respondedAt: '',
    },
    closure: { expLetter: false, slips: false },
    sla: {
      tdsCalc: 'pending',
      taxEmailSentAt: '',
      taxApproval: 'pending',
      reminderSent: false,
      tentativeSentAt: '',
      ack: 'pending',
      finalProcessed: false,
    },
  };
  return Object.assign(base, seed, { id: seed.id != null ? seed.id : base.id });
}

function tenureYearsFromJoining(dateOfJoining) {
  if (!dateOfJoining) return 0;
  const years = (Date.now() - new Date(dateOfJoining).getTime()) / (365.25 * 24 * 3600 * 1000);
  return Math.round(years * 10) / 10;
}

async function resolveEmpCode(email) {
  if (!email) return null;
  try {
    const match = await pms.getEmployeeCodeByEmail(email);
    return match ? String(match.EmpCode) : null;
  } catch (err) {
    console.error(`Failed to resolve emp_code for ${email}:`, err.message);
    return null;
  }
}

async function resolveAppraisal(empCode) {
  if (!empCode) return null;
  try {
    return await pms.getAppraisalData(empCode);
  } catch (err) {
    console.error(`Failed to resolve appraisal data for emp_code ${empCode}:`, err.message);
    return null;
  }
}

async function resolveWfhBalance(empCode) {
  if (!empCode) return 0;
  try {
    const balance = await pms.getWfhBalance(empCode);
    return balance == null ? 0 : balance;
  } catch (err) {
    console.error(`Failed to resolve WFH balance for emp_code ${empCode}:`, err.message);
    return 0;
  }
}

async function resolveExamDues(empCode) {
  if (!empCode) return { required: false, amount: 0 };
  try {
    return await pms.getExamPendingDues(empCode);
  } catch (err) {
    console.error(`Failed to resolve exam dues for emp_code ${empCode}:`, err.message);
    return { required: false, amount: 0 };
  }
}

const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

// [startDate, endDate] (YYYY-MM-DD) covering the full calendar month `monthsBack`
// months before `lwd` (monthsBack=1 is the month immediately before LWD's month).
function previousMonthRange(lwd, monthsBack) {
  if (!lwd) return null;
  const [y, m] = lwd.split('-').map(Number); // m is 1-indexed
  const m0 = m - 1; // 0-indexed current month
  const back = monthsBack == null ? 1 : monthsBack;
  const totalMonths = y * 12 + m0 - back;
  const targetY = Math.floor(totalMonths / 12);
  const targetM0 = ((totalMonths % 12) + 12) % 12;
  const lastDay = new Date(Date.UTC(targetY, targetM0 + 1, 0)).getUTCDate();
  return {
    start: `${targetY}-${String(targetM0 + 1).padStart(2, '0')}-01`,
    end: `${targetY}-${String(targetM0 + 1).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`,
    label: `${MONTH_NAMES[targetM0]} ${targetY}`,
  };
}

// Net Payable figures for the month before LWD's month. When that month shows
// no TDS, the month before it is checked too (priorMonth/priorTds): payroll
// usually stops deducting TDS once someone has resigned, so a taxpayer's last
// deduction can sit one month earlier.
async function resolvePreviousMonthDues(empCode, lwd) {
  const empty = { month: '', tds: 0, nps: 0, priorMonth: '', priorTds: 0 };
  const range = previousMonthRange(lwd, 1);
  if (!empCode || !range) return empty;

  const result = { ...empty };
  try {
    const details = await pms.getNetPayableDetails(empCode, range.start, range.end);
    if (details) Object.assign(result, { month: range.label, tds: Number(details.TDS) || 0, nps: Number(details.Total_NPS) || 0 });
  } catch (err) {
    console.error(`Failed to resolve previous-month dues for emp_code ${empCode}:`, err.message);
  }

  if (result.tds <= 0) {
    const prior = previousMonthRange(lwd, 2);
    try {
      const earlier = await pms.getNetPayableDetails(empCode, prior.start, prior.end);
      if (earlier) Object.assign(result, { priorMonth: prior.label, priorTds: Number(earlier.TDS) || 0 });
    } catch (err) {
      console.error(`Failed to resolve earlier-month TDS for emp_code ${empCode}:`, err.message);
    }
  }
  return result;
}

// Collates the last `monthsBack` calendar months of PMS "Net Payable Details"
// (one payroll snapshot per month) so HR can review the full recent history
// alongside the settlement, not just the single month immediately before LWD.
// Months with no payroll run on file (e.g. joined recently) are skipped.
const NET_PAYABLE_HISTORY_MONTHS = 6;

async function resolveNetPayableHistory(empCode, lwd) {
  if (!empCode || !lwd) return [];
  const history = [];
  for (let back = 1; back <= NET_PAYABLE_HISTORY_MONTHS; back++) {
    const range = previousMonthRange(lwd, back);
    if (!range) continue;
    try {
      const details = await pms.getNetPayableDetails(empCode, range.start, range.end);
      if (details) history.push({ month: range.label, ...details });
    } catch (err) {
      console.error(`Failed to resolve net payable history for emp_code ${empCode}, ${range.label}:`, err.message);
    }
  }
  return history;
}

// Fills rec.settlement.tdsByMonth with the TDS deducted in every month from the
// start of the TDS sheet (April / joining month) up to the month before LWD, so
// the Accounts TDS working sheet can list month-wise deductions. Only months
// not already cached are fetched (sequentially — one PMS session). Returns the
// labels of any months whose lookup failed so the caller can flag them.
async function resolveTdsByMonth(rec) {
  const failed = [];
  if (!isTdsApplicable(rec) || !rec.empId) return { tdsByMonth: rec.settlement.tdsByMonth || {}, failed };
  rec.settlement.tdsByMonth = rec.settlement.tdsByMonth || {};
  const cache = rec.settlement.tdsByMonth;

  const prevRange = previousMonthRange(rec.lwd, 1);
  if (prevRange && rec.settlement.previousMonthDues) {
    cache[prevRange.start.slice(0, 7)] = Number(rec.settlement.previousMonthDues.tds) || 0;
  }

  const lwdIdx = monthIndex(rec.lwd);
  for (let idx = tdsCalcStartIndex(rec); idx < lwdIdx; idx++) {
    const range = previousMonthRange(rec.lwd, lwdIdx - idx);
    const ym = range.start.slice(0, 7);
    if (cache[ym] != null) continue;
    try {
      const details = await pms.getNetPayableDetails(rec.empId, range.start, range.end);
      cache[ym] = details ? Number(details.TDS) || 0 : 0;
    } catch (err) {
      console.error(`Failed to resolve TDS for emp_code ${rec.empId}, ${range.label}:`, err.message);
      failed.push(range.label);
    }
  }
  return { tdsByMonth: cache, failed };
}

// Record ids must mean the same employee on every server instance and after a
// restart, so they come from the employee code (not a load-order counter).
function stableId(code, fallbackKey) {
  if (/^\d{1,9}$/.test(String(code))) return Number(code);
  let h = 0;
  for (const ch of String(fallbackKey)) h = (h * 31 + ch.charCodeAt(0)) >>> 0;
  return 1_000_000_000 + (h % 1_000_000_000); // above any numeric employee code, so no clash
}

async function fromPmsEmployee(emp, idx) {
  const name = [emp.first_name, emp.middle_name, emp.last_name].filter(Boolean).join(' ').replace(/\s+/g, ' ').trim();
  const empCode = await resolveEmpCode(emp.email_address);
  const appraisal = await resolveAppraisal(empCode);
  const lwd = emp.last_working_day ? emp.last_working_day.slice(0, 10) : '';
  const dor = emp.date_of_resigantion ? emp.date_of_resigantion.slice(0, 10) : '';
  const previousMonthDues = await resolvePreviousMonthDues(empCode, lwd);
  const wfhBalance = await resolveWfhBalance(empCode);
  const examDues = await resolveExamDues(empCode);
  const { workingDays, presentDays } = computeWorkingPresentDays(lwd);
  const grossSalary = appraisal ? Number(appraisal.Amount) || 0 : 0;
  const epf = appraisal ? Number(appraisal.EPF) || 0 : 0;
  const { salaryCurrentMonth, salaryPreviousMonth } = computeSalarySplit(grossSalary, dor, lwd, workingDays, presentDays);
  // PF follows the monthly amount in force in each month (₹3,000 from Oct 2026), so the LWD month and the held month can differ.
  const pfMonths = pfMonthsForLwd(lwd);
  const { salaryCurrentMonth: pfCurrentMonth } = computeSalarySplit(pfForMonth(epf, pfMonths.current), dor, lwd, workingDays, presentDays);
  const { salaryPreviousMonth: pfPreviousMonth } = computeSalarySplit(pfForMonth(epf, pfMonths.previous), dor, lwd, workingDays, presentDays);
  const tenureYears = tenureYearsFromJoining(emp.date_of_joining);
  const gratuity = computeGratuity(grossSalary, tenureYears);
  const empId = empCode || emp.email_address || `PMS${String(idx + 1).padStart(4, '0')}`;
  return newRecord({
    id: stableId(empCode, empId),
    empId,
    name,
    department: emp.deparment_name || '',
    designation: emp.designation_name || '',
    location: emp.city_name || '',
    dor,
    lwd,
    contact: emp.email_address || '',
    manager: (empCode && MANAGER_FULL_NAME_OVERRIDES[empCode]) || emp.manager_name || '',
    isRayontara: String(emp.Is_rayontara).toLowerCase() === 'true',
    started: true,
    dues: {
      exam: { required: examDues.required ? 'yes_required' : 'not_required', amount: examDues.amount },
      wfh: { balance: wfhBalance },
      loanAdvance: { amount: 0 },
      dcTeam: { assetIssued: 'No' },
      noticePay: { applicable: false, date: '', amount: 0 },
    },
    comp: {
      grossSalary,
      epf,
      workingDays, presentDays, tenureYears, taBill: 0, nps: 0, vpf: 0, tds: 0, pli: 0,
    },
    settlement: {
      doj: emp.date_of_joining ? emp.date_of_joining.slice(0, 10) : '',
      earnings: { salaryCurrentMonth, salaryPreviousMonth, incentives: 0, taBillPositive: 0, gratuity, otherReimbursement: 0, noticePeriodPay: 0 },
      deductions: { pfCurrentMonth, pfPreviousMonth, examDues: examDues.required ? examDues.amount : 0, wfhDues: wfhBalance < 0 ? Math.abs(wfhBalance) : 0, nps: previousMonthDues.nps || 0, vpf: 0, taBillNegative: 0, pluxeeMeal: 0, esi: 0, tds: 0, pt: 0 },
      previousMonthDues,
      netPayableHistory: [], // lazy-loaded on demand — see GET /api/records/:id/netpay-history
    },
  });
}

// ---- Durable state ----
// PMS stays the source for everything about the employee. What is saved in the
// shared key-value store (lib/kv.js) is the state that passes between people:
// PLI, the TDS reply, acknowledgement tracking, and the workflow fields (personal
// email, send date, approval, payment, closure, DC Team). It is stored per
// employee code in small groups, so two people saving different things never
// overwrite each other. Amounts that feed the sheet (TDS, Incentives) are
// re-applied from these groups on load, so they survive restarts and agree
// across server instances.
const OVERLAY_VERSION = 'v1';
const OVERLAY_GROUPS = {
  pli: {
    read: (r) => r.pli,
    apply(r, v) {
      r.pli = v;
      if (v.status === 'submitted' && v.amount != null) { r.settlement.earnings.incentives = v.amount; r.comp.pli = v.amount; }
    },
  },
  tds: {
    read: (r) => r.tdsReply,
    apply(r, v) {
      r.tdsReply = v;
      if (v.status === 'applied' && v.amount != null) r.settlement.deductions.tds = v.amount;
    },
  },
  ack: { read: (r) => r.ackTracking, apply(r, v) { r.ackTracking = v; } },
  flow: {
    read: (r) => ({
      personalEmail: r.personalEmail, emailCc: r.emailCc, emailBcc: r.emailBcc, emailSentDate: r.emailSentDate, sheetApproval: r.sheetApproval,
      ack: r.ack, payment: r.payment, closure: r.closure, sla: r.sla, dcTeam: r.dues && r.dues.dcTeam,
    }),
    apply(r, v) {
      Object.assign(r, { personalEmail: v.personalEmail, emailCc: v.emailCc || '', emailBcc: v.emailBcc || '', emailSentDate: v.emailSentDate, sheetApproval: v.sheetApproval, ack: v.ack, payment: v.payment, closure: v.closure, sla: v.sla });
      if (v.dcTeam) r.dues.dcTeam = v.dcTeam;
    },
  },
};
const DEFAULT_JSON = (() => {
  const base = newRecord();
  return Object.fromEntries(Object.entries(OVERLAY_GROUPS).map(([g, d]) => [g, JSON.stringify(d.read(base))]));
})();
// Local runs write under 'dev', the deployed app under 'prod', even though both
// talk to the same database — so testing here never touches live data.
const KV_NAMESPACE = process.env.KV_NAMESPACE || (process.env.VERCEL ? 'prod' : 'dev');
const overlayKey = (rec, group) => `fnf:${KV_NAMESPACE}:${OVERLAY_VERSION}:${rec.empId}:${group}`;
const lastSeen = new Map(); // key -> JSON last read from / written to the store

async function persistRecord(rec) {
  if (!rec.empId) return;
  for (const [group, def] of Object.entries(OVERLAY_GROUPS)) {
    const json = JSON.stringify(def.read(rec));
    const key = overlayKey(rec, group);
    const seen = lastSeen.get(key);
    if (json === seen) continue;
    if (seen === undefined && json === DEFAULT_JSON[group]) continue; // untouched, nothing to store
    await kv.set(key, JSON.parse(json));
    lastSeen.set(key, json);
  }
}

async function refreshOverlays(records) {
  const slots = [];
  for (const rec of records) {
    if (!rec.empId) continue;
    for (const group of Object.keys(OVERLAY_GROUPS)) slots.push({ rec, group, key: overlayKey(rec, group) });
  }
  if (!slots.length) return;
  const values = await kv.mget(slots.map((s) => s.key));
  values.forEach((value, i) => {
    if (value == null) return;
    const { rec, group, key } = slots[i];
    const json = JSON.stringify(value);
    if (lastSeen.get(key) === json) return;
    OVERLAY_GROUPS[group].apply(rec, value);
    lastSeen.set(key, json);
  });
}

// Another server instance may have saved something since this one last looked.
const REFRESH_MS = 1500;
let lastRefreshAll = 0;
const lastRefreshOne = new Map();

async function freshen(records, one) {
  try {
    if (one) {
      if (Date.now() - (lastRefreshOne.get(one.id) || 0) < REFRESH_MS) return;
      lastRefreshOne.set(one.id, Date.now());
      await refreshOverlays([one]);
    } else {
      if (Date.now() - lastRefreshAll < REFRESH_MS) return;
      lastRefreshAll = Date.now();
      await refreshOverlays(records);
    }
  } catch (err) {
    console.error('Could not refresh saved state:', err.message); // serve what we have rather than fail the page
  }
}

let recordsPromise = null;

function loadRecords() {
  if (!recordsPromise) {
    recordsPromise = pms.getEmployeeDetails('')
      .then(async (employees) => {
        const resigned = employees.filter((e) => e.date_of_resigantion);
        // Resolved sequentially — the emp_code-by-email API appears to rate-limit
        // (returns a bare "Forbidden") when hit with many concurrent requests.
        const records = [];
        const taken = new Set();
        for (let i = 0; i < resigned.length; i++) {
          const rec = await fromPmsEmployee(resigned[i], i);
          if (taken.has(rec.id)) continue; // the same person listed twice
          taken.add(rec.id);
          records.push(rec);
        }
        try { await refreshOverlays(records); } catch (err) { console.error('Could not load saved state:', err.message); }
        lastRefreshAll = Date.now();
        return records;
      })
      .catch((err) => {
        console.error('Failed to load employees from PMS:', err.message);
        recordsPromise = null; // allow a retry on the next call
        return [];
      });
  }
  return recordsPromise;
}

function deepMerge(target, patch) {
  for (const key of Object.keys(patch)) {
    const val = patch[key];
    if (val && typeof val === 'object' && !Array.isArray(val) && target[key] && typeof target[key] === 'object') {
      deepMerge(target[key], val);
    } else {
      target[key] = val;
    }
  }
  return target;
}

async function listRecords() {
  const records = await loadRecords();
  await freshen(records);
  return records;
}

async function getRecord(id) {
  const records = await loadRecords();
  const rec = records.find((r) => r.id === Number(id)) || null;
  if (rec) await freshen(records, rec);
  return rec;
}

async function createRecord(seed) {
  const records = await loadRecords();
  const rec = newRecord(seed);
  records.push(rec);
  return rec;
}

async function updateRecord(id, patch) {
  const rec = await getRecord(id);
  if (!rec) return null;
  deepMerge(rec, patch);
  await persistRecord(rec); // a failed save is reported, never silently dropped
  return rec;
}

async function deleteRecord(id) {
  const records = await loadRecords();
  const idx = records.findIndex((r) => r.id === Number(id));
  if (idx === -1) return false;
  records.splice(idx, 1);
  return true;
}

module.exports = { listRecords, getRecord, createRecord, updateRecord, deleteRecord, newRecord, resolveNetPayableHistory, resolveTdsByMonth };
