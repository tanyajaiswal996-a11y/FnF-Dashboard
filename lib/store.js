// Data access layer for FnF records.
//
// Backed by live PMS data (lib/pms.js) — on first access this fetches every
// employee with a recorded resignation date and seeds one FnF record per
// person. The resulting array is cached in memory for the life of the
// process; edits made in the app (checklist, comp figures, etc.) mutate that
// cached array directly. Swap loadRecords() for a real DB read when Turso
// access is granted.

const pms = require('./pms');
const { computeWorkingPresentDays, computeSalarySplit, computeGratuity, tdsCalcStartIndex, monthIndex, isTdsApplicable } = require('./calc');

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
    empId: '', name: '', department: '', designation: '', location: '', dor: '', lwd: '', contact: '', personalEmail: '',
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
    pli: { status: 'not_shared', token: '', sharedAt: '', amount: null, submittedAt: '' },
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
  return Object.assign(base, seed, { id: base.id });
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

async function resolvePreviousMonthDues(empCode, lwd) {
  const range = previousMonthRange(lwd, 1);
  if (!empCode || !range) return { month: '', tds: 0, nps: 0 };
  try {
    const details = await pms.getNetPayableDetails(empCode, range.start, range.end);
    if (!details) return { month: '', tds: 0, nps: 0 };
    return { month: range.label, tds: Number(details.TDS) || 0, nps: Number(details.Total_NPS) || 0 };
  } catch (err) {
    console.error(`Failed to resolve previous-month dues for emp_code ${empCode}:`, err.message);
    return { month: '', tds: 0, nps: 0 };
  }
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
  const { salaryCurrentMonth: pfCurrentMonth, salaryPreviousMonth: pfPreviousMonth } = computeSalarySplit(epf, dor, lwd, workingDays, presentDays);
  const tenureYears = tenureYearsFromJoining(emp.date_of_joining);
  const gratuity = computeGratuity(grossSalary, tenureYears);
  return newRecord({
    empId: empCode || emp.email_address || `PMS${String(idx + 1).padStart(4, '0')}`,
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

let recordsPromise = null;

function loadRecords() {
  if (!recordsPromise) {
    recordsPromise = pms.getEmployeeDetails('')
      .then(async (employees) => {
        const resigned = employees.filter((e) => e.date_of_resigantion);
        // Resolved sequentially — the emp_code-by-email API appears to rate-limit
        // (returns a bare "Forbidden") when hit with many concurrent requests.
        const records = [];
        for (let i = 0; i < resigned.length; i++) {
          records.push(await fromPmsEmployee(resigned[i], i));
        }
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
  return loadRecords();
}

async function getRecord(id) {
  const records = await loadRecords();
  return records.find((r) => r.id === Number(id)) || null;
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
