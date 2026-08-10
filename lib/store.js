// Data access layer for FnF records.
//
// Backed by live PMS data (lib/pms.js) — on first access this fetches every
// employee with a recorded resignation date and seeds one FnF record per
// person. The resulting array is cached in memory for the life of the
// process; edits made in the app (checklist, comp figures, etc.) mutate that
// cached array directly. Swap loadRecords() for a real DB read when Turso
// access is granted.

const pms = require('./pms');

let nextId = 1;

function newRecord(seed) {
  seed = seed || {};
  const base = {
    id: nextId++,
    empId: '', name: '', department: '', designation: '', location: '', dor: '', lwd: '', contact: '',
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
      noticePay: { applicable: false, date: '', amount: 0 },
    },
    fnfDue: { pendingFrom: 'none' },
    comp: { grossSalary: 0, epf: 0, workingDays: 26, presentDays: 26, tenureYears: 0, taBill: 0, nps: 0, vpf: 0, tds: 0, pli: 0 },
    ack: { shared: false, acknowledged: false },
    emailSentDate: '',
    payment: { apInformed: false, apConfirmed: false },
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

function fromPmsEmployee(emp, idx) {
  const name = [emp.first_name, emp.middle_name, emp.last_name].filter(Boolean).join(' ').replace(/\s+/g, ' ').trim();
  return newRecord({
    empId: `PMS${String(idx + 1).padStart(4, '0')}`,
    name,
    department: emp.deparment_name || '',
    designation: emp.designation_name || '',
    location: emp.city_name || '',
    dor: emp.date_of_resigantion ? emp.date_of_resigantion.slice(0, 10) : '',
    lwd: emp.last_working_day ? emp.last_working_day.slice(0, 10) : '',
    contact: emp.email_address || '',
    manager: emp.manager_name || '',
    started: true,
    comp: { grossSalary: 0, epf: 0, workingDays: 26, presentDays: 26, tenureYears: tenureYearsFromJoining(emp.date_of_joining), taBill: 0, nps: 0, vpf: 0, tds: 0, pli: 0 },
  });
}

let recordsPromise = null;

function loadRecords() {
  if (!recordsPromise) {
    recordsPromise = pms.getEmployeeDetails('')
      .then((employees) => employees.filter((e) => e.date_of_resigantion).map(fromPmsEmployee))
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

module.exports = { listRecords, getRecord, createRecord, updateRecord, deleteRecord, newRecord };
