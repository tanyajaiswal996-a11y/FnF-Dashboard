// Client for the PMS "Kites" API. Two distinct credential sets are used
// against the same host, each scoped to a different role:
//   - "Get Employee Details (PMS)"    -> employee profile + resignation data
//   - "Get Employee Code by Email"    -> emp_code lookup, keyed by email
// Access + device tokens are cached per-role in memory and only refreshed
// when a call comes back unauthorized, since the API gives no explicit TTL.

const { decryptAmount } = require('./pms-crypto');

const BASE_URL = process.env.PMS_API_BASE_URL || 'https://api.koenig-solutions.com';

// The API rejects requests carrying Node's default fetch User-Agent
// (undici) with a bare "Forbidden" — sending a browser/curl-style UA works.
const REQUEST_HEADERS = { 'Content-Type': 'application/json', 'User-Agent': 'curl/8.0.1' };

function createSession({ username, password, role, apiKey }) {
  let cachedTokens = null; // { accessToken, deviceToken }

  async function fetchToken() {
    const res = await fetch(`${BASE_URL}/api/Kites/Operator/GetToken`, {
      method: 'POST',
      headers: REQUEST_HEADERS,
      body: JSON.stringify({ userName: username, userPassword: password, userRole: role }),
    });
    const body = await res.json();
    if (body.statuscode !== 200 || !body.content) {
      throw new Error(`PMS GetToken failed (${role}): ${body.message || res.status}`);
    }
    cachedTokens = { accessToken: body.content.accessToken, deviceToken: body.content.deviceToken };
    return cachedTokens;
  }

  async function getTokens() {
    if (cachedTokens) return cachedTokens;
    return fetchToken();
  }

  async function callCommon(payload, tokens) {
    const url = `${BASE_URL}/api/Kites/Operator/common?apikey=${apiKey}&accessToken=${encodeURIComponent(tokens.accessToken)}&deviceToken=${encodeURIComponent(tokens.deviceToken)}`;
    const res = await fetch(url, {
      method: 'POST',
      headers: REQUEST_HEADERS,
      body: JSON.stringify(payload),
    });
    const body = await res.json();
    // `content` comes back as a JSON-encoded string rather than a parsed array.
    if (typeof body.content === 'string') {
      body.content = JSON.parse(body.content);
    }
    return { status: res.status, body };
  }

  // Calls the "common" endpoint with the given payload, retrying once with a
  // fresh token on ANY failure — not just 401. The API returns a bare
  // "Forbidden" (not 401) when a session gets invalidated, e.g. by a
  // concurrent login on the same account elsewhere.
  async function call(payload) {
    let tokens = await getTokens();
    let { status, body } = await callCommon(payload, tokens);

    if (body.statuscode !== 200) {
      tokens = await fetchToken();
      ({ status, body } = await callCommon(payload, tokens));
    }

    if (body.statuscode !== 200) {
      throw new Error(`PMS common call failed (${role}): ${body.message || status}`);
    }
    return body.content || [];
  }

  return { call };
}

const detailsSession = createSession({
  username: process.env.PMS_USERNAME,
  password: process.env.PMS_PASSWORD,
  role: process.env.PMS_ROLE,
  apiKey: process.env.PMS_API_KEY,
});

const codeSession = createSession({
  username: process.env.PMS_CODE_USERNAME,
  password: process.env.PMS_CODE_PASSWORD,
  role: process.env.PMS_CODE_ROLE,
  apiKey: process.env.PMS_CODE_API_KEY,
});

const exitListSession = createSession({
  username: process.env.PMS_EXIT_USERNAME,
  password: process.env.PMS_EXIT_PASSWORD,
  role: process.env.PMS_EXIT_ROLE,
  apiKey: process.env.PMS_EXIT_API_KEY,
});

const appraisalSession = createSession({
  username: process.env.PMS_APPRAISAL_USERNAME,
  password: process.env.PMS_APPRAISAL_PASSWORD,
  role: process.env.PMS_APPRAISAL_ROLE,
  apiKey: process.env.PMS_APPRAISAL_API_KEY,
});

const netPayableSession = createSession({
  username: process.env.PMS_NETPAY_USERNAME,
  password: process.env.PMS_NETPAY_PASSWORD,
  role: process.env.PMS_NETPAY_ROLE,
  apiKey: process.env.PMS_NETPAY_API_KEY,
});

const wfhSession = createSession({
  username: process.env.PMS_WFH_USERNAME,
  password: process.env.PMS_WFH_PASSWORD,
  role: process.env.PMS_WFH_ROLE,
  apiKey: process.env.PMS_WFH_API_KEY,
});

const examSession = createSession({
  username: process.env.PMS_EXAM_USERNAME,
  password: process.env.PMS_EXAM_PASSWORD,
  role: process.env.PMS_EXAM_ROLE,
  apiKey: process.env.PMS_EXAM_API_KEY,
});

// Fetches employee details for a given emp_code ('' = all employees).
async function getEmployeeDetails(empCode) {
  return detailsSession.call({ emp_code: empCode || '' });
}

// Looks up { EmpCode, EmployeeName, OfficialEmail, UserId } for one email.
// Returns null if the email has no match.
async function getEmployeeCodeByEmail(email) {
  const results = await codeSession.call({ OfficialEmail: email });
  return results.find((r) => r.EmpCode != null) || null;
}

// Lists { EmpCode, Name, LWD, 'Deptt.', DOR, BaseLocation } for every exit
// with a DOR in [fromDate, toDate] (YYYY-MM-DD).
async function getExitEmployeeList(fromDate, toDate) {
  return exitListSession.call({ FromDate: fromDate || '', ToDate: toDate || '' });
}

// Looks up { Amount, Currency, EPF, AllowNPS, EmployeeShare, EmployerShare }
// for one emp_code. Returns null if there's no appraisal record for them.
// `Amount` comes back AES-encrypted from the API — decrypted here so callers
// always see the plain gross salary figure.
async function getAppraisalData(empId) {
  const results = await appraisalSession.call({ EmpId: empId || '' });
  const match = results.find((r) => r.Amount != null);
  if (!match) return null;
  return { ...match, Amount: decryptAmount(match.Amount) };
}

// Looks up payroll details (Leave_BF, Payable_Days, PF, TDS, Total_NPS, Salary,
// etc.) for one emp_code over [startDate, endDate] (YYYY-MM-DD). Returns null
// if there's no payroll run for that period.
async function getNetPayableDetails(empId, startDate, endDate) {
  const results = await netPayableSession.call({ EmpCode: empId || '', StartDate: startDate || '', EndDate: endDate || '' });
  const match = results.find((r) => r.Salary != null);
  return match || null;
}

// Looks up the WFH infra balance for one emp_code. Returns null if the API
// gives no record (rather than 0) so callers can distinguish "unknown" from
// "zero balance".
async function getWfhBalance(empId) {
  const results = await wfhSession.call({ EmpCode: empId || '' });
  const match = results.find((r) => r.WFHinfrabalance != null);
  return match ? Number(match.WFHinfrabalance) : null;
}

// Looks up pending exam dues for one emp_id. Each exam/course entry carries
// its own RecoverFees flag ("Yes Required" / "Not Required") and ExamCost —
// only entries flagged "Yes Required" count toward the amount owed.
async function getExamPendingDues(empId) {
  const results = await examSession.call({ EmpId: empId || '' });
  const recoverable = (results || []).filter((r) => String(r.RecoverFees).trim().toLowerCase() === 'yes required');
  if (!recoverable.length) return { required: false, amount: 0 };
  const amount = recoverable.reduce((sum, r) => sum + (Number(r.ExamCost) || 0), 0);
  return { required: true, amount };
}

module.exports = { getEmployeeDetails, getEmployeeCodeByEmail, getExitEmployeeList, getAppraisalData, getNetPayableDetails, getWfhBalance, getExamPendingDues };
