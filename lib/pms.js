// Client for the PMS "Kites" API — fetches employee resignation/notice-period
// data. Access + device tokens are cached in memory and only refreshed when a
// call comes back unauthorized, since the API gives no explicit token TTL.

const BASE_URL = process.env.PMS_API_BASE_URL || 'https://api.koenig-solutions.com';
const USERNAME = process.env.PMS_USERNAME;
const PASSWORD = process.env.PMS_PASSWORD;
const ROLE = process.env.PMS_ROLE;
const API_KEY = process.env.PMS_API_KEY;

let cachedTokens = null; // { accessToken, deviceToken }

// The API rejects requests carrying Node's default fetch User-Agent
// (undici) with a bare "Forbidden" — sending a browser/curl-style UA works.
const REQUEST_HEADERS = { 'Content-Type': 'application/json', 'User-Agent': 'curl/8.0.1' };

async function fetchToken() {
  const res = await fetch(`${BASE_URL}/api/Kites/Operator/GetToken`, {
    method: 'POST',
    headers: REQUEST_HEADERS,
    body: JSON.stringify({ userName: USERNAME, userPassword: PASSWORD, userRole: ROLE }),
  });
  const body = await res.json();
  if (body.statuscode !== 200 || !body.content) {
    throw new Error(`PMS GetToken failed: ${body.message || res.status}`);
  }
  cachedTokens = { accessToken: body.content.accessToken, deviceToken: body.content.deviceToken };
  return cachedTokens;
}

async function getTokens() {
  if (cachedTokens) return cachedTokens;
  return fetchToken();
}

async function callEmployeeDetails(empCode, tokens) {
  const url = `${BASE_URL}/api/Kites/Operator/common?apikey=${API_KEY}&accessToken=${encodeURIComponent(tokens.accessToken)}&deviceToken=${encodeURIComponent(tokens.deviceToken)}`;
  const res = await fetch(url, {
    method: 'POST',
    headers: REQUEST_HEADERS,
    body: JSON.stringify({ emp_code: empCode || '' }),
  });
  const body = await res.json();
  // `content` comes back as a JSON-encoded string rather than a parsed array.
  if (typeof body.content === 'string') {
    body.content = JSON.parse(body.content);
  }
  return { status: res.status, body };
}

// Fetches employee details for a given emp_code ('' = all employees), retrying
// once with a fresh token if the cached one has been rejected.
async function getEmployeeDetails(empCode) {
  let tokens = await getTokens();
  let { status, body } = await callEmployeeDetails(empCode, tokens);

  if (status === 401 || body.statuscode === 401) {
    tokens = await fetchToken();
    ({ status, body } = await callEmployeeDetails(empCode, tokens));
  }

  if (body.statuscode !== 200) {
    throw new Error(`PMS GetEmployeeDetails failed: ${body.message || status}`);
  }
  return body.content || [];
}

module.exports = { getEmployeeDetails };
