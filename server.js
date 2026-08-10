try {
  process.loadEnvFile(require('path').join(__dirname, '.env')); // Node 20.6+
} catch {
  // no .env present — fine while running on mock data
}

const http = require('http');
const fs = require('fs');
const path = require('path');
const { URL } = require('url');
const store = require('./lib/store');
const { calc, recStatus, recSla } = require('./lib/calc');
const pms = require('./lib/pms');

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
