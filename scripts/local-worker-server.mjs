/* ══════════════════════════════════════════════════════════
   Node.js Local Runner for worker-whatsapp
   جمعية إرث وحضارة بالقريات
═══════════════════════════════════════════════════════════ */
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// استيراد كود الـ Worker مباشرة
const workerModule = await import('../worker-whatsapp/src/index.js');
const worker = workerModule.default;

const PORT = process.env.PORT || 8787;

// قراءة مفتاح OpenWA من المتغيرات أو من ملف OpenWA المحلي تلقائياً
let apiKey = process.env.OPENWA_API_KEY || '';
const localOpenWaKey = path.join(process.env.USERPROFILE || 'C:\\Users\\huawei', 'OpenWA', 'data', '.api-key');
if (!apiKey && fs.existsSync(localOpenWaKey)) {
  apiKey = fs.readFileSync(localOpenWaKey, 'utf8').trim();
}

const envPath = path.join(__dirname, '..', 'worker-whatsapp', '.dev.vars');
if (!apiKey && fs.existsSync(envPath)) {
  const content = fs.readFileSync(envPath, 'utf8');
  const match = content.match(/OPENWA_API_KEY=["']?([^"'\r\n]+)/);
  if (match) apiKey = match[1];
}

const env = {
  OPENWA_BASE_URL: process.env.OPENWA_BASE_URL || 'http://127.0.0.1:2785',
  OPENWA_API_KEY: apiKey,
  OPENWA_SESSION_ID: process.env.OPENWA_SESSION_ID || '2831a802-6a68-4a62-a062-5d9a23736f93',
  DEFAULT_SESSION_ID: '2831a802-6a68-4a62-a062-5d9a23736f93',
  ALLOWED_ORIGIN: '*'
};

const server = http.createServer(async (req, res) => {
  try {
    const fullUrl = `http://${req.headers.host || '127.0.0.1:' + PORT}${req.url}`;
    
    // جمع بيانات الطلب إن كانت الطريقة POST / PUT
    let bodyBuffer = null;
    if (req.method !== 'GET' && req.method !== 'HEAD') {
      const chunks = [];
      for await (const chunk of req) {
        chunks.push(chunk);
      }
      bodyBuffer = Buffer.concat(chunks);
    }

    const headers = new Headers();
    for (const [k, v] of Object.entries(req.headers)) {
      if (Array.isArray(v)) {
        v.forEach(val => headers.append(k, val));
      } else if (v !== undefined) {
        headers.set(k, v);
      }
    }

    const cfRequest = new Request(fullUrl, {
      method: req.method,
      headers: headers,
      body: bodyBuffer && bodyBuffer.length > 0 ? bodyBuffer : undefined,
      duplex: 'half'
    });

    const cfResponse = await worker.fetch(cfRequest, env);

    res.statusCode = cfResponse.status;
    cfResponse.headers.forEach((val, key) => {
      res.setHeader(key, val);
    });

    const resArrayBuffer = await cfResponse.arrayBuffer();
    res.end(Buffer.from(resArrayBuffer));
  } catch (err) {
    console.error('[Local Server Error]:', err);
    res.statusCode = 500;
    res.setHeader('Content-Type', 'application/json');
    res.end(JSON.stringify({ success: false, error: 'Internal Server Error: ' + err.message }));
  }
});

server.listen(PORT, '127.0.0.1', () => {
  console.log(`[Worker Local Runner] running on http://127.0.0.1:${PORT}`);
  console.log(`[Worker Local Runner] forwarding to OpenWA at ${env.OPENWA_BASE_URL}`);
});
