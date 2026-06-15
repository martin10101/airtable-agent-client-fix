// Connectivity diagnostic — run via check-environment.bat.
// Tests whether this PC can reach every endpoint the agent depends on.
// Tolerates corporate SSL inspection (sets NODE_TLS_REJECT_UNAUTHORIZED=0).

process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';

const fs = require('fs');
const path = require('path');

const endpoints = [
  { name: 'npm registry          ', url: 'https://registry.npmjs.org/',         mustWork: 'to install dependencies' },
  { name: 'Anthropic API         ', url: 'https://api.anthropic.com/v1/models', mustWork: 'for Claude AI calls',          expectAuth: true },
  { name: 'Airtable API          ', url: 'https://api.airtable.com/v0/meta/bases', mustWork: 'to read records',            expectAuth: true },
  { name: 'Airtable Content API  ', url: 'https://content.airtable.com/',       mustWork: 'to upload completed documents' }
];

function pad(s, n) {
  s = String(s);
  return s.length >= n ? s : s + ' '.repeat(n - s.length);
}

async function testEndpoint(ep) {
  const started = Date.now();
  try {
    const res = await fetch(ep.url, { method: 'GET' });
    const ms = Date.now() - started;
    // Any HTTP response means we reached the server. 401/403 is expected
    // for unauthenticated calls to auth-required endpoints.
    const ok = res.status < 500;
    return { ok, status: res.status, ms };
  } catch (e) {
    const ms = Date.now() - started;
    return { ok: false, error: e.message, ms };
  }
}

function checkProxy() {
  const vars = ['HTTP_PROXY', 'HTTPS_PROXY', 'http_proxy', 'https_proxy', 'NO_PROXY', 'no_proxy'];
  const set = {};
  for (const v of vars) {
    if (process.env[v]) set[v] = process.env[v];
  }
  return set;
}

function checkWritePermission() {
  try {
    const testFile = path.join(__dirname, '.write-test-' + Date.now());
    fs.writeFileSync(testFile, 'ok');
    fs.unlinkSync(testFile);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function checkNodeVersion() {
  const v = process.versions.node;
  const major = parseInt(v.split('.')[0], 10);
  return { version: v, ok: major >= 18, recommended: major >= 22 };
}

(async () => {
  console.log('');
  console.log('  Running diagnostics...');
  console.log('');

  const nodeInfo = checkNodeVersion();
  const nodeMark = nodeInfo.recommended ? '[OK]' : nodeInfo.ok ? '[WARN]' : '[FAIL]';
  console.log(`  ${nodeMark} Node.js version:          ${nodeInfo.version}` +
    (nodeInfo.recommended ? '' : nodeInfo.ok ? '  (works, but 22 LTS is recommended)' : '  (need 18+)'));

  const writeInfo = checkWritePermission();
  const writeMark = writeInfo.ok ? '[OK]' : '[FAIL]';
  console.log(`  ${writeMark} Write permission here:    ${writeInfo.ok ? 'yes' : 'no — ' + writeInfo.error}`);

  console.log('');
  console.log('  Network endpoints:');

  let networkFailures = [];
  for (const ep of endpoints) {
    const r = await testEndpoint(ep);
    const mark = r.ok ? '[OK]' : '[FAIL]';
    let line = `  ${mark} ${ep.name} ${pad(ep.url, 45)}`;
    if (r.ok) line += ` HTTP ${r.status}  (${r.ms} ms)`;
    else line += ` ${r.error || 'HTTP ' + r.status}`;
    console.log(line);
    if (!r.ok) networkFailures.push(ep);
  }

  const proxy = checkProxy();
  console.log('');
  if (Object.keys(proxy).length) {
    console.log('  Proxy env vars detected:');
    for (const [k, v] of Object.entries(proxy)) console.log(`     ${k} = ${v}`);
  } else {
    console.log('  No proxy env vars set.');
    console.log('  (If the endpoints above failed, IT may require a proxy — ask them for the URL.)');
  }

  console.log('');
  const anyFail = !nodeInfo.ok || !writeInfo.ok || networkFailures.length > 0;

  if (anyFail) {
    console.log('  Some checks failed. What to do:');
    if (!nodeInfo.ok) {
      console.log('   * Install Node.js 22 LTS from https://nodejs.org');
    }
    if (!writeInfo.ok) {
      console.log('   * Move this folder to a location where you have write permission (e.g. Documents or H:\\).');
    }
    for (const ep of networkFailures) {
      console.log(`   * Cannot reach ${ep.url.trim()} — needed ${ep.mustWork}.`);
      console.log('     Ask IT to whitelist this URL, or provide a proxy URL.');
    }
    console.log('');
    process.exitCode = 1;
  } else {
    console.log('  All checks passed — you are ready to run first-time-setup.bat');
    console.log('');
  }
})();
