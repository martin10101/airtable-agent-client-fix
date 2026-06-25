process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';

const fs = require('fs');
const path = require('path');
const { spawn, spawnSync } = require('child_process');
const express = require('express');

// Load .env manually (no dotenv dependency)
(function loadEnv() {
  const envPath = path.join(__dirname, '.env');
  if (!fs.existsSync(envPath)) {
    console.warn('[WARN] .env not found at', envPath, 'â€” copy .env.example to .env and fill in your keys.');
    return;
  }
  const envFile = fs.readFileSync(envPath, 'utf8');
  envFile.split(/\r?\n/).forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) return;
    const eq = trimmed.indexOf('=');
    if (eq === -1) return;
    const key = trimmed.slice(0, eq).trim();
    const val = trimmed.slice(eq + 1).trim();
    if (key && !(key in process.env)) process.env[key] = val;
  });
})();

const airtable = require('./airtable');
const agent = require('./agent');
const docxHandler = require('./docx-handler');
const smartDocx = require('./smart-docx');
const smartTemplateCopier = require('./smart-template-copier');
const xlsxHandler = require('./xlsx-handler');
const pdfHandler = require('./pdf-handler');
const projectRules = require('./project-rules');
const { findFolderForTips } = require('./find-tip-folder');

const APP_VERSION = '2026-06-25-smart-visual-copy-v23';
const PORT = parseInt(process.env.PORT || '3000', 10);
const TEMPLATE_FIELD = process.env.TEMPLATE_FIELD || 'Template Attachment';
const TEMPLATE_SELECT_FIELD = process.env.TEMPLATE_SELECT_FIELD || 'Template';
const OUTPUT_FIELD = process.env.OUTPUT_FIELD || 'Completed Document';
const OUTPUT_NAME_FIELD = process.env.OUTPUT_NAME_FIELD || 'Address';
const AIRTABLE_TABLE_NAME = process.env.AIRTABLE_TABLE_NAME;

const ROOT = __dirname;
const OUTPUT_DIR = (process.env.OUTPUT_FOLDER && process.env.OUTPUT_FOLDER.trim())
  ? process.env.OUTPUT_FOLDER.trim()
  : path.join(ROOT, 'output');
const ICAP_BOUNDARIES_FILE = (process.env.ICAP_BOUNDARIES_FILE && process.env.ICAP_BOUNDARIES_FILE.trim())
  ? process.env.ICAP_BOUNDARIES_FILE.trim()
  : path.join(ROOT, 'ICAP BOUNDARIES.xls');
const WORK_DIR = path.join(ROOT, 'work');
const LOG_DIR = path.join(ROOT, 'logs');
const TEMPLATES_DIR = (process.env.TEMPLATES_FOLDER && process.env.TEMPLATES_FOLDER.trim())
  ? process.env.TEMPLATES_FOLDER.trim()
  : path.join(ROOT, 'templates');

fs.mkdirSync(WORK_DIR, { recursive: true });
fs.mkdirSync(LOG_DIR, { recursive: true });

// Output folder: try to create it; if the parent doesn't exist (e.g. wrong username
// in .env, missing user profile), fail loudly instead of silently creating something
// in an unexpected place.
function ensureWritableDir(label, dir) {
  try {
    fs.mkdirSync(dir, { recursive: true });
    fs.accessSync(dir, fs.constants.W_OK);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.code === 'ENOENT'
      ? `parent path does not exist`
      : (e.code === 'EACCES' ? 'no write permission' : e.message) };
  }
}

const outputCheck = ensureWritableDir('OUTPUT_FOLDER', OUTPUT_DIR);

// Templates folder: this is the SOURCE â€” if it's a network drive (Z:\), we want a
// loud, clear warning when it's not mounted, not a silent empty fallback.
function checkTemplatesDir(dir) {
  if (!fs.existsSync(dir)) {
    const isNetworkDrive = /^[A-Z]:\\/i.test(dir) && !/^[Cc]:/.test(dir);
    return {
      ok: false,
      reason: isNetworkDrive
        ? `network drive not reachable (is ${dir.slice(0, 2)} mapped and online?)`
        : 'folder does not exist'
    };
  }
  try {
    const entries = fs.readdirSync(dir);
    return { ok: true, fileCount: entries.length };
  } catch (e) {
    return { ok: false, reason: `cannot read folder: ${e.message}` };
  }
}

const templatesCheck = checkTemplatesDir(TEMPLATES_DIR);

function log(...args) {
  const ts = new Date().toISOString();
  const line = `[${ts}] ${args.map((a) => (typeof a === 'string' ? a : JSON.stringify(a))).join(' ')}`;
  console.log(line);
  try { fs.appendFileSync(path.join(LOG_DIR, 'server.log'), line + '\n'); } catch (_) {}
}

function safeFilename(name) {
  if (name == null) return '';
  return String(name)
    .replace(/[<>:"/\\|?*\x00-\x1f]/g, '')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^[._-]+|[._-]+$/g, '')
    .slice(0, 120);
}

function pickOutputBaseName(fields, templateName, recordId) {
  const candidateNames = [
    'Property Address',
    'Project Address',
    'Address',
    OUTPUT_NAME_FIELD
  ].filter((name, index, arr) => name && arr.indexOf(name) === index);

  for (const fieldName of candidateNames) {
    const value = fields[fieldName];
    const text = projectRules.asString ? projectRules.asString(value).trim() : String(value || '').trim();
    if (!text) continue;
    const cleaned = safeFilename(text);
    if (cleaned) return cleaned;
  }
  const ext = path.extname(templateName);
  return `${path.basename(templateName, ext)}_${recordId}`;
}

// Recursive walk â€” returns relative paths like "421a/cover.docx" so users can
// organize 100+ templates into subfolders.
function walkTemplates(dir, base) {
  base = base || dir;
  const results = [];
  let entries;
  try {
    entries = fs.readdirSync(dir, { withFileTypes: true });
  } catch (_) {
    return results;
  }
  for (const entry of entries) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      if (entry.name.startsWith('.') || entry.name === 'node_modules') continue;
      results.push(...walkTemplates(full, base));
    } else if (/\.(docx|xlsx|pdf)$/i.test(entry.name)) {
      results.push(path.relative(base, full).split(path.sep).join('/'));
    }
  }
  return results.sort((a, b) => a.localeCompare(b));
}

function openInExplorer(targetPath, selectMode) {
  const args = selectMode ? ['/select,', targetPath] : [targetPath];
  const child = spawn('explorer.exe', args, { detached: true, stdio: 'ignore' });
  child.unref();
}

// Returns true if `filePath` either doesn't exist (we can create it) OR
// exists and is writable (no exclusive lock from Word/Excel/etc.).
function isWritable(filePath) {
  if (!fs.existsSync(filePath)) return true;
  let fd = null;
  try {
    fd = fs.openSync(filePath, 'r+'); // read+write, no truncate, no create
    return true;
  } catch (e) {
    if (e.code === 'EBUSY' || e.code === 'EPERM' || e.code === 'EACCES') return false;
    throw e;
  } finally {
    if (fd != null) { try { fs.closeSync(fd); } catch (_) {} }
  }
}

function ensureSmartTemplates(reason) {
  if (!templatesCheck.ok) return { created: [], skipped: [], warnings: [templatesCheck.reason] };
  const result = smartTemplateCopier.ensureSmartTemplateCopies(TEMPLATES_DIR, { log });
  if (result.created.length || result.warnings.length) {
    log(`[smart-template] ${reason}: created ${result.created.length}, skipped ${result.skipped.length}, warnings ${result.warnings.length}`);
    if (result.warnings.length) log('[smart-template] warnings:', result.warnings);
  }
  return result;
}

// Open a file with its default Windows app — Word for .docx, Excel for .xlsx,
// Adobe (or whichever PDF reader is set as default) for .pdf. Detached so the
// server doesn't block on the GUI process.
function openFile(filePath) {
  const safe = String(filePath).replace(/"/g, '\\"');
  const child = spawn(`start "" "${safe}"`, [], {
    detached: true,
    stdio: 'ignore',
    windowsHide: true,
    shell: true
  });
  child.unref();
}

function repairDocxWithWord(filePath, logFn) {
  const dlog = typeof logFn === 'function' ? logFn : (() => {});
  if (!/\.docx$/i.test(filePath)) return { skipped: true, reason: 'not a docx' };
  if (/^(1|true|yes)$/i.test(String(process.env.DISABLE_WORD_REPAIR || ''))) {
    return { skipped: true, reason: 'DISABLE_WORD_REPAIR is set' };
  }

  const scriptPath = path.join(WORK_DIR, 'word-repair-saveas.ps1');
  const script = `
param([Parameter(Mandatory=$true)][string]$Path)
$ErrorActionPreference = "Stop"
$word = $null
$doc = $null
$closed = $false
$quit = $false
$temp = [System.IO.Path]::Combine(
  [System.IO.Path]::GetDirectoryName($Path),
  ("~word-repair-" + [System.Guid]::NewGuid().ToString("N") + ".docx")
)
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  try { $word.AutomationSecurity = 3 } catch {}
  $missing = [System.Reflection.Missing]::Value
  $args = @($Path, $false, $false, $false, $missing, $missing, $false, $missing, $missing, $missing, $missing, $false, $true, $missing, $true, $missing)
  $doc = $word.Documents.Open.Invoke($args)
  if ($null -eq $doc) { throw "Word returned no document after repair open." }
  $doc.SaveAs2($temp, 12)
  $doc.Close($false)
  $closed = $true
  $word.Quit()
  $quit = $true
  if (!(Test-Path -LiteralPath $temp)) { throw "Word did not create repaired file: $temp" }
  Move-Item -LiteralPath $temp -Destination $Path -Force
  Write-Output ("OK " + (Split-Path -Leaf $Path))
} finally {
  if (($doc -ne $null) -and (-not $closed)) {
    try { $doc.Close($false) } catch {}
  }
  if (($word -ne $null) -and (-not $quit)) {
    try { $word.Quit() } catch {}
  }
  if ($doc -ne $null) {
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null } catch {}
  }
  if ($word -ne $null) {
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
  }
  if (($temp -ne $null) -and (Test-Path -LiteralPath $temp)) {
    try { Remove-Item -LiteralPath $temp -Force } catch {}
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
`.trim();

  try {
    fs.writeFileSync(scriptPath, script, 'utf8');
  } catch (e) {
    dlog('[WARN] Could not write Word repair script:', e.message);
    return { ok: false, error: e.message };
  }

  const result = spawnSync('powershell.exe', [
    '-NoProfile',
    '-ExecutionPolicy', 'Bypass',
    '-File', scriptPath,
    filePath
  ], {
    encoding: 'utf8',
    timeout: 90000,
    windowsHide: true
  });

  const stdout = String(result.stdout || '').trim();
  const stderr = String(result.stderr || '').trim();
  if (result.error) {
    dlog('[WARN] Word repair failed to run:', result.error.message);
    return { ok: false, error: result.error.message, stdout, stderr };
  }
  if (result.status !== 0) {
    dlog('[WARN] Word repair failed:', stderr || stdout || `exit ${result.status}`);
    return { ok: false, status: result.status, stdout, stderr };
  }
  dlog('[word-repair] Repaired/saved DOCX before opening:', stdout || path.basename(filePath));
  return { ok: true, stdout };
}

// If `targetPath` is locked (e.g. open in Word), pick the next available
// suffixed name: "name (2).docx", "name (3).docx", etc. Always returns a
// path we can definitely write to. Behavior is unchanged when the file
// isn't open — we just overwrite it.
function pickWritablePath(targetPath) {
  if (isWritable(targetPath)) return targetPath;
  const ext = path.extname(targetPath);
  const dir = path.dirname(targetPath);
  const base = path.basename(targetPath, ext);
  for (let i = 2; i <= 50; i++) {
    const candidate = path.join(dir, `${base} (${i})${ext}`);
    if (isWritable(candidate)) return candidate;
  }
  throw new Error(`Could not find a writable output path after 50 attempts (target was locked: ${targetPath})`);
}

const app = express();
app.use(express.json({ limit: '20mb' }));
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});

// -------- UI and static assets --------
app.get('/', (_req, res) => res.sendFile(path.join(ROOT, 'index.html')));
app.use('/files', express.static(OUTPUT_DIR));

// -------- Templates: list local folder (recursive; returns relative paths) --------
app.get('/templates', (_req, res) => {
  try {
    const _allFiles = walkTemplates(TEMPLATES_DIR).filter(f => {
      const base = f.split('/').pop();
      if (base.startsWith('~$')) return false;
      if (/[%,&()\[\]{}$;"'`]/.test(base)) return false;
      if (/[\r\n\t]/.test(f)) return false;
      if (f.length > 90) return false;
      if (/\.pdf$/i.test(base)) return false;
      return true;
    });
    const files = _allFiles
      .map(f => { try { return { p: f, m: fs.statSync(path.join(TEMPLATES_DIR, f)).mtimeMs }; } catch (_) { return { p: f, m: 0 }; } })
      .sort((a, b) => b.m - a.m)
      .slice(0, 999)
      .map(x => x.p);
    res.json({ folder: TEMPLATES_DIR, files });
  } catch (e) {
    res.status(500).json({ error: e.message, folder: TEMPLATES_DIR });
  }
});

// -------- One-click Airtable setup: create Template field + show button formulas --------
app.get('/setup', async (_req, res) => {
  try {
    if (!AIRTABLE_TABLE_NAME) throw new Error('AIRTABLE_TABLE_NAME not set in .env');
    const report = await airtable.autoSetupTable(AIRTABLE_TABLE_NAME, {
      templateSelectField: TEMPLATE_SELECT_FIELD,
      outputNameField: OUTPUT_NAME_FIELD,
      outputField: OUTPUT_FIELD
    });
    res.send(renderSetupPage(report));
  } catch (e) {
    log('[ERROR] /setup:', e.message);
    res.status(500).send(renderErrorPage('Setup failed', e.message));
  }
});

// -------- Sync the local file list into Airtable's single-select field --------
app.get('/sync-templates', async (_req, res) => {
  try {
    if (!AIRTABLE_TABLE_NAME) throw new Error('AIRTABLE_TABLE_NAME not set in .env');
    ensureSmartTemplates('/sync-templates');
    const _allFiles = walkTemplates(TEMPLATES_DIR).filter(f => {
      const base = f.split('/').pop();
      if (base.startsWith('~$')) return false;
      if (/[%,&()\[\]{}$;"'`]/.test(base)) return false;
      if (/[\r\n\t]/.test(f)) return false;
      if (f.length > 90) return false;
      if (/\.pdf$/i.test(base)) return false;
      return true;
    });
    const files = _allFiles
      .map(f => { try { return { p: f, m: fs.statSync(path.join(TEMPLATES_DIR, f)).mtimeMs }; } catch (_) { return { p: f, m: 0 }; } })
      .sort((a, b) => b.m - a.m)
      .slice(0, 999)
      .map(x => x.p);
    log(`/sync-templates: ${files.length} local files â†’ field "${TEMPLATE_SELECT_FIELD}"`);
    const result = await airtable.syncTemplateOptions(AIRTABLE_TABLE_NAME, TEMPLATE_SELECT_FIELD, files);
    log(`Sync done. Added: ${result.added.length}, orphaned: ${result.orphaned.length}`);
    res.send(renderSyncPage(TEMPLATES_DIR, files, result));
  } catch (e) {
    log('[ERROR] /sync-templates:', e.message);
    res.status(500).send(renderErrorPage('Sync failed', e.message));
  }
});

// -------- Record info: what Template is set on this record? --------
app.get('/record-info/:recordId', async (req, res) => {
  try {
    const recordId = req.params.recordId;
    if (!/^rec[A-Za-z0-9]{10,}$/.test(recordId)) {
      return res.status(400).json({ error: 'Invalid recordId' });
    }
    const record = await airtable.getRecord(recordId);
    const fields = record.fields || {};
    res.json({
      recordId,
      template: fields[TEMPLATE_SELECT_FIELD] || null,
      outputName: fields[OUTPUT_NAME_FIELD] || null,
      templateField: TEMPLATE_SELECT_FIELD,
      outputNameField: OUTPUT_NAME_FIELD
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// -------- Open templates folder in Windows Explorer --------
app.get('/open-templates', (_req, res) => {
  try {
    openInExplorer(TEMPLATES_DIR, false);
    res.send(renderOpenedPage('Templates folder', TEMPLATES_DIR));
  } catch (e) {
    res.status(500).send('Failed to open folder: ' + e.message);
  }
});

// -------- Open a specific output file's folder with the file selected --------
app.get('/open-output/:filename', (req, res) => {
  try {
    const safe = path.basename(req.params.filename);
    const full = path.join(OUTPUT_DIR, safe);
    if (!fs.existsSync(full)) return res.status(404).send('File not found');
    openInExplorer(full, true);
    res.send(renderOpenedPage('Output file', full));
  } catch (e) {
    res.status(500).send('Failed: ' + e.message);
  }
});

function renderSyncPage(folder, files, result) {
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const addedList = result.added.length
    ? '<ul>' + result.added.map((f) => `<li>${esc(f)}</li>`).join('') + '</ul>'
    : '<p style="color:#8a90a0">None â€” Airtable already has every local file.</p>';
  const orphanedList = result.orphaned.length
    ? '<ul>' + result.orphaned.map((f) => `<li>${esc(f)}</li>`).join('') + '</ul>'
    : '<p style="color:#8a90a0">None.</p>';
  return `<!doctype html>
<html><head><meta charset="utf-8"><title>Templates synced</title>
<style>body{font-family:system-ui,sans-serif;background:#0f1115;color:#e6e8ee;padding:32px;max-width:720px;margin:0 auto}
h2{margin-top:0}h3{margin-top:24px;color:#4f8cff}code{background:#1b1f28;padding:2px 6px;border-radius:4px}
ul{line-height:1.7}a{color:#4f8cff}</style></head><body>
<h2>Templates synced to Airtable</h2>
<p>Folder: <code>${esc(folder)}</code></p>
<p>Local files: <strong>${files.length}</strong> &middot; Airtable options now: <strong>${result.total}</strong></p>
<h3>Added to Airtable (${result.added.length})</h3>
${addedList}
<h3>Orphaned in Airtable (${result.orphaned.length})</h3>
<p style="color:#8a90a0;font-size:13px">These options exist in Airtable but have no matching file in the folder.
Airtable's API won't let us delete options that records are using â€” remove them manually if needed.</p>
${orphanedList}
<p style="margin-top:32px"><a href="/">&larr; Back to Agent</a></p>
<script>setTimeout(function(){ try{ window.close(); }catch(e){} }, 4000);</script>
</body></html>`;
}

function renderSetupPage(report) {
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const PORT_NUM = PORT;
  const buttons = [
    {
      label: 'Open Templates',
      desc: 'Opens the Windows folder where blank templates live.',
      formula: `"http://localhost:${PORT_NUM}/open-templates"`
    },
    {
      label: 'Refresh Templates',
      desc: 'Pushes the file list into the Template dropdown. Click after adding/removing template files.',
      formula: `"http://localhost:${PORT_NUM}/sync-templates"`
    },
    {
      label: 'Run Agent',
      desc: 'On a record where Template is set: fills the doc and uploads it back.',
      formula: `CONCATENATE(\n  "http://localhost:${PORT_NUM}/?recordId=", RECORD_ID(),\n  "&auto=1"\n)`
    }
  ];
  const rows = (label, items, color) => items.length
    ? `<h3 style="color:${color};margin-top:24px">${label}</h3><ul>${items.map((x) => `<li>${esc(x)}</li>`).join('')}</ul>`
    : '';
  const buttonBlocks = buttons.map((b, i) => `
    <div class="btn-card">
      <div class="btn-head">
        <span class="num">${i + 1}</span>
        <strong>${b.label}</strong>
        <span class="hint">â€” Single-line text? No. <b>Add a Button field</b> with action = Open URL</span>
      </div>
      <p>${b.desc}</p>
      <div class="formula-row">
        <pre id="f${i}">${esc(b.formula)}</pre>
        <button class="copy" onclick="navigator.clipboard.writeText(document.getElementById('f${i}').innerText).then(()=>this.textContent='Copied!')">Copy</button>
      </div>
    </div>`).join('');

  return `<!doctype html>
<html><head><meta charset="utf-8"><title>Setup â€” ${esc(report.tableName)}</title>
<style>
body{font-family:system-ui,sans-serif;background:#0f1115;color:#e6e8ee;padding:32px;max-width:880px;margin:0 auto;line-height:1.5}
h2{margin:0 0 4px;color:#7be08a}
h3{font-size:16px}
ul{margin:8px 0;padding-left:20px}
li{margin:4px 0}
.btn-card{background:#1b1f28;border:1px solid #2a2f3c;border-radius:8px;padding:16px;margin:14px 0}
.btn-head{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.num{background:#4f8cff;color:#fff;width:24px;height:24px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-weight:bold;font-size:14px}
.hint{color:#9aa1b3;font-size:13px;font-weight:normal}
pre{background:#0f1115;padding:10px;border-radius:6px;flex:1;margin:0;white-space:pre-wrap;font-family:Consolas,monospace;font-size:13px;color:#e6e8ee}
.formula-row{display:flex;gap:8px;align-items:flex-start;margin-top:8px}
.copy{background:#4f8cff;color:#fff;border:none;padding:8px 14px;border-radius:6px;cursor:pointer;font-size:13px;height:38px;flex-shrink:0}
.copy:hover{background:#6fa0ff}
a{color:#4f8cff}
.banner{background:#143323;border:1px solid #2d6b48;color:#7be08a;padding:12px;border-radius:6px;margin:16px 0}
</style></head><body>
<h2>Airtable setup â€” ${esc(report.tableName)}</h2>
<p style="color:#9aa1b3">Table id: <code>${esc(report.tableId)}</code></p>

<div class="banner">Auto-setup ran. Below is what changed and what's left for you to add by hand (Airtable's API can't create button fields â€” you'll have to add those 3 yourself).</div>

${rows('Created automatically', report.created, '#7be08a')}
${rows('Already correct', report.verified, '#4f8cff')}
${rows('Warnings', report.warnings, '#ffaa3b')}

<h3 style="margin-top:32px;color:#e6e8ee">Now add these 3 buttons to the table</h3>
<p>For each one, click <code>+</code> on the rightmost column â†’ choose <b>Button</b> â†’ action <b>Open URL</b> â†’ paste the formula. Hit "Copy" to grab it:</p>
${buttonBlocks}

<p style="margin-top:32px"><a href="/">â† Back to Agent</a></p>
</body></html>`;
}

function renderErrorPage(title, message) {
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  return `<!doctype html>
<html><head><meta charset="utf-8"><title>${esc(title)}</title>
<style>body{font-family:system-ui,sans-serif;background:#0f1115;color:#e6e8ee;padding:32px;max-width:720px;margin:0 auto}
h2{margin-top:0;color:#ff5c5c}pre{background:#1b1f28;padding:12px;border-radius:6px;white-space:pre-wrap;word-break:break-word}
a{color:#4f8cff}</style></head><body>
<h2>${esc(title)}</h2>
<pre>${esc(message)}</pre>
<p><a href="/">&larr; Back to Agent</a></p>
</body></html>`;
}

function renderOpenedPage(what, loc) {
  const safeLoc = String(loc).replace(/</g, '&lt;');
  return `<!doctype html>
<html><head><meta charset="utf-8"><title>Opened</title>
<style>body{font-family:system-ui,sans-serif;background:#0f1115;color:#e6e8ee;text-align:center;padding:48px}
a{color:#4f8cff}code{background:#1b1f28;padding:4px 8px;border-radius:6px;display:inline-block;margin-top:8px}</style>
</head><body>
<h2>${what} opened</h2>
<p>Switch to Windows Explorer.</p>
<code>${safeLoc}</code>
<p style="margin-top:32px"><a href="/">&larr; Back to Agent</a></p>
<script>setTimeout(function(){ try{ window.close(); }catch(e){} }, 2500);</script>
</body></html>`;
}

// -------- Health --------
app.get('/health', (_req, res) => res.json({ ok: true, version: APP_VERSION, uptime: process.uptime(), templatesFolder: TEMPLATES_DIR }));

// -------- Background trigger ? fires /generate in background, auto-closing tab --------
app.get('/trigger/:recordId', (req, res) => {
  const recordId = req.params.recordId;
  if (!/^rec[A-Za-z0-9]{10,}$/.test(recordId)) {
    return res.status(400).send('Invalid recordId');
  }
  log(`/trigger ${recordId}`);
  fetch(`http://localhost:${PORT}/generate`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ recordId })
  }).catch((e) => log('[trigger] generate error:', e.message));
  res.send(`<!doctype html>
<html><head><meta charset="utf-8"><title>Agent running</title>
<style>body{font-family:system-ui,sans-serif;background:#0f1115;color:#e6e8ee;text-align:center;padding:48px}</style>
</head><body>
<h2 style="color:#7be08a">Agent running</h2>
<p>Switch back to Airtable. The filled file will appear in ~30 seconds.</p>
<p style="color:#9aa1b3;font-size:13px">This tab will close automatically.</p>
<script>
setTimeout(function(){ try{ window.close(); }catch(e){} }, 200);
setTimeout(function(){ try{ window.close(); }catch(e){} }, 1500);
setTimeout(function(){ try{ window.close(); }catch(e){} }, 3000);
</script>
</body></html>`);
});

// -------- Core: generate filled document --------
app.post('/generate', async (req, res) => {
  const started = Date.now();
  try {
    const recordId = req.body && req.body.recordId;
    let templateFilename = req.body && req.body.templateFilename;

    if (!recordId || !/^rec[A-Za-z0-9]{10,}$/.test(recordId)) {
      return res.status(400).json({ success: false, error: 'recordId is required (e.g. recXXXXXXXXXXXXXX)' });
    }

    log('=== /generate start ===', APP_VERSION, recordId, templateFilename || '(check record for Template field)');

    // Fetch record
    const record = await airtable.getRecord(recordId);
    const fields = record.fields || {};

    // Resolve template source.
    // Priority: (1) explicit request arg, (2) record's single-select field,
    // (3) record's attachment fallback.
    let templatePath, templateName, templateSource;
    if (!templateFilename && fields[TEMPLATE_SELECT_FIELD]) {
      templateFilename = fields[TEMPLATE_SELECT_FIELD];
      log(`Using template from record's "${TEMPLATE_SELECT_FIELD}" field:`, templateFilename);
    }

    if (templateFilename) {
      // Relative paths ok (e.g. "421a/cover.docx"). Reject any ".." escape.
      const relNorm = String(templateFilename).replace(/\\/g, '/');
      if (relNorm.split('/').some((seg) => seg === '..')) {
        throw new Error(`Invalid template path "${templateFilename}"`);
      }
      const candidate = path.resolve(TEMPLATES_DIR, relNorm);
      const templatesResolved = path.resolve(TEMPLATES_DIR);
      if (!candidate.startsWith(templatesResolved + path.sep) && candidate !== templatesResolved) {
        throw new Error(`Template path escapes the templates folder: "${templateFilename}"`);
      }
      if (!fs.existsSync(candidate)) {
        throw new Error(
          `Template "${templateFilename}" not found in ${TEMPLATES_DIR}. ` +
          `If you just added it, click "Refresh Templates" in Airtable.`
        );
      }
      templatePath = candidate;
      templateName = path.basename(relNorm);
      templateSource = 'local';
    } else {
      const attachments = fields[TEMPLATE_FIELD];
      if (!Array.isArray(attachments) || !attachments.length) {
        throw new Error(
          `No template selected. Set the "${TEMPLATE_SELECT_FIELD}" field on this record, ` +
          `or pass templateFilename in the request, ` +
          `or attach a file to the "${TEMPLATE_FIELD}" field as a fallback.`
        );
      }
      const dl = await airtable.downloadAttachment(attachments[0], WORK_DIR);
      templatePath = dl.filePath;
      templateName = dl.filename;
      templateSource = 'airtable-attachment';
      log('Using Airtable attachment:', templateName);
    }

    const ext = path.extname(templateName).toLowerCase();
    const baseName = pickOutputBaseName(fields, templateName, recordId);
    const outputFilename = `${baseName}${ext}`;
    let actualOutputDir = TEMPLATES_DIR;
    const _tipRaw = record && record.fields && record.fields.TIP;
    const _tipsArr = Array.isArray(_tipRaw) ? _tipRaw : (_tipRaw ? [_tipRaw] : []);
    const _sub = findFolderForTips(_tipsArr, TEMPLATES_DIR);
    if (_sub) actualOutputDir = path.join(TEMPLATES_DIR, _sub);
    fs.mkdirSync(actualOutputDir, { recursive: true });
    const targetPath = path.join(actualOutputDir, outputFilename);
    const outputPath = pickWritablePath(targetPath);
    if (outputPath !== targetPath) {
      log(`[INFO] ${outputFilename} appears to be open in another program. Writing to ${path.basename(outputPath)} instead.`);
    }
    log('Output dir (TIP-based): ' + actualOutputDir);
    log(`Output name: ${path.basename(outputPath)}  (base: "${baseName}" from field "${OUTPUT_NAME_FIELD}")`);

    // Dispatch by file type
    let swapSummary = null;
    let docxValidation = null;
    let wordRepair = null;
    let renderMode = ext === '.docx' ? 'legacy' : null;
    let filledTags = [];
    let missingTags = [];
    let templateWarnings = [];
    let validationWarnings = [];
    if (ext === '.docx') {
      const docxContext = docxHandler.extractDocxContext(templatePath);
      const templateText = docxContext.text;
      const projectFacts = projectRules.deriveProjectFacts(fields, {
        icapBoundariesFile: ICAP_BOUNDARIES_FILE
      });
      const strictTemplate = smartDocx.containsSmartTags(templatePath);
      renderMode = strictTemplate ? 'strict' : 'legacy';
      log(`Template text length: ${templateText.length} chars; marked yellow/red targets: ${docxContext.markedTargets.length}; render mode: ${renderMode}.`);
      log('Project rule facts:', projectFacts);
      if (projectFacts.icap && projectFacts.icap.isIcap && !projectFacts.icap.checked) {
        log('[WARN] ICAP term was not resolved:', projectFacts.icap.reason);
      }

      if (strictTemplate) {
        log('Strict smart DOCX mode: filling {{...}} tags only. Claude and legacy cleanup are skipped.');
        const tagBuild = projectRules.buildSmartDocxTags(fields, projectFacts, { log });
        validationWarnings = tagBuild.warnings || [];
        templateWarnings = smartDocx.inspectSmartTemplateText(templateText);
        if (templateWarnings.length) log('[WARN] Smart template warnings:', templateWarnings);
        if (validationWarnings.length) log('[WARN] Smart value warnings:', validationWarnings);

        const strictResult = smartDocx.renderSmartDocx(templatePath, tagBuild.values, outputPath, { log });
        filledTags = strictResult.filledTags;
        missingTags = strictResult.missingTags;
        if (missingTags.length) {
          validationWarnings = validationWarnings.concat(`Unresolved smart tags remain: ${missingTags.join(', ')}`);
          log('[WARN] Missing smart tags:', missingTags);
        }

        docxValidation = docxHandler.validateDocx(outputPath);
        if (docxValidation.ok) {
          log(`[docx-check] OK: ${docxValidation.checkedParts} XML part(s) checked; no invalid XML or paragraph-property order issue found.`);
        } else {
          validationWarnings = validationWarnings.concat(docxValidation.problems || []);
          log(`[WARN] DOCX validation found ${docxValidation.problems.length} issue(s) in ${path.basename(outputPath)}:`, docxValidation.problems);
        }
        swapSummary = {
          renderMode,
          filledTags,
          missingTags,
          templateWarnings,
          validationWarnings,
          docxValidation
        };
      } else {
        const deterministicSwaps = projectRules.buildTemplatePlaceholderSwaps(templateText, fields, projectFacts, { log });
        if (deterministicSwaps.length) {
          log(`Template placeholders produced ${deterministicSwaps.length} deterministic swaps`);
        }
        const schema = await airtable.getTableSchema(AIRTABLE_TABLE_NAME);
        log(`Schema has ${schema.length} fields. Calling Claude (swap mode)...`);
        const aiSwaps = await agent.mapDocxSwaps(templateText, fields, schema, {
          templateFieldName: TEMPLATE_FIELD,
          outputFieldName: OUTPUT_FIELD,
          templateSelectFieldName: TEMPLATE_SELECT_FIELD,
          projectFacts,
          markedTargets: docxContext.markedTargets,
          log
        });
        const seenSwaps = new Set();
        const swaps = [];
        for (const swap of deterministicSwaps.concat(aiSwaps)) {
          const key = String(swap.oldValue || '');
          if (!key || seenSwaps.has(key)) continue;
          seenSwaps.add(key);
          swaps.push(swap);
        }
        log(`Claude returned ${aiSwaps.length} swaps; total after deterministic template rules: ${swaps.length}`);
        const CFG_LABELS = ['Building Configuration', 'Project Details', 'Proposed Construction', 'Building Description', 'Project Description'];
        const cfgSwaps = swaps.filter((s) => CFG_LABELS.includes(s.fieldName));
        if (cfgSwaps.length) {
          log(`Building-config rewrites (${cfgSwaps.length}):`);
          for (const s of cfgSwaps) {
            log(`  ${s.fieldName}: "${String(s.oldValue || '').slice(0, 100)}" -> "${String(s.newValue || '').slice(0, 100)}"`);
          }
        } else {
          log('No building-config rewrites produced (neither AI configuration_claims nor labeled-line backstop fired).');
        }
        const result = docxHandler.fillDocxSwaps(templatePath, swaps, outputPath, {
          markerEvaluator: (paragraphText) => projectRules.evaluateTemplateMarkerText(paragraphText, fields, projectFacts),
          log
        });
        log(`Applied ${result.applied.length} swaps; ${result.missed.length} old values not found in doc`);
        if (result.missed.length) {
          log('Missed swaps (old value not located in doc):', result.missed.map((s) => s.oldValue));
        }
        let outputText = docxHandler.extractDocxText(outputPath);
        const cleanupPasses = [];
        for (let pass = 1; pass <= 3; pass++) {
          const cleanupSwaps = projectRules.buildPostGenerationCleanupSwaps(outputText, fields, projectFacts, { log });
          if (!cleanupSwaps.length) break;
          log(`Applying ${cleanupSwaps.length} post-generation cleanup swap(s), pass ${pass}`);
          const cleanupResult = docxHandler.fillDocxSwaps(outputPath, cleanupSwaps, outputPath, { log });
          log(`Cleanup pass ${pass} applied ${cleanupResult.applied.length}; missed ${cleanupResult.missed.length}`);
          cleanupPasses.push({ pass, ...cleanupResult });
          outputText = docxHandler.extractDocxText(outputPath);
          if (!cleanupResult.applied.length) break;
        }
        const cleanupResult = cleanupPasses.length ? cleanupPasses : null;
        const qualityWarnings = projectRules.inspectGeneratedDocxText(outputText, fields, projectFacts);
        if (qualityWarnings.length) {
          log('[WARN] Output quality warnings:', qualityWarnings);
        }
        docxValidation = docxHandler.validateDocx(outputPath);
        if (docxValidation.ok) {
          log(`[docx-check] OK: ${docxValidation.checkedParts} XML part(s) checked; no invalid XML or paragraph-property order issue found.`);
        } else {
          log(`[WARN] DOCX validation found ${docxValidation.problems.length} issue(s) in ${path.basename(outputPath)}:`, docxValidation.problems);
        }
        swapSummary = {
          renderMode,
          applied: result.applied,
          missed: result.missed,
          cleanup: cleanupResult,
          qualityWarnings,
          docxValidation
        };
      }
    } else if (ext === '.xlsx') {
      const workbookJson = await xlsxHandler.extractXlsxContent(templatePath);
      log(`Workbook JSON length: ${workbookJson.length} chars. Fetching table schema...`);
      const schema = await airtable.getTableSchema(AIRTABLE_TABLE_NAME);
      log(`Schema has ${schema.length} fields. Calling Claude (xlsx swap mode)...`);
      const swapsBySheet = await agent.mapXlsxSwaps(workbookJson, fields, schema, {
        templateFieldName: TEMPLATE_FIELD,
        outputFieldName: OUTPUT_FIELD,
        templateSelectFieldName: TEMPLATE_SELECT_FIELD
      });
      const totalProposed = Object.values(swapsBySheet).reduce((n, a) => n + a.length, 0);
      log(`Claude returned ${totalProposed} cell swaps across ${Object.keys(swapsBySheet).length} sheet(s)`);
      const result = await xlsxHandler.fillXlsxSwaps(templatePath, swapsBySheet, outputPath);
      log(`Applied ${result.applied.length}, mismatched ${result.mismatched.length}, skipped-formula ${result.skippedFormula.length}`);
      swapSummary = {
        applied: result.applied.map((s) => ({ fieldName: s.fieldName, oldValue: s.oldValue, newValue: s.newValue, count: 1, sheet: s.sheet, cellRef: s.cellRef })),
        missed: result.mismatched.map((s) => ({ fieldName: s.fieldName, oldValue: s.oldValue, sheet: s.sheet, cellRef: s.cellRef, note: 'cell value did not match Claude\'s oldValue' }))
      };
    } else if (ext === '.pdf') {
      const pdfInfo = await pdfHandler.extractPdfContent(templatePath);
      if (pdfInfo.type !== 'form') {
        throw new Error(
          'This PDF has no fillable form fields. Open it in Adobe Acrobat â†’ "Prepare Form" ' +
          'to add fields, then save and try again. (The agent will not scribble on a flat PDF.)'
        );
      }
      log(`PDF has ${pdfInfo.fields.length} fillable fields. Calling Claude...`);
      const formFields = await agent.mapPdfFormFields(pdfInfo.fields, fields, TEMPLATE_FIELD, OUTPUT_FIELD);
      const flatten = String(req.query.flatten || req.body.flatten || '').toLowerCase() === 'true';
      const result = await pdfHandler.fillPdfForm(templatePath, formFields, outputPath, { flatten });
      if (result.skipped.length) log('Skipped PDF fields:', result.skipped);
      log(`Filled ${result.filled.length} fields${flatten ? ' (flattened)' : ''}`);
    } else {
      throw new Error(`Unsupported template type: ${ext}. Expected .docx, .xlsx, or .pdf`);
    }

    log('Written to:', outputPath);
    if (ext === '.docx') {
      const shouldRunWordRepair = renderMode !== 'strict' || !docxValidation || !docxValidation.ok;
      if (shouldRunWordRepair) {
        wordRepair = repairDocxWithWord(outputPath, log);
      } else {
        wordRepair = { skipped: true, reason: 'strict DOCX passed validation' };
        log('[docx-check] Skipping Word repair because strict DOCX passed validation.');
      }
      if (wordRepair && wordRepair.ok) {
        const afterWordRepairCheck = docxHandler.validateDocx(outputPath);
        if (afterWordRepairCheck.ok) {
          docxValidation = afterWordRepairCheck;
          log(`[docx-check] OK after Word repair: ${afterWordRepairCheck.checkedParts} XML part(s) checked.`);
        } else {
          log('[WARN] DOCX validation after Word repair found issue(s):', afterWordRepairCheck.problems);
        }
      }
    }

    // Open the file directly in its default app (Word/Excel/Adobe) so the user
    // can review immediately. File stays in its folder; we no longer upload a
    // copy back to Airtable (per client request — file lives in Z:\... only).
    try { openFile(outputPath); }
    catch (e) { log('[WARN] Could not auto-open file', outputPath, '-', e.message); }

    const uploadedToAirtable = false;
    const uploadError = null;

    const elapsed = ((Date.now() - started) / 1000).toFixed(1);
    log(`=== /generate done in ${elapsed}s ===`);

    res.json({
      success: true,
      recordId,
      templateUsed: templateName,
      templateSource,
      filePath: outputPath,
      outputFilename: path.basename(outputPath),
      fileUrl: `http://localhost:${PORT}/files/${encodeURIComponent(path.basename(outputPath))}`,
      openFolderUrl: `http://localhost:${PORT}/open-output/${encodeURIComponent(path.basename(outputPath))}`,
      uploadedToAirtable,
      uploadError,
      renderMode,
      filledTags,
      missingTags,
      templateWarnings,
      validationWarnings,
      swaps: swapSummary,
      docxValidation,
      wordRepair,
      elapsedSeconds: Number(elapsed)
    });
  } catch (err) {
    log('[ERROR]', err.stack || err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

require('./pick-and-run')(app, { airtable, log, TEMPLATES_DIR, TEMPLATE_SELECT_FIELD, PORT });

app.listen(PORT, () => {
  ensureSmartTemplates('startup');
  console.log('');
  console.log('  Airtable Document Agent');
  console.log('  ------------------------');
  console.log(`  Version:    ${APP_VERSION}`);
  console.log(`  UI:         http://localhost:${PORT}`);
  if (templatesCheck.ok) {
    console.log(`  Templates:  ${TEMPLATES_DIR}  (${templatesCheck.fileCount} entries)`);
  } else {
    console.log(`  Templates:  ${TEMPLATES_DIR}`);
  }
  console.log(`  Output:     ${OUTPUT_DIR}`);
  console.log(`  Name by:    Airtable field "${OUTPUT_NAME_FIELD}"`);
  console.log(`  Template:   Airtable field "${TEMPLATE_SELECT_FIELD}" (single-select, synced from folder)`);
  console.log('');

  let problems = 0;
  if (!templatesCheck.ok) {
    console.log(`  [!] TEMPLATES_FOLDER problem: ${templatesCheck.reason}`);
    console.log(`      Fix: open .env, set TEMPLATES_FOLDER to a folder that exists,`);
    console.log(`           or map the network drive before starting the server.`);
    console.log('');
    problems++;
  }
  if (!outputCheck.ok) {
    console.log(`  [!] OUTPUT_FOLDER problem: ${outputCheck.error}`);
    console.log(`      Path: ${OUTPUT_DIR}`);
    console.log(`      Fix: open .env, set OUTPUT_FOLDER to a folder you can write to.`);
    console.log('');
    problems++;
  }
  if (!process.env.ANTHROPIC_API_KEY || !process.env.AIRTABLE_API_KEY) {
    console.log('  [!] Missing API keys. Edit .env before generating documents.');
    console.log('');
    problems++;
  }
  if (problems === 0) {
    console.log('  Ready. Open http://localhost:' + PORT + ' in your browser.');
    console.log('');
  } else {
    console.log(`  ${problems} issue(s) above. The server is running so you can`);
    console.log('  hit endpoints, but generation will fail until they are fixed.');
    console.log('');
  }
});
