const path = require('path');
const fs = require('fs');
const os = require('os');
const { exec } = require('child_process');
const { findFolderForTips } = require('./find-tip-folder');

// After finding TIP folder, descend into a Tmplt/Templates-like subfolder if one exists
function findTmpltSubfolder(parentPath) {
  let entries;
  try { entries = fs.readdirSync(parentPath, { withFileTypes: true }); }
  catch (_) { return null; }
  const subfolders = entries.filter(e => e.isDirectory()).map(e => e.name);
  const norm = (s) => String(s).toLowerCase().replace(/[\s\-_()]+/g, '');
  // Priority: exact "Tmplt"/"Templates" -> "new" variants -> any non-"old" template folder
  return (
    subfolders.find(s => ['tmplt','tmplts','template','templates'].includes(norm(s))) ||
    subfolders.find(s => { const n = norm(s); return (n.includes('tmplt') || n.includes('template')) && n.includes('new'); }) ||
    subfolders.find(s => { const n = norm(s); return (n.includes('tmplt') || n.includes('template')) && !n.includes('old'); }) ||
    null
  );
}

module.exports = function registerPickAndRun(app, deps) {
  const { airtable, log, TEMPLATES_DIR, TEMPLATE_SELECT_FIELD, PORT } = deps;

  app.get('/pick-and-run/:recordId', async (req, res) => {
    const recordId = req.params.recordId;
    if (!/^rec[A-Za-z0-9]{10,}$/.test(recordId)) return res.status(400).send('Invalid recordId');

    // Send auto-closing page IMMEDIATELY so the browser tab vanishes ? work happens in background
    res.send('<!doctype html><html><body style="background:#0f1115"><script>window.close();</script></body></html>');

    try {
      const record = await airtable.getRecord(recordId);
      const tipRaw = record.fields && record.fields.TIP;
      const tipsArr = Array.isArray(tipRaw) ? tipRaw : (tipRaw ? [tipRaw] : []);
      let startFolder = TEMPLATES_DIR;
      const sub = findFolderForTips(tipsArr, TEMPLATES_DIR);
      if (sub) startFolder = path.join(TEMPLATES_DIR, sub);
      // Descend into Tmplt subfolder if there is one
      const tmpltSub = findTmpltSubfolder(startFolder);
      if (tmpltSub) startFolder = path.join(startFolder, tmpltSub);
      log('/pick-and-run ' + recordId + ' TIP=' + tipsArr.join(',') + ' -> ' + startFolder);

      const stamp = Date.now();
      const tmpScript = path.join(os.tmpdir(), 'pick-' + stamp + '.ps1');
      const tmpOutput = path.join(os.tmpdir(), 'picked-' + stamp + '.txt');
      const tmpLog    = path.join(os.tmpdir(), 'picklog-' + stamp + '.txt');
      const esc = (s) => s.replace(/'/g, "''");
      const psScript = [
        "$logPath = '" + esc(tmpLog) + "'",
        "$outPath = '" + esc(tmpOutput) + "'",
        '"start: $(Get-Date -Format o)" | Out-File -FilePath $logPath -Encoding UTF8',
        'try {',
        '  Add-Type -AssemblyName System.Windows.Forms',
        '  $f = New-Object System.Windows.Forms.OpenFileDialog',
        "  $f.InitialDirectory = '" + esc(startFolder) + "'",
        "  $f.Filter = 'Documents (*.docx;*.xlsx;*.pdf)|*.docx;*.xlsx;*.pdf|All files (*.*)|*.*'",
        "  $f.Title = 'Pick a template'",
        '  $f.RestoreDirectory = $false',
        '  $f.Multiselect = $false',
        '  $result = $f.ShowDialog()',
        '  "result=$result fn=$($f.FileName)" | Out-File -FilePath $logPath -Encoding UTF8 -Append',
        "  if ($result -eq [System.Windows.Forms.DialogResult]::OK) {",
        '    $f.FileName | Out-File -FilePath $outPath -Encoding UTF8 -NoNewline',
        '  }',
        '} catch {',
        '  "ERROR: $($_.Exception.Message)" | Out-File -FilePath $logPath -Encoding UTF8 -Append',
        '}'
      ].join('\r\n');
      fs.writeFileSync(tmpScript, '\ufeff' + psScript, 'utf8');

      const cmdLine = 'start "PickTemplate" /WAIT powershell -ExecutionPolicy Bypass -File "' + tmpScript + '"';
      log('exec: ' + cmdLine);
      exec(cmdLine, { windowsHide: true, timeout: 5 * 60 * 1000 }, async (err) => {
        log('PS done. err=' + (err ? err.message : 'none'));

        if (fs.existsSync(tmpLog)) {
          const psLog = fs.readFileSync(tmpLog, 'utf8').replace(/^\ufeff/, '');
          psLog.split(/\r?\n/).forEach(line => { if (line) log('  PS: ' + line); });
          try { fs.unlinkSync(tmpLog); } catch (_) {}
        }

        let picked = '';
        if (fs.existsSync(tmpOutput)) {
          picked = fs.readFileSync(tmpOutput, 'utf8').replace(/^\ufeff/, '').trim();
          try { fs.unlinkSync(tmpOutput); } catch (_) {}
        }
        try { fs.unlinkSync(tmpScript); } catch (_) {}

        if (!picked) {
          log('User cancelled or dialog failed ? no action taken');
          return;
        }
        let toSave = picked;
        if (picked.toLowerCase().startsWith(TEMPLATES_DIR.toLowerCase())) {
          toSave = picked.slice(TEMPLATES_DIR.length).replace(/^[\\/]+/, '').replace(/\\/g, '/');
        }
        try {
          await airtable.updateRecordField(recordId, TEMPLATE_SELECT_FIELD, toSave);
          log('/pick-and-run wrote: ' + toSave);
          fetch('http://localhost:' + PORT + '/generate', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ recordId })
          }).catch((e) => log('[pick-and-run] generate error:', e.message));
        } catch (e) {
          log('[ERROR] write failed: ' + e.message);
        }
      });
    } catch (e) {
      log('[ERROR] /pick-and-run:', e.message);
    }
  });
};
