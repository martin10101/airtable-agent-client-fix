const fs = require('fs');
const path = require('path');

const norm = (s) => String(s).toLowerCase().replace(/[\s\-_()/]+/g, '');

// tip-folder-map.json supports two rule types:
//   {
//     "exact":    { "<exact TIP value>": "<subfolder>" }   // matches one TIP value verbatim (case/punct insensitive)
//     "contains": { "<substring>":      "<subfolder>" }   // matches if ANY TIP value in the array contains this substring
//   }
// `contains` rules win over `exact` rules and the fuzzy fallback. They scan
// across all TIP values, so a record tagged ["ICAP", "421a New Construction"]
// still routes to the 421a folder when "421a" is configured under contains.
//
// Backward compat: if the JSON is a flat object (not wrapped with exact/contains),
// the whole thing is treated as the exact-match map.
let EXACT = {};
let CONTAINS = {};
try {
  const mapPath = path.join(__dirname, 'tip-folder-map.json');
  if (fs.existsSync(mapPath)) {
    const raw = JSON.parse(fs.readFileSync(mapPath, 'utf8')) || {};
    const isStructured =
      (raw.exact && typeof raw.exact === 'object') ||
      (raw.contains && typeof raw.contains === 'object');
    if (isStructured) {
      EXACT = raw.exact || {};
      CONTAINS = raw.contains || {};
    } else {
      EXACT = raw;
    }
  }
} catch (e) {
  console.warn('[WARN] tip-folder-map.json malformed; ignoring overrides:', e.message);
}

function folderExists(templatesDir, folder) {
  try {
    return fs.statSync(path.join(templatesDir, folder)).isDirectory();
  } catch (_) {
    return false;
  }
}

function findTipFolder(tip, templatesDir) {
  if (!tip) return null;

  for (const [k, v] of Object.entries(EXACT)) {
    if (norm(k) === norm(tip) && folderExists(templatesDir, v)) return v;
  }

  const tipN = norm(tip);
  if (!tipN) return null;
  let entries;
  try { entries = fs.readdirSync(templatesDir, { withFileTypes: true }); }
  catch (_) { return null; }
  const folders = entries.filter(e => e.isDirectory()).map(e => e.name);
  return (
    folders.find(f => norm(f) === tipN) ||
    folders.find(f => norm(f).startsWith(tipN)) ||
    folders.find(f => norm(f).endsWith(tipN)) ||
    folders.find(f => norm(f).includes(tipN)) ||
    null
  );
}

function findFolderForTips(tips, templatesDir) {
  const arr = Array.isArray(tips) ? tips : (tips ? [tips] : []);
  if (!arr.length) return null;

  for (const [needle, folder] of Object.entries(CONTAINS)) {
    const needleN = norm(needle);
    if (!needleN) continue;
    const hit = arr.some(t => norm(t).includes(needleN));
    if (hit && folderExists(templatesDir, folder)) return folder;
  }

  for (const t of arr) {
    const sub = findTipFolder(t, templatesDir);
    if (sub) return sub;
  }
  return null;
}

module.exports = { findTipFolder, findFolderForTips };
