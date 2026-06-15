const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const BOROUGHS = {
  1: 'Manhattan',
  2: 'Bronx',
  3: 'Brooklyn',
  4: 'Queens',
  5: 'Staten Island'
};

const BOROUGH_CODE_BY_NAME = {
  manhattan: 1,
  newyork: 1,
  nyc: 1,
  bronx: 2,
  brooklyn: 3,
  kings: 3,
  queens: 4,
  statenisland: 5,
  richmond: 5
};

const ICAP_PAIR_COLUMNS = {
  Manhattan: [[0, 2], [4, 6]],
  Bronx: [[0, 2]],
  Brooklyn: [[0, 2]],
  Queens: [[0, 2]],
  'Staten Island': [[0, 1], [4, 6]]
};

let icapCache = null;

function getField(fields, name) {
  if (!fields) return undefined;
  const lower = String(name).toLowerCase().trim();
  for (const key of Object.keys(fields)) {
    if (String(key).toLowerCase().trim() === lower) return fields[key];
  }
  return undefined;
}

function asString(value) {
  if (value == null) return '';
  if (typeof value === 'string') return value;
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  if (Array.isArray(value)) {
    return value.map(asString).filter(Boolean).join(', ');
  }
  if (typeof value === 'object') {
    return String(value.name || value.value || value.text || value.label || '');
  }
  return String(value);
}

function asArray(value) {
  if (value == null || value === '') return [];
  return Array.isArray(value) ? value : [value];
}

function parseNumber(value) {
  if (value == null || value === '') return null;
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  const text = asString(value).replace(/[$,\s]/g, '');
  if (!text) return null;
  const n = Number(text);
  return Number.isFinite(n) ? n : null;
}

function normalizeBorough(value) {
  const text = asString(value).trim();
  if (!text) return null;
  const numeric = Number(text);
  if (Number.isInteger(numeric) && BOROUGHS[numeric]) {
    return { code: numeric, name: BOROUGHS[numeric] };
  }
  const key = text.toLowerCase().replace(/[^a-z]/g, '');
  const code = BOROUGH_CODE_BY_NAME[key];
  if (!code) return null;
  return { code, name: BOROUGHS[code] };
}

function normalizeBlock(value) {
  if (value == null || value === '') return null;
  if (typeof value === 'number') {
    return Number.isFinite(value) ? String(Math.trunc(value)) : null;
  }
  const text = asString(value).trim();
  if (!text) return null;
  const match = text.match(/\d+/);
  return match ? String(Number(match[0])) : null;
}

function normalizeTipValues(fields) {
  return asArray(getField(fields, 'TIP')).map((v) => asString(v).toLowerCase());
}

function normalizePermitType(fields) {
  const raw = asString(getField(fields, 'Permit Type')).trim();
  const key = raw.toLowerCase().replace(/[^a-z0-9]/g, '');
  if (key === 'nb' || key.includes('newbuilding') || key.includes('newconstruction')) {
    return { raw, kind: 'new-building' };
  }
  if (
    key.includes('alt') ||
    key.includes('alteration') ||
    key.includes('conversion') ||
    key.includes('enlargement')
  ) {
    return { raw, kind: 'alteration-conversion' };
  }
  return raw ? { raw, kind: 'unknown' } : { raw: '', kind: null };
}

function isIcapRecord(fields) {
  return normalizeTipValues(fields).some((tip) => tip.includes('icap'));
}

function loadIcapBoundaries(filePath) {
  if (!filePath) return null;
  const resolved = path.resolve(filePath);
  const stat = fs.statSync(resolved);
  if (
    icapCache &&
    icapCache.filePath === resolved &&
    icapCache.mtimeMs === stat.mtimeMs &&
    icapCache.size === stat.size
  ) {
    return icapCache;
  }

  const workbook = XLSX.readFile(resolved, { cellDates: false });
  const blockKeys = new Set();

  for (const [sheetName, pairs] of Object.entries(ICAP_PAIR_COLUMNS)) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
    const expectedBorough = normalizeBorough(sheetName);
    if (!expectedBorough) continue;

    for (const row of rows) {
      for (const [boroughCol, blockCol] of pairs) {
        const rowBorough = normalizeBorough(row[boroughCol]);
        const block = normalizeBlock(row[blockCol]);
        if (!rowBorough || rowBorough.code !== expectedBorough.code || !block) continue;
        blockKeys.add(`${expectedBorough.code}:${block}`);
      }
    }
  }

  icapCache = {
    filePath: resolved,
    mtimeMs: stat.mtimeMs,
    size: stat.size,
    blockKeys
  };
  return icapCache;
}

function resolveIcapTerm(fields, filePath) {
  const icap = isIcapRecord(fields);
  const borough = normalizeBorough(getField(fields, 'Borough'));
  const block = normalizeBlock(getField(fields, 'Block'));

  if (!icap) {
    return {
      isIcap: false,
      checked: false,
      found: false,
      term: null,
      reason: 'TIP does not include ICAP',
      borough,
      block
    };
  }

  if (!borough || !block) {
    return {
      isIcap: true,
      checked: false,
      found: false,
      term: null,
      reason: 'Missing Borough or Block',
      borough,
      block
    };
  }

  if (!filePath || !fs.existsSync(filePath)) {
    return {
      isIcap: true,
      checked: false,
      found: false,
      term: null,
      reason: `ICAP boundaries file not found: ${filePath || '(not configured)'}`,
      borough,
      block
    };
  }

  const boundaries = loadIcapBoundaries(filePath);
  const found = boundaries.blockKeys.has(`${borough.code}:${block}`);
  return {
    isIcap: true,
    checked: true,
    found,
    term: found ? '25-year' : '15-year',
    reason: found ? 'Borough/block found in ICAP boundaries workbook' : 'Borough/block not found in ICAP boundaries workbook',
    borough,
    block
  };
}

function deriveProjectFacts(fields, opts) {
  opts = opts || {};
  const units = parseNumber(getField(fields, 'Units'));
  const commercialGrossSqft = parseNumber(getField(fields, 'Commercial Gross SQFT'));
  const buildingType = asString(getField(fields, 'Building Type')).toLowerCase();
  const hasCommercial =
    (commercialGrossSqft != null && commercialGrossSqft > 0) ||
    buildingType.includes('mixed') ||
    buildingType.includes('commercial');
  const permitType = normalizePermitType(fields);
  const icap = resolveIcapTerm(fields, opts.icapBoundariesFile);

  return {
    units,
    buildingSize: units == null ? null : (units > 10 ? 'Large (>10 units)' : 'Small (<=10 units)'),
    keepUnitSection: units == null ? null : (units > 10 ? 'transitional' : 'capped'),
    deleteUnitSection: units == null ? null : (units > 10 ? 'capped' : 'transitional'),
    permitType,
    keepPermitScenario: permitType.kind,
    deletePermitScenario: permitType.kind === 'new-building'
      ? 'alteration-conversion'
      : (permitType.kind === 'alteration-conversion' ? 'new-building' : null),
    commercialGrossSqft,
    hasCommercial,
    hasCommercialText: hasCommercial ? 'Yes' : 'No',
    icap
  };
}

module.exports = {
  getField,
  asString,
  parseNumber,
  normalizeBorough,
  normalizeBlock,
  normalizePermitType,
  deriveProjectFacts,
  resolveIcapTerm
};
