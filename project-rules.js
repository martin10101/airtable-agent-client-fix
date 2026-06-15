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

const CONTROL_MARKER_RE = /\[\[\s*(END|KEEP_IF_[^\]]+|DELETE_IF_[^\]]+)\s*\]\]/i;

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

function normalizeToken(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function formatNumber(value) {
  const n = parseNumber(value);
  if (n == null) return asString(value);
  return new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 }).format(n);
}

function projectSummary(fields, facts) {
  const units = facts && facts.units != null ? facts.units : parseNumber(getField(fields, 'Units'));
  const condoRental = asString(getField(fields, 'Condo/Rental')).toLowerCase();
  const isCondo = condoRental.includes('condo') && !condoRental.includes('rental');
  const unitLabel = isCondo ? 'residential condominium units' : 'residential rental units';
  const parts = [];
  if (units != null) parts.push(`${formatNumber(units)} ${unitLabel}`);
  if (facts && facts.hasCommercial) parts.push('commercial space');
  if (!parts.length && facts && facts.hasCommercial) return 'commercial space';
  return parts.join(' and ');
}

function getFactsValue(token, fields, facts) {
  const key = normalizeToken(token);
  const borough = normalizeBorough(getField(fields, 'Borough'));
  const block = asString(getField(fields, 'Block'));
  const lot = asString(getField(fields, 'Lot'));
  const genericFields = {};
  for (const [name, value] of Object.entries(fields || {})) {
    genericFields[normalizeToken(name)] = value;
  }

  const direct = {
    projectsummary: projectSummary(fields, facts),
    propertysummary: projectSummary(fields, facts),
    buildingconfiguration: projectSummary(fields, facts),
    buildingdescription: projectSummary(fields, facts),
    units: formatNumber(getField(fields, 'Units')),
    unitcount: formatNumber(getField(fields, 'Units')),
    propertyaddress: asString(getField(fields, 'Property Address') || getField(fields, 'Address') || getField(fields, 'Project Address')),
    projectaddress: asString(getField(fields, 'Property Address') || getField(fields, 'Address') || getField(fields, 'Project Address')),
    address: asString(getField(fields, 'Property Address') || getField(fields, 'Address') || getField(fields, 'Project Address')),
    owner: asString(getField(fields, 'Owner')),
    borough: borough ? borough.name : asString(getField(fields, 'Borough')),
    block,
    lot,
    blocklot: block && lot ? `${block} - ${lot}` : '',
    blockandlot: block && lot ? `${block} - ${lot}` : '',
    bbl: borough && block && lot ? `${borough.code}${String(block).padStart(5, '0')}${String(lot).padStart(4, '0')}` : '',
    residentialgrosssqft: formatNumber(getField(fields, 'Residential Gross SQFT') || getField(fields, 'Gross SQFT') || getField(fields, 'Total GSF') || getField(fields, 'GSF')),
    grosssqft: formatNumber(getField(fields, 'Residential Gross SQFT') || getField(fields, 'Gross SQFT') || getField(fields, 'Total GSF') || getField(fields, 'GSF')),
    totalgsf: formatNumber(getField(fields, 'Residential Gross SQFT') || getField(fields, 'Gross SQFT') || getField(fields, 'Total GSF') || getField(fields, 'GSF')),
    commercialgrosssqft: formatNumber(getField(fields, 'Commercial Gross SQFT')),
    commercialsqft: formatNumber(getField(fields, 'Commercial Gross SQFT')),
    hascommercial: facts && facts.hasCommercial ? 'Yes' : 'No',
    buildingsize: facts ? asString(facts.buildingSize) : '',
    buildingtype: asString(getField(fields, 'Building Type')),
    permittype: asString(getField(fields, 'Permit Type')),
    permitscenario: facts ? asString(facts.keepPermitScenario) : '',
    icapterm: facts && facts.icap ? asString(facts.icap.term) : '',
    icapyears: facts && facts.icap ? asString(facts.icap.term) : '',
    abatementterm: facts && facts.icap ? asString(facts.icap.term) : '',
    affordabilityoption: asString(getField(fields, 'Affordability Option') || getField(fields, '485-X Affordability Option') || getField(fields, '485X Affordability Option'))
  };

  if (Object.prototype.hasOwnProperty.call(direct, key)) return direct[key];
  if (Object.prototype.hasOwnProperty.call(genericFields, key)) return asString(genericFields[key]);
  return '';
}

function valueForCondition(name, fields, facts) {
  const key = normalizeToken(name);
  if (key === 'permit' || key === 'permittype') {
    const permit = facts && facts.permitType ? facts.permitType : normalizePermitType(fields);
    return [permit.raw, permit.kind].filter(Boolean).join(' ');
  }
  if (key === 'permitkind' || key === 'permitscenario') return facts && facts.keepPermitScenario ? facts.keepPermitScenario : '';
  if (key === 'units' || key === 'unitcount') return facts && facts.units != null ? facts.units : parseNumber(getField(fields, 'Units'));
  if (key === 'hascommercial' || key === 'commercial') return facts && facts.hasCommercial ? 'Yes' : 'No';
  if (key === 'icapterm' || key === 'icapyears' || key === 'abatementterm') return facts && facts.icap ? facts.icap.term : '';
  if (key === 'tip') return normalizeTipValues(fields).join(', ');
  if (key === 'buildingsize') return facts ? facts.buildingSize : '';
  if (key === 'buildingtype') return asString(getField(fields, 'Building Type'));
  return getFactsValue(name, fields, facts);
}

function compareCondition(actual, operator, expected) {
  const actualNum = parseNumber(actual);
  const expectedNum = parseNumber(expected);
  if (['>', '>=', '<', '<='].includes(operator) && actualNum != null && expectedNum != null) {
    if (operator === '>') return actualNum > expectedNum;
    if (operator === '>=') return actualNum >= expectedNum;
    if (operator === '<') return actualNum < expectedNum;
    return actualNum <= expectedNum;
  }

  const a = normalizeToken(actual);
  const e = normalizeToken(expected);
  const eq = a === e || (a && e && a.includes(e));
  return operator === '!=' ? !eq : eq;
}

function evaluateTemplateMarkerText(paragraphText, fields, facts) {
  const match = String(paragraphText || '').match(CONTROL_MARKER_RE);
  if (!match) return null;
  const inner = match[1].trim();
  if (/^END$/i.test(inner)) return { kind: 'end', raw: match[0] };

  const mode = /^DELETE_IF_/i.test(inner) ? 'delete' : 'keep';
  const condition = inner.replace(/^(KEEP_IF_|DELETE_IF_)/i, '');
  const parsed = condition.match(/^(.+?)(>=|<=|!=|=|>|<)(.+)$/);
  if (!parsed) {
    return { kind: 'start', keep: true, raw: match[0], reason: `Unrecognized marker condition: ${inner}` };
  }

  const actual = valueForCondition(parsed[1].trim(), fields, facts);
  const expected = parsed[3].trim();
  const conditionTrue = compareCondition(actual, parsed[2], expected);
  const keep = mode === 'keep' ? conditionTrue : !conditionTrue;
  return {
    kind: 'start',
    keep,
    raw: match[0],
    reason: `${parsed[1].trim()} ${parsed[2]} ${expected}; actual=${asString(actual)}`
  };
}

function buildTemplatePlaceholderSwaps(templateText, fields, facts, opts) {
  opts = opts || {};
  const log = typeof opts.log === 'function' ? opts.log : (() => {});
  const swaps = [];
  const seen = new Set();
  const markerRe = /\[\[\s*([A-Za-z0-9 _/&().-]+?)\s*\]\]/g;
  let m;
  while ((m = markerRe.exec(String(templateText || ''))) !== null) {
    const full = m[0];
    const token = m[1].trim();
    if (!token || /^END$/i.test(token) || /^(KEEP_IF_|DELETE_IF_)/i.test(token)) continue;
    if (seen.has(full)) continue;
    seen.add(full);
    const value = getFactsValue(token, fields, facts);
    if (value == null || value === '') {
      log(`[template-rule] placeholder left unchanged (no Airtable value): ${full}`);
      continue;
    }
    swaps.push({ fieldName: `Template Placeholder ${token}`, oldValue: full, newValue: String(value) });
    log(`[template-rule] placeholder ${full} -> ${JSON.stringify(String(value).slice(0, 80))}`);
  }
  return swaps;
}

function inspectGeneratedDocxText(text, fields, facts) {
  const warnings = [];
  const body = String(text || '');
  const lower = body.toLowerCase();

  if (/commercial space\s+and\s+commercial space/i.test(body)) {
    warnings.push('Duplicate commercial language found: "commercial space and commercial space".');
  }
  if (/\ba\s+\d+\s+residential/i.test(body)) {
    warnings.push('Bad grammar found: "a [number] residential...".');
  }
  const grossSqft = getFactsValue('Residential Gross SQFT', fields, facts);
  if (/\baggregate\s+gross square feet\b/i.test(body)) {
    warnings.push('Incomplete gross-square-feet sentence found.');
  }
  if (!grossSqft && /\baggregate\s+[\d,.\s]+gross square feet\b/i.test(body)) {
    warnings.push('Gross-square-foot sentence remains even though Airtable has no Residential Gross SQFT value.');
  }
  if (!grossSqft && /\baggregate\s+\[[^\]]*gross[^\]]*\]\s+gross square feet\b/i.test(body)) {
    warnings.push('Gross-square-foot placeholder remains even though Airtable has no Residential Gross SQFT value.');
  }
  if (/For buildings with\s+\d+\s+residential[^.]*\s+or more/i.test(body)) {
    warnings.push('Possible statutory unit threshold was replaced with the project summary.');
  }
  if (/\[\[[^\]]+\]\]/.test(body)) {
    warnings.push('Unresolved [[...]] template marker or placeholder remains.');
  }

  if (facts && facts.deletePermitScenario === 'alteration-conversion') {
    if (lower.includes('alteration') || lower.includes('conversion')) {
      warnings.push('Permit Type is NB, but alteration/conversion language remains.');
    }
  }
  if (facts && facts.deletePermitScenario === 'new-building') {
    if (lower.includes('new building') || lower.includes('new construction')) {
      warnings.push('Permit Type is alteration/conversion, but new-building language remains.');
    }
  }
  if (facts && facts.icap && facts.icap.term) {
    const wrong = facts.icap.term === '25-year' ? '15-year' : '25-year';
    if (lower.includes('icap') && lower.includes(wrong)) {
      warnings.push(`ICAP lookup selected ${facts.icap.term}, but ${wrong} language remains.`);
    }
  }
  if (facts && facts.units != null) {
    const unitRe = /\b(\d{1,4})\s+(?:residential\s+)?(?:rental\s+)?(?:apartments|units)\b/gi;
    let m;
    while ((m = unitRe.exec(body)) !== null) {
      const n = Number(m[1]);
      if (Number.isFinite(n) && n !== facts.units) {
        warnings.push(`Possible old unit count remains: "${m[0]}" does not match Airtable Units=${facts.units}.`);
        break;
      }
    }
  }

  return warnings;
}

function buildPostGenerationCleanupSwaps(text, fields, facts, opts) {
  opts = opts || {};
  const log = typeof opts.log === 'function' ? opts.log : (() => {});
  const body = String(text || '');
  const swaps = [];
  const add = (fieldName, oldValue, newValue, reason) => {
    if (!oldValue || swaps.some((s) => s.oldValue === oldValue && s.newValue === newValue)) return;
    if (!body.toLowerCase().includes(String(oldValue).toLowerCase())) return;
    swaps.push({ fieldName, oldValue, newValue });
    log(`[quality-fix] ${reason}: ${JSON.stringify(String(oldValue).slice(0, 120))} -> ${JSON.stringify(String(newValue).slice(0, 120))}`);
  };

  add(
    'Quality Fix',
    'commercial space and commercial space',
    'commercial space',
    'Removed duplicate commercial language'
  );

  const badArticleRe = /\ba\s+(\d+\s+residential[^\r\n.]*)/gi;
  let m;
  while ((m = badArticleRe.exec(body)) !== null) {
    add('Quality Fix', m[0], m[1], 'Removed bad article before unit phrase');
  }

  if (facts && facts.deletePermitScenario === 'alteration-conversion') {
    for (const phrase of [
      'The project involves the proposed alteration and conversion',
      'proposed alteration and conversion',
      'alteration and conversion',
      'alteration or conversion'
    ]) {
      if (body.toLowerCase().includes(phrase.toLowerCase())) {
        add('Permit Type Cleanup', phrase, '', 'Removed alteration/conversion language for NB project');
        break;
      }
    }
  } else if (facts && facts.deletePermitScenario === 'new-building') {
    for (const phrase of [
      'A new mixed-use building will be constructed',
      'A new residential building will be constructed',
      'A new commercial building will be constructed',
      'new building will be constructed',
      'new construction'
    ]) {
      if (body.toLowerCase().includes(phrase.toLowerCase())) {
        add('Permit Type Cleanup', phrase, '', 'Removed new-building language for alteration/conversion project');
        break;
      }
    }
  }

  const grossSqft = getFactsValue('Residential Gross SQFT', fields, facts);
  const aggregateGrossSqftSentence = body.match(/The building will aggregate\s+(?:[\d,.\s]*|\[[^\]]+\]\s*)gross square feet\.?/i);
  if (/\baggregate\s+(?:[\d,.\s]*|\[[^\]]+\]\s*)gross square feet\b/i.test(body)) {
    if (grossSqft) {
      add(
        'Gross SQFT Cleanup',
        aggregateGrossSqftSentence ? aggregateGrossSqftSentence[0] : 'aggregate gross square feet',
        aggregateGrossSqftSentence ? `The building will aggregate ${grossSqft} gross square feet.` : `aggregate ${grossSqft} gross square feet`,
        'Filled missing or stale gross-square-foot phrase'
      );
      add(
        'Gross SQFT Cleanup',
        'aggregate gross square feet',
        `aggregate ${grossSqft} gross square feet`,
        'Filled missing gross-square-foot phrase'
      );
    } else {
      if (aggregateGrossSqftSentence) {
        add(
          'Gross SQFT Cleanup',
          aggregateGrossSqftSentence[0],
          '',
          'Removed gross-square-foot sentence because Airtable has no Residential Gross SQFT value'
        );
      }
      add(
        'Gross SQFT Cleanup',
        'The building will aggregate gross square feet.',
        '',
        'Removed incomplete gross-square-foot sentence because Airtable has no value'
      );
      add(
        'Gross SQFT Cleanup',
        'The building will aggregate gross square feet',
        '',
        'Removed incomplete gross-square-foot sentence because Airtable has no value'
      );
    }
  }

  if (facts && facts.units != null) {
    const unitRe = /\b(\d{1,4})\s+(?:residential\s+)?(?:rental\s+)?(?:apartments|units)\b/gi;
    while ((m = unitRe.exec(body)) !== null) {
      const n = Number(m[1]);
      if (Number.isFinite(n) && n !== facts.units) {
        const summary = projectSummary(fields, facts);
        if (summary) add('Unit Count Cleanup', m[0], summary, 'Replaced likely old unit count');
      }
    }
  }

  const summary = projectSummary(fields, facts);
  if (summary && /Property Summary:\s*[^\r\n]+/i.test(body)) {
    const line = body.match(/Property Summary:\s*([^\r\n]+)/i);
    if (line && /commercial space\s+and\s+commercial space/i.test(line[0])) {
      add('Property Summary Cleanup', line[0], line[0].replace(/commercial space\s+and\s+commercial space/ig, 'commercial space'), 'Cleaned property summary duplicate');
    }
  }

  const badWageThreshold = body.match(/For buildings with\s+\d+\s+residential[^.]*?\s+or more,/i);
  if (badWageThreshold) {
    add(
      'Statutory Threshold Cleanup',
      badWageThreshold[0],
      'For buildings with 100 or more dwelling units,',
      'Restored statutory wage threshold that should not use project unit count'
    );
  }

  return swaps;
}

module.exports = {
  getField,
  asString,
  parseNumber,
  normalizeBorough,
  normalizeBlock,
  normalizePermitType,
  deriveProjectFacts,
  resolveIcapTerm,
  evaluateTemplateMarkerText,
  buildTemplatePlaceholderSwaps,
  inspectGeneratedDocxText,
  buildPostGenerationCleanupSwaps
};
