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

function grossSqftDescription(fields) {
  const residential = parseNumber(getField(fields, 'Residential Gross SQFT'));
  const commercial = parseNumber(getField(fields, 'Commercial Gross SQFT'));
  const fallback = parseNumber(getField(fields, 'Gross SQFT') || getField(fields, 'Total GSF') || getField(fields, 'GSF'));
  const parts = [];
  if (residential != null) parts.push(`${formatNumber(residential)} residential gross square feet`);
  if (commercial != null && commercial > 0) parts.push(`${formatNumber(commercial)} commercial gross square feet`);
  if (parts.length) return parts.join(' and ');
  return fallback != null ? `${formatNumber(fallback)} gross square feet` : '';
}

function buildingKind(fields, facts) {
  if (facts && facts.hasCommercial && facts.units != null) return 'mixed-use building';
  if (facts && facts.units != null) return 'residential building';
  if (facts && facts.hasCommercial) return 'commercial building';
  const raw = asString(getField(fields, 'Building Type')).trim();
  return raw ? `${raw.toLowerCase()} building` : 'building';
}

function projectDetailSentence(fields, facts) {
  const summary = projectSummary(fields, facts);
  const kind = buildingKind(fields, facts);
  if (facts && facts.keepPermitScenario === 'alteration-conversion') {
    return summary
      ? `The project involves the proposed renovation of an existing building into a ${kind}, featuring ${summary}.`
      : `The project involves the proposed renovation of an existing building into a ${kind}.`;
  }
  return summary
    ? `A new ${kind} will be constructed, featuring ${summary}.`
    : `A new ${kind} will be constructed.`;
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
    projectdetailsentence: projectDetailSentence(fields, facts),
    constructiondescription: projectDetailSentence(fields, facts),
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
    grosssqftdescription: grossSqftDescription(fields),
    grosssquarefeetdescription: grossSqftDescription(fields),
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
  const commercialSqft = getFactsValue('Commercial Gross SQFT', fields, facts);
  if (!commercialSqft && /\band\s+[\d,]+(?:\.\d+)?\s+square feet of commercial space/i.test(body)) {
    warnings.push('Commercial square-foot value remains even though Airtable has no Commercial Gross SQFT value.');
  }
  if (/\ba\s+\d+\s+residential/i.test(body)) {
    warnings.push('Bad grammar found: "a [number] residential...".');
  }
  const grossSqft = getFactsValue('Residential Gross SQFT', fields, facts);
  const grossSqftText = grossSqftDescription(fields);
  if (/\baggregate\s+gross square feet\b/i.test(body)) {
    warnings.push('Incomplete gross-square-feet sentence found.');
  }
  if (grossSqftText && /(?:residential|commercial) gross square feet/i.test(grossSqftText) && /\baggregate\s+[\d,.\s]+gross square feet\b/i.test(body) && !/\baggregate\b[^\r\n.]*\b(?:residential|commercial) gross square feet\b/i.test(body)) {
    warnings.push('Gross-square-foot sentence does not say whether the area is residential or commercial.');
  }
  if (!grossSqftText && /\baggregate\s+[\d,.\s]+gross square feet\b/i.test(body)) {
    warnings.push('Gross-square-foot sentence remains even though Airtable has no gross square footage value.');
  }
  if (!grossSqftText && /\baggregate\s+\[[^\]]*gross[^\]]*\]\s+gross square feet\b/i.test(body)) {
    warnings.push('Gross-square-foot placeholder remains even though Airtable has no gross square footage value.');
  }
  if (/For buildings with\s+\d+\s+residential[^.]*\s+or more/i.test(body)) {
    warnings.push('Possible statutory unit threshold was replaced with the project summary.');
  }
  if (facts && facts.units != null && facts.units < 100 && /(?:construction wage requirements|building service employees|prevailing wages)[^.]{0,240}(?:100 or more|150 or more)|(?:100 or more|150 or more)[^.]{0,240}(?:construction wage requirements|building service employees|prevailing wages)/i.test(body)) {
    warnings.push('A 100+ wage paragraph remains in an under-100-unit letter.');
  }
  if (facts && facts.units != null && facts.units < 100 && /Affordability Option B \(applicable for projects comprising 100 or more dwelling units\)/i.test(body)) {
    warnings.push('Affordability Option B bracket is wrong for a project with fewer than 100 units.');
  }
  if (facts && facts.units != null && facts.units < 100 && /Affordability Option B \(applicable for projects comprising 6 to 99 dwelling units\s+or more\)/i.test(body)) {
    warnings.push('Affordability Option B bracket contains an invalid "6 to 99 ... or more" range.');
  }
  if (facts && facts.units != null && facts.units > 10 && /For Modest Rental Projects with no more than ten residential dwelling units/i.test(body)) {
    warnings.push('Under-10 modest rental workbook timing remains in a project with more than 10 units.');
  }
  if (facts && facts.units != null && facts.units <= 10 && /For Modest Rental Projects with more than ten and fewer than one hundred residential dwelling units/i.test(body)) {
    warnings.push('Over-10 modest rental workbook timing remains in a project with 10 or fewer units.');
  }
  const hasTransitionalSection = /Transitional Assessed Valuation|transition assessment system|phased-in transition assessment|multiple Class A residential units will be classified as Tax Class 2/i.test(body);
  const hasCapSection = /Tax Class 2a,\s*2b and 2c CAP|Tax Class 2a|tax classes 2a,\s*2b,\s*and 2c|8% cap on annual assessment increases|30% within any five-year period/i.test(body);
  if (facts && facts.units != null && facts.units > 10 && hasCapSection) {
    warnings.push('2A/2B/2C cap section remains in a project with more than 10 units.');
  }
  if (facts && facts.units != null && facts.units <= 10 && hasTransitionalSection) {
    warnings.push('Transitional assessed valuation section remains in a project with 10 or fewer units.');
  }
  if (facts && facts.keepPermitScenario === 'alteration-conversion' && /\bnew construction\b/i.test(body)) {
    warnings.push('Permit Type is ALT/renovation, but "new construction" language remains.');
  }
  const owner = asString(getField(fields, 'Owner'));
  if (owner && new RegExp(`Sincerely yours[\\s\\S]{0,120}${owner.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}[\\s\\S]{0,120}Metropolitan Realty`, 'i').test(body)) {
    warnings.push('Signature block appears to contain the Owner instead of the firm signer.');
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
  const addMatchingParagraphs = (fieldName, regex, newValue, reason) => {
    const seen = new Set();
    for (const raw of body.split(/\r?\n/)) {
      const paragraph = raw.replace(/\s+/g, ' ').trim();
      if (!paragraph || seen.has(paragraph) || !regex.test(paragraph)) continue;
      seen.add(paragraph);
      add(fieldName, paragraph, newValue, reason);
    }
  };

  add(
    'Quality Fix',
    'commercial space and commercial space',
    'commercial space',
    'Removed duplicate commercial language'
  );

  const commercialSqft = getFactsValue('Commercial Gross SQFT', fields, facts);
  if (!commercialSqft) {
    const staleCommercialSqftRe = /\bcommercial space\s+and\s+[\d,]+(?:\.\d+)?\s+square feet of commercial space/gi;
    let commercialSqftMatch;
    let foundFullCommercialSqftPhrase = false;
    while ((commercialSqftMatch = staleCommercialSqftRe.exec(body)) !== null) {
      foundFullCommercialSqftPhrase = true;
      add(
        'Commercial SQFT Cleanup',
        commercialSqftMatch[0],
        'commercial space',
        'Removed commercial-square-foot phrase because Airtable has no Commercial Gross SQFT value'
      );
    }
    if (!foundFullCommercialSqftPhrase) {
      const staleCommercialTailRe = /\s+and\s+[\d,]+(?:\.\d+)?\s+square feet of commercial space/gi;
      while ((commercialSqftMatch = staleCommercialTailRe.exec(body)) !== null) {
        add(
          'Commercial SQFT Cleanup',
          commercialSqftMatch[0],
          ' ',
          'Removed commercial-square-foot phrase because Airtable has no Commercial Gross SQFT value'
        );
      }
    }
    add(
      'Commercial SQFT Cleanup',
      'commercial space and square feet of commercial space',
      'commercial space',
      'Removed incomplete commercial-square-foot phrase'
    );
  }

  const badArticleRe = /\ba\s+(\d+\s+residential[^\r\n.]*)/gi;
  let m;
  while ((m = badArticleRe.exec(body)) !== null) {
    add('Quality Fix', m[0], m[1], 'Removed bad article before unit phrase');
  }

  const desiredProjectDetail = projectDetailSentence(fields, facts);
  if (desiredProjectDetail) {
    const newBuildingLine = body.match(/A new [^\r\n.]*building[^\r\n.]*will be constructed[^\r\n.]*\./i);
    const renovationLine = body.match(/The project involves[^\r\n.]*(?:alteration|conversion|renovation)[^\r\n.]*\./i);
    if (facts && facts.keepPermitScenario === 'new-building') {
      if (newBuildingLine) {
        add(
          'Project Detail Cleanup',
          newBuildingLine[0],
          desiredProjectDetail,
          'Normalized project detail line for NB/new-building permit'
        );
      } else if (renovationLine) {
        add(
          'Project Detail Cleanup',
          renovationLine[0],
          desiredProjectDetail,
          'Moved NB/new-building project detail into the remaining numbered line'
        );
      }
    } else if (facts && facts.keepPermitScenario === 'alteration-conversion') {
      if (renovationLine) {
        add(
          'Project Detail Cleanup',
          renovationLine[0],
          desiredProjectDetail,
          'Normalized project detail line for ALT/renovation permit'
        );
      } else if (newBuildingLine) {
        add(
          'Project Detail Cleanup',
          newBuildingLine[0],
          desiredProjectDetail,
          'Moved ALT/renovation project detail into the remaining numbered line'
        );
      }
    }
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
      'new building will be constructed'
    ]) {
      if (body.toLowerCase().includes(phrase.toLowerCase())) {
        add('Permit Type Cleanup', phrase, '', 'Removed new-building language for alteration/conversion project');
        break;
      }
    }
    add(
      'Permit Type Cleanup',
      'new construction',
      'renovation',
      'Changed new-construction wording to renovation for ALT/renovation project'
    );
  }

  const grossSqft = getFactsValue('Residential Gross SQFT', fields, facts);
  const grossSqftText = grossSqftDescription(fields);
  const aggregateGrossSqftSentence = body.match(/The building will aggregate\s+(?:[\d,.\s]*|\[[^\]]+\]\s*)gross square feet\.?/i);
  if (/\baggregate\s+(?:[\d,.\s]*|\[[^\]]+\]\s*)gross square feet\b/i.test(body)) {
    if (grossSqftText) {
      add(
        'Gross SQFT Cleanup',
        aggregateGrossSqftSentence ? aggregateGrossSqftSentence[0] : 'aggregate gross square feet',
        aggregateGrossSqftSentence ? `The building will aggregate ${grossSqftText}.` : `aggregate ${grossSqftText}`,
        'Filled missing or stale gross-square-foot phrase'
      );
      add(
        'Gross SQFT Cleanup',
        'aggregate gross square feet',
        `aggregate ${grossSqftText}`,
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
  if (summary && facts && facts.units != null) {
    const unitOnly = `${formatNumber(facts.units)} residential rental units`;
    add(
      'Project Summary Cleanup',
      `${unitOnly} ${summary}`,
      summary,
      'Collapsed repeated unit/project summary fragment'
    );
    add(
      'Project Summary Cleanup',
      `${summary} ${summary}`,
      summary,
      'Collapsed repeated project summary fragment'
    );
    add(
      'Project Summary Cleanup',
      `${summary} ${summary} - modest rental project`,
      `${summary} - modest rental project`,
      'Collapsed repeated project summary in property summary'
    );
    add(
      'Project Summary Cleanup',
      `${summary} ${summary} – modest rental project`,
      `${summary} – modest rental project`,
      'Collapsed repeated project summary in property summary'
    );
  }
  if (summary && /Property Summary:\s*[^\r\n]+/i.test(body)) {
    const line = body.match(/Property Summary:\s*([^\r\n]+)/i);
    if (line && /commercial space\s+and\s+commercial space/i.test(line[0])) {
      add('Property Summary Cleanup', line[0], line[0].replace(/commercial space\s+and\s+commercial space/ig, 'commercial space'), 'Cleaned property summary duplicate');
    }
  }

  const under100Project = facts && facts.units != null && facts.units < 100;
  const badWageSentence = body.match(/For buildings with\s+\d+\s+residential[^\r\n.]*?\s+or more,\s+all building service employees/i);
  if (badWageSentence && !under100Project) {
    add(
      'Statutory Threshold Cleanup',
      badWageSentence[0],
      'For buildings with 100 or more dwelling units, all building service employees',
      'Restored statutory wage threshold sentence that should not use project unit count'
    );
  }

  const badWageThreshold = body.match(/For buildings with\s+\d+\s+residential[^\r\n.]*?\s+or more,/i);
  if (badWageThreshold && !under100Project) {
    add(
      'Statutory Threshold Cleanup',
      badWageThreshold[0],
      'For buildings with 100 or more dwelling units,',
      'Restored statutory wage threshold that should not use project unit count'
    );
  }

  const badWageCore = body.match(/\d+\s+residential[^\r\n.]*?\s+or more/i);
  if (badWageCore && !under100Project && /building service employees|prevailing wages/i.test(body)) {
    add(
      'Statutory Threshold Cleanup',
      badWageCore[0],
      '100 or more dwelling units',
      'Restored short statutory wage threshold fragment'
    );
  }
  if (facts && facts.units != null && !under100Project) {
    const summaryThreshold = `${projectSummary(fields, facts)} or more`;
    add(
      'Statutory Threshold Cleanup',
      summaryThreshold,
      '100 or more dwelling units',
      'Restored statutory wage threshold fragment'
    );
  }
  if (under100Project) {
    add(
      'Wage Paragraph Cleanup',
      'Under the ANNY Rules, for any eligible site containing 100 or more dwelling units, regardless of location within New York City, the applicant must comply with specific construction wage requirements for workers at an Eligible Site.',
      '',
      'Removed 100+ construction wage paragraph for under-100-unit project'
    );
    add(
      'Wage Paragraph Cleanup',
      'Eligible sites containing 150 or more dwelling units that are located in Zone A or Zone B are subject to higher minimum construction wage floors.',
      '',
      'Removed 150+ construction wage sentence for under-100-unit project'
    );
    add(
      'Wage Paragraph Cleanup',
      'For buildings with 100 or more dwelling units, all building service employees',
      '',
      'Removed 100+ building-service wage paragraph for under-100-unit project'
    );
    add(
      'Wage Paragraph Cleanup',
      `For buildings with ${projectSummary(fields, facts)} or more, all building service employees`,
      '',
      'Removed project-specific rewrite of 100+ wage paragraph for under-100-unit project'
    );
  }

  if (facts && facts.units != null && facts.units > 10) {
    add(
      'Modest Rental Workbook Cleanup',
      'For Modest Rental Projects with no more than ten residential dwelling units, a 485-X Workbook must be submitted to HPD no earlier than 6 months before the expected completion date and no later than 2 months after the completion date.',
      '',
      'Removed under-10 workbook timing for project with more than 10 units'
    );
  } else if (facts && facts.units != null && facts.units <= 10) {
    add(
      'Modest Rental Workbook Cleanup',
      'For Modest Rental Projects with more than ten and fewer than one hundred residential dwelling units, a 485-X Workbook must be submitted to HPD no earlier than 9 months before the expected completion date and no later than 2 months after the completion date.',
      '',
      'Removed over-10 workbook timing for project with 10 or fewer units'
    );
  }

  if (facts && facts.units != null && facts.units > 10) {
    addMatchingParagraphs(
      'Valuation Section Cleanup',
      /^(?:Tax Class\s*)?2a,\s*2b\s*and\s*2c\s*CAP$|^Tax Class 2a,\s*2b\s*and\s*2c\s*CAP$/i,
      '',
      'Removed 2A/2B/2C cap heading for project with more than 10 units'
    );
    addMatchingParagraphs(
      'Valuation Section Cleanup',
      /tax classes 2a,\s*2b,\s*and 2c|tax class 2a,\s*2b,\s*and 2c|8% cap on annual assessment increases|30% within any five-year period|assessment cap does not apply/i,
      '',
      'Removed 2A/2B/2C cap paragraph for project with more than 10 units'
    );
  } else if (facts && facts.units != null && facts.units <= 10) {
    addMatchingParagraphs(
      'Valuation Section Cleanup',
      /^Transitional Assessed Valuation$/i,
      '',
      'Removed transitional assessed valuation heading for project with 10 or fewer units'
    );
    addMatchingParagraphs(
      'Valuation Section Cleanup',
      /multiple Class A residential units will be classified as Tax Class 2|transition assessment system|phased-in transition assessment|substantial increase in value.*future years.*phased-in transition/i,
      '',
      'Removed transitional assessed valuation paragraph for project with 10 or fewer units'
    );
  }

  if (facts && facts.units != null && facts.units < 100) {
    add(
      'Affordability Option Cleanup',
      'Affordability Option B (applicable for projects comprising 100 or more dwelling units)',
      'Affordability Option B (applicable for projects comprising 6 to 99 dwelling units)',
      'Restored Option B bracket for fewer than 100 units'
    );
    add(
      'Affordability Option Cleanup',
      'Affordability Option B (applicable for projects comprising 6 to 99 dwelling units or more)',
      'Affordability Option B (applicable for projects comprising 6 to 99 dwelling units)',
      'Removed invalid "or more" from Option B bracket range'
    );
    add(
      'Affordability Option Cleanup',
      'projects comprising 6 to 99 dwelling units or more',
      'projects comprising 6 to 99 dwelling units',
      'Removed invalid "or more" from Option B bracket range'
    );
  }

  const owner = asString(getField(fields, 'Owner'));
  if (owner && new RegExp(`Sincerely yours[\\s\\S]{0,120}${owner.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}[\\s\\S]{0,120}Metropolitan Realty`, 'i').test(body)) {
    add(
      'Signature Cleanup',
      owner,
      'Martin Joseph',
      'Restored firm signature; owner should not replace signer'
    );
  }

  if (/Tax liability without 485-X benefits[\s\S]{0,500}Tax liability with 485-X benefits/i.test(body)) {
    const taxSectionMatch = body.match(/Tax liability without 485-X benefits[\s\S]{0,900}?(?=Disclaimer|$)/i);
    const taxSection = taxSectionMatch ? taxSectionMatch[0] : '';
    const taxAmountRe = /\b(?:Assessed Value|Tax liability)\s+(\$[\d,]+(?:\.\d{2})?)/gi;
    let taxMatch;
    while ((taxMatch = taxAmountRe.exec(taxSection)) !== null) {
      add(
        'Tax Projection Cleanup',
        taxMatch[1],
        '[To be determined]',
        'Removed old projected tax value because no Airtable tax projection field is available'
      );
    }
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
