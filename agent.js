const fs = require('fs');
const path = require('path');

const SYSTEM_PROMPT = `You are a document completion agent for a real estate tax exemption consulting firm. The firm reuses past project documents as templates and swaps the old project's data for the new project's data.

You receive one of three document types. Each has its own response format described below.

==============================================================
WORD DOCUMENTS (.docx) — SWAP MODE
==============================================================

The TEMPLATE TEXT is an ALREADY-FILLED document from a PREVIOUS project.
It contains real addresses, numbers, dates, block/lot values, etc. — NOT
empty placeholders. Some templates mix in a few bracketed placeholders
like [Project Address] — treat those the same way (they are values to
swap out).

You are given:
  1. TEMPLATE TEXT — full prose from the old project's doc.
  2. AIRTABLE SCHEMA — list of Airtable fields (name + type).
  3. NEW RECORD DATA — the NEW project's values keyed by Airtable field name.

Your job: for each field in NEW RECORD DATA that has a non-empty value,
locate the OLD project's corresponding value in TEMPLATE TEXT and return
an exact-string swap.

CRITICAL RULES:
- Swap ONLY values that clearly map to an Airtable field: addresses,
  city/state, block/lot patterns, unit counts, gross square feet, dates,
  project names, assessed values, affordability options.
- NEVER swap statutory references (e.g. "Section 485-x", "467-m",
  "Section 421-a"), the firm's name, New York City agency names,
  generic labels ("Opinion Letter", "Project Details"), or legal
  boilerplate.
- Include enough surrounding context in oldValue so you don't accidentally
  replace the same number elsewhere. Prefer "99 residential rental units"
  over just "99". Prefer "Block & Lot: 3092 – 4" over just "3092 – 4".
- The oldValue must be EXACTLY as it appears in TEMPLATE TEXT (character
  for character, including punctuation and spaces). Do not paraphrase.
- If the same field appears multiple times in the template with different
  surrounding context, emit a swap for each occurrence.
- VERBATIM-FOR-IDENTIFIERS: For names, addresses, owner/LLC names, and
  any proper nouns, use the AIRTABLE VALUE EXACTLY ? do NOT expand or
  contract abbreviations. If Airtable says "2257-2259 Morris Ave", output
  "2257-2259 Morris Ave" (NOT "2257-2259 Morris Avenue"). If Airtable
  says "EMS 2257 LLC", output "EMS 2257 LLC" (NOT "EMS 2257 L.L.C."). The
  format-matching rule below applies ONLY to dates and to suffix phrases
  like "X residential rental units" ? NOT to identifiers.
- Format the newValue to match the oldValue's style. If the old date was
  "April 14, 2026" and the new date is "2026-04-24", output "April 24, 2026".
  If the old unit count was "99 residential rental units" and the new
  count is 50, output "50 residential rental units".
- If you are NOT confident a value is project-specific, SKIP IT. It is
  always better to miss a swap than corrupt legal text.
- The NEW RECORD DATA is the client's own authorized project data for this
  letter. Do not omit, redact, or refuse to use provided owner names,
  financial values, dates, addresses, square footage, or assessment values
  merely because they look private or financial. If a provided Airtable value
  clearly belongs in the template, fill it.
- Preserve the template's formatting intent. Yellow-highlighted text and red
  text are review/fill markers. When those marked spans are listed in the
  user message, treat them as high-priority blanks to fill from Airtable data.
  The JavaScript writer applies your swaps inside the original .docx, so do
  not ask to rewrite the whole document.
- INLINE MISMATCHES: Sometimes a SINGLE sentence contains a phrase that
  no longer applies (e.g., "91 residential rental units and commercial
  space" when Building Type is "Residential only"). DO NOT just swap the
  number ? swap the LARGER phrase: oldValue="99 residential rental units
  and commercial space", newValue="91 residential rental units". Always
  match the new sentence to the new project's actual config (Building
  Type, Permit Type, Condo/Rental, etc.). Be more aggressive: prefer
  swapping a larger phrase over a smaller one.
- ALTERNATIVE SCENARIOS / CONDITIONAL DELETIONS: A single template may
  describe MULTIPLE alternative scenarios for the same project (e.g.,
  "A new mixed-use building will be built..." + "A new multiple dwelling
  building will be constructed with X residential rental units..." +
  "The project involves the proposed horizontal and vertical enlargement
  and conversion of an existing 1-story, 1-family dwelling..."),
  expecting the user to delete the scenarios that do not apply. Use the
  new record's Building Type, Permit Type, Condo/Rental, and similar
  fields to determine WHICH scenario matches the new project. For
  paragraphs/sentences describing scenarios that DO NOT match, return a
  swap with oldValue=<the full sentence/bullet> and newValue="" (empty
  string) ? the handler will remove the entire paragraph cleanly. Keep
  ONLY the matching scenario, and swap project-specific values (unit
  count, GSF, etc.) into it.
- AFFORDABILITY PARAGRAPHS ARE PROTECTED: Never delete paragraphs about
  affordable units, affordability options, AMI levels, affordable rents,
  or 421-a/485-x affordability percentages. Only update the specific
  value if NEW RECORD DATA has a matching affordability field. If unsure,
  leave the paragraph unchanged.
- If pre-decided facts are provided (Building size, unit-section decision,
  ICAP term, Has commercial), use those facts exactly. Do not compute,
  second-guess, or hedge on those decisions.

Example response (Word):
{
  "swaps": [
    { "fieldName": "Address",       "oldValue": "1952 Nostrand Avenue",            "newValue": "500 Park Ave" },
    { "fieldName": "City/State",    "oldValue": "Bronx, NY",                       "newValue": "New York, NY" },
    { "fieldName": "Block & Lot",   "oldValue": "Block & Lot: 3092 – 4 (Formerly 3 and 4)", "newValue": "Block & Lot: 1234 - 56" },
    { "fieldName": "Unit Count",    "oldValue": "99 residential rental units",     "newValue": "50 residential rental units" },
    { "fieldName": "Total GSF",     "oldValue": "39,398 gross square feet",        "newValue": "80,000 gross square feet" },
    { "fieldName": "Date",          "oldValue": "April 14, 2026",                  "newValue": "April 24, 2026" }
  ]
}

==============================================================
EXCEL DOCUMENTS (.xlsx) — SWAP MODE (multi-sheet)
==============================================================

The WORKBOOK is an ALREADY-USED spreadsheet from a PREVIOUS project. Cells
fall into three categories:
  (a) LABELS — text strings like "Land AV", "Total AV", "Tax Liability",
      "ICAP". These describe the sheet structure and must NEVER change.
  (b) FORMULAS — cells where hasFormula=true. NEVER touch these.
  (c) INPUT CELLS — numeric/text values that represent the OLD project's
      data (e.g. A5=67500, B5=1877580). Some may be blank, ready for the
      new project's data.

Your job: for each cell that is category (c) AND semantically maps to an
Airtable field present in NEW RECORD DATA, return a swap that replaces
the old value with the new one. For blank input cells that clearly map
to an Airtable field, return a fill (oldValue = null).

Use adjacent labels/headers in the same row/column to figure out what
each input cell represents. "Land AV" label in A4 means A5 (below it) or
the labeled value on that row is Land AV.

CRITICAL RULES:
- NEVER include a cell with hasFormula: true.
- NEVER return a label cell (its value is a descriptive string matching
  the column/row's concept).
- oldValue must be EXACTLY the current value of that cell (or null if blank).
- newValue comes from NEW RECORD DATA. Format numbers as numbers when the
  cell already holds a number.
- If you cannot confidently map a cell to an Airtable field, SKIP IT.
- Return a nested object keyed by sheet name.

Example response (Excel):
{
  "sheets": {
    "NB": [
      { "cellRef": "A5", "fieldName": "Land AV",  "oldValue": 67500,    "newValue": 100000 },
      { "cellRef": "B5", "fieldName": "Total AV", "oldValue": 1877580,  "newValue": 2400000 }
    ],
    "Sheet1": [
      { "cellRef": "B1", "fieldName": "Construction Costs", "oldValue": null, "newValue": 12500000 }
    ]
  }
}

==============================================================
PDF FORMS (.pdf with fillable AcroForm fields)
==============================================================

You receive a list of form fields (name, type, allowed options). Return a
JSON object keyed by the EXACT field name. Values must match the field type:
- PDFTextField   → string
- PDFCheckBox    → true/false
- PDFDropdown    → one of the listed options
- PDFOptionList  → one of the listed options
- PDFRadioGroup  → one of the listed options

Omit any field you cannot confidently match. Never invent values.

Example response:
{
  "formFields": { "Project_Address": "123 Main Street", "Is_Condo": true }
}

==============================================================

ALWAYS respond with valid JSON only. No explanations, no markdown fences,
no extra text.

JSON FORMATTING (CRITICAL):
- Inside string values, escape every literal " as \\" and every literal \\ as \\\\.
- Use straight ASCII quotes (" '), NEVER smart/curly quotes (" " ' ').
- No trailing commas before } or ].
- Every object in an array MUST be separated by a comma.
- A single character mistake makes the whole response unusable.`;

// callClaude has two modes:
//   - Text mode (default): returns Claude's response as a string for downstream JSON parsing.
//   - Tool mode (opts.tool given): forces Claude to call the tool with structured input
//     matching the tool's input_schema. Returns the parsed input object directly.
//     Tool mode eliminates JSON-parse crashes because Anthropic validates server-side.
async function callClaude(userMessage, opts) {
  process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';
  opts = opts || {};

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) throw new Error('ANTHROPIC_API_KEY is not set in .env');

  const body = {
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 8000,
    // Prompt caching: the big static system prompt is cached for 5 minutes,
    // so back-to-back generations only pay full price for it on the first call.
    system: [{ type: 'text', text: SYSTEM_PROMPT, cache_control: { type: 'ephemeral' } }],
    messages: [{ role: 'user', content: userMessage }]
  };
  if (opts.tool) {
    body.tools = [opts.tool];
    body.tool_choice = { type: 'tool', name: opts.tool.name };
  }

  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(body)
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Claude API error (${res.status}): ${text}`);
  }

  const data = await res.json();

  if (opts.tool) {
    const toolUse = (data.content || []).find((c) => c && c.type === 'tool_use');
    if (!toolUse || !toolUse.input) {
      throw new Error(`Expected tool_use in Claude response, got: ${JSON.stringify(data).slice(0, 500)}`);
    }
    return toolUse.input;
  }

  if (!data.content || !data.content[0] || !data.content[0].text) {
    throw new Error(`Claude response malformed: ${JSON.stringify(data).slice(0, 500)}`);
  }
  return data.content[0].text;
}

function repairCommonJsonMistakes(text) {
  return text
    .replace(/[“”]/g, '"')           // curly double quotes -> "
    .replace(/[‘’]/g, "'")           // curly single quotes -> '
    .replace(/,(\s*[}\]])/g, '$1')             // trailing comma before } or ]
    .replace(/}(\s*){/g, '},$1{')              // missing comma between adjacent objects
    .replace(/](\s*)\[/g, '],$1[');            // missing comma between adjacent arrays
}

function parseJsonFromResponse(text) {
  let cleaned = String(text || '').trim();
  if (cleaned.startsWith('```')) {
    cleaned = cleaned.replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/, '');
  }
  try { return JSON.parse(cleaned); } catch (_) {}

  const match = cleaned.match(/\{[\s\S]*\}/);
  const candidate = match ? match[0] : cleaned;
  try { return JSON.parse(candidate); } catch (_) {}

  try { return JSON.parse(repairCommonJsonMistakes(candidate)); } catch (_) {}

  const head = String(text).slice(0, 300).replace(/\n/g, '\\n');
  throw new Error(`Could not parse JSON from Claude response (len=${String(text).length}, head: ${head})`);
}

function dumpBadResponse(label, response) {
  try {
    const dumpDir = path.join(__dirname, 'logs');
    fs.mkdirSync(dumpDir, { recursive: true });
    const stamp = new Date().toISOString().replace(/[:.]/g, '-');
    const dumpPath = path.join(dumpDir, `bad-claude-${label}-${stamp}.txt`);
    fs.writeFileSync(dumpPath, String(response == null ? '' : response));
    return dumpPath;
  } catch (_) { return null; }
}

// Call Claude, parse the JSON it returns. On parse failure, retry once with
// an explicit "your last reply was invalid JSON" prompt; if that ALSO fails,
// dump the raw response to logs/ for debugging and surface a clear error.
async function callAndParse(label, userMessage) {
  let response = await callClaude(userMessage);
  try {
    return parseJsonFromResponse(response);
  } catch (firstErr) {
    const retryMsg = userMessage +
      `\n\n--- RETRY ---\n` +
      `Your previous reply could not be parsed as JSON. Parser said: ${String(firstErr.message).slice(0, 300)}\n` +
      `Re-emit ONLY the JSON object, valid this time. Critical: escape every internal " as \\". ` +
      `No trailing commas. No markdown fences. No commentary.`;
    response = await callClaude(retryMsg);
    try {
      return parseJsonFromResponse(response);
    } catch (secondErr) {
      const dumpPath = dumpBadResponse(label, response);
      const where = dumpPath ? ` Raw response saved to ${dumpPath}.` : '';
      throw new Error(`Claude returned invalid JSON twice for ${label}.${where} Original error: ${firstErr.message}`);
    }
  }
}

// Tolerant field accessor: case-insensitive name match, and unwraps linked-record
// arrays / option objects so the synthesizer sees a plain string regardless of
// how Airtable shaped the field.
function getField(fields, name) {
  if (!fields) return undefined;
  const lower = String(name).toLowerCase().trim();
  for (const k of Object.keys(fields)) {
    if (String(k).toLowerCase().trim() === lower) return fields[k];
  }
  return undefined;
}

function asString(v) {
  if (v == null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number' || typeof v === 'boolean') return String(v);
  if (Array.isArray(v)) {
    if (!v.length) return '';
    const first = v[0];
    if (typeof first === 'string' || typeof first === 'number') return String(first);
    if (first && typeof first === 'object') return String(first.name || first.value || first.text || first.label || '');
  }
  if (typeof v === 'object') return String(v.name || v.value || v.text || v.label || '');
  return String(v);
}

function aiAnswersText(fields) {
  const raw =
    getField(fields, 'AI Answers') ??
    getField(fields, 'AI Answer') ??
    getField(fields, 'AI answers') ??
    getField(fields, 'AI answer');
  return asString(raw).replace(/\s+/g, ' ').trim();
}

function synthesizeProjectDetails(fields, projectFacts) {
  const bt = asString(getField(fields, 'Building Type')).toLowerCase().trim();
  const cr = asString(getField(fields, 'Condo/Rental')).toLowerCase().trim();
  const u  = getField(fields, 'Units');
  const hasCommercialSqft = !!(projectFacts && projectFacts.hasCommercial);

  const isMixed       = bt.includes('mixed') || hasCommercialSqft;
  const isCommercial  = !isMixed && bt.includes('commercial');
  const isResidential = !isMixed && bt.includes('residential');

  if (isCommercial) return 'Commercial space';
  if ((u == null || u === '') && hasCommercialSqft) return 'Commercial space';
  if (u == null || u === '' || (!isResidential && !isMixed)) return null;

  // Tolerant condo detection: "Condo", "condo", "Condo building" all count.
  // "Condo & Rental letters" mentions both so we default to rental (safer).
  const isCondo = cr.includes('condo') && !cr.includes('rental') && !cr.includes('rent ');
  const word = isCondo ? 'residential condominium units' : 'residential rental units';

  if (isMixed) {
    return isCondo
      ? `${u} ${word} and one commercial condominium`
      : `${u} ${word} and commercial space`;
  }
  if (isResidential) return `${u} ${word}`;
  return null;
}

function sanitizeRecordFields(fields, templateFieldName, outputFieldName, templateSelectFieldName) {
  const clean = {};
  for (const [key, value] of Object.entries(fields || {})) {
    if (key === templateFieldName || key === outputFieldName || key === templateSelectFieldName) continue;
    if (Array.isArray(value) && value.length && value[0] && value[0].url) continue; // attachments
    if (value == null || value === '') continue;
    clean[key] = value;
  }
  return clean;
}

function simplifySchema(schema, skipFieldNames) {
  const skip = new Set(skipFieldNames.filter(Boolean));
  return (schema || [])
    .filter((f) => !skip.has(f.name))
    .map((f) => ({ name: f.name, type: f.type }));
}

// Tool definition: Claude returns BOTH literal value swaps (addresses, dates, blocks,
// lots, etc.) AND a list of "configuration claims" — spans of text anywhere in the
// template that describe the OLD project's building configuration (residential vs.
// commercial vs. mixed-use, unit count, condo vs. rental, community facility, story
// count, etc.) regardless of how they're labeled. JS code then compares each claim
// to Airtable and rewrites with a deterministic synthesized phrase.
const DOCX_ANALYSIS_TOOL = {
  name: 'submit_template_analysis',
  description:
    'Submit the value swaps and building-configuration claims found in the template. ' +
    'Use value_swaps for project-specific values like addresses, dates, blocks, lots, owner names. ' +
    'Use configuration_claims for any span describing the OLD building configuration — JS will rewrite those from Airtable fields.',
  input_schema: {
    type: 'object',
    properties: {
      value_swaps: {
        type: 'array',
        description:
          'Literal-text swaps for project-specific values: addresses, city/state, block & lot, dates, ' +
          'owner/LLC names, GSF, assessed values, AND bracketed placeholders like [Project Address], ' +
          '[Block & Lot Number], [City, State, ZIP], [Total Gross Square Feet], [Number]. ALWAYS swap ' +
          'bracketed placeholders verbatim with values from NEW RECORD DATA. ' +
          'ALSO USE value_swaps FOR ALTERNATIVE SCENARIO SENTENCES: when the template has multiple full ' +
          'sentences describing alternative versions of the project (e.g., one for new construction + ' +
          'residential, one for new construction + mixed-use, one for alteration + residential, etc.), ' +
          'check each against NEW RECORD DATA\'s Building Type / Permit Type. For sentences that DO NOT ' +
          'match, emit a value_swap with oldValue=<the full sentence verbatim> and newValue="" (empty ' +
          'string — the handler deletes the whole paragraph). For sentences that DO match, emit small ' +
          'value_swaps for the specific values inside (unit count number, GSF, etc.) — KEEP the sentence ' +
          'structure intact. Use enough context in oldValue to make matches unambiguous. ' +
          'Do NOT include statutory references, firm names, agency names, or boilerplate.',
        items: {
          type: 'object',
          properties: {
            fieldName: { type: 'string', description: 'The Airtable field this swap relates to (or "delete alternative scenario").' },
            oldValue: { type: 'string', description: 'Exact verbatim text from the template to replace.' },
            newValue: { type: 'string', description: 'Replacement text for the new project. Use "" to delete a paragraph entirely.' }
          },
          required: ['fieldName', 'oldValue', 'newValue']
        }
      },
      configuration_claims: {
        type: 'array',
        description:
          'ONLY use this for SHORT noun-phrase summaries (3-15 words, no verbs) appearing AFTER a label like ' +
          '"Project Details:", "Proposed Construction:", "Building Description:", or "Project Description:". ' +
          'Examples of VALID configuration_claims: "99 residential rental units and commercial space", ' +
          '"7 (seven) story commercial and community facility building", "residential rental project". ' +
          'NEVER use this for full sentences with verbs (e.g., "A new mixed-use building will be built...", ' +
          '"The existing building will be converted...", "The project involves..."). Sentences go in ' +
          'value_swaps (either delete or number-swap). NEVER use this for sentences describing alternative ' +
          'scenarios. NEVER use this for sentences describing specific square footage breakdowns. ' +
          'Return only the verbatim summary text (NOT the label). JS replaces each with the synthesized phrase.',
        items: {
          type: 'object',
          properties: {
            exactText: {
              type: 'string',
              description: 'Short verbatim noun-phrase summary (3-15 words, no full sentence with verb).'
            },
            claim: {
              type: 'string',
              description: 'Brief tag (e.g. "building-type", "unit-count", "use-mix").'
            }
          },
          required: ['exactText']
        }
      }
    },
    required: ['value_swaps']
  }
};

// Phase 1: swap mode for Word docs — find old project values, replace with
// the new project's values. Returns an array of { fieldName, oldValue, newValue }.
function formatProjectFactsForPrompt(projectFacts, opts) {
  opts = opts || {};
  if (!projectFacts || typeof projectFacts !== 'object') return '';
  const lines = [];

  if (projectFacts.units != null) {
    lines.push(`Units: ${projectFacts.units}`);
    lines.push(`Building size: ${projectFacts.buildingSize}`);
    lines.push(`Unit-section decision: keep the ${projectFacts.keepUnitSection} section; delete the ${projectFacts.deleteUnitSection} section.`);
  } else {
    lines.push('Units: missing or not numeric; do not delete capped/transitional sections based on units.');
  }

  lines.push(`Has commercial: ${projectFacts.hasCommercialText || (projectFacts.hasCommercial ? 'Yes' : 'No')}`);
  if (!opts.suppressProjectGrossSqft && projectFacts.commercialGrossSqft != null) {
    lines.push(`Commercial Gross SQFT: ${projectFacts.commercialGrossSqft}`);
  }
  if (projectFacts.permitType && projectFacts.permitType.raw) {
    lines.push(`Permit Type: ${projectFacts.permitType.raw}`);
    if (projectFacts.keepPermitScenario === 'new-building') {
      lines.push('Permit-scenario decision: keep new-building/new-construction language; delete alteration/conversion language.');
    } else if (projectFacts.keepPermitScenario === 'alteration-conversion') {
      lines.push('Permit-scenario decision: keep alteration/conversion language; delete new-building/new-construction language.');
    }
  }

  const icap = projectFacts.icap;
  if (icap && icap.isIcap) {
    const where = icap.borough && icap.block ? `${icap.borough.name} block ${icap.block}` : 'missing borough/block';
    if (icap.term) {
      lines.push(`ICAP term: ${icap.term} (${where}; ${icap.reason}).`);
    } else {
      lines.push(`ICAP term: unknown (${where}; ${icap.reason}). Leave existing ICAP term text unchanged if uncertain.`);
    }
  } else {
    lines.push('ICAP term: not applicable because TIP does not include ICAP.');
  }

  return lines.length
    ? `\n\nPRE-DETERMINED PROJECT FACTS (use these exactly; do not recalculate):\n- ${lines.join('\n- ')}\n`
    : '';
}

function formatMarkedTargetsForPrompt(markedTargets) {
  const items = (Array.isArray(markedTargets) ? markedTargets : [])
    .filter((item) => item && item.text)
    .slice(0, 100)
    .map((item) => ({
      text: item.text,
      marker: item.yellowHighlight && item.redText
        ? 'yellow highlight + red text'
        : (item.yellowHighlight ? 'yellow highlight' : 'red text')
    }));

  if (!items.length) return '';
  return `\n\nYELLOW/RED MARKED TEXT TARGETS (high priority blanks/review text; fill from Airtable when possible, otherwise leave unchanged):\n` +
    `${JSON.stringify(items, null, 2)}\n`;
}

function addRuleBackstopSwaps(swaps, templateText, projectFacts, log) {
  const dlog = typeof log === 'function' ? log : (() => {});
  const out = [...swaps];
  const paragraphs = String(templateText || '')
    .split(/\n+/)
    .map((p) => p.replace(/\s+/g, ' ').trim())
    .filter(Boolean);

  const alreadyHasOld = (oldValue) => out.some((s) => s.oldValue === oldValue);
  const addDelete = (fieldName, paragraph, reason) => {
    if (!paragraph || alreadyHasOld(paragraph)) return;
    out.push({ fieldName, oldValue: paragraph, newValue: '' });
    dlog(`[rule-debug] delete ${fieldName}: ${reason}: ${JSON.stringify(paragraph.slice(0, 100))}`);
  };

  if (projectFacts && projectFacts.deleteUnitSection) {
    for (const paragraph of paragraphs) {
      const lower = paragraph.toLowerCase();
      if (projectFacts.deleteUnitSection === 'capped') {
        if (/\b(cap{1,2}ed|caped)\b/i.test(paragraph) || /\b2\s*[abc]\b/i.test(paragraph)) {
          addDelete('Unit Section Rule', paragraph, 'Units > 10 so capped/2A/2B/2C language does not apply');
        }
      } else if (projectFacts.deleteUnitSection === 'transitional') {
        if (lower.includes('transitional')) {
          addDelete('Unit Section Rule', paragraph, 'Units <= 10 so transitional language does not apply');
        }
      }
    }
  }

  if (projectFacts && projectFacts.deletePermitScenario) {
    for (const paragraph of paragraphs) {
      const lower = paragraph.toLowerCase();
      if (projectFacts.deletePermitScenario === 'alteration-conversion') {
        if (
          lower.includes('alteration') ||
          lower.includes('conversion') ||
          lower.includes('converted') ||
          lower.includes('enlargement')
        ) {
          addDelete('Permit Type Rule', paragraph, 'Permit Type is NB/new-building');
        }
      } else if (projectFacts.deletePermitScenario === 'new-building') {
        if (
          lower.includes('new building') ||
          lower.includes('new construction') ||
          /a new .* building/.test(lower)
        ) {
          addDelete('Permit Type Rule', paragraph, 'Permit Type is alteration/conversion');
        }
      }
    }
  }

  const icap = projectFacts && projectFacts.icap;
  if (icap && icap.isIcap && icap.term) {
    const wrongTerm = icap.term === '25-year' ? '15-year' : '25-year';
    for (const paragraph of paragraphs) {
      const lower = paragraph.toLowerCase();
      if (!lower.includes('icap')) continue;
      if (!lower.includes(wrongTerm)) continue;
      if (lower.includes(icap.term)) continue;
      addDelete('ICAP Term Rule', paragraph, `ICAP lookup selected ${icap.term}`);
    }
  }

  if (projectFacts && !projectFacts.hasCommercial) {
    for (const paragraph of paragraphs) {
      const lower = paragraph.toLowerCase();
      if (!lower.includes('commercial')) continue;
      if (lower.includes('residential') && lower.includes('commercial')) {
        addDelete('Commercial Rule', paragraph, 'Commercial Gross SQFT is blank or zero');
      }
    }
  }

  return out;
}

function looksLikeStatutoryThreshold(text) {
  const s = String(text || '');
  return (
    /\b(or more|no more than|fewer than|at least|not less than|more than ten|fewer than one hundred|100 or more|150 or more|six or more)\b/i.test(s) ||
    /\b(wage requirements|prevailing wages|affordability option|AMI|dwelling units|eligible site|modest rental projects?)\b/i.test(s)
  );
}

function looksLikeProtectedAffordabilityText(text) {
  const s = String(text || '');
  return (
    /\b(affordable units?|affordability option|AMI|area median income|affordable rates?|affordable rents?|leased at affordable|Option\s+[ABC])\b/i.test(s) ||
    /\b\d+%\s+of\s+the\s+units\b/i.test(s)
  );
}

async function mapDocxSwaps(templateText, recordFields, schema, opts) {
  opts = opts || {};
  const data = sanitizeRecordFields(
    recordFields,
    opts.templateFieldName,
    opts.outputFieldName,
    opts.templateSelectFieldName
  );
  const simpleSchema = simplifySchema(schema, [
    opts.templateFieldName,
    opts.outputFieldName,
    opts.templateSelectFieldName
  ]);

  const dlog = (opts && typeof opts.log === 'function') ? opts.log : (() => {});
  const projectFacts = opts.projectFacts || {};
  const aiAnswer = aiAnswersText(recordFields);
  if (aiAnswer) {
    for (const key of Object.keys(data)) {
      if (/^(Residential Gross SQFT|Commercial Gross SQFT|Gross SQFT|Total GSF|GSF)$/i.test(key.trim())) {
        delete data[key];
      }
    }
  }
  const markedTargets = Array.isArray(opts.markedTargets) ? opts.markedTargets : [];
  const synth = synthesizeProjectDetails(recordFields, projectFacts);
  dlog(`[cfg-debug] Building Type field: ${JSON.stringify(getField(recordFields, 'Building Type'))}`);
  dlog(`[cfg-debug] Condo/Rental field: ${JSON.stringify(getField(recordFields, 'Condo/Rental'))}`);
  dlog(`[cfg-debug] Units field: ${JSON.stringify(getField(recordFields, 'Units'))}`);
  dlog(`[cfg-debug] Commercial Gross SQFT field: ${JSON.stringify(getField(recordFields, 'Commercial Gross SQFT'))}`);
  if (aiAnswer) dlog('[cfg-debug] AI Answers present: suppressing Residential/Commercial Gross SQFT as project-detail sources');
  dlog(`[rule-debug] Project facts: ${JSON.stringify(projectFacts)}`);
  dlog(`[mark-debug] Yellow/red marked targets found: ${markedTargets.length}`);
  dlog(`[cfg-debug] Synthesized phrase: ${JSON.stringify(synth)}`);
  const factsLine = formatProjectFactsForPrompt(projectFacts, { suppressProjectGrossSqft: !!aiAnswer });
  const aiAnswerLine = aiAnswer
    ? `\n\nAI ANSWERS PROJECT DETAIL OVERRIDE:\n` +
      `The record has AI Answers. Treat AI Answers as the source for the numbered Project details paragraph. ` +
      `Do not use Residential Gross SQFT, Commercial Gross SQFT, Total GSF, or GSF to create or fill a separate gross-square-foot project-detail sentence.\n`
    : '';
  const markedLine = formatMarkedTargetsForPrompt(markedTargets);
  const synthLine = synth
    ? `\n\nSYNTHESIZED BUILDING CONFIGURATION (computed from Airtable Building Type / Condo/Rental / Units): "${synth}"\n` +
      `JS code will replace every configuration_claims span you return with this exact string.\n` +
      `RULE: return EVERY span making any claim about the building's categorical type, use mix, unit count, ` +
      `condo vs. rental, residential vs. commercial vs. mixed-use vs. community facility, story count, or ` +
      `phys configuration. Be MAXIMALLY AGGRESSIVE — when in doubt, INCLUDE the span.\n` +
      `Skip a span ONLY if its wording is character-for-character identical to "${synth}". Do NOT skip because ` +
      `the categorical type "mostly matches" or because the span "adds useful info" — Airtable is the source of ` +
      `truth and any extra detail in the OLD template (community facility, X-story, etc.) MUST be discarded if ` +
      `Airtable doesn't say so.`
    : `\n\nNo synthesized building configuration available for this record (Building Type or Units missing). ` +
      `Still return any configuration_claims you find — JS will leave them alone if no synthesis is possible.`;

  const userMessage =
    `DOCUMENT TYPE: Word (.docx) — SWAP MODE WITH CONFIGURATION CLAIMS\n\n` +
    `AIRTABLE SCHEMA (name + type of every field in the table):\n` +
    `${JSON.stringify(simpleSchema, null, 2)}\n\n` +
    `NEW RECORD DATA (JSON, values for the NEW project only — any field missing ` +
    `here should NOT be swapped):\n${JSON.stringify(data, null, 2)}` +
    factsLine +
    aiAnswerLine +
    synthLine +
    markedLine + `\n\n` +
    `TEMPLATE TEXT (the OLD project's filled-in document):\n"""\n${templateText}\n"""\n\n` +
    `Call the submit_template_analysis tool with two arrays: value_swaps (literal text swaps for ` +
    `project-specific values) and configuration_claims (spans describing the OLD building configuration ` +
    `that may need rewriting from Airtable). Use PRE-DETERMINED PROJECT FACTS as binding decisions: ` +
    `delete the wrong unit-size section, delete the wrong permit scenario, use the listed ICAP term, and keep/remove commercial language ` +
    `based on Has commercial. Fill marked yellow/red targets whenever Airtable provides a matching value.`;

  const parsed = await callClaude(userMessage, { tool: DOCX_ANALYSIS_TOOL });

  const rawSwaps = Array.isArray(parsed.value_swaps) ? parsed.value_swaps : [];
  let swaps = rawSwaps.filter((s) =>
    s && typeof s.oldValue === 'string' && s.oldValue.length &&
    typeof s.newValue === 'string' && s.oldValue !== s.newValue
  );
  swaps = swaps.filter((s) => {
    const old = String(s.oldValue || '');
    const next = String(s.newValue || '');
    if (looksLikeProtectedAffordabilityText(old) && next.trim() === '') {
      dlog(`[cfg-debug] dropped protected affordability deletion: ${JSON.stringify(old.slice(0, 120))}`);
      return false;
    }
    return true;
  });
  if (aiAnswer) {
    swaps = swaps.filter((s) => {
      const field = String(s.fieldName || '');
      if (/\b(?:Residential Gross SQFT|Commercial Gross SQFT|Gross SQFT|Total GSF|GSF|gross square feet|square footage)\b/i.test(field)) {
        dlog(`[cfg-debug] dropped gross-SQFT swap because AI Answers is present: ${JSON.stringify(s.oldValue.slice(0, 100))}`);
        return false;
      }
      return true;
    });
  }
  if (synth) {
    swaps = swaps.filter((s) => {
      const old = String(s.oldValue || '');
      const next = String(s.newValue || '');
      const genericBuildingPhrase =
        /\bbuilding\b/i.test(old) &&
        /\b(mixed[- ]use|residential|commercial|multiple dwelling)\b/i.test(old) &&
        !/\d/.test(old);
      const replacingWithUnitPhrase = next === synth && /\b(residential|commercial).*\bunits?\b|\bcommercial space\b/i.test(next);
      if (genericBuildingPhrase && replacingWithUnitPhrase) {
        dlog(`[cfg-debug] dropped grammar-risk swap: ${JSON.stringify(old)} -> ${JSON.stringify(next)}`);
        return false;
      }
      if (looksLikeStatutoryThreshold(old) && /\b(residential|commercial).*\bunits?\b|\bcommercial space\b/i.test(next)) {
        dlog(`[cfg-debug] dropped statutory-threshold swap: ${JSON.stringify(old)} -> ${JSON.stringify(next)}`);
        return false;
      }
      if (/\b(Martin Joseph|Metropolitan Realty Exemptions)\b/i.test(old)) {
        dlog(`[cfg-debug] dropped signature/firm swap: ${JSON.stringify(old)} -> ${JSON.stringify(next)}`);
        return false;
      }
      return true;
    });
  }

  // For each configuration_claim Claude found, if it's a SHORT noun-phrase
  // summary that appears in the template AND differs from synth, force-rewrite.
  // Sentence guard: skip spans that look like full sentences — those should
  // have been handled via value_swaps (delete or number-swap), not replacement.
  // Replacing a full sentence with a noun phrase produces grammatical nonsense.
  const SENTENCE_STARTERS = /^(A\s|An\s|The\s|We\s|It\s|This\s|These\s|That\s|Those\s|Since\s|During\s|If\s|For\s|In\s|At\s|After\s|Before\s|When\s|Where\s|While\s)/i;
  const looksLikeSentence = (s) => {
    if (!s) return false;
    if (s.length > 120) return true;
    if (SENTENCE_STARTERS.test(s)) return true;
    // Verb hint: ends with period or contains common verb forms
    if (/\.\s*$/.test(s)) return true;
    if (/\b(will be|is being|was|were|will|shall|may|can|should|involves|features|consists|includes|contains)\b/i.test(s)) return true;
    return false;
  };
  const claims = Array.isArray(parsed.configuration_claims) ? parsed.configuration_claims : [];
  dlog(`[cfg-debug] Claude returned ${claims.length} configuration_claims`);
  if (synth) {
    for (const claim of claims) {
      const exact = String((claim && claim.exactText) || '').trim();
      if (!exact || exact === synth) { dlog(`[cfg-debug]   skip (empty or identical to synth): ${JSON.stringify(exact.slice(0,60))}`); continue; }
      if (!templateText.includes(exact)) { dlog(`[cfg-debug]   skip (not in templateText): ${JSON.stringify(exact.slice(0,60))}`); continue; }
      if (looksLikeStatutoryThreshold(exact)) {
        dlog(`[cfg-debug]   skip (statutory threshold, not project config): ${JSON.stringify(exact.slice(0,80))}`);
        continue;
      }
      if (looksLikeSentence(exact)) {
        dlog(`[cfg-debug]   skip (looks like a sentence, not a label-line summary): ${JSON.stringify(exact.slice(0,80))}`);
        continue;
      }
      // Drop any value_swaps whose oldValue is contained in this claim span,
      // so we don't double-swap (model's small swap + our big rewrite).
      swaps = swaps.filter((s) => !(s.oldValue.length >= 4 && exact.includes(s.oldValue)));
      swaps.push({
        fieldName: 'Building Configuration',
        oldValue: exact,
        newValue: synth
      });
      dlog(`[cfg-debug]   PUSHED swap from configuration_claim: ${JSON.stringify(exact.slice(0,60))}`);
    }
  }

  // Deterministic bracketed placeholder backstop. Templates often have
  // [Project Address], [Block & Lot Number], [City, State, ZIP], etc. that
  // Claude inconsistently swaps. Fill these from Airtable fields directly.
  const placeholderMap = [
    { brackets: ['[Project Address]', '[Property Address]', '[Address]'], value: asString(getField(recordFields, 'Property Address') || getField(recordFields, 'Address') || getField(recordFields, 'Project Address')) },
    { brackets: ['[Block & Lot Number]', '[Block & Lot]', '[Block and Lot]'], value: (() => {
      const b = asString(getField(recordFields, 'Block')); const l = asString(getField(recordFields, 'Lot'));
      return (b && l) ? `${b} - ${l}` : '';
    })() },
    { brackets: ['[City, State, ZIP]', '[City, State]', '[City]'], value: (() => {
      const b = asString(getField(recordFields, 'Borough'));
      return b ? `${b}, NY` : '';
    })() },
    { brackets: ['[Residential Gross SQFT]', '[Total Gross Square Feet]', '[Gross Square Feet]', '[GSF]'], value: aiAnswer ? '' : asString(getField(recordFields, 'Residential Gross SQFT') || getField(recordFields, 'Gross SQFT') || getField(recordFields, 'Total GSF') || getField(recordFields, 'GSF')) },
    { brackets: ['[Units]', '[Number of Units]', '[Unit Count]'], value: asString(getField(recordFields, 'Units')) },
    { brackets: ['[Commercial Gross SQFT]', '[Commercial Gross Sq Ft]', '[Commercial Square Feet]', '[Commercial GSF]'], value: aiAnswer ? '' : asString(getField(recordFields, 'Commercial Gross SQFT')) },
    { brackets: ['[ICAP Term]', '[ICAP Years]', '[Abatement Term]'], value: projectFacts && projectFacts.icap ? asString(projectFacts.icap.term) : '' }
  ];
  for (const map of placeholderMap) {
    if (!map.value) continue;
    for (const ph of map.brackets) {
      if (templateText.includes(ph)) {
        // Drop any existing swap on this placeholder
        swaps = swaps.filter((s) => s.oldValue !== ph);
        swaps.push({ fieldName: 'Bracketed Placeholder', oldValue: ph, newValue: map.value });
        dlog(`[cfg-debug] placeholder swap: ${ph} -> ${JSON.stringify(map.value)}`);
      }
    }
  }

  // Deterministic backstop for known labeled "what is this project" lines. Even
  // if the model misses any of these as a configuration_claim, this catches them
  // by literal label match. Each label is matched globally to handle templates
  // that repeat the same label (e.g., footer + body).
  if (synth) {
    const LABELS = [
      'Project Details',
      'Proposed Construction',
      'Building Description',
      'Project Description'
    ];
    const escRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    for (const label of LABELS) {
      // Tolerant colon: regular ASCII : OR fullwidth ：. Optional whitespace either side.
      const re = new RegExp(escRe(label) + '\\s*[:：]\\s*([^\\r\\n]+)', 'gi');
      let m;
      let hits = 0;
      while ((m = re.exec(templateText)) !== null) {
        hits++;
        const oldFull = m[1].trim();
        if (!oldFull || oldFull === synth) {
          dlog(`[cfg-debug] backstop "${label}": match found but skipping (empty or identical to synth)`);
          continue;
        }
        swaps = swaps.filter((s) => {
          const fn = String(s.fieldName || '').toLowerCase();
          if (fn === label.toLowerCase()) return false;
          if (s.oldValue.length >= 4 && oldFull.includes(s.oldValue)) return false;
          return true;
        });
        swaps.push({
          fieldName: label,
          oldValue: oldFull,
          newValue: synth
        });
        dlog(`[cfg-debug] backstop "${label}": PUSHED swap, oldFull=${JSON.stringify(oldFull.slice(0,80))}`);
      }
      if (!hits) dlog(`[cfg-debug] backstop "${label}": no matches in templateText`);
    }
  } else {
    dlog(`[cfg-debug] Backstop SKIPPED entirely because synth is null. Check Building Type field above.`);
  }

  swaps = addRuleBackstopSwaps(swaps, templateText, projectFacts, dlog);
  return swaps;
}

async function mapXlsxSwaps(workbookJson, recordFields, schema, opts) {
  opts = opts || {};
  const data = sanitizeRecordFields(
    recordFields,
    opts.templateFieldName,
    opts.outputFieldName,
    opts.templateSelectFieldName
  );
  const simpleSchema = simplifySchema(schema || [], [
    opts.templateFieldName,
    opts.outputFieldName,
    opts.templateSelectFieldName
  ]);

  const userMessage =
    `DOCUMENT TYPE: Excel (.xlsx) — SWAP MODE (multi-sheet)\n\n` +
    `AIRTABLE SCHEMA (name + type of every field in the table):\n` +
    `${JSON.stringify(simpleSchema, null, 2)}\n\n` +
    `NEW RECORD DATA (JSON, values for the NEW project only):\n` +
    `${JSON.stringify(data, null, 2)}\n\n` +
    `WORKBOOK STRUCTURE (JSON — every sheet, every cell with value + hasFormula flag):\n` +
    `${workbookJson}\n\n` +
    `Call the submit_xlsx_swaps tool with cell-level swaps grouped by sheet. ` +
    `Never include cells where hasFormula is true. Never include label cells.`;

  const XLSX_SWAPS_TOOL = {
    name: 'submit_xlsx_swaps',
    description: 'Submit cell-level swaps for the workbook, grouped by sheet name.',
    input_schema: {
      type: 'object',
      properties: {
        sheets: {
          type: 'object',
          description: 'Keys are sheet names; values are arrays of cell swaps for that sheet.',
          additionalProperties: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                cellRef: { type: 'string' },
                fieldName: { type: 'string' },
                oldValue: {},
                newValue: {}
              },
              required: ['cellRef', 'newValue']
            }
          }
        }
      },
      required: ['sheets']
    }
  };

  const parsed = await callClaude(userMessage, { tool: XLSX_SWAPS_TOOL });
  const sheets = parsed.sheets || {};
  // Filter out invalid entries defensively
  const clean = {};
  for (const [sheetName, arr] of Object.entries(sheets)) {
    if (!Array.isArray(arr)) continue;
    clean[sheetName] = arr.filter((s) =>
      s && typeof s.cellRef === 'string' && s.cellRef.length &&
      // newValue must exist (null/undefined is meaningless to write)
      s.newValue != null
    );
  }
  return clean;
}

async function mapPdfFormFields(formFieldList, recordFields, templateFieldName, outputFieldName) {
  const data = sanitizeRecordFields(recordFields, templateFieldName, outputFieldName);
  const userMessage =
    `DOCUMENT TYPE: PDF form (fillable AcroForm .pdf)\n\n` +
    `FORM FIELDS (JSON array — each entry shows name, type, and allowed options for choice fields):\n` +
    `${JSON.stringify(formFieldList, null, 2)}\n\n` +
    `RECORD DATA (JSON):\n${JSON.stringify(data, null, 2)}\n\n` +
    `Call the submit_pdf_form_fields tool with values keyed by EXACT field name from the FORM FIELDS list. ` +
    `Values must match the field type (string for text, boolean for checkbox, allowed option string for dropdown/radio).`;

  const PDF_FORM_TOOL = {
    name: 'submit_pdf_form_fields',
    description: 'Submit field values for the fillable PDF form.',
    input_schema: {
      type: 'object',
      properties: {
        formFields: {
          type: 'object',
          description: 'Object keyed by exact PDF field name. Values are strings or booleans matching the field type. Omit any field you cannot confidently map.',
          additionalProperties: true
        }
      },
      required: ['formFields']
    }
  };

  const parsed = await callClaude(userMessage, { tool: PDF_FORM_TOOL });
  return parsed.formFields || {};
}

module.exports = { callClaude, mapDocxSwaps, mapXlsxSwaps, mapPdfFormFields };
