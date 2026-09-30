import { readFileSync } from 'node:fs';
import { spawnSync } from 'node:child_process';

const schema = JSON.parse(readFileSync(new URL('./region-schema.json', import.meta.url), 'utf8'));
const site = process.env.SITE;
const issues = process.env.ISSUES || 'Q-Star Issues';
const config = process.env.CONFIG || 'Q-Star Config';
const mode = process.env.REGION_MIGRATION || 'preserve';
if (!site || !['preserve', 'preview', 'apply'].includes(mode)) {
  throw new Error('Set SITE and REGION_MIGRATION=preserve, preview, or apply.');
}

function m365(...args) {
  const result = spawnSync('m365', [...args, '--webUrl', site, '--output', 'json'], {
    encoding: 'utf8', maxBuffer: 64 * 1024 * 1024,
  });
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error(result.stderr || result.stdout || `m365 failed (${result.status})`);
  return result.stdout.trim() ? JSON.parse(result.stdout) : undefined;
}
function xml(value) {
  return String(value).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}
function choicesOf(field) {
  return Array.isArray(field.Choices) ? field.Choices : field.Choices?.results || [];
}
function setChoices(name, choices) {
  const field = m365('spo', 'field', 'get', '--listTitle', issues, '--internalName', name);
  if (field.TypeAsString !== 'Choice') throw new Error(`${name} must be a Choice field; no type conversion was attempted.`);
  const unique = [...new Set(choices)];
  if (JSON.stringify(choicesOf(field)) === JSON.stringify(unique)) return;
  const choiceXml = `<CHOICES>${unique.map(value => `<CHOICE>${xml(value)}</CHOICE>`).join('')}</CHOICES>`;
  if (!/<CHOICES>[\s\S]*?<\/CHOICES>/i.test(field.SchemaXml)) throw new Error(`${name} SchemaXml has no CHOICES element.`);
  m365('spo', 'field', 'set', '--listTitle', issues, '--internalName', name,
    '--SchemaXml', field.SchemaXml.replace(/<CHOICES>[\s\S]*?<\/CHOICES>/i, choiceXml));
}

// Omitting pageSize/pageNumber makes CLI for Microsoft 365 retrieve every page.
const rows = m365('spo', 'listitem', 'list', '--listTitle', issues, '--fields', 'Id,Region,QsNumber');
const changes = rows.filter(row => Object.hasOwn(schema.aliases, row.Region));
for (const row of changes) console.log(`Region item ${row.Id}: ${row.Region} -> ${schema.aliases[row.Region]}`);
if (mode === 'preview') {
  console.log(`Preview only: ${changes.length} region changes; no writes performed.`);
  process.exit(0);
}

function decodeXml(value) {
  return value.replaceAll('&lt;', '<').replaceAll('&gt;', '>').replaceAll('&quot;', '"').replaceAll('&apos;', "'").replaceAll('&amp;', '&');
}
for (const [name, input] of [['Status', process.argv[2]], ['DepartmentBU', process.argv[3]], ['EscalationBU', process.argv[3]]]) {
  if (!input) continue;
  const desired = [...input.matchAll(/<CHOICE>([\s\S]*?)<\/CHOICE>/g)].map(match => decodeXml(match[1]));
  const field = m365('spo', 'field', 'get', '--listTitle', issues, '--internalName', name);
  setChoices(name, [...desired, ...choicesOf(field)]);
}

const regionField = m365('spo', 'field', 'get', '--listTitle', issues, '--internalName', 'Region');
// Accept both old and new values before touching records, preserving custom choices.
setChoices('Region', [...schema.choices, ...choicesOf(regionField), ...rows.map(row => row.Region).filter(Boolean)]);
if (mode === 'apply') {
  for (const row of changes) {
    m365('spo', 'listitem', 'set', '--listTitle', issues, '--id', String(row.Id),
      '--Region', schema.aliases[row.Region], '--systemUpdate');
    row.Region = schema.aliases[row.Region];
  }
  // Unknown historical values remain selectable; never silently reclassify them.
  const custom = choicesOf(regionField).filter(value => !Object.hasOwn(schema.aliases, value));
  setChoices('Region', [...schema.choices, ...custom, ...rows.map(row => row.Region).filter(Boolean)]);
} else if (changes.length) {
  console.log('Legacy regions preserved. Review REGION_MIGRATION=preview, then run REGION_MIGRATION=apply during maintenance.');
}

const settings = m365('spo', 'listitem', 'list', '--listTitle', config, '--fields', 'Id,Title,ReferenceOffset');
if (settings.length > 1) throw new Error(`${config} must contain one settings item; reconcile duplicates before provisioning.`);
const existing = settings[0]?.ReferenceOffset;
if (existing !== null && existing !== undefined) {
  if (!Number.isSafeInteger(existing) || existing < 1000) throw new Error('Existing ReferenceOffset is invalid; it was not overwritten.');
  console.log(`ReferenceOffset ${existing} preserved (immutable after first use).`);
} else {
  let offset = 1000;
  for (const row of rows) {
    if (row.QsNumber === null || row.QsNumber === undefined) continue;
    if (!Number.isSafeInteger(row.QsNumber) || row.QsNumber < 0) throw new Error(`Invalid legacy QsNumber on item ${row.Id}; resolve before allocating references.`);
    offset = Math.max(offset, row.QsNumber);
  }
  if (settings.length) {
    m365('spo', 'listitem', 'set', '--listTitle', config, '--id', String(settings[0].Id), '--ReferenceOffset', String(offset));
  } else {
    m365('spo', 'listitem', 'add', '--listTitle', config, '--Title', 'Q-Star Settings', '--ReferenceOffset', String(offset));
  }
  console.log(`ReferenceOffset initialized to ${offset}; existing QS references were not changed.`);
}
