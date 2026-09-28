import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, writeFileSync, chmodSync, copyFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, delimiter } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';

const entry = fileURLToPath(new URL('../provision-qstar-beta-m365.sh', import.meta.url));
const schema = JSON.parse(readFileSync(new URL('../region-schema.json', import.meta.url), 'utf8'));
function fixture(t) {
  const dir = mkdtempSync(join(tmpdir(), 'qstar-provision-'));
  t.after(() => rmSync(dir, { recursive: true, force: true }));
  const stateFile = join(dir, 'state.json');
  const stub = join(dir, 'm365');
  copyFileSync(new URL('./m365-stub.mjs', import.meta.url), stub);
  chmodSync(stub, 0o755);
  writeFileSync(stateFile, JSON.stringify({ lists: {}, commands: [] }));
  return {
    read: () => JSON.parse(readFileSync(stateFile, 'utf8')),
    write: state => writeFileSync(stateFile, JSON.stringify(state)),
    run: (mode = 'preserve') => spawnSync('bash', [entry], { encoding: 'utf8',
      env: {...process.env, PATH: `${dir}${delimiter}${process.env.PATH}`, SITE: 'https://example.invalid/sites/qstar',
        QSTAR_TEST_STATE: stateFile, REGION_MIGRATION: mode, ISSUES: 'Q-Star Issues', CONFIG: 'Q-Star Config', PERSON_AS_TEXT: '0' } }),
  };
}
function succeeds(result) { assert.equal(result.status, 0, result.stderr + result.stdout); }

test('beta launcher provisions complete native-person schema, indexes, and stable reference offset', t => {
  const f = fixture(t);
  succeeds(f.run());
  const {lists, commands} = f.read();
  const fields = lists['Q-Star Issues'].fields;
  for (const name of ['ReportedBy', 'TaskOwner', 'VerifiedBy']) assert.equal(fields[name].TypeAsString, 'User');
  for (const name of ['Status', 'Triaged', 'DueDate']) assert.equal(fields[name].Indexed, true);
  for (const name of ['Severity', 'Region', 'DepartmentBU']) assert.equal(fields[name].Required, true);
  assert.equal(fields.QsNumber.Required, false);
  assert.equal(fields.ReminderCycle.TypeAsString, 'Text');
  for (const list of Object.values(lists)) assert.equal(list.fields.Title.Required, false);
  assert.equal(lists['Q-Star Progress Log'].EnableFolderCreation, true);
  assert.equal(lists['Q-Star Progress Log'].ContentTypesEnabled, true);
  assert.deepEqual(fields.Region.Choices, schema.choices);
  assert.equal(lists['Q-Star Progress Log'].fields.ParentItemId.Required, false);
  assert.equal(lists['Q-Star Config'].items[0].ReferenceOffset, 1000);
  assert.ok(!commands.some(args => args.includes('group') || args.includes('roleassignment')));
});

test('region preview performs no writes and apply preserves references, content, and unknown regions', t => {
  const f = fixture(t);
  succeeds(f.run());
  const state = f.read();
  const items = Object.keys(schema.aliases).map((region, index) => ({Id: index + 1, Region: region, QsNumber: 2000 + index, Description: `Keep ${index}`, Modified: '2026-01-01'}));
  items.push({Id: 7, Region: 'Historic custom region', QsNumber: 3000, Description: 'Keep custom'});
  state.lists['Q-Star Issues'].items = items;
  state.lists['Q-Star Config'].items = [{Id: 2, SettingsJson: '{"msFormUrl":"preserve"}'}];
  f.write(state);
  succeeds(f.run('preview'));
  assert.deepEqual(f.read().lists, state.lists);
  succeeds(f.run('apply'));
  const applied = f.read().lists;
  assert.deepEqual(applied['Q-Star Issues'].items, items.map(row => ({...row, Region: schema.aliases[row.Region] || row.Region})));
  assert.deepEqual(applied['Q-Star Config'].items, [{Id: 2, SettingsJson: '{"msFormUrl":"preserve"}', ReferenceOffset: 3000}]);
  assert.deepEqual(applied['Q-Star Issues'].fields.Region.Choices, [...schema.choices, 'Historic custom region']);
  succeeds(f.run('apply'));
  assert.deepEqual(f.read().lists, applied);
});

test('existing indexes and choices reconcile; immutable offset is not raised after new intake', t => {
  const f = fixture(t);
  succeeds(f.run());
  const state = f.read();
  state.lists['Q-Star Issues'].fields.DueDate.Indexed = false;
  state.lists['Q-Star Issues'].fields.DepartmentBU.Choices = ['IT'];
  state.lists['Q-Star Issues'].fields.DepartmentBU.SchemaXml = "<Field Type='Choice' Name='DepartmentBU'><CHOICES><CHOICE>IT</CHOICE></CHOICES></Field>";
  state.lists['Q-Star Issues'].items = [{Id: 11, Region: 'Germany', QsNumber: 1011}];
  f.write(state);
  succeeds(f.run());
  const lists = f.read().lists;
  assert.equal(lists['Q-Star Issues'].fields.DueDate.Indexed, true);
  assert.ok(lists['Q-Star Issues'].fields.DepartmentBU.Choices.includes('Finance & Controlling'));
  assert.ok(lists['Q-Star Issues'].fields.Region.Choices.includes('Germany'));
  assert.equal(lists['Q-Star Issues'].items[0].Region, 'Germany');
  assert.equal(lists['Q-Star Config'].items[0].ReferenceOffset, 1000);
});

test('a field creation failure stops provisioning instead of claiming it exists', t => {
  const f = fixture(t);
  f.write({lists: {}, commands: [], failField: 'ShortSummary'});
  const result = f.run();
  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /403 permission denied/);
  assert.equal(f.read().lists['Q-Star Config'], undefined);
});

test('unresolved existing journal parent stops the entry point before schema or permission writes', t => {
  const f = fixture(t);
  const original = {
    'Q-Star Issues': {fields:{},items:[{Id:42,TaskOwnerId:7}]},
    'Q-Star Progress Log': {fields:{},items:[{Id:9,FSObjType:0,ParentItemId:999,FileDirRef:'/sites/qstar/Lists/Q-StarProgressLog',FileRef:'/sites/qstar/Lists/Q-StarProgressLog/9_.000'}]},
  };
  f.write({lists:original,commands:[]});
  const result = f.run();
  assert.notEqual(result.status,0);
  assert.match(result.stderr,/no valid existing parent/);
  assert.deepEqual(f.read().lists,original);
  assert.ok(f.read().commands.every(args=>args[0]==='request' ? args[args.indexOf('--method')+1]==='get' : args[2]==='get'));
});
