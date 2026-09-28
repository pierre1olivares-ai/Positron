import test from 'node:test';
import assert from 'node:assert/strict';
import { reconcileAcl, desiredRoles, APPEND_PERMISSIONS, secureQstar, all } from '../qstar-security-engine.mjs';

test('reconciliation revokes former owners, handles clearing, and preserves Admin/QM access', async () => {
  const full = 1073741829, edit = 1073741830, read = 1073741826, append = 1073741900;
  const grants = new Map([[1,new Set([full])],[2,new Set([edit])],[3,new Set([edit])],[4,new Set([read])],[99,new Set([full])]]);
  let writes = 0;
  const request = async (method, path) => {
    if (method === 'GET' && path.includes('HasUniqueRoleAssignments')) return {HasUniqueRoleAssignments: true};
    if (method === 'GET') return {value: [...grants].map(([id, roles]) => ({Member:{Id:id}, RoleDefinitionBindings:[...roles].map(Id=>({Id}))}))};
    writes++;
    const match = path.match(/(add|remove)roleassignment\(principalid=(\d+),roledefid=(\d+)\)/);
    assert.ok(match, path);
    const [,action,principalText,roleText] = match;
    const principal = Number(principalText), role = Number(roleText);
    if (!grants.has(principal)) grants.set(principal, new Set());
    if (action === 'add') grants.get(principal).add(role); else grants.get(principal).delete(role);
  };
  const base = [[1,full],[2,append],[4,read],[99,full]];
  await reconcileAcl(request, 'folder', [...base,[5,append]]);
  assert.equal(grants.get(3).has(edit), false);
  assert.equal(grants.get(2).has(edit), false);
  assert.ok(grants.get(2).has(append));
  assert.ok(grants.get(1).has(full));
  assert.ok(grants.get(5).has(append));
  const previousWrites = writes;
  await reconcileAcl(request, 'folder', [...base,[5,append]]);
  assert.equal(writes, previousWrites, 'repeating a reconciliation must be idempotent');
  await reconcileAcl(request, 'folder', base);
  assert.equal(grants.get(5).has(append), false, 'empty owner removes last owner grant');
});

test('one principal retains the union of intended roles and append excludes edit/delete', () => {
  assert.deepEqual([...desiredRoles([[1,10],[1,20]]).get(1)], [10,20]);
  const low = BigInt(APPEND_PERMISSIONS.Low);
  assert.equal(low & 2n, 2n, 'AddListItems');
  assert.equal(low & 4n, 0n, 'EditListItems');
  assert.equal(low & 8n, 0n, 'DeleteListItems');
});

test('read-only migration preview validates journal parent IDs before permission changes', async () => {
  const methods = [];
  const request = async (method, path) => {
    methods.push(method);
    if (path.includes('RootFolder/ServerRelativeUrl')) return {RootFolder:{ServerRelativeUrl:'/sites/q/Lists/Progress'}};
    if (path.includes('TaskOwnerId')) return {value:[{Id:42,TaskOwnerId:7}]};
    return {value:[{Id:9,FSObjType:0,ParentItemId:42,FileDirRef:'/sites/q/Lists/Progress',FileRef:'/sites/q/Lists/Progress/9_.000'}]};
  };
  const options = {request,site:'https://example.invalid/sites/q',issuesTitle:'Issues',progressTitle:'Progress',configTitle:'Config',production:true,migration:'preview',log:()=>{}};
  await secureQstar(options);
  assert.ok(methods.every(method=>method==='GET'));
  await assert.rejects(secureQstar({...options,migration:'preserve'}), /reviewed migration/);
  assert.ok(methods.every(method=>method==='GET'));
});

test('paging follows every next link', async () => {
  const first = Array.from({length:2000}, (_,Id)=>({Id}));
  const rows = await all(async (_method,path) => path==='first' ? {value:first,'odata.nextLink':'second'} : {value:[{Id:2000}]}, 'first');
  assert.equal(rows.length,2001);
  assert.equal(rows[2000].Id,2000);
});

test('production upgrade moves legacy entries in place and replaces broad list and old-owner grants', async () => {
  const root = '/sites/q/Lists/Progress With Spaces';
  const entry = {Id:90,FSObjType:0,ParentItemId:42,FileDirRef:root,FileRef:`${root}/90_.000`,AuthorId:17,Created:'2026-01-02T03:04:05Z'};
  const original = {...entry};
  const groups = {'Q-Star Admins':1,'Q-Star Quality Managers':2,'Q-Star Task Owners':3,'Q-Star Readers':4};
  const grants = new Map();
  const calls = [];
  const append = 1073741900;
  const request = async (method,path,body) => {
    calls.push({method,path,body});
    if (path.includes('HasUniqueRoleAssignments')) return {HasUniqueRoleAssignments:true};
    if (path.includes('/roleassignments?')) {
      const object = path.split('/roleassignments?')[0];
      if (!grants.has(object)) grants.set(object,new Map([[55,new Set([1073741830])]]));
      return {value:[...grants.get(object)].map(([Id,roles])=>({Member:{Id},RoleDefinitionBindings:[...roles].map(Id=>({Id}))}))};
    }
    if (path.includes('/roleassignments/')) {
      const object = path.split('/roleassignments/')[0];
      const [,action,p,r] = path.match(/(add|remove)roleassignment\(principalid=(\d+),roledefid=(\d+)\)/);
      const map = grants.get(object); if (!map.has(+p)) map.set(+p,new Set());
      if (action === 'add') map.get(+p).add(+r); else map.get(+p).delete(+r);
      return {};
    }
    if (path.includes('RootFolder/ServerRelativeUrl')) return {RootFolder:{ServerRelativeUrl:root}};
    if (path.includes('TaskOwnerId')) return {value:[{Id:42,TaskOwnerId:7}]};
    if (path.includes('FSObjType')) return {value:[entry]};
    if (path === "web/lists/getbytitle('Config')/items?$select=Id") return {value:[{Id:8}]};
    if (path.includes('currentuser')) return {Id:99};
    if (path.includes('sitegroups')) return {Id:groups[decodeURIComponent(path.match(/getbyname\('(.*?)'\)/)[1])]};
    if (path === 'web/roledefinitions?$select=Id,Name,BasePermissions') return {value:[{Id:append,Name:'Q-Star Append Progress',BasePermissions:APPEND_PERMISSIONS}]};
    if (path.includes('/RootFolder/Folders')) return {value:[]};
    if (path.startsWith('SP.MoveCopyUtil')) {
      assert.equal(body.options.RetainEditorAndModifiedOnMove,true);
      assert.ok(body.destPath.DecodedUrl.includes('Progress With Spaces'), 'ResourcePath requires the decoded absolute URL');
      const dest = body.destPath.DecodedUrl.slice('https://example.invalid'.length);
      entry.FileDirRef = dest.slice(0,dest.lastIndexOf('/')); entry.FileRef = dest;
      return {};
    }
    if (method==='GET' && path.includes('/items(90)')) return entry;
    if (method !== 'GET') return {};
    throw new Error(`Unexpected request ${method} ${path}`);
  };
  await secureQstar({request,site:'https://example.invalid/sites/q',issuesTitle:'Issues',progressTitle:'Progress',configTitle:'Config',production:true,migration:'apply',groupNames:Object.keys(groups),log:()=>{}});
  assert.equal(entry.Id,original.Id);
  assert.equal(entry.AuthorId,original.AuthorId);
  assert.equal(entry.Created,original.Created);
  assert.equal(entry.FileDirRef,`${root}/issue-42`);
  const folder = [...grants.entries()].find(([path])=>path.includes('ListItemAllFields'))[1];
  assert.deepEqual([...folder.get(7)], [append]);
  assert.equal(folder.get(55).size,0,'obsolete owner removed');
  const progress = grants.get("web/lists/getbytitle('Progress')");
  assert.deepEqual([...progress.get(2)], [append]);
  assert.deepEqual([...progress.get(3)], [1073741826]);
  assert.deepEqual([...grants.get("web/lists/getbytitle('Config')").get(2)], [1073741826]);
  assert.ok(calls.some(call=>call.path.endsWith('/items(90)/resetroleinheritance()')));
  assert.ok(calls.some(call=>call.path === "web/lists/getbytitle('Config')/items(8)/resetroleinheritance()"), 'existing settings item must inherit Admin-only writes');
});
