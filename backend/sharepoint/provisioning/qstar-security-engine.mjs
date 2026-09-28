// Transport-independent SharePoint ACL reconciliation and in-place journal upgrade.
export const FULL = 1073741829;
export const EDIT = 1073741830;
export const READ = 1073741826;
const LIMITED = 1073741825;
export const APPEND_NAME = 'Q-Star Append Progress';
// ViewListItems, AddListItems, OpenItems, ViewVersions, ViewFormPages, Open,
// ViewPages, BrowseUserInfo, UseClientIntegration, UseRemoteAPIs. No edit/delete.
export const APPEND_PERMISSIONS = { High: '48', Low: '134418531' };
const literal = value => encodeURIComponent(String(value).replaceAll("'", "''")).replaceAll("'", '%27');
export const listPath = title => `web/lists/getbytitle('${literal(title)}')`;
const folderPath = path => `web/GetFolderByServerRelativePath(decodedUrl='${literal(path)}')`;
const collection = response => response.value || response.results || response.d?.results || [];
const unwrap = response => response?.d || response;

export async function all(request, path) {
  const rows = [];
  do {
    const response = await request('GET', path);
    rows.push(...collection(response));
    path = response['@odata.nextLink'] || response['odata.nextLink'] || response.d?.__next;
  } while (path);
  return rows;
}

export function desiredRoles(pairs) {
  const roles = new Map();
  for (const [principal, role] of pairs) {
    if (!Number.isInteger(principal) || principal <= 0) throw new Error('Invalid permission principal ID.');
    if (!roles.has(principal)) roles.set(principal, new Set());
    roles.get(principal).add(role);
  }
  return roles;
}

export async function reconcileAcl(request, path, pairs) {
  const desired = desiredRoles(pairs);
  const object = unwrap(await request('GET', `${path}?$select=HasUniqueRoleAssignments`));
  if (!object.HasUniqueRoleAssignments) {
    await request('POST', `${path}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=false)`);
  }
  const assignments = await all(request, `${path}/roleassignments?$select=Member/Id,RoleDefinitionBindings/Id&$expand=Member,RoleDefinitionBindings`);
  const existing = new Map(assignments.map(row => [row.Member.Id,
    new Set((row.RoleDefinitionBindings.results || row.RoleDefinitionBindings).map(role => role.Id))]));
  // Add required access first; remove obsolete bindings explicitly afterwards.
  for (const [principal, roles] of desired) for (const role of roles) {
    if (!existing.get(principal)?.has(role)) {
      await request('POST', `${path}/roleassignments/addroleassignment(principalid=${principal},roledefid=${role})`);
    }
  }
  for (const [principal, roles] of existing) for (const role of roles) {
    // SharePoint manages ancestor Limited Access for uniquely secured children.
    if (role !== LIMITED && !desired.get(principal)?.has(role)) {
      await request('POST', `${path}/roleassignments/removeroleassignment(principalid=${principal},roledefid=${role})`);
    }
  }
}

export async function secureQstar({ request, site, issuesTitle, progressTitle, configTitle,
  production, migration = 'preserve', preflight = false, groupNames, log = console.log }) {
  if (!['preserve', 'preview', 'apply'].includes(migration)) throw new Error('Progress migration must be preserve, preview, or apply.');
  const issuesPath = listPath(issuesTitle), progressPath = listPath(progressTitle), configPath = listPath(configTitle);
  const progress = unwrap(await request('GET', `${progressPath}?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`));
  const root = progress.RootFolder.ServerRelativeUrl.replace(/\/$/, '');
  const issues = await all(request, `${issuesPath}/items?$select=Id,TaskOwnerId&$top=2000`);
  const issueIds = new Set(issues.map(row => row.Id));
  const entries = await all(request, `${progressPath}/items?$select=Id,FSObjType,ParentItemId,FileDirRef,FileRef,AuthorId,Created&$top=2000`);
  const moves = [];
  for (const entry of entries.filter(row => row.FSObjType !== 1)) {
    if (entry.FileDirRef === root) {
      if (!Number.isInteger(entry.ParentItemId) || !issueIds.has(entry.ParentItemId)) {
        throw new Error(`Journal item ${entry.Id} has no valid existing parent; no history was discarded. Correct its mapping before migration.`);
      }
      moves.push(entry);
    } else {
      const leaf = entry.FileDirRef?.slice(root.length + 1);
      if (!entry.FileDirRef?.startsWith(`${root}/`) || !/^issue-[1-9][0-9]*$/.test(leaf) || !issueIds.has(Number(leaf.slice(6)))) {
        throw new Error(`Journal item ${entry.Id} has an unrecognized folder. Review its parent mapping before migration.`);
      }
    }
  }
  for (const entry of moves) log(`Journal item ${entry.Id}: move to ${root}/issue-${entry.ParentItemId}`);
  if (migration === 'preview') { log('Preview only: no changes made.'); return; }
  if (moves.length && migration !== 'apply') {
    throw new Error('Legacy journal rows require a reviewed migration. Run PROGRESS_MIGRATION=preview, then apply during maintenance.');
  }

  if (preflight) { log('Journal migration preflight passed.'); return; }

  const configItems = production ? await all(request, `${configPath}/items?$select=Id`) : [];
  if (production && configItems.length !== 1) throw new Error('Config must contain exactly one provisioned settings item before permissions are applied.');
  const currentUser = unwrap(await request('GET', 'web/currentuser?$select=Id')).Id;
  let groups, appendRole;
  if (production) {
    groups = [];
    for (const name of groupNames) groups.push(unwrap(await request('GET', `web/sitegroups/getbyname('${literal(name)}')?$select=Id`)).Id);
    const roles = await all(request, 'web/roledefinitions?$select=Id,Name,BasePermissions');
    let role = roles.find(item => item.Name === APPEND_NAME);
    if (!role) {
      await request('POST', 'web/roledefinitions', { Name: APPEND_NAME,
        Description: 'Read and append progress; no edit, delete or permission management.', BasePermissions: APPEND_PERMISSIONS });
      role = unwrap(await request('GET', `web/roledefinitions/getbyname('${literal(APPEND_NAME)}')`));
    }
    if (!role.Id || String(role.BasePermissions.High) !== APPEND_PERMISSIONS.High || String(role.BasePermissions.Low) !== APPEND_PERMISSIONS.Low) {
      throw new Error('Q-Star Append Progress has unexpected permissions. Review this permission level before proceeding.');
    }
    appendRole = role.Id;
    const [admin, qm, ownerGroup, reader] = groups;
    await reconcileAcl(request, issuesPath, [[currentUser,FULL],[admin,FULL],[qm,EDIT],[ownerGroup,READ],[reader,READ]]);
    await reconcileAcl(request, progressPath, [[currentUser,FULL],[admin,FULL],[qm,appendRole],[ownerGroup,READ],[reader,READ]]);
    await reconcileAcl(request, configPath, [[currentUser,FULL],[admin,FULL],[qm,READ],[ownerGroup,READ],[reader,READ]]);
    // An old uniquely secured settings item must not retain its previous writers.
    await request('POST', `${configPath}/items(${configItems[0].Id})/resetroleinheritance()`);
  }
  await request('MERGE', progressPath, { ContentTypesEnabled: true, EnableFolderCreation: true, EnableVersioning: true });
  await request('MERGE', issuesPath, { EnableVersioning: true });
  const folders = await all(request, `${progressPath}/RootFolder/Folders?$select=ServerRelativeUrl`);
  const folderUrls = new Set(folders.map(folder => folder.ServerRelativeUrl));
  for (const issue of issues) {
    const path = `${root}/issue-${issue.Id}`;
    if (!folderUrls.has(path)) {
      await request('POST', `web/folders/addUsingPath(DecodedUrl='${literal(path)}',overwrite=false)`);
    }
    if (production) {
      const [admin, qm, ownerGroup, reader] = groups;
      const base = [[currentUser,FULL],[admin,FULL],[ownerGroup,READ],[reader,READ]];
      await reconcileAcl(request, `${issuesPath}/items(${issue.Id})`, [...base,[qm,EDIT],...(issue.TaskOwnerId ? [[issue.TaskOwnerId,EDIT]] : [])]);
      await reconcileAcl(request, `${folderPath(path)}/ListItemAllFields`, [...base,[qm,appendRole],...(issue.TaskOwnerId ? [[issue.TaskOwnerId,appendRole]] : [])]);
    }
  }
  for (const entry of moves) {
    const target = `${root}/issue-${entry.ParentItemId}/${entry.FileRef.slice(entry.FileRef.lastIndexOf('/') + 1)}`;
    await request('POST', 'SP.MoveCopyUtil.MoveFileByPath(overwrite=@a1)?@a1=false', {
      srcPath: { __metadata: { type: 'SP.ResourcePath' }, DecodedUrl: new URL(site).origin + entry.FileRef },
      destPath: { __metadata: { type: 'SP.ResourcePath' }, DecodedUrl: new URL(site).origin + target },
      options: { __metadata: { type: 'SP.MoveCopyOptions' }, KeepBoth: false,
        ResetAuthorAndCreatedOnCopy: false, RetainEditorAndModifiedOnMove: true, ShouldBypassSharedLocks: false },
    });
    const moved = unwrap(await request('GET', `${progressPath}/items(${entry.Id})?$select=Id,FileDirRef,AuthorId,Created`));
    if (moved.FileDirRef !== `${root}/issue-${entry.ParentItemId}` || moved.AuthorId !== entry.AuthorId || moved.Created !== entry.Created) {
      throw new Error(`Verify journal item ${entry.Id}: move did not preserve its expected folder/author/time.`);
    }
  }
  if (production) for (const entry of entries.filter(row => row.FSObjType !== 1)) {
    await request('POST', `${progressPath}/items(${entry.Id})/resetroleinheritance()`);
  }
  log(`Journal ready: ${issues.length} issue folders, ${moves.length} existing entries moved; ${production ? 'production ACLs reconciled' : 'beta site permissions retained'}.`);
}
