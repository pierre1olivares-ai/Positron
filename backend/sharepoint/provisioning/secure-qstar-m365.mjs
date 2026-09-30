import { spawnSync } from 'node:child_process';
import { secureQstar } from './qstar-security-engine.mjs';
const env = process.env;
const site = env.SITE?.replace(/\/$/, '');
if (!site) throw new Error('SITE is required.');
const api = `${site}/_api/`;
async function request(method, path, body) {
  const url = path.startsWith('https://') ? path : `${api}${path}`;
  if (!url.toLowerCase().startsWith(api.toLowerCase())) throw new Error('Unexpected SharePoint pagination URL.');
  const args = ['request', '--url', url, '--method', method === 'MERGE' ? 'post' : method.toLowerCase(),
    '--accept', 'application/json;odata=nometadata', '--content-type', 'application/json;odata=nometadata', '--output', 'json'];
  if (method === 'MERGE') args.push('--x-http-method', 'MERGE', '--if-match', '*');
  if (body !== undefined) args.push('--body', JSON.stringify(body));
  const result = spawnSync('m365', args, {encoding:'utf8', maxBuffer:64 * 1024 * 1024});
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error(result.stderr || result.stdout || `SharePoint request failed: ${method} ${path}`);
  return result.stdout.trim() ? JSON.parse(result.stdout) : {};
}
await secureQstar({ request, site, issuesTitle: env.ISSUES || 'Q-Star Issues',
  progressTitle: env.PROGRESS || 'Q-Star Progress Log', configTitle: env.CONFIG || 'Q-Star Config',
  production: env.CREATE_ROLE_GROUPS !== '0', migration: env.PROGRESS_MIGRATION || 'preserve', preflight: env.QSTAR_SECURITY_PREFLIGHT === '1',
  groupNames: [env.ADMIN_GROUP || 'Q-Star Admins', env.QM_GROUP || 'Q-Star Quality Managers',
    env.OWNER_GROUP || 'Q-Star Task Owners', env.READER_GROUP || 'Q-Star Readers'] });
